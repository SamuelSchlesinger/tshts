//! Background HTTP fetcher with caching.
//!
//! `GET()` formula calls go through here. The first call for a URL returns
//! `Loading` (and enqueues a background fetch); subsequent calls return the
//! cached value until TTL expires. A single worker thread services all
//! requests so the formula evaluator never blocks the UI.
//!
//! Safety limits:
//! - URLs resolving to literal private/loopback/link-local IPs are rejected
//!   to block trivial SSRF (e.g. `=GET("http://169.254.169.254/...")`).
//! - Responses larger than `MAX_RESPONSE_BYTES` are truncated and reported.
//! - The cache and in-flight set are capped so a recalc loop with
//!   per-iteration unique URLs cannot exhaust memory.

use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::net::{IpAddr, ToSocketAddrs};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::{Arc, Mutex, OnceLock, mpsc};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::thread;
use std::time::{Duration, Instant};

/// Network requests are gated off by default — a workbook the user just
/// opened might contain hostile `=GET(...)` formulas; opening a file should
/// not silently exfiltrate to the network. The user enables fetches per
/// session via the `:net on` command.
static NETWORK_ENABLED: AtomicBool = AtomicBool::new(false);

/// Toggle whether `GET()` calls actually hit the network. When false, `fetch`
/// returns a sentinel error and no request is enqueued. Bound to `:net on/off`.
pub fn set_network_enabled(enabled: bool) {
    NETWORK_ENABLED.store(enabled, Ordering::Relaxed);
}

/// Whether `:net on` has been issued this session.
pub fn network_enabled() -> bool {
    NETWORK_ENABLED.load(Ordering::Relaxed)
}

const CACHE_TTL: Duration = Duration::from_secs(300);
/// How long to cache an error before re-attempting. Short so transient
/// failures (DNS, 5xx, brief in_flight overflow) clear themselves without
/// requiring `:cache clear`. Excel re-evaluates errors every recalc; this
/// is a compromise that avoids hammering a dead endpoint while still
/// allowing recovery within ~30s.
const ERROR_TTL: Duration = Duration::from_secs(30);
const REQUEST_TIMEOUT: Duration = Duration::from_secs(5);
const MAX_RESPONSE_BYTES: u64 = 5 * 1024 * 1024;
const MAX_IN_FLIGHT: usize = 256;
const MAX_CACHE_ENTRIES: usize = 1024;

#[derive(Clone, Debug)]
enum CacheEntry {
    Loading,
    Done { fetched_at: Instant, body: String },
    /// We don't carry the error body — `FetchResult::Error` is unit-only
    /// (no place for text to land in `Value::Error`). The timestamp is
    /// load-bearing for `ERROR_TTL`: we suppress retries to a known-bad
    /// URL until that window expires.
    Error { fetched_at: Instant },
}

struct Inner {
    cache: Mutex<HashMap<String, CacheEntry>>,
    in_flight: Mutex<HashSet<String>>,
    sender: mpsc::Sender<String>,
    completion_count: AtomicU64,
}

fn inner() -> &'static Arc<Inner> {
    static INSTANCE: OnceLock<Arc<Inner>> = OnceLock::new();
    INSTANCE.get_or_init(|| {
        let (tx, rx) = mpsc::channel::<String>();
        let arc = Arc::new(Inner {
            cache: Mutex::new(HashMap::new()),
            in_flight: Mutex::new(HashSet::new()),
            sender: tx,
            completion_count: AtomicU64::new(0),
        });

        // Spawn the single worker thread.
        let worker = Arc::clone(&arc);
        thread::spawn(move || {
            // Re-validate every redirect hop against the same SSRF rules as
            // the initial URL. Default reqwest policy follows up to 10 hops
            // unvalidated, so `https://attacker.com/redir` could 302 to
            // `http://169.254.169.254/...` and bypass every check below.
            let redirect_policy =
                reqwest::redirect::Policy::custom(|attempt| {
                    if attempt.previous().len() >= 10 {
                        return attempt.error("too many redirects");
                    }
                    match check_url_safety(attempt.url().as_str()) {
                        Ok(()) => attempt.follow(),
                        Err(reason) => attempt.error(format!("redirect blocked: {}", reason)),
                    }
                });
            let client = reqwest::blocking::Client::builder()
                .timeout(REQUEST_TIMEOUT)
                .redirect(redirect_policy)
                .user_agent(concat!("tshts/", env!("CARGO_PKG_VERSION")))
                .build()
                .ok();
            while let Ok(url) = rx.recv() {
                // Trap panics from perform_fetch — a bare panic would leave
                // the cache entry stuck on `Loading` forever, since the
                // post-fetch insert and in_flight.remove never run.
                // AssertUnwindSafe is safe here: a panic mid-fetch may leave
                // an HTTP connection half-read, but we discard the Response
                // anyway, and no shared state is mutated until the cache
                // insert below.
                let entry =
                    catch_unwind(AssertUnwindSafe(|| perform_fetch(client.as_ref(), &url)))
                        .unwrap_or_else(|_| CacheEntry::Error {
                            fetched_at: Instant::now(),
                        });
                {
                    let mut cache = worker
                        .cache
                        .lock()
                        .expect("fetcher cache mutex poisoned");
                    cache.insert(url.clone(), entry);
                }
                worker
                    .in_flight
                    .lock()
                    .expect("fetcher in_flight mutex poisoned")
                    .remove(&url);
                worker.completion_count.fetch_add(1, Ordering::Relaxed);
            }
        });

        arc
    })
}

fn perform_fetch(client: Option<&reqwest::blocking::Client>, url: &str) -> CacheEntry {
    // Refuse to fall back to bare `reqwest::blocking::get` when the
    // configured client failed to build: that path has no timeout, no
    // redirect-revalidation, and no user-agent, which silently strips
    // every SSRF defense for the rest of the session.
    let Some(c) = client else {
        return CacheEntry::Error { fetched_at: Instant::now() };
    };
    let send_result = c.get(url).send();
    let response = match send_result {
        Ok(r) => r,
        Err(_) => return CacheEntry::Error { fetched_at: Instant::now() },
    };
    // Reject by Content-Length BEFORE reading the body. A declared-oversize
    // response would otherwise allocate up to MAX_RESPONSE_BYTES + 1 below
    // before the post-read size check rejected it — that lets a 256-deep
    // in_flight queue inflate to ~1.3 GB of peak allocation.
    if let Some(len) = response.content_length()
        && len > MAX_RESPONSE_BYTES
    {
        return CacheEntry::Error { fetched_at: Instant::now() };
    }
    // Cap servers that don't honestly declare Content-Length: read at most
    // MAX_RESPONSE_BYTES + 1 so we can detect overflow without buffering the
    // full body.
    let mut buf = Vec::new();
    if response
        .take(MAX_RESPONSE_BYTES + 1)
        .read_to_end(&mut buf)
        .is_err()
    {
        return CacheEntry::Error { fetched_at: Instant::now() };
    }
    if buf.len() as u64 > MAX_RESPONSE_BYTES {
        return CacheEntry::Error { fetched_at: Instant::now() };
    }
    let body = String::from_utf8_lossy(&buf).into_owned();
    CacheEntry::Done { fetched_at: Instant::now(), body }
}

/// Reject URLs that resolve to IP literals in private/loopback/link-local
/// ranges, or to obvious hostnames like `localhost`. This catches the common
/// SSRF mistakes (e.g. accidentally pasting an internal URL into a sheet)
/// without trying to do full DNS-rebinding defense.
fn check_url_safety(url_str: &str) -> Result<(), String> {
    let url = reqwest::Url::parse(url_str).map_err(|e| format!("invalid URL: {}", e))?;
    match url.scheme() {
        "http" | "https" => {}
        s => return Err(format!("scheme '{}' not allowed", s)),
    }
    let host = url.host_str().ok_or_else(|| "URL has no host".to_string())?;
    let lower = host.to_ascii_lowercase();
    if lower == "localhost" || lower.ends_with(".localhost") || lower.ends_with(".local") {
        return Err("host is localhost/.local".to_string());
    }
    // url::Url::host_str returns IPv6 literals with surrounding brackets,
    // which `IpAddr::from_str` does not accept. Strip them before parsing.
    let ip_candidate = lower
        .strip_prefix('[')
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(lower.as_str());
    // Zone IDs on link-local IPv6 (`fe80::1%eth0`) make IpAddr::from_str
    // fail, which would let the SSRF check pass by default. Zone-scoped URLs
    // aren't legitimate web traffic in this app — reject outright.
    if ip_candidate.contains('%') {
        return Err("IPv6 zone IDs are not supported".to_string());
    }
    if let Ok(ip) = ip_candidate.parse::<IpAddr>() {
        if is_blocked_ip(&ip) {
            return Err(format!("{} is a private/loopback/link-local address", ip));
        }
        // Literal IP — nothing more to resolve.
        return Ok(());
    }
    // DNS-based SSRF: a hostname like `localtest.me` resolves to 127.0.0.1.
    // We can't fully defend against rebinding (DNS may return different A
    // records by the time reqwest actually connects), but rejecting URLs
    // whose first-resolved address is private catches the obvious attacks.
    let port = url.port_or_known_default().unwrap_or(80);
    let resolve_target = format!("{}:{}", host, port);
    match (resolve_target.as_str()).to_socket_addrs() {
        Ok(addrs) => {
            for addr in addrs {
                if is_blocked_ip(&addr.ip()) {
                    return Err(format!(
                        "{} resolves to {} (private/loopback/link-local)",
                        host,
                        addr.ip()
                    ));
                }
            }
            Ok(())
        }
        Err(_) => {
            // DNS failure isn't itself a safety issue — the actual GET will
            // fail at connect-time. Let it through so the user sees a useful
            // error rather than a confusing "DNS rejected" message.
            Ok(())
        }
    }
}

fn is_blocked_ip(ip: &IpAddr) -> bool {
    match ip {
        IpAddr::V4(v4) => {
            if v4.is_loopback()
                || v4.is_private()
                || v4.is_link_local()
                || v4.is_unspecified()
                || v4.is_broadcast()
            {
                return true;
            }
            let octets = v4.octets();
            // 0.0.0.0/8 "this network"
            if octets[0] == 0 {
                return true;
            }
            // 100.64.0.0/10 CGNAT
            if octets[0] == 100 && (octets[1] & 0xc0) == 64 {
                return true;
            }
            false
        }
        IpAddr::V6(v6) => {
            if v6.is_loopback() || v6.is_unspecified() {
                return true;
            }
            if let Some(v4) = v6.to_ipv4_mapped() {
                return is_blocked_ip(&IpAddr::V4(v4));
            }
            let seg = v6.segments();
            // fc00::/7 unique-local
            if (seg[0] & 0xfe00) == 0xfc00 {
                return true;
            }
            // fe80::/10 link-local
            if (seg[0] & 0xffc0) == 0xfe80 {
                return true;
            }
            false
        }
    }
}

/// Result returned synchronously to a `GET()` call.
///
/// `Error` is unit-only: the GET() formula function maps every error to
/// `Value::Error(ErrorKind::Value)`, which doesn't carry text, so an
/// error body would have nowhere to land. The cache still tracks
/// `CacheEntry::Error` with a timestamp for TTL purposes (so we don't
/// retry the same failing URL within `ERROR_TTL`).
pub enum FetchResult {
    Loading,
    Value(String),
    Error,
}

/// Concrete [`HttpFetcher`](crate::domain::services::HttpFetcher) backed by
/// this module's process-wide cache + worker thread. Wraps the raw `fetch`
/// fn so the domain layer can call HTTP through the trait without importing
/// `crate::infrastructure` directly.
struct ProductionFetcher;

impl crate::domain::services::HttpFetcher for ProductionFetcher {
    fn fetch(&self, url: &str) -> crate::domain::services::HttpFetchResult {
        use crate::domain::services::HttpFetchResult;
        match fetch(url) {
            FetchResult::Value(body) => HttpFetchResult::Value(body),
            FetchResult::Loading => HttpFetchResult::Loading,
            FetchResult::Error => HttpFetchResult::Error,
        }
    }
}

/// Install this module's fetcher as the global HTTP fetcher used by the
/// `GET()` formula function. Call once at process startup before any
/// formula evaluation. Subsequent calls are silently ignored (the global
/// slot uses `OnceLock` semantics).
pub fn install_as_http_fetcher() {
    crate::domain::services::set_http_fetcher(Box::new(ProductionFetcher));
}

/// Look up a URL, returning the cached value if fresh, or `Loading` while
/// the background worker is fetching. Enqueues a request if no entry exists
/// or the cached one has expired.
pub fn fetch(url: &str) -> FetchResult {
    // Hard gate: an unopened-into-this-session workbook can contain hostile
    // `=GET(...)` formulas. Until the user types `:net on`, every GET()
    // returns a static error and no request is enqueued.
    if !network_enabled() {
        return FetchResult::Error;
    }
    let i = inner();
    let now = Instant::now();
    let mut cache = i.cache.lock().expect("fetcher cache mutex poisoned");
    match cache.get(url) {
        Some(CacheEntry::Done { fetched_at, body }) if now.duration_since(*fetched_at) < CACHE_TTL => {
            FetchResult::Value(body.clone())
        }
        Some(CacheEntry::Loading) => FetchResult::Loading,
        Some(CacheEntry::Error { fetched_at })
            if now.duration_since(*fetched_at) < ERROR_TTL =>
        {
            FetchResult::Error
        }
        _ => {
            // Stale (TTL'd Done or Error) or missing — validate URL and
            // enqueue if safe.
            if check_url_safety(url).is_err() {
                cache.insert(
                    url.to_string(),
                    CacheEntry::Error { fetched_at: now },
                );
                return FetchResult::Error;
            }
            if cache.len() >= MAX_CACHE_ENTRIES {
                // Best-effort cleanup: evict any TTL'd Done or Error entries
                // before refusing. Avoids a permanently-broken cache when
                // older entries would have expired by now but nothing has
                // re-triggered their refresh.
                cache.retain(|_, e| match e {
                    CacheEntry::Done { fetched_at, .. } => {
                        now.duration_since(*fetched_at) < CACHE_TTL
                    }
                    CacheEntry::Error { fetched_at, .. } => {
                        now.duration_since(*fetched_at) < ERROR_TTL
                    }
                    CacheEntry::Loading => true,
                });
                if cache.len() >= MAX_CACHE_ENTRIES {
                    return FetchResult::Error;
                }
            }
            cache.insert(url.to_string(), CacheEntry::Loading);
            drop(cache);
            let mut flight = i
                .in_flight
                .lock()
                .expect("fetcher in_flight mutex poisoned");
            if flight.len() >= MAX_IN_FLIGHT {
                // Replace the Loading marker with an Error so any racing
                // caller that observed Loading sees a real error next call
                // instead of polling indefinitely. The error has TTL via
                // ERROR_TTL, so transient overflow auto-recovers.
                drop(flight);
                let mut cache = i.cache.lock().expect("fetcher cache mutex poisoned");
                cache.insert(
                    url.to_string(),
                    CacheEntry::Error { fetched_at: now },
                );
                return FetchResult::Error;
            }
            if flight.insert(url.to_string()) {
                let _ = i.sender.send(url.to_string());
            }
            FetchResult::Loading
        }
    }
}

/// Monotonic counter of completed fetches. The main loop watches this to
/// know when to trigger a recalc.
pub fn completion_count() -> u64 {
    inner().completion_count.load(Ordering::Relaxed)
}

/// Clears the entire cache. Bound to `:cache clear`. Also drops the
/// in_flight set so the in-flight cap doesn't stay artificially saturated
/// if the user invokes this during a slow fetch. Outstanding fetches will
/// still complete and write their results into the (now-fresh) cache.
pub fn clear_cache() {
    let i = inner();
    i.cache
        .lock()
        .expect("fetcher cache mutex poisoned")
        .clear();
    i.in_flight
        .lock()
        .expect("fetcher in_flight mutex poisoned")
        .clear();
}

/// Test-only scaffolding for exercising the GET() three-state contract
/// (Loading / Value / Error) hermetically. We can't stand up a real
/// local HTTP mock because `check_url_safety` blocks 127.0.0.1 — and
/// weakening that defense for tests would be a real footgun. Instead
/// we reach into the cache directly: seed the entry the test wants,
/// then drive `fetch()` (or anything that wraps it, like GET()) and
/// assert on the result.
#[cfg(test)]
pub(crate) mod test_hooks {
    use super::{inner, CacheEntry};
    use std::time::Instant;

    /// Seed `url` so the next `fetch(url)` returns `FetchResult::Value(body)`.
    pub(crate) fn seed_value(url: &str, body: &str) {
        let i = inner();
        i.cache
            .lock()
            .expect("fetcher cache mutex poisoned")
            .insert(
                url.to_string(),
                CacheEntry::Done {
                    fetched_at: Instant::now(),
                    body: body.to_string(),
                },
            );
    }

    /// Seed a cached error so the next `fetch(url)` returns
    /// `FetchResult::Error` synchronously. Useful for exercising the
    /// IFERROR-traps-GET path without depending on a flaky upstream.
    pub(crate) fn seed_error(url: &str) {
        let i = inner();
        i.cache
            .lock()
            .expect("fetcher cache mutex poisoned")
            .insert(
                url.to_string(),
                CacheEntry::Error { fetched_at: Instant::now() },
            );
    }

    /// Park `url` in the Loading state so the next `fetch(url)` returns
    /// `FetchResult::Loading` without enqueueing. Models the racy
    /// in-flight window where two evals see Loading before the worker
    /// publishes the result.
    pub(crate) fn seed_loading(url: &str) {
        let i = inner();
        i.cache
            .lock()
            .expect("fetcher cache mutex poisoned")
            .insert(url.to_string(), CacheEntry::Loading);
    }

    /// True if `url` currently has a Loading entry — i.e. a fetch is
    /// in flight (or seeded as such). Used to assert that `fetch()`
    /// enqueued a request on first call.
    pub(crate) fn is_loading(url: &str) -> bool {
        let i = inner();
        matches!(
            i.cache.lock().expect("fetcher cache mutex poisoned").get(url),
            Some(CacheEntry::Loading)
        )
    }

    /// Drop a single URL from the cache. Lets tests reset just their
    /// own seeded entry without nuking the global cache (which the
    /// parallel test runner is sharing).
    pub(crate) fn forget(url: &str) {
        let i = inner();
        i.cache
            .lock()
            .expect("fetcher cache mutex poisoned")
            .remove(url);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn blocks_loopback_v4() {
        assert!(check_url_safety("http://127.0.0.1/").is_err());
        assert!(check_url_safety("http://127.55.55.55/").is_err());
    }

    #[test]
    fn blocks_private_v4() {
        assert!(check_url_safety("http://10.0.0.1/").is_err());
        assert!(check_url_safety("http://192.168.1.1/").is_err());
        assert!(check_url_safety("http://172.16.0.5/").is_err());
        assert!(check_url_safety("http://172.31.255.255/").is_err());
    }

    #[test]
    fn blocks_link_local_v4() {
        assert!(check_url_safety("http://169.254.169.254/").is_err());
    }

    #[test]
    fn blocks_cgnat() {
        assert!(check_url_safety("http://100.64.0.1/").is_err());
        assert!(check_url_safety("http://100.127.255.255/").is_err());
    }

    #[test]
    fn blocks_loopback_v6() {
        assert!(check_url_safety("http://[::1]/").is_err());
    }

    #[test]
    fn blocks_ipv4_mapped_v6() {
        assert!(check_url_safety("http://[::ffff:127.0.0.1]/").is_err());
    }

    #[test]
    fn blocks_link_local_v6() {
        assert!(check_url_safety("http://[fe80::1]/").is_err());
    }

    #[test]
    fn blocks_localhost_name() {
        assert!(check_url_safety("http://localhost/").is_err());
        assert!(check_url_safety("http://something.localhost/").is_err());
        assert!(check_url_safety("http://printer.local/").is_err());
    }

    #[test]
    fn blocks_non_http_schemes() {
        assert!(check_url_safety("file:///etc/passwd").is_err());
        assert!(check_url_safety("ftp://example.com/").is_err());
    }

    #[test]
    fn allows_public_ips() {
        assert!(check_url_safety("http://1.1.1.1/").is_ok());
        assert!(check_url_safety("https://93.184.216.34/").is_ok());
    }

    #[test]
    fn blocks_ipv6_zone_ids() {
        // Zone IDs cause IpAddr::from_str to fail, which would otherwise
        // let the URL through. We reject them outright.
        assert!(check_url_safety("http://[fe80::1%eth0]/").is_err());
        assert!(check_url_safety("http://[%25some_zone]/").is_err());
    }

    #[test]
    fn allows_public_hostnames() {
        assert!(check_url_safety("https://example.com/").is_ok());
        assert!(check_url_safety("http://api.github.com/repos").is_ok());
    }
}
