//! HTTP-fetcher abstraction. Lets the `GET()` formula function call out
//! to a fetcher without the domain layer depending on infrastructure.
//!
//! At runtime the application installs the production fetcher (which
//! lives in `crate::infrastructure::fetcher`) via [`set_http_fetcher`]
//! during startup. Tests can install a mock implementation the same way,
//! or rely on the no-op default (every call returns `HttpFetchResult::Error`).

use std::sync::OnceLock;

/// Outcome of an `HttpFetcher::fetch` call. Three states keep the cell's
/// rendered value useful while the background fetcher works:
///
/// - `Value(body)` — cached body, fresh.
/// - `Loading` — request is in flight; caller should display a sentinel.
/// - `Error` — fetch failed (network disabled, blocked URL, HTTP error,
///   etc.). Unit-only because `Value::Error` doesn't carry text.
#[derive(Debug, Clone)]
pub enum HttpFetchResult {
    Loading,
    Value(String),
    Error,
}

/// HTTP fetcher contract that the GET() builtin calls into. Implementations
/// must be thread-safe (recalc dispatches `GET()` cells serially today, but
/// the trait object is shared by reference and could be invoked concurrently
/// in the future). The production impl lives in `infrastructure::fetcher`.
pub trait HttpFetcher: Send + Sync {
    /// Look up `url`, returning the current state. The trait makes no
    /// guarantees about whether a call mutates internal state (the
    /// production fetcher enqueues a background request on a cache miss).
    fn fetch(&self, url: &str) -> HttpFetchResult;
}

static GLOBAL_FETCHER: OnceLock<Box<dyn HttpFetcher>> = OnceLock::new();

/// Install the process-wide HTTP fetcher. Idempotent in the sense that the
/// first caller wins; subsequent calls are no-ops (matching `OnceLock`
/// semantics). The application calls this exactly once at startup with the
/// real fetcher; tests that want a mock should do the same before any
/// `GET()` evaluation runs.
pub fn set_http_fetcher(fetcher: Box<dyn HttpFetcher>) {
    let _ = GLOBAL_FETCHER.set(fetcher);
}

/// Look up `url` through whichever fetcher was installed via
/// [`set_http_fetcher`]. If no fetcher has been installed (e.g. unit tests
/// that don't exercise GET()), returns `HttpFetchResult::Error`.
pub fn http_fetch(url: &str) -> HttpFetchResult {
    match GLOBAL_FETCHER.get() {
        Some(f) => f.fetch(url),
        None => HttpFetchResult::Error,
    }
}
