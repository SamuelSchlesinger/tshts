//! Background HTTP fetcher with caching.
//!
//! `GET()` formula calls go through here. The first call for a URL returns
//! `Loading` (and enqueues a background fetch); subsequent calls return the
//! cached value until TTL expires. A single worker thread services all
//! requests so the formula evaluator never blocks the UI.

use std::collections::{HashMap, HashSet};
use std::sync::{Arc, Mutex, OnceLock, mpsc};
use std::sync::atomic::{AtomicU64, Ordering};
use std::thread;
use std::time::{Duration, Instant};

const CACHE_TTL: Duration = Duration::from_secs(300);
const REQUEST_TIMEOUT: Duration = Duration::from_secs(5);

#[derive(Clone, Debug)]
enum CacheEntry {
    Loading,
    Done { fetched_at: Instant, body: String },
    Error { body: String },
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
            let client = reqwest::blocking::Client::builder()
                .timeout(REQUEST_TIMEOUT)
                .build()
                .ok();
            while let Ok(url) = rx.recv() {
                let result = match &client {
                    Some(c) => c.get(&url).send().and_then(|r| r.text()),
                    None => reqwest::blocking::get(&url).and_then(|r| r.text()),
                };
                let entry = match result {
                    Ok(body) => CacheEntry::Done { fetched_at: Instant::now(), body },
                    Err(e) => CacheEntry::Error { body: format!("#ERROR: {}", e) },
                };
                {
                    let mut cache = worker.cache.lock().unwrap();
                    cache.insert(url.clone(), entry);
                }
                worker.in_flight.lock().unwrap().remove(&url);
                worker.completion_count.fetch_add(1, Ordering::Relaxed);
            }
        });

        arc
    })
}

/// Result returned synchronously to a `GET()` call.
pub enum FetchResult {
    Loading,
    Value(String),
    Error(String),
}

/// Look up a URL, returning the cached value if fresh, or `Loading` while
/// the background worker is fetching. Enqueues a request if no entry exists
/// or the cached one has expired.
pub fn fetch(url: &str) -> FetchResult {
    let i = inner();
    let now = Instant::now();
    let mut cache = i.cache.lock().unwrap();
    match cache.get(url) {
        Some(CacheEntry::Done { fetched_at, body }) if now.duration_since(*fetched_at) < CACHE_TTL => {
            FetchResult::Value(body.clone())
        }
        Some(CacheEntry::Loading) => FetchResult::Loading,
        Some(CacheEntry::Error { body }) => FetchResult::Error(body.clone()),
        _ => {
            // Stale or missing — enqueue once.
            cache.insert(url.to_string(), CacheEntry::Loading);
            drop(cache);
            let mut flight = i.in_flight.lock().unwrap();
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

/// Clears the entire cache. Bound to `:cache clear`.
pub fn clear_cache() {
    let i = inner();
    i.cache.lock().unwrap().clear();
}
