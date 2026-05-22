//! Debounced auto-save background thread.
//!
//! When the user toggles auto-save on, the main loop calls
//! `autosave::mark_dirty()` after every modifying action. The worker thread
//! waits for activity to settle (no marks for `IDLE`), then writes the
//! current workbook to disk via a snapshot the main loop posts.
//!
//! Design: the main thread owns the canonical `Workbook` and posts a
//! cheap-clone snapshot via channel. The worker performs the disk write so
//! the UI stays responsive.

use std::sync::{Mutex, OnceLock, mpsc};
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use std::time::{Duration, Instant};

use crate::domain::Workbook;
use crate::infrastructure::FileRepository;

const IDLE: Duration = Duration::from_secs(30);

struct Snapshot {
    workbook: Workbook,
    filename: String,
}

/// Latest autosave outcome the main thread hasn't yet drained. Surface via
/// `take_status_message` so the UI can show "Auto-saved" on success or
/// "Auto-save failed: ..." on disk-full / permission errors, instead of
/// silently dropping data and reporting success.
struct Status {
    message: String,
    is_error: bool,
}

struct Inner {
    sender: mpsc::Sender<Snapshot>,
    status_rx: Mutex<mpsc::Receiver<Status>>,
    enabled: AtomicBool,
    last_mark: Mutex<Option<Instant>>,
    /// True between `maybe_save` queueing a snapshot and the worker reporting
    /// completion. Caller must NOT clear the global dirty flag while this is
    /// true — the user could edit during the worker's write and we'd lose
    /// those edits if we cleared dirty optimistically.
    in_flight: AtomicBool,
}

fn inner() -> &'static Inner {
    static INSTANCE: OnceLock<Inner> = OnceLock::new();
    INSTANCE.get_or_init(|| {
        let (tx, rx) = mpsc::channel::<Snapshot>();
        let (status_tx, status_rx) = mpsc::channel::<Status>();
        thread::spawn(move || {
            while let Ok(snap) = rx.recv() {
                let mut latest = snap;
                while let Ok(next) = rx.try_recv() {
                    latest = next;
                }
                let status = match FileRepository::save_workbook(&latest.workbook, &latest.filename) {
                    Ok(_) => Status {
                        message: format!("Auto-saved {}", latest.filename),
                        is_error: false,
                    },
                    Err(e) => Status {
                        message: format!("Auto-save failed: {}", e),
                        is_error: true,
                    },
                };
                let _ = status_tx.send(status);
                inner().in_flight.store(false, Ordering::Release);
            }
        });
        Inner {
            sender: tx,
            status_rx: Mutex::new(status_rx),
            enabled: AtomicBool::new(false),
            last_mark: Mutex::new(None),
            in_flight: AtomicBool::new(false),
        }
    })
}

pub fn enable() {
    inner().enabled.store(true, Ordering::Relaxed);
}

pub fn disable() {
    inner().enabled.store(false, Ordering::Relaxed);
}

#[allow(dead_code)]
pub fn is_enabled() -> bool {
    inner().enabled.load(Ordering::Relaxed)
}

/// Record that the workbook has been modified. The main loop calls this
/// once per dirtying action; the autosave thread waits for the idle window
/// before writing.
pub fn mark_dirty() {
    *inner()
        .last_mark
        .lock()
        .expect("autosave last_mark mutex poisoned") = Some(Instant::now());
}

/// Called from the main loop each tick. If autosave is enabled, the workbook
/// is dirty, the last edit was more than IDLE ago, and no save is already
/// in flight, post a snapshot to the worker for saving. Returns true if a
/// save was queued. Callers must NOT clear their dirty flag based on this
/// return — the write hasn't landed yet; poll `take_status_message` for the
/// outcome and only clear dirty on the non-error message.
pub fn maybe_save(workbook: &Workbook, filename: Option<&str>) -> bool {
    let i = inner();
    if !i.enabled.load(Ordering::Relaxed) {
        return false;
    }
    let Some(filename) = filename else { return false; };
    // Don't pile up snapshots if the worker hasn't finished the last one.
    if i.in_flight.load(Ordering::Acquire) {
        return false;
    }
    let mut guard = i
        .last_mark
        .lock()
        .expect("autosave last_mark mutex poisoned");
    let Some(when) = *guard else { return false; };
    if when.elapsed() < IDLE {
        return false;
    }
    *guard = None;
    i.in_flight.store(true, Ordering::Release);
    if i.sender.send(Snapshot {
        workbook: workbook.clone(),
        filename: filename.to_string(),
    }).is_err() {
        // Worker thread died — clear the flag so future attempts aren't
        // permanently blocked.
        i.in_flight.store(false, Ordering::Release);
        return false;
    }
    true
}

/// Drain the next autosave status (success or failure), if any. Returns
/// `(message, is_error)`. The main loop should call this each tick and
/// surface non-empty results to the user.
pub fn take_status_message() -> Option<(String, bool)> {
    let i = inner();
    let rx = i.status_rx.lock().ok()?;
    match rx.try_recv() {
        Ok(s) => Some((s.message, s.is_error)),
        Err(_) => None,
    }
}

/// True while the background worker is mid-write. The main loop must keep
/// the dirty flag set during this window so edits made during the write
/// don't get silently overwritten by the in-flight (older) snapshot.
#[allow(dead_code)]
pub fn is_in_flight() -> bool {
    inner().in_flight.load(Ordering::Acquire)
}

/// Force-queue a snapshot, bypassing the IDLE debounce and enabled gate.
/// Called from the shutdown path so a Ctrl+C / SIGTERM with dirty state
/// doesn't drop the user's last edits. Returns true if a save was queued.
pub fn flush_now(workbook: &Workbook, filename: Option<&str>) -> bool {
    let i = inner();
    let Some(filename) = filename else { return false; };
    i.in_flight.store(true, Ordering::Release);
    if i.sender.send(Snapshot {
        workbook: workbook.clone(),
        filename: filename.to_string(),
    }).is_err() {
        i.in_flight.store(false, Ordering::Release);
        return false;
    }
    true
}

/// Block until the worker thread reports completion (or the timeout fires).
/// Called from the shutdown path after `flush_now`.
pub fn wait_until_idle(timeout: std::time::Duration) {
    let i = inner();
    let start = std::time::Instant::now();
    while i.in_flight.load(Ordering::Acquire) {
        if start.elapsed() > timeout {
            return;
        }
        std::thread::sleep(std::time::Duration::from_millis(20));
    }
}
