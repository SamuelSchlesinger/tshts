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

struct Inner {
    sender: mpsc::Sender<Snapshot>,
    enabled: AtomicBool,
    last_mark: Mutex<Option<Instant>>,
}

fn inner() -> &'static Inner {
    static INSTANCE: OnceLock<Inner> = OnceLock::new();
    INSTANCE.get_or_init(|| {
        let (tx, rx) = mpsc::channel::<Snapshot>();
        // Worker: receive snapshot, sleep until idle window elapses, save.
        thread::spawn(move || {
            while let Ok(snap) = rx.recv() {
                // Drain newer snapshots — we only care about the latest.
                let mut latest = snap;
                while let Ok(next) = rx.try_recv() {
                    latest = next;
                }
                // Save (errors are silent — autosave is a safety net,
                // explicit Ctrl+S surfaces failures to the user).
                let _ = FileRepository::save_workbook(&latest.workbook, &latest.filename);
            }
        });
        Inner {
            sender: tx,
            enabled: AtomicBool::new(false),
            last_mark: Mutex::new(None),
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
/// is dirty, and the last edit was more than IDLE ago, post a snapshot to
/// the worker for saving. Returns true if a save was queued.
pub fn maybe_save(workbook: &Workbook, filename: Option<&str>) -> bool {
    let i = inner();
    if !i.enabled.load(Ordering::Relaxed) {
        return false;
    }
    let Some(filename) = filename else { return false; };
    let mut guard = i
        .last_mark
        .lock()
        .expect("autosave last_mark mutex poisoned");
    let Some(when) = *guard else { return false; };
    if when.elapsed() < IDLE {
        return false;
    }
    *guard = None;
    let _ = i.sender.send(Snapshot {
        workbook: workbook.clone(),
        filename: filename.to_string(),
    });
    true
}
