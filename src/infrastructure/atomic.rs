//! Atomic file writes.
//!
//! Writes go to a sibling temp file then `rename` into place. `rename` is
//! atomic on POSIX when source and destination are on the same filesystem,
//! so a crash mid-write can never leave a half-written target on disk —
//! callers see either the old content or the new content, never garbage.

use std::fs::File;
use std::io::{self, Write};
use std::path::Path;
use std::process;
use std::sync::atomic::{AtomicU64, Ordering};

/// Write `contents` to `path` atomically. The temp filename includes the
/// process id and a monotonic counter so concurrent writers (e.g. multiple
/// tshts instances racing on the shared sidecar clipboard) don't trample
/// each other's temp file. `fsync` runs before the rename so a power loss
/// in the window between `rename` landing and the temp file's data blocks
/// reaching disk can't leave a 0-byte target after reboot (the standard
/// ext4 `data=ordered` + delayed-allocation hazard).
///
/// After the rename we also fsync the *parent directory* — without that, the
/// directory entry pointing at the new inode is only in the page cache;
/// a power loss between rename and dir-fsync can leave the file referenced
/// by the old (deleted) inode after recovery. POSIX requires fsync of the
/// directory to make a rename durable.
///
/// If `path` already exists, the new file inherits its mode bits. Without
/// this, `File::create` writes 0644 and a previously-chmodded 0600 workbook
/// silently becomes world-readable after a save.
pub fn atomic_write(path: &str, contents: &[u8]) -> io::Result<()> {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    let tmp = format!("{}.{}.{}.tmp", path, process::id(), n);

    // Capture target's existing permissions so we can preserve them across
    // the rename. Failure to read (e.g. file didn't exist) is fine.
    let prior_perms = std::fs::metadata(path).ok().map(|m| m.permissions());

    // Write + fsync the temp file. On error, clean up.
    let write_result = (|| -> io::Result<()> {
        let mut f = File::create(&tmp)?;
        f.write_all(contents)?;
        if let Some(perms) = &prior_perms {
            // Ignore errors — perm preservation is best-effort.
            let _ = f.set_permissions(perms.clone());
        }
        f.sync_all()?;
        Ok(())
    })();
    if let Err(e) = write_result {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    if let Err(e) = std::fs::rename(&tmp, Path::new(path)) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }

    // fsync the parent directory so the rename itself is durable. Errors
    // here are non-fatal: some filesystems (FAT/exFAT on USB, some
    // distributed mounts) don't support directory fsync.
    if let Some(parent) = Path::new(path).parent() {
        let dir_path = if parent.as_os_str().is_empty() {
            Path::new(".")
        } else {
            parent
        };
        if let Ok(dir) = File::open(dir_path) {
            let _ = dir.sync_all();
        }
    }
    Ok(())
}

/// Concrete [`FileWriter`](crate::domain::services::FileWriter) backed by
/// this module's `atomic_write`. Wraps the raw function so the domain
/// layer can write through the trait without importing
/// `crate::infrastructure` directly.
struct AtomicFileWriter;

impl crate::domain::services::FileWriter for AtomicFileWriter {
    fn write(&self, path: &str, contents: &[u8]) -> Result<(), String> {
        atomic_write(path, contents).map_err(|e| e.to_string())
    }
}

/// Install this module's `atomic_write` as the global file writer used by
/// `CsvExporter::export_to_csv` (and any future domain-layer file output).
/// Call once at process startup before any export. Subsequent calls are
/// silently ignored (the global slot uses `OnceLock` semantics).
pub fn install_as_file_writer() {
    crate::domain::services::set_file_writer(Box::new(AtomicFileWriter));
}
