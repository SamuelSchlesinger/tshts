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
pub fn atomic_write(path: &str, contents: &[u8]) -> io::Result<()> {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    let tmp = format!("{}.{}.{}.tmp", path, process::id(), n);

    // Write + fsync the temp file. On error, clean up.
    let write_result = (|| -> io::Result<()> {
        let mut f = File::create(&tmp)?;
        f.write_all(contents)?;
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
    Ok(())
}
