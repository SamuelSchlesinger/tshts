//! File-writer abstraction. Mirrors the HTTP fetcher pattern: domain code
//! that needs to write a file (currently `CsvExporter::export_to_csv`) goes
//! through this trait instead of importing `crate::infrastructure` directly.
//!
//! At runtime the application installs the production writer (which lives
//! in `crate::infrastructure::atomic`) via [`set_file_writer`] during
//! startup. Tests can install a mock implementation the same way, or rely
//! on the default that returns a "no writer installed" error.

use std::sync::OnceLock;

/// File-writer contract used by CSV export. Implementations must be
/// thread-safe; today's only caller is single-threaded but the trait
/// object is shared by reference and stored in a global, so a future
/// concurrent call site shouldn't need to revisit the contract.
pub trait FileWriter: Send + Sync {
    /// Write `contents` to `path` atomically. The production impl uses
    /// `infrastructure::atomic::atomic_write` (temp file + fsync + rename
    /// + parent dir fsync); a test mock can be a memory map.
    fn write(&self, path: &str, contents: &[u8]) -> Result<(), String>;
}

static GLOBAL_FILE_WRITER: OnceLock<Box<dyn FileWriter>> = OnceLock::new();

/// Install the process-wide file writer. First caller wins (`OnceLock`
/// semantics); subsequent calls are no-ops. The application installs the
/// real writer at startup; tests that exercise the file-writing paths
/// install before any export call runs.
pub fn set_file_writer(writer: Box<dyn FileWriter>) {
    let _ = GLOBAL_FILE_WRITER.set(writer);
}

/// Write through whichever writer was installed via [`set_file_writer`].
/// Returns an error if no writer has been installed — tests that don't
/// exercise file output don't need to install one.
pub fn write_file(path: &str, contents: &[u8]) -> Result<(), String> {
    match GLOBAL_FILE_WRITER.get() {
        Some(w) => w.write(path, contents),
        None => Err("no FileWriter installed; call set_file_writer at startup".to_string()),
    }
}
