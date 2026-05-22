//! Sidecar clipboard for richer copy/paste between tshts instances.
//!
//! When a tshts user copies cells, we write two clipboard formats:
//! 1. A TSV string prefixed with a sentinel header for the system clipboard
//!    so other apps see usable text and tshts can detect formula-rich copies.
//! 2. A JSON sidecar at `~/.cache/tshts/clipboard.json` containing full cell
//!    data (formulas, formats, comments). Paste prefers the sidecar if it's
//!    newer than the system clipboard's timestamp.
//!
//! The two formats coexist so tshts↔tshts pastes round-trip formulas while
//! tshts↔other-app pastes still work via plain TSV.

use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

use serde::{Deserialize, Serialize};

use crate::domain::CellData;

/// Sentinel prefix added to TSV copies. Paste detects this and falls back to
/// the sidecar JSON for full data.
pub const SENTINEL: &str = "#TSHTS-FORMULAS-V1\n";

#[derive(Serialize, Deserialize)]
pub struct SidecarClipboard {
    /// (row_offset, col_offset, cell_data) relative to top-left of source.
    pub cells: Vec<(usize, usize, CellData)>,
    pub source_row: usize,
    pub source_col: usize,
    /// Unix-millis timestamp. Compared against system-clipboard age to decide
    /// whether the sidecar is still fresh.
    pub timestamp_ms: u128,
}

fn config_dir() -> Option<PathBuf> {
    let home = std::env::var_os("HOME")?;
    let mut p = PathBuf::from(home);
    p.push(".cache");
    p.push("tshts");
    Some(p)
}

fn path() -> Option<PathBuf> {
    let mut p = config_dir()?;
    p.push("clipboard.json");
    Some(p)
}

/// Cap on sidecar JSON file size. Beyond this, the write is skipped — the
/// in-memory clipboard still works, but cross-instance paste for very large
/// copies falls back to TSV values. 1 MB is enough for ~10k typical cells.
const MAX_SIDECAR_BYTES: usize = 1_000_000;

pub fn write(cells: Vec<(usize, usize, CellData)>, source_row: usize, source_col: usize) {
    let Some(dir) = config_dir() else { return; };
    let _ = std::fs::create_dir_all(&dir);
    let Some(path) = path() else { return; };
    let payload = SidecarClipboard {
        cells,
        source_row,
        source_col,
        timestamp_ms: SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_millis())
            .unwrap_or(0),
    };
    if let Ok(json) = serde_json::to_string_pretty(&payload) {
        if json.len() > MAX_SIDECAR_BYTES {
            // Skip — caller will still see the in-memory clipboard.
            // Clear any stale sidecar so paste doesn't pick up old data.
            let _ = std::fs::remove_file(&path);
            return;
        }
        let Some(path_str) = path.to_str() else { return; };
        let _ = crate::infrastructure::atomic::atomic_write(path_str, json.as_bytes());
    }
}

/// Remove the sidecar file. Used by `:clipboard clear`.
pub fn clear() {
    if let Some(p) = path() {
        let _ = std::fs::remove_file(p);
    }
}

/// Returns the sidecar payload if it exists and looks well-formed.
pub fn read() -> Option<SidecarClipboard> {
    let path = path()?;
    // Cap the read at MAX_SIDECAR_BYTES so a hostile or runaway tshts
    // instance can't plant a multi-GB clipboard.json that OOMs us on the
    // next paste. The writer enforces the same cap; this is defense-in-
    // depth in case a stale file from a previous build slipped past.
    use std::io::Read;
    let mut f = std::fs::File::open(&path).ok()?;
    let mut buf = Vec::with_capacity(4096);
    (&mut f)
        .take((MAX_SIDECAR_BYTES + 1) as u64)
        .read_to_end(&mut buf)
        .ok()?;
    if buf.len() > MAX_SIDECAR_BYTES {
        return None;
    }
    let content = String::from_utf8(buf).ok()?;
    serde_json::from_str(&content).ok()
}

/// Strip the sentinel prefix from a TSV string if present.
pub fn strip_sentinel(tsv: &str) -> Option<&str> {
    tsv.strip_prefix(SENTINEL)
}
