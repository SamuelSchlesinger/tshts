//! Recent files list, persisted at `~/.config/tshts/recent.json`.
//!
//! Keeps up to 10 entries, most-recent first. Failures (no $HOME, IO error,
//! malformed JSON) degrade gracefully — recent-files is a UX nicety, not
//! load-bearing.

use std::path::PathBuf;

const MAX_ENTRIES: usize = 10;
const FILE_NAME: &str = "recent.json";

fn config_dir() -> Option<PathBuf> {
    let home = std::env::var_os("HOME")?;
    let mut p = PathBuf::from(home);
    p.push(".config");
    p.push("tshts");
    Some(p)
}

fn config_path() -> Option<PathBuf> {
    let mut p = config_dir()?;
    p.push(FILE_NAME);
    Some(p)
}

/// Load the recent-files list. Returns an empty vec on any error. Lazily
/// purges entries whose backing file no longer exists so the UI doesn't show
/// dead paths.
pub fn load() -> Vec<String> {
    let Some(path) = config_path() else { return Vec::new(); };
    let Ok(content) = std::fs::read_to_string(&path) else { return Vec::new(); };
    let list: Vec<String> = serde_json::from_str(&content).unwrap_or_default();
    list.into_iter()
        .filter(|p| std::path::Path::new(p).exists())
        .collect()
}

/// Record a file as most-recently-used, deduping prior entries and capping
/// to `MAX_ENTRIES`. Errors are silently ignored. Written atomically so a
/// crash mid-write can't corrupt the recent-files list.
pub fn add(filename: &str) {
    let mut list = load();
    list.retain(|f| f != filename);
    list.insert(0, filename.to_string());
    list.truncate(MAX_ENTRIES);
    let Some(dir) = config_dir() else { return; };
    let _ = std::fs::create_dir_all(&dir);
    let Some(path) = config_path() else { return; };
    if let Ok(json) = serde_json::to_string_pretty(&list)
        && let Some(path_str) = path.to_str()
    {
        let _ = crate::infrastructure::atomic::atomic_write(path_str, json.as_bytes());
    }
}
