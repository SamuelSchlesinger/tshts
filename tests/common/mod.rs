//! PTY-based end-to-end test harness for `tshts`.
//!
//! Spawns the real release binary inside a pseudo-tty, captures the bytes it
//! writes back into a `vt100::Parser`, and exposes a small ergonomic API so
//! tests read like a transcript of a human session:
//!
//! ```no_run
//! mod common;
//! use common::Harness;
//! use std::time::Duration;
//!
//! let mut h = Harness::new();
//! h.assert_contains("-- NORMAL --");
//! h.send_text("i");
//! h.assert_contains("-- INSERT --");
//! h.send_text("hello");
//! h.send_esc();
//! h.send_text(":q!");
//! h.send_enter();
//! assert_eq!(h.wait_for_exit(Duration::from_secs(2)), Some(0));
//! ```
//!
//! Design notes:
//! - The harness fixes a 30x120 virtual screen. Pick a small enough size that
//!   tests don't depend on incidental layout, but large enough that the status
//!   bar plus a few grid rows are visible.
//! - A background thread continuously drains the PTY master and feeds bytes
//!   into the `vt100::Parser` behind a `Mutex`. Tests poll the parser's screen.
//! - Every `send_*` sleeps a few ms after writing so the child has time to
//!   process the input and re-render before the next assertion. `wait_for_text`
//!   polls up to a timeout for slower transitions.
//! - `Drop` SIGKILLs the child process if it's still alive, so a failed
//!   assertion doesn't leak processes.

#![allow(dead_code)]

pub mod scenarios;

use portable_pty::{native_pty_system, CommandBuilder, MasterPty, PtySize};
use std::io::{Read, Write};
use std::sync::{Arc, Mutex};
use std::thread::JoinHandle;
use std::time::{Duration, Instant};

const ROWS: u16 = 30;
const COLS: u16 = 120;

/// The default settle time after each input — gives the child time to handle
/// the keystroke and emit the next frame before the test asserts. Sized to
/// cover the app's 100ms event-poll tick plus render time, with margin for
/// parallel-test CPU contention.
const INPUT_SETTLE: Duration = Duration::from_millis(120);

/// Default timeout for `wait_for_text` and friends.
pub const DEFAULT_TIMEOUT: Duration = Duration::from_secs(3);

/// A live PTY session running the `tshts` binary.
pub struct Harness {
    writer: Box<dyn Write + Send>,
    parser: Arc<Mutex<vt100::Parser>>,
    reader_thread: Option<JoinHandle<()>>,
    child: Box<dyn portable_pty::Child + Send + Sync>,
    /// Keep the master alive so the reader thread doesn't EOF prematurely.
    _master: Box<dyn MasterPty + Send>,
}

impl Harness {
    /// Spawn `tshts` with no arguments.
    pub fn new() -> Self {
        Self::with_args(&[])
    }

    /// Spawn `tshts` with the given command-line arguments (e.g. a filename to
    /// open on startup).
    pub fn with_args(args: &[&str]) -> Self {
        let pty_system = native_pty_system();
        // Under heavy parallel test load (16+ concurrent PTYs) macOS can
        // return a transient PTY-allocation error. Retry a few times.
        let mut pair_opt = None;
        let mut last_err: Option<String> = None;
        for attempt in 0..8 {
            match pty_system.openpty(PtySize {
                rows: ROWS,
                cols: COLS,
                pixel_width: 0,
                pixel_height: 0,
            }) {
                Ok(p) => {
                    pair_opt = Some(p);
                    break;
                }
                Err(e) => {
                    last_err = Some(format!("{}", e));
                    std::thread::sleep(Duration::from_millis(50 * (attempt + 1) as u64));
                }
            }
        }
        let pair = pair_opt.unwrap_or_else(|| {
            panic!("openpty failed after retries: {}", last_err.unwrap_or_default())
        });

        // CARGO_BIN_EXE_<name> is set by cargo for integration tests in
        // tests/*.rs — it points at the freshly-built binary.
        let bin = env!("CARGO_BIN_EXE_tshts");
        let mut cmd = CommandBuilder::new(bin);
        for a in args {
            cmd.arg(a);
        }
        // Force a known TERM so crossterm's terminfo lookups are predictable
        // across CI environments.
        cmd.env("TERM", "xterm-256color");
        cmd.env("NO_COLOR", "");

        let child = pair.slave.spawn_command(cmd).expect("spawn failed");
        // Drop slave so the reader sees EOF when the child exits.
        drop(pair.slave);

        let mut reader = pair.master.try_clone_reader().expect("clone reader");
        let writer = pair.master.take_writer().expect("take writer");

        let parser = Arc::new(Mutex::new(vt100::Parser::new(ROWS, COLS, 0)));
        let parser_clone = parser.clone();

        let reader_thread = std::thread::spawn(move || {
            let mut buf = [0u8; 8192];
            loop {
                match reader.read(&mut buf) {
                    Ok(0) => break,
                    Ok(n) => {
                        if let Ok(mut p) = parser_clone.lock() {
                            p.process(&buf[..n]);
                        }
                    }
                    Err(_) => break,
                }
            }
        });

        let h = Self {
            writer,
            parser,
            reader_thread: Some(reader_thread),
            child,
            _master: pair.master,
        };
        // Wait for the initial render. The first frame should include the
        // mode chip — we use that as the readiness signal.
        h.wait_for_text("-- NORMAL --", Duration::from_secs(5));
        h
    }

    // ----- Input -----

    /// Write raw bytes to the PTY's master end. Most tests should prefer the
    /// typed helpers below, but this is here for sending arbitrary sequences.
    pub fn send(&mut self, bytes: &[u8]) {
        self.writer.write_all(bytes).expect("PTY write failed");
        self.writer.flush().expect("PTY flush failed");
        std::thread::sleep(INPUT_SETTLE);
    }

    /// Send literal text. Special characters are NOT escaped — newlines in
    /// `s` will be sent as `\n`. For Enter use [`send_enter`].
    pub fn send_text(&mut self, s: &str) {
        self.send(s.as_bytes());
    }

    /// Send a `\r` (Enter / Return).
    pub fn send_enter(&mut self) {
        self.send(b"\r");
    }

    /// Send a `\x1b` (Escape).
    pub fn send_esc(&mut self) {
        self.send(b"\x1b");
    }

    pub fn send_tab(&mut self) {
        self.send(b"\t");
    }

    pub fn send_backspace(&mut self) {
        self.send(b"\x7f");
    }

    /// Send a Ctrl-modified ASCII letter, e.g. `send_ctrl('c')` sends 0x03.
    /// Case-insensitive on the argument.
    pub fn send_ctrl(&mut self, c: char) {
        let upper = c.to_ascii_uppercase();
        assert!(
            upper.is_ascii_uppercase(),
            "send_ctrl only supports ASCII letters, got {:?}",
            c
        );
        let byte = (upper as u8) - b'A' + 1;
        self.send(&[byte]);
    }

    /// Send an arrow key (xterm-style escape sequences).
    pub fn send_arrow(&mut self, dir: Arrow) {
        let seq: &[u8] = match dir {
            Arrow::Up => b"\x1b[A",
            Arrow::Down => b"\x1b[B",
            Arrow::Right => b"\x1b[C",
            Arrow::Left => b"\x1b[D",
        };
        self.send(seq);
    }

    // ----- Screen state -----

    /// Full screen contents as a `\n`-joined string, one line per row,
    /// trailing whitespace included. Useful for `assert!(screen.contains(...))`.
    pub fn screen_contents(&self) -> String {
        self.parser.lock().unwrap().screen().contents()
    }

    /// Get the text of a specific row (0-indexed from the top).
    pub fn row(&self, row: u16) -> String {
        let p = self.parser.lock().unwrap();
        let mut s = String::new();
        for col in 0..COLS {
            if let Some(cell) = p.screen().cell(row, col) {
                s.push_str(cell.contents());
            }
        }
        s
    }

    // ----- Semantic row helpers -----
    //
    // Concrete row indices change when the UI layout changes (e.g. adding a
    // sheet-tabs row pushed the formula bar down by 1). These helpers
    // centralize the layout knowledge so tests don't break on cosmetic
    // changes. The current layout is:
    //   row 0  — header (filename, cell ref, mode chip)
    //   row 1  — sheet tabs
    //   row 2  — formula bar
    //   row 3  — grid top border
    //   row 4  — column-letter row (A B C ...)
    //   row 5+ — data rows (5 = row 1, 6 = row 2, ...)

    /// Returns the formula-bar row. Use for assertions like
    /// `assert!(h.formula_bar().contains("=A1+B1"))`.
    pub fn formula_bar(&self) -> String {
        self.row(2)
    }

    /// Returns the rendered N-th data row (1-indexed, like a spreadsheet).
    /// `data_row(1)` returns A1's row, etc.
    pub fn data_row(&self, n: u16) -> String {
        self.row(4 + n)
    }

    /// Returns the status-bar text row. The status bar is a 3-row bordered
    /// block at the bottom of the screen (rows ROWS-3..ROWS); the actual
    /// text lives on the middle row. Use this to read the mode label,
    /// status messages, and (load-bearing for the scenario framework) the
    /// `SUM=… AVG=… COUNT=…` line that Visual mode publishes for the
    /// current selection.
    pub fn status_bar(&self) -> String {
        self.row(ROWS - 2)
    }

    /// Current cursor position as (row, col).
    pub fn cursor(&self) -> (u16, u16) {
        self.parser.lock().unwrap().screen().cursor_position()
    }

    /// Poll the screen up to `timeout` waiting for `needle` to appear.
    /// Returns true if found, false on timeout.
    pub fn wait_for_text(&self, needle: &str, timeout: Duration) -> bool {
        let start = Instant::now();
        loop {
            if self.screen_contents().contains(needle) {
                return true;
            }
            if start.elapsed() > timeout {
                return false;
            }
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    /// Wait for `needle` to be ABSENT from the screen.
    pub fn wait_for_absent(&self, needle: &str, timeout: Duration) -> bool {
        let start = Instant::now();
        loop {
            if !self.screen_contents().contains(needle) {
                return true;
            }
            if start.elapsed() > timeout {
                return false;
            }
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    /// Assert the screen contains `needle` within the default timeout. On
    /// failure, dumps the current screen for debugging.
    #[track_caller]
    pub fn assert_contains(&self, needle: &str) {
        if !self.wait_for_text(needle, DEFAULT_TIMEOUT) {
            panic!(
                "Expected screen to contain {:?} within {:?}. Current screen:\n{}",
                needle,
                DEFAULT_TIMEOUT,
                self.screen_contents()
            );
        }
    }

    /// Assert the screen does NOT contain `needle`. Polls briefly so that a
    /// pending redraw whose escape sequence is still in flight has a chance
    /// to land before we check.
    #[track_caller]
    pub fn assert_absent(&self, needle: &str) {
        // Drain pending output for up to ~400ms — that's enough for the app's
        // 100ms event-poll plus a render at heavy CPU contention. After
        // that we trust the state.
        let start = Instant::now();
        let drain = Duration::from_millis(400);
        while start.elapsed() < drain {
            // If the needle appears at any point during the drain window,
            // a render is in flight — keep waiting in case it gets cleared
            // (but most callers expect it to stay absent).
            std::thread::sleep(Duration::from_millis(50));
        }
        let contents = self.screen_contents();
        if contents.contains(needle) {
            panic!(
                "Expected screen NOT to contain {:?}. Current screen:\n{}",
                needle, contents
            );
        }
    }

    // ----- Process lifecycle -----

    /// Block up to `timeout` for the child process to exit. Returns the exit
    /// code if the process terminated, or `None` on timeout.
    pub fn wait_for_exit(&mut self, timeout: Duration) -> Option<u32> {
        let start = Instant::now();
        loop {
            match self.child.try_wait() {
                Ok(Some(status)) => return Some(status.exit_code()),
                Ok(None) => {}
                Err(_) => return None,
            }
            if start.elapsed() > timeout {
                return None;
            }
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    /// Returns true if the child process is no longer alive.
    pub fn has_exited(&mut self) -> bool {
        matches!(self.child.try_wait(), Ok(Some(_)))
    }
}

impl Drop for Harness {
    fn drop(&mut self) {
        // SIGKILL if the test left it running.
        let _ = self.child.kill();
        // Best-effort join; don't block test teardown for long.
        if let Some(t) = self.reader_thread.take() {
            let _ = t.join();
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub enum Arrow {
    Up,
    Down,
    Left,
    Right,
}
