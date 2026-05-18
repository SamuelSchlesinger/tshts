//! TSHTS - Terminal Spreadsheet
//!
//! A terminal-based spreadsheet application with formula support, built in Rust.
//! Features include cell editing, formula evaluation, file persistence, and
//! a comprehensive expression language for calculations.

use std::io;
use std::time::Duration;
use crossterm::{
    event::{self, DisableMouseCapture, EnableMouseCapture, Event, KeyEventKind, MouseEventKind},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::{Backend, CrosstermBackend},
    Terminal,
};

mod domain;
mod application;
mod infrastructure;
mod presentation;

use application::App;
use infrastructure::{autosave, fetcher, FileRepository};
use presentation::{render_ui, InputHandler};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    // Mouse capture is on by default. Users who want host-terminal text
    // selection can hold Shift (or Option on macOS) to bypass the capture.
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let mut app = App::default();

    // Optional file argument. Auto-detects `.xlsx` vs `.tshts` by extension.
    if let Some(path) = std::env::args().nth(1) {
        let lower = path.to_lowercase();
        let result = if lower.ends_with(".xlsx") {
            infrastructure::xlsx::load_xlsx(&path).map(|wb| (wb, path.clone()))
        } else {
            FileRepository::load_workbook(&path)
        };
        match result {
            Ok((workbook, filename)) => {
                app.workbook = workbook;
                app.filename = Some(filename.clone());
                app.dirty = false;
                infrastructure::recent::add(&filename);
                app.status_message = Some(format!("Loaded {}", filename));
            }
            Err(e) => {
                app.status_message = Some(format!("Failed to load {}: {}", path, e));
            }
        }
    }

    let res = run_app(&mut terminal, &mut app);

    disable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        LeaveAlternateScreen,
        DisableMouseCapture
    )?;
    terminal.show_cursor()?;

    if let Err(err) = res {
        println!("{err:?}");
    }

    Ok(())
}

fn run_app<B: Backend>(terminal: &mut Terminal<B>, app: &mut App) -> io::Result<()> {
    let mut last_fetch_count = fetcher::completion_count();
    loop {
        terminal.draw(|f| render_ui(f, app))?;

        if app.should_quit {
            return Ok(());
        }

        // Poll with a 100ms timeout so background fetch completions can trigger
        // a redraw and recalc even if the user is idle.
        if event::poll(Duration::from_millis(100))? {
            match event::read()? {
                Event::Key(key) if key.kind == KeyEventKind::Press => {
                    InputHandler::handle_key_event(app, key.code, key.modifiers);
                }
                Event::Mouse(m) => match m.kind {
                    MouseEventKind::ScrollDown => {
                        if app.scroll_row + 1 < app.workbook.current_sheet().rows {
                            app.scroll_row += 1;
                        }
                    }
                    MouseEventKind::ScrollUp => {
                        if app.scroll_row > 0 {
                            app.scroll_row -= 1;
                        }
                    }
                    MouseEventKind::Down(_) => {
                        // Click: select the cell and arm a selection anchor so
                        // a subsequent Drag can extend the range.
                        if let Some((row, col)) =
                            presentation::cell_at(app, m.column as usize, m.row as usize)
                        {
                            app.clear_selection();
                            app.selected_row = row;
                            app.selected_col = col;
                            app.start_selection();
                            app.ensure_cursor_visible();
                        }
                    }
                    MouseEventKind::Drag(_) => {
                        // Drag: extend the selection from the click anchor to
                        // the cell under the cursor.
                        if let Some((row, col)) =
                            presentation::cell_at(app, m.column as usize, m.row as usize)
                        {
                            if app.selection_start.is_none() {
                                app.start_selection();
                            }
                            app.selected_row = row;
                            app.selected_col = col;
                            app.update_selection(row, col);
                            app.ensure_cursor_visible();
                        }
                    }
                    MouseEventKind::Up(_) => {
                        // Finalize: if the user clicked without dragging, the
                        // selection collapses back to a single cell so a fresh
                        // click doesn't carry the previous range forward.
                        if let (Some(start), Some(end)) = (app.selection_start, app.selection_end)
                            && start == end
                        {
                            app.clear_selection();
                        }
                    }
                    _ => {}
                },
                _ => {}
            }
        }

        // If background fetches completed since last frame, recalc formulas
        // so newly-arrived GET() values appear without user action.
        let count = fetcher::completion_count();
        if count != last_fetch_count {
            last_fetch_count = count;
            app.recalc_all();
        }

        // Idle auto-save check.
        if app.dirty && autosave::maybe_save(&app.workbook, app.filename.as_deref()) {
            app.dirty = false;
            app.status_message = Some("Auto-saved".to_string());
        }
    }
}