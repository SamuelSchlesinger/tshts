//! TSHTS - Terminal Spreadsheet
//!
//! A terminal-based spreadsheet application with formula support, built in Rust.
//! Features include cell editing, formula evaluation, file persistence, and
//! a comprehensive expression language for calculations.

use std::io;
use crossterm::{
    event::{self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind},
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
use presentation::{render_ui, InputHandler};


/// Entry point for the TSHTS terminal spreadsheet application.
///
/// Sets up the terminal interface, initializes the application state,
/// and runs the main event loop until the user quits.
///
/// # Errors
///
/// Returns an error if terminal setup fails or if there are issues
/// with the terminal interface during runtime.
fn main() -> Result<(), Box<dyn std::error::Error>> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let mut app = App::default();
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

/// Main application event loop.
///
/// Handles terminal rendering and keyboard input processing.
/// Continues running until the user presses 'q' in normal mode.
///
/// # Arguments
///
/// * `terminal` - Terminal interface for rendering
/// * `app` - Mutable reference to application state
///
/// # Errors
///
/// Returns an IO error if terminal operations fail.
fn run_app<B: Backend>(terminal: &mut Terminal<B>, app: &mut App) -> io::Result<()> {
    loop {
        terminal.draw(|f| render_ui(f, app))?;

        if let Event::Key(key) = event::read()? {
            if key.kind == KeyEventKind::Press {
                match key.code {
                    KeyCode::Char('q') if matches!(app.mode, application::AppMode::Normal) => return Ok(()),
                    _ => InputHandler::handle_key_event(app, key.code, key.modifiers),
                }
            }
        }
    }
}