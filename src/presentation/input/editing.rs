//! Submodule of `input` — see input/mod.rs.

use super::*;
use crate::application::App;
use crossterm::event::KeyCode;

impl InputHandler {
    pub(super) fn handle_editing_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.finish_editing();
            }
            KeyCode::Tab => {
                // Finish editing and move right instead of down
                app.finish_editing_move_right();
            }
            KeyCode::Esc => {
                app.cancel_editing();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.input.remove(char_to_byte_pos(&app.input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < char_count(&app.input) {
                    app.input.remove(char_to_byte_pos(&app.input, app.cursor_position));
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < char_count(&app.input) {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = char_count(&app.input);
            }
            KeyCode::Char(c) => {
                app.input.insert(char_to_byte_pos(&app.input, app.cursor_position), c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }

}
