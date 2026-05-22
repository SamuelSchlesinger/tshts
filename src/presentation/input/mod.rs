use crate::application::{App, AppMode};
#[allow(unused_imports)] // re-exported via super::* for submodules
use crate::application::{VimOperator, VisualKind};
use crate::infrastructure::FileRepository;
use crate::domain::CsvExporter;
use crossterm::event::{KeyCode, KeyModifiers};

fn char_to_byte_pos(s: &str, char_pos: usize) -> usize {
    s.char_indices()
        .nth(char_pos)
        .map(|(byte_idx, _)| byte_idx)
        .unwrap_or(s.len())
}

fn char_count(s: &str) -> usize {
    s.chars().count()
}

pub struct InputHandler;


mod normal;
mod visual;
mod editing;
mod dialogs;

impl InputHandler {
    pub fn handle_key_event(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        // Drop modifier-laden Char events for text-input dialogs that have
        // no Ctrl/Alt bindings of their own. Without this guard, pressing
        // Ctrl+S inside the search/goto/filename/command-palette buffer
        // literally inserts the character `s` into the buffer.
        // Editing-mode and FindReplace handle their own modifier semantics
        // (FindReplace owns Ctrl+A for "replace all"); Normal/Visual route
        // through their own modifier-aware dispatchers.
        if matches!(
            app.mode,
            AppMode::SaveAs
                | AppMode::LoadFile
                | AppMode::ExportCsv
                | AppMode::ImportCsv
                | AppMode::Search
                | AppMode::GoToCell
                | AppMode::CommandPalette
                | AppMode::Editing
        ) && matches!(key, KeyCode::Char(_))
            && modifiers
                .intersects(KeyModifiers::CONTROL | KeyModifiers::ALT | KeyModifiers::SUPER)
        {
            return;
        }
        match app.mode {
            AppMode::Normal => Self::handle_normal_mode(app, key, modifiers),
            AppMode::Editing => Self::handle_editing_mode(app, key),
            AppMode::Visual { .. } => Self::handle_visual_mode(app, key, modifiers),
            AppMode::Help => Self::handle_help_mode(app, key),
            AppMode::SaveAs => Self::handle_filename_input_mode(app, key, "save"),
            AppMode::LoadFile => Self::handle_filename_input_mode(app, key, "load"),
            AppMode::ExportCsv => Self::handle_filename_input_mode(app, key, "csv_export"),
            AppMode::ImportCsv => Self::handle_filename_input_mode(app, key, "csv_import"),
            AppMode::Search => Self::handle_search_mode(app, key),
            AppMode::GoToCell => Self::handle_goto_cell_mode(app, key),
            AppMode::FindReplace => Self::handle_find_replace_mode(app, key, modifiers),
            AppMode::CommandPalette => Self::handle_command_palette_mode(app, key),
            AppMode::ConfirmDiscard => Self::handle_confirm_discard_mode(app, key),
        }
    }

}

