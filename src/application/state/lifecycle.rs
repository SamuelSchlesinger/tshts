//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn request_quit(&mut self) {
        if self.dirty {
            self.pending_action = Some(PendingAction::Quit);
            self.mode = AppMode::ConfirmDiscard;
            self.status_message = Some(
                "Unsaved changes — quit anyway? (y=quit, n=cancel, s=save & quit)".to_string(),
            );
        } else {
            self.should_quit = true;
        }
    }

    pub fn request_load_file(&mut self) {
        if self.dirty {
            self.pending_action = Some(PendingAction::LoadFile);
            self.mode = AppMode::ConfirmDiscard;
            self.status_message = Some(
                "Unsaved changes — load new file anyway? (y=discard, n=cancel)".to_string(),
            );
        } else {
            self.start_load_file();
        }
    }

    pub fn confirm_pending_action(&mut self) {
        let action = self.pending_action.take();
        self.mode = AppMode::Normal;
        self.status_message = None;
        match action {
            Some(PendingAction::Quit) => self.should_quit = true,
            Some(PendingAction::LoadFile) => self.start_load_file(),
            None => {}
        }
    }

    pub fn cancel_pending_action(&mut self) {
        self.pending_action = None;
        self.mode = AppMode::Normal;
        self.status_message = Some("Cancelled".to_string());
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_confirm_discard_save_then_quit_with_known_filename() {
        use tempfile::NamedTempFile;
        let tmp = NamedTempFile::new().unwrap();
        let path = tmp.path().to_str().unwrap().to_string();

        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "x".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        app.dirty = true;
        app.filename = Some(path);

        // Quit with dirty → prompt.
        app.request_quit();
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        assert!(!app.should_quit);

        // Simulate "s" (save & quit). save_in_place_or_prompt should succeed
        // because filename is known.
        let pending = app.pending_action.take();
        app.save_in_place_or_prompt();
        assert!(!app.dirty);
        // Trigger the deferred quit.
        if let Some(action) = pending {
            app.pending_action = Some(action);
            app.confirm_pending_action();
        }
        assert!(app.should_quit);
    }

}
