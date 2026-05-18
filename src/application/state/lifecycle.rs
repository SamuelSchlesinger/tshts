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
