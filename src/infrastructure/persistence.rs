use crate::domain::Spreadsheet;
use std::fs;

pub struct FileRepository;

impl FileRepository {
    pub fn save_spreadsheet(spreadsheet: &Spreadsheet, filename: &str) -> Result<String, String> {
        match serde_json::to_string_pretty(spreadsheet) {
            Ok(json) => {
                match fs::write(filename, &json) {
                    Ok(_) => Ok(filename.to_string()),
                    Err(e) => Err(e.to_string()),
                }
            }
            Err(e) => Err(format!("Serialization failed: {}", e)),
        }
    }

    pub fn load_spreadsheet(filename: &str) -> Result<(Spreadsheet, String), String> {
        match fs::read_to_string(filename) {
            Ok(content) => {
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(spreadsheet) => Ok((spreadsheet, filename.to_string())),
                    Err(e) => Err(format!("Invalid file format - {}", e)),
                }
            }
            Err(e) => Err(e.to_string()),
        }
    }
}