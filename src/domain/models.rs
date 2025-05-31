use std::collections::HashMap;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellData {
    pub value: String,
    pub formula: Option<String>,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            value: String::new(),
            formula: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Spreadsheet {
    #[serde(serialize_with = "serialize_cells", deserialize_with = "deserialize_cells")]
    pub cells: HashMap<(usize, usize), CellData>,
    pub rows: usize,
    pub cols: usize,
    pub column_widths: HashMap<usize, usize>,
    pub default_column_width: usize,
}

impl Default for Spreadsheet {
    fn default() -> Self {
        Self {
            cells: HashMap::new(),
            rows: 100,
            cols: 26,
            column_widths: HashMap::new(),
            default_column_width: 8,
        }
    }
}

impl Spreadsheet {
    pub fn get_cell(&self, row: usize, col: usize) -> CellData {
        self.cells.get(&(row, col)).cloned().unwrap_or_default()
    }

    pub fn set_cell(&mut self, row: usize, col: usize, data: CellData) {
        self.cells.insert((row, col), data.clone());
        
        let current_width = self.get_column_width(col);
        let value_width = data.value.len();
        let formula_width = data.formula.as_ref().map(|f| f.len()).unwrap_or(0);
        let content_width = value_width.max(formula_width);
        let header_width = Self::column_label(col).len();
        let needed_width = content_width.max(header_width).max(3).min(50);
        
        if needed_width > current_width {
            self.set_column_width(col, needed_width);
        }
    }

    pub fn get_cell_value_for_formula(&self, row: usize, col: usize) -> f64 {
        let cell = self.get_cell(row, col);
        cell.value.parse::<f64>().unwrap_or(0.0)
    }

    pub fn column_label(col: usize) -> String {
        let mut result = String::new();
        let mut c = col;
        loop {
            result = char::from(b'A' + (c % 26) as u8).to_string() + &result;
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        result
    }

    pub fn parse_cell_reference(cell_ref: &str) -> Option<(usize, usize)> {
        if cell_ref.is_empty() {
            return None;
        }
        
        let mut chars = cell_ref.chars();
        let mut col_str = String::new();
        let mut row_str = String::new();
        
        for ch in chars.by_ref() {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch.to_ascii_uppercase());
            } else if ch.is_ascii_digit() {
                row_str.push(ch);
                break;
            } else {
                return None;
            }
        }
        
        for ch in chars {
            if ch.is_ascii_digit() {
                row_str.push(ch);
            } else {
                return None;
            }
        }
        
        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }
        
        let col = Self::column_str_to_index(&col_str)?;
        let row = row_str.parse::<usize>().ok()?.checked_sub(1)?;
        
        Some((row, col))
    }
    
    fn column_str_to_index(col_str: &str) -> Option<usize> {
        if col_str.is_empty() {
            return None;
        }
        
        let mut result = 0;
        for ch in col_str.chars() {
            if !ch.is_ascii_alphabetic() {
                return None;
            }
            result = result * 26 + (ch as usize - 'A' as usize + 1);
        }
        Some(result - 1)
    }

    pub fn get_column_width(&self, col: usize) -> usize {
        self.column_widths.get(&col).copied().unwrap_or(self.default_column_width)
    }

    pub fn set_column_width(&mut self, col: usize, width: usize) {
        self.column_widths.insert(col, width);
    }

    pub fn auto_resize_column(&mut self, col: usize) {
        let current_width = self.get_column_width(col);
        let mut max_width = Self::column_label(col).len().max(current_width);
        
        for row in 0..self.rows {
            let cell = self.get_cell(row, col);
            let value_width = cell.value.len();
            let formula_width = cell.formula.as_ref().map(|f| f.len()).unwrap_or(0);
            let content_width = value_width.max(formula_width);
            max_width = max_width.max(content_width);
        }
        
        max_width = max_width.max(3).min(50);
        if max_width > current_width {
            self.set_column_width(col, max_width);
        }
    }

    pub fn auto_resize_all_columns(&mut self) {
        for col in 0..self.cols {
            self.auto_resize_column(col);
        }
    }
}

fn serialize_cells<S>(cells: &HashMap<(usize, usize), CellData>, serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    use serde::ser::SerializeSeq;
    let mut seq = serializer.serialize_seq(Some(cells.len()))?;
    for (key, value) in cells {
        seq.serialize_element(&(key.0, key.1, value))?;
    }
    seq.end()
}

fn deserialize_cells<'de, D>(deserializer: D) -> Result<HashMap<(usize, usize), CellData>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    use serde::de::{SeqAccess, Visitor};
    use std::fmt;

    struct CellsVisitor;

    impl<'de> Visitor<'de> for CellsVisitor {
        type Value = HashMap<(usize, usize), CellData>;

        fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
            formatter.write_str("a sequence of cell data")
        }

        fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
        where
            A: SeqAccess<'de>,
        {
            let mut cells = HashMap::new();
            while let Some((row, col, data)) = seq.next_element::<(usize, usize, CellData)>()? {
                cells.insert((row, col), data);
            }
            Ok(cells)
        }
    }

    deserializer.deserialize_seq(CellsVisitor)
}