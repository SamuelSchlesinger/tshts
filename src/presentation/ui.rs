use crate::application::{App, AppMode};
use crate::domain::Spreadsheet;
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Style},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame,
};

pub fn render_ui(f: &mut Frame, app: &App) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Min(0),
            Constraint::Length(3),
        ])
        .split(f.area());

    render_header(f, app, chunks[0]);
    render_spreadsheet(f, app, chunks[1]);
    render_status_bar(f, app, chunks[2]);

    if matches!(app.mode, AppMode::Help) {
        render_help_popup(f, app.help_scroll);
    }
}

fn render_header(f: &mut Frame, app: &App, area: Rect) {
    let header = Paragraph::new(format!(
        "tshts - Terminal Spreadsheet | Cell: {}{}",
        Spreadsheet::column_label(app.selected_col),
        app.selected_row + 1
    ))
    .style(Style::default().fg(Color::Cyan));
    f.render_widget(header, area);
}

fn render_spreadsheet(f: &mut Frame, app: &App, area: Rect) {
    let visible_rows = area.height as usize - 1;
    
    let mut total_width = 4;
    let mut visible_cols = 0;
    let available_width = area.width as usize;
    
    for col in app.scroll_col..app.spreadsheet.cols {
        let col_width = app.spreadsheet.get_column_width(col);
        if total_width + col_width + 1 > available_width {
            break;
        }
        total_width += col_width + 1;
        visible_cols += 1;
    }
    
    let mut headers = vec![Cell::from("")];
    for col in app.scroll_col..app.scroll_col + visible_cols {
        let header_style = if col == app.selected_col {
            Style::default().bg(Color::LightBlue).fg(Color::Black)
        } else {
            Style::default().fg(Color::Yellow)
        };
        headers.push(Cell::from(Spreadsheet::column_label(col)).style(header_style));
    }

    let header_row = Row::new(headers).height(1);
    
    let mut rows = vec![header_row];
    
    for row in app.scroll_row..std::cmp::min(app.scroll_row + visible_rows, app.spreadsheet.rows) {
        let row_number_style = if row == app.selected_row {
            Style::default().bg(Color::LightBlue).fg(Color::Black)
        } else {
            Style::default().fg(Color::Yellow)
        };
        let mut cells = vec![Cell::from(format!("{}", row + 1)).style(row_number_style)];
        
        for col in app.scroll_col..app.scroll_col + visible_cols {
            let cell_data = app.spreadsheet.get_cell(row, col);
            let cell_value = if cell_data.value.is_empty() { " ".to_string() } else { cell_data.value };
            
            let style = if row == app.selected_row && col == app.selected_col {
                Style::default().bg(Color::Blue).fg(Color::White)
            } else {
                Style::default()
            };
            
            cells.push(Cell::from(cell_value).style(style));
        }
        
        rows.push(Row::new(cells).height(1));
    }

    let mut widths = vec![Constraint::Length(4)];
    for col in app.scroll_col..app.scroll_col + visible_cols {
        widths.push(Constraint::Length(app.spreadsheet.get_column_width(col) as u16));
    }
    let table = Table::new(rows, widths)
        .block(Block::default().borders(Borders::ALL).title("Spreadsheet"))
        .column_spacing(1);

    f.render_widget(table, area);
}

fn render_status_bar(f: &mut Frame, app: &App, area: Rect) {
    let input_text = match app.mode {
        AppMode::Normal => {
            if let Some(ref status) = app.status_message {
                status.clone()
            } else {
                let filename = app.filename.as_ref().map(|f| f.as_str()).unwrap_or("unsaved");
                format!("File: {} | Ctrl+S: save | Ctrl+O: load | Ctrl+E: export CSV | Ctrl+L: import CSV | F1/?: help | q: quit", filename)
            }
        }
        AppMode::Editing => format!("Editing: {} (Enter to save, Esc to cancel)", app.input),
        AppMode::Help => "↑↓/jk: scroll | PgUp/PgDn: fast scroll | Home: top | Esc/q: close help".to_string(),
        AppMode::SaveAs => format!("Save as: {} (Enter to save, Esc to cancel)", app.filename_input),
        AppMode::LoadFile => format!("Load file: {} (Enter to load, Esc to cancel)", app.filename_input),
        AppMode::ExportCsv => format!("Export CSV as: {} (Enter to export, Esc to cancel)", app.filename_input),
        AppMode::ImportCsv => format!("Import CSV from: {} (Enter to import, Esc to cancel)", app.filename_input),
    };

    let input = Paragraph::new(input_text)
        .block(Block::default().borders(Borders::ALL).title("Status"))
        .style(match app.mode {
            AppMode::Normal => Style::default(),
            AppMode::Editing => Style::default().fg(Color::Green),
            AppMode::Help => Style::default().fg(Color::Cyan),
            AppMode::SaveAs => Style::default().fg(Color::Yellow),
            AppMode::LoadFile => Style::default().fg(Color::Yellow),
            AppMode::ExportCsv => Style::default().fg(Color::Magenta),
            AppMode::ImportCsv => Style::default().fg(Color::Green),
        });
    f.render_widget(input, area);
}

fn render_help_popup(f: &mut Frame, scroll: usize) {
    let area = f.area();
    let popup_area = Rect {
        x: area.width / 10,
        y: area.height / 10,
        width: area.width * 4 / 5,
        height: area.height * 4 / 5,
    };

    f.render_widget(Clear, popup_area);
    
    let help_text = get_help_text();
    let help_lines: Vec<&str> = help_text.lines().collect();
    let visible_height = popup_area.height.saturating_sub(2) as usize;
    
    let start_line = scroll.min(help_lines.len().saturating_sub(visible_height));
    let end_line = (start_line + visible_height).min(help_lines.len());
    
    let visible_text = help_lines[start_line..end_line].join("\n");
    
    let help_widget = Paragraph::new(visible_text)
        .block(Block::default()
            .borders(Borders::ALL)
            .title(format!("tshts Expression Language Help (Line {}/{})", start_line + 1, help_lines.len()))
            .style(Style::default().fg(Color::Cyan)))
        .style(Style::default().fg(Color::White));
    
    f.render_widget(help_widget, popup_area);
}

fn get_help_text() -> String {
    r#"TSHTS EXPRESSION LANGUAGE REFERENCE

=== BASIC CONCEPTS ===
• All formulas start with = (equals sign)
• Cell references use column letter + row number (A1, B2, Z99, AA1, etc.)
• Numbers can be integers or decimals (42, 3.14, -5.5)
• Case insensitive for functions and cell references

=== ARITHMETIC OPERATORS ===
+       Addition                    =5+3 → 8, =A1+B1
-       Subtraction                 =10-3 → 7, =A1-5
*       Multiplication              =4*3 → 12, =A1*B1
/       Division                    =15/3 → 5, =A1/B1
**      Exponentiation              =2**3 → 8, =A1**2
^       Power (same as **)          =3^2 → 9, =A1^B1
%       Modulo (remainder)          =10%3 → 1, =A1%B1

=== COMPARISON OPERATORS ===
<       Less than                   =A1<B1 → 1 or 0
>       Greater than                =A1>B1 → 1 or 0
<=      Less than or equal          =A1<=B1 → 1 or 0
>=      Greater than or equal       =A1>=B1 → 1 or 0
<>      Not equal                   =A1<>B1 → 1 or 0

Note: Comparisons return 1 for true, 0 for false

=== BASIC FUNCTIONS ===
SUM(...)        Sum of values           =SUM(A1,B1,C1) or =SUM(A1:C1)
AVERAGE(...)    Average of values       =AVERAGE(A1:A10)
MIN(...)        Minimum value           =MIN(A1,B1,C1,5)
MAX(...)        Maximum value           =MAX(A1:C3)

=== LOGICAL FUNCTIONS ===
IF(cond,true,false) Conditional         =IF(A1>5,100,0)
AND(...)        All values true         =AND(A1>0,B1<10)
OR(...)         Any value true          =OR(A1=0,B1=0)
NOT(value)      Logical not             =NOT(A1>5)

Note: 0 is false, anything else is true

=== CELL RANGES ===
A1:C3           Rectangle from A1 to C3
A1:A10          Column A, rows 1-10
B2:D2           Row 2, columns B-D

=== EXAMPLES ===
=A1+B1*2        Math with precedence
=IF(A1>0,A1*2,0) Conditional calculation
=SUM(A1:A5)/5   Same as AVERAGE(A1:A5)
=MAX(A1:C3)     Largest in 3x3 range
=A1**2+B1**2    Pythagorean calculation

=== FILE OPERATIONS ===
Ctrl+S          Save spreadsheet to file
Ctrl+O          Load spreadsheet from file
Ctrl+E          Export spreadsheet to CSV file
Ctrl+L          Import data from CSV file
                Files are saved as "spreadsheet.tshts" in JSON format
                CSV exports contain only cell values (not formulas)
                CSV imports replace current spreadsheet data

=== NAVIGATION SHORTCUTS ===
F1 or ?         Show this help (scroll with ↑↓, PgUp/PgDn, Home)
Enter/F2        Edit selected cell
Arrow keys      Navigate cells (hjkl also work)
= key           Auto-resize column to fit content
+ key           Auto-resize all columns to fit content
- / _ keys      Manually shrink/grow column width
q               Quit application

=== HELP NAVIGATION ===
↑↓ or j/k       Scroll help text up/down one line
Page Up/Down    Scroll help text up/down 5 lines
Home            Jump to top of help text
Esc/F1/?/q      Close this help window

Note: Your spreadsheet is automatically saved when you use Ctrl+S.
Use Ctrl+O to load the saved spreadsheet on next session."#.to_string()
}