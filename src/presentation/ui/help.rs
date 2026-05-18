//! Submodule of `ui` — see ui/mod.rs.

#![allow(unused_imports)]
use crate::application::{App, AppMode};
use crate::domain::{CellStyle, NumberFormat, Spreadsheet, TerminalColor, format_cell_value};
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table, Tabs},
    Frame,
};
use super::{caret, layer_for_render, terminal_color_to_ratatui};

pub(super) fn render_help_popup(f: &mut Frame, app: &App) {
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
    let scroll = app.help_scroll;
    let start_line = scroll.min(help_lines.len().saturating_sub(visible_height));
    let end_line = (start_line + visible_height).min(help_lines.len());
    let visible_text = help_lines[start_line..end_line].join("\n");

    let title = if app.help_search_active {
        format!(
            "Help (Line {}/{}) — /{}",
            start_line + 1,
            help_lines.len(),
            app.help_search
        )
    } else if !app.help_search.is_empty() {
        format!(
            "Help (Line {}/{}) — /{} (press n for next)",
            start_line + 1,
            help_lines.len(),
            app.help_search
        )
    } else {
        format!(
            "Help (Line {}/{}) — 1-9: jump, /: search, ↑↓: scroll, Esc: close",
            start_line + 1,
            help_lines.len()
        )
    };

    let help_widget = Paragraph::new(visible_text)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(title)
                .style(Style::default().fg(Color::Cyan)),
        )
        .style(Style::default().fg(Color::White));
    f.render_widget(help_widget, popup_area);
}

pub fn help_section_offset(key: char) -> Option<usize> {
    let text = get_help_text();
    let target = HELP_SECTIONS.iter().find(|(k, _)| *k == key)?.1;
    text.lines().position(|l| l.trim() == target)
}

pub fn find_help_match(needle: &str, from: usize) -> Option<usize> {
    if needle.is_empty() {
        return None;
    }
    let text = get_help_text();
    let needle_lc = needle.to_lowercase();
    let lines: Vec<&str> = text.lines().collect();
    let len = lines.len();
    // Search from `from` forward, then wrap.
    for i in 0..len {
        let idx = (from + i) % len;
        if lines[idx].to_lowercase().contains(&needle_lc) {
            return Some(idx);
        }
    }
    None
}

pub const HELP_SECTIONS: &[(char, &str)] = &[
    ('1', "=== BASIC & OPERATORS ==="),
    ('2', "=== NUMERIC FUNCTIONS ==="),
    ('3', "=== STRING FUNCTIONS ==="),
    ('4', "=== LOGICAL FUNCTIONS ==="),
    ('5', "=== LOOKUP / CONDITIONAL ==="),
    ('6', "=== DATE FUNCTIONS ==="),
    ('7', "=== FILE OPERATIONS ==="),
    ('8', "=== NAVIGATION SHORTCUTS ==="),
    ('9', "=== ADVANCED ==="),
    ('0', "=== VIM MODE ==="),
];

pub(crate) fn get_help_text() -> String {
    r#"TSHTS EXPRESSION LANGUAGE REFERENCE (v0.2)

Press 0-9 to jump to a section (0 = Vim Mode). Press / to search within this help.

=== WHAT'S NEW IN 0.2 ===
• Cross-sheet refs: Sheet2!A1, 'My Sheet'!B5:B10, 3-D Sheet1:Sheet3!A1
• Excel error types: #DIV/0!, #REF!, #VALUE!, #NAME?, #NUM!, #N/A, #NULL!, #SPILL!
• 2-D arrays with shape: =A1:C3 * 2 broadcasts; VLOOKUP/HLOOKUP/XLOOKUP/INDEX
• Array literals: {1,2,3;4,5,6}
• Dynamic arrays: SEQUENCE, FILTER, SORT, UNIQUE, TRANSPOSE, SUMPRODUCT
• LET local bindings, LAMBDA functions, MAP/REDUCE/BYROW/BYCOL/SCAN/MAKEARRAY
• Named LAMBDAs: `:name DOUBLE LAMBDA(x, x*2)` then =DOUBLE(7) → 14
• Tables: `:table create A1:D100 name=Sales` + structured refs Table1[Col]
• Pivots: `:pivot SOURCE TARGET row=COL value=COL agg=sum|count|avg|min|max`
  (auto-refreshes via formulas)
• Charts: `:chart bar A1:A10`, `:chart line`, popup with auto-scale
• Goal Seek: `:goalseek TARGET EXPECTED INPUT`
• Iterative calc: `:iterative on/off` for intentional circular refs
• Data validation: `:validate <COL> "_ > 0"` flags violators in red
• Frozen rows AND columns; hyperlinks (http://) open in browser on Enter
• Mouse: click-to-select, scroll wheel
• .xlsx import + export (auto-detects extension)

=== BASIC & OPERATORS ===
• All formulas start with = (equals sign)
• Cell references use column letter + row number (A1, B2, Z99, AA1, etc.)
• Numbers can be integers or decimals (42, 3.14, -5.5)
• Strings use double quotes ("Hello World", "", "Quote""Test")
• Case insensitive for functions and cell references
• Supports both numbers and strings with automatic type conversion

=== ARITHMETIC OPERATORS ===
+       Addition                    =5+3 → 8, =A1+B1
-       Subtraction                 =10-3 → 7, =A1-5
*       Multiplication              =4*3 → 12, =A1*B1
/       Division                    =15/3 → 5, =A1/B1
**      Exponentiation              =2**3 → 8, =A1**2
^       Power (same as **)          =3^2 → 9, =A1^B1
%       Modulo (remainder)          =10%3 → 1, =A1%B1

=== STRING OPERATORS ===
&       Concatenation               ="Hello" & " " & "World" → Hello World
                                   ="Number: " & 42 → Number: 42
""      String literals             ="Hello World", =""
                                   Use "" for quotes: "Quote""Test" → Quote"Test

=== COMPARISON OPERATORS ===
<       Less than                   =A1<B1 → 1 or 0 (works with numbers only)
>       Greater than                =A1>B1 → 1 or 0 (works with numbers only)
<=      Less than or equal          =A1<=B1 → 1 or 0 (works with numbers only)
>=      Greater than or equal       =A1>=B1 → 1 or 0 (works with numbers only)
=       Equal                       =A1=B1 → 1 or 0 (works with strings and numbers)
<>      Not equal                   =A1<>B1 → 1 or 0 (works with strings and numbers)

Note: Comparisons return 1 for true, 0 for false

=== NUMERIC FUNCTIONS ===
SUM(...)        Sum of values           =SUM(A1,B1,C1) or =SUM(A1:C1)
AVERAGE(...)    Average of values       =AVERAGE(A1:A10)
MIN(...)        Minimum value           =MIN(A1,B1,C1,5)
MAX(...)        Maximum value           =MAX(A1:C3)
ABS(value)      Absolute value          =ABS(-5) → 5
SQRT(value)     Square root             =SQRT(16) → 4
ROUND(num)      Round to integer        =ROUND(3.14) → 3
ROUND(num,places) Round to decimals     =ROUND(3.14159,2) → 3.14

=== STRING FUNCTIONS ===
LEN(text)       String length           =LEN("Hello") → 5
UPPER(text)     Convert to uppercase    =UPPER("hello") → HELLO
LOWER(text)     Convert to lowercase    =LOWER("WORLD") → world
TRIM(text)      Remove leading/trailing spaces  =TRIM("  hi  ") → hi
LEFT(text,num)  First N characters      =LEFT("Hello World",5) → Hello
RIGHT(text,num) Last N characters       =RIGHT("Hello World",5) → World
MID(text,start,len) Substring           =MID("Hello World",6,5) → World
FIND(search,text) Find position         =FIND("lo","Hello") → 3
FIND(search,text,start) Find from pos   =FIND("l","Hello",2) → 3
CONCAT(...)     Concatenate values      =CONCAT("A","B","C") → ABC

=== WEB FUNCTIONS ===
GET(url)        Fetch content from URL  =GET("https://api.example.com/data")
                                       =GET("https://raw.githubusercontent.com/...")

Note: String functions use 0-based indexing (positions start at 0)

=== LOGICAL FUNCTIONS ===
IF(cond,true,false) Conditional         =IF(A1>5,"High","Low")
                                       =IF(A1="Hello","Found","Not Found")
AND(...)        All values true         =AND(A1>0,B1<10)
OR(...)         Any value true          =OR(A1="",A1="N/A")
NOT(value)      Logical not             =NOT(A1>5)
TRUE() / FALSE() Boolean literals
ISBLANK, ISNUMBER, ISTEXT, TYPE          Type tests
COUNT(...), COUNTA(...)                  Counters

Note: For logical tests: 0 and empty strings are false, everything else is true

=== LOOKUP / CONDITIONAL ===
SUMIF(range, criteria, [sum_range])  Sum values matching criteria
COUNTIF(range, criteria)             Count matches; supports ">5", "<=10", "*glob*"
AVERAGEIF(range, criteria, [avg])    Mean of matches
VLOOKUP(value, range, col_index, [exact])
INDEX(range, row, [col])             1-based; col defaults to 1
MATCH(value, range, [type])          type 0 = exact, 1 = approx (default)
INDIRECT(ref_text)                   Build cell ref at eval time: =INDIRECT("A"&ROW())
OFFSET(base, rows, cols, [h], [w])   Range/value offset from base

=== DATE FUNCTIONS ===
TODAY()                              Days since 1899-12-30 (Excel serial)
NOW()                                Days + fractional time-of-day
DATE(year, month, day)               Construct serial
YEAR(serial), MONTH(serial), DAY(serial)
Tip: format a date column with `:format number 0` to see the serial integer.

=== ADVANCED ===
Named ranges: `:name MyRange A1:B10`, then use `MyRange` in formulas.
              `:names` lists; `:unname X` removes.
Conditional formatting: `:cf A "_ > 100" bg=red bold` — `_` binds to cell value.
                       `:cf list`, `:cf clear` to manage rules.
Search options: `:regex on/off`, `:case on/off`.
Auto-save:   `:autosave on/off` — 30s idle window, requires a known filename.
Cache:       `:cache clear` clears GET() cache; F5 recomputes all formulas.
CSV append:  `:import-append PATH` adds rows below current data.
Recent files: tab in load/save dialog cycles through recent + cwd matches.

=== CELL RANGES ===
A1:C3           Rectangle from A1 to C3
A1:A10          Column A, rows 1-10
B2:D2           Row 2, columns B-D

=== TYPE CONVERSION ===
• Numbers in strings are automatically converted: "123" + 1 → 124
• Numbers in string operations: 42 & " items" → "42 items"
• Invalid strings become 0 in math: "hello" + 1 → 1
• String comparisons are case-sensitive: "Hello" <> "hello" → 1

=== FORMULA EXAMPLES ===

Numeric Examples:
=A1+B1*2        Math with precedence
=IF(A1>0,A1*2,0) Conditional calculation
=SUM(A1:A5)/5   Same as AVERAGE(A1:A5)
=MAX(A1:C3)     Largest in 3x3 range
=A1**2+B1**2    Pythagorean calculation

String Examples:
=UPPER(A1) & " - " & LOWER(B1)     Combined formatted text
=IF(LEN(A1)>0,A1,"Empty")          Check for non-empty strings
=LEFT(A1,FIND(" ",A1)-1)           Extract first word
="Hello " & A1 & ", you scored " & B1 & "%"   Dynamic messages
=IF(AND(LEN(A1)>3,A1<>""),"Valid","Invalid")  Validate input

Mixed Type Examples:
="Total: " & SUM(A1:A10) & " items"   Numeric result with description
=IF(AVERAGE(A1:A10)>50,"PASS","FAIL") Grade based on average
=CONCAT("Value: ",A1," Total: ",SUM(B1:B5))  Dynamic labels

=== SEARCH FUNCTIONALITY ===
/               Start text search across all cells
                Search is case-insensitive and searches both cell values and formulas
                Live search: results update as you type
                ↑↓          Navigate through search results while searching
                Enter       Finish search and return to normal mode
                Esc         Cancel search and return to normal mode
                n/N         Navigate search results in normal mode (after search)

=== FILE OPERATIONS ===
Ctrl+S          Save in place (or Save As if file is new)
Ctrl+O          Load spreadsheet from file (prompts to confirm if dirty)
Ctrl+E          Export spreadsheet to CSV file
Ctrl+L          Import data from CSV file (prompts to confirm if dirty)
                Files are saved as "spreadsheet.tshts" in JSON format
                CSV exports contain only cell values (not formulas)
                CSV imports replace current spreadsheet data

Open at startup: pass a filename on the command line: `tshts foo.tshts`.

=== NAVIGATION SHORTCUTS ===
TSHTS uses vim-style modes. In NORMAL mode, single letters trigger commands;
to type into a cell, enter INSERT mode first (see VIM MODE section below).
The classic Excel/Sheets Ctrl-bindings still work alongside the vim layer.

F1 or ?         Show this help (scroll with ↑↓, PgUp/PgDn, Home)
Enter / F2 / i  Start editing the current cell (INSERT mode)
Arrow keys      Navigate cells (h / j / k / l also work)
Shift+arrows    Quick range-select (or use `v` to enter VISUAL mode)
+ key           Auto-resize all columns to fit content
- / _ keys      Manually shrink/grow column width
Ctrl+Z / u      Undo last action
Ctrl+Y / Ctrl+R Redo last undone action
Ctrl+G          Go to cell reference
Ctrl+C / X      Copy / Cut (uses system clipboard for plain text)
Ctrl+V          Enter VISUAL BLOCK mode (paste is `p` or `P`)
Ctrl+D          Autofill the selection from its top-left cell
Ctrl+B / U      Toggle bold / underline on selection
Ctrl+Home/End   Jump to A1 / last cell with data
Ctrl+PgUp/PgDn  Switch to previous / next sheet
/               Start text search (live)
n / N           Next / previous search result
:               Ex-command palette (`:q` quit, `:w` save, `:wq` save+quit, ...)
F5              Recalculate all formulas (refresh RAND / GET caches)
Esc             Clear selection / search highlights / pending op / status
q  or  :q       Quit (prompts if there are unsaved changes)
:q!             Force quit, discarding changes

=== VIM MODE ===
TSHTS speaks vim. There are four primary modes — the current one is shown
as `-- NORMAL --`, `-- INSERT --`, `-- VISUAL --`, or `-- COMMAND --` at
the left of the status bar. Pending operators and counts also show there
(e.g. `[5d]` while you're typing `5dd`).

--- Mode transitions ---
i / a           Enter INSERT mode (edit current cell)
I               INSERT with cursor at start of the cell's text
A               INSERT with cursor at end of the cell's text
o               Open a new row below, enter INSERT mode
O               Move to the row above, enter INSERT mode
s               Substitute cell: clear it and enter INSERT
S               Substitute row: clear current row, enter INSERT at col 0
Enter / F2      Enter INSERT mode without clearing
Esc             Return to NORMAL from any mode (also cancels pending op)
v               Enter VISUAL (cell-granularity selection)
V               Enter VISUAL LINE (whole-row selection)
Ctrl+V          Enter VISUAL BLOCK (rectangular selection)
:               Enter COMMAND (ex-command palette)

--- Motions (work in NORMAL and VISUAL) ---
h j k l         Left / down / up / right
Arrow keys      Same as h/j/k/l
0 / Home        First column of the current row
$ / End         Last column with data in the current row
^               First column with data in the current row
gg              Jump to first row
G               Jump to last row with data
NG              Jump to row N (e.g. `42G`)
PgUp / PgDn     Page up / down
Tab / Shift+Tab Move right / left (Excel-style)

--- Operators (NORMAL mode) ---
Operators set a pending state shown in the status bar; press a motion
next to apply, or repeat the operator key for a whole-row operation.

d{motion}       Delete the range from cursor to motion target
y{motion}       Yank (copy) the range
c{motion}       Change: delete and enter INSERT at the start
dd / yy / cc    Whole-row delete / yank / change
x               Delete current cell (no motion needed)
p               Paste after / below cursor
P               Paste before / above cursor
u               Undo
Ctrl+R          Redo

--- Counts ---
Prefix any motion or operator with a count.
  5j            Move down 5 rows
  3l            Move right 3 columns
  10G           Jump to row 10
  3dd           Delete 3 rows starting here
  2yj           Yank current row + 2 rows below it

--- VISUAL mode ---
Once in visual mode, motions extend the selection. Status bar shows the
size and SUM / AVG / COUNT of the current selection. Then:
  y             Yank the selection
  d  or  x      Delete the selection
  c             Change: delete and enter INSERT at the top-left
  p             Paste over the selection
  v / Esc       Exit VISUAL back to NORMAL

--- COMMAND mode (`:`) ---
:q              Quit (prompts on unsaved changes)
:q!             Force quit, discard changes
:w              Save (in place if known, else prompts)
:w <filename>   Save as <filename> (.xlsx auto-detected)
:wq  or  :x     Save and quit
:wq!  or  :x!   Save and force quit
:e <filename>   Open another file
Every command in this help (`:sort asc`, `:freeze`, `:format number`,
`:cf …`, `:name …`, etc.) is also reachable here. Type and press Tab
to cycle through suggestions.

=== ABSOLUTE REFERENCES ===
$A$1            Both row and column absolute
$A1             Column A is absolute; row shifts with autofill/paste
A$1             Column A shifts; row 1 is absolute
Use `$` to anchor part of a reference during autofill or paste.

=== SELECTION AND AUTOFILL ===
Shift+arrows    Select a range of cells by holding Shift and using arrow keys
Ctrl+D          Autofill: Copy formula from top-left cell of selection to all
                selected cells, automatically adjusting cell references
                Example: =SUM(B4:B6) becomes =SUM(C4:C6), =SUM(D4:D6), etc.
                when dragged right, or =SUM(B5:B7), =SUM(B6:B8), etc. when
                dragged down. Works with any formula containing cell references.

=== HELP NAVIGATION ===
↑↓ or j/k       Scroll help text up/down one line
Page Up/Down    Scroll help text up/down 5 lines
Home            Jump to top of help text
Esc/F1/?/q      Close this help window

=== ERROR HANDLING ===
#ERROR          Displayed when formula evaluation fails
                Common causes: division by zero, invalid functions,
                circular references, invalid FIND operations

Note: Your spreadsheet is automatically saved when you use Ctrl+S.
Use Ctrl+O to load the saved spreadsheet on next session."#.to_string()
}
