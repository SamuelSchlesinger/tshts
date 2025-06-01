# TSHTS Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added - String Support and Multi-Type Formula System

#### 🎯 Major Features

**Multi-Type Value System**
- Added comprehensive string and number support throughout the formula engine
- Introduced `Value` enum with `Number(f64)` and `String(String)` variants
- Implemented automatic type conversion between strings and numbers
- Enhanced cell value detection to distinguish between numeric and string content

**String Literals and Operations**
- String literal support with double quotes: `"Hello World"`
- String concatenation operator `&`: `"Hello" & " " & "World"`
- Escaped quote support: `"Quote""Test"` produces `Quote"Test`
- Mixed-type concatenation: `"Number: " & 42` produces `"Number: 42"`

#### 🔧 String Functions

**Text Processing Functions**
- `LEN(text)` - Returns string length in characters
- `UPPER(text)` - Converts text to uppercase
- `LOWER(text)` - Converts text to lowercase  
- `TRIM(text)` - Removes leading and trailing whitespace

**String Extraction Functions (0-based indexing)**
- `LEFT(text, num_chars)` - Extracts first N characters
- `RIGHT(text, num_chars)` - Extracts last N characters
- `MID(text, start_pos, length)` - Extracts substring from specified position
- `FIND(search_text, within_text, [start_pos])` - Finds position of substring

**String Manipulation Functions**
- `CONCAT(value1, value2, ...)` - Concatenates multiple values
- Enhanced concatenation with automatic type conversion

#### 🌐 Web Integration Functions

**HTTP Request Functions**
- `GET(url)` - Fetches raw content from any HTTP/HTTPS URL
- Returns response body as string for further processing
- Enables integration with APIs, web services, and online data sources
- Equivalent to Google Sheets' IMPORTDATA function

#### 🧮 Enhanced Numeric Functions

**Updated Function Signatures**
- All existing functions now work with the new `Value` type system
- `SUM`, `AVERAGE`, `MIN`, `MAX` automatically convert string inputs to numbers
- `IF` function now supports string conditions and results
- `AND`, `OR`, `NOT` functions work with string truthiness (empty strings are false)

**New Math Functions**
- Enhanced `ROUND` function with optional decimal places: `ROUND(3.14159, 2)`
- All math functions handle mixed string/number inputs gracefully

#### 🔍 Enhanced Comparison Operations

**String and Number Comparisons**
- `=` and `<>` operators now work with both strings and numbers
- String equality is case-sensitive: `"Hello" = "hello"` returns `0`
- Mixed-type comparisons convert to strings for comparison
- Numeric comparisons (`<`, `>`, `<=`, `>=`) convert strings to numbers

#### 🎨 User Experience Improvements

**Enhanced Formula Examples**
- Updated in-app help with comprehensive string function documentation
- Added 30+ new examples covering string manipulation, data validation, and mixed operations
- Clear documentation of 0-based indexing for string functions
- Type conversion rules and behavior documentation

**Backward Compatibility**
- All existing numeric formulas continue to work unchanged
- Existing spreadsheets load and function normally
- No breaking changes to file format or existing functionality

#### 🧪 Testing and Quality

**Comprehensive Test Coverage**
- Added 14 new test functions specifically for string functionality
- All 135 tests passing (including existing functionality)
- Test coverage for string literals, concatenation, functions, cell references, and edge cases
- Performance testing to ensure no regression in numeric operations

**Error Handling**
- Robust error handling for invalid string operations
- Clear error messages for malformed string functions
- Graceful handling of type conversion edge cases

#### 📝 Documentation Updates

**README.md Enhancements**
- Expanded formula reference with string operations section
- Added 50+ new formula examples
- Comprehensive type conversion documentation
- Updated feature list to reflect new capabilities

**In-App Help System**
- Completely rewritten help text with string function documentation
- Added sections for string operators, functions, and type conversion
- Interactive examples for all new string functions
- Clear notation for 0-based indexing system

### Technical Implementation Details

#### Architecture Changes
- Refactored `ExpressionEvaluator` to work with `Value` type instead of `f64`
- Enhanced `FunctionRegistry` to support multi-type function signatures
- Updated parser to handle string literals and concatenation operator
- Extended AST with `Expr::String` variant and `BinaryOp::Concatenate`

#### Parser Enhancements
- Added string literal tokenization with proper quote escaping
- Implemented concatenation operator parsing with correct precedence
- Enhanced lexer to handle ampersand (`&`) operator
- Updated grammar to support string expressions

#### Performance Optimizations
- Efficient string handling with minimal allocations
- Optimized type conversion for common operations
- Maintained performance parity with existing numeric operations

### Usage Examples

#### Basic String Operations
```
="Hello World"                    → Hello World
="Hello" & " " & "World"          → Hello World
="Result: " & (2 + 3)             → Result: 5
=LEN("Hello")                     → 5
=UPPER("hello world")             → HELLO WORLD
```

#### Advanced String Processing
```
=LEFT("Hello World", 5)           → Hello
=RIGHT("Hello World", 5)          → World
=MID("Hello World", 6, 5)         → World
=FIND("World", "Hello World")     → 6
=TRIM("  spaces  ")               → spaces
```

#### Mixed-Type Operations
```
=IF(LEN(A1)>0, A1, "Empty")              → Conditional string handling
=UPPER(A1) & " - " & LOWER(B1)           → Format combination
="Total: " & SUM(A1:A10) & " items"      → Numeric results with labels
=IF(AVERAGE(A1:A10)>50, "PASS", "FAIL")  → Grade calculation with text
```

#### Web Data Integration
```
=GET("https://api.github.com/users/octocat")        → Fetch user data from GitHub API
=GET("https://jsonplaceholder.typicode.com/posts/1") → Get sample JSON data
=LEN(GET("https://example.com"))                    → Get content length of webpage
=UPPER(GET("https://httpbin.org/uuid"))             → Fetch and format UUID
=FIND("error", GET("https://api.service.com/status")) → Check API status
```

#### Data Validation and Processing
```
=IF(AND(LEN(A1)>3, A1<>""), "Valid", "Invalid")    → Input validation
=LEFT(A1, FIND(" ", A1)-1)                         → Extract first word
=IF(OR(A1="", A1="N/A"), "Missing", A1)            → Handle missing data
```

### Migration Guide

#### For Existing Users
- No changes required - all existing formulas continue to work
- Existing files load normally with full compatibility
- New string features are opt-in and don't affect existing workflows

#### For New Users
- Start with string literals using double quotes: `="Hello"`
- Use `&` operator for concatenation: `="Hello" & " World"`
- Explore string functions: `=LEN()`, `=UPPER()`, `=LOWER()`, etc.
- Remember that string functions use 0-based indexing

### Breaking Changes
- None. This release maintains full backward compatibility.

### Known Issues
- String comparisons are case-sensitive (by design)
- FIND function returns error for non-existent matches (by design)
- String indexing is 0-based (different from some spreadsheet applications)

---

## Previous Releases

### [0.1.0] - Initial Release
- Basic spreadsheet functionality with numeric formulas
- Terminal-based UI with keyboard navigation
- JSON file format for persistence
- Core arithmetic and logical functions
- Cell references and range operations