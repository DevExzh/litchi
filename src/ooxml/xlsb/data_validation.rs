//! Data validation support for XLSB

/// Data validation rule
///
/// Represents data validation constraints on a cell or range.
#[derive(Debug, Clone)]
pub struct DataValidation {
    /// Type of validation (0=none, 1=whole, 2=decimal, 3=list, 4=date, 5=time, 6=text length, 7=custom)
    pub validation_type: u8,
    /// Operator (0=between, 1=not between, 2=equal, 3=not equal, 4=greater than, 5=less than, 6=greater or equal, 7=less or equal)
    pub operator: u8,
    /// First formula (constraint)
    pub formula1: Option<String>,
    /// Second formula (for between/not between)
    pub formula2: Option<String>,
    /// Allow blank cells
    pub allow_blank: bool,
    /// Show dropdown (for list validation)
    pub show_dropdown: bool,
    /// Show input message
    pub show_input_message: bool,
    /// Show error message
    pub show_error_message: bool,
    /// Error style (0=stop, 1=warning, 2=information)
    pub error_style: u8,
    /// Input message title
    pub input_title: Option<String>,
    /// Input message text
    pub input_text: Option<String>,
    /// Error message title
    pub error_title: Option<String>,
    /// Error message text
    pub error_text: Option<String>,
    /// Cell ranges (e.g., "A1:B2,C3:D4")
    pub cell_ranges: String,
}

impl DataValidation {
    /// Create a new data validation rule
    pub fn new(validation_type: u8, cell_ranges: String) -> Self {
        DataValidation {
            validation_type,
            operator: 0,
            formula1: None,
            formula2: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges,
        }
    }
}
