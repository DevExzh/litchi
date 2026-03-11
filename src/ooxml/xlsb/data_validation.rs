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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_data_validation_new() {
        let dv = DataValidation::new(3, "A1:A10".to_string());
        assert_eq!(dv.validation_type, 3);
        assert_eq!(dv.cell_ranges, "A1:A10");
        // Check defaults
        assert_eq!(dv.operator, 0);
        assert!(dv.formula1.is_none());
        assert!(dv.formula2.is_none());
        assert!(dv.allow_blank);
        assert!(dv.show_dropdown);
        assert!(!dv.show_input_message);
        assert!(dv.show_error_message);
        assert_eq!(dv.error_style, 0);
        assert!(dv.input_title.is_none());
        assert!(dv.input_text.is_none());
        assert!(dv.error_title.is_none());
        assert!(dv.error_text.is_none());
    }

    #[test]
    fn test_data_validation_whole_number() {
        let mut dv = DataValidation::new(1, "B1:B20".to_string()); // whole number
        dv.operator = 2; // greater than
        dv.formula1 = Some("10".to_string());
        dv.allow_blank = false;

        assert_eq!(dv.validation_type, 1);
        assert_eq!(dv.operator, 2);
        assert_eq!(dv.formula1, Some("10".to_string()));
        assert!(!dv.allow_blank);
    }

    #[test]
    fn test_data_validation_decimal() {
        let mut dv = DataValidation::new(2, "C1:C10".to_string()); // decimal
        dv.operator = 0; // between
        dv.formula1 = Some("0".to_string());
        dv.formula2 = Some("100".to_string());

        assert_eq!(dv.validation_type, 2);
        assert_eq!(dv.operator, 0);
        assert_eq!(dv.formula1, Some("0".to_string()));
        assert_eq!(dv.formula2, Some("100".to_string()));
    }

    #[test]
    fn test_data_validation_list() {
        let mut dv = DataValidation::new(3, "D1:D10".to_string()); // list
        dv.formula1 = Some("Yes,No,Maybe".to_string());
        dv.show_dropdown = true;

        assert_eq!(dv.validation_type, 3);
        assert_eq!(dv.formula1, Some("Yes,No,Maybe".to_string()));
        assert!(dv.show_dropdown);
    }

    #[test]
    fn test_data_validation_date() {
        let mut dv = DataValidation::new(4, "E1:E10".to_string()); // date
        dv.operator = 4; // greater than
        dv.formula1 = Some("2024-01-01".to_string());

        assert_eq!(dv.validation_type, 4);
        assert_eq!(dv.operator, 4);
    }

    #[test]
    fn test_data_validation_time() {
        let mut dv = DataValidation::new(5, "F1:F10".to_string()); // time
        dv.operator = 5; // less than
        dv.formula1 = Some("12:00".to_string());

        assert_eq!(dv.validation_type, 5);
        assert_eq!(dv.operator, 5);
    }

    #[test]
    fn test_data_validation_text_length() {
        let mut dv = DataValidation::new(6, "G1:G10".to_string()); // text length
        dv.operator = 6; // greater than or equal
        dv.formula1 = Some("5".to_string());

        assert_eq!(dv.validation_type, 6);
        assert_eq!(dv.formula1, Some("5".to_string()));
    }

    #[test]
    fn test_data_validation_custom() {
        let mut dv = DataValidation::new(7, "H1:H10".to_string()); // custom
        dv.formula1 = Some("=A1>0".to_string());

        assert_eq!(dv.validation_type, 7);
        assert_eq!(dv.formula1, Some("=A1>0".to_string()));
    }

    #[test]
    fn test_data_validation_with_messages() {
        let mut dv = DataValidation::new(1, "I1:I10".to_string());
        dv.show_input_message = true;
        dv.input_title = Some("Enter value".to_string());
        dv.input_text = Some("Please enter a number greater than 10".to_string());
        dv.show_error_message = true;
        dv.error_style = 0; // stop
        dv.error_title = Some("Invalid input".to_string());
        dv.error_text = Some("The value must be greater than 10".to_string());

        assert!(dv.show_input_message);
        assert_eq!(dv.input_title, Some("Enter value".to_string()));
        assert_eq!(
            dv.input_text,
            Some("Please enter a number greater than 10".to_string())
        );
        assert!(dv.show_error_message);
        assert_eq!(dv.error_style, 0);
        assert_eq!(dv.error_title, Some("Invalid input".to_string()));
        assert_eq!(
            dv.error_text,
            Some("The value must be greater than 10".to_string())
        );
    }

    #[test]
    fn test_data_validation_multiple_ranges() {
        let dv = DataValidation::new(3, "A1:A10,C1:C10,E1:E10".to_string());
        assert_eq!(dv.cell_ranges, "A1:A10,C1:C10,E1:E10");
    }

    #[test]
    fn test_data_validation_clone() {
        let dv = DataValidation::new(3, "A1:A10".to_string());
        let cloned = dv.clone();
        assert_eq!(cloned.validation_type, dv.validation_type);
        assert_eq!(cloned.cell_ranges, dv.cell_ranges);
    }
}
