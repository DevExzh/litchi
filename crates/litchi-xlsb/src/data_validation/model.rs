//! Semantic values used by the XLSB data-validation records.

use crate::formula::ParsedFormula;

/// Binary formula storage used by a validation rule.
///
/// The owner stores BIFF12 formula streams without interpreting or discarding
/// their ancillary bytes. Host adapters can implement this trait for their
/// own formula wrapper without coupling the owner to package orchestration.
pub trait FormulaBinary: Clone + PartialEq + Eq {
    /// Construct a formula from its losslessly preserved BIFF12 streams.
    fn from_parts(rgce: Vec<u8>, rgcb: Vec<u8>) -> Self;
    /// Borrow the formula's token stream.
    fn rgce(&self) -> &[u8];
    /// Borrow the formula's ancillary stream.
    fn rgcb(&self) -> &[u8];
}

impl FormulaBinary for ParsedFormula {
    fn from_parts(rgce: Vec<u8>, rgcb: Vec<u8>) -> Self {
        Self { rgce, rgcb }
    }

    fn rgce(&self) -> &[u8] {
        &self.rgce
    }

    fn rgcb(&self) -> &[u8] {
        &self.rgcb
    }
}

/// Worksheet-level UI settings stored by a data-validation collection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Settings {
    /// Whether every validation input prompt is disabled on the sheet.
    pub input_prompts_disabled: bool,
    /// Horizontal prompt-window position in pixels.
    pub prompt_x: u16,
    /// Vertical prompt-window position in pixels.
    pub prompt_y: u16,
}

/// Binary record family used to store a validation rule.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum RecordKind {
    /// `BrtDVal`, limited to 8,191 target ranges.
    #[default]
    Classic,
    /// Office 2013 `BrtDVal14`, which permits more target ranges.
    Extension14,
}

/// A typed data-validation rule with lossless binary formula storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Validation<F = ParsedFormula>
where
    F: FormulaBinary,
{
    /// Type of validation (0=none, 1=whole, 2=decimal, 3=list, 4=date, 5=time, 6=text length, 7=custom).
    pub validation_type: u8,
    /// Operator (0=between, 1=not between, 2=equal, 3=not equal, 4=greater than, 5=less than, 6=greater or equal, 7=less or equal).
    pub operator: u8,
    /// First formula (constraint).
    pub formula1: Option<String>,
    /// Second formula (for between/not between).
    pub formula2: Option<String>,
    /// Original first binary formula, retained for lossless round-tripping.
    pub formula1_binary: Option<F>,
    /// Original second binary formula, retained for lossless round-tripping.
    pub formula2_binary: Option<F>,
    /// Allow blank cells.
    pub allow_blank: bool,
    /// Show dropdown (for list validation).
    pub show_dropdown: bool,
    /// Show input message.
    pub show_input_message: bool,
    /// Show error message.
    pub show_error_message: bool,
    /// Error style (0=stop, 1=warning, 2=information).
    pub error_style: u8,
    /// Input Method Editor mode (0=no control, 1..=10 are XLSB IME modes).
    pub ime_mode: u8,
    /// Compatibility bit used by Excel and LibreOffice for inline list strings.
    pub string_list: bool,
    /// Optional `BrtDValList` text which overrides the first binary formula.
    pub list_formula: Option<String>,
    /// Input message title.
    pub input_title: Option<String>,
    /// Input message text.
    pub input_text: Option<String>,
    /// Error message title.
    pub error_title: Option<String>,
    /// Error message text.
    pub error_text: Option<String>,
    /// Cell ranges (for example, `A1:B2 C3:D4`) in source order.
    pub cell_ranges: String,
    /// Binary record family from which this rule was read or should be written.
    pub record_kind: RecordKind,
}

impl<F: FormulaBinary> Validation<F> {
    /// Create a new validation rule with the XLSB-compatible defaults.
    #[must_use]
    pub fn new(validation_type: u8, cell_ranges: String) -> Self {
        Self {
            validation_type,
            operator: 0,
            formula1: None,
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: validation_type == 3,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges,
            record_kind: RecordKind::Classic,
        }
    }
}
