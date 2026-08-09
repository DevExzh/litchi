//! Contextual semantic values produced by BIFF8 DVAL and DV records.

pub(crate) const DVAL_RECORD_TYPE: u16 = 0x01B2;
pub(crate) const DV_RECORD_TYPE: u16 = 0x01BE;

/// Worksheet-level settings declared by a BIFF8 DVAL record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings {
    pub(super) window_closed: bool,
    pub(super) x_left: u32,
    pub(super) y_top: u32,
    pub(super) dropdown_object_id: Option<u16>,
    pub(super) declared_rule_count: u16,
}

impl Settings {
    #[must_use]
    pub fn window_closed(&self) -> bool {
        self.window_closed
    }
    #[must_use]
    pub fn x_left(&self) -> u32 {
        self.x_left
    }
    #[must_use]
    pub fn y_top(&self) -> u32 {
        self.y_top
    }
    #[must_use]
    pub fn dropdown_object_id(&self) -> Option<u16> {
        self.dropdown_object_id
    }
    #[must_use]
    pub fn declared_rule_count(&self) -> u16 {
        self.declared_rule_count
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Any,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorStyle {
    Stop,
    Warning,
    Information,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterOrEqual,
    LessOrEqual,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImeMode {
    NoControl,
    On,
    Off,
    Hiragana,
    WideKatakana,
    NarrowKatakana,
    FullWidthAlphanumeric,
    HalfWidthAlphanumeric,
    FullWidthHangul,
    HalfWidthHangul,
}

/// An unevaluated BIFF formula token stream from a DV record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    pub(super) tokens: Vec<u8>,
    pub(super) rendered: Option<String>,
}

impl Formula {
    #[must_use]
    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Best-effort inert rendering using the workbook's existing BIFF token renderer.
    #[must_use]
    pub fn rendered(&self) -> Option<&str> {
        self.rendered.as_deref()
    }
}

/// An inclusive BIFF8 cell range targeted by a validation rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u8,
    pub(super) last_column: u8,
}

impl Range {
    #[must_use]
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    #[must_use]
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    #[must_use]
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    #[must_use]
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

/// One BIFF8 worksheet data-validation rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rule {
    pub(super) kind: Kind,
    pub(super) error_style: ErrorStyle,
    pub(super) explicit_list: bool,
    pub(super) allow_blank: bool,
    pub(super) suppress_dropdown: bool,
    pub(super) ime_mode: ImeMode,
    pub(super) show_input_message: bool,
    pub(super) show_error_message: bool,
    pub(super) operator: Operator,
    pub(super) prompt_title: Option<String>,
    pub(super) error_title: Option<String>,
    pub(super) prompt: Option<String>,
    pub(super) error: Option<String>,
    pub(super) formula1: Option<Formula>,
    pub(super) formula2: Option<Formula>,
    pub(super) ranges: Vec<Range>,
}

impl Rule {
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }
    #[must_use]
    pub fn error_style(&self) -> ErrorStyle {
        self.error_style
    }
    #[must_use]
    pub fn explicit_list(&self) -> bool {
        self.explicit_list
    }
    #[must_use]
    pub fn allow_blank(&self) -> bool {
        self.allow_blank
    }
    #[must_use]
    pub fn suppress_dropdown(&self) -> bool {
        self.suppress_dropdown
    }
    #[must_use]
    pub fn ime_mode(&self) -> ImeMode {
        self.ime_mode
    }
    #[must_use]
    pub fn show_input_message(&self) -> bool {
        self.show_input_message
    }
    #[must_use]
    pub fn show_error_message(&self) -> bool {
        self.show_error_message
    }
    #[must_use]
    pub fn operator(&self) -> Operator {
        self.operator
    }
    #[must_use]
    pub fn prompt_title(&self) -> Option<&str> {
        self.prompt_title.as_deref()
    }
    #[must_use]
    pub fn error_title(&self) -> Option<&str> {
        self.error_title.as_deref()
    }
    #[must_use]
    pub fn prompt(&self) -> Option<&str> {
        self.prompt.as_deref()
    }
    #[must_use]
    pub fn error(&self) -> Option<&str> {
        self.error.as_deref()
    }
    #[must_use]
    pub fn formula1(&self) -> Option<&Formula> {
        self.formula1.as_ref()
    }
    #[must_use]
    pub fn formula2(&self) -> Option<&Formula> {
        self.formula2.as_ref()
    }
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
}
