//! Shared formatting types for XLSX (used in both reading and writing).

/// Cell format information.
#[derive(Debug, Clone, Default)]
pub struct CellFormat {
    pub font: Option<CellFont>,
    pub fill: Option<CellFill>,
    pub border: Option<CellBorder>,
    pub number_format: Option<String>,
}

/// Font properties for a cell.
#[derive(Debug, Clone)]
pub struct CellFont {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: Option<String>,
}

#[allow(clippy::derivable_impls)]
impl Default for CellFont {
    fn default() -> Self {
        Self {
            name: None,
            size: None,
            bold: false,
            italic: false,
            underline: false,
            color: None,
        }
    }
}

/// Fill properties for a cell.
#[derive(Debug, Clone)]
pub struct CellFill {
    pub pattern_type: CellFillPatternType,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

/// Cell fill pattern types.
#[derive(Debug, Clone, Copy)]
pub enum CellFillPatternType {
    None,
    Solid,
    Gray125,
    DarkGray,
    MediumGray,
    LightGray,
    Gray0625,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
}

impl CellFillPatternType {
    #[allow(dead_code)]
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::Gray125 => "gray125",
            Self::DarkGray => "darkGray",
            Self::MediumGray => "mediumGray",
            Self::LightGray => "lightGray",
            Self::Gray0625 => "gray0625",
            Self::DarkHorizontal => "darkHorizontal",
            Self::DarkVertical => "darkVertical",
            Self::DarkDown => "darkDown",
            Self::DarkUp => "darkUp",
            Self::DarkGrid => "darkGrid",
            Self::DarkTrellis => "darkTrellis",
        }
    }
}

/// Border properties for a cell.
#[derive(Debug, Clone, Default)]
pub struct CellBorder {
    pub left: Option<CellBorderSide>,
    pub right: Option<CellBorderSide>,
    pub top: Option<CellBorderSide>,
    pub bottom: Option<CellBorderSide>,
    pub diagonal: Option<CellBorderSide>,
}

/// Border side properties.
#[derive(Debug, Clone)]
pub struct CellBorderSide {
    pub style: CellBorderLineStyle,
    pub color: Option<String>,
}

/// Border line styles.
#[derive(Debug, Clone, Copy)]
pub enum CellBorderLineStyle {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl CellBorderLineStyle {
    #[allow(dead_code)]
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Thin => "thin",
            Self::Medium => "medium",
            Self::Dashed => "dashed",
            Self::Dotted => "dotted",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Hair => "hair",
            Self::MediumDashed => "mediumDashed",
            Self::DashDot => "dashDot",
            Self::MediumDashDot => "mediumDashDot",
            Self::DashDotDot => "dashDotDot",
            Self::MediumDashDotDot => "mediumDashDotDot",
            Self::SlantDashDot => "slantDashDot",
        }
    }
}

/// Chart types supported in Excel.
#[derive(Debug, Clone, Copy)]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
}

/// Chart configuration.
#[derive(Debug, Clone)]
pub struct Chart {
    pub chart_type: ChartType,
    pub title: Option<String>,
    pub data_range: String,
    pub position: (u32, u32, u32, u32),
    pub show_legend: bool,
}

/// Data validation types.
#[derive(Debug, Clone)]
pub enum DataValidationType {
    Whole {
        operator: DataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    Decimal {
        operator: DataValidationOperator,
        value1: f64,
        value2: Option<f64>,
    },
    List {
        values: Vec<String>,
    },
    Date {
        operator: DataValidationOperator,
        value1: String,
        value2: Option<String>,
    },
    TextLength {
        operator: DataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    Custom {
        formula: String,
    },
}

/// Data validation operators.
#[derive(Debug, Clone, Copy)]
pub enum DataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

impl DataValidationOperator {
    #[allow(dead_code)]
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThan => "greaterThan",
            Self::LessThan => "lessThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::LessThanOrEqual => "lessThanOrEqual",
        }
    }
}

/// Data validation rule.
#[derive(Debug, Clone)]
pub struct DataValidation {
    pub range: String,
    pub validation_type: DataValidationType,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
}
