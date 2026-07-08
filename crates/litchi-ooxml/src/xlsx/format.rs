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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_format_default() {
        let format = CellFormat::default();
        assert!(format.font.is_none());
        assert!(format.fill.is_none());
        assert!(format.border.is_none());
        assert!(format.number_format.is_none());
    }

    #[test]
    fn test_cell_font_default() {
        let font = CellFont::default();
        assert!(font.name.is_none());
        assert!(font.size.is_none());
        assert!(!font.bold);
        assert!(!font.italic);
        assert!(!font.underline);
        assert!(font.color.is_none());
    }

    #[test]
    fn test_cell_font_clone() {
        let font = CellFont {
            name: Some("Arial".to_string()),
            size: Some(12.0),
            bold: true,
            italic: false,
            underline: true,
            color: Some("FF0000".to_string()),
        };
        let font2 = font.clone();
        assert_eq!(font.name, font2.name);
        assert_eq!(font.bold, font2.bold);
    }

    #[test]
    fn test_cell_fill_pattern_type_as_str() {
        assert_eq!(CellFillPatternType::None.as_str(), "none");
        assert_eq!(CellFillPatternType::Solid.as_str(), "solid");
        assert_eq!(CellFillPatternType::Gray125.as_str(), "gray125");
        assert_eq!(CellFillPatternType::DarkGray.as_str(), "darkGray");
        assert_eq!(CellFillPatternType::MediumGray.as_str(), "mediumGray");
        assert_eq!(CellFillPatternType::LightGray.as_str(), "lightGray");
        assert_eq!(CellFillPatternType::Gray0625.as_str(), "gray0625");
        assert_eq!(
            CellFillPatternType::DarkHorizontal.as_str(),
            "darkHorizontal"
        );
        assert_eq!(CellFillPatternType::DarkVertical.as_str(), "darkVertical");
        assert_eq!(CellFillPatternType::DarkDown.as_str(), "darkDown");
        assert_eq!(CellFillPatternType::DarkUp.as_str(), "darkUp");
        assert_eq!(CellFillPatternType::DarkGrid.as_str(), "darkGrid");
        assert_eq!(CellFillPatternType::DarkTrellis.as_str(), "darkTrellis");
    }

    #[test]
    fn test_cell_fill_pattern_type_debug() {
        let debug_str = format!("{:?}", CellFillPatternType::Solid);
        assert!(debug_str.contains("Solid"));
    }

    #[test]
    fn test_cell_border_default() {
        let border = CellBorder::default();
        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_none());
        assert!(border.diagonal.is_none());
    }

    #[test]
    fn test_cell_border_side() {
        let side = CellBorderSide {
            style: CellBorderLineStyle::Thin,
            color: Some("000000".to_string()),
        };
        assert_eq!(side.style.as_str(), "thin");
        assert_eq!(side.color, Some("000000".to_string()));
    }

    #[test]
    fn test_cell_border_line_style_as_str() {
        assert_eq!(CellBorderLineStyle::None.as_str(), "none");
        assert_eq!(CellBorderLineStyle::Thin.as_str(), "thin");
        assert_eq!(CellBorderLineStyle::Medium.as_str(), "medium");
        assert_eq!(CellBorderLineStyle::Dashed.as_str(), "dashed");
        assert_eq!(CellBorderLineStyle::Dotted.as_str(), "dotted");
        assert_eq!(CellBorderLineStyle::Thick.as_str(), "thick");
        assert_eq!(CellBorderLineStyle::Double.as_str(), "double");
        assert_eq!(CellBorderLineStyle::Hair.as_str(), "hair");
        assert_eq!(CellBorderLineStyle::MediumDashed.as_str(), "mediumDashed");
        assert_eq!(CellBorderLineStyle::DashDot.as_str(), "dashDot");
        assert_eq!(CellBorderLineStyle::MediumDashDot.as_str(), "mediumDashDot");
        assert_eq!(CellBorderLineStyle::DashDotDot.as_str(), "dashDotDot");
        assert_eq!(
            CellBorderLineStyle::MediumDashDotDot.as_str(),
            "mediumDashDotDot"
        );
        assert_eq!(CellBorderLineStyle::SlantDashDot.as_str(), "slantDashDot");
    }

    #[test]
    fn test_chart_type_debug() {
        let chart = Chart {
            chart_type: ChartType::Bar,
            title: Some("Test Chart".to_string()),
            data_range: "A1:B10".to_string(),
            position: (0, 0, 10, 10),
            show_legend: true,
        };
        assert!(format!("{:?}", chart).contains("Bar"));
    }

    #[test]
    fn test_data_validation_operator_as_str() {
        assert_eq!(DataValidationOperator::Between.as_str(), "between");
        assert_eq!(DataValidationOperator::NotBetween.as_str(), "notBetween");
        assert_eq!(DataValidationOperator::Equal.as_str(), "equal");
        assert_eq!(DataValidationOperator::NotEqual.as_str(), "notEqual");
        assert_eq!(DataValidationOperator::GreaterThan.as_str(), "greaterThan");
        assert_eq!(DataValidationOperator::LessThan.as_str(), "lessThan");
        assert_eq!(
            DataValidationOperator::GreaterThanOrEqual.as_str(),
            "greaterThanOrEqual"
        );
        assert_eq!(
            DataValidationOperator::LessThanOrEqual.as_str(),
            "lessThanOrEqual"
        );
    }

    #[test]
    fn test_data_validation_types() {
        let whole = DataValidationType::Whole {
            operator: DataValidationOperator::Between,
            value1: 1,
            value2: Some(10),
        };
        assert!(matches!(whole, DataValidationType::Whole { .. }));

        let decimal = DataValidationType::Decimal {
            operator: DataValidationOperator::GreaterThan,
            value1: 0.0,
            value2: None,
        };
        assert!(matches!(decimal, DataValidationType::Decimal { .. }));

        let list = DataValidationType::List {
            values: vec!["A".to_string(), "B".to_string()],
        };
        assert!(matches!(list, DataValidationType::List { .. }));

        let custom = DataValidationType::Custom {
            formula: "A1>0".to_string(),
        };
        assert!(matches!(custom, DataValidationType::Custom { .. }));
    }
}
