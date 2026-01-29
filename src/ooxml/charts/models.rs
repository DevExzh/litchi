//! Chart data models.
//!
//! This module contains the core data structures for representing chart data,
//! including series, data sources, and numeric/text data.

/// A reference to a data source (cell range formula).
#[derive(Debug, Clone, PartialEq)]
pub struct DataSourceRef {
    /// Formula reference (e.g., "Sheet1!$A$1:$A$10")
    pub formula: String,
}

impl DataSourceRef {
    /// Create a new data source reference.
    #[inline]
    pub fn new(formula: impl Into<String>) -> Self {
        Self {
            formula: formula.into(),
        }
    }
}

/// Numeric data with optional cached values.
#[derive(Debug, Clone)]
pub struct NumericData {
    /// Optional reference to cell range
    pub source_ref: Option<DataSourceRef>,
    /// Cached numeric values
    pub values: Vec<f64>,
    /// Format code for display
    pub format_code: Option<String>,
}

impl NumericData {
    /// Create a new numeric data set with values.
    #[inline]
    pub fn from_values(values: Vec<f64>) -> Self {
        Self {
            source_ref: None,
            values,
            format_code: None,
        }
    }

    /// Create a new numeric data set with a reference.
    #[inline]
    pub fn from_ref(formula: impl Into<String>) -> Self {
        Self {
            source_ref: Some(DataSourceRef::new(formula)),
            values: Vec::new(),
            format_code: None,
        }
    }

    /// Set the format code.
    #[inline]
    pub fn with_format_code(mut self, format_code: impl Into<String>) -> Self {
        self.format_code = Some(format_code.into());
        self
    }

    /// Add cached values.
    #[inline]
    pub fn with_cached_values(mut self, values: Vec<f64>) -> Self {
        self.values = values;
        self
    }
}

/// String data with optional cached values.
#[derive(Debug, Clone)]
pub struct StringData {
    /// Optional reference to cell range
    pub source_ref: Option<DataSourceRef>,
    /// Cached string values
    pub values: Vec<String>,
}

impl StringData {
    /// Create a new string data set with values.
    #[inline]
    pub fn from_values(values: Vec<String>) -> Self {
        Self {
            source_ref: None,
            values,
        }
    }

    /// Create a new string data set with a reference.
    #[inline]
    pub fn from_ref(formula: impl Into<String>) -> Self {
        Self {
            source_ref: Some(DataSourceRef::new(formula)),
            values: Vec::new(),
        }
    }

    /// Add cached values.
    #[inline]
    pub fn with_cached_values(mut self, values: Vec<String>) -> Self {
        self.values = values;
        self
    }
}

/// Multi-level string data (for hierarchical categories).
#[derive(Debug, Clone)]
pub struct MultiLevelStringData {
    /// Multiple levels of string data
    pub levels: Vec<StringData>,
}

impl MultiLevelStringData {
    /// Create a new multi-level string data set.
    #[inline]
    pub fn new() -> Self {
        Self { levels: Vec::new() }
    }

    /// Add a level.
    #[inline]
    pub fn add_level(mut self, level: StringData) -> Self {
        self.levels.push(level);
        self
    }
}

impl Default for MultiLevelStringData {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Number format specification.
#[derive(Debug, Clone)]
pub struct NumberFormat {
    /// Format code (e.g., "General", "0.00", "#,##0")
    pub format_code: String,
    /// Whether the format is linked to source data
    pub source_linked: bool,
}

impl NumberFormat {
    /// Create a new number format.
    #[inline]
    pub fn new(format_code: impl Into<String>) -> Self {
        Self {
            format_code: format_code.into(),
            source_linked: true,
        }
    }

    /// Create a General format.
    #[inline]
    pub fn general() -> Self {
        Self::new("General")
    }

    /// Set whether the format is linked to source.
    #[inline]
    pub fn with_source_linked(mut self, linked: bool) -> Self {
        self.source_linked = linked;
        self
    }
}

impl Default for NumberFormat {
    #[inline]
    fn default() -> Self {
        Self::general()
    }
}

/// Layout information for chart elements.
#[derive(Debug, Clone)]
pub struct Layout {
    /// X position (0.0 to 1.0 for factor mode)
    pub x: Option<f64>,
    /// Y position (0.0 to 1.0 for factor mode)
    pub y: Option<f64>,
    /// Width (0.0 to 1.0 for factor mode)
    pub width: Option<f64>,
    /// Height (0.0 to 1.0 for factor mode)
    pub height: Option<f64>,
    /// X mode (edge or factor)
    pub x_mode: Option<crate::ooxml::charts::types::LayoutMode>,
    /// Y mode (edge or factor)
    pub y_mode: Option<crate::ooxml::charts::types::LayoutMode>,
    /// Width mode (edge or factor)
    pub width_mode: Option<crate::ooxml::charts::types::LayoutMode>,
    /// Height mode (edge or factor)
    pub height_mode: Option<crate::ooxml::charts::types::LayoutMode>,
    /// Layout target (inner or outer)
    pub target: Option<crate::ooxml::charts::types::LayoutTarget>,
}

impl Layout {
    /// Create a new manual layout.
    #[inline]
    pub fn new() -> Self {
        Self {
            x: None,
            y: None,
            width: None,
            height: None,
            x_mode: None,
            y_mode: None,
            width_mode: None,
            height_mode: None,
            target: None,
        }
    }

    /// Set position.
    #[inline]
    pub fn with_position(mut self, x: f64, y: f64) -> Self {
        self.x = Some(x);
        self.y = Some(y);
        self
    }

    /// Set size.
    #[inline]
    pub fn with_size(mut self, width: f64, height: f64) -> Self {
        self.width = Some(width);
        self.height = Some(height);
        self
    }
}

impl Default for Layout {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Text content for titles and labels.
#[derive(Debug, Clone)]
pub struct RichText {
    /// Text content
    pub text: String,
}

impl RichText {
    /// Create a new rich text.
    #[inline]
    pub fn new(text: impl Into<String>) -> Self {
        Self { text: text.into() }
    }
}

/// Title text source (can be from formula or literal).
#[derive(Debug, Clone)]
pub enum TitleText {
    /// Literal text
    Literal(RichText),
    /// Reference to a cell
    Reference(DataSourceRef),
}

impl TitleText {
    /// Create from a string.
    #[inline]
    pub fn from_string(text: impl Into<String>) -> Self {
        Self::Literal(RichText::new(text))
    }

    /// Create from a formula reference.
    #[inline]
    pub fn from_ref(formula: impl Into<String>) -> Self {
        Self::Reference(DataSourceRef::new(formula))
    }
}
