//! Semantic table-template vocabulary and ergonomic typed accessors.

/// Legacy row/column selector used by deprecated table-template edge attributes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Axis {
    Row,
    Column,
}

/// Cell and optional paragraph styles for one table-template region.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    pub style_name: String,
    pub paragraph_style_name: Option<String>,
}

impl Style {
    /// Create a cell style reference without a paragraph style.
    pub fn new(style_name: impl Into<String>) -> Self {
        Self {
            style_name: style_name.into(),
            paragraph_style_name: None,
        }
    }

    /// Attach the optional paragraph style used by text-bearing regions.
    pub fn with_paragraph_style(mut self, style_name: impl Into<String>) -> Self {
        self.paragraph_style_name = Some(style_name.into());
        self
    }
}

/// Typed table-template region selector.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Region {
    FirstRow,
    LastRow,
    FirstColumn,
    LastColumn,
    Body,
    EvenRows,
    OddRows,
    EvenColumns,
    OddColumns,
    Background,
}

impl Region {
    /// All regions in deterministic ODF serialization order.
    pub const ALL: [Self; 10] = [
        Self::FirstRow,
        Self::LastRow,
        Self::FirstColumn,
        Self::LastColumn,
        Self::Body,
        Self::EvenRows,
        Self::OddRows,
        Self::EvenColumns,
        Self::OddColumns,
        Self::Background,
    ];

    pub(crate) const fn name(self) -> &'static str {
        match self {
            Self::FirstRow => "first-row",
            Self::LastRow => "last-row",
            Self::FirstColumn => "first-column",
            Self::LastColumn => "last-column",
            Self::Body => "body",
            Self::EvenRows => "even-rows",
            Self::OddRows => "odd-rows",
            Self::EvenColumns => "even-columns",
            Self::OddColumns => "odd-columns",
            Self::Background => "background",
        }
    }

    pub(crate) fn from_name(value: &str) -> Option<Self> {
        Some(match value {
            "first-row" => Self::FirstRow,
            "last-row" => Self::LastRow,
            "first-column" => Self::FirstColumn,
            "last-column" => Self::LastColumn,
            "body" => Self::Body,
            "even-rows" => Self::EvenRows,
            "odd-rows" => Self::OddRows,
            "even-columns" => Self::EvenColumns,
            "odd-columns" => Self::OddColumns,
            "background" => Self::Background,
            _ => return None,
        })
    }
}

/// Named cell-style regions which make up an ODF table template.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Template {
    pub name: String,
    pub first_row_start_column: Option<Axis>,
    pub first_row_end_column: Option<Axis>,
    pub last_row_start_column: Option<Axis>,
    pub last_row_end_column: Option<Axis>,
    pub use_first_row_styles: Option<bool>,
    pub use_last_row_styles: Option<bool>,
    pub use_first_column_styles: Option<bool>,
    pub use_last_column_styles: Option<bool>,
    pub use_banding_rows_styles: Option<bool>,
    pub use_banding_columns_styles: Option<bool>,
    pub first_row: Option<Style>,
    pub last_row: Option<Style>,
    pub first_column: Option<Style>,
    pub last_column: Option<Style>,
    pub body: Option<Style>,
    pub even_rows: Option<Style>,
    pub odd_rows: Option<Style>,
    pub even_columns: Option<Style>,
    pub odd_columns: Option<Style>,
    pub background: Option<Style>,
}

impl Template {
    /// Create an empty template with a semantic name.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            first_row_start_column: None,
            first_row_end_column: None,
            last_row_start_column: None,
            last_row_end_column: None,
            use_first_row_styles: None,
            use_last_row_styles: None,
            use_first_column_styles: None,
            use_last_column_styles: None,
            use_banding_rows_styles: None,
            use_banding_columns_styles: None,
            first_row: None,
            last_row: None,
            first_column: None,
            last_column: None,
            body: None,
            even_rows: None,
            odd_rows: None,
            even_columns: None,
            odd_columns: None,
            background: None,
        }
    }

    /// Return the style assigned to a typed region.
    pub fn region(&self, region: Region) -> Option<&Style> {
        match region {
            Region::FirstRow => self.first_row.as_ref(),
            Region::LastRow => self.last_row.as_ref(),
            Region::FirstColumn => self.first_column.as_ref(),
            Region::LastColumn => self.last_column.as_ref(),
            Region::Body => self.body.as_ref(),
            Region::EvenRows => self.even_rows.as_ref(),
            Region::OddRows => self.odd_rows.as_ref(),
            Region::EvenColumns => self.even_columns.as_ref(),
            Region::OddColumns => self.odd_columns.as_ref(),
            Region::Background => self.background.as_ref(),
        }
    }

    /// Return a mutable typed region slot for snapshot-style edits.
    pub fn region_mut(&mut self, region: Region) -> &mut Option<Style> {
        match region {
            Region::FirstRow => &mut self.first_row,
            Region::LastRow => &mut self.last_row,
            Region::FirstColumn => &mut self.first_column,
            Region::LastColumn => &mut self.last_column,
            Region::Body => &mut self.body,
            Region::EvenRows => &mut self.even_rows,
            Region::OddRows => &mut self.odd_rows,
            Region::EvenColumns => &mut self.even_columns,
            Region::OddColumns => &mut self.odd_columns,
            Region::Background => &mut self.background,
        }
    }

    /// Set or clear a typed region without exposing its storage field.
    pub fn set_region(&mut self, region: Region, style: Option<Style>) {
        *self.region_mut(region) = style;
    }

    /// Return an owned template with one typed region assigned.
    pub fn with_region(mut self, region: Region, style: Style) -> Self {
        self.set_region(region, Some(style));
        self
    }
}
