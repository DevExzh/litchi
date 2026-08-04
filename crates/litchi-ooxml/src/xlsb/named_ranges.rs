//! Compatibility facade for the XLSB defined-name owner.
//!
//! `litchi-xlsb` owns the `BrtName` model and codec. This adapter retains the
//! historical host type and maps standalone owner errors into `XlsbError`;
//! workbook record ordering and package traversal remain in this crate.

use crate::xlsb::error::{XlsbError, XlsbResult};

/// Historical host representation of a defined name.
#[derive(Debug, Clone)]
pub struct NamedRange {
    /// Name of the range.
    pub name: String,
    /// Raw `NameParsedFormula.rgce` bytes.
    pub formula: Option<Vec<u8>>,
    /// Sheet ID (`None` for workbook scope).
    pub sheet_id: Option<u32>,
    /// Whether the name is hidden.
    pub hidden: bool,
    /// Whether the name is a function.
    pub function: bool,
}

impl From<litchi_xlsb::named_ranges::Definition> for NamedRange {
    fn from(definition: litchi_xlsb::named_ranges::Definition) -> Self {
        Self {
            name: definition.name,
            formula: definition.formula,
            sheet_id: definition.sheet_id,
            hidden: definition.hidden,
            function: definition.function,
        }
    }
}

impl NamedRange {
    /// Create a new named range.
    #[must_use]
    pub fn new(name: String, sheet_id: Option<u32>) -> Self {
        litchi_xlsb::named_ranges::Definition::new(name, sheet_id).into()
    }

    /// Set formula bytes.
    #[must_use]
    pub fn with_formula(mut self, formula: Vec<u8>) -> Self {
        self.formula = Some(formula);
        self
    }

    /// Set the hidden flag.
    #[must_use]
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }

    /// Create a 3D area formula token stream.
    pub fn area3d_formula(
        sheet_id: u32,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsbResult<Vec<u8>> {
        area3d_formula(sheet_id, first_row, last_row, first_col, last_col)
    }

    /// Historical method spelling retained for source compatibility.
    pub fn create_area3d_formula(
        sheet_id: u32,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsbResult<Vec<u8>> {
        Self::area3d_formula(sheet_id, first_row, last_row, first_col, last_col)
    }

    /// Parse one complete `BrtName` payload.
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        litchi_xlsb::named_ranges::parse(data)
            .map(Self::from)
            .map_err(Into::into)
    }
}

/// Validate a defined name using the standalone XLSB owner.
pub(crate) fn validate_name(name: &str) -> XlsbResult<()> {
    litchi_xlsb::named_ranges::validate_name(name).map_err(Into::into)
}

/// Historical validator spelling retained for the host implementation.
pub(crate) use validate_name as validate_defined_name;

/// Create a 3D area formula token stream.
pub fn area3d_formula(
    sheet_id: u32,
    first_row: u32,
    last_row: u32,
    first_col: u16,
    last_col: u16,
) -> XlsbResult<Vec<u8>> {
    litchi_xlsb::named_ranges::area3d_formula(sheet_id, first_row, last_row, first_col, last_col)
        .map_err(Into::into)
}

/// Historical helper spelling retained for source compatibility.
pub use area3d_formula as create_area3d_formula;

impl From<litchi_xlsb::named_ranges::Error> for XlsbError {
    fn from(error: litchi_xlsb::named_ranges::Error) -> Self {
        let message = error.to_string();
        match error {
            litchi_xlsb::named_ranges::Error::Wire(error) => Self::from(error),
            litchi_xlsb::named_ranges::Error::Formula(error) => Self::from(error),
            litchi_xlsb::named_ranges::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            litchi_xlsb::named_ranges::Error::InvalidFormula(message) => {
                Self::InvalidFormula(message)
            },
            _ => Self::InvalidFormula(message),
        }
    }
}
