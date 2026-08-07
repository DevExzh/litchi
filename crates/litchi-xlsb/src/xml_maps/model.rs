//! Validated XLSB XML-map binding values.

use super::validation;
use crate::package::error::Result;

/// Maximum zero-based worksheet row.
pub(super) const MAX_ROW: u32 = 1_048_575;
/// Maximum zero-based worksheet column.
pub(super) const MAX_COLUMN: u32 = 16_383;
/// Default upper bound for an `XmlMappedXpath`, in UTF-16 code units.
pub(super) const MAX_XPATH_UNITS: usize = 31_999;

/// Resource ceilings for one XLSB binding part operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_bindings: usize,
    pub max_xpath_units: usize,
    pub max_records: usize,
    pub max_part_bytes: usize,
    pub max_opaque_records: usize,
    pub max_opaque_bytes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_bindings: 65_536,
        max_xpath_units: MAX_XPATH_UNITS,
        max_records: 1_000_000,
        max_part_bytes: 64 * 1024 * 1024,
        max_opaque_records: 65_536,
        max_opaque_bytes: 64 * 1024 * 1024,
    };
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// BIFF12 `XmlDataType`, whose wire domain is exactly `1..=0x2D`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct XmlDataType(u32);

impl XmlDataType {
    pub fn new(value: u32) -> Result<Self> {
        validation::xml_data_type(value)?;
        Ok(Self(value))
    }

    pub fn try_new(value: u32) -> Result<Self> {
        Self::new(value)
    }

    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for XmlDataType {
    type Error = crate::package::error::Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// Bounded, absolute XPath retained as inert lexical metadata.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct XPath(String);

impl XPath {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        Self::new_with_limit(value, MAX_XPATH_UNITS)
    }

    pub fn new_with_limits(value: impl Into<String>, limits: Limits) -> Result<Self> {
        Self::new_with_limit(value, limits.max_xpath_units)
    }

    fn new_with_limit(value: impl Into<String>, max_units: usize) -> Result<Self> {
        let value = value.into();
        validation::xpath(&value, max_units)?;
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for XPath {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Checked zero-based worksheet cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellReference {
    row: u32,
    column: u32,
}

impl CellReference {
    pub fn new(row: u32, column: u32) -> Result<Self> {
        validation::cell(row, column)?;
        Ok(Self { row, column })
    }

    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// One `BrtBeginListXmlCPr` table-column binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColumnBinding {
    column_id: u32,
    map_id: u32,
    data_type: XmlDataType,
    xpath: XPath,
    can_be_single: bool,
}

impl ColumnBinding {
    pub fn new(
        column_id: u32,
        map_id: u32,
        data_type: XmlDataType,
        xpath: XPath,
        can_be_single: bool,
    ) -> Result<Self> {
        validation::nonzero_id(column_id, "XML-mapped table column ID")?;
        validation::nonzero_id(map_id, "XML map ID")?;
        Ok(Self {
            column_id,
            map_id,
            data_type,
            xpath,
            can_be_single,
        })
    }

    #[must_use]
    pub const fn column_id(&self) -> u32 {
        self.column_id
    }

    #[must_use]
    pub const fn map_id(&self) -> u32 {
        self.map_id
    }

    #[must_use]
    pub const fn data_type(&self) -> XmlDataType {
        self.data_type
    }

    #[must_use]
    pub const fn datatype(&self) -> XmlDataType {
        self.data_type
    }

    #[must_use]
    pub const fn xpath(&self) -> &XPath {
        &self.xpath
    }

    #[must_use]
    pub const fn can_be_single(&self) -> bool {
        self.can_be_single
    }
}

/// Mapped columns from one ordinary Table part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MappedTable {
    table_id: u32,
    columns: Vec<ColumnBinding>,
}

impl MappedTable {
    pub fn new(table_id: u32, columns: Vec<ColumnBinding>) -> Result<Self> {
        let value = Self { table_id, columns };
        validation::mapped_table(&value, Limits::DEFAULT)?;
        Ok(value)
    }

    pub fn new_with_limits(
        table_id: u32,
        columns: Vec<ColumnBinding>,
        limits: Limits,
    ) -> Result<Self> {
        let value = Self { table_id, columns };
        validation::mapped_table(&value, limits)?;
        Ok(value)
    }

    #[must_use]
    pub const fn table_id(&self) -> u32 {
        self.table_id
    }

    #[must_use]
    pub fn columns(&self) -> &[ColumnBinding] {
        &self.columns
    }

    #[must_use]
    pub fn binding(&self, column_id: u32) -> Option<&ColumnBinding> {
        self.columns
            .iter()
            .find(|value| value.column_id == column_id)
    }
}

/// One XML map bound to one worksheet cell through a single-cell table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SingleCellBinding {
    table_id: u32,
    cell: CellReference,
    column: ColumnBinding,
}

impl SingleCellBinding {
    pub fn new(
        table_id: u32,
        column_id: u32,
        cell: CellReference,
        map_id: u32,
        data_type: XmlDataType,
        xpath: XPath,
    ) -> Result<Self> {
        validation::list_id(table_id, "single-cell table ID")?;
        Ok(Self {
            table_id,
            cell,
            column: ColumnBinding::new(column_id, map_id, data_type, xpath, true)?,
        })
    }

    #[must_use]
    pub const fn table_id(&self) -> u32 {
        self.table_id
    }

    #[must_use]
    pub const fn column_id(&self) -> u32 {
        self.column.column_id()
    }

    #[must_use]
    pub const fn cell(&self) -> CellReference {
        self.cell
    }

    #[must_use]
    pub const fn map_id(&self) -> u32 {
        self.column.map_id()
    }

    #[must_use]
    pub const fn data_type(&self) -> XmlDataType {
        self.column.data_type()
    }

    #[must_use]
    pub const fn datatype(&self) -> XmlDataType {
        self.column.data_type()
    }

    #[must_use]
    pub const fn xpath(&self) -> &XPath {
        self.column.xpath()
    }

    #[must_use]
    pub const fn can_be_single(&self) -> bool {
        true
    }

    #[must_use]
    pub const fn column_binding(&self) -> &ColumnBinding {
        &self.column
    }
}

/// Contextual alias emphasizing the BIFF table envelope.
pub type SingleCellTable = SingleCellBinding;
