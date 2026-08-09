use crate::error::Result;

use super::codec::{
    CORE_NAMESPACE, STRICT_NAMESPACE, bounded, parse_sqref, validate_collection, validate_metadata,
    validate_range, validate_sheet_protection, validate_strong_verifier, validate_xml_text,
};

/// `SpreadsheetML` namespace form used by the deterministic writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => CORE_NAMESPACE,
            Self::Strict => STRICT_NAMESPACE,
        }
    }
}

/// Source schema for a protected-range collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProtectedRangeSource {
    /// ISO/IEC 29500 `SpreadsheetML` collection.
    Core,
    /// Office 2010 `x14` worksheet extension.
    Office2010,
}

/// Password-verifier metadata. This type does not verify passwords.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProtectionPasswordVerifier {
    /// Legacy 16-bit `SpreadsheetML` password verifier.
    Legacy(u16),
    /// Salted iterative password hash metadata.
    Strong(StrongProtectionPasswordVerifier),
}

/// Salted iterative password hash metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StrongProtectionPasswordVerifier {
    pub(crate) algorithm_name: String,
    pub(crate) hash_value: Vec<u8>,
    pub(crate) salt_value: Vec<u8>,
    pub(crate) spin_count: u32,
}

impl StrongProtectionPasswordVerifier {
    pub fn new(
        algorithm_name: impl Into<String>,
        hash_value: Vec<u8>,
        salt_value: Vec<u8>,
        spin_count: u32,
    ) -> Result<Self> {
        let value = Self {
            algorithm_name: algorithm_name.into(),
            hash_value,
            salt_value,
            spin_count,
        };
        validate_strong_verifier(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn algorithm_name(&self) -> &str {
        &self.algorithm_name
    }
    #[must_use]
    pub fn hash_value(&self) -> &[u8] {
        &self.hash_value
    }
    #[must_use]
    pub fn salt_value(&self) -> &[u8] {
        &self.salt_value
    }
    #[must_use]
    pub fn spin_count(&self) -> u32 {
        self.spin_count
    }
}

/// Effective operation locks from a worksheet `sheetProtection` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Protection {
    pub(crate) verifier: Option<ProtectionPasswordVerifier>,
    pub(crate) sheet: bool,
    pub(crate) objects: bool,
    pub(crate) scenarios: bool,
    pub(crate) format_cells: bool,
    pub(crate) format_columns: bool,
    pub(crate) format_rows: bool,
    pub(crate) insert_columns: bool,
    pub(crate) insert_rows: bool,
    pub(crate) insert_hyperlinks: bool,
    pub(crate) delete_columns: bool,
    pub(crate) delete_rows: bool,
    pub(crate) select_locked_cells: bool,
    pub(crate) sort: bool,
    pub(crate) auto_filter: bool,
    pub(crate) pivot_tables: bool,
    pub(crate) select_unlocked_cells: bool,
}

impl Protection {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.verifier.as_ref()
    }
    #[must_use]
    pub fn sheet_locked(&self) -> bool {
        self.sheet
    }
    #[must_use]
    pub fn objects_locked(&self) -> bool {
        self.objects
    }
    #[must_use]
    pub fn scenarios_locked(&self) -> bool {
        self.scenarios
    }
    #[must_use]
    pub fn format_cells_locked(&self) -> bool {
        self.format_cells
    }
    #[must_use]
    pub fn format_columns_locked(&self) -> bool {
        self.format_columns
    }
    #[must_use]
    pub fn format_rows_locked(&self) -> bool {
        self.format_rows
    }
    #[must_use]
    pub fn insert_columns_locked(&self) -> bool {
        self.insert_columns
    }
    #[must_use]
    pub fn insert_rows_locked(&self) -> bool {
        self.insert_rows
    }
    #[must_use]
    pub fn insert_hyperlinks_locked(&self) -> bool {
        self.insert_hyperlinks
    }
    #[must_use]
    pub fn delete_columns_locked(&self) -> bool {
        self.delete_columns
    }
    #[must_use]
    pub fn delete_rows_locked(&self) -> bool {
        self.delete_rows
    }
    #[must_use]
    pub fn select_locked_cells_locked(&self) -> bool {
        self.select_locked_cells
    }
    #[must_use]
    pub fn sort_locked(&self) -> bool {
        self.sort
    }
    #[must_use]
    pub fn auto_filter_locked(&self) -> bool {
        self.auto_filter
    }
    #[must_use]
    pub fn pivot_tables_locked(&self) -> bool {
        self.pivot_tables
    }
    #[must_use]
    pub fn select_unlocked_cells_locked(&self) -> bool {
        self.select_unlocked_cells
    }

    pub fn set_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) -> Result<()> {
        if let Some(ProtectionPasswordVerifier::Strong(value)) = verifier.as_ref() {
            validate_strong_verifier(value)?;
        }
        self.verifier = verifier;
        Ok(())
    }

    pub fn set_sheet_locked(&mut self, value: bool) {
        self.sheet = value;
    }
    pub fn set_objects_locked(&mut self, value: bool) {
        self.objects = value;
    }
    pub fn set_scenarios_locked(&mut self, value: bool) {
        self.scenarios = value;
    }
    pub fn set_format_cells_locked(&mut self, value: bool) {
        self.format_cells = value;
    }
    pub fn set_format_columns_locked(&mut self, value: bool) {
        self.format_columns = value;
    }
    pub fn set_format_rows_locked(&mut self, value: bool) {
        self.format_rows = value;
    }
    pub fn set_insert_columns_locked(&mut self, value: bool) {
        self.insert_columns = value;
    }
    pub fn set_insert_rows_locked(&mut self, value: bool) {
        self.insert_rows = value;
    }
    pub fn set_insert_hyperlinks_locked(&mut self, value: bool) {
        self.insert_hyperlinks = value;
    }
    pub fn set_delete_columns_locked(&mut self, value: bool) {
        self.delete_columns = value;
    }
    pub fn set_delete_rows_locked(&mut self, value: bool) {
        self.delete_rows = value;
    }
    pub fn set_select_locked_cells_locked(&mut self, value: bool) {
        self.select_locked_cells = value;
    }
    pub fn set_sort_locked(&mut self, value: bool) {
        self.sort = value;
    }
    pub fn set_auto_filter_locked(&mut self, value: bool) {
        self.auto_filter = value;
    }
    pub fn set_pivot_tables_locked(&mut self, value: bool) {
        self.pivot_tables = value;
    }
    pub fn set_select_unlocked_cells_locked(&mut self, value: bool) {
        self.select_unlocked_cells = value;
    }
}

impl Default for Protection {
    fn default() -> Self {
        Self {
            verifier: None,
            sheet: false,
            objects: false,
            scenarios: false,
            format_cells: true,
            format_columns: true,
            format_rows: true,
            insert_columns: true,
            insert_rows: true,
            insert_hyperlinks: true,
            delete_columns: true,
            delete_rows: true,
            select_locked_cells: false,
            sort: true,
            auto_filter: true,
            pivot_tables: true,
            select_unlocked_cells: false,
        }
    }
}

/// Typed kind of an individual protected-range reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionRangeReferenceKind {
    Cells {
        start_row: u32,
        start_column: u32,
        end_row: u32,
        end_column: u32,
    },
    Columns {
        start_column: u32,
        end_column: u32,
    },
    Rows {
        start_row: u32,
        end_row: u32,
    },
}

/// One validated reference in a protected range's `sqref`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeReference {
    pub(crate) raw: String,
    pub(crate) kind: ProtectionRangeReferenceKind,
}

impl ProtectionRangeReference {
    #[must_use]
    pub fn raw(&self) -> &str {
        &self.raw
    }
    #[must_use]
    pub fn kind(&self) -> ProtectionRangeReferenceKind {
        self.kind
    }
}

/// Validated whitespace-separated protected-range references.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeSqref {
    pub(crate) raw: String,
    pub(crate) references: Vec<ProtectionRangeReference>,
}

impl ProtectionRangeSqref {
    pub fn parse(value: impl AsRef<str>) -> Result<Self> {
        parse_sqref(value.as_ref())
    }

    #[must_use]
    pub fn raw(&self) -> &str {
        &self.raw
    }
    #[must_use]
    pub fn references(&self) -> &[ProtectionRangeReference] {
        &self.references
    }
}

/// A single editable range associated with worksheet protection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectedRange {
    pub(crate) source: ProtectedRangeSource,
    pub(crate) name: String,
    pub(crate) sqref: ProtectionRangeSqref,
    pub(crate) verifier: Option<ProtectionPasswordVerifier>,
    pub(crate) security_descriptor: Option<String>,
}

impl ProtectedRange {
    pub fn new(
        source: ProtectedRangeSource,
        name: impl Into<String>,
        sqref: ProtectionRangeSqref,
    ) -> Result<Self> {
        let value = Self {
            source,
            name: name.into(),
            sqref,
            verifier: None,
            security_descriptor: None,
        };
        validate_range(&value)?;
        Ok(value)
    }

    pub fn set_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) -> Result<()> {
        if let Some(ProtectionPasswordVerifier::Strong(value)) = verifier.as_ref() {
            validate_strong_verifier(value)?;
        }
        self.verifier = verifier;
        Ok(())
    }

    pub fn set_security_descriptor(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            bounded(value, "securityDescriptor")?;
            validate_xml_text(value, "securityDescriptor")?;
        }
        self.security_descriptor = value;
        Ok(())
    }

    #[must_use]
    pub fn source(&self) -> ProtectedRangeSource {
        self.source
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub fn sqref(&self) -> &ProtectionRangeSqref {
        &self.sqref
    }
    #[must_use]
    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.verifier.as_ref()
    }
    #[must_use]
    pub fn security_descriptor(&self) -> Option<&str> {
        self.security_descriptor.as_deref()
    }
}

/// A protected-range container in worksheet document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectedRangeCollection {
    pub(crate) source: ProtectedRangeSource,
    pub(crate) ranges: Vec<ProtectedRange>,
}

impl ProtectedRangeCollection {
    pub fn new(source: ProtectedRangeSource, ranges: Vec<ProtectedRange>) -> Result<Self> {
        let value = Self { source, ranges };
        validate_collection(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn source(&self) -> ProtectedRangeSource {
        self.source
    }
    #[must_use]
    pub fn ranges(&self) -> &[ProtectedRange] {
        &self.ranges
    }
}

/// Complete worksheet protection metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Metadata {
    pub(crate) sheet_protection: Option<Protection>,
    pub(crate) protected_range_collections: Vec<ProtectedRangeCollection>,
}

impl Metadata {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn sheet_protection(&self) -> Option<&Protection> {
        self.sheet_protection.as_ref()
    }
    #[must_use]
    pub fn protected_range_collections(&self) -> &[ProtectedRangeCollection] {
        &self.protected_range_collections
    }
    pub fn protected_ranges(&self) -> impl Iterator<Item = &ProtectedRange> {
        self.protected_range_collections
            .iter()
            .flat_map(|collection| collection.ranges.iter())
    }

    pub fn set_sheet_protection(&mut self, value: Option<Protection>) -> Result<()> {
        if let Some(value) = value.as_ref() {
            validate_sheet_protection(value)?;
        }
        self.sheet_protection = value;
        Ok(())
    }

    pub fn set_protected_range_collections(
        &mut self,
        value: Vec<ProtectedRangeCollection>,
    ) -> Result<()> {
        let candidate = Self {
            sheet_protection: self.sheet_protection.clone(),
            protected_range_collections: value,
        };
        validate_metadata(&candidate)?;
        self.protected_range_collections = candidate.protected_range_collections;
        Ok(())
    }

    pub fn clear_sheet_protection(&mut self) {
        self.sheet_protection = None;
    }

    pub fn clear_protected_ranges(&mut self) {
        self.protected_range_collections.clear();
    }
}
