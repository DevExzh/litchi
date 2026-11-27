//! Contextual semantic operations for `PivotTable` OLAP values.

use super::model::{PivotFieldOlapExt, PivotHierarchy, PivotHierarchyAxis, PivotViewOlapHeader};

impl PivotViewOlapHeader {
    /// Whether the record carries bytes reserved for a future OLAP version.
    #[must_use]
    pub const fn has_future_extensions(&self) -> bool {
        !self.future_bytes.is_empty()
    }
}

impl PivotHierarchyAxis {
    /// Whether no `PivotTable` axis is selected by this `SXAxis` value.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        !self.row && !self.column && !self.page && !self.data
    }

    /// Number of selected axes represented by this value.
    #[must_use]
    pub const fn axis_count(self) -> u8 {
        self.row as u8 + self.column as u8 + self.page as u8 + self.data as u8
    }
}

impl PivotHierarchy {
    /// Number of OLAP levels represented by `rgisxvd` (`cisxvd`).
    #[must_use]
    pub fn level_count(&self) -> usize {
        self.level_fields.len()
    }

    /// Number of hierarchy levels with hidden-member metadata.
    #[must_use]
    pub fn hidden_level_count(&self) -> usize {
        self.hidden_member_sets.len()
    }
}

impl PivotFieldOlapExt {
    /// Number of `SXVIFlags` entries in `rgsxvi` (`csxvi`).
    #[must_use]
    pub fn item_count(&self) -> usize {
        self.item_flags.len()
    }
}
