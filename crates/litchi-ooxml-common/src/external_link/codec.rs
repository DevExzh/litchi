//! External-workbook relationship classification.

use super::model::EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES;

/// Whether `reltype` may be used to target an external workbook.
pub fn is_external_workbook_relationship(reltype: &str) -> bool {
    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES.contains(&reltype)
}
