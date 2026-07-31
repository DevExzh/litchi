//! Relationship types that may target an external workbook.
//!
//! ECMA-376 defines a single `externalLinkPath` relationship type, but Excel
//! also emits the Microsoft extension families documented by MS-OE376 and
//! MS-XLSB to record *how* a link was resolved: workbooks opened from the
//! startup, alternate-startup, or library folders, links whose target could
//! not be located (`xlPathMissing`), and the long-path variants that carry
//! targets exceeding the classic path limit.
//!
//! A reader that only accepts the base type rejects otherwise-valid workbooks
//! outright, so the XLSX and XLSB external-link readers and writers all
//! validate against the single list below.
//!
//! Targets are inert. Recognising a relationship type never opens, resolves,
//! contacts, refreshes, or otherwise follows the referenced workbook; the
//! stored target string is only ever surfaced to the caller verbatim.

use litchi_opc::constants::relationship_type;

/// Every relationship type accepted as an external-workbook target.
///
/// Ordered standard-first so the common case matches on the first comparison.
pub const EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES: &[&str] = &[
    // ECMA-376 transitional and strict.
    relationship_type::EXTERNAL_LINK_PATH,
    relationship_type::STRICT_EXTERNAL_LINK_PATH,
    // MS-OE376 resolution-origin families.
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
    // Long-path families for targets beyond the classic path limit.
    "http://schemas.microsoft.com/office/2019/04/relationships/externalLinkLongPath",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlStartup",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlPathMissing",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlLibrary",
];

/// Whether `reltype` may be used to target an external workbook.
pub fn is_external_workbook_relationship(reltype: &str) -> bool {
    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES.contains(&reltype)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_the_standard_transitional_and_strict_types() {
        assert!(is_external_workbook_relationship(
            relationship_type::EXTERNAL_LINK_PATH
        ));
        assert!(is_external_workbook_relationship(
            relationship_type::STRICT_EXTERNAL_LINK_PATH
        ));
    }

    #[test]
    fn accepts_microsoft_resolution_origin_families() {
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing"
        ));
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup"
        ));
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup"
        ));
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary"
        ));
    }

    #[test]
    fn accepts_long_path_families() {
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2019/04/relationships/externalLinkLongPath"
        ));
        assert!(is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlPathMissing"
        ));
    }

    #[test]
    fn rejects_unrelated_relationship_types() {
        assert!(!is_external_workbook_relationship(
            relationship_type::HYPERLINK
        ));
        assert!(!is_external_workbook_relationship(""));
        assert!(!is_external_workbook_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath"
        ));
    }

    #[test]
    fn the_list_has_no_duplicates() {
        let mut sorted = EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES.to_vec();
        sorted.sort_unstable();
        let total = sorted.len();
        sorted.dedup();
        assert_eq!(sorted.len(), total);
    }
}
