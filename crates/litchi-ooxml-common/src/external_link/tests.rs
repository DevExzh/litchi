use super::*;
use litchi_opc::constants::relationship_type;

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
