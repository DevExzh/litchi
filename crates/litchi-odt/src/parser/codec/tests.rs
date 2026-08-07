//! Focused regression coverage for codec-wide validation invariants.

use super::validation::{
    checked_semantic_depth, parse_boolean, parse_tracked_change_bool, validate_protection_key,
    validate_tracked_change_text,
};

#[test]
fn accepts_schema_boolean_forms() {
    assert!(parse_boolean("true", "flag").unwrap());
    assert!(!parse_boolean("0", "flag").unwrap());
    assert!(parse_tracked_change_bool("track-changes", "1").unwrap());
}

#[test]
fn rejects_invalid_boolean_forms() {
    assert!(parse_boolean("yes", "flag").is_err());
    assert!(parse_tracked_change_bool("track-changes", "yes").is_err());
}

#[test]
fn bounds_depth_and_text() {
    assert_eq!(checked_semantic_depth(0, "document").unwrap(), 1);
    assert!(validate_tracked_change_text("", "id", false).is_err());
    assert!(validate_tracked_change_text("ok", "id", false).is_ok());
}

#[test]
fn validates_base64_protection_keys() {
    assert!(validate_protection_key("YWJj").is_ok());
    assert!(validate_protection_key("not-base64").is_err());
}
