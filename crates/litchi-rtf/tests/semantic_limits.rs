#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{EditOptions, Package, SemanticLimitKind, SemanticLimits, edit::Error};

fn semantic_limits(max_operations: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .build()
        .unwrap()
}

#[test]
fn checked_profiles_cover_zero_default_and_exact_supported_boundaries() {
    let default = SemanticLimits::default();
    assert_eq!(default.max_operations(), 256);

    let zero = semantic_limits(0);
    assert_eq!(zero.max_operations(), 0);

    let one = semantic_limits(1);
    assert_eq!(one.max_operations(), 1);

    let maximum = semantic_limits(isize::MAX as usize);
    assert_eq!(maximum.max_operations(), isize::MAX as usize);
}

#[test]
fn usize_max_is_rejected_during_profile_construction() {
    assert!(matches!(
        SemanticLimits::builder().max_operations(usize::MAX),
        Err(Error::InvalidLimit {
            kind: SemanticLimitKind::Operations,
            value: usize::MAX,
            maximum: isize::MAX as usize,
        })
    ));
}

#[test]
fn limit_refusal_keeps_source_immutable_and_edit_reusable() {
    let source = Package::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(1)));
    assert_eq!(edit.options().semantic_limits_value(), semantic_limits(1));

    edit.replace_text(litchi_rtf::edit::TextSpan::new(0, 5).unwrap(), "First")
        .unwrap();
    assert!(matches!(
        edit.replace_text(litchi_rtf::edit::TextSpan::new(6, 10).unwrap(), "Second"),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 2,
            maximum: 1,
        })
    ));
    assert_eq!(edit.operation_count(), 1);

    let commit = edit.commit().unwrap();
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    assert_eq!(commit.package().text(), "First Beta");
}

#[test]
fn zero_profile_refuses_the_first_mutation_without_touching_source() {
    let source = Package::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(0)));
    assert_eq!(edit.options().semantic_limits_value(), semantic_limits(0));

    assert!(matches!(
        edit.replace_text(litchi_rtf::edit::TextSpan::new(0, 5).unwrap(), "Other"),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 1,
            maximum: 0,
        })
    ));
    assert_eq!(edit.operation_count(), 0);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    assert!(edit.commit().unwrap().package().same_snapshot(&source));
}
