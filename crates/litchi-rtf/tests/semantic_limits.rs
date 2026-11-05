#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::{
    EditOptions, HeaderFooterType, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SmartTagAttribute, SourceLimits, TableCellPath,
    edit::{Error, HeaderFooterParagraph, ParagraphTextReplacement, TextSpan},
};
use std::borrow::Cow;

fn semantic_limits(max_operations: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .build()
        .unwrap()
}

fn replacement_limits(max_operations: usize, max_replacement_bytes: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .max_replacement_bytes(max_replacement_bytes)
        .unwrap()
        .build()
        .unwrap()
}

#[test]
fn checked_profiles_cover_zero_default_and_exact_supported_boundaries() {
    let default = SemanticLimits::default();
    assert_eq!(default.max_operations(), 256);
    assert_eq!(default.max_replacement_bytes(), 512 * 1_048_576);

    let zero = semantic_limits(0);
    assert_eq!(zero.max_operations(), 0);

    let zero_replacements = replacement_limits(256, 0);
    assert_eq!(zero_replacements.max_replacement_bytes(), 0);

    let one = semantic_limits(1);
    assert_eq!(one.max_operations(), 1);

    let maximum = semantic_limits(isize::MAX as usize);
    assert_eq!(maximum.max_operations(), isize::MAX as usize);

    let maximum_replacements = replacement_limits(256, isize::MAX as usize);
    assert_eq!(
        maximum_replacements.max_replacement_bytes(),
        isize::MAX as usize
    );
}

#[test]
fn usize_max_is_rejected_during_profile_construction() {
    assert!(matches!(
        SemanticLimits::builder().max_operations(usize::MAX),
        Err(Error::InvalidLimit {
            kind: SemanticLimitKind::Operations,
            value: usize::MAX,
            maximum: SemanticLimits::MAX_OPERATIONS,
        })
    ));

    assert!(matches!(
        SemanticLimits::builder().max_replacement_bytes(usize::MAX),
        Err(Error::InvalidLimit {
            kind: SemanticLimitKind::ReplacementBytes,
            value: usize::MAX,
            maximum: SemanticLimits::MAX_REPLACEMENT_BYTES,
        })
    ));
}

#[test]
fn limit_refusal_keeps_source_immutable_and_edit_reusable() {
    let source = Package::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(1)));
    assert_eq!(edit.options().semantic_limits_value(), semantic_limits(1));

    edit.replace_text(TextSpan::new(0, 5).unwrap(), "First")
        .unwrap();
    assert!(matches!(
        edit.replace_text(TextSpan::new(6, 10).unwrap(), "Second"),
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
        edit.replace_text(TextSpan::new(0, 5).unwrap(), "Other"),
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

#[test]
fn zero_replacement_profile_allows_empty_no_ops_but_refuses_nonempty_payload() {
    let source = Package::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = replacement_limits(8, 0);

    let empty = source
        .edit_with_options(EditOptions::new().semantic_limits(limits))
        .commit()
        .unwrap();
    assert!(empty.package().same_snapshot(&source));

    let mut empty_replacement =
        source.edit_with_options(EditOptions::new().semantic_limits(limits));
    empty_replacement
        .replace_text(TextSpan::new(0, 0).unwrap(), "")
        .unwrap();
    assert!(
        empty_replacement
            .commit()
            .unwrap()
            .package()
            .same_snapshot(&source)
    );

    let mut refused = source.edit_with_options(EditOptions::new().semantic_limits(limits));
    assert!(matches!(
        refused.replace_text(TextSpan::new(0, 5).unwrap(), "Other"),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            observed: 5,
            maximum: 0,
        })
    ));
    assert_eq!(refused.operation_count(), 0);
    assert!(refused.commit().unwrap().package().same_snapshot(&source));
}

#[test]
fn scalar_replacements_are_aggregated_at_exact_and_above_boundaries() {
    let source = Package::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let first = TextSpan::new(0, 5).unwrap();
    let second = TextSpan::new(6, 10).unwrap();

    let mut exact =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 10)));
    {
        let borrowed = String::from("First");
        exact.replace_text(first, borrowed.as_str()).unwrap();
    }
    exact.replace_text(second, String::from("Right")).unwrap();
    assert_eq!(exact.commit().unwrap().package().text(), "First Right");

    let mut above =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 9)));
    above.replace_text(first, "First").unwrap();
    assert!(matches!(
        above.replace_text(second, "Right"),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            observed: 10,
            maximum: 9,
        })
    ));
    assert_eq!(above.operation_count(), 1);
    above.replace_text(second, "R").unwrap();
    assert_eq!(above.commit().unwrap().package().text(), "First R");
}

#[test]
fn paragraph_batches_admit_exact_payload_and_reject_aggregate_overflow_atomically() {
    let source = Package::parse(r"{\rtf1\ansi One\par Two\par Three}").unwrap();
    let exact_batch = [
        ParagraphTextReplacement::new(Position::new(0), "Uno"),
        ParagraphTextReplacement::new(Position::new(1), "Deux"),
    ];
    let mut exact =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 7)));
    exact.replace_body_paragraph_texts(&exact_batch).unwrap();
    assert_eq!(exact.commit().unwrap().package().text(), "Uno\nDeux\nThree");

    let overflow_batch = [
        ParagraphTextReplacement::new(Position::new(0), "Uno"),
        ParagraphTextReplacement::new(Position::new(1), "Deux"),
    ];
    let mut overflow =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 6)));
    assert!(matches!(
        overflow.replace_body_paragraph_texts(&overflow_batch),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            observed: 7,
            maximum: 6,
        })
    ));
    assert_eq!(overflow.operation_count(), 0);
    overflow
        .replace_paragraph_text(Position::new(0), "A")
        .unwrap();
    assert_eq!(overflow.commit().unwrap().package().text(), "A\nTwo\nThree");
}

#[test]
fn paragraph_structural_scalar_uses_the_same_replacement_budget() {
    let source = Package::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let mut exact =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 3)));
    exact
        .replace_paragraph_text(Position::new(0), "A\nB")
        .unwrap();
    assert_eq!(exact.commit().unwrap().package().text(), "A\nB");

    let mut overflow =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 2)));
    assert!(matches!(
        overflow.replace_paragraph_text(Position::new(0), "A\nB"),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            observed: 3,
            maximum: 2,
        })
    ));
    assert_eq!(overflow.operation_count(), 0);
}

#[test]
fn retained_destinations_share_the_replacement_budget() {
    let source = Package::parse(
        r#"{\rtf1\ansi{\header Head}\pard Body\par{\*\atnid AM}{\*\atnauthor Ada}\chatn{\*\annotation Note}\trowd\cellx1000\cellx2000\intbl A\cell B\cell\row}"#,
    )
    .unwrap();
    let limits = replacement_limits(8, 3);
    let header = HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0);
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(limits));
    {
        let cell = String::from("X");
        edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), cell.as_str())
            .unwrap();
    }
    edit.set_header_footer_text(header, String::from("Y"))
        .unwrap();
    edit.set_annotation_text(0, "Z").unwrap();
    assert_eq!(edit.operation_count(), 3);
    let committed = edit.commit().unwrap();
    assert_eq!(
        committed.package().tables()[0].rows()[0].cells()[0].text(),
        "X"
    );
    assert_eq!(
        committed.package().sections()[0].headers_footers[0].text(),
        "Y"
    );
    assert_eq!(committed.package().annotations()[0].text, "Z");
}

#[test]
fn smart_tag_metadata_is_admitted_and_reopened_under_the_same_profile() {
    let source =
        Package::parse(r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}}"#)
            .unwrap();
    let mut tag = source.smart_tags()[0].clone();
    tag.name = Cow::Borrowed("time");
    tag.attributes = vec![
        SmartTagAttribute::new(Some(1), Cow::Borrowed("format"), Cow::Borrowed("short")).unwrap(),
    ];
    let mut refused =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 15)));
    assert!(matches!(
        refused.set_smart_tag(0, tag.clone()),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            observed: 16,
            maximum: 15,
        })
    ));
    assert_eq!(refused.operation_count(), 0);
    let mut edit =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 16)));
    edit.set_smart_tag(0, tag).unwrap();
    let committed = edit.commit().unwrap();
    let reopened = Package::from_bytes(&committed.package().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.smart_tags()[0].name, "time");
    assert_eq!(reopened.smart_tags()[0].attributes[0].name, "format");
}

#[test]
fn retained_source_limit_can_reject_output_after_semantic_payload_is_admitted() {
    let source_bytes = br"{\rtf1\ansi Alpha}";
    let source_limits = SourceLimits::builder()
        .max_input_bytes(source_bytes.len())
        .unwrap()
        .build()
        .unwrap();
    let source = Package::from_bytes_with_options(
        source_bytes,
        OpenOptions::new().source_limits(source_limits),
    )
    .unwrap();
    let mut edit =
        source.edit_with_options(EditOptions::new().semantic_limits(replacement_limits(8, 128)));
    edit.replace_body_text("\\\\\\\\\\").unwrap();
    assert!(matches!(edit.commit(), Err(Error::InputTooLarge { .. })));
}

#[test]
fn composition_preserves_operation_and_replacement_counts() {
    let source = Package::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let limits = litchi_rtf::edit::CompositionLimits::new(4, 8, 16, 8);
    let mut left = source.edit();
    left.replace_text(TextSpan::new(0, 5).unwrap(), "Left")
        .unwrap();
    left.replace_text(TextSpan::new(6, 10).unwrap(), "Right")
        .unwrap();
    let mut composition = litchi_rtf::edit::Composition::new(&source, limits);
    composition
        .join(left.into_sub_edit("left", limits).unwrap())
        .unwrap();
    let edit = composition.into_edit();
    let selected = edit.options().semantic_limits_value();
    assert_eq!(selected.max_operations(), 2);
    assert_eq!(
        selected.max_replacement_bytes(),
        source.source_limits().max_input_bytes()
    );
    assert_eq!(edit.commit().unwrap().package().text(), "Left Right");
}
