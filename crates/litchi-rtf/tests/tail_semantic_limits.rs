#![allow(clippy::unwrap_used, reason = "test assertions panic on failure")]

use litchi_rtf::{
    Package,
    tail_append::{
        EditOptions, Error, LimitKind, SemanticLimits, SemanticLimitsBuilder, TailSelector,
    },
};

#[test]
fn every_dimension_checks_zero_maximum_and_overflow() {
    type Setter = fn(SemanticLimitsBuilder, usize) -> Result<SemanticLimitsBuilder, Error>;
    type Getter = fn(SemanticLimits) -> usize;
    let dimensions: [(LimitKind, Setter, Getter, usize); 6] = [
        (
            LimitKind::Paragraphs,
            SemanticLimitsBuilder::max_paragraphs,
            SemanticLimits::max_paragraphs,
            256,
        ),
        (
            LimitKind::Runs,
            SemanticLimitsBuilder::max_runs,
            SemanticLimits::max_runs,
            4096,
        ),
        (
            LimitKind::InputBytes,
            SemanticLimitsBuilder::max_input_bytes,
            SemanticLimits::max_input_bytes,
            256 * 1024,
        ),
        (
            LimitKind::InsertedBytes,
            SemanticLimitsBuilder::max_inserted_bytes,
            SemanticLimits::max_inserted_bytes,
            512 * 1024,
        ),
        (
            LimitKind::OutputBytes,
            SemanticLimitsBuilder::max_output_bytes,
            SemanticLimits::max_output_bytes,
            64 * 1024 * 1024,
        ),
        (
            LimitKind::PatchBytes,
            SemanticLimitsBuilder::max_patch_bytes,
            SemanticLimits::max_patch_bytes,
            2 * 1024 * 1024,
        ),
    ];
    for (kind, setter, getter, default) in dimensions {
        assert_eq!(getter(SemanticLimits::default()), default);
        for boundary in [0, SemanticLimits::MAXIMUM] {
            let profile = setter(SemanticLimits::builder(), boundary)
                .unwrap()
                .build()
                .unwrap();
            assert_eq!(getter(profile), boundary);
            assert_eq!(profile.into_builder().build().unwrap(), profile);
        }
        assert!(matches!(setter(SemanticLimits::builder(), usize::MAX),
            Err(Error::InvalidLimit { kind: actual, value: usize::MAX, maximum })
                if actual == kind && maximum == isize::MAX as usize));
    }
}

#[test]
fn retained_edit_uses_checked_options_and_is_reusable_after_refusal() {
    let source = Package::parse(r"{\rtf1 A}").unwrap();
    let limits = SemanticLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let options = EditOptions::new().semantic_limits(limits);
    let mut edit = source.edit_tail_append_with_options(TailSelector::Body, options);
    assert_eq!(edit.options(), options);
    assert!(matches!(
        edit.append_text_paragraphs(&["AB"]),
        Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: 2,
            maximum: 1,
        })
    ));
    assert_eq!(edit.paragraph_count(), 0);
    edit.append_text_paragraphs(&["B"]).unwrap();
    assert_eq!(edit.commit().unwrap().package().text(), "A\nB\n");
    assert_eq!(source.text(), "A");
}

#[test]
fn zero_paragraph_budget_denies_payload_but_preserves_empty_edit() {
    let source = Package::parse(r"{\rtf1 A}").unwrap();
    let limits = SemanticLimits::builder()
        .max_paragraphs(0)
        .unwrap()
        .build()
        .unwrap();
    let mut edit = source.edit_tail_append_with_options(
        TailSelector::Body,
        EditOptions::new().semantic_limits(limits),
    );
    assert!(matches!(
        edit.append_text_paragraphs(&["B"]),
        Err(Error::Limit {
            kind: LimitKind::Paragraphs,
            observed: 1,
            maximum: 0,
        })
    ));
    assert!(edit.commit().unwrap().package().same_snapshot(&source));
}
