#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Document,
    edit::{Error, Limits, MAX_PICTURE_REMOVAL_OPERATIONS},
};
use serde_json::Value;

const FIRST: &str = r"{\pict\pngblip\picw1\pich1 89504E47 0d0A1a0A
 0102aBcd}";
const SECOND: &str = r"{\pict\jpegblip\picw2\pich2 FFD8 0102 FFD9}";
const THIRD: &str = r"{\pict\pngblip 89504e470d0a1a0a DEADBEEF}";

fn source() -> String {
    format!(r"{{\rtf1\ansi Before{FIRST}Middle{SECOND}Later{THIRD}After}}")
}

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        max_operations,
        8,
        512 * 1024,
        1024 * 1024,
    )
}

#[test]
fn removes_one_exact_group_without_reserializing_surrounding_bytes() {
    let input = source();
    let document = Document::parse(&input).unwrap();
    let retained_first = document.pictures()[0].clone();
    let retained_third = document.pictures()[2].clone();

    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(
        commit.snapshot().to_bytes().unwrap(),
        input.replace(SECOND, "").as_bytes()
    );
    assert_eq!(
        commit.snapshot().pictures(),
        &[retained_first, retained_third]
    );
    assert_eq!(commit.snapshot().text(), document.text());
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
}

#[test]
fn source_relative_batch_is_atomic_and_exactly_reversible() {
    let input = source();
    let document = Document::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_pictures(&[0, 2]).unwrap();
    let commit = edit.commit().unwrap();
    let expected = input.replace(FIRST, "").replace(THIRD, "");
    assert_eq!(commit.snapshot().to_bytes().unwrap(), expected.as_bytes());
    assert_eq!(commit.snapshot().pictures().len(), 1);

    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .to_bytes()
            .unwrap(),
        input.as_bytes()
    );

    let limits = durable_limits(2);
    let durable = commit.patch().to_durable(limits).unwrap();
    let encoded = durable.to_deterministic_json().unwrap();
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &encoded, limits,
        )
        .unwrap();
    let removed = document.apply_durable(&decoded).unwrap();
    assert_eq!(removed.to_bytes().unwrap(), expected.as_bytes());
    let restored = removed.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), input.as_bytes());
}

#[test]
fn adjacent_batch_restores_original_group_order_at_one_collapsed_offset() {
    let input = source();
    let document = Document::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_pictures(&[0, 1, 2]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.snapshot().pictures().is_empty());

    let durable = commit.patch().to_durable(durable_limits(3)).unwrap();
    let restored = commit.snapshot().apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), input.as_bytes());
}

#[test]
fn rejects_empty_unordered_duplicate_out_of_range_and_bounded_batches_atomically() {
    let document = Document::parse(&source()).unwrap();
    assert!(matches!(
        document.edit().remove_pictures(&[]),
        Err(Error::EmptyPictureRemovalBatch)
    ));
    assert!(matches!(
        document.edit().remove_pictures(&[2, 1]),
        Err(Error::PictureRemovalBatchOutOfOrder {
            previous: 2,
            incoming: 1
        })
    ));
    assert!(matches!(
        document.edit().remove_pictures(&[1, 1]),
        Err(Error::PictureRemovalBatchOutOfOrder {
            previous: 1,
            incoming: 1
        })
    ));
    let mut out_of_range = document.edit();
    assert!(matches!(
        out_of_range.remove_pictures(&[0, 3]),
        Err(Error::PictureOutOfRange {
            position: 3,
            count: 3
        })
    ));
    assert_eq!(out_of_range.operation_count(), 0);

    let mut limited = document.edit_with_limits(Limits::new(1));
    assert!(matches!(
        limited.remove_pictures(&[0, 1]),
        Err(Error::OperationLimit {
            observed: 2,
            limit: 1
        })
    ));
    assert_eq!(limited.operation_count(), 0);

    let too_many = (0..=MAX_PICTURE_REMOVAL_OPERATIONS).collect::<Vec<_>>();
    assert!(matches!(
        document.edit().remove_pictures(&too_many),
        Err(Error::OperationLimit {
            observed: 65,
            limit: 64
        })
    ));
}

#[test]
fn refuses_nested_mixed_dependent_unknown_and_protected_picture_sources() {
    let nested = Document::parse(&format!(r"{{\rtf1{{\*\shppict{FIRST}}}}}")).unwrap();
    assert!(matches!(
        nested.edit().remove_picture(0),
        Err(Error::UnsupportedSource(_))
    ));

    let field = Document::parse(&format!(
        r"{{\rtf1{{\field{{\*\fldinst INCLUDEPICTURE x}}{{\fldrslt{FIRST}}}}}}}"
    ))
    .unwrap();
    assert!(matches!(
        field.edit().remove_picture(0),
        Err(Error::UnsupportedSource(_))
    ));

    let unknown = Document::parse(&format!(r"{{\rtf1\future42{FIRST}}}")).unwrap();
    assert!(matches!(
        unknown.edit().remove_picture(0),
        Err(Error::UnsupportedSource(_))
    ));

    let protected = Document::parse(&format!(r"{{\rtf1\allprot\enforceprot1{FIRST}}}")).unwrap();
    let mut edit = protected.edit();
    edit.remove_picture(0).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::ProtectedDocument { .. })
    ));
}

#[test]
fn durable_removal_is_artifact_stale_checked_and_rejects_forged_group_preconditions() {
    let input = source();
    let document = Document::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();
    let limits = durable_limits(1);
    let durable = commit.patch().to_durable(limits).unwrap();

    let stale = Document::parse(input.replace("Before", "Changed").as_str()).unwrap();
    assert!(matches!(
        stale.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        stale.apply_durable(&durable.inverse()),
        Err(Error::PatchConflict)
    ));

    let mut forged = durable.operations()[0].clone();
    forged
        .preconditions
        .insert("group_sha256".to_string(), Value::String("00".repeat(32)));
    let malicious = litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [litchi_core::patch::ReversibleOperation::new(
            forged,
            durable.inverse().operations()[0].clone(),
        )],
        litchi_core::patch::BlobBundle::new(limits.blobs()),
        litchi_core::patch::BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        document.apply_durable(&malicious),
        Err(Error::StalePrecondition("picture group differs"))
    ));
}

#[test]
fn durable_inverse_refuses_a_foreign_artifact_even_with_the_same_visible_text() {
    let input = source();
    let document = Document::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();

    let foreign_bytes = String::from_utf8(commit.snapshot().to_bytes().unwrap())
        .unwrap()
        .replace("89504E47", "89504e47");
    let foreign = Document::parse(&foreign_bytes).unwrap();
    assert_eq!(foreign.text(), commit.snapshot().text());
    assert!(matches!(
        foreign.apply_durable(&durable.inverse()),
        Err(Error::PatchConflict)
    ));
}
