#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    EditOptions, Package, SemanticLimitKind, SemanticLimits,
    edit::{Error, MAX_PICTURE_REMOVAL_OPERATIONS},
};
use serde_json::Value;

fn semantic_limits(max_operations: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .build()
        .unwrap()
}

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
    let document = Package::parse(&input).unwrap();
    let retained_first = document.pictures()[0].clone();
    let retained_third = document.pictures()[2].clone();

    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(
        commit.package().to_bytes().unwrap(),
        input.replace(SECOND, "").as_bytes()
    );
    assert_eq!(
        commit.package().pictures(),
        &[retained_first, retained_third]
    );
    assert_eq!(commit.package().text(), document.text());
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
}

#[test]
fn source_relative_batch_is_atomic_and_exactly_reversible() {
    let input = source();
    let document = Package::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_pictures(&[0, 2]).unwrap();
    let commit = edit.commit().unwrap();
    let expected = input.replace(FIRST, "").replace(THIRD, "");
    assert_eq!(commit.package().to_bytes().unwrap(), expected.as_bytes());
    assert_eq!(commit.package().pictures().len(), 1);

    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
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
    let removed = document.apply_durable(&decoded).unwrap().into_package();
    assert_eq!(removed.to_bytes().unwrap(), expected.as_bytes());
    let restored = removed
        .apply_durable(&decoded.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), input.as_bytes());
}

#[test]
fn adjacent_batch_restores_original_group_order_at_one_collapsed_offset() {
    let input = source();
    let document = Package::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_pictures(&[0, 1, 2]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.package().pictures().is_empty());

    let durable = commit.patch().to_durable(durable_limits(3)).unwrap();
    let restored = commit
        .package()
        .apply_durable(&durable.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), input.as_bytes());
}

#[test]
fn rejects_empty_unordered_duplicate_out_of_range_and_bounded_batches_atomically() {
    let document = Package::parse(&source()).unwrap();
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

    let mut limited =
        document.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(1)));
    assert!(matches!(
        limited.remove_pictures(&[0, 1]),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 2,
            maximum: 1
        })
    ));
    assert_eq!(limited.operation_count(), 0);

    let too_many = (0..=MAX_PICTURE_REMOVAL_OPERATIONS).collect::<Vec<_>>();
    assert!(matches!(
        document.edit().remove_pictures(&too_many),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 65,
            maximum: 64
        })
    ));
}

#[test]
fn refuses_nested_mixed_dependent_unknown_and_protected_picture_sources() {
    let nested = Package::parse(&format!(r"{{\rtf1{{\*\shppict{FIRST}}}}}")).unwrap();
    assert!(matches!(
        nested.edit().remove_picture(0),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let field = Package::parse(&format!(
        r"{{\rtf1{{\field{{\*\fldinst INCLUDEPICTURE x}}{{\fldrslt{FIRST}}}}}}}"
    ))
    .unwrap();
    assert!(matches!(
        field.edit().remove_picture(0),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let unknown = Package::parse(&format!(r"{{\rtf1\future42{FIRST}}}")).unwrap();
    assert!(matches!(
        unknown.edit().remove_picture(0),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let protected = Package::parse(&format!(r"{{\rtf1\allprot\enforceprot1{FIRST}}}")).unwrap();
    let mut edit = protected.edit();
    edit.remove_picture(0).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));
}

#[test]
fn durable_removal_is_artifact_stale_checked_and_rejects_forged_group_preconditions() {
    let input = source();
    let document = Package::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();
    let limits = durable_limits(1);
    let durable = commit.patch().to_durable(limits).unwrap();

    let stale = Package::parse(input.replace("Before", "Changed").as_str()).unwrap();
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
    let document = Package::parse(&input).unwrap();
    let mut edit = document.edit();
    edit.remove_picture(1).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();

    let foreign_bytes = String::from_utf8(commit.package().to_bytes().unwrap())
        .unwrap()
        .replace("89504E47", "89504e47");
    let foreign = Package::parse(&foreign_bytes).unwrap();
    assert_eq!(foreign.text(), commit.package().text());
    assert!(matches!(
        foreign.apply_durable(&durable.inverse()),
        Err(Error::PatchConflict)
    ));
}

/// Record 0290 F7: removal moves the whole group's bytes, so a payload that
/// does not match its declared format (LibreOffice `n192129.rtf` stores GIF
/// data under `\pngblip`) or a format payload edits do not support (EMF) is
/// removed exactly, and the durable inverse restores it; payload replacement
/// still refuses such a picture.
#[test]
fn removes_groups_whose_payload_format_payload_edits_refuse() {
    const MISMATCHED: &str =
        r"{\pict \pngblip \picw-64 \pich-1061137057 \picwgoal0 \pichgoal0 474946383961100010003b}";
    const EMF: &str = r"{\pict\emfblip\picw4\pich4 0100000058000000}";
    let input = format!(r"{{\rtf1 Before{MISMATCHED}Middle{EMF}After\par }}");
    let document = Package::parse(&input).unwrap();
    assert_eq!(document.pictures().len(), 2);

    let mut edit = document.edit();
    edit.remove_pictures(&[0, 1]).unwrap();
    let commit = edit.commit().unwrap();
    let expected = input.replace(MISMATCHED, "").replace(EMF, "");
    assert_eq!(commit.package().to_bytes().unwrap(), expected.as_bytes());
    assert!(commit.package().pictures().is_empty());
    assert_eq!(commit.package().text(), document.text());

    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
    let restored = commit
        .package()
        .apply_durable(&durable.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), input.as_bytes());

    let mut payload = document.edit();
    assert!(matches!(
        payload.replace_picture_payload(0, [0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
}
