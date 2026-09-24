#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "integration fixtures use contextual fail-fast assertions"
)]

//! Correctness and ownership coverage for the bounded DOC rendered handoff.
//!
//! These tests intentionally exercise the public body-text transaction. They
//! do not call the common-container handoff directly and do not measure
//! elapsed time or allocator counters.

use litchi_core::Position;
use litchi_core::patch::CompositionLimits;
use litchi_doc::Package;
use litchi_doc::body_text::{
    Edit, Error, Projection, Refusal, Snapshot, SubEditJoinFailure, TransactionLimits,
};
use litchi_doc::tracked_revision::Limits;
use litchi_doc::writer::Writer;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;

mod common;

const HANDOFF_REPLACEMENT: &str = "litchi copy-through baseline replacement text";

fn transaction_policy(
    operations: usize,
    replacement_units: usize,
    total_units: usize,
    retained_render_bytes: usize,
) -> TransactionLimits {
    TransactionLimits::new(operations, replacement_units, total_units)
        .with_max_retained_render_bytes(retained_render_bytes)
}

fn assert_transaction_policy(
    actual: TransactionLimits,
    expected: TransactionLimits,
    context: &str,
) {
    assert_eq!(
        actual.max_operations(),
        expected.max_operations(),
        "{context}: operation ceiling"
    );
    assert_eq!(
        actual.max_replacement_units(),
        expected.max_replacement_units(),
        "{context}: per-replacement ceiling"
    );
    assert_eq!(
        actual.max_total_replacement_units(),
        expected.max_total_replacement_units(),
        "{context}: aggregate replacement ceiling"
    );
    assert_eq!(
        actual.max_retained_render_bytes(),
        expected.max_retained_render_bytes(),
        "{context}: retained-render ceiling"
    );
}

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn fixture_bytes(relative: &str) -> Vec<u8> {
    std::fs::read(fixture(relative)).expect("DOC fixture should be readable")
}

fn snapshot(bytes: &[u8], transaction_limits: TransactionLimits) -> Snapshot {
    Snapshot::open_bounded(bytes.to_vec(), Limits::default(), transaction_limits)
        .expect("DOC fixture should open with bounded transaction limits")
}

fn writer_document(paragraphs: &[&str]) -> Vec<u8> {
    let mut writer = Writer::new();
    for paragraph in paragraphs {
        writer
            .add_paragraph(paragraph)
            .expect("synthetic DOC paragraph should be valid");
    }
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("synthetic DOC should serialize");
    output.into_inner()
}

fn replacement_position(source: &Snapshot) -> Position {
    source
        .paragraphs(Projection::All)
        .expect("DOC paragraphs should be available")
        .into_iter()
        .find(|paragraph| {
            !paragraph.text().is_empty()
                && !paragraph
                    .text()
                    .chars()
                    .any(|character| character.is_control() && character != '\t')
        })
        .expect("source should contain an ordinary body paragraph")
        .position()
}

fn staged_edit(
    bytes: &[u8],
    transaction_limits: TransactionLimits,
    position: Position,
    replacement: &str,
) -> Edit {
    let source = snapshot(bytes, transaction_limits);
    let mut edit = source.edit().expect("body edit should open");
    edit.replace_paragraph(position, replacement)
        .expect("ordinary body replacement should stage");
    edit
}

fn reference_output(bytes: &[u8], position: Position, replacement: &str) -> Vec<u8> {
    let limits = TransactionLimits::default().with_max_retained_render_bytes(0);
    let edit = staged_edit(bytes, limits, position, replacement);
    edit.commit()
        .expect("reference body edit should commit")
        .snapshot()
        .bytes()
        .to_vec()
}

fn sequential_reference_output(bytes: &[u8], position: Position, replacements: &[&str]) -> Vec<u8> {
    let limits = TransactionLimits::default().with_max_retained_render_bytes(0);
    let source = snapshot(bytes, limits);
    let mut edit = source.edit().expect("reference body edit should open");
    for replacement in replacements {
        edit.replace_paragraph(position, replacement)
            .expect("reference body replacement should stage");
    }
    edit.commit()
        .expect("reference body edit should commit")
        .snapshot()
        .bytes()
        .to_vec()
}

fn assert_public_reopens(bytes: &[u8]) {
    let parsed = Snapshot::parse(bytes).expect("strict DOC snapshot should reopen");
    assert_eq!(parsed.bytes(), bytes);
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec()))
        .expect("public DOC package should reopen");
    package
        .document()
        .expect("public DOC reader should validate the candidate");
}

#[test]
fn transaction_policy_intersection_keeps_the_stricter_retention_ceiling() {
    let broad = TransactionLimits::default().with_max_retained_render_bytes(4096);
    let narrow = TransactionLimits::new(3, 32, 64).with_max_retained_render_bytes(128);
    let intersection = broad.intersect(narrow);
    assert_eq!(intersection.max_operations(), 3);
    assert_eq!(intersection.max_replacement_units(), 32);
    assert_eq!(intersection.max_total_replacement_units(), 64);
    assert_eq!(intersection.max_retained_render_bytes(), 128);
}

#[test]
fn equal_byte_lineage_prepared_edits_meet_all_transaction_limits() {
    let bytes = writer_document(&["alpha", "bravo", "charlie"]);
    let broad = transaction_policy(8, 256, 512, 8 * 1024 * 1024);
    let strict = transaction_policy(4, 64, 128, 0);
    let broad_source = snapshot(&bytes, broad);
    let strict_source = snapshot(&bytes, strict);
    let composition_limits = CompositionLimits::new(4, 2, 8, 8);

    let left = broad_source
        .prepare_replace(
            composition_limits,
            "broad-left",
            Position::new(0),
            "alpha policy",
        )
        .expect("broad prepared edit should validate");
    let right = strict_source
        .prepare_replace(
            composition_limits,
            "strict-right",
            Position::new(2),
            "charlie policy",
        )
        .expect("strict prepared edit should validate");

    let mut composition = broad_source.compose(composition_limits);
    composition
        .join(left)
        .expect("broad prepared edit should join");
    composition
        .join(right)
        .expect("equal-byte lineage should ignore distinct transaction policies");
    let commit = composition
        .commit()
        .expect("policy-meeting composition should commit");
    assert_transaction_policy(
        commit.snapshot().transaction_limits(),
        strict,
        "prepared-edit composition",
    );
    assert_public_reopens(commit.snapshot().bytes());
    assert_eq!(
        paragraph_text(commit.snapshot().bytes(), Position::new(0)),
        "alpha policy"
    );
    assert_eq!(
        paragraph_text(commit.snapshot().bytes(), Position::new(2)),
        "charlie policy"
    );

    let donor = Snapshot::parse(&writer_document(&["donor policy", "untouched donor"]))
        .expect("donor should parse");
    let receiver_bytes = writer_document(&["receiver policy", "untouched receiver"]);
    let broad_receiver = snapshot(&receiver_bytes, broad);
    let strict_receiver = snapshot(&receiver_bytes, strict);
    let plan = strict_receiver
        .plan_text_transfer_from(
            &donor,
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
        )
        .expect("strict transfer plan should validate");
    let mut transfer_edit = broad_receiver
        .edit()
        .expect("broad receiver edit should open");
    transfer_edit
        .apply_transfer(&plan)
        .expect("equal-byte receiver lineage should accept transfer");
    let transfer_commit = transfer_edit
        .commit()
        .expect("policy-meeting transfer should commit");
    assert_transaction_policy(
        transfer_commit.snapshot().transaction_limits(),
        strict,
        "transfer policy meet",
    );
    assert_public_reopens(transfer_commit.snapshot().bytes());
    assert_eq!(
        paragraph_text(transfer_commit.snapshot().bytes(), Position::new(0)),
        "donor policy"
    );
}

#[test]
fn rejected_join_and_transfer_preserve_policy_and_retained_state() {
    let bytes = writer_document(&["alpha", "bravo", "charlie"]);
    let broad = transaction_policy(8, 256, 512, 8 * 1024 * 1024);
    let strict = transaction_policy(4, 64, 128, 0);
    let broad_source = snapshot(&bytes, broad);
    let strict_source = snapshot(&bytes, strict);
    let composition_limits = CompositionLimits::new(4, 2, 8, 8);

    let accepted = broad_source
        .prepare_replace(
            composition_limits,
            "accepted",
            Position::new(0),
            "accepted policy",
        )
        .expect("accepted prepared edit should validate");
    let conflicting = strict_source
        .prepare_replace(
            composition_limits,
            "conflicting",
            Position::new(0),
            "conflicting policy",
        )
        .expect("conflicting prepared edit should validate");
    let mut composition = broad_source.compose(composition_limits);
    composition
        .join(accepted)
        .expect("accepted prepared edit should join");
    let join_error = composition
        .join(conflicting)
        .expect_err("overlapping prepared edit should be rejected");
    assert!(matches!(
        join_error.failure(),
        SubEditJoinFailure::Overlap(_)
    ));
    let rejected = join_error.into_rejected();
    assert_eq!(rejected.identifier(), "conflicting");
    let broad_commit = composition
        .commit()
        .expect("rejected join must not poison accepted composition");
    assert_transaction_policy(
        broad_commit.snapshot().transaction_limits(),
        broad,
        "rejected composition",
    );
    assert_eq!(
        paragraph_text(broad_commit.snapshot().bytes(), Position::new(0)),
        "accepted policy"
    );

    let mut recovered = strict_source.compose(composition_limits);
    recovered
        .join(rejected)
        .expect("rejected prepared token should remain recoverable");
    let recovered_commit = recovered
        .commit()
        .expect("recovered rejected edit should commit");
    assert_transaction_policy(
        recovered_commit.snapshot().transaction_limits(),
        strict,
        "recovered prepared token",
    );

    let donor = Snapshot::parse(&writer_document(&["donor transfer", "untouched donor"]))
        .expect("donor should parse");
    let wrong_receiver_bytes = writer_document(&["wrong receiver", "untouched receiver"]);
    let wrong_receiver = snapshot(&wrong_receiver_bytes, strict);
    let wrong_plan = wrong_receiver
        .plan_text_transfer_from(
            &donor,
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
        )
        .expect("wrong receiver transfer plan should validate");

    let mut edit = broad_source.edit().expect("broad edit should open");
    edit.replace_paragraph(Position::new(0), "before rejected transfer")
        .expect("pre-transfer mutation should stage");
    let held_before_rejection = edit.retained_render_bytes();
    assert!(held_before_rejection.is_some());
    let source_policy_before = edit.source().transaction_limits();
    assert!(matches!(
        edit.apply_transfer(&wrong_plan),
        Err(Error::Conflict)
    ));
    assert_eq!(edit.retained_render_bytes(), held_before_rejection);
    assert_transaction_policy(
        edit.source().transaction_limits(),
        source_policy_before,
        "rejected transfer source policy",
    );
    let edit_commit = edit
        .commit()
        .expect("rejected transfer must leave the staged edit committable");
    assert_transaction_policy(
        edit_commit.snapshot().transaction_limits(),
        broad,
        "rejected transfer commit",
    );
    assert_eq!(
        paragraph_text(edit_commit.snapshot().bytes(), Position::new(0)),
        "before rejected transfer"
    );
}

#[test]
fn forward_patch_and_three_way_merge_meet_before_after_policies() {
    let bytes = writer_document(&["alpha", "bravo", "charlie"]);
    let broad = transaction_policy(10, 256, 512, 8 * 1024 * 1024);
    let strict = transaction_policy(4, 64, 128, 0);
    let source = snapshot(&bytes, broad);
    let mut edit = source.edit().expect("broad edit should open");
    edit.replace_paragraph(Position::new(0), "alpha patched")
        .expect("broad patch should stage");
    let patch = edit
        .commit()
        .expect("broad patch should commit")
        .patch()
        .clone();
    assert_transaction_policy(
        patch.before().transaction_limits(),
        broad,
        "patch before policy",
    );
    assert_transaction_policy(
        patch.after().transaction_limits(),
        broad,
        "patch after policy",
    );

    let strict_source = snapshot(&bytes, strict);
    let applied = patch
        .apply(&strict_source)
        .expect("forward patch should accept equal-byte strict source");
    assert_transaction_policy(
        applied.transaction_limits(),
        strict,
        "forward patch target policy",
    );
    assert_eq!(
        paragraph_text(applied.bytes(), Position::new(0)),
        "alpha patched"
    );
    let restored = patch
        .inverse()
        .apply(&applied)
        .expect("inverse patch should accept the strict result");
    assert_transaction_policy(
        restored.transaction_limits(),
        strict,
        "inverse patch target policy",
    );
    assert_eq!(restored.bytes(), bytes.as_slice());

    let left_policy = transaction_policy(8, 80, 160, 512);
    let right_policy = transaction_policy(6, 60, 120, 0);
    let left_source = snapshot(&bytes, left_policy);
    let right_source = snapshot(&bytes, right_policy);
    let mut left_edit = left_source.edit().expect("left edit should open");
    left_edit
        .replace_paragraph(Position::new(0), "alpha left")
        .expect("left edit should stage");
    let left = left_edit.commit().expect("left edit should commit");
    let mut right_edit = right_source.edit().expect("right edit should open");
    right_edit
        .replace_paragraph(Position::new(2), "charlie right")
        .expect("right edit should stage");
    let right = right_edit.commit().expect("right edit should commit");
    assert_transaction_policy(
        left.patch().before().transaction_limits(),
        left_policy,
        "left before policy",
    );
    assert_transaction_policy(
        left.patch().after().transaction_limits(),
        left_policy,
        "left after policy",
    );
    assert_transaction_policy(
        right.patch().before().transaction_limits(),
        right_policy,
        "right before policy",
    );
    assert_transaction_policy(
        right.patch().after().transaction_limits(),
        right_policy,
        "right after policy",
    );

    let merge_source = snapshot(&bytes, broad);
    let expected = transaction_policy(6, 60, 120, 0);
    let merged = merge_source
        .plan_three_way(left.patch(), right.patch())
        .expect("equal-byte differently bounded patches should merge");
    let merged = merged.commit().expect("three-way merge should commit");
    assert_transaction_policy(
        merged.snapshot().transaction_limits(),
        expected,
        "three-way before/after policy meet",
    );
    assert_public_reopens(merged.snapshot().bytes());
    assert_eq!(
        paragraph_text(merged.snapshot().bytes(), Position::new(0)),
        "alpha left"
    );
    assert_eq!(
        paragraph_text(merged.snapshot().bytes(), Position::new(2)),
        "charlie right"
    );
}

fn paragraph_text(bytes: &[u8], position: Position) -> String {
    Snapshot::parse(bytes)
        .expect("DOC output should parse")
        .paragraphs(Projection::All)
        .expect("DOC output paragraphs should be available")
        .into_iter()
        .find(|paragraph| paragraph.position() == position)
        .expect("edited paragraph should remain addressable")
        .text()
        .to_owned()
}

#[test]
fn retention_boundaries_preserve_exact_output_and_capacity() {
    for relative in [
        "test-data/ole/doc/NoHeadFoot.doc",
        "test-data/ole/doc/FloatingPictures.doc",
    ] {
        let bytes = fixture_bytes(relative);
        let source = snapshot(&bytes, TransactionLimits::default());
        let position = Position::new(0);
        let reference = reference_output(&bytes, position, HANDOFF_REPLACEMENT);

        let mut calibration = source.edit().expect("calibration body edit should open");
        calibration
            .replace_paragraph(position, HANDOFF_REPLACEMENT)
            .expect("calibration replacement should stage");
        let capacity = calibration
            .retained_render_bytes()
            .expect("default DOC retention should hold this candidate");
        assert!(capacity > 1, "fixture render should have a useful capacity");

        let cases = [
            ("zero", 0usize),
            ("under", capacity - 1),
            ("exact", capacity),
            ("over", capacity + 1),
        ];
        for (name, ceiling) in cases {
            let limits = TransactionLimits::default().with_max_retained_render_bytes(ceiling);
            assert_eq!(limits.max_retained_render_bytes(), ceiling);
            let edit = staged_edit(&bytes, limits, position, HANDOFF_REPLACEMENT);
            let held = edit.retained_render_bytes();
            match name {
                "zero" | "under" => assert_eq!(held, None, "{relative}: {name}"),
                "exact" => assert_eq!(held, Some(capacity), "{relative}: {name}"),
                "over" => {
                    let held = held.expect("over-ceiling candidate should be retained");
                    assert_eq!(held, capacity, "{relative}: {name}");
                    assert!(held <= ceiling, "{relative}: {name}");
                },
                _ => unreachable!("all retention cases are listed above"),
            }

            let commit = edit
                .commit()
                .expect("retention ceiling must not refuse a supported edit");
            let output = commit.snapshot().bytes();
            assert_eq!(output, reference.as_slice(), "{relative}: {name}");
            assert!(commit.changed());
            assert_public_reopens(output);
            assert_eq!(
                paragraph_text(output, position),
                HANDOFF_REPLACEMENT,
                "{relative}: {name} semantic readback"
            );
        }
    }
}

#[test]
fn release_successive_noop_and_refused_edits_are_atomic() {
    let bytes = fixture_bytes("test-data/ole/doc/NoHeadFoot.doc");
    let position = Position::new(0);
    let first = "first bounded handoff replacement";
    let second = "second bounded handoff replacement with more text";
    let first_reference = reference_output(&bytes, position, first);
    let second_reference = sequential_reference_output(&bytes, position, &[first, second]);

    let mut edit = staged_edit(&bytes, TransactionLimits::default(), position, first);
    let first_held = edit.retained_render_bytes();
    edit.replace_paragraph(position, second)
        .expect("successive replacement should stage");
    let second_held = edit.retained_render_bytes();
    edit.replace_paragraph(position, second)
        .expect("same-value replacement should be a no-op");
    assert_eq!(edit.retained_render_bytes(), second_held);
    edit.release_retained_render();
    assert_eq!(edit.retained_render_bytes(), None);
    let second_commit = edit
        .commit()
        .expect("released successive edit should commit");
    assert_eq!(
        second_commit.snapshot().bytes(),
        second_reference.as_slice()
    );
    assert_public_reopens(second_commit.snapshot().bytes());
    assert_ne!(first_held, None, "the first edit should exercise retention");

    let limited =
        TransactionLimits::new(1, 4096, 4096).with_max_retained_render_bytes(8 * 1024 * 1024);
    let mut refused = staged_edit(&bytes, limited, position, first);
    let held_before_refusal = refused.retained_render_bytes();
    let error = refused
        .replace_paragraph(position, second)
        .expect_err("the second operation must exceed the operation limit");
    assert!(matches!(
        error,
        Error::Refused(Refusal::OperationLimit { .. })
    ));
    assert_eq!(refused.retained_render_bytes(), held_before_refusal);
    let refused_commit = refused
        .commit()
        .expect("a refused later operation must leave the edit committable");
    assert_eq!(
        refused_commit.snapshot().bytes(),
        first_reference.as_slice()
    );

    let source = snapshot(&bytes, TransactionLimits::default());
    let original = source
        .paragraphs(Projection::All)
        .expect("source paragraphs")
        .into_iter()
        .find(|paragraph| paragraph.position() == position)
        .expect("position zero paragraph")
        .text()
        .to_owned();
    let mut noop = source.edit().expect("no-op edit should open");
    noop.replace_paragraph(position, &original)
        .expect("same source text should be a no-op");
    assert_eq!(noop.retained_render_bytes(), None);
    let noop_source = noop
        .commit()
        .expect("no-op edit should commit")
        .into_parts()
        .0;
    assert_eq!(noop_source.bytes(), bytes.as_slice());
    assert!(Arc::ptr_eq(
        &source.bytes_shared(),
        &noop_source.bytes_shared()
    ));
}

#[test]
fn multi_generation_and_writer_producer_candidates_reopen() {
    let generated = writer_document(&["alpha", "bravo 😀", "charlie"]);
    let cases = [
        ("word97", fixture_bytes("test-data/ole/doc/NoHeadFoot.doc")),
        // The Word 2002 producer fixture carries a nonconforming FIB and DOP
        // length that the protection classifier refuses to edit; normalize
        // them as the other edit tests do.
        (
            "word0101",
            common::with_valid_word97_dop(fixture_bytes(
                "test-data/ole/doc/documentProperties.doc",
            )),
        ),
        ("litchi-writer", generated),
    ];

    for (name, bytes) in cases {
        let source = snapshot(&bytes, TransactionLimits::default());
        let position = replacement_position(&source);
        let mut edit = source.edit().expect("producer source should edit");
        edit.replace_paragraph(position, HANDOFF_REPLACEMENT)
            .expect("producer source replacement should stage");
        let commit = edit.commit().expect("producer source should commit");
        let output = commit.snapshot().bytes();
        assert_public_reopens(output);
        assert_eq!(
            paragraph_text(output, position),
            HANDOFF_REPLACEMENT,
            "{name} semantic body readback"
        );
    }
}

#[test]
fn patches_inverse_conflict_and_composition_preserve_handoff_results() {
    let bytes = writer_document(&["alpha", "bravo", "charlie"]);
    let source = snapshot(&bytes, TransactionLimits::default());
    let limits = CompositionLimits::new(4, 2, 8, 8);
    let left = source
        .prepare_replace(limits, "left", Position::new(0), "alpha changed")
        .expect("left replacement should prepare");
    let right = source
        .prepare_replace(limits, "right", Position::new(2), "charlie changed")
        .expect("right replacement should prepare");
    let mut composition = source.compose(limits);
    composition
        .join(left)
        .expect("left replacement should join");
    composition
        .join(right)
        .expect("right replacement should join");
    let composed = composition.commit().expect("composition should commit");

    let mut sequential = source.edit().expect("sequential edit should open");
    sequential
        .replace_paragraph(Position::new(0), "alpha changed")
        .expect("sequential left replacement");
    sequential
        .replace_paragraph(Position::new(2), "charlie changed")
        .expect("sequential right replacement");
    let sequential = sequential.commit().expect("sequential edit should commit");
    assert_eq!(
        composed.snapshot().bytes(),
        sequential.snapshot().bytes(),
        "composition and sequential publication should agree"
    );

    let forward = composed
        .patch()
        .apply(&source)
        .expect("forward patch should apply to its exact source");
    assert_eq!(forward, *composed.snapshot());
    let inverse = composed
        .patch()
        .inverse()
        .apply(composed.snapshot())
        .expect("inverse patch should restore its exact source");
    assert_eq!(inverse.bytes(), source.bytes());

    let wrong = Snapshot::parse(&writer_document(&["different", "bravo", "charlie"]))
        .expect("mismatched source should parse");
    let wrong_bytes = wrong.bytes().to_vec();
    assert!(matches!(
        composed.patch().apply(&wrong),
        Err(Error::Conflict)
    ));
    assert_eq!(wrong.bytes(), wrong_bytes.as_slice());
}

#[test]
fn transfer_after_a_handoff_reopens_and_checks_source_lineage() {
    let donor = Snapshot::parse(&writer_document(&["donor text", "untouched donor"]))
        .expect("donor should parse");
    let receiver_bytes = writer_document(&["receiver text", "untouched receiver"]);
    let receiver = Snapshot::parse(&receiver_bytes).expect("receiver should parse");
    let plan = receiver
        .plan_text_transfer_from(
            &donor,
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
            litchi_doc::body_text::TextTarget::body_paragraph(Position::new(0)),
        )
        .expect("inert text transfer should prepare");

    let mut edit = receiver.edit().expect("receiver edit should open");
    edit.replace_paragraph(Position::new(0), "intermediate handoff")
        .expect("intermediate replacement should stage");
    let held_before_transfer = edit.retained_render_bytes();
    edit.apply_transfer(&plan)
        .expect("transfer should invalidate the earlier staged state");
    let transferred = edit.commit().expect("transferred receiver should commit");
    let transferred_bytes = transferred.snapshot().bytes();
    assert_public_reopens(transferred_bytes);
    assert_eq!(
        paragraph_text(transferred_bytes, Position::new(0)),
        "donor text"
    );
    assert!(held_before_transfer.is_some());

    let expected_source = snapshot(
        &receiver_bytes,
        TransactionLimits::default().with_max_retained_render_bytes(0),
    );
    let mut expected_edit = expected_source.edit().expect("expected receiver edit");
    expected_edit
        .replace_paragraph(Position::new(0), "intermediate handoff")
        .expect("same intermediate edit in the recomputation control");
    expected_edit
        .apply_transfer(&plan)
        .expect("expected transfer should stage");
    let expected = expected_edit
        .commit()
        .expect("expected transfer should commit");
    assert_eq!(transferred_bytes, expected.snapshot().bytes());

    let wrong_receiver = Snapshot::parse(&writer_document(&["wrong receiver"]))
        .expect("wrong receiver should parse");
    let wrong_bytes = wrong_receiver.bytes().to_vec();
    let mut wrong_edit = wrong_receiver.edit().expect("wrong edit should open");
    assert!(matches!(
        wrong_edit.apply_transfer(&plan),
        Err(Error::Conflict)
    ));
    assert_eq!(wrong_edit.source().bytes(), wrong_bytes.as_slice());
}
