//! Public and deferred-reader coverage for the inert DOC `Selsf` cache.

#![allow(
    clippy::expect_used,
    clippy::uninlined_format_args,
    reason = "bounded binary fixtures fail fast with contextual assertions"
)]

use litchi_core::Position;
use litchi_doc::body_text::{Error, Projection, Refusal, Snapshot};
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::tracked_revision::Limits;
use litchi_doc::writer::Writer;
use litchi_doc::{Package, SavedSelection, SelectionGeometry, SelectionStyle};
use litchi_ole_common::object::{Editor as PackageEditor, Targets};
use litchi_ole_common::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value,
    document_summary::DIGITAL_SIGNATURE,
};
use std::io::Cursor;

fn record(flags: u16) -> Vec<u8> {
    let mut data = vec![0; 36];
    data[0..2].copy_from_slice(&flags.to_le_bytes());
    data[2] = 0x81;
    data[4..8].copy_from_slice(&4i32.to_le_bytes());
    data[8..12].copy_from_slice(&8i32.to_le_bytes());
    data[12..16].copy_from_slice(&[0xA5, 0x5A, 0xC3, 0x3C]);
    data[20..24].copy_from_slice(&4i32.to_le_bytes());
    data[24..26].copy_from_slice(&(SelectionStyle::Character as u16).to_le_bytes());
    data[26..28].copy_from_slice(&[0xD7, 0x7D]);
    data[32..34].copy_from_slice(&(-100i16).to_le_bytes());
    data[34..36].copy_from_slice(&100i16.to_le_bytes());
    data
}

fn set_cps(data: &mut [u8], cp_first: i32, cp_lim: i32, cp_anchor: i32) {
    data[4..8].copy_from_slice(&cp_first.to_le_bytes());
    data[8..12].copy_from_slice(&cp_lim.to_le_bytes());
    data[20..24].copy_from_slice(&cp_anchor.to_le_bytes());
}

fn signature_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    let put = |bytes: &mut [u8], offset: usize, value: u32| {
        bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    };
    put(&mut bytes, 0, 48);
    put(&mut bytes, 4, 8);
    put(&mut bytes, 8, 3);
    put(&mut bytes, 12, 44);
    put(&mut bytes, 16, 4);
    put(&mut bytes, 20, 48);
    put(&mut bytes, 24, 0);
    put(&mut bytes, 28, 52);
    put(&mut bytes, 32, 0xA5A5_5A5A);
    put(&mut bytes, 36, 0);
    put(&mut bytes, 40, 54);
    bytes[44..47].copy_from_slice(&[1, 2, 3]);
    bytes[47] = 0xCC;
    bytes[48..52].copy_from_slice(&[4, 5, 6, 7]);
    bytes
}

fn signed_document_with_selsf(source: &[u8]) -> Vec<u8> {
    let base = document_with_selsf(source, 36);
    let mut package =
        PackageEditor::open(base, Targets::default(), Limits::default()).expect("package");
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
        .add(DIGITAL_SIGNATURE, Value::Blob(signature_blob()))
        .expect("signature property");
    let stream = Stream::new(section).to_bytes().expect("signature stream");
    package
        .add_stream(
            vec!["\u{0005}DocumentSummaryInformation".to_string()],
            stream,
        )
        .expect("signature stream insertion");
    package.finish().expect("signed package finish")
}

fn document_with_selsf(source: &[u8], length: u32) -> Vec<u8> {
    document_with_selsf_body("Body", source, length)
}

fn document_with_selsf_body(body: &str, source: &[u8], length: u32) -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph(body).expect("fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("fixture DOC");
    let mut package =
        PackageEditor::open(output.into_inner(), Targets::default(), Limits::default())
            .expect("package");

    let word_path = ["WordDocument".to_string()];
    let mut word = package.stream(&word_path).expect("WordDocument").to_vec();
    let fib = FileInformationBlock::parse(&word).expect("FIB");
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    let offset = u32::try_from(table.len()).expect("table offset");
    table.extend_from_slice(source);
    package
        .put_stream(&table_path, table)
        .expect("Selsf table stream");

    let pair = 154 + 30 * 8;
    word[pair..pair + 4].copy_from_slice(&offset.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&length.to_le_bytes());
    package.put_stream(&word_path, word).expect("Selsf pointer");
    package
        .add_stream(
            vec!["OpaqueVendorData".to_string()],
            b"untouched auxiliary stream".to_vec(),
        )
        .expect("opaque stream");
    package.finish().expect("package finish")
}

#[test]
fn parses_fixed_record_and_preserves_ignored_fields() {
    let mut data = record(1 << 13);
    data[16..20].copy_from_slice(&0x0002_0001u32.to_le_bytes());
    let selection = SavedSelection::parse_bytes(&data).expect("valid Selsf");
    assert_eq!(selection.bytes(), data.as_slice());
    assert_eq!(selection.cp_first(), 4);
    assert_eq!(selection.cp_lim(), 8);
    assert_eq!(selection.cp_anchor(), 4);
    assert_eq!(selection.direction_flags(), 0x81);
    assert!(selection.is_forward());
    assert!(selection.prefix_w2007());
    assert_eq!(
        selection.geometry(),
        SelectionGeometry::Block { first: 1, limit: 2 }
    );

    let mut table = record(1 << 11);
    table[16..20].copy_from_slice(&0x0040_0000u32.to_le_bytes());
    table[32..34].copy_from_slice(&(-31_680i16).to_le_bytes());
    table[34..36].copy_from_slice(&31_680i16.to_le_bytes());
    assert_eq!(
        SavedSelection::parse_bytes(&table)
            .expect("whole-row Selsf")
            .geometry(),
        SelectionGeometry::Table {
            first: 0,
            limit: 64
        }
    );
}

#[test]
fn rejects_invalid_domains_and_cross_field_constraints() {
    let mut reserved = record(1 << 14);
    assert!(SavedSelection::parse_bytes(&reserved).is_err());
    reserved[0] = 0;
    reserved[3] = 2;
    assert!(SavedSelection::parse_bytes(&reserved).is_err());

    let mut shape = record(1 << 8);
    shape[3] = 2;
    let shape_selection = SavedSelection::parse_bytes(&shape).expect("shape Selsf");
    assert!(!shape_selection.is_insertion_end());
    shape[3] = 1;
    assert!(SavedSelection::parse_bytes(&shape).is_ok());

    let mut insertion = record(1 << 15);
    insertion[8..12].copy_from_slice(&9i32.to_le_bytes());
    assert!(SavedSelection::parse_bytes(&insertion).is_err());

    for flags in [
        1 << 2 | 1 << 11,
        1 << 4,
        1 << 8 | 1 << 12,
        1 << 10,
        1 << 10 | 1 << 11,
    ] {
        assert!(SavedSelection::parse_bytes(&record(flags)).is_err());
    }

    let mut invalid_style = record(0);
    invalid_style[24..26].copy_from_slice(&6u16.to_le_bytes());
    assert!(SavedSelection::parse_bytes(&invalid_style).is_err());
}

#[test]
fn fib_range_requires_exact_size_and_deferred_errors_are_cached() {
    let source = record(0);
    let offset = 8u32;
    let mut fib_data = vec![0; 154 + 31 * 8];
    fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    fib_data[152..154].copy_from_slice(&31u16.to_le_bytes());
    let pointer = 154 + 30 * 8;
    fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
    fib_data[pointer + 4..pointer + 8].copy_from_slice(&36u32.to_le_bytes());
    let fib = FileInformationBlock::parse(&fib_data).expect("FIB");
    let mut table = vec![0xCC; 80];
    table[offset as usize..offset as usize + 36].copy_from_slice(&source);
    assert_eq!(
        SavedSelection::parse(&fib, &table)
            .expect("selected Selsf")
            .expect("record")
            .bytes(),
        source.as_slice()
    );

    let bytes = document_with_selsf(&source[..35], 35);

    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package
        .document()
        .expect("document open despite optional error");
    let first = document
        .saved_selection()
        .expect_err("malformed optional Selsf");
    let second = document
        .saved_selection()
        .expect_err("cached malformed optional Selsf");
    assert_eq!(first.to_string(), second.to_string());
}

#[test]
fn document_accessor_returns_selection_and_enforces_main_text_bound() {
    let mut valid = record(0);
    set_cps(&mut valid, 0, 1, 0);
    let bytes = document_with_selsf(&valid, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let selection = document
        .saved_selection()
        .expect("valid Selsf")
        .expect("selected record");
    assert_eq!(selection.cp_first(), 0);
    assert_eq!(selection.cp_lim(), 1);
    assert!(selection.cp_lim() <= document.fib().get_main_doc_range().1);

    let mut outside = record(0);
    set_cps(&mut outside, i32::MAX, i32::MAX, i32::MAX);
    let bytes = document_with_selsf(&outside, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let first = document
        .saved_selection()
        .expect_err("cpFirst beyond ccpText");
    let second = document
        .saved_selection()
        .expect_err("cpFirst error is cached");
    assert!(first.to_string().contains("cpFirst"));
    assert_eq!(first.to_string(), second.to_string());

    let mut outside_lim = record(0);
    set_cps(&mut outside_lim, 0, i32::MAX, 0);
    let bytes = document_with_selsf(&outside_lim, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let first = document
        .saved_selection()
        .expect_err("cpLim beyond ccpText");
    let second = document
        .saved_selection()
        .expect_err("cpLim error is cached");
    assert!(first.to_string().contains("cpLim"));
    assert_eq!(first.to_string(), second.to_string());
}

#[test]
fn body_saved_selection_edit_is_source_bound_reversible_and_atomic() {
    let mut source_record = record(0);
    set_cps(&mut source_record, 0, 1, 0);
    let source_bytes = document_with_selsf(&source_record, 36);
    let source = Snapshot::open(source_bytes.clone(), Limits::default()).expect("source DOC");
    let selection = source
        .saved_selection()
        .expect("saved-selection read")
        .expect("saved-selection record");

    let mut no_op = source.edit().expect("no-op edit");
    let no_op_patch = selection
        .transaction()
        .commit()
        .expect("no-op Selsf transaction");
    assert!(
        !no_op
            .apply_saved_selection_patch(no_op_patch.patch())
            .expect("exact no-op patch")
    );
    let no_op_commit = no_op.commit().expect("no-op commit");
    assert!(!no_op_commit.changed());
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &no_op_commit.snapshot().bytes_shared()
    ));

    let mut transaction = selection.transaction();
    transaction
        .set_range(0, 1)
        .expect("valid range")
        .set_cp_anchor(1)
        .expect("valid anchor");
    let selection_commit = transaction.commit().expect("selection commit");
    assert!(selection_commit.changed());
    let patch = selection_commit.patch().clone();

    let mut edit = source.edit().expect("source edit");
    assert!(
        edit.edit_saved_selection(|transaction| transaction.set_cp_anchor(1).map(|_| ()))
            .expect("publish Selsf patch")
    );
    let commit = edit.commit().expect("publish commit");
    assert!(commit.changed());
    let published_selection = commit
        .snapshot()
        .saved_selection()
        .expect("published Selsf read")
        .expect("published Selsf");
    assert_eq!(published_selection.cp_anchor(), 1);
    let published_package = PackageEditor::open(
        commit.snapshot().finish(),
        Targets::default(),
        Limits::default(),
    )
    .expect("published package");
    assert_eq!(
        published_package
            .stream(&["OpaqueVendorData".to_string()])
            .expect("opaque stream"),
        b"untouched auxiliary stream"
    );

    let applied = commit.patch().apply(&source).expect("forward body patch");
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    let reverted = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse body patch");
    assert_eq!(reverted.bytes(), source.bytes());
    assert!(std::sync::Arc::ptr_eq(
        &reverted.bytes_shared(),
        &source.bytes_shared()
    ));

    let reopened = Snapshot::open(source_bytes, Limits::default()).expect("reopened source");
    let mut reopened_edit = reopened.edit().expect("reopened edit");
    assert!(
        reopened_edit
            .apply_saved_selection_patch(&patch)
            .expect("identical reopened source accepts patch")
    );

    let foreign = Snapshot::open(
        document_with_selsf_body("different", &source_record, 36),
        Limits::default(),
    )
    .expect("foreign source");
    let mut foreign_edit = foreign.edit().expect("foreign edit");
    assert!(matches!(
        foreign_edit.apply_saved_selection_patch(&patch),
        Err(Error::Conflict)
    ));
    let foreign_commit = foreign_edit.commit().expect("foreign rollback commit");
    assert_eq!(foreign_commit.snapshot().bytes(), foreign.bytes());

    let mut malformed_record = source_record.clone();
    malformed_record[4..8].copy_from_slice(&2i32.to_le_bytes());
    malformed_record[8..12].copy_from_slice(&1i32.to_le_bytes());
    let malformed = Snapshot::open(
        document_with_selsf(&malformed_record, 36),
        Limits::default(),
    )
    .expect("malformed optional selection does not block DOC edit admission");
    let mut malformed_edit = malformed.edit().expect("malformed-source edit");
    assert!(matches!(
        malformed_edit.apply_saved_selection_patch(&patch),
        Err(Error::Invalid(_))
    ));
    let malformed_commit = malformed_edit.commit().expect("malformed rollback commit");
    assert_eq!(malformed_commit.snapshot().bytes(), malformed.bytes());
}

#[test]
fn body_saved_selection_bounds_and_length_changes_are_checked_before_publish() {
    let mut outside = record(0);
    set_cps(&mut outside, i32::MAX, i32::MAX, i32::MAX);
    let invalid = Snapshot::open(document_with_selsf(&outside, 36), Limits::default())
        .expect("optional malformed Selsf does not block DOC admission");
    let first = invalid
        .saved_selection()
        .expect_err("cpFirst beyond ccpText");
    assert!(first.to_string().contains("cpFirst"));

    let mut valid = record(0);
    set_cps(&mut valid, 4, 4, 4);
    let source =
        Snapshot::open(document_with_selsf(&valid, 36), Limits::default()).expect("valid source");
    let before = source.bytes().to_vec();
    let selection = source
        .saved_selection()
        .expect("valid saved-selection read")
        .expect("saved-selection record");
    let max_cp = i32::MAX as u32;
    let mut invalid_replacement = selection.transaction();
    invalid_replacement
        .set_cp_anchor(max_cp)
        .expect("structurally valid replacement anchor")
        .set_range(max_cp, max_cp)
        .expect("structurally valid replacement range");
    let invalid_patch = invalid_replacement.commit().expect("invalid-context patch");
    let mut invalid_edit = source.edit().expect("invalid-context edit");
    assert!(matches!(
        invalid_edit.apply_saved_selection_patch(invalid_patch.patch()),
        Err(Error::Invalid(_))
    ));
    let invalid_commit = invalid_edit.commit().expect("invalid patch remains atomic");
    assert_eq!(invalid_commit.snapshot().bytes(), before.as_slice());

    let mut edit = source.edit().expect("body edit");
    edit.replace_paragraph(Position::new(0), "Longer body")
        .expect("boundary Selsf CPs remap with the paragraph");
    let commit = edit.commit().expect("remapped Selsf commit");
    let remapped = commit
        .snapshot()
        .saved_selection()
        .expect("remapped saved-selection read")
        .expect("remapped saved-selection record");
    assert_eq!(remapped.cp_first(), 11);
    assert_eq!(remapped.cp_lim(), 11);
    assert_eq!(remapped.cp_anchor(), 11);
    assert!(remapped.cp_lim() <= 12);
    let reverted = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse remapped text patch");
    assert_eq!(reverted.bytes(), before.as_slice());

    let durable_limits = litchi_core::PatchLimits::new(
        litchi_core::BlobLimits::new(0, 0, 0),
        128 * 1024,
        16,
        8,
        16 * 1024,
        64 * 1024,
    );
    let durable = commit
        .patch()
        .to_durable(durable_limits)
        .expect("deterministic text patch retains Selsf remap");
    let replayed = source
        .apply_durable(&durable)
        .expect("durable text replay remaps Selsf");
    assert_eq!(replayed.bytes(), commit.snapshot().bytes());
    let replayed_reverted = replayed
        .apply_durable(&durable.inverse())
        .expect("durable inverse remaps Selsf back");
    assert_eq!(
        replayed_reverted
            .paragraphs(Projection::All)
            .expect("durable inverse paragraph read")[0]
            .text(),
        "Body"
    );
    let reverted_selection = replayed_reverted
        .saved_selection()
        .expect("durable inverse Selsf read")
        .expect("durable inverse Selsf record");
    assert_eq!(reverted_selection.cp_first(), 4);
    assert_eq!(reverted_selection.cp_lim(), 4);
    assert_eq!(reverted_selection.cp_anchor(), 4);

    let mut ambiguous = record(0);
    set_cps(&mut ambiguous, 1, 3, 1);
    let ambiguous = Snapshot::open(document_with_selsf(&ambiguous, 36), Limits::default())
        .expect("ambiguous source");
    let ambiguous_before = ambiguous.bytes().to_vec();
    let mut ambiguous_edit = ambiguous.edit().expect("ambiguous body edit");
    assert!(matches!(
        ambiguous_edit.replace_paragraph(Position::new(0), "Longer body"),
        Err(Error::Refused(Refusal::PositionDependency {
            fib_index: 30
        }))
    ));
    let ambiguous_commit = ambiguous_edit
        .commit()
        .expect("ambiguous edit remains atomic");
    assert_eq!(
        ambiguous_commit.snapshot().bytes(),
        ambiguous_before.as_slice()
    );
}

#[test]
fn saved_selection_changes_are_refused_by_body_three_way_planning() {
    let mut source_record = record(0);
    set_cps(&mut source_record, 0, 1, 0);
    let source = Snapshot::open(document_with_selsf(&source_record, 36), Limits::default())
        .expect("source DOC");

    let mut left = source.edit().expect("left edit");
    left.edit_saved_selection(|transaction| transaction.set_cp_anchor(1).map(|_| ()))
        .expect("left Selsf edit");
    let left = left.commit().expect("left commit");

    let mut right_text = source.edit().expect("right text edit");
    right_text
        .replace_paragraph(Position::new(0), "Boby")
        .expect("same-length text edit");
    let right_text = right_text.commit().expect("right text commit");
    assert!(matches!(
        source.plan_three_way(left.patch(), right_text.patch()),
        Err(Error::Refused(Refusal::AuxiliaryThreeWayMergeUnsupported))
    ));

    let mut right_selection = source.edit().expect("right Selsf edit");
    right_selection
        .edit_saved_selection(|transaction| transaction.set_forward(false).map(|_| ()))
        .expect("right Selsf edit");
    let right_selection = right_selection.commit().expect("right Selsf commit");
    assert!(matches!(
        source.plan_three_way(left.patch(), right_selection.patch()),
        Err(Error::Refused(Refusal::AuxiliaryThreeWayMergeUnsupported))
    ));
}

#[test]
fn body_saved_selection_edit_refuses_signed_changed_source_but_allows_noop() {
    let mut source_record = record(0);
    set_cps(&mut source_record, 0, 1, 0);
    let source = Snapshot::open(
        signed_document_with_selsf(&source_record),
        Limits::default(),
    )
    .expect("signed source");
    let selection = source
        .saved_selection()
        .expect("saved-selection read")
        .expect("saved-selection record");

    let mut no_op = source.edit().expect("signed no-op edit");
    let no_op_patch = selection
        .transaction()
        .commit()
        .expect("no-op Selsf transaction");
    assert!(
        !no_op
            .apply_saved_selection_patch(no_op_patch.patch())
            .expect("signed exact no-op")
    );
    let no_op_commit = no_op.commit().expect("signed no-op commit");
    assert_eq!(no_op_commit.snapshot().bytes(), source.bytes());

    let mut transaction = selection.transaction();
    transaction.set_cp_anchor(1).expect("valid anchor");
    let patch = transaction
        .commit()
        .expect("selection commit")
        .patch()
        .clone();
    let mut changed = source.edit().expect("signed changed edit");
    assert!(matches!(
        changed.apply_saved_selection_patch(&patch),
        Err(Error::Refused(Refusal::SignedSource))
    ));
    let unchanged = changed.commit().expect("refused edit remains atomic");
    assert_eq!(unchanged.snapshot().bytes(), source.bytes());
}
