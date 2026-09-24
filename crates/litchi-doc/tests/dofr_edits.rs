//! Bounded, source-checked `RgDofr` edit regressions.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::expect_used,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::uninlined_format_args,
    reason = "integration fixtures fail fast with contextual assertions"
)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::Position;
use litchi_doc::body_text::{Error, Refusal, Snapshot, TransactionLimits};
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::tracked_revision::{
    Limits, RevisionEditor, RevisionKind, RevisionMetadata, Snapshot as RevisionSnapshot,
};
use litchi_doc::writer::{CharacterFormatting, ParagraphFormatting, Writer};
use litchi_ole_common::object::{Editor as PackageEditor, Targets};
use litchi_ole_common::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value,
    document_summary::DIGITAL_SIGNATURE,
};
use std::io::Cursor;

fn base_doc() -> Vec<u8> {
    base_doc_with_text("alpha")
}

fn base_doc_with_text(text: &str) -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![(text.to_string(), CharacterFormatting::default())],
            ParagraphFormatting::default(),
        )
        .expect("fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("fixture DOC");
    output.into_inner()
}

fn doc_with_protected_storage(marker: &str) -> Vec<u8> {
    let mut source = OleFile::open(Cursor::new(base_doc())).expect("base CFB");
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let bytes = source.open_stream(&refs).expect("base stream");
        writer
            .create_stream(&refs, &bytes)
            .expect("copy base stream");
    }
    writer.create_storage(&[marker]).expect("protected storage");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("protected CFB");
    output.into_inner()
}

fn record(kind: u32, payload: &[u8]) -> Vec<u8> {
    let size = u32::try_from(8 + payload.len()).expect("record size");
    let mut bytes = Vec::with_capacity(usize::try_from(size).expect("record capacity"));
    bytes.extend_from_slice(&size.to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
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

fn word_signature_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    bytes[0..2].copy_from_slice(&27u16.to_le_bytes());
    bytes[2..6].copy_from_slice(&45u32.to_le_bytes());
    bytes[6..10].copy_from_slice(&8u32.to_le_bytes());
    bytes[10..14].copy_from_slice(&2u32.to_le_bytes());
    bytes[14..18].copy_from_slice(&44u32.to_le_bytes());
    bytes[18..22].copy_from_slice(&3u32.to_le_bytes());
    bytes[22..26].copy_from_slice(&46u32.to_le_bytes());
    bytes[30..34].copy_from_slice(&49u32.to_le_bytes());
    bytes[42..46].copy_from_slice(&51u32.to_le_bytes());
    bytes[46..48].copy_from_slice(&[9, 8]);
    bytes[48..51].copy_from_slice(&[7, 6, 5]);
    bytes[55] = 0xEE;
    bytes
}

fn xst(value: &str) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut bytes = Vec::with_capacity(2 + units.len() * 2);
    bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
    bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
    bytes
}

fn stw_user(signature_name: &str) -> Vec<u8> {
    stw_user_with_names(&[signature_name, "Other"])
}

fn stw_user_with_names(names: &[&str]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0xFFFFu16.to_le_bytes());
    bytes.extend_from_slice(
        &u16::try_from(names.len())
            .expect("StwUser fixture name count")
            .to_le_bytes(),
    );
    bytes.extend_from_slice(&4u16.to_le_bytes());
    for name in names {
        bytes.extend(xst(name));
        bytes.extend_from_slice(&0u32.to_le_bytes());
    }
    for name in names {
        if ["Sign", "SigAgile", "SigV3"].contains(name) {
            bytes.extend(word_signature_blob());
        } else {
            bytes.extend(xst("plain"));
        }
    }
    bytes
}

fn doc_with_dofr(frame_value: u32, signed: bool) -> Vec<u8> {
    doc_with_dofr_options(
        frame_value,
        signed.then_some("\u{0005}DocumentSummaryInformation"),
        None,
    )
}

fn doc_with_dofr_options(
    frame_value: u32,
    signature_stream_name: Option<&str>,
    stw_user_signature_name: Option<&str>,
) -> Vec<u8> {
    doc_with_dofr_and_stw_user(
        frame_value,
        base_doc(),
        signature_stream_name,
        stw_user_signature_name.map(stw_user),
    )
}

fn doc_with_dofr_and_stw_user(
    frame_value: u32,
    base: Vec<u8>,
    signature_stream_name: Option<&str>,
    stw_user_data: Option<Vec<u8>>,
) -> Vec<u8> {
    let mut package =
        PackageEditor::open(base, Targets::default(), Limits::default()).expect("package");
    let word_path = ["WordDocument".to_string()];
    let mut word = package.stream(&word_path).expect("WordDocument").to_vec();
    let fib = FileInformationBlock::parse(&word).expect("FIB");
    assert!(fib.table_pointer_count().expect("FIB pair count") > 99);
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();

    let mut frame_payload = vec![0u8; 36];
    frame_payload[12..16].copy_from_slice(&frame_value.to_le_bytes());
    let mut dofr = record(0, &[]);
    dofr.extend_from_slice(&record(1, &frame_payload));
    let dofr_offset = u32::try_from(table.len()).expect("table offset");
    table.extend_from_slice(&dofr);
    let pair = 154usize + 99 * 8;
    word[pair..pair + 4].copy_from_slice(&dofr_offset.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(
        &u32::try_from(dofr.len())
            .expect("table length")
            .to_le_bytes(),
    );

    if let Some(stw_user) = stw_user_data {
        let stw_user_offset = u32::try_from(table.len()).expect("StwUser offset");
        table.extend_from_slice(&stw_user);
        let pair = 154usize + 60 * 8;
        word[pair..pair + 4].copy_from_slice(&stw_user_offset.to_le_bytes());
        word[pair + 4..pair + 8].copy_from_slice(
            &u32::try_from(stw_user.len())
                .expect("StwUser length")
                .to_le_bytes(),
        );
    }

    package
        .put_stream(&word_path, word)
        .expect("WordDocument update");
    package
        .put_stream(&table_path, table)
        .expect("table update");
    if let Some(signature_name) = signature_stream_name {
        let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
        section.set_page(CodePage::Utf16Le);
        section
            .add(DIGITAL_SIGNATURE, Value::Blob(signature_blob()))
            .expect("signature property");
        let stream = Stream::new(section).to_bytes().expect("signature stream");
        package
            .add_stream(vec![signature_name.to_string()], stream)
            .expect("signature stream insertion");
    }
    package.finish().expect("DOC package finish")
}

fn changed_frame_patch(source: &Snapshot, replacement_value: u32) -> litchi_doc::DofrPatch {
    let records = source
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    let mut replacement = records.get(1).expect("frame record").bytes().to_vec();
    replacement[20..24].copy_from_slice(&replacement_value.to_le_bytes());
    let mut transaction = records.transaction();
    transaction
        .replace_record_bytes(1, &replacement)
        .expect("record replacement");
    transaction.commit().expect("Dofr commit").patch().clone()
}

#[test]
fn signed_dofr_noop_is_allowed_but_changed_source_is_refused() {
    let source = Snapshot::open(doc_with_dofr(2, true), Limits::default()).expect("signed source");
    let records = source
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    let unchanged = records.get(1).expect("frame record").bytes().to_vec();

    let mut no_op = source.edit().expect("no-op edit");
    assert!(
        !no_op
            .replace_dofr_record(1, &unchanged)
            .expect("signed exact no-op")
    );
    assert_eq!(
        no_op.commit().expect("no-op commit").snapshot().bytes(),
        source.bytes()
    );

    let patch = changed_frame_patch(&source, 1);
    let mut changed = source.edit().expect("changed edit");
    assert!(matches!(
        changed.apply_dofr_patch(&patch),
        Err(Error::Refused(Refusal::SignedSource))
    ));
    assert_eq!(
        changed.commit().expect("refused commit").snapshot().bytes(),
        source.bytes()
    );
}

#[test]
fn changed_dofr_patch_requires_the_complete_source_owner() {
    let source = Snapshot::open(doc_with_dofr(2, false), Limits::default()).expect("source");
    let foreign = Snapshot::open_bounded(
        doc_with_dofr_and_stw_user(2, base_doc_with_text("bravo"), None, None),
        Limits::default(),
        TransactionLimits::new(1, 1024, 1024),
    )
    .expect("foreign source");
    let foreign_patch = changed_frame_patch(&source, 1);
    let mut edit = foreign.edit().expect("foreign edit");

    assert!(matches!(
        edit.apply_dofr_patch(&foreign_patch),
        Err(Error::Conflict)
    ));

    // The rejected foreign patch was refused before the operation budget was
    // charged; a patch generated from this exact source still fits.
    let local_patch = changed_frame_patch(&foreign, 1);
    assert!(edit.apply_dofr_patch(&local_patch).expect("local patch"));
}

#[test]
fn reopened_identical_source_with_different_limits_accepts_dofr_patch() {
    let bytes = doc_with_dofr(2, false);
    let source = Snapshot::open(bytes.clone(), Limits::default()).expect("source");
    let retained_limits = Limits {
        max_streams: 32_768,
        ..Limits::default()
    };
    let reopened = Snapshot::open_bounded(
        bytes,
        retained_limits,
        TransactionLimits::new(1, 1024, 1024),
    )
    .expect("independently reopened source");
    assert_eq!(reopened.bytes(), source.bytes());

    let patch = changed_frame_patch(&source, 1);
    let mut edit = reopened.edit().expect("reopened edit");
    assert!(edit.apply_dofr_patch(&patch).expect("same-source patch"));
    let committed = edit.commit().expect("reopened commit");
    let records = committed
        .snapshot()
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    assert_eq!(
        &records.get(1).expect("frame record").bytes()[20..24],
        1u32.to_le_bytes().as_slice()
    );
}

#[test]
fn dofr_three_way_plan_refuses_auxiliary_and_text_merge() {
    let source = Snapshot::open(doc_with_dofr(2, false), Limits::default()).expect("source");
    let dofr_patch = changed_frame_patch(&source, 1);
    let mut dofr_edit = source.edit().expect("Dofr edit");
    assert!(dofr_edit.apply_dofr_patch(&dofr_patch).expect("Dofr patch"));
    let dofr_commit = dofr_edit.commit().expect("Dofr commit");
    assert!(dofr_commit.patch().has_auxiliary_changes());

    let mut text_edit = source.edit().expect("text edit");
    text_edit
        .replace_paragraph(Position::new(0), "bravo")
        .expect("text replacement");
    let text_commit = text_edit.commit().expect("text commit");

    assert!(matches!(
        source.plan_three_way(dofr_commit.patch(), text_commit.patch()),
        Err(Error::Refused(Refusal::AuxiliaryThreeWayMergeUnsupported))
    ));
}

#[test]
fn dofr_three_way_plan_refuses_two_auxiliary_merges() {
    let source = Snapshot::open(doc_with_dofr(2, false), Limits::default()).expect("source");
    let mut left_edit = source.edit().expect("left Dofr edit");
    let left_patch = changed_frame_patch(&source, 1);
    assert!(
        left_edit
            .apply_dofr_patch(&left_patch)
            .expect("left Dofr patch")
    );
    let left = left_edit.commit().expect("left Dofr commit");

    let mut right_edit = source.edit().expect("right Dofr edit");
    let right_patch = changed_frame_patch(&source, 0);
    assert!(
        right_edit
            .apply_dofr_patch(&right_patch)
            .expect("right Dofr patch")
    );
    let right = right_edit.commit().expect("right Dofr commit");

    assert!(matches!(
        source.plan_three_way(left.patch(), right.patch()),
        Err(Error::Refused(Refusal::AuxiliaryThreeWayMergeUnsupported))
    ));
}

#[test]
fn duplicate_unsigned_stw_user_names_are_conservatively_signed() {
    let source = Snapshot::open(
        doc_with_dofr_and_stw_user(
            2,
            base_doc(),
            None,
            Some(stw_user_with_names(&["Other", "Other"])),
        ),
        Limits::default(),
    )
    .expect("duplicate-name source");
    let records = source
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    let unchanged = records.get(1).expect("frame record").bytes().to_vec();
    let mut no_op = source.edit().expect("no-op edit");
    assert!(
        !no_op
            .replace_dofr_record(1, &unchanged)
            .expect("malformed exact no-op")
    );
    assert_eq!(
        no_op.commit().expect("no-op commit").snapshot().bytes(),
        source.bytes()
    );

    let patch = changed_frame_patch(&source, 1);
    let mut changed = source.edit().expect("changed edit");
    assert!(matches!(
        changed.apply_dofr_patch(&patch),
        Err(Error::Refused(Refusal::SignedSource))
    ));
}

#[test]
fn signed_source_refuses_changed_ordinary_body_publication() {
    let source = Snapshot::open(doc_with_dofr(2, true), Limits::default()).expect("signed source");
    let mut edit = source.edit().expect("body edit");
    edit.replace_paragraph(Position::new(0), "bravo")
        .expect("stage body edit");
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::SignedSource))
    ));
}

#[test]
fn signed_lower_editors_refuse_changed_publication_but_allow_exact_noops() {
    let source_bytes = doc_with_dofr(2, true);

    let mut direct = RevisionEditor::open(source_bytes.clone(), Limits::default())
        .expect("direct signed editor");
    direct
        .add_text(
            0,
            "x",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .expect("stage direct edit");
    assert!(direct.finish().is_err());

    let direct_noop =
        RevisionEditor::open(source_bytes.clone(), Limits::default()).expect("direct signed no-op");
    assert_eq!(
        direct_noop.finish().expect("direct exact no-op"),
        source_bytes
    );

    let tracked = RevisionSnapshot::open(source_bytes.clone(), Limits::default())
        .expect("tracked signed snapshot");
    let mut transaction = tracked.edit().expect("tracked edit");
    transaction
        .add_text(
            0,
            "x",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .expect("stage tracked edit");
    assert!(transaction.commit().is_err());

    let tracked_noop =
        RevisionSnapshot::open(source_bytes, Limits::default()).expect("tracked signed no-op");
    let noop_commit = tracked_noop
        .edit()
        .expect("tracked no-op edit")
        .commit()
        .expect("tracked exact no-op");
    assert_eq!(noop_commit.snapshot().bytes(), tracked_noop.bytes());
}

#[test]
fn case_variant_piddsi_stream_is_a_signed_source() {
    let source = Snapshot::open(
        doc_with_dofr_options(2, Some("\u{0005}dOcUmEnTsUmMaRyInFoRmAtIoN"), None),
        Limits::default(),
    )
    .expect("case-variant PIDDSI source");
    let records = source
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    let unchanged = records.get(1).expect("frame record").bytes().to_vec();
    let mut no_op = source.edit().expect("no-op edit");
    assert!(
        !no_op
            .replace_dofr_record(1, &unchanged)
            .expect("case-variant exact no-op")
    );

    let patch = changed_frame_patch(&source, 1);
    let mut changed = source.edit().expect("changed edit");
    assert!(matches!(
        changed.apply_dofr_patch(&patch),
        Err(Error::Refused(Refusal::SignedSource))
    ));
}

#[test]
fn all_word_vba_signature_variable_names_refuse_changed_dofr() {
    for name in ["Sign", "SigAgile", "SigV3"] {
        let source = Snapshot::open(
            doc_with_dofr_options(2, None, Some(name)),
            Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{name} source: {error}"));
        let patch = changed_frame_patch(&source, 1);
        let mut changed = source.edit().expect("changed edit");
        assert!(matches!(
            changed.apply_dofr_patch(&patch),
            Err(Error::Refused(Refusal::SignedSource))
        ));

        let records = source
            .dofr_records()
            .expect("Dofr read")
            .expect("Dofr records");
        let unchanged = records.get(1).expect("frame record").bytes().to_vec();
        let mut no_op = source.edit().expect("no-op edit");
        assert!(
            !no_op
                .replace_dofr_record(1, &unchanged)
                .unwrap_or_else(|error| panic!("{name} exact no-op: {error}"))
        );
    }
}

#[test]
fn recursive_xml_and_legacy_signature_storages_are_refused() {
    for marker in ["_xmlsignatures", "_signatures"] {
        assert!(matches!(
            Snapshot::open(doc_with_protected_storage(marker), Limits::default()),
            Err(Error::Invalid(_))
        ));
    }
}

#[test]
fn dofr_preflight_errors_and_noops_do_not_consume_operation_budget() {
    let source = Snapshot::open_bounded(
        doc_with_dofr(2, false),
        Limits::default(),
        TransactionLimits::new(1, 1024, 1024),
    )
    .expect("bounded source");
    let records = source
        .dofr_records()
        .expect("Dofr read")
        .expect("Dofr records");
    let unchanged = records.get(1).expect("frame record").bytes().to_vec();

    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.replace_dofr_record(99, &unchanged),
        Err(Error::Invalid(_))
    ));
    assert!(matches!(
        edit.replace_dofr_record(1, &[0; 8]),
        Err(Error::Invalid(_))
    ));
    assert!(
        !edit
            .replace_dofr_record(1, &unchanged)
            .expect("exact no-op")
    );
    let patch = changed_frame_patch(&source, 1);
    assert!(edit.apply_dofr_patch(&patch).expect("changed patch"));
    assert!(matches!(
        edit.apply_dofr_patch(&patch),
        Err(Error::Invalid(_))
    ));

    let no_budget = Snapshot::open_bounded(
        doc_with_dofr(2, false),
        Limits::default(),
        TransactionLimits::new(0, 1024, 1024),
    )
    .expect("zero-budget source");
    let mut no_budget_edit = no_budget.edit().expect("zero-budget edit");
    let foreign = Snapshot::open(doc_with_dofr(1, false), Limits::default()).expect("foreign");
    let stale_patch = changed_frame_patch(&foreign, 2);
    assert!(matches!(
        no_budget_edit.apply_dofr_patch(&stale_patch),
        Err(Error::Invalid(_))
    ));
    assert!(
        !no_budget_edit
            .replace_dofr_record(1, &unchanged)
            .expect("zero-budget exact no-op")
    );
}

#[test]
fn dofr_save_reopen_and_inverse_restore_exact_source() {
    let source = Snapshot::open(doc_with_dofr(2, false), Limits::default()).expect("source");
    let patch = changed_frame_patch(&source, 1);
    let mut edit = source.edit().expect("edit");
    assert!(edit.apply_dofr_patch(&patch).expect("apply patch"));
    let commit = edit.commit().expect("commit");
    let reopened = Snapshot::open(commit.snapshot().finish(), Limits::default()).expect("reopen");
    let reopened_records = reopened
        .dofr_records()
        .expect("reopened Dofr")
        .expect("reopened records");
    let expected_frame_value = 1u32.to_le_bytes();
    assert_eq!(
        &reopened_records.get(1).expect("reopened frame").bytes()[20..24],
        expected_frame_value.as_slice()
    );

    let restored = commit.patch().inverse().apply(&reopened).expect("inverse");
    assert_eq!(restored.bytes(), source.bytes());
}
