#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "native SST regression fixtures fail fast by design"
)]

use super::*;
use crate::cell_values::{Reference, ResourceChange, SST, Selector, Snapshot, Value};
use crate::records::{
    Encoding, PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedStringFormatRun,
    SharedStringProperties, SharedStringTable,
};
use litchi_biff::{Record as Frame, Records};
use litchi_cfb::{OleFile, OleWriter};
use std::io::Cursor;

const EXT_SST: u16 = 0x00ff;

fn package(strings: &[&str]) -> Vec<u8> {
    let mut writer = crate::Builder::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    for (row, text) in strings.iter().enumerate() {
        writer
            .write_string(sheet, u32::try_from(row).unwrap(), 0, text)
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn sst_records(workbook: &[u8]) -> Vec<litchi_biff::RecordRef<'_>> {
    let mut records = Vec::new();
    let mut in_sst = false;
    for item in Records::new(workbook, litchi_biff::Limits::default()).expect("valid BIFF limits") {
        let record = item.unwrap();
        if !in_sst {
            if record.kind().get() == SST {
                in_sst = true;
                records.push(record);
            }
        } else if record.kind().get() == CONTINUE {
            records.push(record);
        } else {
            break;
        }
    }
    records
}

fn sst_table(snapshot: &Snapshot) -> SharedStringTable {
    SharedStringTable::parse_from_records(
        &sst_records(snapshot.workbook_stream()),
        &Encoding::from_codepage(1252).unwrap(),
    )
    .unwrap()
}

fn sst_counts(workbook: &[u8]) -> (u32, u32) {
    let records = sst_records(workbook);
    let record = records.first().unwrap();
    let payload = record.payload();
    (
        u32::from_le_bytes(payload[0..4].try_into().unwrap()),
        u32::from_le_bytes(payload[4..8].try_into().unwrap()),
    )
}

fn label_indexes(workbook: &[u8]) -> Vec<u32> {
    raw_records(workbook)
        .unwrap()
        .into_iter()
        .filter(|record| record.kind == LABEL_SST)
        .map(|record| {
            u32::from_le_bytes(
                workbook[record.start + 10..record.start + 14]
                    .try_into()
                    .unwrap(),
            )
        })
        .collect()
}

fn replace_sst_family(package: &[u8], replacement: &[u8]) -> Vec<u8> {
    let mut ole =
        OleFile::open(Cursor::new(package), litchi_cfb::OleFileLimits::default()).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let records = raw_records(&workbook).unwrap();
    let sst_index = records
        .iter()
        .position(|record| record.kind == SST)
        .unwrap();
    let mut old_end = records[sst_index].end;
    for record in &records[sst_index + 1..] {
        if record.kind == CONTINUE || record.kind == EXT_SST {
            old_end = record.end;
        } else {
            break;
        }
    }
    let mut stream = workbook;
    replace_range_and_adjust_bounds(&mut stream, records[sst_index].start, old_end, replacement)
        .unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn rich_phonetic_properties() -> (SharedStringProperties, Vec<u8>) {
    let raw_extension = vec![
        0x34, 0x12, // reserved marker
        0x01, 0x00, // deliberately stale inner byte length
        0x07, 0x00, // font
        0x0a, 0xa0, // Hiragana + Center, with producer option bits
        0x01, 0x00, // one PhRun
        0x01, 0x00, // one phonetic character
        0x01, 0x00, // repeated count
        0x68, 0x30, // と
        0x00, 0x00, 0x00, 0x00, 0x01, 0x00, // PhRun
        0xaa, 0xbb, // opaque producer data covered by cbExtRst
    ];
    let properties = SharedStringProperties {
        formatting_runs: vec![SharedStringFormatRun {
            character_index: 0,
            font_index: 1,
        }],
        phonetic: Some(PhoneticString {
            font_index: 7,
            phonetic_type: PhoneticType::Hiragana,
            alignment: PhoneticAlignment::Center,
            text: "と".into(),
            runs: vec![PhoneticRun {
                phonetic_text_index: 0,
                base_text_index: 0,
                base_text_length: 1,
            }],
            extra_data: vec![0xaa, 0xbb],
        }),
    };
    (properties, raw_extension)
}

fn custom_rich_source() -> (Vec<u8>, Vec<u8>) {
    let (properties, raw_extension) = rich_phonetic_properties();
    let entries = vec![
        SstEntry {
            text: "rich".into(),
            properties: Some(Box::new(properties)),
            source_index: None,
            raw: Some(RawSstEntry {
                prefix: vec![4, 0, 0x0c, 1, 0, 24, 0, 0, 0],
                characters: vec![RawSstCharacters {
                    wide: false,
                    bytes: b"rich".to_vec(),
                }],
                formatting_runs: vec![0, 0, 1, 0],
                extension: Some(raw_extension.clone()),
            }),
        },
        SstEntry {
            text: "old".into(),
            properties: None,
            source_index: None,
            raw: None,
        },
    ];
    let encoded = encode_sst_records(&entries, 2).unwrap();
    let mut replacement = Vec::new();
    for record in encoded.records {
        replacement.extend_from_slice(&record);
    }
    (
        replace_sst_family(&package(&["rich", "old"]), &replacement),
        raw_extension,
    )
}

#[test]
fn replacing_the_last_sst_reference_prunes_the_old_unique_entry() {
    let source = Snapshot::from_bytes(
        package(&["old"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(0, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let target = commit.package();
    assert_eq!(sst_counts(target.workbook_stream()), (1, 1));
    assert_eq!(sst_table(target).strings, ["new"]);
    assert_eq!(label_indexes(target.workbook_stream()), [0]);
    assert_eq!(
        target
            .worksheet(Selector::Name("Sheet1"))
            .unwrap()
            .unwrap()
            .cell(Reference::new(0, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("new".into())
    );
}

#[test]
fn clearing_the_last_sst_reference_writes_zero_counts() {
    let source = Snapshot::from_bytes(
        package(&["old"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.remove_cell(Selector::Name("Sheet1"), Reference::new(0, 0).unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    let target = commit.package();
    assert_eq!(sst_counts(target.workbook_stream()), (0, 0));
    assert!(sst_table(target).strings.is_empty());
    assert!(label_indexes(target.workbook_stream()).is_empty());
    assert!(
        target
            .worksheet(Selector::Name("Sheet1"))
            .unwrap()
            .unwrap()
            .cell(Reference::new(0, 0).unwrap())
            .unwrap()
            .is_none()
    );
}

#[test]
fn shared_references_keep_the_old_unique_entry_and_remap_labels() {
    let source = Snapshot::from_bytes(
        package(&["old", "old"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(0, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let target = edit.commit().unwrap().package().clone();
    assert_eq!(sst_counts(target.workbook_stream()), (2, 2));
    assert_eq!(sst_table(&target).strings, ["old", "new"]);
    assert_eq!(label_indexes(target.workbook_stream()), [1, 0]);
}

#[test]
fn rich_phonetic_properties_and_opaque_ext_rst_survive_compaction() {
    let (bytes, raw_extension) = custom_rich_source();
    let source = Snapshot::from_bytes(bytes, litchi_ole_common::object::Limits::default()).unwrap();
    assert_eq!(
        source_sst_extensions(&source).unwrap()[0],
        Some(raw_extension.clone())
    );
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(1, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let target = edit.commit().unwrap().package().clone();
    let table = sst_table(&target);
    assert_eq!(table.strings, ["rich", "new"]);
    assert_eq!(table.properties[0], sst_table(&source).properties[0]);
    assert_eq!(
        source_sst_extensions(&target).unwrap()[0],
        Some(raw_extension)
    );
}

#[test]
fn pruning_the_last_rich_reference_is_refused_until_semantic_inverse_can_restore_it() {
    let (bytes, _raw_extension) = custom_rich_source();
    let source = Snapshot::from_bytes(bytes, litchi_ole_common::object::Limits::default()).unwrap();
    let plain_removal = effective_sst_entries(
        &source,
        &[ResourceChange::SharedString {
            text: "rich".into(),
            insert: false,
        }],
    );
    assert!(plain_removal.is_err());
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(0, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let error = edit.commit().unwrap_err();
    assert!(error.to_string().contains("semantic inverse metadata"));
}

#[test]
fn non_tail_sst_resource_removal_is_refused_before_label_remapping() {
    let source = Snapshot::from_bytes(
        package(&["first", "second"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let result = effective_sst_entries(
        &source,
        &[ResourceChange::SharedString {
            text: "first".into(),
            insert: false,
        }],
    );
    assert!(result.is_err());
}

#[test]
fn continued_sst_entries_and_ext_sst_buckets_are_rebuilt_with_true_offsets() {
    let long = "x".repeat(10_000);
    let entries = vec![SstEntry {
        text: long.clone(),
        properties: None,
        source_index: None,
        raw: None,
    }];
    let encoded = encode_sst_records(&entries, 1).unwrap();
    assert!(encoded.records.len() > 1);
    let frames: Vec<_> = encoded
        .records
        .iter()
        .map(|record| Frame::open(record.clone(), litchi_biff::Limits::default()).unwrap())
        .collect();
    let refs: Vec<_> = frames.iter().map(Frame::as_ref).collect();
    let table =
        SharedStringTable::parse_from_records(&refs, &Encoding::from_codepage(1252).unwrap())
            .unwrap();
    assert_eq!(table.strings, [long]);
    assert_eq!(encoded.locations.len(), 1);
    assert!(encoded.locations[0].1 >= 4);
    let index = encoded.ext_sst_index(100, 1).unwrap();
    assert_eq!(index.buckets().len(), 1);
    assert_eq!(
        index.buckets()[0].stream_position(),
        100 + encoded.locations[0].0 as u32
    );

    let many_entries: Vec<_> = (0..9)
        .map(|index| SstEntry {
            text: format!("entry-{index}"),
            properties: None,
            source_index: None,
            raw: None,
        })
        .collect();
    let many = encode_sst_records(&many_entries, 9).unwrap();
    let many_index = many.ext_sst_index(100, 9).unwrap();
    assert_eq!(many_index.buckets().len(), 2);
    assert_eq!(
        many_index.buckets()[1].stream_position(),
        100 + many.locations[8].0 as u32
    );

    // The wide entry begins at an odd character-data offset after the
    // preceding compressed entry. Its retained source segment is longer than
    // the odd remaining capacity, so the encoder must round the split down to
    // a whole UTF-16 code unit and continue through another CONTINUE frame.
    let wide_text = "界".repeat(5_000);
    let wide_bytes: Vec<u8> = wide_text
        .encode_utf16()
        .flat_map(|unit| unit.to_le_bytes())
        .collect();
    let odd_entries = vec![
        SstEntry {
            text: "a".into(),
            properties: None,
            source_index: None,
            raw: None,
        },
        SstEntry {
            text: wide_text.clone(),
            properties: None,
            source_index: None,
            raw: Some(RawSstEntry {
                prefix: vec![0x88, 0x13, 0x01],
                characters: vec![RawSstCharacters {
                    wide: true,
                    bytes: wide_bytes,
                }],
                formatting_runs: Vec::new(),
                extension: None,
            }),
        },
    ];
    let odd = encode_sst_records(&odd_entries, 2).unwrap();
    let odd_frames: Vec<_> = odd
        .records
        .iter()
        .map(|record| Frame::open(record.clone(), litchi_biff::Limits::default()).unwrap())
        .collect();
    let odd_refs: Vec<_> = odd_frames.iter().map(Frame::as_ref).collect();
    let odd_table =
        SharedStringTable::parse_from_records(&odd_refs, &Encoding::from_codepage(1252).unwrap())
            .unwrap();
    assert_eq!(odd_table.strings, vec!["a".to_string(), wide_text]);
}

#[test]
fn wide_continuation_after_a_changed_preceding_cell_remains_lossless() {
    let long_wide = "界".repeat(5_000);
    let source = Snapshot::from_bytes(
        package(&["a", &long_wide]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(0, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let target = edit.commit().unwrap().package().clone();
    assert_eq!(sst_counts(target.workbook_stream()), (2, 2));
    assert_eq!(
        sst_table(&target).strings,
        vec![long_wide.clone(), "new".to_string()]
    );
    assert!(sst_records(target.workbook_stream()).len() > 1);
}

#[test]
fn exact_and_semantic_inverse_restore_a_pruned_sst_tail() {
    let source = Snapshot::from_bytes(
        package(&["old"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_value(
        Selector::Name("Sheet1"),
        Reference::new(0, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let exact = commit.patch().inverse().apply(commit.package()).unwrap();
    assert_eq!(exact.bytes(), source.bytes());
    let semantic = commit.patch().semantic();
    let restored = semantic.inverse().apply(commit.package()).unwrap();
    assert_eq!(sst_table(restored.package()).strings, ["old"]);
    assert_eq!(
        restored
            .package()
            .worksheet(Selector::Name("Sheet1"))
            .unwrap()
            .unwrap()
            .cell(Reference::new(0, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("old".into())
    );
}

/// A source whose SST of `count` distinct strings carries an ExtSST index,
/// as Excel writes it (the builder writes none).
fn package_with_ext_sst(count: usize) -> (Vec<u8>, Vec<String>) {
    let texts: Vec<String> = (0..count).map(|index| format!("s{index:05}")).collect();
    let borrowed: Vec<&str> = texts.iter().map(String::as_str).collect();
    let plain = package(&borrowed);
    let mut ole = OleFile::open(
        Cursor::new(plain.as_slice()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let sst_start = raw_records(&workbook)
        .unwrap()
        .iter()
        .find(|record| record.kind == SST)
        .unwrap()
        .start;
    let entries: Vec<_> = texts
        .iter()
        .map(|text| SstEntry {
            text: text.clone(),
            properties: None,
            source_index: None,
            raw: None,
        })
        .collect();
    let unique = u32::try_from(count).unwrap();
    let encoded = encode_sst_records(&entries, unique).unwrap();
    let mut replacement = Vec::new();
    for record in &encoded.records {
        replacement.extend_from_slice(record);
    }
    replacement.extend_from_slice(
        &encoded
            .ext_sst_index(sst_start, unique)
            .unwrap()
            .to_record_bytes()
            .unwrap(),
    );
    (replace_sst_family(&plain, &replacement), texts)
}

fn ext_sst(workbook: &[u8], unique: u32) -> crate::shared_string_index::SharedStringIndex {
    let record = raw_records(workbook)
        .unwrap()
        .into_iter()
        .find(|record| record.kind == EXT_SST)
        .expect("ExtSST record");
    crate::shared_string_index::SharedStringIndex::parse_payload(
        unique,
        &workbook[record.start + 4..record.end],
    )
    .unwrap()
}

#[test]
fn new_text_across_an_ext_sst_bucket_size_boundary_commits_through_compaction() {
    // 1,023 strings use 8-string buckets; the 1,024th needs 9.  Commit
    // rebuilds the SST and its ExtSST, so staging no longer refuses the
    // insertion with the tail-append geometry check (record 0359fa).
    let (bytes, texts) = package_with_ext_sst(1_023);
    let source = Snapshot::from_bytes(bytes, litchi_ole_common::object::Limits::default()).unwrap();
    assert_eq!(
        ext_sst(source.workbook_stream(), 1_023).strings_per_bucket(),
        8
    );
    let mut edit = source.edit();
    edit.insert_cell(
        Selector::Name("Sheet1"),
        Reference::new(1_023, 0).unwrap(),
        Value::Text("new".into()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let target = commit.package();
    assert_eq!(sst_counts(target.workbook_stream()), (1_024, 1_024));
    let mut expected = texts;
    expected.push("new".into());
    assert_eq!(sst_table(target).strings, expected);
    let index = ext_sst(target.workbook_stream(), 1_024);
    assert_eq!(index.strings_per_bucket(), 9);
    assert_eq!(index.buckets().len(), 114);
    assert_eq!(
        target
            .worksheet(Selector::Name("Sheet1"))
            .unwrap()
            .unwrap()
            .cell(Reference::new(1_023, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("new".into())
    );
}

#[test]
fn shared_string_membership_matches_the_effective_table() {
    // More strings than the scan limit build the hash index; duplicates and
    // staged removals exercise the occurrence count.
    let mut texts: Vec<String> = (0..80).map(|index| format!("t{index}")).collect();
    texts.push("t3".into());
    texts.push("t3".into());
    let borrowed: Vec<&str> = texts.iter().map(String::as_str).collect();
    let source = Snapshot::from_bytes(
        package(&borrowed),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.resource_changes.push(ResourceChange::SharedString {
        text: "t3".into(),
        insert: false,
    });
    edit.resource_changes.push(ResourceChange::SharedString {
        text: "t7".into(),
        insert: false,
    });
    edit.resource_changes.push(ResourceChange::SharedString {
        text: "added".into(),
        insert: true,
    });
    let effective = edit.effective_shared_strings();
    for probe in ["t0", "t3", "t7", "t79", "t80", "added", "absent", ""] {
        assert_eq!(
            edit.has_shared_string(probe),
            effective.iter().any(|text| text == probe),
            "{probe}"
        );
    }
    edit.resource_changes.push(ResourceChange::SharedString {
        text: "t3".into(),
        insert: false,
    });
    edit.resource_changes.push(ResourceChange::SharedString {
        text: "t3".into(),
        insert: false,
    });
    assert!(
        !edit
            .effective_shared_strings()
            .iter()
            .any(|text| text == "t3")
    );
    assert!(!edit.has_shared_string("t3"));
}

#[test]
fn a_refused_text_insertion_leaves_no_shared_string_resource() {
    let source = Snapshot::from_bytes(
        package(&["a", "b"]),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.insert_rows(Selector::Name("Sheet1"), 5, 1).unwrap();
    let resources = edit.resource_changes.len();
    let error = edit
        .insert_cell(
            Selector::Name("Sheet1"),
            Reference::new(9, 0).unwrap(),
            Value::Text("fresh".into()),
        )
        .unwrap_err();
    assert!(
        error.to_string().contains("overlapping structural"),
        "{error}"
    );
    assert_eq!(edit.resource_changes.len(), resources);
    assert!(!edit.has_shared_string("fresh"));
}
