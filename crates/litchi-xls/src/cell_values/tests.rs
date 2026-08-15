#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test fixtures and assertions intentionally fail fast"
)]

use super::*;
use litchi_cfb::{OleFile, OleWriter, SharedOleFile};
use std::io::Write;

fn package() -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    writer.write_number(sheet, 4, 0, 1.0).unwrap();
    writer.write_number(sheet, 4, 1, 2.0).unwrap();
    writer.write_string(sheet, 6, 0, "alpha").unwrap();
    writer.write_string(sheet, 6, 1, "beta").unwrap();
    writer.write_string(sheet, 6, 2, "alpha").unwrap();
    writer.write_boolean(sheet, 7, 0, true).unwrap();
    writer.write_formula(sheet, 8, 0, "1+1").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package =
        PackageEditor::open(output.into_inner(), Targets::default(), Limits::default()).unwrap();
    package
        .add_stream(vec!["Opaque".to_string()], b"untouched".to_vec())
        .unwrap();
    package.finish().unwrap()
}

fn fixed_numeric_family_package(storage: Storage) -> (Vec<u8>, Reference) {
    let source = package();
    let mut ole = OleFile::open(Cursor::new(source)).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    let parsed = Records::new(&workbook);
    for record in parsed {
        let record = record.unwrap();
        records.push((record.kind().get(), record.payload().to_vec()));
    }

    let number_indices = records
        .iter()
        .enumerate()
        .filter(|(_, (kind, _))| *kind == NUMBER)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let (indices, reference) = match storage {
        Storage::Rk => {
            let index = *number_indices.first().unwrap();
            let payload = &records[index].1;
            (
                vec![index],
                Reference::new(
                    u32::from(u16::from_le_bytes([payload[0], payload[1]])),
                    u32::from(u16::from_le_bytes([payload[2], payload[3]])),
                )
                .unwrap(),
            )
        },
        Storage::MulRk => {
            let pair = number_indices
                .windows(2)
                .find(|pair| {
                    pair[1] == pair[0] + 1
                        && records[pair[0]].1[0..2] == records[pair[1]].1[0..2]
                        && u16::from_le_bytes([records[pair[1]].1[2], records[pair[1]].1[3]])
                            == u16::from_le_bytes([records[pair[0]].1[2], records[pair[0]].1[3]])
                                + 1
                })
                .map(|pair| [pair[0], pair[1]])
                .unwrap();
            let payload = &records[pair[0]].1;
            (
                pair.to_vec(),
                Reference::new(
                    u32::from(u16::from_le_bytes([payload[0], payload[1]])),
                    u32::from(u16::from_le_bytes([payload[2], payload[3]])),
                )
                .unwrap(),
            )
        },
        Storage::Number => panic!("the native package already contains Number records"),
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            panic!("test fixture requests a nonnumeric storage family")
        },
    };
    let mut transformed = Vec::new();
    for (index, (kind, payload)) in records.iter().enumerate() {
        if storage == Storage::MulRk && index == indices[1] {
            continue;
        }
        let mut payload = payload.clone();
        if index == indices[0] {
            if storage == Storage::Rk {
                payload.truncate(10);
                let rk = encode_rk(f64::from_le_bytes(
                    records[index].1[6..14].try_into().unwrap(),
                ))
                .unwrap();
                payload[6..10].copy_from_slice(&rk.to_le_bytes());
                transformed.push((RK, payload));
                continue;
            }
            let first = &records[indices[0]].1;
            let second = &records[indices[1]].1;
            let mut packed = Vec::with_capacity(18);
            packed.extend_from_slice(&first[0..4]);
            packed.extend_from_slice(&first[4..10]);
            packed.extend_from_slice(&second[4..10]);
            packed.extend_from_slice(&u16::from_le_bytes([second[2], second[3]]).to_le_bytes());
            transformed.push((MUL_RK, packed));
            continue;
        }
        transformed.push((*kind, payload));
    }
    let mut rebuilt = Vec::new();
    for (kind, payload) in transformed {
        rebuilt.extend_from_slice(&kind.to_le_bytes());
        rebuilt.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
        rebuilt.extend_from_slice(&payload);
    }
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &rebuilt).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    (output.into_inner(), reference)
}

fn raw_rk_family_package(storage: Storage, raw_values: &[u32]) -> (Vec<u8>, Reference) {
    let (source, reference) = fixed_numeric_family_package(storage);
    let mut ole = OleFile::open(Cursor::new(source)).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut rebuilt = Vec::new();
    let mut replaced = false;
    for record in Records::new(&workbook) {
        let record = record.unwrap();
        let kind = record.kind().get();
        let mut payload = record.payload().to_vec();
        match storage {
            Storage::Rk if kind == RK => {
                assert_eq!(raw_values.len(), 1);
                payload[6..10].copy_from_slice(&raw_values[0].to_le_bytes());
                replaced = true;
            },
            Storage::MulRk if kind == MUL_RK => {
                assert_eq!(raw_values.len(), 2);
                for (index, raw) in raw_values.iter().enumerate() {
                    let offset = 6 + index * 6;
                    payload[offset..offset + 4].copy_from_slice(&raw.to_le_bytes());
                }
                replaced = true;
            },
            _ => {},
        }
        rebuilt.extend_from_slice(&kind.to_le_bytes());
        rebuilt.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
        rebuilt.extend_from_slice(&payload);
    }
    assert!(replaced);
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &rebuilt).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    (output.into_inner(), reference)
}

fn formula_free_package() -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Plain").unwrap();
    writer.write_number(sheet, 1, 0, 1.0).unwrap();
    writer.write_number(sheet, 3, 1, 2.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn formula_package(formula: &str) -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Formula").unwrap();
    writer.write_formula(sheet, 0, 0, formula).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn positioned_formula_package(formula: &str, row: u32, column: u16) -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Formula").unwrap();
    writer.write_formula(sheet, row, column, formula).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn signed_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn macro_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer.create_storage(&["_VBA_PROJECT_CUR"]).unwrap();
    writer.create_storage(&["_VBA_PROJECT_CUR", "VBA"]).unwrap();
    writer
        .create_stream(&["_VBA_PROJECT_CUR", "VBA", "_VBA_PROJECT"], b"project")
        .unwrap();
    writer
        .create_stream(&["_VBA_PROJECT_CUR", "VBA", "dir"], b"dir")
        .unwrap();
    writer
        .create_stream(&["_VBA_PROJECT_CUR", "VBA", "Module1"], b"module")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn empty_macro_storage_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer.create_storage(&["_VBA_PROJECT_CUR"]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn topology_with_clsids_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer.create_storage(&["Metadata"]).unwrap();
    writer
        .create_stream(&["Metadata", "Payload"], b"metadata")
        .unwrap();
    writer.set_root_clsid([0x11; 16]);
    writer.set_storage_clsid(&["Metadata"], [0x22; 16]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn directory_signature(
    bytes: &[u8],
) -> Vec<(u32, String, u8, String, u32, u32, u32, u32, u64, bool)> {
    let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
    let shared = SharedOleFile::open(Arc::new(SnapshotSource::new(
        Arc::clone(&snapshot.inner.bytes),
        snapshot.inner.source_version,
    )))
    .unwrap();
    shared
        .directory_entries()
        .map(|entry| {
            (
                entry.sid,
                entry.name.clone(),
                entry.entry_type,
                entry.clsid.clone(),
                entry.sid_left,
                entry.sid_right,
                entry.sid_child,
                entry.start_sector,
                entry.size,
                entry.is_minifat,
            )
        })
        .collect()
}

fn stream(bytes: &[u8], name: &str) -> Vec<u8> {
    OleFile::open(Cursor::new(bytes))
        .unwrap()
        .open_stream(&[name])
        .unwrap()
}

#[test]
fn edits_only_one_number_field_and_round_trips_patch() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let worksheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(worksheet.position(), 0);
    assert_eq!(worksheet.number(reference).unwrap().unwrap().value(), 4.5);

    let before_workbook = source.workbook_stream().to_vec();
    let mut edit = source.edit();
    edit.set_number(0usize.into(), reference, 9.25).unwrap();
    let commit = edit.commit().unwrap();
    assert!(Arc::ptr_eq(
        &source.inner.shared_strings,
        &commit.snapshot().inner.shared_strings
    ));
    assert!(Arc::ptr_eq(
        &source.inner.shared_string_properties,
        &commit.snapshot().inner.shared_string_properties
    ));
    assert!(Arc::ptr_eq(
        &source.inner.xf_records,
        &commit.snapshot().inner.xf_records
    ));
    assert_eq!(commit.diagnostics().changed_number_fields(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    let after_workbook = commit.snapshot().workbook_stream();
    assert_eq!(before_workbook.len(), after_workbook.len());
    assert_eq!(
        before_workbook
            .iter()
            .zip(after_workbook)
            .filter(|(left, right)| left != right)
            .count(),
        2
    );
    assert_eq!(stream(source.bytes(), "Opaque"), b"untouched");
    assert_eq!(stream(commit.snapshot().bytes(), "Opaque"), b"untouched");
    assert_eq!(
        source
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(reference)
            .unwrap()
            .unwrap()
            .value(),
        4.5
    );

    let applied = commit.patch().apply(&source).unwrap();
    let value = applied
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .number(reference)
        .unwrap()
        .unwrap()
        .value();
    assert_eq!(value, 9.25);
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );
}

#[test]
fn source_backed_number_publication_preserves_the_artifact_contract() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit();
    edit.set_number("Sheet1".into(), reference, 9.25).unwrap();
    let commit = edit.commit_source_backed().unwrap();

    assert_eq!(commit.diagnostics().changed_cells(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    assert_eq!(commit.diagnostics().splice_count(), 1);
    assert_eq!(commit.diagnostics().replacement_bytes(), 8);
    assert!(commit.diagnostics().changed_spans() > 0);
    assert_eq!(
        commit.diagnostics().source_version(),
        source.source_version()
    );
    assert_eq!(
        commit.diagnostics().target_version(),
        commit.snapshot().source_version()
    );
    assert_ne!(
        commit.diagnostics().source_fingerprint(),
        commit.diagnostics().target_fingerprint()
    );

    let mut output = Vec::new();
    let report = commit.write_to(&mut output).unwrap();
    assert_eq!(output, commit.snapshot().bytes());
    assert_eq!(
        report.target_fingerprint(),
        commit.diagnostics().target_fingerprint()
    );
    assert_eq!(stream(&output, "Opaque"), b"untouched");
    assert_eq!(
        commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(reference)
            .unwrap()
            .unwrap()
            .value(),
        9.25
    );

    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );
}

#[test]
fn source_backed_numeric_plan_validates_without_retaining_target_snapshot() {
    let cases = [
        (Storage::Number, Reference::new(3, 2).unwrap(), 9.25),
        (Storage::Rk, Reference::new(3, 2).unwrap(), 8.0),
        (Storage::MulRk, Reference::new(3, 2).unwrap(), 6.0),
    ];
    for (storage, package_reference, replacement) in cases {
        let (bytes, reference) = if storage == Storage::Number {
            (package(), package_reference)
        } else {
            fixed_numeric_family_package(storage)
        };
        let source = Snapshot::from_bytes(bytes).unwrap();
        let mut edit = source.edit();
        edit.set_numeric("Sheet1".into(), reference, replacement)
            .unwrap();
        let plan = edit.commit_source_backed_plan().unwrap();
        assert_eq!(plan.diagnostics().splice_count(), 1);
        assert_eq!(
            plan.diagnostics().replacement_bytes(),
            if storage == Storage::Number { 8 } else { 4 }
        );
        assert_eq!(
            plan.diagnostics().target_workbook_bytes(),
            plan.diagnostics().source_workbook_bytes()
        );

        let mut output = Vec::new();
        plan.write_to(&mut output).unwrap();
        assert_eq!(stream(&output, "Opaque"), b"untouched");
        let reopened = Snapshot::from_bytes(output).unwrap();
        let cell = reopened
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap();
        assert_eq!(cell.value(), &Value::Number(replacement));
        assert_eq!(cell.storage(), storage);
    }
}

#[test]
fn source_backed_numeric_plan_preserves_complete_directory_topology_and_clsids() {
    let source = Snapshot::from_bytes(topology_with_clsids_package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let before = directory_signature(source.bytes());
    let mut edit = source.edit();
    edit.set_number("Sheet1".into(), reference, 9.25).unwrap();
    let plan = edit.commit_source_backed_plan().unwrap();

    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert_eq!(directory_signature(&output), before);
    assert_eq!(
        OleFile::open(Cursor::new(&output))
            .unwrap()
            .open_stream(&["Metadata", "Payload"])
            .unwrap(),
        b"metadata"
    );
}

#[test]
fn source_backed_numeric_plan_rejects_stale_and_foreign_change_metadata() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();

    let mut stale_reference = source.edit();
    stale_reference
        .set_number("Sheet1".into(), reference, 9.25)
        .unwrap();
    stale_reference.changes[0].reference = Reference::new(3, 3).unwrap();
    assert!(stale_reference.commit_source_backed_plan().is_err());

    let foreign_entry = source.inner.sheets[0]
        .entries
        .iter()
        .position(|entry| {
            entry.cell.reference != reference && entry.cell.storage == Storage::Number
        })
        .unwrap();
    let mut stale_entry = source.edit();
    stale_entry
        .set_number("Sheet1".into(), reference, 9.25)
        .unwrap();
    stale_entry.changes[0].entry = foreign_entry;
    assert!(stale_entry.commit_source_backed_plan().is_err());
}

#[test]
fn source_backed_numeric_plan_handles_noncanonical_noop_and_sink_failure() {
    let raw = ((1.0_f64.to_bits() >> 32) as u32) & 0xffff_fffc;
    let (bytes, reference) = raw_rk_family_package(Storage::Rk, &[raw]);
    let source = Snapshot::from_bytes(bytes).unwrap();
    let mut noop_edit = source.edit();
    noop_edit
        .set_numeric("Sheet1".into(), reference, 1.0)
        .unwrap();
    let noop = noop_edit.commit_source_backed_plan().unwrap();
    assert!(noop.is_noop());
    assert_eq!(noop.diagnostics().splice_count(), 0);
    assert!(std::ptr::eq(
        noop.source().bytes().as_ptr(),
        source.bytes().as_ptr()
    ));
    assert_eq!(
        noop.diagnostics().source_fingerprint(),
        noop.diagnostics().target_fingerprint()
    );
    assert_eq!(
        noop.diagnostics().source_version(),
        noop.diagnostics().target_version()
    );

    let mut edit = source.edit();
    edit.set_numeric("Sheet1".into(), reference, 2.0).unwrap();
    let plan = edit.commit_source_backed_plan().unwrap();
    let mut sink = PrefixFailingSink {
        accepted: Vec::new(),
        limit: 17,
    };
    let error = plan.write_to(&mut sink).unwrap_err();
    assert!(matches!(error, OverlayError::IncompleteOutput { .. }));
    assert_eq!(sink.accepted.len(), sink.limit);
}

#[test]
fn source_backed_numeric_plan_refuses_structural_macro_and_signed_inputs() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut structural = source.edit();
    structural.insert_rows("Sheet1".into(), 1, 1).unwrap();
    assert!(structural.commit_source_backed_plan().is_err());

    let source = Snapshot::from_bytes(macro_package()).unwrap();
    let mut macro_edit = source.edit();
    macro_edit
        .set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    assert!(macro_edit.commit_source_backed_plan().is_err());

    let source = Snapshot::from_bytes(empty_macro_storage_package()).unwrap();
    let mut empty_macro_edit = source.edit();
    empty_macro_edit
        .set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    assert!(empty_macro_edit.commit_source_backed_plan().is_err());

    let mut empty_macro_eager_edit = source.edit();
    empty_macro_eager_edit
        .set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    assert!(empty_macro_eager_edit.commit_source_backed().is_err());

    // Signed packages are rejected during Snapshot opening, so this also
    // guards the plan path from accepting a marker through a different entry.
    assert!(Snapshot::from_bytes(signed_package()).is_err());
}

#[test]
fn source_backed_numeric_noop_reuses_identity_and_fingerprint() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit();
    edit.set_number("Sheet1".into(), reference, 4.5).unwrap();
    let commit = edit.commit_source_backed().unwrap();

    assert!(commit.is_noop());
    assert!(commit.patch().is_empty());
    assert!(std::ptr::eq(
        commit.snapshot().bytes().as_ptr(),
        source.bytes().as_ptr()
    ));
    assert_eq!(
        commit.diagnostics().source_fingerprint(),
        commit.diagnostics().target_fingerprint()
    );
    assert_eq!(
        commit.diagnostics().source_version(),
        commit.diagnostics().target_version()
    );
    assert_eq!(commit.diagnostics().splice_count(), 0);
    assert_eq!(commit.diagnostics().replacement_bytes(), 0);
}

#[test]
fn source_backed_numeric_refuses_non_numeric_and_structural_operations() {
    let source = Snapshot::from_bytes(package()).unwrap();

    let mut text = source.edit();
    text.set_value(
        "Sheet1".into(),
        Reference::new(6, 0).unwrap(),
        Value::Text("beta".into()),
    )
    .unwrap();
    assert!(text.commit_source_backed().is_err());

    let mut structural = source.edit();
    structural.insert_rows("Sheet1".into(), 1, 1).unwrap();
    assert!(structural.commit_source_backed().is_err());
}

#[test]
fn source_backed_rk_and_mulrk_publication_retain_their_families() {
    for (storage, replacement) in [(Storage::Rk, 8.0), (Storage::MulRk, 6.0)] {
        let (bytes, reference) = fixed_numeric_family_package(storage);
        let source = Snapshot::from_bytes(bytes).unwrap();
        assert_eq!(
            source
                .worksheet("Sheet1".into())
                .unwrap()
                .unwrap()
                .cell(reference)
                .unwrap()
                .unwrap()
                .storage(),
            storage
        );
        let mut edit = source.edit();
        edit.set_numeric("Sheet1".into(), reference, replacement)
            .unwrap();
        let commit = edit.commit_source_backed().unwrap();
        assert_eq!(commit.diagnostics().replacement_bytes(), 4);
        assert_eq!(commit.diagnostics().splice_count(), 1);
        let cell = commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap();
        assert_eq!(cell.storage(), storage);
        assert_eq!(cell.value(), &Value::Number(replacement));
        assert_eq!(stream(commit.snapshot().bytes(), "Opaque"), b"untouched");
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .bytes(),
            source.bytes()
        );
    }
}

#[test]
fn source_backed_rk_accepts_signed_integer_boundaries_and_noncanonical_source_bytes() {
    let rk_integer = |value: i32, divide_by_100: bool| {
        (crate::utils::reinterpret_i32_as_u32(value) << 2) | if divide_by_100 { 0x03 } else { 0x02 }
    };
    let maximum = (1_i32 << 29) - 1;
    let minimum = -(1_i32 << 29);
    let cases = [
        (rk_integer(1, false), 1.0, 2.0),
        (rk_integer(-1, false), -1.0, -2.0),
        (
            rk_integer(maximum, false),
            f64::from(maximum),
            f64::from(maximum - 1),
        ),
        (
            rk_integer(minimum, false),
            f64::from(minimum),
            f64::from(minimum + 1),
        ),
        (rk_integer(-123, true), -1.23, -1.24),
        // An IEEE upper-word encoding of an integral value is valid but
        // noncanonical relative to the signed integer form above.
        (((1.0_f64.to_bits() >> 32) as u32) & 0xffff_fffc, 1.0, 2.0),
    ];

    for (raw, source_value, replacement) in cases {
        let (bytes, reference) = raw_rk_family_package(Storage::Rk, &[raw]);
        let source = Snapshot::from_bytes(bytes).unwrap();
        let sheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
        assert_eq!(
            sheet.cell(reference).unwrap().unwrap().value(),
            &Value::Number(source_value)
        );
        let entry_index = unique_entry_index(&source.inner.sheets[0].entries, reference)
            .unwrap()
            .unwrap();
        let entry = &source.inner.sheets[0].entries[entry_index];
        let start = entry.value_offset.unwrap();
        assert_eq!(
            u32::from_le_bytes(
                source.inner.workbook_stream[start..start + 4]
                    .try_into()
                    .unwrap()
            ),
            raw
        );

        let mut edit = source.edit();
        edit.set_numeric("Sheet1".into(), reference, replacement)
            .unwrap();
        let commit = edit.commit_source_backed().unwrap();
        let target_cell = commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap();
        assert_eq!(target_cell.value(), &Value::Number(replacement));
        let target_entry = &commit.snapshot().inner.sheets[0].entries[entry_index];
        let target_start = target_entry.value_offset.unwrap();
        let expected_replacement = encode_rk(replacement).unwrap();
        assert_eq!(
            u32::from_le_bytes(
                commit.snapshot().inner.workbook_stream[target_start..target_start + 4]
                    .try_into()
                    .unwrap()
            ),
            expected_replacement
        );
        assert_eq!(commit.diagnostics().splice_count(), 1);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .bytes(),
            source.bytes()
        );
    }
}

#[test]
fn source_backed_scaled_rk_rounds_only_exact_decodes() {
    let rk_integer = |value: i32, divide_by_100: bool| {
        (crate::utils::reinterpret_i32_as_u32(value) << 2) | if divide_by_100 { 0x03 } else { 0x02 }
    };
    let raw_029 = rk_integer(29, true);
    let raw_negative = rk_integer(-999_997, true);
    assert_eq!(
        encode_rk(0.29),
        Some(raw_029),
        "rounded scaled RK must reproduce 0.29 exactly"
    );
    assert_eq!(
        encode_rk(-9999.97),
        Some(raw_negative),
        "rounded scaled RK must reproduce -9999.97 exactly"
    );

    let maximum = (1_i32 << 29) - 1;
    let minimum = -(1_i32 << 29);
    assert_eq!(
        encode_rk(f64::from(maximum) / 100.0),
        Some(rk_integer(maximum, true))
    );
    assert_eq!(
        encode_rk(f64::from(minimum) / 100.0),
        Some(rk_integer(minimum, true))
    );
    // A nearby value that rounds to 29 must still refuse because its exact
    // f64 bits are not the value decoded from the candidate RK field.
    assert!(encode_rk(0.29000000000000004).is_none());

    let cases = [
        (rk_integer(1, false), 1.0, 0.29),
        (rk_integer(-1, false), -1.0, -9999.97),
        (raw_029, 0.29, 0.3),
        (raw_negative, -9999.97, -9999.96),
    ];
    for (raw, source_value, replacement) in cases {
        let (bytes, reference) = raw_rk_family_package(Storage::Rk, &[raw]);
        let source = Snapshot::from_bytes(bytes).unwrap();
        assert_eq!(
            source
                .worksheet("Sheet1".into())
                .unwrap()
                .unwrap()
                .cell(reference)
                .unwrap()
                .unwrap()
                .value(),
            &Value::Number(source_value)
        );
        let mut edit = source.edit();
        edit.set_numeric("Sheet1".into(), reference, replacement)
            .unwrap();
        let commit = edit.commit_source_backed().unwrap();
        assert_eq!(
            commit
                .snapshot()
                .worksheet("Sheet1".into())
                .unwrap()
                .unwrap()
                .cell(reference)
                .unwrap()
                .unwrap()
                .value(),
            &Value::Number(replacement)
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .bytes(),
            source.bytes()
        );
    }
}

#[test]
fn source_backed_mulrk_accepts_signed_integer_fields_and_inverse() {
    let rk_integer = |value: i32| (crate::utils::reinterpret_i32_as_u32(value) << 2) | 0x02;
    let maximum = (1_i32 << 29) - 1;
    let raw_values = [rk_integer(-1), rk_integer(maximum)];
    let (bytes, first_reference) = raw_rk_family_package(Storage::MulRk, &raw_values);
    let second_reference = Reference::new(
        u32::from(first_reference.row()),
        u32::from(first_reference.column()) + 1,
    )
    .unwrap();
    let source = Snapshot::from_bytes(bytes).unwrap();
    let sheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(
        sheet.cell(first_reference).unwrap().unwrap().value(),
        &Value::Number(-1.0)
    );
    assert_eq!(
        sheet.cell(second_reference).unwrap().unwrap().value(),
        &Value::Number(f64::from(maximum))
    );

    let mut edit = source.edit();
    edit.set_numeric("Sheet1".into(), first_reference, -2.0)
        .unwrap();
    edit.set_numeric("Sheet1".into(), second_reference, f64::from(maximum - 1))
        .unwrap();
    let commit = edit.commit_source_backed().unwrap();
    assert_eq!(commit.diagnostics().splice_count(), 2);
    let target = commit
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert_eq!(
        target.cell(first_reference).unwrap().unwrap().value(),
        &Value::Number(-2.0)
    );
    assert_eq!(
        target.cell(second_reference).unwrap().unwrap().value(),
        &Value::Number(f64::from(maximum - 1))
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .bytes(),
        source.bytes()
    );
}

#[test]
fn source_backed_noncanonical_rk_noop_keeps_exact_source_identity() {
    let raw = ((1.0_f64.to_bits() >> 32) as u32) & 0xffff_fffc;
    let (bytes, reference) = raw_rk_family_package(Storage::Rk, &[raw]);
    let source = Snapshot::from_bytes(bytes).unwrap();
    let mut edit = source.edit();
    edit.set_numeric("Sheet1".into(), reference, 1.0).unwrap();
    let commit = edit.commit_source_backed().unwrap();
    assert!(commit.is_noop());
    assert!(std::ptr::eq(
        commit.snapshot().bytes().as_ptr(),
        source.bytes().as_ptr()
    ));
    assert_eq!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(commit.diagnostics().splice_count(), 0);
}

#[test]
fn source_backed_real_producer_rk_fixture_reopens_and_inverts() {
    let source = Snapshot::from_bytes(
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/54016.xls").to_vec(),
    )
    .unwrap();
    let (sheet, reference, replacement) = source
        .worksheets()
        .find_map(|sheet| {
            sheet.cells().find_map(|cell| {
                let Value::Number(value) = cell.value() else {
                    return None;
                };
                if !matches!(cell.storage(), Storage::Rk | Storage::MulRk) {
                    return None;
                }
                [*value + 1.0, *value - 1.0]
                    .into_iter()
                    .find(|candidate| {
                        candidate.to_bits() != value.to_bits() && encode_rk(*candidate).is_some()
                    })
                    .map(|candidate| (sheet.position(), cell.reference(), candidate))
            })
        })
        .expect("POI 54016 fixture has an RK/MulRk cell with an RK replacement");
    let mut edit = source.edit();
    edit.set_numeric(Selector::Position(sheet), reference, replacement)
        .unwrap();
    let commit = edit.commit_source_backed().unwrap();
    let reopened = Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let cell = reopened
        .worksheets()
        .nth(sheet)
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.value(), &Value::Number(replacement));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(&commit.snapshot())
            .unwrap()
            .bytes(),
        source.bytes()
    );
}

#[test]
fn source_backed_numeric_sink_reports_partial_output() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut edit = source.edit();
    edit.set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    let commit = edit.commit_source_backed().unwrap();
    let mut sink = PrefixFailingSink {
        accepted: Vec::new(),
        limit: 17,
    };
    let error = commit.write_to(&mut sink).unwrap_err();
    assert!(matches!(error, OverlayError::IncompleteOutput { .. }));
    assert_eq!(sink.accepted.len(), sink.limit);
}

struct PrefixFailingSink {
    accepted: Vec<u8>,
    limit: usize,
}

impl Write for PrefixFailingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        if self.accepted.len() >= self.limit {
            return Err(std::io::Error::other("test sink refusal"));
        }
        let remaining = self.limit - self.accepted.len();
        let count = remaining.min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[test]
fn sequential_number_commits_reuse_resources_and_reopen_publicly() {
    use litchi_core::sheet::Cell as _;

    let source = Snapshot::from_bytes(package()).unwrap();
    let first_reference = Reference::new(3, 2).unwrap();
    let second_reference = Reference::new(4, 1).unwrap();
    let mut first = source.edit();
    first
        .set_number("Sheet1".into(), first_reference, 9.25)
        .unwrap();
    let first = first.commit().unwrap();
    let mut second = first.snapshot().edit();
    second
        .set_number("Sheet1".into(), second_reference, 7.5)
        .unwrap();
    let second = second.commit().unwrap();

    assert!(Arc::ptr_eq(
        &source.inner.shared_strings,
        &second.snapshot().inner.shared_strings
    ));
    assert!(Arc::ptr_eq(
        &source.inner.xf_records,
        &second.snapshot().inner.xf_records
    ));
    let workbook = Workbook::new(Cursor::new(second.snapshot().bytes())).unwrap();
    let metadata = workbook.sheet(0).unwrap();
    let worksheet = workbook
        .xls_worksheet(metadata.parsed_worksheet_index().unwrap())
        .unwrap();
    assert_eq!(
        worksheet.get_cell(3, 2).unwrap().value().as_float(),
        Some(9.25)
    );
    assert_eq!(
        worksheet.get_cell(4, 1).unwrap().value().as_float(),
        Some(7.5)
    );
    assert_eq!(stream(second.snapshot().bytes(), "Opaque"), b"untouched");
}

#[test]
fn number_commit_shares_untouched_worksheet_inventories() {
    let mut writer = crate::Writer::new();
    let first = writer.add_worksheet("First").unwrap();
    writer.write_number(first, 0, 0, 1.0).unwrap();
    let second = writer.add_worksheet("Second").unwrap();
    writer.write_number(second, 0, 0, 2.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = Snapshot::from_bytes(output.into_inner()).unwrap();

    let mut edit = source.edit();
    edit.set_number("First".into(), Reference::new(0, 0).unwrap(), 3.0)
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(!Arc::ptr_eq(
        &source.inner.sheets[0].entries,
        &commit.snapshot().inner.sheets[0].entries
    ));
    assert!(Arc::ptr_eq(
        &source.inner.sheets[1].entries,
        &commit.snapshot().inner.sheets[1].entries
    ));
}

#[test]
fn inventory_carry_refuses_bytes_outside_changed_numeric_field() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let sheet = source.resolve_sheet("Sheet1".into()).unwrap().unwrap();
    let entry = unique_entry_index(&source.inner.sheets[sheet].entries, reference)
        .unwrap()
        .unwrap();
    let change = Change {
        sheet,
        entry,
        reference,
        storage: Storage::Number,
        value: Value::Number(9.25),
    };
    let mut workbook = source.workbook_stream().to_vec();
    write_cell_value(
        &mut workbook,
        &source.inner.sheets[sheet].entries[entry],
        &change,
        &[],
    )
    .unwrap();
    let kind_offset = source.inner.sheets[sheet].entries[entry].kind_offset;
    workbook[kind_offset] ^= 1;

    let error = carry_fixed_numeric_inventory(&source, &workbook, &[change]).unwrap_err();
    assert!(error.to_string().contains("outside its value fields"));
}

#[test]
fn rejected_and_noop_edits_are_failure_atomic() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::NAN)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::INFINITY)
            .is_err()
    );
    assert!(edit.set_number("Sheet1".into(), reference, -0.0).is_err());
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::MIN_POSITIVE / 2.0)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), Reference::new(9, 9).unwrap(), 1.0)
            .is_err()
    );
    edit.set_number("sHeEt1".into(), reference, 0.0).unwrap();
    edit.set_number("Sheet1".into(), reference, 4.5).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(commit.diagnostics(), Diagnostics::default());
}

#[test]
fn references_enforce_the_biff8_grid() {
    let reference = Reference::new(u32::from(u16::MAX), u32::from(u8::MAX)).unwrap();
    assert_eq!(reference.row(), u16::MAX);
    assert_eq!(reference.column(), u8::MAX);
    assert!(Reference::new(u32::from(u16::MAX) + 1, 0).is_err());
    assert!(Reference::new(0, u32::from(u8::MAX) + 1).is_err());
}

#[test]
fn signed_packages_are_refused_before_editing() {
    assert!(Snapshot::from_bytes(signed_package()).is_err());
}

#[test]
fn macro_packages_are_refused_by_source_backed_numeric_publication() {
    let source = Snapshot::from_bytes(macro_package()).unwrap();
    let mut edit = source.edit();
    edit.set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    assert!(edit.commit_source_backed().is_err());

    let mut eager = source.edit();
    eager
        .set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 9.25)
        .unwrap();
    assert!(eager.commit().is_ok());
}

#[test]
fn edits_boolean_text_and_formula_cache() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let sheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(
        sheet
            .cell(Reference::new(4, 0).unwrap())
            .unwrap()
            .unwrap()
            .storage(),
        Storage::Number
    );
    assert_eq!(
        sheet
            .cell(Reference::new(6, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("alpha".to_string())
    );
    assert_eq!(
        sheet
            .cell(Reference::new(8, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Empty)
    );

    let mut edit = source.edit();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(4, 0).unwrap(),
        Value::Number(11.0),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(6, 0).unwrap(),
        Value::Text("beta".to_string()),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(8, 0).unwrap(),
        Value::FormulaCache(FormulaCache::Error(CellError::new(0x07).unwrap())),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_cells(), 4);
    let sheet = commit
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert_eq!(
        sheet
            .cell(Reference::new(4, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Number(11.0)
    );
    assert_eq!(
        sheet
            .cell(Reference::new(7, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Boolean(false)
    );
    assert_eq!(
        sheet
            .cell(Reference::new(6, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("beta".to_string())
    );
    assert_eq!(
        sheet
            .cell(Reference::new(8, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Error(CellError::new(0x07).unwrap()))
    );
}

#[test]
fn formula_string_cache_resizes_and_semantically_inverts() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(8, 0).unwrap();
    let mut transaction = source.transaction();
    transaction
        .set_value(
            "Sheet1".into(),
            reference,
            Value::FormulaCache(FormulaCache::String("cached λ value".to_string())),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::String("cached λ value".to_string()))
    );
    let wire = commit.patch().semantic().to_deterministic_json().unwrap();
    let durable = SemanticPatch::from_deterministic_json(&wire).unwrap();
    let restored = durable.inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Empty)
    );
}

#[test]
fn formula_cache_transfer_requires_identical_tokens() {
    let source = Snapshot::from_bytes(formula_package("1+1")).unwrap();
    let mut transaction = source.transaction();
    transaction
        .set_value(
            "Formula".into(),
            Reference::new(0, 0).unwrap(),
            Value::FormulaCache(FormulaCache::Boolean(true)),
        )
        .unwrap();
    let patch = transaction.commit().unwrap().patch().semantic().clone();
    let divergent = Snapshot::from_bytes(formula_package("2+0")).unwrap();
    let transfer = patch.plan_transfer(&divergent);
    assert!(!transfer.is_executable());
    assert!(patch.apply(&divergent).is_err());
}

#[test]
fn authored_formulas_join_reopen_and_durably_invert() {
    let source = Snapshot::from_bytes(formula_free_package()).unwrap();
    let first = Reference::new(1, 1).unwrap();
    let second = Reference::new(3, 2).unwrap();
    let mut left = source.transaction();
    left.insert_formula("Plain".into(), first, "A2+1").unwrap();
    let mut right = source.transaction();
    right
        .insert_formula("Plain".into(), second, "B4*2")
        .unwrap();
    left.join(right).unwrap();
    let commit = left.commit().unwrap();
    for reference in [first, second] {
        assert_eq!(
            commit
                .snapshot()
                .worksheet("Plain".into())
                .unwrap()
                .unwrap()
                .cell(reference)
                .unwrap()
                .unwrap()
                .value(),
            &Value::FormulaCache(FormulaCache::Empty)
        );
    }
    Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    let wire = commit.patch().semantic().to_deterministic_json().unwrap();
    let durable = SemanticPatch::from_deterministic_json(&wire).unwrap();
    let replay = durable.apply(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    let restored = durable.inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.snapshot().bytes(), source.bytes());
}

#[test]
fn new_sst_and_xf_resources_round_trip_with_cells() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let base_style = source
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(Reference::new(3, 2).unwrap())
        .unwrap()
        .unwrap()
        .style();
    let mut transaction = source.transaction();
    let mut xf_payload = source.inner.xf_records[usize::from(base_style.get())].clone();
    xf_payload[6] ^= 0x08;
    let authored_style = transaction.author_style(&xf_payload).unwrap();
    let text_reference = Reference::new(3, 3).unwrap();
    let style_reference = Reference::new(4, 3).unwrap();
    let continued_text = "λ".repeat(5_000);
    transaction
        .insert_cell(
            "Sheet1".into(),
            text_reference,
            Value::Text(continued_text.clone()),
        )
        .unwrap();
    transaction
        .insert_cell_with_style(
            "Sheet1".into(),
            style_reference,
            Value::Number(33.0),
            authored_style,
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    let sheet = commit
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert_eq!(
        sheet.cell(text_reference).unwrap().unwrap().value(),
        &Value::Text(continued_text)
    );
    assert_eq!(
        sheet.cell(style_reference).unwrap().unwrap().style(),
        authored_style
    );
    let restored = commit
        .patch()
        .semantic()
        .inverse()
        .apply(commit.snapshot())
        .unwrap();
    let restored = restored
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert!(restored.cell(text_reference).unwrap().is_none());
    assert!(restored.cell(style_reference).unwrap().is_none());
}

#[test]
fn rich_sst_resource_round_trips_and_semantically_inverts() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let font_index = binary::read_u16_le_at(&source.inner.xf_records[0], 0).unwrap();
    let runs = vec![crate::records::SharedStringFormatRun {
        character_index: 0,
        font_index,
    }];
    let reference = Reference::new(6, 3).unwrap();
    let reused_reference = Reference::new(7, 3).unwrap();
    let text = "λ".repeat(5_000);
    let mut transaction = source.transaction();
    transaction
        .insert_rich_text_cell("Sheet1".into(), reference, text.clone(), runs.clone())
        .unwrap();
    transaction
        .insert_rich_text_cell(
            "Sheet1".into(),
            reused_reference,
            text.clone(),
            runs.clone(),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text(text.clone())
    );
    assert_eq!(
        commit
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reused_reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text(text.clone())
    );
    assert_eq!(
        commit
            .snapshot()
            .inner
            .shared_strings
            .iter()
            .filter(|candidate| candidate.as_str() == text.as_str())
            .count(),
        1
    );
    let index = commit
        .snapshot()
        .inner
        .shared_strings
        .iter()
        .position(|candidate| candidate == &text)
        .unwrap();
    assert_eq!(
        commit.snapshot().inner.shared_string_properties[index]
            .as_deref()
            .unwrap()
            .formatting_runs,
        runs
    );
    let restored = commit
        .patch()
        .semantic()
        .inverse()
        .apply(commit.snapshot())
        .unwrap();
    let restored_bytes = restored.snapshot().bytes();
    let source_bytes = source.bytes();
    assert_eq!(restored_bytes.len(), source_bytes.len());
    let first_difference = restored_bytes
        .iter()
        .zip(source_bytes)
        .position(|(restored, source)| restored != source);
    assert_eq!(first_difference, None);
}

#[test]
fn converts_fixed_width_labelsst_and_rk_with_sst_accounting() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(6, 0).unwrap();
    let mut to_number = source.edit();
    to_number
        .set_value("Sheet1".into(), reference, Value::Number(42.0))
        .unwrap();
    let number = to_number.commit().unwrap();
    let cell = number
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.storage(), Storage::Rk);
    assert_eq!(cell.value(), &Value::Number(42.0));

    let mut to_text = number.snapshot().edit();
    to_text
        .set_value("Sheet1".into(), reference, Value::Text("beta".to_string()))
        .unwrap();
    let text = to_text.commit().unwrap();
    let cell = text
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.storage(), Storage::LabelSst);
    assert_eq!(cell.value(), &Value::Text("beta".to_string()));
}

#[test]
fn durable_semantic_patch_round_trips_and_checks_preconditions() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(7, 0).unwrap();
    let mut edit = source.transaction();
    edit.set_value("Sheet1".into(), reference, Value::Boolean(false))
        .unwrap();
    let commit = edit.commit().unwrap();
    let json = commit.patch().semantic().to_deterministic_json().unwrap();
    let parsed = SemanticPatch::from_deterministic_json(&json).unwrap();
    assert_eq!(parsed.to_deterministic_json().unwrap(), json);
    let replay = parsed.apply(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    assert!(parsed.apply(commit.snapshot()).is_err());
    let restored = parsed.inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::Boolean(true)
    );
}

#[test]
fn joins_disjoint_work_and_reports_cell_conflicts() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut left = source.transaction();
    left.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    let mut right = source.transaction();
    right
        .set_value(
            "Sheet1".into(),
            Reference::new(6, 0).unwrap(),
            Value::Text("beta".to_string()),
        )
        .unwrap();
    left.join(right).unwrap();
    assert_eq!(left.commit().unwrap().diagnostics().changed_cells(), 2);

    let mut left = source.transaction();
    left.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    let mut right = source.transaction();
    right
        .set_value(
            "Sheet1".into(),
            Reference::new(7, 0).unwrap(),
            Value::Error(CellError::new(0x07).unwrap()),
        )
        .unwrap();
    let error = left.join(right).unwrap_err();
    let JoinError::Conflicts(conflicts) = error else {
        panic!("expected structured cell conflict")
    };
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts[0].reference(), Reference::new(7, 0).unwrap());
    assert_eq!(left.commit().unwrap().diagnostics().changed_cells(), 1);
}

#[test]
fn bounded_history_undoes_and_redoes_immutable_snapshots() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut history = source.history(HistoryLimits::new(2, u64::MAX));
    let mut edit = source.transaction();
    edit.set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 99.0)
        .unwrap();
    edit.commit().unwrap().record_in(&mut history).unwrap();
    assert!(history.can_undo());
    assert!(history.undo());
    assert_eq!(history.current().bytes(), source.bytes());
    assert!(history.redo());
    assert_eq!(
        history
            .current()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(Reference::new(3, 2).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        99.0
    );
}

#[test]
fn structural_cell_style_and_sheet_name_round_trip_semantically() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 3).unwrap();
    let style = source
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(Reference::new(3, 2).unwrap())
        .unwrap()
        .unwrap()
        .style();
    let mut transaction = source.transaction();
    transaction
        .insert_cell_with_style("Sheet1".into(), reference, Value::Number(27.5), style)
        .unwrap();
    transaction
        .rename_sheet("Sheet1".into(), "Renamed")
        .unwrap();
    let commit = transaction.commit().unwrap();
    let cell = commit
        .snapshot()
        .worksheet("Renamed".into())
        .unwrap()
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.value(), &Value::Number(27.5));
    assert_eq!(cell.style(), style);

    let json = commit.patch().semantic().to_deterministic_json().unwrap();
    let durable = SemanticPatch::from_deterministic_json(&json).unwrap();
    let replay = durable.apply(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    let restored = durable.inverse().apply(commit.snapshot()).unwrap();
    let restored_sheet = restored
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert!(restored_sheet.cell(reference).unwrap().is_none());
}

#[test]
fn structural_join_three_way_and_transfer_are_deterministic() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut left = source.transaction();
    left.insert_cell(
        "Sheet1".into(),
        Reference::new(3, 3).unwrap(),
        Value::Number(10.0),
    )
    .unwrap();
    let mut right = source.transaction();
    right
        .insert_cell(
            "Sheet1".into(),
            Reference::new(4, 3).unwrap(),
            Value::Boolean(true),
        )
        .unwrap();
    left.join(right).unwrap();
    assert_eq!(left.commit().unwrap().diagnostics().changed_cells(), 2);

    let mut left = source.transaction();
    left.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    let left = left.commit().unwrap().patch().semantic().clone();
    let mut right = source.transaction();
    right
        .set_value(
            "Sheet1".into(),
            Reference::new(6, 0).unwrap(),
            Value::Text("beta".to_string()),
        )
        .unwrap();
    let right = right.commit().unwrap().patch().semantic().clone();
    let plan = SemanticPatch::plan_three_way(&source, &left, &right).unwrap();
    assert!(plan.conflicts().is_empty());
    let merged = plan.merged().unwrap();
    assert_eq!(merged.len(), 2);
    let transfer = merged.plan_transfer(&source);
    assert!(transfer.is_executable());
    transfer.execute(&source).unwrap();
}

#[test]
fn dependency_safe_row_insertion_reopens_and_inverts() {
    let source = Snapshot::from_bytes(formula_free_package()).unwrap();
    let mut transaction = source.transaction();
    transaction.insert_rows("Plain".into(), 2, 2).unwrap();
    let commit = transaction.commit().unwrap();
    let moved = commit
        .snapshot()
        .worksheet("Plain".into())
        .unwrap()
        .unwrap()
        .cell(Reference::new(5, 1).unwrap())
        .unwrap()
        .unwrap();
    assert_eq!(moved.value(), &Value::Number(2.0));
    let inverse = commit.patch().semantic().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored
            .snapshot()
            .worksheet("Plain".into())
            .unwrap()
            .unwrap()
            .cell(Reference::new(3, 1).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Number(2.0)
    );
}

#[test]
fn reference_free_formula_shifts_reopen_and_invert_exactly() {
    let source = Snapshot::from_bytes(formula_package("1+1")).unwrap();
    let mut transaction = source.transaction();
    transaction.insert_rows("Formula".into(), 0, 2).unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .worksheet("Formula".into())
            .unwrap()
            .unwrap()
            .cell(Reference::new(2, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Empty)
    );
    Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    let wire = commit.patch().semantic().to_deterministic_json().unwrap();
    let durable = SemanticPatch::from_deterministic_json(&wire).unwrap();
    assert!(durable.plan_transfer(&source).is_executable());
    let replay = durable.apply(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    let restored = commit
        .patch()
        .semantic()
        .inverse()
        .apply(commit.snapshot())
        .unwrap();
    assert_eq!(restored.snapshot().bytes(), source.bytes());

    let column_source = Snapshot::from_bytes(formula_package("1+1")).unwrap();
    let mut columns = column_source.transaction();
    columns.insert_columns("Formula".into(), 0, 2).unwrap();
    let column_commit = columns.commit().unwrap();
    assert_eq!(
        column_commit
            .snapshot()
            .worksheet("Formula".into())
            .unwrap()
            .unwrap()
            .cell(Reference::new(0, 2).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Empty)
    );
    let column_restored = column_commit
        .patch()
        .semantic()
        .inverse()
        .apply(column_commit.snapshot())
        .unwrap();
    assert_eq!(column_restored.snapshot().bytes(), column_source.bytes());
}

#[test]
fn reference_formula_shifts_are_durable_mergeable_and_history_safe() {
    let source = Snapshot::from_bytes(positioned_formula_package("A1+$B$2+C1:C3", 4, 4)).unwrap();
    let mut transaction = source.transaction();
    transaction.insert_rows("Formula".into(), 1, 2).unwrap();
    let commit = transaction.commit().unwrap();

    let reopened = Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    assert_eq!(
        reopened
            .xls_worksheet(0)
            .unwrap()
            .get_cell(6, 4)
            .unwrap()
            .formula(),
        Some("=((A1+$B$4)+(C1:C5))")
    );

    let wire = commit.patch().semantic().to_deterministic_json().unwrap();
    let durable = SemanticPatch::from_deterministic_json(&wire).unwrap();
    let transfer = durable.plan_transfer(&source);
    assert!(transfer.is_executable());
    let replay = transfer.execute(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    let restored = durable.inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.snapshot().bytes(), source.bytes());

    let mut history = source.history(HistoryLimits::new(2, u64::MAX));
    commit.clone().record_in(&mut history).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().bytes(), source.bytes());
    assert!(history.redo());
    assert_eq!(history.current().bytes(), commit.snapshot().bytes());

    let mut rename = source.transaction();
    rename.rename_sheet("Formula".into(), "Shifted").unwrap();
    let rename = rename.commit().unwrap().patch().semantic().clone();
    let plan = SemanticPatch::plan_three_way(&source, &durable, &rename).unwrap();
    let merged = plan.merged().unwrap();
    let merged_commit = merged.apply(&source).unwrap();
    let merged_reopen = Workbook::new(Cursor::new(merged_commit.snapshot().bytes())).unwrap();
    assert_eq!(merged_reopen.sheets()[0].name(), "Shifted");
    assert_eq!(
        merged_reopen
            .xls_worksheet(0)
            .unwrap()
            .get_cell(6, 4)
            .unwrap()
            .formula(),
        Some("=((A1+$B$4)+(C1:C5))")
    );

    let column_source =
        Snapshot::from_bytes(positioned_formula_package("A1+$B$2+C1:C3", 4, 4)).unwrap();
    let mut columns = column_source.transaction();
    columns.insert_columns("Formula".into(), 1, 2).unwrap();
    let column_commit = columns.commit().unwrap();
    let column_reopen = Workbook::new(Cursor::new(column_commit.snapshot().bytes())).unwrap();
    assert_eq!(
        column_reopen
            .xls_worksheet(0)
            .unwrap()
            .get_cell(4, 6)
            .unwrap()
            .formula(),
        Some("=((A1+$D$2)+(E1:E3))")
    );
    assert_eq!(
        column_commit
            .patch()
            .semantic()
            .inverse()
            .apply(column_commit.snapshot())
            .unwrap()
            .snapshot()
            .bytes(),
        column_source.bytes()
    );

    let dependent = Snapshot::from_bytes(positioned_formula_package("A2", 4, 4)).unwrap();
    assert!(
        dependent
            .transaction()
            .delete_rows("Formula".into(), 1, 1)
            .is_err()
    );
}
