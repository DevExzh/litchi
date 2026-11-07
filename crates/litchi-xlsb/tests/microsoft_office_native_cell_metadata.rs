//! Native cell/value metadata stays bound to its cell; payloads remain inert.
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet};
use litchi_xlsb::cell_values::{Commit, Reference, StyleIndex, Value};
use litchi_xlsb::raw::{Limits, Records, kind};
use litchi_xlsb::{Package, SourceBackedWorkbook};
use std::{collections::BTreeMap, io::Cursor, path::Path};

type Prefixes = BTreeMap<(u32, u32), Vec<(u16, u32)>>;
const FORMULAS: &str = "excel-formula-owners.xlsb";
const RICH: &str = "excel-links-objects.xlsb";
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn u32_at(bytes: &[u8], at: usize) -> u32 {
    u32::from_le_bytes(bytes[at..at + 4].try_into().unwrap())
}
fn reference(row: u32, col: u32) -> Reference {
    Reference::new(row, col).unwrap()
}
fn value(p: &Package, sheet: usize, row: u32, column: u32) -> Option<CellValue> {
    let book = p.workbook().unwrap();
    book.worksheet_by_name(&book.worksheet_names()[sheet])
        .unwrap()
        .get_cell(row, column)
        .map(|cell| cell.value().clone())
}
fn prefixes(bytes: &[u8]) -> Prefixes {
    let mut out = Prefixes::new();
    let mut row = None;
    let mut pending = Vec::new();
    let mut in_data = false;
    for record in Records::new(bytes, Limits::DEFAULT) {
        let r = record.unwrap();
        match r.kind() {
            kind::BEGIN_SHEET_DATA => in_data = true,
            kind::END_SHEET_DATA => {
                assert!(pending.is_empty(), "orphan metadata at sheet-data end");
                in_data = false;
            },
            kind::ROW_HDR if in_data => {
                assert!(pending.is_empty(), "orphan metadata at row boundary");
                row = Some(u32_at(r.payload(), 0));
            },
            kind::CELL_META | kind::VALUE_META if in_data => {
                assert_eq!(r.payload().len(), 4);
                pending.push((r.kind().get(), u32_at(r.payload(), 0)));
            },
            k if in_data && (matches!(k.get(), 1..=11) || k == kind::CELL_R_STRING) => {
                if !pending.is_empty() {
                    assert!(
                        out.insert(
                            (row.unwrap(), u32_at(r.payload(), 0)),
                            std::mem::take(&mut pending)
                        )
                        .is_none()
                    );
                }
            },
            _ => assert!(pending.is_empty(), "metadata is separated from its cell"),
        }
    }
    assert!(pending.is_empty());
    out
}
fn publish(
    source: &Package,
    commit: &Commit,
    sheet: usize,
    expected: &Prefixes,
    changed_cells: &[(u32, u32)],
    extra_parts: &[&str],
) -> Package {
    let before = source.to_bytes().unwrap();
    let output = commit.package().to_bytes().unwrap();
    let a = members(&before);
    let b = members(&output);
    let sheet_part = format!("xl/worksheets/sheet{}.bin", sheet + 1);
    let index_part = format!("xl/worksheets/binaryIndex{}.bin", sheet + 1);
    assert_eq!(prefixes(&b[&sheet_part]), *expected);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in &a {
        if name != &sheet_part && name != &index_part && !extra_parts.contains(&name.as_str()) {
            assert_eq!(&b[name], data, "changed {name}");
        }
    }
    let path = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(path.path()).unwrap();
    let reopened = Package::open(path.path()).unwrap();
    assert_eq!(members(&reopened.to_bytes().unwrap()), b);
    let old = source.workbook().unwrap();
    let new = reopened.workbook().unwrap();
    let name = &old.worksheet_names()[sheet];
    let old_sheet = old.worksheet_by_name(name).unwrap();
    let new_sheet = new.worksheet_by_name(name).unwrap();
    for c in old_sheet.cells() {
        let c = c.unwrap();
        let c = c.as_ref();
        if changed_cells.contains(&(c.row(), c.column())) {
            continue;
        }
        let retained = new_sheet.get_cell(c.row(), c.column()).unwrap();
        assert_eq!(retained.value(), c.value());
        assert_eq!(retained.formula_bytes(), c.formula_bytes());
        assert_eq!(retained.style_id(), c.style_id());
    }
    let lazy = SourceBackedWorkbook::from_reader(Cursor::new(&output)).unwrap();
    let indexed = lazy
        .sheet(sheet)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap();
    for (row, col) in expected.keys().chain(changed_cells.iter()) {
        assert_eq!(
            indexed.cached_value(*row, *col).unwrap(),
            new_sheet.get_cell(*row, *col).map(|c| c.value().clone())
        );
    }
    assert_eq!(
        source
            .apply_cell_values(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        output
    );
    assert_eq!(
        reopened
            .apply_cell_values(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
    reopened
}

#[test]
fn native_xlsb_metadata_types_blocks_and_cell_associations_are_exact() {
    for (file, sheet, flags, name, cell_meta, expected) in [
        (
            FORMULAS,
            2,
            0xd86a_c0b0,
            "XLDAPR",
            1u32,
            BTreeMap::from([((3, 0), vec![(0x31, 1)]), ((3, 3), vec![(0x31, 1)])]),
        ),
        (
            RICH,
            1,
            0x986a_c0b0,
            "XLRICHVALUE",
            0,
            BTreeMap::from([((4, 1), vec![(0x32, 1)])]),
        ),
    ] {
        let bytes = native(file);
        let parts = members(&bytes);
        let records = Records::new(&parts["xl/metadata.bin"], Limits::DEFAULT)
            .collect::<Result<Vec<_>, _>>()
            .unwrap();
        assert_eq!(records[0].kind().get(), 0x14c);
        assert_eq!(records.last().unwrap().kind().get(), 0x14d);
        let types = records
            .iter()
            .filter(|r| r.kind().get() == 0x14f)
            .collect::<Vec<_>>();
        assert_eq!(types.len(), 1);
        let info = types[0].payload();
        assert_eq!((u32_at(info, 0), u32_at(info, 4)), (flags, 120000));
        let units = info[12..]
            .chunks_exact(2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
            .collect::<Vec<_>>();
        assert_eq!(units.len(), u32_at(info, 8) as usize);
        assert_eq!(String::from_utf16(&units).unwrap(), name);
        let blocks = records
            .iter()
            .filter(|r| r.kind().get() == 0x151)
            .collect::<Vec<_>>();
        assert_eq!(blocks.len(), 1);
        assert_eq!(
            blocks[0].payload(),
            [1u32.to_le_bytes(), cell_meta.to_le_bytes()].concat()
        );
        let block = records
            .iter()
            .filter(|r| r.kind().get() == 0x33)
            .collect::<Vec<_>>();
        assert_eq!(block.len(), 1);
        assert_eq!(
            block[0].payload(),
            [1u32.to_le_bytes(), 1u32.to_le_bytes(), 0u32.to_le_bytes()].concat()
        );
        assert_eq!(
            prefixes(&parts[&format!("xl/worksheets/sheet{sheet}.bin")]),
            expected
        );
        let opc = litchi_opc::OpcPackage::from_reader(
            Cursor::new(&bytes),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        let m = opc
            .get_part(&litchi_opc::PackURI::new("/xl/metadata.bin").unwrap())
            .unwrap();
        assert_eq!(m.content_type(), "application/vnd.ms-excel.sheetMetadata");
        let book = opc
            .get_part(&litchi_opc::PackURI::new("/xl/workbook.bin").unwrap())
            .unwrap();
        let relationships = book.rels().iter().filter(|r|r.reltype()=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata").collect::<Vec<_>>();
        assert_eq!(relationships.len(), 1);
        assert_eq!(relationships[0].target_ref(), "metadata.bin");
        let p = Package::from_bytes(&bytes).unwrap();
        let edit = p.edit_cell_values(sheet - 1).unwrap();
        let noop = edit.commit().unwrap();
        assert!(noop.patch().is_noop());
        assert_eq!(noop.package().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_xlsb_inserting_before_dynamic_array_anchor_keeps_cell_metadata_on_its_formula() {
    let bytes = native(FORMULAS);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected = prefixes(&members(&bytes)["xl/worksheets/sheet2.bin"]);
    let mut edit = p.edit_cell_values(1usize).unwrap();
    edit.insert(
        reference(3, 2),
        StyleIndex::new(0).unwrap(),
        Value::Number(88.0),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let q = publish(&p, &commit, 1, &expected, &[(3, 2)], &[]);
    assert_eq!(value(&q, 1, 3, 2), Some(CellValue::Float(88.0)));
    assert_eq!(native(FORMULAS), bytes);
}

#[test]
fn native_xlsb_inserting_before_rich_value_keeps_the_native_value_metadata_association() {
    let bytes = native(RICH);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected = prefixes(&members(&bytes)["xl/worksheets/sheet1.bin"]);
    // A5 is a native label; remove it to create an insertion point before B5.
    let mut edit = p.edit_cell_values(0usize).unwrap();
    assert!(edit.remove(reference(4, 0)).unwrap());
    let removed = edit.commit().unwrap();
    let source = publish(
        &p,
        &removed,
        0,
        &expected,
        &[(4, 0)],
        &["xl/sharedStrings.bin"],
    );
    let mut edit = source.edit_cell_values(0usize).unwrap();
    edit.insert(
        reference(4, 0),
        StyleIndex::new(0).unwrap(),
        Value::Number(88.0),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let q = publish(&source, &commit, 0, &expected, &[(4, 0)], &[]);
    assert_eq!(value(&q, 0, 4, 0), Some(CellValue::Float(88.0)));
    assert_eq!(native(RICH), bytes);
}

#[test]
fn native_xlsb_removing_rich_value_removes_its_prefix_without_retargeting_the_next_cell() {
    let bytes = native(RICH);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected = prefixes(&members(&bytes)["xl/worksheets/sheet1.bin"]);
    let mut edit = p.edit_cell_values(0usize).unwrap();
    edit.insert(
        reference(4, 2),
        StyleIndex::new(0).unwrap(),
        Value::Number(76.0),
    )
    .unwrap();
    let added = edit.commit().unwrap();
    let source = publish(&p, &added, 0, &expected, &[(4, 2)], &[]);
    let mut edit = source.edit_cell_values(0usize).unwrap();
    assert!(edit.remove(reference(4, 1)).unwrap());
    let commit = edit.commit().unwrap();
    let q = publish(&source, &commit, 0, &Prefixes::new(), &[(4, 1)], &[]);
    assert_eq!(value(&q, 0, 4, 1), None);
    assert_eq!(value(&q, 0, 4, 2), Some(CellValue::Float(76.0)));
    assert_eq!(native(RICH), bytes);
}

#[test]
fn native_xlsb_cache_and_header_edits_keep_metadata_prefixes_and_opaque_parts() {
    let bytes = native(FORMULAS);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected = prefixes(&members(&bytes)["xl/worksheets/sheet2.bin"]);
    let mut edit = p.edit_cell_values(1usize).unwrap();
    edit.set_formula_string_cache(reference(3, 0), "Stored cache 🍎".into())
        .unwrap();
    let commit = edit.commit().unwrap();
    let q = publish(&p, &commit, 1, &expected, &[(3, 0)], &[]);
    assert_eq!(
        value(&q, 1, 3, 0),
        Some(CellValue::String("Stored cache 🍎".into()))
    );
    let original_book = p.workbook().unwrap();
    let changed_book = q.workbook().unwrap();
    let name = &original_book.worksheet_names()[1];
    assert_eq!(
        original_book
            .worksheet_by_name(name)
            .unwrap()
            .get_cell(3, 0)
            .unwrap()
            .formula_bytes(),
        changed_book
            .worksheet_by_name(name)
            .unwrap()
            .get_cell(3, 0)
            .unwrap()
            .formula_bytes()
    );
    assert_eq!(native(FORMULAS), bytes);

    let bytes = native(RICH);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected = prefixes(&members(&bytes)["xl/worksheets/sheet1.bin"]);
    let original = value(&p, 0, 4, 1);
    assert_eq!(original, Some(CellValue::Error("#VALUE!".into())));
    let mut edit = p.edit_cell_values(0usize).unwrap();
    edit.set_show_phonetic(reference(4, 1), true).unwrap();
    let commit = edit.commit().unwrap();
    let q = publish(&p, &commit, 0, &expected, &[], &[]);
    assert_eq!(value(&q, 0, 4, 1), original);
    let parts = members(&q.to_bytes().unwrap());
    let error = Records::new(&parts["xl/worksheets/sheet1.bin"], Limits::DEFAULT)
        .map(Result::unwrap)
        .find(|r| r.kind() == kind::CELL_ERROR)
        .unwrap();
    assert_eq!(error.payload()[7] & 1, 1);
    assert_eq!(native(RICH), bytes);
}

#[test]
fn native_xlsb_malformed_metadata_prefixes_refuse_cell_publication_without_changing_source() {
    use litchi_xlsb::raw::{Encoder, Kind};
    let bytes = native(RICH);
    let parts = members(&bytes);
    let original = &parts["xl/worksheets/sheet1.bin"];
    let mut cases = Vec::new();
    for (payload, message) in [
        (vec![1, 0, 0], "cell metadata prefix has invalid framing"),
        (vec![0; 4], "cell metadata index must be positive"),
        (vec![0xff; 4], "cell metadata index must be positive"),
    ] {
        cases.push((payload, None, false, message));
    }
    cases.push((
        vec![1, 0, 0, 0],
        Some(kind::VALUE_META),
        false,
        "cell metadata prefix is duplicated or out of order",
    ));
    cases.push((
        vec![1, 0, 0, 0],
        Some(kind::CELL_META),
        false,
        "cell metadata prefix is duplicated or out of order",
    ));
    cases.push((
        vec![1, 0, 0, 0],
        None,
        true,
        "cell metadata prefix is not followed by a supported cell record",
    ));
    for (payload, extra, orphan, message) in cases {
        let mut changed = Vec::new();
        let mut writer = Encoder::new(&mut changed, Limits::DEFAULT);
        let mut after_metadata = false;
        for r in Records::new(original, Limits::DEFAULT) {
            let r = r.unwrap();
            if r.kind() == kind::VALUE_META {
                writer.write_record(kind::VALUE_META, &payload).unwrap();
                if let Some(k) = extra {
                    writer.write_record(k, &1u32.to_le_bytes()).unwrap();
                }
                after_metadata = true;
            } else {
                let k = if after_metadata && orphan {
                    Kind::new(0x1234).unwrap()
                } else {
                    r.kind()
                };
                writer.write_record(k, r.payload()).unwrap();
                after_metadata = false;
            }
        }
        let mut opc = litchi_opc::OpcPackage::from_reader(
            Cursor::new(&bytes),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        opc.get_part_mut(&litchi_opc::PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
            .unwrap()
            .set_blob(changed);
        let mut derivative = Vec::new();
        opc.write_to(&mut derivative).unwrap();
        let source = Package::from_bytes(&derivative).unwrap();
        let mut edit = source.edit_cell_values(0usize).unwrap();
        edit.insert(
            reference(120, 25),
            StyleIndex::new(0).unwrap(),
            Value::Number(88.0),
        )
        .unwrap();
        assert!(
            matches!(edit.commit(),Err(litchi_xlsb::cell_values::Error::InvalidFormat(m)) if m==message)
        );
        assert_eq!(source.to_bytes().unwrap(), derivative);
    }
    assert_eq!(native(RICH), bytes);
}
