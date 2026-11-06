//! Native styles, rich strings and merge lifecycle; no rendering or evaluation.
use litchi_xlsb::Package;
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xlsb::merged_cells::MergedCell;
#[test]
fn native_merge_anchor_and_formats_resolve_the_binary_style_table() {
    let bytes = native("excel-values-formulas-chart.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let b = p.workbook().unwrap();
    let s = b.worksheet(1).unwrap();
    assert_eq!(
        b.merged_cell_ranges_by_name("Sheet2").unwrap(),
        [MergedCell::new(0, 0, 0, 2)]
    );
    let c = s.get_cell(0, 0).unwrap();
    assert_eq!(
        c.value(),
        &CellValue::String("Native workbook summary".into())
    );
    let f = b.style_for_cell(c).unwrap();
    let font = b.styles().get_font(f.font_id as usize).unwrap();
    assert_eq!(
        (&*font.name, font.size, font.bold, font.italic),
        ("Aptos Narrow", 12., true, false)
    );
    assert_eq!(
        f.alignment.as_ref().unwrap().horizontal,
        litchi_xlsb::package::styles::HorizontalAlignment::Center
    );
    let data = b.worksheet(0).unwrap();
    for (row, col, id) in [
        (12, 1, 14),
        (12, 2, 14),
        (13, 1, 18),
        (13, 2, 21),
        (14, 1, 9),
        (14, 2, 10),
    ] {
        assert_eq!(
            b.style_for_cell(data.get_cell(row, col).unwrap())
                .unwrap()
                .num_fmt_id,
            id
        );
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_rich_sst_runs_resolve_exact_font_changes() {
    let p = Package::from_bytes(&native("excel-metadata.xlsb")).unwrap();
    let b = p.workbook().unwrap();
    let s = b.worksheet(0).unwrap();
    let c = s.get_cell(1, 0).unwrap();
    assert_eq!(
        c.value(),
        &CellValue::String("RICH_TEXT_015 bold italic".into())
    );
    let strings = b
        .shared_strings()
        .iter()
        .filter(|s| s.text == "RICH_TEXT_015 bold italic")
        .collect::<Vec<_>>();
    assert_eq!(strings.len(), 1);
    let rich = strings[0];
    assert_eq!(
        rich.runs
            .iter()
            .map(|r| (r.character_index, r.font_id))
            .collect::<Vec<_>>(),
        [(14, 4), (18, 0), (19, 3)]
    );
    assert!(rich.phonetic.is_none());
    for (run, (bold, italic)) in
        rich.runs
            .iter()
            .zip([(true, false), (false, false), (false, true)])
    {
        let f = b.styles().get_font(run.font_id as usize).unwrap();
        assert_eq!(
            (&*f.name, f.size, f.bold, f.italic),
            ("Aptos Narrow", 12., bold, italic)
        );
    }
}
#[test]
fn native_merge_remove_create_overlap_refusal_and_inverse_preserve_cell_records() {
    let bytes = native("excel-values-formulas-chart.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    assert_eq!(
        edit.unmerge_cells("Sheet2", MergedCell::new(0, 0, 0, 2))
            .unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.merge_cells("Sheet2", MergedCell::new(9, 9, 0, 2))
            .unwrap(),
        Some(true)
    );
    assert!(
        matches!(edit.merge_cells("Sheet2",MergedCell::new(9,9,1,3)),Err(litchi_xlsb::package::PackageError::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::InvalidValue)
    );
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let read = Package::from_bytes(&after).unwrap();
    assert_eq!(
        read.workbook()
            .unwrap()
            .merged_cell_ranges_by_name("Sheet2")
            .unwrap(),
        [MergedCell::new(9, 9, 0, 2)]
    );
    let a = members(&bytes);
    let b = members(&after);
    let changed = a
        .iter()
        .filter(|(n, v)| b[*n] != **v)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    assert_eq!(changed, ["xl/worksheets/sheet2.bin"]);
    let outside = |data: &[u8]| {
        litchi_xlsb::raw::Records::new(data, litchi_xlsb::raw::Limits::DEFAULT)
            .map(|r| r.unwrap())
            .filter(|r| r.kind().get() != 176)
            .map(|r| (r.kind().get(), r.payload().to_vec()))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        outside(&a["xl/worksheets/sheet2.bin"]),
        outside(&b["xl/worksheets/sheet2.bin"])
    );
    assert!(
        read.apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
#[test]
fn native_rich_string_copy_preserves_run_fonts_and_other_parts() {
    use litchi_xlsb::cell_values::Reference;
    let bytes = native("excel-metadata.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let book = p.workbook().unwrap();
    let rich = book
        .shared_strings()
        .iter()
        .find(|s| s.text == "RICH_TEXT_015 bold italic")
        .unwrap()
        .clone();
    // This source cell is BrtCellIsst. The existing-cell setter owns only
    // BrtCellRString and must refuse a storage-family conversion atomically.
    let mut edit = p.edit_cell_values("Sheet1").unwrap();
    assert!(
        matches!(edit.set_rich_string(Reference::new(1, 0).unwrap(), rich.clone()),
        Err(litchi_xlsb::package::PackageError::Refused(r))
        if r.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
    );
    assert!(edit.commit().unwrap().package().to_bytes().unwrap() == bytes);
    let mut edit = p.edit().unwrap();
    edit.transfer_cell(
        book,
        0,
        Reference::new(1, 0).unwrap(),
        0,
        Reference::new(99, 25).unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let read = Package::from_bytes(&after).unwrap();
    let wb = read.workbook().unwrap();
    let sheet = wb.worksheet(0).unwrap();
    for (row, col) in [(1, 0), (99, 25)] {
        assert_eq!(
            sheet.get_cell(row, col).unwrap().value(),
            &CellValue::String(rich.text.clone())
        );
    }
    assert_eq!(wb.shared_strings(), book.shared_strings());
    assert_eq!(wb.styles().fonts, book.styles().fonts);
    let a = members(&bytes);
    let b = members(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let changed = a
        .iter()
        .filter(|(n, v)| b[*n] != **v)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        changed,
        [
            "xl/sharedStrings.bin",
            "xl/worksheets/binaryIndex1.bin",
            "xl/worksheets/sheet1.bin"
        ]
    );
    assert!(
        read.apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
