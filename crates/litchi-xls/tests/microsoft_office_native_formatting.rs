//! Native styles, rich strings and merged cells; no rendering or evaluation.
use litchi_core::sheet::{Cell as _, WorkbookTrait, Worksheet as _};
use litchi_xls::{Package, merged_cells::MergedCellRange, raw::Reader};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut file =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    file.list_streams()
        .into_iter()
        .map(|p| {
            let refs = p.iter().map(String::as_str).collect::<Vec<_>>();
            let b = file.open_stream(&refs).unwrap();
            (p, b)
        })
        .collect()
}
fn globals(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let all = streams(bytes);
    let workbook = &all[&vec!["Workbook".into()]];
    let mut pos = 0;
    let mut out = Vec::new();
    loop {
        let kind = u16::from_le_bytes(workbook[pos..pos + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            workbook[pos + 2..pos + 4].try_into().unwrap(),
        ));
        let payload = workbook[pos + 4..pos + 4 + len].to_vec();
        out.push((kind, payload));
        pos += 4 + len;
        if kind == 0x000a {
            break;
        }
    }
    out
}

// Independent bit-at-a-time MS-OSHARED 2.4.3 oracle (no production CRC calls).
fn assert_xfcrc(records: &[(u16, Vec<u8>)]) -> (usize, u32) {
    let mut checksum = 0_u32;
    let mut count = 0;
    for (_, payload) in records.iter().filter(|(kind, _)| *kind == 0x00e0) {
        count += 1;
        for byte in payload {
            let mut cache = ((checksum >> 24) ^ u32::from(*byte)) << 24;
            for _ in 0..8 {
                cache = (cache << 1) ^ if cache & 0x8000_0000 != 0 { 0xaf } else { 0 };
            }
            checksum = (checksum << 8) ^ (cache & 0xffff);
        }
    }
    let crcs = records
        .iter()
        .filter(|(kind, _)| *kind == 0x087c)
        .collect::<Vec<_>>();
    assert_eq!(crcs.len(), 1);
    let data = &crcs[0].1;
    assert_eq!(data.len(), 20);
    assert_eq!(
        usize::from(u16::from_le_bytes(data[14..16].try_into().unwrap())),
        count
    );
    assert_eq!(
        u32::from_le_bytes(data[16..20].try_into().unwrap()),
        checksum
    );
    (count, checksum)
}

#[derive(Debug, PartialEq)]
struct CellSnapshot {
    sheet: String,
    row: u32,
    column: u32,
    value: litchi_core::sheet::CellValue,
    formula: Option<String>,
}
fn values(bytes: &[u8]) -> Vec<CellSnapshot> {
    let book = Reader::new(Cursor::new(bytes)).unwrap();
    let mut out = Vec::new();
    for name in book.worksheet_names() {
        for c in book.worksheet_by_name(name).unwrap().cells() {
            let c = c.unwrap();
            out.push(CellSnapshot {
                sheet: name.clone(),
                row: c.row(),
                column: c.column(),
                value: c.value().clone(),
                formula: c.formula().map(str::to_owned),
            });
        }
    }
    out
}
#[test]
fn native_merge_anchor_resolves_effective_font_alignment_and_number_formats() {
    let bytes = native("excel-values-formulas-chart.xls");
    let book = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = book.worksheet_by_name("Sheet2").unwrap();
    let merge = MergedCellRange {
        first_row: 0,
        last_row: 0,
        first_col: 0,
        last_col: 2,
    };
    assert_eq!(sheet.merged_cells(), [merge]);
    assert!(merge.contains(0, 1));
    assert!(!merge.contains(1, 0));
    let cell = sheet.get_cell(0, 0).unwrap();
    assert_eq!(
        cell.value(),
        &litchi_core::sheet::CellValue::String("Native workbook summary".into())
    );
    let style = book.effective_extended_format(cell.xf_index()).unwrap();
    let font = book.font(style.font_index()).unwrap();
    assert_eq!(font.name(), "Aptos Narrow");
    assert_eq!(font.height_twips(), 240);
    assert!(font.is_bold() && !font.is_italic());
    assert_eq!(
        style.alignment().horizontal(),
        litchi_xls::HorizontalAlignment::Center
    );
    let data = book.worksheet_by_name("Sheet1").unwrap();
    for (row, col, id) in [
        (12, 1, 14),
        (12, 2, 14),
        (13, 1, 18),
        (13, 2, 21),
        (14, 1, 9),
        (14, 2, 10),
    ] {
        let c = data.get_cell(row, col).unwrap();
        assert_eq!(
            book.effective_extended_format(c.xf_index())
                .unwrap()
                .number_format_id(),
            id
        );
    }
    let quoted = data.get_cell(19, 1).unwrap();
    assert!(
        book.effective_extended_format(quoted.xf_index())
            .unwrap()
            .direct()
            .quote_prefix()
    );
    assert_eq!(
        quoted.value(),
        &litchi_core::sheet::CellValue::String("00123".into())
    );
    // Parsed worksheets are immutable views; the create-only method must not
    // masquerade as an edit of the native merge records.
    let mut copied = sheet.clone();
    assert!(matches!(
        copied.add_merged_cells(&[MergedCellRange {
            first_row: 9,
            last_row: 9,
            first_col: 0,
            last_col: 2
        }]),
        Err(litchi_xls::Error::SourceBoundWorksheetMutation {
            operation: "add_merged_cells"
        })
    ));
    assert!(Package::from_bytes(&bytes).unwrap().to_bytes().unwrap() == bytes);
}
#[test]
fn native_rich_sst_runs_resolve_utf16_boundaries_and_logical_font_ids() {
    let bytes = native("excel-metadata.xls");
    let book = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = book.worksheet_by_name("Sheet1").unwrap();
    let c = sheet.get_cell(1, 0).unwrap();
    assert_eq!(
        c.value(),
        &litchi_core::sheet::CellValue::String("RICH_TEXT_015 bold italic".into())
    );
    let rich = book.shared_string_properties_for_cell(c).unwrap();
    assert!(rich.phonetic.is_none());
    assert_eq!(
        rich.formatting_runs
            .iter()
            .map(|r| (r.character_index, r.font_index))
            .collect::<Vec<_>>(),
        [(14, 8), (18, 0), (19, 7)]
    );
    assert!(book.font(4).is_none());
    for (run, (bold, italic)) in
        rich.formatting_runs
            .iter()
            .zip([(true, false), (false, false), (false, true)])
    {
        let f = book.font(run.font_index).unwrap();
        assert_eq!((f.name(), f.height_twips()), ("Aptos Narrow", 240));
        assert_eq!((f.is_bold(), f.is_italic()), (bold, italic));
    }
}
#[test]
fn native_alignment_edit_preserves_merge_cells_values_and_other_cfb_streams() {
    let bytes = native("excel-values-formulas-chart.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let before_records = globals(&bytes);
    assert_eq!(assert_xfcrc(&before_records), (69, 0x82cd_d373));
    let before_values = values(&bytes);
    let mut change = litchi_xls::cell_values::FormatChange::default();
    change.horizontal = Some(litchi_xls::HorizontalAlignment::Right);
    change.wrap_text = Some(true);
    let mut edit = p.edit().unwrap();
    edit.set_cell_format(
        "Sheet2",
        litchi_xls::cell_values::Reference::new(1, 0).unwrap(),
        &change,
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    assert_eq!(values(&after), before_values);
    let after_records = globals(&after);
    assert_eq!(assert_xfcrc(&after_records).0, 70);
    let xfs = |records: &[(u16, Vec<u8>)]| {
        records
            .iter()
            .filter(|(kind, _)| *kind == 0x00e0)
            .map(|(_, data)| data.clone())
            .collect::<Vec<_>>()
    };
    assert_eq!(xfs(&before_records), xfs(&after_records)[..69]);
    let extensions = |records: &[(u16, Vec<u8>)]| {
        records
            .iter()
            .filter(|(kind, _)| *kind == 0x087d)
            .cloned()
            .collect::<Vec<_>>()
    };
    assert!(!extensions(&before_records).is_empty());
    assert_eq!(extensions(&before_records), extensions(&after_records));
    let before = Reader::new(Cursor::new(&bytes)).unwrap();
    let read = Reader::new(Cursor::new(&after)).unwrap();
    let sheet = read.worksheet_by_name("Sheet2").unwrap();
    let style = read
        .effective_extended_format(sheet.get_cell(1, 0).unwrap().xf_index())
        .unwrap();
    assert_eq!(
        style.alignment().horizontal(),
        litchi_xls::HorizontalAlignment::Right
    );
    assert!(style.alignment().wraps_text());
    // The existing FormatChange contract authors a palette-based XF. It
    // retains the source extension records but does not clone their theme
    // overrides onto the new style; this test makes no rendering claim.
    assert!(!style.has_xf_extension());
    assert_eq!(
        sheet.merged_cells(),
        before.worksheet_by_name("Sheet2").unwrap().merged_cells()
    );
    assert_eq!(read.fonts(), before.fonts());
    let original = streams(&bytes);
    let changed = streams(&after);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    for (name, data) in original {
        if name != ["Workbook"] {
            assert!(changed[&name] == data, "changed {name:?}");
        }
    }
    let reopened = Package::from_bytes(&after).unwrap();
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

#[test]
fn fresh_xf_extension_writer_checksums_actual_format_payloads() {
    use litchi_xls::{
        ExtProp, XfExt,
        writer::{Builder, ExtendedFormat},
    };
    let mut writer = Builder::new();
    let sheet = writer.add_worksheet("Styles").unwrap();
    writer
        .set_xf_extensions(vec![XfExt::try_new(15, vec![ExtProp::Indent(2)]).unwrap()])
        .unwrap();
    for extra in 0..2 {
        if extra != 0 {
            writer.add_cell_format(ExtendedFormat::default()).unwrap();
        }
        writer.write_number(sheet, 0, 0, 42.).unwrap();
        let mut out = Cursor::new(Vec::new());
        writer.build().unwrap().write_to(&mut out).unwrap();
        assert_eq!(assert_xfcrc(&globals(&out.into_inner())).0, 21 + extra);
    }
}
