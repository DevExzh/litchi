#![allow(
    clippy::cast_possible_truncation,
    clippy::expect_used,
    clippy::wildcard_enum_match_arm,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Focused workbook facade and package/codec integration tests.

use super::model::Workbook;
use crate::calc::Props;
use crate::external_link::Kind;
use crate::package::error::Result;
use crate::package::formula::{Compiler, Context, ExternalBook, Parser};
use crate::package::shared_strings::SharedString;
use crate::package::styles_table::StylesTable;
use crate::raw::{Kind as RawKind, Records, Writer, kind};
use litchi_core::sheet::{Cell, WorkbookTrait, Worksheet};
use litchi_ooxml_common::embedded::{Kind as EmbeddedKind, Target};
use litchi_ooxml_common::web;
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::fs::File;
use std::io::Cursor;
use std::sync::Arc;

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn external_link_records(records: &[(RawKind, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    for (record_type, payload) in records {
        writer.write_record(*record_type, payload).unwrap();
    }
    data
}

fn empty_workbook() -> Workbook {
    Workbook {
        package: OpcPackage::new(),
        worksheets: Vec::new(),
        worksheet_names: Vec::new(),
        worksheet_positions: Vec::new(),
        worksheet_rel_ids: Vec::new(),
        active_catalog_position: None,
        formula_context: Context::default(),
        shared_strings: Vec::new(),
        styles: StylesTable::default(),
        calc: Props::default(),
        is_1904: false,
        pivot_cache_definitions: Vec::new(),
        structured_tables: Vec::new(),
        chart_sheets: Vec::new(),
        sheet_drawings: Vec::new(),
        connections: None,
    }
}

fn generated_workbook() -> Workbook {
    let mut writer = crate::writer::WorkbookWriter::new();
    writer.add_worksheet(crate::writer::MutableWorksheet::new("Sheet1"));
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).unwrap();
    Workbook::new(Cursor::new(bytes.into_inner())).unwrap()
}

#[test]
fn workbook_ingress_honors_exact_read_limits() {
    let mut writer = crate::writer::WorkbookWriter::new();
    writer.add_worksheet(crate::writer::MutableWorksheet::new("Sheet1"));
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("serialize workbook");
    let bytes = output.into_inner();
    let input_bytes = u64::try_from(bytes.len()).expect("input length fits u64");
    let exact = crate::ReadLimits::builder()
        .max_input_bytes(input_bytes)
        .expect("exact input limit")
        .build()
        .expect("valid exact limit");
    let over = crate::ReadLimits::builder()
        .max_input_bytes(input_bytes - 1)
        .expect("smaller input limit")
        .build()
        .expect("valid smaller limit");

    assert!(Workbook::new(Cursor::new(bytes.clone())).is_ok());
    assert!(Workbook::new_with_limits(Cursor::new(bytes.clone()), exact).is_ok());
    assert!(Workbook::new_with_limits(Cursor::new(bytes), over).is_err());
}

#[test]
fn raw_opc_edit_publishes_a_reparsed_candidate() {
    let mut workbook = generated_workbook();
    let marker = PackURI::new("/xl/raw-edit-marker.bin").unwrap();

    let value = workbook
        .edit_opc(|package| {
            package.try_add_part(Box::new(BlobPart::new(
                marker.clone(),
                "application/octet-stream".to_string(),
                b"published".to_vec(),
            )))?;
            Ok::<_, crate::package::error::Error>("published")
        })
        .unwrap();

    assert_eq!(value, "published");
    assert_eq!(
        workbook.opc_package().get_part(&marker).unwrap().blob(),
        b"published"
    );
    assert_eq!(workbook.worksheet_names(), &["Sheet1".to_string()]);
}

#[test]
fn raw_opc_edit_rejects_invalid_workbook_without_publishing() {
    let mut workbook = generated_workbook();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = workbook
        .opc_package()
        .get_part(&workbook_uri)
        .unwrap()
        .blob_arc();

    let error = workbook
        .edit_opc(|candidate| {
            candidate.remove_part(&workbook_uri);
            Ok::<_, crate::package::error::Error>(())
        })
        .expect_err("an XLSB candidate without workbook.bin must be rejected");

    assert!(!error.to_string().is_empty());
    assert!(Arc::ptr_eq(
        &original,
        &workbook
            .opc_package()
            .get_part(&workbook_uri)
            .unwrap()
            .blob_arc()
    ));
    assert_eq!(workbook.worksheet_names(), &["Sheet1".to_string()]);
}

#[test]
fn task_pane_facade_round_trips_common_model() {
    let mut workbook = empty_workbook();
    let add_in = web::AddIn::new(
        "add-in-1",
        web::Reference::new("ref-1", "1", web::Store::Registry).unwrap(),
    )
    .unwrap();
    let mut panes = web::Panes::new();
    panes.push(web::Pane::new(add_in)).unwrap();

    workbook
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let loaded = workbook.task_panes().unwrap().unwrap();
    assert_eq!(loaded.get("add-in-1").unwrap().add_in().id(), "add-in-1");
    assert!(workbook.remove_task_panes().unwrap());
    assert!(workbook.task_panes().unwrap().is_none());
}

fn parse_external_link(records: &[(RawKind, Vec<u8>)]) -> Result<ExternalBook> {
    parse_external_link_with_relationship_type(records, Some(relationship_type::EXTERNAL_LINK_PATH))
}

fn parse_external_link_with_relationship_type(
    records: &[(RawKind, Vec<u8>)],
    target_relationship_type: Option<&str>,
) -> Result<ExternalBook> {
    let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
    let mut part = BlobPart::new(
        uri.clone(),
        "application/vnd.ms-excel.externalLink".to_string(),
        external_link_records(records),
    );
    if let Some(target_relationship_type) = target_relationship_type {
        part.rels_mut().add_relationship(
            target_relationship_type.to_string(),
            "Book.xlsx".to_string(),
            "rIdPath".to_string(),
            true,
        );
    }
    let mut package = OpcPackage::new();
    package.add_part(Box::new(part));
    let workbook = Workbook {
        package,
        worksheets: Vec::new(),
        worksheet_names: Vec::new(),
        worksheet_positions: Vec::new(),
        worksheet_rel_ids: Vec::new(),
        active_catalog_position: None,
        formula_context: Context::default(),
        shared_strings: Vec::new(),
        styles: StylesTable::default(),
        calc: Props::default(),
        is_1904: false,
        pivot_cache_definitions: Vec::new(),
        structured_tables: Vec::new(),
        chart_sheets: Vec::new(),
        sheet_drawings: Vec::new(),
        connections: None,
    };
    workbook.load_external_book(&uri)
}

fn parse_shared_string_records(records: &[(RawKind, Vec<u8>)]) -> Result<Vec<SharedString>> {
    let data = external_link_records(records);
    let mut iter = Records::new(data.as_slice());
    let mut strings = Vec::new();
    Workbook::read_shared_strings(&mut iter, &mut strings)?;
    Ok(strings)
}

#[test]
fn embedded_facade_accepts_binary_worksheet_sources() {
    let mut bundle_sheet = 0u32.to_le_bytes().to_vec();
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));

    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        external_link_records(&[(kind::BUNDLE_SH, bundle_sheet)]),
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_string(),
        "worksheets/sheet1.bin".to_string(),
        "rIdSheet1".to_string(),
        false,
    );

    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let mut sheet_part = BlobPart::new(
        sheet_uri.clone(),
        "application/vnd.ms-excel.worksheet".to_string(),
        external_link_records(&[
            (kind::BEGIN_SHEET, Vec::new()),
            (kind::END_SHEET, Vec::new()),
        ]),
    );
    sheet_part.rels_mut().add_relationship(
        relationship_type::OLE_OBJECT.to_string(),
        "../embeddings/oleObject1.bin".to_string(),
        "rIdObject".to_string(),
        false,
    );

    let payload = BlobPart::new(
        PackURI::new("/xl/embeddings/oleObject1.bin").unwrap(),
        content_type::OFC_OLE_OBJECT.to_string(),
        b"opaque XLSB payload".to_vec(),
    );
    let workbook_target = workbook_part
        .partname()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            .to_owned(),
        workbook_target,
        "rIdWorkbook".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    package.add_part(Box::new(payload));

    let workbook = Workbook::from_opc_package(package).unwrap();
    let entries = workbook.embedded().unwrap();

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].source(), &sheet_uri);
    assert_eq!(entries[0].id(), "rIdObject");
    assert_eq!(entries[0].kind(), EmbeddedKind::Object);
    let Target::Internal(payload) = entries[0].target() else {
        panic!("synthetic XLSB object must be internal")
    };
    assert_eq!(payload.part().as_str(), "/xl/embeddings/oleObject1.bin");
    assert_eq!(payload.content_type(), content_type::OFC_OLE_OBJECT);
    assert_eq!(payload.bytes(), b"opaque XLSB payload");
}

fn external_workbook_records() -> Vec<(RawKind, Vec<u8>)> {
    let mut begin = 0u16.to_le_bytes().to_vec();
    begin.extend_from_slice(&wide_string("rIdPath"));
    begin.extend_from_slice(&u32::MAX.to_le_bytes());
    let mut tabs = 1u32.to_le_bytes().to_vec();
    tabs.extend_from_slice(&wide_string("Data Sheet"));
    vec![
        (kind::BEGIN_SUP_BOOK, begin),
        (kind::SUP_TABS, tabs),
        (kind::SUP_NAME_START, wide_string("Rate")),
        (kind::SUP_NAME_FORMULA, 0u32.to_le_bytes().to_vec()),
        (kind::SUP_NAME_BITS, vec![0; 7]),
        (kind::SUP_NAME_END, Vec::new()),
        (kind::END_SUP_BOOK, Vec::new()),
    ]
}

fn external_data_source_records(
    kind: u16,
    source: &str,
    detail: &str,
    item_name: &str,
) -> Vec<(RawKind, Vec<u8>)> {
    assert!(matches!(kind, 1 | 2));
    let mut begin = kind.to_le_bytes().to_vec();
    begin.extend_from_slice(&wide_string(source));
    begin.extend_from_slice(&wide_string(detail));
    let mut bits = vec![0; 7];
    if kind == 2 {
        bits[0] = 1 << 4;
    }
    bits[6] = 1;
    vec![
        (kind::BEGIN_SUP_BOOK, begin),
        (kind::SUP_NAME_START, wide_string(item_name)),
        (kind::SUP_NAME_BITS, bits),
        (kind::SUP_NAME_END, Vec::new()),
        (kind::END_SUP_BOOK, Vec::new()),
    ]
}

#[test]
fn reads_formula_records_from_real_workbook_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/universal-content.xlsb"
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    let mut formula_cells = Vec::new();
    for index in 0..workbook.worksheet_names.len() {
        let worksheet = workbook.worksheet(index).unwrap();
        if let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let Some(cell) = worksheet.get_cell(row, col) else {
                        continue;
                    };
                    if cell.is_formula() {
                        formula_cells.push((
                            worksheet.name().to_string(),
                            cell.coordinate(),
                            cell.value().clone(),
                            cell.formula_bytes().unwrap().to_vec(),
                        ));
                    }
                }
            }
        }
    }
    assert_eq!(formula_cells.len(), 4);
    let formulas: Vec<_> = formula_cells
        .iter()
        .map(|cell| match &cell.2 {
            litchi_core::sheet::CellValue::Formula {
                formula,
                cached_value,
                ..
            } => (cell.1.as_str(), formula.as_str(), cached_value.as_deref()),
            value => panic!("expected decoded formula, found {value:?}"),
        })
        .collect();
    assert_eq!(formulas[0].0, "C1");
    assert_eq!(formulas[0].1, "(2*3)");
    assert_eq!(formulas[1].1, "(2+3)");
    assert_eq!(formulas[2].1, "(2-3)");
    assert_eq!(formulas[3].1, "(C1+C2)");
    assert!(matches!(
        formulas[3].2,
        Some(litchi_core::sheet::CellValue::Float(11.0))
    ));
    assert!(formula_cells.iter().all(|cell| !cell.3.is_empty()));
}

#[test]
fn reads_conditional_formatting_from_real_workbook_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/cond_format.xlsb"
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    let formatting = worksheet.conditional_formattings();
    assert_eq!(formatting.len(), 1);
    assert_eq!(formatting[0].ranges, ["E3:E18"]);
    assert!(!formatting[0].pivot_only);
    assert_eq!(formatting[0].rules.len(), 1);
    let rule = &formatting[0].rules[0];
    assert_eq!(
        rule.rule_type,
        crate::conditional_formatting::RuleType::CellIs
    );
    assert_eq!(rule.template, 0);
    assert_eq!(rule.dxf_id, Some(0));
    assert_eq!(rule.priority, 1);
    assert_eq!(rule.parameter, 5);
    assert_eq!(rule.formula_texts, ["5"]);
}

#[test]
fn reads_rich_and_phonetic_shared_strings_from_local_fixtures() {
    let rich_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/sample.xlsb"
    );
    let workbook = Workbook::new(File::open(rich_path).unwrap()).unwrap();
    let rich = workbook
        .shared_strings()
        .iter()
        .find(|value| !value.runs.is_empty())
        .expect("sample.xlsb should contain rich shared strings");
    assert_eq!(rich.text, "hello, xssf");
    assert_eq!(rich.runs[0].character_index, 0);
    let mut found_cell_text = false;
    for index in 0..workbook.worksheet_names.len() {
        let worksheet = workbook.worksheet(index).unwrap();
        let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() else {
            continue;
        };
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                found_cell_text |= worksheet.get_cell(row, col).is_some_and(|cell| {
                        matches!(cell.value(), litchi_core::sheet::CellValue::String(value) if value == "hello, xssf")
                    });
            }
        }
    }
    assert!(
        found_cell_text,
        "rich SST text should remain the cell value"
    );

    let phonetic_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/51519.xlsb"
    );
    let workbook = Workbook::new(File::open(phonetic_path).unwrap()).unwrap();
    let phonetic = workbook
        .shared_strings()
        .iter()
        .find_map(|value| {
            value
                .phonetic
                .as_ref()
                .filter(|value| !value.runs.is_empty())
        })
        .expect("51519.xlsb should contain phonetic shared strings");
    assert_eq!(phonetic.font_id, 1);
    assert_eq!(
        phonetic.phonetic_type,
        crate::package::PhoneticType::FullWidthKatakana
    );
    assert_eq!(phonetic.alignment, crate::package::PhoneticAlignment::Left);
}

#[test]
fn reads_binary_comments_from_real_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/comments.xlsb"
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    assert_eq!(worksheet.comments().len(), 4);
    let first = &worksheet.comments()[0];
    assert_eq!((first.row, first.col), (0, 0));
    assert_eq!(first.author, "Sven Nissel");
    assert!(first.text.contains("comment top row1"));
    assert_eq!(first.runs.len(), 2);
}

#[test]
fn validates_shared_string_stream_structure_and_counts() {
    let mut item = vec![0];
    item.extend_from_slice(&wide_string("value"));
    let valid = vec![
        (
            kind::BEGIN_SST,
            [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        ),
        (kind::SST_ITEM, item.clone()),
        (kind::END_SST, Vec::new()),
    ];
    let strings = parse_shared_string_records(&valid).unwrap();
    assert_eq!(strings[0].text, "value");

    let invalid_counts = vec![(
        kind::BEGIN_SST,
        [0u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
    )];
    assert!(matches!(
        parse_shared_string_records(&invalid_counts),
        Err(crate::package::error::Error::Unrecognized { .. })
    ));

    let missing_item = vec![
        (
            kind::BEGIN_SST,
            [2u32.to_le_bytes(), 2u32.to_le_bytes()].concat(),
        ),
        (kind::SST_ITEM, item),
        (kind::END_SST, Vec::new()),
    ];
    assert!(matches!(
        parse_shared_string_records(&missing_item),
        Err(crate::package::error::Error::Unrecognized { .. })
    ));

    let malformed_item = vec![
        (
            kind::BEGIN_SST,
            [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        ),
        (kind::SST_ITEM, vec![1]),
        (kind::END_SST, Vec::new()),
    ];
    assert!(parse_shared_string_records(&malformed_item).is_err());
}

#[test]
fn resolves_cell_style_references_from_real_fixtures() {
    let mut saw_nondefault_style = false;
    for fixture in [
        "Simple.xlsb",
        "date.xlsb",
        "universal-content.xlsb",
        "cond_format.xlsb",
    ] {
        let path = format!(
            "{}/../../test-data/ooxml/xlsb/{fixture}",
            env!("CARGO_MANIFEST_DIR")
        );
        let workbook = Workbook::new(File::open(path).unwrap())
            .unwrap_or_else(|error| panic!("{fixture}: {error}"));
        assert!(!workbook.styles().cell_xfs.is_empty(), "{fixture}");
        for index in 0..workbook.worksheet_names.len() {
            let worksheet = workbook.worksheet(index).unwrap();
            if let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() {
                for row in min_row..=max_row {
                    for col in min_col..=max_col {
                        let Some(cell) = worksheet.get_cell(row, col) else {
                            continue;
                        };
                        saw_nondefault_style |= cell.style_id() != 0;
                        assert!(workbook.style_for_cell(cell).is_some(), "{fixture}");
                    }
                }
            }
        }
    }
    assert!(saw_nondefault_style);
}

#[test]
fn opens_custom_number_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/62815.xlsb"
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    assert!(workbook.styles().num_fmts.keys().any(|id| *id >= 164));
    let worksheet = workbook.worksheet(0).unwrap();
    assert!(worksheet.dimensions().is_some());
}

#[test]
fn reads_external_book_metadata_from_local_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/bug66682.xlsb"
    );

    let package = OpcPackage::open(path).unwrap();
    let workbook = Workbook {
        package,
        worksheets: Vec::new(),
        worksheet_names: Vec::new(),
        worksheet_positions: Vec::new(),
        worksheet_rel_ids: Vec::new(),
        active_catalog_position: None,
        formula_context: Context::default(),
        shared_strings: Vec::new(),
        styles: StylesTable::default(),
        calc: Props::default(),
        is_1904: false,
        pivot_cache_definitions: Vec::new(),
        structured_tables: Vec::new(),
        chart_sheets: Vec::new(),
        sheet_drawings: Vec::new(),
        connections: None,
    };
    let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
    let book = workbook.load_external_book(&uri).unwrap();
    assert!(book.metadata.is_workbook());
    assert_eq!(book.metadata.source(), "ab");
    assert_eq!(book.metadata.sheet_names(), &["ab"]);
}

#[test]
fn parses_external_workbook_sheet_and_name_metadata() {
    let book = parse_external_link(&external_workbook_records()).unwrap();
    assert!(book.metadata.is_workbook());
    assert_eq!(book.metadata.source(), "Book.xlsx");
    assert_eq!(book.metadata.sheet_names(), &["Data Sheet"]);
    assert_eq!(book.metadata().defined_names()[0].name(), "Rate");

    let link = book.metadata();
    assert_eq!(link.kind(), Kind::Workbook);
    assert!(link.is_workbook());
    assert_eq!(link.source(), "Book.xlsx");
    assert_eq!(link.dde_topic(), None);
    assert_eq!(link.ole_program_id(), None);
    assert_eq!(link.sheet_names(), &["Data Sheet".to_string()]);
    assert_eq!(link.defined_names()[0].name(), "Rate");
}

#[test]
fn exposes_inert_dde_and_ole_link_metadata() {
    let dde_records = external_data_source_records(1, "Excel", "System", "RatesItem");
    let dde = parse_external_link_with_relationship_type(&dde_records, None)
        .unwrap()
        .metadata();
    assert_eq!(dde.kind(), Kind::Dde);
    assert!(!dde.is_workbook());
    assert_eq!(dde.source(), "Excel");
    assert_eq!(dde.dde_topic(), Some("System"));
    assert_eq!(dde.ole_program_id(), None);
    assert!(dde.sheet_names().is_empty());
    assert_eq!(dde.dde_items()[0].name(), "RatesItem");

    let ole_records = external_data_source_records(2, "rIdPath", "Acme.Server", "ReportItem");
    let ole = parse_external_link_with_relationship_type(
        &ole_records,
        Some(relationship_type::OLE_OBJECT),
    )
    .unwrap()
    .metadata();
    assert_eq!(ole.kind(), Kind::Ole);
    assert!(!ole.is_workbook());
    assert_eq!(ole.source(), "Book.xlsx");
    assert_eq!(ole.dde_topic(), None);
    assert_eq!(ole.ole_program_id(), Some("Acme.Server"));
    assert!(ole.sheet_names().is_empty());
    assert_eq!(ole.ole_items()[0].name(), "ReportItem");
}

#[test]
fn rejects_invalid_external_item_flags_and_cache_framing() {
    let mut invalid_dde = external_data_source_records(1, "Excel", "System", "StatusItem");
    invalid_dde
        .iter_mut()
        .find(|(record_type, _)| *record_type == kind::SUP_NAME_BITS)
        .unwrap()
        .1[6] = 0;
    assert!(matches!(
        parse_external_link_with_relationship_type(&invalid_dde, None),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));

    let mut truncated_cache =
        external_data_source_records(2, "rIdPath", "Acme.Server", "ReportItem");
    let end = truncated_cache.len() - 2;
    truncated_cache.splice(
        end..end,
        [
            (
                kind::SUP_NAME_VALUE_START,
                [1u32.to_le_bytes(), 2u32.to_le_bytes()].concat(),
            ),
            (kind::SUP_NAME_NUM, 1.0f64.to_le_bytes().to_vec()),
            (kind::SUP_NAME_VALUE_END, Vec::new()),
        ],
    );
    assert!(matches!(
        parse_external_link_with_relationship_type(
            &truncated_cache,
            Some(relationship_type::OLE_OBJECT),
        ),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));
}

#[test]
fn validates_external_link_relationship_types() {
    assert!(matches!(
        parse_external_link_with_relationship_type(
            &external_workbook_records(),
            Some(relationship_type::OLE_OBJECT),
        ),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));

    let dde_records = external_data_source_records(1, "Excel", "System", "Rates");
    assert!(matches!(
        parse_external_link_with_relationship_type(
            &dde_records,
            Some(relationship_type::EXTERNAL_LINK_PATH),
        ),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));

    let ole_records = external_data_source_records(2, "rIdPath", "Acme.Server", "Report");
    assert!(matches!(
        parse_external_link_with_relationship_type(
            &ole_records,
            Some(relationship_type::EXTERNAL_LINK_PATH),
        ),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));
}

#[test]
fn resolves_external_formula_tokens_from_package_relationships() {
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let mut workbook_data = Vec::new();
    {
        let mut writer = Writer::new(&mut workbook_data);
        writer
            .write_record(kind::SUP_BOOK_SRC, &wide_string("rIdExternal"))
            .unwrap();
        let mut extern_sheet = 1u32.to_le_bytes().to_vec();
        extern_sheet.extend_from_slice(&0u32.to_le_bytes());
        extern_sheet.extend_from_slice(&0u32.to_le_bytes());
        extern_sheet.extend_from_slice(&0u32.to_le_bytes());
        writer
            .write_record(kind::EXTERN_SHEET, &extern_sheet)
            .unwrap();
    }
    let mut workbook_part = BlobPart::new(
        workbook_uri,
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        workbook_data,
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::EXTERNAL_LINK.to_string(),
        "externalLinks/externalLink1.bin".to_string(),
        "rIdExternal".to_string(),
        false,
    );

    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
    let mut external_part = BlobPart::new(
        external_uri,
        "application/vnd.ms-excel.externalLink".to_string(),
        external_link_records(&external_workbook_records()),
    );
    external_part.rels_mut().add_relationship(
        relationship_type::EXTERNAL_LINK_PATH.to_string(),
        "Book.xlsx".to_string(),
        "rIdPath".to_string(),
        true,
    );

    let workbook_target = workbook_part
        .partname()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            .to_owned(),
        workbook_target,
        "rIdWorkbook".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(external_part));
    let workbook = Workbook::from_opc_package(package).unwrap();

    let links = workbook.external_links();
    assert_eq!(links.len(), 1);
    let link = &links[0];
    assert_eq!(link.kind(), Kind::Workbook);
    assert_eq!(link.source(), "Book.xlsx");
    assert_eq!(link.sheet_names(), &["Data Sheet".to_string()]);
    assert_eq!(link.defined_names()[0].name(), "Rate");

    let reference = Parser::new(&[0x5A, 0, 0, 0, 0, 0, 0, 0, 0])
        .parse()
        .unwrap();
    assert_eq!(
        Compiler::try_tokens_to_string_with_resolution(&reference, &workbook.formula_context)
            .unwrap(),
        "'[Book.xlsx]Data Sheet'!$A$1"
    );
    let name = Parser::new(&[0x59, 0, 0, 1, 0, 0, 0]).parse().unwrap();
    assert_eq!(
        Compiler::try_tokens_to_string_with_resolution(&name, &workbook.formula_context).unwrap(),
        "'[Book.xlsx]'!Rate"
    );
}

#[test]
fn rejects_malformed_external_workbook_record_sequences() {
    let mut duplicate_tabs = external_workbook_records();
    duplicate_tabs.insert(2, duplicate_tabs[1].clone());
    assert!(matches!(
        parse_external_link(&duplicate_tabs),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));

    let mut unclosed_name = external_workbook_records();
    unclosed_name.remove(5);
    assert!(matches!(
        parse_external_link(&unclosed_name),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));

    let mut trailing_record = external_workbook_records();
    trailing_record.push((kind::SUP_NAME_END, Vec::new()));
    assert!(matches!(
        parse_external_link(&trailing_record),
        Err(crate::package::error::Error::InvalidFormula(_))
    ));
}

#[test]
fn loads_typed_pivot_cache_definitions_from_package_relationships() {
    // workbook.bin declares one PivotCache (idSx 12) related to a
    // pivotCacheDefinition part.
    let mut cache_id = 12u32.to_le_bytes().to_vec();
    cache_id.extend_from_slice(&wide_string("rIdCache"));
    let workbook_data = external_link_records(&[
        (kind::BEGIN_PIVOT_CACHE_IDS, Vec::new()),
        (kind::BEGIN_PIVOT_CACHE_ID, cache_id),
        (kind::END_PIVOT_CACHE_ID, Vec::new()),
        (kind::END_PIVOT_CACHE_IDS, Vec::new()),
    ]);
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let mut workbook_part = BlobPart::new(
        workbook_uri,
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        workbook_data,
    );
    workbook_part.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
            .to_string(),
        "pivotCache/pivotCacheDefinition1.bin".to_string(),
        "rIdCache".to_string(),
        false,
    );

    // Minimal worksheet-range PivotCache definition stream.
    let mut definition = vec![
        3,           // bVerCacheLastRefresh
        0,           // bVerCacheRefreshableMin
        2,           // bVerCacheCreated
        0b0001_0001, // fSaveData | fEnableRefresh
    ];
    definition.extend_from_slice(&(-1i32).to_le_bytes()); // citmGhostMax
    definition.extend_from_slice(&44_000.0f64.to_le_bytes()); // xnumRefreshedDate
    definition.push(0x00); // no optional strings
    definition.extend_from_slice(&5u32.to_le_bytes()); // cRecords
    definition.extend_from_slice(&[0; 4]); // unused (fLoadRefreshedWho = 0)
    let mut source = Vec::new();
    source.extend_from_slice(&0u32.to_le_bytes()); // iSrcType = sheet
    source.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
    let mut range = vec![0x00, 0x00, 0b0000_0010]; // fLoadSheet
    range.extend_from_slice(&wide_string("Data"));
    for value in [0i32, 9, 0, 3] {
        range.extend_from_slice(&value.to_le_bytes());
    }
    let definition_part = BlobPart::new(
        PackURI::new("/xl/pivotCache/pivotCacheDefinition1.bin").unwrap(),
        "application/vnd.ms-excel.pivotCacheDefinition".to_string(),
        external_link_records(&[
            (kind::BEGIN_PIVOT_CACHE_DEF, definition),
            (kind::BEGIN_PCD_SOURCE, source),
            (kind::BEGIN_PCDS_RANGE, range),
            (kind::END_PCDS_RANGE, Vec::new()),
            (kind::END_PCD_SOURCE, Vec::new()),
            (kind::END_PIVOT_CACHE_DEF, Vec::new()),
        ]),
    );

    let workbook_target = workbook_part
        .partname()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            .to_owned(),
        workbook_target,
        "rIdWorkbook".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(definition_part));
    let workbook = Workbook::from_opc_package(package).unwrap();

    let definitions = workbook.pivot_cache_definitions();
    assert_eq!(definitions.len(), 1);
    assert_eq!(definitions[0].0, 12);
    let definition = workbook.pivot_cache_definition(12).unwrap();
    assert!(definition.save_data);
    assert_eq!(definition.record_count, 5);
    let source = definition.source.as_ref().unwrap();
    assert_eq!(
        source.source_type,
        crate::package::pivot::PivotCacheSourceType::Worksheet
    );
    let worksheet = source.worksheet.as_ref().unwrap();
    assert_eq!(worksheet.sheet_name.as_deref(), Some("Data"));
    assert_eq!(
        worksheet.range,
        Some(crate::package::pivot::PivotCacheRange {
            first_row: 0,
            last_row: 9,
            first_column: 0,
            last_column: 3,
        })
    );
    assert!(workbook.pivot_cache_definition(99).is_none());
}
