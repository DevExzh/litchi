//! Focused tests for the layered XLS workbook owner.

use super::super::{OpenOptions, Workbook};
use super::codec::pivot_cache_stream_paths;
use crate::number_format::Formatting;
use crate::records::Encoding;
use litchi_biff::Records;
use litchi_core::sheet::Cell;
use std::io::Cursor;
use std::sync::Arc;

#[cfg(test)]
mod defined_name_tests {
    use super::{OpenOptions, Workbook};
    use crate::{BuiltInName, DefinedNameKind, NameScope};
    use std::fs::File;
    use std::path::{Path, PathBuf};

    fn poi_fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    }

    fn open(name: &str) -> Workbook<File> {
        Workbook::new(File::open(poi_fixture(name)).unwrap()).unwrap()
    }

    #[test]
    fn opens_poi_named_input_with_exact_names_and_ranges() {
        let workbook = open("namedinput.xls");
        assert_eq!(workbook.defined_names().len(), 2);
        let first = workbook.defined_name("namedrangename", Some(0)).unwrap();
        assert_eq!(first.name, "NamedRangeName");
        assert_eq!(first.scope, NameScope::Workbook);
        assert!(!first.hidden);
        assert!(first.formula.as_deref().unwrap().contains("$A$1:$D$10"));
        let second = workbook.defined_name("SECONDNAMEDRANGE", None).unwrap();
        assert!(second.formula.as_deref().unwrap().contains("$D$17:$G$27"));
    }

    #[test]
    fn recognizes_deleted_unicode_and_named_formula_fixtures() {
        let deleted = open("24207.xls");
        assert_eq!(deleted.defined_name("a", None).unwrap().name, "a");
        assert!(deleted.defined_name("b", None).unwrap().is_deleted());

        let unicode = open("unicodeNameRecord.xls");
        assert!(!unicode.defined_names().is_empty());

        for fixture in ["named-cell-test.xls", "named-cell-in-formula-test.xls"] {
            let workbook = open(fixture);
            assert!(!workbook.defined_names().is_empty());
        }
    }

    #[test]
    fn recognizes_poi_print_area() {
        let workbook = open("SimpleWithPrintArea.xls");
        let print_area = workbook.print_area(0).unwrap();
        assert_eq!(
            print_area.kind,
            DefinedNameKind::BuiltIn(BuiltInName::PrintArea)
        );
        assert!(print_area.formula.is_some());
    }

    #[test]
    fn lookup_prefers_last_local_name_then_last_global_name() {
        use crate::{DefinedName, DefinedNameKind};
        let mut workbook = open("namedinput.xls");
        let template = workbook.defined_names[0].clone();
        workbook.defined_names.extend([
            DefinedName {
                name: "Rate".to_string(),
                scope: NameScope::Workbook,
                ..template.clone()
            },
            DefinedName {
                name: "RATE".to_string(),
                scope: NameScope::Workbook,
                record_index: 20,
                ..template.clone()
            },
            DefinedName {
                name: "rate".to_string(),
                scope: NameScope::Worksheet(0),
                record_index: 21,
                kind: DefinedNameKind::User,
                ..template
            },
        ]);
        assert_eq!(
            workbook.defined_name("RaTe", None).unwrap().record_index,
            20
        );
        assert_eq!(
            workbook.defined_name("RaTe", Some(0)).unwrap().record_index,
            21
        );
        assert_eq!(
            workbook.defined_name("RaTe", Some(1)).unwrap().record_index,
            20
        );
    }

    #[test]
    fn parses_names_after_xor_decryption() {
        let file = File::open(poi_fixture("xor-encryption-abc.xls")).unwrap();
        let workbook =
            Workbook::new_with_options(file, OpenOptions::new().with_password("abc")).unwrap();
        assert!(workbook.defined_names().len() <= usize::from(u16::MAX));
    }
}

fn push_record(stream: &mut Vec<u8>, record_type: u16, data: &[u8]) {
    stream.extend_from_slice(&record_type.to_le_bytes());
    stream.extend_from_slice(&(data.len() as u16).to_le_bytes());
    stream.extend_from_slice(data);
}

fn dval_data(rule_count: u32) -> Vec<u8> {
    let mut data = vec![0; 10];
    data.extend_from_slice(&(-1i32).to_le_bytes());
    data.extend_from_slice(&rule_count.to_le_bytes());
    data
}

fn dimensions_data(first_row: u32, last_row: u32, first_col: u16, last_col: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&first_row.to_le_bytes());
    data.extend_from_slice(&last_row.to_le_bytes());
    data.extend_from_slice(&first_col.to_le_bytes());
    data.extend_from_slice(&last_col.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data
}

fn string_formula_data(row: u16, col: u16) -> Vec<u8> {
    let mut data = formula_data(row, col, &[]);
    data[6] = 0;
    data
}

fn string_formula_data_with_tokens(row: u16, col: u16, tokens: &[u8]) -> Vec<u8> {
    let mut data = formula_data(row, col, tokens);
    data[6] = 0;
    data
}

fn formula_data(row: u16, col: u16, tokens: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&row.to_le_bytes());
    data.extend_from_slice(&col.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[3, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(tokens);
    data
}

fn array_data(first_row: u16, last_row: u16, first_col: u8, last_col: u8) -> Vec<u8> {
    let tokens = [0x1e, 7, 0];
    let mut data = Vec::new();
    data.extend_from_slice(&first_row.to_le_bytes());
    data.extend_from_slice(&last_row.to_le_bytes());
    data.extend_from_slice(&[first_col, last_col]);
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(&tokens);
    data
}

#[test]
fn pivot_cache_stream_paths_are_sorted_without_unchecked_reparsing() {
    let paths = pivot_cache_stream_paths(vec![
        vec!["_SX_DB_CUR".into(), "0002".into()],
        vec!["_sx_db_cur".into(), "0001".into()],
        vec!["_SX_DB_CUR".into(), "not-hex".into()],
        vec!["_SX_DB_CUR".into()],
        vec!["Other".into(), "0003".into()],
    ]);

    assert_eq!(
        paths
            .iter()
            .map(|(stream_id, _)| *stream_id)
            .collect::<Vec<_>>(),
        vec![1, 2]
    );
    assert_eq!(paths[0].1[1], "0001");
    assert_eq!(paths[1].1[1], "0002");
}

#[test]
fn worksheet_expands_packed_numeric_and_blank_cells() {
    let mut mul_rk = Vec::new();
    mul_rk.extend_from_slice(&2u16.to_le_bytes());
    mul_rk.extend_from_slice(&4u16.to_le_bytes());
    mul_rk.extend_from_slice(&0u16.to_le_bytes());
    mul_rk.extend_from_slice(&((7u32 << 2) | 0x02).to_le_bytes());
    mul_rk.extend_from_slice(&1u16.to_le_bytes());
    mul_rk.extend_from_slice(&((250u32 << 2) | 0x03).to_le_bytes());
    mul_rk.extend_from_slice(&5u16.to_le_bytes());

    let mut mul_blank = Vec::new();
    mul_blank.extend_from_slice(&3u16.to_le_bytes());
    mul_blank.extend_from_slice(&1u16.to_le_bytes());
    mul_blank.extend_from_slice(&2u16.to_le_bytes());
    mul_blank.extend_from_slice(&3u16.to_le_bytes());
    mul_blank.extend_from_slice(&2u16.to_le_bytes());

    let mut stream = Vec::new();
    push_record(&mut stream, 0x00BD, &mul_rk);
    push_record(&mut stream, 0x00BE, &mul_blank);
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);

    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();

    assert!(matches!(
        worksheet.get_cell(2, 4).unwrap().value(),
        litchi_core::sheet::CellValue::Float(value) if *value == 7.0
    ));
    assert!(matches!(
        worksheet.get_cell(2, 5).unwrap().value(),
        litchi_core::sheet::CellValue::Float(value) if *value == 2.5
    ));
    assert!(worksheet.get_cell(3, 1).unwrap().is_empty());
    assert!(worksheet.get_cell(3, 2).unwrap().is_empty());
}

#[test]
fn worksheet_resolves_formula_value_from_following_string_record() {
    let mut string_data = Vec::new();
    string_data.extend_from_slice(&3u16.to_le_bytes());
    string_data.push(1);
    for code_unit in "文😀".encode_utf16() {
        string_data.extend_from_slice(&code_unit.to_le_bytes());
    }

    let mut stream = Vec::new();
    push_record(&mut stream, 0x0006, &string_formula_data(4, 5));
    push_record(&mut stream, 0x0207, &string_data);
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);

    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();
    let cell = worksheet.get_cell(4, 5).unwrap();

    assert!(cell.is_formula());
    assert!(matches!(
        cell.value(),
        litchi_core::sheet::CellValue::String(value) if value == "文😀"
    ));
}

#[test]
fn worksheet_rejects_formula_missing_its_string_record() {
    let mut stream = Vec::new();
    push_record(&mut stream, 0x0006, &string_formula_data(0, 0));
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);

    let result = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    );

    assert!(result.is_err());
}

#[test]
fn custom_view_closing_record_restores_primary_page_setup_collection() {
    let mut custom_view_begin = vec![0; 64];
    custom_view_begin[20..24].copy_from_slice(&100u32.to_le_bytes());

    let mut stream = Vec::new();
    push_record(
        &mut stream,
        crate::custom_view::USER_S_VIEW_BEGIN_RECORD_TYPE,
        &custom_view_begin,
    );
    // Custom-view content is consumed inertly by the worksheet parser.
    push_record(&mut stream, 0x0014, &[]);
    push_record(
        &mut stream,
        crate::custom_view::USER_S_VIEW_END_RECORD_TYPE,
        &1u16.to_le_bytes(),
    );
    // This record belongs to the worksheet's primary page-setup block.
    push_record(&mut stream, 0x0014, &[]);
    push_record(&mut stream, 0x000A, &[]);

    let mut records = Records::new(&stream);
    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();

    assert!(worksheet.page_setup().is_some());
}

#[test]
fn worksheet_resolves_string_formula_across_intervening_array_record() {
    // MS-XLS 2.1: FORMULA = Formula [Array / Table / ShrFmla / SUB]
    // [String *Continue], so an Array record may sit between the Formula
    // and its cached String result.
    let mut array = Vec::new();
    array.extend_from_slice(&4u16.to_le_bytes()); // rwFirst
    array.extend_from_slice(&4u16.to_le_bytes()); // rwLast
    array.push(5); // colFirst
    array.push(5); // colLast
    array.extend_from_slice(&[0; 6]); // reserved/options
    array.extend_from_slice(&3u16.to_le_bytes()); // cce
    array.extend_from_slice(&[0x1E, 0x01, 0x00]); // PtgInt 1

    let mut string_data = Vec::new();
    string_data.extend_from_slice(&5u16.to_le_bytes());
    string_data.push(0);
    string_data.extend_from_slice(b"array");

    let mut stream = Vec::new();
    push_record(&mut stream, 0x0200, &dimensions_data(4, 5, 5, 6));
    let anchor = [0x01, 4, 0, 5, 0];
    push_record(
        &mut stream,
        0x0006,
        &string_formula_data_with_tokens(4, 5, &anchor),
    );
    push_record(&mut stream, 0x0221, &array);
    push_record(&mut stream, 0x0207, &string_data);
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);

    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();
    let cell = worksheet.get_cell(4, 5).unwrap();

    assert!(cell.is_formula());
    assert!(matches!(
        cell.value(),
        litchi_core::sheet::CellValue::String(value) if value == "array"
    ));
    assert!(cell.is_array_formula());
    assert_eq!(worksheet.array_formulas().len(), 1);
}

#[test]
fn worksheet_resolves_string_formula_spanning_continue_records() {
    // The declared characters do not fit in the String record; the rest
    // arrive in a Continue record with its own option-flags byte.
    let mut string_data = Vec::new();
    string_data.extend_from_slice(&6u16.to_le_bytes());
    string_data.push(0);
    string_data.extend_from_slice(b"abc");

    let mut continues = Vec::new();
    continues.push(0); // fHighByte = 0 for this chunk
    continues.extend_from_slice(b"def");

    let mut stream = Vec::new();
    push_record(&mut stream, 0x0006, &string_formula_data(1, 2));
    push_record(&mut stream, 0x0207, &string_data);
    push_record(&mut stream, 0x003C, &continues);
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);

    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();
    let cell = worksheet.get_cell(1, 2).unwrap();

    assert!(cell.is_formula());
    assert!(matches!(
        cell.value(),
        litchi_core::sheet::CellValue::String(value) if value == "abcdef"
    ));
}

#[test]
fn worksheet_enforces_dval_dv_ordering() {
    let mut stream = Vec::new();
    push_record(
        &mut stream,
        super::super::data_validation::DVAL_RECORD_TYPE,
        &dval_data(0),
    );
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);
    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();
    assert_eq!(
        worksheet
            .data_validation_settings()
            .unwrap()
            .declared_rule_count(),
        0
    );
    assert!(worksheet.data_validations().is_empty());

    let mut stream = Vec::new();
    push_record(
        &mut stream,
        super::super::data_validation::DVAL_RECORD_TYPE,
        &dval_data(1),
    );
    push_record(&mut stream, 0x000A, &[]);
    let mut records = Records::new(&stream);
    assert!(
        Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            stream.len() as u64,
            0,
            0,
            &Encoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(Formatting::default()),
        )
        .is_err()
    );
}

#[test]
fn reads_data_validation_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let empty = Workbook::new(
        std::fs::File::open(root.join("test-data/poi/test-data/spreadsheet/dvEmpty.xls")).unwrap(),
    )
    .unwrap();
    let sheet = empty.xls_worksheet(0).unwrap();
    assert_eq!(
        sheet
            .data_validation_settings()
            .unwrap()
            .declared_rule_count(),
        0
    );
    assert!(sheet.data_validations().is_empty());

    let validation = Workbook::new(
        std::fs::File::open(
            root.join("test-data/libreoffice-core/sc/qa/unit/data/xls/validation.xls"),
        )
        .unwrap(),
    )
    .unwrap();
    let sheet = validation.xls_worksheet(0).unwrap();
    assert!(!sheet.data_validations().is_empty());
    for rule in sheet.data_validations() {
        assert!(
            rule.formula1().is_some()
                || rule.kind() == super::super::data_validation::DataValidationKind::Any
        );
        assert!(!rule.ranges().is_empty());
    }
    assert!(
        sheet
            .data_validations()
            .iter()
            .flat_map(|rule| rule.ranges())
            .any(|range| {
                range.first_row() <= 4
                    && range.last_row() >= 4
                    && range.first_column() <= 3
                    && range.last_column() >= 3
            })
    );
    assert!(
        sheet
            .data_validations()
            .iter()
            .flat_map(|rule| rule.ranges())
            .any(|range| {
                range.first_row() <= 8
                    && range.last_row() >= 8
                    && range.first_column() <= 5
                    && range.last_column() >= 5
            })
    );
}

#[test]
fn worksheet_rejects_malformed_optional_validation_records_and_continue() {
    for stream in [
        {
            let mut stream = Vec::new();
            push_record(
                &mut stream,
                super::super::data_validation::DVAL_RECORD_TYPE,
                &[0; 17],
            );
            push_record(&mut stream, 0x000A, &[]);
            stream
        },
        {
            let mut stream = Vec::new();
            push_record(
                &mut stream,
                super::super::data_validation::DVAL_RECORD_TYPE,
                &dval_data(1),
            );
            push_record(
                &mut stream,
                super::super::data_validation::DV_RECORD_TYPE,
                &[],
            );
            push_record(&mut stream, 0x000A, &[]);
            stream
        },
        {
            let mut stream = Vec::new();
            push_record(
                &mut stream,
                super::super::data_validation::DVAL_RECORD_TYPE,
                &dval_data(1),
            );
            push_record(&mut stream, 0x003C, &[0]);
            push_record(&mut stream, 0x000A, &[]);
            stream
        },
    ] {
        let mut records = Records::new(&stream);
        assert!(
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                &mut records,
                stream.len() as u64,
                0,
                0,
                &Encoding::Utf16Le,
                "Sheet1",
                Arc::new(Vec::new()),
                Arc::new(Vec::new()),
                None,
                Arc::new(Formatting::default()),
            )
            .is_err()
        );
    }
}

#[test]
fn worksheet_expands_shared_formula_relative_references() {
    let anchor = [0x01, 0, 0, 1, 0];
    let template = [
        0x4c, 0, 0, 0xff, 0xc0, // same row, previous column
        0x1e, 2, 0, 0x05, // * 2
    ];
    let mut shared = Vec::new();
    shared.extend_from_slice(&0u16.to_le_bytes());
    shared.extend_from_slice(&1u16.to_le_bytes());
    shared.extend_from_slice(&[1, 1, 0, 2]); // columns, reserved, cUse
    shared.extend_from_slice(&(template.len() as u16).to_le_bytes());
    shared.extend_from_slice(&template);

    let mut stream = Vec::new();
    let mut owner_formula = formula_data(0, 1, &anchor);
    owner_formula[14..16].copy_from_slice(&0x0008u16.to_le_bytes());
    push_record(&mut stream, 0x0006, &owner_formula);
    push_record(&mut stream, 0x04bc, &shared);
    let mut participant_formula = formula_data(1, 1, &anchor);
    participant_formula[14..16].copy_from_slice(&0x0008u16.to_le_bytes());
    push_record(&mut stream, 0x0006, &participant_formula);
    push_record(&mut stream, 0x000a, &[]);
    let mut records = Records::new(&stream);
    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();

    let first = worksheet.get_cell(0, 1).unwrap();
    let second = worksheet.get_cell(1, 1).unwrap();
    assert_eq!(first.formula(), Some("=(A1*2)"));
    assert_eq!(second.formula(), Some("=(A2*2)"));
    assert_eq!(first.formula_bytes(), Some(anchor.as_slice()));
    assert_eq!(second.formula_bytes(), Some(anchor.as_slice()));
}

#[test]
fn worksheet_resolves_array_formula_for_every_cell() {
    let anchor = [0x01, 0, 0, 2, 0];
    let template = [0x1e, 7, 0];
    let mut array = Vec::new();
    array.extend_from_slice(&0u16.to_le_bytes());
    array.extend_from_slice(&1u16.to_le_bytes());
    array.extend_from_slice(&[2, 2]);
    array.extend_from_slice(&0u16.to_le_bytes());
    array.extend_from_slice(&0u32.to_le_bytes());
    array.extend_from_slice(&(template.len() as u16).to_le_bytes());
    array.extend_from_slice(&template);

    let mut stream = Vec::new();
    push_record(&mut stream, 0x0200, &dimensions_data(0, 2, 2, 3));
    push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
    push_record(&mut stream, 0x0221, &array);
    push_record(&mut stream, 0x0006, &formula_data(1, 2, &anchor));
    push_record(&mut stream, 0x000a, &[]);
    let mut records = Records::new(&stream);
    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        stream.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();

    assert_eq!(worksheet.get_cell(0, 2).unwrap().formula(), Some("=7"));
    assert_eq!(worksheet.get_cell(1, 2).unwrap().formula(), Some("=7"));
    assert!(worksheet.get_cell(0, 2).unwrap().is_array_formula());
    assert!(worksheet.get_cell(1, 2).unwrap().is_array_formula());
    assert!(worksheet.array_formula_at(1, 2).is_some());
    assert_eq!(worksheet.array_formulas().len(), 1);
    let first = worksheet.get_cell(0, 2).unwrap();
    let second = worksheet.get_cell(1, 2).unwrap();
    let owner = worksheet.array_formulas().next().unwrap();
    assert!(std::ptr::eq(
        first.formula().unwrap(),
        second.formula().unwrap()
    ));
    assert!(std::ptr::eq(first.array_formula().unwrap(), owner));
    assert!(std::ptr::eq(second.array_formula().unwrap(), owner));
    assert!(std::ptr::eq(
        first.formula_metadata().unwrap().array_owner().unwrap(),
        owner,
    ));
}

#[test]
fn worksheet_rejects_orphan_and_incomplete_array_ptg_exp_links() {
    let anchor = [0x01, 0, 0, 2, 0];
    let orphan = {
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0200, &dimensions_data(0, 2, 2, 3));
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x000a, &[]);
        stream
    };

    let incomplete = {
        let template = [0x1e, 7, 0];
        let mut array = Vec::new();
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&1u16.to_le_bytes());
        array.extend_from_slice(&[2, 2]);
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&0u32.to_le_bytes());
        array.extend_from_slice(&(template.len() as u16).to_le_bytes());
        array.extend_from_slice(&template);
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x000a, &[]);
        stream
    };

    for stream in [orphan, incomplete] {
        let mut records = Records::new(&stream);
        assert!(
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                &mut records,
                stream.len() as u64,
                0,
                0,
                &Encoding::Utf16Le,
                "Sheet1",
                Arc::new(Vec::new()),
                Arc::new(Vec::new()),
                None,
                Arc::new(Formatting::default()),
            )
            .is_err()
        );
    }
}

#[test]
fn worksheet_rejects_array_without_dimensions_and_second_formula_companion() {
    let anchor = [0x01, 0, 0, 2, 0];
    let template = [0x1e, 7, 0];
    let mut array = Vec::new();
    array.extend_from_slice(&0u16.to_le_bytes());
    array.extend_from_slice(&0u16.to_le_bytes());
    array.extend_from_slice(&[2, 2]);
    array.extend_from_slice(&0u16.to_le_bytes());
    array.extend_from_slice(&0u32.to_le_bytes());
    array.extend_from_slice(&(template.len() as u16).to_le_bytes());
    array.extend_from_slice(&template);

    let without_dimensions = {
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x000a, &[]);
        stream
    };

    let with_second_companion = {
        let mut shared = Vec::new();
        shared.extend_from_slice(&0u16.to_le_bytes());
        shared.extend_from_slice(&0u16.to_le_bytes());
        shared.extend_from_slice(&[2, 2, 0, 1]);
        shared.extend_from_slice(&(template.len() as u16).to_le_bytes());
        shared.extend_from_slice(&template);
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0200, &dimensions_data(0, 1, 2, 3));
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x04bc, &shared);
        push_record(&mut stream, 0x000a, &[]);
        stream
    };

    let cross_kind_same_anchor = {
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0200, &dimensions_data(0, 1, 2, 3));
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0091, &[]);
        push_record(&mut stream, 0x000a, &[]);
        stream
    };

    for stream in [
        without_dimensions,
        with_second_companion,
        cross_kind_same_anchor,
    ] {
        let mut records = Records::new(&stream);
        assert!(
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                &mut records,
                stream.len() as u64,
                0,
                0,
                &Encoding::Utf16Le,
                "Sheet1",
                Arc::new(Vec::new()),
                Arc::new(Vec::new()),
                None,
                Arc::new(Formatting::default()),
            )
            .is_err()
        );
    }
}

#[test]
fn worksheet_rejects_formula_companions_without_an_adjacent_formula() {
    for record_type in [0x0221, crate::data_table::TABLE_RECORD_TYPE, 0x04bc, 0x0091] {
        let mut stream = Vec::new();
        push_record(&mut stream, record_type, &[]);
        push_record(&mut stream, 0x000a, &[]);
        let mut records = Records::new(&stream);
        assert!(
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                &mut records,
                stream.len() as u64,
                0,
                0,
                &Encoding::Utf16Le,
                "Sheet1",
                Arc::new(Vec::new()),
                Arc::new(Vec::new()),
                None,
                Arc::new(Formatting::default()),
            )
            .is_err()
        );
    }
}

#[test]
fn worksheet_indexes_many_array_ranges_and_rejects_indexed_overlap() {
    let mut nonoverlap = Vec::new();
    push_record(&mut nonoverlap, 0x0200, &dimensions_data(0, 64, 0, 8));
    for row in 0u16..64 {
        for col in 0u8..8 {
            let row_bytes = row.to_le_bytes();
            let anchor = [0x01, row_bytes[0], row_bytes[1], col, 0];
            push_record(
                &mut nonoverlap,
                0x0006,
                &formula_data(row, u16::from(col), &anchor),
            );
            push_record(&mut nonoverlap, 0x0221, &array_data(row, row, col, col));
        }
    }
    push_record(&mut nonoverlap, 0x000a, &[]);
    let mut records = Records::new(&nonoverlap);
    let worksheet = Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
        &mut records,
        nonoverlap.len() as u64,
        0,
        0,
        &Encoding::Utf16Le,
        "Sheet1",
        Arc::new(Vec::new()),
        Arc::new(Vec::new()),
        None,
        Arc::new(Formatting::default()),
    )
    .unwrap();
    assert_eq!(worksheet.array_formulas().len(), 512);

    let first_anchor = [0x01, 0, 0, 0, 0];
    let second_anchor = [0x01, 0, 0, 1, 0];
    let mut overlap = Vec::new();
    push_record(&mut overlap, 0x0200, &dimensions_data(0, 1, 0, 3));
    push_record(&mut overlap, 0x0006, &formula_data(0, 0, &first_anchor));
    push_record(&mut overlap, 0x0221, &array_data(0, 0, 0, 1));
    push_record(&mut overlap, 0x0006, &formula_data(0, 1, &first_anchor));
    push_record(&mut overlap, 0x0006, &formula_data(0, 1, &second_anchor));
    push_record(&mut overlap, 0x0221, &array_data(0, 0, 1, 2));
    push_record(&mut overlap, 0x000a, &[]);
    let mut records = Records::new(&overlap);
    assert!(
        Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            overlap.len() as u64,
            0,
            0,
            &Encoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(Formatting::default()),
        )
        .is_err()
    );
}
