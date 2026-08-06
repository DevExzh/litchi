use super::model::{AutoFilter, ColumnInfo, RowInfo};
use super::*;
use crate::conditional_formatting::Formatting;
use crate::package::comments::Record;
use crate::package::data_validation::Validation;
use crate::package::error::Error;
use crate::package::formula::{Group, GroupKind, ParsedFormula, Range};
use crate::package::hyperlinks::Hyperlink;
use crate::package::merged_cells::MergedCell;
use crate::package::web_extension_bindings::Binding;
use crate::raw::Records;
use crate::raw::{Writer, kind};
use litchi_core::binary;
use litchi_core::sheet::CellValue;
use std::io::Cursor;

#[test]
fn test_set_and_get_cell() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Hello");
    sheet.set_cell(1, 1, 42.0);

    assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Hello"));
    assert_eq!(sheet.get_cell(1, 1).and_then(|v| v.as_float()), Some(42.0));
}

#[test]
fn writes_worksheet_web_extension_collection() {
    let formula = ParsedFormula {
        rgce: vec![0x3B, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 1, 0],
        rgcb: Vec::new(),
    };
    let binding = Binding::new("sales-table", formula, |index| index == 0).unwrap();
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet
        .set_web_extension_bindings(vec![binding.clone()])
        .unwrap();
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet.write(&mut writer, &mut shared_strings).unwrap();

    let records = Records::new(&buffer)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    let begin = records
        .iter()
        .position(|record| record.kind() == kind::BEGIN_WEB_EXTENSIONS)
        .unwrap();
    assert!(records[begin].payload().is_empty());
    assert_eq!(records[begin + 1].kind(), kind::WEB_EXTENSION);
    assert_eq!(
        Binding::parse_payload(records[begin + 1].payload(), |index| index == 0).unwrap(),
        binding
    );
    assert_eq!(records[begin + 2].kind(), kind::END_WEB_EXTENSIONS);
    assert!(records[begin + 2].payload().is_empty());
}

#[test]
fn test_set_cell_with_style() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell_with_style(0, 0, "Styled", 5);

    assert_eq!(
        sheet.get_cell(0, 0).and_then(|v| v.as_str()),
        Some("Styled")
    );
    assert_eq!(sheet.cell_count(), 1);
}

#[test]
fn test_delete_cell() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Hello");

    assert!(sheet.delete_cell(0, 0).is_some());
    assert!(sheet.get_cell(0, 0).is_none());
}

#[test]
fn test_delete_row() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Row 0");
    sheet.set_cell(1, 0, "Row 1");
    sheet.set_cell(2, 0, "Row 2");

    sheet.delete_row(1);

    assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Row 0"));
    assert_eq!(sheet.get_cell(1, 0).and_then(|v| v.as_str()), Some("Row 2"));
    assert!(sheet.get_cell(2, 0).is_none());
}

#[test]
fn test_delete_column() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Col 0");
    sheet.set_cell(0, 1, "Col 1");
    sheet.set_cell(0, 2, "Col 2");

    sheet.delete_column(1);

    assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Col 0"));
    assert_eq!(sheet.get_cell(0, 1).and_then(|v| v.as_str()), Some("Col 2"));
    assert!(sheet.get_cell(0, 2).is_none());
}

#[test]
fn test_insert_row() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Row 0");
    sheet.set_cell(1, 0, "Row 1");

    sheet.insert_row(1);

    assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Row 0"));
    assert!(sheet.get_cell(1, 0).is_none()); // Inserted row is empty
    assert_eq!(sheet.get_cell(2, 0).and_then(|v| v.as_str()), Some("Row 1"));
}

#[test]
fn test_insert_column() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Col 0");
    sheet.set_cell(0, 1, "Col 1");

    sheet.insert_column(1);

    assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Col 0"));
    assert!(sheet.get_cell(0, 1).is_none()); // Inserted column is empty
    assert_eq!(sheet.get_cell(0, 2).and_then(|v| v.as_str()), Some("Col 1"));
}

#[test]
fn test_dimensions() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    assert!(sheet.dimensions().is_none());

    sheet.set_cell(5, 10, "Test");
    assert_eq!(sheet.dimensions(), Some((0, 0, 5, 10)));
}

#[test]
fn test_cell_count() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    assert_eq!(sheet.cell_count(), 0);

    sheet.set_cell(0, 0, "A");
    sheet.set_cell(0, 1, "B");
    sheet.set_cell(1, 0, "C");

    assert_eq!(sheet.cell_count(), 3);
}

#[test]
fn test_name() {
    let sheet = MutableWorksheet::new("Sheet1");
    assert_eq!(sheet.name(), "Sheet1");
}

#[test]
fn test_set_name() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_name("RenamedSheet");
    assert_eq!(sheet.name(), "RenamedSheet");
}

#[test]
fn test_set_column_width() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_column_width(0, 15.5);
    sheet.set_column_width(2, 20.0);

    // Verify columns are set
    assert_eq!(sheet.columns.len(), 2);
}

#[test]
fn test_set_row_height() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_row_height(0, 25.0);
    sheet.set_row_height(3, 30.5);

    // Verify rows are set
    assert_eq!(sheet.rows.len(), 2);
}

#[test]
fn test_clear() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Test");
    sheet.add_merged_cell(MergedCell::new(0, 1, 0, 1));

    sheet.clear();

    assert_eq!(sheet.cell_count(), 0);
    assert!(sheet.merged_cells.is_empty());
    assert!(sheet.dimensions().is_none());
}

#[test]
fn test_add_merged_cell() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let merged = MergedCell::new(0, 1, 0, 1);
    sheet.add_merged_cell(merged);

    assert_eq!(sheet.merged_cells().len(), 1);
    assert_eq!(sheet.merged_cells()[0].row_first, 0);
    assert_eq!(sheet.merged_cells()[0].row_last, 1);
}

#[test]
fn test_add_hyperlink() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string());
    sheet.add_hyperlink(link);

    assert_eq!(sheet.hyperlinks().len(), 1);
}

#[test]
fn test_hyperlinks_mut() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string());
    sheet.add_hyperlink(link);

    let links = sheet.hyperlinks_mut();
    assert_eq!(links.len(), 1);
}

#[test]
fn test_add_comment() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let comment = Record::new(0, 0, "Author".to_string(), "Comment text".to_string());
    sheet.add_comment(comment);

    assert_eq!(sheet.comments().len(), 1);
}

#[test]
fn test_set_auto_filter() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_auto_filter(0, 10, 0, 5);

    assert!(sheet.auto_filter.is_some());
    let af = sheet.auto_filter.unwrap();
    assert_eq!(af.row_first, 0);
    assert_eq!(af.row_last, 10);
    assert_eq!(af.col_first, 0);
    assert_eq!(af.col_last, 5);
}

#[test]
fn test_add_data_validation() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let dv = Validation::new(3, "A1:A10".to_string());
    sheet.add_data_validation(dv);

    assert_eq!(sheet.data_validations().len(), 1);
    assert_eq!(sheet.data_validations()[0].validation_type, 3);
}

#[test]
fn test_add_conditional_formatting() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let cf = Formatting::new(vec!["A1:A10".to_string()]);
    sheet.add_conditional_formatting(cf);

    assert_eq!(sheet.conditional_formattings().len(), 1);
}

#[test]
fn test_cell_data_types() {
    let mut sheet = MutableWorksheet::new("Sheet1");

    // String
    sheet.set_cell(0, 0, "String");
    assert_eq!(
        sheet.get_cell(0, 0).and_then(|v| v.as_str()),
        Some("String")
    );

    // Integer - stored as CellValue::Int
    sheet.set_cell(0, 1, 42i32);
    match sheet.get_cell(0, 1) {
        Some(CellValue::Int(i)) => assert_eq!(*i, 42),
        _ => panic!("Expected Int(42)"),
    }

    // Float
    sheet.set_cell(0, 2, 1.5f64);
    assert_eq!(sheet.get_cell(0, 2).and_then(|v| v.as_float()), Some(1.5));

    // Bool - check by matching the enum variant directly
    sheet.set_cell(0, 3, true);
    match sheet.get_cell(0, 3) {
        Some(CellValue::Bool(b)) => assert!(*b),
        _ => panic!("Expected Bool(true)"),
    }
}

#[test]
fn test_worksheet_write_empty() {
    let sheet = MutableWorksheet::new("Sheet1");
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();

    let result = sheet.write(&mut writer, &mut shared_strings);
    assert!(result.is_ok());
    assert!(!buffer.is_empty());
}

#[test]
fn test_worksheet_write_with_data() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Hello");
    sheet.set_cell(0, 1, 42.0);
    sheet.set_cell(1, 0, true);

    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();

    let result = sheet.write(&mut writer, &mut shared_strings);
    assert!(result.is_ok());
    assert!(!buffer.is_empty());

    // Verify shared strings were added
    assert_eq!(shared_strings.len(), 1); // "Hello"
}

#[test]
fn test_column_info_struct() {
    let info = ColumnInfo {
        width: Some(15.0),
        hidden: false,
        best_fit: true,
    };
    assert_eq!(info.width, Some(15.0));
    assert!(!info.hidden);
    assert!(info.best_fit);
}

#[test]
fn writes_best_fit_in_the_specified_column_flag() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.columns.insert(
        2,
        ColumnInfo {
            width: Some(15.0),
            hidden: false,
            best_fit: true,
        },
    );
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet.write(&mut writer, &mut shared_strings).unwrap();

    let record = Records::new(buffer.as_slice())
        .find_map(|record| {
            let record = record.unwrap();
            (record.kind() == kind::COL_INFO).then_some(record)
        })
        .unwrap();
    assert_eq!(
        binary::read_u16_le_at(record.payload(), 16).unwrap(),
        0x0006
    );
}

#[test]
fn test_row_info_struct() {
    let info = RowInfo {
        height: Some(20.0),
        hidden: true,
    };
    assert_eq!(info.height, Some(20.0));
    assert!(info.hidden);
}

#[test]
fn test_auto_filter_struct() {
    let af = AutoFilter {
        row_first: 0,
        row_last: 10,
        col_first: 0,
        col_last: 5,
    };
    assert_eq!(af.row_first, 0);
    assert_eq!(af.row_last, 10);
    assert_eq!(af.col_first, 0);
    assert_eq!(af.col_last, 5);
}

#[test]
fn test_cell_data_struct() {
    let cell = CellData {
        value: CellValue::String("Test".to_string()),
        style: 5,
        formula_binary: None,
        formula_flags: 0,
    };
    assert_eq!(cell.style, 5);
    assert_eq!(cell.value.as_str(), Some("Test"));
}

#[test]
fn writes_ms_xlsb_brt_fmla_num_layout_without_downgrading_formula() {
    let sheet = MutableWorksheet::new("Sheet1");
    let cell = CellData {
        value: CellValue::Formula {
            formula: "C13*2".to_string(),
            cached_value: Some(Box::new(CellValue::Float(4.0))),
            is_array: false,
            array_range: None,
        },
        style: 0,
        formula_binary: None,
        formula_flags: 0,
    };
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet
        .write_cell(&mut writer, 12, 1, &cell, &mut shared_strings)
        .unwrap();

    let mut expected = vec![0x09, 0x25]; // BrtFmlaNum, 37-byte payload
    expected.extend_from_slice(&1_u32.to_le_bytes()); // Cell.column
    expected.extend_from_slice(&[0; 4]); // style and phonetic flags
    expected.extend_from_slice(&4_f64.to_le_bytes()); // cached xnum
    expected.extend_from_slice(&0_u16.to_le_bytes()); // GrbitFmla
    expected.extend_from_slice(&11_u32.to_le_bytes()); // cce
    expected.extend_from_slice(&[
        0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
    ]);
    expected.extend_from_slice(&0_u32.to_le_bytes()); // cb
    assert_eq!(buffer, expected);
}

#[test]
fn unsupported_formula_is_an_error_instead_of_a_cached_constant() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(
        0,
        0,
        CellValue::Formula {
            formula: "UNSUPPORTED(A1)".to_string(),
            cached_value: Some(Box::new(CellValue::Float(42.0))),
            is_array: false,
            array_range: None,
        },
    );
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    let error = sheet.write(&mut writer, &mut shared_strings).unwrap_err();
    assert!(matches!(error, Error::UnsupportedFeature(_)));
}

#[test]
fn formula_without_cached_result_is_marked_for_recalculation() {
    let sheet = MutableWorksheet::new("Sheet1");
    let cell = CellData {
        value: CellValue::Formula {
            formula: "1+1".to_string(),
            cached_value: None,
            is_array: false,
            array_range: None,
        },
        style: 0,
        formula_binary: None,
        formula_flags: 0,
    };
    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet
        .write_cell(&mut writer, 0, 0, &cell, &mut shared_strings)
        .unwrap();

    // Two-byte record header, then Cell (8) + cached xnum (8).
    assert_eq!(u16::from_le_bytes([buffer[18], buffer[19]]), 0x0002);
}

#[test]
fn writes_shared_definition_immediately_after_anchor_and_exp_followers() {
    use crate::package::records::Stream;

    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(2, 2, 10.0);
    sheet.set_cell(3, 2, 20.0);
    sheet.set_shared_formula(2, 2, 3, 2, "B3").unwrap();

    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet.write_cells(&mut writer, &mut shared_strings).unwrap();

    let mut iter = Stream::new(Cursor::new(buffer));
    let mut records = Vec::new();
    while let Ok(record_type) = iter.read_type() {
        let mut data = Vec::new();
        iter.fill_buffer(&mut data).unwrap();
        records.push((record_type, data));
    }
    assert_eq!(
        records.iter().map(|record| record.0).collect::<Vec<_>>(),
        vec![
            kind::ROW_HDR,
            kind::FMLA_NUM,
            kind::SHR_FMLA,
            kind::ROW_HDR,
            kind::FMLA_NUM,
        ]
    );

    let group = Group::parse_shared(&records[2].1).unwrap();
    assert_eq!(group.range.to_a1(), "C3:C4");
    for formula_record in [&records[1].1, &records[4].1] {
        let (placeholder, consumed) = ParsedFormula::parse(&formula_record[18..]).unwrap();
        assert_eq!(18 + consumed, formula_record.len());
        assert_eq!(placeholder.exp_cell().unwrap(), Some((2, 2)));
    }
}

#[test]
fn writes_unsupported_group_definition_losslessly() {
    use crate::package::records::Stream;
    use std::io::Cursor;

    let group = Group {
        kind: GroupKind::Array,
        range: Range::new(8, 8, 2, 2).unwrap(),
        formula: ParsedFormula {
            rgce: vec![0x23, 0x02, 0x00, 0x00, 0x00, 0x42, 0x01, 0xFF, 0x00],
            rgcb: Vec::new(),
        },
        always_calculate: true,
    };
    let expected = group.to_record_data().unwrap();
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_formula_group_binary(group).unwrap();

    let mut buffer = Vec::new();
    let mut writer = Writer::new(&mut buffer);
    let mut shared_strings = crate::writer::MutableSharedStringsWriter::new();
    sheet.write_cells(&mut writer, &mut shared_strings).unwrap();

    let mut iter = Stream::new(Cursor::new(buffer));
    let mut definition = None;
    while let Ok(record_type) = iter.read_type() {
        let mut data = Vec::new();
        iter.fill_buffer(&mut data).unwrap();
        if record_type == kind::ARR_FMLA {
            definition = Some(data);
        }
    }
    assert_eq!(definition.as_deref(), Some(expected.as_slice()));
}
