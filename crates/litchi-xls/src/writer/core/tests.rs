use super::*;
use crate::writer::biff::AutoFilterConditionWrite;
use crate::{Error, Result};
use litchi_core::sheet::Cell;
use std::io::Cursor;

use crate::formula_metadata::{Cell as FormulaCell, Range as FormulaRange};

#[test]
fn test_create_writer() {
    let writer = Writer::new();
    assert_eq!(writer.worksheets.len(), 0);
    assert_eq!(writer.shared_strings.len(), 0);
}

#[test]
fn test_add_worksheet() {
    let mut writer = Writer::new();
    let idx = writer.add_worksheet("Sheet1").unwrap();
    assert_eq!(idx, 0);
    assert_eq!(writer.worksheets.len(), 1);
    assert_eq!(writer.worksheets[0].name, "Sheet1");
}

#[test]
fn test_add_multiple_worksheets() {
    let mut writer = Writer::new();
    let idx1 = writer.add_worksheet("Sheet1").unwrap();
    let idx2 = writer.add_worksheet("Sheet2").unwrap();
    let idx3 = writer.add_worksheet("Sheet3").unwrap();

    assert_eq!(idx1, 0);
    assert_eq!(idx2, 1);
    assert_eq!(idx3, 2);
    assert_eq!(writer.worksheets.len(), 3);
}

#[test]
fn test_add_worksheet_empty_name() {
    let mut writer = Writer::new();
    let result = writer.add_worksheet("");
    assert!(result.is_err());
}

#[test]
fn test_add_worksheet_long_name() {
    let mut writer = Writer::new();
    let long_name = "A".repeat(50);
    let result = writer.add_worksheet(&long_name);
    assert!(result.is_err()); // Name too long
}

#[test]
fn test_add_worksheet_duplicate_name() {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    let result = writer.add_worksheet("Sheet1");
    assert!(result.is_err());
}

#[test]
fn test_write_string() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_string(sheet, 0, 0, "Hello").unwrap();
    assert_eq!(writer.worksheets[0].cells.len(), 1);

    let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
    assert_eq!(cell.row(), 0);
    assert_eq!(cell.col(), 0);
    assert!(matches!(&cell.value, CellValue::String(s) if s == "Hello"));
}

#[test]
fn test_write_number() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 0, 0, 42.5).unwrap();
    assert_eq!(writer.worksheets[0].cells.len(), 1);

    let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
    assert!(matches!(&cell.value, CellValue::Number(n) if *n == 42.5));
}

#[test]
fn non_finite_cell_numbers_are_rejected_before_mutation() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();

    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.write_number(sheet, 0, 0, value)
        }));
        assert!(matches!(
            result,
            Ok(Err(Error::InvalidData(message)))
                if message == "cell number must be finite for BIFF8 serialization"
        ));
        assert!(writer.worksheets[sheet].cells.is_empty());
    }
}

#[test]
fn cell_grid_bounds_are_atomic_and_never_unwind() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();

    let max = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.write_number(sheet, u32::from(u16::MAX), u16::from(u8::MAX), 42.5)
    }));
    assert!(matches!(max, Ok(Ok(()))));
    let state = |writer: &Writer| {
        let worksheet = &writer.worksheets[sheet];
        (
            worksheet.cells.len(),
            worksheet.first_row,
            worksheet.last_row,
            worksheet.first_col,
            worksheet.last_col,
        )
    };
    let max_state = (1, 65_535, 65_536, 255, 256);
    assert_eq!(state(&writer), max_state);

    let oversized_row = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.write_string(sheet, 65_536, 0, "outside")
    }));
    assert!(matches!(
        oversized_row,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(state(&writer), max_state);

    let adversarial_row = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.write_number(sheet, u32::MAX, 0, 1.0)
    }));
    assert!(matches!(
        adversarial_row,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(state(&writer), max_state);

    let oversized_col = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.write_formula(sheet, 0, 256, "1")
    }));
    assert!(matches!(
        oversized_col,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(state(&writer), max_state);

    let table = crate::DataTable::one_variable(
        crate::DataTableRange::new(2, 8, 3, 5).unwrap(),
        false,
        crate::DataTableInputCell::Deleted,
    );
    let adversarial_col = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_data_table(sheet, 0, u16::MAX, table)
    }));
    assert!(matches!(
        adversarial_col,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(state(&writer), max_state);
    assert!(writer.worksheets[sheet].data_tables.is_empty());

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.set_position(0);
    let workbook = crate::Workbook::new(output).unwrap();
    let cell = workbook
        .xls_worksheet(0)
        .unwrap()
        .get_cell(u32::from(u16::MAX), u32::from(u8::MAX))
        .unwrap();
    assert!(matches!(
        cell.value(),
        litchi_core::sheet::CellValue::Float(value) if *value == 42.5
    ));
}

#[test]
fn worksheet_location_bounds_are_atomic_and_never_unwind() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Grid").unwrap();

    writer.merge_cells(sheet, 65_534, 65_535, 254, 255).unwrap();
    writer
        .add_horizontal_page_break(sheet, 65_535, 0, 16_383)
        .unwrap();
    writer
        .add_vertical_page_break(sheet, 255, 0, 65_535)
        .unwrap();
    let max_range = DataValidationRange::new(65_535, 65_535, 255, 255).unwrap();
    writer
        .add_data_validation(
            sheet,
            DataValidation::new(max_range, DataValidationType::Any),
        )
        .unwrap();

    let state = |writer: &Writer| {
        let worksheet = &writer.worksheets[sheet];
        (
            worksheet.merged_ranges.len(),
            worksheet.horizontal_page_breaks.len(),
            worksheet.vertical_page_breaks.len(),
            worksheet.data_validations.len(),
            writer.real_time_data.len(),
            writer.web_publications.len(),
        )
    };
    let max_state = (1, 1, 1, 1, 0, 0);
    assert_eq!(state(&writer), max_state);

    for outcome in [
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.merge_cells(sheet, 65_536, 65_536, 0, 0)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.merge_cells(sheet, 0, 0, 256, 256)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.merge_cells(sheet, 2, 1, 0, 0)
        })),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }
    let merge_overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.merge_cells(sheet, 65_535, 65_535, 255, 255)
    }));
    assert!(matches!(merge_overlap, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(state(&writer), max_state);

    for outcome in [
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_horizontal_page_break(sheet, 65_536, 0, 1)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_horizontal_page_break(sheet, 0, 0, 16_384)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_vertical_page_break(sheet, 256, 0, 1)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.add_vertical_page_break(sheet, 0, 0, 65_536)
        })),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }

    let overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_horizontal_page_break(sheet, 65_535, 1, 2)
    }));
    assert!(matches!(overlap, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(state(&writer), max_state);
    let overlap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_vertical_page_break(sheet, 255, 1, 2)
    }));
    assert!(matches!(overlap, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(state(&writer), max_state);

    for outcome in [
        std::panic::catch_unwind(|| DataValidationRange::new(65_536, 65_536, 0, 0)),
        std::panic::catch_unwind(|| DataValidationRange::new(0, 0, 256, 256)),
        std::panic::catch_unwind(|| DataValidationRange::new(1, 0, 0, 0)),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }

    let too_many_ranges = vec![max_range; 432];
    let count_error = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_data_validation_with_options(
            sheet,
            DataValidation::new(max_range, DataValidationType::Any),
            &too_many_ranges,
            DataValidationOptions::default(),
        )
    }));
    assert!(matches!(count_error, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(state(&writer), max_state);

    let invalid_rule = DataValidation::new(
        max_range,
        DataValidationType::Whole {
            operator: DataValidationOperator::Between,
            value1: 1,
            value2: None,
        },
    );
    let rule_error = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_data_validation(sheet, invalid_rule)
    }));
    assert!(matches!(rule_error, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(state(&writer), max_state);

    assert!(matches!(
        crate::DataTableInputCell::present(65_535, 255),
        Ok(crate::DataTableInputCell::Present {
            row: 65_535,
            col: 255
        })
    ));
    for outcome in [
        std::panic::catch_unwind(|| crate::DataTableInputCell::present(65_536, 0)),
        std::panic::catch_unwind(|| crate::DataTableInputCell::present(u32::MAX, 0)),
        std::panic::catch_unwind(|| crate::DataTableInputCell::present(0, 256)),
        std::panic::catch_unwind(|| crate::DataTableInputCell::present(0, u16::MAX)),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }
    let invalid_publication = crate::WebPub {
        source: crate::WebSourceType::Workbook,
        page_type: crate::WebPageType::ViewOnly,
        range: Some(crate::WebPubRange::new(0, 0, 0, 0).unwrap()),
        auto_republish: false,
        single_file: false,
        style_id: 1,
        source_name: None,
        file_destination: "x.htm".to_string(),
        div_id: String::new(),
        title: "x".to_string(),
        chart_shape_id: None,
        reserved: Vec::new(),
    };
    let invalid_publication = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_web_publication(invalid_publication)
    }));
    assert!(matches!(
        invalid_publication,
        Ok(Err(Error::InvalidData(_)))
    ));
    assert_eq!(state(&writer), max_state);

    for outcome in [
        std::panic::catch_unwind(|| crate::WebPubRange::new(65_536, 65_536, 0, 0)),
        std::panic::catch_unwind(|| crate::WebPubRange::new(0, 0, 256, 256)),
        std::panic::catch_unwind(|| crate::WebPubRange::new(1, 0, 0, 0)),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }

    for outcome in [
        std::panic::catch_unwind(|| crate::RtdCell::new(65_536, 0, 0)),
        std::panic::catch_unwind(|| crate::RtdCell::new(0, 256, 0)),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(state(&writer), max_state);
    }
    let invalid_topic = crate::RealTimeData {
        common_prefix_len: 0,
        topic_segments: vec!["server".to_string(), String::new()],
        topic: "server".to_string(),
        value: crate::RtdValue::Integer(1),
        cells: vec![crate::RtdCell::new(0, 0, 1).unwrap()],
    };
    let missing_sheet = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_real_time_data(invalid_topic)
    }));
    assert!(matches!(
        missing_sheet,
        Ok(Err(Error::WorksheetNotFound(_)))
    ));
    assert_eq!(state(&writer), max_state);

    writer
        .set_page_setup(sheet, PageSetupOptions::default())
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.set_position(0);
    let workbook = crate::Workbook::new(output).unwrap();
    let worksheet = workbook.xls_worksheet(sheet).unwrap();
    assert!(
        worksheet
            .merged_cells()
            .iter()
            .any(|range| range.first_row == 65_534
                && range.last_row == 65_535
                && range.first_col == 254
                && range.last_col == 255)
    );
    let page_setup = worksheet.page_setup().unwrap();
    assert_eq!(page_setup.horizontal_page_breaks()[0].position(), 65_535);
    assert_eq!(page_setup.horizontal_page_breaks()[0].range_end(), 16_383);
    assert_eq!(page_setup.vertical_page_breaks()[0].position(), 255);
    assert_eq!(page_setup.vertical_page_breaks()[0].range_end(), 65_535);
}

#[test]
fn page_break_entry_limits_are_failure_atomic() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Break limits").unwrap();

    for row in 0..1_026 {
        writer.add_horizontal_page_break(sheet, row, 0, 1).unwrap();
    }
    for column in 0..255 {
        writer.add_vertical_page_break(sheet, column, 0, 1).unwrap();
    }
    assert_eq!(writer.worksheets[sheet].horizontal_page_breaks.len(), 1_026);
    assert_eq!(writer.worksheets[sheet].vertical_page_breaks.len(), 255);

    let horizontal = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_horizontal_page_break(sheet, 1_026, 0, 1)
    }));
    assert!(matches!(horizontal, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(writer.worksheets[sheet].horizontal_page_breaks.len(), 1_026);

    let vertical = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_vertical_page_break(sheet, 255, 0, 1)
    }));
    assert!(matches!(vertical, Ok(Err(Error::InvalidData(_)))));
    assert_eq!(writer.worksheets[sheet].vertical_page_breaks.len(), 255);
}

#[test]
fn filter_sort_and_pivot_locations_fail_before_mutation() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Locations").unwrap();
    writer.set_auto_filter(sheet, 0, 65_535, 0, 255).unwrap();
    let filter_state = |writer: &Writer| {
        let worksheet = &writer.worksheets[sheet];
        (
            worksheet.auto_filter.map(|range| {
                (
                    range.first_row,
                    range.last_row,
                    range.first_col,
                    range.last_col,
                )
            }),
            worksheet.auto_filter_columns.len(),
            worksheet.sort_config.is_some(),
            worksheet.pivot_tables.len(),
            worksheet.cells.len(),
            writer.defined_names.len(),
        )
    };
    let initial = (Some((0, 65_535, 0, 255)), 0, false, 0, 0, 1);
    assert_eq!(filter_state(&writer), initial);

    for outcome in [
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.set_auto_filter(sheet, 0, 65_536, 0, 0)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.set_auto_filter(sheet, 0, 0, 0, 256)
        })),
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            writer.set_auto_filter(sheet, 1, 0, 0, 0)
        })),
    ] {
        assert!(matches!(outcome, Ok(Err(Error::InvalidCellReference(_)))));
        assert_eq!(filter_state(&writer), initial);
    }

    let invalid_filter = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_filter_condition(
            sheet,
            256,
            false,
            AutoFilterConditionWrite::None,
            AutoFilterConditionWrite::None,
        )
    }));
    assert!(matches!(
        invalid_filter,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(filter_state(&writer), initial);

    let invalid_sort = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.set_sort(sheet, false, false, &[(256, false)])
    }));
    assert!(matches!(
        invalid_sort,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(filter_state(&writer), initial);

    let invalid_pivot = PivotTableConfig {
        name: "InvalidPivot".to_string(),
        source_type: 1,
        source_sheet_name: "Locations".to_string(),
        source_first_row: 0,
        source_last_row: 0,
        source_first_col: 0,
        source_last_col: 256,
        first_row: 0,
        last_row: 1,
        first_col: 0,
        last_col: 1,
        first_header_row: 0,
        first_data_row: 1,
        first_data_col: 1,
        data_field_name: "Values".to_string(),
        data_axis: 0,
        data_position: 0,
        fields: Vec::new(),
        data_items: Vec::new(),
        page_entries: Vec::new(),
        source_data: Vec::new(),
    };
    let invalid_pivot = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        writer.add_pivot_table(sheet, invalid_pivot)
    }));
    assert!(matches!(
        invalid_pivot,
        Ok(Err(Error::InvalidCellReference(_)))
    ));
    assert_eq!(filter_state(&writer), initial);
}

#[test]
fn test_write_boolean() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_boolean(sheet, 0, 0, true).unwrap();
    writer.write_boolean(sheet, 1, 0, false).unwrap();

    assert_eq!(writer.worksheets[0].cells.len(), 2);
    assert!(matches!(
        writer.worksheets[0].cells.get(&(0, 0)).unwrap().value,
        CellValue::Boolean(true)
    ));
    assert!(matches!(
        writer.worksheets[0].cells.get(&(1, 0)).unwrap().value,
        CellValue::Boolean(false)
    ));
}

#[test]
fn test_write_formula() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_formula(sheet, 0, 0, "SUM(A1:B1)").unwrap();
    let metadata = crate::FormulaMetadata::new()
        .with_fill_alignment(true)
        .with_clear_errors(true)
        .with_calculation_cache(0xCAFE_BABE);
    writer
        .write_formula_with_metadata(sheet, 0, 1, "A1+1", metadata.clone())
        .unwrap();

    let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
    assert!(matches!(&cell.value, CellValue::Formula(f) if f == "SUM(A1:B1)"));
    assert_eq!(
        writer.worksheets[0]
            .cells
            .get(&(0, 1))
            .and_then(|cell| cell.formula_metadata.clone()),
        Some(metadata)
    );
}

#[test]
fn invalid_formula_reference_returns_error_without_unwinding() {
    let outcome = std::panic::catch_unwind(|| -> Result<()> {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1")?;
        writer.write_formula(sheet, 0, 0, "ZZZZ1")?;
        writer.write_to(&mut Cursor::new(Vec::new()))
    });

    assert!(matches!(
        outcome,
        Ok(Err(Error::InvalidCellReference(reference))) if reference == "ZZZZ1"
    ));
}

#[test]
fn test_formula_round_trips_through_xls_reader() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 0, 0, 2.0).unwrap();
    writer.write_number(sheet, 0, 1, 3.0).unwrap();
    writer.write_formula(sheet, 0, 2, "SUM(A1:B1)").unwrap();
    let metadata = crate::FormulaMetadata::new()
        .with_fill_alignment(true)
        .with_clear_errors(true)
        .with_calculation_cache(0x1020_3040);
    writer
        .write_formula_with_metadata(sheet, 0, 4, "A1+1", metadata.clone())
        .unwrap();
    writer
        .write_formula(sheet, 0, 3, "IF(TRUE,\"a\"\"b\",FALSE)")
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.set_position(0);
    let workbook = crate::Workbook::new(output).unwrap();
    let formula_cell = workbook.xls_worksheet(0).unwrap().get_cell(0, 2).unwrap();

    assert!(formula_cell.is_formula());
    assert_eq!(formula_cell.formula(), Some("=SUM((A1:B1))"));
    assert!(!formula_cell.formula_bytes().unwrap().is_empty());
    assert_eq!(
        formula_cell.formula_metadata(),
        Some(&crate::FormulaMetadata::new().with_always_calculate(true))
    );
    let metadata_cell = workbook.xls_worksheet(0).unwrap().get_cell(0, 4).unwrap();
    assert_eq!(metadata_cell.formula_metadata(), Some(&metadata));
    assert_eq!(
        workbook
            .xls_worksheet(0)
            .unwrap()
            .get_cell(0, 3)
            .unwrap()
            .formula(),
        Some("=IF(TRUE,\"a\"\"b\",FALSE)")
    );
}

#[test]
fn shared_formula_round_trips_with_inert_relative_references() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    let anchor = FormulaCell::new(0, 0);
    let participants = [anchor, FormulaCell::new(1, 0)];
    writer
        .write_shared_formula(
            sheet,
            FormulaRange::try_new(0, 0, 1, 0).unwrap(),
            anchor,
            "A1*2",
            &participants,
        )
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.set_position(0);
    let workbook = crate::Workbook::new(output).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let first = worksheet.get_cell(0, 0).unwrap();
    let second = worksheet.get_cell(1, 0).unwrap();

    assert_eq!(first.formula(), Some("=(A1*2)"));
    assert_eq!(second.formula(), Some("=(A2*2)"));
    assert_eq!(first.formula_bytes(), Some(&[0x01, 0, 0, 0, 0][..]));
    assert_eq!(second.formula_bytes(), Some(&[0x01, 0, 0, 0, 0][..]));
    assert!(first.formula_metadata().unwrap().shared_formula());
    assert!(second.formula_metadata().unwrap().shared_formula());
}

#[test]
fn test_write_multiple_cells() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();

    writer.write_string(sheet, 0, 0, "A1").unwrap();
    writer.write_string(sheet, 0, 1, "B1").unwrap();
    writer.write_string(sheet, 1, 0, "A2").unwrap();
    writer.write_string(sheet, 1, 1, "B2").unwrap();

    assert_eq!(writer.worksheets[0].cells.len(), 4);
}

#[test]
fn test_shared_strings_build() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();

    writer.write_string(sheet, 0, 0, "Hello").unwrap();
    writer.write_string(sheet, 0, 1, "Hello").unwrap();
    writer.write_string(sheet, 1, 0, "World").unwrap();

    // Build shared strings table (normally done during write)
    writer.build_shared_strings();

    // Should only have 2 unique strings
    assert_eq!(writer.shared_strings.len(), 2);
}

#[test]
fn test_write_to_memory() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_string(sheet, 0, 0, "Test").unwrap();
    writer.write_number(sheet, 0, 1, 123.45).unwrap();

    let mut cursor = Cursor::new(Vec::new());
    let result = writer.write_to(&mut cursor);
    assert!(result.is_ok());

    let data = cursor.into_inner();
    assert!(!data.is_empty());
    // Should start with OLE compound document signature
    assert_eq!(
        &data[0..8],
        [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
    );
}

#[test]
fn test_save_to_file() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_string(sheet, 0, 0, "Hello").unwrap();

    let temp_path = std::env::temp_dir().join("test_xls_writer.xls");
    let result = writer.save(&temp_path);
    assert!(result.is_ok());

    // Verify file was created
    assert!(temp_path.exists());

    // Clean up
    let _ = std::fs::remove_file(temp_path);
}

#[test]
fn test_xls_writer_default() {
    let writer: Writer = Default::default();
    assert_eq!(writer.worksheets.len(), 0);
    assert_eq!(writer.shared_strings.len(), 0);
}

#[test]
fn test_xlscellvalue_variants() {
    let string_val = CellValue::String("test".to_string());
    let number_val = CellValue::Number(42.0);
    let bool_val = CellValue::Boolean(true);
    let formula_val = CellValue::Formula("A1+B1".to_string());
    let blank_val = CellValue::Blank;

    assert!(matches!(string_val, CellValue::String(_)));
    assert!(matches!(number_val, CellValue::Number(_)));
    assert!(matches!(bool_val, CellValue::Boolean(_)));
    assert!(matches!(formula_val, CellValue::Formula(_)));
    assert!(matches!(blank_val, CellValue::Blank));
}

#[test]
fn test_xlscellvalue_debug() {
    let val = CellValue::String("test".to_string());
    let debug = format!("{:?}", val);
    assert!(debug.contains("String"));
}

#[test]
fn test_xlscellvalue_clone() {
    let val = CellValue::Number(42.0);
    let cloned = val.clone();
    assert!(matches!(cloned, CellValue::Number(42.0)));
}

#[test]
fn test_writablecell_creation() {
    let cell = WritableCell::new(
        CellPos::try_new(5, 3).unwrap(),
        CellValue::String("Test".to_string()),
        15,
        None,
    );

    assert_eq!(cell.row(), 5);
    assert_eq!(cell.col(), 3);
    assert_eq!(cell.format_idx, 15);
    assert!(matches!(
        CellPos::try_new(65_536, 0),
        Err(Error::InvalidCellReference(_))
    ));
    assert!(matches!(
        CellPos::try_new(0, 256),
        Err(Error::InvalidCellReference(_))
    ));
}

#[test]
fn test_writableworksheet_creation() {
    let ws = WritableWorksheet::new("TestSheet".to_string());
    assert_eq!(ws.name, "TestSheet");
    assert!(ws.cells.is_empty());
    assert!(ws.merged_ranges.is_empty());
    assert!(ws.column_widths.is_empty());
}

#[test]
fn test_writableworksheet_add_cell() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    let cell = WritableCell::new(
        CellPos::try_new(0, 0).unwrap(),
        CellValue::Number(100.0),
        0,
        None,
    );
    ws.add_cell(cell);
    assert_eq!(ws.cells.len(), 1);
}

#[test]
fn test_writableworksheet_set_column_width() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    ws.set_column_width(0, 2560); // ~10 characters
    assert_eq!(ws.column_widths.get(&0), Some(&2560));
}

#[test]
fn test_writableworksheet_merge_cells() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    ws.add_merged_range(MergedRange::try_new(0, 1, 0, 2).unwrap()); // Merge A1:C2
    assert_eq!(ws.merged_ranges.len(), 1);
    assert_eq!(ws.merged_ranges[0].fields(), (0, 1, 0, 2));
}

#[test]
fn test_writableworksheet_freeze_panes() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    assert!(ws.view.pane().is_none());
    ws.set_freeze_panes(1, 2).unwrap();
    let pane = ws.view.pane().unwrap();
    assert_eq!(pane.vertical(), 1);
    assert_eq!(pane.horizontal(), 2);
}

#[test]
fn test_writableworksheet_add_conditional_format() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    let cf = ConditionalFormat {
        first_row: 0,
        last_row: 10,
        first_col: 0,
        last_col: 0,
        format_type: ConditionalFormatType::Formula {
            formula: "A1>100".to_string(),
        },
        pattern: None,
    };
    ws.add_conditional_format(cf);
    assert_eq!(ws.conditional_formats.len(), 1);
}

#[test]
fn test_writableworksheet_add_data_validation() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    let dv = DataValidation {
        range: DataValidationRange::new(0, 10, 0, 0).unwrap(),
        validation_type: DataValidationType::List {
            values: vec!["Option1".to_string(), "Option2".to_string()],
        },
        show_input_message: true,
        input_title: None,
        input_message: None,
        show_error_alert: true,
        error_title: None,
        error_message: None,
    };
    let payload = dv.validation_type.to_biff_payload().unwrap();
    ws.add_data_validation(
        dv,
        payload,
        vec![DataValidationRange::new(0, 9, 0, 0).unwrap()],
        DataValidationOptions::default(),
    );
    assert_eq!(ws.data_validations.len(), 1);
}

#[test]
fn test_writableworksheet_add_hyperlink() {
    let mut ws = WritableWorksheet::new("Sheet1".to_string());
    let link = Hyperlink {
        first_row: 0,
        last_row: 0,
        first_col: 0,
        last_col: 0,
        url: "https://example.com".to_string(),
    };
    ws.add_hyperlink(link);
    assert_eq!(ws.hyperlinks.len(), 1);
    assert_eq!(ws.hyperlinks[0].url, "https://example.com");
}

#[test]
fn test_xls_defined_name_basic() {
    let name = DefinedName {
        name: "TestRange".to_string(),
        reference: "A1:B10".to_string(),
        comment: None,
        local_sheet: None,
        target_sheet: Some(0),
        hidden: false,
        is_function: false,
        is_built_in: false,
        built_in_code: None,
    };
    assert_eq!(name.name, "TestRange");
    assert_eq!(name.reference, "A1:B10");
    assert_eq!(name.target_sheet, Some(0));
}

#[test]
fn test_xls_defined_name_to_biff_formula_area() {
    let name = DefinedName {
        name: "TestRange".to_string(),
        reference: "A1:B10".to_string(),
        comment: None,
        local_sheet: None,
        target_sheet: Some(0),
        hidden: false,
        is_function: false,
        is_built_in: false,
        built_in_code: None,
    };
    let formula = name.to_biff_formula().unwrap();
    assert!(!formula.is_empty());
}

#[test]
fn test_xls_defined_name_normalizes_reversed_area_corners() {
    let forward = DefinedName {
        name: "Forward".to_string(),
        reference: "A1:B10".to_string(),
        comment: None,
        local_sheet: None,
        target_sheet: Some(0),
        hidden: false,
        is_function: false,
        is_built_in: false,
        built_in_code: None,
    };
    let reversed = DefinedName {
        name: "Reversed".to_string(),
        reference: "B10:A1".to_string(),
        ..forward.clone()
    };

    assert_eq!(
        reversed.to_biff_formula().unwrap(),
        forward.to_biff_formula().unwrap()
    );
}

#[test]
fn test_xls_defined_name_to_biff_formula_single() {
    let name = DefinedName {
        name: "SingleCell".to_string(),
        reference: "C5".to_string(),
        comment: None,
        local_sheet: None,
        target_sheet: None,
        hidden: false,
        is_function: false,
        is_built_in: false,
        built_in_code: None,
    };
    let formula = name.to_biff_formula().unwrap();
    assert!(!formula.is_empty());
}
