//! Tests for text-based spreadsheet formats

use super::*;
use crate::sheet::{CellValue, Workbook};

#[test]
fn test_csv_parsing() {
    let csv_data = "name,age,city\nJohn,25,New York\nJane,30,London";
    let workbook = TextWorkbook::from_bytes(csv_data.as_bytes(), TextConfig::default()).unwrap();

    // Test worksheet access
    let worksheet = workbook.active_worksheet().unwrap();

    // Test dimensions
    assert_eq!(worksheet.row_count(), 3);
    assert_eq!(worksheet.column_count(), 3);

    // Test cell access
    let cell_a1 = worksheet.cell(1, 1).unwrap();
    assert_eq!(cell_a1.value(), &CellValue::String("name".to_string()));

    let cell_a2 = worksheet.cell(2, 1).unwrap();
    assert_eq!(cell_a2.value(), &CellValue::String("John".to_string()));

    let cell_b2 = worksheet.cell(2, 2).unwrap();
    assert_eq!(cell_b2.value(), &CellValue::Int(25));

    // Test coordinate access
    let cell_c2 = worksheet.cell_by_coordinate("C2").unwrap();
    assert_eq!(cell_c2.value(), &CellValue::String("New York".to_string()));

    // Test row access
    let row1 = worksheet.row(0).unwrap();
    assert_eq!(row1.len(), 3);
    assert_eq!(row1[0], CellValue::String("name".to_string()));
    assert_eq!(row1[1], CellValue::String("age".to_string()));
    assert_eq!(row1[2], CellValue::String("city".to_string()));

    let row2 = worksheet.row(1).unwrap();
    assert_eq!(row2.len(), 3);
    assert_eq!(row2[0], CellValue::String("John".to_string()));
    assert_eq!(row2[1], CellValue::Int(25));
    assert_eq!(row2[2], CellValue::String("New York".to_string()));
}

#[test]
fn test_tsv_parsing() {
    let tsv_data = "name\tage\tcity\nJohn\t25\tNew York\nJane\t30\tLondon";
    let config = TextConfig::tsv();
    let workbook = TextWorkbook::from_bytes(tsv_data.as_bytes(), config).unwrap();

    let worksheet = workbook.active_worksheet().unwrap();
    assert_eq!(worksheet.row_count(), 3);
    assert_eq!(worksheet.column_count(), 3);

    let cell_a2 = worksheet.cell(2, 1).unwrap();
    assert_eq!(cell_a2.value(), &CellValue::String("John".to_string()));
}

#[test]
fn test_quoted_fields() {
    let csv_data = "\"Hello, World\",\"Value with \"\"quotes\"\"\",\"Normal\"";
    let workbook = TextWorkbook::from_bytes(csv_data.as_bytes(), TextConfig::default()).unwrap();

    let worksheet = workbook.active_worksheet().unwrap();
    let row = worksheet.row(0).unwrap();

    assert_eq!(row.len(), 3);
    assert_eq!(row[0], CellValue::String("Hello, World".to_string()));
    assert_eq!(row[1], CellValue::String("Value with \"quotes\"".to_string()));
    assert_eq!(row[2], CellValue::String("Normal".to_string()));
}

#[test]
fn test_empty_cells() {
    let csv_data = "a,,c\n,,";
    let workbook = TextWorkbook::from_bytes(csv_data.as_bytes(), TextConfig::default()).unwrap();

    let worksheet = workbook.active_worksheet().unwrap();

    // Test empty cell access
    let empty_cell = worksheet.cell(1, 2).unwrap();
    assert_eq!(empty_cell.value(), &CellValue::Empty);

    let row1 = worksheet.row(0).unwrap();
    assert_eq!(row1[1], CellValue::Empty);

    let row2 = worksheet.row(1).unwrap();
    assert_eq!(row2[0], CellValue::Empty);
    assert_eq!(row2[1], CellValue::Empty);
    assert_eq!(row2[2], CellValue::Empty);
}

#[test]
fn test_type_inference() {
    let csv_data = "int,float,bool,string\n42,3.14,true,hello\n,2.0,false,";
    let workbook = TextWorkbook::from_bytes(csv_data.as_bytes(), TextConfig::default()).unwrap();

    let worksheet = workbook.active_worksheet().unwrap();
    let row2 = worksheet.row(1).unwrap();

    assert_eq!(row2[0], CellValue::Int(42));
    assert_eq!(row2[1], CellValue::Float(3.14));
    assert_eq!(row2[2], CellValue::Bool(true));
    assert_eq!(row2[3], CellValue::String("hello".to_string()));

    let row3 = worksheet.row(2).unwrap();
    assert_eq!(row3[0], CellValue::Empty);
    assert_eq!(row3[1], CellValue::Float(2.0));
    assert_eq!(row3[2], CellValue::Bool(false));
    assert_eq!(row3[3], CellValue::Empty);
}

#[test]
fn test_iterators() {
    let csv_data = "a,b,c\n1,2,3\n4,5,6";
    let workbook = TextWorkbook::from_bytes(csv_data.as_bytes(), TextConfig::default()).unwrap();

    let worksheet = workbook.active_worksheet().unwrap();

    // Test row iterator
    let mut rows_iter = worksheet.rows();
    let mut row_count = 0;
    while let Some(row_result) = rows_iter.next() {
        let row = row_result.unwrap();
        row_count += 1;
        match row_count {
            1 => {
                assert_eq!(row[0], CellValue::String("a".to_string()));
                assert_eq!(row[1], CellValue::String("b".to_string()));
                assert_eq!(row[2], CellValue::String("c".to_string()));
            }
            2 => {
                assert_eq!(row[0], CellValue::Int(1));
                assert_eq!(row[1], CellValue::Int(2));
                assert_eq!(row[2], CellValue::Int(3));
            }
            3 => {
                assert_eq!(row[0], CellValue::Int(4));
                assert_eq!(row[1], CellValue::Int(5));
                assert_eq!(row[2], CellValue::Int(6));
            }
            _ => panic!("Too many rows"),
        }
    }
    assert_eq!(row_count, 3);

    // Test cell iterator
    let mut cells_iter = worksheet.cells();
    let mut cell_count = 0;
    while let Some(cell_result) = cells_iter.next() {
        let cell = cell_result.unwrap();
        cell_count += 1;
        // Just verify we can access the cells
        let _ = cell.value();
    }
    assert_eq!(cell_count, 9); // 3 rows * 3 columns
}
