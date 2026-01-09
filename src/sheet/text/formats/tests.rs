//! Tests for all text format handlers

use super::*;
use crate::common::BomKind;
use crate::sheet::CellValue;
use std::io::Cursor;

#[test]
fn test_sylk_basic_read_write() {
    let sylk_data = "ID;PWXL;N;E\nB;Y3;X3\nC;Y1;X1;K\"Name\"\nC;Y1;X2;K\"Age\"\nC;Y1;X3;K\"City\"\nC;Y2;X1;K\"Alice\"\nC;Y2;X2;K25\nC;Y2;X3;K\"NYC\"\nC;Y3;X1;K\"Bob\"\nC;Y3;X2;K30\nC;Y3;X3;K\"LA\"\nE\n";

    let mut cursor = Cursor::new(sylk_data.as_bytes());
    let data = read_sylk(&mut cursor, SylkConfig::default()).unwrap();

    assert_eq!(data.len(), 3);
    assert_eq!(data[0].len(), 3);
    assert_eq!(data[0][0], CellValue::String("Name".to_string()));
    assert_eq!(data[1][1], CellValue::Int(25));

    // Test writing
    let mut output = Vec::new();
    write_sylk(&data, &mut output, SylkConfig::default()).unwrap();
    let output_str = String::from_utf8(output).unwrap();
    assert!(output_str.contains("C;Y1;X1;K\"Name\""));
}

#[test]
fn test_sylk_with_bom() {
    let mut output = Vec::new();
    let data = vec![vec![CellValue::String("Test".to_string())]];

    let config = SylkConfig {
        strip_bom: true,
        write_bom: Some(BomKind::Utf8),
    };

    write_sylk(&data, &mut output, config).unwrap();

    assert_eq!(&output[0..3], &[0xEF, 0xBB, 0xBF]);
}

#[test]
fn test_sylk_formulas() {
    let sylk_data = "ID;PWXL;N;E\nC;Y1;X1;E1+2\nE\n";
    let mut cursor = Cursor::new(sylk_data.as_bytes());
    let data = read_sylk(&mut cursor, SylkConfig::default()).unwrap();

    match &data[0][0] {
        CellValue::Formula { formula, .. } => {
            assert_eq!(formula, "1+2");
        },
        _ => panic!("Expected formula"),
    }
}

#[test]
fn test_dif_basic_read_write() {
    let dif_data = "TABLE\n0,1\n\"LITCHI\"\nVECTORS\n0,3\n\"\"\nTUPLES\n0,2\n\"\"\nDATA\n0,0\n\"\"\n-1,0\nBOT\n1,0\n\"Name\"\n1,0\n\"Age\"\n1,0\n\"City\"\n-1,0\nBOT\n1,0\n\"Alice\"\n0,25\nV\n1,0\n\"NYC\"\n-1,0\nEOD\n";

    let mut cursor = Cursor::new(dif_data.as_bytes());
    let data = read_dif(&mut cursor, DifConfig::default()).unwrap();

    assert_eq!(data.len(), 2);
    assert_eq!(data[0][0], CellValue::String("Name".to_string()));
    assert_eq!(data[1][1], CellValue::Float(25.0));

    // Test writing
    let mut output = Vec::new();
    write_dif(&data, &mut output, DifConfig::default()).unwrap();
    let output_str = String::from_utf8(output).unwrap();
    assert!(output_str.contains("TABLE"));
    assert!(output_str.contains("VECTORS"));
    assert!(output_str.contains("TUPLES"));
}

#[test]
fn test_dif_with_bom() {
    let mut output = Vec::new();
    let data = vec![vec![CellValue::Int(42)]];

    let config = DifConfig {
        strip_bom: true,
        write_bom: Some(BomKind::Utf8),
    };

    write_dif(&data, &mut output, config).unwrap();

    assert_eq!(&output[0..3], &[0xEF, 0xBB, 0xBF]);
}

#[test]
fn test_dif_booleans() {
    let dif_data = "TABLE\n0,1\n\"TEST\"\nVECTORS\n0,2\n\"\"\nTUPLES\n0,1\n\"\"\nDATA\n0,0\n\"\"\n-1,0\nBOT\n1,0\n\"TRUE\"\n1,0\n\"FALSE\"\n-1,0\nEOD\n";

    let mut cursor = Cursor::new(dif_data.as_bytes());
    let data = read_dif(&mut cursor, DifConfig::default()).unwrap();

    assert_eq!(data[0][0], CellValue::Bool(true));
    assert_eq!(data[0][1], CellValue::Bool(false));
}

#[test]
fn test_fixed_width_auto_detect() {
    let prn_data = "Name       Age  City\nAlice      25   NYC\nBob        30   LA\n";

    let mut cursor = Cursor::new(prn_data.as_bytes());
    let data = read_fixed_width(&mut cursor, FixedWidthConfig::default()).unwrap();

    assert_eq!(data.len(), 3);
    assert!(matches!(data[0][0], CellValue::String(_)));
}

#[test]
fn test_fixed_width_with_specified_widths() {
    let prn_data = "Alice  25NYC\nBob    30LA\n";

    let config = FixedWidthConfig {
        column_widths: vec![7, 2, 3],
        auto_detect_widths: false,
        ..Default::default()
    };

    let mut cursor = Cursor::new(prn_data.as_bytes());
    let data = read_fixed_width(&mut cursor, config).unwrap();

    assert_eq!(data.len(), 2);
    assert_eq!(data[0].len(), 3);
}

#[test]
fn test_fixed_width_write() {
    let data = vec![
        vec![
            CellValue::String("Name".to_string()),
            CellValue::String("Age".to_string()),
        ],
        vec![CellValue::String("Alice".to_string()), CellValue::Int(25)],
    ];

    let mut output = Vec::new();
    write_fixed_width(&data, &mut output, FixedWidthConfig::default()).unwrap();

    let output_str = String::from_utf8(output).unwrap();
    assert!(output_str.contains("Name"));
    assert!(output_str.contains("Alice"));
}

#[test]
fn test_fixed_width_with_bom() {
    let data = vec![vec![CellValue::Int(123)]];

    let config = FixedWidthConfig {
        write_bom: Some(BomKind::Utf8),
        ..Default::default()
    };

    let mut output = Vec::new();
    write_fixed_width(&data, &mut output, config).unwrap();

    assert_eq!(&output[0..3], &[0xEF, 0xBB, 0xBF]);
}

#[test]
fn test_delimited_csv() {
    let csv_data = "name,age,city\nAlice,25,NYC\nBob,30,LA\n";

    let mut cursor = Cursor::new(csv_data.as_bytes());
    let data = read_delimited(&mut cursor, DelimitedConfig::csv()).unwrap();

    assert_eq!(data.len(), 3);
    assert_eq!(data[0][0], CellValue::String("name".to_string()));
    assert_eq!(data[1][1], CellValue::Int(25));
}

#[test]
fn test_delimited_tsv() {
    let tsv_data = "name\tage\tcity\nAlice\t25\tNYC\n";

    let mut cursor = Cursor::new(tsv_data.as_bytes());
    let data = read_delimited(&mut cursor, DelimitedConfig::tsv()).unwrap();

    assert_eq!(data.len(), 2);
    assert_eq!(data[0][0], CellValue::String("name".to_string()));
}

#[test]
fn test_delimited_quoted_fields() {
    let csv_data = "\"Hello, World\",\"Value with \"\"quotes\"\"\"\n";

    let mut cursor = Cursor::new(csv_data.as_bytes());
    let data = read_delimited(&mut cursor, DelimitedConfig::csv()).unwrap();

    assert_eq!(data[0][0], CellValue::String("Hello, World".to_string()));
    assert_eq!(
        data[0][1],
        CellValue::String("Value with \"quotes\"".to_string())
    );
}

#[test]
fn test_delimited_write_with_bom() {
    let data = vec![vec![CellValue::String("Test".to_string())]];

    let config = DelimitedConfig {
        write_bom: Some(BomKind::Utf8),
        ..Default::default()
    };

    let mut output = Vec::new();
    write_delimited(&data, &mut output, config).unwrap();

    assert_eq!(&output[0..3], &[0xEF, 0xBB, 0xBF]);
    assert!(String::from_utf8_lossy(&output[3..]).contains("Test"));
}

#[test]
fn test_all_bom_variants() {
    let boms = vec![
        (BomKind::Utf8, vec![0xEF, 0xBB, 0xBF]),
        (BomKind::Utf16Le, vec![0xFF, 0xFE]),
        (BomKind::Utf16Be, vec![0xFE, 0xFF]),
        (BomKind::Utf32Le, vec![0xFF, 0xFE, 0x00, 0x00]),
        (BomKind::Utf32Be, vec![0x00, 0x00, 0xFE, 0xFF]),
    ];

    for (kind, expected_bytes) in boms {
        let data = vec![vec![CellValue::Int(1)]];
        let mut output = Vec::new();

        let config = DifConfig {
            strip_bom: true,
            write_bom: Some(kind),
        };

        write_dif(&data, &mut output, config).unwrap();
        assert_eq!(&output[..expected_bytes.len()], &expected_bytes[..]);
    }
}

#[test]
fn test_empty_data() {
    let empty: Vec<Vec<CellValue>> = Vec::new();

    let mut output = Vec::new();
    write_sylk(&empty, &mut output, SylkConfig::default()).unwrap();
    let output_str = String::from_utf8(output).unwrap();
    assert!(output_str.contains("ID;PWXL"));

    let mut output = Vec::new();
    write_dif(&empty, &mut output, DifConfig::default()).unwrap();
    let output_str = String::from_utf8(output).unwrap();
    assert!(output_str.contains("TABLE"));
}
