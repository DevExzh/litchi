//! Workbook, format and limit tests.

use litchi_core::BomKind;
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet as _};

use crate::delimited::Options;
use crate::{
    Error, Format, LimitKind, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
    Workbook, dif, fixed_width, sylk,
};

fn text(value: &str) -> CellValue {
    CellValue::String(value.to_owned())
}

#[test]
fn csv_workbook_reads_one_sheet() {
    let workbook =
        Workbook::from_bytes(b"name,age,city\nJohn,25,New York\nJane,30,London").unwrap();
    let sheet = workbook.active_worksheet().unwrap();
    assert_eq!(sheet.row_count(), 3);
    assert_eq!(sheet.column_count(), 3);
    assert_eq!(sheet.cell(1, 1).unwrap().value(), &text("name"));
    assert_eq!(sheet.cell(2, 2).unwrap().value(), &CellValue::Int(25));
    assert_eq!(
        sheet.cell_by_coordinate("C2").unwrap().value(),
        &text("New York")
    );
    assert_eq!(
        sheet.row(1).unwrap()[..],
        [text("John"), CellValue::Int(25), text("New York")]
    );
    assert!(sheet.row(3).is_err());
    assert_eq!(workbook.worksheet_names(), ["Sheet1"]);
    assert!(workbook.worksheet_by_name("Other").is_err());
    assert!(workbook.worksheet_by_index(1).is_err());
}

#[test]
fn tsv_quotes_empty_cells_and_inference() {
    let tsv = Workbook::from_bytes_with_options(b"a\tb\n1\t2", Options::tsv()).unwrap();
    assert_eq!(tsv.rows()[1], [CellValue::Int(1), CellValue::Int(2)]);

    let quoted =
        Workbook::from_bytes(b"\"Hello, World\",\"Value with \"\"quotes\"\"\",\"Normal\"").unwrap();
    assert_eq!(
        quoted.rows()[0],
        [
            text("Hello, World"),
            text("Value with \"quotes\""),
            text("Normal")
        ]
    );

    let empty = Workbook::from_bytes(b"a,,c\n,,").unwrap();
    assert_eq!(
        empty.rows()[1],
        [CellValue::Empty, CellValue::Empty, CellValue::Empty]
    );

    let inferred =
        Workbook::from_bytes(b"int,float,bool,string\n42,3.15,true,hello\n,2.0,false,").unwrap();
    assert_eq!(
        inferred.rows()[1],
        [
            CellValue::Int(42),
            CellValue::Float(3.15),
            CellValue::Bool(true),
            text("hello")
        ]
    );
    assert_eq!(
        inferred.rows()[2],
        [
            CellValue::Empty,
            CellValue::Float(2.0),
            CellValue::Bool(false),
            CellValue::Empty
        ]
    );
}

#[test]
fn iterators_visit_every_row_and_cell() {
    let workbook = Workbook::from_bytes(b"a,b,c\n1,2,3\n4,5,6").unwrap();
    let sheet = workbook.worksheet();
    let mut count = 0;
    for row in sheet.rows() {
        assert_eq!(row.unwrap().len(), 3);
        count += 1;
    }
    assert_eq!(count, 3);
    let mut seen = Vec::new();
    for cell in sheet.cells() {
        let cell = cell.unwrap();
        seen.push(cell.coordinate());
    }
    assert_eq!(seen.len(), 9);
    assert_eq!(seen[4], "B2");
    let mut sheets = workbook.worksheets();
    assert!(sheets.next().is_some());
    assert!(sheets.next().is_none());
}

#[test]
fn comments_escapes_and_trimming_follow_the_options() {
    let source = b"# heading\n\"a\\tb\", x \n";
    let workbook = Workbook::from_bytes_with_options(source, Options::csv().trim(true)).unwrap();
    assert_eq!(workbook.rows(), [vec![text("a\tb"), text("x")]]);

    let plain = Options::csv().comment(None).escapes(false);
    let workbook = Workbook::from_bytes_with_options(source, plain).unwrap();
    assert_eq!(workbook.rows()[0], [text("# heading")]);
    assert_eq!(workbook.rows()[1], [text("a\\tb"), text(" x ")]);
}

#[test]
fn records_end_at_every_line_break_and_quotes_open_fields_only() {
    let workbook = Workbook::from_bytes(b"a\rb\r\nc\nx\"y,\"q\"z").unwrap();
    assert_eq!(
        workbook.rows(),
        [
            vec![text("a")],
            vec![text("b")],
            vec![text("c")],
            vec![text("x\"y"), text("qz")]
        ]
    );
}

#[test]
fn write_round_trips_values_that_need_quotes() {
    let rows = vec![
        vec![text("#not a comment"), text("a,b"), text("back\\slash")],
        vec![
            text("line\nbreak"),
            CellValue::Bool(true),
            CellValue::Float(1.5),
        ],
    ];
    let workbook = Workbook::new(rows.clone(), Options::csv());
    let bytes = workbook.to_bytes().unwrap();
    let read = Workbook::from_bytes(&bytes).unwrap();
    assert_eq!(read.rows(), rows);
}

#[test]
fn byte_order_marks_are_detected_and_written() {
    let workbook = Workbook::from_bytes(b"\xEF\xBB\xBFa,b").unwrap();
    assert_eq!(workbook.bom(), Some(BomKind::Utf8));
    assert_eq!(workbook.rows()[0][0], text("a"));

    let utf16: Vec<u8> = [0xFF, 0xFE]
        .into_iter()
        .chain("x,1".encode_utf16().flat_map(u16::to_le_bytes))
        .collect();
    let workbook = Workbook::from_bytes(&utf16).unwrap();
    assert_eq!(workbook.bom(), Some(BomKind::Utf16Le));
    assert_eq!(workbook.rows()[0], [text("x"), CellValue::Int(1)]);

    let utf16 = Options::csv().write_bom(Some(BomKind::Utf16Be));
    let mut out = Vec::new();
    crate::delimited::write(&[vec![text("é"), CellValue::Int(2)]], &mut out, &utf16).unwrap();
    assert_eq!(out, [0xFE, 0xFF, 0x00, 0xE9, 0x00, b',', 0x00, b'2']);
    let workbook = Workbook::from_bytes(&out).unwrap();
    assert_eq!(workbook.bom(), Some(BomKind::Utf16Be));
    assert_eq!(workbook.rows()[0], [text("é"), CellValue::Int(2)]);
    let mut out = Vec::new();
    assert!(matches!(
        dif::write(
            &[],
            &mut out,
            &dif::Options::default().write_bom(Some(BomKind::Utf32Le))
        ),
        Err(Error::InvalidOption { .. })
    ));

    let options = Options::csv().write_bom(Some(BomKind::Utf8));
    let mut out = Vec::new();
    crate::delimited::write(&[vec![CellValue::Int(1)]], &mut out, &options).unwrap();
    assert_eq!(out, b"\xEF\xBB\xBF1");
}

#[test]
fn malformed_input_is_a_typed_error() {
    assert!(matches!(
        Workbook::from_bytes(b"\"open"),
        Err(Error::UnterminatedQuote { line: 1 })
    ));
    assert!(matches!(
        Workbook::from_bytes(b"ok\n\xFF"),
        Err(Error::Encoding { offset: 3 })
    ));
    assert!(matches!(
        Workbook::from_bytes_with_options(b"a", Options::csv().delimiter(b'"')),
        Err(Error::InvalidOption { .. })
    ));
}

#[test]
fn limits_bound_input_bytes_and_cells() {
    let small = SourceLimits::builder()
        .max_input_bytes(4)
        .unwrap()
        .build()
        .unwrap();
    let error = Workbook::from_bytes_with_options(b"a,b,c", Options::csv().source_limits(small))
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Source(SourceLimitKind::InputBytes),
            observed: 5,
            maximum: 4
        }
    ));

    let cells = SemanticLimits::builder()
        .max_cells(2)
        .unwrap()
        .build()
        .unwrap();
    let error = Workbook::from_bytes_with_options(b"a,b,c", Options::csv().semantic_limits(cells))
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Cells),
            ..
        }
    ));

    // A sparse SYLK cell charges every padding cell before it.
    let sylk_options = sylk::Options::default().semantic_limits(cells);
    assert!(matches!(
        sylk::read(&b"C;Y1;X5;K1\n"[..], &sylk_options),
        Err(Error::Limit { .. })
    ));
    assert!(matches!(
        sylk::read(&b"ID\nC;Y99999999;X1;K1\n"[..], &sylk::Options::default()),
        Err(Error::OutOfGrid {
            format: Format::Sylk,
            line: 2
        })
    ));
    assert!(SourceLimits::builder().max_input_bytes(0).is_err());
    assert!(SemanticLimits::builder().max_cells(0).is_err());
}

#[test]
fn sylk_reads_values_formulas_and_cached_results() {
    let source = "ID;PWXL;N;E\nB;Y3;X3\nC;Y1;X1;K\"Name\"\nC;Y1;X2;K\"Age\"\nC;Y2;X1;K\"Alice\"\nC;Y2;X2;K25\nC;Y3;X1;K3;E1+2\nE\n";
    let rows = sylk::read(source.as_bytes(), &sylk::Options::default()).unwrap();
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[0], [text("Name"), text("Age")]);
    assert_eq!(rows[1][1], CellValue::Int(25));
    assert_eq!(
        rows[2][0],
        CellValue::Formula {
            formula: "1+2".to_owned(),
            cached_value: Some(Box::new(CellValue::Int(3))),
            is_array: false,
            array_range: None,
        }
    );

    let special = vec![vec![text("a;b \"q\""), CellValue::Int(-3)]];
    let mut out = Vec::new();
    sylk::write(&special, &mut out, &sylk::Options::default()).unwrap();
    assert!(String::from_utf8_lossy(&out).contains("K\"a;;b \"\"q\"\"\""));
    assert_eq!(
        sylk::read(&out[..], &sylk::Options::default()).unwrap(),
        special
    );

    let mut out = Vec::new();
    sylk::write(
        &rows,
        &mut out,
        &sylk::Options::default().write_bom(Some(BomKind::Utf8)),
    )
    .unwrap();
    assert_eq!(&out[..3], [0xEF, 0xBB, 0xBF]);
    let written = String::from_utf8(out[3..].to_vec()).unwrap();
    assert!(written.contains("C;Y1;X1;K\"Name\""));
    assert!(written.contains("C;Y3;X1;E1+2"));
}

#[test]
fn dif_reads_rows_booleans_and_missing_data() {
    let source = "TABLE\n0,1\n\"LITCHI\"\nVECTORS\n0,3\n\"\"\nTUPLES\n0,2\n\"\"\nDATA\n0,0\n\"\"\n-1,0\nBOT\n1,0\n\"Name\"\n1,0\n\"TRUE\"\n1,0\n\"NA\"\n-1,0\nBOT\n1,0\n\"Alice\"\n0,25\nV\n-1,0\nEOD\n";
    let rows = dif::read(source.as_bytes(), &dif::Options::default()).unwrap();
    assert_eq!(
        rows,
        [
            vec![text("Name"), CellValue::Bool(true), CellValue::Empty],
            vec![text("Alice"), CellValue::Float(25.0)]
        ]
    );
    let mut out = Vec::new();
    dif::write(&rows, &mut out, &dif::Options::default()).unwrap();
    assert_eq!(dif::read(&out[..], &dif::Options::default()).unwrap(), rows);
    assert!(matches!(
        dif::read(&b"TABLE\n0,1\n"[..], &dif::Options::default()),
        Err(Error::Syntax {
            format: Format::Dif,
            line: 2
        })
    ));
}

#[test]
fn fixed_width_detects_and_cuts_by_characters() {
    let rows = fixed_width::read(
        "Name   Age  City\nAlice  25   NYC\nBob    30   LA\n".as_bytes(),
        &fixed_width::Options::default(),
    )
    .unwrap();
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[1], [text("Alice"), CellValue::Int(25), text("NYC")]);

    let options = fixed_width::Options::default().column_widths(vec![3, 2]);
    let rows = fixed_width::read("éèà12\nabc34\n".as_bytes(), &options).unwrap();
    assert_eq!(rows[0], [text("éèà"), CellValue::Int(12)]);

    let mut out = Vec::new();
    fixed_width::write(&[vec![text("ééé"), CellValue::Int(7)]], &mut out, &options).unwrap();
    assert_eq!(String::from_utf8(out).unwrap(), "ééé7\n");
    let mut out = Vec::new();
    fixed_width::write(
        &[vec![text("wide"), CellValue::Int(7)]],
        &mut out,
        &fixed_width::Options::default(),
    )
    .unwrap();
    assert_eq!(String::from_utf8(out).unwrap(), "wide  7\n");
}
