//! Parser, decoder and inference tests for CSV/TSV import.

use super::{Encoding, Error, Options, Table, Value, parse, serial};
use crate::Code;
use crate::transaction::SemanticLimits;

fn table(source: &[u8], options: &Options) -> Table {
    parse(source, options, &SemanticLimits::default()).expect("parse")
}

/// Every field as `(row, column, rendering)`.
fn cells(table: &Table) -> Vec<(u32, u32, String)> {
    table
        .fields()
        .iter()
        .map(|field| {
            let rendering = match &field.value {
                Value::Empty => "∅".to_owned(),
                Value::Text(range) => format!("t:{}", table.text(range)),
                Value::Number { value, lexical } => format!("n:{value}|{}", table.text(lexical)),
                Value::Bool(value) => format!("b:{value}"),
                Value::Date {
                    days,
                    seconds,
                    time,
                    ..
                } => format!("d:{days}:{seconds}:{time}"),
                Value::Formula(range) => format!("f:{}", table.text(range)),
            };
            (field.row, field.column, rendering)
        })
        .collect()
}

fn code(source: &[u8], options: &Options) -> Code {
    parse(source, options, &SemanticLimits::default())
        .expect_err("refused")
        .code()
}

#[test]
fn quoted_fields_delimiters_and_line_breaks() {
    let parsed = table(
        b"a,\"b,c\",\"d\"\"e\"\r\n\"multi\nline\",,x\n",
        &Options::csv().infer(false),
    );
    assert_eq!(
        cells(&parsed),
        [
            (0, 0, "t:a".to_owned()),
            (0, 1, "t:b,c".to_owned()),
            (0, 2, "t:d\"e".to_owned()),
            (1, 0, "t:multi\nline".to_owned()),
            (1, 1, "∅".to_owned()),
            (1, 2, "t:x".to_owned()),
        ]
    );
    assert_eq!((parsed.rows, parsed.columns), (2, 3));
}

#[test]
fn records_end_at_cr_lf_and_crlf_and_blank_lines_keep_positions() {
    let parsed = table(b"a\rb\r\n\nc", &Options::csv());
    assert_eq!(
        cells(&parsed),
        [
            (0, 0, "t:a".to_owned()),
            (1, 0, "t:b".to_owned()),
            (2, 0, "∅".to_owned()),
            (3, 0, "t:c".to_owned()),
        ]
    );
}

#[test]
fn ragged_records_and_trailing_delimiter() {
    let parsed = table(b"1,2,3\n4\n5,\n", &Options::csv());
    assert_eq!(parsed.columns, 3);
    assert_eq!(parsed.rows, 3);
    assert_eq!(parsed.fields().len(), 6);
    assert_eq!(parsed.fields()[5].value, Value::Empty);
}

#[test]
fn quote_tail_is_kept_and_unquoted_quotes_are_literal() {
    let parsed = table(b"\"a\"b,c\"d\n", &Options::csv());
    assert_eq!(
        cells(&parsed),
        [(0, 0, "t:ab".to_owned()), (0, 1, "t:c\"d".to_owned())]
    );
    let verbatim = table(b"\"a,b\"\n", &Options::csv().quote(None));
    assert_eq!(
        cells(&verbatim),
        [(0, 0, "t:\"a".to_owned()), (0, 1, "t:b\"".to_owned())]
    );
}

#[test]
fn unterminated_quote_reports_its_line() {
    let error = parse(b"a\n\"b\nc", &Options::csv(), &SemanticLimits::default()).unwrap_err();
    assert_eq!(error.code(), Code::InvalidInput);
    assert!(matches!(
        error.leaf(),
        crate::error::Source::Import(Error::UnterminatedQuote { line: 2 })
    ));
}

#[test]
fn separator_declaration_is_used_and_stripped() {
    let parsed = table(b"sep=;\na;b\n", &Options::csv());
    assert_eq!(
        cells(&parsed),
        [(0, 0, "t:a".to_owned()), (0, 1, "t:b".to_owned())]
    );
    // An explicit delimiter wins, and the declaration is still not data.
    let explicit = table(b"sep=;\na;b\n", &Options::csv().delimiter(','));
    assert_eq!(cells(&explicit), [(0, 0, "t:a;b".to_owned())]);
    assert!(table(b"sep=;", &Options::csv()).is_empty());
}

#[test]
fn tsv_and_custom_delimiters() {
    let parsed = table(b"a\tb,c\n", &Options::tsv());
    assert_eq!(
        cells(&parsed),
        [(0, 0, "t:a".to_owned()), (0, 1, "t:b,c".to_owned())]
    );
    let piped = table(b"1|2\n", &Options::csv().delimiter('|'));
    assert_eq!(piped.fields().len(), 2);
}

#[test]
fn invalid_options_are_arguments() {
    for options in [
        Options::csv().delimiter('"'),
        Options::csv().delimiter('\n'),
        Options::csv().delimiter('é'),
        Options::csv().quote(Some('\r')),
    ] {
        assert_eq!(code(b"a", &options), Code::InvalidArgument);
    }
}

#[test]
fn encodings_and_byte_order_marks() {
    let expected = vec![(0, 0, "t:é".to_owned()), (0, 1, "n:1|1".to_owned())];
    let utf8 = "é,1".as_bytes();
    let mut bom = vec![0xEF, 0xBB, 0xBF];
    bom.extend_from_slice(utf8);
    let le: Vec<u8> = "\u{FEFF}é,1"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let be: Vec<u8> = "\u{FEFF}é,1"
        .encode_utf16()
        .flat_map(u16::to_be_bytes)
        .collect();
    let be_bare: Vec<u8> = "é,1".encode_utf16().flat_map(u16::to_be_bytes).collect();
    assert_eq!(cells(&table(utf8, &Options::csv())), expected);
    assert_eq!(cells(&table(&bom, &Options::csv())), expected);
    assert_eq!(cells(&table(&le, &Options::csv())), expected);
    assert_eq!(cells(&table(&be, &Options::csv())), expected);
    assert_eq!(
        cells(&table(
            &be_bare,
            &Options::csv().encoding(Encoding::Utf16Be)
        )),
        expected
    );
    assert_eq!(
        cells(&table(&bom, &Options::csv().encoding(Encoding::Utf8))),
        expected
    );
}

#[test]
fn malformed_encodings_are_input_errors() {
    assert_eq!(
        code(&[b'a', 0xFF, b'b'], &Options::csv()),
        Code::InvalidInput
    );
    // An odd UTF-16 length and an unpaired surrogate.
    assert_eq!(
        code(&[0xFF, 0xFE, b'a'], &Options::csv()),
        Code::InvalidInput
    );
    assert_eq!(
        code(&[0xFF, 0xFE, 0x00, 0xD8, b'a', 0], &Options::csv()),
        Code::InvalidInput
    );
}

#[test]
fn characters_no_cell_stores_are_refused() {
    assert_eq!(code(b"a\0b", &Options::csv()), Code::InvalidInput);
    assert_eq!(code(b"\"a\x01\"", &Options::csv()), Code::InvalidInput);
    assert_eq!(
        code("a\u{FFFE}".as_bytes(), &Options::csv()),
        Code::InvalidInput
    );
    // Tabs inside a CSV field are ordinary text.
    assert_eq!(
        cells(&table(b"a\tb", &Options::csv())),
        [(0, 0, "t:a\tb".to_owned())]
    );
}

#[test]
fn number_inference_keeps_canonical_spellings() {
    let parsed = table(
        b"1,-2.5,1e3,+4,007,.5,5.,1E-2,-0,123456789012345678901,1.2.3,1e,,0x10,inf,NaN,1e400,1e-320",
        &Options::csv(),
    );
    let rendered: Vec<String> = cells(&parsed).into_iter().map(|cell| cell.2).collect();
    assert_eq!(
        rendered,
        [
            "n:1|1",
            "n:-2.5|-2.5",
            "n:1000|1e3",
            "n:4|4",
            "n:7|7",
            "n:0.5|0.5",
            "n:5|5",
            "n:0.01|1E-2",
            "n:0|0",
            "n:123456789012345680000|123456789012345678901",
            "t:1.2.3",
            "t:1e",
            "∅",
            "t:0x10",
            "t:inf",
            "t:NaN",
            "t:1e400",
            "t:1e-320",
        ]
    );
}

#[test]
fn booleans_dates_and_order() {
    let parsed = table(
        b"TRUE,false,True,2026-10-13,2026-10-13T08:30:00,2026-10-13 08:30:00.5Z,2026-02-30,2026-10-13T24:00:00,1",
        &Options::csv(),
    );
    let rendered: Vec<String> = cells(&parsed).into_iter().map(|cell| cell.2).collect();
    assert_eq!(
        rendered,
        [
            "b:true",
            "b:false",
            "b:true",
            "d:20739:0:false",
            "d:20739:30600:true",
            "d:20739:30600.5:true",
            "t:2026-02-30",
            "t:2026-10-13T24:00:00",
            "n:1|1",
        ]
    );
}

#[test]
fn inference_can_be_switched_off_per_kind_and_header_stays_text() {
    let source = b"n,b,d\n1,TRUE,2026-01-01\n";
    let none = table(source, &Options::csv().infer(false));
    assert!(
        none.fields()
            .iter()
            .all(|field| matches!(field.value, Value::Text(_)))
    );
    let dates_off = table(source, &Options::csv().infer_dates(false));
    assert!(matches!(dates_off.fields()[5].value, Value::Text(_)));
    assert!(matches!(dates_off.fields()[3].value, Value::Number { .. }));
    let header = table(b"1,TRUE\n1,TRUE\n", &Options::csv().header(true));
    assert!(matches!(header.fields()[0].value, Value::Text(_)));
    assert!(matches!(header.fields()[1].value, Value::Text(_)));
    assert!(matches!(header.fields()[2].value, Value::Number { .. }));
}

#[test]
fn formulas_are_opt_in() {
    let off = table(b"=SUM(A1:A2),=\n", &Options::csv());
    assert_eq!(
        cells(&off),
        [(0, 0, "t:=SUM(A1:A2)".to_owned()), (0, 1, "t:=".to_owned())]
    );
    let on = table(b"=SUM(A1:A2),=\n", &Options::csv().formulas(true));
    assert_eq!(
        cells(&on),
        [(0, 0, "f:SUM(A1:A2)".to_owned()), (0, 1, "t:=".to_owned())]
    );
}

#[test]
fn limits_bound_fields_and_bytes() {
    let limits = SemanticLimits::builder()
        .max_expanded_targets(3)
        .unwrap()
        .build()
        .unwrap();
    let error = parse(b"1,2\n3,4\n", &Options::csv(), &limits).unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    let bytes = SemanticLimits::builder()
        .max_staged_bytes(100)
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(
        parse(&[b'a'; 101], &Options::csv(), &bytes)
            .unwrap_err()
            .code(),
        Code::LimitExceeded
    );
    // Two fields carry two fixed charges.
    assert_eq!(
        parse(b"a,b", &Options::csv(), &bytes).unwrap_err().code(),
        Code::LimitExceeded
    );
}

#[test]
fn over_long_fields_are_values() {
    let long = vec![b'x'; super::MAX_CELL_CHARS + 1];
    assert_eq!(code(&long, &Options::csv()), Code::InvalidValue);
    let exact = vec![b'x'; super::MAX_CELL_CHARS];
    assert_eq!(table(&exact, &Options::csv()).fields().len(), 1);
}

#[test]
fn empty_input_is_empty() {
    assert!(table(b"", &Options::csv()).is_empty());
    assert!(table(&[0xEF, 0xBB, 0xBF], &Options::csv()).is_empty());
}

#[test]
fn serials_follow_both_date_systems() {
    let day = |text: &str| match &table(text.as_bytes(), &Options::csv()).fields()[0].value {
        Value::Date { days, seconds, .. } => (*days, *seconds),
        other => panic!("not a date: {other:?}"),
    };
    let (days, seconds) = day("1900-01-01");
    assert_eq!(serial(days, seconds, false), Some(1.0));
    let (days, seconds) = day("1900-02-28");
    assert_eq!(serial(days, seconds, false), Some(59.0));
    let (days, seconds) = day("1900-03-01");
    assert_eq!(serial(days, seconds, false), Some(61.0));
    let (days, seconds) = day("2026-10-13T12:00:00");
    assert_eq!(serial(days, seconds, false), Some(46_308.5));
    assert_eq!(serial(days, seconds, true), Some(46_308.5 - 1_462.0));
    let (days, seconds) = day("1899-12-31");
    assert_eq!(serial(days, seconds, false), None);
    let (days, seconds) = day("1903-12-31");
    assert_eq!(serial(days, seconds, true), None);
}
