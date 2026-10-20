//! Tokenizer tests: the grammar `File::import` relies on.

use super::{Dialect, Tokenizer, declared_delimiter};
use crate::Error;

fn records(text: &str, dialect: Dialect) -> Result<Vec<Vec<String>>, Error> {
    let mut tokenizer = Tokenizer::new(text, dialect)?;
    let mut records = Vec::new();
    while tokenizer.next_record() {
        let mut fields = Vec::new();
        loop {
            fields.push(tokenizer.field()?.to_owned());
            if !tokenizer.next_field() {
                break;
            }
        }
        tokenizer.end_record();
        records.push(fields);
    }
    Ok(records)
}

#[test]
fn quoted_fields_hold_delimiters_breaks_and_doubled_quotes() {
    let parsed = records("\"a,b\",\"c\"\"d\",\"e\nf\",\"g\"h\n", Dialect::default()).unwrap();
    assert_eq!(parsed, [vec!["a,b", "c\"d", "e\nf", "gh"]]);
}

#[test]
fn blank_lines_are_records_and_the_last_break_is_not() {
    let parsed = records("a\n\nb\r\n", Dialect::default()).unwrap();
    assert_eq!(parsed, [vec!["a"], vec![""], vec!["b"]]);
}

#[test]
fn quote_none_reads_verbatim() {
    let parsed = records("\"a\",b", Dialect::default().quote(None)).unwrap();
    assert_eq!(parsed, [vec!["\"a\"", "b"]]);
}

#[test]
fn controls_are_refused_only_when_asked_with_their_line() {
    assert_eq!(
        records("a\u{1}", Dialect::default()).unwrap(),
        [vec!["a\u{1}"]]
    );
    let error = records(
        "ok\n\"x\ny\u{1}\"",
        Dialect::default().refuse_controls(true),
    )
    .unwrap_err();
    assert!(matches!(error, Error::InvalidCharacter { line: 3 }));
    let error = records("\u{FFFE}", Dialect::default().refuse_controls(true)).unwrap_err();
    assert!(matches!(error, Error::InvalidCharacter { line: 1 }));
}

#[test]
fn unterminated_quotes_report_their_starting_line() {
    let error = records("a\n\"b\nc", Dialect::default()).unwrap_err();
    assert!(matches!(error, Error::UnterminatedQuote { line: 2 }));
}

#[test]
fn escapes_and_comments_are_opt_in() {
    let dialect = Dialect::default().escapes(true).comment(Some(b'#'));
    let parsed = records("#c\r\n\"a\\\"b\\n\\q\"\n#d", dialect).unwrap();
    assert_eq!(parsed, [vec!["a\"b\n\\q"]]);
    let parsed = records("\"a\\\"", Dialect::default()).unwrap();
    assert_eq!(parsed, [vec!["a\\"]]);
}

#[test]
fn dialects_refuse_ambiguous_bytes() {
    for dialect in [
        Dialect::new(b'\n'),
        Dialect::new(b'"'),
        Dialect::new(b',').quote(Some(0)),
        Dialect::new(b',').comment(Some(b',')),
        Dialect::new(0xC3),
    ] {
        assert!(
            matches!(dialect.validate(), Err(Error::InvalidOption { .. })),
            "{dialect:?}"
        );
    }
}

#[test]
fn the_sep_line_declares_the_delimiter() {
    assert_eq!(declared_delimiter("sep=;\r\na"), (Some(b';'), "a"));
    assert_eq!(declared_delimiter("sep=|"), (Some(b'|'), ""));
    assert_eq!(declared_delimiter("sep=;;\n"), (None, "sep=;;\n"));
    assert_eq!(declared_delimiter("sep=\n"), (None, "sep=\n"));
}
