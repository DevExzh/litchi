#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "limit fixtures favor explicit inputs and panic-driven assertions"
)]

//! Every bounded string field of the fresh XLS writer at its limit: one code
//! unit below and exactly at the limit are written whole and read back, and
//! one unit past it is refused with a typed error, before any output and
//! without changing the writer. Until change 0757 the shared-string, formula
//! string, number-format and worksheet-name encoders truncated such a string
//! instead, and the defined-name encoder miscounted or misencoded non-ASCII
//! names.

use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait as _};
use litchi_xls::Error;
use litchi_xls::writer::{PageSetupOptions, Writer};

type Workbook = litchi_xls::Workbook<Cursor<Vec<u8>>>;

fn written(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read(bytes: Vec<u8>) -> Workbook {
    Workbook::new(Cursor::new(bytes)).unwrap()
}

/// The field, length and limit of a [`Error::StringTooLong`] refusal.
fn too_long<T: std::fmt::Debug>(result: Result<T, Error>) -> (&'static str, usize, usize) {
    match result {
        Err(Error::StringTooLong {
            field,
            utf16_units,
            limit,
        }) => (field, utf16_units, limit),
        other => panic!("expected StringTooLong, got {other:?}"),
    }
}

/// Writes `writer`, requiring a [`Error::StringTooLong`] refusal that leaves
/// the destination untouched.
fn refused_write(writer: &mut Writer) -> (&'static str, usize, usize) {
    let mut output = Cursor::new(Vec::new());
    let result = writer.write_to(&mut output);
    assert!(
        output.get_ref().is_empty(),
        "a refused write produced output"
    );
    too_long(result)
}

/// Strings of exactly `units` UTF-16 code units: ASCII (compressed on the
/// wire), Latin-1 and CJK (two and three UTF-8 bytes per unit), and one that
/// ends with a surrogate pair.
fn strings_of(units: usize) -> Vec<String> {
    vec![
        "b".repeat(units),
        "é".repeat(units),
        "漢".repeat(units),
        format!("{}😀", "a".repeat(units - 2)),
    ]
}

fn cell_text(workbook: &Workbook, sheet: usize, row: u32, column: u32) -> CellValue {
    workbook
        .xls_worksheet(sheet)
        .unwrap()
        .get_cell(row, column)
        .unwrap()
        .value()
        .clone()
}

#[test]
fn shared_strings_up_to_65535_units_are_written_whole_and_one_more_is_refused() {
    for units in [65_534, 65_535] {
        for value in strings_of(units) {
            let mut writer = Writer::new();
            let sheet = writer.add_worksheet("Text").unwrap();
            writer.write_string(sheet, 0, 0, &value).unwrap();
            writer.write_string(sheet, 1, 0, "after").unwrap();
            let workbook = read(written(&mut writer));
            assert_eq!(cell_text(&workbook, 0, 0, 0), CellValue::String(value));
            assert_eq!(
                cell_text(&workbook, 0, 1, 0),
                CellValue::String("after".to_string())
            );
        }
    }
    for value in strings_of(65_536) {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Text").unwrap();
        writer.write_string(sheet, 0, 0, "kept").unwrap();
        assert_eq!(
            too_long(writer.write_string(sheet, 0, 0, &value)),
            ("shared string", 65_536, 65_535)
        );
        assert_eq!(
            too_long(writer.write_string_with_format(sheet, 2, 0, &value, 0)),
            ("shared string", 65_536, 65_535)
        );
        // The refused writes changed nothing.
        let workbook = read(written(&mut writer));
        assert_eq!(
            cell_text(&workbook, 0, 0, 0),
            CellValue::String("kept".to_string())
        );
        assert!(workbook.xls_worksheet(0).unwrap().get_cell(2, 0).is_none());
    }
}

/// The 0753 review's reproduction. Cutting this string at 0xFFFF code units
/// left the emoji's high surrogate alone at the end of the SST entry, and the
/// crate's own reader then refused the whole workbook ("lone surrogate
/// found"). The writer refuses the string instead; one `a` fewer fits.
#[test]
fn a_string_whose_surrogate_pair_straddles_the_limit_is_refused_not_split() {
    let straddling = format!("{}😀", "a".repeat(65_534));
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Pair").unwrap();
    assert_eq!(
        too_long(writer.write_string(sheet, 0, 0, &straddling)),
        ("shared string", 65_536, 65_535)
    );

    let fitting = format!("{}😀", "a".repeat(65_533));
    writer.write_string(sheet, 0, 0, &fitting).unwrap();
    let workbook = read(written(&mut writer));
    assert_eq!(cell_text(&workbook, 0, 0, 0), CellValue::String(fitting));
}

#[test]
fn worksheet_names_up_to_31_units_are_written_whole_and_one_more_is_refused() {
    for units in [30, 31] {
        for name in strings_of(units) {
            let mut writer = Writer::new();
            writer.add_worksheet(&name).unwrap();
            let workbook = read(written(&mut writer));
            assert_eq!(workbook.worksheet_names(), [name]);
        }
    }
    for name in strings_of(32) {
        let mut writer = Writer::new();
        writer.add_worksheet("First").unwrap();
        assert_eq!(
            too_long(writer.add_worksheet(&name)),
            ("worksheet name", 32, 31)
        );
        assert_eq!(writer.worksheet_count(), 1);
    }
    // 31 UTF-16 code units in more than 31 UTF-8 bytes used to be refused by
    // byte length; BIFF8 counts code units.
    let mut writer = Writer::new();
    writer.add_worksheet(&"é".repeat(31)).unwrap();
    assert!(matches!(
        writer.add_worksheet(""),
        Err(Error::InvalidData(_))
    ));
}

#[test]
fn defined_names_up_to_255_units_are_written_whole_and_one_more_is_refused() {
    for units in [254, 255] {
        for name in strings_of(units) {
            let mut writer = Writer::new();
            writer.add_worksheet("Data").unwrap();
            writer.define_name(&name, "A1:B2").unwrap();
            let workbook = read(written(&mut writer));
            let names: Vec<&str> = workbook
                .defined_names()
                .iter()
                .map(|defined| defined.name.as_str())
                .collect();
            assert_eq!(names, [name.as_str()]);
        }
    }
    for name in strings_of(256) {
        let mut writer = Writer::new();
        writer.add_worksheet("Data").unwrap();
        assert_eq!(
            too_long(writer.define_name(&name, "A1")),
            ("defined name", 256, 255)
        );
        assert_eq!(
            too_long(writer.define_name_local(&name, "A1", 0)),
            ("defined name", 256, 255)
        );
        assert_eq!(
            too_long(writer.define_name_with_comment(&name, "A1", "note")),
            ("defined name", 256, 255)
        );
        assert!(writer.named_ranges().is_empty());
    }
}

/// Latin-1 names were written as UTF-8 bytes under a compressed-string flag,
/// and supplementary-plane names with a character count instead of a UTF-16
/// count; both desynchronised the record.
#[test]
fn non_ascii_defined_names_round_trip() {
    let mut writer = Writer::new();
    writer.add_worksheet("Data").unwrap();
    let names = [
        "Café_Total",
        "Größe",
        "日本語の名前",
        "Rocket😀Launch",
        "ascii_name",
    ];
    for name in names {
        writer.define_name(name, "A1:B2").unwrap();
    }
    let workbook = read(written(&mut writer));
    let read_back: Vec<&str> = workbook
        .defined_names()
        .iter()
        .map(|defined| defined.name.as_str())
        .collect();
    assert_eq!(read_back, names);
}

#[test]
fn defined_name_comments_up_to_255_units_are_written_and_one_more_is_refused() {
    for units in [254, 255] {
        for comment in strings_of(units) {
            let mut writer = Writer::new();
            writer.add_worksheet("Data").unwrap();
            writer
                .define_name_with_comment("Noted", "A1", &comment)
                .unwrap();
            let workbook = read(written(&mut writer));
            assert_eq!(
                workbook.defined_names()[0].comment.as_deref(),
                Some(comment.as_str())
            );
        }
    }
    for comment in strings_of(256) {
        let mut writer = Writer::new();
        writer.add_worksheet("Data").unwrap();
        writer
            .define_name_with_comment("Noted", "A1", &comment)
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        let result = writer.write_to(&mut output);
        assert!(
            matches!(result, Err(Error::InvalidData(ref message)) if message.contains("NameCmt")),
            "{result:?}"
        );
        assert!(output.get_ref().is_empty());
    }
}

/// Decodes the `PtgStr` a one-token formula consists of.
fn formula_string(workbook: &Workbook) -> String {
    let tokens = workbook
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .unwrap()
        .formula_bytes()
        .unwrap()
        .to_vec();
    assert_eq!(tokens[0], 0x17, "PtgStr");
    let units = usize::from(tokens[1]);
    if tokens[2] == 0 {
        assert_eq!(tokens.len(), 3 + units);
        tokens[3..].iter().map(|&byte| char::from(byte)).collect()
    } else {
        assert_eq!(tokens.len(), 3 + 2 * units);
        let wide: Vec<u16> = tokens[3..]
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16(&wide).unwrap()
    }
}

#[test]
fn formula_string_literals_up_to_255_units_are_written_whole_and_one_more_is_refused() {
    for units in [254, 255] {
        for literal in strings_of(units) {
            let mut writer = Writer::new();
            let sheet = writer.add_worksheet("Formulas").unwrap();
            writer
                .write_formula(sheet, 0, 0, &format!("\"{literal}\""))
                .unwrap();
            let workbook = read(written(&mut writer));
            assert_eq!(formula_string(&workbook), literal);
        }
    }
    for literal in strings_of(256) {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Formulas").unwrap();
        // Formula text is tokenized when the workbook is written.
        writer
            .write_formula(sheet, 0, 0, &format!("\"{literal}\""))
            .unwrap();
        assert_eq!(
            refused_write(&mut writer),
            ("formula string literal", 256, 255)
        );
    }
}

#[test]
fn number_formats_up_to_255_units_are_written_whole_and_one_more_is_refused() {
    for units in [254, 255] {
        for pattern in strings_of(units) {
            let mut writer = Writer::new();
            writer.add_worksheet("Formats").unwrap();
            let id = writer.register_number_format(&pattern);
            let workbook = read(written(&mut writer));
            assert!(
                workbook
                    .number_formats()
                    .iter()
                    .any(|format| format.id() == id && format.code() == pattern),
                "format {id} did not round-trip"
            );
        }
    }
    for pattern in strings_of(256) {
        let mut writer = Writer::new();
        writer.add_worksheet("Formats").unwrap();
        writer.register_number_format(&pattern);
        assert_eq!(refused_write(&mut writer), ("number format", 256, 255));
    }
}

#[test]
fn comment_authors_and_texts_keep_their_limits() {
    for (author, text) in [
        ("a".repeat(53), "b".repeat(65_534)),
        ("é".repeat(54), format!("{}😀", "c".repeat(65_533))),
    ] {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Notes").unwrap();
        writer.add_comment(sheet, 0, 0, &author, &text).unwrap();
        let workbook = read(written(&mut writer));
        let comments = workbook.xls_worksheet(0).unwrap().comments();
        assert_eq!(comments[0].author(), author);
        assert_eq!(comments[0].text(), text);
    }
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Notes").unwrap();
    assert!(matches!(
        writer.add_comment(sheet, 0, 0, &"a".repeat(55), "text"),
        Err(Error::InvalidData(message)) if message.contains("1..=54")
    ));
    for text in ["b".repeat(65_536), format!("{}😀", "c".repeat(65_534))] {
        assert!(matches!(
            writer.add_comment(sheet, 0, 0, "author", &text),
            Err(Error::InvalidData(message)) if message.contains("65535")
        ));
    }
}

#[test]
fn headers_and_footers_keep_their_limits() {
    for units in [254, 255] {
        for text in strings_of(units) {
            let mut writer = Writer::new();
            let sheet = writer.add_worksheet("Print").unwrap();
            writer
                .set_page_setup(
                    sheet,
                    PageSetupOptions {
                        header: text.clone(),
                        footer: text.clone(),
                        ..PageSetupOptions::default()
                    },
                )
                .unwrap();
            let workbook = read(written(&mut writer));
            let setup = workbook.xls_worksheet(0).unwrap().page_setup().unwrap();
            assert_eq!(setup.header(), text);
            assert_eq!(setup.footer(), text);

            let mut even = litchi_xls::HeaderFooter::default();
            even.set_even(text.clone(), text.clone()).unwrap();
        }
    }
    for text in strings_of(256) {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Print").unwrap();
        for options in [
            PageSetupOptions {
                header: text.clone(),
                ..PageSetupOptions::default()
            },
            PageSetupOptions {
                footer: text.clone(),
                ..PageSetupOptions::default()
            },
        ] {
            assert!(matches!(
                writer.set_page_setup(sheet, options),
                Err(Error::InvalidData(message)) if message.contains("255")
            ));
        }
        let mut even = litchi_xls::HeaderFooter::default();
        assert!(even.set_even(text.clone(), String::new()).is_err());
        assert!(even.set_first(String::new(), text.clone()).is_err());
    }
}
