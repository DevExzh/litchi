use litchi_rtf::{EmbeddedFontFormat, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\uc1
{\fonttbl
{\f0\fswiss\fcharset0{\*\fontemb\fttruetype{\*\fontfile\cpg1252 arial.ttf} 00010000abCDef42}Arial;}
{\f1\fnil{\*\fontemb\ftnil{\*\fontfile symbol.fon}}Symbol;}
{\f2\fmodern Courier New;}}
\f0 Body}"#;

#[test]
fn parses_decodes_and_round_trips_embedded_fonts() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.font_table();

    let first = table.get(0).unwrap();
    let embedded = first.embedded.as_ref().unwrap();
    assert_eq!(embedded.format, EmbeddedFontFormat::TrueType);
    assert_eq!(embedded.file_name.as_deref(), Some("arial.ttf"));
    assert_eq!(embedded.file_code_page, Some(1252));
    assert_eq!(
        embedded.data.as_deref(),
        Some(&[0x00, 0x01, 0x00, 0x00, 0xab, 0xcd, 0xef, 0x42][..])
    );

    let second = table.get(1).unwrap();
    let embedded = second.embedded.as_ref().unwrap();
    assert_eq!(embedded.format, EmbeddedFontFormat::Nil);
    assert_eq!(embedded.file_name.as_deref(), Some("symbol.fon"));
    assert_eq!(embedded.file_code_page, None);
    assert_eq!(embedded.data, None);

    assert_eq!(table.get(2).unwrap().embedded, None);

    let mut first_pass = Vec::new();
    RtfWriter::new(&mut first_pass)
        .write_document(&doc)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first_pass).unwrap();
    assert_eq!(doc.font_table(), reparsed.font_table());
    let mut second_pass = Vec::new();
    RtfWriter::new(&mut second_pass)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_pass, second_pass);
}

#[test]
fn accepts_fontemb_without_ignorable_destination_marker() {
    let source = r#"{\rtf1{\fonttbl{\f0\fnil{\fontemb\fttruetype{\fontfile times.ttf}}Times;}}}"#;
    let doc = RtfDocument::parse(source).unwrap();
    let embedded = doc.font_table().get(0).unwrap().embedded.as_ref().unwrap();
    assert_eq!(embedded.format, EmbeddedFontFormat::TrueType);
    assert_eq!(embedded.file_name.as_deref(), Some("times.ttf"));
}

#[test]
fn rejects_malformed_embedded_fonts() {
    let malformed = [
        // Duplicate fontemb destination in one font entry.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil}{\*\fontemb\ftnil}Arial;}}}"#,
        // Duplicate embedded font format keyword.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\fttruetype\ftnil}Arial;}}}"#,
        // Odd number of hexadecimal digits.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\fttruetype 001}Arial;}}}"#,
        // Non-hexadecimal payload character.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\fttruetype 00zz}Arial;}}}"#,
        // Duplicate fontfile destination.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil{\*\fontfile a.ttf}{\*\fontfile b.ttf}}Arial;}}}"#,
        // Empty fontfile name.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil{\*\fontfile }}Arial;}}}"#,
        // Duplicate fontfile code page.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil{\*\fontfile\cpg1252\cpg1250 a.ttf}}Arial;}}}"#,
        // Out-of-range fontfile code page.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil{\*\fontfile\cpg70000 a.ttf}}Arial;}}}"#,
        // Nested group inside the fontfile name destination.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil{\*\fontfile {\field X}}}Arial;}}}"#,
        // Unterminated fontemb destination.
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\fontemb\ftnil"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}
