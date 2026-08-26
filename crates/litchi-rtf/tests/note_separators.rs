#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{NoteSeparatorElement, NoteSeparatorKind, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r"{\rtf1\ansi\ansicpg1250\uc1
{\*\ftnsep \'8a\chftnsep A\par B}
{\*\ftnsepc \chftnsepc C\line D}
{\*\ftncn Notice \u20320?}
{\*\aftnsep E\chftnsep\par}
{\*\aftnsepc EC\chftnsepc\par}
{\*\aftncn End}
Body}";

const FONT_SCOPED_SEPARATOR: &str = r"{\rtf1\ansi\ansicpg1252
{\fonttbl{\f0\fnil\fcharset0 ANSI;}{\f1\fnil\fcharset128 JIS;}}
{\*\ftnsep\f1\'82\'a0\chftnsep}\'e9 Body}";

#[test]
fn parses_decodes_and_round_trips_note_separators() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.note_separators();
    assert_eq!(table.entries().len(), 6);
    let first = table.get(NoteSeparatorKind::FootnoteSeparator).unwrap();
    assert!(
        first
            .elements
            .iter()
            .any(|element| matches!(element, NoteSeparatorElement::Text(text) if text == "Š"))
    );
    assert!(
        first
            .elements
            .contains(&NoteSeparatorElement::SeparatorMark)
    );
    assert!(
        first
            .elements
            .contains(&NoteSeparatorElement::ParagraphBreak)
    );
    let notice = table
        .get(NoteSeparatorKind::FootnoteContinuationNotice)
        .unwrap();
    assert!(notice.elements.iter().any(
        |element| matches!(element, NoteSeparatorElement::Text(text) if text.contains("Notice 你"))
    ));
    assert_eq!(doc.text().trim(), "Body");

    let mut first_bytes = Vec::new();
    RtfWriter::new(&mut first_bytes)
        .write_document(&doc)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first_bytes).unwrap();
    assert_eq!(table, reparsed.note_separators());
    let mut second_bytes = Vec::new();
    RtfWriter::new(&mut second_bytes)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_bytes, second_bytes);
}

#[test]
fn rejects_malformed_note_separators() {
    let malformed = [
        r"{\rtf1{\ftnsep\chftnsep}}",
        r"{\rtf1{\*\ftnsep\chftnsep}{\*\ftnsep\chftnsep}}",
        r"{\rtf1{\*\aftnsep X}{\*\ftnsep X}}",
        r"{\rtf1 Body{\*\ftnsep X}}",
        r"{\rtf1{\*\ftnsep{\field X}}}",
        r"{\rtf1{\*\ftnsep{\object X}}}",
        r"{\rtf1{\*\ftnsep\bin2 AB}}",
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn parses_real_libreoffice_note_separator_fixture() {
    let fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-3.rtf"
    );
    let marker = br"{\*\ftnsep";
    let start = fixture
        .windows(marker.len())
        .position(|window| window == marker)
        .unwrap();
    let mut depth = 0usize;
    let mut end = None;
    for (offset, byte) in fixture[start..].iter().enumerate() {
        match byte {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(start + offset + 1);
                    break;
                }
            },
            _ => {},
        }
    }
    let mut isolated = br"{\rtf1\ansi\ansicpg1252".to_vec();
    isolated.extend_from_slice(&fixture[start..end.unwrap()]);
    isolated.push(b'}');
    let doc = RtfDocument::parse_bytes(&isolated).unwrap();
    let separator = doc
        .note_separators()
        .get(NoteSeparatorKind::FootnoteSeparator)
        .unwrap();
    assert!(
        separator
            .elements
            .contains(&NoteSeparatorElement::SeparatorMark)
    );
    assert!(
        separator
            .elements
            .contains(&NoteSeparatorElement::ParagraphBreak)
    );
}

#[test]
fn starred_note_separator_font_controls_are_inert_and_do_not_leak_into_body() {
    let document = RtfDocument::parse(FONT_SCOPED_SEPARATOR).unwrap();
    let separator = document
        .note_separators()
        .get(NoteSeparatorKind::FootnoteSeparator)
        .unwrap();

    assert!(
        separator
            .elements
            .contains(&NoteSeparatorElement::SeparatorMark)
    );
    assert_eq!(separator.text(), "‚ ");
    assert_eq!(document.text(), "é Body");
}
