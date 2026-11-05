#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{LegacyNumberingAlignment, LegacyNumberingFormat, write::Writer};

const SYNTHETIC: &str = r"{\rtf1\ansi\ansicpg1250
{\*\pnseclvl1\pnucrm\pnqc\pnstart3\pnindent720\pnsp144\pnhang\pnprev\pnf2
{\pntxtb \'8a(}{\pntxta )\u20320?}}
{\*\pnseclvl2\pndec\pnstart1\pnindent360{\pntxta .}}
Body}";

#[test]
fn parses_decodes_and_round_trips_legacy_section_numbering() {
    let doc = Tree::parse(SYNTHETIC, litchi_rtf::read::OpenOptions::default()).unwrap();
    let numbering = doc.legacy_section_numbering();
    assert_eq!(numbering.levels().len(), 2);
    let first = numbering.get(1).unwrap();
    assert_eq!(first.format, LegacyNumberingFormat::UpperRoman);
    assert_eq!(first.alignment, Some(LegacyNumberingAlignment::Center));
    assert_eq!(first.start_at, Some(3));
    assert_eq!(first.indent, Some(720));
    assert_eq!(first.space, Some(144));
    assert!(first.hanging);
    assert!(first.previous);
    assert_eq!(first.font_ref, Some(2));
    assert_eq!(first.text_before, "Š(");
    assert_eq!(first.text_after, ")你");
    assert_eq!(doc.text().trim(), "Body");

    let mut first_bytes = Vec::new();
    Writer::new(&mut first_bytes).write_document(&doc).unwrap();
    let reparsed =
        Tree::parse_bytes(&first_bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(numbering, reparsed.legacy_section_numbering());
    let mut second_bytes = Vec::new();
    Writer::new(&mut second_bytes)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_bytes, second_bytes);
}

#[test]
fn rejects_malformed_legacy_section_numbering() {
    let malformed = [
        r"{\rtf1{\pnseclvl1\pndec}}",
        r"{\rtf1{\*\pnseclvl0\pndec}}",
        r"{\rtf1{\*\pnseclvl1\pndec}{\*\pnseclvl1\pndec}}",
        r"{\rtf1{\*\pnseclvl2\pndec}{\*\pnseclvl1\pndec}}",
        r"{\rtf1{\*\pnseclvl1\pnstart1}}",
        r"{\rtf1{\*\pnseclvl1\pndec{\pntxta X}{\pntxta Y}}}",
        r"{\rtf1{\*\pnseclvl1\pndec{\pntxta {\field X}}}}",
        r"{\rtf1 Body{\*\pnseclvl1\pndec}}",
    ];
    for source in malformed {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn accepts_only_unchanged_section_defaults_repeated_after_body_text() {
    let source = r"{\rtf1{\*\pnseclvl1\pndec\pnstart1{\pntxta .}}Body\par\sectd{\*\pnseclvl1\pndec\pnstart1{\pntxta .}}Tail}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(document.legacy_section_numbering().levels().len(), 1);
    assert_eq!(
        document.legacy_section_numbering().get(1).unwrap().start_at,
        Some(1)
    );
    for changed in [
        source.replace(
            "\\sectd{\\*\\pnseclvl1\\pndec\\pnstart1",
            "\\sectd{\\*\\pnseclvl1\\pndec\\pnstart2",
        ),
        source.replace("\\sectd{\\*\\pnseclvl1", "\\sectd{\\*\\pnseclvl2"),
    ] {
        assert!(Tree::parse(&changed, litchi_rtf::read::OpenOptions::default()).is_err());
    }
}

#[test]
fn parses_real_libreoffice_pnseclvl_fixture() {
    let doc = Tree::parse_bytes(
        include_bytes!(
            "../../../test-data/libreoffice-core/sw/qa/uibase/uiview/data/tdf152839_formtext.rtf"
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let numbering = doc.legacy_section_numbering();
    assert_eq!(numbering.levels().len(), 9);
    assert_eq!(
        numbering.get(1).unwrap().format,
        LegacyNumberingFormat::UpperRoman
    );
    assert_eq!(
        numbering.get(3).unwrap().format,
        LegacyNumberingFormat::Decimal
    );
    assert_eq!(numbering.get(5).unwrap().text_before, "(");
    assert_eq!(numbering.get(5).unwrap().text_after, ")");
}
