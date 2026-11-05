#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::{
    LegacyParagraphNumberingAlignment, LegacyParagraphNumberingFormat,
    LegacyParagraphNumberingLevel, LegacyParagraphNumberingUnderline, write::Writer,
};

const PRODUCER: &str = r"{\rtf1\ansi\ansicpg1252
{\pn\pnlvl3\pndec\pnqc\pnstart7\pnindent720\pnsp120\pnacross\pnnumonce\pnprev\pnrestart\pnhang\pnbidia\pnf2\pnfs24\pncf3\pnb\pni0\pncaps\pnscaps0\pnstrike\pnuldashd\pnrauth4\pnrdate5\pnrnfc6\pnrnot\pnrpnbr7\pnrrgb255\pnrstart8\pnrstop9\pnrxst10{\pntxtb (}{\pntxta )\u20320?}}Alpha\par
\pard{\pn\pnlvlblt\pngblip\pnqr{\pntxtb \'b7}}Beta\par}";

#[test]
fn parses_owns_and_canonically_round_trips_pn_metadata() {
    let document = Tree::parse(PRODUCER, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(document.legacy_paragraph_numbering_records().len(), 2);
    let alpha = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Alpha"))
        .unwrap();
    let first = document
        .legacy_paragraph_numbering(&alpha.paragraph)
        .unwrap();
    assert_eq!(first.level, LegacyParagraphNumberingLevel::Explicit(3));
    assert_eq!(first.format, Some(LegacyParagraphNumberingFormat::Decimal));
    assert_eq!(
        first.alignment,
        Some(LegacyParagraphNumberingAlignment::Center)
    );
    assert_eq!(
        first.underline,
        Some(LegacyParagraphNumberingUnderline::DashDot)
    );
    assert_eq!(first.text_before.as_deref(), Some("("));
    assert_eq!(first.text_after.as_deref(), Some(")你"));
    assert_eq!(first.bold, Some(true));
    assert_eq!(first.italic, Some(false));
    assert!(first.revision.no_tracking);
    let beta = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Beta"))
        .unwrap();
    let second = document
        .legacy_paragraph_numbering(&beta.paragraph)
        .unwrap();
    assert_eq!(second.level, LegacyParagraphNumberingLevel::Bullet);
    assert_eq!(second.format, Some(LegacyParagraphNumberingFormat::GbLip));
    let mut first_bytes = Vec::new();
    Writer::new(&mut first_bytes)
        .write_document(&document)
        .unwrap();
    let reparsed =
        Tree::parse_bytes(&first_bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut second_bytes = Vec::new();
    Writer::new(&mut second_bytes)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_bytes, second_bytes);
}

#[test]
fn pn_is_inert_and_paragraph_scoped() {
    let document = Tree::parse(
        r"{\rtf1{\pn\pnlvlbody\pndec{\pntxtb NOT-BODY}}One\par\pard Two}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "One\nTwo");
    let one = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("One"))
        .unwrap();
    let two = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Two"))
        .unwrap();
    assert!(
        document
            .legacy_paragraph_numbering(&one.paragraph)
            .is_some()
    );
    assert!(
        document
            .legacy_paragraph_numbering(&two.paragraph)
            .is_none()
    );
}

#[test]
fn rejects_adversarial_pn_destinations() {
    let malformed = [
        r"{\rtf1\pn\pnlvlbody\pndec X}",
        r"{\rtf1{\*\pn\pnlvlbody\pndec}X}",
        r"{\rtf1{\pn2\pnlvlbody\pndec}X}",
        r"{\rtf1{\pn\pndec}X}",
        r"{\rtf1{\pn\pnlvl0\pndec}X}",
        r"{\rtf1{\pn\pnlvlbody}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnucrm}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnstart}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnstart32768}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnfs0}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnb2}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec\pnql\pnqr}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec{\pntxta A}{\pntxta B}}X}",
        r"{\rtf1{\pn\pnlvlbody\pndec{\pntxta {\field X}}}X}",
        r"{\rtf1 X{\pn\pnlvlbody\pndec}}",
        r"{\rtf1{\pn\pnlvlbody\pndec}{\pn\pnlvlbody\pndec}X}",
        r"{\rtf1\trowd\intbl{\pn\pnlvlbody\pndec}X\cell\row}",
    ];
    for source in malformed {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                source,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed RTF: {source}"
        );
    }
}
