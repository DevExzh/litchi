//! Paragraph removal, moves and swaps on producer RTF (record 0293).
//!
//! Producer bodies (`\pard` paragraph properties, character-formatting
//! groups, fonts) are spliced as whole source segments; the candidate must
//! read back as the expected paragraphs, and a move followed by its inverse
//! restores the source bytes.

use litchi_core::Position;
use litchi_rtf::Package;
use litchi_rtf::edit::Error;

/// A Word-style body: a header, `\pard` per paragraph, formatting groups and
/// a paragraph whose `\par` closes its formatting group.
const WORD: &str = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0\\froman Times;}{\\f1\\fswiss Arial;}}\r\n\
\\pard\\plain \\ql \\f0\\fs24 {\\b First} paragraph\\par\r\n\
\\pard\\plain \\qc \\f1\\fs20 Second {\\i centred}\\par\r\n\
\\pard\\plain \\qr \\f0\\fs24 {\\ul Third\\par }\r\n\
\\pard\\plain \\qj \\f0\\fs24 Fourth\\par\r\n}";

fn texts(package: &Package) -> Vec<String> {
    package
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect()
}

#[test]
fn producer_paragraphs_move_and_move_back_byte_exactly() {
    let source = Package::parse(WORD).unwrap();
    assert_eq!(texts(&source).len(), 4);
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let moved = edit.commit().unwrap().into_package();
    assert_eq!(
        texts(&moved),
        ["Second centred", "Third", "First paragraph", "Fourth"]
    );
    let mut edit = moved.edit();
    edit.move_paragraph(Position::new(2), Position::new(0))
        .unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(back.to_bytes().unwrap(), WORD.as_bytes());
}

#[test]
fn producer_paragraphs_swap_and_swap_back_byte_exactly() {
    let source = Package::parse(WORD).unwrap();
    let mut edit = source.edit();
    edit.swap_paragraphs(Position::new(1), Position::new(3))
        .unwrap();
    let swapped = edit.commit().unwrap().into_package();
    assert_eq!(
        texts(&swapped),
        ["First paragraph", "Fourth", "Third", "Second centred"]
    );
    let mut edit = swapped.edit();
    edit.swap_paragraphs(Position::new(1), Position::new(3))
        .unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(back.to_bytes().unwrap(), WORD.as_bytes());
}

#[test]
fn producer_paragraph_removal_keeps_every_other_byte() {
    let source = Package::parse(WORD).unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(2)).unwrap();
    let removed = edit.commit().unwrap().into_package();
    assert_eq!(
        texts(&removed),
        ["First paragraph", "Second centred", "Fourth"]
    );
    let expected = WORD.replace("\\pard\\plain \\qr \\f0\\fs24 {\\ul Third\\par }\r\n", "");
    assert_eq!(removed.to_bytes().unwrap(), expected.as_bytes());
}

#[test]
fn a_paragraph_that_inherits_properties_cannot_move() {
    // The second paragraph has no `\pard`: it inherits `\qc` from the
    // first, so moving it to the front would change its alignment.
    let source =
        Package::parse("{\\rtf1\\ansi\\pard\\qc One\\par Two\\par\\pard Three\\par}").unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(1), Position::new(2))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
}

#[test]
fn a_paragraph_holding_a_field_is_not_removed() {
    let source = Package::parse(
        "{\\rtf1\\ansi\\pard One {\\field{\\*\\fldinst PAGE}{\\fldrslt 1}}\\par\\pard Two\\par}",
    )
    .unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
}

#[test]
fn an_empty_first_paragraph_moves_and_moves_back() {
    let input = "{\\rtf1\n\\fs50 one \\par \\par two \\par\n}";
    let source = Package::parse(input).unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    let moved = edit.commit().unwrap().into_package();
    assert_eq!(texts(&moved)[..2], ["", "one "]);
    let mut edit = moved.edit();
    edit.move_paragraph(Position::new(1), Position::new(0))
        .unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(back.to_bytes().unwrap(), input.as_bytes());
}
