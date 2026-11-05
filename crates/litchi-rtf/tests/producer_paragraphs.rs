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
fn a_paragraph_that_inherits_properties_moves_with_its_state() {
    // The second paragraph has no `\pard`: it inherits `\qc` from the
    // first, so the move replays that state before it (record 0308).
    let source =
        Package::parse("{\\rtf1\\ansi\\pard\\qc One\\par Two\\par\\pard Three\\par}").unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(1), Position::new(2))
        .unwrap();
    let moved = edit.commit().unwrap().into_package();
    assert_eq!(texts(&moved), ["One", "Three", "Two"]);
    let two = moved.body().paragraphs().nth(2).unwrap();
    let one = source.body().paragraphs().nth(1).unwrap();
    assert_eq!(format!("{:?}", two.format()), format!("{:?}", one.format()));
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

/// Record 0308: a formatting group that holds several paragraphs, paragraphs
/// that inherit root-level state, insertion, removal of anchored content
/// and table rows.
mod record_0308 {
    use super::*;

    fn moved(source: &Package, from: usize, to: usize) -> Package {
        let mut edit = source.edit();
        edit.move_paragraph(Position::new(from), Position::new(to))
            .unwrap();
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn a_group_holding_several_paragraphs_is_split() {
        let rtf = "{\\rtf1\\ansi{\\fonttbl{\\f0 Times;}{\\f1 Arial;}}\\pard\\plain {\\f1\\b One\\par Two\\par}Three\\par}";
        let source = Package::parse(rtf).unwrap();
        assert_eq!(texts(&source), ["One", "Two", "Three"]);
        let changed = moved(&source, 0, 2);
        assert_eq!(texts(&changed), ["Two", "Three", "One"]);
        // Each paragraph keeps its runs' formatting.
        let runs = |package: &Package, text: &str| {
            package
                .body()
                .paragraphs()
                .find(|paragraph| paragraph.to_text() == text)
                .unwrap()
                .runs()
                .map(|run| format!("{:?}", run.format()))
                .collect::<Vec<_>>()
        };
        for text in ["One", "Two", "Three"] {
            assert_eq!(runs(&changed, text), runs(&source, text), "{text}");
        }
        let back = moved(&changed, 2, 0);
        assert_eq!(texts(&back), texts(&source));
    }

    #[test]
    fn inherited_root_state_is_replayed() {
        // The second paragraph has no `\pard`: it inherits `\qc` and `\b`.
        let rtf = "{\\rtf1\\ansi{\\fonttbl{\\f0 Times;}}\\pard\\plain\\qc\\b One\\par Two\\par\\pard\\plain Three\\par}";
        let source = Package::parse(rtf).unwrap();
        let changed = moved(&source, 1, 0);
        assert_eq!(texts(&changed), ["Two", "One", "Three"]);
        let format = |package: &Package, index: usize| {
            let paragraph = package.body().paragraphs().nth(index).unwrap();
            (
                format!("{:?}", paragraph.format()),
                paragraph
                    .runs()
                    .map(|run| format!("{:?}", run.format()))
                    .collect::<Vec<_>>(),
            )
        };
        assert_eq!(format(&changed, 0), format(&source, 1));
        assert_eq!(format(&changed, 1), format(&source, 0));
        assert_eq!(format(&changed, 2), format(&source, 2));
    }

    #[test]
    fn insertion_is_a_group_and_its_removal_is_exact() {
        let source = Package::parse(WORD).unwrap();
        let mut edit = source.edit();
        edit.insert_paragraph(Position::new(1), "New {text}")
            .unwrap();
        let inserted = edit.commit().unwrap().into_package();
        assert_eq!(
            texts(&inserted),
            [
                "First paragraph",
                "New {text}",
                "Second centred",
                "Third",
                "Fourth"
            ]
        );
        let mut edit = inserted.edit();
        edit.remove_paragraph(Position::new(1)).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(removed.to_bytes().unwrap(), WORD.as_bytes());
    }

    #[test]
    fn a_removed_paragraph_takes_its_field_but_not_half_a_bookmark() {
        let field = "{\\rtf1\\ansi\\pard\\plain One {\\field{\\*\\fldinst PAGE}{\\fldrslt 1}}\\par\\pard\\plain Two\\par}";
        let source = Package::parse(field).unwrap();
        assert_eq!(source.fields().len(), 1);
        let mut edit = source.edit();
        edit.remove_paragraph(Position::new(0)).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(texts(&removed), ["Two"]);
        assert!(removed.fields().is_empty());

        let bookmark = "{\\rtf1\\ansi\\pard\\plain {\\*\\bkmkstart b}One\\par\\pard\\plain Two{\\*\\bkmkend b}\\par\\pard\\plain Three\\par}";
        let source = Package::parse(bookmark).unwrap();
        let mut edit = source.edit();
        edit.remove_paragraph(Position::new(0)).unwrap();
        assert!(
            matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
    }

    #[test]
    fn table_rows_keep_their_place() {
        let rtf = "{\\rtf1\\ansi\\pard\\plain One\\par\\trowd\\cellx1000\\pard\\intbl A\\cell\\row\\pard\\plain Two\\par\\pard\\plain Three\\par}";
        let source = Package::parse(rtf).unwrap();
        assert_eq!(texts(&source), ["One", "Two", "Three"]);
        let changed = moved(&source, 2, 1);
        assert_eq!(texts(&changed), ["One", "Three", "Two"]);
        assert_eq!(changed.tables().len(), 1);
        let bytes = String::from_utf8(changed.to_bytes().unwrap()).unwrap();
        // The row stays after One and before Two; Three moved before it.
        let table = bytes.find("\\trowd").unwrap();
        assert!(bytes.find("Three").unwrap() < table && table < bytes.find("Two").unwrap());
    }
}
