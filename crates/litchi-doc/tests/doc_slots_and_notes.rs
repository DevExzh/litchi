#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! M1 L-DOC-R8 (header and footer slots by section and kind) and L-DOC-R9
//! (notes and comments by index, one malformed comment failing only itself).

use litchi_core::Position;
use litchi_doc::raw::Reader;
use litchi_doc::writer::{Builder, CommentEntry};
use litchi_doc::{Document, Error, HeaderFooterType, Paragraph};
use litchi_ole_common::object::{
    Edit as ObjectEditor, Limits, OpenOptions as ObjectOpenOptions, Package as ObjectPackage,
    Targets,
};
use std::io::Cursor;
use std::path::PathBuf;

fn open_object(
    bytes: Vec<u8>,
    targets: Targets,
    limits: Limits,
) -> Result<ObjectEditor, litchi_ole_common::Error> {
    ObjectPackage::open_with_options(
        bytes,
        ObjectOpenOptions::new().targets(targets).limits(limits),
    )?
    .edit()
}

fn object_bytes(editor: &ObjectEditor) -> Vec<u8> {
    let package = editor.clone().commit().unwrap().into_package();
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap();
    bytes
}

fn fixture_directory() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc")
}

fn open(name: &str) -> Document {
    let mut package = Reader::open(fixture_directory().join(name)).expect("DOC fixture opens");
    package.document().expect("DOC fixture document")
}

fn open_bytes(bytes: Vec<u8>) -> Document {
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("DOC bytes open");
    package.document().expect("DOC document")
}

fn corpus() -> Vec<(String, Document)> {
    let mut names = std::fs::read_dir(fixture_directory())
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .filter(|name| name.ends_with(".doc"))
        .collect::<Vec<_>>();
    names.sort();
    let documents = names
        .into_iter()
        .filter_map(|name| {
            let mut package = Reader::open(fixture_directory().join(&name)).ok()?;
            let document = package.document().ok()?;
            Some((name, document))
        })
        .collect::<Vec<_>>();
    assert!(documents.len() > 25, "DOC corpus went missing");
    documents
}

fn paragraph_keys(paragraphs: &[Paragraph]) -> Vec<String> {
    paragraphs
        .iter()
        .map(|paragraph| {
            format!(
                "{:?} {:?} {:?}",
                paragraph.text().unwrap(),
                paragraph.cp_range(),
                paragraph.properties()
            )
        })
        .collect()
}

/// Slots cover every section, agree with `headers_footers()`, and resolve
/// empty slots to the nearest earlier section with a story of their kind.
#[test]
fn header_footer_slots_keep_sections_and_resolve_inheritance() {
    for (name, document) in corpus() {
        let sections = document.header_footer_section_count();
        let slots = document.header_footer_slots();
        assert_eq!(slots.len(), sections * 6, "{name}");

        // The non-empty slots, in storage order, are `headers_footers()`.
        let mut own = slots
            .iter()
            .filter(|slot| !slot.is_empty())
            .copied()
            .collect::<Vec<_>>();
        own.sort_by_key(|slot| slot.cp_range().unwrap().start());
        let legacy = document.headers_footers().unwrap();
        assert_eq!(own.len(), legacy.len(), "{name}");
        for (slot, expected) in own.iter().zip(&legacy) {
            assert_eq!(slot.kind(), expected.header_footer_type, "{name}");
            let built = document
                .section_header_footer(slot.section(), slot.kind())
                .unwrap()
                .unwrap();
            assert_eq!(built.text(), expected.text(), "{name}");
            assert_eq!(
                paragraph_keys(built.paragraphs().unwrap()),
                paragraph_keys(expected.paragraphs().unwrap()),
                "{name}"
            );
        }

        for slot in &slots {
            assert_eq!(
                document.header_footer_slot(slot.section(), slot.kind()),
                Some(*slot)
            );
            match slot.source_section() {
                Some(source) => {
                    assert!(source <= slot.section(), "{name}");
                    let source_slot = document.header_footer_slot(source, slot.kind()).unwrap();
                    assert!(!source_slot.is_empty(), "{name}");
                    assert_eq!(slot.is_inherited(), source != slot.section());
                    assert!((source.get() + 1..slot.section().get()).all(|between| {
                        document
                            .header_footer_slot(Position::new(between), slot.kind())
                            .unwrap()
                            .is_empty()
                    }));
                    let inherited = document
                        .section_header_footer(slot.section(), slot.kind())
                        .unwrap()
                        .unwrap();
                    let own_story = document
                        .section_header_footer(source, slot.kind())
                        .unwrap()
                        .unwrap();
                    assert_eq!(inherited.text(), own_story.text(), "{name}");
                },
                None => {
                    assert!(slot.is_empty() && !slot.is_inherited(), "{name}");
                    assert!(
                        document
                            .section_header_footer(slot.section(), slot.kind())
                            .unwrap()
                            .is_none()
                    );
                },
            }
        }
        assert!(
            document
                .header_footer_slot(Position::new(sections), HeaderFooterType::OddPageHeader)
                .is_none()
        );
        assert!(
            document
                .section_header_footer(Position::new(sections), HeaderFooterType::OddPageHeader)
                .unwrap()
                .is_none()
        );
    }
}

#[test]
fn later_sections_inherit_empty_slots() {
    let document = open("watermark.doc");
    assert_eq!(document.header_footer_section_count(), 3);
    let header = document
        .header_footer_slot(Position::new(2), HeaderFooterType::OddPageHeader)
        .unwrap();
    assert!(header.is_empty());
    assert!(header.is_inherited());
    assert_eq!(header.source_section(), Some(Position::new(0)));
    assert_eq!(
        document
            .section_header_footer(Position::new(2), HeaderFooterType::OddPageHeader)
            .unwrap()
            .unwrap()
            .text(),
        document
            .section_header_footer(Position::new(0), HeaderFooterType::OddPageHeader)
            .unwrap()
            .unwrap()
            .text()
    );
    // Section 0 has no even-page header of its own, and none earlier.
    let even = document
        .header_footer_slot(Position::new(0), HeaderFooterType::EvenPageHeader)
        .unwrap();
    assert_eq!(even.source_section(), None);

    let document = open("first-header-footer.doc");
    assert_eq!(document.header_footer_section_count(), 2);
    assert!(
        document
            .header_footer_slots()
            .iter()
            .all(|slot| !slot.is_empty())
    );
    let first = document
        .section_header_footer(Position::new(0), HeaderFooterType::FirstPageHeader)
        .unwrap()
        .unwrap();
    let second = document
        .section_header_footer(Position::new(1), HeaderFooterType::FirstPageHeader)
        .unwrap()
        .unwrap();
    assert_ne!(first.text(), second.text());
}

/// Notes by index are the values the lists return.
#[test]
fn notes_by_index_match_the_lists() {
    let mut notes = 0;
    for (name, document) in corpus() {
        let footnotes = document.footnotes().unwrap();
        assert_eq!(document.footnote_count(), footnotes.len(), "{name}");
        for (index, expected) in footnotes.iter().enumerate() {
            let note = document.footnote(Position::new(index)).unwrap().unwrap();
            assert_eq!(
                (note.reference_position, note.number, note.text()),
                (
                    expected.reference_position,
                    expected.number,
                    expected.text()
                )
            );
            assert_eq!(
                paragraph_keys(note.paragraphs().unwrap()),
                paragraph_keys(expected.paragraphs().unwrap())
            );
            notes += 1;
        }
        assert!(
            document
                .footnote(Position::new(footnotes.len()))
                .unwrap()
                .is_none()
        );

        let endnotes = document.endnotes().unwrap();
        assert_eq!(document.endnote_count(), endnotes.len(), "{name}");
        for (index, expected) in endnotes.iter().enumerate() {
            let note = document.endnote(Position::new(index)).unwrap().unwrap();
            assert_eq!(
                (note.reference_position, note.number, note.text()),
                (
                    expected.reference_position,
                    expected.number,
                    expected.text()
                )
            );
            notes += 1;
        }
        assert!(
            document
                .endnote(Position::new(endnotes.len()))
                .unwrap()
                .is_none()
        );
    }
    assert!(notes >= 8, "too few notes compared: {notes}");
}

/// Comments by index are the values the list returns, where it reads.
#[test]
fn comments_by_index_match_the_list() {
    let redlines = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/core/doc/data/bookmark-delete-redline.doc");
    let mut package = Reader::open(redlines).expect("LibreOffice fixture opens");
    let mut documents = corpus();
    documents.push((
        "bookmark-delete-redline.doc".to_string(),
        package.document().unwrap(),
    ));
    let mut compared = 0;
    for (name, document) in documents {
        let Ok(comments) = document.comments() else {
            continue;
        };
        assert_eq!(document.comment_count(), comments.len(), "{name}");
        for (index, expected) in comments.iter().enumerate() {
            let comment = document.comment(Position::new(index)).unwrap().unwrap();
            assert_eq!(comment.reference_position, expected.reference_position);
            assert_eq!(comment.author, expected.author);
            assert_eq!(comment.text(), expected.text());
            assert_eq!(
                paragraph_keys(comment.paragraphs().unwrap()),
                paragraph_keys(expected.paragraphs().unwrap())
            );
            compared += 1;
        }
        assert!(
            document
                .comment(Position::new(comments.len()))
                .unwrap()
                .is_none()
        );
    }
    assert_eq!(compared, 65, "comments compared");
}

/// A real file whose only comment is anchored where no annotation mark is:
/// the list fails, the comment's own read fails, and nothing else does.
#[test]
fn a_malformed_corpus_comment_fails_only_its_own_read() {
    let document = open("image-comment-at-char.doc");
    assert!(document.comments().is_err());
    assert_eq!(document.comment_count(), 1);
    let error = document.comment(Position::new(0)).unwrap_err();
    assert!(matches!(error, Error::Corrupted(_)), "{error}");
    assert!(!document.paragraphs().unwrap().is_empty());
    assert!(!document.body_blocks().unwrap().is_empty());
}

/// Three writer comments, the second with a damaged story mark: only that
/// one fails.
#[test]
fn one_bad_comment_leaves_the_others_readable() {
    let mut writer = Builder::new();
    writer.add_paragraph("Alpha beta gamma delta").unwrap();
    for (position, text) in [(1, "First note"), (7, "Second note"), (13, "Third note")] {
        writer.add_comment(CommentEntry::new(position, text, "Reviewer", "RV"));
    }
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let bytes = output.into_inner();

    let intact = open_bytes(bytes.clone());
    let texts = intact
        .comments()
        .unwrap()
        .iter()
        .map(|comment| comment.text().to_string())
        .collect::<Vec<_>>();
    assert_eq!(texts.len(), 3);
    assert!(texts[1].starts_with("Second note"), "{texts:?}");

    // Replace the second comment's story mark (U+0005 before its text) with
    // a letter, in the UTF-16LE text of the WordDocument stream.
    let mut editor = open_object(bytes, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let mut word_document = editor.stream(&word_path).unwrap().to_vec();
    let marked = "\u{5}Second note"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    let at = word_document
        .windows(marked.len())
        .position(|window| window == marked)
        .expect("the second comment's story");
    word_document[at..at + 2].copy_from_slice(&u16::from(b'X').to_le_bytes());
    editor.put_stream(&word_path, word_document).unwrap();
    let damaged = open_bytes(object_bytes(&editor));

    assert!(damaged.comments().is_err());
    assert_eq!(damaged.comment_count(), 3);
    assert_eq!(
        damaged.comment(Position::new(0)).unwrap().unwrap().text(),
        texts[0]
    );
    assert!(matches!(
        damaged.comment(Position::new(1)),
        Err(Error::Corrupted(_))
    ));
    assert_eq!(
        damaged.comment(Position::new(2)).unwrap().unwrap().text(),
        texts[2]
    );
}
