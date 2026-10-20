#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! M1 L-DOC-R1, R2, R3: per-story reads, character-position ranges, and the
//! memoized story index.

use litchi_core::Position;
use litchi_doc::body_text::{Projection, Snapshot, Story};
use litchi_doc::raw::Reader;
use litchi_doc::writer::Builder;
use litchi_doc::{CpRange, Document, Element, Paragraph, StoryBlockKind, Table};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture_directory() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc")
}

fn open(name: &str) -> Document {
    let mut package = Reader::open(fixture_directory().join(name)).expect("DOC fixture opens");
    package.document().unwrap()
}

fn open_bytes(bytes: Vec<u8>) -> Document {
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("DOC bytes open");
    package.document().unwrap()
}

/// Every readable DOC fixture, by name, with its document.
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

/// A paragraph's text, run texts and properties, for comparisons.
fn paragraph_key(paragraph: &Paragraph) -> String {
    let runs = paragraph
        .runs()
        .unwrap()
        .iter()
        .map(|run| run.text().unwrap().to_string())
        .collect::<Vec<_>>();
    format!(
        "{:?} {runs:?} {:?}",
        paragraph.text().unwrap(),
        paragraph.properties()
    )
}

fn table_key(table: &Table) -> Vec<Vec<String>> {
    table
        .rows()
        .unwrap()
        .iter()
        .map(|row| {
            row.cells()
                .unwrap()
                .iter()
                .map(|cell| {
                    cell.paragraphs()
                        .unwrap()
                        .iter()
                        .map(paragraph_key)
                        .collect::<Vec<_>>()
                        .join("|")
                })
                .collect()
        })
        .collect()
}

fn element_key(element: &Element) -> String {
    match element {
        Element::Paragraph(paragraph) => format!("P {}", paragraph_key(paragraph)),
        Element::Table(table) => format!("T {:?}", table_key(table)),
    }
}

fn utf16_len(text: &str) -> u32 {
    u32::try_from(text.encode_utf16().count()).unwrap()
}

/// Per-story reads partition the all-story reads exactly: every story's
/// paragraphs, in story order, are `paragraphs()`, and every story's blocks
/// are `elements()`.
#[test]
fn story_reads_partition_the_all_story_reads() {
    for (name, document) in corpus() {
        let mut paragraphs = Vec::new();
        let mut elements = Vec::new();
        for story in Story::ALL {
            let story_paragraphs = document.story_paragraphs(story).unwrap();
            let blocks = document.story_blocks(story).unwrap();
            assert_eq!(blocks.story(), story);
            assert_eq!(blocks.paragraph_count(), story_paragraphs.len(), "{name}");
            assert_eq!(
                blocks.cp_range(),
                document.story_range(story).unwrap(),
                "{name}"
            );
            paragraphs.extend(story_paragraphs.iter().map(paragraph_key));
            elements.extend(blocks.elements().unwrap().iter().map(element_key));
        }
        assert_eq!(
            paragraphs,
            document
                .paragraphs()
                .unwrap()
                .iter()
                .map(paragraph_key)
                .collect::<Vec<_>>(),
            "{name}"
        );
        assert_eq!(paragraphs.len(), document.paragraph_count().unwrap());
        assert_eq!(
            elements,
            document
                .elements()
                .unwrap()
                .iter()
                .map(element_key)
                .collect::<Vec<_>>(),
            "{name}"
        );
    }
}

/// Block lookups by position, by kind and ordinal, and the body shortcuts
/// agree with the built elements.
#[test]
fn block_lookups_agree_with_the_built_elements() {
    for (name, document) in corpus() {
        for story in Story::ALL {
            let blocks = document.story_blocks(story).unwrap();
            let elements = blocks.elements().unwrap();
            assert_eq!(elements.len(), blocks.len(), "{name} {story:?}");
            let mut ordinals = [0usize; 2];
            for (position, element) in elements.iter().enumerate() {
                let block = blocks.get(Position::new(position)).unwrap();
                assert_eq!(block.position(), position);
                assert_eq!(
                    element_key(&blocks.element(Position::new(position)).unwrap().unwrap()),
                    element_key(element)
                );
                let slot = match (block.kind(), element) {
                    (StoryBlockKind::Paragraph, Element::Paragraph(paragraph)) => {
                        assert_eq!(
                            paragraph_key(
                                &blocks
                                    .paragraph(Position::new(block.ordinal()))
                                    .unwrap()
                                    .unwrap()
                            ),
                            paragraph_key(paragraph)
                        );
                        assert_eq!(paragraph.cp_range(), Some(block.cp_range()));
                        0
                    },
                    (StoryBlockKind::Table, Element::Table(table)) => {
                        assert_eq!(
                            table_key(
                                &blocks
                                    .table(Position::new(block.ordinal()))
                                    .unwrap()
                                    .unwrap()
                            ),
                            table_key(table)
                        );
                        assert_eq!(table.cp_range(), Some(block.cp_range()));
                        1
                    },
                    other => panic!("{name}: block kind and element differ: {other:?}"),
                };
                assert_eq!(block.ordinal(), ordinals[slot], "{name}");
                ordinals[slot] += 1;
            }
            assert_eq!(ordinals[0], blocks.count(StoryBlockKind::Paragraph));
            assert_eq!(ordinals[1], blocks.count(StoryBlockKind::Table));
            assert!(blocks.get(Position::new(blocks.len())).is_none());
            assert!(
                blocks
                    .paragraph(Position::new(blocks.count(StoryBlockKind::Paragraph)))
                    .unwrap()
                    .is_none()
            );
            assert!(
                blocks
                    .table(Position::new(blocks.count(StoryBlockKind::Table)))
                    .unwrap()
                    .is_none()
            );
            if story == Story::Main {
                assert_eq!(document.body_block_count().unwrap(), blocks.len());
                assert_eq!(
                    document
                        .body_table(Position::new(0))
                        .unwrap()
                        .map(|table| table_key(&table)),
                    blocks
                        .table(Position::new(0))
                        .unwrap()
                        .map(|table| table_key(&table))
                );
            }
        }
    }
}

/// `/body/p[N]` is the main-story text editor's paragraph space: the same
/// paragraphs, in the same order, with the same text.
#[test]
fn body_paragraphs_are_the_body_text_editor_paragraphs() {
    let mut compared = 0;
    for (name, document) in corpus() {
        let bytes = std::fs::read(fixture_directory().join(&name)).unwrap();
        let Ok(snapshot) = Snapshot::from_bytes(bytes) else {
            continue;
        };
        let Ok(editor_paragraphs) = snapshot.paragraphs(Projection::All) else {
            continue;
        };
        let body = document.body_blocks().unwrap();
        assert_eq!(
            body.count(StoryBlockKind::Paragraph),
            editor_paragraphs.len(),
            "{name}"
        );
        for (ordinal, expected) in editor_paragraphs.iter().enumerate() {
            let paragraph = document
                .body_paragraph(Position::new(ordinal))
                .unwrap()
                .unwrap();
            assert_eq!(paragraph.text().unwrap(), expected.text(), "{name}");
        }
        compared += 1;
    }
    assert!(compared > 20, "too few fixtures compared: {compared}");
}

/// Paragraphs tile their story, runs sit inside their paragraph's text, and
/// tables, rows and cells cover their own paragraphs.
#[test]
fn cp_ranges_locate_every_value_in_its_story() {
    for (name, document) in corpus() {
        for story in Story::ALL {
            let Some(story_range) = document.story_range(story).unwrap() else {
                assert_eq!(document.story_blocks(story).unwrap().len(), 0);
                continue;
            };
            let mut expected_start = story_range.start();
            for paragraph in document.story_paragraphs(story).unwrap() {
                let range = paragraph.cp_range().expect("story paragraphs have ranges");
                let content = paragraph.content_cp_range().unwrap();
                assert_eq!(range.start(), expected_start, "{name} {story:?}");
                assert!(story_range.encloses(range), "{name}");
                assert!(range.encloses(content), "{name}");
                assert!(range.len() - content.len() <= 1, "{name}");
                assert_eq!(
                    utf16_len(paragraph.text().unwrap()),
                    content.len(),
                    "{name}"
                );
                let mut run_start = content.start();
                for run in paragraph.runs().unwrap() {
                    let run_range = run.cp_range().expect("story runs have ranges");
                    assert!(run_range.start() >= run_start, "{name}");
                    assert!(content.encloses(run_range), "{name}");
                    assert_eq!(utf16_len(run.text().unwrap()), run_range.len(), "{name}");
                    run_start = run_range.end();
                }
                expected_start = range.end();
            }

            for table in document.story_blocks(story).unwrap().tables().unwrap() {
                let table_range = table.cp_range().expect("story tables have ranges");
                let rows = table.rows().unwrap();
                assert_eq!(
                    rows.first().unwrap().cp_range().unwrap().start(),
                    table_range.start()
                );
                assert_eq!(
                    rows.last().unwrap().cp_range().unwrap().end(),
                    table_range.end()
                );
                let mut row_start = table_range.start();
                for row in rows {
                    let row_range = row.cp_range().expect("story rows have ranges");
                    assert_eq!(row_range.start(), row_start, "{name}");
                    let mut cell_start = row_start;
                    for cell in row.cells().unwrap() {
                        let Some(cell_range) = cell.cp_range() else {
                            // Only the empty cell of a row-end-only row.
                            assert_eq!(cell.text().unwrap(), "");
                            continue;
                        };
                        assert_eq!(cell_range.start(), cell_start, "{name}");
                        assert!(row_range.encloses(cell_range), "{name}");
                        for paragraph in cell.paragraphs().unwrap() {
                            assert!(cell_range.encloses(paragraph.cp_range().unwrap()));
                        }
                        cell_start = cell_range.end();
                    }
                    row_start = row_range.end();
                }
            }
        }
    }
}

#[test]
fn word97_tables_are_blocks_of_their_own_story() {
    let document = open("HeaderFooterProblematic.doc");
    let body = document.body_blocks().unwrap();
    let kinds = body.iter().map(|block| block.kind()).collect::<Vec<_>>();
    assert_eq!(
        kinds,
        [
            StoryBlockKind::Paragraph,
            StoryBlockKind::Table,
            StoryBlockKind::Paragraph,
        ]
    );
    let table_block = body.get(Position::new(1)).unwrap();
    assert_eq!(table_block.paragraph_range(), 1..352);
    assert_eq!(body.paragraph_count(), 353);
    assert_eq!(
        body.table(Position::new(0))
            .unwrap()
            .unwrap()
            .row_count()
            .unwrap(),
        116
    );
    // The fourth paragraph of `paragraphs()` is the header story's.
    let headers = document.story_blocks(Story::Header).unwrap();
    assert_eq!(headers.len(), 1);
    assert_eq!(headers.paragraph_count(), 1);

    let document = open("tdf71749_with_footnote.doc");
    assert_eq!(
        document.body_blocks().unwrap().count(StoryBlockKind::Table),
        0
    );
    let footnotes = document.story_blocks(Story::Footnote).unwrap();
    assert_eq!(footnotes.count(StoryBlockKind::Table), 1);
    let table = footnotes.table(Position::new(0)).unwrap().unwrap();
    assert_eq!(table.row_count().unwrap(), 2);
    // The footnote story starts after the main story's CPs.
    let footnote_range = document.story_range(Story::Footnote).unwrap().unwrap();
    assert_eq!(
        footnote_range.start(),
        document.story_range(Story::Main).unwrap().unwrap().end()
    );
    assert!(footnote_range.encloses(table.cp_range().unwrap()));
}

/// A writer document whose CP layout is known: exact ranges for body
/// paragraphs, the table, its rows and cells, and a header story.
#[test]
fn writer_document_ranges_are_exact() {
    let mut writer = Builder::new();
    writer.add_paragraph("Before").unwrap();
    let table = writer.add_table(2, 2).unwrap();
    for row in 0..2 {
        for column in 0..2 {
            writer
                .set_table_cell_text(table, row, column, &format!("r{row}c{column}"))
                .unwrap();
        }
    }
    writer.add_paragraph("After").unwrap();
    writer.set_odd_header("Header");
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let document = open_bytes(output.into_inner());

    // The writer lays the main story out as
    // "Before\rAfter\r" + "r0c0\x07r0c1\x07\x07r1c0\x07r1c1\x07\x07" + "\r".
    let main = document.story_range(Story::Main).unwrap().unwrap();
    assert_eq!(main, CpRange::new(0, 36).unwrap());
    let body = document.body_blocks().unwrap();
    let kinds = body.iter().map(|block| block.kind()).collect::<Vec<_>>();
    assert_eq!(
        kinds,
        [
            StoryBlockKind::Paragraph,
            StoryBlockKind::Paragraph,
            StoryBlockKind::Table,
            StoryBlockKind::Paragraph,
        ]
    );

    let before = body.paragraph(Position::new(0)).unwrap().unwrap();
    assert_eq!(before.text().unwrap(), "Before");
    assert_eq!(before.cp_range(), CpRange::new(0, 7));
    assert_eq!(before.content_cp_range(), CpRange::new(0, 6));
    assert_eq!(before.runs().unwrap()[0].cp_range(), CpRange::new(0, 6));
    let after = body.paragraph(Position::new(1)).unwrap().unwrap();
    assert_eq!(after.text().unwrap(), "After");
    assert_eq!(after.cp_range(), CpRange::new(7, 13));
    assert_eq!(after.content_cp_range(), CpRange::new(7, 12));

    let table = body.table(Position::new(0)).unwrap().unwrap();
    assert_eq!(table.cp_range(), CpRange::new(13, 35));
    assert_eq!(body.get(Position::new(2)).unwrap().paragraph_range(), 2..8);
    let rows = table.rows().unwrap();
    // Each row: two cells of "rXcY" and a cell mark, then its row-end mark.
    assert_eq!(rows[0].cp_range(), CpRange::new(13, 24));
    assert_eq!(rows[1].cp_range(), CpRange::new(24, 35));
    let cells = rows[1].cells().unwrap();
    assert_eq!(cells[0].cp_range(), CpRange::new(24, 29));
    assert_eq!(cells[1].cp_range(), CpRange::new(29, 34));
    let cell_paragraph = &cells[1].paragraphs().unwrap()[0];
    assert_eq!(cell_paragraph.text().unwrap(), "r1c1");
    assert_eq!(cell_paragraph.cp_range(), CpRange::new(29, 34));
    assert_eq!(cell_paragraph.content_cp_range(), CpRange::new(29, 33));

    // The empty paragraph the writer ends the story with.
    let last = body.paragraph(Position::new(2)).unwrap().unwrap();
    assert_eq!(last.text().unwrap(), "");
    assert_eq!(last.cp_range(), CpRange::new(35, 36));

    let header = document.story_range(Story::Header).unwrap().unwrap();
    assert_eq!(header.start(), main.end());
    let header_paragraphs = document.story_paragraphs(Story::Header).unwrap();
    assert_eq!(header_paragraphs[0].text().unwrap(), "Header");
    assert_eq!(header_paragraphs[0].cp_range(), CpRange::new(36, 43));
    assert!(
        header_paragraphs
            .iter()
            .all(|paragraph| { header.encloses(paragraph.cp_range().unwrap()) })
    );
    assert!(document.story_range(Story::Comment).unwrap().is_none());
    assert!(document.story_blocks(Story::Comment).unwrap().is_empty());
}

/// The memo: one scan per story, views pin what they hold, trimming releases
/// the rest, and a released story reads the same again.
#[test]
fn story_indexes_are_memoized_pinned_and_trimmed() {
    let document = open("table-merged-cells.doc");
    let initial = document.story_index_diagnostics();
    assert_eq!(initial.retained_entries, 0);
    assert!(initial.max_bytes > 0);

    let body = document.body_blocks().unwrap();
    let first = document.story_index_diagnostics();
    assert_eq!(first.misses, initial.misses + 1);
    assert_eq!(first.builds, initial.builds + 1);
    assert_eq!(first.retained_entries, 1);
    assert_eq!(first.pinned_entries, 1);

    let table = document.body_table(Position::new(0)).unwrap().unwrap();
    let second = document.story_index_diagnostics();
    assert_eq!(second.hits, first.hits + 1);
    assert_eq!(second.builds, first.builds);

    // The view holds the index, so trimming keeps it.
    assert_eq!(document.trim_story_indexes(), 0);
    assert_eq!(document.story_index_diagnostics().retained_entries, 1);
    let kept = table_key(&body.table(Position::new(0)).unwrap().unwrap());
    drop(body);
    assert_eq!(document.trim_story_indexes(), 1);
    let trimmed = document.story_index_diagnostics();
    assert_eq!(trimmed.retained_entries, 0);
    assert_eq!(trimmed.retained_bytes, 0);
    assert_eq!(trimmed.trimmed, 1);

    // A released story is scanned again and reads the same.
    let again = document.body_table(Position::new(0)).unwrap().unwrap();
    assert_eq!(table_key(&again), kept);
    assert_eq!(table_key(&again), table_key(&table));
    let rebuilt = document.story_index_diagnostics();
    assert_eq!(rebuilt.builds, trimmed.builds + 1);
    assert_eq!(rebuilt.pinned_entries, 0);
}

/// One document shared across threads reads the same blocks everywhere.
#[test]
fn story_reads_are_shared_across_threads() {
    let document = open("HeaderFooterProblematic.doc");
    let expected = document
        .body_blocks()
        .unwrap()
        .elements()
        .unwrap()
        .iter()
        .map(element_key)
        .collect::<Vec<_>>();
    document.trim_story_indexes();
    std::thread::scope(|scope| {
        let handles = (0..4)
            .map(|_| {
                scope.spawn(|| {
                    document
                        .body_blocks()
                        .unwrap()
                        .elements()
                        .unwrap()
                        .iter()
                        .map(element_key)
                        .collect::<Vec<_>>()
                })
            })
            .collect::<Vec<_>>();
        for handle in handles {
            assert_eq!(handle.join().unwrap(), expected);
        }
    });
    let diagnostics = document.story_index_diagnostics();
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(
        diagnostics.builds, diagnostics.misses,
        "every miss built once"
    );
}

/// Over the corpus: the story index's borrowed paragraph text, where it
/// vouches for a paragraph, is the body traversal's text at the same
/// position, so the constant-time lookup and the scan address one paragraph.
#[test]
fn indexed_body_paragraph_text_matches_the_body_scan_over_the_corpus() {
    let mut compared = 0usize;
    for (name, document) in corpus() {
        let Ok(blocks) = document.body_blocks() else {
            continue;
        };
        let scanned = document.body_paragraphs().collect::<Result<Vec<_>, _>>();
        let Ok(scanned) = scanned else { continue };
        for (position, text) in scanned.iter().enumerate() {
            if let Some(indexed) = blocks.paragraph_text(Position::new(position)) {
                assert_eq!(indexed, *text, "{name} paragraph {position}");
                compared += 1;
            }
        }
        // The index never vouches for a paragraph the scan does not emit.
        assert!(
            blocks
                .paragraph_text(Position::new(scanned.len()))
                .is_none(),
            "{name}"
        );
    }
    assert!(compared > 400, "too few paragraphs compared: {compared}");
}
