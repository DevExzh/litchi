#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! L-DOC-R7: table paragraphs that store no table depth form top-level tables.
//!
//! Word 97 predates nested tables and marks every table paragraph with
//! `sprmPFInTable` alone, without `sprmPItap`. [MS-DOC] §2.4.3 derives such a
//! paragraph's table depth as 1; the reader used to require a stored depth of
//! 1 and so dropped these tables, and their paragraphs, from `elements()`.

use litchi_doc::raw::Reader;
use litchi_doc::writer::Builder;
use litchi_doc::{Document, Element, Table};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name)
}

fn open(name: &str) -> Document {
    let mut package = Reader::open(fixture(name)).expect("DOC fixture opens");
    package.document().unwrap()
}

fn open_bytes(bytes: Vec<u8>) -> Document {
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("DOC bytes open");
    package.document().unwrap()
}

/// Every row's cell texts.
fn cell_texts(table: &Table) -> Vec<Vec<String>> {
    table
        .rows()
        .unwrap()
        .iter()
        .map(|row| {
            row.cells()
                .unwrap()
                .iter()
                .map(|cell| cell.text().unwrap().to_string())
                .collect()
        })
        .collect()
}

/// The number of paragraphs held by a table's cells.
fn cell_paragraph_count(table: &Table) -> usize {
    table
        .rows()
        .unwrap()
        .iter()
        .flat_map(|row| row.cells().unwrap())
        .map(|cell| cell.paragraphs().unwrap().len())
        .sum()
}

/// `P` for a paragraph element and `T` for a table element, in order.
fn element_kinds(document: &Document) -> String {
    document
        .elements()
        .unwrap()
        .iter()
        .map(|element| match element {
            Element::Paragraph(_) => 'P',
            Element::Table(_) => 'T',
        })
        .collect()
}

#[test]
fn word97_body_table_of_header_footer_problematic_is_read() {
    let document = open("HeaderFooterProblematic.doc");
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(paragraphs.len(), 354);
    let in_table = paragraphs
        .iter()
        .filter(|paragraph| paragraph.properties().in_table)
        .collect::<Vec<_>>();
    assert_eq!(in_table.len(), 351);
    // The file stores no `sprmPItap`: every table paragraph keeps a stored
    // depth of 0 and derives depth 1.
    assert!(in_table.iter().all(|paragraph| {
        paragraph.properties().table_nesting_level == 0 && paragraph.properties().table_depth() == 1
    }));

    // Before L-DOC-R7: no table, and `elements()` kept 3 of 354 paragraphs.
    assert_eq!(document.table_count().unwrap(), 1);
    let tables = document.tables().unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.row_count().unwrap(), 116);
    // Every table paragraph lands in the table: 235 cell paragraphs plus
    // 116 row-end marks.
    assert_eq!(cell_paragraph_count(table), 235);
    let rows = cell_texts(table);
    assert!(rows[0][0].contains("STATEMENT OF INSOLVENCY PRACTICE 1"));
    assert_eq!(rows[9][0], "1.");
    assert!(rows[9][1].starts_with("This statement of insolvency practice is"));
    assert!(rows.iter().all(|row| (1..=3).contains(&row.len())));

    assert_eq!(element_kinds(&document), "PTPP");
    let Element::Table(element_table) = &document.elements().unwrap()[1] else {
        panic!("the second element is the table");
    };
    assert_eq!(cell_texts(element_table), rows);
}

#[test]
fn word97_footnote_table_of_tdf71749_is_read() {
    let document = open("tdf71749_with_footnote.doc");
    let tables = document.tables().unwrap();
    assert_eq!(document.table_count().unwrap(), 1);
    assert_eq!(tables.len(), 1);
    assert_eq!(
        cell_texts(&tables[0]),
        [vec!["numcom", "L1", "L2", "Total"], vec!["", "", "", ""],]
    );
    // 8 cell paragraphs and 2 row-end marks: the file's 10 table paragraphs.
    assert_eq!(cell_paragraph_count(&tables[0]), 8);

    // The table sits in the second footnote's story, after its first
    // paragraph; `elements()` used to drop its 10 paragraphs.
    let kinds = element_kinds(&document);
    assert_eq!(kinds.len(), 28);
    assert_eq!(kinds.matches('T').count(), 1);
    assert_eq!(kinds.find('T'), Some(25));
    let elements = document.elements().unwrap();
    let Element::Paragraph(before) = &elements[24] else {
        panic!("a paragraph precedes the footnote table");
    };
    assert_eq!(before.text().unwrap(), "\u{2} Production  details");
}

/// Replace every `sprmPItap` the writer stores after `sprmPFInTable` with a
/// `sprmPRsid` of the same size, which leaves the Word 97 encoding:
/// `sprmPFInTable` and no stored depth. Returns the number of replacements.
fn strip_stored_table_depth(bytes: &mut [u8]) -> usize {
    // sprmPFInTable(1), then sprmPItap(1): 0x2416 01, 0x6649 01000000.
    const STORED: [u8; 9] = [0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00];
    // sprmPFInTable(1), then sprmPRsid(1): 0x2416 01, 0x6467 01000000.
    const WORD97: [u8; 9] = [0x16, 0x24, 0x01, 0x67, 0x64, 0x01, 0x00, 0x00, 0x00];
    let mut replaced = 0;
    let mut offset = 0;
    while let Some(found) = bytes[offset..]
        .windows(STORED.len())
        .position(|window| window == STORED)
    {
        let start = offset + found;
        bytes[start..start + STORED.len()].copy_from_slice(&WORD97);
        replaced += 1;
        offset = start + STORED.len();
    }
    replaced
}

#[test]
fn a_table_without_stored_depths_reads_like_the_same_table_with_them() {
    let mut writer = Builder::new();
    writer.add_paragraph("Before").unwrap();
    let table = writer.add_table(2, 3).unwrap();
    for row in 0..2 {
        for column in 0..3 {
            writer
                .set_table_cell_text(table, row, column, &format!("r{row}c{column}"))
                .unwrap();
        }
    }
    writer.add_paragraph("After").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let stored = output.into_inner();
    let mut word97 = stored.clone();
    assert!(strip_stored_table_depth(&mut word97) > 0);

    let with_depth = open_bytes(stored);
    let without_depth = open_bytes(word97);
    let depths = |document: &Document| {
        document
            .paragraphs()
            .unwrap()
            .iter()
            .filter(|paragraph| paragraph.properties().in_table)
            .map(|paragraph| {
                (
                    paragraph.properties().table_nesting_level,
                    paragraph.properties().table_depth(),
                )
            })
            .collect::<Vec<_>>()
    };
    // 6 cells and 2 row-end marks, stored at depth 1 or not at all.
    assert_eq!(depths(&with_depth), vec![(1, 1); 8]);
    assert_eq!(depths(&without_depth), vec![(0, 1); 8]);

    for document in [&with_depth, &without_depth] {
        assert_eq!(document.table_count().unwrap(), 1);
        let tables = document.tables().unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(
            cell_texts(&tables[0]),
            [vec!["r0c0", "r0c1", "r0c2"], vec!["r1c0", "r1c1", "r1c2"]]
        );
    }
    assert_eq!(element_kinds(&with_depth), element_kinds(&without_depth));
    let row_properties = |document: &Document| {
        document.tables().unwrap()[0]
            .rows()
            .unwrap()
            .iter()
            .map(|row| format!("{:?}", row.properties()))
            .collect::<Vec<_>>()
    };
    assert_eq!(row_properties(&with_depth), row_properties(&without_depth));
}

/// Over every DOC fixture that opens: `table_count()` is `tables().len()`,
/// and `elements()` keeps every paragraph outside a table plus one element
/// per top-level table, in the order `tables()` returns them.
#[test]
fn corpus_tables_and_elements_agree() {
    let directory = fixture("");
    let mut names = std::fs::read_dir(&directory)
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .filter(|name| name.ends_with(".doc"))
        .collect::<Vec<_>>();
    names.sort();
    let mut checked = 0;
    for name in names {
        let Ok(mut package) = Reader::open(directory.join(&name)) else {
            continue;
        };
        let Ok(document) = package.document() else {
            continue;
        };
        let tables = document.tables().unwrap();
        assert_eq!(document.table_count().unwrap(), tables.len(), "{name}");
        let elements = document.elements().unwrap();
        let outside_tables = document
            .paragraphs()
            .unwrap()
            .iter()
            .filter(|paragraph| paragraph.properties().table_depth() == 0)
            .count();
        let element_tables = elements
            .iter()
            .filter_map(|element| match element {
                Element::Table(table) => Some(cell_texts(table)),
                Element::Paragraph(_) => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            elements.len() - element_tables.len(),
            outside_tables,
            "{name}"
        );
        assert_eq!(
            element_tables,
            tables.iter().map(cell_texts).collect::<Vec<_>>(),
            "{name}"
        );
        checked += 1;
    }
    assert!(checked > 25, "too few DOC fixtures opened: {checked}");
}
