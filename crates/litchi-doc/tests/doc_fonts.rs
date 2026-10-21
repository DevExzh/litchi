#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! M1 L-DOC-R4: the font table (`SttbfFfn`) and `font_name(ftc)`.

use litchi_core::Position;
use litchi_doc::raw::Reader;
use litchi_doc::writer::{Builder, CharacterFormatting, ParagraphFormatting};
use litchi_doc::{Document, Error, FontFamily};
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

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document")
        .join(name)
}

fn open_poi(name: &str) -> Document {
    let mut package = Reader::open(poi_fixture(name)).expect("DOC fixture opens");
    package.document().expect("DOC fixture document")
}

fn open_bytes(bytes: Vec<u8>) -> Document {
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("DOC bytes open");
    package.document().expect("DOC document")
}

/// Over every readable fixture: the font table parses, and every font index
/// a run carries names a font.
#[test]
fn corpus_font_tables_name_every_run_font() {
    let mut names = std::fs::read_dir(fixture_directory())
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .filter(|name| name.ends_with(".doc"))
        .collect::<Vec<_>>();
    names.sort();
    let mut tables = 0;
    for name in names {
        let Ok(mut package) = Reader::open(fixture_directory().join(&name)) else {
            continue;
        };
        let Ok(document) = package.document() else {
            continue;
        };
        let fonts = document
            .font_table()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
            .unwrap_or_else(|| panic!("{name}: Word writes a font table"));
        // Word reserves the first three slots.
        assert_eq!(fonts.name(0), Some("Times New Roman"), "{name}");
        assert_eq!(fonts.name(1), Some("Symbol"), "{name}");
        assert_eq!(fonts.name(2), Some("Arial"), "{name}");
        for paragraph in document.paragraphs().unwrap() {
            for run in paragraph.runs().unwrap() {
                if let Some(ftc) = run.font_index() {
                    assert!(
                        document.font_name(ftc).unwrap().is_some(),
                        "{name}: run font {ftc} is not in the font table"
                    );
                }
            }
        }
        tables += 1;
    }
    assert!(tables > 25, "too few font tables read: {tables}");
}

#[test]
fn run_fonts_resolve_to_their_names() {
    let document = open_poi("test.doc");
    let mut used = std::collections::BTreeSet::new();
    for paragraph in document.paragraphs().unwrap() {
        for run in paragraph.runs().unwrap() {
            if let Some(ftc) = run.font_index() {
                used.insert(document.font_name(ftc).unwrap().unwrap().to_string());
            }
        }
    }
    assert!(used.contains("Comic Sans MS"), "{used:?}");

    let fonts = open_poi("saved-by-table.doc");
    let fonts = fonts.font_table().unwrap().unwrap();
    let arial = fonts.iter().find(|font| font.name() == "Arial").unwrap();
    assert_eq!(arial.alternate_name(), Some("Helvetica"));
    assert_eq!(arial.family(), FontFamily::Swiss);
    assert!(arial.is_true_type());
}

#[test]
fn writer_run_fonts_round_trip_through_the_font_table() {
    let mut writer = Builder::new();
    writer
        .add_paragraph_runs(
            vec![
                (
                    "mono".to_string(),
                    CharacterFormatting {
                        font_name: Some("Courier New".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " sans".to_string(),
                    CharacterFormatting {
                        font_name: Some("Arial".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let document = open_bytes(output.into_inner());

    let paragraph = document.body_paragraph(Position::new(0)).unwrap().unwrap();
    let names = paragraph
        .runs()
        .unwrap()
        .iter()
        .map(|run| {
            let ftc = run.font_index().expect("the writer sets each run's font");
            (
                run.text().unwrap().to_string(),
                document.font_name(ftc).unwrap().unwrap().to_string(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        names,
        [
            ("mono".to_string(), "Courier New".to_string()),
            (" sans".to_string(), "Arial".to_string()),
        ]
    );
    assert_eq!(document.font_name(u16::MAX).unwrap(), None);
}

/// The `FibRgFcLcb97` offset of entry `index` in a Word 97+ FIB.
fn fc_lcb_offset(word_document: &[u8], index: usize) -> usize {
    let csw = usize::from(u16::from_le_bytes([word_document[32], word_document[33]]));
    let cslw_at = 34 + csw * 2;
    let cslw = usize::from(u16::from_le_bytes([
        word_document[cslw_at],
        word_document[cslw_at + 1],
    ]));
    let fc_lcb = cslw_at + 2 + cslw * 4 + 2;
    fc_lcb + index * 8
}

/// A malformed font table is reported where it is read and nowhere else.
#[test]
fn a_malformed_font_table_is_a_typed_error_that_spares_the_rest() {
    let bytes = std::fs::read(poi_fixture("test.doc")).unwrap();
    let mut editor = open_object(bytes, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let mut word_document = editor.stream(&word_path).unwrap().to_vec();
    // Shorten lcbSttbfFfn (entry 15) by one byte: the last font record is
    // then truncated.
    let lcb_at = fc_lcb_offset(&word_document, 15) + 4;
    let lcb = u32::from_le_bytes(word_document[lcb_at..lcb_at + 4].try_into().unwrap());
    word_document[lcb_at..lcb_at + 4].copy_from_slice(&(lcb - 1).to_le_bytes());
    editor.put_stream(&word_path, word_document).unwrap();
    let document = open_bytes(object_bytes(&editor));

    let original = open_poi("test.doc");
    assert_eq!(document.text().unwrap(), original.text().unwrap());
    for _ in 0..2 {
        let error = document.font_table().unwrap_err();
        assert!(matches!(error, Error::Corrupted(_)), "{error}");
        assert!(error.to_string().contains("invalid font table"), "{error}");
        assert!(document.font_name(0).is_err());
    }
}
