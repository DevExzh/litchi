//! Regression coverage for ordinal semantic paragraph queries.

use litchi_core::Position;
use litchi_doc::Package;
use litchi_doc::writer::{CharacterFormatting, Writer};
use std::io::Cursor;

#[test]
fn paragraph_at_matches_materialized_paragraphs_and_preserves_runs() {
    let mut writer = Writer::new();
    writer.add_paragraph("first").unwrap();
    writer
        .add_paragraph_runs(
            vec![
                (
                    "bold".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " 😀".to_string(),
                    CharacterFormatting {
                        italic: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            Default::default(),
        )
        .unwrap();
    writer.add_paragraph("third").unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();
    let materialized = document.paragraphs().unwrap();

    assert_eq!(materialized.len(), 3);
    for (position, expected) in materialized.iter().enumerate() {
        let actual = document
            .paragraph_at(Position::new(position))
            .unwrap()
            .expect("materialized paragraph must be selectable");
        assert_eq!(actual.text().unwrap(), expected.text().unwrap());

        let expected_runs = expected.runs().unwrap();
        let actual_runs = actual.runs().unwrap();
        assert_eq!(actual_runs.len(), expected_runs.len());
        for (actual_run, expected_run) in actual_runs.iter().zip(expected_runs.iter()) {
            assert_eq!(actual_run.text().unwrap(), expected_run.text().unwrap());
            assert_eq!(actual_run.properties(), expected_run.properties());
        }
    }
    assert!(
        document
            .paragraph_at(Position::new(materialized.len()))
            .unwrap()
            .is_none()
    );
}
