#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! Paragraph structure (insert, remove, move), character-run spans and valued
//! character formatting through the primary DOC package transaction.

use litchi_core::Position;
use litchi_doc::body_text::{CharacterValue, Error as BodyError, Refusal, TextTarget};
use litchi_doc::writer::{CharacterFormatting, ParagraphFormatting};
use litchi_doc::{Builder, Package};
use std::io::Cursor;

fn build(paragraphs: &[&str]) -> Package {
    let mut builder = Builder::new();
    for text in paragraphs {
        builder.add_paragraph(text).unwrap();
    }
    builder.build().unwrap()
}

fn mixed_runs() -> Package {
    let mut builder = Builder::new();
    builder.add_paragraph("first").unwrap();
    builder
        .add_paragraph_runs(
            vec![
                (
                    "plain ".to_string(),
                    CharacterFormatting {
                        font_name: Some("Arial".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "bold".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        font_name: Some("Courier New".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " tail".to_string(),
                    CharacterFormatting {
                        italic: Some(true),
                        font_name: Some("Arial".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    builder.add_paragraph("last").unwrap();
    builder.build().unwrap()
}

/// Reopen the published bytes through the physical reader, so every
/// assertion reads the edited file, not the edit's own model.
fn reopen(package: &Package) -> litchi_doc::Document {
    let bytes = package.to_bytes().unwrap();
    let mut raw = litchi_doc::raw::Reader::from_reader(Cursor::new(bytes)).unwrap();
    raw.document().unwrap()
}

fn texts(package: &Package) -> Vec<String> {
    let document = reopen(package);
    let texts = document
        .body_paragraphs()
        .map(|text| text.unwrap().to_owned())
        .collect::<Vec<_>>();
    // The full semantic reader must agree with the body traversal.
    assert_eq!(document.body_paragraph_count().unwrap(), texts.len());
    document.paragraphs().unwrap();
    texts
}

fn commit(package: &Package, edit: impl FnOnce(&mut litchi_doc::Edit<'static>)) -> (Package, bool) {
    let mut transaction = package.edit().unwrap();
    edit(&mut transaction);
    let commit = transaction.commit().unwrap();
    let changed = commit.diagnostics().changed();
    (commit.into_package(), changed)
}

fn refusal(error: litchi_doc::package::edit::Error) -> Refusal {
    match error {
        litchi_doc::package::edit::Error::Body(BodyError::Refused(refusal)) => refusal,
        other => panic!("expected a refusal, got {other:?}"),
    }
}

#[test]
fn insert_paragraph_before_between_and_after() {
    let package = build(&["a", "b", "c"]);
    let (front, changed) = commit(&package, |edit| {
        edit.insert_paragraph(Position::new(0), "front").unwrap();
    });
    assert!(changed);
    assert_eq!(texts(&front), ["front", "a", "b", "c"]);

    let (middle, _) = commit(&package, |edit| {
        edit.insert_paragraph(Position::new(2), "middle").unwrap();
    });
    assert_eq!(texts(&middle), ["a", "b", "middle", "c"]);

    let (end, _) = commit(&package, |edit| {
        edit.insert_paragraph(Position::new(3), "end").unwrap();
        edit.insert_paragraph(Position::new(4), "").unwrap();
    });
    assert_eq!(texts(&end), ["a", "b", "c", "end", ""]);

    // The source snapshot is unchanged.
    assert_eq!(texts(&package), ["a", "b", "c"]);

    let mut edit = package.edit().unwrap();
    assert_eq!(
        refusal(edit.insert_paragraph(Position::new(5), "x").unwrap_err()),
        Refusal::ParagraphNotFound
    );
    assert_eq!(
        refusal(edit.insert_paragraph(Position::new(0), "a\rb").unwrap_err()),
        Refusal::ReplacementContainsStructuralContent
    );
}

#[test]
fn remove_paragraph_first_middle_last_and_sole() {
    let package = build(&["a", "b", "c"]);
    let (removed, _) = commit(&package, |edit| {
        edit.remove_paragraph(Position::new(1)).unwrap();
    });
    assert_eq!(texts(&removed), ["a", "c"]);
    let (removed, _) = commit(&package, |edit| {
        edit.remove_paragraph(Position::new(0)).unwrap();
        edit.remove_paragraph(Position::new(1)).unwrap();
    });
    assert_eq!(texts(&removed), ["b"]);
    let mut edit = removed.edit().unwrap();
    assert_eq!(
        refusal(edit.remove_paragraph(Position::new(0)).unwrap_err()),
        Refusal::SoleParagraph
    );
    assert_eq!(
        refusal(edit.remove_paragraph(Position::new(1)).unwrap_err()),
        Refusal::ParagraphNotFound
    );
}

#[test]
fn move_paragraph_keeps_runs_and_formatting() {
    let package = mixed_runs();
    let (moved, changed) = commit(&package, |edit| {
        edit.move_paragraph(Position::new(1), Position::new(2))
            .unwrap();
    });
    assert!(changed);
    assert_eq!(texts(&moved), ["first", "last", "plain bold tail"]);
    let document = reopen(&moved);
    let paragraph = document.body_paragraph(Position::new(2)).unwrap().unwrap();
    let runs = paragraph.runs().unwrap();
    let shape = runs
        .iter()
        .map(|run| {
            (
                run.text().unwrap().to_owned(),
                run.bold(),
                run.italic(),
                run.font_index()
                    .and_then(|ftc| document.font_name(ftc).unwrap().map(str::to_owned)),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        shape,
        [
            ("plain ".to_owned(), None, None, Some("Arial".to_owned())),
            (
                "bold".to_owned(),
                Some(true),
                None,
                Some("Courier New".to_owned())
            ),
            (
                " tail".to_owned(),
                None,
                Some(true),
                Some("Arial".to_owned())
            ),
        ]
    );

    let (front, _) = commit(&package, |edit| {
        edit.move_paragraph(Position::new(2), Position::new(0))
            .unwrap();
    });
    assert_eq!(texts(&front), ["last", "first", "plain bold tail"]);

    // A move to the same place is an exact no-op.
    let (same, changed) = commit(&package, |edit| {
        edit.move_paragraph(Position::new(1), Position::new(1))
            .unwrap();
    });
    assert!(!changed);
    assert_eq!(same.to_bytes().unwrap(), package.to_bytes().unwrap());
}

#[test]
fn span_targets_replace_one_run_and_format_it() {
    let package = mixed_runs();
    // "plain " is 0..6, "bold" 6..10, " tail" 10..15 of paragraph 1.
    let bold = TextTarget::BodySpan {
        paragraph: Position::new(1),
        start: 6,
        end: 10,
    };
    let (edited, _) = commit(&package, |edit| {
        edit.set_character_value(bold, CharacterValue::Size(36))
            .unwrap();
        edit.set_character_value(bold, CharacterValue::Color(Some([255, 0, 0])))
            .unwrap();
        edit.set_character_value(bold, CharacterValue::Font("Arial"))
            .unwrap();
        edit.replace_text(bold, "BOLDER").unwrap();
    });
    assert_eq!(texts(&edited), ["first", "plain BOLDER tail", "last"]);
    let document = reopen(&edited);
    let paragraph = document.body_paragraph(Position::new(1)).unwrap().unwrap();
    let runs = paragraph.runs().unwrap();
    let target = runs
        .iter()
        .find(|run| run.text().unwrap() == "BOLDER")
        .expect("the replaced run");
    assert_eq!(target.font_size(), Some(36));
    assert_eq!(target.color(), Some((255, 0, 0)));
    assert_eq!(target.bold(), Some(true));
    assert_eq!(
        document.font_name(target.font_index().unwrap()).unwrap(),
        Some("Arial")
    );
    let tail = runs
        .iter()
        .find(|run| run.text().unwrap() == " tail")
        .unwrap();
    assert_eq!(tail.font_size(), None);
    assert_eq!(tail.color(), None);

    // Setting the value already carried is an exact no-op.
    let (again, changed) = commit(&edited, |edit| {
        let bolder = TextTarget::BodySpan {
            paragraph: Position::new(1),
            start: 6,
            end: 12,
        };
        edit.set_character_value(bolder, CharacterValue::Size(36))
            .unwrap();
        edit.set_character_value(bolder, CharacterValue::Color(Some([255, 0, 0])))
            .unwrap();
    });
    assert!(!changed);
    assert_eq!(again.to_bytes().unwrap(), edited.to_bytes().unwrap());
}

#[test]
fn valued_formatting_refusals_are_typed() {
    let package = mixed_runs();
    let mut edit = package.edit().unwrap();
    let whole = TextTarget::body_paragraph(Position::new(0));
    assert_eq!(
        refusal(
            edit.set_character_value(whole, CharacterValue::Font("No Such Font"))
                .unwrap_err()
        ),
        Refusal::FontNotFound
    );
    assert_eq!(
        refusal(
            edit.set_character_value(whole, CharacterValue::Size(1))
                .unwrap_err()
        ),
        Refusal::InvalidFontSize
    );
    let split = TextTarget::BodySpan {
        paragraph: Position::new(1),
        start: 4,
        end: 99,
    };
    assert_eq!(
        refusal(edit.replace_text(split, "x").unwrap_err()),
        Refusal::TargetNotFound
    );
    // A whole-paragraph color over runs that agree on it (none) is admitted.
    let mixed = TextTarget::body_paragraph(Position::new(1));
    edit.set_character_value(mixed, CharacterValue::Color(Some([1, 2, 3])))
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let document = reopen(&package);
    let paragraph = document.body_paragraph(Position::new(1)).unwrap().unwrap();
    for run in paragraph.runs().unwrap() {
        assert_eq!(run.color(), Some((1, 2, 3)));
    }
    // Clearing the direct color returns the text to automatic.
    let (cleared, changed) = commit(&package, |edit| {
        edit.set_character_value(mixed, CharacterValue::Color(None))
            .unwrap();
    });
    assert!(changed);
    let document = reopen(&cleared);
    let paragraph = document.body_paragraph(Position::new(1)).unwrap().unwrap();
    for run in paragraph.runs().unwrap() {
        assert_eq!(run.color(), None);
    }
}

#[test]
fn structural_patches_invert_exactly() {
    let package = build(&["a", "b", "c"]);
    let mut edit = package.edit().unwrap();
    edit.insert_paragraph(Position::new(1), "x").unwrap();
    edit.move_paragraph(Position::new(0), Position::new(3))
        .unwrap();
    edit.remove_paragraph(Position::new(1)).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&after), ["x", "c", "a"]);
    assert_eq!(patch.changes().len(), 3);
    let restored = after.apply(&patch.inverse()).unwrap();
    assert_eq!(
        restored.package().to_bytes().unwrap(),
        package.to_bytes().unwrap()
    );
    // The durable envelope refuses the structural vocabulary explicitly.
    assert!(
        patch
            .to_durable(litchi_core::PatchLimits::new(
                litchi_core::BlobLimits::new(0, 0, 0),
                128 * 1024,
                16,
                8,
                16 * 1024,
                64 * 1024,
            ))
            .is_err()
    );
}

type Operation = dyn Fn(&mut litchi_doc::Edit<'static>) -> litchi_doc::package::edit::Result<()>;

/// Over the DOC corpus: every admitted structural edit publishes a document
/// whose body reads back as expected, and every refusal is typed.
#[test]
fn corpus_structural_edits_read_back_or_refuse_typed() {
    let directory =
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc");
    let mut names = std::fs::read_dir(&directory)
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .filter(|name| name.ends_with(".doc"))
        .collect::<Vec<_>>();
    names.sort();
    let (mut admitted, mut refused) = (0usize, 0usize);
    for name in names {
        let Ok(package) = Package::open(directory.join(&name)) else {
            continue;
        };
        if package.edit().is_err() {
            continue;
        }
        let before = texts(&package);
        if before.len() < 2 {
            continue;
        }
        let last = before.len() - 1;
        let cases: [(&str, Box<Operation>, Vec<String>); 3] = [
            (
                "insert",
                Box::new(|edit| edit.insert_paragraph(Position::new(0), "inserted")),
                std::iter::once("inserted".to_owned())
                    .chain(before.iter().cloned())
                    .collect(),
            ),
            (
                "remove",
                Box::new(|edit| edit.remove_paragraph(Position::new(0))),
                before[1..].to_vec(),
            ),
            (
                "move",
                Box::new(move |edit| edit.move_paragraph(Position::new(0), Position::new(last))),
                before[1..]
                    .iter()
                    .cloned()
                    .chain(std::iter::once(before[0].clone()))
                    .collect(),
            ),
        ];
        for (label, operation, expected) in cases {
            let mut edit = package.edit().unwrap();
            match operation(&mut edit) {
                Ok(()) => {
                    let published = edit.commit().unwrap().into_package();
                    assert_eq!(texts(&published), expected, "{name} {label}");
                    admitted += 1;
                },
                Err(litchi_doc::package::edit::Error::Body(BodyError::Refused(_))) => {
                    refused += 1;
                },
                Err(other) => panic!("{name} {label}: untyped failure {other:?}"),
            }
        }
    }
    assert!(admitted > 20, "too few corpus edits admitted: {admitted}");
    eprintln!("corpus structural edits: {admitted} admitted, {refused} refused");
}
