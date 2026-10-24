#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration-test fixtures favor concise panic-driven assertions over production-style ergonomics"
)]

//! Record 0324: paragraph and character styles, tab stops, `HYPERLINK`
//! fields, standard bookmarks and section page setup through the primary DOC
//! package transaction. Every edit is read back from the published bytes,
//! inverted byte-exactly, and setting a value the target has is a no-op.

use litchi_core::Position;
use litchi_doc::Package;
use litchi_doc::body_text::{
    CharacterValue, Error as BodyError, FieldPart, HyperlinkInsert, Refusal, SectionValue,
    TextTarget,
};
use litchi_doc::parts::pap::{TabAlignment, TabLeader, TabStop};
use litchi_doc::section::{BreakKind, PageOrientation};
use std::io::Cursor;

fn fixture(name: &str) -> Package {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name);
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

fn reopen(package: &Package) -> litchi_doc::Document {
    let bytes = package.to_bytes().unwrap();
    let mut raw = litchi_doc::raw::Reader::from_reader(Cursor::new(bytes)).unwrap();
    raw.document().unwrap()
}

/// Commit `edit`, check that the patch inverts to the exact source bytes,
/// and return the published package.
fn commit(source: &Package, edit: litchi_doc::Edit<'static>) -> Package {
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (package, patch) = commit.into_parts();
    let restored = package.apply(&patch.inverse()).unwrap();
    assert_eq!(
        restored.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
    package
}

fn refusal(error: litchi_doc::package::edit::Error) -> Refusal {
    match error {
        litchi_doc::package::edit::Error::Body(BodyError::Refused(refusal)) => refusal,
        other => panic!("expected a refusal, got {other:?}"),
    }
}

fn style_index(document: &litchi_doc::Document, name: &str) -> u16 {
    let sheet = document.stylesheet().unwrap();
    (0..u16::try_from(sheet.styles().len()).unwrap())
        .find(|index| sheet.get(*index).is_some_and(|style| style.name == name))
        .unwrap()
}

#[test]
fn a_paragraph_style_is_applied_kept_and_inverted() {
    let source = fixture("Lists.doc");
    let document = source.document();
    let heading = style_index(&document, "Heading 1");
    let before: Vec<_> = (0..4)
        .map(|index| {
            document
                .body_paragraph(Position::new(index))
                .unwrap()
                .unwrap()
                .properties()
                .style_index
        })
        .collect();
    let mut edit = source.edit().unwrap();
    edit.set_paragraph_style(Position::new(1), heading).unwrap();
    let package = commit(&source, edit);
    let after = reopen(&package);
    for (index, style) in before.iter().enumerate() {
        let paragraph = after.body_paragraph(Position::new(index)).unwrap().unwrap();
        let expected = if index == 1 { Some(heading) } else { *style };
        assert_eq!(paragraph.properties().style_index, expected, "{index}");
    }
    assert_eq!(after.body_paragraph_count().unwrap(), 40);

    // The style the paragraph already has is an exact no-op.
    let mut edit = package.edit().unwrap();
    edit.set_paragraph_style(Position::new(1), heading).unwrap();
    assert!(!edit.commit().unwrap().diagnostics().changed());

    // A character style is not a paragraph style.
    let character = style_index(&document, "Heading 1 Char");
    let mut edit = source.edit().unwrap();
    assert_eq!(
        refusal(
            edit.set_paragraph_style(Position::new(1), character)
                .unwrap_err()
        ),
        Refusal::StyleNotFound
    );
    assert_eq!(
        refusal(
            edit.set_paragraph_style(Position::new(1), 4000)
                .unwrap_err()
        ),
        Refusal::StyleNotFound
    );
}

#[test]
fn a_character_style_is_applied_to_one_run_and_cleared() {
    let source = fixture("Lists.doc");
    let document = source.document();
    let character = style_index(&document, "Heading 1 Char");
    let paragraph = document.body_paragraph(Position::new(1)).unwrap().unwrap();
    let length = u32::try_from(paragraph.text().unwrap().encode_utf16().count()).unwrap();
    let target = TextTarget::BodySpan {
        paragraph: Position::new(1),
        start: 0,
        end: length,
    };
    let mut edit = source.edit().unwrap();
    edit.set_character_value(target, CharacterValue::Style(character))
        .unwrap();
    let package = commit(&source, edit);
    let after = reopen(&package);
    let runs = after
        .body_paragraph(Position::new(1))
        .unwrap()
        .unwrap()
        .runs()
        .unwrap();
    assert!(
        runs.iter()
            .all(|run| run.properties().style_index == Some(character))
    );

    // The default character style (10) clears the direct style.
    let mut edit = package.edit().unwrap();
    edit.set_character_value(target, CharacterValue::Style(10))
        .unwrap();
    let cleared = edit.commit().unwrap().into_package();
    let runs = reopen(&cleared)
        .body_paragraph(Position::new(1))
        .unwrap()
        .unwrap()
        .runs()
        .unwrap();
    assert!(
        runs.iter()
            .all(|run| run.properties().style_index.is_none())
    );

    let heading = style_index(&document, "Heading 1");
    let mut edit = source.edit().unwrap();
    assert_eq!(
        refusal(
            edit.set_character_value(target, CharacterValue::Style(heading))
                .unwrap_err()
        ),
        Refusal::StyleNotFound
    );
}

#[test]
fn tab_stops_are_set_replaced_and_cleared() {
    let source = fixture("Lists.doc");
    let stops = [
        TabStop {
            position: 2880,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dots,
        },
        TabStop {
            position: 1440,
            alignment: TabAlignment::Center,
            leader: TabLeader::None,
        },
    ];
    let mut edit = source.edit().unwrap();
    edit.set_paragraph_tabs(Position::new(1), &stops).unwrap();
    assert_eq!(edit.paragraph_tabs(Position::new(1)).unwrap().len(), 2);
    let package = commit(&source, edit);
    let read = reopen(&package)
        .body_paragraph(Position::new(1))
        .unwrap()
        .unwrap()
        .properties()
        .tab_stops
        .clone();
    assert_eq!(read, vec![stops[1], stops[0]]);

    // The same stops are a no-op; an empty list clears them.
    let mut edit = package.edit().unwrap();
    edit.set_paragraph_tabs(Position::new(1), &[stops[1], stops[0]])
        .unwrap();
    assert!(!edit.commit().unwrap().diagnostics().changed());
    let mut edit = package.edit().unwrap();
    edit.set_paragraph_tabs(Position::new(1), &[]).unwrap();
    let cleared = edit.commit().unwrap().into_package();
    assert!(
        reopen(&cleared)
            .body_paragraph(Position::new(1))
            .unwrap()
            .unwrap()
            .properties()
            .tab_stops
            .is_empty()
    );

    let mut edit = source.edit().unwrap();
    let duplicate = [stops[0], stops[0]];
    assert_eq!(
        refusal(
            edit.set_paragraph_tabs(Position::new(1), &duplicate)
                .unwrap_err()
        ),
        Refusal::InvalidProperty
    );
}

#[test]
fn a_hyperlink_is_inserted_retargeted_retitled_and_removed() {
    let source = fixture("Lists.doc");
    let original = source
        .document()
        .body_paragraph_text(Position::new(1))
        .unwrap()
        .unwrap()
        .to_owned();
    let mut edit = source.edit().unwrap();
    edit.insert_hyperlink(
        Position::new(1),
        4,
        &HyperlinkInsert {
            instruction: r#"HYPERLINK "http://example.com/""#.to_string(),
            text: "link".to_string(),
        },
    )
    .unwrap();
    let package = commit(&source, edit);
    let document = reopen(&package);
    let links = document.hyperlinks().unwrap();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].destination(), "http://example.com/");
    assert_eq!(links[0].display_text(), "link");
    let start = document
        .fields()
        .unwrap()
        .iter()
        .find(|field| field.field.field_type == litchi_doc::FieldType::Hyperlink)
        .unwrap()
        .field
        .start_cp;
    // Every other paragraph keeps its text.
    assert_eq!(
        document.body_paragraph_text(Position::new(0)).unwrap(),
        source
            .document()
            .body_paragraph_text(Position::new(0))
            .unwrap()
    );

    let mut edit = package.edit().unwrap();
    edit.replace_text(
        TextTarget::Field {
            start,
            part: FieldPart::Instruction,
        },
        r#" HYPERLINK "http://example.org/" "#,
    )
    .unwrap();
    edit.replace_text(
        TextTarget::Field {
            start,
            part: FieldPart::Result,
        },
        "other text",
    )
    .unwrap();
    let changed = commit(&package, edit);
    let links = reopen(&changed).hyperlinks().unwrap();
    assert_eq!(links[0].destination(), "http://example.org/");
    assert_eq!(links[0].display_text(), "other text");

    let mut edit = changed.edit().unwrap();
    edit.remove_field(start).unwrap();
    let removed = commit(&changed, edit);
    let document = reopen(&removed);
    assert!(document.hyperlinks().unwrap().is_empty());
    assert_eq!(
        document
            .body_paragraph_text(Position::new(1))
            .unwrap()
            .unwrap(),
        original
    );

    let mut edit = removed.edit().unwrap();
    assert_eq!(
        refusal(edit.remove_field(start).unwrap_err()),
        Refusal::FieldNotFound
    );
    assert_eq!(
        refusal(
            edit.insert_hyperlink(
                Position::new(1),
                10_000,
                &HyperlinkInsert {
                    instruction: "HYPERLINK \"x\"".to_string(),
                    text: "x".to_string(),
                },
            )
            .unwrap_err()
        ),
        Refusal::TargetNotFound
    );
}

#[test]
fn bookmarks_are_added_renamed_and_removed() {
    let source = fixture("Lists.doc");
    let mut edit = source.edit().unwrap();
    edit.add_bookmark("First", Position::new(1), 0, 4).unwrap();
    edit.add_bookmark("Second", Position::new(0), 2, 2).unwrap();
    assert_eq!(
        refusal(
            edit.add_bookmark("first", Position::new(2), 0, 1)
                .unwrap_err()
        ),
        Refusal::BookmarkName
    );
    assert_eq!(
        refusal(
            edit.add_bookmark("has space", Position::new(2), 0, 1)
                .unwrap_err()
        ),
        Refusal::BookmarkName
    );
    let package = commit(&source, edit);
    let bookmarks = reopen(&package).bookmarks().unwrap();
    let names: Vec<_> = bookmarks
        .iter()
        .map(|bookmark| bookmark.name.as_str())
        .collect();
    assert_eq!(names, ["Second", "First"]);
    assert_eq!(bookmarks[1].end - bookmarks[1].start, 4);

    let mut edit = package.edit().unwrap();
    edit.rename_bookmark("First", "Renamed").unwrap();
    edit.remove_bookmark("Second").unwrap();
    let edited = commit(&package, edit);
    let bookmarks = reopen(&edited).bookmarks().unwrap();
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].name, "Renamed");

    let mut edit = edited.edit().unwrap();
    assert_eq!(
        refusal(edit.remove_bookmark("Second").unwrap_err()),
        Refusal::BookmarkNotFound
    );
    edit.remove_bookmark("Renamed").unwrap();
    let empty = edit.commit().unwrap().into_package();
    assert!(reopen(&empty).bookmarks().unwrap().is_empty());
}

#[test]
fn section_page_setup_is_set_and_equal_values_are_no_ops() {
    let source = fixture("Lists.doc");
    let mut edit = source.edit().unwrap();
    edit.set_section_values(
        Position::new(0),
        &[
            SectionValue::Landscape(true),
            SectionValue::PageWidth(15_840),
            SectionValue::PageHeight(12_240),
            SectionValue::MarginLeft(720),
            SectionValue::Columns(2),
            SectionValue::Break(BreakKind::OddPage),
        ],
    )
    .unwrap();
    let package = commit(&source, edit);
    let document = reopen(&package);
    let section = &document.sections()[0];
    assert_eq!(section.page.orientation, PageOrientation::Landscape);
    assert_eq!(section.page.width_twips, 15_840);
    assert_eq!(section.page.height_twips, 12_240);
    assert_eq!(section.page.margins.left_twips, 720);
    assert_eq!(section.columns.count(), 2);
    assert_eq!(section.break_kind, BreakKind::OddPage);

    let mut edit = package.edit().unwrap();
    edit.set_section_values(Position::new(0), &[SectionValue::PageWidth(15_840)])
        .unwrap();
    assert!(!edit.commit().unwrap().diagnostics().changed());
    let mut edit = package.edit().unwrap();
    assert_eq!(
        refusal(
            edit.set_section_values(Position::new(9), &[SectionValue::TitlePage(true)])
                .unwrap_err()
        ),
        Refusal::SectionNotFound
    );
    assert_eq!(
        refusal(
            edit.set_section_values(Position::new(0), &[SectionValue::Columns(0)])
                .unwrap_err()
        ),
        Refusal::InvalidProperty
    );
}
