//! Styles, numbering, header/footer and note edits staged with
//! main-document edits (record 0319): each publishes its part and its
//! main-story markup in one commit, reopens, replays and inverts exactly.

use litchi_core::Position;
use litchi_docx::document::{ElementPath, ElementStep, NewNote, NewStory, NoteKind};
use litchi_docx::header_footer::{Kind, Role};
use litchi_docx::numbering::catalog::{LevelChange, NewAbstractNum, NewNum};
use litchi_docx::styles::Type;
use litchi_docx::styles::catalog::{Flag, NewStyle, StyleChange};
use litchi_docx::{Builder, Package};

fn source() -> Package {
    let mut builder = Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("First");
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("Second");
    builder.build().unwrap()
}

fn paragraph(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

fn reopened(package: Package) -> Package {
    Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
}

fn bytes(package: &Package) -> Vec<u8> {
    package.clone().to_bytes().unwrap()
}

#[test]
fn styles_are_added_changed_removed_and_inverted() {
    let base = source();
    let mut edit = base.edit().unwrap();
    let mut run = litchi_docx::properties::Run::default();
    run.bold = Some(true);
    edit.add_style(
        &NewStyle::new("Accent", Type::Character)
            .with_custom(true)
            .with_change(
                StyleChange::default()
                    .with_name("Accent Text")
                    .with_run(run),
            ),
    )
    .unwrap();
    let (added, add_patch) = edit.commit().unwrap().into_parts();
    let catalog = reopened(added.clone()).style_catalog().unwrap().unwrap();
    let accent = catalog.resolve("Accent Text").unwrap();
    assert_eq!(accent.id(), "Accent");
    assert_eq!(catalog.run_properties(accent).unwrap().bold, Some(true));

    let mut edit = added.edit().unwrap();
    edit.set_style(
        "Accent",
        &StyleChange::default().with_flag(Flag::QuickFormat, Some(true)),
    )
    .unwrap();
    let (changed, change_patch) = edit.commit().unwrap().into_parts();
    assert_eq!(
        changed
            .style_catalog()
            .unwrap()
            .unwrap()
            .get("Accent")
            .unwrap()
            .flag(Flag::QuickFormat),
        Some(true)
    );
    assert_eq!(
        bytes(changed.apply(&change_patch.inverse()).unwrap().package()),
        bytes(&added)
    );

    let mut edit = changed.edit().unwrap();
    edit.remove_style("Accent").unwrap();
    assert!(edit.remove_style("Missing").is_err());
    let removed = edit.commit().unwrap().into_package();
    assert!(
        removed
            .style_catalog()
            .unwrap()
            .unwrap()
            .get("Accent")
            .is_none()
    );

    // The add replays onto its base and inverts exactly.
    assert_eq!(
        bytes(base.apply(&add_patch).unwrap().package()),
        bytes(&added)
    );
    assert_eq!(
        bytes(added.apply(&add_patch.inverse()).unwrap().package()),
        bytes(&base)
    );
}

#[test]
fn numbering_definitions_instances_and_levels_round_trip() {
    let base = source();
    let mut edit = base.edit().unwrap();
    let definition = edit
        .add_abstract_num(&NewAbstractNum::new(false).with_name("Steps"))
        .unwrap();
    let num = edit
        .add_num(&NewNum::new(definition).with_start_override(0, 3))
        .unwrap();
    edit.set_level(
        definition,
        1,
        &LevelChange::default().with_format(Some("upperLetter".into())),
    )
    .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let catalog = reopened(changed.clone())
        .numbering_catalog()
        .unwrap()
        .unwrap();
    assert_eq!(
        catalog.abstract_num(definition).unwrap().name(),
        Some("Steps")
    );
    assert_eq!(catalog.format_of(num, 1), Some("upperLetter"));
    assert_eq!(
        catalog.num(num).unwrap().overrides()[0].1.as_deref(),
        Some("3")
    );
    assert_eq!(
        bytes(changed.apply(&patch.inverse()).unwrap().package()),
        bytes(&base)
    );

    let mut edit = changed.edit().unwrap();
    // The definition goes with its instance.
    edit.remove_abstract_num(definition).unwrap();
    let emptied = edit.commit().unwrap().into_package();
    let catalog = emptied.numbering_catalog().unwrap().unwrap();
    assert!(catalog.abstract_num(definition).is_none() && catalog.num(num).is_none());
}

#[test]
fn headers_and_footers_are_added_retyped_written_and_removed() {
    let base = source();
    let mut edit = base.edit().unwrap();
    let header = edit
        .add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Draft"))
        .unwrap();
    let footer = edit
        .add_header_footer(Role::Footer, Kind::FirstPage, &NewStory::new("Page"))
        .unwrap();
    assert!(
        edit.add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Again"))
            .is_err()
    );
    let (added, patch) = edit.commit().unwrap().into_parts();
    let again = reopened(added.clone());
    let document = again.document().unwrap();
    assert_eq!(document.headers().unwrap()[0].text().unwrap(), "Draft");
    let footers = document.footers().unwrap();
    assert_eq!(footers[0].kind(), Kind::FirstPage);
    let inventory = document.section_inventory().unwrap();
    assert_eq!(inventory.sections().last().unwrap().headers().len(), 1);
    let main = String::from_utf8(
        again
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    assert!(main.contains("<w:titlePg/>"), "{main}");
    assert_eq!(
        bytes(added.apply(&patch.inverse()).unwrap().package()),
        bytes(&base)
    );
    assert_eq!(bytes(base.apply(&patch).unwrap().package()), bytes(&added));

    let mut edit = again.edit().unwrap();
    edit.set_story_text(&header, "Final").unwrap();
    edit.set_header_footer_kind(&header, Kind::EvenPage)
        .unwrap();
    let italic = |run: &mut litchi_docx::properties::Run| run.italic = Some(true);
    edit.format_story(&footer, None, Some(&italic)).unwrap();
    let written = reopened(edit.commit().unwrap().into_package());
    let headers = written.document().unwrap().headers().unwrap();
    assert_eq!(headers[0].text().unwrap(), "Final");
    assert_eq!(headers[0].kind(), Kind::EvenPage);

    let mut edit = written.edit().unwrap();
    edit.remove_header_footer(&header).unwrap();
    let (removed, remove_patch) = edit.commit().unwrap().into_parts();
    let removed_again = reopened(removed.clone());
    assert!(
        removed_again
            .document()
            .unwrap()
            .headers()
            .unwrap()
            .is_empty()
    );
    assert_eq!(removed_again.header_footer_entries().unwrap().len(), 1);
    assert_eq!(
        bytes(removed.apply(&remove_patch.inverse()).unwrap().package()),
        bytes(&written)
    );
}

#[test]
fn notes_are_added_written_and_removed_with_their_references() {
    let base = source();
    let mut edit = base.edit().unwrap();
    let first = edit
        .add_note_at(
            NoteKind::Footnote,
            &paragraph(0),
            None,
            &NewNote::new("One"),
        )
        .unwrap();
    let second = edit
        .add_note_at(
            NoteKind::Endnote,
            &paragraph(1),
            Some(Position::new(0)),
            &NewNote::new("Two"),
        )
        .unwrap();
    let (added, patch) = edit.commit().unwrap().into_parts();
    assert_eq!((first, second), (1, 1));
    let again = reopened(added.clone());
    let document = again.document().unwrap();
    assert_eq!(
        document.footnotes().unwrap()[0].text().unwrap().trim(),
        "One"
    );
    assert_eq!(
        document.endnotes().unwrap()[0].text().unwrap().trim(),
        "Two"
    );
    assert_eq!(document.text().unwrap().matches("First").count(), 1);
    assert_eq!(
        bytes(added.apply(&patch.inverse()).unwrap().package()),
        bytes(&base)
    );

    let mut edit = again.edit().unwrap();
    edit.set_note_text(NoteKind::Footnote, first, "Uno")
        .unwrap();
    assert!(edit.set_note_text(NoteKind::Footnote, 9, "x").is_err());
    let written = edit.commit().unwrap().into_package();
    assert_eq!(
        written.document().unwrap().footnotes().unwrap()[0]
            .text()
            .unwrap()
            .trim(),
        "Uno"
    );

    let mut edit = written.edit().unwrap();
    edit.remove_note(NoteKind::Footnote, first).unwrap();
    let removed = reopened(edit.commit().unwrap().into_package());
    assert!(removed.document().unwrap().footnotes().unwrap().is_empty());
    let main = String::from_utf8(
        removed
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    assert!(!main.contains("footnoteReference"), "{main}");
    assert!(main.contains("endnoteReference"));
}
