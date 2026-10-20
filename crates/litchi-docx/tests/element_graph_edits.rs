//! Package-graph changes staged with main-document edits (record 0251b):
//! hyperlink insertion and retargeting publish their relationship and their
//! markup in one commit, replay exactly, and invert exactly.

use litchi_core::Position;
use litchi_docx::document::{ElementPath, ElementStep, HyperlinkTarget, NewRun};
use litchi_docx::{Builder, Inline, Package};

fn source() -> Package {
    let mut builder = Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("See ");
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("Second");
    builder.build().unwrap()
}

fn paragraph(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

fn links(package: &Package, paragraph: usize) -> Vec<(Option<String>, Option<String>, String)> {
    let document = package.document().unwrap();
    let paragraph = &document.paragraphs().unwrap()[paragraph];
    paragraph
        .inlines_with_relationships(document.part().part().rels())
        .unwrap()
        .into_iter()
        .filter_map(|inline| match inline {
            Inline::Hyperlink(link) => Some((
                link.link().url().map(str::to_owned),
                link.link().anchor().map(str::to_owned),
                link.runs()
                    .iter()
                    .map(|run| run.text().unwrap())
                    .collect::<String>(),
            )),
            _ => None,
        })
        .collect()
}

fn main_relationships(package: &Package) -> usize {
    package
        .opc_package()
        .main_document_part()
        .unwrap()
        .rels()
        .len()
}

#[test]
fn url_hyperlinks_publish_markup_and_relationship_atomically() {
    let base = source();
    let before_relationships = main_relationships(&base);
    let mut edit = base.edit().unwrap();
    let position = edit
        .insert_hyperlink_at(
            &paragraph(0),
            None,
            &HyperlinkTarget::Url("https://example.test/a?b=1&c=2".to_owned()),
            &NewRun::new("the site"),
        )
        .unwrap();
    assert_eq!(position, Position::new(0));
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (changed, patch) = commit.into_parts();
    assert_eq!(main_relationships(&changed), before_relationships + 1);
    assert_eq!(main_relationships(&base), before_relationships);
    assert_eq!(
        links(&changed, 0),
        [(
            Some("https://example.test/a?b=1&c=2".to_owned()),
            None,
            "the site".to_owned()
        )]
    );
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    assert_eq!(links(&reopened, 0), links(&changed, 0));
    assert_eq!(
        reopened.document().unwrap().text().unwrap(),
        "See the siteSecond"
    );

    // The patch replays onto the source and its inverse restores it.
    let replayed = base.apply(&patch).unwrap().into_package();
    assert_eq!(links(&replayed, 0), links(&changed, 0));
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(main_relationships(&restored), before_relationships);
    assert!(links(&restored, 0).is_empty());
    assert_eq!(
        restored.to_bytes().unwrap(),
        base.clone().to_bytes().unwrap()
    );
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    // Replaying onto the changed package fails its preconditions.
    assert!(changed.apply(&patch).is_err());
}

#[test]
fn bookmark_hyperlinks_and_retargeting() {
    let base = source();
    let relationships = main_relationships(&base);
    let mut edit = base.edit().unwrap();
    edit.insert_hyperlink_at(
        &paragraph(1),
        Some(Position::new(0)),
        &HyperlinkTarget::Anchor("Chapter_1".to_owned()),
        &NewRun::new("go "),
    )
    .unwrap();
    let anchored = edit.commit().unwrap().into_package();
    assert_eq!(main_relationships(&anchored), relationships);
    assert_eq!(
        links(&anchored, 1),
        [(None, Some("Chapter_1".to_owned()), "go ".to_owned())]
    );
    assert_eq!(
        anchored.document().unwrap().text().unwrap(),
        "See go Second"
    );

    let mut edit = anchored.edit().unwrap();
    edit.set_hyperlink_target_at(
        &paragraph(1).child(ElementStep::Hyperlink(Position::new(0))),
        &HyperlinkTarget::Url("https://example.test/".to_owned()),
    )
    .unwrap();
    let (retargeted, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(
        links(&retargeted, 1),
        [(
            Some("https://example.test/".to_owned()),
            None,
            "go ".to_owned()
        )]
    );
    let back = retargeted.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(links(&back, 1), links(&anchored, 1));
}

#[test]
fn hyperlinks_go_into_nested_paragraphs() {
    let mut builder = Builder::new().unwrap();
    let table = builder.document().unwrap().add_table(1, 1);
    table.cell(0, 0).unwrap().set_text("cell");
    let base = builder.build().unwrap();
    let cell_paragraph = ElementPath::new([
        ElementStep::Table(Position::new(0)),
        ElementStep::Row(Position::new(0)),
        ElementStep::Cell(Position::new(0)),
        ElementStep::Paragraph(Position::new(0)),
    ]);
    let mut edit = base.edit().unwrap();
    edit.insert_hyperlink_at(
        &cell_paragraph,
        None,
        &HyperlinkTarget::Url("https://example.test/cell".to_owned()),
        &NewRun::new(" link"),
    )
    .unwrap();
    let changed = edit.commit().unwrap().into_package();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.document().unwrap().text().unwrap(), "cell link");
}

#[test]
fn graph_only_changes_are_not_noops_and_source_owners_refuse_them() {
    let base = source();
    let mut edit = base.edit().unwrap();
    edit.add_external_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://example.test/",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_noop());
    assert_eq!(
        main_relationships(commit.package()),
        main_relationships(&base) + 1
    );

    // A source-backed owner publishes the main document only, so a patch
    // carrying a graph change is refused there.
    let source_backed =
        litchi_docx::source_backed::Package::from_bytes(&base.clone().to_bytes().unwrap()).unwrap();
    assert!(source_backed.apply(commit.patch()).is_err());
}

fn main_xml(package: &Package) -> String {
    String::from_utf8(
        package
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
}

#[test]
fn comments_are_added_into_new_and_existing_parts_and_removed_exactly() {
    use litchi_docx::document::NewComment;

    let base = source();
    let mut edit = base.edit().unwrap();
    let first = edit
        .add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Ada", "First note")
                .with_initials("AL")
                .with_date("2026-10-13T00:00:00Z"),
        )
        .unwrap();
    let (one, add_patch) = edit.commit().unwrap().into_parts();
    let comments = one.document().unwrap().comments().unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].id(), first);
    assert_eq!(comments[0].author(), "Ada");
    assert_eq!(comments[0].initials(), Some("AL"));
    assert_eq!(comments[0].text().unwrap(), "First note");
    let xml = main_xml(&one);
    assert!(
        xml.contains(&format!(r#"<w:commentRangeStart w:id="{first}"/>"#)),
        "{xml}"
    );
    assert!(xml.contains(&format!(r#"<w:commentReference w:id="{first}"/>"#)));

    // A second comment, on one run of the second paragraph, joins the part.
    let mut edit = one.edit().unwrap();
    let second = edit
        .add_comment_at(
            &paragraph(1),
            Some((Position::new(0), Position::new(0))),
            &NewComment::new("Bo", "Second note"),
        )
        .unwrap();
    assert_eq!(second, first + 1);
    let two = edit.commit().unwrap().into_package();
    let reopened = Package::from_bytes(&two.to_bytes().unwrap()).unwrap();
    let texts: Vec<String> = reopened
        .document()
        .unwrap()
        .comments()
        .unwrap()
        .iter()
        .map(|comment| comment.text().unwrap())
        .collect();
    assert_eq!(texts, ["First note", "Second note"]);
    assert_eq!(reopened.document().unwrap().text().unwrap(), "See Second");

    // Removal strips the entry, the markers and the reference run.
    let mut edit = reopened.edit().unwrap();
    edit.remove_comment(first).unwrap();
    let (removed, remove_patch) = edit.commit().unwrap().into_parts();
    let remaining = removed.document().unwrap().comments().unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].id(), second);
    let xml = main_xml(&removed);
    assert!(!xml.contains(&format!(r#"w:id="{first}""#)), "{xml}");
    assert_eq!(removed.document().unwrap().text().unwrap(), "See Second");
    let restored = removed
        .apply(&remove_patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.document().unwrap().comments().unwrap().len(), 2);

    // The first add's inverse removes the part it created.
    let back = one.apply(&add_patch.inverse()).unwrap().into_package();
    assert!(back.document().unwrap().comments().unwrap().is_empty());
    assert_eq!(back.to_bytes().unwrap(), base.clone().to_bytes().unwrap());
    assert!(one.to_bytes().is_ok());
}

/// A 2x1 PNG header (enough for the size reader and the format check).
fn png() -> Vec<u8> {
    let mut png = vec![0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13];
    png.extend_from_slice(b"IHDR");
    png.extend_from_slice(&2u32.to_be_bytes());
    png.extend_from_slice(&1u32.to_be_bytes());
    png.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    png
}

#[test]
fn pictures_publish_their_media_part_and_drawing_atomically() {
    use litchi_docx::document::NewPicture;

    let base = source();
    let parts = base.opc_package().iter_parts().count();
    let mut edit = base.edit().unwrap();
    let run = edit
        .insert_picture_at(
            &paragraph(0),
            Some(Position::new(0)),
            &NewPicture::new(png()).with_description("A <logo>"),
        )
        .unwrap();
    assert_eq!(run, Position::new(0));
    // A second picture gets the next media name and drawing identifier.
    edit.insert_picture_at(
        &paragraph(1),
        None,
        &NewPicture::new(png()).with_size(Some(914_400), None),
    )
    .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(changed.opc_package().iter_parts().count(), parts + 2);
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let document = reopened.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    let first = paragraphs[0].images().unwrap();
    assert_eq!(first.len(), 1);
    assert_eq!(
        (first[0].width_emu(), first[0].height_emu()),
        (19_050, 9_525)
    );
    assert_eq!(first[0].description(), "A <logo>");
    let second = paragraphs[1].images().unwrap();
    assert_eq!(
        (second[0].width_emu(), second[0].height_emu()),
        (914_400, 457_200)
    );
    assert_ne!(first[0].r_embed(), second[0].r_embed());
    let main = reopened.opc_package().main_document_part().unwrap();
    let target = main
        .rels()
        .get(first[0].r_embed())
        .unwrap()
        .target_partname()
        .unwrap();
    assert_eq!(target.as_str(), "/word/media/image1.png");
    assert_eq!(
        reopened.opc_package().get_part(&target).unwrap().blob(),
        png().as_slice()
    );
    let xml = main_xml(&reopened);
    assert!(
        xml.contains(r#"<wp:docPr id="1" name="Picture 1""#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<wp:docPr id="2" name="Picture 2""#),
        "{xml}"
    );
    assert_eq!(reopened.document().unwrap().text().unwrap(), "See Second");

    // Replay and exact inverse.
    let replayed = base.apply(&patch).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(
        restored.to_bytes().unwrap(),
        base.clone().to_bytes().unwrap()
    );

    // Unrecognized bytes and out-of-range extents are refused.
    let mut edit = base.edit().unwrap();
    assert!(
        edit.insert_picture_at(
            &paragraph(0),
            None,
            &NewPicture::new(b"not an image".to_vec())
        )
        .is_err()
    );
    assert!(
        edit.insert_picture_at(
            &paragraph(0),
            None,
            &NewPicture::new(png()).with_size(Some(0), Some(10)),
        )
        .is_err()
    );
}

#[test]
fn lists_create_the_numbering_part_and_extend_it_exactly() {
    use litchi_docx::numbering::{Format, ListKind};
    use litchi_docx::properties::{Numbering, Paragraph as ParagraphProperties};

    fn joined(num_id: u32) -> impl FnOnce(&mut ParagraphProperties) {
        move |properties| {
            let mut numbering = Numbering::default();
            numbering.id = Some(num_id);
            numbering.level = Some(0);
            properties.numbering = Some(numbering);
        }
    }

    let base = source();
    let mut edit = base.edit().unwrap();
    let bullets = edit.add_list(ListKind::Bullet, None).unwrap();
    let numbers = edit.add_list(ListKind::Decimal, Some((0, 5))).unwrap();
    // The builder's template already has a numbering part; new lists take
    // the next identifiers.
    assert!(bullets.num_id() > 1);
    assert_eq!(
        (numbers.abstract_num_id(), numbers.num_id()),
        (bullets.abstract_num_id() + 1, bullets.num_id() + 1)
    );
    edit.format_paragraph_at(&paragraph(0), joined(bullets.num_id()))
        .unwrap();
    edit.format_paragraph_at(&paragraph(1), joined(numbers.num_id()))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .opc_package()
            .get_part(&litchi_opc::PackURI::new("/word/numbering.xml".to_owned()).unwrap())
            .is_ok()
    );
    let document = reopened.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(
        paragraphs[0].numbering().unwrap().unwrap().num_id,
        bullets.num_id()
    );
    assert_eq!(
        paragraphs[1].numbering().unwrap().unwrap().num_id,
        numbers.num_id()
    );
    let collection = document.numbering().unwrap().unwrap();
    let level = |num_id: u32| {
        let instance = collection.get_num(num_id).unwrap();
        collection
            .get_abstract_num(instance.abstract_num_id())
            .unwrap()
            .level(0)
            .unwrap()
            .clone()
    };
    assert_eq!(level(bullets.num_id()).format, Format::Bullet);
    assert_eq!(level(numbers.num_id()).format, Format::Decimal);
    assert_eq!(level(numbers.num_id()).start, 5);
    let instances = collection.nums().len();

    // A second edit extends the existing part with the next identifiers.
    let mut edit = reopened.edit().unwrap();
    let third = edit.add_list(ListKind::Decimal, None).unwrap();
    assert_eq!(third.num_id(), numbers.num_id() + 1);
    let extended = edit.commit().unwrap().into_package();
    assert_eq!(
        extended
            .document()
            .unwrap()
            .numbering()
            .unwrap()
            .unwrap()
            .nums()
            .len(),
        instances + 1
    );

    // Replay and exact inverse of the first edit.
    let replayed = base.apply(&patch).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(
        restored.to_bytes().unwrap(),
        base.clone().to_bytes().unwrap()
    );
}
