//! Picture, comment and SmartArt changes and removals staged with
//! main-document edits (record 0335): each publishes in one commit,
//! replays exactly, inverts exactly, and an identical change is a no-op.

use litchi_core::Position;
use litchi_docx::document::{
    CommentChange, ElementPath, ElementStep, NewComment, NewPicture, NewSmartArt, PictureChange,
};
use litchi_docx::{Builder, Package};
use litchi_drawingml::diagram::Limits;
use litchi_drawingml::diagram::authoring::{Layout, Spec};

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

/// A PNG header of `width` x `height` pixels.
fn png(width: u32, height: u32) -> Vec<u8> {
    let mut png = vec![0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13];
    png.extend_from_slice(b"IHDR");
    png.extend_from_slice(&width.to_be_bytes());
    png.extend_from_slice(&height.to_be_bytes());
    png.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    png
}

/// A GIF header (another format, so the new part gets a new extension).
fn gif() -> Vec<u8> {
    let mut gif = b"GIF89a".to_vec();
    gif.extend_from_slice(&3u16.to_le_bytes());
    gif.extend_from_slice(&4u16.to_le_bytes());
    gif.extend_from_slice(&[0, 0, 0]);
    gif
}

fn with_pictures() -> Package {
    let mut edit = source().edit().unwrap();
    edit.insert_picture_at(&paragraph(0), None, &NewPicture::new(png(2, 1)))
        .unwrap();
    edit.insert_picture_at(
        &paragraph(0),
        None,
        &NewPicture::new(png(4, 4)).with_name("Second"),
    )
    .unwrap();
    edit.commit().unwrap().into_package()
}

#[test]
fn pictures_are_resized_renamed_and_described_in_place() {
    let base = with_pictures();
    let mut edit = base.edit().unwrap();
    let change = PictureChange::new()
        .with_size(914_400, 457_200)
        .with_name("Logo & mark")
        .with_description("A <logo>");
    assert!(
        edit.set_picture_at(&paragraph(0), Position::new(1), &change)
            .unwrap()
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let images = reopened.document().unwrap().paragraphs().unwrap()[0]
        .images()
        .unwrap();
    assert_eq!(images.len(), 2);
    assert_eq!(
        (images[1].width_emu(), images[1].height_emu()),
        (914_400, 457_200)
    );
    assert_eq!(images[1].name(), "Logo & mark");
    assert_eq!(images[1].description(), "A <logo>");
    // The first picture is untouched, and the shape's own extent follows.
    assert_eq!(images[0].width_emu(), 19_050);
    let xml = main_xml(&reopened);
    assert!(xml.contains(r#"<a:ext cx="914400" cy="457200"/>"#), "{xml}");
    // Replay and exact inverse.
    let replayed = base.apply(&patch).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
    // The same values again change nothing.
    let mut edit = changed.edit().unwrap();
    assert!(
        !edit
            .set_picture_at(&paragraph(0), Position::new(1), &change)
            .unwrap()
    );
    assert!(edit.commit().unwrap().into_parts().1.is_noop());
}

#[test]
fn a_picture_image_is_replaced_by_a_new_media_part() {
    let base = with_pictures();
    let before = base.document().unwrap().paragraphs().unwrap()[0]
        .images()
        .unwrap()[0]
        .r_embed()
        .to_owned();
    let parts = base.opc_package().iter_parts().count();
    let mut edit = base.edit().unwrap();
    assert!(
        edit.set_picture_at(
            &paragraph(0),
            Position::new(0),
            &PictureChange::new().with_image(gif())
        )
        .unwrap()
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let images = reopened.document().unwrap().paragraphs().unwrap()[0]
        .images()
        .unwrap();
    assert_ne!(images[0].r_embed(), before);
    let main = reopened.opc_package().main_document_part().unwrap();
    let target = main
        .rels()
        .get(images[0].r_embed())
        .unwrap()
        .target_partname()
        .unwrap();
    assert!(target.as_str().ends_with(".gif"), "{target}");
    assert_eq!(
        reopened.opc_package().get_part(&target).unwrap().blob(),
        gif().as_slice()
    );
    // The previous image stays, unreferenced; nothing else changed.
    assert_eq!(changed.opc_package().iter_parts().count(), parts + 1);
    assert_eq!(images[0].width_emu(), 19_050);
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
    // Refusals stage nothing.
    let mut edit = base.edit().unwrap();
    assert!(
        edit.set_picture_at(
            &paragraph(0),
            Position::new(0),
            &PictureChange::new().with_image(b"not an image".to_vec())
        )
        .is_err()
    );
    assert!(
        edit.set_picture_at(
            &paragraph(0),
            Position::new(0),
            &PictureChange::new().with_size(0, 5)
        )
        .is_err()
    );
    assert!(
        edit.set_picture_at(
            &paragraph(0),
            Position::new(5),
            &PictureChange::new().with_image(png(1, 1))
        )
        .is_err()
    );
    assert_eq!(base.opc_package().iter_parts().count(), parts);
    assert!(edit.commit().unwrap().into_parts().1.is_noop());
}

#[test]
fn a_removed_picture_takes_its_run_and_inverts_exactly() {
    let base = with_pictures();
    let mut edit = base.edit().unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let images = changed.document().unwrap().paragraphs().unwrap()[0]
        .images()
        .unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].name(), "Second");
    assert_eq!(changed.document().unwrap().text().unwrap(), "See Second");
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
    let mut edit = changed.edit().unwrap();
    assert!(
        edit.remove_picture_at(&paragraph(0), Position::new(1))
            .is_err()
    );
}

#[test]
fn comments_take_new_authors_initials_dates_and_text() {
    let mut edit = source().edit().unwrap();
    let id = edit
        .add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Ada", "Draft").with_initials("AL"),
        )
        .unwrap();
    let base = edit.commit().unwrap().into_package();
    let mut edit = base.edit().unwrap();
    let change = CommentChange::new()
        .with_author("Grace <H>")
        .with_initials("GH")
        .with_date("2026-10-17T00:00:00Z")
        .with_text("Final");
    assert!(edit.set_comment(id, &change).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let comments = reopened.document().unwrap().comments().unwrap();
    assert_eq!(comments[0].author(), "Grace <H>");
    assert_eq!(comments[0].initials(), Some("GH"));
    assert_eq!(comments[0].date(), Some("2026-10-17T00:00:00Z"));
    assert_eq!(comments[0].text().unwrap(), "Final");
    // The annotation mark stays; the anchors in the main story too.
    let name = litchi_opc::PackURI::new("/word/comments.xml".to_owned()).unwrap();
    let part = String::from_utf8(
        reopened
            .opc_package()
            .get_part(&name)
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    assert!(part.contains("annotationRef"), "{part}");
    assert!(main_xml(&reopened).contains("commentReference"));
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
    // The same values change nothing; refusals stage nothing.
    let mut edit = changed.edit().unwrap();
    assert!(!edit.set_comment(id, &change).unwrap());
    assert!(
        edit.set_comment(id, &CommentChange::new().with_author(""))
            .is_err()
    );
    assert!(
        edit.set_comment(id + 7, &CommentChange::new().with_text("x"))
            .is_err()
    );
    assert!(edit.commit().unwrap().into_parts().1.is_noop());
}

fn with_smartart() -> (Package, String) {
    let mut edit = source().edit().unwrap();
    let spec = Spec::parse_outline(Layout::Process, "Plan\nBuild", &Limits::default()).unwrap();
    let id = edit
        .insert_smartart_at(&paragraph(1), None, &NewSmartArt::new(spec))
        .unwrap();
    (edit.commit().unwrap().into_package(), id)
}

#[test]
fn smartart_parts_are_replaced_and_their_drawing_kept_consistent() {
    let (base, id) = with_smartart();
    // The inventory reads the drawing's extent (the insertion default).
    assert_eq!(
        base.smart_arts().unwrap()[0].extent,
        Some((5_760_000, 3_240_000))
    );
    let mut edit = base.edit().unwrap();
    let parts = edit.smartart_parts(&id).unwrap().unwrap();
    assert!(parts.drawing.is_some());
    // Identical parts are a no-op.
    assert!(!edit.set_smartart(&id, parts.clone()).unwrap());
    // A new outline with its regenerated drawing.
    let limits = Limits::default();
    let spec = Spec::parse_outline(Layout::Process, "Plan\nBuild\nShip", &limits).unwrap();
    let mut next = parts.clone();
    next.data = spec.data_xml(Some("rId0"), &limits).unwrap().into_bytes();
    next.drawing = Some(
        spec.drawing_xml(5_760_000, 3_240_000, &limits)
            .unwrap()
            .into_bytes(),
    );
    assert!(edit.set_smartart(&id, next).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let diagrams = reopened.smart_arts().unwrap();
    assert_eq!(diagrams[0].text(), "Plan\nBuild\nShip");
    assert!(diagrams[0].drawing_part_name.is_some());
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());

    // Without a drawing, the drawing part and its relationship go, and the
    // data model no longer names it.
    let mut edit = base.edit().unwrap();
    let mut bare = parts.clone();
    bare.drawing = None;
    assert!(edit.set_smartart(&id, bare).unwrap());
    let (undrawn, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&undrawn.to_bytes().unwrap()).unwrap();
    let diagrams = reopened.smart_arts().unwrap();
    assert_eq!(diagrams[0].drawing_part_name, None);
    assert!(
        reopened
            .opc_package()
            .iter_parts()
            .all(|part| !part.partname().as_str().contains("diagrams/drawing"))
    );
    let restored = undrawn.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());

    // A drawing comes back as a new part and relationship.
    let mut edit = reopened.edit().unwrap();
    let mut drawn = edit.smartart_parts(&id).unwrap().unwrap();
    assert!(drawn.drawing.is_none());
    drawn.drawing = parts.drawing.clone();
    assert!(edit.set_smartart(&id, drawn).unwrap());
    let redrawn = edit.commit().unwrap().into_package();
    assert!(redrawn.smart_arts().unwrap()[0].drawing_part_name.is_some());

    // A data model that does not read, or an unknown graphic, is refused.
    let mut edit = base.edit().unwrap();
    let mut broken = parts.clone();
    broken.data = b"<nope/>".to_vec();
    assert!(edit.set_smartart(&id, broken).is_err());
    assert!(edit.set_smartart("rId999", parts).is_err());
}

#[test]
fn a_removed_smartart_takes_its_parts_and_relationships() {
    let (base, id) = with_smartart();
    let parts = base.opc_package().iter_parts().count();
    let relationships = base
        .opc_package()
        .main_document_part()
        .unwrap()
        .rels()
        .len();
    let mut edit = base.edit().unwrap();
    edit.remove_smartart(&id).unwrap();
    let (removed, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&removed.to_bytes().unwrap()).unwrap();
    assert!(reopened.smart_arts().unwrap().is_empty());
    assert_eq!(reopened.opc_package().iter_parts().count(), parts - 5);
    assert_eq!(
        reopened
            .opc_package()
            .main_document_part()
            .unwrap()
            .rels()
            .len(),
        relationships - 5
    );
    assert!(!main_xml(&reopened).contains("relIds"));
    assert_eq!(reopened.document().unwrap().text().unwrap(), "See Second");
    let replayed = base.apply(&patch).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), removed.to_bytes().unwrap());
    let restored = removed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
    let mut edit = removed.edit().unwrap();
    assert!(edit.remove_smartart(&id).is_err());
}

#[test]
fn a_note_text_set_keeps_its_reference_mark() {
    use litchi_docx::document::{NewNote, NoteKind};
    let mut edit = source().edit().unwrap();
    let id = edit
        .add_note_at(
            NoteKind::Footnote,
            &paragraph(0),
            Some(Position::new(0)),
            &NewNote::new("One"),
        )
        .unwrap();
    let base = edit.commit().unwrap().into_package();
    let mut edit = base.edit().unwrap();
    edit.set_note_text(NoteKind::Footnote, id, "Uno").unwrap();
    let written = edit.commit().unwrap().into_package();
    let name = litchi_opc::PackURI::new("/word/footnotes.xml".to_owned()).unwrap();
    let part = String::from_utf8(
        written
            .opc_package()
            .get_part(&name)
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    assert!(part.contains("footnoteRef"), "{part}");
    assert!(part.contains("Uno"), "{part}");
}
