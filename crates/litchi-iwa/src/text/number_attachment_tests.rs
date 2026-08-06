use super::{
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText, TextPosition,
};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::comment::DrawableId;

const PREFIX: &str = "Page ";
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn settings(kind: TextNumberAttachmentKind) -> TextNumberAttachmentSettings {
    TextNumberAttachmentSettings::new(kind)
}

#[test]
fn scratch_pages_insert_update_remove_is_byte_exact() {
    let mut pages = PagesEditor::create_with_text(PREFIX).unwrap();
    let baseline = pages.to_bytes().unwrap();
    let position = TextPosition::from_utf16_index(PREFIX.encode_utf16().count()).unwrap();
    let created = pages
        .insert_body_number_attachment(position, settings(TextNumberAttachmentKind::PageNumber))
        .unwrap();
    assert_eq!(pages.body_text().unwrap(), format!("{PREFIX}\u{fffc}"));
    assert_eq!(
        pages.body_number_attachments().unwrap().as_slice(),
        std::slice::from_ref(&created)
    );

    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    let updated_settings = settings(TextNumberAttachmentKind::PageCount)
        .with_string_equivalent(TextNumberAttachmentText::new("").unwrap());
    let updated = pages
        .update_body_number_attachment(created.id, updated_settings.clone())
        .unwrap();
    assert_eq!(updated.id, created.id);
    assert_eq!(updated.position, created.position);
    assert_eq!(updated.settings, updated_settings);
    assert_eq!(
        pages.remove_body_number_attachment(created.id).unwrap(),
        updated
    );
    assert!(pages.body_number_attachments().unwrap().is_empty());
    assert_eq!(pages.body_text().unwrap(), PREFIX);
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn invalid_positions_and_text_replacement_are_transactional() {
    let mut pages = PagesEditor::create_with_text("A😀B").unwrap();
    let baseline = pages.to_bytes().unwrap();
    assert!(
        pages
            .insert_body_number_attachment(
                TextPosition::from_utf16_index(2).unwrap(),
                settings(TextNumberAttachmentKind::PageNumber),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);

    let attachment = pages
        .insert_body_number_attachment(
            TextPosition::from_utf16_index(1).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    pages.replace_body_text(1..2, "").unwrap();
    assert!(pages.body_number_attachments().unwrap().is_empty());
    assert!(!pages.package().iwa_entry_names().any(|name| {
        pages
            .package()
            .archive(name)
            .unwrap()
            .object(attachment.id.object_id())
            .is_some()
    }));
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn pages_header_footer_ownership_is_enforced() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let header = pages.header_footers().unwrap()[0].storage.object_id;
    pages.set_header_footer_text(header, PREFIX).unwrap();
    let attachment = pages
        .insert_header_footer_number_attachment(
            header,
            TextPosition::from_utf16_index(PREFIX.len()).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    assert_eq!(
        pages
            .header_footer_number_attachments(header)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&attachment)
    );
    let body_id = pages.body_storage().unwrap().object_id;
    assert!(pages.header_footer_number_attachments(body_id).is_err());
    pages
        .remove_header_footer_number_attachment(header, attachment.id)
        .unwrap();
    assert_eq!(
        pages
            .header_footers()
            .unwrap()
            .into_iter()
            .find(|item| item.storage.object_id == header)
            .unwrap()
            .storage
            .storage
            .text(),
        PREFIX
    );
}

#[test]
fn scratch_pages_text_box_number_attachments_round_trip() {
    let position = TextPosition::from_utf16_index(PREFIX.encode_utf16().count()).unwrap();

    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let body_attachment = pages
        .insert_body_number_attachment(
            TextPosition::from_utf16_index(4).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let pages_box = pages.add_text_box(5, PREFIX, POSITION, SIZE).unwrap();
    let pages_box_id = DrawableId::from_raw(pages_box.drawable_object_id).unwrap();
    let pages_attachment = pages
        .insert_text_box_number_attachment(
            pages_box_id,
            position,
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages.body_number_attachments().unwrap().as_slice(),
        std::slice::from_ref(&body_attachment)
    );
    assert_eq!(
        pages
            .text_box_number_attachments(pages_box_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&pages_attachment)
    );
    pages
        .remove_text_box_number_attachment(pages_box_id, pages_attachment.id)
        .unwrap();
}
