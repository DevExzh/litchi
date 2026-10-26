//! Page-number attachments (`litchi_pages::number_attachment`, record 0371).

use litchi_pages::number_attachment::{
    Target, TextNumberAttachmentKind, TextNumberAttachmentSettings,
};
use litchi_pages::{HeaderFooterSelector, Kind, Package, Template, TextPosition};

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn body_and_header_insert_update_remove_round_trip() {
    let package = Package::from_bytes(BASIC).unwrap();
    for target in [
        Target::Body,
        Target::HeaderFooter(HeaderFooterSelector::index(
            0,
            Template::First,
            Kind::Footer,
            1,
        )),
    ] {
        let mut edit = package.edit_number_attachments(target).unwrap();
        edit.insert(
            TextPosition::from_utf16_index(0).unwrap(),
            TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
        let inserted = edit.commit().unwrap();
        let attachments = inserted.package().number_attachments(target).unwrap();
        assert_eq!(attachments.len(), 1);
        assert_eq!(attachments[0].position().utf16_index(), 0);
        let restored = inserted
            .package()
            .apply_number_attachments(&inserted.patch().inverse())
            .unwrap();
        assert_eq!(bytes(restored.package()), BASIC);

        let mut edit = inserted.package().edit_number_attachments(target).unwrap();
        edit.update(
            0,
            TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount),
        )
        .unwrap();
        let updated = edit.commit().unwrap();
        assert_eq!(
            updated.package().number_attachments(target).unwrap()[0]
                .settings()
                .kind,
            TextNumberAttachmentKind::PageCount
        );

        let mut edit = updated.package().edit_number_attachments(target).unwrap();
        edit.remove(0).unwrap();
        let removed = edit.commit().unwrap();
        assert!(
            removed
                .package()
                .number_attachments(target)
                .unwrap()
                .is_empty()
        );
    }
}
