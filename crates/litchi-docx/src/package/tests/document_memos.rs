//! The owning package's memos of the visible main document and its body
//! block index.

use super::*;
use crate::parts::memo::MEMO_ENTRY_OVERHEAD_BYTES;
use litchi_core::Position;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn package_with(document: String) -> Package {
    let mut opc = OpcPackage::new();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document.into_bytes(),
    )))
    .unwrap();
    opc.relate_to(
        "word/document.xml",
        litchi_opc::constants::relationship_type::OFFICE_DOCUMENT,
    );
    let bytes = litchi_opc::PackageWriter::to_bytes(&opc).unwrap();
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

fn word_document(paragraphs: &[&str]) -> String {
    let body: String = paragraphs
        .iter()
        .map(|text| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"))
        .collect();
    format!(r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}"><w:body>{body}</w:body></w:document>"#)
}

#[test]
fn views_of_an_unchanged_part_share_one_projection_and_one_body_index() {
    let package = package_with(word_document(&["alpha", "beta"]));
    let first = package.document().unwrap();
    let second = package.document().unwrap();
    assert_eq!(
        first.part().xml_bytes().as_ptr(),
        second.part().xml_bytes().as_ptr(),
        "the second view reuses the first projection"
    );
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!((visible.builds, visible.hits), (1, 1));
    assert_eq!(visible.pinned_entries, 1, "live views pin the projection");

    let first_body = first.body_blocks().unwrap();
    let second_body = second.body_blocks().unwrap();
    assert!(first_body.shares_index_with(&second_body));
    assert_eq!(package.paragraph_index_cache.body_diagnostics().builds, 1);
    assert_eq!(
        first
            .body_paragraph(Position::new(1))
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "beta"
    );
    assert_eq!(second.body_block_count().unwrap(), 2);
    assert!(second.body_table(Position::new(0)).unwrap().is_none());
    assert!(second.body_sdt(Position::new(0)).unwrap().is_none());
    assert!(second.body_block(Position::new(2)).unwrap().is_none());

    // A live view keeps the memos pinned through a release request.
    package.trim_document_memos();
    assert_eq!(
        package
            .paragraph_index_cache
            .visible_diagnostics()
            .pinned_entries,
        1
    );
    drop((first, second, first_body, second_body));
    package.trim_document_memos();
    let visible = package.paragraph_index_cache.visible_diagnostics();
    let body = package.paragraph_index_cache.body_diagnostics();
    assert_eq!((visible.clean_evictions, body.clean_evictions), (1, 1));
    assert_eq!(visible.clean_entries + body.clean_entries, 0);

    // A released memo is rebuilt with the same result.
    let rebuilt = package.document().unwrap();
    assert_eq!(rebuilt.text().unwrap(), "alphabeta");
    assert_eq!(
        package.paragraph_index_cache.visible_diagnostics().builds,
        2
    );
}

#[test]
fn a_published_edit_installs_a_new_generation_and_old_values_stay_pinned() {
    let package = package_with(word_document(&["alpha", "beta"]));
    let old_paragraph = package
        .document()
        .unwrap()
        .body_paragraph(Position::new(1))
        .unwrap()
        .unwrap();
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(Position::new(1), "changed")
        .unwrap();
    let package = edit.commit().unwrap().into_package();

    let after = package.document().unwrap();
    assert_eq!(after.text().unwrap(), "alphachanged");
    assert_eq!(
        after
            .body_paragraph(Position::new(1))
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "changed"
    );
    assert_eq!(old_paragraph.text().unwrap(), "beta");
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!((visible.builds, visible.hits), (2, 0));
    assert_eq!(package.paragraph_index_cache.body_diagnostics().builds, 2);
}

#[test]
fn a_part_without_markup_compatibility_is_read_without_a_copy() {
    let package = package_with(format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:body></w:document>"#
    ));
    let document = package.document().unwrap();
    let raw = package.opc.main_document_part().unwrap().blob_arc();
    assert_eq!(document.part().xml_bytes().as_ptr(), raw.as_ptr());
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!(visible.retained_weight, MEMO_ENTRY_OVERHEAD_BYTES);
    assert_eq!(visible.clean_entries, 1, "the verdict pins no payload");
    assert_eq!(
        document
            .body_paragraph(Position::new(0))
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "plain"
    );
}
