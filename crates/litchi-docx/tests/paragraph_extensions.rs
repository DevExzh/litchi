use litchi_docx::Package;
use litchi_docx::paragraph::extensions::Id;
use tempfile::NamedTempFile;

#[test]
fn package_round_trip_preserves_extension_state_and_hyperlink_relationships() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let para_id = Id::new(0x1234).unwrap();
    let text_id = Id::new(0x5678).unwrap();

    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_hyperlink("https://example.test/typed", "typed link");
        paragraph
            .set_para_id(Some(para_id))
            .unwrap()
            .set_text_id(Some(text_id))
            .unwrap()
            .set_no_spell_err(Some(true));

        document
            .add_table(1, 1)
            .row(0)
            .unwrap()
            .set_para_id(Some(Id::new(0x9abc).unwrap()))
            .unwrap()
            .set_text_id(Some(Id::new(0xdef0).unwrap()))
            .unwrap();
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let paragraph = document.paragraph(0).unwrap().unwrap();
    let extensions = paragraph.extensions().unwrap();
    assert_eq!(extensions.ids().para_id(), Some(para_id));
    assert_eq!(extensions.ids().text_id(), Some(text_id));
    assert_eq!(extensions.no_spell_err(), Some(true));

    let main = reopened.opc_package().main_document_part().unwrap();
    let links = paragraph.hyperlinks(main.rels()).unwrap();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].url(), Some("https://example.test/typed"));

    let row = document
        .table(0)
        .unwrap()
        .unwrap()
        .rows()
        .unwrap()
        .into_iter()
        .next()
        .unwrap();
    let ids = row.extension_ids().unwrap();
    assert_eq!(ids.para_id(), Id::new(0x9abc));
    assert_eq!(ids.text_id(), Id::new(0xdef0));
}
