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

/// Change 0653: the markup-compatibility writer stopped re-declaring every
/// in-scope namespace on every element it emits, so a `w:p` span sliced out of
/// the processed `word/document.xml` no longer carries `xmlns:w` by accident.
/// `Paragraph::extensions` re-declares what the span inherits before it parses
/// it, so a real marker-bearing document still reads its `w14:paraId`.
#[test]
fn extensions_read_word_2010_ids_from_a_real_marker_bearing_document() {
    let path = "../../test-data/ooxml/docx/table-alignment.docx";
    let package = Package::open(path).unwrap();
    let document = package.document().unwrap();

    let mut seen = 0usize;
    for index in 0..document.paragraph_count().unwrap() {
        let Some(paragraph) = document.paragraph(index).unwrap() else {
            continue;
        };
        let extensions = paragraph.extensions().unwrap();
        if extensions.ids().para_id().is_some() {
            seen += 1;
        }
    }
    assert!(
        seen > 0,
        "no paragraph of a Word-authored document reported a w14:paraId"
    );

    let mut rows = 0usize;
    for table in document.tables().unwrap() {
        for row in table.rows().unwrap() {
            if row.extension_ids().unwrap().para_id().is_some() {
                rows += 1;
            }
        }
    }
    assert!(rows > 0, "no table row reported a w14:paraId");
}
