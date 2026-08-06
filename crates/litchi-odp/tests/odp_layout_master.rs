use litchi_odp::{
    MasterPage, Presentation, constants,
    layout::{Layout, Placeholder, Role},
};

fn package() -> Vec<u8> {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide1"/></office:presentation></office:body></office:document-content>"#;
    let styles = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:styles/><office:automatic-styles><style:page-layout style:name="physical"/><style:style style:name="page-style" style:family="drawing-page"/></office:automatic-styles><office:master-styles/></office:document-styles>"#;

    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype(constants::ODF_PRESENTATION).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer.finish().unwrap()
}

#[test]
fn master_pages_are_exported_and_commit_lossless_snapshot_edits() {
    let mut presentation = Presentation::from_bytes(package()).unwrap();
    let mut layout = Layout::new("layout-a").unwrap();
    layout.placeholders.push(Placeholder::new(
        Role::Title,
        "1cm".parse().unwrap(),
        "1cm".parse().unwrap(),
        "10cm".parse().unwrap(),
        "2cm".parse().unwrap(),
    ));
    presentation.add_layout(&layout).unwrap();

    let mut master = MasterPage::new("master-a", "physical").unwrap();
    master.master_page.drawing_style_name = Some("page-style".to_string());
    master.page_layout_name = Some("layout-a".to_string());
    master.master_page.xml = master.master_page.xml.replace(
        "/>",
        "><draw:rect xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"/></style:master-page>",
    );
    presentation.add_master_page(&master).unwrap();
    presentation
        .assign_slide_master_page(0, Some("master-a"))
        .unwrap();
    presentation
        .assign_slide_page_layout(0, Some("layout-a"))
        .unwrap();

    let pages = presentation.master_pages().unwrap();
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].name(), "master-a");
    assert_eq!(pages[0].page_layout_name.as_deref(), Some("layout-a"));
    assert!(presentation.content_xml().contains("master-a"));
}
