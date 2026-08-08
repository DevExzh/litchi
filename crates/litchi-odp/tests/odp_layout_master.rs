use litchi_odp::{
    MasterPage, Presentation, constants,
    core::{OwnedPackage, PackageWriter},
    layout::{Layout, Measure, Placeholder, Role},
};

fn package() -> Vec<u8> {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide1"/></office:presentation></office:body></office:document-content>"#;
    let styles = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:styles/><office:automatic-styles><style:page-layout style:name="physical"/><style:style style:name="page-style" style:family="drawing-page"/></office:automatic-styles><office:master-styles/></office:document-styles>"#;

    let mut writer = PackageWriter::new();
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

fn layout(name: &str, width: &str) -> Layout {
    let mut layout = Layout::new(name).unwrap();
    layout.display_name = Some(format!("{name} display"));
    layout.placeholders.push(Placeholder::new(
        Role::Title,
        "1cm".parse().unwrap(),
        "1cm".parse().unwrap(),
        width.parse().unwrap(),
        "2cm".parse().unwrap(),
    ));
    layout
}

fn host_package() -> Vec<u8> {
    const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:foreign="urn:example:foreign"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide1" foreign:keep="content"><draw:rect draw:name="untouched"/></draw:page></office:presentation></office:body></office:document-content>"#;
    const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:foreign="urn:example:foreign" foreign:keep="styles"><office:styles/><office:automatic-styles><style:page-layout style:name="physical"/><style:style style:name="page-style" style:family="drawing-page"/></office:automatic-styles><office:master-styles/></office:document-styles>"#;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_PRESENTATION).unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    writer.finish().unwrap()
}

#[test]
fn reorder_remove_reassign_and_unknown_xml_round_trip() {
    let mut presentation = Presentation::from_bytes(host_package()).unwrap();
    presentation
        .add_layout(&layout("layout-a", "20cm"))
        .unwrap();
    presentation
        .add_layout(&layout("layout-b", "18cm"))
        .unwrap();
    let mut first = MasterPage::new("master-a", "physical").unwrap();
    first.master_page.drawing_style_name = Some("page-style".to_string());
    first.page_layout_name = Some("layout-a".to_string());
    first.master_page.xml = first.master_page.xml.replace(
        "/>",
        "><draw:rect xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" draw:name=\"kept-master-shape\"/></style:master-page>",
    );
    presentation.add_master_page(&first).unwrap();
    presentation
        .add_master_page(&MasterPage::new("master-b", "physical").unwrap())
        .unwrap();
    presentation
        .assign_slide_master_page(0, Some("master-a"))
        .unwrap();
    presentation
        .assign_slide_page_layout(0, Some("layout-a"))
        .unwrap();
    presentation
        .reorder_layouts(&["layout-b".to_string(), "layout-a".to_string()])
        .unwrap();
    presentation
        .reorder_master_pages(&["master-b".to_string(), "master-a".to_string()])
        .unwrap();
    presentation
        .remove_page_layout("layout-a", Some("layout-b"))
        .unwrap();
    presentation
        .remove_master_page("master-a", Some("master-b"))
        .unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    let styles = String::from_utf8(package.get_file("styles.xml").unwrap()).unwrap();
    let manifest = String::from_utf8(package.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    assert!(content.contains("foreign:keep=\"content\""));
    assert!(styles.contains("foreign:keep=\"styles\""));
    assert!(content.contains("master-b"));
    assert!(content.contains("layout-b"));
    assert!(!content.contains("master-a"));
    assert!(!content.contains("layout-a"));
    assert!(manifest.contains("styles.xml"));
    for xml in [&content, &styles, &manifest] {
        assert!(!xml.contains('\n'));
        assert!(!xml.contains("> <"));
    }
    assert_eq!(
        Presentation::from_bytes(bytes)
            .unwrap()
            .master_pages()
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn invalid_layout_master_edits_are_atomic() {
    let mut presentation = Presentation::from_bytes(host_package()).unwrap();
    presentation.add_layout(&layout("signed", "-1cm")).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!("NaNcm".parse::<Measure>().is_err());
    assert!("infinity%".parse::<Measure>().is_err());
    assert!(
        format!("{}cm", "9".repeat(65_537))
            .parse::<Measure>()
            .is_err()
    );

    let mut oversized = layout("valid", "1cm");
    oversized.name = "x".repeat(4_097);
    assert!(presentation.add_layout(&oversized).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    assert!(
        presentation
            .assign_slide_master_page(0, Some("missing"))
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let mut scripted = MasterPage::new("scripted", "physical").unwrap();
    scripted.master_page.xml = scripted.master_page.xml.replace(
        "/>",
        "><script:event-listener xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\"/></style:master-page>",
    );
    assert!(presentation.add_master_page(&scripted).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    assert!(
        presentation
            .reorder_layouts(&["unknown".to_string()])
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
}
