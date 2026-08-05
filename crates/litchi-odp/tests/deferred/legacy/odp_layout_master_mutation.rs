use litchi_odp::{Layout, MasterPage, Measure, Placeholder, Presentation, Role, constants};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive};

fn measure(value: &str) -> Measure {
    value.parse().unwrap()
}

fn layout(name: &str, width: &str) -> Layout {
    Layout {
        name: name.to_string(),
        display_name: Some(format!("{name} display")),
        placeholders: vec![Placeholder {
            role: Role::Title,
            x: measure("1cm"),
            y: measure("1cm"),
            width: measure(width),
            height: measure("2cm"),
        }],
    }
}

#[test]
fn packaged_layout_master_roundtrip_reassignment_reorder_and_unknown_xml() {
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
    first.master_page.xml = first.master_page.xml.replace("/>", "><draw:rect xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" draw:name=\"kept-master-shape\"/></style:master-page>");
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
    let content = zip_text(&bytes, "content.xml");
    let styles = zip_text(&bytes, "styles.xml");
    let manifest = zip_text(&bytes, "META-INF/manifest.xml");
    assert!(content.contains("foreign:keep=\"content\""));
    assert!(styles.contains("foreign:keep=\"styles\""));
    assert!(content.contains("master-b"));
    assert!(content.contains("layout-b"));
    assert!(!content.contains("master-a"));
    assert!(!content.contains("layout-a"));
    assert!(manifest.contains("styles.xml"));
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
fn malformed_geometry_active_content_missing_refs_and_bad_reorder_are_atomic() {
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
    scripted.master_page.xml = scripted.master_page.xml.replace("/>", "><script:event-listener xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\"/></style:master-page>");
    assert!(presentation.add_master_page(&scripted).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    assert!(
        presentation
            .reorder_layouts(&["unknown".to_string()])
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

fn host_package() -> Vec<u8> {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:foreign="urn:example:foreign"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide1" foreign:keep="content"><draw:rect draw:name="untouched"/></draw:page></office:presentation></office:body></office:document-content>"#;
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:foreign="urn:example:foreign" foreign:keep="styles"><office:styles/><office:automatic-styles><style:page-layout style:name="physical"/><style:style style:name="page-style" style:family="drawing-page"/></office:automatic-styles><office:master-styles/></office:document-styles>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(constants::ODF_PRESENTATION.as_bytes())
        .unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("styles.xml", deflated).unwrap();
    zip.write_all(styles.as_bytes()).unwrap();
    let manifest = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        constants::ODF_PRESENTATION
    );
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

fn zip_text(bytes: &[u8], path: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut file = archive.by_name(path).unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    text
}
