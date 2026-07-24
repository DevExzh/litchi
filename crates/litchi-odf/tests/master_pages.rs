use litchi_odf::{
    FlatOpenDocument, MasterPage, MasterPageChild, MasterPageChildKind, OpenDocumentPackage,
    insert_master_page_xml, remove_master_page_xml, replace_master_page_xml,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

fn styles(body: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><o:automatic-styles/><o:master-styles>{body}</o:master-styles></o:document-styles>"#
    )
}

#[test]
fn parses_and_classifies_exact_rng_child_order_inertly() {
    let body = r#"<s:master-page s:name="M" s:page-layout-name="pm1"><s:header/><s:header-left/><s:header-first/><s:footer/><s:footer-left/><s:footer-first/><d:layer-set/><o:forms/><d:rect/><d:frame/><a:par/><p:notes/></s:master-page>"#;
    let flat = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:master-styles>{body}</o:master-styles><o:body><o:text/></o:body></o:document>"#
    );
    let document = FlatOpenDocument::from_bytes(flat.into_bytes()).unwrap();
    let page = &document.master_pages().unwrap()[0];
    assert_eq!(page.page_layout_name.as_deref(), Some("pm1"));
    assert_eq!(page.children.len(), 12);
    assert_eq!(page.children[6].kind, MasterPageChildKind::LayerSet);
    assert_eq!(page.children[7].kind, MasterPageChildKind::Forms);
    assert_eq!(page.children[8].kind, MasterPageChildKind::Shape);
    assert_eq!(page.children[9].kind, MasterPageChildKind::Shape);
    assert_eq!(page.children[10].kind, MasterPageChildKind::Animation);
    assert_eq!(page.children[11].kind, MasterPageChildKind::Notes);
}

#[test]
fn canonical_insert_replace_remove_preserves_unrelated_bytes() {
    let original = styles(r#"<s:master-page s:name="Keep" s:page-layout-name="pm0"/>"#);
    let mut page = MasterPage::try_new("Added", "pm1").unwrap();
    page.display_name = Some("Added & Main".to_string());
    page.children.push(MasterPageChild::new(
        MasterPageChildKind::Shape,
        r#"<draw:rect xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/>"#,
    ));
    let fragment = page.to_xml_fragment().unwrap();
    let inserted = insert_master_page_xml(&original, &fragment).unwrap();
    assert!(inserted.contains(r#"s:name="Keep""#));
    assert!(inserted.contains("Added &amp; Main"));

    page.children.clear();
    let replaced =
        replace_master_page_xml(&inserted, "Added", &page.to_xml_fragment().unwrap()).unwrap();
    assert!(replaced.contains(r#"s:name="Keep""#));
    assert!(!replaced.contains("<draw:rect"));
    let removed = remove_master_page_xml(&replaced, "Added").unwrap();
    assert_eq!(removed, original);
}

#[test]
fn rejects_missing_attributes_wrong_order_duplicates_foreign_children_and_depth() {
    for body in [
        r#"<s:master-page s:name="M"/>"#,
        r#"<s:master-page s:page-layout-name="pm1"/>"#,
        r#"<s:master-page s:name="M" s:page-layout-name="pm1"><s:header-left/><s:header/></s:master-page>"#,
        r#"<s:master-page s:name="M" s:page-layout-name="pm1"><d:layer-set/><d:layer-set/></s:master-page>"#,
        r#"<s:master-page s:name="M" s:page-layout-name="pm1"><x:foreign xmlns:x="urn:example"/></s:master-page>"#,
    ] {
        let xml = styles(body);
        assert!(litchi_odf::FlatOpenDocument::from_bytes(format!(r#"<o:document xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:master-styles>{body}</o:master-styles><o:body><o:text/></o:body></o:document>"#).into_bytes()).unwrap().master_pages().is_err(), "accepted {xml}");
    }
    let deep = format!(
        "{}{}{}",
        r#"<s:master-page s:name="M" s:page-layout-name="pm1"><s:header>"#,
        "<d:g>".repeat(260),
        "</d:g>".repeat(260) + "</s:header></s:master-page>"
    );
    let flat = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:master-styles>{deep}</o:master-styles><o:body><o:text/></o:body></o:document>"#
    );
    assert!(
        FlatOpenDocument::from_bytes(flat.into_bytes())
            .unwrap()
            .master_pages()
            .is_err()
    );
}

#[test]
fn reads_real_conforming_libreoffice_package_through_generic_accessor() {
    let package = OpenDocumentPackage::open(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/libreoffice-core/writerperfect/qa/unit/data/writer/epubexport/simple-ruby.odt")).unwrap();
    let pages = package.master_pages().unwrap();
    assert!(!pages.is_empty());
    assert!(pages.iter().all(|page| page.page_layout_name.is_some()));
}
