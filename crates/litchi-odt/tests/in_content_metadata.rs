use litchi_core::Position;
use litchi_odt::{
    Document, RdfaAttributes, TextMeta,
    core::PackageWriter,
    generic::{FlatDocument, Package},
    xforms::ModelChild,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XHTML: &str = "http://www.w3.org/1999/xhtml";
const XFORMS: &str = "http://www.w3.org/2002/xforms";

fn flat(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:h="{XHTML}" xmlns:xf="{XFORMS}" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document>"#
    )
    .into_bytes()
}

fn content(body: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:h="{XHTML}" xmlns:xf="{XFORMS}" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
    )
}

fn package(content_xml: &str, styles_xml: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file("content.xml", content_xml.as_bytes())
        .unwrap();
    if let Some(styles) = styles_xml {
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn body() -> &'static str {
    r##"<o:forms><xf:model id="m"><xf:instance id="i" xmlns:d="urn:example:data"><d:data><d:value>one</d:value></d:data></xf:instance><xf:bind id="b" nodeset="/data/value" type="xsd:string"/></xf:model></o:forms><t:p h:about="#doc" h:property="dc:title" h:content="A > B"><!--keep--><t:bookmark-start t:name="mark" h:about="#mark"/><t:meta xml:id="meta" h:property="dc:description">value</t:meta>visible</t:p>"##
}

#[test]
fn public_flat_facade_reads_and_preserves_in_content_metadata_and_xforms() {
    let source = flat(body());
    let mut document = FlatDocument::from_bytes(source).unwrap();
    let metadata = document.in_content_metadata().unwrap();
    assert_eq!(metadata.text_meta.len(), 1);
    assert!(
        metadata
            .rdfa
            .iter()
            .any(|item| matches!(item.host, litchi_odt::RdfaHost::Paragraph { index: 0 }))
    );
    let models = document.xforms_models().unwrap();
    assert_eq!(models.len(), 1);
    assert_eq!(
        models[0].instances().next().unwrap().id.as_deref(),
        Some("i")
    );
    assert!(
        models[0]
            .instances()
            .next()
            .unwrap()
            .content_xml
            .as_deref()
            .unwrap()
            .contains("d:data")
    );

    let value = RdfaAttributes {
        about: Some("#changed".to_string()),
        property: Some("dc:subject".to_string()),
        content: None,
        datatype: None,
    };
    document
        .set_paragraph_rdfa(Position::new(0), &value)
        .unwrap();
    let mut metadata = TextMeta::from_text("inserted").unwrap();
    metadata.rdfa.property = Some("dc:comment".to_string());
    document
        .insert_text_meta(Position::new(0), &metadata)
        .unwrap();
    let mut replacement = document.xforms_models().unwrap().remove(0);
    replacement.id = Some("changed-model".to_string());
    document
        .replace_xforms_model(Position::new(0), &replacement)
        .unwrap();

    let updated = document.xml().to_owned();
    assert!(updated.contains("<!--keep-->"));
    assert!(updated.contains("xhtml:about=\"#changed\""));
    assert!(updated.contains("<text:meta") || updated.contains("<t:meta"));
    assert!(updated.contains("id=\"changed-model\""));
    let reparsed = FlatDocument::from_bytes(updated.into_bytes()).unwrap();
    assert_eq!(reparsed.in_content_metadata().unwrap().text_meta.len(), 2);
    assert_eq!(
        reparsed.xforms_models().unwrap()[0].id.as_deref(),
        Some("changed-model")
    );
}

#[test]
fn document_and_package_facades_cover_styles_and_mutable_public_edits() {
    let styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:h="{XHTML}" o:version="1.3"><o:styles><t:meta h:property="dc:style">style value</t:meta></o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>"#
    );
    let bytes = package(&content(body()), Some(&styles));
    let document = Document::from_bytes(bytes.clone()).unwrap();
    let metadata = document.in_content_metadata().unwrap();
    assert!(
        metadata
            .text_meta
            .iter()
            .any(|item| item.part == litchi_odt::MetadataPart::Styles)
    );
    assert_eq!(document.xforms_models().unwrap().len(), 1);

    let mut generic = Package::from_bytes(bytes).unwrap();
    let original_package = generic.to_bytes();
    generic
        .set_paragraph_rdfa(
            Position::new(0),
            &RdfaAttributes {
                about: Some("#doc".to_string()),
                property: Some("dc:title".to_string()),
                content: Some("A > B".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .unwrap();
    assert_eq!(generic.to_bytes(), original_package);
    generic
        .set_paragraph_rdfa(
            Position::new(0),
            &RdfaAttributes {
                about: Some("#package".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .unwrap();
    let mut model = generic.xforms_models().unwrap().remove(0);
    model
        .children
        .retain(|child| !matches!(child, ModelChild::Bind(_)));
    generic
        .replace_xforms_model(Position::new(0), &model)
        .unwrap();
    let reopened = Package::from_bytes(generic.to_bytes()).unwrap();
    assert_eq!(reopened.xforms_models().unwrap()[0].binds().count(), 0);
    assert!(
        reopened
            .in_content_metadata()
            .unwrap()
            .rdfa
            .iter()
            .any(|item| matches!(item.host, litchi_odt::RdfaHost::Paragraph { index: 0 }))
    );

    let source = Document::from_bytes(package(&content(body()), None)).unwrap();
    let mut mutable = litchi_odt::mutable::MutableDocument::from_document(source).unwrap();
    mutable
        .set_bookmark_rdfa(
            "mark",
            &RdfaAttributes {
                property: Some("dc:identifier".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .unwrap();
    mutable.remove_text_meta(Position::new(0)).unwrap();
    mutable.remove_xforms_model(Position::new(0)).unwrap();
    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(reopened.in_content_metadata().unwrap().text_meta.is_empty());
    assert!(reopened.xforms_models().unwrap().is_empty());
}
