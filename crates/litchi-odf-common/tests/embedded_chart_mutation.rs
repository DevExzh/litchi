use litchi_odf_common::{
    chart::authoring::{Definition, Text, serialize_content},
    constants,
    core::{OwnedPackage, PackageWriter},
    embedded::{Root, Source, scan_package},
    package::edit::{Addition, rebuild_package, splice},
};

const OBJECT: &str = r#"<draw:object xlink:href="./Object_1"/>"#;
const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p>host</text:p><draw:frame draw:name="Chart">OBJECT</draw:frame></office:text></office:body></office:document-content>"#;

fn chart(title: &str) -> Definition {
    let mut definition = Definition::new("chart:bar");
    definition.title = Some(Text::new(title));
    definition
}

fn package(title: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(
            constants::ODF_CONTENT,
            CONTENT.replace("OBJECT", OBJECT).as_bytes(),
        )
        .unwrap();
    writer
        .add_manifest_directory("Object_1/", constants::ODF_CHART)
        .unwrap();
    writer
        .add_file_with_media_type(
            "Object_1/content.xml",
            serialize_content(&chart(title)).unwrap().as_bytes(),
            "text/xml",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn packaged_chart_replacement_and_removal_use_common_package_primitives() {
    let source = OwnedPackage::from_bytes(package("First")).unwrap();
    let content = String::from_utf8(source.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    let lookup = source.package().unwrap();
    let objects = scan_package(&content, None, &lookup).unwrap();
    assert_eq!(objects.len(), 1);
    assert!(matches!(
        &objects[0].source,
        Source::PackageSubdocument {
            root_path,
            content_path,
            manifest_media_type: Some(media_type),
            ..
        } if root_path == "Object_1/"
            && content_path == "Object_1/content.xml"
            && media_type == constants::ODF_CHART
    ));

    let replacement = serialize_content(&chart("Second")).unwrap();
    let replaced = rebuild_package(
        &source,
        &content,
        vec![Addition {
            path: "Object_1/content.xml".to_string(),
            bytes: replacement.as_bytes().to_vec(),
            media_type: "text/xml".to_string(),
        }],
        vec![("Object_1/".to_string(), constants::ODF_CHART.to_string())],
        Vec::new(),
        vec!["Object_1/".to_string()],
    )
    .unwrap();
    let replaced = OwnedPackage::from_bytes(replaced).unwrap();
    assert_eq!(
        replaced.get_file("Object_1/content.xml").unwrap(),
        replacement.as_bytes()
    );

    let start = content.find(OBJECT).unwrap();
    let removed_content = splice(&content, start, start + OBJECT.len(), "").unwrap();
    let removed = rebuild_package(
        &replaced,
        &removed_content,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        vec!["Object_1/".to_string()],
    )
    .unwrap();
    let removed = OwnedPackage::from_bytes(removed).unwrap();
    assert!(!removed.has_file("Object_1/content.xml").unwrap());
    let removed_xml = String::from_utf8(removed.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    assert!(
        scan_package(&removed_xml, None, &removed.package().unwrap())
            .unwrap()
            .is_empty()
    );
}

#[test]
fn inline_chart_is_inert_and_invalid_mutations_do_not_publish_bytes() {
    let inline = format!(
        r#"<draw:object><office:document office:mimetype="{}">{}</office:document></draw:object>"#,
        constants::ODF_CHART,
        serialize_content(&chart("Inline")).unwrap()
    );
    let flat = format!(
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">{inline}</office:document>"#
    );
    let objects = litchi_odf_common::embedded::scan_flat(&flat).unwrap();
    assert!(matches!(
        &objects[0].source,
        Source::InlineXml {
            root: Root::OpenDocument,
            ..
        }
    ));

    let source = package("Stable");
    let mut invalid = chart("Rejected");
    invalid.class = "not a qualified name".to_string();
    assert!(serialize_content(&invalid).is_err());
    assert!(splice(&flat, flat.len() + 1, flat.len() + 1, "").is_err());
    assert_eq!(source, package("Stable"));
}
