use litchi_odf_common::constants;
use litchi_ods::rdf::{Object, Subject, Triple};
use litchi_ods::{Builder, Spreadsheet};
use std::io::{Cursor, Write};

#[test]
fn builder_and_package_facade_round_trip() {
    let bytes = Builder::new()
        .build()
        .expect("test fixture or operation should succeed");
    let spreadsheet =
        Spreadsheet::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
    assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    assert_eq!(spreadsheet.into_bytes(), bytes);
}

#[test]
fn spreadsheet_facade_owns_rdf_crud() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:spreadsheet><table:table xml:id="sheet" table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#;
    let bytes = Builder::new()
        .content_xml(content)
        .build()
        .expect("test fixture or operation should succeed");
    let mut spreadsheet =
        Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
    let triple = Triple {
        subject: Subject::Iri("#sheet".to_string()),
        predicate: "https://example.invalid/schema#label".to_string(),
        object: Object::Literal {
            value: "Sheet".to_string(),
            datatype: None,
            language: None,
        },
    };
    let path = spreadsheet
        .add_rdf_graph(None, &[triple.clone()])
        .expect("test fixture or operation should succeed");
    assert_eq!(path, "Metadata/metadata_1.rdf");
    assert_eq!(
        spreadsheet
            .rdf_graphs()
            .expect("test fixture or operation should succeed")[0]
            .triples,
        [triple]
    );
    spreadsheet
        .remove_rdf_graph(&path)
        .expect("test fixture or operation should succeed");
    assert!(
        spreadsheet
            .rdf_graphs()
            .expect("test fixture or operation should succeed")
            .is_empty()
    );
    assert_eq!(
        constants::ODF_SPREADSHEET,
        "application/vnd.oasis.opendocument.spreadsheet"
    );
}

#[test]
fn spreadsheet_facade_discovers_resources_and_extracts_local_images() {
    let spreadsheet = Spreadsheet::from_bytes(resource_package())
        .expect("test fixture or operation should succeed");

    let images = spreadsheet
        .images()
        .expect("test fixture or operation should succeed");
    assert_eq!(images.len(), 2);
    assert_eq!(images[0].part, litchi_ods::Part::Content);
    let frame = images[0]
        .frame
        .as_ref()
        .expect("test fixture or operation should succeed");
    assert_eq!(frame.sheet_name.as_deref(), Some("Sheet1"));
    assert!(frame.sheet_shape);
    assert!(matches!(
        images[0].source,
        litchi_ods::media::Source::PackagePart {
            ref path,
            manifest_media_type: Some(ref media_type),
            ..
        } if path == "Pictures/photo.png" && media_type == "image/png"
    ));
    assert_eq!(
        spreadsheet
            .image_bytes(&images[0])
            .expect("test fixture or operation should succeed"),
        Some(vec![1, 2, 3])
    );
    assert!(matches!(
        images[1].source,
        litchi_ods::media::Source::Linked { ref href }
            if href == "https://example.invalid/photo.png"
    ));
    assert_eq!(
        spreadsheet
            .image_bytes(&images[1])
            .expect("test fixture or operation should succeed"),
        None
    );

    let objects = spreadsheet
        .embedded_objects()
        .expect("test fixture or operation should succeed");
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].kind, litchi_ods::Kind::Object);
    assert!(matches!(
        objects[0].source,
        litchi_ods::embedded::Source::PackageFile {
            ref path,
            manifest_media_type: Some(ref media_type),
            ..
        } if path == "Objects/widget.bin" && media_type == "application/octet-stream"
    ));
}

fn resource_package() -> Vec<u8> {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    office:version="1.3">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:shapes>
          <draw:frame draw:name="Resources" svg:width="2cm" svg:height="1cm" table:end-cell-address="B2">
            <draw:image xlink:href="Pictures/photo.png" xlink:type="simple"/>
            <draw:image xlink:href="https://example.invalid/photo.png" xlink:type="simple"/>
            <draw:object xlink:href="Objects/widget.bin" xlink:type="simple"/>
          </draw:frame>
        </table:shapes>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;
    let manifest = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
    xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Pictures/photo.png" manifest:media-type="image/png"/>
  <manifest:file-entry manifest:full-path="Objects/widget.bin" manifest:media-type="application/octet-stream"/>
</manifest:manifest>"#
    );

    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", options)
        .expect("test fixture or operation should succeed");
    zip.write_all(mimetype.as_bytes())
        .expect("test fixture or operation should succeed");
    zip.start_file("META-INF/manifest.xml", options)
        .expect("test fixture or operation should succeed");
    zip.write_all(manifest.as_bytes())
        .expect("test fixture or operation should succeed");
    zip.start_file("content.xml", options)
        .expect("test fixture or operation should succeed");
    zip.write_all(content.as_bytes())
        .expect("test fixture or operation should succeed");
    zip.start_file("Pictures/photo.png", options)
        .expect("test fixture or operation should succeed");
    zip.write_all(&[1, 2, 3])
        .expect("test fixture or operation should succeed");
    zip.start_file("Objects/widget.bin", options)
        .expect("test fixture or operation should succeed");
    zip.write_all(&[9, 8, 7])
        .expect("test fixture or operation should succeed");
    zip.finish()
        .expect("test fixture or operation should succeed");
    output.into_inner()
}
