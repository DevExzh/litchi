use litchi_odf::Document;
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document(body: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
    );
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    write!(
        zip,
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#
    )
    .unwrap();
    zip.finish().unwrap();
    Document::from_bytes(output.into_inner()).unwrap()
}

#[test]
fn document_exposes_explicit_page_sequence_without_paginating() {
    let document = document(
        r#"<p:page-sequence xmlns:p="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <p:page p:master-page-name="First"/>
            <p:page p:master-page-name="Left"/>
            <p:page p:master-page-name="Right"/>
        </p:page-sequence>"#,
    );

    let page_sequence = document.page_sequence().unwrap().unwrap();
    assert_eq!(page_sequence.master_page_names, ["First", "Left", "Right"]);
}
