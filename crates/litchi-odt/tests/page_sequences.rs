use litchi_odt::Document;
mod support;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document(body: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
    );
    Document::from_bytes(support::package(
        MIMETYPE,
        &[("content.xml", content.as_bytes())],
    ))
    .unwrap()
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
