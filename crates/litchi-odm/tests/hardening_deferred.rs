#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error;
use litchi_odf_common::core::PackageWriter;
use litchi_odm::Master;

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
    r#"<office:body><office:text><text:section text:name="intro" xml:id="sec_intro">"#,
    r#"<text:p>placeholder</text:p></text:section></office:text></office:body>"#,
    r#"</office:document-content>"#,
);

fn package(mimetype: &str, content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn content(inner: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text>{inner}</office:text></office:body></office:document-content>"#
    )
}

#[test]
fn odm_truncation_and_mutation_sweeps_never_panic() {
    let bytes = package(MIME, CONTENT);
    for end in 0..bytes.len() {
        drop(Master::from_bytes(bytes[..end].to_vec()));
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.clone();
        mutated[position] ^= 1;
        drop(Master::from_bytes(mutated));
    }
    assert!(Master::from_bytes(bytes).is_ok());
}

#[test]
fn odm_malformed_and_misplaced_inputs_return_typed_errors() {
    let cases = [
        package(
            MIME,
            &content(r#"<text:section text:name="dup"/><text:section text:name="dup"/>"#),
        ),
        package(
            MIME,
            &content(
                r#"<text:section text:name="a" xml:id="same"/><text:section text:name="b" xml:id="same"/>"#,
            ),
        ),
        package(
            MIME,
            r#"<?xml version="1.0"?><!DOCTYPE x><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#,
        ),
        package(
            MIME,
            r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><text:section text:name="outside"/><office:text/></office:body></office:document-content>"#,
        ),
        package("application/vnd.oasis.opendocument.text", CONTENT),
        b"not a zip at all".to_vec(),
        Vec::new(),
    ];
    for case in cases {
        assert!(matches!(
            Master::from_bytes(case),
            Err(Error::InvalidFormat(_))
        ));
    }
}
