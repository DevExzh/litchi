use litchi_docx::{Package, Result};
use litchi_opc::PackURI;
use tempfile::NamedTempFile;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Build a package whose `/word/document.xml` has the given blob and run
/// the document accessor battery over it. A panic anywhere fails the test.
fn with_document_blob(blob: &[u8]) -> (Result<usize>, Result<usize>, Result<String>) {
    let output = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/word/document.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(blob.to_vec());
            Ok(())
        })
        .unwrap();

    let document = package.document().unwrap();
    (
        document.paragraph_count(),
        document.table_count(),
        document.text(),
    )
}

#[test]
fn truncated_document_xml_never_panics() {
    let cases: Vec<Vec<u8>> = vec![
        Vec::new(),
        b"<w:document".to_vec(),
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>hi"#).into_bytes(),
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr>"#).into_bytes(),
        format!(r#"<w:document xmlns:w="{W}"><w:body></w:document>"#).into_bytes(),
        b"\xff\x00not xml\x01".to_vec(),
    ];
    for blob in cases {
        // Must return (Ok or typed Err) — never panic.
        let _ = with_document_blob(&blob);
    }
}

#[test]
fn excessive_nesting_depth_is_rejected() {
    let mut blob = format!(r#"<w:document xmlns:w="{W}"><w:body><w:tbl>"#).into_bytes();
    blob.extend_from_slice(b"<w:tbl>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</w:tbl>".repeat(100_001).as_slice());
    blob.extend_from_slice(b"</w:body></w:document>");
    let (paragraphs, tables, _) = with_document_blob(&blob);
    assert!(tables.is_err(), "deeply nested document XML was accepted");
    assert!(
        paragraphs.is_err(),
        "deeply nested document XML was accepted"
    );
}

#[test]
fn excessive_element_count_is_rejected() {
    let mut blob = format!(r#"<w:document xmlns:w="{W}"><w:body>"#).into_bytes();
    blob.extend_from_slice(b"<w:p/>".repeat(2_000_000).as_slice());
    blob.extend_from_slice(b"</w:body></w:document>");
    let (paragraphs, _, _) = with_document_blob(&blob);
    assert!(paragraphs.is_err(), "oversized document XML was accepted");
}
