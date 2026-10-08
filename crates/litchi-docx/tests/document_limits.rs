use litchi_core::Position;
use litchi_docx::document::{Error, LimitKind, Snapshot};

#[test]
fn depth_error_exposes_typed_kind_and_canonical_counts() {
    let xml = format!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}{}</w:body></w:document>",
        "<w:customXml>".repeat(256),
        "</w:customXml>".repeat(256)
    );
    assert!(matches!(
        Snapshot::from_xml(xml.into_bytes()),
        Err(Error::Limit {
            kind: LimitKind::XmlDepth,
            observed: 257,
            maximum: 256
        })
    ));
}

#[test]
fn operation_limit_preserves_staging_and_has_content_free_diagnostics() {
    let xml = b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>initial</w:t></w:r></w:p></w:body></w:document>".to_vec();
    let snapshot = Snapshot::from_xml(xml).unwrap();
    let mut edit = snapshot.edit();
    for i in 0..4096 {
        edit.replace_paragraph_text(
            Position::new(0),
            if i % 2 == 0 {
                "private-author-a"
            } else {
                "private-author-b"
            },
        )
        .unwrap();
    }
    let before = edit.projected().xml_bytes().to_vec();
    let error = edit
        .replace_paragraph_text(Position::new(0), "private-author-c")
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Operations,
            observed: 4097,
            maximum: 4096
        }
    ));
    assert!(!format!("{error} {error:?}").contains("private-author"));
    assert_eq!(
        format!("{error:?}"),
        "Limit { kind: Operations, observed: 4097, maximum: 4096 }"
    );
    assert_eq!(edit.projected().xml_bytes(), before);
    assert_eq!(edit.commit().unwrap().diagnostics().operations(), 4096);
}
