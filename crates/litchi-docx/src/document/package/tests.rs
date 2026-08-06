//! Focused behavioral checks for the document-package facade seams.

#[test]
fn minimal_package_exposes_all_document_facade_layers() {
    let package = crate::package::Package::new().expect("minimal DOCX package");
    let document = package.document().expect("main document");

    assert!(document.text().is_ok());
    assert!(document.paragraphs().is_ok());
    assert!(document.sections().is_ok());
    assert!(document.styles().is_ok());
    assert!(document.settings().is_ok());
    assert!(document.numbering().is_ok());
    assert!(document.numbering_snapshot().is_ok());
    assert!(document.statistics().is_ok());
}
