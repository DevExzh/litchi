//! Regression coverage for ordinary edits of a real DOCX source package.
//!
//! The source relationship member is deliberately checked as raw bytes.  A
//! semantically equivalent relationship graph is insufficient here: the
//! preservation contract includes the producer's relationship order and
//! lexical spelling for an untouched member.

use std::error::Error;

use litchi_docx::Package;
use litchi_opc::phys_pkg::OwnedPhysPkgReader;

const SOURCE: &[u8] =
    include_bytes!("../../../test-data/ooxml/docx/documentProperties.docx");
const RELATIONSHIPS_MEMBER: &str = "word/_rels/document.xml.rels";
const EDIT_MARKER: &str = "litchi-docx-preservation-regression";

fn member(bytes: &[u8], name: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    Ok(OwnedPhysPkgReader::from_bytes(bytes.to_vec())?.read_member(name)?)
}

#[test]
fn docx_ordinary_save_preserves_document_relationship_bytes() -> Result<(), Box<dyn Error>> {
    let directory = tempfile::tempdir()?;
    let source_path = directory.path().join("source.docx");
    let output_path = directory.path().join("output.docx");
    std::fs::write(&source_path, SOURCE)?;

    let source_relationships = member(SOURCE, RELATIONSHIPS_MEMBER)?;
    let mut package = Package::open(&source_path)?;
    package
        .document_mut()?
        .add_paragraph_with_text(EDIT_MARKER);
    package.save(&output_path)?;

    let output = std::fs::read(&output_path)?;
    assert_eq!(
        member(&output, RELATIONSHIPS_MEMBER)?,
        source_relationships,
        "an untouched document relationship member must retain exact bytes and order"
    );

    let reopened = Package::open(&output_path)?;
    assert!(
        reopened.document()?.text()?.contains(EDIT_MARKER),
        "the semantic edit must survive the save"
    );
    Ok(())
}
