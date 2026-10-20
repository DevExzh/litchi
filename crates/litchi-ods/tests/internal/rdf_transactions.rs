use crate::{
    Builder, Package,
    builder::Workbook as Draft,
    metadata_graphs::Snapshot,
    rdf::{Object, Subject, Triple},
};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4"><office:body><office:spreadsheet><table:table xml:id="sheet" table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#;

fn triple(predicate: &str, value: &str) -> Triple {
    Triple {
        subject: Subject::Iri("#sheet".to_string()),
        predicate: predicate.to_string(),
        object: Object::Literal {
            value: value.to_string(),
            datatype: None,
            language: None,
        },
    }
}

#[test]
fn rdf_no_ops_preserve_exact_bytes_and_failed_staging_is_atomic() -> litchi_core::Result<()> {
    let source = Builder::new().content_xml(CONTENT).build()?.to_bytes()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let empty = snapshot.edit().commit();
    assert!(!empty.changed());
    assert_eq!(empty.snapshot().as_bytes(), source);

    let mut edit = snapshot.edit();
    assert!(edit.remove_graph("missing.rdf").is_err());
    assert!(edit.graphs().is_empty());
    assert_eq!(edit.commit().snapshot().as_bytes(), source);
    Ok(())
}

#[test]
fn mutable_facade_publishes_a_batched_rdf_transaction() -> litchi_core::Result<()> {
    let source = Builder::new().content_xml(CONTENT).build()?.to_bytes()?;
    let mut mutable = Draft::from_workbook(Package::from_bytes(&source)?.workbook()?);
    let value = triple("https://example.invalid/schema#label", "Sheet");
    mutable.edit_rdf(|edit| {
        edit.add_graph(Some("Metadata/sheet.rdf"), std::slice::from_ref(&value))?;
        Ok(())
    })?;

    assert_eq!(
        mutable.rdf_snapshot()?.graphs()[0].path,
        "Metadata/sheet.rdf"
    );
    Ok(())
}
