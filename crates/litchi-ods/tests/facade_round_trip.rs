use litchi_odf_common::constants;
use litchi_ods::rdf::{Object, Subject, Triple};
use litchi_ods::{Spreadsheet, SpreadsheetBuilder};

#[test]
fn builder_and_package_facade_round_trip() {
    let bytes = SpreadsheetBuilder::new().build().unwrap();
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).unwrap();
    assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    assert_eq!(spreadsheet.into_bytes(), bytes);
}

#[test]
fn spreadsheet_facade_owns_rdf_crud() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:spreadsheet><table:table xml:id="sheet" table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#;
    let bytes = SpreadsheetBuilder::new()
        .content_xml(content)
        .build()
        .unwrap();
    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let triple = Triple {
        subject: Subject::Iri("#sheet".to_string()),
        predicate: "https://example.invalid/schema#label".to_string(),
        object: Object::Literal {
            value: "Sheet".to_string(),
            datatype: None,
            language: None,
        },
    };
    let path = spreadsheet.add_rdf_graph(None, &[triple.clone()]).unwrap();
    assert_eq!(path, "Metadata/metadata_1.rdf");
    assert_eq!(spreadsheet.rdf_graphs().unwrap()[0].triples, [triple]);
    spreadsheet.remove_rdf_graph(&path).unwrap();
    assert!(spreadsheet.rdf_graphs().unwrap().is_empty());
    assert_eq!(
        constants::ODF_SPREADSHEET,
        "application/vnd.oasis.opendocument.spreadsheet"
    );
}
