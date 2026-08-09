#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::{Database, Limits, TableKind};

#[test]
fn real_libreoffice_database_exposes_inert_query_and_table_catalogs() {
    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb")
            .to_vec();
    let database = Database::from_bytes(bytes.clone()).unwrap();
    let catalog = database.catalog().unwrap();

    let query = catalog.query("AliasTest").unwrap().unwrap();
    assert_eq!(
        query.command(),
        "SELECT \"tid\" \"TestId\", \"tname\" \"TestName\" FROM \"test\""
    );
    assert_eq!(query.escape_processing(), None);
    let table = catalog.table("test").unwrap().unwrap();
    assert_eq!(table.kind(), TableKind::Representation);
    assert_eq!(table.columns()[0].name(), "tid");
    assert_eq!(table.columns()[1].name(), "tname");
    assert_eq!(database.as_bytes(), bytes.as_slice());
}

#[test]
fn real_libreoffice_schema_exposes_columns_without_connecting() {
    let database = Database::from_bytes(
        include_bytes!("../../../test-data/libreoffice-core/extras/source/database/biblio.odb")
            .to_vec(),
    )
    .unwrap();
    let catalog = database.catalog().unwrap();
    let table = catalog.table("biblio").unwrap().unwrap();
    assert_eq!(table.columns().len(), 32);
    assert_eq!(table.columns()[0].name(), "Address");
    assert_eq!(catalog.queries(), []);
}

#[test]
fn namespace_aliases_and_schema_definitions_are_semantic() {
    let content = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:schema-definition><d:table-definitions><d:table-definition d:name="ledger"><d:column-definitions><d:column-definition d:name="amount"/></d:column-definitions></d:table-definition></d:table-definitions></d:schema-definition><d:queries><d:query d:name="inert" d:command="SELECT &quot;amount&quot;" d:escape-processing="true"/></d:queries></o:database></o:body></o:document-content>"#,
    );
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(content)
            .build()
            .unwrap(),
    )
    .unwrap();
    let catalog = database.catalog().unwrap();
    assert_eq!(
        catalog.table("ledger").unwrap().unwrap().kind(),
        TableKind::Definition
    );
    assert_eq!(
        catalog.table("ledger").unwrap().unwrap().columns()[0].name(),
        "amount"
    );
    assert_eq!(
        catalog.query("inert").unwrap().unwrap().command(),
        "SELECT \"amount\""
    );
    assert_eq!(
        catalog.query("inert").unwrap().unwrap().escape_processing(),
        Some(true)
    );
    assert_eq!(catalog.to_owned().tables().len(), 1);
}

#[test]
fn catalog_enforces_its_finite_limits() {
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(
                r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:queries><d:query d:name="q" d:command="SELECT 1"/></d:queries></o:database></o:body></o:document-content>"#,
            )
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(
        database
            .catalog_with(Limits::default().with_max_queries(0))
            .is_err()
    );
}
