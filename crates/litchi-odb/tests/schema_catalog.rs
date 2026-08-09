#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::{
    ComponentKind, DataType, Database, KeyKind, Limits, ReferentialAction, TableKind,
    connection::Connection,
};

#[test]
fn real_libreoffice_database_exposes_inert_query_and_table_catalogs() {
    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb")
            .to_vec();
    let database = Database::from_bytes(bytes.clone()).unwrap();
    let catalog = database.catalog().unwrap();

    assert!(matches!(
        catalog.connection(),
        Some(Connection::Resource(href)) if href == "sdbc:embedded:firebird"
    ));

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
    assert!(matches!(
        catalog.connection(),
        Some(Connection::File(href)) if href == "$(userurl)/database/biblio"
    ));
    let table = catalog.table("biblio").unwrap().unwrap();
    assert_eq!(table.columns().len(), 32);
    assert_eq!(table.columns()[0].name(), "Address");
    assert_eq!(catalog.queries(), []);
}

#[test]
fn real_libreoffice_reports_are_typed_inert_components() {
    let database = Database::from_bytes(Vec::from(include_bytes!(
        "../../../3rdparty/libreoffice-core/reportdesign/qa/unit/data/roundTrip.odb"
    )))
    .unwrap();
    let catalog = database.catalog().unwrap();

    assert_eq!(catalog.components().len(), 2);
    let report = &catalog.components()[0];
    assert_eq!(report.kind(), ComponentKind::Report);
    assert_eq!(report.name(), Some("BasicFields"));
    assert_eq!(report.href(), Some("reports/Obj21"));
    assert_eq!(report.as_template(), Some(false));
    assert_eq!(report.title(), None);
    assert_eq!(report.description(), None);
}

#[test]
fn component_catalog_is_bounded_and_parent_checked() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:forms><d:component d:name="form"/></d:forms></o:database></o:body></o:document-content>"#;
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(content)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(
        database
            .catalog_with(Limits::default().with_max_components(0))
            .is_err()
    );

    let malformed = content.replace("<d:forms>", "").replace("</d:forms>", "");
    let malformed_database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(malformed)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(malformed_database.catalog().is_err());
}

#[test]
fn schema_columns_keys_and_indices_are_typed_and_bounded() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:schema-definition><d:table-definitions><d:table-definition d:name="orders"><d:column-definitions><d:column-definition d:name="id" d:data-type="bigint" d:type-name="BIGINT" d:precision="19" d:is-nullable="no-nulls" d:is-empty-allowed="false" d:is-autoincrement="true"/><d:column-definition d:name="customer_id" d:data-type="integer"/></d:column-definitions><d:keys><d:key d:name="orders_pk" d:type="primary"><d:key-columns><d:key-column d:name="id"/></d:key-columns></d:key><d:key d:name="orders_customer_fk" d:type="foreign" d:referenced-table-name="customers" d:update-rule="cascade" d:delete-rule="restrict"><d:key-columns><d:key-column d:name="customer_id" d:related-column-name="id"/></d:key-columns></d:key></d:keys><d:indices><d:index d:name="orders_customer_idx" d:is-unique="false" d:is-clustered="true"><d:index-columns><d:index-column d:name="customer_id" d:is-ascending="true"/></d:index-columns></d:index></d:indices></d:table-definition></d:table-definitions></d:schema-definition></o:database></o:body></o:document-content>"#;
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(content)
            .build()
            .unwrap(),
    )
    .unwrap();
    let catalog = database.catalog().unwrap();
    let table = catalog.table("orders").unwrap().unwrap();

    let id = &table.columns()[0];
    assert_eq!(id.data_type(), Some(DataType::BigInt));
    assert_eq!(id.type_name(), Some("BIGINT"));
    assert_eq!(id.precision(), Some(19));
    assert_eq!(id.nullable(), Some(false));
    assert_eq!(id.empty_allowed(), Some(false));
    assert_eq!(id.autoincrement(), Some(true));
    assert_eq!(table.keys().len(), 2);
    assert_eq!(table.keys()[0].kind(), KeyKind::Primary);
    assert_eq!(table.keys()[0].columns()[0].name(), Some("id"));
    assert_eq!(table.keys()[1].kind(), KeyKind::Foreign);
    assert_eq!(table.keys()[1].referenced_table(), Some("customers"));
    assert_eq!(
        table.keys()[1].update_rule(),
        Some(ReferentialAction::Cascade)
    );
    assert_eq!(
        table.keys()[1].delete_rule(),
        Some(ReferentialAction::Restrict)
    );
    assert_eq!(table.keys()[1].columns()[0].related_column(), Some("id"));
    assert_eq!(table.indices().len(), 1);
    assert_eq!(table.indices()[0].name(), "orders_customer_idx");
    assert_eq!(table.indices()[0].unique(), Some(false));
    assert_eq!(table.indices()[0].clustered(), Some(true));
    assert_eq!(table.indices()[0].columns()[0].ascending(), Some(true));

    assert!(
        database
            .catalog_with(Limits::default().with_max_keys(1))
            .is_err()
    );
    assert!(
        database
            .catalog_with(Limits::default().with_max_indices(0))
            .is_err()
    );
}

#[test]
fn schema_rejects_invalid_typed_constraint_scalars() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:schema-definition><d:table-definitions><d:table-definition d:name="broken"><d:column-definitions><d:column-definition d:name="id" d:data-type="not-an-odf-type"/></d:column-definitions></d:table-definition></d:table-definitions></d:schema-definition></o:database></o:body></o:document-content>"#;
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(content)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(database.catalog().is_err());
}

#[test]
fn server_connection_is_schema_bound_and_never_opened() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source><d:connection-data><d:database-description><d:server-database d:hostname="db.example.test" d:database-name="ledger"/></d:database-description></d:connection-data></d:data-source></o:database></o:body></o:document-content>"#;
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(content)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(matches!(
        database.catalog().unwrap().connection(),
        Some(Connection::Server { host, database: server_database }) if host == "db.example.test" && server_database == "ledger"
    ));
}

#[test]
fn namespace_aliases_and_schema_definitions_are_semantic() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:schema-definition><d:table-definitions><d:table-definition d:name="ledger"><d:column-definitions><d:column-definition d:name="amount"/></d:column-definitions></d:table-definition></d:table-definitions></d:schema-definition><d:queries><d:query d:name="inert" d:command="SELECT &quot;amount&quot;" d:escape-processing="true"/></d:queries></o:database></o:body></o:document-content>"#;
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

#[test]
fn catalog_rejects_semantic_nodes_outside_their_canonical_collections() {
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(
                r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:query d:name="wrong" d:command="SELECT 1"/></o:database></o:body></o:document-content>"#,
            )
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(database.catalog().is_err());
}

#[test]
fn catalog_rejects_columns_beyond_the_configured_budget_before_insertion() {
    let database = Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(
                r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:table-representations><d:table-representation d:name="ledger"><d:columns><d:column d:name="amount"/></d:columns></d:table-representation></d:table-representations></o:database></o:body></o:document-content>"#,
            )
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(
        database
            .catalog_with(Limits::default().with_max_columns(0))
            .is_err()
    );
}
