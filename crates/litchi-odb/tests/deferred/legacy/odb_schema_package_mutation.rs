use litchi_odf::{
    DatabaseDocument, OdfDatabaseColumnDefinition, OdfDatabaseDataType, OdfDatabaseIndex,
    OdfDatabaseIndexColumn, OdfDatabaseKey, OdfDatabaseKeyColumn, OdfDatabaseNullable,
    OdfDatabaseReferentialRule, OdfDatabaseSchemaDefinition, OdfDatabaseTableDefinition,
    OdfDocumentSigner, OdfEncryptionProfile, OdfSignatureAlgorithm, OdfSignatureValidity,
    OwnedPackage, PackageWriter,
};

const MIME: &str = "application/vnd.oasis.opendocument.base";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const RSA_KEY: &[u8] = include_bytes!("fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] = include_bytes!("fixtures/signatures/rsa-cert.der");

fn column(name: &str, data_type: OdfDatabaseDataType) -> OdfDatabaseColumnDefinition {
    let mut value = OdfDatabaseColumnDefinition::new(name);
    value.data_type = Some(data_type);
    value
}

fn schema() -> OdfDatabaseSchemaDefinition {
    let mut users = OdfDatabaseTableDefinition::new("用户 & users");
    let mut id = column("id", OdfDatabaseDataType::Integer);
    id.nullable = Some(OdfDatabaseNullable::NoNulls);
    id.autoincrement = Some(true);
    users.columns = vec![
        id,
        column("tenant", OdfDatabaseDataType::Integer),
        column("name <display>", OdfDatabaseDataType::VarChar),
    ];
    users.keys = Some(vec![OdfDatabaseKey::primary(
        Some("pk_users".into()),
        vec!["id".into(), "tenant".into()],
    )]);
    users.indices = Some(vec![OdfDatabaseIndex::new(
        "ix_name",
        vec!["name <display>".into()],
    )]);

    let mut orders = OdfDatabaseTableDefinition::new("orders");
    orders.columns = vec![
        column("user_id", OdfDatabaseDataType::Integer),
        column("tenant", OdfDatabaseDataType::Integer),
        column("total", OdfDatabaseDataType::Decimal),
    ];
    let mut relation = OdfDatabaseKey::foreign(
        Some("fk_orders_users".into()),
        "用户 & users",
        vec![
            ("user_id".into(), "id".into()),
            ("tenant".into(), "tenant".into()),
        ],
    );
    relation.update_rule = Some(OdfDatabaseReferentialRule::Cascade);
    relation.delete_rule = Some(OdfDatabaseReferentialRule::Restrict);
    orders.keys = Some(vec![relation]);
    let mut total_index = OdfDatabaseIndex::new("ix_total", vec!["total".into()]);
    total_index.column_groups[0][0].ascending = Some(false);
    orders.indices = Some(vec![total_index]);

    let mut view = OdfDatabaseTableDefinition::new_view("order summary");
    view.columns = vec![
        column("tenant", OdfDatabaseDataType::Integer),
        column("total", OdfDatabaseDataType::Decimal),
    ];
    OdfDatabaseSchemaDefinition {
        tables: vec![users, orders, view],
    }
}

fn content(schema: &str, table_representation: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:db=\"{DB}\" xmlns:xlink=\"{XLINK}\" xmlns:v=\"urn:test:vendor\"><office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href=\"sdbc:embedded:firebird\" xlink:type=\"simple\"/></db:connection-data></db:data-source><!--keep-->{table_representation}<v:cache v:keep=\"yes\"/>{schema}</office:database></office:body></office:document-content>"
    )
}

fn package(xml: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", xml.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("database/data", b"opaque", "application/x-firebird")
        .unwrap();
    writer
        .add_file("settings.xml", b"<settings keep='yes'/>")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn whole_schema_and_granular_crud_save_reopen() {
    let mut document = DatabaseDocument::from_bytes(package(&content("", ""))).unwrap();
    assert!(
        document
            .set_schema_definition(Some(&schema()))
            .unwrap()
            .is_none()
    );
    assert!(document.schema_definition().unwrap().unwrap().tables[2].is_view());
    document.move_schema_table(2, 1).unwrap();
    document.move_schema_column(0, 2, 1).unwrap();
    document
        .add_schema_column(1, column("count", OdfDatabaseDataType::Integer))
        .unwrap();
    document
        .add_schema_index(1, OdfDatabaseIndex::new("ix_count", vec!["count".into()]))
        .unwrap();
    document
        .add_schema_index_column(
            1,
            0,
            0,
            OdfDatabaseIndexColumn {
                name: "tenant".into(),
                ascending: Some(true),
            },
        )
        .unwrap();
    document.move_schema_index_column(1, 0, 0, 1, 0).unwrap();
    let path = std::env::temp_dir().join("litchi-odb-schema-crud.odb");
    document.save(&path).unwrap();
    let reopened = DatabaseDocument::open(&path).unwrap();
    std::fs::remove_file(path).unwrap();
    let parsed = reopened.schema_definition().unwrap().unwrap();
    assert_eq!(parsed.tables.len(), 3);
    assert!(
        String::from_utf8(reopened.get_file("content.xml").unwrap())
            .unwrap()
            .contains("用户 &amp; users")
    );
    assert_eq!(reopened.get_file("database/data").unwrap(), b"opaque");
}

#[test]
fn dangling_duplicates_and_destructive_edits_are_atomic() {
    let mut document = DatabaseDocument::from_bytes(package(&content("", ""))).unwrap();
    document.set_schema_definition(Some(&schema())).unwrap();
    for mutation in [0, 1, 2] {
        let before = document.to_bytes();
        let result = match mutation {
            0 => document.remove_schema_table(0).map(|_| ()),
            1 => document.remove_schema_column(0, 0).map(|_| ()),
            _ => document
                .add_schema_column(0, column("id", OdfDatabaseDataType::Integer))
                .map(|_| ()),
        };
        assert!(result.is_err());
        assert_eq!(document.to_bytes(), before);
    }
    let mut invalid = schema();
    invalid.tables[1].keys.as_mut().unwrap()[0].column_groups[0][0].related_column_name = None;
    assert!(invalid.validate().is_err());
    let mut invalid = schema();
    invalid.tables[0].keys.as_mut().unwrap()[0].referenced_table_name = Some("orders".into());
    assert!(invalid.validate().is_err());
    let mut invalid = schema();
    invalid.tables[1].indices.as_mut().unwrap()[0].column_groups[0][0].name = "missing".into();
    assert!(invalid.validate().is_err());
}

#[test]
fn relation_column_crud_and_clear_are_validated() {
    let mut document = DatabaseDocument::from_bytes(package(&content("", ""))).unwrap();
    document.set_schema_definition(Some(&schema())).unwrap();
    document
        .update_schema_key_column(1, 0, 0, 0, OdfDatabaseKeyColumn::foreign("user_id", "id"))
        .unwrap();
    document.remove_schema_key_column(1, 0, 0, 0).unwrap();
    let before = document.to_bytes();
    assert!(document.remove_schema_key_column(1, 0, 0, 0).is_err());
    assert_eq!(document.to_bytes(), before);
    document.remove_schema_key(1, 0).unwrap();
    document.remove_schema_table(0).unwrap();
    assert_eq!(
        document
            .clear_schema_definition()
            .unwrap()
            .unwrap()
            .tables
            .len(),
        2
    );
    assert!(document.schema_definition().unwrap().is_none());
}

#[test]
fn libreoffice_table_representation_variation_is_preserved() {
    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/extras/source/database/biblio.odb")
            .to_vec();
    let mut document = DatabaseDocument::from_bytes(bytes).unwrap();
    let original = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    let settings = document.get_file("settings.xml").unwrap();
    document.set_schema_definition(Some(&schema())).unwrap();
    let mutated = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    let insertion = original.rfind("</office:database>").unwrap();
    let mut expected = original.clone();
    expected.insert_str(insertion, &schema().to_xml_fragment().unwrap());
    assert_eq!(mutated, expected);
    assert!(mutated.find("<db:columns>").unwrap() < mutated.find("<db:order-statement").unwrap());
    assert_eq!(document.get_file("settings.xml").unwrap(), settings);
}

#[test]
fn encrypted_and_signed_schema_package_reopens() {
    let signer = OdfDocumentSigner::from_pkcs8_der(
        OdfSignatureAlgorithm::RsaSha256,
        RSA_KEY,
        vec![RSA_CERT.to_vec()],
        "2026-07-19T12:00:00Z",
    )
    .unwrap();
    let xml = content(&schema().to_xml_fragment().unwrap(), "");
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.set_document_signer(signer).unwrap();
    writer
        .set_encryption("database-password", OdfEncryptionProfile::compatible())
        .unwrap();
    writer.add_file("content.xml", xml.as_bytes()).unwrap();
    let bytes = writer.finish_to_bytes().unwrap();
    assert!(
        OwnedPackage::from_bytes(bytes.clone())
            .unwrap()
            .verify_document_signatures()
            .unwrap()
            .iter()
            .all(|result| result.validity == OdfSignatureValidity::Valid)
    );
    let reopened = DatabaseDocument::from_bytes_with_password(bytes, "database-password").unwrap();
    assert_eq!(
        reopened.schema_definition().unwrap().unwrap().tables.len(),
        3
    );
}

#[test]
fn odb_schema_crud_contains_no_unsafe_code() {
    for source in [
        include_str!("../../../src/migration/legacy/document.rs"),
        include_str!("../../../src/migration/legacy/schema.rs"),
    ] {
        assert!(!source.contains("unsafe {"));
    }
}
