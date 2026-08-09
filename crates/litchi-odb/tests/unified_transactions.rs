#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::{
    Builder, ChangeKind, Column, Component, ComponentKind, Connection, DataType, Database, History,
    HistoryLimits, Index, IndexColumn, Key, KeyColumn, KeyKind, Query, ReferentialAction, Table,
    TableKind,
};

const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source/></o:database></o:body></o:document-content>"#;

fn source() -> Database {
    Database::from_bytes(Builder::new().content_xml(SOURCE).build().unwrap()).unwrap()
}

#[test]
fn unified_transaction_composes_schema_relation_query_connection_component_and_extension_crud() {
    let source = source();
    let mut edit = source.edit();
    edit.set_connection(Some(Connection::Resource(
        "sdbc:embedded:firebird".to_string(),
    )))
    .unwrap();
    edit.add_table(
        Table::new("customers", TableKind::Definition).with_column(
            Column::new("id")
                .with_data_type(Some(DataType::BigInt))
                .with_nullable(Some(false)),
        ),
    )
    .unwrap();
    edit.add_table(
        Table::new("orders", TableKind::Definition)
            .with_column(
                Column::new("id")
                    .with_data_type(Some(DataType::BigInt))
                    .with_autoincrement(Some(true)),
            )
            .with_column(Column::new("customer_id").with_data_type(Some(DataType::BigInt))),
    )
    .unwrap();
    edit.add_key(
        "orders",
        Key::new("orders_customer_fk", KeyKind::Foreign)
            .with_referenced_table("customers")
            .with_update_rule(Some(ReferentialAction::Cascade))
            .with_delete_rule(Some(ReferentialAction::Restrict))
            .with_column(KeyColumn::new("customer_id").with_related_column("id")),
    )
    .unwrap();
    edit.add_index(
        "orders",
        Index::new("orders_customer_idx")
            .with_unique(Some(false))
            .with_column(IndexColumn::new("customer_id").with_ascending(Some(true))),
    )
    .unwrap();
    edit.add_query(Query::new("recent", "SELECT * FROM orders").with_escape_processing(Some(true)))
        .unwrap();
    edit.add_component(
        Component::new(ComponentKind::Form, "orders_form")
            .with_href("forms/Orders")
            .with_as_template(Some(false)),
    )
    .unwrap();
    edit.add_component(Component::new(ComponentKind::Report, "orders_report"))
        .unwrap();
    edit.add_producer_extension(
        r#"<lo:database-settings xmlns:lo="urn:org:documentfoundation:names:experimental:database:xmlns:loext:1.0" lo:keep="true"/>"#,
    )
    .unwrap();

    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(commit.patch().changes().len() >= 8);
    assert_eq!(commit.patch().changes()[0].kind(), ChangeKind::Connection);
    let catalog = commit.database().catalog().unwrap();
    assert_eq!(catalog.tables().len(), 2);
    assert_eq!(catalog.queries()[0].command(), "SELECT * FROM orders");
    assert_eq!(catalog.components().len(), 2);
    assert_eq!(catalog.relations().len(), 1);
    assert_eq!(catalog.relations()[0].table(), "orders");
    assert_eq!(catalog.relations()[0].referenced_table(), "customers");
    assert_eq!(
        catalog.relations()[0].columns()[0].name(),
        Some("customer_id")
    );
    assert!(matches!(
        catalog.connection(),
        Some(Connection::Resource(value)) if value == "sdbc:embedded:firebird"
    ));
    let extensions = commit.database().producer_extensions().unwrap();
    assert_eq!(extensions.len(), 1);
    assert_eq!(extensions[0].local_name(), "database-settings");
    assert!(extensions[0].xml().contains("lo:keep=\"true\""));

    let restored = commit.patch().inverse().apply(commit.database()).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
    assert!(commit.patch().apply(commit.database()).is_err());
}

#[test]
fn dependency_aware_renames_and_complete_deletion_reopen_cleanly() {
    let source = source();
    let mut create = source.edit();
    create
        .add_table(
            Table::new("parent", TableKind::Definition)
                .with_column(Column::new("id").with_data_type(Some(DataType::Integer))),
        )
        .unwrap();
    create
        .add_table(
            Table::new("child", TableKind::Definition)
                .with_column(Column::new("parent_id").with_data_type(Some(DataType::Integer)))
                .with_key(
                    Key::new("child_parent_fk", KeyKind::Foreign)
                        .with_referenced_table("parent")
                        .with_column(KeyColumn::new("parent_id").with_related_column("id")),
                ),
        )
        .unwrap();
    create.add_query(Query::new("q", "SELECT 1")).unwrap();
    let created = create.commit().unwrap().into_database();

    let mut edit = created.edit();
    edit.rename_table("parent", "account").unwrap();
    edit.rename_column("account", "id", "account_id").unwrap();
    let renamed = edit.commit().unwrap().into_database();
    let renamed_catalog = renamed.catalog().unwrap();
    let relation = &renamed_catalog.relations()[0];
    assert_eq!(relation.referenced_table(), "account");
    assert_eq!(relation.columns()[0].related_column(), Some("account_id"));

    let mut remove = renamed.edit();
    remove.remove_key("child", "child_parent_fk").unwrap();
    remove.remove_column("child", "parent_id").unwrap();
    remove.remove_table("child").unwrap();
    remove.remove_table("account").unwrap();
    remove.remove_query("q").unwrap();
    let empty = remove.commit().unwrap().into_database();
    assert!(empty.catalog().unwrap().tables().is_empty());
    assert!(empty.catalog().unwrap().queries().is_empty());
}

#[test]
fn real_pretty_libreoffice_package_changes_only_the_selected_query_xml() {
    let bytes = Vec::from(include_bytes!(
        "../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb"
    ));
    let database = Database::from_bytes(bytes.clone()).unwrap();
    let original = database.content_xml().to_owned();
    let mut edit = database.edit();
    edit.set_query_command("AliasTest", "SELECT 7").unwrap();
    let commit = edit.commit().unwrap();

    let expected = original.replace(
        "db:command=\"SELECT &quot;tid&quot; &quot;TestId&quot;, &quot;tname&quot; &quot;TestName&quot; FROM &quot;test&quot;\"",
        "db:command=\"SELECT 7\"",
    );
    assert_eq!(commit.database().content_xml(), expected);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.database())
            .unwrap()
            .as_bytes(),
        bytes
    );
}

#[test]
fn budgeted_history_undoes_and_redoes_full_reopened_snapshots() {
    let source = source();
    let mut edit = source.edit();
    edit.add_query(Query::new("history", "SELECT 1")).unwrap();
    let commit = edit.commit().unwrap();
    let weight = u64::try_from(commit.database().as_bytes().len()).unwrap();
    let mut history = History::new(source.clone(), HistoryLimits::new(2, weight));
    assert!(
        history
            .record(commit.into_database(), weight)
            .unwrap()
            .is_empty()
    );
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), source.as_bytes());
    assert!(history.redo());
    assert!(
        history
            .current()
            .catalog()
            .unwrap()
            .query("history")
            .unwrap()
            .is_some()
    );
}
