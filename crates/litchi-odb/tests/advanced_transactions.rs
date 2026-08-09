use litchi_odb::{
    Builder, Column, Component, ComponentKind, CompositionLimits, Connection, Database,
    DependencyDisposition, Index, IndexColumn, JoinedEdits, Key, KeyColumn, KeyKind, MergeChoice,
    MergePlan, Patch, Query, SealedPatch, Table, TableKind,
};

const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source/><d:queries/></o:database></o:body></o:document-content>"#;

fn source() -> Database {
    Database::from_bytes(Builder::new().content_xml(SOURCE).build().unwrap()).unwrap()
}

fn limits() -> CompositionLimits {
    CompositionLimits::new(16, 16, 64, 16)
}

fn query_patch(source: &Database, name: &str, command: &str) -> Patch {
    let mut edit = source.edit();
    edit.add_query(Query::new(name, command)).unwrap();
    edit.commit().unwrap().patch().clone()
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn durable_semantic_patch_is_canonical_reversible_sealable_and_stale_checked() {
    let source = source();
    let patch = query_patch(&source, "durable", "SELECT 1");
    let durable = patch.durable().unwrap();
    let first = durable.to_deterministic_json().unwrap();
    assert_eq!(first, durable.to_deterministic_json().unwrap());
    assert!(
        std::str::from_utf8(&first)
            .unwrap()
            .contains("odb.query.create")
    );

    let decoded = litchi_odb::DurablePatch::from_deterministic_json(&first).unwrap();
    let applied = decoded.apply(&source).unwrap();
    assert!(
        applied
            .catalog()
            .unwrap()
            .query("durable")
            .unwrap()
            .is_some()
    );
    assert_eq!(
        decoded.inverse().apply(&applied).unwrap().as_bytes(),
        source.as_bytes()
    );
    assert!(decoded.apply(&applied).is_err());

    let sealed_json = decoded.seal().to_deterministic_json().unwrap();
    let sealed = SealedPatch::from_deterministic_json(&sealed_json).unwrap();
    assert_eq!(
        sealed.apply(&source).unwrap().as_bytes(),
        applied.as_bytes()
    );
    assert!(sealed.apply(&applied).is_err());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn independent_edits_join_deterministically_and_overlap_is_structured() {
    let source = source();
    let alpha = query_patch(&source, "alpha", "SELECT 1");
    let beta = query_patch(&source, "beta", "SELECT 2");
    let alpha_prepared = alpha.prepare("alpha", limits()).unwrap();
    let beta_prepared = beta.prepare("beta", limits()).unwrap();
    let mut joined = JoinedEdits::new(alpha_prepared.lineage().clone(), limits());
    joined.join(beta_prepared).unwrap();
    joined.join(alpha_prepared).unwrap();
    let composed = Patch::compose(joined).unwrap();
    let database = composed.apply(&source).unwrap();
    assert_eq!(database.catalog().unwrap().queries()[0].name(), "alpha");
    assert_eq!(database.catalog().unwrap().queries()[1].name(), "beta");

    let left = query_patch(&source, "same", "SELECT 3")
        .prepare("left", limits())
        .unwrap();
    let right = query_patch(&source, "same", "SELECT 4")
        .prepare("right", limits())
        .unwrap();
    let mut conflicts = JoinedEdits::new(left.lineage().clone(), limits());
    conflicts.join(left).unwrap();
    let error = conflicts.join(right).unwrap_err();
    assert!(matches!(
        error.failure(),
        litchi_core::SubEditJoinFailure::Overlap(values) if !values.is_empty()
    ));
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn three_way_plan_is_non_mutating_and_requires_explicit_resolution() {
    let source = source();
    let left_edit = query_patch(&source, "same", "SELECT 10")
        .prepare("left", limits())
        .unwrap();
    let right_edit = query_patch(&source, "same", "SELECT 20")
        .prepare("right", limits())
        .unwrap();
    let mut left = JoinedEdits::new(left_edit.lineage().clone(), limits());
    left.join(left_edit).unwrap();
    let mut right = JoinedEdits::new(right_edit.lineage().clone(), limits());
    right.join(right_edit).unwrap();

    let plan = MergePlan::new(left, right).unwrap();
    assert!(!plan.conflicts().is_empty());
    let mut unresolved = *plan.finish().unwrap_err();
    unresolved.resolve(MergeChoice::Right);
    let selected = unresolved.finish().unwrap();
    let merged = Patch::compose(selected).unwrap().apply(&source).unwrap();
    assert_eq!(
        merged
            .catalog()
            .unwrap()
            .query("same")
            .unwrap()
            .unwrap()
            .command(),
        "SELECT 20"
    );
    assert!(source.catalog().unwrap().queries().is_empty());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn cascade_disposition_removes_modeled_relation_and_index_dependencies() {
    let source = source();
    let mut create = source.edit();
    create
        .add_table(Table::new("parent", TableKind::Definition).with_column(Column::new("id")))
        .unwrap();
    create
        .add_table(
            Table::new("child", TableKind::Definition)
                .with_column(Column::new("parent_id"))
                .with_key(
                    Key::new("child_parent", KeyKind::Foreign)
                        .with_referenced_table("parent")
                        .with_column(KeyColumn::new("parent_id").with_related_column("id")),
                )
                .with_index(
                    Index::new("child_parent_idx").with_column(IndexColumn::new("parent_id")),
                ),
        )
        .unwrap();
    let database = create.commit().unwrap().into_database();

    assert!(database.edit().remove_table("parent").is_err());
    let mut cascade = database.edit();
    cascade
        .remove_table_with("parent", DependencyDisposition::Cascade)
        .unwrap();
    let removed = cascade.commit().unwrap().into_database();
    let removed_catalog = removed.catalog().unwrap();
    let child = removed_catalog.table("child").unwrap().unwrap();
    assert!(child.keys().is_empty());

    let mut remove_column = removed.edit();
    assert!(remove_column.remove_column("child", "parent_id").is_err());
    remove_column
        .remove_column_with("child", "parent_id", DependencyDisposition::Cascade)
        .unwrap();
    let removed = remove_column.commit().unwrap().into_database();
    let catalog = removed.catalog().unwrap();
    let child = catalog.table("child").unwrap().unwrap();
    assert!(child.columns().is_empty());
    assert!(child.indices().is_empty());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn every_authored_fragment_family_rejects_invalid_raw_scalar_or_markup() {
    let source = source();
    assert!(
        source
            .edit()
            .add_table(Table::new("bad\0table", TableKind::Definition))
            .is_err()
    );
    assert!(
        source
            .edit()
            .add_query(Query::new("q", "SELECT \u{1}"))
            .is_err()
    );
    assert!(
        source
            .edit()
            .set_connection(Some(Connection::Server {
                host: "bad\0host".to_string(),
                database: "db".to_string(),
            }))
            .is_err()
    );
    assert!(
        source
            .edit()
            .add_component(Component::new(ComponentKind::Form, "bad\0form"))
            .is_err()
    );
    assert!(source.edit().add_producer_extension("<lo:broken>").is_err());

    let mut create = source.edit();
    create
        .add_table(Table::new("schema", TableKind::Definition).with_column(Column::new("id")))
        .unwrap();
    let database = create.commit().unwrap().into_database();
    assert!(
        database
            .edit()
            .add_column("schema", Column::new("bad\0column"))
            .is_err()
    );
    assert!(
        database
            .edit()
            .add_key(
                "schema",
                Key::new("bad-key", KeyKind::Primary).with_column(KeyColumn::new("bad\0column")),
            )
            .is_err()
    );
    assert!(
        database
            .edit()
            .add_index(
                "schema",
                Index::new("bad-index").with_column(IndexColumn::new("bad\0column")),
            )
            .is_err()
    );
    assert_eq!(database.catalog().unwrap().tables().len(), 1);
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn signed_real_odb_is_inventoried_and_changed_publication_is_refused() {
    let database = Database::from_bytes(Vec::from(include_bytes!(
        "../../../3rdparty/libreoffice-core/xmlsecurity/qa/unit/signing/data/odb_signed_macros.odb"
    )))
    .unwrap();
    assert!(database.protection_status().unwrap().is_signed());
    let unchanged = database.edit().commit().unwrap();
    assert!(!unchanged.changed());
    assert_eq!(unchanged.database().as_bytes(), database.as_bytes());

    let mut changed = database.edit();
    changed.add_query(Query::new("inert", "SELECT 1")).unwrap();
    assert!(changed.commit().is_err());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn bounded_transfer_copies_only_inert_semantic_declarations() {
    let donor = source();
    let mut author = donor.edit();
    author
        .add_table(Table::new("transferred", TableKind::Definition).with_column(Column::new("id")))
        .unwrap();
    author
        .add_query(Query::new(
            "transferred-query",
            "SELECT id FROM transferred",
        ))
        .unwrap();
    author
        .add_component(Component::new(ComponentKind::Report, "transferred-report"))
        .unwrap();
    let donor = author.commit().unwrap().into_database();

    let destination = source();
    let mut transfer = destination.edit();
    transfer.transfer_table_from(&donor, "transferred").unwrap();
    transfer
        .transfer_query_from(&donor, "transferred-query")
        .unwrap();
    transfer
        .transfer_component_from(&donor, ComponentKind::Report, "transferred-report")
        .unwrap();
    let received = transfer.commit().unwrap().into_database();
    let catalog = received.catalog().unwrap();
    assert!(catalog.table("transferred").unwrap().is_some());
    assert!(catalog.query("transferred-query").unwrap().is_some());
    assert!(catalog.components().iter().any(|component| {
        component.kind() == ComponentKind::Report && component.name() == Some("transferred-report")
    }));
}
