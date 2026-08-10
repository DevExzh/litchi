use litchi_odb::{
    ActiveContentDisposition, ActiveContentKind, Builder, ChangeKind, Column, Component,
    ComponentKind, CompositionLimits, Connection, Database, DependencyDisposition, EditPolicy,
    History, HistoryLimits, Index, IndexColumn, JoinedEdits, Key, KeyColumn, KeyKind, MergeChoice,
    MergePlan, Patch, ProtectionOperation, ProtectionSupport, Query, QueryUpdateTarget,
    SealedPatch, SignaturePolicy, Table, TableKind,
};

const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source/><d:queries/></o:database></o:body></o:document-content>"#;

#[expect(clippy::unwrap_used, reason = "fixed compact fixture construction")]
fn source() -> Database {
    Database::from_bytes(Builder::new().content_xml(SOURCE).build().unwrap()).unwrap()
}

fn limits() -> CompositionLimits {
    CompositionLimits::new(16, 16, 64, 16)
}

#[expect(
    clippy::unwrap_used,
    reason = "fixed semantic patch fixture construction"
)]
fn query_patch(source: &Database, name: &str, command: &str) -> Patch {
    let mut edit = source.edit();
    edit.add_query(Query::new(name, command)).unwrap();
    edit.commit().unwrap().patch().clone()
}

#[expect(
    clippy::shadow_reuse,
    clippy::unwrap_used,
    reason = "fixed two-sided conflict fixture construction"
)]
fn conflicting(left: &Patch, right: &Patch) -> bool {
    let left = left.prepare("left", limits()).unwrap();
    let right = right.prepare("right", limits()).unwrap();
    let mut left_joined = JoinedEdits::new(left.lineage().clone(), limits());
    left_joined.join(left).unwrap();
    let mut right_joined = JoinedEdits::new(right.lineage().clone(), limits());
    right_joined.join(right).unwrap();
    !MergePlan::new(left_joined, right_joined)
        .unwrap()
        .conflicts()
        .is_empty()
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
fn three_way_conflicts_cover_schema_connection_component_and_extension_owners() {
    let source = source();
    let table_patch = |column: &str| {
        let mut edit = source.edit();
        edit.add_table(
            Table::new("same-table", TableKind::Definition).with_column(Column::new(column)),
        )
        .unwrap();
        edit.commit().unwrap().patch().clone()
    };
    assert!(conflicting(&table_patch("left"), &table_patch("right")));

    let connection_patch = |target: &str| {
        let mut edit = source.edit();
        edit.set_connection(Some(Connection::Resource(target.to_string())))
            .unwrap();
        edit.commit().unwrap().patch().clone()
    };
    assert!(conflicting(
        &connection_patch("sdbc:inert:left"),
        &connection_patch("sdbc:inert:right")
    ));

    let component_patch = |title: &str| {
        let mut edit = source.edit();
        edit.add_component(Component::new(ComponentKind::Form, "same-form").with_title(title))
            .unwrap();
        edit.commit().unwrap().patch().clone()
    };
    assert!(conflicting(
        &component_patch("Left"),
        &component_patch("Right")
    ));

    let extension_patch = |value: &str| {
        let mut edit = source.edit();
        edit.add_producer_extension(&format!(
            r#"<x:same xmlns:x="urn:example:conflict" x:value="{value}"/>"#
        ))
        .unwrap();
        edit.commit().unwrap().patch().clone()
    };
    assert!(conflicting(
        &extension_patch("left"),
        &extension_patch("right")
    ));
    assert!(source.catalog().unwrap().tables().is_empty());
}

#[test]
#[expect(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
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
#[expect(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
fn signed_real_odb_is_inventoried_and_changed_publication_is_refused() {
    let database = Database::from_bytes(Vec::from(include_bytes!(
        "../../../3rdparty/libreoffice-core/xmlsecurity/qa/unit/signing/data/odb_signed_macros.odb"
    )))
    .unwrap();
    assert!(database.protection_status().unwrap().is_signed());
    assert!(!database.digital_signatures().unwrap().is_empty());
    let unchanged = database.edit().commit().unwrap();
    assert!(!unchanged.changed());
    assert_eq!(unchanged.database().as_bytes(), database.as_bytes());

    let mut changed = database.edit();
    changed.add_query(Query::new("inert", "SELECT 1")).unwrap();
    assert!(changed.commit().is_err());

    let policy = EditPolicy::default().with_signature(SignaturePolicy::RemoveInvalidated);
    let mut changed = database.edit_with_policy(policy);
    changed
        .add_query(Query::new("inert-after-signature", "SELECT 2"))
        .unwrap();
    let committed = changed.commit().unwrap();
    let transition = committed.protection_transition().unwrap();
    assert!(transition.before().is_signed());
    assert!(!transition.after().is_signed());
    assert!(transition.signature_was_removed());
    assert!(
        Database::from_bytes(committed.database().as_bytes().to_vec())
            .unwrap()
            .catalog()
            .unwrap()
            .query("inert-after-signature")
            .unwrap()
            .is_some()
    );
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn encrypted_package_open_and_edit_lifecycle_is_explicit_and_fail_closed() {
    use litchi_odf_common::core::{PackageWriter, Profile};

    let mut writer = PackageWriter::new();
    writer
        .set_mimetype(litchi_odf_common::constants::ODF_DATABASE)
        .unwrap();
    writer
        .set_encryption("correct horse", Profile::compatible())
        .unwrap();
    writer
        .add_file_with_media_type("content.xml", SOURCE.as_bytes(), "text/xml")
        .unwrap();
    let encrypted = writer.finish().unwrap();
    assert!(Database::from_bytes(encrypted.clone()).is_err());
    assert!(Database::from_bytes_with_password(encrypted.clone(), "wrong battery").is_err());
    let database = Database::from_bytes_with_password(encrypted, "correct horse").unwrap();
    assert!(database.protection_status().unwrap().is_encrypted());

    let unchanged = database.edit().commit().unwrap();
    let transition = unchanged.protection_transition().unwrap();
    assert!(transition.before().is_encrypted());
    assert!(transition.after().is_encrypted());
    assert_eq!(unchanged.database().as_bytes(), database.as_bytes());

    let mut changed = database.edit();
    changed
        .add_query(Query::new("encrypted-inert", "SELECT 1"))
        .unwrap();
    assert!(changed.commit().is_err());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn compact_inert_producer_extension_full_crud_is_durable() {
    const NAMESPACE: &str = "urn:example:producer-extension";
    let source = source();
    let mut create = source.edit();
    create
        .add_producer_extension(
            r#"<x:settings xmlns:x="urn:example:producer-extension" x:value="one"/>"#,
        )
        .unwrap();
    assert!(
        create
            .add_producer_extension(
                r#"<x:settings xmlns:x="urn:example:producer-extension" x:value="duplicate"/>"#,
            )
            .is_err()
    );
    let created = create.commit().unwrap().into_database();

    let mut replace = created.edit();
    replace
        .replace_producer_extension(
            NAMESPACE,
            "settings",
            r#"<p:settings xmlns:p="urn:example:producer-extension" p:value="two"/>"#,
        )
        .unwrap();
    assert!(
        replace
            .replace_producer_extension(
                NAMESPACE,
                "settings",
                r#"<p:different xmlns:p="urn:example:producer-extension"/>"#,
            )
            .is_err()
    );
    let replacement = replace.commit().unwrap();
    let durable = replacement.patch().durable().unwrap();
    let replaced = durable.apply(&created).unwrap();
    assert!(
        replaced.producer_extensions().unwrap()[0]
            .xml()
            .contains("value=\"two\"")
    );
    assert_eq!(
        durable.inverse().apply(&replaced).unwrap().as_bytes(),
        created.as_bytes()
    );

    let mut remove = replaced.edit();
    remove
        .remove_producer_extension(NAMESPACE, "settings")
        .unwrap();
    let removed = remove.commit().unwrap().into_database();
    assert!(removed.producer_extensions().unwrap().is_empty());
}

#[test]
#[expect(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
fn bounded_transfer_copies_only_inert_semantic_declarations() {
    let donor = source();
    let mut author = donor.edit();
    author
        .add_table(Table::new("transferred", TableKind::Definition).with_column(Column::new("id")))
        .unwrap();
    author
        .add_table(
            Table::new("presented", TableKind::Representation)
                .with_filter_statement("id > 0")
                .with_order_statement("id ASC"),
        )
        .unwrap();
    author
        .add_query(
            Query::new("transferred-query", "SELECT id FROM transferred")
                .with_column(Column::new("id"))
                .with_filter_statement("id > 0")
                .with_order_statement("id ASC")
                .with_update_target(
                    QueryUpdateTarget::new("transferred")
                        .with_schema_name("public")
                        .with_catalog_name("main"),
                ),
        )
        .unwrap();
    author
        .add_component(Component::new(ComponentKind::Report, "transferred-report"))
        .unwrap();
    let donor = author.commit().unwrap().into_database();

    let destination = source();
    let mut transfer = destination.edit();
    transfer.transfer_table_from(&donor, "transferred").unwrap();
    transfer.transfer_table_from(&donor, "presented").unwrap();
    transfer
        .transfer_query_from(&donor, "transferred-query")
        .unwrap();
    transfer
        .transfer_component_from(&donor, ComponentKind::Report, "transferred-report")
        .unwrap();
    let received = transfer.commit().unwrap().into_database();
    let catalog = received.catalog().unwrap();
    assert!(catalog.table("transferred").unwrap().is_some());
    let presented = catalog.table("presented").unwrap().unwrap();
    assert_eq!(presented.filter_statement(), Some("id > 0"));
    assert_eq!(presented.order_statement(), Some("id ASC"));
    assert!(catalog.query("transferred-query").unwrap().is_some());
    let query = catalog.query("transferred-query").unwrap().unwrap();
    assert_eq!(query.columns().len(), 1);
    assert_eq!(query.filter_statement(), Some("id > 0"));
    assert_eq!(query.order_statement(), Some("id ASC"));
    let update_target = query.update_target().unwrap();
    assert_eq!(update_target.name(), "transferred");
    assert_eq!(update_target.schema_name(), Some("public"));
    assert_eq!(update_target.catalog_name(), Some("main"));
    assert!(catalog.components().iter().any(|component| {
        component.kind() == ComponentKind::Report && component.name() == Some("transferred-report")
    }));
}

#[test]
#[expect(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
fn linked_component_transfer_remaps_exact_payload_and_remains_reversible() {
    use litchi_odf_common::{
        core::OwnedPackage,
        package::edit::{Addition, rebuild_package},
    };

    let donor = source();
    let mut author = donor.edit();
    author
        .add_component(
            Component::new(ComponentKind::Report, "payload-report").with_href("reports/source"),
        )
        .unwrap();
    author
        .add_component(
            Component::new(ComponentKind::Report, "payload-report-two")
                .with_href("reports/source-two"),
        )
        .unwrap();
    let donor = author.commit().unwrap().into_database();
    let owned = OwnedPackage::from_bytes(donor.as_bytes().to_vec()).unwrap();
    let payload = b"<?xml version=\"1.0\"?><report:document xmlns:report=\"urn:example:report\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\"><form:button/></report:document>";
    let donor = Database::from_bytes(
        rebuild_package(
            &owned,
            donor.content_xml(),
            vec![
                Addition {
                    path: "reports/source/content.xml".to_string(),
                    bytes: payload.to_vec(),
                    media_type: "text/xml".to_string(),
                },
                Addition {
                    path: "reports/source-two/content.xml".to_string(),
                    bytes: payload.to_vec(),
                    media_type: "text/xml".to_string(),
                },
            ],
            vec![
                (
                    "reports/source/".to_string(),
                    "application/vnd.sun.xml.report".to_string(),
                ),
                (
                    "reports/source-two/".to_string(),
                    "application/vnd.sun.xml.report".to_string(),
                ),
            ],
            Vec::<String>::new(),
            Vec::<String>::new(),
        )
        .unwrap(),
    )
    .unwrap();

    let destination = source();
    assert!(
        !donor
            .component_active_content(ComponentKind::Report, "payload-report")
            .unwrap()
            .is_empty()
    );
    let mut refused = destination.edit();
    assert!(
        refused
            .transfer_component_from_to(
                &donor,
                ComponentKind::Report,
                "payload-report",
                "reports/refused",
            )
            .is_err()
    );
    let mut transfer = destination.edit();
    transfer
        .transfer_component_from_to_with(
            &donor,
            ComponentKind::Report,
            "payload-report",
            "reports/imported",
            ActiveContentDisposition::CopyInert,
        )
        .unwrap();
    let commit = transfer.commit().unwrap();
    let received = commit.database();
    let catalog = received.catalog().unwrap();
    let component = catalog
        .components()
        .iter()
        .find(|component| component.name() == Some("payload-report"))
        .unwrap();
    assert_eq!(component.href(), Some("reports/imported"));
    let package = OwnedPackage::from_bytes(received.as_bytes().to_vec()).unwrap();
    assert_eq!(
        package.get_file("reports/imported/content.xml").unwrap(),
        payload
    );
    assert!(!package.has_file("reports/source/content.xml").unwrap());
    assert_eq!(
        commit.patch().inverse().apply(received).unwrap().as_bytes(),
        destination.as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .apply(&destination)
            .unwrap()
            .as_bytes(),
        received.as_bytes()
    );
    let durable_reopened = Database::from_bytes(
        commit
            .patch()
            .durable()
            .unwrap()
            .apply(&destination)
            .unwrap()
            .into_bytes(),
    )
    .unwrap();
    let durable_package = OwnedPackage::from_bytes(durable_reopened.into_bytes()).unwrap();
    assert_eq!(
        durable_package
            .get_file("reports/imported/content.xml")
            .unwrap(),
        payload
    );

    let mut second = destination.edit();
    second
        .transfer_component_from_to_with(
            &donor,
            ComponentKind::Report,
            "payload-report-two",
            "reports/imported-two",
            ActiveContentDisposition::CopyInert,
        )
        .unwrap();
    let second = second.commit().unwrap();
    let first = commit.patch().prepare("payload-one", limits()).unwrap();
    let second = second.patch().prepare("payload-two", limits()).unwrap();
    let mut joined = JoinedEdits::new(first.lineage().clone(), limits());
    joined.join(second).unwrap();
    joined.join(first).unwrap();
    let composed = Patch::compose(joined).unwrap().apply(&destination).unwrap();
    let package = OwnedPackage::from_bytes(composed.as_bytes().to_vec()).unwrap();
    assert_eq!(
        package.get_file("reports/imported/content.xml").unwrap(),
        payload
    );
    assert_eq!(
        package
            .get_file("reports/imported-two/content.xml")
            .unwrap(),
        payload
    );
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn active_content_inventory_is_inert_bounded_and_source_located() {
    const ACTIVE_SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><o:body><o:database><d:data-source/><o:scripts><s:script/><s:event-listener/></o:scripts></o:database></o:body></o:document-content>"#;
    use litchi_odf_common::{
        core::OwnedPackage,
        package::edit::{Addition, rebuild_package},
    };

    let database =
        Database::from_bytes(Builder::new().content_xml(ACTIVE_SOURCE).build().unwrap()).unwrap();
    let owned = OwnedPackage::from_bytes(database.as_bytes().to_vec()).unwrap();
    let database = Database::from_bytes(
        rebuild_package(
            &owned,
            database.content_xml(),
            vec![Addition {
                path: "Basic/Standard/Module1.xml".to_string(),
                bytes: b"<module/>".to_vec(),
                media_type: "text/xml".to_string(),
            }],
            vec![("Basic/".to_string(), String::new())],
            Vec::<String>::new(),
            Vec::<String>::new(),
        )
        .unwrap(),
    )
    .unwrap();
    let inventory = database.active_content().unwrap();
    assert!(inventory.entries().iter().any(|entry| {
        entry.kind() == ActiveContentKind::BasicMacro
            && entry.package_path() == "Basic/Standard/Module1.xml"
    }));
    assert!(inventory.entries().iter().any(|entry| {
        entry.kind() == ActiveContentKind::Script && entry.declaration_name() == Some("script")
    }));
    assert!(inventory.entries().iter().any(|entry| {
        entry.kind() == ActiveContentKind::EventListener
            && entry.declaration_name() == Some("event-listener")
    }));

    let capabilities = database.protection_capabilities();
    assert!(capabilities.can_verify_signatures());
    assert!(capabilities.can_remove_invalidated_signatures());
    assert!(!capabilities.can_re_sign());
    assert!(!capabilities.can_re_encrypt());
    assert_eq!(
        database.protection_support(ProtectionOperation::VerifySignatures),
        ProtectionSupport::Supported
    );
    assert_eq!(
        database.protection_support(ProtectionOperation::ReSign),
        ProtectionSupport::Unsupported
    );
    assert!(
        database
            .require_protection_operation(ProtectionOperation::ReSign)
            .is_err()
    );
    assert!(
        database
            .require_protection_operation(ProtectionOperation::ReEncrypt)
            .is_err()
    );
}

#[test]
#[expect(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
fn durable_history_and_join_span_the_inert_database_root() {
    let source = source();
    let mut edit = source.edit();
    edit.set_connection(Some(Connection::Resource(
        "sdbc:embedded:firebird".to_string(),
    )))
    .unwrap();
    edit.add_table(
        Table::new("people", TableKind::Definition)
            .with_column(Column::new("id"))
            .with_key(
                Key::new("people-primary", KeyKind::Primary).with_column(KeyColumn::new("id")),
            )
            .with_index(Index::new("people-id").with_column(IndexColumn::new("id"))),
    )
    .unwrap();
    edit.add_query(Query::new("people-query", "SELECT id FROM people"))
        .unwrap();
    edit.add_component(
        Component::new(ComponentKind::Form, "people-form")
            .with_href("Forms/people")
            .with_title("People"),
    )
    .unwrap();
    edit.add_component(
        Component::new(ComponentKind::Report, "people-report").with_href("Reports/people"),
    )
    .unwrap();
    edit.add_producer_extension(
        r#"<lo:settings xmlns:lo="urn:example:litchi:odb"><lo:mode lo:value="inert"/></lo:settings>"#,
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let kinds = commit
        .patch()
        .changes()
        .iter()
        .map(litchi_odb::Change::kind)
        .collect::<Vec<_>>();
    for expected in [
        ChangeKind::Connection,
        ChangeKind::Table,
        ChangeKind::Query,
        ChangeKind::Component,
        ChangeKind::ProducerExtension,
    ] {
        assert!(kinds.contains(&expected));
    }
    let durable = commit.patch().durable().unwrap();
    let applied = durable.apply(&source).unwrap();
    assert_eq!(applied.as_bytes(), commit.database().as_bytes());
    assert_eq!(
        durable.inverse().apply(&applied).unwrap().as_bytes(),
        source.as_bytes()
    );

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 64 * 1024 * 1024));
    history
        .record(applied.clone(), applied.as_bytes().len() as u64)
        .unwrap();
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), source.as_bytes());
    assert!(history.redo());
    assert_eq!(history.current().as_bytes(), applied.as_bytes());

    let independent = [
        {
            let mut edit = source.edit();
            edit.add_table(Table::new("joined-table", TableKind::Definition))
                .unwrap();
            edit.commit().unwrap().patch().clone()
        },
        query_patch(&source, "joined-query", "SELECT 7"),
        {
            let mut edit = source.edit();
            edit.set_connection(Some(Connection::File("file:///tmp/inert".to_string())))
                .unwrap();
            edit.commit().unwrap().patch().clone()
        },
        {
            let mut edit = source.edit();
            edit.add_component(Component::new(ComponentKind::Form, "joined-form"))
                .unwrap();
            edit.commit().unwrap().patch().clone()
        },
        {
            let mut edit = source.edit();
            edit.add_producer_extension(r#"<x:joined xmlns:x="urn:example:joined"/>"#)
                .unwrap();
            edit.commit().unwrap().patch().clone()
        },
    ];
    let mut joined = JoinedEdits::new(
        independent[0]
            .prepare("00-table", limits())
            .unwrap()
            .lineage()
            .clone(),
        limits(),
    );
    for (index, patch) in independent.iter().enumerate().rev() {
        joined
            .join(patch.prepare(format!("{index:02}"), limits()).unwrap())
            .unwrap();
    }
    let joined_database = Patch::compose(joined).unwrap().apply(&source).unwrap();
    let catalog = joined_database.catalog().unwrap();
    assert!(catalog.table("joined-table").unwrap().is_some());
    assert!(catalog.query("joined-query").unwrap().is_some());
    assert!(catalog.connection().is_some());
    assert!(
        catalog
            .components()
            .iter()
            .any(|component| component.name() == Some("joined-form"))
    );
    assert_eq!(joined_database.producer_extensions().unwrap().len(), 1);
}

#[test]
#[expect(
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "unexpected fixture failure"
)]
fn dependency_closed_transfer_spans_schema_resources_and_components() {
    let donor = source();
    let mut edit = donor.edit();
    edit.add_table(Table::new("parent", TableKind::Definition).with_column(Column::new("id")))
        .unwrap();
    edit.add_table(
        Table::new("child", TableKind::Definition)
            .with_column(Column::new("parent_id"))
            .with_key(
                Key::new("child-parent", KeyKind::Foreign)
                    .with_referenced_table("parent")
                    .with_column(KeyColumn::new("parent_id").with_related_column("id")),
            )
            .with_index(
                Index::new("child-parent-index").with_column(IndexColumn::new("parent_id")),
            ),
    )
    .unwrap();
    edit.set_connection(Some(Connection::Resource(
        "sdbc:firebird:inert".to_string(),
    )))
    .unwrap();
    edit.add_query(Query::new("child-query", "SELECT parent_id FROM child"))
        .unwrap();
    edit.add_component(Component::new(ComponentKind::Form, "child-form"))
        .unwrap();
    edit.add_component(Component::new(ComponentKind::Report, "child-report"))
        .unwrap();
    edit.add_producer_extension(r#"<x:resource xmlns:x="urn:example:transfer" x:id="one"/>"#)
        .unwrap();
    let donor = edit.commit().unwrap().into_database();

    let destination = source();
    let mut seed = destination.edit();
    seed.add_table(Table::new("child", TableKind::Definition))
        .unwrap();
    let destination = seed.commit().unwrap().into_database();
    let mut transfer = destination.edit();
    let before_refusal = transfer.staged_content_xml().to_owned();
    assert!(
        transfer
            .transfer_key_from(
                &donor,
                "child",
                "child-parent",
                DependencyDisposition::Refuse
            )
            .is_err()
    );
    assert_eq!(transfer.staged_content_xml(), before_refusal);
    assert!(transfer.changes().is_empty());
    transfer
        .transfer_column_from(&donor, "child", "parent_id")
        .unwrap();
    transfer
        .transfer_key_from(
            &donor,
            "child",
            "child-parent",
            DependencyDisposition::Cascade,
        )
        .unwrap();
    transfer
        .transfer_index_from(
            &donor,
            "child",
            "child-parent-index",
            DependencyDisposition::Refuse,
        )
        .unwrap();
    transfer.transfer_connection_from(&donor).unwrap();
    transfer.transfer_query_from(&donor, "child-query").unwrap();
    transfer
        .transfer_component_from(&donor, ComponentKind::Form, "child-form")
        .unwrap();
    transfer
        .transfer_component_from(&donor, ComponentKind::Report, "child-report")
        .unwrap();
    transfer
        .transfer_producer_extension_from(&donor, "urn:example:transfer", "resource")
        .unwrap();
    let received = transfer.commit().unwrap().into_database();
    let catalog = received.catalog().unwrap();
    assert!(catalog.table("parent").unwrap().is_some());
    let child = catalog.table("child").unwrap().unwrap();
    assert_eq!(child.columns().len(), 1);
    assert_eq!(child.keys().len(), 1);
    assert_eq!(child.indices().len(), 1);
    assert!(matches!(
        catalog.connection(),
        Some(Connection::Resource(_))
    ));
    assert_eq!(catalog.components().len(), 2);
    assert_eq!(received.producer_extensions().unwrap().len(), 1);

    let empty = source();
    let mut closed = empty.edit();
    closed
        .transfer_table_from_with(&donor, "child", DependencyDisposition::Cascade)
        .unwrap();
    let closed = closed.commit().unwrap().into_database();
    assert!(closed.catalog().unwrap().table("parent").unwrap().is_some());
    assert!(closed.catalog().unwrap().table("child").unwrap().is_some());
}

#[test]
#[expect(clippy::unwrap_used, reason = "unexpected fixture failure")]
fn multiple_genuine_libreoffice_base_packages_survive_changed_full_reopen() {
    let fixtures = [
        "../../3rdparty/libreoffice-core/extras/source/database/biblio.odb",
        "../../3rdparty/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb",
        "../../3rdparty/libreoffice-core/reportdesign/qa/unit/data/roundTrip.odb",
    ];
    for (index, relative) in fixtures.into_iter().enumerate() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(relative);
        let database = Database::open(path).unwrap();
        let query_name = format!("__litchi_inert_reopen_{index}");
        let mut edit = database.edit();
        edit.add_query(
            Query::new(&query_name, format!("SELECT {}", index + 1))
                .with_column(Column::new("value")),
        )
        .unwrap();
        let committed = edit.commit().unwrap();
        let reopened = Database::from_bytes(committed.database().as_bytes().to_vec()).unwrap();
        let catalog = reopened.catalog().unwrap();
        let query = catalog.query(&query_name).unwrap().unwrap();
        assert_eq!(query.columns().len(), 1);
        assert!(reopened.as_bytes().len() <= 256 * 1024 * 1024);
    }
}
