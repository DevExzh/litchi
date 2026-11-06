//! Native Excel external-data metadata stays inert and source-preserving.
use litchi_xlsx::{Package, SourceLimits};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_query_table_catalog_follows_both_worksheet_and_table_owners() {
    let bytes = native("excel-connections.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    let book = package.workbook();
    for (sheet, name, id) in [
        ("native-query-010", "ExternalData_1", 2),
        ("Sheet1", "LEGACY_CSV_010", 1),
    ] {
        let owner = book.sheet(sheet).unwrap().unwrap();
        let tables = owner.query_tables().unwrap();
        assert_eq!(tables.len(), 1, "{sheet}");
        let table = tables[0].query_table();
        assert_eq!(table.name(), name);
        assert_eq!(table.connection_id(), id);
        if id == 2 {
            let refresh = table.refresh().unwrap();
            assert_eq!(refresh.next_id(), Some(4));
            assert_eq!(
                refresh
                    .fields()
                    .iter()
                    .map(|f| (f.id(), f.name(), f.table_column_id()))
                    .collect::<Vec<_>>(),
                [
                    (1, Some("ID"), Some(1)),
                    (2, Some("Category"), Some(2)),
                    (3, Some("Amount"), Some(3))
                ]
            );
        }
    }
    assert!(package.to_bytes().unwrap() == bytes);
}
#[test]
fn native_connections_expose_text_import_and_inert_mashup_metadata() {
    let bytes = native("excel-connections.xlsx");
    let opc = Package::from_bytes(&bytes).unwrap().into_plain_opc();
    let snapshot =
        litchi_xlsx::connections::Snapshot::load(&opc, &SourceLimits::default()).unwrap();
    let connections = snapshot.connections().unwrap();
    assert_eq!(connections.connections.len(), 2);
    let text = connections.find(1).unwrap();
    assert_eq!(text.name.as_deref(), Some("native-query-010"));
    assert_eq!(text.connection_type, Some(6));
    let properties = text.text.as_ref().unwrap();
    assert_eq!(properties.code_page, Some(10000));
    assert_eq!(
        (properties.tab, properties.comma),
        (Some(false), Some(true))
    );
    assert_eq!(properties.fields.as_ref().unwrap().len(), 3);
    assert_eq!(
        properties.source_file.as_deref(),
        Some(
            "/Users/ryker/CodeProjects/litchi/test-data/microsoft-office/assets/native-query-010.csv"
        )
    );
    let query = connections.find(2).unwrap();
    assert_eq!(query.name.as_deref(), Some("Query - native-query-010"));
    assert_eq!(query.connection_type, Some(5));
    assert_eq!(query.keep_alive, Some(true));
    let db = query.database.as_ref().unwrap();
    assert_eq!(
        db.command.as_deref(),
        Some("SELECT * FROM [native-query-010]")
    );
    assert_eq!(
        db.connection,
        "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=native-query-010;Extended Properties=\"\""
    );
    assert_eq!(snapshot.query_table_parts().count(), 2);
}
#[test]
fn native_external_link_keeps_alternate_url_and_cached_sheets() {
    let opc = Package::from_bytes(&native("excel-links-objects.xlsx"))
        .unwrap()
        .into_plain_opc();
    let snapshot =
        litchi_xlsx::external_links::Snapshot::load(&opc, &SourceLimits::default()).unwrap();
    assert_eq!(snapshot.entries().len(), 1);
    let litchi_xlsx::external_links::Link::Workbook(book) = &snapshot.entries()[0].link else {
        panic!("workbook link")
    };
    assert_eq!(book.target.target, "excel-values-formulas-chart.xlsx");
    assert_eq!(book.sheet_names, ["Sales chart", "Sheet1", "Sheet2"]);
    let urls = book.alternate_urls().unwrap();
    assert!(urls.relative_url.is_none());
    assert_eq!(urls.absolute_url.as_ref().unwrap().relationship_id, "rId2");
    assert_eq!(
        urls.absolute_url.as_ref().unwrap().target,
        "/Users/ryker/CodeProjects/litchi/test-data/microsoft-office/excel/excel-values-formulas-chart.xlsx"
    );
    let caches = book
        .cached_sheets
        .iter()
        .flat_map(|s| {
            s.rows.iter().flat_map(move |r| {
                r.cells.iter().map(move |c| {
                    (
                        s.sheet_id,
                        r.row,
                        c.reference.as_deref(),
                        c.raw_value.as_deref(),
                    )
                })
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(
        caches,
        [
            (1, 6, Some("D6"), Some("49.75")),
            (2, 4, Some("B4"), Some("14"))
        ]
    );
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
fn preserves_other_members(before: &[u8], after: &[u8], changed: &str) {
    let (before, after) = (members(before), members(after));
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, bytes) in before {
        if name != changed {
            assert!(after[&name] == bytes, "changed {name}");
        }
    }
}
#[test]
fn native_query_table_rename_keeps_table_bindings_revisions_and_legacy_import() {
    let bytes = native("excel-connections.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    for (sheet, part, owner) in [
        (
            "native-query-010",
            "xl/queryTables/queryTable1.xml",
            "/xl/tables/table1.xml",
        ),
        (
            "Sheet1",
            "xl/queryTables/queryTable2.xml",
            "/xl/worksheets/sheet2.xml",
        ),
    ] {
        let mut edit = package.edit_query_tables(sheet).unwrap();
        let entry = edit.query_tables().unwrap().remove(0);
        assert_eq!(entry.owner_part_name(), Some(owner));
        let mut changed = entry.query_table().clone();
        changed.set_name("NATIVE_QUERY_RENAMED_020".into());
        edit.replace_part(
            entry.part_name(),
            changed.clone(),
            litchi_xlsx::query_table::Conformance::Transitional,
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("query-renamed.xlsx");
        commit.package().save(&path).unwrap();
        let saved = std::fs::read(path).unwrap();
        let reopened = Package::from_bytes(&saved).unwrap();
        assert_eq!(
            reopened
                .workbook()
                .sheet(sheet)
                .unwrap()
                .unwrap()
                .query_tables()
                .unwrap()[0]
                .query_table(),
            &changed
        );
        preserves_other_members(&bytes, &saved, part);
        let before = String::from_utf8(members(&bytes)[part].clone()).unwrap();
        let expected = before.replace(entry.query_table().name(), "NATIVE_QUERY_RENAMED_020");
        assert!(
            members(&saved)[part] == expected.as_bytes(),
            "query rename changed unrelated source markup"
        );
        assert_restored_members(
            &bytes,
            &(reopened
                .apply_query_tables(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()),
        );
    }
    assert!(package.to_bytes().unwrap() == bytes);
}
#[test]
fn native_connection_description_edit_keeps_import_data_and_mashup_payloads() {
    let bytes = native("excel-connections.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_connections().unwrap();
    edit.edit(2, |connection| {
        connection.description = Some("NATIVE_DESCRIPTION_020".into());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let saved = commit.package().to_bytes().unwrap();
    preserves_other_members(&bytes, &saved, "xl/connections.xml");
    let reopened = Package::from_bytes(&saved).unwrap();
    let snapshot = litchi_xlsx::connections::Snapshot::load(
        &reopened.clone().into_plain_opc(),
        &SourceLimits::default(),
    )
    .unwrap();
    assert_eq!(
        snapshot
            .connections()
            .unwrap()
            .find(2)
            .unwrap()
            .description
            .as_deref(),
        Some("NATIVE_DESCRIPTION_020")
    );
    assert_eq!(snapshot.query_table_parts().count(), 2);
    assert_restored_members(
        &bytes,
        &(reopened
            .apply_connections(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()),
    );
}
#[test]
fn native_alternate_url_metadata_edit_preserves_both_inert_targets_and_caches() {
    let bytes = native("excel-links-objects.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    let snapshot = litchi_xlsx::external_links::Snapshot::load(
        &package.clone().into_plain_opc(),
        &SourceLimits::default(),
    )
    .unwrap();
    let litchi_xlsx::external_links::Link::Workbook(book) = &snapshot.entries()[0].link else {
        panic!("workbook link")
    };
    let mut urls = book.alternate_urls.clone().unwrap();
    urls.drive_id = Some("NATIVE_DRIVE_020".into());
    let mut edit = package.edit_external_links().unwrap();
    edit.set_alternate_urls(0, Some(urls.clone())).unwrap();
    let commit = edit.commit().unwrap();
    let saved = commit.package().to_bytes().unwrap();
    preserves_other_members(&bytes, &saved, "xl/externalLinks/externalLink1.xml");
    let reopened = Package::from_bytes(&saved).unwrap();
    let after = litchi_xlsx::external_links::Snapshot::load(
        &reopened.clone().into_plain_opc(),
        &SourceLimits::default(),
    )
    .unwrap();
    let mut expected = book.clone();
    expected.alternate_urls = Some(urls);
    assert_eq!(
        after.entries()[0].link,
        litchi_xlsx::external_links::Link::Workbook(expected)
    );
    assert_restored_members(
        &bytes,
        &(reopened
            .apply_external_links(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()),
    );
}

fn assert_restored_members(before: &[u8], after: &[u8]) {
    let (before_members, after_members) = (members(before), members(after));
    let differences = before_members
        .iter()
        .filter(|(name, data)| after_members.get(*name) != Some(*data))
        .map(|(name, _)| name)
        .collect::<Vec<_>>();

    assert_eq!(
        before_members.keys().collect::<Vec<_>>(),
        after_members.keys().collect::<Vec<_>>()
    );
    assert!(
        differences.is_empty(),
        "inverse changed members: {differences:?}"
    );
}

#[test]
fn native_connection_owner_creation_removal_and_inverse_preserve_original_members() {
    let donor = Package::from_bytes(&native("excel-connections.xlsx"))
        .unwrap()
        .into_plain_opc();
    let snapshot =
        litchi_xlsx::connections::Snapshot::load(&donor, &SourceLimits::default()).unwrap();
    let mut connection = snapshot.connections().unwrap().find(1).unwrap().clone();
    connection.id = 7;
    connection.name = Some("NATIVE_INERT_CONNECTION_020".into());
    let bytes = native("excel-links-objects.xlsx");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit_connections().unwrap();
    assert!(edit.connections().unwrap().is_none());
    edit.set(connection.clone()).unwrap();
    let created = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("inert-connection.xlsx");
    created.package().save(&path).unwrap();
    let saved = std::fs::read(path).unwrap();
    let reopened = Package::from_bytes(&saved).unwrap();
    let snapshot = litchi_xlsx::connections::Snapshot::load(
        &reopened.clone().into_plain_opc(),
        &SourceLimits::default(),
    )
    .unwrap();
    assert_eq!(snapshot.connections().unwrap().find(7), Some(&connection));
    let before = members(&bytes);
    let after = members(&saved);
    for (name, data) in &before {
        if name != "[Content_Types].xml" && name != "xl/_rels/workbook.xml.rels" {
            assert!(after.get(name) == Some(data), "changed {name}");
        }
    }
    let restored = reopened
        .apply_connections(&created.patch().inverse())
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    assert_restored_members(&bytes, &restored);
    let mut edit = reopened.edit_connections().unwrap();
    assert!(edit.remove(7).unwrap().is_some());
    let removed = edit.commit().unwrap();
    let removal_bytes = removed.package().to_bytes().unwrap();
    let removed_reopened = Package::from_bytes(&removal_bytes).unwrap();
    assert!(
        litchi_xlsx::connections::Snapshot::load(
            &removed_reopened.clone().into_plain_opc(),
            &SourceLimits::default()
        )
        .unwrap()
        .connections()
        .is_none()
    );
    for (name, data) in before {
        if name != "[Content_Types].xml" && name != "xl/_rels/workbook.xml.rels" {
            assert!(
                members(&removal_bytes).get(&name) == Some(&data),
                "changed {name}"
            );
        }
    }
    let restored = removed_reopened
        .apply_connections(&removed.patch().inverse())
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    assert_restored_members(&saved, &restored);
    assert!(source.to_bytes().unwrap() == bytes);
}
