use super::{Condition, ConditionSource, Expression, Key, Order, Range, Rule, Rules, Sort, Source};
use crate::{Builder, MutableSpreadsheet, Spreadsheet};
use std::io::{Cursor, Read, Write};

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:v="urn:example:vendor"><office:body><office:spreadsheet><t:table t:name="Input"/><t:database-ranges><t:database-range t:name="Sales" t:target-range-address="Input.A1:Input.B20"><t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"/><t:filter t:condition-source="self"><t:filter-condition t:field-number="0" t:value="East" t:operator="="/></t:filter><t:sort><t:sort-by t:field-number="1" t:order="descending"/></t:sort><t:subtotal-rules><t:subtotal-rule t:group-by-field-number="0"><t:subtotal-field t:field-number="1" t:function="sum"/></t:subtotal-rule></t:subtotal-rules></t:database-range></t:database-ranges><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

const WITHOUT_OWNER: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><t:table t:name="Input"/><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

fn package(xml: &str) -> Vec<u8> {
    Builder::new()
        .content_xml(xml)
        .build()
        .expect("test fixture or operation should succeed")
}

fn signed_package(xml: &str) -> Vec<u8> {
    let source = package(xml);
    let mut input = zip::ZipArchive::new(Cursor::new(source)).expect("fixture ZIP should open");
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for index in 0..input.len() {
            let mut entry = input.by_index(index).expect("fixture member should exist");
            let name = entry.name().to_string();
            let mut bytes = Vec::new();
            entry
                .read_to_end(&mut bytes)
                .expect("fixture member should read");
            writer
                .start_file(name, options)
                .expect("fixture member should start");
            writer
                .write_all(&bytes)
                .expect("fixture member should write");
        }
        writer
            .start_file("META-INF/documentsignatures.xml", options)
            .expect("signature member should start");
        writer
            .write_all(b"<document-signatures/>")
            .expect("signature member should write");
        writer.finish().expect("fixture ZIP should finish");
    }
    output.into_inner()
}

fn range(name: &str) -> Range {
    let mut range = Range::new("Input.A1:Input.B20");
    range.name = Some(name.to_string());
    range.source = Some(Source::Query {
        database_name: "sales.odb".to_string(),
        query_name: "OpenOrders".to_string(),
    });
    range.filter = Some(super::Filter {
        target_range_address: None,
        condition_source: Some(ConditionSource::SelfContained),
        condition_source_range_address: None,
        display_duplicates: None,
        expression: Expression::Condition(Condition::new(0, "=", "East")),
    });
    range.sort = Some(Sort {
        keys: vec![Key {
            field_number: 1,
            data_type: Some("number".to_string()),
            order: Some(Order::Descending),
        }],
        ..Sort::default()
    });
    range.subtotals = Some(Rules {
        rules: vec![Rule {
            group_by_field_number: 0,
            fields: vec![super::Field {
                field_number: 1,
                function: "sum".to_string(),
            }],
        }],
        ..Rules::default()
    });
    range
}

fn known_parent_fixture(parent: &str, child: &str) -> String {
    let parent_attributes = if parent == "database-range" {
        r#" t:target-range-address="Input.A1""#
    } else {
        ""
    };
    let nested = format!(
        r#"<t:{parent}{parent_attributes}>{child}</t:{parent}>"#,
        parent = parent,
        parent_attributes = parent_attributes,
        child = child,
    );
    if parent == "database-ranges" {
        WITHOUT_OWNER.replace("<t:shapes/>", &format!("{nested}<t:shapes/>"))
    } else {
        WITHOUT_OWNER.replace(
            "<t:shapes/>",
            &format!(
                r#"<t:database-ranges><t:database-range t:target-range-address="Input.A1">{nested}</t:database-range></t:database-ranges><t:shapes/>"#
            ),
        )
    }
}

fn assert_contextual_change_is_refused(xml: &str) {
    let location = super::codec::locate(xml).expect("context fixture should be well formed");
    assert!(
        location.opaque,
        "known child context must be admitted as opaque"
    );

    let bytes = package(xml);
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).expect("outer package should open");
    let catalog = spreadsheet.database_ranges();
    if catalog.is_err() {
        // A typed parser may refuse a malformed expression outright.  That is
        // safe admission; the regression is specifically that a successful
        // read must never permit a lossy changed rewrite.
        return;
    }

    let mut noop = MutableSpreadsheet::from_bytes(bytes.clone()).expect("package should open");
    let before = noop.spreadsheet().content_xml().to_owned();
    noop.edit_database_ranges(|_| Ok(()))
        .expect("opaque no-op should preserve source");
    assert_eq!(noop.spreadsheet().content_xml(), before);

    let mut changed = MutableSpreadsheet::from_bytes(bytes).expect("package should open");
    let before = changed.spreadsheet().content_xml().to_owned();
    let result = changed.edit_database_ranges(|editor| {
        editor.update("Sales", |range| {
            range.name = Some("Changed".to_string());
            Ok(())
        })
    });
    assert!(
        result.is_err(),
        "opaque context must refuse a changed rewrite"
    );
    assert_eq!(changed.spreadsheet().content_xml(), before);
}

#[test]
fn catalog_reads_owned_filter_sort_and_subtotal_metadata() {
    let bytes = package(XML);
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).expect("fixture should open");
    let catalog = spreadsheet.database_ranges().expect("catalog should parse");
    assert!(catalog.has_owner());
    assert_eq!(catalog.len(), 1);
    assert_eq!(
        catalog
            .named("Sales")
            .expect("selector should work")
            .unwrap()
            .sort
            .as_ref()
            .unwrap()
            .keys[0]
            .order,
        Some(Order::Descending)
    );
    assert_eq!(
        catalog.ranges()[0].subtotals.as_ref().unwrap().rules[0].fields[0].function,
        "sum"
    );
    let commit = catalog.transaction().commit().expect("no-op should commit");
    assert!(!commit.changed());
    assert_eq!(commit.bytes(), bytes.as_slice());
}

#[test]
fn clone_staged_crud_inserts_and_removes_owner_atomically() {
    let mut mutable = MutableSpreadsheet::from_bytes(package(WITHOUT_OWNER)).expect("open");
    mutable
        .edit_database_ranges(|editor| editor.add(range("Sales")))
        .expect("insert should succeed");
    let content = mutable.spreadsheet().content_xml();
    assert!(content.contains("database-ranges"));
    assert!(content.find("database-ranges").unwrap() < content.find("t:shapes").unwrap());
    mutable
        .edit_database_ranges(|editor| {
            let removed = editor.remove("Sales")?;
            assert_eq!(removed.name.as_deref(), Some("Sales"));
            Ok(())
        })
        .expect("remove should succeed");
    assert!(!mutable.database_ranges().unwrap().has_owner());
}

#[test]
fn insertion_keeps_database_ranges_after_tracked_changes() {
    let xml = WITHOUT_OWNER.replace("<t:shapes/>", "<t:tracked-changes/><t:shapes/>");
    let mut mutable = MutableSpreadsheet::from_bytes(package(&xml)).expect("open");
    mutable
        .edit_database_ranges(|editor| editor.add(range("Sales")))
        .expect("insertion should succeed");
    let content = mutable.spreadsheet().content_xml();
    assert!(content.find("t:tracked-changes").unwrap() < content.find("database-ranges").unwrap());
    assert!(content.find("database-ranges").unwrap() < content.find("t:shapes").unwrap());
}

#[test]
fn builder_stages_database_range_without_external_execution() {
    let mut builder = Builder::new().content_xml(WITHOUT_OWNER);
    builder
        .add_database_range(range("Sales"))
        .expect("builder insertion should succeed");
    let spreadsheet = Spreadsheet::from_bytes(builder.build().expect("build should succeed"))
        .expect("reopen should succeed");
    let catalog = spreadsheet.database_ranges().expect("catalog should parse");
    assert_eq!(
        catalog.named("Sales").unwrap().unwrap().source,
        range("Sales").source
    );
}

#[test]
fn explicit_empty_owner_is_distinct_from_absence() {
    let xml = WITHOUT_OWNER.replace(
        "<t:shapes/>",
        "<t:database-ranges/><!-- retained --><t:shapes/>",
    );
    let spreadsheet = Spreadsheet::from_bytes(package(&xml)).expect("open");
    let catalog = spreadsheet.database_ranges().expect("catalog");
    assert!(catalog.has_owner());
    assert!(catalog.is_empty());
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet);
    mutable
        .edit_database_ranges(|editor| editor.add(range("Sales")))
        .expect("replace empty owner");
    mutable
        .edit_database_ranges(|editor| editor.clear())
        .expect("remove owner");
    assert!(!mutable.database_ranges().unwrap().has_owner());
}

#[test]
fn inherited_namespace_on_empty_owner_is_injected_before_close_token() {
    let xml = WITHOUT_OWNER.replace("<t:shapes/>", "<t:database-ranges/><t:shapes/>");
    let location = super::codec::locate(&xml).expect("host should be located");
    let owner = location
        .container
        .as_ref()
        .expect("owner should be located");
    let fragment = super::codec::owner_fragment(&xml, owner).expect("owner should be sliced");
    assert_eq!(
        fragment,
        "<t:database-ranges xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"/>"
    );
    assert!(
        crate::model::database_range::parse_database_ranges(&fragment)
            .unwrap()
            .is_empty()
    );
}

#[test]
fn owner_fragment_ignores_quoted_greater_than_in_attributes() {
    let xml = WITHOUT_OWNER.replace(
        "<t:shapes/>",
        r#"<t:database-ranges t:future="quoted > marker"/><t:shapes/>"#,
    );
    let location = super::codec::locate(&xml).expect("host should be located");
    let owner = location
        .container
        .as_ref()
        .expect("owner should be located");
    let fragment = super::codec::owner_fragment(&xml, owner).expect("owner should be sliced");
    assert!(fragment.contains("quoted > marker"));
    assert!(crate::model::database_range::parse_database_ranges(&fragment).is_ok());
}

#[test]
fn owner_fragment_does_not_treat_quoted_namespace_text_as_a_declaration() {
    let xml = WITHOUT_OWNER.replace(
        "<t:shapes/>",
        r#"<t:database-ranges t:future="value xmlns:t=not-a-declaration"/><t:shapes/>"#,
    );
    let location = super::codec::locate(&xml).expect("host should be located");
    let owner = location
        .container
        .as_ref()
        .expect("owner should be located");
    let fragment = super::codec::owner_fragment(&xml, owner).expect("owner should be sliced");
    assert!(fragment.contains("xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\""));
    assert!(crate::model::database_range::parse_database_ranges(&fragment).is_ok());
}

#[test]
fn unknown_owner_markup_blocks_changed_edit_but_noop_is_byte_exact() {
    let xml = XML.replace(
        "</t:database-range>",
        "<v:future v:flag=\"keep\"/><!-- future --></t:database-range>",
    );
    let mut mutable = MutableSpreadsheet::from_bytes(package(&xml)).expect("open");
    let before = mutable.spreadsheet().content_xml().to_owned();
    mutable
        .edit_database_ranges(|_| Ok(()))
        .expect("no-op should preserve opaque owner");
    assert_eq!(mutable.spreadsheet().content_xml(), before);
    let error = mutable.edit_database_ranges(|editor| {
        editor.update("Sales", |range| {
            range.name = Some("Changed".to_string());
            Ok(())
        })
    });
    assert!(error.is_err());
    assert_eq!(mutable.spreadsheet().content_xml(), before);
}

#[test]
fn opaque_source_leaf_comment_allows_exact_noop_but_not_changed_rewrite() {
    let xml = XML.replace(
        r#"<t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"/>"#,
        r#"<t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"><!-- retained --></t:database-source-query>"#,
    );
    let mut mutable = MutableSpreadsheet::from_bytes(package(&xml)).expect("open");
    let before = mutable.spreadsheet().content_xml().to_owned();
    mutable
        .edit_database_ranges(|_| Ok(()))
        .expect("opaque no-op should preserve source");
    assert_eq!(mutable.spreadsheet().content_xml(), before);

    let result = mutable.edit_database_ranges(|editor| {
        editor.update("Sales", |range| {
            range.name = Some("Changed".to_string());
            Ok(())
        })
    });
    assert!(result.is_err());
    assert_eq!(mutable.spreadsheet().content_xml(), before);
}

#[test]
fn contextual_read_refuses_known_children_nested_under_source_leaves() {
    let xml = XML.replace(
        r#"<t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"/>"#,
        r#"<t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"><t:database-source-table t:database-name="sales.odb" t:database-table-name="nested"/></t:database-source-query>"#,
    );
    let spreadsheet = Spreadsheet::from_bytes(package(&xml)).expect("outer package should open");
    assert!(spreadsheet.database_ranges().is_err());
}

#[test]
fn every_known_parent_context_marks_wrong_table_children_opaque() {
    // Keep one adversarial expanded-name edge for each known owner/child
    // context.  The public parser may reject some of these immediately, but
    // the source locator must classify all of them as opaque before any typed
    // graph can silently discard a known element.
    let cases = [
        ("database-ranges", r#"<t:sort-by/>"#),
        ("database-range", r#"<t:filter-set-item/>"#),
        ("database-source-sql", r#"<t:filter/>"#),
        ("database-source-table", r#"<t:filter/>"#),
        ("database-source-query", r#"<t:filter/>"#),
        ("filter", r#"<t:sort-by/>"#),
        ("filter-condition", r#"<t:sort/>"#),
        ("filter-and", r#"<t:filter-and/>"#),
        ("filter-or", r#"<t:filter-or/>"#),
        ("filter-set-item", r#"<t:filter/>"#),
        ("sort", r#"<t:filter-condition/>"#),
        ("sort-by", r#"<t:sort/>"#),
        ("subtotal-rules", r#"<t:filter-condition/>"#),
        ("sort-groups", r#"<t:filter-condition/>"#),
        ("subtotal-rule", r#"<t:filter-condition/>"#),
        ("subtotal-field", r#"<t:filter/>"#),
    ];
    for (parent, child) in cases {
        let xml = known_parent_fixture(parent, child);
        let location = super::codec::locate(&xml).expect("context fixture should be located");
        assert!(
            location.opaque,
            "wrong table:{child} under table:{parent} was not opaque"
        );
    }
}

#[test]
fn valid_expanded_name_parent_contexts_remain_editable() {
    let xml = XML.replace(
        r#"<t:filter-condition t:field-number="0" t:value="East" t:operator="="/>"#,
        r#"<t:filter-and><t:filter-condition t:field-number="0" t:value="East" t:operator="="><t:filter-set-item t:value="East"/></t:filter-condition><t:filter-or><t:filter-condition t:field-number="1" t:value="West" t:operator="="/></t:filter-or></t:filter-and>"#,
    );
    let location = super::codec::locate(&xml).expect("valid context fixture should be located");
    assert!(!location.opaque);
    let spreadsheet = Spreadsheet::from_bytes(package(&xml)).expect("valid package should open");
    assert!(spreadsheet.database_ranges().is_ok());
}

#[test]
fn known_children_in_wrong_parents_never_silently_drop_on_public_edit() {
    let cases = [
        (
            "sort-by under filter",
            XML.replace(
                r#"<t:filter-condition t:field-number="0" t:value="East" t:operator="="/>"#,
                r#"<t:sort-by t:field-number="0"/>"#,
            ),
        ),
        (
            "filter-condition under sort",
            XML.replace(
                r#"<t:sort><t:sort-by t:field-number="1" t:order="descending"/></t:sort>"#,
                r#"<t:sort><t:filter-condition t:field-number="1" t:value="East" t:operator="="/></t:sort>"#,
            ),
        ),
        (
            "filter-condition under subtotal-rule",
            XML.replace(
                r#"<t:subtotal-field t:field-number="1" t:function="sum"/>"#,
                r#"<t:filter-condition t:field-number="1" t:value="East" t:operator="="/>"#,
            ),
        ),
        (
            "nested database-range",
            XML.replace(
                r#"<t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"/>"#,
                r#"<t:database-range t:target-range-address="Input.C1"/>"#,
            ),
        ),
    ];
    for (label, xml) in cases {
        assert_contextual_change_is_refused(&xml);
        assert!(
            super::codec::locate(&xml)
                .expect("context fixture should be located")
                .opaque,
            "{label} should be opaque"
        );
    }
}

#[test]
fn signed_database_range_noop_is_exact_but_changed_publication_is_refused() {
    let bytes = signed_package(XML);
    let mut noop = MutableSpreadsheet::from_bytes(bytes.clone()).expect("signed source opens");
    noop.edit_database_ranges(|_| Ok(()))
        .expect("signed no-op should be accepted");
    assert_eq!(noop.to_bytes(), bytes);

    let mut changed = MutableSpreadsheet::from_bytes(bytes).expect("signed source opens");
    let before = changed.spreadsheet().content_xml().to_owned();
    let result = changed.edit_database_ranges(|editor| {
        editor.update("Sales", |range| {
            range.target_range_address = "Input.A1:Input.C1".to_string();
            Ok(())
        })
    });
    assert!(result.is_err());
    assert_eq!(changed.spreadsheet().content_xml(), before);
}

#[test]
fn facade_rejects_overbudget_owner_before_materializing_snapshot_graph() {
    let oversized_target = "x".repeat(1_048_577);
    let xml = WITHOUT_OWNER.replace(
        "<t:shapes/>",
        &format!(
            "<t:database-ranges><t:database-range t:target-range-address=\"{oversized_target}\"/></t:database-ranges><t:shapes/>"
        ),
    );
    let spreadsheet = Spreadsheet::from_bytes(package(&xml)).expect("facade should open source");
    assert!(spreadsheet.database_range_snapshot().is_err());
}

#[test]
fn owned_snapshot_patch_inverse_and_conflict_are_atomic() -> litchi_core::Result<()> {
    let snapshot = crate::database_range::Snapshot::from_bytes(package(WITHOUT_OWNER))?;
    let mut edit = snapshot.edit();
    edit.editor().add(range("Sales"))?;
    let commit = edit.commit()?;
    assert!(commit.changed());
    assert!(commit.snapshot().has_owner());
    assert_eq!(commit.snapshot().ranges()[0].name.as_deref(), Some("Sales"));
    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), snapshot.as_bytes());
    let other = crate::database_range::Snapshot::from_bytes(package(XML))?;
    assert!(commit.patch().apply(&other).is_err());
    Ok(())
}

#[test]
fn failed_typed_update_leaves_staged_catalog_unchanged() {
    let spreadsheet = Spreadsheet::from_bytes(package(WITHOUT_OWNER)).expect("open");
    let catalog = spreadsheet.database_ranges().expect("catalog");
    let mut transaction = catalog.transaction();
    let result = transaction.editor().update(0, |range| {
        range.target_range_address.clear();
        Ok(())
    });
    assert!(result.is_err());
    assert!(transaction.ranges().is_empty());
}
