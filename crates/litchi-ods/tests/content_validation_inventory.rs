use litchi_ods::{Builder, Spreadsheet, content_validation};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn document(inner: &str) -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}"><o:body><o:spreadsheet>{inner}</o:spreadsheet></o:body></o:document-content>"#
    )
}

#[test]
fn inventories_namespace_owned_catalog_and_compact_repeated_bindings() {
    let xml = document(
        r#"<t:content-validations><t:content-validation t:name="whole" t:condition="cell-content-is-whole-number()" t:allow-empty-cell="true" t:display-list="unsorted"><t:help-message><x:p>Whole</x:p></t:help-message></t:content-validation></t:content-validations><t:table t:name="Data"><t:table-row t:number-rows-repeated="1000000"><t:table-cell t:number-columns-repeated="500000" t:content-validation-name="whole"/></t:table-row></t:table>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&xml).unwrap();
    assert_eq!(snapshot.source_xml().as_ptr(), xml.as_ptr());
    assert_eq!(snapshot.definitions().len(), 1);
    assert_eq!(snapshot.definitions()[0].name(), "whole");
    assert_eq!(
        snapshot.definitions()[0].display_list(),
        Some(content_validation::DisplayList::Unsorted)
    );
    let sheet = snapshot.sheet("Data").unwrap();
    assert_eq!(sheet.logical_row_count(), 1_000_000);
    assert_eq!(sheet.bindings().len(), 1);
    let binding = &sheet.bindings()[0];
    assert_eq!((binding.row(), binding.column()), (0, 0));
    assert_eq!(
        (binding.row_count(), binding.column_count()),
        (1_000_000, 500_000)
    );
    assert_eq!(binding.definition_index(), Some(0));
    assert!(snapshot.definitions()[0].has_opaque_content());
    assert!(!snapshot.has_complete_reference_closure());
}

#[test]
fn facade_reads_bindings_that_the_legacy_worksheet_projection_drops() {
    let xml = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:content-validations><table:content-validation table:name="v"/></table:content-validations><table:table table:name="Sheet"><table:table-row><table:table-cell table:content-validation-name="v"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#,
    );
    let bytes = Builder::new().content_xml(xml).build().unwrap();
    let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let inventory = spreadsheet.content_validations().unwrap();
    assert_eq!(inventory.sheet("Sheet").unwrap().bindings().len(), 1);
    assert_eq!(
        inventory.sheet("Sheet").unwrap().bindings()[0].validation_name(),
        "v"
    );
}

#[test]
fn retains_opaque_owner_source_and_reports_dangling_references() {
    let xml = document(
        r#"<t:content-validations><!--keep--><t:content-validation t:name="known"><?keep data?><vendor:future xmlns:vendor="urn:vendor"><vendor:item/></vendor:future></t:content-validation></t:content-validations><t:table t:name="Sheet"><t:table-row><t:table-cell t:content-validation-name="missing"/></t:table-row></t:table>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&xml).unwrap();
    assert_eq!(snapshot.source_xml(), xml);
    assert!(snapshot.has_opaque_catalog_content());
    assert!(snapshot.definitions()[0].has_opaque_content());
    assert_eq!(snapshot.dangling_binding_count(), 1);
    assert!(snapshot.sheets()[0].bindings()[0].is_dangling());
    assert!(!snapshot.has_complete_reference_closure());
}

#[test]
fn opaque_and_active_xml_are_conservative_and_never_interpreted() {
    let cdata = document(
        r#"<t:content-validations><t:content-validation t:name="v"><![CDATA[opaque]]></t:content-validation></t:content-validations>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&cdata).unwrap();
    assert!(snapshot.definitions()[0].has_opaque_content());
    let foreign_attribute = document(
        r#"<t:content-validations><t:content-validation t:name="v" x:name="vendor"/></t:content-validations>"#,
    );
    assert!(
        content_validation::Snapshot::parse(&foreign_attribute)
            .unwrap()
            .definitions()[0]
            .has_opaque_content()
    );
    let dtd = format!(
        r#"<!DOCTYPE o:document-content [<!ENTITY e "x">]><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:content-validations><t:content-validation t:name="v">&e;</t:content-validation></t:content-validations></o:spreadsheet></o:body></o:document-content>"#
    );
    assert!(content_validation::Snapshot::parse(&dtd).is_err());
    let text = document(
        r#"<t:content-validations><t:content-validation t:name="v">garbage</t:content-validation></t:content-validations>"#,
    );
    assert!(
        content_validation::Snapshot::parse(&text)
            .unwrap()
            .definitions()[0]
            .has_opaque_content()
    );
}

#[test]
fn rejects_illegal_ownership_spoofs_mce_and_expanded_attribute_duplicates() {
    let misplaced = document(r#"<t:table t:name="Sheet"><t:content-validations/></t:table>"#);
    assert!(content_validation::Snapshot::parse(&misplaced).is_err());

    let vendor_collision = document(
        r#"<x:content-validations><x:content-validation/></x:content-validations><t:content-validations><t:content-validation t:name="v"/></t:content-validations>"#,
    );
    assert!(content_validation::Snapshot::parse(&vendor_collision).is_ok());

    let mce = document(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent>"#,
    );
    assert_eq!(
        content_validation::Snapshot::parse(&mce),
        Err(content_validation::Error::UnsupportedMarkupCompatibility)
    );

    let duplicates = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:u="{TABLE}"><o:body><o:spreadsheet><t:content-validations><t:content-validation t:name="one" u:name="two"/></t:content-validations></o:spreadsheet></o:body></o:document-content>"#
    );
    assert!(content_validation::Snapshot::parse(&duplicates).is_err());

    let duplicate_catalogs =
        document(r#"<t:content-validations/><t:content-validations/><t:table t:name="Sheet"/>"#);
    assert!(content_validation::Snapshot::parse(&duplicate_catalogs).is_err());

    let misplaced_binding = document(
        r#"<t:table t:name="Sheet"><t:table-row t:content-validation-name="v"/></t:table>"#,
    );
    assert!(content_validation::Snapshot::parse(&misplaced_binding).is_err());
}

#[test]
fn returns_typed_failures_at_exact_resource_boundaries_without_expanding_runs() {
    let xml = document(
        r#"<t:content-validations><t:content-validation t:name="v"/></t:content-validations><t:table t:name="Sheet"><t:table-row t:number-rows-repeated="8"><t:table-cell t:number-columns-repeated="9" t:content-validation-name="v"/></t:table-row></t:table>"#,
    );
    assert!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default()
                .with_logical_rows(8)
                .with_logical_columns(9)
                .with_definitions(1)
                .with_sheets(1)
                .with_bindings(1),
        )
        .is_ok()
    );
    assert!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_attributes(8),
        )
        .is_ok()
    );
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_attributes(7),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Attributes,
            observed: 8,
            maximum: 7,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_logical_rows(7),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::LogicalRows,
            observed: 8,
            maximum: 7,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_bindings(0),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Bindings,
            maximum: 0,
            ..
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_input_bytes(xml.len() - 1),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::InputBytes,
            ..
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_logical_columns(8),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::LogicalColumns,
            observed: 9,
            maximum: 8,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_definitions(0),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Definitions,
            observed: 1,
            maximum: 0,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_depth(1),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Depth,
            observed: 2,
            maximum: 1,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_events(1),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Events,
            observed: 2,
            maximum: 1,
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_attributes(0),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::Attributes,
            maximum: 0,
            ..
        })
    ));
    assert!(matches!(
        content_validation::Snapshot::parse_with_limits(
            &xml,
            content_validation::Limits::default().with_text_bytes(0),
        ),
        Err(content_validation::Error::LimitExceeded {
            kind: content_validation::LimitKind::TextBytes,
            ..
        })
    ));
}

#[test]
fn enforces_catalog_cardinality_order_and_skips_schema_valid_dde_cache_tables() {
    assert!(content_validation::Snapshot::parse(&document("<t:content-validations/>")).is_err());
    let late = document(
        r#"<t:table t:name="Sheet"/><t:content-validations><t:content-validation t:name="v"/></t:content-validations>"#,
    );
    assert!(content_validation::Snapshot::parse(&late).is_err());
    let late_after_labels = document(
        r#"<t:label-ranges/><t:content-validations><t:content-validation t:name="v"/></t:content-validations>"#,
    );
    assert!(content_validation::Snapshot::parse(&late_after_labels).is_err());
    let valid_prelude = document(
        r#"<t:tracked-changes/><t:calculation-settings/><t:content-validations><t:content-validation t:name="v"/></t:content-validations>"#,
    );
    assert!(content_validation::Snapshot::parse(&valid_prelude).is_ok());
    let dde = document(
        r#"<t:content-validations><t:content-validation t:name="v"/></t:content-validations><t:table t:name="Sheet"><t:table-row><t:covered-table-cell t:number-columns-repeated="3" t:content-validation-name="v"/></t:table-row></t:table><t:dde-links><t:dde-link><t:table><t:table-row><t:table-cell/></t:table-row></t:table></t:dde-link></t:dde-links>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&dde).unwrap();
    assert_eq!(snapshot.sheets().len(), 1);
    assert!(snapshot.sheets()[0].bindings()[0].is_covered_cell());
    assert_eq!(snapshot.sheets()[0].bindings()[0].column_count(), 3);
}

#[test]
fn offsets_duplicates_and_repeat_overflow_fail_closed() {
    let xml = document(
        r#"<t:content-validations><t:content-validation t:name="a"/><t:content-validation t:name="b"/></t:content-validations><t:table t:name="Sheet"><t:table-row><t:table-cell/><t:table-cell t:number-columns-repeated="2" t:content-validation-name="a"/></t:table-row><t:table-row t:number-rows-repeated="3"><t:table-cell t:content-validation-name="b"/></t:table-row></t:table>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&xml).unwrap();
    assert_eq!(snapshot.sheets()[0].logical_row_count(), 4);
    assert_eq!(
        (
            snapshot.sheets()[0].bindings()[0].row(),
            snapshot.sheets()[0].bindings()[0].column()
        ),
        (0, 1)
    );
    assert_eq!(
        (
            snapshot.sheets()[0].bindings()[1].row(),
            snapshot.sheets()[0].bindings()[1].row_count()
        ),
        (1, 3)
    );
    let duplicate_definition = document(
        r#"<t:content-validations><t:content-validation t:name="v"/><t:content-validation t:name="v"/></t:content-validations>"#,
    );
    assert!(content_validation::Snapshot::parse(&duplicate_definition).is_err());
    let duplicate_sheet = document(
        r#"<t:content-validations><t:content-validation t:name="v"/></t:content-validations><t:table t:name="S"/><t:table t:name="S"/>"#,
    );
    assert!(content_validation::Snapshot::parse(&duplicate_sheet).is_err());
    let overflow = document(
        r#"<t:content-validations><t:content-validation t:name="v"/></t:content-validations><t:table t:name="S"><t:table-row t:number-rows-repeated="18446744073709551615"><t:table-cell/></t:table-row><t:table-row/></t:table>"#,
    );
    assert!(content_validation::Snapshot::parse(&overflow).is_err());
}

#[test]
fn long_namespace_uri_is_interned_across_many_attributes() {
    let namespace = format!("urn:vendor:{}", "n".repeat(128 * 1024));
    let mut attributes = String::new();
    for index in 0..2_000 {
        use std::fmt::Write as _;
        write!(attributes, " v:a{index}=\"x\"").unwrap();
    }
    let xml = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:v="{namespace}"><o:body><o:spreadsheet><t:content-validations><t:content-validation t:name="v"{attributes}/></t:content-validations></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = content_validation::Snapshot::parse_with_limits(
        &xml,
        content_validation::Limits::default().with_attributes(2_010),
    )
    .unwrap();
    assert!(snapshot.definitions()[0].has_opaque_content());
}
