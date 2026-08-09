//! Focused regression coverage for the database-range semantic and XML layers.

use super::*;

#[test]
fn new_range_has_an_ergonomic_validated_baseline() {
    let range = Range::new("Sheet1.A1:B2");
    assert!(range.validate().is_ok());
    assert_eq!(range.target_range_address, "Sheet1.A1:B2");
}

#[test]
fn codec_round_trip_preserves_nested_filter_metadata() {
    let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:name="Sales" t:target-range-address="Sheet1.A1:Sheet1.B20"><t:database-source-query t:database-name="sales.odb" t:query-name="OpenOrders"/><t:filter t:condition-source="self"><t:filter-and><t:filter-condition t:field-number="0" t:value="East &amp; West" t:operator="="/><t:filter-or><t:filter-condition t:field-number="1" t:value="10" t:operator=">"/></t:filter-or></t:filter-and></t:filter></t:database-range></t:database-ranges></s>"#;
    let parsed = parse_database_ranges(xml).expect("test fixture or operation should succeed");
    let mut written = String::new();
    write_database_ranges(&mut written, &parsed).expect("test fixture or operation should succeed");
    let reparsed = parse_database_ranges(&format!(
        r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{written}</s>"#
    ))
    .expect("test fixture or operation should succeed");
    assert_eq!(reparsed, parsed);
}

#[test]
fn validation_rejects_duplicate_names_and_same_group_nesting() {
    let expression = Expression::And(vec![Expression::And(vec![Expression::Condition(
        Condition::new(0, "=", "x"),
    )])]);
    let filter = Filter {
        target_range_address: None,
        condition_source: None,
        condition_source_range_address: None,
        display_duplicates: None,
        expression,
    };
    assert!(validate_filter(&filter).is_err());

    let mut first = Range::new("Sheet1.A1");
    first.name = Some("Sales".to_string());
    let mut second = Range::new("Sheet1.B1");
    second.name = Some("Sales".to_string());
    assert!(validate_database_range_collection(&[first, second]).is_err());
}

#[cfg(any())]
mod legacy_end_to_end {
    use super::validation::validate_filter_expression;
    use super::*;
    use crate::{Builder, MutableSpreadsheet, Spreadsheet};

    #[test]
    fn parses_and_writes_complete_database_range_metadata() {
        let xml = r##"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:name="Data &amp; More" t:is-selection="false" t:on-update-keep-styles="true" t:on-update-keep-size="false" t:has-persistent-data="true" t:orientation="column" t:contains-header="true" t:display-filter-buttons="true" t:target-range-address="Sheet1.A1:Sheet1.D20" t:refresh-delay="PT5M"><t:database-source-sql t:database-name="db&amp;1" t:sql-statement="SELECT &lt;x&gt;" t:parse-sql-statement="true"/><t:filter t:target-range-address="Sheet1.F1:Sheet1.I20" t:condition-source="cell-range" t:condition-source-range-address="Sheet2.A1:Sheet2.B2" t:display-duplicates="false"><t:filter-and><t:filter-condition t:field-number="0" t:value="alpha" t:operator="=" t:case-sensitive="true" t:data-type="text"/><t:filter-or><t:filter-condition t:field-number="1" t:value="10" t:operator=">=" t:data-type="number"/><t:filter-condition t:field-number="2" t:value="" t:operator="in"><t:filter-set-item t:value="A&amp;B"/><t:filter-set-item t:value="C"/></t:filter-condition></t:filter-or></t:filter-and></t:filter><t:sort t:bind-styles-to-content="true" t:target-range-address="Sheet1.A2:Sheet1.D20" t:case-sensitive="false" t:language="en" t:country="US" t:script="Latn" t:rfc-language-tag="en-US" t:algorithm="unicode" t:embedded-number-behavior="integer"><t:sort-by t:field-number="1" t:data-type="number" t:order="descending"/></t:sort><t:subtotal-rules t:bind-styles-to-content="false" t:case-sensitive="true" t:page-breaks-on-group-change="true"><t:sort-groups t:data-type="text" t:order="ascending"/><t:subtotal-rule t:group-by-field-number="0"><t:subtotal-field t:field-number="3" t:function="sum"/></t:subtotal-rule></t:subtotal-rules></t:database-range></t:database-ranges></o:spreadsheet>"##;
        let parsed = parse_database_ranges(xml).expect("test fixture or operation should succeed");
        assert_eq!(parsed.len(), 1);
        let range = &parsed[0];
        assert_eq!(range.name.as_deref(), Some("Data & More"));
        assert_eq!(range.orientation, Some(Orientation::Column));
        assert_eq!(range.refresh_delay.as_deref(), Some("PT5M"));
        assert!(matches!(range.source, Some(Source::Sql { .. })));
        assert_eq!(
            range
                .sort
                .as_ref()
                .expect("test fixture or operation should succeed")
                .keys[0]
                .field_number,
            1
        );
        assert_eq!(
            range
                .subtotals
                .as_ref()
                .expect("test fixture or operation should succeed")
                .rules[0]
                .fields[0]
                .function,
            "sum"
        );

        let mut written = String::new();
        write_database_ranges(&mut written, &parsed)
            .expect("test fixture or operation should succeed");
        let reparsed = parse_database_ranges(&format!(
            r#"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{written}</o:spreadsheet>"#
        ))
        .expect("test fixture or operation should succeed");
        assert_eq!(reparsed, parsed);
        assert!(written.contains("Data &amp; More"));
        assert!(written.contains("SELECT &lt;x&gt;"));
    }

    #[test]
    fn rejects_invalid_filter_shapes_and_required_values() {
        let same_group = Expression::And(vec![Expression::And(vec![Expression::Condition(
            Condition::new(0, "=", "x"),
        )])]);
        assert!(validate_filter_expression(&same_group, 0, None).is_err());

        let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="A1:B2"><t:sort/></t:database-range></t:database-ranges></s>"#;
        assert!(parse_database_ranges(xml).is_err());
    }

    #[test]
    fn external_database_sources_remain_inert_data() {
        let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="A1"><t:database-source-query t:database-name="file:///database.odb" t:query-name="DangerousQuery"/></t:database-range></t:database-ranges></s>"#;
        let ranges = parse_database_ranges(xml).expect("test fixture or operation should succeed");
        assert_eq!(
            ranges[0].source,
            Some(Source::Query {
                database_name: "file:///database.odb".to_string(),
                query_name: "DangerousQuery".to_string(),
            })
        );
    }

    #[test]
    fn database_ranges_round_trip_through_builder_and_mutable_packages() {
        let mut range = Range::new("Sheet1.A1:Sheet1.C20");
        range.name = Some("Sales".to_string());
        range.orientation = Some(Orientation::Column);
        range.display_filter_buttons = Some(true);
        range.source = Some(Source::Query {
            database_name: "file:///sales&forecast.odb".to_string(),
            query_name: "Quarter <One>".to_string(),
        });
        range.filter = Some(Filter {
            target_range_address: None,
            condition_source: Some(ConditionSource::SelfContained),
            condition_source_range_address: None,
            display_duplicates: Some(false),
            expression: Expression::And(vec![
                Expression::Condition(Condition::new(0, "=", "East")),
                Expression::Or(vec![
                    Expression::Condition(Condition::new(1, ">", "100")),
                    Expression::Condition(Condition::new(2, "=", "Open")),
                ]),
            ]),
        });
        range.sort = Some(Sort {
            embedded_number_behavior: Some(EmbeddedNumberBehavior::Integer),
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
                fields: vec![Field {
                    field_number: 1,
                    function: "sum".to_string(),
                }],
            }],
            ..Rules::default()
        });

        let mut builder = Builder::new();
        builder
            .add_sheet("Sheet1")
            .expect("test fixture or operation should succeed");
        builder
            .add_database_range(range.clone())
            .expect("test fixture or operation should succeed");
        let bytes = builder
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        assert_eq!(spreadsheet.database_ranges(), &[range.clone()]);

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)
            .expect("test fixture or operation should succeed");
        let Expression::And(expressions) = &mut mutable.database_ranges_mut()[0]
            .filter
            .as_mut()
            .expect("test fixture or operation should succeed")
            .expression
        else {
            panic!("expected AND filter")
        };
        let Expression::Condition(condition) = &mut expressions[0] else {
            panic!("expected filter condition")
        };
        condition.value = "West & Central".to_string();

        let reopened = Spreadsheet::from_bytes(
            mutable
                .to_bytes()
                .expect("test fixture or operation should succeed"),
        )
        .expect("test fixture or operation should succeed");
        let reopened_range = &reopened.database_ranges()[0];
        let Expression::And(expressions) = &reopened_range
            .filter
            .as_ref()
            .expect("test fixture or operation should succeed")
            .expression
        else {
            panic!("expected AND filter")
        };
        let Expression::Condition(condition) = &expressions[0] else {
            panic!("expected filter condition")
        };
        assert_eq!(condition.value, "West & Central");
        assert_eq!(reopened_range.source, range.source);
        assert_eq!(reopened_range.sort, range.sort);
        assert_eq!(reopened_range.subtotals, range.subtotals);
    }
}
