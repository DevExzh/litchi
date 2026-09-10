//! Focused regression coverage for the database-range semantic and XML layers.

use super::*;

#[test]
fn new_range_has_an_ergonomic_validated_baseline() {
    let range = Range::new("Sheet1.A1:Sheet1.B2");
    assert!(range.validate().is_ok());
    assert_eq!(range.target_range_address, "Sheet1.A1:Sheet1.B2");
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

#[test]
fn validation_rejects_multiple_unnamed_ranges() {
    let first = Range::new("Sheet1.A1");
    let second = Range::new("Sheet1.B1");
    assert!(validate_database_range_collection(&[first.clone(), second]).is_err());

    let mut output = "prefix".to_string();
    assert!(write_database_ranges(&mut output, &[first, Range::new("Sheet1.B1")]).is_err());
    assert_eq!(output, "prefix");

    let mut named = Range::new("Sheet1.A1");
    named.name = Some("Named".to_string());
    assert!(validate_database_range_collection(&[named, Range::new("Sheet1.B1")]).is_ok());
}

#[test]
fn standalone_parser_rejects_multiple_unnamed_ranges() {
    let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Sheet1.A1"/><t:database-range t:target-range-address="Sheet1.B1"/></t:database-ranges></s>"#;
    assert!(parse_database_ranges(xml).is_err());
}

#[test]
fn filter_color_data_types_follow_odf_vocabulary() {
    let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Sheet1.A1"><t:filter><t:filter-condition t:field-number="0" t:data-type="text-color" t:value="window-font-color" t:operator="="/></t:filter></t:database-range></t:database-ranges></s>"#;
    let ranges = parse_database_ranges(xml).expect("ODF color condition should parse");
    assert_eq!(
        ranges[0].filter.as_ref().unwrap().expression,
        Expression::Condition(Condition {
            field_number: 0,
            value: "window-font-color".to_string(),
            operator: "=".to_string(),
            case_sensitive: None,
            data_type: Some(DataType::TextColor),
            set_items: Vec::new(),
        })
    );
    let invalid = xml.replace("window-font-color", "not-a-color");
    assert!(parse_database_ranges(&invalid).is_err());
}

#[test]
fn database_range_children_follow_schema_order() {
    let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Sheet1.A1"><t:sort><t:sort-by t:field-number="0"/></t:sort><t:filter><t:filter-condition t:field-number="0" t:value="x" t:operator="="/></t:filter></t:database-range></t:database-ranges></s>"#;
    assert!(parse_database_ranges(xml).is_err());
}

#[test]
fn standalone_parser_rejects_nested_leaf_source_content() {
    let xml = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Sheet1.A1"><t:database-source-query t:database-name="db" t:query-name="query"><t:database-source-table t:database-name="db" t:database-table-name="nested"/></t:database-source-query></t:database-range></t:database-ranges></s>"#;
    assert!(parse_database_ranges(xml).is_err());
}

#[test]
fn standalone_parser_rejects_known_children_under_all_leaf_metadata() {
    for leaf in [
        r#"<t:database-source-query t:database-name="db" t:query-name="query"><t:database-source-table t:database-name="db" t:database-table-name="nested"/></t:database-source-query>"#,
        r#"<t:sort><t:sort-by t:field-number="0"><t:sort-by t:field-number="1"/></t:sort-by></t:sort>"#,
        r#"<t:filter><t:filter-condition t:field-number="0" t:value="x" t:operator="="><t:filter-set-item t:value="nested"><t:filter-set-item t:value="deeper"/></t:filter-set-item></t:filter-condition></t:filter>"#,
        r#"<t:subtotal-rules><t:subtotal-rule t:group-by-field-number="0"><t:subtotal-field t:field-number="1" t:function="sum"><t:subtotal-field t:field-number="2" t:function="sum"/></t:subtotal-field></t:subtotal-rule></t:subtotal-rules>"#,
    ] {
        let xml = format!(
            r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Sheet1.A1">{leaf}</t:database-range></t:database-ranges></s>"#
        );
        assert!(
            parse_database_ranges(&xml).is_err(),
            "nested leaf should fail: {leaf}"
        );
    }
}

#[test]
fn whole_column_and_row_target_addresses_are_valid_database_ranges() {
    assert!(Range::new("Input.A:Input.C").validate().is_ok());
    assert!(Range::new("Input.1:Input.3").validate().is_ok());
    assert!(
        Range::new("'Input Data'.A:'Input Data'.C")
            .validate()
            .is_ok()
    );
}

#[test]
fn sort_language_fields_follow_schema_lexical_types() {
    let mut range = Range::new("Input.A1");
    range.sort = Some(Sort {
        language: Some("en".to_string()),
        country: Some("US".to_string()),
        script: Some("Latn".to_string()),
        rfc_language_tag: Some("en-US".to_string()),
        keys: vec![Key::new(0)],
        ..Sort::default()
    });
    assert!(range.validate().is_ok());

    for (field, value) in [
        ("language", "en_US"),
        ("country", "U-"),
        ("script", ""),
        ("rfc_language_tag", "en--US"),
    ] {
        let mut invalid = range.clone();
        let sort = invalid.sort.as_mut().unwrap();
        match field {
            "language" => sort.language = Some(value.to_string()),
            "country" => sort.country = Some(value.to_string()),
            "script" => sort.script = Some(value.to_string()),
            "rfc_language_tag" => sort.rfc_language_tag = Some(value.to_string()),
            _ => unreachable!(),
        }
        assert!(invalid.validate().is_err(), "{field}={value} should fail");
    }

    let mut invalid = range;
    invalid.sort.as_mut().unwrap().algorithm = Some("bad\u{1}algorithm".to_string());
    assert!(invalid.validate().is_err());
}

#[test]
fn public_writer_rejects_sort_algorithm_controls_atomically() {
    let mut range = Range::new("Input.A1");
    range.sort = Some(Sort {
        algorithm: Some("bad\u{1}algorithm".to_string()),
        keys: vec![Key::new(0)],
        ..Sort::default()
    });
    let mut output = "prefix".to_string();
    assert!(write_database_ranges(&mut output, &[range]).is_err());
    assert_eq!(output, "prefix");
}

#[test]
fn standalone_parser_checks_matching_end_names_and_escaped_output_size() {
    let mismatched = r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="Input.A1"></t:database-ranges></s>"#;
    assert!(parse_database_ranges(mismatched).is_err());

    let mut range = Range::new("Input.A1");
    range.source = Some(Source::Table {
        database_name: "db & \"quoted\"".to_string(),
        table_name: "table <one>".to_string(),
    });
    let mut output = String::new();
    write_database_ranges(&mut output, &[range.clone()]).expect("bounded writer should succeed");
    assert!(output.contains("&amp;") && output.contains("&quot;"));
    let reparsed = parse_database_ranges(&format!(
        r#"<s xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{output}</s>"#
    ))
    .expect("escaped source should parse");
    assert_eq!(reparsed, vec![range]);
}

#[test]
fn standalone_parser_bounds_raw_attribute_before_decoding() {
    let oversized = "x".repeat(1_048_577);
    let xml = format!(
        r#"<s xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:database-ranges><t:database-range t:target-range-address="{oversized}"/></t:database-ranges></s>"#
    );
    assert!(parse_database_ranges(&xml).is_err());
}

#[test]
fn standalone_source_and_filter_writers_apply_validation_limits() {
    let oversized = "x".repeat(1_048_577);
    let source = Source::Query {
        database_name: oversized,
        query_name: "query".to_string(),
    };
    let mut output = String::new();
    assert!(write_database_source(&mut output, &source).is_err());
    assert!(output.is_empty(), "rejected source must not mutate output");

    let filter = Filter {
        target_range_address: Some("not-a-range".to_string()),
        condition_source: Some(ConditionSource::SelfContained),
        condition_source_range_address: None,
        display_duplicates: None,
        expression: Expression::Condition(Condition::new(0, "=", "value")),
    };
    assert!(write_filter(&mut output, &filter).is_err());
    assert!(output.is_empty(), "rejected filter must not mutate output");
}

#[test]
fn escaped_output_cap_is_checked_before_emission() {
    let item = "'".repeat(900_000);
    let mut condition = Condition::new(0, "=", "value");
    condition.set_items = vec![item; 13];
    let mut range = Range::new("Input.A1");
    range.filter = Some(Filter {
        target_range_address: None,
        condition_source: Some(ConditionSource::SelfContained),
        condition_source_range_address: None,
        display_duplicates: None,
        expression: Expression::Condition(condition),
    });
    let mut output = "prefix".to_string();
    assert!(write_database_ranges(&mut output, &[range]).is_err());
    assert_eq!(output, "prefix");
}
