//! Regression coverage for the layered ODF field owner.

use super::*;
use litchi_core::Result;

mod database_field_tests {
    use super::*;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
        xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:text><t:p>"#;
    const SUFFIX: &str = "</t:p></o:text></o:body></o:document-content>";

    #[test]
    fn parses_all_database_field_kinds_without_resolving_resources() {
        let xml = format!(
            r#"{PREFIX}<t:database-display t:database-name="Contacts" t:table-name="People"
                t:table-type="query" t:column-name="FullName" s:data-style-name="N1">A&amp;B</t:database-display>
            <t:database-next t:database-name="Contacts" t:table-name="People" t:condition="of:=TRUE()"/>
            <t:database-row-select t:table-name="People" t:row-number="42"><f:connection-resource x:href="sdbc:embedded:firebird"/></t:database-row-select>
            <t:database-row-number t:database-name="Contacts" t:table-name="People"
                t:value="42" s:num-format="a" s:num-letter-sync="false">42</t:database-row-number>
            <t:database-name t:database-name="Contacts" t:table-name="People">Contacts</t:database-name>{SUFFIX}"#
        );
        let fields = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(fields.len(), 5);
        assert_eq!(fields[0].kind, DatabaseFieldKind::Display);
        assert_eq!(fields[0].display_text, "A&B");
        assert_eq!(
            fields[0].source.effective_table_type(),
            DatabaseTableType::Query
        );
        assert_eq!(
            fields[2]
                .row_number
                .as_ref()
                .map(NonNegativeInteger::as_str),
            Some("42")
        );
        assert_eq!(
            fields[2]
                .source
                .connection_resource
                .as_ref()
                .map(|resource| resource.href.as_str()),
            Some("sdbc:embedded:firebird")
        );
        assert_eq!(fields[3].number_letter_sync, Some(false));
    }

    #[test]
    fn rejects_missing_invalid_nested_and_active_database_fields() {
        let bodies = [
            r#"<t:database-display t:database-name="db" t:table-name="t"/>"#,
            r#"<t:database-next t:database-name="db"/>"#,
            r#"<t:database-row-select t:database-name="db" t:table-name="t" t:row-number="-1"/>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t" t:table-type="view"/>"#,
            r#"<t:database-next t:database-name="db" t:table-name="t">text</t:database-next>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t"><t:span>x</t:span></t:database-name>"#,
            r#"<t:database-name t:table-name="t"><f:connection-resource x:href="https://example.invalid/db"/><f:connection-resource x:href="other"/></t:database-name>"#,
            r#"<t:database-name t:table-name="t"><f:connection-resource x:href="db" x:type="simple"/></t:database-name>"#,
            r#"<t:database-row-number t:table-name="t" s:num-format="1" s:num-letter-sync="true">1</t:database-row-number>"#,
            r#"<t:database-name t:table-name="t">text<f:connection-resource x:href="db"/></t:database-name>"#,
            r#"<t:database-name t:table-name="t" xmlns:z="urn:foreign" z:extra="x"/>"#,
        ];
        for body in bodies {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                FieldParser::parse_database_fields(&xml).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn database_fields_roundtrip_all_kinds_and_schema_optional_values() {
        let source = || DatabaseSource {
            database_name: None,
            table_name: String::new(),
            table_type: Some(DatabaseTableType::Command),
            connection_resource: None,
        };
        let fields = vec![
            DatabaseField {
                kind: DatabaseFieldKind::Display,
                source: source(),
                column_name: Some(String::new()),
                condition: None,
                row_number: None,
                value: None,
                data_style_name: Some("N1".into()),
                number_format: None,
                number_letter_sync: None,
                display_text: "A&B".into(),
            },
            DatabaseField {
                kind: DatabaseFieldKind::Next,
                source: source(),
                column_name: None,
                condition: Some("of:=TRUE()".into()),
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: String::new(),
            },
            DatabaseField {
                kind: DatabaseFieldKind::RowSelect,
                source: source(),
                column_name: None,
                condition: None,
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: String::new(),
            },
            DatabaseField {
                kind: DatabaseFieldKind::RowNumber,
                source: source(),
                column_name: None,
                condition: None,
                row_number: None,
                value: Some(NonNegativeInteger::new("7").unwrap()),
                data_style_name: None,
                number_format: Some("A".into()),
                number_letter_sync: Some(true),
                display_text: "VII".into(),
            },
            DatabaseField {
                kind: DatabaseFieldKind::Name,
                source: DatabaseSource {
                    connection_resource: Some(DatabaseConnectionResource {
                        href: "sdbc:embedded:firebird".into(),
                        simple_link: true,
                    }),
                    ..source()
                },
                column_name: None,
                condition: None,
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: "db".into(),
            },
        ];
        for field in fields {
            let fragment = field.to_xml_fragment().unwrap();
            let parsed =
                FieldParser::parse_database_fields(&format!("{PREFIX}{fragment}{SUFFIX}")).unwrap();
            assert_eq!(parsed, vec![field]);
        }
        let optional = format!(
            "{PREFIX}<t:database-next t:table-name=\"\"/><t:database-row-select t:table-name=\"\"/>{SUFFIX}"
        );
        assert_eq!(
            FieldParser::parse_database_fields(&optional).unwrap().len(),
            2
        );
    }

    #[test]
    fn database_parser_ignores_spoofed_names_and_rejects_bad_placement() {
        let spoof =
            format!("{PREFIX}<z:database-display xmlns:z=\"urn:not-text\" z:any=\"x\"/>{SUFFIX}");
        assert!(
            FieldParser::parse_database_fields(&spoof)
                .unwrap()
                .is_empty()
        );
        let misplaced = format!("{PREFIX}{SUFFIX}")
            .replace("<t:p>", "<t:database-name t:table-name=\"t\"/><t:p>");
        assert!(FieldParser::parse_database_fields(&misplaced).is_err());
    }

    #[test]
    fn database_non_negative_integer_is_arbitrary_width_and_canonical() {
        let beyond_u64 = "18446744073709551616000000000000000000";
        for (lexical, canonical) in [
            (beyond_u64, beyond_u64),
            ("+00042", "42"),
            ("-000", "0"),
            ("  0000\t", "0"),
        ] {
            assert_eq!(
                NonNegativeInteger::new(lexical).unwrap().as_str(),
                canonical
            );
        }
        for invalid in ["", "+", "-", "-1", "1.0", "1 2", "１２", "++1"] {
            assert!(
                NonNegativeInteger::new(invalid).is_err(),
                "accepted {invalid}"
            );
        }
        let boundary = "9".repeat(MAX_DATABASE_INTEGER_DIGITS);
        assert_eq!(
            NonNegativeInteger::new(&boundary).unwrap().as_str(),
            boundary
        );
        assert!(NonNegativeInteger::new(&"9".repeat(MAX_DATABASE_INTEGER_DIGITS + 1)).is_err());

        let xml = format!(
            "{PREFIX}<t:database-row-select t:table-name=\"t\" t:row-number=\"+000{beyond_u64}\"/><t:database-row-number t:table-name=\"t\" t:value=\"-000\">0</t:database-row-number>{SUFFIX}"
        );
        let fields = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(fields[0].row_number.as_ref().unwrap().as_str(), beyond_u64);
        assert_eq!(fields[1].value.as_ref().unwrap().as_str(), "0");
        let canonical = fields
            .iter()
            .map(DatabaseField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        assert!(canonical.contains(&format!("text:row-number=\"{beyond_u64}\"")));
        assert!(canonical.contains("text:value=\"0\""));
        assert!(!canonical.contains("+000"));
    }
}

#[cfg(test)]
mod script_field_tests {
    use super::*;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:l="http://www.w3.org/1999/xlink"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0">
        <o:body><o:text><t:p>"#;
    const SUFFIX: &str = "</t:p></o:text></o:body></o:document-content>";

    fn document(body: &str) -> String {
        format!("{PREFIX}{body}{SUFFIX}")
    }

    fn embedded_field() -> DynamicTextField {
        DynamicTextField::Script {
            href: None,
            language: Some("application/javascript".to_string()),
            content: "alert('stored & inert');".to_string(),
        }
    }

    fn linked_field() -> DynamicTextField {
        DynamicTextField::Script {
            href: Some("https://example.invalid/scripts/main.js?one=1&two=2".to_string()),
            language: Some("application/javascript".to_string()),
            content: String::new(),
        }
    }

    #[test]
    fn script_fields_preserve_inert_links_languages_and_payloads() {
        let embedded = document(
            r#"<t:script s:language="application/javascript">alert('stored &amp; inert');</t:script>"#,
        );
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&embedded).unwrap(),
            vec![embedded_field()]
        );

        let fragment = linked_field().to_xml_fragment().unwrap();
        assert!(fragment.contains(r#"xlink:type="simple""#));
        assert!(fragment.contains(r#"script:language="application/javascript""#));
        assert!(fragment.contains("one=1&amp;two=2"));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&document(&fragment)).unwrap(),
            vec![linked_field()]
        );

        let embedded_fragment = embedded_field().to_xml_fragment().unwrap();
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&document(&embedded_fragment)).unwrap(),
            vec![embedded_field()]
        );

        let empty = DynamicTextField::Script {
            href: None,
            language: None,
            content: String::new(),
        };
        let empty_fragment = empty.to_xml_fragment().unwrap();
        assert!(empty_fragment.ends_with("/>"));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&document(&empty_fragment)).unwrap(),
            vec![empty]
        );
    }

    #[test]
    fn script_fields_reject_invalid_link_metadata_and_nested_content() {
        for invalid in [
            r#"<t:script l:href="https://example.invalid">payload</t:script>"#,
            r#"<t:script l:type="simple">payload</t:script>"#,
            r#"<t:script l:type="extended" l:href="https://example.invalid">payload</t:script>"#,
            r#"<t:script l:type="simple" l:href="https://example.invalid" l:actuate="onLoad">payload</t:script>"#,
            r#"<t:script l:type="simple" l:href="https://example.invalid">payload</t:script>"#,
            r#"<t:script xmlns:fake="urn:not-xlink" fake:type="simple" fake:href="https://example.invalid">payload</t:script>"#,
            r"<t:script><t:span>nested</t:span></t:script>",
            r#"<t:script><foreign:node xmlns:foreign="urn:foreign"/></t:script>"#,
            r"<t:script>before<?unsafe data?>after</t:script>",
        ] {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(invalid)).is_err(),
                "accepted {invalid}"
            );
        }

        let oversized = DynamicTextField::Script {
            href: None,
            language: None,
            content: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod drop_down_field_tests {
    use super::*;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><t:p>"#;
    const SUFFIX: &str = "</t:p></o:text></o:body></o:document-content>";

    fn document(body: &str) -> String {
        format!("{PREFIX}{body}{SUFFIX}")
    }

    fn field() -> DynamicTextField {
        DynamicTextField::DropDown {
            name: "Priority & state".to_string(),
            labels: vec![
                DropDownLabel {
                    value: Some("Low".to_string()),
                    current_selected: Some(false),
                },
                DropDownLabel {
                    value: Some("High & urgent".to_string()),
                    current_selected: Some(true),
                },
                DropDownLabel::default(),
            ],
            display_text: "High & urgent".to_string(),
        }
    }

    #[test]
    fn drop_down_fields_preserve_labels_selected_state_and_cached_text() {
        let xml = document(
            r#"<t:drop-down t:name="Priority &amp; state"><t:label t:value="Low" t:current-selected="false"/><t:label t:value="High &amp; urgent" t:current-selected="1"></t:label><t:label/>High &amp; urgent</t:drop-down>"#,
        );
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&xml).unwrap(),
            vec![field()]
        );

        let fragment = field().to_xml_fragment().unwrap();
        assert!(fragment.contains(r#"text:name="Priority &amp; state""#));
        assert!(fragment.contains(r#"text:current-selected="true""#));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&document(&fragment)).unwrap(),
            vec![field()]
        );

        let empty = DynamicTextField::DropDown {
            name: String::new(),
            labels: Vec::new(),
            display_text: String::new(),
        };
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&document(&empty.to_xml_fragment().unwrap()))
                .unwrap(),
            vec![empty]
        );

        let with_outer_instruction = format!("<?outside test?>{}", document(&fragment));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&with_outer_instruction).unwrap(),
            vec![field()]
        );
    }

    #[test]
    fn drop_down_fields_follow_dynamic_field_order_and_work_inside_meta_fields() {
        let xml = document(
            r#"<t:date>2026-07-22</t:date><t:meta-field xml:id="meta"><t:drop-down t:name="choice"><t:label t:value="one" t:current-selected="true"/>one</t:drop-down></t:meta-field><t:time>10:00</t:time>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert!(matches!(fields[0], DynamicTextField::Date { .. }));
        assert!(matches!(fields[1], DynamicTextField::MetaField { .. }));
        assert!(matches!(fields[2], DynamicTextField::DropDown { .. }));
        assert!(matches!(fields[3], DynamicTextField::Time { .. }));
    }

    #[test]
    fn drop_down_fields_reject_schema_and_resource_limit_violations() {
        for invalid in [
            r"<t:drop-down>missing name</t:drop-down>",
            r#"<t:drop-down xmlns:x="urn:not-text" x:name="spoof">value</t:drop-down>"#,
            r#"<t:drop-down t:name="choice" t:extra="no">value</t:drop-down>"#,
            r#"<t:drop-down t:name="choice"><t:label>text is forbidden</t:label></t:drop-down>"#,
            r#"<t:drop-down t:name="choice">selected<t:label t:value="late"/></t:drop-down>"#,
            r#"<t:drop-down t:name="choice"><t:span>not a label</t:span></t:drop-down>"#,
            r#"<t:drop-down t:name="choice"><t:label t:current-selected="maybe"/></t:drop-down>"#,
            r#"<t:drop-down t:name="choice"><t:label t:extra="no"/></t:drop-down>"#,
            r#"<t:drop-down t:name="choice"><?unsafe value?>value</t:drop-down>"#,
        ] {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(invalid)).is_err(),
                "accepted {invalid}"
            );
        }

        let oversized = DynamicTextField::DropDown {
            name: "choice".to_string(),
            labels: vec![DropDownLabel {
                value: Some("x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1)),
                current_selected: None,
            }],
            display_text: String::new(),
        };
        assert!(oversized.validate().is_err());
    }
}

#[cfg(test)]
mod fixed_page_date_time_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn fixed_page_date_time_round_trips_every_standard_field() {
        let fields = vec![
            DynamicTextField::PageNumber {
                number_format: Some(SequenceNumberFormat::new("A", Some(true)).unwrap()),
                fixed: Some(false),
                page_adjust: Some(-2),
                select_page: Some(PageSelection::Previous),
                display_text: "IV & cached".to_string(),
            },
            DynamicTextField::Date {
                value: Some(FieldDateValue::new("2024-02-29Z").unwrap()),
                adjustment: Some(FieldDuration::new("-P1Y2M3DT4H5M6.7S").unwrap()),
                fixed: Some(true),
                data_style_name: Some("Date & Time".to_string()),
                display_text: "29 < February".to_string(),
            },
            DynamicTextField::Time {
                value: Some(FieldTimeValue::new("2024-02-29T24:00:00+14:00").unwrap()),
                adjustment: Some(FieldDuration::new("PT15M").unwrap()),
                fixed: Some(false),
                data_style_name: Some("Clock".to_string()),
                display_text: "midnight".to_string(),
            },
            DynamicTextField::PageContinuation {
                select_page: PageContinuationSelection::Next,
                string_value: Some("Continued on & next".to_string()),
                display_text: "continued <cached>".to_string(),
            },
        ];
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[1].display_text(), "29 < February");
    }

    #[test]
    fn fixed_page_date_time_accepts_aliases_and_exact_temporal_lexicals() {
        let xml = document(
            r#"<t:date t:date-value="-12345-01-01+05:30" t:date-adjust="P999999999999Y" s:data-style-name="D">historic</t:date>
               <t:time t:time-value="23:59:59.123456789Z" t:time-adjust="-PT0.5S">clock</t:time>
               <t:page-number t:select-page="next" t:page-adjust="9223372036854775807" s:num-format="a" s:num-letter-sync="0">a</t:page-number>
               <t:page-continuation t:select-page="previous">back</t:page-continuation>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert!(matches!(
            &fields[0],
            DynamicTextField::Date { value: Some(value), .. }
                if value.kind() == DateValueKind::Date
        ));
        assert!(matches!(
            &fields[1],
            DynamicTextField::Time { value: Some(value), .. }
                if value.kind() == TimeValueKind::Time
        ));
    }

    #[test]
    fn fixed_page_date_time_rejects_hostile_invalid_and_extension_inputs() {
        for value in [
            "2023-02-29",
            "0000-01-01",
            "2024-01-01+14:01",
            "+2024-01-01",
        ] {
            assert!(FieldDateValue::new(value).is_err(), "accepted {value}");
        }
        for value in ["24:00:01", "12:60:00", "12:00:60", "12:00:00+15:00"] {
            assert!(FieldTimeValue::new(value).is_err(), "accepted {value}");
        }
        assert!(FieldDuration::new("P").is_err());
        assert!(FieldDateValue::new("2024-01-01\u{0}").is_err());

        let invalid = [
            r#"<t:page-number t:select-page="later">1</t:page-number>"#,
            r#"<t:page-number t:page-adjust="9223372036854775808">1</t:page-number>"#,
            r#"<t:page-number s:num-letter-sync="true">a</t:page-number>"#,
            r#"<t:date t:date-value="2024-01-01" t:extra="x">date</t:date>"#,
            r#"<t:time t:time-value="12:00:00" t:date-adjust="P1D">time</t:time>"#,
            r#"<t:page-continuation t:select-page="current">continued</t:page-continuation>"#,
            r"<t:page-continuation>continued</t:page-continuation>",
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:date t:date-value="2024-01-01" x:data-style-name="spoof">date</t:date>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let extension = document(
            r#"<t:page-continuation-string t:select-page="next">extension</t:page-continuation-string>"#,
        );
        assert!(
            FieldParser::parse_dynamic_text_fields(&extension)
                .unwrap()
                .is_empty()
        );

        let oversized = DynamicTextField::PageContinuation {
            select_page: PageContinuationSelection::Next,
            string_value: Some("x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1)),
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod page_variable_family_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn page_variable_family_round_trips_both_standard_elements() {
        let fields = vec![
            DynamicTextField::PageVariableSet {
                active: Some(false),
                page_adjust: Some(i64::MIN),
                display_text: "inert setter cache & <safe>".to_string(),
            },
            DynamicTextField::PageVariableGet {
                number_format: Some(SequenceNumberFormat::new("A", Some(true)).unwrap()),
                display_text: "cached A & <not calculated>".to_string(),
            },
        ];
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[1].display_text(), "cached A & <not calculated>");
    }

    #[test]
    fn page_variable_family_preserves_omission_and_exposes_defaults() {
        let xml = document(
            r#"<t:page-variable-set>opaque</t:page-variable-set>
               <t:page-variable-get s:num-format="a" s:num-letter-sync="0">x</t:page-variable-get>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].effective_page_variable_active(), Some(true));
        assert_eq!(fields[0].effective_page_variable_adjustment(), Some(0));
        assert!(matches!(
            &fields[0],
            DynamicTextField::PageVariableSet {
                active: None,
                page_adjust: None,
                display_text,
            } if display_text == "opaque"
        ));
        assert_eq!(
            fields[0].to_xml_fragment().unwrap(),
            concat!(
                "<text:page-variable-set xmlns:text=\"",
                "urn:oasis:names:tc:opendocument:xmlns:text:1.0\">",
                "opaque</text:page-variable-set>"
            )
        );
    }

    #[test]
    fn page_variable_family_rejects_nonstandard_hostile_and_oversized_input() {
        let invalid = [
            r#"<t:page-variable-set t:active="TRUE"/>"#,
            r#"<t:page-variable-set t:page-adjust="9223372036854775808"/>"#,
            r#"<t:page-variable-set t:value="3"/>"#,
            r#"<t:page-variable-set t:select-page="next"/>"#,
            r#"<t:page-variable-get s:num-letter-sync="true">a</t:page-variable-get>"#,
            r#"<t:page-variable-get s:num-format="I" s:num-letter-sync="true">I</t:page-variable-get>"#,
            r#"<t:page-variable-get t:page-adjust="1">1</t:page-variable-get>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:page-variable-get x:num-format="1">1</t:page-variable-get>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::PageVariableSet {
            active: None,
            page_adjust: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::PageVariableGet {
            number_format: None,
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod document_metadata_fixed_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    fn metadata_field(
        kind: MetadataFieldKind,
        value: Option<MetadataFieldValue>,
        display_text: &str,
    ) -> DynamicTextField {
        DynamicTextField::DocumentMetadata {
            kind,
            value,
            fixed: Some(true),
            data_style_name: kind
                .permits_data_style()
                .then(|| format!("style-{display_text}")),
            display_text: display_text.to_string(),
        }
    }

    #[test]
    fn document_metadata_fixed_fields_round_trip_all_eight_standard_elements() {
        let fields = vec![
            metadata_field(
                MetadataFieldKind::CreationDate,
                Some(MetadataFieldValue::Date(
                    FieldDateValue::new("2024-02-29T23:59:59Z").unwrap(),
                )),
                "created date & <cached>",
            ),
            metadata_field(
                MetadataFieldKind::CreationTime,
                Some(MetadataFieldValue::Time(
                    FieldTimeValue::new("2024-02-29T24:00:00+14:00").unwrap(),
                )),
                "created time",
            ),
            metadata_field(
                MetadataFieldKind::PrintDate,
                Some(MetadataFieldValue::Date(
                    FieldDateValue::new("2025-01-31-05:00").unwrap(),
                )),
                "print date",
            ),
            metadata_field(
                MetadataFieldKind::PrintTime,
                Some(MetadataFieldValue::Time(
                    FieldTimeValue::new("12:34:56.789Z").unwrap(),
                )),
                "print time",
            ),
            metadata_field(MetadataFieldKind::EditingCycles, None, "42"),
            metadata_field(
                MetadataFieldKind::EditingDuration,
                Some(MetadataFieldValue::Duration(
                    FieldDuration::new("P999999999999Y11M30DT23H59M59.5S").unwrap(),
                )),
                "edited duration",
            ),
            metadata_field(
                MetadataFieldKind::ModificationDate,
                Some(MetadataFieldValue::Date(
                    FieldDateValue::new("-12345-12-31Z").unwrap(),
                )),
                "modified date",
            ),
            metadata_field(
                MetadataFieldKind::ModificationTime,
                Some(MetadataFieldValue::Time(
                    FieldTimeValue::new("00:00:00+05:30").unwrap(),
                )),
                "modified time",
            ),
        ];
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[0].display_text(), "created date & <cached>");
    }

    #[test]
    fn document_metadata_fixed_fields_preserve_optional_attributes_and_aliases() {
        let xml = document(
            r#"<t:creation-date>created</t:creation-date>
               <t:creation-time t:fixed="0">time</t:creation-time>
               <t:print-date s:data-style-name="D">date</t:print-date>
               <t:print-time>time</t:print-time>
               <t:editing-cycles t:fixed="1">7</t:editing-cycles>
               <t:editing-duration s:data-style-name="Elapsed">duration</t:editing-duration>
               <t:modification-date>date</t:modification-date>
               <t:modification-time>time</t:modification-time>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 8);
        assert!(matches!(
            &fields[0],
            DynamicTextField::DocumentMetadata {
                kind: MetadataFieldKind::CreationDate,
                value: None,
                fixed: None,
                data_style_name: None,
                ..
            }
        ));
        assert!(matches!(
            &fields[4],
            DynamicTextField::DocumentMetadata {
                kind: MetadataFieldKind::EditingCycles,
                fixed: Some(true),
                ..
            }
        ));
    }

    #[test]
    fn document_metadata_fixed_fields_reject_invalid_lexicals_attributes_and_bounds() {
        let invalid = [
            r#"<t:creation-date t:date-value="2023-02-29">bad</t:creation-date>"#,
            r#"<t:creation-time t:time-value="12:60:00">bad</t:creation-time>"#,
            r#"<t:print-date t:date-value="2024-01-01T00:00:00">bad</t:print-date>"#,
            r#"<t:print-time t:time-value="2024-01-01T00:00:00">bad</t:print-time>"#,
            r#"<t:editing-cycles s:data-style-name="N">7</t:editing-cycles>"#,
            r#"<t:editing-cycles t:duration="P1D">7</t:editing-cycles>"#,
            r#"<t:editing-duration t:duration="P">bad</t:editing-duration>"#,
            r#"<t:modification-date t:time-value="12:00:00">bad</t:modification-date>"#,
            r#"<t:modification-time t:fixed="TRUE">bad</t:modification-time>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let mismatched = metadata_field(
            MetadataFieldKind::PrintDate,
            Some(MetadataFieldValue::Date(
                FieldDateValue::new("2024-01-01T00:00:00Z").unwrap(),
            )),
            "bad",
        );
        assert!(mismatched.to_xml_fragment().is_err());

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:print-date x:data-style-name="spoof">date</t:print-date>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::DocumentMetadata {
            kind: MetadataFieldKind::EditingCycles,
            value: None,
            fixed: None,
            data_style_name: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::DocumentMetadata {
            kind: MetadataFieldKind::EditingCycles,
            value: None,
            fixed: None,
            data_style_name: None,
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod document_identity_fixed_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn document_identity_fixed_fields_round_trip_all_nine_standard_elements() {
        let kinds = [
            IdentityFieldKind::InitialCreator,
            IdentityFieldKind::Description,
            IdentityFieldKind::PrintedBy,
            IdentityFieldKind::Title,
            IdentityFieldKind::Subject,
            IdentityFieldKind::Keywords,
            IdentityFieldKind::Creator,
            IdentityFieldKind::AuthorName,
            IdentityFieldKind::AuthorInitials,
        ];
        let fields = kinds
            .into_iter()
            .enumerate()
            .map(|(index, kind)| DynamicTextField::DocumentIdentity {
                kind,
                fixed: Some(index % 2 == 0),
                display_text: format!("cached {index} & <inert>"),
            })
            .collect::<Vec<_>>();
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[3].display_text(), "cached 3 & <inert>");
    }

    #[test]
    fn document_identity_fixed_fields_preserve_omission_and_namespace_aliases() {
        let xml = document(
            r"<t:initial-creator>first</t:initial-creator>
               <t:description>description</t:description>
               <t:printed-by>printer</t:printed-by>
               <t:title>title</t:title>
               <t:subject>subject</t:subject>
               <t:keywords>one, two</t:keywords>
               <t:creator>last</t:creator>
               <t:author-name>author</t:author-name>
               <t:author-initials>au</t:author-initials>",
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 9);
        for field in fields {
            assert!(matches!(
                field,
                DynamicTextField::DocumentIdentity { fixed: None, .. }
            ));
        }
    }

    #[test]
    fn document_identity_fixed_fields_reject_hostile_attributes_and_bounds() {
        let invalid = [
            r#"<t:initial-creator t:fixed="TRUE">first</t:initial-creator>"#,
            r#"<t:description t:name="not-standard">description</t:description>"#,
            r#"<t:printed-by t:fixed="yes">printer</t:printed-by>"#,
            r#"<t:title t:display="value">title</t:title>"#,
            r#"<t:subject t:fixed="2">subject</t:subject>"#,
            r#"<t:keywords t:string-value="one">one</t:keywords>"#,
            r#"<t:creator t:extra="x">creator</t:creator>"#,
            r#"<t:author-name t:display="value">author</t:author-name>"#,
            r#"<t:author-initials t:fixed="yes">au</t:author-initials>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-text"><o:body><o:text><t:p>
            <t:title x:fixed="true">spoof</t:title>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::DocumentIdentity {
            kind: IdentityFieldKind::Description,
            fixed: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::DocumentIdentity {
            kind: IdentityFieldKind::Title,
            fixed: Some(true),
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod document_context_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn document_context_fields_round_trip_every_standard_display() {
        let fields = vec![
            DynamicTextField::FileName {
                display: Some(FileNameDisplay::Full),
                fixed: Some(true),
                display_text: "file:///cached/report.odt".to_string(),
            },
            DynamicTextField::FileName {
                display: Some(FileNameDisplay::Path),
                fixed: Some(false),
                display_text: "file:///cached/".to_string(),
            },
            DynamicTextField::FileName {
                display: Some(FileNameDisplay::Name),
                fixed: None,
                display_text: "report".to_string(),
            },
            DynamicTextField::FileName {
                display: Some(FileNameDisplay::NameAndExtension),
                fixed: None,
                display_text: "report.odt".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::Area),
                display_text: "Business".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::Full),
                display_text: "file:///templates/Letter.ott".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::Name),
                display_text: "Letter".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::NameAndExtension),
                display_text: "Letter.ott".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::Path),
                display_text: "file:///templates/".to_string(),
            },
            DynamicTextField::TemplateName {
                display: Some(TemplateNameDisplay::Title),
                display_text: "Cached template & <title>".to_string(),
            },
            DynamicTextField::SheetName {
                display_text: "Sheet 1".to_string(),
            },
        ];
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[10].display_text(), "Sheet 1");
    }

    #[test]
    fn document_context_fields_preserve_attribute_omission() {
        let xml = document(
            r"<t:file-name>cached.odt</t:file-name>
               <t:template-name>Letter</t:template-name>
               <t:sheet-name>Budget</t:sheet-name>",
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(matches!(
            &fields[0],
            DynamicTextField::FileName {
                display: None,
                fixed: None,
                ..
            }
        ));
        assert!(matches!(
            &fields[1],
            DynamicTextField::TemplateName { display: None, .. }
        ));
        assert!(matches!(&fields[2], DynamicTextField::SheetName { .. }));
    }

    #[test]
    fn document_context_fields_reject_hostile_attributes_and_bounds() {
        let invalid = [
            r#"<t:file-name t:display="directory">report</t:file-name>"#,
            r#"<t:file-name t:fixed="yes">report</t:file-name>"#,
            r#"<t:file-name t:template-name="Letter">report</t:file-name>"#,
            r#"<t:template-name t:display="extension">Letter</t:template-name>"#,
            r#"<t:template-name t:fixed="true">Letter</t:template-name>"#,
            r#"<t:sheet-name t:display="name">Budget</t:sheet-name>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-text"><o:body><o:text><t:p>
            <t:file-name x:display="full">spoof</t:file-name>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::FileName {
            display: None,
            fixed: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::TemplateName {
            display: Some(TemplateNameDisplay::Title),
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod sender_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn sender_fields_round_trip_all_fifteen_standard_elements() {
        let kinds = [
            SenderFieldKind::FirstName,
            SenderFieldKind::LastName,
            SenderFieldKind::Initials,
            SenderFieldKind::Title,
            SenderFieldKind::Position,
            SenderFieldKind::Email,
            SenderFieldKind::PrivatePhone,
            SenderFieldKind::Fax,
            SenderFieldKind::Company,
            SenderFieldKind::WorkPhone,
            SenderFieldKind::Street,
            SenderFieldKind::City,
            SenderFieldKind::PostalCode,
            SenderFieldKind::Country,
            SenderFieldKind::StateOrProvince,
        ];
        let fields = kinds
            .into_iter()
            .enumerate()
            .map(|(index, kind)| DynamicTextField::Sender {
                kind,
                fixed: match index % 3 {
                    0 => None,
                    1 => Some(true),
                    _ => Some(false),
                },
                display_text: format!("cached sender {index} & <inert>"),
            })
            .collect::<Vec<_>>();
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[5].display_text(), "cached sender 5 & <inert>");
    }

    #[test]
    fn sender_fields_preserve_fixed_omission_and_boolean_aliases() {
        let xml = document(
            r#"<t:sender-firstname>first</t:sender-firstname>
               <t:sender-lastname t:fixed="1">last</t:sender-lastname>
               <t:sender-email t:fixed="0">author@example.invalid</t:sender-email>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(matches!(
            &fields[0],
            DynamicTextField::Sender {
                kind: SenderFieldKind::FirstName,
                fixed: None,
                ..
            }
        ));
        assert!(matches!(
            &fields[1],
            DynamicTextField::Sender {
                kind: SenderFieldKind::LastName,
                fixed: Some(true),
                ..
            }
        ));
        assert!(matches!(
            &fields[2],
            DynamicTextField::Sender {
                kind: SenderFieldKind::Email,
                fixed: Some(false),
                ..
            }
        ));
    }

    #[test]
    fn sender_fields_reject_hostile_attributes_and_bounds() {
        let invalid = [
            r#"<t:sender-firstname t:fixed="yes">first</t:sender-firstname>"#,
            r#"<t:sender-email t:display="value">author@example.invalid</t:sender-email>"#,
            r#"<t:sender-country t:fixed="TRUE">Country</t:sender-country>"#,
            r#"<t:sender-state-or-province t:name="region">State</t:sender-state-or-province>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-text"><o:body><o:text><t:p>
            <t:sender-company x:fixed="true">spoof</t:sender-company>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::Sender {
            kind: SenderFieldKind::Company,
            fixed: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::Sender {
            kind: SenderFieldKind::Email,
            fixed: Some(false),
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod chapter_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn chapter_fields_round_trip_every_standard_display() {
        let displays = [
            None,
            Some(ChapterDisplay::Name),
            Some(ChapterDisplay::Number),
            Some(ChapterDisplay::NumberAndName),
            Some(ChapterDisplay::PlainNumber),
            Some(ChapterDisplay::PlainNumberAndName),
        ];
        let fields = displays
            .into_iter()
            .enumerate()
            .map(|(index, display)| DynamicTextField::Chapter {
                display,
                outline_level: (index % 2 == 0)
                    .then(|| NonNegativeInteger::new(&index.to_string()).unwrap()),
                display_text: format!("cached chapter {index} & <inert>"),
            })
            .collect::<Vec<_>>();
        let body = fields
            .iter()
            .map(DynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[5].display_text(), "cached chapter 5 & <inert>");
    }

    #[test]
    fn chapter_fields_preserve_omission_and_canonical_outline_levels() {
        let xml = document(
            r#"<t:chapter>untargeted</t:chapter>
               <t:chapter t:display="number-and-name" t:outline-level="0002">two</t:chapter>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert!(matches!(
            &fields[0],
            DynamicTextField::Chapter {
                display: None,
                outline_level: None,
                ..
            }
        ));
        let DynamicTextField::Chapter {
            display,
            outline_level,
            ..
        } = &fields[1]
        else {
            panic!("expected chapter field");
        };
        assert_eq!(*display, Some(ChapterDisplay::NumberAndName));
        assert_eq!(outline_level.as_ref().unwrap().as_str(), "2");
    }

    #[test]
    fn chapter_fields_reject_hostile_attributes_and_bounds() {
        let invalid = [
            r#"<t:chapter t:display="full">chapter</t:chapter>"#,
            r#"<t:chapter t:outline-level="-1">chapter</t:chapter>"#,
            r#"<t:chapter t:outline-level="1.5">chapter</t:chapter>"#,
            r#"<t:chapter t:fixed="true">chapter</t:chapter>"#,
            r#"<t:chapter s:data-style-name="N">chapter</t:chapter>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-text"><o:body><o:text><t:p>
            <t:chapter x:outline-level="2">spoof</t:chapter>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::Chapter {
            display: None,
            outline_level: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::Chapter {
            display: Some(ChapterDisplay::Name),
            outline_level: Some(NonNegativeInteger::new("1").unwrap()),
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod user_defined_metadata_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn user_defined_metadata_field_round_trips_every_independent_value_attribute() {
        let field = DynamicTextField::UserDefinedMetadata {
            name: "custom & name".to_string(),
            values: UserDefinedMetadataValues {
                number: Some("-INF".to_string()),
                date: Some(FieldDateValue::new("2024-02-29T23:59:59Z").unwrap()),
                time: Some(FieldDuration::new("P999999999999Y1M2DT3H4M5.6S").unwrap()),
                boolean: Some(false),
                string: Some("cached & <string>".to_string()),
            },
            fixed: Some(true),
            data_style_name: Some("Custom & Style".to_string()),
            display_text: "inert & <presentation>".to_string(),
        };
        let fragment = field.to_xml_fragment().unwrap();
        assert!(fragment.contains("office:value=\"-INF\""));
        assert!(fragment.contains("office:boolean-value=\"false\""));
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&fragment)).unwrap();
        assert_eq!(parsed, vec![field]);
        assert_eq!(parsed[0].display_text(), "inert & <presentation>");
    }

    #[test]
    fn user_defined_metadata_field_preserves_empty_name_values_and_omission() {
        let xml = document(
            r#"<t:user-defined t:name="" o:string-value="" t:fixed="0"
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0">cached</t:user-defined>
               <t:user-defined t:name="minimal"/>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert!(matches!(
            &fields[0],
            DynamicTextField::UserDefinedMetadata {
                name,
                values: UserDefinedMetadataValues { string: Some(value), .. },
                fixed: Some(false),
                ..
            } if name.is_empty() && value.is_empty()
        ));
        assert!(matches!(
            &fields[1],
            DynamicTextField::UserDefinedMetadata {
                values: UserDefinedMetadataValues {
                    number: None,
                    date: None,
                    time: None,
                    boolean: None,
                    string: None,
                },
                fixed: None,
                data_style_name: None,
                ..
            }
        ));
    }

    #[test]
    fn user_defined_metadata_field_rejects_nonstandard_invalid_and_hostile_input() {
        let invalid = [
            r"<t:user-defined>missing name</t:user-defined>",
            r#"<t:user-defined t:name="x" o:value-type="float" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:currency="USD" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:value="1e" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:date-value="2023-02-29" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:time-value="P" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:boolean-value="TRUE" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" t:fixed="yes"/>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-office"><o:body><o:text><t:p>
            <t:user-defined t:name="x" x:string-value="spoof">value</t:user-defined>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = DynamicTextField::UserDefinedMetadata {
            name: "x".to_string(),
            values: UserDefinedMetadataValues {
                string: Some("x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1)),
                ..UserDefinedMetadataValues::default()
            },
            fixed: None,
            data_style_name: None,
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = DynamicTextField::UserDefinedMetadata {
            name: "bad\u{0}".to_string(),
            values: UserDefinedMetadataValues::default(),
            fixed: None,
            data_style_name: None,
            display_text: String::new(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod meta_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <office:body><office:text>{body}</office:text></office:body>
</office:document-content>"#
        )
    }

    #[test]
    fn meta_field_preserves_ordered_mixed_content_and_roundtrips() {
        let xml = document(
            r#"<text:p><text:meta-field xml:id="meta1" style:data-style-name="N1">before<text:span text:style-name="Em">middle</text:span>after<text:a xlink:href="https://example.invalid" xlink:type="simple">link</text:a>end</text:meta-field></text:p>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 1);
        let DynamicTextField::MetaField {
            xml_id,
            data_style_name,
            content,
        } = &fields[0]
        else {
            panic!("expected metadata field");
        };
        assert_eq!(xml_id, "meta1");
        assert_eq!(data_style_name.as_deref(), Some("N1"));
        assert_eq!(content.display_text(), "beforemiddleafterlinkend");
        assert!(matches!(
            content.nodes(),
            [
                MetaFieldNode::Text(_),
                MetaFieldNode::Element(_),
                MetaFieldNode::Text(_),
                MetaFieldNode::Element(_),
                MetaFieldNode::Text(_),
            ]
        ));

        let fragment = fields[0].to_xml_fragment().unwrap();
        assert!(fragment.contains("xml:id=\"meta1\""));
        assert!(fragment.contains("style:data-style-name=\"N1\""));
        assert!(fragment.contains("xlink:href=\"https://example.invalid\""));
        let reparsed = FieldParser::parse_dynamic_text_fields(&document(&format!(
            "<text:p>{fragment}</text:p>"
        )))
        .unwrap();
        assert_eq!(reparsed, fields);
    }

    #[test]
    fn meta_field_recursion_is_inert_and_fields_remain_in_document_order() {
        let xml = document(
            r#"<text:p><text:meta-field xml:id="outer">A<text:meta-field xml:id="inner">B</text:meta-field>C</text:meta-field></text:p>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        let DynamicTextField::MetaField {
            xml_id, content, ..
        } = &fields[0]
        else {
            panic!("expected outer metadata field");
        };
        assert_eq!(xml_id, "outer");
        assert_eq!(content.display_text(), "ABC");
        assert!(
            matches!(content.nodes().get(1), Some(MetaFieldNode::Element(element)) if element.local_name == "meta-field")
        );
        let DynamicTextField::MetaField {
            xml_id, content, ..
        } = &fields[1]
        else {
            panic!("expected inner metadata field");
        };
        assert_eq!(xml_id, "inner");
        assert_eq!(content.display_text(), "B");
    }

    #[test]
    fn meta_field_rejects_invalid_identity_placement_and_markup() {
        let invalid = [
            r"<text:p><text:meta-field>missing</text:meta-field></text:p>",
            r#"<text:p><text:meta-field xml:id="">empty</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="1bad">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="bad:id">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m" text:style-name="bad">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="same">a</text:meta-field><text:meta-field xml:id="same">b</text:meta-field></text:p>"#,
            r#"<text:meta-field xml:id="m">not paragraph content</text:meta-field>"#,
            r#"<text:p><text:meta-field xml:id="m"><table:table/></text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m"><evil:x xmlns:evil="urn:evil"/></text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m"><bad:x/></text:meta-field></text:p>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let dtd = document(r#"<text:p><text:meta-field xml:id="m">x</text:meta-field></text:p>"#)
            .replacen(
                "<office:document-content",
                "<!DOCTYPE x><office:document-content",
                1,
            );
        assert!(FieldParser::parse_dynamic_text_fields(&dtd).is_err());
        assert!(FieldParser::parse_dynamic_text_fields(&document(
            r#"<text:p><text:meta-field xml:id="m">a<?unsafe data?>b</text:meta-field></text:p>"#,
        ))
        .is_err());
    }

    #[test]
    fn meta_field_dispatch_ignores_foreign_vocabulary_but_keeps_real_roots_strict() {
        let spoof = document(
            r#"<text:p><fake:meta-field xmlns:fake="urn:not-text" fake:attribute="ignored">spoof</fake:meta-field></text:p>"#,
        );
        assert!(
            FieldParser::parse_dynamic_text_fields(&spoof)
                .unwrap()
                .is_empty()
        );

        let genuine = document(
            r#"<text:p><text:meta-field xmlns:fake="urn:not-text" xml:id="m" fake:attribute="rejected">real</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&genuine).is_err());
    }

    #[test]
    fn meta_field_accepts_rng_inline_child_grammars() {
        let allowed = [
            r#"text<text:span><text:a xlink:type="simple" xlink:href="urn:test"><text:date>2026-07-18</text:date></text:a></text:span>"#,
            r#"<text:meta xml:id="nested-meta"><text:meta-field xml:id="nested-field">nested</text:meta-field></text:meta>"#,
            r"<text:ruby><text:ruby-base>base<text:span>span</text:span></text:ruby-base><text:ruby-text>reading</text:ruby-text></text:ruby>",
            r#"<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>body</text:p></text:note-body></text:note>"#,
            r#"<text:execute-macro xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><office:event-listeners><script:event-listener script:event-name="dom:click" script:language="ooo:script" script:macro-name="M"/></office:event-listeners>cached</text:execute-macro>"#,
            r#"<office:annotation><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">A</dc:creator><text:p>comment</text:p></office:annotation>"#,
            r#"<draw:line xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/>"#,
            r#"<presentation:header xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"/>"#,
        ];
        for (index, content) in allowed.into_iter().enumerate() {
            let xml = document(&format!(
                r#"<text:p><text:meta-field xml:id="allowed{index}">{content}</text:meta-field></text:p>"#,
            ));
            assert!(
                FieldParser::parse_dynamic_text_fields(&xml).is_ok(),
                "rejected allowed RNG content {content}"
            );
        }
    }

    #[test]
    fn meta_field_rejects_wrong_rng_descendant_elements_and_cardinality() {
        let disallowed = [
            r#"<text:a xlink:type="simple" xlink:href="urn:outer"><text:a xlink:type="simple" xlink:href="urn:inner">nested</text:a></text:a>"#,
            r"<text:date><text:span>not cached text</text:span></text:date>",
            r"<text:s>not empty</text:s>",
            r"<text:number>heading-only vocabulary</text:number>",
            r"<text:ruby><text:ruby-text>wrong order</text:ruby-text><text:ruby-base>base</text:ruby-base></text:ruby>",
            r"<text:ruby><text:ruby-base>missing reading</text:ruby-base></text:ruby>",
            r#"<text:note text:note-class="footnote"><text:note-body/><text:note-citation>1</text:note-citation></text:note>"#,
            r#"<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><style:style/></text:note-body></text:note>"#,
            r"<text:execute-macro>text<office:event-listeners/></text:execute-macro>",
            r"<office:event-listeners/>",
            r"<text:span><table:table/></text:span>",
            r"<text:span><style:style/></text:span>",
            r#"<draw:line xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><table:table/></draw:line>"#,
        ];
        for (index, content) in disallowed.into_iter().enumerate() {
            let xml = document(&format!(
                r#"<text:p><text:meta-field xml:id="bad{index}">{content}</text:meta-field></text:p>"#,
            ));
            assert!(
                FieldParser::parse_dynamic_text_fields(&xml).is_err(),
                "accepted disallowed RNG content {content}"
            );
        }
    }

    #[test]
    fn meta_field_scan_enforces_document_wide_xml_id_uniqueness() {
        let duplicate_with_meta = document(
            r#"<text:p xml:id="same">before</text:p><text:p><text:meta-field xml:id="same">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&duplicate_with_meta).is_err());

        let duplicate_outside_meta = document(
            r#"<text:p xml:id="same">one</text:p><text:p xml:id="same">two</text:p><text:p><text:meta-field xml:id="unique">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&duplicate_outside_meta).is_err());

        let invalid_outside_meta = document(
            r#"<text:p xml:id="1invalid">one</text:p><text:p><text:meta-field xml:id="unique">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&invalid_outside_meta).is_err());

        let unique = document(
            r#"<text:p xml:id="paragraph">one</text:p><text:p><text:meta-field xml:id="field">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&unique).is_ok());
    }

    #[test]
    fn meta_field_content_constructor_enforces_resource_and_xml_bounds() {
        assert!(
            MetaFieldContent::new(vec![MetaFieldNode::Text(
                "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
            )])
            .is_err()
        );
        assert!(
            MetaFieldContent::new(vec![MetaFieldNode::Text("bad\u{1}control".to_string(),)])
                .is_err()
        );

        let mut nested = MetaFieldNode::Text("leaf".to_string());
        for _ in 0..=MAX_META_FIELD_DEPTH {
            nested = MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: Vec::new(),
                children: vec![nested],
            });
        }
        assert!(MetaFieldContent::new(vec![nested]).is_err());

        let oversized_attribute = MetaFieldElement {
            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
            local_name: "span".to_string(),
            attributes: vec![MetaFieldAttribute {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "style-name".to_string(),
                value: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
            }],
            children: Vec::new(),
        };
        assert!(MetaFieldContent::new(vec![MetaFieldNode::Element(oversized_attribute,)]).is_err());
    }

    #[test]
    fn note_body_content_enforces_block_root_grammar_and_projects_text() {
        let content = NoteBodyContent::new(vec![
            MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![
                    MetaFieldNode::Text("First ".to_string()),
                    MetaFieldNode::Element(MetaFieldElement {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "span".to_string(),
                        attributes: vec![MetaFieldAttribute {
                            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                            local_name: "style-name".to_string(),
                            value: "Emphasis".to_string(),
                        }],
                        children: vec![MetaFieldNode::Text("styled".to_string())],
                    }),
                ],
            }),
            MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "list".to_string(),
                attributes: Vec::new(),
                children: vec![MetaFieldNode::Element(MetaFieldElement {
                    namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                    local_name: "list-item".to_string(),
                    attributes: Vec::new(),
                    children: vec![MetaFieldNode::Element(MetaFieldElement {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "p".to_string(),
                        attributes: Vec::new(),
                        children: vec![MetaFieldNode::Text("Second".to_string())],
                    })],
                })],
            }),
        ])
        .unwrap();
        assert_eq!(content.display_text(), "First styled\nSecond");
        assert!(content.validate().is_ok());

        assert!(
            NoteBodyContent::new(vec![MetaFieldNode::Text("not a block".to_string(),)]).is_err()
        );
        assert!(
            NoteBodyContent::new(vec![MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: Vec::new(),
                children: vec![MetaFieldNode::Text("not a root block".to_string())],
            },)])
            .is_err()
        );
    }

    #[test]
    fn note_body_content_projects_odf_whitespace_controls() {
        let text_control = |local_name: &str, attributes: Vec<MetaFieldAttribute>| {
            MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: local_name.to_string(),
                attributes,
                children: Vec::new(),
            })
        };
        let content = NoteBodyContent::new(vec![MetaFieldNode::Element(MetaFieldElement {
            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
            local_name: "p".to_string(),
            attributes: Vec::new(),
            children: vec![
                MetaFieldNode::Text("A".to_string()),
                text_control(
                    "s",
                    vec![MetaFieldAttribute {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "c".to_string(),
                        value: "2".to_string(),
                    }],
                ),
                text_control("tab", Vec::new()),
                text_control("line-break", Vec::new()),
                MetaFieldNode::Text("B".to_string()),
            ],
        })])
        .unwrap();
        assert_eq!(content.display_text(), "A  \t\nB");

        let invalid = NoteBodyContent::new(vec![MetaFieldNode::Element(MetaFieldElement {
            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
            local_name: "p".to_string(),
            attributes: Vec::new(),
            children: vec![text_control(
                "s",
                vec![MetaFieldAttribute {
                    namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                    local_name: "c".to_string(),
                    value: "two".to_string(),
                }],
            )],
        })]);
        assert!(invalid.is_err());
    }

    #[test]
    fn meta_field_serialization_is_canonical_and_escaped() {
        let content = MetaFieldContent::new(vec![
            MetaFieldNode::Text("a<&".to_string()),
            MetaFieldNode::Element(MetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: vec![MetaFieldAttribute {
                    namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                    local_name: "style-name".to_string(),
                    value: "A&B\"".to_string(),
                }],
                children: vec![MetaFieldNode::Text("z>".to_string())],
            }),
        ])
        .unwrap();
        let field = DynamicTextField::MetaField {
            xml_id: "m1".to_string(),
            data_style_name: None,
            content,
        };
        let xml = field.to_xml_fragment().unwrap();
        assert!(xml.contains("a&lt;&amp;"));
        assert!(xml.contains("text:style-name=\"A&amp;B&quot;\""));
        assert!(xml.contains("z&gt;"));
    }
}
