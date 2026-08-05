//! Regression coverage for the layered dynamic-text content owner.

use super::*;
use crate::elements::field::{DatabaseField, DynamicTextField, FieldParser};

mod meta_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, MetaFieldContent, MetaFieldNode};

    fn content_xml(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:text><text:p>{body}</text:p></office:text></office:body>
</office:document-content>"#
        )
    }

    #[test]
    fn nested_meta_fields_and_dynamic_children_are_individually_mutable() {
        let xml = content_xml(
            r#"<text:meta-field xml:id="m">before<text:date>2026-07-18</text:date>after</text:meta-field>"#,
        );
        let spans = scan(&xml, None).unwrap().fields;
        assert_eq!(spans.len(), 2);
        assert!(spans[0].end > spans[1].end);

        let xml = replace_dynamic_text_field_xml(
            &xml,
            1,
            &DynamicTextField::PageVariableGet {
                number_format: None,
                display_text: "value".to_string(),
            },
        )
        .unwrap();
        assert!(xml.contains("text:meta-field"));
        assert!(xml.contains("text:page-variable-get"));
        assert!(!xml.contains("text:date"));

        let xml = remove_dynamic_text_field_xml(&xml, 0).unwrap();
        assert!(!xml.contains("text:meta-field"));
        assert!(!xml.contains("text:page-variable-get"));
    }

    #[test]
    fn meta_field_can_be_inserted_and_out_of_bounds_is_rejected() {
        let xml = content_xml("plain");
        let field = DynamicTextField::MetaField {
            xml_id: "inserted".to_string(),
            data_style_name: None,
            content: MetaFieldContent::new(vec![MetaFieldNode::Text("metadata".to_string())])
                .unwrap(),
        };
        let xml = insert_dynamic_text_field_xml(&xml, 0, &field).unwrap();
        assert!(xml.contains("xml:id=\"inserted\""));
        assert!(xml.contains(">metadata</text:meta-field>"));
        assert!(replace_dynamic_text_field_xml(&xml, 99, &field).is_err());
        assert!(remove_dynamic_text_field_xml(&xml, 99).is_err());
    }

    #[test]
    fn meta_field_mutation_rejects_document_wide_xml_id_collisions() {
        let xml = content_xml(r#"<text:span xml:id="existing">plain</text:span>"#);
        let field = DynamicTextField::MetaField {
            xml_id: "existing".to_string(),
            data_style_name: None,
            content: MetaFieldContent::new(vec![MetaFieldNode::Text("metadata".to_string())])
                .unwrap(),
        };
        assert!(insert_dynamic_text_field_xml(&xml, 0, &field).is_err());
    }
}

mod database_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DatabaseFieldKind, DatabaseSource, NonNegativeInteger};

    const XML: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>before<t:database-name t:table-name="old">old</t:database-name>after</t:p></o:text></o:body></o:document-content>"#;

    fn field(kind: DatabaseFieldKind, table: &str, text: &str) -> DatabaseField {
        DatabaseField {
            kind,
            source: DatabaseSource {
                database_name: None,
                table_name: table.into(),
                table_type: None,
                connection_resource: None,
            },
            column_name: None,
            condition: None,
            row_number: None,
            value: None,
            data_style_name: None,
            number_format: None,
            number_letter_sync: None,
            display_text: text.into(),
        }
    }

    #[test]
    fn database_fields_insert_replace_remove_and_check_bounds() {
        let replacement = field(DatabaseFieldKind::Name, "new", "new");
        let xml = replace_database_field_xml(XML, 0, &replacement).unwrap();
        assert!(xml.contains("text:table-name=\"new\""));
        let inserted =
            insert_database_field_xml(&xml, 0, &field(DatabaseFieldKind::Next, "next", ""))
                .unwrap();
        assert_eq!(
            FieldParser::parse_database_fields(&inserted).unwrap().len(),
            2
        );
        let removed = remove_database_field_xml(&inserted, 0).unwrap();
        assert_eq!(
            FieldParser::parse_database_fields(&removed).unwrap().len(),
            1
        );
        assert!(replace_database_field_xml(&removed, 9, &replacement).is_err());
        assert!(remove_database_field_xml(&removed, 9).is_err());
    }

    #[test]
    fn database_mutation_preserves_arbitrary_width_row_numbers() {
        let mut row = field(DatabaseFieldKind::RowSelect, "table", "");
        let huge = "18446744073709551616000000000000000000";
        row.row_number = Some(NonNegativeInteger::new(&format!("+000{huge}")).unwrap());
        let xml = insert_database_field_xml(XML, 0, &row).unwrap();
        assert!(xml.contains(&format!("text:row-number=\"{huge}\"")));
        let parsed = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(
            parsed.last().unwrap().row_number.as_ref().unwrap().as_str(),
            huge
        );
    }
}

mod dde_connection_mutation_tests {
    use super::*;

    const XML: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>before<t:dde-connection t:connection-name="old">cached</t:dde-connection>after</t:p></o:text></o:body></o:document-content>"#;

    #[test]
    fn dde_cached_fields_insert_replace_remove_and_check_bounds() {
        let replacement = DynamicTextField::DdeConnection {
            connection_name: "new".into(),
            display_text: "new cache".into(),
        };
        let xml = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(xml.contains("text:connection-name=\"new\""));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&xml).unwrap(),
            vec![replacement.clone()]
        );
        let inserted = insert_dynamic_text_field_xml(&xml, 0, &replacement).unwrap();
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&inserted)
                .unwrap()
                .len(),
            2
        );
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&removed)
                .unwrap()
                .len(),
            1
        );
        assert!(replace_dynamic_text_field_xml(&removed, 9, &replacement).is_err());

        let missing = XML.replace(" t:connection-name=\"old\"", "");
        assert!(FieldParser::parse_dynamic_text_fields(&missing).is_err());
        let spoof = XML
            .replace(
                "<t:dde-connection",
                "<fake:dde-connection xmlns:fake=\"urn:not-text\"",
            )
            .replace("</t:dde-connection>", "</fake:dde-connection>");
        assert!(
            FieldParser::parse_dynamic_text_fields(&spoof)
                .unwrap()
                .is_empty()
        );
        let oversized = DynamicTextField::DdeConnection {
            connection_name: "x".repeat(65_537),
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
    }
}

mod fixed_page_date_time_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, FieldDateValue, PageContinuationSelection};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><t:p>before<t:date t:date-value="2024-01-01">old</t:date>after</t:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn fixed_page_date_time_mutation_insert_replace_remove_is_namespace_aware() {
        let replacement = DynamicTextField::PageContinuation {
            select_page: PageContinuationSelection::Next,
            string_value: Some("continued".to_string()),
            display_text: "cached".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:page-continuation"));
        assert!(replaced.contains("</text:page-continuation>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::Date {
                value: Some(FieldDateValue::new("2024-12-31").unwrap()),
                adjustment: None,
                fixed: Some(true),
                data_style_name: None,
                display_text: "year end".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:date"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:page-continuation"));
        assert!(removed.contains("<text:date"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

mod page_variable_family_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, SequenceNumberFormat};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:q="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><q:p>before<q:page-variable-set q:active="true"/>after</q:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn page_variable_family_mutation_inserts_replaces_removes_and_checks_bounds() {
        let getter = DynamicTextField::PageVariableGet {
            number_format: Some(SequenceNumberFormat::new("1", None).unwrap()),
            display_text: "12".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &getter).unwrap();
        assert!(replaced.contains("before<text:page-variable-get"));
        assert!(replaced.contains("</text:page-variable-get>after"));

        let setter = DynamicTextField::PageVariableSet {
            active: Some(false),
            page_adjust: Some(-4),
            display_text: String::new(),
        };
        let inserted = insert_dynamic_text_field_xml(&replaced, 0, &setter).unwrap();
        assert!(inserted.contains("<text:page-variable-set"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:page-variable-get"));
        assert!(removed.contains("text:page-variable-set"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

mod document_metadata_fixed_field_mutation_tests {
    use super::*;
    use crate::elements::field::{
        DynamicTextField, FieldDuration, MetadataFieldKind, MetadataFieldValue,
    };

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:m="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><m:p>before<m:editing-cycles m:fixed="true">3</m:editing-cycles>after</m:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn document_metadata_fixed_field_mutation_is_namespace_aware_and_bounded() {
        let replacement = DynamicTextField::DocumentMetadata {
            kind: MetadataFieldKind::EditingDuration,
            value: Some(MetadataFieldValue::Duration(
                FieldDuration::new("PT3H").unwrap(),
            )),
            fixed: Some(true),
            data_style_name: Some("Elapsed".to_string()),
            display_text: "three hours".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:editing-duration"));
        assert!(replaced.contains("</text:editing-duration>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::DocumentMetadata {
                kind: MetadataFieldKind::ModificationTime,
                value: None,
                fixed: None,
                data_style_name: None,
                display_text: "now".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:modification-time"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:editing-duration"));
        assert!(removed.contains("text:modification-time"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

mod document_identity_fixed_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, IdentityFieldKind};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:i="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><i:p>before<i:title i:fixed="true">old</i:title>after</i:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn document_identity_fixed_field_mutation_is_namespace_aware_and_bounded() {
        let replacement = DynamicTextField::DocumentIdentity {
            kind: IdentityFieldKind::Creator,
            fixed: Some(false),
            display_text: "new creator".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:creator"));
        assert!(replaced.contains("</text:creator>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::DocumentIdentity {
                kind: IdentityFieldKind::Keywords,
                fixed: None,
                display_text: "one, two".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:keywords"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:creator"));
        assert!(removed.contains("text:keywords"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

mod user_defined_metadata_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, UserDefinedMetadataValues};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:u="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><u:p>before<u:user-defined u:name="old">old</u:user-defined>after</u:p></o:text></o:body>
    </o:document-content>"#;

    fn field(name: &str, display_text: &str) -> DynamicTextField {
        DynamicTextField::UserDefinedMetadata {
            name: name.to_string(),
            values: UserDefinedMetadataValues {
                string: Some(display_text.to_string()),
                ..UserDefinedMetadataValues::default()
            },
            fixed: Some(true),
            data_style_name: None,
            display_text: display_text.to_string(),
        }
    }

    #[test]
    fn user_defined_metadata_field_mutation_is_namespace_aware_and_bounded() {
        let replaced = replace_dynamic_text_field_xml(XML, 0, &field("new", "cached")).unwrap();
        assert!(replaced.contains("before<text:user-defined"));
        assert!(replaced.contains("</text:user-defined>after"));

        let inserted =
            insert_dynamic_text_field_xml(&replaced, 0, &field("second", "two")).unwrap();
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:name=\"new\""));
        assert!(removed.contains("text:name=\"second\""));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}
