use litchi_odf::elements::field::{
    OdfCalculatedFieldValue, OdfCrossReferenceFormat, OdfDocumentStatisticKind, OdfDropDownLabel,
    OdfDynamicTextField, OdfFieldValueType, OdfFormulaFieldDisplay, OdfMeasureKind,
    OdfNoteReferenceClass, OdfNoteReferenceFormat, OdfPlaceholderType, OdfSequenceNumberFormat,
    OdfSequenceReferenceFormat, OdfStatisticNumberFormat, OdfUserFieldDisplay,
    OdfVariableSetDisplay,
};
use litchi_odf::{Document, DocumentBuilder, MutableDocument};
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document(inner: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.2"><o:body><o:text><t:p>{inner}</t:p></o:text></o:body></o:document-content>"#
    );
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    write!(
        zip,
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.2"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#
    )
    .unwrap();
    zip.finish().unwrap();
    Document::from_bytes(output.into_inner()).unwrap()
}

#[test]
fn parses_typed_dynamic_text_fields_in_document_order_without_evaluation() {
    let document = document(concat!(
        r#"<t:placeholder t:placeholder-type="text-box" t:description="Prompt">Enter &amp; edit</t:placeholder>"#,
        r#"<t:conditional-text t:condition="of:=WEBSERVICE(&quot;https://never.invalid&quot;)" t:string-value-if-true="yes" t:string-value-if-false="no" t:current-value="1">cached yes</t:conditional-text>"#,
        r#"<t:hidden-text t:condition="ooow:flag" t:string-value="secret" t:is-hidden="false">visible cache</t:hidden-text>"#,
        r#"<t:hidden-paragraph t:condition="ooow:hide" t:is-hidden="true">paragraph cache</t:hidden-paragraph>"#,
    ));

    let fields = document.dynamic_text_fields().unwrap();
    assert_eq!(fields.len(), 4);
    assert!(matches!(
        &fields[0],
        OdfDynamicTextField::Placeholder {
            placeholder_type: OdfPlaceholderType::TextBox,
            description: Some(description),
            display_text,
        } if description == "Prompt" && display_text == "Enter & edit"
    ));
    assert!(matches!(
        &fields[1],
        OdfDynamicTextField::ConditionalText {
            condition,
            value_if_true,
            value_if_false,
            current_value: Some(true),
            display_text,
        } if condition.contains("WEBSERVICE")
            && value_if_true == "yes"
            && value_if_false == "no"
            && display_text == "cached yes"
    ));
    assert!(matches!(
        &fields[2],
        OdfDynamicTextField::HiddenText {
            string_value,
            is_hidden: Some(false),
            ..
        } if string_value == "secret"
    ));
    assert_eq!(fields[3].display_text(), "paragraph cache");
}

#[test]
fn rejects_missing_required_attributes_invalid_enums_and_invalid_booleans() {
    for field in [
        r#"<t:placeholder>missing type</t:placeholder>"#,
        r#"<t:placeholder t:placeholder-type="video">bad type</t:placeholder>"#,
        r#"<t:conditional-text t:condition="x" t:string-value-if-true="yes">missing false</t:conditional-text>"#,
        r#"<t:conditional-text t:condition="x" t:string-value-if-true="yes" t:string-value-if-false="no" t:current-value="yes">bad bool</t:conditional-text>"#,
        r#"<t:hidden-text t:condition="x">missing string</t:hidden-text>"#,
        r#"<t:hidden-paragraph t:is-hidden="true">missing condition</t:hidden-paragraph>"#,
    ] {
        assert!(
            document(field).dynamic_text_fields().is_err(),
            "accepted {field}"
        );
    }
}

#[test]
fn accepts_namespace_aliases_and_ignores_spoofed_vocabulary() {
    let content = format!(
        r#"<x:placeholder xmlns:x="{TEXT}" x:placeholder-type="image">image</x:placeholder><fake:conditional-text xmlns:fake="urn:not-text" fake:condition="x" fake:string-value-if-true="a" fake:string-value-if-false="b">spoof</fake:conditional-text>"#
    );
    let fields = document(&content).dynamic_text_fields().unwrap();
    assert_eq!(fields.len(), 1);
    assert!(matches!(
        fields[0],
        OdfDynamicTextField::Placeholder {
            placeholder_type: OdfPlaceholderType::Image,
            ..
        }
    ));
}

#[test]
fn serializes_every_dynamic_field_with_escaping_and_round_trips_inert_values() {
    let fields = vec![
        OdfDynamicTextField::Placeholder {
            placeholder_type: OdfPlaceholderType::Object,
            description: Some("Choose \"A&B\" <object>".to_string()),
            display_text: "cached <object> & text".to_string(),
        },
        OdfDynamicTextField::ConditionalText {
            condition: "of:=WEBSERVICE(\"https://never.invalid/?a=1&b=2\")<3".to_string(),
            value_if_true: "yes & <true>".to_string(),
            value_if_false: "no \"false\"".to_string(),
            current_value: Some(true),
            display_text: "cached & true".to_string(),
        },
        OdfDynamicTextField::HiddenText {
            condition: "ooow:flag & 1".to_string(),
            string_value: "secret <&>".to_string(),
            is_hidden: Some(false),
            display_text: "visible <cache>".to_string(),
        },
        OdfDynamicTextField::HiddenParagraph {
            condition: "ooow:hide > 0".to_string(),
            is_hidden: None,
            display_text: "paragraph & cache".to_string(),
        },
    ];

    for field in fields {
        let fragment = field.to_xml_fragment().unwrap();
        assert!(
            fragment.contains(r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#)
        );
        assert!(!fragment.contains("https://never.invalid/?a=1&b=2"));
        let parsed = document(&fragment).dynamic_text_fields().unwrap();
        assert_eq!(parsed, vec![field]);
    }
}

#[test]
fn writer_rejects_empty_conditions_forbidden_xml_characters_and_oversized_values() {
    let empty_condition = OdfDynamicTextField::HiddenParagraph {
        condition: " \t".to_string(),
        is_hidden: None,
        display_text: String::new(),
    };
    assert!(empty_condition.to_xml_fragment().is_err());

    let forbidden_character = OdfDynamicTextField::HiddenText {
        condition: "ooow:flag".to_string(),
        string_value: "bad\u{0}value".to_string(),
        is_hidden: None,
        display_text: String::new(),
    };
    assert!(forbidden_character.validate().is_err());

    let oversized = OdfDynamicTextField::Placeholder {
        placeholder_type: OdfPlaceholderType::Text,
        description: None,
        display_text: "x".repeat(65_537),
    };
    assert!(oversized.to_xml_fragment().is_err());
}

#[test]
fn mutable_document_inserts_replaces_and_removes_fields_without_rewriting_neighbors() {
    let source = document(concat!(
        r#"before<t:span t:style-name="Keep">unchanged &amp; exact</t:span>"#,
        r#"<t:placeholder t:placeholder-type="text">old</t:placeholder>after"#,
    ));
    let mut mutable = MutableDocument::from_document(source).unwrap();
    let conditional = OdfDynamicTextField::ConditionalText {
        condition: "of:=1<2 & 3>2".to_string(),
        value_if_true: "yes".to_string(),
        value_if_false: "no".to_string(),
        current_value: Some(true),
        display_text: "yes & cached".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &conditional).unwrap();

    let hidden = OdfDynamicTextField::HiddenText {
        condition: "ooow:flag".to_string(),
        string_value: "secret".to_string(),
        is_hidden: Some(false),
        display_text: "visible".to_string(),
    };
    let replaced = mutable.replace_dynamic_text_field(0, &hidden).unwrap();
    assert!(matches!(replaced, OdfDynamicTextField::Placeholder { .. }));
    let removed = mutable.remove_dynamic_text_field(1).unwrap();
    assert_eq!(removed, conditional);

    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(round_trip.dynamic_text_fields().unwrap(), vec![hidden]);
    let content = String::from_utf8(round_trip.get_file("content.xml").unwrap()).unwrap();
    assert!(
        content.contains(r#"before<t:span t:style-name="Keep">unchanged &amp; exact</t:span>"#)
    );
    assert!(content.contains("after"));
}

#[test]
fn insertion_supports_empty_prefixed_paragraphs_and_builder_round_trips() {
    let field = OdfDynamicTextField::HiddenParagraph {
        condition: "ooow:hide".to_string(),
        is_hidden: Some(true),
        display_text: "cached".to_string(),
    };
    let mut mutable = MutableDocument::from_document(document("")).unwrap();
    mutable.insert_dynamic_text_field(0, &field).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![field.clone()]
    );

    let mut builder = DocumentBuilder::new();
    builder.add_dynamic_text_field(&field).unwrap();
    let built = Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(built.dynamic_text_fields().unwrap(), vec![field]);
}

#[test]
fn mutation_rejects_out_of_bounds_targets_without_changing_content() {
    let mut mutable = MutableDocument::from_document(document("plain")).unwrap();
    let before = Document::from_bytes(mutable.to_bytes().unwrap())
        .unwrap()
        .get_file("content.xml")
        .unwrap();
    let field = OdfDynamicTextField::Placeholder {
        placeholder_type: OdfPlaceholderType::TextBox,
        description: None,
        display_text: "prompt".to_string(),
    };
    assert!(mutable.insert_dynamic_text_field(1, &field).is_err());
    assert!(mutable.replace_dynamic_text_field(0, &field).is_err());
    assert!(mutable.remove_dynamic_text_field(0).is_err());
    let after = Document::from_bytes(mutable.to_bytes().unwrap())
        .unwrap()
        .get_file("content.xml")
        .unwrap();
    assert_eq!(after, before);
}

#[test]
fn sequence_fields_parse_serialize_and_retain_formulas_inertly() {
    let field = OdfDynamicTextField::Sequence {
        name: "Figure & Diagram".to_string(),
        formula: Some("ooow:Figure+WEBSERVICE(\"https://never.invalid/?a=1&b=2\")".to_string()),
        number_format: Some(OdfSequenceNumberFormat::new("A", Some(true)).unwrap()),
        reference_name: Some("fig<&>1".to_string()),
        display_text: "Figure <1> & cached".to_string(),
    };
    let fragment = field.to_xml_fragment().unwrap();
    assert!(fragment.contains("text:sequence"));
    assert!(fragment.contains("xmlns:style="));
    assert!(!fragment.contains("?a=1&b=2"));
    let parsed = document(&fragment).dynamic_text_fields().unwrap();
    assert_eq!(parsed, vec![field]);
}

#[test]
fn sequence_fields_support_document_mutation_and_namespace_aliases() {
    let source = document(
        r#"<t:sequence xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" t:name="Old" t:formula="ooow:1+1" s:num-format="a" s:num-letter-sync="1" t:ref-name="old-ref">1</t:sequence>"#,
    );
    let old = OdfDynamicTextField::Sequence {
        name: "Old".to_string(),
        formula: Some("ooow:1+1".to_string()),
        number_format: Some(OdfSequenceNumberFormat::new("a", Some(true)).unwrap()),
        reference_name: Some("old-ref".to_string()),
        display_text: "1".to_string(),
    };
    assert_eq!(source.dynamic_text_fields().unwrap(), vec![old.clone()]);

    let replacement = OdfDynamicTextField::Sequence {
        name: "New".to_string(),
        formula: Some("of:=2+2".to_string()),
        number_format: Some(OdfSequenceNumberFormat::new("I", None).unwrap()),
        reference_name: None,
        display_text: "IV".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert_eq!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        old
    );
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone()]
    );

    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
    assert!(
        Document::from_bytes(mutable.to_bytes().unwrap())
            .unwrap()
            .dynamic_text_fields()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn sequence_numbering_rejects_schema_invalid_letter_sync_combinations() {
    assert!(OdfSequenceNumberFormat::new("1", Some(true)).is_err());
    assert!(OdfSequenceNumberFormat::new("i", Some(false)).is_err());
    assert!(document(
        r#"<t:sequence xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" t:name="x" s:num-letter-sync="true">1</t:sequence>"#
    )
    .dynamic_text_fields()
    .is_err());
    assert!(document(
        r#"<t:sequence xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" t:name="x" s:num-format="1" s:num-letter-sync="true">1</t:sequence>"#
    )
    .dynamic_text_fields()
    .is_err());
}

#[test]
fn sequence_references_round_trip_every_schema_format() {
    let formats = [
        OdfSequenceReferenceFormat::Page,
        OdfSequenceReferenceFormat::Chapter,
        OdfSequenceReferenceFormat::Direction,
        OdfSequenceReferenceFormat::Text,
        OdfSequenceReferenceFormat::CategoryAndValue,
        OdfSequenceReferenceFormat::Caption,
        OdfSequenceReferenceFormat::Value,
    ];
    for format in formats.into_iter().map(Some).chain(std::iter::once(None)) {
        let field = OdfDynamicTextField::SequenceReference {
            reference_name: "figure<&>1".to_string(),
            reference_format: format,
            display_text: "Figure <1> & cached".to_string(),
        };
        let fragment = field.to_xml_fragment().unwrap();
        let parsed = document(&fragment).dynamic_text_fields().unwrap();
        assert_eq!(parsed, vec![field]);
    }
}

#[test]
fn sequence_references_support_namespace_aware_document_mutation() {
    let source = document(
        r#"<t:sequence-ref t:ref-name="old" t:reference-format="caption">Old caption</t:sequence-ref>"#,
    );
    let old = OdfDynamicTextField::SequenceReference {
        reference_name: "old".to_string(),
        reference_format: Some(OdfSequenceReferenceFormat::Caption),
        display_text: "Old caption".to_string(),
    };
    assert_eq!(source.dynamic_text_fields().unwrap(), vec![old.clone()]);

    let replacement = OdfDynamicTextField::SequenceReference {
        reference_name: "new&ref".to_string(),
        reference_format: Some(OdfSequenceReferenceFormat::CategoryAndValue),
        display_text: "Figure 2".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert_eq!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        old
    );
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone()]
    );

    let inserted = OdfDynamicTextField::SequenceReference {
        reference_name: "page-ref".to_string(),
        reference_format: Some(OdfSequenceReferenceFormat::Page),
        display_text: "12".to_string(),
    };
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
    let final_document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        final_document.dynamic_text_fields().unwrap(),
        vec![inserted]
    );
}

#[test]
fn sequence_references_reject_missing_empty_and_invalid_reference_metadata() {
    assert!(
        document(r#"<t:sequence-ref>missing</t:sequence-ref>"#)
            .dynamic_text_fields()
            .is_err()
    );
    assert!(
        document(
            r#"<t:sequence-ref t:ref-name="x" t:reference-format="number">bad</t:sequence-ref>"#
        )
        .dynamic_text_fields()
        .is_err()
    );
    let empty = OdfDynamicTextField::SequenceReference {
        reference_name: String::new(),
        reference_format: None,
        display_text: String::new(),
    };
    assert!(empty.to_xml_fragment().is_err());
    let oversized = OdfDynamicTextField::SequenceReference {
        reference_name: "x".repeat(65_537),
        reference_format: Some(OdfSequenceReferenceFormat::Text),
        display_text: String::new(),
    };
    assert!(oversized.validate().is_err());
}

#[test]
fn calculated_variable_fields_round_trip_typed_values_and_inert_formulas() {
    let fields = vec![
        OdfDynamicTextField::VariableSet {
            name: "Total".to_string(),
            formula: Some("of:=SUM([.A1:.A3])&WEBSERVICE(\"https://never.invalid\")".to_string()),
            value: OdfCalculatedFieldValue::Currency {
                value: "12.50".to_string(),
                currency: Some("C&Y".to_string()),
            },
            display: Some(OdfVariableSetDisplay::Value),
            data_style_name: Some("Money&Style".to_string()),
            display_text: "$12.50 <cached>".to_string(),
        },
        OdfDynamicTextField::VariableGet {
            name: "Total".to_string(),
            display: Some(OdfFormulaFieldDisplay::Formula),
            data_style_name: None,
            display_text: "Total".to_string(),
        },
        OdfDynamicTextField::Expression {
            formula: Some("of:=1<2".to_string()),
            value: Some(OdfCalculatedFieldValue::Boolean(true)),
            display: Some(OdfFormulaFieldDisplay::Value),
            data_style_name: None,
            display_text: "true & cached".to_string(),
        },
    ];
    for field in fields {
        let fragment = field.to_xml_fragment().unwrap();
        assert!(!fragment.contains("1<2"));
        assert_eq!(
            document(&fragment).dynamic_text_fields().unwrap(),
            vec![field]
        );
    }
}

#[test]
fn calculated_field_values_cover_every_odf_value_group() {
    let values = [
        OdfCalculatedFieldValue::Float("-1.25E2".to_string()),
        OdfCalculatedFieldValue::Percentage("0.25".to_string()),
        OdfCalculatedFieldValue::Currency {
            value: "5".to_string(),
            currency: None,
        },
        OdfCalculatedFieldValue::Date("2026-07-18T10:20:30Z".to_string()),
        OdfCalculatedFieldValue::Time("PT1H2M3.5S".to_string()),
        OdfCalculatedFieldValue::Boolean(false),
        OdfCalculatedFieldValue::String(None),
    ];
    for value in values {
        let field = OdfDynamicTextField::Expression {
            formula: None,
            value: Some(value),
            display: None,
            data_style_name: None,
            display_text: String::new(),
        };
        let fragment = field.to_xml_fragment().unwrap();
        assert_eq!(
            document(&fragment).dynamic_text_fields().unwrap(),
            vec![field]
        );
    }
}

#[test]
fn calculated_variables_support_namespace_aware_document_mutation() {
    let source = document(
        r#"<t:expression xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" t:formula="of:=1+1" o:value-type="float" o:value="2">2</t:expression>"#,
    );
    let replacement = OdfDynamicTextField::VariableSet {
        name: "Counter".to_string(),
        formula: Some("ooow:Counter+1".to_string()),
        value: OdfCalculatedFieldValue::Float("3".to_string()),
        display: Some(OdfVariableSetDisplay::None),
        data_style_name: None,
        display_text: String::new(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::Expression { .. }
    ));
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone()]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
    assert!(
        Document::from_bytes(mutable.to_bytes().unwrap())
            .unwrap()
            .dynamic_text_fields()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn calculated_fields_reject_mismatched_value_groups_displays_and_lexicals() {
    for invalid in [
        r#"<t:variable-set t:name="x">missing value</t:variable-set>"#,
        r#"<t:variable-get t:name="x" t:display="none">bad display</t:variable-get>"#,
        r#"<t:expression xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:value="1">missing type</t:expression>"#,
        r#"<t:expression xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:value-type="boolean" o:boolean-value="yes">bad bool</t:expression>"#,
        r#"<t:expression xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:value-type="date" o:date-value="2026-99-99">bad date</t:expression>"#,
        r#"<t:expression xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:value-type="float" o:value="abc">bad number</t:expression>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
}

#[test]
fn interactive_input_fields_round_trip_all_typed_attributes() {
    let fields = vec![
        OdfDynamicTextField::VariableInput {
            name: "Amount".to_string(),
            description: Some("Enter <amount> & currency".to_string()),
            value_type: OdfFieldValueType::Currency,
            display: Some(OdfVariableSetDisplay::Value),
            data_style_name: Some("Money&Style".to_string()),
            display_text: "$5 <cached>".to_string(),
        },
        OdfDynamicTextField::UserFieldGet {
            name: "Company".to_string(),
            display: Some(OdfUserFieldDisplay::None),
            data_style_name: None,
            display_text: "Example & Co".to_string(),
        },
        OdfDynamicTextField::UserFieldInput {
            name: "Company".to_string(),
            description: Some("Edit \"company\"".to_string()),
            data_style_name: Some("TextStyle".to_string()),
            display_text: "Example <Co>".to_string(),
        },
        OdfDynamicTextField::TextInput {
            description: Some("Free & safe".to_string()),
            display_text: "cached <input>".to_string(),
        },
    ];
    for field in fields {
        let fragment = field.to_xml_fragment().unwrap();
        assert_eq!(
            document(&fragment).dynamic_text_fields().unwrap(),
            vec![field]
        );
    }
}

#[test]
fn drop_down_fields_round_trip_and_support_namespace_aware_mutation() {
    let source = document(
        r#"<t:drop-down t:name="old"><t:label t:value="old" t:current-selected="true"/>old</t:drop-down>"#,
    );
    let old = OdfDynamicTextField::DropDown {
        name: "old".to_string(),
        labels: vec![OdfDropDownLabel {
            value: Some("old".to_string()),
            current_selected: Some(true),
        }],
        display_text: "old".to_string(),
    };
    assert_eq!(source.dynamic_text_fields().unwrap(), vec![old.clone()]);

    let replacement = OdfDynamicTextField::DropDown {
        name: "Priority & state".to_string(),
        labels: vec![
            OdfDropDownLabel {
                value: Some("Low".to_string()),
                current_selected: Some(false),
            },
            OdfDropDownLabel {
                value: Some("High & urgent".to_string()),
                current_selected: Some(true),
            },
        ],
        display_text: "High & urgent".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert_eq!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        old
    );
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone()]
    );

    let inserted = OdfDynamicTextField::DropDown {
        name: "Status".to_string(),
        labels: vec![OdfDropDownLabel {
            value: Some("Open".to_string()),
            current_selected: None,
        }],
        display_text: "Open".to_string(),
    };
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted.clone()]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
    assert_eq!(
        Document::from_bytes(mutable.to_bytes().unwrap())
            .unwrap()
            .dynamic_text_fields()
            .unwrap(),
        vec![inserted]
    );
}

#[test]
fn inline_script_metadata_round_trips_and_supports_namespace_aware_mutation() {
    let source = document(
        r#"<t:script xmlns:l="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0" l:type="simple" l:href="https://example.invalid/never-open?one=1&amp;two=2" s:language="application/javascript"/>"#,
    );
    let old = OdfDynamicTextField::Script {
        href: Some("https://example.invalid/never-open?one=1&two=2".to_string()),
        language: Some("application/javascript".to_string()),
        content: String::new(),
    };
    assert_eq!(source.dynamic_text_fields().unwrap(), vec![old.clone()]);

    let replacement = OdfDynamicTextField::Script {
        href: None,
        language: Some("text/x-basic".to_string()),
        content: "REM stored macro payload".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert_eq!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        old
    );
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone()]
    );

    let inserted = OdfDynamicTextField::Script {
        href: Some("vnd.example:stored-only".to_string()),
        language: None,
        content: String::new(),
    };
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted.clone()]
    );
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
    assert_eq!(
        Document::from_bytes(mutable.to_bytes().unwrap())
            .unwrap()
            .dynamic_text_fields()
            .unwrap(),
        vec![inserted]
    );
}

#[test]
fn variable_input_round_trips_every_odf_value_type() {
    for value_type in [
        OdfFieldValueType::Float,
        OdfFieldValueType::Time,
        OdfFieldValueType::Date,
        OdfFieldValueType::Percentage,
        OdfFieldValueType::Currency,
        OdfFieldValueType::Boolean,
        OdfFieldValueType::String,
    ] {
        let field = OdfDynamicTextField::VariableInput {
            name: "Value".to_string(),
            description: None,
            value_type,
            display: None,
            data_style_name: None,
            display_text: String::new(),
        };
        assert_eq!(
            document(&field.to_xml_fragment().unwrap())
                .dynamic_text_fields()
                .unwrap(),
            vec![field]
        );
    }
}

#[test]
fn interactive_fields_support_namespace_aware_mutation() {
    let source = document(r#"<t:user-field-get t:name="Company">Old</t:user-field-get>"#);
    let replacement = OdfDynamicTextField::VariableInput {
        name: "Count".to_string(),
        description: Some("Count".to_string()),
        value_type: OdfFieldValueType::Float,
        display: Some(OdfVariableSetDisplay::None),
        data_style_name: None,
        display_text: String::new(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::UserFieldGet { .. }
    ));
    let input = OdfDynamicTextField::TextInput {
        description: None,
        display_text: "Prompt".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &input).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), input]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
}

#[test]
fn interactive_fields_reject_required_type_name_and_display_violations() {
    for invalid in [
        r#"<t:variable-input t:name="x">missing type</t:variable-input>"#,
        r#"<t:variable-input xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" t:name="x" o:value-type="number">bad type</t:variable-input>"#,
        r#"<t:variable-input xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" t:name="x" o:value-type="float" o:value="1">value not allowed</t:variable-input>"#,
        r#"<t:user-field-get t:display="value">missing name</t:user-field-get>"#,
        r#"<t:user-field-get t:name="x" t:display="caption">bad display</t:user-field-get>"#,
        r#"<t:user-field-input>missing name</t:user-field-input>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
    let oversized = OdfDynamicTextField::TextInput {
        description: Some("x".repeat(65_537)),
        display_text: String::new(),
    };
    assert!(oversized.validate().is_err());
}

#[test]
fn table_formula_round_trips_inert_formula_display_and_style() {
    let field = OdfDynamicTextField::TableFormula {
        formula: Some(
            "of:=SUM([.A1:.A3])&WEBSERVICE(\"https://never.invalid/?a=1&b=2\")".to_string(),
        ),
        display: Some(OdfFormulaFieldDisplay::Formula),
        data_style_name: Some("Number<&>Style".to_string()),
        display_text: "SUM <cached> & safe".to_string(),
    };
    let fragment = field.to_xml_fragment().unwrap();
    assert!(fragment.contains("text:table-formula"));
    assert!(fragment.contains("xmlns:style="));
    assert!(!fragment.contains("?a=1&b=2"));
    assert_eq!(
        document(&fragment).dynamic_text_fields().unwrap(),
        vec![field]
    );
}

#[test]
fn table_formula_supports_namespace_aware_insert_replace_and_remove() {
    let source = document(
        r#"<t:table-formula xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" t:formula="ooow:=1+1" t:display="value" s:data-style-name="N1">2</t:table-formula>"#,
    );
    let replacement = OdfDynamicTextField::TableFormula {
        formula: Some("of:=2+2".to_string()),
        display: Some(OdfFormulaFieldDisplay::Value),
        data_style_name: None,
        display_text: "4".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::TableFormula { .. }
    ));
    let inserted = OdfDynamicTextField::TableFormula {
        formula: None,
        display: None,
        data_style_name: None,
        display_text: "cached".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
}

#[test]
fn table_formula_rejects_invalid_display_cached_value_groups_and_bounds() {
    for invalid in [
        r#"<t:table-formula t:display="none">bad display</t:table-formula>"#,
        r#"<t:table-formula xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:value-type="float" o:value="1">bad cache</t:table-formula>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
    let oversized = OdfDynamicTextField::TableFormula {
        formula: Some("x".repeat(65_537)),
        display: None,
        data_style_name: None,
        display_text: String::new(),
    };
    assert!(oversized.to_xml_fragment().is_err());
    let forbidden = OdfDynamicTextField::TableFormula {
        formula: Some("of:=1\u{0}+1".to_string()),
        display: None,
        data_style_name: None,
        display_text: String::new(),
    };
    assert!(forbidden.validate().is_err());
}

#[test]
fn measure_fields_round_trip_every_kind_and_escape_cached_text() {
    for kind in [
        OdfMeasureKind::Value,
        OdfMeasureKind::Unit,
        OdfMeasureKind::Gap,
    ] {
        let field = OdfDynamicTextField::Measure {
            kind,
            display_text: "12 <cm> & cached".to_string(),
        };
        let fragment = field.to_xml_fragment().unwrap();
        assert!(fragment.contains("text:measure"));
        assert!(fragment.contains("&lt;cm&gt; &amp; cached"));
        assert_eq!(
            document(&fragment).dynamic_text_fields().unwrap(),
            vec![field]
        );
    }
}

#[test]
fn measure_fields_support_namespace_aware_insert_replace_and_remove() {
    let source = document(r#"<t:measure t:kind="unit">cm</t:measure>"#);
    let replacement = OdfDynamicTextField::Measure {
        kind: OdfMeasureKind::Value,
        display_text: "12.5".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::Measure {
            kind: OdfMeasureKind::Unit,
            ..
        }
    ));
    let inserted = OdfDynamicTextField::Measure {
        kind: OdfMeasureKind::Gap,
        display_text: " ".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
}

#[test]
fn measure_fields_reject_missing_spoofed_invalid_and_extra_attributes() {
    for invalid in [
        r#"<t:measure>missing</t:measure>"#,
        r#"<t:measure t:kind="distance">bad kind</t:measure>"#,
        r#"<t:measure xmlns:x="urn:not-text" x:kind="unit">spoof</t:measure>"#,
        r#"<t:measure xmlns:x="urn:extension" t:kind="unit" x:extra="1">extra</t:measure>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
    let oversized = OdfDynamicTextField::Measure {
        kind: OdfMeasureKind::Value,
        display_text: "x".repeat(65_537),
    };
    assert!(oversized.validate().is_err());
    let forbidden = OdfDynamicTextField::Measure {
        kind: OdfMeasureKind::Unit,
        display_text: "cm\u{0}".to_string(),
    };
    assert!(forbidden.to_xml_fragment().is_err());
}

#[test]
fn mark_and_bookmark_references_round_trip_all_schema_formats() {
    let formats = [
        OdfCrossReferenceFormat::Page,
        OdfCrossReferenceFormat::Chapter,
        OdfCrossReferenceFormat::Direction,
        OdfCrossReferenceFormat::Text,
        OdfCrossReferenceFormat::NumberNoSuperior,
        OdfCrossReferenceFormat::NumberAllSuperior,
        OdfCrossReferenceFormat::Number,
    ];
    for format in formats.into_iter().map(Some).chain(std::iter::once(None)) {
        for field in [
            OdfDynamicTextField::Reference {
                reference_name: Some("mark<&>1".to_string()),
                reference_format: format,
                display_text: "Mark <1> & cached".to_string(),
            },
            OdfDynamicTextField::BookmarkReference {
                reference_name: Some("bookmark<&>1".to_string()),
                reference_format: format,
                display_text: "Bookmark <1> & cached".to_string(),
            },
        ] {
            let fragment = field.to_xml_fragment().unwrap();
            assert_eq!(
                document(&fragment).dynamic_text_fields().unwrap(),
                vec![field]
            );
        }
    }
}

#[test]
fn note_references_round_trip_classes_formats_and_schema_optional_targets() {
    let formats = [
        OdfNoteReferenceFormat::Page,
        OdfNoteReferenceFormat::Chapter,
        OdfNoteReferenceFormat::Direction,
        OdfNoteReferenceFormat::Text,
    ];
    for note_class in [
        OdfNoteReferenceClass::Footnote,
        OdfNoteReferenceClass::Endnote,
    ] {
        for format in formats.into_iter().map(Some).chain(std::iter::once(None)) {
            let field = OdfDynamicTextField::NoteReference {
                reference_name: None,
                note_class,
                reference_format: format,
                display_text: "cached note".to_string(),
            };
            assert_eq!(
                document(&field.to_xml_fragment().unwrap())
                    .dynamic_text_fields()
                    .unwrap(),
                vec![field]
            );
        }
    }
}

#[test]
fn cross_references_support_namespace_aware_insert_replace_and_remove() {
    let source = document(
        r#"<t:bookmark-ref t:ref-name="old" t:reference-format="page">1</t:bookmark-ref>"#,
    );
    let replacement = OdfDynamicTextField::Reference {
        reference_name: Some("new&mark".to_string()),
        reference_format: Some(OdfCrossReferenceFormat::NumberAllSuperior),
        display_text: "1.2.3".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::BookmarkReference { .. }
    ));
    let inserted = OdfDynamicTextField::NoteReference {
        reference_name: Some("note-1".to_string()),
        note_class: OdfNoteReferenceClass::Footnote,
        reference_format: Some(OdfNoteReferenceFormat::Text),
        display_text: "footnote".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
}

#[test]
fn cross_references_reject_invalid_classes_formats_namespaces_attributes_and_bounds() {
    for invalid in [
        r#"<t:reference-ref t:reference-format="caption">bad format</t:reference-ref>"#,
        r#"<t:note-ref t:note-class="margin">bad class</t:note-ref>"#,
        r#"<t:note-ref t:note-class="footnote" t:reference-format="number">bad format</t:note-ref>"#,
        r#"<t:note-ref t:ref-name="n">missing class</t:note-ref>"#,
        r#"<t:bookmark-ref xmlns:x="urn:not-text" x:reference-format="page">spoof</t:bookmark-ref>"#,
        r#"<t:reference-ref xmlns:x="urn:extension" t:ref-name="x" x:extra="1">extra</t:reference-ref>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
    let oversized = OdfDynamicTextField::Reference {
        reference_name: Some("x".repeat(65_537)),
        reference_format: None,
        display_text: String::new(),
    };
    assert!(oversized.validate().is_err());
    let forbidden = OdfDynamicTextField::BookmarkReference {
        reference_name: Some("bad\u{0}name".to_string()),
        reference_format: None,
        display_text: String::new(),
    };
    assert!(forbidden.to_xml_fragment().is_err());
}

#[test]
fn document_statistics_round_trip_all_seven_kinds_and_numbering_modes() {
    let kinds = [
        OdfDocumentStatisticKind::Page,
        OdfDocumentStatisticKind::Paragraph,
        OdfDocumentStatisticKind::Word,
        OdfDocumentStatisticKind::Character,
        OdfDocumentStatisticKind::Table,
        OdfDocumentStatisticKind::Image,
        OdfDocumentStatisticKind::Object,
    ];
    for (index, kind) in kinds.into_iter().enumerate() {
        let number_format = match index % 3 {
            0 => None,
            1 => Some(OdfStatisticNumberFormat::new("I", None).unwrap()),
            _ => Some(OdfStatisticNumberFormat::new("A", Some(true)).unwrap()),
        };
        let field = OdfDynamicTextField::DocumentStatistic {
            kind,
            number_format,
            display_text: format!("{} <cached> & safe", index + 1),
        };
        let fragment = field.to_xml_fragment().unwrap();
        assert!(fragment.contains(kind.element_name()));
        assert_eq!(
            document(&fragment).dynamic_text_fields().unwrap(),
            vec![field]
        );
    }
}

#[test]
fn document_statistics_parse_style_aliases_and_support_mutation() {
    let source = document(
        r#"<t:word-count xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:num-format="a" s:num-letter-sync="1">ten</t:word-count>"#,
    );
    let replacement = OdfDynamicTextField::DocumentStatistic {
        kind: OdfDocumentStatisticKind::Page,
        number_format: Some(OdfStatisticNumberFormat::new("1", None).unwrap()),
        display_text: "12".to_string(),
    };
    let mut mutable = MutableDocument::from_document(source).unwrap();
    assert!(matches!(
        mutable.replace_dynamic_text_field(0, &replacement).unwrap(),
        OdfDynamicTextField::DocumentStatistic {
            kind: OdfDocumentStatisticKind::Word,
            ..
        }
    ));
    let inserted = OdfDynamicTextField::DocumentStatistic {
        kind: OdfDocumentStatisticKind::Image,
        number_format: None,
        display_text: "3".to_string(),
    };
    mutable.insert_dynamic_text_field(0, &inserted).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.dynamic_text_fields().unwrap(),
        vec![replacement.clone(), inserted]
    );
    let mut mutable = MutableDocument::from_document(round_trip).unwrap();
    assert_eq!(mutable.remove_dynamic_text_field(0).unwrap(), replacement);
}

#[test]
fn document_statistics_reject_invalid_numbering_namespaces_attributes_and_bounds() {
    for invalid in [
        r#"<t:page-count xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:num-letter-sync="true">bad</t:page-count>"#,
        r#"<t:paragraph-count xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:num-format="1" s:num-letter-sync="true">bad</t:paragraph-count>"#,
        r#"<t:word-count xmlns:x="urn:not-style" x:num-format="I">spoof</t:word-count>"#,
        r#"<t:object-count xmlns:x="urn:extension" x:extra="1">extra</t:object-count>"#,
    ] {
        assert!(
            document(invalid).dynamic_text_fields().is_err(),
            "accepted {invalid}"
        );
    }
    let oversized = OdfDynamicTextField::DocumentStatistic {
        kind: OdfDocumentStatisticKind::Character,
        number_format: None,
        display_text: "x".repeat(65_537),
    };
    assert!(oversized.validate().is_err());
    let forbidden = OdfDynamicTextField::DocumentStatistic {
        kind: OdfDocumentStatisticKind::Table,
        number_format: None,
        display_text: "1\u{0}".to_string(),
    };
    assert!(forbidden.to_xml_fragment().is_err());
}
