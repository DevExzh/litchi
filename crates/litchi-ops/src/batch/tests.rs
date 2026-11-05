use std::borrow::Cow;

use super::*;
use crate::PropValue;

#[test]
fn every_officecli_key_survives_transport_and_owned_detachment() {
    let input = r#"{
        "op":"raw-set", "path":"/body/p[1]", "path2":"/body/p[2]",
        "parent":"/body", "type":"paragraph", "from":"/body/p[3]",
        "to":"/body/p[4]", "index":0, "after":"/body/p[5]",
        "before":"/body/p[6]", "mode":"text", "depth":-1,
        "part":"word/document.xml", "xpath":"//w:p", "action":"replace",
        "xml":"<w:p/>", "props":{"text":"Hello", "bold":true},
        "selector":"/body/p:contains(Hello)", "text":"Hello", "dumpVersion":2
    }"#;
    let item: BatchItem<'_> = serde_json::from_str(input).unwrap();
    let wire = serde_json::to_value(&item).unwrap();
    let mut expected: serde_json::Value = serde_json::from_str(input).unwrap();
    let object = expected.as_object_mut().unwrap();
    let command = object.remove("op").unwrap();
    object.insert("command".to_owned(), command);
    assert_eq!(wire, expected);
    assert_eq!(serde_json::to_value(item.into_owned()).unwrap(), expected);
}

#[test]
fn unescaped_values_borrow_and_escaped_values_detach() {
    let input = String::from(
        r#"{"command":"set","path":"/body/p[1]","text":"a\nb","props":{"bold":true,"text":"borrow me"}}"#,
    );
    let item: BatchItem<'_> = serde_json::from_str(&input).unwrap();
    assert!(matches!(item.command, Cow::Borrowed("set")));
    assert!(matches!(item.path, Some(Cow::Borrowed("/body/p[1]"))));
    assert!(matches!(item.text, Some(Cow::Owned(_))));
    assert!(matches!(
        item.props.as_ref().unwrap().get("text"),
        Some(PropValue::Text(Cow::Borrowed("borrow me")))
    ));
    let owned = item.into_owned();
    drop(input);
    assert_eq!(owned.text.as_deref(), Some("a\nb"));
    assert_eq!(
        owned.props.unwrap().get("bold"),
        Some(&PropValue::Bool(true))
    );
}

#[test]
fn property_inputs_keep_types_order_and_first_equals_split() {
    let item: BatchItem<'_> =
        serde_json::from_str(r#"{"command":"set","props":["text=a=b","bold=true","bold=false"]}"#)
            .unwrap();
    let props = item.props.unwrap();
    assert_eq!(props.len(), 3);
    assert_eq!(props.get("text"), Some(&PropValue::from("a=b")));
    assert_eq!(props.get("bold"), Some(&PropValue::from("false")));
    let item: BatchItem<'_> = serde_json::from_str(r#"{"command":"set","props":{"a":null,"b":true,"c":-2,"d":18446744073709551615,"e":1.5,"f":"2"}}"#).unwrap();
    let props = item.props.unwrap();
    assert_eq!(props.get("a"), Some(&PropValue::Null));
    assert_eq!(props.get("b"), Some(&PropValue::Bool(true)));
    assert_eq!(props.get("c"), Some(&PropValue::Int(-2)));
    assert_eq!(props.get("d"), Some(&PropValue::Unsigned(u64::MAX)));
    assert_eq!(props.get("e"), Some(&PropValue::Number(1.5)));
    assert_eq!(props.get("f"), Some(&PropValue::from("2")));
}

/// JSON numbers decode correctly rounded and the canonical wire form is a
/// fixed point (record 0338 F1: the workspace's `serde_json` has
/// `float_roundtrip`).
#[test]
fn numbers_decode_correctly_rounded_and_round_trip_exactly() {
    // Correctly rounded values, checked against Python's `float()`.
    let cases = [
        ("98446744073709551615", 9.844_674_407_370_955e19),
        ("0.1", 0.1),
        ("2.2250738585072011e-308", 2.225_073_858_507_201e-308),
        ("9007199254740993.0", 9_007_199_254_740_992.0),
        ("1.7976931348623157e308", f64::MAX),
        ("3.0000000000000004", 3.000_000_000_000_000_4),
    ];
    for (text, expected) in cases {
        let json = format!(r#"[{{"command":"set","props":{{"n":{text}}}}}]"#);
        let items = from_bytes(json.as_bytes(), &SourceLimits::default()).unwrap();
        let Some(PropValue::Number(number)) = items[0].props.as_ref().unwrap().get("n") else {
            panic!("{text}: not a number");
        };
        assert_eq!(number.to_bits(), expected.to_bits(), "{text}");
        // Serializing and decoding again gives the same double, twice.
        let wire = serde_json::to_vec(&items).unwrap();
        let again = from_bytes(&wire, &SourceLimits::default()).unwrap();
        assert!(again == items, "{text}: round trip moved the number");
        let rewire = serde_json::to_vec(&again).unwrap();
        assert_eq!(rewire, wire, "{text}: the wire form is not a fixed point");
    }
    // The 121-digit literal of record 0334 (finding F1): stable from the
    // first round trip.
    let long = format!("9{}", "7".repeat(120));
    let json = format!(r#"[{{"command":"set","props":{{"n":{long}}}}}]"#);
    let items = from_bytes(json.as_bytes(), &SourceLimits::default()).unwrap();
    let wire = serde_json::to_vec(&items).unwrap();
    assert!(from_bytes(&wire, &SourceLimits::default()).unwrap() == items);
}

#[test]
fn keys_are_ascii_caseless_and_serialization_is_canonical() {
    let item: BatchItem<'_> =
        serde_json::from_str(r#"{"OP":"SeT","PaTh":"/body","TYPE":"p","DUMPVERSION":2}"#).unwrap();
    assert_eq!(
        serde_json::to_string(&item).unwrap(),
        r#"{"command":"SeT","path":"/body","type":"p","dumpVersion":2}"#
    );
}

#[test]
fn optional_nulls_are_absent_but_property_nulls_are_not() {
    let item: BatchItem<'_> = serde_json::from_str(r#"{"command":"set","path":null,"props":null,"index":null,"depth":null,"dumpVersion":null}"#).unwrap();
    assert_eq!(
        serde_json::to_string(&item).unwrap(),
        r#"{"command":"set"}"#
    );
    let item: BatchItem<'_> =
        serde_json::from_str(r#"{"command":"set","props":{"bold":null}}"#).unwrap();
    assert_eq!(item.props.unwrap().get("bold"), Some(&PropValue::Null));
}

#[test]
fn duplicate_fields_never_choose_an_order_dependent_winner() {
    for input in [
        r#"{"command":"set","op":"remove"}"#,
        r#"{"op":"remove","COMMAND":"set"}"#,
        r#"{"command":"set","path":"a","PATH":"b"}"#,
        r#"{"command":"set","index":null,"Index":0}"#,
        r#"{"command":"set","props":null,"props":{}}"#,
        r#"{"command":"set","dumpVersion":2,"DUMPVERSION":1}"#,
    ] {
        assert!(
            serde_json::from_str::<BatchItem<'_>>(input).is_err(),
            "{input}"
        );
    }
}

#[test]
fn unknown_fields_are_refused_even_when_null_or_empty() {
    for value in ["null", "{}", "[]", "true", "42", r#""authored""#] {
        let input = format!(r#"{{"command":"set","propps":{value}}}"#);
        assert!(serde_json::from_str::<BatchItem<'_>>(&input).is_err());
    }
}

#[test]
fn envelope_requires_a_nonblank_string_command() {
    for input in [
        "null",
        "[]",
        "42",
        "{}",
        r#"{"command":null}"#,
        r#"{"command":false}"#,
        r#"{"command":""}"#,
        r#"{"command":"  \n\t"}"#,
    ] {
        assert!(
            serde_json::from_str::<BatchItem<'_>>(input).is_err(),
            "{input}"
        );
    }
    let empty = BatchItem::default();
    assert_eq!(empty.validate().unwrap_err().code(), Code::CorruptBatch);
    assert!(serde_json::to_string(&empty).is_err());
}

#[test]
fn malformed_fields_and_nested_properties_are_refused() {
    for (field, value) in [
        ("index", "2147483648"),
        ("index", "-2147483649"),
        ("depth", "1.5"),
        ("dumpVersion", r#""2""#),
        ("path", "4"),
        ("xml", "[]"),
        ("props", "true"),
        ("props", r#"{"text":{}}"#),
        // An array value is a list of strings (record 0354); anything else
        // inside it is refused.
        ("props", r#"{"text":[1]}"#),
        ("props", r#"{"text":["a",[]]}"#),
        ("props", r#"["missing-equals"]"#),
        ("props", r#"["=missing-key"]"#),
        ("props", "[42]"),
    ] {
        let input = format!(r#"{{"command":"set","{field}":{value}}}"#);
        assert!(
            serde_json::from_str::<BatchItem<'_>>(&input).is_err(),
            "{input}"
        );
    }
}

#[test]
fn an_array_of_strings_is_a_list_value() {
    let item: BatchItem<'_> = serde_json::from_str(
        r#"{"command":"set","path":"/","props":{"keywords":["a","b"],"none":[]}}"#,
    )
    .expect("decode");
    let props = item.props.expect("props");
    assert_eq!(
        props.get("keywords"),
        Some(&PropValue::from(vec!["a", "b"]))
    );
    assert_eq!(
        props.get("none"),
        Some(&PropValue::from(Vec::<String>::new()))
    );
}

#[test]
fn debug_and_decode_errors_never_include_authored_values() {
    let marker = "PRIVATE_PAYLOAD";
    let item = BatchItem::new(marker);
    assert!(!format!("{item:?}").contains(marker));
    for input in [
        format!(r#""{marker}""#),
        format!(r#"{{"command":"set","{marker}":true}}"#),
        format!(r#"{{"command":"set","index":"{marker}"}}"#),
        format!(r#"{{"command":"set","props":["{marker}"]}}"#),
        format!(r#"{{"command":"set","text":"{marker}\z"}}"#),
    ] {
        let error = serde_json::from_str::<BatchItem<'_>>(&input).unwrap_err();
        assert!(!error.to_string().contains(marker));
        assert!(!format!("{error:?}").contains(marker));
    }
}

#[test]
fn dump_version_and_unknown_commands_are_preserved_without_execution_or_rewriting() {
    for version in [i32::MIN, 1, 2, i32::MAX] {
        let mut item = BatchItem::new("future-command");
        item.dump_version = Some(version);
        item.text = Some("paragraph\nsoft\u{000b}line".into());
        item.depth = Some(-1);
        let wire = serde_json::to_string(&item).unwrap();
        let again: BatchItem<'_> = serde_json::from_str(&wire).unwrap();
        assert_eq!(again, item);
    }
}

#[test]
fn heterogeneous_batch_array_decodes_items_in_order() {
    let input = r#"[{"command":"meta","dumpVersion":2},{"command":"add","parent":"/body","type":"p","props":["text=first"]},{"op":"set","path":"/body/p[last()]","props":{"bold":true}}]"#;
    let items: Vec<BatchItem<'_>> = serde_json::from_str(input).unwrap();
    assert_eq!(
        items
            .iter()
            .map(|item| item.command.as_ref())
            .collect::<Vec<_>>(),
        ["meta", "add", "set"]
    );
    assert_eq!(items[2].path.as_deref(), Some("/body/p[last()]"));
}

#[test]
fn array_props_roundtrip_preserves_repeated_ordered_assignments() {
    let input = r#"{"command":"set","props":["bold=true","text=a=b","bold=false"]}"#;
    let item: BatchItem<'_> = serde_json::from_str(input).unwrap();
    let wire = serde_json::to_string(&item).unwrap();
    assert_eq!(
        wire,
        r#"{"command":"set","props":{"bold":"true","text":"a=b","bold":"false"}}"#
    );
    let replay: BatchItem<'_> = serde_json::from_str(&wire).unwrap();
    assert_eq!(replay, item);
}

#[test]
fn bounded_decoder_enforces_bytes_before_parsing_and_items_before_decoding() {
    let limits = SourceLimits::builder()
        .max_input_bytes(2)
        .unwrap()
        .build()
        .unwrap();
    let error = from_bytes(b"bad", &limits).unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.item(), None);
    assert!(matches!(
        error.leaf(),
        crate::error::Source::Batch(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: 3,
            maximum: 2,
        })
    ));
    let limits = SourceLimits::builder()
        .max_items(1)
        .unwrap()
        .build()
        .unwrap();
    let error = from_bytes(br#"[{"command":"set"},{"PRIVATE":"bad"}]"#, &limits).unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.item(), Some(1));
    assert!(matches!(
        error.leaf(),
        crate::error::Source::Batch(Error::Limit {
            kind: SourceLimitKind::Items,
            observed: 2,
            maximum: 1,
        })
    ));
    assert_eq!(
        from_bytes(br#"[{"command":"set"}]"#, &limits)
            .unwrap()
            .len(),
        1
    );
    assert!(from_bytes(b"[]", &limits).unwrap().is_empty());
}

#[test]
fn typed_ingress_attributes_item_errors_and_refuses_trailing_data() {
    for (input, index) in [
        (r#"[{"command":"set"},{"command":""}]"#, Some(1)),
        (r#"[{"command":"set"},false]"#, Some(1)),
        (r#"[{"command":"set","index":"PRIVATE"}]"#, Some(0)),
        (r#"{"command":"set"}"#, None),
        (r#"[] {"PRIVATE":"extra"}"#, None),
        (r#"["PRIVATE"]"#, Some(0)),
    ] {
        let error = from_bytes(input.as_bytes(), &SourceLimits::default()).unwrap_err();
        assert_eq!(error.code(), Code::CorruptBatch);
        assert_eq!(error.item(), index, "{input}");
        assert!(!error.to_string().contains("PRIVATE"));
        assert!(!format!("{error:?}").contains("PRIVATE"));
        let source = std::error::Error::source(&error).unwrap();
        assert!(source.downcast_ref::<Error>().is_some());
        assert!(!source.to_string().contains("PRIVATE"));
    }
}

#[test]
fn ingress_limits_are_checked_and_property_payloads_count_toward_bytes() {
    for value in [0, usize::MAX] {
        assert_eq!(
            SourceLimits::builder()
                .max_input_bytes(value)
                .unwrap_err()
                .code(),
            Code::InvalidArgument
        );
        assert_eq!(
            SourceLimits::builder().max_items(value).unwrap_err().code(),
            Code::InvalidArgument
        );
    }
    let input = br#"[{"command":"set","props":{"text":"some content"}}]"#;
    let exact = SourceLimits::builder()
        .max_input_bytes(input.len())
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(from_bytes(input, &exact).unwrap().len(), 1);
    let short = SourceLimits::builder()
        .max_input_bytes(input.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(
        from_bytes(input, &short).unwrap_err().code(),
        Code::LimitExceeded
    );
    assert!(from_bytes(b"[\xff]", &SourceLimits::default()).is_err());
}
