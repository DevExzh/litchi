use super::json::{self, Value};
use super::{Data, Error, Missing, Options, Planner, Record, normalize_breaks, placeholders};
use crate::transaction::SemanticLimits;

fn keys(text: &str) -> Vec<(&str, &str)> {
    placeholders(text)
        .map(|placeholder| (&text[placeholder.range], &text[placeholder.key]))
        .collect()
}

#[test]
fn placeholders_follow_officecli_grammar() {
    assert_eq!(keys("{{name}}"), [("{{name}}", "name")]);
    assert_eq!(keys("{{ name }}"), [("{{ name }}", "name")]);
    assert_eq!(
        keys("{{  first name  }}"),
        [("{{  first name  }}", "first name")]
    );
    assert_eq!(keys("{{user-id}}"), [("{{user-id}}", "user-id")]);
    assert_eq!(
        keys("{{items[0].name}}"),
        [("{{items[0].name}}", "items[0].name")]
    );
    assert_eq!(keys("{{\tname\n}}"), [("{{\tname\n}}", "name")]);
    assert_eq!(keys("{{名前}}"), [("{{名前}}", "名前")]);
    assert_eq!(keys("{{_x}}{{9}}"), [("{{_x}}", "_x"), ("{{9}}", "9")]);
    // Not placeholders: an empty key, a key starting with punctuation, a
    // single brace, a missing close, a tab inside the key.
    for text in [
        "{{}}", "{{ }}", "{{-x}}", "{x}", "{{x}", "{{a\tb}}", "{{a}b}}", "{{a+b}}",
    ] {
        assert!(keys(text).is_empty(), "{text}");
    }
    // A failed match moves one character on.
    assert_eq!(keys("{{{name}}}"), [("{{name}}", "name")]);
    assert_eq!(keys("{{ {{a}}"), [("{{a}}", "a")]);
    assert_eq!(keys("a {{b}} c {{d}}"), [("{{b}}", "b"), ("{{d}}", "d")]);
}

#[test]
fn json_keeps_number_spellings_and_member_order() {
    let value =
        json::parse(br#" {"b": 5200.00, "a": [1e3, -0, true, null, "x\u00e9\ud83d\ude00"]} "#)
            .unwrap();
    let Value::Object(members) = &value else {
        panic!("object");
    };
    assert_eq!(members[0].0, "b");
    assert_eq!(members[0].1, Value::Number("5200.00".to_owned()));
    let Value::Array(items) = &members[1].1 else {
        panic!("array");
    };
    assert_eq!(items[0], Value::Number("1e3".to_owned()));
    assert_eq!(items[4], Value::String("xé😀".to_owned()));
    assert_eq!(
        json::render(&members[1].1),
        "[\n  1e3,\n  -0,\n  true,\n  null,\n  \"xé😀\"\n]"
    );
    for bad in [
        &b"{"[..],
        b"[1,]",
        b"01",
        b"1.",
        b"\"\\x\"",
        b"{\"a\" 1}",
        b"nul",
        b"\"\x01\"",
    ] {
        assert!(
            matches!(json::parse(bad), Err(Error::Json { .. })),
            "{bad:?}"
        );
    }
    assert!(matches!(
        json::parse(b"{\n  \"a\": tru\n}"),
        Err(Error::Json { line: 2, column: 8 })
    ));
}

#[test]
fn records_flatten_nested_members_with_literal_keys_winning() {
    let data = Data::from_json(br#"{"a": {"b": "nested", "c": {"d": 1}}, "a.b": "literal", "list": [{"k": "v"}, [true]], "none": null}"#, &SemanticLimits::default())
    .unwrap();
    let record = &data.records()[0];
    assert_eq!(record.get("a.b"), Some("literal"));
    assert_eq!(record.get("a.c.d"), Some("1"));
    assert_eq!(record.get("a.c"), Some("{\n  \"d\": 1\n}"));
    assert_eq!(record.get("list[0].k"), Some("v"));
    assert_eq!(record.get("list[1][0]"), Some("true"));
    assert_eq!(record.get("none"), Some(""));
    assert!(record.get("list[1]").is_none());
    let many = Data::from_json(br#"[{"a": 1}, {"a": 2}]"#, &SemanticLimits::default()).unwrap();
    assert_eq!(many.len(), 2);
    assert!(matches!(
        Data::from_json(br#"[{"a": 1}, 2]"#, &SemanticLimits::default())
            .unwrap_err()
            .leaf(),
        crate::error::Source::Merge(Error::NotARecord { record: 2 })
    ));
}

#[test]
fn csv_records_take_their_keys_from_the_header() {
    let data = Data::from_csv(
        b"\xEF\xBB\xBFname,,name,city\nAda,x,y,London\nGrace\n\n,,,\n",
        &crate::import::Options::csv(),
        &SemanticLimits::default(),
    )
    .unwrap();
    assert_eq!(data.len(), 2);
    let first = &data.records()[0];
    assert_eq!(
        (first.get("name"), first.get("city")),
        (Some("Ada"), Some("London"))
    );
    assert_eq!(first.len(), 2);
    assert_eq!(data.records()[1].get("city"), Some(""));
    let options = crate::import::Options::tsv();
    let tabbed = Data::from_csv(b"a\tb\n1\t2\n", &options, &SemanticLimits::default()).unwrap();
    assert_eq!(tabbed.records()[0].get("b"), Some("2"));
    let limits = SemanticLimits::builder()
        .max_staged_bytes(8)
        .unwrap()
        .build()
        .unwrap();
    assert!(Data::from_json(br#"{"key": "a long value"}"#, &limits).is_err());
}

#[test]
fn the_planner_counts_keys_and_honours_the_missing_policy() {
    let record: Record = [("a", "1")].into_iter().collect();
    let limits = SemanticLimits::default();
    for (missing, expected) in [
        (Missing::Keep, vec!["1", "1"]),
        (Missing::Empty, vec!["1", "", "1"]),
    ] {
        let mut planner = Planner::new(&record, &Options::new().missing(missing), &limits);
        let mut values = Vec::new();
        planner
            .text(
                "{{a}} {{b}} {{a}}",
                |_| true,
                &[],
                |_, value| values.push(value),
            )
            .unwrap();
        assert_eq!(values, expected);
        let summary = planner.finish();
        assert_eq!(summary.keys(), ["a"]);
        assert_eq!(summary.unresolved(), ["b"]);
    }
    // Taken ranges and non-text ranges are skipped before lookup.
    let mut planner = Planner::new(&record, &Options::new(), &limits);
    let mut values = Vec::new();
    planner
        .text(
            "{{a}} {{b}} {{a}}",
            |range| range.start != 12,
            &[3..9, 30..31],
            |_, value| values.push(value),
        )
        .unwrap();
    assert!(values.is_empty());
    assert!(planner.finish().unresolved().is_empty());
}

#[cfg(feature = "docx")]
#[test]
fn merge_fields_apply_their_switches() {
    let record: Record = [("Name", "ada lovelace"), ("Empty", "")]
        .into_iter()
        .collect();
    let limits = SemanticLimits::default();
    let mut planner = Planner::new(&record, &Options::new(), &limits);
    let cases = [
        (r#" MERGEFIELD Name \* Upper "#, Some("ADA LOVELACE")),
        (
            r#" MERGEFIELD name \* Caps \b "Dear " \f "," "#,
            Some("Dear Ada Lovelace,"),
        ),
        (r#" MERGEFIELD Name \* FirstCap "#, Some("Ada lovelace")),
        (r#" MERGEFIELD Empty \b "x" "#, Some("")),
        (r#" MERGEFIELD Missing "#, None),
        (r#" PAGE "#, None),
    ];
    for (instruction, expected) in cases {
        assert_eq!(
            planner.field(instruction, true).unwrap().as_deref(),
            expected,
            "{instruction}"
        );
    }
    assert_eq!(planner.field(" MERGEFIELD Name ", false).unwrap(), None);
    let summary = planner.finish();
    assert_eq!(summary.fields(), 4);
    assert_eq!(summary.skipped_fields(), 1);
    assert_eq!(summary.unresolved(), ["Missing"]);
    assert!(!format!("{summary:?}").contains("ada"));
}

#[test]
fn breaks_are_normalized() {
    assert_eq!(normalize_breaks("a\r\nb\rc\nd"), "a\nb\nc\nd");
    assert!(matches!(
        normalize_breaks("plain"),
        std::borrow::Cow::Borrowed(_)
    ));
}
