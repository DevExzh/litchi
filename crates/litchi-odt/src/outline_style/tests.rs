//! Focused regression tests for the outline-style model and codec.

use super::{
    Attribute, LevelStyle, NumberFormat, PositiveInteger, Style, TextAlign, parse_outline_styles,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

fn document(outline: &str) -> String {
    format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}" xmlns:f="{FO}"><o:styles>{outline}</o:styles><o:body><o:text/></o:body></o:document>"#
    )
}

#[test]
fn typed_model_serializes_and_reparses() {
    let style = Style {
        name: "Outline".to_string(),
        levels: vec![LevelStyle {
            level: 1,
            text_style_name: Some("Heading_20_1".to_string()),
            number_format: Some(NumberFormat::new("A").unwrap()),
            number_prefix: Some("(".to_string()),
            number_suffix: Some(")".to_string()),
            letter_sync: Some(true),
            display_levels: Some(PositiveInteger::new("2").unwrap()),
            start_value: Some(PositiveInteger::new("1").unwrap()),
            list_level_properties: None,
            text_properties: None,
            extensions: vec![Attribute::new("urn:producer", "num-list-format", "%1%").unwrap()],
        }],
        extensions: Vec::new(),
    };
    let xml = style.to_xml().unwrap();
    let parsed = parse_outline_styles(&document(&xml)).unwrap();
    assert_eq!(parsed.get("Outline"), Some(&style));
}

#[test]
fn invalid_typed_model_is_rejected_before_writing() {
    let style = Style {
        name: "Outline".to_string(),
        levels: vec![LevelStyle {
            level: 1,
            text_style_name: None,
            number_format: Some(NumberFormat::new("1").unwrap()),
            number_prefix: None,
            number_suffix: None,
            letter_sync: Some(true),
            display_levels: None,
            start_value: None,
            list_level_properties: None,
            text_properties: None,
            extensions: Vec::new(),
        }],
        extensions: Vec::new(),
    };
    assert!(style.to_xml().is_err());
    assert_eq!(TextAlign::Start.as_str(), "start");
}

/// Every sequence of at most `max_len` indices below `alphabet`, shortest
/// first.
fn sequences(alphabet: usize, max_len: usize) -> Vec<Vec<usize>> {
    let mut all = vec![Vec::new()];
    let mut previous = vec![Vec::new()];
    for _ in 0..max_len {
        let mut next = Vec::new();
        for prefix in &previous {
            for index in 0..alphabet {
                let mut sequence: Vec<usize> = prefix.clone();
                sequence.push(index);
                next.push(sequence);
            }
        }
        all.extend(next.iter().cloned());
        previous = next;
    }
    all
}

fn assert_lines(actual: &[String], expected: &[&str]) {
    let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
    assert!(
        actual
            .iter()
            .map(String::as_str)
            .eq(expected.iter().copied()),
        "observed lines:\n{}",
        listing.join("\n")
    );
}

#[test]
fn outline_level_child_order_matrix_is_stable() {
    let children = [
        ("list", "<s:list-level-properties/>"),
        (
            "list-open",
            "<s:list-level-properties></s:list-level-properties>",
        ),
        ("text", "<s:text-properties/>"),
        ("text-open", "<s:text-properties></s:text-properties>"),
    ];
    let mut cases: Vec<Vec<(&str, &str)>> = sequences(children.len(), 3)
        .into_iter()
        .map(|sequence| sequence.into_iter().map(|index| children[index]).collect())
        .collect();
    let unknown = ("unknown", "<s:unknown/>");
    let foreign = (
        "foreign list",
        "<x:list-level-properties xmlns:x=\"urn:other\"/>",
    );
    let aligned = (
        "list-aligned",
        "<s:list-level-properties><s:list-level-label-alignment t:label-followed-by=\"space\"/></s:list-level-properties>",
    );
    cases.push(vec![unknown]);
    cases.push(vec![children[0], unknown]);
    cases.push(vec![children[2], unknown]);
    cases.push(vec![foreign]);
    cases.push(vec![aligned, children[2]]);
    cases.push(vec![aligned, children[0]]);
    cases.push(vec![children[2], aligned]);
    let actual: Vec<String> = cases
        .iter()
        .map(|case| {
            let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
            let body: String = case.iter().map(|(_, xml)| *xml).collect();
            let xml = document(&format!(
                "<t:outline-style s:name=\"O\"><t:outline-level-style t:level=\"1\">{body}</t:outline-level-style></t:outline-style>"
            ));
            let outcome = match parse_outline_styles(&xml) {
                Ok(styles) => {
                    let level = &styles.get("O").unwrap().levels[0];
                    format!(
                        "ok list={} text={}",
                        level.list_level_properties.is_some(),
                        level.text_properties.is_some()
                    )
                },
                Err(error) => format!("{error:?}"),
            };
            format!("{} => {outcome}", label.join(","))
        })
        .collect();
    assert_lines(
        &actual,
        &[
            " => ok list=false text=false",
            "list => ok list=true text=false",
            "list-open => ok list=true text=false",
            "text => ok list=false text=true",
            "text-open => ok list=false text=true",
            "list,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text => ok list=true text=true",
            "list,text-open => ok list=true text=true",
            "list-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text => ok list=true text=true",
            "list-open,text-open => ok list=true text=true",
            "text,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list,list,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list-open,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,list-open,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list,text,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list,text-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list,text-open,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list,text-open,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list-open,list,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list-open,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,list-open,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list-open,text,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list-open,text-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "list-open,text-open,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "list-open,text-open,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,list,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-open,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-open,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,text,list => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text,list-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text-open,list => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text-open,list-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text-open,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text,text-open,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,list,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list-open,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list-open,list-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list-open,text => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,list-open,text-open => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text-open,text,list => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text,list-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text-open,list => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text-open,list-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text-open,text => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "text-open,text-open,text-open => InvalidFormat(\"duplicate or out-of-order style:text-properties\")",
            "unknown => InvalidFormat(\"text:outline-level-style has an invalid child\")",
            "list,unknown => InvalidFormat(\"text:outline-level-style has an invalid child\")",
            "text,unknown => InvalidFormat(\"text:outline-level-style has an invalid child\")",
            "foreign list => InvalidFormat(\"text:outline-level-style has an invalid child\")",
            "list-aligned,text => ok list=true text=true",
            "list-aligned,list => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
            "text,list-aligned => InvalidFormat(\"duplicate or out-of-order style:list-level-properties\")",
        ],
    );
}
