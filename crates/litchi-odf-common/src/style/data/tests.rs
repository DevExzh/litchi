use super::*;

const HEAD_12: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" office:version="1.2"><office:styles>"#;
const HEAD_13: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3"><office:styles>"#;
const TAIL: &str = "</office:styles><office:automatic-styles/></office:document>";

fn doc13(body: &str) -> String {
    format!("{HEAD_13}{body}{TAIL}")
}

fn doc12(body: &str) -> String {
    format!("{HEAD_12}{body}{TAIL}")
}

fn test_ok<T>(result: Result<T>) -> T {
    match result {
        Ok(value) => value,
        Err(error) => panic!("test operation failed: {error}"),
    }
}

fn test_some<T>(value: Option<T>) -> T {
    match value {
        Some(found_value) => found_value,
        None => panic!("test fixture did not contain a required value"),
    }
}

fn fixture_fragment<'a>(fixture: &'a str, marker: &str, close: &str) -> &'a str {
    let begin = test_some(fixture.find(marker));
    let end = begin + test_some(fixture[begin..].find(close)) + close.len();
    &fixture[begin..end]
}

#[test]
fn parses_all_seven_containers_and_all_standard_tokens() {
    let xml = doc13(
        r##"
        <number:number-style style:name="n" style:display-name="N" number:language="en" number:country="US" number:script="Latn" number:rfc-language-tag="en-US" number:title="title" style:volatile="true" number:transliteration-format="一" number:transliteration-language="zh" number:transliteration-country="CN" number:transliteration-style="medium">
          <style:text-properties fo:color="#ff0000"/><number:text>[</number:text><number:fill-character> </number:fill-character><number:number number:decimal-replacement="--" number:display-factor="1000" number:decimal-places="2" number:min-decimal-places="1" number:min-integer-digits="1" number:grouping="true"><number:embedded-text number:position="1">x</number:embedded-text></number:number><number:text>]</number:text><style:map style:condition="value()&gt;=0" style:apply-style-name="positive" style:base-cell-address="Sheet1.A1"/>
        </number:number-style>
        <number:number-style style:name="s"><number:scientific-number number:min-exponent-digits="2" number:exponent-interval="3" number:forced-exponent-sign="true" number:decimal-places="4" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="false"/></number:number-style>
        <number:number-style style:name="f"><number:fraction number:min-numerator-digits="1" number:min-denominator-digits="2" number:denominator-value="8" number:max-denominator-value="64" number:min-integer-digits="0" number:grouping="false"/></number:number-style>
        <number:currency-style style:name="c" number:automatic-order="true"><number:currency-symbol number:language="fr" number:country="FR">€</number:currency-symbol><number:text> </number:text><number:number/></number:currency-style>
        <number:percentage-style style:name="p"><number:number/><number:text>%</number:text></number:percentage-style>
        <number:date-style style:name="d" number:automatic-order="false" number:format-source="language"><number:day number:style="long" number:calendar="gregorian"/><number:month number:style="short" number:textual="true" number:possessive-form="false" number:calendar="gengou"/><number:year/><number:era/><number:day-of-week/><number:week-of-year number:calendar="ROC"/><number:quarter/><number:hours/><number:minutes/><number:seconds number:style="long" number:decimal-places="3"/><number:am-pm/></number:date-style>
        <number:time-style style:name="t" number:format-source="fixed" number:truncate-on-overflow="false"><number:hours number:style="long"/><number:text>:</number:text><number:minutes/><number:seconds/><number:am-pm/></number:time-style>
        <number:boolean-style style:name="b"><number:text>?</number:text><number:boolean/><number:text>!</number:text></number:boolean-style>
        <number:text-style style:name="x"><number:text-content/><number:text> </number:text><number:text-content/></number:text-style>
    "##,
    );
    let styles = test_ok(parse_data_styles_xml(&xml, Part::Flat));
    assert_eq!(styles.styles.len(), 9);
    for style in &styles.styles {
        let fragment = test_ok(style.to_xml_fragment(Version::V1_3));
        let reparsed = test_ok(parse_data_styles_xml(&doc13(&fragment), Part::Flat));
        assert_eq!(reparsed.styles[0].kind, style.kind);
        assert_eq!(reparsed.styles[0].parts, style.parts);
    }
}

#[test]
fn reads_libreoffice_12_aliases_but_writes_them_only_as_standard_13() {
    let xml = doc12(
        r#"<number:number-style style:name="n"><loext:fill-character> </loext:fill-character><number:number loext:min-decimal-places="2"/></number:number-style>"#,
    );
    let style = test_ok(parse_data_styles_xml(&xml, Part::Flat))
        .styles
        .remove(0);
    assert!(style.to_xml_fragment(Version::V1_2).is_err());
    let out = test_ok(style.to_xml_fragment(Version::V1_3));
    assert!(out.contains("number:fill-character"));
    assert!(out.contains("number:min-decimal-places=\"2\""));
    assert!(!out.contains("loext:"));
}

#[test]
fn parses_yielddisc_n122_n126_n170() {
    let fixture = include_str!(
        "../../../../../test-data/libreoffice-core/sc/qa/unit/data/functions/financial/fods/yielddisc.fods"
    );
    let body = format!(
        "{}{}{}",
        fixture_fragment(
            fixture,
            r#"<number:currency-style style:name="N122">"#,
            "</number:currency-style>"
        ),
        fixture_fragment(
            fixture,
            r#"<number:text-style style:name="N126">"#,
            "</number:text-style>"
        ),
        fixture_fragment(
            fixture,
            r#"<number:date-style style:name="N170">"#,
            "</number:date-style>"
        )
    );
    let parsed = test_ok(parse_data_styles_xml(&doc12(&body), Part::Flat));
    assert_eq!(parsed.styles.len(), 3);
    assert_eq!(parsed.styles[0].maps.len(), 1);
    assert_eq!(parsed.styles[1].maps.len(), 3);
    assert!(matches!(parsed.styles[2].parts[0], Token::DayOfWeek(_)));
}

#[test]
fn accepts_odfdo_default_style_shapes() {
    let body = r#"<number:boolean-style style:name="bool"><number:boolean/></number:boolean-style><number:currency-style style:name="cur"><number:text>-</number:text><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text> </number:text><number:currency-symbol number:language="fr" number:country="FR">€</number:currency-symbol></number:currency-style><number:date-style style:name="date"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style><number:number-style style:name="num"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style><number:percentage-style style:name="pct"><number:number number:decimal-places="2" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style><number:time-style style:name="time"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>"#;
    assert_eq!(
        test_ok(parse_data_styles_xml(&doc12(body), Part::Flat))
            .styles
            .len(),
        6
    );
}

#[test]
fn rejects_wrong_namespace_order_cardinality_and_lexicals() {
    let invalid = [
        r#"<x:number-style xmlns:x="urn:wrong" style:name="n"/>"#,
        "<number:number-style/>",
        r#"<number:number-style style:name="n"><number:number/><number:currency-symbol>$</number:currency-symbol></number:number-style>"#,
        r#"<number:date-style style:name="d"/>"#,
        r#"<number:time-style style:name="t"><number:day/></number:time-style>"#,
        r#"<number:number-style style:name="n"><style:map style:condition="x" style:apply-style-name="a"/><number:number/></number:number-style>"#,
        r#"<number:number-style style:name="n"><number:number number:grouping="yes"/></number:number-style>"#,
        r#"<number:number-style style:name="n"><number:number number:decimal-places="1.5"/></number:number-style>"#,
        r#"<number:number-style style:name="n"><number:fraction number:max-denominator-value="0"/></number:number-style>"#,
        r#"<number:number-style style:name="n"><number:number/><style:map style:condition="x"/></number:number-style>"#,
        r#"<number:boolean-style style:name="b"><number:fill-character> </number:fill-character><number:boolean/></number:boolean-style>"#,
    ];
    for invalid_body in invalid {
        assert!(
            parse_data_styles_xml(&doc13(invalid_body), Part::Flat).is_err(),
            "accepted {invalid_body}"
        );
    }
    assert!(parse_data_styles_xml(&doc12(r#"<number:number-style style:name="n"><number:number number:min-decimal-places="1"/></number:number-style>"#), Part::Flat).is_err());
}

#[test]
fn accepts_exact_xsd_integer_and_double_lexicals() {
    let body = r#"<number:number-style style:name="plus"><number:number number:decimal-places="+2" number:min-integer-digits="+1" number:display-factor="+1.5"/></number:number-style><number:number-style style:name="inf"><number:number number:display-factor="INF"/></number:number-style><number:number-style style:name="neg"><number:number number:display-factor="-INF"/></number:number-style><number:number-style style:name="nan"><number:number number:display-factor="NaN"/></number:number-style>"#;
    let parsed = test_ok(parse_data_styles_xml(&doc13(body), Part::Flat));
    let factor = |index: usize| {
        let Token::Number(number) = &parsed.styles[index].parts[0] else {
            panic!("expected number token");
        };
        test_some(number.display_factor)
    };
    assert!((factor(0) - 1.5).abs() < f64::EPSILON);
    assert!(factor(1).is_infinite() && factor(1).is_sign_positive());
    assert!(factor(2).is_infinite() && factor(2).is_sign_negative());
    assert!(factor(3).is_nan());
    assert!(
        test_ok(parsed.styles[1].to_xml_fragment(Version::V1_3)).contains("display-factor=\"INF\"")
    );
    assert!(
        test_ok(parsed.styles[2].to_xml_fragment(Version::V1_3))
            .contains("display-factor=\"-INF\"")
    );
    assert!(
        test_ok(parsed.styles[3].to_xml_fragment(Version::V1_3)).contains("display-factor=\"NaN\"")
    );
    for lexical in ["inf", "-inf", "+INF", "1e", "++1"] {
        let invalid_body = format!(
            r#"<number:number-style style:name="bad"><number:number number:display-factor="{lexical}"/></number:number-style>"#
        );
        assert!(parse_data_styles_xml(&doc13(&invalid_body), Part::Flat).is_err());
    }
    assert!(parse_data_styles_xml(&doc13(r#"<number:number-style style:name="bad"><number:number number:decimal-places="++1"/></number:number-style>"#), Part::Flat).is_err());
}

#[test]
fn lossless_insert_replace_remove_preserves_unrelated_markup() {
    let original = doc13(
        "<!--keep--><number:number-style style:name=\"other\"><number:number/></number:number-style><x:keep xmlns:x=\"urn:x\"/>",
    );
    let mut style = test_ok(Style::new("new", Kind::Number, Section::Styles));
    style.parts.push(Token::Number(NumberToken::default()));
    let inserted = test_ok(set_data_style_xml(&original, &style));
    assert!(inserted.contains("<!--keep--><number:number-style style:name=\"other\""));
    assert!(inserted.contains("<x:keep xmlns:x=\"urn:x\"/>"));
    style.parts = vec![Token::Text("-".into())];
    let replaced = test_ok(set_data_style_xml(&inserted, &style));
    assert!(replaced.contains("<number:text>-</number:text>"));
    assert_eq!(
        test_ok(remove_data_style_xml(&replaced, Section::Styles, "new")),
        original
    );
}

#[test]
fn expands_empty_target_container_and_enforces_caps() {
    let xml = format!("{HEAD_13}</office:styles><office:automatic-styles/></office:document>");
    let mut style = test_ok(Style::new("auto", Kind::Text, Section::AutomaticStyles));
    style.parts.push(Token::TextContent);
    let output = test_ok(set_data_style_xml(&xml, &style));
    assert!(output.contains("<office:automatic-styles><number:text-style"));
    let huge = "x".repeat(MAX_VALUE_BYTES + 1);
    style.parts = vec![Token::Text(huge)];
    assert!(style.validate(Version::V1_3).is_err());
}
