#![allow(
    clippy::expect_used,
    reason = "focused source-splice regression assertions"
)]

use super::*;

fn model() -> Theme {
    Theme {
        name: "Office".to_owned(),
        colors: Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                palette.with(slot, Color::rgb("4472C4").expect("RGB"))
            }),
        fonts: FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos")),
    }
}

fn encoded(value: &Theme) -> String {
    String::from_utf8(codec::encode_part(&value.name, &value.colors, &value.fonts).expect("encode"))
        .expect("UTF-8")
}

fn retain_schema_fixture(name: &str, bytes: &[u8]) {
    if let Some(directory) = std::env::var_os("LITCHI_THEME_SCHEMA_OUTPUT_DIR") {
        let directory = std::path::Path::new(&directory);
        std::fs::create_dir_all(directory).expect("schema fixture directory");
        std::fs::write(directory.join(name), bytes).expect("schema fixture output");
    }
}

#[test]
fn name_splice_preserves_quote_style_and_attribute_whitespace_values() {
    let before = model();
    let source = encoded(&before).replacen("name=\"Office\"", "name='Office'", 1);
    let mut after = before.clone();
    after.name = "A 'quoted' &\r\n\t theme".to_owned();
    let output =
        rewrite_source(source.as_bytes(), &before, &after, Limits::DEFAULT).expect("splice");
    assert_eq!(codec::read(&output).expect("read back"), after);
    retain_schema_fixture("renamed-theme.xml", &output);
    assert!(
        String::from_utf8(output)
            .expect("UTF-8")
            .contains("name='A &apos;quoted&apos; &amp;&#13;&#10;&#9; theme'")
    );
}

#[test]
fn missing_optional_root_name_can_be_inserted() {
    let mut before = model();
    let source = encoded(&before).replacen(" name=\"Office\"", "", 1);
    before.name.clear();
    let mut after = before.clone();
    after.name = "Named".to_owned();
    let output =
        rewrite_source(source.as_bytes(), &before, &after, Limits::DEFAULT).expect("insert");
    assert_eq!(codec::read(&output).expect("read back"), after);
    retain_schema_fixture("inserted-name-theme.xml", &output);
}

#[test]
fn root_namespace_probe_has_a_finite_declaration_limit() {
    let mut xml = format!("<a:theme xmlns:a=\"{}\"", codec::NAMESPACE);
    for index in 0..255 {
        use std::fmt::Write;
        write!(&mut xml, " xmlns:n{index}=\"urn:test:{index}\"").expect("namespace");
    }
    assert_eq!(
        source_namespace(format!("{xml}/>").as_bytes()).expect("exact limit"),
        Some(codec::NAMESPACE)
    );
    xml.push_str(" xmlns:overflow=\"urn:overflow\"/>");
    assert!(source_namespace(xml.as_bytes()).is_err());
}

#[test]
fn conformance_rejects_a_scheme_from_the_other_drawingml_namespace() {
    let source = encoded(&model());
    let mixed = source.replacen(
        "<a:clrScheme",
        &format!("<a:clrScheme xmlns:a=\"{}\"", codec::STRICT_NAMESPACE),
        1,
    );
    assert!(validate_conformance(mixed.as_bytes(), RELATIONSHIP_TYPE).is_err());
    let strict = source.replacen(codec::NAMESPACE, codec::STRICT_NAMESPACE, 1);
    assert!(validate_conformance(strict.as_bytes(), STRICT_THEME_RELATIONSHIP).is_ok());
    let mixed_strict = strict.replacen(
        "<a:fontScheme",
        &format!("<a:fontScheme xmlns:a=\"{}\"", codec::NAMESPACE),
        1,
    );
    assert!(validate_conformance(mixed_strict.as_bytes(), STRICT_THEME_RELATIONSHIP).is_err());
}

#[test]
fn strict_arbitrary_prefix_preserves_namespace_uri_in_user_values() {
    let before = model();
    let source = encoded(&before)
        .replacen(codec::NAMESPACE, codec::STRICT_NAMESPACE, 1)
        .replace("xmlns:a=", "xmlns:z=")
        .replace("<a:", "<z:")
        .replace("</a:", "</z:");
    let mut after = before.clone();
    after.fonts = FontSet::new(
        codec::NAMESPACE,
        Face::new(codec::NAMESPACE),
        Face::new("Aptos"),
    );
    let output =
        rewrite_source(source.as_bytes(), &before, &after, Limits::DEFAULT).expect("Strict splice");
    assert_eq!(codec::read(&output).expect("read back"), after);
    retain_schema_fixture("strict-edited-theme.xml", &output);
    let text = String::from_utf8(output).expect("UTF-8");
    assert!(text.contains(&format!("xmlns:a=\"{}\"", codec::STRICT_NAMESPACE)));
    assert!(text.contains(&format!("typeface=\"{}\"", codec::NAMESPACE)));
}

#[test]
fn multiple_replacements_have_one_checked_output_size() {
    let before = model();
    let source = encoded(&before);
    let mut after = before.clone();
    after.name = "Longer name".to_owned();
    after.colors = after
        .colors
        .with(Slot::Accent1, Color::rgb("ABCDEF").expect("RGB"));
    after.fonts = FontSet::new(
        "Different fonts",
        Face::new("Cambria"),
        Face::new("Calibri"),
    );
    let output =
        rewrite_source(source.as_bytes(), &before, &after, Limits::DEFAULT).expect("splice");
    assert_eq!(codec::read(&output).expect("read back"), after);
    retain_schema_fixture("multiple-edits-theme.xml", &output);
    let exact = Limits::new(output.len(), 100_000, 128);
    assert_eq!(
        rewrite_source(source.as_bytes(), &before, &after, exact).expect("exact"),
        output
    );
    assert!(matches!(
        rewrite_source(source.as_bytes(), &before, &after, Limits::new(output.len() - 1, 100_000, 128)),
        Err(Error::LimitExceeded { actual, maximum, .. }) if actual == output.len() && maximum == output.len() - 1
    ));
}

#[test]
fn empty_elements_count_toward_the_caller_depth_limit() {
    assert!(matches!(
        bounded_xml_preflight(b"<root><child/></root>", Limits::new(128, 10, 1)),
        Err(Error::LimitExceeded {
            actual: 2,
            maximum: 1,
            ..
        })
    ));
    bounded_xml_preflight(b"<root><child/></root>", Limits::new(128, 10, 2)).expect("exact depth");
}
