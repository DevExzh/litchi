#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected parse and package failures"
)]

use litchi_odf_common::constants::ODF_TEXT;
use litchi_odf_common::core::PackageWriter;
use litchi_odf_formula::authoring::{self, Display, Variant};
use litchi_odf_formula::{Content, Formula, Kind};

const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mi mathvariant="italic">x</mi><mo>+</mo><mn>1</mn></mrow><annotation encoding="StarMath 5.0">x + 1</annotation></semantics></math>"#;

#[test]
fn validated_formula_package_round_trips_without_losing_xml() {
    let formula = Formula::create(MATHML).expect("valid formula");
    assert_eq!(
        formula.mimetype(),
        "application/vnd.oasis.opendocument.formula"
    );
    assert!(!formula.is_template());
    assert_eq!(formula.root().kind(), Kind::Math);
    assert_eq!(formula.starmath_source().as_deref(), Some("x + 1"));
    assert_eq!(formula.content_xml().expect("content XML"), MATHML);

    let reopened = Formula::from_bytes(formula.to_bytes()).expect("reopen formula");
    assert_eq!(reopened.root(), formula.root());
    assert_eq!(reopened.content_xml().expect("content XML"), MATHML);
}

#[test]
fn typed_authoring_builds_a_template_and_edits_atomically() {
    let root = authoring::document_root(
        authoring::row(vec![
            authoring::identifier_with_variant("x", Variant::Italic),
            authoring::operator("+"),
            authoring::number("1"),
        ]),
        Display::Inline,
    );
    let mut formula = authoring::Builder::template(root)
        .build()
        .expect("template");
    assert!(formula.is_template());
    assert_eq!(formula.root().attribute(None, "display"), Some("inline"));

    let mut replacement = authoring::document_root(authoring::identifier("y"), Display::Block);
    replacement.push_text("");
    formula.set_root(replacement).expect("replace root");
    assert_eq!(formula.root().children().next().unwrap().local_name(), "mi");
}

#[test]
fn malformed_content_and_wrong_family_are_rejected() {
    assert!(Formula::create("<math/>").is_err());
    assert!(Formula::create("not XML").is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODF_TEXT).expect("set text MIME type");
    writer
        .add_file("content.xml", MATHML.as_bytes())
        .expect("add text content");
    let wrong_family = writer.finish_to_bytes().expect("finish text package");
    assert!(Formula::from_bytes(wrong_family).is_err());
}

#[test]
fn mixed_content_remains_ordered() {
    let xml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>before<mi>a</mi>after</mrow></math>"#;
    let formula = Formula::create(xml).expect("valid formula");
    let row = formula.root().children().next().expect("row");
    assert!(matches!(&row.content()[0], Content::Text(value) if value == "before"));
    assert!(matches!(row.content()[1], Content::Element(_)));
    assert!(matches!(&row.content()[2], Content::Text(value) if value == "after"));
}
