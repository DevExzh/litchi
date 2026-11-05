#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected parse and package failures"
)]

use litchi_odf_common::constants::ODF_TEXT;
use litchi_odf_common::core::PackageWriter;
use litchi_odf_formula::authoring::{self, Display, Variant};
use litchi_odf_formula::{Builder, Kind, Package};

const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mi mathvariant="italic">x</mi><mo>+</mo><mn>1</mn></mrow><annotation encoding="StarMath 5.0">x + 1</annotation></semantics></math>"#;

#[test]
fn validated_formula_package_round_trips_without_losing_xml() {
    let package = Builder::from_mathml(MATHML).build().expect("valid formula");
    let formula = package.formula();
    assert_eq!(
        formula.mimetype(),
        "application/vnd.oasis.opendocument.formula"
    );
    assert!(!formula.is_template());
    assert_eq!(formula.root().kind(), Kind::Math);
    assert_eq!(
        formula.starmath().expect("StarMath").opaque().source(),
        "x + 1"
    );
    assert_eq!(formula.content_xml().expect("content XML"), MATHML);

    let reopened = Package::from_bytes(package.as_bytes()).expect("reopen formula");
    assert_eq!(reopened.formula().root(), formula.root());
    assert_eq!(
        reopened.formula().content_xml().expect("content XML"),
        MATHML
    );
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
    let package = Builder::template(root).build().expect("template");
    assert!(package.formula().is_template());
    assert_eq!(
        package.formula().root().attribute(None, "display"),
        Some("inline")
    );

    let mut replacement = authoring::document_root(authoring::identifier("y"), Display::Block);
    replacement.push_text("");
    let mut edit = package.edit();
    edit.set_root(replacement).expect("replace root");
    let package = edit.commit().expect("commit root").into_package();
    assert!(package.formula().is_template());
    assert_eq!(
        package
            .formula()
            .root()
            .children()
            .next()
            .unwrap()
            .local_name(),
        "mi"
    );
}

#[test]
fn malformed_content_and_wrong_family_are_rejected() {
    assert!(create("<math/>").is_err());
    assert!(create("not XML").is_err());

    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(ODF_TEXT).expect("set text MIME type");
    writer
        .add_file("content.xml", MATHML.as_bytes())
        .expect("add text content");
    let wrong_family = writer.finish_to_bytes().expect("finish text package");
    assert!(Package::from_bytes(&wrong_family).is_err());
}

#[test]
fn layout_character_data_is_rejected_by_the_mathml_schema() {
    let xml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>before<mi>a</mi>after</mrow></math>"#;
    assert!(create(xml).is_err());
}

fn create(mathml: &str) -> Result<Package, litchi_odf_formula::Error> {
    Builder::from_mathml(mathml).build()
}
