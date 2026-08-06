use super::{BinaryBreak, BinarySubtractionBreak, Properties};
use crate::presentation_properties::{Extension, OpaqueExtension};

const A14_NS: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const OMML_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const STRICT_OMML_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/math";
const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MATH_URI: &str = "{4599F94E-CEE6-441E-89CC-EB005ECD8F06}";

#[test]
fn typed_payload_round_trips_in_transitional_and_strict_omml() {
    let value = Properties::new()
        .with_binary_break(BinaryBreak::Repeat)
        .with_binary_subtraction_break(BinarySubtractionBreak::MinusPlus);

    for strict in [false, true] {
        let mut xml = String::new();
        super::codec::write(&mut xml, &value, strict).unwrap();
        assert_eq!(super::codec::parse(xml.as_bytes()).unwrap(), value);
        assert!(xml.contains(if strict { STRICT_OMML_NS } else { OMML_NS }));
    }
}

#[test]
fn empty_math_properties_use_schema_defaults() {
    let xml = format!(r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr/></a14:m>"#);
    assert_eq!(
        super::codec::parse(xml.as_bytes()).unwrap(),
        Properties::new()
    );
}

#[test]
fn strict_math_payload_accepts_strict_omml_closing_names() {
    let xml = format!(
        r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{STRICT_OMML_NS}"><m:mathPr><m:brkBin m:val="after"/></m:mathPr></a14:m>"#
    );
    assert_eq!(
        super::codec::parse(xml.as_bytes()).unwrap(),
        Properties::new().with_binary_break(BinaryBreak::After)
    );
}

#[test]
fn math_payload_rejects_ambiguous_or_unsafe_grammar() {
    let cases = [
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:brkBin val="before"/></m:mathPr></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:brkBin m:val="before"/><m:brkBin m:val="after"/></m:mathPr></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:brkBinSub m:val="+-"/><m:brkBin m:val="before"/></m:mathPr></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:brkBin m:val="invalid"/></m:mathPr></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:unknown/></m:mathPr></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}" unexpected="1"><m:mathPr/></a14:m>"#
        ),
        format!(
            r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr><m:brkBin m:val="before">x</m:brkBin></m:mathPr></a14:m>"#
        ),
    ];

    for xml in cases {
        assert!(
            super::codec::parse(xml.as_bytes()).is_err(),
            "accepted {xml}"
        );
    }
}

#[test]
fn parent_facade_preserves_opaque_extensions_around_typed_math() {
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}" xmlns:r="{R_NS}" xmlns:v="urn:vendor"><p:extLst><p:ext uri="{MATH_URI}"><a14:m><m:mathPr><m:brkBin m:val="repeat"/></m:mathPr></a14:m></p:ext><p:ext uri="urn:vendor:opaque"><v:payload r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p:presentationPr>"#
    );
    let value = crate::presentation_properties::Properties::parse(xml.as_bytes()).unwrap();
    assert_eq!(
        value.math(),
        Some(&Properties::new().with_binary_break(BinaryBreak::Repeat))
    );
    assert!(matches!(
        &value.extensions[1],
        Extension::Unknown(OpaqueExtension { uri, xml })
            if uri == "urn:vendor:opaque"
                && String::from_utf8_lossy(xml).contains("r:id=\"rIdNeverFetched\"")
    ));

    let written = value.to_xml(false).unwrap();
    let again = crate::presentation_properties::Properties::parse(&written).unwrap();
    assert_eq!(again.math(), value.math());
    assert!(String::from_utf8_lossy(&written).contains("https://example.invalid/not-opened"));
}

#[test]
fn package_facade_edits_math_transactionally() {
    let mut package = crate::Package::new().unwrap();
    let value = Properties::new()
        .with_binary_break(BinaryBreak::Before)
        .with_binary_subtraction_break(BinarySubtractionBreak::PlusMinus);

    assert_eq!(package.math_properties().unwrap(), None);
    assert_eq!(package.put_math_properties(value.clone()).unwrap(), None);
    assert_eq!(package.math_properties().unwrap(), Some(value.clone()));

    let bytes = package.to_bytes().unwrap();
    let mut reopened = crate::Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.math_properties().unwrap(), Some(value.clone()));
    assert_eq!(
        reopened.put_math_properties(value.clone()).unwrap(),
        Some(value.clone())
    );
    assert_eq!(reopened.remove_math_properties().unwrap(), Some(value));
    assert_eq!(reopened.math_properties().unwrap(), None);
}
