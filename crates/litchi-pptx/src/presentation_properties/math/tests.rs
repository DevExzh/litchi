#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

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
        let limits = crate::presentation_properties::Limits::default();
        let mut xml = super::super::codec::XmlSink::output(0, limits.max_output_bytes()).unwrap();
        super::codec::write(&mut xml, &value, strict, limits).unwrap();
        let xml = xml.finish().unwrap();
        assert_eq!(super::codec::parse(&xml, limits).unwrap(), value);
        assert!(String::from_utf8_lossy(&xml).contains(if strict {
            STRICT_OMML_NS
        } else {
            OMML_NS
        }));
    }
}

#[test]
fn empty_math_properties_use_schema_defaults() {
    let xml = format!(r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}"><m:mathPr/></a14:m>"#);
    assert_eq!(
        super::codec::parse(
            xml.as_bytes(),
            crate::presentation_properties::Limits::default(),
        )
        .unwrap(),
        Properties::new()
    );
}

#[test]
fn replay_validator_uses_recipient_presentation_profile() {
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}"><p:extLst><p:ext uri="urn:future"/></p:extLst></p:presentationPr>"#
    );
    let name = litchi_opc::PackURI::new("/ppt/presProps.xml").unwrap();
    let part = litchi_opc::part::BlobPart::new(
        name,
        crate::presentation_properties::CONTENT_TYPE.to_owned(),
        xml.into_bytes(),
    );
    assert!(
        crate::presentation_properties::validate_replay_part(
            &part,
            crate::presentation_properties::Limits::default(),
            litchi_opc::SourceLimits::default(),
        )
        .is_ok()
    );
    let tight =
        crate::presentation_properties::Limits::new(1024, 1024, 128, 1, 1024, 1, 1).unwrap();
    assert!(
        crate::presentation_properties::validate_replay_part(
            &part,
            tight,
            litchi_opc::SourceLimits::default(),
        )
        .is_err(),
        "the recipient's tighter node profile must be enforced during replay"
    );

    let event_limited = litchi_opc::SourceLimits::builder()
        .max_events(4)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        crate::presentation_properties::validate_replay_part(
            &part,
            crate::presentation_properties::Limits::default(),
            event_limited,
        ),
        Err(crate::Error::Opc(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::Events,
            ..
        }))
    ));
}

#[test]
fn strict_math_payload_accepts_strict_omml_closing_names() {
    let xml = format!(
        r#"<a14:m xmlns:a14="{A14_NS}" xmlns:m="{STRICT_OMML_NS}"><m:mathPr><m:brkBin m:val="after"/></m:mathPr></a14:m>"#
    );
    assert_eq!(
        super::codec::parse(
            xml.as_bytes(),
            crate::presentation_properties::Limits::default(),
        )
        .unwrap(),
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
            super::codec::parse(
                xml.as_bytes(),
                crate::presentation_properties::Limits::default(),
            )
            .is_err(),
            "accepted {xml}"
        );
    }
}

#[test]
fn parent_facade_preserves_opaque_extensions_around_typed_math() {
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:a14="{A14_NS}" xmlns:m="{OMML_NS}" xmlns:r="{R_NS}" xmlns:v="urn:vendor"><p:extLst><p:ext uri="{MATH_URI}"><a14:m><m:mathPr><m:brkBin m:val="repeat"/></m:mathPr></a14:m></p:ext><p:ext uri="urn:vendor:opaque"><v:payload r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p:presentationPr>"#
    );
    let limits = crate::presentation_properties::Limits::default();
    let value = crate::presentation_properties::Properties::parse(xml.as_bytes(), limits).unwrap();
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

    let written = value.to_xml(false, limits).unwrap();
    let again = crate::presentation_properties::Properties::parse(&written, limits).unwrap();
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
