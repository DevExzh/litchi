#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::presentation_properties::metadata::protection::{Algorithm, Settings, Type};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/modification-verifier/presentation.xml");

#[test]
fn presentation_modification_verifier_is_exposed_by_protection_owner() {
    let xml = std::str::from_utf8(PRESENTATION_XML)
        .unwrap()
        .replace(
            "legacy-hash",
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==",
        )
        .replace("legacy-salt", "AA==");
    let settings = Settings::parse_xml(&xml).unwrap();
    let verifier = settings.modify().unwrap();

    // The standalone protection model retains the security-bearing verifier
    // fields; descriptive legacy provider attributes are intentionally inert.
    assert_eq!(settings.protection_type(), Type::ModifyPassword);
    assert_eq!(verifier.algorithm(), Algorithm::Sha512);
    assert_eq!(
        verifier.hash(),
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="
    );
    assert_eq!(verifier.salt(), "AA==");
    assert_eq!(verifier.spins(), 100_000);
    assert!(settings.to_xml().contains(r#"cryptAlgorithmSid="14""#));
    assert!(settings.to_xml().contains(r#"spinCount="100000""#));
}

#[test]
fn read_only_recommendation_uses_the_p1710_view_mode_owner() {
    let settings = Settings::new().with_read_only_recommended(true);
    let xml = settings.to_pres_props_xml();
    assert!(xml.contains("{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}"));
    assert!(xml.contains("p1710:readonlyRecommended"));
    assert!(xml.contains("http://schemas.microsoft.com/office/powerpoint/2017/10/main"));
    assert!(!xml.contains("discardImageEditData"));
}
