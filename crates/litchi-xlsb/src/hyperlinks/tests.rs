#![allow(
    clippy::expect_used,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

use super::{Error, Hyperlink, PREFIX_LEN};

#[test]
fn builders_preserve_hyperlink_semantics() {
    let link = Hyperlink::new_external(0, 1, 2, 3, "https://example.test".to_string())
        .with_location("#section".to_string())
        .with_tooltip("Open".to_string())
        .with_display("Example".to_string());

    assert_eq!(link.row_last, 1);
    assert_eq!(link.col_first, 2);
    assert_eq!(link.target.as_deref(), Some("https://example.test"));
    assert_eq!(link.location.as_deref(), Some("#section"));
}

#[test]
fn serialize_parse_round_trip() {
    let link = Hyperlink::new(0, 4, 1, 3, "rId7".to_string())
        .with_location("Sheet2!A1".to_string())
        .with_tooltip("Go".to_string())
        .with_display("Display".to_string());

    let encoded = link.try_serialize().expect("valid hyperlink strings");
    let parsed = Hyperlink::parse(&encoded).expect("valid hyperlink payload");
    assert_eq!(
        parsed,
        Hyperlink {
            target: None,
            ..link
        }
    );
}

#[test]
fn missing_optional_strings_are_accepted_for_compatibility() {
    let mut encoded = Vec::new();
    encoded.extend_from_slice(&0_u32.to_le_bytes());
    encoded.extend_from_slice(&1_u32.to_le_bytes());
    encoded.extend_from_slice(&2_u32.to_le_bytes());
    encoded.extend_from_slice(&3_u32.to_le_bytes());
    encoded.extend_from_slice(&0_u32.to_le_bytes());

    let parsed = Hyperlink::parse(&encoded).expect("relationship ID is required");
    assert_eq!(parsed.r_id, "");
    assert!(parsed.location.is_none());
}

#[test]
fn rejects_truncated_prefix_and_string() {
    assert!(matches!(
        Hyperlink::parse(&[0; PREFIX_LEN - 1]),
        Err(Error::InvalidLength { .. })
    ));

    let mut encoded = vec![0; PREFIX_LEN];
    encoded.extend_from_slice(&2_u32.to_le_bytes());
    encoded.extend_from_slice(&[0, 0]);
    assert!(matches!(Hyperlink::parse(&encoded), Err(Error::Wire(_))));
}
