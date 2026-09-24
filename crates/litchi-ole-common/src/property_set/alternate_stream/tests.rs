use super::{ALTERNATE_STREAM_CONTROL_NAME, AlternateStreamControl, NonSimpleAlternateStreamName};
use crate::property_set::{
    Binding, DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, PROPERTY_BAG_FMTID,
    SUMMARY_INFORMATION_FMTID, USER_DEFINED_PROPERTIES_FMTID,
};

#[test]
fn property_bag_example_control_packet_is_the_eight_byte_absent_clsid_form() {
    let bytes = [0u8; 8];
    let control = AlternateStreamControl::parse(&bytes).unwrap();

    assert_eq!(
        ALTERNATE_STREAM_CONTROL_NAME,
        "{4c8cc155-6c1e-11d1-8e41-00c04fb9386d}"
    );
    assert_eq!(control.reserved2(), 0);
    assert_eq!(control.application_state(), 0);
    assert_eq!(control.class_identifier(), None);
    assert_eq!(control.to_bytes(), bytes);
}

#[test]
fn control_packet_preserves_reserved2_and_opaque_optional_clsid() {
    let class_identifier = Guid::from_bytes([
        0x53, 0xFF, 0x4B, 0x99, 0xF9, 0xDD, 0xAD, 0x42, 0xA5, 0x6A, 0xFF, 0xEA, 0x36, 0x17, 0xAC,
        0x16,
    ]);
    let mut bytes = Vec::from([0u8; 24]);
    bytes[2..4].copy_from_slice(&0xA55Au16.to_le_bytes());
    bytes[4..8].copy_from_slice(&0x8000_0001u32.to_le_bytes());
    bytes[8..24].copy_from_slice(class_identifier.as_bytes());

    let control = AlternateStreamControl::parse(&bytes).unwrap();
    assert_eq!(control.reserved2(), 0xA55A);
    assert_eq!(control.application_state(), 0x8000_0001);
    assert_eq!(control.class_identifier(), Some(class_identifier));
    assert_eq!(control.to_bytes(), bytes);
}

#[test]
fn control_edits_retain_reserved2_while_changing_state_and_clsid_presence() {
    let mut bytes = [0u8; 8];
    bytes[2..4].copy_from_slice(&0xA55Au16.to_le_bytes());
    bytes[4..8].copy_from_slice(&7u32.to_le_bytes());
    let source = AlternateStreamControl::parse(&bytes).unwrap();
    let class_identifier = Guid::from_bytes([0xCD; 16]);

    let edited = source
        .with_application_state(42)
        .with_class_identifier(Some(class_identifier));
    assert_eq!(edited.reserved2(), 0xA55A);
    assert_eq!(edited.application_state(), 42);
    assert_eq!(edited.class_identifier(), Some(class_identifier));
    assert_eq!(&edited.to_bytes()[..4], [0, 0x00, 0x5A, 0xA5]);

    let cleared = edited.with_class_identifier(None);
    assert_eq!(cleared.reserved2(), 0xA55A);
    assert_eq!(cleared.class_identifier(), None);
    assert_eq!(cleared.to_bytes(), [0, 0, 0x5A, 0xA5, 42, 0, 0, 0]);
}

#[test]
fn fresh_control_packets_zero_reserved_words_and_select_packet_size() {
    let absent = AlternateStreamControl::new(7, None);
    assert_eq!(absent.to_bytes(), [0, 0, 0, 0, 7, 0, 0, 0]);

    let present = AlternateStreamControl::new(7, Some(Guid::from_bytes([0xAB; 16])));
    let bytes = present.to_bytes();
    assert_eq!(bytes.len(), 24);
    assert_eq!(&bytes[..8], [0, 0, 0, 0, 7, 0, 0, 0]);
    assert_eq!(&bytes[8..], [0xAB; 16]);
}

#[test]
fn control_packet_rejects_other_sizes_and_nonzero_reserved1() {
    for length in [0, 1, 7, 9, 23, 25] {
        assert!(AlternateStreamControl::parse(&vec![0; length]).is_err());
    }

    let mut bytes = [0u8; 8];
    bytes[..2].copy_from_slice(&1u16.to_le_bytes());
    assert!(AlternateStreamControl::parse(&bytes).is_err());
}

#[test]
fn non_simple_name_matches_the_property_bag_alternate_stream_example() {
    let name = NonSimpleAlternateStreamName::new(Binding::custom(PROPERTY_BAG_FMTID));
    assert_eq!(name.as_str(), "Docf_\u{0005}bagaaqy23kudbhchaaq5u2chnd");
    assert_eq!(name.binding(), Binding::custom(PROPERTY_BAG_FMTID));
    assert_eq!(
        NonSimpleAlternateStreamName::parse("Docf_\u{0005}Bagaaqy23kudbhchAaq5u2chNd").unwrap(),
        name
    );
}

#[test]
fn non_simple_name_keeps_standard_binding_names_distinct_from_cfb_paths() {
    let name = NonSimpleAlternateStreamName::new(Binding::SummaryInformation);
    assert_eq!(name.as_str(), "Docf_\u{0005}SummaryInformation");
    assert_eq!(name.len(), name.as_bytes().len());
    assert_eq!(name.as_ref(), name.as_str());
    assert_eq!(name.to_string(), "Docf_\u{0005}SummaryInformation");
}

#[test]
fn non_simple_name_canonicalizes_binding_aliases_to_the_wire_name() {
    let document_summary = NonSimpleAlternateStreamName::new(Binding::DocumentSummaryInformation);
    let user_defined = NonSimpleAlternateStreamName::new(Binding::UserDefinedProperties);
    assert_eq!(document_summary, user_defined);
    assert_eq!(user_defined.binding(), Binding::DocumentSummaryInformation);
    assert_eq!(
        NonSimpleAlternateStreamName::parse(user_defined.as_str()).unwrap(),
        document_summary
    );

    let summary = NonSimpleAlternateStreamName::new(Binding::SummaryInformation);
    let custom_special =
        NonSimpleAlternateStreamName::new(Binding::custom(SUMMARY_INFORMATION_FMTID));
    assert_eq!(summary, custom_special);
    assert_eq!(
        NonSimpleAlternateStreamName::parse(custom_special.as_str()).unwrap(),
        summary
    );

    let custom_document_summary =
        NonSimpleAlternateStreamName::new(Binding::custom(DOCUMENT_SUMMARY_INFORMATION_FMTID));
    let custom_user_defined =
        NonSimpleAlternateStreamName::new(Binding::custom(USER_DEFINED_PROPERTIES_FMTID));
    assert_eq!(custom_document_summary, custom_user_defined);
    assert_eq!(
        custom_user_defined.binding(),
        Binding::DocumentSummaryInformation
    );
}

#[test]
fn non_simple_name_rejects_wrong_prefix_and_invalid_binding_suffix() {
    let valid = "Docf_\u{0005}SummaryInformation";
    assert_eq!(
        NonSimpleAlternateStreamName::parse("dOcF_\u{0005}SummaryInformation")
            .unwrap()
            .as_str(),
        valid
    );
    assert!(NonSimpleAlternateStreamName::parse("Docf_SummaryInformation").is_err());
    assert!(NonSimpleAlternateStreamName::parse("Docf_\u{0005}too-short").is_err());
    assert!(NonSimpleAlternateStreamName::parse(valid).is_ok());
}
