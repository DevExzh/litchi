use super::{Binding, GLOBAL_INFO_FMTID, IMAGE_CONTENTS_FMTID, IMAGE_INFO_FMTID};
use crate::property_set::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, SUMMARY_INFORMATION_FMTID,
    USER_DEFINED_PROPERTIES_FMTID,
};

#[test]
fn named_bindings_use_the_ms_oleps_cfb_names() {
    let cases = [
        (
            Binding::SummaryInformation,
            SUMMARY_INFORMATION_FMTID,
            "\u{0005}SummaryInformation",
        ),
        (
            Binding::DocumentSummaryInformation,
            DOCUMENT_SUMMARY_INFORMATION_FMTID,
            "\u{0005}DocumentSummaryInformation",
        ),
        (
            Binding::UserDefinedProperties,
            USER_DEFINED_PROPERTIES_FMTID,
            "\u{0005}DocumentSummaryInformation",
        ),
        (Binding::GlobalInfo, GLOBAL_INFO_FMTID, "\u{0005}GlobalInfo"),
        (
            Binding::ImageContents,
            IMAGE_CONTENTS_FMTID,
            "\u{0005}ImageContents",
        ),
        (Binding::ImageInfo, IMAGE_INFO_FMTID, "\u{0005}ImageInfo"),
    ];

    for (binding, format_identifier, expected_name) in cases {
        let name = binding.name();
        assert_eq!(binding.format_identifier(), format_identifier);
        assert_eq!(name.as_str(), expected_name);
        let parsed = parse(name.as_str());
        assert_eq!(
            parsed,
            if binding == Binding::UserDefinedProperties {
                Binding::DocumentSummaryInformation
            } else {
                binding
            }
        );
    }
    assert_eq!(
        parse("\u{0005}documentsummaryinformation"),
        Binding::DocumentSummaryInformation
    );
    assert!(Binding::UserDefinedProperties.uses_document_summary_stream());
}

#[test]
fn generic_binding_names_round_trip_without_heap_storage() {
    let format_identifier = Guid::from_bytes([
        0x10, 0x32, 0x54, 0x76, 0x98, 0xBA, 0xDC, 0xFE, 0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD,
        0xEF,
    ]);
    let binding = Binding::from_format_identifier(format_identifier);
    let name = binding.name();

    assert_eq!(name.len(), 27);
    assert_eq!(name.as_bytes()[0], 0x05);
    assert!(
        name.as_bytes()[1..]
            .iter()
            .all(|byte| byte.is_ascii_lowercase() || *byte >= b'0' && *byte <= b'5')
    );
    assert_eq!(parse(name.as_str()), binding);

    let uppercase = name.as_str().to_ascii_uppercase();
    assert_eq!(parse(&uppercase), Binding::Custom(format_identifier));
    assert_eq!(binding.name(), name);
}

#[test]
fn generic_binding_names_reject_bad_alphabet_and_trailing_bits() {
    assert!(Binding::from_name("SummaryInformation").is_err());
    assert!(Binding::from_name("\u{0005}too-short").is_err());
    assert!(Binding::from_name(&format!("\u{0005}{}", "a".repeat(25) + "6")).is_err());

    let mut bytes = Binding::custom(Guid::from_bytes([0; 16]))
        .name()
        .as_bytes()
        .to_vec();
    bytes[26] = b'b';
    assert!(Binding::from_name(std::str::from_utf8(&bytes).unwrap()).is_err());
}

fn parse(name: &str) -> Binding {
    match Binding::from_name(name) {
        Ok(binding) => binding,
        Err(error) => panic!("{error}"),
    }
}
