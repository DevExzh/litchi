use litchi_docx::settings::{
    DocumentId, Extension, Extensions, Guid, Settings, WORD_2010_NAMESPACE, WORD_2012_NAMESPACE,
};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = WORD_2010_NAMESPACE;
const W15: &str = WORD_2012_NAMESPACE;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn settings_xml(body: &str) -> String {
    format!(
        r#"<w:settings xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w14 w15">{body}</w:settings>"#
    )
}

#[test]
fn parses_all_typed_extensions_in_order_and_round_trips_opaque_markup() {
    let guid = Guid::parse("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap();
    let xml = settings_xml(
        r#"<w15:chartTrackingRefBased w:val="0"/><w14:docId w14:val="7F00AA10"/><x:future xmlns:x="urn:future"><x:child><![CDATA[a < b]]></x:child></x:future><w14:conflictMode/><w14:discardImageEditingData w14:val="0"/><w14:defaultImageDpi w14:val="-150"/><w15:docId w15:val="{01234567-89AB-CDEF-0123-456789ABCDEF}"/>"#,
    );

    let settings = Settings::parse(xml.as_bytes()).unwrap();
    let extensions = settings.extensions();
    assert_eq!(extensions.len(), 7);
    assert_eq!(extensions.chart_tracking_ref_based(), Some(false));
    assert_eq!(extensions.document_id(), Some(0x7F00AA10));
    assert_eq!(extensions.conflict_mode(), Some(true));
    assert_eq!(extensions.discard_image_editing_data(), Some(false));
    assert_eq!(extensions.default_image_dpi(), Some(-150));
    assert_eq!(extensions.source_document_id(), Some(&guid));
    assert!(extensions.has_source_document_id());
    assert_eq!(extensions.unknown().count(), 1);
    let unknown = extensions.unknown().next().unwrap().xml();
    assert!(unknown.starts_with(b"<x:future"));
    assert!(
        unknown
            .windows(b"xmlns:x=\"urn:future\"".len())
            .any(|window| { window == b"xmlns:x=\"urn:future\"" })
    );

    let fragment = settings.to_xml("w");
    let reparsed = Settings::parse(settings_xml(&fragment).as_bytes()).unwrap();
    assert_eq!(reparsed.extensions(), extensions);
    assert!(fragment.contains("w15:chartTrackingRefBased"));
    assert!(fragment.contains("w14:defaultImageDpi"));
    assert!(fragment.contains("x:future"));
}

#[test]
fn preserves_an_optional_source_document_id_without_a_guid() {
    let settings = Settings::parse(settings_xml(r#"<w15:docId/>"#).as_bytes()).unwrap();
    let extensions = settings.extensions();
    assert!(extensions.has_source_document_id());
    assert_eq!(extensions.source_document_id(), None);
    assert!(matches!(
        extensions.iter().next(),
        Some(Extension::DocumentId(DocumentId::Source(None)))
    ));
    assert!(extensions.to_xml("w").contains("<w15:docId"));
}

#[test]
fn authoring_uses_checked_bounded_values_and_prefix_free_types() {
    let guid = Guid::from_bytes([0xAB; 16]);
    let mut extensions = Extensions::new();
    extensions
        .set_chart_tracking_ref_based(Some(false))
        .unwrap()
        .set_document_id(Some(1))
        .unwrap()
        .set_source_document_id(Some(guid))
        .unwrap()
        .set_conflict_mode(Some(true))
        .unwrap()
        .set_discard_image_editing_data(Some(false))
        .unwrap()
        .set_default_image_dpi(Some(i32::MIN))
        .unwrap();

    assert!(extensions.document_id().is_some());
    assert_eq!(extensions.source_document_id(), Some(&guid));
    assert_eq!(extensions.default_image_dpi(), Some(i32::MIN));
    let fragment = extensions.to_xml("w");
    assert!(fragment.contains("w14:defaultImageDpi"));
    assert!(fragment.contains("w14:val=\"-2147483648\""));
}

#[test]
fn rejects_malformed_values_duplicate_known_extensions_and_nested_typed_markup() {
    let malformed = [
        r#"<w14:conflictMode w14:val="on"/>"#,
        r#"<w14:docId w14:val="00000000"/>"#,
        r#"<w14:docId w14:val="80000000"/>"#,
        r#"<w14:docId w14:val="1234"/>"#,
        r#"<w15:docId w15:val="{01234567-89ab-CDEF-0123-456789ABCDEF}"/>"#,
        r#"<w14:defaultImageDpi/>"#,
        r#"<w14:defaultImageDpi w14:val="2147483648"/>"#,
        r#"<w14:conflictMode><w14:child/></w14:conflictMode>"#,
        r#"<w14:conflictMode/><w14:conflictMode/>"#,
        r#"<x:future xmlns:x="urn:future"><x:child></x:future>"#,
    ];
    for body in malformed {
        assert!(
            Settings::parse(settings_xml(body).as_bytes()).is_err(),
            "malformed settings extension was accepted: {body}"
        );
    }
}
