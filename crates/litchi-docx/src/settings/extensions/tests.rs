use super::*;
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = WORD_2010_NAMESPACE;
const W15: &str = WORD_2012_NAMESPACE;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn package_processing_understands_ignorable_word_settings_extensions() {
    let xml = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w14 w15"><w14:conflictMode/><w15:chartTrackingRefBased/></w:settings>"#
    );
    let part = BlobPart::new(
        PackURI::new("/word/settings.xml").unwrap(),
        ct::WML_SETTINGS.to_owned(),
        xml.into_bytes(),
    );
    let processed = process_part(&part).unwrap();
    let processed = std::str::from_utf8(processed.as_ref()).unwrap();
    assert!(processed.contains("conflictMode"));
    assert!(processed.contains("chartTrackingRefBased"));
}

#[test]
fn serialized_extension_siblings_each_bind_their_own_namespaces() {
    use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
    let mut values = Extensions::new();
    values
        .set_chart_tracking_ref_based(Some(OnOff::off()))
        .unwrap()
        .set_source_document_id(Some(Guid::from_bytes([0xab; 16])))
        .unwrap()
        .set_document_id(Some(1))
        .unwrap()
        .set_conflict_mode(Some(OnOff::default_on()))
        .unwrap()
        .set_discard_image_editing_data(Some(OnOff::on()))
        .unwrap()
        .set_default_image_dpi(Some(144))
        .unwrap();
    for prefix in ["w", "word", ""] {
        let fragment = values.to_xml(prefix);
        // No enclosing namespace bindings can mask a missing declaration.
        let mut reader = NsReader::from_str(&fragment);
        let mut count = 0;
        loop {
            let (namespace, event) = reader.read_resolved_event().unwrap();
            match event {
                Event::Empty(element) => {
                    assert!(matches!(namespace, ResolveResult::Bound(_)));
                    for attr in element.attributes() {
                        let attr = attr.unwrap();
                        if !attr.key.as_ref().starts_with(b"xmlns:") {
                            assert!(matches!(
                                reader.resolver().resolve_attribute(attr.key).0,
                                ResolveResult::Bound(_)
                            ));
                        }
                    }
                    count += 1;
                },
                Event::Eof => break,
                _ => panic!("expected a self-contained extension child"),
            }
        }
        assert_eq!(count, 6);
        let xml = format!(r#"<w:settings xmlns:w="{W}">{fragment}</w:settings>"#);
        assert_eq!(Extensions::parse(xml.as_bytes()).unwrap(), values);
    }
}
