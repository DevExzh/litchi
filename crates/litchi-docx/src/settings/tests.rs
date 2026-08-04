use super::*;

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

#[test]
fn parses_requested_settings_values_in_both_dialects() {
    for namespace in [TRANSITIONAL, STRICT] {
        let compatibility_flag = if namespace == STRICT {
            "spaceForUL"
        } else {
            "useFELayout"
        };
        let xml = format!(
            r#"<w:settings xmlns:w="{namespace}"><w:documentProtection w:edit="readOnly" w:enforcement="on"/><w:trackRevisions/><w:zoom w:percent="125"/><w:compat><w:{compatibility_flag}/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat><w:footnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachPage"/></w:footnotePr><w:view w:val="print"/><w:proofState w:spelling="clean"/><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1"/></w:settings>"#
        );
        let settings = Settings::parse(xml.as_bytes()).unwrap();
        assert!(settings.is_protected());
        assert_eq!(settings.protection_type(), Some(ProtectionType::ReadOnly));
        assert!(settings.track_revisions());
        assert_eq!(settings.zoom_percent(), Some(125));
        assert_eq!(settings.compatibility_options().len(), 1);
        assert_eq!(settings.compatibility_mode(), Some(14));
        assert_eq!(
            settings.footnote_properties().unwrap().format(),
            Some(NoteNumberFormat::LowerRoman)
        );
        assert_eq!(settings.view(), Some(DocumentView::Print));
        assert_eq!(
            settings.proofing_state().unwrap().spelling(),
            Some(ProofState::Clean)
        );
        assert_eq!(settings.default_tab_stop_twips(), Some(720));
        assert_eq!(
            settings.theme_font_languages().unwrap().latin(),
            Some("en-US")
        );
        assert_eq!(
            settings
                .color_scheme_mapping()
                .unwrap()
                .get(ColorSchemeSlot::Background1),
            Some(ColorSchemeIndex::Light1)
        );
    }
}

#[test]
fn rejects_strict_transitional_compatibility_flags_and_duplicates() {
    let strict = format!(
        r#"<w:settings xmlns:w="{STRICT}"><w:compat><w:useFELayout/></w:compat></w:settings>"#
    );
    assert!(Settings::parse(strict.as_bytes()).is_err());

    let duplicate = format!(
        r#"<w:settings xmlns:w="{TRANSITIONAL}"><w:view w:val="print"/><w:view w:val="web"/></w:settings>"#
    );
    assert!(Settings::parse(duplicate.as_bytes()).is_err());
}

#[test]
fn writes_modeled_values_in_schema_order() {
    let xml = format!(
        r#"<w:settings xmlns:w="{TRANSITIONAL}"><w:documentProtection w:edit="comments" w:enforcement="on"/><w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:view w:val="outline"/><w:proofState w:grammar="dirty"/><w:themeFontLang w:val="en-US"/></w:settings>"#
    );
    let settings = Settings::parse(xml.as_bytes()).unwrap();
    let output = settings.to_xml("w");
    assert!(output.starts_with("<w:documentProtection"));
    assert!(output.contains("<w:compat><w:useFELayout/>"));
    assert!(output.contains("<w:footnotePr><w:numFmt w:val=\"decimal\"/></w:footnotePr>"));
    assert!(output.ends_with("<w:themeFontLang w:val=\"en-US\"/>"));
}

#[test]
fn smart_tag_type_validates_client_lengths_without_rejecting_empty_present_values() {
    let value =
        SmartTagType::new("é".repeat(MAX_SMART_TAG_NAMESPACE_URI_CHARS), "name", "url").unwrap();
    assert_eq!(
        value.namespace_uri().chars().count(),
        MAX_SMART_TAG_NAMESPACE_URI_CHARS
    );

    assert!(
        SmartTagType::new("namespace", "n".repeat(MAX_SMART_TAG_NAME_CHARS + 1), "url",).is_err()
    );
    assert!(
        validate_smart_tag_type(
            "namespace",
            "name",
            &"u".repeat(MAX_SMART_TAG_URL_CHARS + 1),
        )
        .is_err()
    );

    let empty = SmartTagType::new("", "", "").unwrap();
    assert!(empty.namespace_uri().is_empty());
    assert!(empty.name().is_empty());
    assert!(empty.url().is_empty());
}

#[test]
fn rejects_oversized_input_before_xml_allocation() {
    let input = vec![b' '; MAX_SETTINGS_XML_BYTES + 1];
    assert!(Settings::parse(&input).is_err());
}
