use super::*;
use crate::Variables;
use crate::error::Error;
use crate::numbering::Format;
use crate::settings::{COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI};
use crate::settings::{
    ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag, NoteNumberingRestart,
    NotePosition, ProofState, ProofingState, ProtectionType, ThemeFontLanguages, View,
};
use litchi_opc::PackURI;
use litchi_opc::part::{BlobPart, Part};
use std::mem::size_of;

#[test]
fn test_settings_creation() {
    let settings = DocumentSettings::new();
    assert!(!settings.is_protected());
    assert!(settings.protection_type().is_none());
    assert!(!settings.track_revisions());
}

#[test]
fn test_protection_type() {
    assert_eq!(
        ProtectionType::from_xml("readOnly"),
        Some(ProtectionType::ReadOnly)
    );
    assert_eq!(
        ProtectionType::from_xml("comments"),
        Some(ProtectionType::Comments)
    );
    assert_eq!(
        ProtectionType::from_xml("trackedChanges"),
        Some(ProtectionType::TrackedChanges)
    );
    assert_eq!(
        ProtectionType::from_xml("forms"),
        Some(ProtectionType::Forms)
    );
    assert_eq!(ProtectionType::from_xml("invalid"), None);
}

#[test]
fn parses_smart_tag_settings_with_strict_namespaces() {
    let xml = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
        xmlns:false="urn:not-wordprocessingml">
        <s:trackRevisions s:val="on"/>
        <s:smartTagType s:namespaceuri="urn:contacts" s:name="person"
            s:url="https://example.test/schema?a=1&amp;b=2"/>
        <false:smartTagType false:namespaceuri="urn:false" false:name="ignored" false:url="ignored"/>
        <s:doNotEmbedSmartTags s:val="off"/>
    </s:settings>"#;

    let settings = DocumentSettings::extract_from_xml(xml).unwrap();
    assert!(settings.track_revisions());
    assert!(!settings.do_not_embed_smart_tags());
    assert_eq!(settings.smart_tag_types().len(), 1);
    assert_eq!(
        settings.smart_tag_types()[0].namespace_uri(),
        "urn:contacts"
    );
    assert_eq!(settings.smart_tag_types()[0].name(), "person");
    assert_eq!(
        settings.smart_tag_types()[0].url(),
        "https://example.test/schema?a=1&b=2"
    );
}

#[test]
fn validates_smart_tag_settings() {
    let enabled = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/></w:settings>"#;
    assert!(
        DocumentSettings::extract_from_xml(enabled)
            .unwrap()
            .do_not_embed_smart_tags()
    );

    let missing_url = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTagType w:namespaceuri="urn:test" w:name="test"/></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(missing_url).is_err());

    let oversized_name = "n".repeat(crate::settings::MAX_SMART_TAG_NAME_CHARS + 1);
    let oversized = format!(
        r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTagType w:namespaceuri="urn:test" w:name="{oversized_name}" w:url="https://example.test"/></w:settings>"#
    );
    assert!(matches!(
        DocumentSettings::extract_from_xml(oversized.as_bytes()),
        Err(Error::InvalidFormat(message)) if message.contains("smart-tag name")
    ));

    let invalid_on_off = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags w:val="maybe"/></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(invalid_on_off).is_err());

    let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/><w:doNotEmbedSmartTags/></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(duplicate).is_err());
}

#[test]
fn parses_compat_options_and_settings() {
    let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:compat><w:useFELayout/><w:doNotExpandShiftReturn w:val="off"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:rsids/></w:settings>"#;
    let settings = DocumentSettings::extract_from_xml(xml).unwrap();
    assert_eq!(settings.zoom_percent(), Some(100));
    assert_eq!(settings.compatibility_options().len(), 2);
    assert_eq!(
        settings.compatibility_options()[0].flag(),
        CompatFlag::UseFarEastLayout
    );
    assert!(settings.compatibility_options()[0].is_enabled());
    assert_eq!(
        settings.compatibility_options()[1].flag(),
        CompatFlag::DoNotExpandShiftReturn
    );
    assert!(!settings.compatibility_options()[1].is_enabled());
    assert_eq!(settings.compatibility_settings().len(), 2);
    let mode = settings
        .compatibility_setting(COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI)
        .unwrap();
    assert_eq!(mode.value(), "14");
    assert_eq!(settings.compatibility_mode(), Some(14));
    assert!(
        settings
            .compatibility_setting("missing", COMPATIBILITY_SETTING_URI)
            .is_none()
    );
}

#[test]
fn strict_and_transitional_compatibility_flag_domains_are_exhaustive() {
    const TRANSITIONAL_TOKENS: [&str; 65] = [
        "useSingleBorderforContiguousCells",
        "wpJustification",
        "noTabHangInd",
        "noLeading",
        "spaceForUL",
        "noColumnBalance",
        "balanceSingleByteDoubleByteWidth",
        "noExtraLineSpacing",
        "doNotLeaveBackslashAlone",
        "ulTrailSpace",
        "doNotExpandShiftReturn",
        "spacingInWholePoints",
        "lineWrapLikeWord6",
        "printBodyTextBeforeHeader",
        "printColBlack",
        "wpSpaceWidth",
        "showBreaksInFrames",
        "subFontBySize",
        "suppressBottomSpacing",
        "suppressTopSpacing",
        "suppressSpacingAtTopOfPage",
        "suppressTopSpacingWP",
        "suppressSpBfAfterPgBrk",
        "swapBordersFacingPages",
        "convMailMergeEsc",
        "truncateFontHeightsLikeWP6",
        "mwSmallCaps",
        "usePrinterMetrics",
        "doNotSuppressParagraphBorders",
        "wrapTrailSpaces",
        "footnoteLayoutLikeWW8",
        "shapeLayoutLikeWW8",
        "alignTablesRowByRow",
        "forgetLastTabAlignment",
        "adjustLineHeightInTable",
        "autoSpaceLikeWord95",
        "noSpaceRaiseLower",
        "doNotUseHTMLParagraphAutoSpacing",
        "layoutRawTableWidth",
        "layoutTableRowsApart",
        "useWord97LineBreakRules",
        "doNotBreakWrappedTables",
        "doNotSnapToGridInCell",
        "selectFldWithFirstOrLastChar",
        "applyBreakingRules",
        "doNotWrapTextWithPunct",
        "doNotUseEastAsianBreakRules",
        "useWord2002TableStyleRules",
        "growAutofit",
        "useFELayout",
        "useNormalStyleForList",
        "doNotUseIndentAsNumberingTabStop",
        "useAltKinsokuLineBreakRules",
        "allowSpaceOfSameStyleInTable",
        "doNotSuppressIndentation",
        "doNotAutofitConstrainedTables",
        "autofitToFirstFixedWidthCell",
        "underlineTabInNumList",
        "displayHangulFixedWidth",
        "splitPgBreakAndParaMark",
        "doNotVertAlignCellWithSp",
        "doNotBreakConstrainedForcedTable",
        "doNotVertAlignInTxbx",
        "useAnsiKerningPairs",
        "cachedColBalance",
    ];
    const STRICT_TOKENS: [&str; 7] = [
        "spaceForUL",
        "balanceSingleByteDoubleByteWidth",
        "doNotLeaveBackslashAlone",
        "ulTrailSpace",
        "doNotExpandShiftReturn",
        "adjustLineHeightInTable",
        "applyBreakingRules",
    ];

    assert_eq!(CompatFlag::ALL.len(), TRANSITIONAL_TOKENS.len());
    let mut flags = std::collections::HashSet::new();
    for (flag, raw) in CompatFlag::ALL.iter().copied().zip(TRANSITIONAL_TOKENS) {
        assert!(flags.insert(flag), "duplicate compatibility flag {raw}");
        assert_eq!(raw.parse(), Ok(flag));
        assert_eq!(flag.as_str(), raw);
        assert_eq!(flag.to_string(), raw);
        assert_eq!(flag.is_strict(), STRICT_TOKENS.contains(&raw));

        let transitional = format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:{raw}/></w:compat></w:settings>"#
        );
        let parsed = DocumentSettings::extract_from_xml(transitional.as_bytes()).unwrap();
        assert_eq!(parsed.compatibility_options()[0].flag(), flag);

        let strict = format!(
            r#"<w:settings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:compat><w:{raw}/></w:compat></w:settings>"#
        );
        let parsed = DocumentSettings::extract_from_xml(strict.as_bytes());
        if flag.is_strict() {
            assert_eq!(parsed.unwrap().compatibility_options()[0].flag(), flag);
        } else {
            assert!(parsed.is_err(), "Strict accepted Transitional-only {raw}");
        }
    }
    assert_eq!(flags.len(), 65);
    assert_eq!(CompatFlag::CachedColumnBalance as u8, 64);
    assert_eq!(size_of::<CompatFlag>(), 1);
    assert!("vendorCompat".parse::<CompatFlag>().is_err());
    assert!("UseFELayout".parse::<CompatFlag>().is_err());

    for namespace in [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let unknown = format!(
            r#"<w:settings xmlns:w="{namespace}"><w:compat><w:vendorCompat/></w:compat></w:settings>"#
        );
        assert!(DocumentSettings::extract_from_xml(unknown.as_bytes()).is_err());

        let duplicate = format!(
            r#"<w:settings xmlns:w="{namespace}"><w:compat><w:spaceForUL/><w:spaceForUL/></w:compat></w:settings>"#
        );
        assert!(DocumentSettings::extract_from_xml(duplicate.as_bytes()).is_err());
    }
}

#[test]
fn parses_empty_and_strict_compat_groups() {
    let empty = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/></w:settings>"#;
    let settings = DocumentSettings::extract_from_xml(empty).unwrap();
    assert!(settings.compatibility_options().is_empty());
    assert!(settings.compatibility_settings().is_empty());
    assert_eq!(settings.compatibility_mode(), None);

    let strict = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat><s:compatSetting s:name="compatibilityMode" s:uri="http://schemas.microsoft.com/office/word" s:val="15"/></s:compat></s:settings>"#;
    let settings = DocumentSettings::extract_from_xml(strict).unwrap();
    assert_eq!(settings.compatibility_mode(), Some(15));
}

#[test]
fn rejects_invalid_compat_groups() {
    let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/><w:compat/></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(duplicate).is_err());

    let missing_value = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"/></w:compat></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(missing_value).is_err());

    let unterminated = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>"#;
    assert!(DocumentSettings::extract_from_xml(unterminated).is_err());
}

#[test]
fn parses_document_level_note_properties() {
    let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:footnote w:type="separator" w:id="-1"/><w:pos w:val="pageBottom"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachPage"/></w:footnotePr><w:endnotePr><w:pos w:val="docEnd"/><w:numFmt w:val="upperLetter"/></w:endnotePr></w:settings>"#;
    let settings = DocumentSettings::extract_from_xml(xml).unwrap();

    let footnotes = settings.footnote_properties().unwrap();
    assert_eq!(footnotes.position(), Some(NotePosition::PageBottom));
    assert_eq!(footnotes.format(), Some(Format::LowerRoman));
    assert_eq!(footnotes.start(), Some(2));
    assert_eq!(footnotes.restart(), Some(NoteNumberingRestart::EachPage));

    let endnotes = settings.endnote_properties().unwrap();
    assert_eq!(endnotes.position(), Some(NotePosition::DocumentEnd));
    assert_eq!(endnotes.format(), Some(Format::UpperLetter));
    assert_eq!(endnotes.start(), None);
    assert_eq!(endnotes.restart(), None);
}

#[test]
fn strict_and_transitional_note_position_domains_are_closed() {
    for (raw, expected) in [
        ("pageBottom", NotePosition::PageBottom),
        ("beneathText", NotePosition::BeneathText),
        ("sectEnd", NotePosition::SectionEnd),
        ("docEnd", NotePosition::DocumentEnd),
    ] {
        assert_eq!(raw.parse(), Ok(expected));
        assert_eq!(expected.as_str(), raw);
        assert_eq!(expected.to_string(), raw);

        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let xml = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:footnotePr><w:pos w:val="{raw}"/></w:footnotePr></w:settings>"#
            );
            assert_eq!(
                DocumentSettings::extract_from_xml(xml.as_bytes())
                    .unwrap()
                    .footnote_properties()
                    .unwrap()
                    .position(),
                Some(expected)
            );
        }
    }
    assert!("vendorPosition".parse::<NotePosition>().is_err());
    assert!("PageBottom".parse::<NotePosition>().is_err());
    assert_eq!(size_of::<NotePosition>(), 1);

    for (raw, expected) in [
        ("sectEnd", NotePosition::SectionEnd),
        ("docEnd", NotePosition::DocumentEnd),
    ] {
        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let xml = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:endnotePr><w:pos w:val="{raw}"/></w:endnotePr></w:settings>"#
            );
            assert_eq!(
                DocumentSettings::extract_from_xml(xml.as_bytes())
                    .unwrap()
                    .endnote_properties()
                    .unwrap()
                    .position(),
                Some(expected)
            );
        }
    }

    for raw in ["pageBottom", "beneathText"] {
        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let xml = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:endnotePr><w:pos w:val="{raw}"/></w:endnotePr></w:settings>"#
            );
            assert!(DocumentSettings::extract_from_xml(xml.as_bytes()).is_err());
        }
    }

    for (raw, expected) in [
        ("continuous", NoteNumberingRestart::Continuous),
        ("eachSect", NoteNumberingRestart::EachSection),
        ("eachPage", NoteNumberingRestart::EachPage),
    ] {
        assert_eq!(NoteNumberingRestart::from_xml(raw).unwrap(), expected);
        assert_eq!(expected.as_str(), raw);
    }
    assert!(NoteNumberingRestart::from_xml("sometimes").is_err());
}

#[test]
fn rejects_invalid_note_property_groups() {
    let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr/><w:footnotePr/></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(duplicate).is_err());

    let bad_start = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnotePr><w:numStart w:val="soon"/></w:endnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(bad_start).is_err());

    let bad_restart = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numRestart w:val="sometimes"/></w:footnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(bad_restart).is_err());

    let bad_position = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:pos w:val="vendorPosition"/></w:footnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(bad_position).is_err());

    let footnote_only_position = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnotePr><w:pos w:val="pageBottom"/></w:endnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(footnote_only_position).is_err());

    let bad_format = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numFmt w:val="vendorNumbering"/></w:footnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(bad_format).is_err());

    let duplicate_child = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numFmt w:val="decimal"/><w:numFmt w:val="bullet"/></w:footnotePr></w:settings>"#;
    assert!(DocumentSettings::extract_from_xml(duplicate_child).is_err());
}

#[test]
fn parses_view_proofing_and_theme_defaults() {
    let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection/><w:view w:val="print"/><w:proofState w:spelling="clean" w:grammar="dirty"/><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:hyperlink="hyperlink"/></w:settings>"#;
    let settings = DocumentSettings::extract_from_xml(xml).unwrap();
    assert!(settings.is_write_protected());
    assert_eq!(settings.view(), Some(View::Print));
    let proofing = settings.proofing_state().unwrap();
    assert_eq!(proofing.spelling(), Some(ProofState::Clean));
    assert_eq!(proofing.grammar(), Some(ProofState::Dirty));
    assert_eq!(settings.default_tab_stop_twips(), Some(720));
    let languages = settings.theme_font_languages().unwrap();
    assert_eq!(languages.latin(), Some("en-US"));
    assert_eq!(languages.east_asia(), Some("ja-JP"));
    assert_eq!(languages.bidi(), Some("ar-SA"));
    let mapping = settings.color_scheme_mapping().unwrap();
    assert!(!mapping.is_empty());
    assert_eq!(
        mapping.get(ColorSchemeSlot::Background1),
        Some(ColorSchemeIndex::Light1)
    );
    assert_eq!(
        mapping.get(ColorSchemeSlot::Text1),
        Some(ColorSchemeIndex::Dark1)
    );
    assert_eq!(
        mapping.get(ColorSchemeSlot::Hyperlink),
        Some(ColorSchemeIndex::Hyperlink)
    );
    assert_eq!(mapping.get(ColorSchemeSlot::Accent1), None);

    let strict = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:writeProtection s:val="off"/><s:view s:val="web"/><s:proofState/></s:settings>"#;
    let settings = DocumentSettings::extract_from_xml(strict).unwrap();
    assert!(!settings.is_write_protected());
    assert_eq!(settings.view(), Some(View::Web));
    let proofing = settings.proofing_state().unwrap();
    assert_eq!(proofing.spelling(), None);
    assert_eq!(proofing.grammar(), None);
}

#[test]
fn editing_settings_enums_round_trip() {
    for (raw, expected) in [
        ("none", View::None),
        ("print", View::Print),
        ("outline", View::Outline),
        ("masterPages", View::MasterPages),
        ("normal", View::Normal),
        ("web", View::Web),
    ] {
        assert_eq!(View::from_xml(raw).unwrap(), expected);
        assert_eq!(expected.as_str(), raw);
    }
    assert!(View::from_xml("immersive").is_err());

    for (raw, expected) in [("clean", ProofState::Clean), ("dirty", ProofState::Dirty)] {
        assert_eq!(ProofState::from_xml(raw).unwrap(), expected);
        assert_eq!(expected.as_str(), raw);
    }
    assert!(ProofState::from_xml("pending").is_err());

    assert_eq!(ColorSchemeSlot::COUNT, 12);
    for (raw, expected) in [
        ("dark1", ColorSchemeIndex::Dark1),
        ("light2", ColorSchemeIndex::Light2),
        ("accent6", ColorSchemeIndex::Accent6),
        ("followedHyperlink", ColorSchemeIndex::FollowedHyperlink),
    ] {
        assert_eq!(ColorSchemeIndex::from_xml(raw).unwrap(), expected);
        assert_eq!(expected.as_str(), raw);
    }
    assert!(ColorSchemeIndex::from_xml("accent7").is_err());
}

#[test]
fn editing_settings_serialize_and_reparse() {
    let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection/><w:view w:val="outline"/><w:proofState w:spelling="dirty"/><w:defaultTabStop w:val="1440"/><w:themeFontLang w:val="en-US" w:bidi="he-IL"/><w:clrSchemeMapping w:bg1="dark2" w:accent3="accent1" w:followedHyperlink="hyperlink"/></w:settings>"#;
    let settings = DocumentSettings::extract_from_xml(xml).unwrap();

    let fragment = settings.to_editing_settings_xml("w");
    assert_eq!(
        fragment,
        r#"<w:writeProtection/><w:view w:val="outline"/><w:proofState w:spelling="dirty"/><w:defaultTabStop w:val="1440"/><w:themeFontLang w:val="en-US" w:bidi="he-IL"/><w:clrSchemeMapping w:bg1="dark2" w:accent3="accent1" w:followedHyperlink="hyperlink"/>"#
    );

    let reparsed_xml = format!(
        r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{fragment}</w:settings>"#
    );
    let reparsed = DocumentSettings::extract_from_xml(reparsed_xml.as_bytes()).unwrap();
    assert_eq!(reparsed.is_write_protected(), settings.is_write_protected());
    assert_eq!(reparsed.view(), settings.view());
    assert_eq!(reparsed.proofing_state(), settings.proofing_state());
    assert_eq!(
        reparsed.default_tab_stop_twips(),
        settings.default_tab_stop_twips()
    );
    assert_eq!(
        reparsed.theme_font_languages(),
        settings.theme_font_languages()
    );
    assert_eq!(
        reparsed.color_scheme_mapping(),
        settings.color_scheme_mapping()
    );
    // Serializing the reparsed settings is stable.
    assert_eq!(reparsed.to_editing_settings_xml("w"), fragment);
}

#[test]
fn editing_settings_builders_write_fragments() {
    let mut proofing = ProofingState::new();
    proofing
        .set_spelling(Some(ProofState::Clean))
        .set_grammar(Some(ProofState::Dirty));
    assert_eq!(
        proofing.to_xml("w"),
        r#"<w:proofState w:spelling="clean" w:grammar="dirty"/>"#
    );

    let mut languages = ThemeFontLanguages::new();
    languages
        .set_latin(Some("fr-FR".to_owned()))
        .unwrap()
        .set_bidi(None)
        .unwrap();
    assert!(languages.set_latin(Some(String::new())).is_err());
    assert_eq!(languages.to_xml("w"), r#"<w:themeFontLang w:val="fr-FR"/>"#);

    let mut mapping = ColorSchemeMapping::new();
    assert!(mapping.is_empty());
    mapping
        .set(ColorSchemeSlot::Background1, ColorSchemeIndex::Light1)
        .set(ColorSchemeSlot::Text1, ColorSchemeIndex::Dark1);
    mapping.clear(ColorSchemeSlot::Text1);
    assert_eq!(
        mapping.iter().collect::<Vec<_>>(),
        [(ColorSchemeSlot::Background1, ColorSchemeIndex::Light1)]
    );
    assert_eq!(
        mapping.to_xml("w"),
        r#"<w:clrSchemeMapping w:bg1="light1"/>"#
    );
}

#[test]
fn rejects_invalid_editing_settings() {
    let prefix =
        br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#;
    let suffix = br#"</w:settings>"#;
    let reject = |body: &[u8]| {
        let mut xml = prefix.to_vec();
        xml.extend_from_slice(body);
        xml.extend_from_slice(suffix);
        DocumentSettings::extract_from_xml(&xml)
    };

    assert!(reject(br#"<w:view/>"#).is_err());
    assert!(reject(br#"<w:view w:val="immersive"/>"#).is_err());
    assert!(reject(br#"<w:view w:val="print"/><w:view w:val="web"/>"#).is_err());
    assert!(reject(br#"<w:proofState w:spelling="pending"/>"#).is_err());
    assert!(reject(br#"<w:proofState w:grammar="maybe"/>"#).is_err());
    assert!(reject(br#"<w:proofState/><w:proofState/>"#).is_err());
    assert!(reject(br#"<w:defaultTabStop/>"#).is_err());
    assert!(reject(br#"<w:defaultTabStop w:val="wide"/>"#).is_err());
    assert!(reject(br#"<w:defaultTabStop w:val="-720"/>"#).is_err());
    assert!(reject(br#"<w:defaultTabStop w:val="99999999999999999999"/>"#).is_err());
    assert!(reject(br#"<w:defaultTabStop w:val="720"/><w:defaultTabStop w:val="720"/>"#).is_err());
    assert!(reject(br#"<w:themeFontLang w:val=""/>"#).is_err());
    assert!(reject(br#"<w:themeFontLang w:eastAsia="ja&#0;-JP"/>"#).is_err());
    assert!(reject(br#"<w:themeFontLang w:val="en-US"/><w:themeFontLang/>"#).is_err());
    assert!(reject(br#"<w:clrSchemeMapping w:bg1="light7"/>"#).is_err());
    assert!(reject(br#"<w:clrSchemeMapping/><w:clrSchemeMapping/>"#).is_err());
    assert!(reject(br#"<w:writeProtection/><w:writeProtection/>"#).is_err());
    assert!(reject(br#"<w:writeProtection w:val="maybe"/>"#).is_err());
}

#[test]
fn font_embedding_inserts_each_missing_flag_in_schema_order() {
    let missing_embed = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:saveSubsetFonts q:val="on"/><q:saveFormsData/></q:settings>"#;
    let patched = patch_font_embedding(missing_embed, true).unwrap();
    assert_eq!(
        patched,
        br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts/><q:saveSubsetFonts q:val="on"/><q:saveFormsData/></q:settings>"#
    );

    let missing_subset = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts q:val="true"/><q:embedSystemFonts/><q:saveFormsData/></q:settings>"#;
    let patched = patch_font_embedding(missing_subset, true).unwrap();
    assert_eq!(
        patched,
        br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts q:val="true"/><q:embedSystemFonts/><q:saveSubsetFonts/><q:saveFormsData/></q:settings>"#
    );
}

#[test]
fn font_embedding_rewrites_false_word_flags_without_touching_foreign_twins() {
    let xml = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts q:val="false" x:val="true"><x:keep/></q:embedTrueTypeFonts><q:saveSubsetFonts q:val="0"/><x:saveSubsetFonts/></q:settings>"#;
    let enabled = patch_font_embedding(xml, true).unwrap();
    assert_eq!(
        enabled,
        br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts/><q:saveSubsetFonts/><x:saveSubsetFonts/></q:settings>"#
    );

    let full_font = patch_font_embedding(&enabled, false).unwrap();
    assert_eq!(
        full_font,
        br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts/><x:saveSubsetFonts/></q:settings>"#
    );
}

#[test]
fn font_embedding_expands_self_closing_strict_root() {
    let xml = br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
    assert_eq!(
        patch_font_embedding(xml, true).unwrap(),
        br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:embedTrueTypeFonts/><s:saveSubsetFonts/></s:settings>"#
    );
    assert_eq!(
        patch_font_embedding(xml, false).unwrap(),
        br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:embedTrueTypeFonts/></s:settings>"#
    );

    let default_namespace =
        br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    assert_eq!(
        patch_font_embedding(default_namespace, true).unwrap(),
        br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><embedTrueTypeFonts/><saveSubsetFonts/></settings>"#
    );
}

#[test]
fn font_embedding_matching_flags_are_an_exact_namespace_aware_noop() {
    let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><!--keep--><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts x:val="off"/><q:saveSubsetFonts q:val="on" x:val="off"/><x:saveSubsetFonts/></q:settings>"#;
    assert_eq!(patch_font_embedding(xml, true).unwrap(), xml);

    let explicit_false = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:embedTrueTypeFonts/><q:saveSubsetFonts q:val="off"/></q:settings>"#;
    assert_eq!(
        patch_font_embedding(explicit_false, false).unwrap(),
        explicit_false
    );
}

#[test]
fn font_embedding_rejects_duplicate_or_invalid_word_flags() {
    let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:embedTrueTypeFonts/><w:embedTrueTypeFonts/></w:settings>"#;
    assert!(patch_font_embedding(duplicate, true).is_err());

    let invalid = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:embedTrueTypeFonts w:val="maybe"/></w:settings>"#;
    assert!(patch_font_embedding(invalid, true).is_err());

    let mut utf16 = vec![0xFF, 0xFE];
    for unit in
        r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
            .encode_utf16()
    {
        utf16.extend_from_slice(&unit.to_le_bytes());
    }
    assert!(patch_font_embedding(&utf16, true).is_err());
}

#[test]
fn parses_bundled_settings_resource() {
    let settings =
        DocumentSettings::extract_from_xml(include_bytes!("../../resources/settings.xml")).unwrap();
    assert_eq!(settings.compatibility_mode(), Some(14));
    assert_eq!(settings.compatibility_settings().len(), 4);
    assert!(
        settings
            .compatibility_options()
            .iter()
            .any(|option| { option.flag() == CompatFlag::UseFarEastLayout && option.is_enabled() })
    );
    let proofing = settings.proofing_state().unwrap();
    assert_eq!(proofing.spelling(), Some(ProofState::Clean));
    assert_eq!(proofing.grammar(), Some(ProofState::Clean));
    assert_eq!(settings.default_tab_stop_twips(), Some(720));
    let languages = settings.theme_font_languages().unwrap();
    assert_eq!(languages.latin(), Some("en-US"));
    assert_eq!(languages.east_asia(), Some("ja-JP"));
    let mapping = settings.color_scheme_mapping().unwrap();
    for (slot, expected) in [
        (ColorSchemeSlot::Background1, ColorSchemeIndex::Light1),
        (ColorSchemeSlot::Text1, ColorSchemeIndex::Dark1),
        (ColorSchemeSlot::Background2, ColorSchemeIndex::Light2),
        (ColorSchemeSlot::Text2, ColorSchemeIndex::Dark2),
        (ColorSchemeSlot::Accent1, ColorSchemeIndex::Accent1),
        (ColorSchemeSlot::Accent6, ColorSchemeIndex::Accent6),
        (ColorSchemeSlot::Hyperlink, ColorSchemeIndex::Hyperlink),
        (
            ColorSchemeSlot::FollowedHyperlink,
            ColorSchemeIndex::FollowedHyperlink,
        ),
    ] {
        assert_eq!(mapping.get(slot), Some(expected));
    }
}

fn attached_template_part(xml: &[u8], reltype: &str, target: &str, external: bool) -> BlobPart {
    let mut part = BlobPart::new(
        PackURI::new("/word/settings.xml").unwrap(),
        litchi_opc::constants::content_type::WML_SETTINGS.to_owned(),
        xml.to_vec(),
    );
    part.rels_mut().add_relationship(
        reltype.to_owned(),
        target.to_owned(),
        "customRel".to_owned(),
        external,
    );
    part
}

#[test]
fn parses_transitional_and_strict_attached_templates() {
    for (word, relationships, reltype) in [
        (
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ATTACHED_TEMPLATE_RELATIONSHIP,
        ),
        (
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            STRICT_ATTACHED_TEMPLATE_RELATIONSHIP,
        ),
    ] {
        let xml = format!(
            r#"<q:settings xmlns:q="{word}" xmlns:rel="{relationships}"><q:attachedTemplate rel:id="customRel"/></q:settings>"#
        );
        let part = attached_template_part(
            xml.as_bytes(),
            reltype,
            "file:///templates/Corporate.dotx",
            true,
        );
        let settings = DocumentSettings::extract_from_part(&part).unwrap();
        let attached = settings.attached_template().unwrap();
        assert_eq!(attached.relationship_id(), "customRel");
        assert_eq!(attached.target_uri(), "file:///templates/Corporate.dotx");
    }
}

#[test]
fn rejects_invalid_attached_template_graphs() {
    let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:attachedTemplate r:id="customRel"/></w:settings>"#;
    let missing = BlobPart::new(
        PackURI::new("/word/settings.xml").unwrap(),
        litchi_opc::constants::content_type::WML_SETTINGS.to_owned(),
        xml.to_vec(),
    );
    assert!(DocumentSettings::extract_from_part(&missing).is_err());

    let wrong_type = attached_template_part(xml, "urn:wrong", "file:///a.dotx", true);
    assert!(DocumentSettings::extract_from_part(&wrong_type).is_err());
    let internal =
        attached_template_part(xml, ATTACHED_TEMPLATE_RELATIONSHIP, "template.dotx", false);
    assert!(DocumentSettings::extract_from_part(&internal).is_err());
    let whitespace = attached_template_part(
        xml,
        ATTACHED_TEMPLATE_RELATIONSHIP,
        "file:///bad path.dotx",
        true,
    );
    assert!(DocumentSettings::extract_from_part(&whitespace).is_err());

    let mut duplicate =
        attached_template_part(xml, ATTACHED_TEMPLATE_RELATIONSHIP, "file:///a.dotx", true);
    duplicate.rels_mut().add_relationship(
        STRICT_ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
        "file:///b.dotx".to_owned(),
        "duplicate".to_owned(),
        true,
    );
    assert!(DocumentSettings::extract_from_part(&duplicate).is_err());
}

#[test]
fn patches_only_the_attached_template_element() {
    let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="125"/><q:attachedTemplate rel:id="old"><x:ignored/></q:attachedTemplate><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#;
    let replaced = patch_attached_template(xml, Some("new-id")).unwrap();
    assert_eq!(
        String::from_utf8(replaced).unwrap(),
        r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="125"/><q:attachedTemplate rel:id="new-id"/><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#
    );

    let empty = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
    let inserted =
        String::from_utf8(patch_attached_template(empty, Some("rId7")).unwrap()).unwrap();
    assert_eq!(
        inserted,
        r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:attachedTemplate r:id="rId7" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"/></s:settings>"#
    );
}

#[test]
fn patches_document_variables_without_touching_unrelated_settings() {
    let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><q:docVars><q:docVar q:name="old" q:val="old-value"/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#;
    let mut variables = Variables::new();
    variables.insert("Company & Team", "A < B").unwrap();
    variables.insert("empty", "").unwrap();
    let patched = String::from_utf8(patch_document_variables(xml, &variables).unwrap()).unwrap();
    assert_eq!(
        patched,
        r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><q:docVars><q:docVar q:name="Company &amp; Team" q:val="A &lt; B"/><q:docVar q:name="empty" q:val=""/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
    );

    variables.clear();
    let removed =
        String::from_utf8(patch_document_variables(patched.as_bytes(), &variables).unwrap())
            .unwrap();
    assert_eq!(
        removed,
        r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
    );
}

#[test]
fn inserts_document_variables_into_empty_strict_root_in_schema_order() {
    let mut variables = Variables::new();
    variables.insert("strict", "value").unwrap();
    let empty = br#"<settings xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
    assert_eq!(
        String::from_utf8(patch_document_variables(empty, &variables).unwrap()).unwrap(),
        r#"<settings xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docVars xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docVar w:name="strict" w:val="value"/></w:docVars></settings>"#
    );

    let ordered = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat/><s:rsids/></s:settings>"#;
    assert_eq!(
        String::from_utf8(patch_document_variables(ordered, &variables).unwrap()).unwrap(),
        r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat/><s:docVars><s:docVar s:name="strict" s:val="value"/></s:docVars><s:rsids/></s:settings>"#
    );
}
