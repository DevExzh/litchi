use litchi_rtf::{DocumentCompatibilityPolicy, DocumentFeatureThrottle, RtfDocument, RtfWriter};

fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_flags_and_preserves_explicit_zero_separately_from_omission() {
    let omitted = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    assert!(omitted.compatibility_policy().is_empty());
    assert_eq!(
        omitted.compatibility_policy().effective_feature_throttle(),
        DocumentFeatureThrottle::CompatibilityLimited
    );

    let explicit =
        RtfDocument::parse(r"{\rtf1\ansi\nocompatoptions\nofeaturethrottle0\forceupgrade Body}")
            .unwrap();
    assert!(explicit.compatibility_policy().reset_options_to_defaults);
    assert_eq!(
        explicit.compatibility_policy().feature_throttle,
        Some(DocumentFeatureThrottle::CompatibilityLimited)
    );
    assert!(explicit.compatibility_policy().force_upgrade);
    assert_eq!(explicit.text(), "Body");
}

#[test]
fn alias_pair_uses_specification_last_wins_semantics() {
    let numeric_last =
        RtfDocument::parse(r"{\rtf1\ansi\nouicompat\nofeaturethrottle0 Body}").unwrap();
    assert_eq!(
        numeric_last.compatibility_policy().feature_throttle,
        Some(DocumentFeatureThrottle::CompatibilityLimited)
    );

    let alias_last =
        RtfDocument::parse(r"{\rtf1\ansi\nofeaturethrottle0\nouicompat Body}").unwrap();
    assert_eq!(
        alias_last.compatibility_policy().feature_throttle,
        Some(DocumentFeatureThrottle::Unrestricted)
    );
}

#[test]
fn typed_api_round_trips_canonically_and_clears_without_upgrading() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_compatibility_policy(DocumentCompatibilityPolicy {
        reset_options_to_defaults: true,
        feature_throttle: Some(DocumentFeatureThrottle::Unrestricted),
        force_upgrade: true,
    });
    let serialized = write(&doc);
    let reset = serialized.find("\\nocompatoptions").unwrap();
    let throttle = serialized.find("\\nofeaturethrottle1").unwrap();
    let upgrade = serialized.find("\\forceupgrade").unwrap();
    assert!(reset < throttle && throttle < upgrade);
    assert!(!serialized.contains("\\nouicompat"));
    assert_eq!(serialized.matches("\\nofeaturethrottle").count(), 1);
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(
        *reparsed.compatibility_policy(),
        *doc.compatibility_policy()
    );
    assert_eq!(reparsed.text(), "Body");

    doc.clear_compatibility_policy();
    assert!(doc.compatibility_policy().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn force_upgrade_remains_passive_and_does_not_synthesize_throttle_policy() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_compatibility_policy(DocumentCompatibilityPolicy {
        force_upgrade: true,
        ..DocumentCompatibilityPolicy::default()
    });
    let serialized = write(&doc);
    assert!(serialized.contains("\\forceupgrade"));
    assert!(!serialized.contains("\\nofeaturethrottle"));
    assert!(!serialized.contains("\\nouicompat"));
    assert_eq!(doc.text(), "Body");
}

#[test]
fn rejects_malformed_duplicates_starred_grouped_and_late_controls() {
    for name in ["nocompatoptions", "nouicompat", "forceupgrade"] {
        for input in [
            format!(r"{{\rtf1\ansi\{name}0 Body}}"),
            format!(r"{{\rtf1\ansi\{name}\{name} Body}}"),
            format!(r"{{\rtf1\ansi{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi{{\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi Body\{name}}}"),
        ] {
            assert!(RtfDocument::parse(&input).is_err(), "accepted {input}");
        }
    }

    for input in [
        r"{\rtf1\ansi\nofeaturethrottle Body}",
        r"{\rtf1\ansi\nofeaturethrottle-1 Body}",
        r"{\rtf1\ansi\nofeaturethrottle2 Body}",
        r"{\rtf1\ansi\nofeaturethrottle1\nofeaturethrottle0 Body}",
        r"{\rtf1\ansi{\*\nofeaturethrottle1}Body}",
        r"{\rtf1\ansi{\nofeaturethrottle1}Body}",
        r"{\rtf1\ansi Body\nofeaturethrottle1}",
        r"{\rtf1\ansi\nofeaturethrottle999999999999 Body}",
    ] {
        assert!(RtfDocument::parse(input).is_err(), "accepted {input}");
    }
}

#[test]
fn parses_bundled_libreoffice_alias_sequence() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/hidden-linebreaks.rtf"
    );
    let doc = RtfDocument::parse_bytes(&std::fs::read(path).unwrap()).unwrap();
    assert_eq!(
        doc.compatibility_policy().feature_throttle,
        Some(DocumentFeatureThrottle::Unrestricted)
    );
    let serialized = write(&doc);
    assert!(serialized.contains("\\nofeaturethrottle1"));
    assert!(!serialized.contains("\\nouicompat"));
}
