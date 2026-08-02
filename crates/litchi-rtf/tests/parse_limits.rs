use litchi_rtf::{ParseLimits, RtfDocument, RtfError, compress};

#[test]
fn source_and_token_limits_are_typed() {
    let source = br"{\rtf1 body}";
    let limits = ParseLimits::default().with_max_source_bytes(source.len() - 1);
    assert!(matches!(
        RtfDocument::parse_bytes_with_limits(source, limits),
        Err(RtfError::LimitExceeded {
            resource: "source bytes",
            observed,
            limit,
        }) if observed == source.len() && limit == source.len() - 1
    ));

    let limits = ParseLimits::default().with_max_tokens(3);
    assert!(matches!(
        RtfDocument::parse_bytes_with_limits(source, limits),
        Err(RtfError::LimitExceeded {
            resource: "lexer tokens",
            observed: 4,
            limit: 3,
        })
    ));
}

#[test]
fn binary_and_compressed_expansion_limits_flow_through_document_parsing() {
    let limits = ParseLimits::default().with_max_binary_bytes(3);
    assert!(matches!(
        RtfDocument::parse_with_limits(r"{\rtf1\bin4 ABCD}", limits),
        Err(RtfError::LimitExceeded {
            resource: "binary payload bytes",
            observed: 4,
            limit: 3,
        })
    ));

    let raw = br"{\rtf1\ansi bounded expansion}";
    let compressed = compress(raw, true).unwrap();
    let limits = ParseLimits::default().with_max_decompressed_bytes(raw.len() - 1);
    assert!(matches!(
        RtfDocument::from_bytes_with_limits(&compressed, limits),
        Err(RtfError::LimitExceeded {
            resource: "decompressed bytes",
            observed,
            limit,
        }) if observed == raw.len() && limit == raw.len() - 1
    ));
}

#[test]
fn bounded_file_open_rejects_metadata_before_reading() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/navigation-index.rtf");
    let observed = usize::try_from(std::fs::metadata(&fixture).unwrap().len()).unwrap();
    let limits = ParseLimits::default().with_max_source_bytes(observed - 1);
    assert!(matches!(
        RtfDocument::open_with_limits(&fixture, limits),
        Err(RtfError::LimitExceeded {
            resource: "source bytes",
            observed: actual,
            limit,
        }) if actual == observed && limit == observed - 1
    ));
}

#[test]
fn default_profile_preserves_existing_entry_points() {
    let source = br"{\rtf1\ansi finite defaults}";
    assert_eq!(
        RtfDocument::parse_bytes(source).unwrap().text(),
        "finite defaults"
    );
    assert_eq!(
        RtfDocument::from_bytes(source).unwrap().text(),
        "finite defaults"
    );
    assert_eq!(
        RtfDocument::parse_bytes_with_limits(source, ParseLimits::default())
            .unwrap()
            .text(),
        "finite defaults"
    );
}
