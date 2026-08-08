#![cfg(feature = "iwork")]

use litchi::iwork::{Document, ErrorKind, Options, Resource, SourceLimits, Stage};

#[test]
fn facade_error_details_are_typed_and_content_free() {
    const PRIVATE: &str = "private-path/member/authored-value";
    let defaults = SourceLimits::default();
    let source = SourceLimits::new(
        1,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_expanded_bytes(),
        defaults.max_decoded_bytes_per_item(),
    )
    .unwrap_or_else(|error| panic!("one-byte source profile must be valid: {error}"));
    let error = Document::from_bytes_with_options(
        PRIVATE.as_bytes(),
        Options::default().with_source(source),
    )
    .err()
    .unwrap_or_else(|| panic!("private marker must exceed a one-byte source profile"));

    assert_eq!(error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(error.stage(), Stage::Input);
    assert_eq!(error.format(), None);
    assert_eq!(error.resource(), Some(Resource::InputBytes));
    assert_eq!(error.observed(), Some(PRIVATE.len() as u64));
    assert_eq!(error.maximum(), Some(1));
    assert!(!error.to_string().contains(PRIVATE));
    assert!(!format!("{error:?}").contains(PRIVATE));
    assert!(std::error::Error::source(&error).is_none());
}

#[test]
fn exact_semantic_resource_vocabulary_is_public_and_copyable() {
    let resources = [
        Resource::Objects,
        Resource::Sheets,
        Resource::References,
        Resource::TextStorages,
        Resource::TextFragments,
        Resource::PayloadBytes,
        Resource::Fields,
        Resource::NestingDepth,
    ];
    let copied = resources;
    assert_eq!(resources, copied);
}
