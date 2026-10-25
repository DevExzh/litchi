#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{Error, OpenOptions, SourceLimitKind, SourceLimits, compress};

fn options(limits: SourceLimits) -> OpenOptions {
    OpenOptions::new().source_limits(limits)
}

fn limits_with_input_bytes(value: usize) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(value)
        .expect("input-byte limit is representable")
        .build()
        .expect("input-byte profile is consistent")
}

fn limits_with_events(value: usize) -> SourceLimits {
    SourceLimits::builder()
        .max_events(value)
        .expect("event limit is representable")
        .build()
        .expect("event profile is consistent")
}

#[test]
fn source_and_token_limits_are_typed() {
    let source = br"{\rtf1 body}";
    let limits = limits_with_input_bytes(source.len() - 1);
    assert!(matches!(
        Tree::parse_bytes(source, options(limits)),
        Err(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum,
        }) if observed == source.len() && maximum == source.len() - 1
    ));

    let limits = limits_with_events(3);
    assert!(matches!(
        Tree::parse_bytes(source, options(limits)),
        Err(Error::Limit {
            kind: SourceLimitKind::Events,
            observed: 4,
            maximum: 3,
        })
    ));
}

#[test]
fn binary_and_compressed_expansion_limits_flow_through_document_parsing() {
    let limits = SourceLimits::builder()
        .max_binary_bytes(3)
        .expect("binary limit is representable")
        .build()
        .expect("binary profile is consistent");
    assert!(matches!(
        Tree::parse(r"{\rtf1\bin4 ABCD}", options(limits)),
        Err(Error::Limit {
            kind: SourceLimitKind::BinaryBytes,
            observed: 4,
            maximum: 3,
        })
    ));

    let raw = br"{\rtf1\ansi bounded expansion}";
    let compressed = compress(raw, true, litchi_rtf::CompressionLimits::default()).unwrap();
    let limits = SourceLimits::builder()
        .max_decompressed_bytes(raw.len() - 1)
        .expect("decompression limit is representable")
        .build()
        .expect("decompression profile is consistent");
    assert!(matches!(
        Tree::parse_bytes(&compressed, options(limits)),
        Err(Error::Limit {
            kind: SourceLimitKind::DecompressedBytes,
            observed,
            maximum,
        }) if observed == raw.len() && maximum == raw.len() - 1
    ));
}

#[test]
fn bounded_file_open_rejects_metadata_before_reading() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/navigation-index.rtf");
    let observed = usize::try_from(std::fs::metadata(&fixture).unwrap().len()).unwrap();
    let limits = limits_with_input_bytes(observed - 1);
    assert!(matches!(
        Tree::open(&fixture, options(limits)),
        Err(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: actual,
            maximum,
        }) if actual == observed && maximum == observed - 1
    ));
}

#[test]
fn default_profile_preserves_existing_entry_points() {
    let source = br"{\rtf1\ansi finite defaults}";
    assert_eq!(
        Tree::parse_bytes(source, OpenOptions::default())
            .unwrap()
            .text(),
        "finite defaults"
    );
    assert_eq!(
        Tree::from_bytes(source, OpenOptions::default())
            .unwrap()
            .text(),
        "finite defaults"
    );
    assert_eq!(
        Tree::parse_bytes(source, options(SourceLimits::default()))
            .unwrap()
            .text(),
        "finite defaults"
    );
}

#[test]
fn checked_profile_names_and_getters_cover_every_rtf_dimension() {
    assert_eq!(SourceLimits::default().max_input_bytes(), 512 * 1_048_576);

    let limits = SourceLimits::builder()
        .max_input_bytes(1)
        .expect("input-byte limit")
        .max_events(2)
        .expect("event limit")
        .max_binary_bytes(3)
        .expect("binary limit")
        .max_aggregate_binary_bytes(4)
        .expect("aggregate binary limit")
        .max_decompressed_bytes(5)
        .expect("decompressed limit")
        .max_opaque_nodes(6)
        .expect("opaque-node limit")
        .max_opaque_node_bytes(7)
        .expect("opaque-node-byte limit")
        .max_aggregate_opaque_bytes(8)
        .expect("aggregate opaque-byte limit")
        .build()
        .expect("profile is consistent");

    assert_eq!(limits.max_input_bytes(), 1);
    assert_eq!(limits.max_events(), 2);
    assert_eq!(limits.max_binary_bytes(), 3);
    assert_eq!(limits.max_aggregate_binary_bytes(), 4);
    assert_eq!(limits.max_decompressed_bytes(), 5);
    assert_eq!(limits.max_opaque_nodes(), 6);
    assert_eq!(limits.max_opaque_node_bytes(), 7);
    assert_eq!(limits.max_aggregate_opaque_bytes(), 8);
}

#[test]
fn zero_opaque_nodes_is_an_explicit_denial_and_usize_max_is_invalid() {
    let zero_opaque = SourceLimits::builder()
        .max_opaque_nodes(0)
        .expect("zero is a supported explicit denial")
        .build()
        .expect("zero opaque-node profile is consistent");
    assert!(matches!(
        Tree::parse(r"{\rtf1\future1 x}", options(zero_opaque)),
        Err(Error::Limit {
            kind: SourceLimitKind::OpaqueNodes,
            observed: 1,
            maximum: 0,
        })
    ));

    assert!(matches!(
        SourceLimits::builder().max_input_bytes(usize::MAX),
        Err(Error::InvalidLimit {
            kind: SourceLimitKind::InputBytes,
            value,
            maximum,
        }) if value == usize::MAX && maximum == isize::MAX as usize
    ));
}
