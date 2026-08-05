//! Focused regression tests for the tracked-revision semantic layer.

use super::{RevisionKind, RevisionMetadata};

#[test]
fn metadata_builder_is_composable() {
    let metadata = RevisionMetadata::new("Alice")
        .with_reason(0x2b)
        .with_revision_save_id(42);

    assert_eq!(metadata.author, "Alice");
    assert_eq!(metadata.reason, Some(0x2b));
    assert_eq!(metadata.revision_save_id, Some(42));
    assert_eq!(metadata.timestamp, None);
}

#[test]
fn revision_kinds_are_copyable_and_distinct() {
    let insertion = RevisionKind::Insertion;
    let deletion = insertion;

    assert_eq!(insertion, deletion);
    assert_ne!(insertion, RevisionKind::Deletion);
}
