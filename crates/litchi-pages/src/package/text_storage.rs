//! Bounded, read-only qualification of native Pages text storage.
//!
//! Pages body discovery only needs a yes/no answer for whether one borrowed
//! payload is a writable `TSWP.StorageArchive`. The payload itself remains
//! owned by the source catalog; this adapter never rewrites or retains it.

use litchi_iwa_text_wire::{RewriteError, RewriteLimits};

/// Minimum schema depth needed by a complete `TSWP.StorageArchive` tree.
const REQUIRED_NESTING: usize = 4;

/// Finite limits used while qualifying one Pages text-storage payload.
///
/// Every axis is derived from the selected source length and then clamped to
/// the shared hard ceiling. This prevents a single malformed Pages object
/// from borrowing the text-wire default's process-wide allowance while still
/// accepting every source that fits within its own encoded byte envelope.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct PagesTextStorageLimits {
    message_bytes: usize,
    fields: usize,
    nesting: usize,
    fragments: usize,
    text_bytes: usize,
    table_entries: usize,
    object_references: usize,
    output_bytes: usize,
    rewrite_work: usize,
}

impl PagesTextStorageLimits {
    fn for_source(source: &[u8]) -> Self {
        let source_bytes = source.len().max(1);
        Self {
            message_bytes: source_bytes.min(RewriteLimits::MAX_MESSAGE_BYTES),
            fields: source_bytes.min(RewriteLimits::MAX_FIELDS),
            nesting: REQUIRED_NESTING,
            fragments: source_bytes.min(RewriteLimits::MAX_FRAGMENTS),
            text_bytes: source_bytes.min(RewriteLimits::MAX_TEXT_BYTES),
            table_entries: source_bytes.min(RewriteLimits::MAX_TABLE_ENTRIES),
            object_references: source_bytes.min(RewriteLimits::MAX_OBJECT_REFERENCES),
            output_bytes: source_bytes.min(RewriteLimits::MAX_OUTPUT_BYTES),
            rewrite_work: source_bytes
                .saturating_mul(16)
                .clamp(1, RewriteLimits::MAX_REWRITE_WORK),
        }
    }

    fn into_wire(self) -> Result<RewriteLimits, RewriteError> {
        RewriteLimits::new(
            self.message_bytes,
            self.fields,
            self.nesting,
            self.fragments,
            self.text_bytes,
            self.table_entries,
            self.object_references,
            self.output_bytes,
            self.rewrite_work,
        )
    }
}

/// Run the strict text-wire qualification without changing the source bytes.
fn validate(source: &[u8]) -> Result<(), RewriteError> {
    let limits = PagesTextStorageLimits::for_source(source).into_wire()?;
    litchi_iwa_text_wire::validate_storage_with_limits(source, limits).map(|_| ())
}

/// Return whether a source payload is a valid, bounded Pages text storage.
pub(crate) fn is_valid(source: &[u8]) -> bool {
    validate(source).is_ok()
}

#[cfg(test)]
mod tests {
    use super::{PagesTextStorageLimits, validate};
    use litchi_iwa_protos::tswp::StorageArchive;
    use prost::Message as _;

    fn storage_source(fragments: &[&str]) -> Vec<u8> {
        StorageArchive {
            text: fragments.iter().map(|text| (*text).to_owned()).collect(),
            ..StorageArchive::default()
        }
        .encode_to_vec()
    }

    #[test]
    fn qualification_is_bounded_and_preserves_source_bytes() {
        let mut source = storage_source(&["before ", "😀after", ""]);
        // Unknown fields stay in the caller-owned source and are not copied
        // into a retained semantic value by the qualification pass.
        source.extend_from_slice(&[0xa0, 0x06, 0x01]);
        let before = source.clone();

        assert!(validate(&source).is_ok());
        assert_eq!(source, before);
    }

    #[test]
    fn qualification_rejects_malformed_storage_wire() {
        assert!(validate(&[0x18, 0x01]).is_err());
        assert!(validate(&[0x1a, 0x01, 0xff]).is_err());
    }

    #[test]
    fn source_profile_is_exact_for_each_bounded_axis() {
        let source = storage_source(&["x"]);
        let limits = PagesTextStorageLimits::for_source(&source);
        assert_eq!(limits.message_bytes, source.len());
        assert_eq!(limits.fields, source.len());
        assert_eq!(limits.fragments, source.len());
        assert_eq!(limits.text_bytes, source.len());
        assert_eq!(limits.table_entries, source.len());
        assert_eq!(limits.object_references, source.len());
        assert_eq!(limits.output_bytes, source.len());
        assert_eq!(limits.nesting, 4);
        assert_eq!(limits.rewrite_work, source.len().saturating_mul(16));

        let wire = limits.into_wire().expect("source-derived profile is valid");
        assert!(litchi_iwa_text_wire::validate_storage_with_limits(&source, wire).is_ok());

        let one_over_fragments = StorageArchive {
            text: vec!["a".to_owned(), "b".to_owned()],
            ..StorageArchive::default()
        }
        .encode_to_vec();
        let mut one_fragment = PagesTextStorageLimits::for_source(&one_over_fragments);
        one_fragment.fragments = 1;
        let error = litchi_iwa_text_wire::validate_storage_with_limits(
            &one_over_fragments,
            one_fragment.into_wire().expect("test profile is valid"),
        )
        .expect_err("one fragment over the typed limit must fail");
        assert!(matches!(
            error,
            litchi_iwa_text_wire::RewriteError::LimitExceeded {
                resource: "text fragments",
                observed: 2,
                limit: 1,
            }
        ));
    }
}
