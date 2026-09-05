//! Bounded, read-only qualification of native Pages text storage.
//!
//! Pages discovery qualifies writable storage and retains only the body text
//! length and section edges. The payload itself remains owned by the source
//! catalog; this adapter never rewrites or retains it.

use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_protos::pages_body_codec::{self, DecodeOptions};
use litchi_iwa_text_wire::{RewriteError, RewriteLimits, StorageValidation};

use super::{PackageError, PackageResult};

/// Minimal body information retained by the legacy Pages discovery bridge.
///
/// This migration-only value contains no text, styles, or generated messages.
#[doc(hidden)]
#[derive(Debug)]
pub struct BodyStorageDiscovery {
    utf16_len: usize,
    section_references: Vec<(u32, u64)>,
}

impl BodyStorageDiscovery {
    /// Validated length of all source text fragments in UTF-16 code units.
    #[must_use]
    pub const fn utf16_len(&self) -> usize {
        self.utf16_len
    }

    /// Consume the compact positional section references.
    #[must_use]
    pub fn into_section_references(self) -> Vec<(u32, u64)> {
        self.section_references
    }
}

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
fn validate(source: &[u8]) -> Result<StorageValidation, RewriteError> {
    let limits = PagesTextStorageLimits::for_source(source).into_wire()?;
    litchi_iwa_text_wire::validate_storage_with_limits(source, limits)
}

/// Return whether a source payload is a valid, bounded Pages text storage.
pub(crate) fn is_valid(source: &[u8]) -> bool {
    validate(source).is_ok()
}

/// Qualify the complete storage, then retain only the body discovery fields.
pub(crate) fn body_discovery(source: &[u8]) -> PackageResult<BodyStorageDiscovery> {
    let validation = validate(source).map_err(discovery_error)?;
    let profile = PagesTextStorageLimits::for_source(source);
    let wire_limits = WireLimits::default()
        .with_input_bytes(profile.message_bytes)
        .and_then(|limits| limits.with_fields(profile.fields))
        .and_then(|limits| limits.with_nesting(profile.nesting))
        .and_then(|limits| {
            limits.with_rewrite_work(profile.rewrite_work.min(WireLimits::MAX_REWRITE_WORK))
        })
        .map_err(discovery_error)?;
    let view = WireView::parse_with_limits(source, wire_limits).map_err(discovery_error)?;
    let mut section_references = Vec::new();
    // Full text-wire qualification above proves singularity, framing, entry
    // order/bounds, and reference validity. Repeated tables stay outside the
    // generated representation; only each selected entry forces a Buffa view.
    if let Some(table) = view.fields().find(|field| field.number() == 17) {
        let table =
            WireView::parse_with_limits(table.payload(), wire_limits).map_err(discovery_error)?;
        for field in table.fields().filter(|field| field.number() == 1) {
            let payload = field.payload();
            let options = DecodeOptions::new(
                payload.len(),
                payload.len(),
                payload.len().checked_mul(4).ok_or_else(|| {
                    PackageError::InvalidFormat("Pages body discovery work overflows".to_owned())
                })?,
                4,
            );
            let boundary = pages_body_codec::decode_section_boundary(payload, options)
                .map_err(discovery_error)?;
            // Object-less entries are valid positional sentinels. The host
            // has always omitted them from its discovered section graph.
            if let Some(section) = boundary.section() {
                section_references.try_reserve(1).map_err(|_allocation| {
                    PackageError::Allocation {
                        amount: section_references.len().saturating_add(1),
                    }
                })?;
                section_references.push((boundary.character_index(), section.identifier().get()));
            }
        }
    }
    Ok(BodyStorageDiscovery {
        utf16_len: validation.utf16_len(),
        section_references,
    })
}

fn discovery_error(error: impl std::fmt::Display) -> PackageError {
    PackageError::InvalidFormat(format!("Invalid Pages body discovery storage: {error}"))
}

#[cfg(test)]
mod tests {
    use super::{PagesTextStorageLimits, body_discovery, validate};
    use litchi_iwa_protos::{
        tsp::Reference,
        tswp::{ObjectAttributeTable, StorageArchive, object_attribute_table::ObjectAttribute},
    };
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
    fn body_projection_preserves_fragment_lengths_and_section_sentinels() {
        let mut source = StorageArchive {
            text: vec!["a".to_owned(), "😀".to_owned(), "b".to_owned()],
            table_section: Some(ObjectAttributeTable {
                entries: [(0, Some(44)), (3, Some(55)), (4, None)]
                    .into_iter()
                    .map(|(character_index, identifier)| ObjectAttribute {
                        character_index,
                        object: identifier.map(|identifier| Reference {
                            identifier,
                            ..Reference::default()
                        }),
                    })
                    .collect(),
            }),
            ..StorageArchive::default()
        }
        .encode_to_vec();
        source.extend_from_slice(&[0xa0, 0x06, 0x01]);
        let before = source.clone();
        let discovery = body_discovery(&source).expect("valid body discovery");
        assert_eq!(discovery.utf16_len(), 4);
        assert_eq!(discovery.into_section_references(), [(0, 44), (3, 55)]);
        assert_eq!(source, before);
    }

    #[test]
    fn body_projection_retains_no_text_when_section_table_is_absent() {
        let source = storage_source(&[&"😀".repeat(65_536), "", "tail"]);
        let discovery = body_discovery(&source).expect("bounded large text body");
        assert_eq!(discovery.utf16_len(), 131_076);
        assert!(discovery.into_section_references().is_empty());
    }

    #[test]
    fn body_projection_rejects_duplicate_section_table_and_malformed_text() {
        let mut duplicate = storage_source(&["body"]);
        duplicate.extend_from_slice(&[0x8a, 0x01, 0x00, 0x8a, 0x01, 0x00]);
        assert!(body_discovery(&duplicate).is_err());
        assert!(body_discovery(&[0x1a, 0x01, 0xff]).is_err());
    }

    #[test]
    fn body_projection_requires_schema_valid_deprecated_reference_types() {
        // Text-wire validates deprecated types as canonical varints. The
        // section codec additionally enforces the declared signed i32 range:
        // positive 2^31 cannot be silently truncated to a negative type.
        let source = [
            0x1a, 0x04, b'b', b'o', b'd', b'y', 0x8a, 0x01, 0x0e, 0x0a, 0x0c, 0x08, 0x00, 0x12,
            0x08, 0x08, 0x2c, 0x10, 0x80, 0x80, 0x80, 0x80, 0x08,
        ];
        assert!(validate(&source).is_ok());
        assert!(body_discovery(&source).is_err());

        let valid = StorageArchive {
            text: vec!["body".to_owned()],
            table_section: Some(ObjectAttributeTable {
                entries: vec![ObjectAttribute {
                    character_index: 0,
                    object: Some(Reference {
                        identifier: 44,
                        deprecated_type: Some(i32::MIN),
                        deprecated_is_external: Some(false),
                    }),
                }],
            }),
            ..StorageArchive::default()
        }
        .encode_to_vec();
        assert_eq!(
            body_discovery(&valid)
                .expect("signed canonical deprecated type")
                .into_section_references(),
            [(0, 44)],
        );
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
