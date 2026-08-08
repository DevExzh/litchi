//! Bounded conversion from native TSWP storage messages to semantic text.
//!
//! This crate is the narrow wire adapter between generated IWA protobufs and
//! [`litchi_iwa_text`]. It owns no package traversal, archive state, object
//! lookup, or application semantics. Callers retain context-specific error
//! wording and transaction ownership at their format boundary.

#![forbid(unsafe_code)]

mod rewrite;

pub use rewrite::{
    RemovedObjectReference, RewriteBehavior, RewriteError, RewriteLimits, RewriteResult,
    StorageRewrite, StorageValidation, rewrite_storage_text_with_behavior_and_limits,
    rewrite_storage_text_with_limits, validate_storage_with_limits,
};

use std::{cell::Cell, mem::size_of};

use litchi_iwa_common::{WireLimits, wire::WireDescent, wire::preflight_wire_tree_with_limits};
use litchi_iwa_protos::text_storage_codec;
use litchi_iwa_protos::tswp::StorageArchive;
use litchi_iwa_text::storage::{Error as StorageError, MAX_RUNS, Run, Storage};

/// Maximum native text fragments accepted by one conversion.
pub const MAX_FRAGMENTS: usize = MAX_RUNS;
/// Default maximum encoded bytes accepted by the raw storage decoder.
pub const DEFAULT_MAX_MESSAGE_BYTES: usize = 64 * 1024 * 1024;
/// Default maximum root wire fields accepted by the raw storage decoder.
pub const DEFAULT_MAX_FIELDS: usize = 100_000;
/// Default maximum field-3 text fragments accepted by the raw storage decoder.
pub const DEFAULT_MAX_WIRE_FRAGMENTS: usize = 100_000;
/// Default and hard maximum aggregate text bytes in one decoded storage.
pub const DEFAULT_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite resource policy for decoding one raw `TSWP.StorageArchive`.
///
/// The profile governs the allocation-free common-wire preflight, Buffa's
/// borrowed repeated-string metadata, and the final semantic allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "Each finite maximum has a matching explicit accessor"
)]
pub struct Limits {
    max_message_bytes: usize,
    max_fields: usize,
    max_fragments: usize,
    max_text_bytes: usize,
}

impl Limits {
    /// Hard encoded-message ceiling inherited from the common wire layer.
    pub const MAX_MESSAGE_BYTES: usize = WireLimits::MAX_INPUT_BYTES;
    /// Hard root-field ceiling inherited from the common wire layer.
    pub const MAX_FIELDS: usize = WireLimits::MAX_FIELDS;
    /// Hard raw fragment ceiling. Every fragment consumes one root field, so
    /// this is no larger than the common field ceiling.
    pub const MAX_FRAGMENTS: usize = WireLimits::MAX_FIELDS;
    /// Hard aggregate semantic text ceiling.
    pub const MAX_TEXT_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;

    /// Build an explicit finite resource profile.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] if any value is zero or exceeds its
    /// non-bypassable hard ceiling.
    pub fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_fragments: usize,
        max_text_bytes: usize,
    ) -> Result<Self> {
        Ok(Self {
            max_message_bytes: checked_limit(
                "message bytes",
                max_message_bytes,
                Self::MAX_MESSAGE_BYTES,
            )?,
            max_fields: checked_limit("fields", max_fields, Self::MAX_FIELDS)?,
            max_fragments: checked_limit("text fragments", max_fragments, Self::MAX_FRAGMENTS)?,
            max_text_bytes: checked_limit("text bytes", max_text_bytes, Self::MAX_TEXT_BYTES)?,
        })
    }

    /// Maximum encoded bytes accepted for one storage message.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Maximum root wire fields accepted for one storage message.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum repeated field-3 occurrences accepted for one storage message.
    #[must_use]
    pub const fn max_fragments(self) -> usize {
        self.max_fragments
    }

    /// Maximum aggregate UTF-8 bytes accepted for one storage message.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Conservative peak bytes for the borrowed fragment-pointer vector,
    /// including a possible old-plus-new buffer during geometric growth.
    #[must_use]
    pub const fn max_borrowed_element_memory(self) -> usize {
        self.max_fragments * size_of::<&str>() * 2
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_message_bytes: DEFAULT_MAX_MESSAGE_BYTES,
            max_fields: DEFAULT_MAX_FIELDS,
            max_fragments: DEFAULT_MAX_WIRE_FRAGMENTS,
            max_text_bytes: DEFAULT_MAX_TEXT_BYTES,
        }
    }
}

/// Why a native text-storage payload could not become a semantic value.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A caller supplied a zero or over-hard-ceiling resource limit.
    #[error("invalid text wire limit {field}: {value}, expected 1..={maximum}")]
    InvalidLimit {
        /// Limit field.
        field: &'static str,
        /// Supplied value.
        value: usize,
        /// Non-bypassable hard ceiling.
        maximum: usize,
    },
    /// The native payload contains more fragments than the semantic range
    /// representation can retain.
    #[error("text storage contains {actual} fragments; maximum is {limit}")]
    TooManyFragments { actual: usize, limit: usize },
    /// Aggregate decoded UTF-8 exceeds the configured semantic ceiling.
    #[error("text storage contains {actual} UTF-8 bytes; maximum is {limit}")]
    TooManyTextBytes { actual: usize, limit: usize },
    /// The aggregate UTF-8 length cannot be represented by the host address
    /// space.
    #[error("text storage text length overflows the host address space")]
    TextLengthOverflow,
    /// A repeated text occurrence is not valid UTF-8.
    #[error("text storage fragment {fragment} is invalid UTF-8 at byte {valid_up_to}")]
    InvalidUtf8 {
        /// Zero-based field-3 occurrence.
        fragment: usize,
        /// Valid prefix length within that occurrence.
        valid_up_to: usize,
    },
    /// Field 3 used a representation other than its required string wire
    /// type.
    #[error("text storage field 3 has wire type {actual}; expected length-delimited wire type 2")]
    WrongTextWireType {
        /// Observed protobuf wire type.
        actual: u8,
    },
    /// Buffa rejected a payload that had already passed structural preflight.
    #[error("text storage projection decode failed: {reason}")]
    ProjectionDecode {
        /// Runtime-neutral diagnostic text.
        reason: String,
    },
    /// The private projection disagreed with the schema-directed preflight.
    #[error(
        "text storage projection returned {decoded} fragments after preflight counted {preflight}"
    )]
    ProjectionMismatch {
        /// Fragment occurrences counted by common-wire preflight.
        preflight: usize,
        /// Fragment occurrences returned by Buffa.
        decoded: usize,
    },
    /// The final semantic buffer length disagreed with the preflighted UTF-8
    /// length.
    #[error(
        "text storage projection materialized {decoded} bytes after preflight counted {preflight}"
    )]
    ProjectionTextLengthMismatch {
        /// Aggregate bytes counted by common-wire preflight.
        preflight: usize,
        /// Aggregate bytes materialized from the Buffa view.
        decoded: usize,
    },
    /// A fallible text or run allocation failed.
    #[error(transparent)]
    Common(#[from] litchi_iwa_common::Error),
    /// The resulting text/range relation failed semantic validation.
    #[error("semantic text storage is invalid: {0}")]
    Storage(#[source] StorageError),
}

/// Result type for native text-storage conversion.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Preflight {
    fragments: usize,
    text_bytes: usize,
    element_memory: usize,
}

/// Decode raw `TSWP.StorageArchive` bytes through the default bounded Buffa
/// projection and return only archive-free semantic text.
///
/// # Errors
///
/// Returns a typed error for malformed root wire data, invalid UTF-8,
/// exceeded resource limits, projection failure, allocation failure, or an
/// invalid semantic text/range relation.
pub fn from_bytes(source: &[u8]) -> Result<Storage> {
    from_bytes_with_limits(source, Limits::default())
}

/// Decode raw `TSWP.StorageArchive` bytes through an explicit finite resource
/// profile.
///
/// Only repeated UTF-8 field 3 is interpreted. Other native fields are
/// intentionally opaque to this derived projection, including their
/// length-delimited payloads; their root framing is still validated and
/// counted before Buffa runs.
///
/// # Errors
///
/// Returns a typed error for malformed root wire data, invalid UTF-8,
/// exceeded resource limits, projection failure, allocation failure, or an
/// invalid semantic text/range relation.
pub fn from_bytes_with_limits(source: &[u8], limits: Limits) -> Result<Storage> {
    let preflight = preflight_storage(source, limits)?;
    let options = text_storage_codec::DecodeOptions::new(
        limits.max_message_bytes(),
        0,
        preflight.element_memory,
        1,
    );
    let view = text_storage_codec::decode_storage_text(source, options).map_err(|error| {
        Error::ProjectionDecode {
            reason: error.to_string(),
        }
    })?;
    if view.len() != preflight.fragments {
        return Err(Error::ProjectionMismatch {
            preflight: preflight.fragments,
            decoded: view.len(),
        });
    }

    materialize_fragments(view.fragments(), preflight.fragments, preflight.text_bytes)
}

/// Convert one decoded native TSWP storage payload without retaining wire
/// fragments or allocating a second concatenated text buffer.
///
/// This Prost-shaped entry point remains temporarily for compatibility and as
/// a differential oracle while format call sites migrate to [`from_bytes`].
///
/// # Errors
///
/// Returns a typed error when the fragment budget, aggregate text length, or
/// allocation budget is exceeded, or when the semantic storage ranges cannot
/// be validated.
pub fn from_archive(archive: StorageArchive) -> Result<Storage> {
    if archive.text.len() > MAX_FRAGMENTS {
        return Err(Error::TooManyFragments {
            actual: archive.text.len(),
            limit: MAX_FRAGMENTS,
        });
    }

    let text_len = archive.text.iter().try_fold(0usize, |length, fragment| {
        length
            .checked_add(fragment.len())
            .ok_or(Error::TextLengthOverflow)
    })?;

    let fragment_count = archive.text.len();
    materialize_fragments(archive.text.into_iter(), fragment_count, text_len)
}

fn materialize_fragments<Fragment>(
    fragments: impl ExactSizeIterator<Item = Fragment>,
    fragment_count: usize,
    text_len: usize,
) -> Result<Storage>
where
    Fragment: AsRef<str>,
{
    if fragments.len() != fragment_count {
        return Err(Error::ProjectionMismatch {
            preflight: fragment_count,
            decoded: fragments.len(),
        });
    }

    let mut text = String::new();
    text.try_reserve_exact(text_len).map_err(|_allocation| {
        litchi_iwa_common::Error::Allocation {
            resource: "native text storage",
            amount: text_len,
        }
    })?;

    let mut runs = Vec::new();
    runs.try_reserve_exact(fragment_count)
        .map_err(|_allocation| litchi_iwa_common::Error::Allocation {
            resource: "native text storage runs",
            amount: fragment_count,
        })?;

    for fragment_value in fragments {
        let fragment = fragment_value.as_ref();
        let start = text.len();
        let length = fragment.len();
        text.push_str(fragment);
        runs.push(Run::new(start, length));
    }

    if text.len() != text_len {
        return Err(Error::ProjectionTextLengthMismatch {
            preflight: text_len,
            decoded: text.len(),
        });
    }

    Storage::try_from_parts(text, runs).map_err(Error::Storage)
}

fn preflight_storage(source: &[u8], limits: Limits) -> Result<Preflight> {
    let wire_limits = WireLimits::default()
        .with_input_bytes(limits.max_message_bytes())?
        .with_fields(limits.max_fields())?
        .with_nesting(1)?;
    let fragments = Cell::new(0usize);
    let text_bytes = Cell::new(0usize);
    let text_overflow = Cell::new(false);
    let invalid_utf8 = Cell::new(None);
    let wrong_wire_type = Cell::new(None);

    let report = preflight_wire_tree_with_limits(source, wire_limits, |visit| {
        let field = visit.field();
        if visit.path().is_empty() && field.number() == 3 {
            if field.wire_type() != 2 {
                if wrong_wire_type.get().is_none() {
                    wrong_wire_type.set(Some(field.wire_type()));
                }
                return Ok(WireDescent::Skip);
            }
            let fragment = fragments.get();
            fragments.set(fragment.saturating_add(1));
            match text_bytes.get().checked_add(field.payload().len()) {
                Some(total) => text_bytes.set(total),
                None => text_overflow.set(true),
            }
            if invalid_utf8.get().is_none()
                && let Err(error) = std::str::from_utf8(field.payload())
            {
                invalid_utf8.set(Some((fragment, error.valid_up_to())));
            }
        }
        Ok(WireDescent::Skip)
    })?;

    if text_overflow.get() {
        return Err(Error::TextLengthOverflow);
    }
    if let Some(actual) = wrong_wire_type.get() {
        return Err(Error::WrongTextWireType { actual });
    }
    if fragments.get() > limits.max_fragments() {
        return Err(Error::TooManyFragments {
            actual: fragments.get(),
            limit: limits.max_fragments(),
        });
    }
    if text_bytes.get() > limits.max_text_bytes() {
        return Err(Error::TooManyTextBytes {
            actual: text_bytes.get(),
            limit: limits.max_text_bytes(),
        });
    }
    if let Some((fragment, valid_up_to)) = invalid_utf8.get() {
        return Err(Error::InvalidUtf8 {
            fragment,
            valid_up_to,
        });
    }

    // `RepeatedView<&str>` may temporarily hold both old and new buffers while
    // growing. The occurrence limit therefore bounds a conservative 2x peak,
    // while Buffa receives the exact logical element footprint it charges.
    let element_memory = fragments
        .get()
        .checked_mul(size_of::<&str>())
        .ok_or(Error::TextLengthOverflow)?;
    let conservative_element_memory = element_memory
        .checked_mul(2)
        .ok_or(Error::TextLengthOverflow)?;
    if conservative_element_memory > limits.max_borrowed_element_memory() {
        return Err(Error::TooManyFragments {
            actual: fragments.get(),
            limit: limits.max_fragments(),
        });
    }
    debug_assert!(report.fields() <= limits.max_fields());

    Ok(Preflight {
        fragments: fragments.get(),
        text_bytes: text_bytes.get(),
        element_memory,
    })
}

fn checked_limit(field: &'static str, value: usize, maximum: usize) -> Result<usize> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            field,
            value,
            maximum,
        });
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use prost::Message as _;

    fn field(number: u8, payload: &[u8]) -> Vec<u8> {
        assert!(number < 16, "test helper supports one-byte field keys");
        assert!(payload.len() < 128, "test helper supports one-byte lengths");
        let mut encoded = Vec::with_capacity(payload.len() + 2);
        encoded.push((number << 3) | 2);
        encoded.push(
            u8::try_from(payload.len())
                .unwrap_or_else(|error| panic!("test payload length should fit u8: {error}")),
        );
        encoded.extend_from_slice(payload);
        encoded
    }

    fn limits(max_fragments: usize, max_text_bytes: usize) -> Limits {
        Limits::new(1_024, 32, max_fragments, max_text_bytes)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"))
    }

    #[test]
    fn conversion_keeps_fragment_ranges_in_one_owned_text_buffer() {
        let storage = from_archive(StorageArchive {
            text: vec!["Hello".to_owned(), " ".to_owned(), "world".to_owned()],
            ..StorageArchive::default()
        })
        .unwrap_or_else(|error| panic!("valid storage should convert: {error}"));

        assert_eq!(storage.text(), "Hello world");
        assert_eq!(
            storage.runs(),
            [Run::new(0, 5), Run::new(5, 1), Run::new(6, 5)]
        );
        assert_eq!(
            storage
                .fragments()
                .map(litchi_iwa_text::storage::Fragment::text)
                .collect::<Vec<_>>(),
            ["Hello", " ", "world"]
        );
    }

    #[test]
    fn buffa_projection_matches_prost_oracle_for_duplicates_and_unknowns() {
        let mut encoded = StorageArchive {
            text: vec!["first".to_owned(), String::new(), "最後".to_owned()],
            ..StorageArchive::default()
        }
        .encode_to_vec();
        encoded.extend_from_slice(&[0xa0, 0x06, 0x01]);

        let oracle = StorageArchive::decode(encoded.as_slice())
            .map_err(|error| error.to_string())
            .and_then(|archive| from_archive(archive).map_err(|error| error.to_string()))
            .unwrap_or_else(|error| panic!("Prost oracle should decode: {error}"));
        let projected = from_bytes(&encoded)
            .unwrap_or_else(|error| panic!("Buffa projection should decode: {error}"));

        assert_eq!(projected, oracle);
        assert_eq!(projected.text(), "first最後");
        assert_eq!(
            projected.runs(),
            [Run::new(0, 5), Run::new(5, 0), Run::new(5, 6)]
        );
    }

    #[test]
    fn invalid_utf8_is_rejected_before_buffa() {
        let encoded = field(3, &[0xff]);

        assert_eq!(
            from_bytes(&encoded),
            Err(Error::InvalidUtf8 {
                fragment: 0,
                valid_up_to: 0,
            })
        );
        assert!(StorageArchive::decode(encoded.as_slice()).is_err());
    }

    #[test]
    fn raw_fragment_limit_is_checked_before_buffa_metadata_allocation() {
        let encoded = [field(3, &[]), field(3, &[]), field(3, &[])].concat();

        assert_eq!(
            from_bytes_with_limits(&encoded, limits(2, 16)),
            Err(Error::TooManyFragments {
                actual: 3,
                limit: 2,
            })
        );
    }

    #[test]
    fn aggregate_text_limit_is_checked_before_semantic_allocation() {
        let encoded = [field(3, b"abc"), field(3, b"de")].concat();

        let exact = from_bytes_with_limits(&encoded, limits(2, 5))
            .unwrap_or_else(|error| panic!("exact resource boundaries should decode: {error}"));
        assert_eq!(exact.text(), "abcde");
        assert_eq!(
            from_bytes_with_limits(&encoded, limits(4, 4)),
            Err(Error::TooManyTextBytes {
                actual: 5,
                limit: 4,
            })
        );
    }

    #[test]
    fn field_three_requires_the_declared_string_wire_type() {
        assert_eq!(
            from_bytes(&[0x18, 0x01]),
            Err(Error::WrongTextWireType { actual: 0 })
        );
    }

    #[test]
    fn noncanonical_string_length_matches_prost_semantics() {
        let encoded = [0x1a, 0x81, 0x00, b'x'];

        let archive = StorageArchive::decode(encoded.as_slice())
            .unwrap_or_else(|error| panic!("Prost should accept the noncanonical length: {error}"));
        let oracle = from_archive(archive)
            .unwrap_or_else(|error| panic!("Prost oracle should convert: {error}"));
        let projected = from_bytes(&encoded)
            .unwrap_or_else(|error| panic!("projection should accept compatible wire: {error}"));

        assert_eq!(projected, oracle);
        assert_eq!(projected.text(), "x");
    }

    #[test]
    fn malformed_unrelated_submessage_remains_opaque_to_projection() {
        let encoded = [field(3, b"safe"), field(5, &[0x0a])].concat();

        let storage = from_bytes(&encoded)
            .unwrap_or_else(|error| panic!("unrelated payload should stay opaque: {error}"));
        assert_eq!(storage.text(), "safe");
        assert!(
            StorageArchive::decode(encoded.as_slice()).is_err(),
            "the full Prost schema eagerly interprets the malformed field-5 child"
        );
    }

    #[test]
    fn malformed_unknown_root_framing_is_rejected_by_common_preflight() {
        let encoded = [0x2a, 0x02, 0x00];

        assert!(matches!(from_bytes(&encoded), Err(Error::Common(_))));
    }

    #[test]
    fn invalid_limit_is_typed_and_cannot_disable_the_hard_profile() {
        assert_eq!(
            Limits::new(0, 1, 1, 1),
            Err(Error::InvalidLimit {
                field: "message bytes",
                value: 0,
                maximum: Limits::MAX_MESSAGE_BYTES,
            })
        );
    }

    #[test]
    fn empty_native_storage_has_no_run_allocation() {
        let storage = from_archive(StorageArchive::default())
            .unwrap_or_else(|error| panic!("empty storage should convert: {error}"));

        assert!(storage.is_empty());
        assert!(storage.runs().is_empty());
    }

    #[test]
    fn fragment_limit_is_checked_before_materialization() {
        let archive = StorageArchive {
            text: vec![String::new(); MAX_FRAGMENTS + 1],
            ..StorageArchive::default()
        };

        assert_eq!(
            from_archive(archive),
            Err(Error::TooManyFragments {
                actual: MAX_FRAGMENTS + 1,
                limit: MAX_FRAGMENTS,
            })
        );
    }
}
