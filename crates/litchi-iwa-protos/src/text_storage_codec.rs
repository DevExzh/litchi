//! Private-type Buffa codec for the TSWP text-storage projection.
//!
//! The generated projection contains only repeated UTF-8 field 3 from
//! `TSWP.StorageArchive`. Callers must complete their schema-directed wire
//! preflight and establish finite limits before entering this module.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_text_storage_generated::LitchiIwaProjection as projection;

/// Finite limits already established by the text wire adapter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_unknown_fields: usize,
    max_element_memory: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted storage payload.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_unknown_fields: usize,
        max_element_memory: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_unknown_fields,
            max_element_memory,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_unknown_fields)
            .with_element_memory_limit(self.max_element_memory)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Failure from the private Buffa projection decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError(buffa::DecodeError);

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self(error)
    }
}

/// Borrowed repeated text projection with no generated type in its public
/// surface.
#[derive(Debug)]
pub struct StorageTextView<'source> {
    view: projection::TSWPStorageArchiveLazyView<'source>,
}

impl<'source> StorageTextView<'source> {
    /// Number of field-3 occurrences in source order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.view.text.len()
    }

    /// Whether the source contains no field-3 occurrences.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.view.text.is_empty()
    }

    /// Borrow every UTF-8 text fragment in source order.
    #[must_use]
    pub fn fragments(&self) -> impl ExactSizeIterator<Item = &'source str> + '_ {
        self.view.text.iter().copied()
    }
}

/// Decode one already-preflighted `TSWP.StorageArchive` text projection.
///
/// Unknown fields remain opaque and are not exposed. The returned text
/// fragments borrow the original input, while the generated Buffa view stays
/// private to this crate.
pub fn decode_storage_text(
    source: &[u8],
    options: DecodeOptions,
) -> Result<StorageTextView<'_>, DecodeError> {
    let view = options.buffa().decode_lazy_view(source)?;
    Ok(StorageTextView { view })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tswp;
    use prost::Message as _;
    use std::mem::size_of;

    fn string_field(payload: &[u8]) -> Vec<u8> {
        length_field(3, payload)
    }

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 0);
        push_varint(&mut output, value);
        output
    }

    fn fixed32_field(field: u32, value: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 5);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn fixed64_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 1);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 2);
        push_varint(
            &mut output,
            u64::try_from(payload.len()).expect("fixture length fits u64"),
        );
        output.extend_from_slice(payload);
        output
    }

    fn start_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 3);
        output
    }

    fn end_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 4);
        output
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).expect("varint chunk fits u8");
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn options(
        source: &[u8],
        max_unknown_fields: usize,
        max_element_memory: usize,
        recursion_limit: u32,
    ) -> DecodeOptions {
        DecodeOptions::new(
            source.len(),
            max_unknown_fields,
            max_element_memory,
            recursion_limit,
        )
    }

    fn canonical_source(fragments: &[&str]) -> Vec<u8> {
        tswp::StorageArchive {
            text: fragments
                .iter()
                .map(|fragment| (*fragment).to_owned())
                .collect(),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn assert_borrowed(source: &[u8], fragment: &str) {
        if fragment.is_empty() {
            return;
        }
        let source_start = source.as_ptr() as usize;
        let source_end = source_start.saturating_add(source.len());
        let fragment_start = fragment.as_ptr() as usize;
        let fragment_end = fragment_start.saturating_add(fragment.len());
        assert!(fragment_start >= source_start);
        assert!(fragment_end <= source_end);
        let offset = fragment_start - source_start;
        assert_eq!(
            &source[offset..offset + fragment.len()],
            fragment.as_bytes()
        );
    }

    #[test]
    fn repeated_unicode_and_empty_fragments_match_prost_and_borrow_source() {
        let expected_fragments = ["", "é", "東京", "👩‍🚀", "", "e\u{301}"];
        let source = canonical_source(&expected_fragments);
        let native = tswp::StorageArchive::decode(source.as_slice()).expect("native decode");
        let view = decode_storage_text(
            &source,
            options(&source, 0, expected_fragments.len() * size_of::<&str>(), 1),
        )
        .expect("Buffa projection");
        let projected = view.fragments().collect::<Vec<_>>();
        let expected = native.text.iter().map(String::as_str).collect::<Vec<_>>();

        assert_eq!(projected, expected);
        assert_eq!(view.len(), expected_fragments.len());
        assert!(!view.is_empty());
        for fragment in projected {
            assert_borrowed(&source, fragment);
        }
    }

    #[test]
    fn unknown_wire_kinds_and_balanced_groups_are_opaque_with_zero_budget() {
        let mut source = string_field(b"visible");
        source.extend(varint_field(99, 0xfeed));
        source.extend(fixed32_field(100, 0xfeed_face));
        source.extend(fixed64_field(101, 0xfeed_face_cafe_beef));
        source.extend(length_field(102, b"opaque"));
        source.extend(start_group(103));
        source.extend(varint_field(104, 7));
        source.extend(fixed32_field(105, 0xdecafbad));
        source.extend(start_group(106));
        source.extend(length_field(107, b"nested"));
        source.extend(end_group(106));
        source.extend(end_group(103));

        let view = decode_storage_text(&source, options(&source, 0, size_of::<&str>(), 3))
            .expect("unknown fields are opaque to the lazy projection");
        assert_eq!(view.fragments().collect::<Vec<_>>(), ["visible"]);
    }

    #[test]
    fn malformed_unknown_groups_are_rejected() {
        let cases = [
            (
                [start_group(99), end_group(100)].concat(),
                buffa::DecodeError::InvalidEndGroup(100),
            ),
            (
                [start_group(99), varint_field(100, 7)].concat(),
                buffa::DecodeError::UnexpectedEof,
            ),
            (end_group(99), buffa::DecodeError::InvalidWireType(4)),
        ];

        for (source, expected) in cases {
            let error = decode_storage_text(&source, options(&source, 0, 0, 1))
                .expect_err("malformed group");
            assert_eq!(error.0, expected);
        }
    }

    #[test]
    fn invalid_utf8_is_rejected_by_the_string_projection() {
        let source = string_field(&[0xff, 0x80]);
        let error =
            decode_storage_text(&source, options(&source, 0, 0, 1)).expect_err("invalid UTF-8");
        assert_eq!(error.0, buffa::DecodeError::InvalidUtf8);
    }

    #[test]
    fn exact_and_one_under_message_byte_limits_are_enforced() {
        let source = string_field(b"boundary");
        let exact = decode_storage_text(&source, options(&source, 0, size_of::<&str>(), 1))
            .expect("exact byte boundary");
        assert_eq!(exact.fragments().collect::<Vec<_>>(), ["boundary"]);

        let error = decode_storage_text(
            &source,
            DecodeOptions::new(source.len() - 1, 0, size_of::<&str>(), 1),
        )
        .expect_err("one byte under");
        assert_eq!(error.0, buffa::DecodeError::MessageTooLarge);
    }

    #[test]
    fn exact_and_one_under_element_memory_limits_are_enforced() {
        let source = string_field(b"element");
        let footprint = size_of::<&str>();
        assert!(footprint > 0);
        assert!(decode_storage_text(&source, options(&source, 0, footprint, 1)).is_ok());

        let error = decode_storage_text(&source, options(&source, 0, footprint - 1, 1))
            .expect_err("one element byte under");
        assert_eq!(error.0, buffa::DecodeError::ElementMemoryLimitExceeded);
    }

    #[test]
    fn exact_and_one_under_recursion_limits_are_enforced_for_groups() {
        let mut source = string_field(b"visible");
        source.extend(start_group(99));
        source.extend(varint_field(100, 7));
        source.extend(end_group(99));

        assert!(decode_storage_text(&source, options(&source, 0, size_of::<&str>(), 1),).is_ok());
        let error = decode_storage_text(&source, options(&source, 0, size_of::<&str>(), 0))
            .expect_err("one nesting level under");
        assert_eq!(error.0, buffa::DecodeError::RecursionLimitExceeded);
    }
}
