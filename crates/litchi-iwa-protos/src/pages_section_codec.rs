//! Private-type Buffa projection for Pages section pagination.
//!
//! The generated projection contains only `TP.SectionArchive` fields 20--22.
//! Callers retain and rewrite the original payload; this module never owns or
//! re-encodes unrelated section fields.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_section_generated::LitchiIwaProjection as projection;

/// Finite limits established by the caller's strict wire preflight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted section payload.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrow-free scalar result of the private Pages projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PaginationSnapshot {
    /// Optional native section-start discriminant.
    pub section_start_kind: Option<u32>,
    /// Optional native page-numbering discriminant.
    pub section_page_number_kind: Option<u32>,
    /// Optional native first page number.
    pub section_page_number_start: Option<u32>,
}

/// Failure from the private Pages section projection decoder.
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

/// Decode the three pagination scalars from an already-preflighted payload.
///
/// The generated lazy view borrows `source`, retains no repeated-element or
/// unknown-field storage, and is dropped before this borrow-free result is
/// returned.
pub fn decode_pagination(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PaginationSnapshot, DecodeError> {
    let view: projection::PagesSectionPaginationArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    Ok(PaginationSnapshot {
        section_start_kind: view.section_start_kind,
        section_page_number_kind: view.section_page_number_kind,
        section_page_number_start: view.section_page_number_start,
    })
}

#[cfg(test)]
mod tests {
    use prost::Message as _;

    use super::{DecodeOptions, PaginationSnapshot, decode_pagination};

    fn decode(source: &[u8]) -> Result<PaginationSnapshot, super::DecodeError> {
        decode_pagination(source, DecodeOptions::new(source.len(), 1))
    }

    #[test]
    fn canonical_prost_section_matches_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = crate::tp::SectionArchive {
            section_start_kind: Some(2),
            section_page_number_kind: Some(1),
            section_page_number_start: Some(42),
            name: Some("opaque to this projection".to_owned()),
            ..crate::tp::SectionArchive::default()
        }
        .encode_to_vec();
        assert_eq!(
            decode(&source)?,
            PaginationSnapshot {
                section_start_kind: Some(2),
                section_page_number_kind: Some(1),
                section_page_number_start: Some(42),
            }
        );
        Ok(())
    }

    #[test]
    fn absent_scalars_and_opaque_unknown_payload_remain_allocation_free()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [0xd2, 0x0c, 0x03, 0xff, 0x00, 0xfe];
        assert_eq!(decode(&source)?, PaginationSnapshot::default());
        Ok(())
    }

    #[test]
    fn malformed_selected_scalar_is_rejected() {
        let Err(error) = decode(&[0xa2, 0x01, 0x01, 0x00]) else {
            panic!("field 20 with a length-delimited wire type must fail");
        };
        assert!(!error.to_string().is_empty());
    }
}
