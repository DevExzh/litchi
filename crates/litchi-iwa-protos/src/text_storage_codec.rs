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
