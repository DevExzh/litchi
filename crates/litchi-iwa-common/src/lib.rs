//! Format-independent primitives shared by Litchi's iWork crates.
//!
//! This crate deliberately contains no Pages, Numbers, or Keynote object-model
//! knowledge.  Concrete iWork crates may depend on it, but it must not depend
//! on a concrete document crate or on the umbrella facade.

#![forbid(unsafe_code)]

pub mod color;
pub mod table;
pub mod varint;
pub mod wire;

pub use varint::{
    DecodeError, MAX_BYTES, decode_svarint, decode_varint, decode_varint_from_bytes,
    encode_svarint, encode_svarint_into, encode_svarint_to_buffer, encode_varint,
    encode_varint_into, encode_varint_to_buffer,
};

/// The resource counted by a bounded wire operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// Bytes in the input message.
    InputBytes,
    /// Parsed protobuf fields.
    Fields,
    /// Bytes in a rewritten output message.
    OutputBytes,
    /// Nested length-delimited traversal depth.
    Nesting,
    /// Aggregate rewrite visits or replacement work.
    RewriteWork,
    /// Addressable table rows.
    TableRows,
    /// Addressable table columns.
    TableColumns,
    /// Addressable table cells.
    TableCells,
    /// Materialized sparse table cells.
    MaterializedCells,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Self::InputBytes => "input bytes",
            Self::Fields => "fields",
            Self::OutputBytes => "output bytes",
            Self::Nesting => "nesting depth",
            Self::RewriteWork => "rewrite work",
            Self::TableRows => "table rows",
            Self::TableColumns => "table columns",
            Self::TableCells => "table cells",
            Self::MaterializedCells => "materialized cells",
        };
        formatter.write_str(name)
    }
}

/// Errors produced while validating shared IWA wire data.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The input or requested mutation is not a valid bounded protobuf wire
    /// representation.
    #[error("invalid IWA wire data: {0}")]
    InvalidFormat(String),
    /// A configured finite resource budget was exceeded.
    #[error("IWA wire {kind} limit exceeded: observed {observed}, limit {limit}")]
    LimitExceeded {
        /// Resource that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// A fallible collection allocation could not be completed.
    #[error("IWA wire allocation failed for {resource}: {amount}")]
    Allocation {
        /// Collection or buffer being allocated.
        resource: &'static str,
        /// Number of elements or bytes requested.
        amount: usize,
    },
    /// A caller supplied an invalid limit value.
    #[error("invalid IWA wire limit {field}: {value}, expected 1..={maximum}")]
    InvalidLimit {
        /// Limit field name.
        field: &'static str,
        /// Supplied value.
        value: usize,
        /// Hard maximum.
        maximum: usize,
    },
}

/// Result type for shared IWA wire operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Finite budgets for parsing and mutating one protobuf-style IWA payload.
///
/// The hard ceilings prevent a caller from turning a configurable limit into
/// an unbounded allocation or traversal. Lower limits are useful for document
/// profiles and adversarial-input tests.
#[allow(
    clippy::struct_field_names,
    reason = "Each budget has a distinct public accessor and must remain independently configurable"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WireLimits {
    max_input_bytes: usize,
    max_fields: usize,
    max_output_bytes: usize,
    max_nesting: usize,
    max_rewrite_work: usize,
}

impl WireLimits {
    /// Absolute input-byte ceiling.
    pub const MAX_INPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Absolute parsed-field ceiling.
    pub const MAX_FIELDS: usize = 1_000_000;
    /// Absolute rewritten-output ceiling.
    pub const MAX_OUTPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Absolute nested traversal ceiling.
    pub const MAX_NESTING: usize = 64;
    /// Absolute aggregate rewrite-work ceiling.
    pub const MAX_REWRITE_WORK: usize = 16_000_000;

    /// Tighten the input-byte budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when `value` is zero or exceeds the
    /// crate's hard input-byte ceiling.
    pub fn with_input_bytes(mut self, value: usize) -> Result<Self> {
        self.max_input_bytes = checked_limit("input bytes", value, Self::MAX_INPUT_BYTES)?;
        Ok(self)
    }

    /// Tighten the parsed-field budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when `value` is zero or exceeds the
    /// crate's hard field-count ceiling.
    pub fn with_fields(mut self, value: usize) -> Result<Self> {
        self.max_fields = checked_limit("fields", value, Self::MAX_FIELDS)?;
        Ok(self)
    }

    /// Tighten the rewritten-output budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when `value` is zero or exceeds the
    /// crate's hard output-byte ceiling.
    pub fn with_output_bytes(mut self, value: usize) -> Result<Self> {
        self.max_output_bytes = checked_limit("output bytes", value, Self::MAX_OUTPUT_BYTES)?;
        Ok(self)
    }

    /// Tighten the nested traversal budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when `value` is zero or exceeds the
    /// crate's hard nesting-depth ceiling.
    pub fn with_nesting(mut self, value: usize) -> Result<Self> {
        self.max_nesting = checked_limit("nesting", value, Self::MAX_NESTING)?;
        Ok(self)
    }

    /// Tighten the aggregate rewrite-work budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when `value` is zero or exceeds the
    /// crate's hard rewrite-work ceiling.
    pub fn with_rewrite_work(mut self, value: usize) -> Result<Self> {
        self.max_rewrite_work = checked_limit("rewrite work", value, Self::MAX_REWRITE_WORK)?;
        Ok(self)
    }

    /// Maximum input bytes accepted by this profile.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum parsed fields accepted by this profile.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum output bytes accepted by this profile.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum nested traversal depth accepted by this profile.
    #[must_use]
    pub const fn max_nesting(self) -> usize {
        self.max_nesting
    }

    /// Maximum aggregate rewrite work accepted by this profile.
    #[must_use]
    pub const fn max_rewrite_work(self) -> usize {
        self.max_rewrite_work
    }
}

impl Default for WireLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
            max_fields: 100_000,
            max_output_bytes: Self::MAX_OUTPUT_BYTES,
            max_nesting: Self::MAX_NESTING,
            max_rewrite_work: Self::MAX_REWRITE_WORK,
        }
    }
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
