//! Caller-selected budgets for raw Numbers-wire operations.
//!
//! The wire crate is deliberately usable below a package owner, so it never
//! chooses a policy at a raw boundary.  Every parser, projection, and
//! variable-size writer receives one of these checked profiles.  Owners may
//! tighten the hard ceilings for their retained source and semantic budgets;
//! the profile is copied by value only after the owner has selected it.

use std::fmt;

use litchi_iwa_common::WireLimits;

/// A resource governed by one Numbers-wire operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in one source cell or wire message.
    InputBytes,
    /// Bytes in one encoded cell or rewritten wire message.
    OutputBytes,
    /// Parsed wire fields in one message.
    Fields,
    /// Opaque bytes after the known fixed-width cell fields.
    TailBytes,
    /// Repeated cells, rows, entries, or other materialized items.
    Items,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Fields => "fields",
            Self::TailBytes => "tail bytes",
            Self::Items => "items",
        })
    }
}

/// Checked budgets for one raw Numbers-wire operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    wire: WireLimits,
    max_tail_bytes: usize,
    max_items: usize,
}

impl Limits {
    /// Absolute input-byte ceiling.
    pub const MAX_INPUT_BYTES: usize = WireLimits::MAX_INPUT_BYTES;
    /// Absolute output-byte ceiling.
    pub const MAX_OUTPUT_BYTES: usize = WireLimits::MAX_OUTPUT_BYTES;
    /// Absolute parsed-field ceiling.
    pub const MAX_FIELDS: usize = WireLimits::MAX_FIELDS;
    /// Absolute opaque-tail ceiling.
    pub const MAX_TAIL_BYTES: usize = WireLimits::MAX_INPUT_BYTES;
    /// Absolute repeated-item ceiling.
    pub const MAX_ITEMS: usize = WireLimits::MAX_REWRITE_WORK;

    /// The most permissive bounded raw profile.
    pub const DEFAULT: Self = Self {
        wire: WireLimits::DEFAULT,
        max_tail_bytes: Self::MAX_TAIL_BYTES,
        max_items: Self::MAX_ITEMS,
    };

    /// Create a Numbers-wire profile from an already checked common IWA
    /// wire profile.  Tail and item ceilings remain at this module's hard
    /// ceilings until the owner tightens them explicitly.
    #[must_use]
    pub const fn from_wire_limits(wire: WireLimits) -> Self {
        Self {
            wire,
            max_tail_bytes: Self::MAX_TAIL_BYTES,
            max_items: Self::MAX_ITEMS,
        }
    }

    /// Tighten the input-byte ceiling.
    pub fn with_input_bytes(mut self, value: usize) -> Result<Self, Error> {
        let value = checked(value, Self::MAX_INPUT_BYTES, LimitKind::InputBytes)?;
        self.wire = self
            .wire
            .with_input_bytes(value)
            .map_err(|_| Error::InvalidLimit {
                kind: LimitKind::InputBytes,
                value,
                maximum: Self::MAX_INPUT_BYTES,
            })?;
        Ok(self)
    }

    /// Tighten the output-byte ceiling.
    pub fn with_output_bytes(mut self, value: usize) -> Result<Self, Error> {
        let value = checked(value, Self::MAX_OUTPUT_BYTES, LimitKind::OutputBytes)?;
        self.wire = self
            .wire
            .with_output_bytes(value)
            .map_err(|_| Error::InvalidLimit {
                kind: LimitKind::OutputBytes,
                value,
                maximum: Self::MAX_OUTPUT_BYTES,
            })?;
        Ok(self)
    }

    /// Tighten the parsed-field ceiling.
    pub fn with_fields(mut self, value: usize) -> Result<Self, Error> {
        let value = checked(value, Self::MAX_FIELDS, LimitKind::Fields)?;
        self.wire = self
            .wire
            .with_fields(value)
            .map_err(|_| Error::InvalidLimit {
                kind: LimitKind::Fields,
                value,
                maximum: Self::MAX_FIELDS,
            })?;
        Ok(self)
    }

    /// Tighten the opaque-tail ceiling.
    pub fn with_tail_bytes(mut self, value: usize) -> Result<Self, Error> {
        self.max_tail_bytes = checked(value, Self::MAX_TAIL_BYTES, LimitKind::TailBytes)?;
        Ok(self)
    }

    /// Tighten the repeated-item ceiling.
    pub fn with_items(mut self, value: usize) -> Result<Self, Error> {
        self.max_items = checked(value, Self::MAX_ITEMS, LimitKind::Items)?;
        Ok(self)
    }

    /// Maximum input bytes accepted by this profile.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.wire.max_input_bytes()
    }

    /// Maximum output bytes accepted by this profile.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.wire.max_output_bytes()
    }

    /// Maximum parsed fields accepted by this profile.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.wire.max_fields()
    }

    /// Maximum opaque tail bytes accepted by this profile.
    #[must_use]
    pub const fn max_tail_bytes(self) -> usize {
        self.max_tail_bytes
    }

    /// Maximum repeated items accepted by this profile.
    #[must_use]
    pub const fn max_items(self) -> usize {
        self.max_items
    }

    /// Return the common IWA wire profile selected by this policy.
    #[must_use]
    pub const fn wire(self) -> WireLimits {
        self.wire
    }

    /// Check a source byte count before parsing or copying it.
    pub(crate) fn check_input(self, observed: usize) -> Result<(), Error> {
        check(observed, self.max_input_bytes(), LimitKind::InputBytes)
    }

    /// Check an encoded output byte count before allocating it.
    pub(crate) fn check_output(self, observed: usize) -> Result<(), Error> {
        check(observed, self.max_output_bytes(), LimitKind::OutputBytes)
    }

    /// Check an opaque tail before retaining it.
    pub(crate) fn check_tail(self, observed: usize) -> Result<(), Error> {
        check(observed, self.max_tail_bytes(), LimitKind::TailBytes)
    }

    /// Check a repeated item count before materializing it.
    pub(crate) fn check_items(self, observed: usize) -> Result<(), Error> {
        check(observed, self.max_items(), LimitKind::Items)
    }

    /// Check a parsed-field count before materializing field views.
    pub(crate) fn check_fields(self, observed: usize) -> Result<(), Error> {
        check(observed, self.max_fields(), LimitKind::Fields)
    }

    /// Admit an operation's generated candidate for readback validation.
    ///
    /// A source-backed owner may deliberately choose an input ceiling equal
    /// to the retained source while allowing a larger, separately checked
    /// output.  Candidate validation is over bytes just produced by this
    /// operation, not a new untrusted source ingress, so it may widen only
    /// the temporary input check to the already bounded output length.
    pub(crate) fn for_generated_output(self, output_len: usize) -> Result<Self, Error> {
        if output_len <= self.max_input_bytes() {
            return Ok(self);
        }
        let wire = self
            .wire
            .with_input_bytes(output_len)
            .map_err(|_| Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: output_len,
                maximum: Self::MAX_INPUT_BYTES,
            })?;
        Ok(Self {
            wire,
            max_tail_bytes: self.max_tail_bytes,
            max_items: self.max_items,
        })
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Failure while constructing or applying a raw Numbers-wire profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A selected limit is outside the module's finite hard ceiling.
    InvalidLimit {
        /// Dimension being configured.
        kind: LimitKind,
        /// Requested value.
        value: usize,
        /// Hard maximum.
        maximum: usize,
    },
    /// A source or requested result exceeds its selected ceiling.
    LimitExceeded {
        /// Dimension that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Selected maximum.
        maximum: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid Numbers-wire {kind} limit: {value}, maximum {maximum}"
            ),
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "Numbers-wire {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
        }
    }
}

impl std::error::Error for Error {}

fn checked(value: usize, maximum: usize, kind: LimitKind) -> Result<usize, Error> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(value)
}

fn check(observed: usize, maximum: usize, kind: LimitKind) -> Result<(), Error> {
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}
