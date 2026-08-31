//! Borrowed semantic inventory types for embedded relationships.

use crate::{Error, Result};
use litchi_opc::{PackURI, Part, PartData, PartView};
use std::fmt;
use std::io::Write;

const DEFAULT_RELATIONSHIPS: usize = 1_024;
const DEFAULT_PAYLOAD_RELATIONSHIPS: usize = 1_024;

/// Normative OOXML embedded-part relationship family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// An Embedded Object Part (ISO/IEC 29500-1 section 15.2.10).
    Object,
    /// An Embedded Package Part (ISO/IEC 29500-1 section 15.2.11).
    Package,
}

/// Resource budgets for an embedded-part inventory.
///
/// [`Limits::default`] is the safe general-purpose policy. Callers that know
/// their workload may explicitly tighten or loosen either independent budget
/// with [`scan_with`](super::scan_with). Payload relationships are charged
/// once per unique internal target, even when several source relationships
/// reference it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum embedded-object or embedded-package relationship occurrences.
    pub relationships: usize,
    /// Maximum aggregate relationships on unique internal payload parts.
    pub payload_relationships: usize,
}

impl Limits {
    /// Absolute ceiling for embedded relationship occurrences in one scan.
    pub const MAX_RELATIONSHIPS: usize = 1_000_000;
    /// Absolute ceiling for relationships on embedded payload parts in one scan.
    pub const MAX_PAYLOAD_RELATIONSHIPS: usize = 1_000_000;

    /// Return the safe general-purpose limits.
    #[inline]
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            relationships: DEFAULT_RELATIONSHIPS,
            payload_relationships: DEFAULT_PAYLOAD_RELATIONSHIPS,
        }
    }

    /// Validate this budget against the immutable embedded-graph ceilings.
    ///
    /// The public fields remain available for callers that need a tightly
    /// scoped profile. A zero budget is meaningful and continues to reject
    /// every matching relationship; only values above the finite ceilings are
    /// invalid.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when either configured budget exceeds its
    /// format-wide ceiling.
    #[inline]
    pub const fn validate(self) -> Result<Self> {
        if self.relationships > Self::MAX_RELATIONSHIPS {
            return Err(Error::Limit {
                resource: "embedded relationships",
                max: Self::MAX_RELATIONSHIPS,
                actual: self.relationships,
            });
        }
        if self.payload_relationships > Self::MAX_PAYLOAD_RELATIONSHIPS {
            return Err(Error::Limit {
                resource: "embedded payload relationships",
                max: Self::MAX_PAYLOAD_RELATIONSHIPS,
                actual: self.payload_relationships,
            });
        }
        Ok(self)
    }
}

impl Default for Limits {
    #[inline]
    fn default() -> Self {
        Self::standard()
    }
}

/// A borrowed internal embedded payload.
///
/// The view retains no independent byte allocation and cannot outlive its OPC
/// package.
#[derive(Clone, Copy)]
pub struct Payload<'a> {
    pub(super) part: &'a dyn Part,
}

impl<'a> Payload<'a> {
    /// Absolute OPC part name of the payload.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a PackURI {
        self.part.partname()
    }

    /// Declared OPC content type, without format sniffing.
    #[inline]
    #[must_use]
    pub fn content_type(&self) -> &'a str {
        self.part.content_type()
    }

    /// Original payload bytes held by the OPC package.
    #[inline]
    #[must_use]
    pub fn bytes(&self) -> &'a [u8] {
        self.part.blob()
    }
}

impl fmt::Debug for Payload<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Payload")
            .field("part", self.part())
            .field("content_type", &self.content_type())
            .field("byte_len", &self.bytes().len())
            .finish()
    }
}

/// A borrowed source-backed embedded payload.
///
/// The view retains only metadata and a position in its caller-owned OPC
/// source. Its bytes remain deferred until [`Self::data`] or [`Self::stream_to`]
/// is requested.
pub struct SourcePayload<'a> {
    pub(super) part: PartView<'a>,
}

impl SourcePayload<'_> {
    /// Absolute OPC part name of the payload.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &PackURI {
        self.part.partname()
    }

    /// Declared OPC content type, without format sniffing.
    #[inline]
    #[must_use]
    pub fn content_type(&self) -> &str {
        self.part.content_type()
    }

    /// ZIP central-directory size declared for this payload.
    ///
    /// This reads metadata only. The returned value is not a proof of the
    /// decoded length; [`Self::data`] and [`Self::stream_to`] perform the
    /// complete bounded source validation.
    ///
    /// # Errors
    ///
    /// Returns an OPC error when the source is stale, cancelled, or its
    /// central-directory metadata cannot be read safely.
    pub fn declared_uncompressed_size(&self) -> Result<u64> {
        Ok(self.part.declared_uncompressed_size()?)
    }

    /// Read the payload through the source-backed OPC budget.
    ///
    /// The returned wrapper intentionally exposes only a byte slice and its
    /// size predicates. It cannot detach the managed OPC allocation.
    ///
    /// # Errors
    ///
    /// Returns an OPC error when the payload cannot be read under the source
    /// package's limits or execution context.
    pub fn data(&self) -> Result<SourcePayloadData> {
        Ok(SourcePayloadData {
            data: self.part.data()?,
        })
    }

    /// Stream the decoded payload to `sink` without materializing it here.
    ///
    /// # Errors
    ///
    /// Returns an OPC or sink error when the source, execution context,
    /// archive, or output fails.
    pub fn stream_to<W: Write>(&self, sink: &mut W) -> Result<u64> {
        Ok(self.part.stream_to(sink)?)
    }
}

impl fmt::Debug for SourcePayload<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePayload")
            .field("part", self.part())
            .field("content_type", &self.content_type())
            .finish()
    }
}

/// A bounded source-backed payload allocation.
///
/// The underlying OPC [`PartData`] handle is deliberately private. In
/// particular, this type has no conversion to `Arc`, no `into_arc` helper,
/// and does not expose the low-level handle to callers.
#[derive(Clone)]
pub struct SourcePayloadData {
    data: PartData,
}

impl SourcePayloadData {
    /// Borrow the decoded payload bytes.
    #[inline]
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.data.as_bytes()
    }

    /// Return the decoded payload length.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.as_bytes().len()
    }

    /// Return whether the decoded payload is empty.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.as_bytes().is_empty()
    }
}

impl fmt::Debug for SourcePayloadData {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePayloadData")
            .field("byte_len", &self.len())
            .finish()
    }
}

/// Internal payload bytes or an inert external relationship target.
#[derive(Debug, Clone, Copy)]
pub enum Target<'a> {
    /// A package-owned payload exposed as a borrowed byte view.
    Internal(Payload<'a>),
    /// An external target retained verbatim and never contacted.
    External(&'a str),
}

/// One source-backed relationship occurrence referencing an embedded part.
#[derive(Debug)]
pub struct SourceEntry<'a> {
    pub(super) source: &'a PackURI,
    pub(super) id: &'a str,
    pub(super) kind: Kind,
    pub(super) target: SourceTarget<'a>,
}

impl<'a> SourceEntry<'a> {
    /// Source part that owns the relationship.
    #[inline]
    #[must_use]
    pub fn source(&self) -> &'a PackURI {
        self.source
    }

    /// Relationship identifier within [`Self::source`].
    #[inline]
    #[must_use]
    pub fn id(&self) -> &'a str {
        self.id
    }

    /// Embedded relationship family.
    #[inline]
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Borrowed source-backed payload or inert external relationship target.
    #[inline]
    #[must_use]
    pub fn target(&self) -> &SourceTarget<'a> {
        &self.target
    }
}

/// A source-backed internal payload or an inert external relationship target.
#[derive(Debug)]
pub enum SourceTarget<'a> {
    /// A package-owned payload exposed as a deferred source view.
    Internal(SourcePayload<'a>),
    /// An external target retained verbatim and never contacted.
    External(&'a str),
}

/// One explicit relationship occurrence referencing an embedded part.
#[derive(Debug, Clone, Copy)]
pub struct Entry<'a> {
    pub(super) source: &'a PackURI,
    pub(super) id: &'a str,
    pub(super) kind: Kind,
    pub(super) target: Target<'a>,
}

impl<'a> Entry<'a> {
    /// Source part that owns the relationship.
    #[inline]
    #[must_use]
    pub fn source(&self) -> &'a PackURI {
        self.source
    }

    /// Relationship identifier within [`Self::source`].
    #[inline]
    #[must_use]
    pub fn id(&self) -> &'a str {
        self.id
    }

    /// Embedded relationship family.
    #[inline]
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Borrowed internal payload or inert external target.
    #[inline]
    #[must_use]
    pub fn target(&self) -> Target<'a> {
        self.target
    }
}
