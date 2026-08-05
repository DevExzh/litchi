//! Borrowed semantic inventory types for embedded relationships.

use litchi_opc::{PackURI, Part};
use std::fmt;

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
    /// Return the safe general-purpose limits.
    #[inline]
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            relationships: DEFAULT_RELATIONSHIPS,
            payload_relationships: DEFAULT_PAYLOAD_RELATIONSHIPS,
        }
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

/// Internal payload bytes or an inert external relationship target.
#[derive(Debug, Clone, Copy)]
pub enum Target<'a> {
    /// A package-owned payload exposed as a borrowed byte view.
    Internal(Payload<'a>),
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
