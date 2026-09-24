//! Opaque, bounded Calculation Chain source data.

use std::sync::Arc;

use crate::package::error::{Error, Result};
use crate::raw::{Limits as RawLimits, Records};

/// Largest individual Workbook or Calculation Chain stream accepted by this
/// owner.
pub const HARD_MAX_PART_BYTES: usize = 64 * 1024 * 1024;
/// Largest number of generic BIFF12 records reported by one explicit
/// Calculation Chain diagnostic.
pub const HARD_MAX_RECORDS: usize = 1_000_000;
/// Largest number of relationships scanned while proving the owner graph.
pub const HARD_MAX_RELATIONSHIPS: usize = 1_000_000;
/// Largest number of physical package parts visited by one owner graph pass.
pub const HARD_MAX_GRAPH_PARTS: usize = 1_000_000;

/// Caller-controlled finite limits for Calculation Chain inspection and
/// source-bound removal.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadLimits {
    /// Maximum encoded bytes in the Workbook or Calculation Chain stream and
    /// in each captured relationship/content-types source token.
    pub max_part_bytes: usize,
    /// Maximum generic BIFF12 records reported by [`ChainPart::record_count`].
    /// This is a diagnostic bound only; Calculation Chain admission never
    /// requires the opaque payload to use BIFF12 framing.
    pub max_records: usize,
    /// Maximum relationships visited during one bounded package graph scan.
    /// Source checks and publication readbacks each perform their own bounded
    /// inventory; no unbounded relationship collection is retained.
    pub max_relationships: usize,
    /// Maximum physical package parts visited during one graph inventory pass.
    pub max_graph_parts: usize,
}

impl ReadLimits {
    /// Conservative finite limits for ordinary Calculation Chain inspection.
    pub const DEFAULT: Self = Self {
        max_part_bytes: 16 * 1024 * 1024,
        max_records: 262_144,
        max_relationships: 65_536,
        max_graph_parts: 65_536,
    };

    /// Validate caller limits against implementation hard ceilings.
    pub const fn validate(self) -> Result<Self> {
        if self.max_part_bytes > HARD_MAX_PART_BYTES {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain part bytes",
                actual: self.max_part_bytes,
                maximum: HARD_MAX_PART_BYTES,
            });
        }
        if self.max_records > HARD_MAX_RECORDS {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain records",
                actual: self.max_records,
                maximum: HARD_MAX_RECORDS,
            });
        }
        if self.max_relationships > HARD_MAX_RELATIONSHIPS {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain relationships",
                actual: self.max_relationships,
                maximum: HARD_MAX_RELATIONSHIPS,
            });
        }
        if self.max_graph_parts > HARD_MAX_GRAPH_PARTS {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain graph parts",
                actual: self.max_graph_parts,
                maximum: HARD_MAX_GRAPH_PARTS,
            });
        }
        Ok(self)
    }
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The physical Calculation Chain part and its opaque source bytes.
///
/// The bytes are shared across snapshots, transactions, and inverse patches.
/// The optional generic BIFF12 inventory is an explicit diagnostic and does
/// not participate in source admission or publication.
#[derive(Clone, Debug)]
pub struct ChainPart {
    pub(crate) part_name: String,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) max_records: usize,
}

impl PartialEq for ChainPart {
    fn eq(&self, other: &Self) -> bool {
        self.part_name == other.part_name
            && self.content_type == other.content_type
            && self.bytes == other.bytes
    }
}

impl Eq for ChainPart {}

impl ChainPart {
    /// Physical OPC part name retained from the source package.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Physical content type token retained from the source package.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Borrow the opaque source bytes without copying them.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Encoded byte length of the opaque source stream.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    /// Whether the opaque source stream is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    /// Count generic BIFF12 envelopes in the source stream on request.
    ///
    /// MS-XLSB does not normatively provide a Calculation Chain record
    /// grammar, so this diagnostic is never required for snapshot admission.
    /// Arbitrary opaque bytes therefore produce a typed framing error here
    /// while remaining available for exact source-bound removal and inverse
    /// publication. The retained snapshot `max_records` bound is enforced
    /// before the next record is yielded.
    pub fn record_count(&self) -> Result<usize> {
        let raw_limits = RawLimits::new(self.bytes.len(), 0);
        let mut records = Records::try_with_limits(self.bytes.as_slice(), raw_limits)?;
        let mut count = 0usize;
        while records.next().transpose()?.is_some() {
            count = count.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "Calculation Chain records",
            })?;
            if count > self.max_records {
                return Err(Error::LimitExceeded {
                    resource: "Calculation Chain records",
                    actual: count,
                    maximum: self.max_records,
                });
            }
        }
        Ok(count)
    }

    /// Copy the opaque source stream for an explicit caller-owned export.
    #[must_use]
    pub fn to_vec(&self) -> Vec<u8> {
        self.bytes.as_ref().clone()
    }
}
