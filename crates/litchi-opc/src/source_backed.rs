//! Immutable, source-backed access to an OPC package.
//!
//! This module intentionally exposes a smaller surface than [`OpcPackage`].
//! The latter owns mutable parts, while this type keeps ordinary payloads in a
//! positional source until a caller explicitly asks for one.

use crate::accounting::OpcOperationAccounting;
use crate::constants::{content_type, relationship_type};
use crate::content_type::{ContentType, ContentTypeMap};
use crate::error::{OpcError, Result};
use crate::limits::{ReadLimits, ReadResource};
use crate::members::{NonPartMember, PartNameIndex};
use crate::package::OpcPackage;
use crate::packuri::{PACKAGE_URI, PackURI};
use crate::part::PartFactory;
use crate::pkgreader::{
    PackageReader, SerializedRelationship, SourceCatalog, ValidationCatalogError,
    ValidationCatalogPhase, is_xml_id,
};
use crate::rel::{Relationships, TargetMode};
use litchi_core::{
    ExecutionContext, ExecutionError, OwnedSource, ReadAt, Reservation, Resource, SourceVersion,
};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use sha2::{Digest as _, Sha256};
use soapberry_zip::ReaderAt as ZipReaderAt;
use soapberry_zip::ZipOperationAccounting as LowLevelZipOperationAccounting;
use soapberry_zip::office::{EntryId, IndexedArchive};
use std::collections::HashMap;
use std::io::Write;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::time::Duration;

const SOURCE_PUBLICATION_CHUNK_BYTES: usize = 64 * 1024;
const MAX_SOURCE_OVERLAY_PARTS: usize = 64;
const MAX_SOURCE_RELATIONSHIP_REMOVALS: usize = 4096;
const MAX_SOURCE_TOPOLOGY_PARTS: usize = 64;
const MAX_SOURCE_TOPOLOGY_RELATIONSHIPS: usize = 4096;
const MAX_SOURCE_RELATIONSHIP_FIELD_BYTES: usize = 64 * 1024;
const MAX_SOURCE_TOPOLOGY_RELATIONSHIP_BYTES: usize = 8 * 1024 * 1024;

struct PendingOverlay {
    target: usize,
    replacement: Arc<Vec<u8>>,
}

/// A bounded source-backed OPC publication plan.
///
/// The plan is deliberately opaque: callers describe logical Part payload and
/// relationship changes while the consuming publisher retains ownership of
/// ZIP preservation, content-types lexical preservation, and relationship
/// member placement. Building a plan never reads or mutates a source package.
#[derive(Debug, Default)]
pub struct SourceTopologyPlan {
    replacements: Vec<TopologyReplacement>,
    additions: Vec<TopologyPartAddition>,
    relationships: Vec<TopologyRelationshipChange>,
    relationship_bytes: usize,
}

#[derive(Debug)]
struct TopologyReplacement {
    partname: PackURI,
    replacement: Arc<Vec<u8>>,
}

#[derive(Debug)]
struct TopologyPartAddition {
    partname: PackURI,
    content_type: ContentType,
    payload: Arc<Vec<u8>>,
}

/// The target requested for a source-backed relationship operation.
///
/// Internal targets are validated against the source/new Part graph and are
/// serialized as owner-relative references. External targets are retained as
/// URI references exactly as supplied, including query and fragment text.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SourceRelationshipTarget {
    /// A target Part in the OPC package.
    Internal(PackURI),
    /// An external URI reference, preserved without package resolution.
    External(String),
}

impl SourceRelationshipTarget {
    fn mode(&self) -> TargetMode {
        match self {
            Self::Internal(_) => TargetMode::Internal,
            Self::External(_) => TargetMode::External,
        }
    }
}

fn escaped_xml_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let encoded = match character {
            '&' => 5,
            '<' | '>' => 4,
            '\"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length
            .checked_add(encoded)
            .ok_or_else(|| overlay_unavailable("escaped XML attribute length overflows usize"))
    })
}

fn relationship_xml_event_count(xml: &[u8], limits: ReadLimits) -> Result<u64> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut events = 0_u64;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| overlay_unavailable("relationship XML event count overflows u64"))?;
        limits.check(
            ReadResource::XmlEvents,
            events,
            limits.max_xml_events() as u64,
        )?;
        if matches!(reader.read_event()?, Event::Eof) {
            return Ok(events);
        }
    }
}

#[derive(Debug)]
enum TopologyRelationshipOperation {
    Add {
        reltype: String,
        target: SourceRelationshipTarget,
    },
    Replace {
        reltype: String,
        target: SourceRelationshipTarget,
        required_mode: Option<TargetMode>,
    },
    Remove {
        required_mode: Option<TargetMode>,
    },
}

#[derive(Debug)]
struct TopologyRelationshipChange {
    owner: PackURI,
    r_id: String,
    operation: TopologyRelationshipOperation,
}

impl SourceTopologyPlan {
    /// Construct an empty topology plan.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Replace the payload of one existing Part.
    ///
    /// The source-backed publisher verifies that the Part exists and that the
    /// replacement remains within the package read policy before output.
    pub fn try_replace_part(&mut self, partname: PackURI, replacement: Vec<u8>) -> Result<()> {
        if partname.as_str() == PACKAGE_URI {
            return Err(OpcError::InvalidPackUri(
                "the package root is not a Part URI".to_string(),
            ));
        }
        let operation_count = self
            .replacements
            .len()
            .checked_add(self.additions.len())
            .ok_or_else(|| overlay_unavailable("topology Part operation count overflows usize"))?;
        if operation_count >= MAX_SOURCE_TOPOLOGY_PARTS {
            return Err(overlay_unavailable(format!(
                "topology replacement set exceeds the {MAX_SOURCE_TOPOLOGY_PARTS}-Part bound"
            )));
        }
        if self
            .replacements
            .iter()
            .any(|candidate| candidate.partname.is_equivalent_to(&partname))
        {
            return Err(OpcError::DuplicatePartName(partname.to_string()));
        }
        if self
            .additions
            .iter()
            .any(|candidate| candidate.partname.is_equivalent_to(&partname))
        {
            return Err(OpcError::DuplicatePartName(partname.to_string()));
        }
        self.replacements
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology replacements",
                source,
            })?;
        self.replacements.push(TopologyReplacement {
            partname,
            replacement: Arc::new(replacement),
        });
        Ok(())
    }

    /// Add a new typed Part payload.
    pub fn try_add_part(
        &mut self,
        partname: PackURI,
        content_type: impl Into<String>,
        payload: Vec<u8>,
    ) -> Result<()> {
        if partname.as_str() == PACKAGE_URI {
            return Err(OpcError::InvalidPackUri(
                "the package root is not a Part URI".to_string(),
            ));
        }
        let operation_count = self
            .replacements
            .len()
            .checked_add(self.additions.len())
            .ok_or_else(|| overlay_unavailable("topology Part operation count overflows usize"))?;
        if operation_count >= MAX_SOURCE_TOPOLOGY_PARTS {
            return Err(overlay_unavailable(format!(
                "topology addition set exceeds the {MAX_SOURCE_TOPOLOGY_PARTS}-Part bound"
            )));
        }
        if self
            .additions
            .iter()
            .any(|candidate| candidate.partname.is_equivalent_to(&partname))
            || self
                .replacements
                .iter()
                .any(|candidate| candidate.partname.is_equivalent_to(&partname))
        {
            return Err(OpcError::DuplicatePartName(partname.to_string()));
        }
        let content_type = ContentType::new(content_type.into())?;
        self.additions
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology Part additions",
                source,
            })?;
        self.additions.push(TopologyPartAddition {
            partname,
            content_type,
            payload: Arc::new(payload),
        });
        Ok(())
    }

    /// Add a typed relationship owned by the package (`/`), an existing Part,
    /// or a Part added by this plan.
    pub fn try_add_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: SourceRelationshipTarget,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Add {
                reltype: reltype.into(),
                target,
            },
        })
    }

    /// Add an internal relationship owned by the package (`/`), an existing
    /// Part, or a Part added by this plan. The target must be an absolute OPC
    /// Part URI and is serialized as the owner-relative target reference.
    pub fn try_add_internal_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: PackURI,
    ) -> Result<()> {
        self.try_add_relationship(
            owner,
            r_id,
            reltype,
            SourceRelationshipTarget::Internal(target),
        )
    }

    /// Add an external relationship while preserving the exact URI reference.
    pub fn try_add_external_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: impl Into<String>,
    ) -> Result<()> {
        self.try_add_relationship(
            owner,
            r_id,
            reltype,
            SourceRelationshipTarget::External(target.into()),
        )
    }

    /// Replace an existing typed relationship without changing its rId.
    pub fn try_replace_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: SourceRelationshipTarget,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Replace {
                reltype: reltype.into(),
                target,
                required_mode: None,
            },
        })
    }

    /// Replace an existing internal relationship without changing its rId.
    pub fn try_replace_internal_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: PackURI,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Replace {
                reltype: reltype.into(),
                target: SourceRelationshipTarget::Internal(target),
                required_mode: Some(TargetMode::Internal),
            },
        })
    }

    /// Replace an existing external relationship while preserving the exact
    /// URI reference. The publisher requires the current relationship to be
    /// external before applying the replacement.
    pub fn try_replace_external_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
        reltype: impl Into<String>,
        target: impl Into<String>,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Replace {
                reltype: reltype.into(),
                target: SourceRelationshipTarget::External(target.into()),
                required_mode: Some(TargetMode::External),
            },
        })
    }

    /// Remove an existing relationship by owner and rId.
    pub fn try_remove_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Remove {
                required_mode: None,
            },
        })
    }

    /// Remove an existing external relationship by owner and rId.
    pub fn try_remove_external_relationship(
        &mut self,
        owner: PackURI,
        r_id: impl Into<String>,
    ) -> Result<()> {
        self.try_push_relationship_change(TopologyRelationshipChange {
            owner,
            r_id: r_id.into(),
            operation: TopologyRelationshipOperation::Remove {
                required_mode: Some(TargetMode::External),
            },
        })
    }

    fn try_push_relationship_change(&mut self, change: TopologyRelationshipChange) -> Result<()> {
        if self.relationships.len() >= MAX_SOURCE_TOPOLOGY_RELATIONSHIPS {
            return Err(overlay_unavailable(format!(
                "topology relationship set exceeds the {MAX_SOURCE_TOPOLOGY_RELATIONSHIPS}-relationship bound"
            )));
        }
        let TopologyRelationshipChange {
            ref owner,
            ref r_id,
            ref operation,
        } = change;
        let mut field_bytes = owner.as_str().len();
        field_bytes = field_bytes
            .checked_add(r_id.len())
            .ok_or_else(|| overlay_unavailable("topology relationship field bytes overflow"))?;
        if r_id.len() > MAX_SOURCE_RELATIONSHIP_FIELD_BYTES || !is_xml_id(r_id) {
            return Err(OpcError::InvalidRelationship(format!(
                "relationship Id '{r_id}' is not a bounded XML ID"
            )));
        }
        match operation {
            TopologyRelationshipOperation::Add { reltype, target }
            | TopologyRelationshipOperation::Replace {
                reltype, target, ..
            } => {
                if reltype.is_empty()
                    || reltype.len() > MAX_SOURCE_RELATIONSHIP_FIELD_BYTES
                    || reltype.chars().any(char::is_whitespace)
                    || reltype.chars().any(char::is_control)
                {
                    return Err(OpcError::InvalidRelationship(
                        "relationship Type is not a bounded URI reference".to_string(),
                    ));
                }
                field_bytes = field_bytes.checked_add(reltype.len()).ok_or_else(|| {
                    overlay_unavailable("topology relationship field bytes overflow")
                })?;
                let target_bytes = match target {
                    SourceRelationshipTarget::Internal(target) => target.as_str().len(),
                    SourceRelationshipTarget::External(target) => {
                        if target.is_empty()
                            || target.len() > MAX_SOURCE_RELATIONSHIP_FIELD_BYTES
                            || target.chars().any(char::is_control)
                        {
                            return Err(OpcError::InvalidRelationship(
                                "relationship Target is not a bounded URI reference".to_string(),
                            ));
                        }
                        target.len()
                    },
                };
                if target_bytes > MAX_SOURCE_RELATIONSHIP_FIELD_BYTES {
                    return Err(OpcError::InvalidRelationship(
                        "relationship Target is not a bounded URI reference".to_string(),
                    ));
                }
                field_bytes = field_bytes.checked_add(target_bytes).ok_or_else(|| {
                    overlay_unavailable("topology relationship field bytes overflow")
                })?;
            },
            TopologyRelationshipOperation::Remove { .. } => {},
        }
        let next_bytes = self
            .relationship_bytes
            .checked_add(field_bytes)
            .ok_or_else(|| overlay_unavailable("topology relationship bytes overflow"))?;
        if next_bytes > MAX_SOURCE_TOPOLOGY_RELATIONSHIP_BYTES {
            return Err(overlay_unavailable(format!(
                "topology relationship fields exceed the {MAX_SOURCE_TOPOLOGY_RELATIONSHIP_BYTES}-byte bound"
            )));
        }
        if self.relationships.iter().any(|candidate| {
            candidate.owner.is_equivalent_to(owner) && candidate.r_id.as_str() == r_id.as_str()
        }) {
            return Err(OpcError::DuplicateRelationshipId(r_id.clone()));
        }
        self.relationships
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology relationships",
                source,
            })?;
        self.relationship_bytes = next_bytes;
        self.relationships.push(change);
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.replacements.is_empty() && self.additions.is_empty() && self.relationships.is_empty()
    }
}

#[derive(Debug)]
struct TopologyRelationshipPublication {
    member_name: String,
    xml: Arc<Vec<u8>>,
    existing_entry: Option<EntryId>,
    relationship_count: usize,
}

#[derive(Clone, Copy)]
struct PhysicalMemberInfo {
    entry_id: EntryId,
    count: usize,
}

struct PhysicalMemberLookup {
    by_folded_name: HashMap<String, PhysicalMemberInfo>,
}

struct ChangedOverlay {
    target: ChangedOverlayTarget,
    replacement: Arc<Vec<u8>>,
}

enum ChangedOverlayTarget {
    Part(usize),
    Member(String),
}

/// Validation failure returned by [`SourceCacheLimits::new`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceCacheLimitError {
    /// A cache needs a positive byte capacity to retain any payload.
    ZeroMaximumBytes,
    /// A cache needs a positive entry capacity to retain any payload.
    ZeroMaximumEntries,
}

impl std::fmt::Display for SourceCacheLimitError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ZeroMaximumBytes => "source cache maximum bytes must be greater than zero",
            Self::ZeroMaximumEntries => "source cache maximum entries must be greater than zero",
        })
    }
}

impl std::error::Error for SourceCacheLimitError {}

/// Finite retention policy for source-backed part payloads.
///
/// Both limits are enforced: a part is retained only when it fits in
/// [`Self::max_bytes`] and the cache can make room below [`Self::max_entries`].
/// Larger requested parts are returned to the caller but are never cached.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceCacheLimits {
    max_bytes: usize,
    max_entries: usize,
}

impl SourceCacheLimits {
    /// Construct a validated finite cache policy.
    ///
    /// # Errors
    ///
    /// Returns an error when either bound is zero.
    pub const fn new(
        max_bytes: usize,
        max_entries: usize,
    ) -> std::result::Result<Self, SourceCacheLimitError> {
        if max_bytes == 0 {
            return Err(SourceCacheLimitError::ZeroMaximumBytes);
        }
        if max_entries == 0 {
            return Err(SourceCacheLimitError::ZeroMaximumEntries);
        }
        Ok(Self {
            max_bytes,
            max_entries,
        })
    }

    /// Maximum total retained payload bytes.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Maximum retained payload entries.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }
}

impl Default for SourceCacheLimits {
    fn default() -> Self {
        // Both values are literal non-zero constants, so this cannot fail.
        Self {
            max_bytes: 8 * 1024 * 1024,
            max_entries: 128,
        }
    }
}

/// Content-free point-in-time diagnostics for a source-backed payload cache.
///
/// Counters are monotonically increasing for the package lifetime and use
/// checked relaxed atomic CAS updates, so a snapshot is observational rather
/// than a globally linearized transaction. The checked CAS is intentional
/// instrumentation overhead: event recording must never wrap silently.
/// `retained_entries`, `retained_bytes`, and `in_flight_loads` are captured
/// together while the cache takes its existing short lock. No member names,
/// part URIs, or ZIP entry IDs are exposed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SourceCacheDiagnostics {
    /// Requests satisfied directly from a retained payload entry.
    pub hits: u64,
    /// Requests that became the loader for a cold payload read.
    pub cold_loads: u64,
    /// Requests that found an existing same-part cold load and waited for it.
    pub waiter_joins: u64,
    /// Cold reads that completed successfully, whether retained or uncached.
    pub successful_loads: u64,
    /// Cold reads that failed before a payload could be published.
    pub failed_loads: u64,
    /// Retained entries removed to satisfy byte or entry capacity.
    pub evictions: u64,
    /// Successful cold reads returned without retention for any bypass reason.
    pub bypasses: u64,
    /// Successful cold reads returned without retention because their payload
    /// exceeded the configured byte limit.
    pub oversized_bypasses: u64,
    /// Requests or successful loads that could not be coordinated or retained
    /// because cache bookkeeping allocation failed.
    pub allocation_bypasses: u64,
    /// Payload entries currently retained by the cache.
    pub retained_entries: usize,
    /// Payload bytes currently retained by the cache.
    pub retained_bytes: usize,
    /// Same-part cold loads currently coordinated by a flight.
    pub in_flight_loads: usize,
    /// Whether this cache charges retained and in-flight payloads to a caller
    /// supplied hierarchical memory budget.
    pub budget_managed: bool,
    /// Managed budget reservations rejected by the hierarchical budget. A
    /// final `InputBytes` or managed direct-publication `OutputBytes` refusal
    /// is counted; a temporary oversized read window that is shrunk to the
    /// remaining capacity is not.
    pub budget_reservation_failures: u64,
    /// Current memory usage observed on the managed context's local budget.
    /// This is content-free and may include sibling operations sharing the
    /// same budget.
    pub budget_memory_used: u64,
    /// Bytes reserved by retained cache entries and active cold-load flights.
    /// This deliberately excludes ordinary caller-owned [`PartData`] handles
    /// that were returned after a cache entry was evicted or bypassed.
    pub budget_cache_reserved_bytes: u64,
    /// Local memory limit observed on the managed context's budget. `None`
    /// means that the compatibility, unmanaged cache path is active.
    pub budget_memory_limit: Option<u64>,
    /// Cumulative physical bytes accepted from positional reads charged to
    /// this context. Shared contexts may include sibling operations. This is
    /// never released when a package, cache entry, or payload handle is
    /// dropped.
    pub budget_input_bytes_used: u64,
    /// Local cumulative input-byte limit, or `None` for an unmanaged cache.
    pub budget_input_bytes_limit: Option<u64>,
    /// Cumulative physical bytes accepted by managed source-backed publication
    /// sinks. Shared contexts may include sibling operations. This is never
    /// released after acceptance, including an incomplete publication.
    pub budget_output_bytes_used: u64,
    /// Local cumulative output-byte limit, or `None` for an unmanaged cache.
    pub budget_output_bytes_limit: Option<u64>,
    /// Cumulative declared cold-load work charged before payload I/O. A
    /// successful ZIP read proves that the declared uncompressed size is also
    /// the actual materialized size; hits and waiters add no work. Shared
    /// contexts may include sibling operations.
    pub budget_work_used: u64,
    /// Local cumulative work limit, or `None` for an unmanaged cache.
    pub budget_work_limit: Option<u64>,
    /// Current retained object usage observed on the managed context's local
    /// budget. Shared contexts may include sibling operations. Unlike input
    /// bytes and work, this usage is released when the package catalog or a
    /// payload object is dropped.
    pub budget_objects_used: u64,
    /// Local retained-object limit, or `None` for an unmanaged cache.
    pub budget_objects_limit: Option<u64>,
    /// Object units retained by the package catalog itself.
    pub budget_catalog_reserved_objects: u64,
    /// Object units retained by cache entries and active flights. Returned
    /// handles that outlive eviction are reflected in `budget_objects_used`
    /// but deliberately excluded here, matching the retained-byte diagnostic.
    /// Shared reservations are counted once.
    pub budget_cache_reserved_objects: u64,
}

/// The event-counter portion of a source-cache diagnostic interval.
///
/// This deliberately excludes point-in-time gauges such as retained bytes and
/// in-flight loads. Use [`SourceCacheDiagnostics::checked_counter_delta`] to
/// construct one from two snapshots; a counter regression is rejected rather
/// than interpreted as a wrapped interval.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SourceCacheCounterDelta {
    /// Requests satisfied directly from a retained payload entry.
    pub hits: u64,
    /// Requests that became the loader for a cold payload read.
    pub cold_loads: u64,
    /// Requests that found an existing same-part cold load and waited for it.
    pub waiter_joins: u64,
    /// Cold reads that completed successfully.
    pub successful_loads: u64,
    /// Cold reads that failed before publication.
    pub failed_loads: u64,
    /// Retained entries removed to satisfy cache capacity.
    pub evictions: u64,
    /// Successful cold reads returned without retention.
    pub bypasses: u64,
    /// Successful cold reads bypassed because their payload was oversized.
    pub oversized_bypasses: u64,
    /// Requests or successful loads bypassed after cache bookkeeping
    /// allocation failed.
    pub allocation_bypasses: u64,
    /// Managed budget reservation refusals attributed to the cache or its
    /// source-backed publication counters.
    pub budget_reservation_failures: u64,
}

/// Failure returned by a fail-closed source-cache diagnostic operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceCacheDiagnosticsError {
    /// A checked diagnostic counter could not be incremented without wrapping.
    CounterOverflow,
    /// The cache state mutex was poisoned by a panic while it was held.
    StatePoisoned,
    /// A supposedly monotonic counter decreased between two snapshots.
    CounterMovedBackwards {
        /// Counter whose interval was invalid.
        counter: &'static str,
    },
}

impl std::fmt::Display for SourceCacheDiagnosticsError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::CounterOverflow => formatter
                .write_str("source-cache diagnostic counter overflowed; metrics are unavailable"),
            Self::StatePoisoned => formatter
                .write_str("source-cache diagnostic state is poisoned; metrics are unavailable"),
            Self::CounterMovedBackwards { counter } => {
                write!(formatter, "source-cache counter {counter} moved backwards")
            },
        }
    }
}

impl std::error::Error for SourceCacheDiagnosticsError {}

impl SourceCacheDiagnostics {
    /// Compute a checked interval for the monotonic event counters.
    ///
    /// The returned delta contains event counters only; current cache gauges
    /// are intentionally not subtracted. A counter regression is an invalid
    /// or wrapped interval and returns an error. Callers collecting metrics
    /// should obtain both snapshots through
    /// [`SourceBackedPackage::try_cache_diagnostics`].
    pub fn checked_counter_delta(
        before: Self,
        after: Self,
    ) -> std::result::Result<SourceCacheCounterDelta, SourceCacheDiagnosticsError> {
        let subtract = |counter: &'static str, after: u64, before: u64| {
            after
                .checked_sub(before)
                .ok_or(SourceCacheDiagnosticsError::CounterMovedBackwards { counter })
        };
        Ok(SourceCacheCounterDelta {
            hits: subtract("hits", after.hits, before.hits)?,
            cold_loads: subtract("cold_loads", after.cold_loads, before.cold_loads)?,
            waiter_joins: subtract("waiter_joins", after.waiter_joins, before.waiter_joins)?,
            successful_loads: subtract(
                "successful_loads",
                after.successful_loads,
                before.successful_loads,
            )?,
            failed_loads: subtract("failed_loads", after.failed_loads, before.failed_loads)?,
            evictions: subtract("evictions", after.evictions, before.evictions)?,
            bypasses: subtract("bypasses", after.bypasses, before.bypasses)?,
            oversized_bypasses: subtract(
                "oversized_bypasses",
                after.oversized_bypasses,
                before.oversized_bypasses,
            )?,
            allocation_bypasses: subtract(
                "allocation_bypasses",
                after.allocation_bypasses,
                before.allocation_bypasses,
            )?,
            budget_reservation_failures: subtract(
                "budget_reservation_failures",
                after.budget_reservation_failures,
                before.budget_reservation_failures,
            )?,
        })
    }
}

#[derive(Debug, Default)]
struct DiagnosticState {
    overflowed: AtomicBool,
}

#[derive(Debug)]
struct DiagnosticCounter {
    value: AtomicU64,
    state: Arc<DiagnosticState>,
}

impl DiagnosticCounter {
    fn new(state: Arc<DiagnosticState>) -> Self {
        Self {
            value: AtomicU64::new(0),
            state,
        }
    }

    fn increment(&self) {
        if self
            .value
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |value| {
                value.checked_add(1)
            })
            .is_err()
        {
            self.state.overflowed.store(true, Ordering::Release);
        }
    }

    fn load(&self) -> u64 {
        self.value.load(Ordering::Relaxed)
    }
}

#[derive(Debug)]
struct CacheCounters {
    hits: DiagnosticCounter,
    cold_loads: DiagnosticCounter,
    waiter_joins: DiagnosticCounter,
    successful_loads: DiagnosticCounter,
    failed_loads: DiagnosticCounter,
    evictions: DiagnosticCounter,
    bypasses: DiagnosticCounter,
    oversized_bypasses: DiagnosticCounter,
    allocation_bypasses: DiagnosticCounter,
    budget_reservation_failures: DiagnosticCounter,
}

impl CacheCounters {
    fn new(state: Arc<DiagnosticState>) -> Self {
        Self {
            hits: DiagnosticCounter::new(Arc::clone(&state)),
            cold_loads: DiagnosticCounter::new(Arc::clone(&state)),
            waiter_joins: DiagnosticCounter::new(Arc::clone(&state)),
            successful_loads: DiagnosticCounter::new(Arc::clone(&state)),
            failed_loads: DiagnosticCounter::new(Arc::clone(&state)),
            evictions: DiagnosticCounter::new(Arc::clone(&state)),
            bypasses: DiagnosticCounter::new(Arc::clone(&state)),
            oversized_bypasses: DiagnosticCounter::new(Arc::clone(&state)),
            allocation_bypasses: DiagnosticCounter::new(Arc::clone(&state)),
            budget_reservation_failures: DiagnosticCounter::new(state),
        }
    }
}

#[derive(Clone)]
struct SourceReader {
    snapshot: SourceSnapshot,
}

impl ZipReaderAt for SourceReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> std::io::Result<usize> {
        let result = read_source_at_with_context(
            &self.snapshot,
            self.snapshot.context.as_ref(),
            offset,
            output,
            "archive",
        );
        result.map_err(|error| {
            let execution = match &error {
                OpcError::Cancelled => Some(ExecutionError::Cancelled),
                OpcError::Execution(execution) => Some(execution.clone()),
                _ => None,
            };
            if let Some(execution) = execution {
                record_source_execution_failure(&self.snapshot, execution);
            }
            match error {
                OpcError::IoError(error) => error,
                OpcError::Execution(error) => execution_io_error(error),
                error => std::io::Error::other(error.to_string()),
            }
        })
    }
}

#[derive(Clone)]
struct SourceSnapshot {
    source: Arc<dyn ReadAt>,
    version: SourceVersion,
    length: u64,
    monitor_reads: Arc<AtomicBool>,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
    execution_failure: Option<Arc<Mutex<Option<ExecutionError>>>>,
    input_reservation_failures: Option<Arc<DiagnosticCounter>>,
    output_reservation_failures: Option<Arc<DiagnosticCounter>>,
}

/// Process-local identity for one opened source-backed package lineage.
///
/// A lineage is intentionally distinct from [`SourceVersion`]. Two source
/// adapters may report the same caller-chosen version token while still being
/// different package instances; patches must never cross that boundary.
#[derive(Clone, Debug)]
pub struct SourceLineage(Arc<()>);

impl PartialEq for SourceLineage {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for SourceLineage {}

/// Exact immutable source artifact retained for a later reversible restore.
#[derive(Clone)]
pub struct SourceArtifact {
    snapshot: SourceSnapshot,
}

/// SHA-256 identity of an exact source artifact.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceArtifactFingerprint([u8; 32]);

impl SourceArtifactFingerprint {
    /// Construct from a completed SHA-256 digest.
    #[must_use]
    pub const fn from_sha256(digest: [u8; 32]) -> Self {
        Self(digest)
    }
}

impl SourceArtifact {
    /// Hash the exact current artifact without materializing it.
    pub fn fingerprint(&self) -> Result<SourceArtifactFingerprint> {
        self.snapshot.ensure_current()?;
        let mut hasher = Sha256::new();
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
            .map_err(|source| OpcError::Allocation {
                resource: "source artifact fingerprint buffer",
                source,
            })?;
        buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
        let mut offset = 0_u64;
        while offset < self.snapshot.length {
            let remaining =
                usize::try_from((self.snapshot.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
            let read = read_source_at_with_context(
                &self.snapshot,
                self.snapshot.context.as_ref(),
                offset,
                &mut buffer[..remaining],
                "fingerprinting",
            )?;
            if read == 0 {
                return Err(OpcError::IoError(std::io::Error::new(
                    std::io::ErrorKind::UnexpectedEof,
                    "source-backed OPC source ended during fingerprinting",
                )));
            }
            self.snapshot.ensure_current()?;
            hasher.update(&buffer[..read]);
            offset = offset
                .checked_add(read as u64)
                .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
        }
        self.snapshot.ensure_current()?;
        Ok(SourceArtifactFingerprint(hasher.finalize().into()))
    }

    /// Copy the retained source artifact exactly to a sequential sink.
    ///
    /// A managed artifact reserves `Resource::OutputBytes` for each sink
    /// write and commits only the accepted count. A refusal before any bytes
    /// are accepted returns the typed output resource-limit error; a sink or
    /// policy failure after accepted bytes returns [`OpcError::IncompleteOutput`].
    pub fn write_to_stream<W: Write>(&self, writer: W) -> Result<()> {
        write_exact_snapshot(&self.snapshot, writer, self.snapshot.context.as_ref())
    }

    /// Copy the retained source artifact while recording exact accepted output
    /// and unchanged-source bytes in a caller-owned report.
    ///
    /// The report is updated as the sink accepts bytes, including a partial
    /// prefix before a source, cancellation, policy, or sink error. The raw
    /// unchanged counter equals the actual accepted source bytes for this
    /// exact-copy operation; it is never inferred from the source length.
    pub fn write_to_stream_with_accounting<W: Write>(
        &self,
        writer: W,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<()> {
        write_exact_snapshot_with_accounting(
            &self.snapshot,
            writer,
            self.snapshot.context.as_ref(),
            accounting,
        )
    }
}

impl SourceSnapshot {
    fn ensure_current(&self) -> Result<()> {
        let actual = self.source.version()?;
        if actual == self.version {
            Ok(())
        } else {
            Err(OpcError::SourceChanged {
                expected: self.version,
                actual,
            })
        }
    }

    fn ensure_current_io_if_monitored(&self) -> std::io::Result<()> {
        if !self.monitor_reads.load(Ordering::Acquire) {
            return Ok(());
        }
        let actual = self.source.version()?;
        if actual == self.version {
            Ok(())
        } else {
            Err(std::io::Error::other(
                "source-backed OPC source changed during publication",
            ))
        }
    }

    fn monitor_publication(&self) {
        self.monitor_reads.store(true, Ordering::Release);
    }
}

struct Counted<'count, W> {
    inner: W,
    written: &'count mut u64,
    accounting: Option<AccountingCounter<'count>>,
}

struct AccountingCounter<'count> {
    report: &'count mut OpcOperationAccounting,
    error: &'count mut Option<OpcError>,
    raw_source: bool,
}

impl AccountingCounter<'_> {
    fn record(&mut self, bytes: usize) {
        let Ok(bytes) = u64::try_from(bytes) else {
            record_accounting_error(
                self.error,
                accounting_overflow("accepted output byte count exceeds u64"),
            );
            return;
        };
        if let Err(error) = self.report.add_output_bytes_accepted(bytes) {
            record_accounting_error(self.error, error);
        }
        if self.raw_source {
            if let Err(error) = self.report.add_raw_unchanged_source_bytes_accepted(bytes) {
                record_accounting_error(self.error, error);
            }
        }
    }
}

impl<'count, W> Counted<'count, W> {
    fn new(inner: W, written: &'count mut u64) -> Self {
        Self {
            inner,
            written,
            accounting: None,
        }
    }

    fn with_accounting(
        inner: W,
        written: &'count mut u64,
        report: &'count mut OpcOperationAccounting,
        error: &'count mut Option<OpcError>,
        raw_source: bool,
    ) -> Self {
        Self {
            inner,
            written,
            accounting: Some(AccountingCounter {
                report,
                error,
                raw_source,
            }),
        }
    }
}

impl<W: Write> Write for Counted<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        if written > bytes.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "source-backed OPC sink reported {written} bytes for a {}-byte write",
                    bytes.len()
                ),
            ));
        }
        *self.written = self.written.saturating_add(written as u64);
        if let Some(accounting) = self.accounting.as_mut() {
            accounting.record(written);
        }
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

struct SourceCheckedSink<W> {
    inner: W,
    snapshot: SourceSnapshot,
}

struct ContextCheckedSink<W> {
    inner: W,
    context: Option<ExecutionContext>,
    failure: Arc<Mutex<Option<ExecutionError>>>,
}

impl<W: Write> ContextCheckedSink<W> {
    fn record_failure(&self, error: ExecutionError) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check_before_write(&self) -> std::io::Result<()> {
        let Some(context) = self.context.as_ref() else {
            return Ok(());
        };
        context.check().map_err(|error| {
            let message = error.to_string();
            self.record_failure(error);
            // `Write::write_all` retries `Interrupted` indefinitely. Use a
            // terminal I/O classification; the shared failure slot below
            // restores the typed execution error after ZIP writing returns.
            std::io::Error::other(message)
        })
    }
}

impl<W: Write> Write for ContextCheckedSink<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.check_before_write()?;
        let result = self.inner.write(bytes);
        if result.is_ok() {
            // Check after each bounded Chunked write as well. If a sink
            // cancels from inside its first write, the accepted byte count is
            // preserved and the next write/flush returns the typed failure.
            if let Some(context) = self.context.as_ref() {
                if let Err(error) = context.check() {
                    self.record_failure(error);
                }
            }
        }
        result
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.check_before_write()?;
        let result = self.inner.flush();
        if result.is_ok() {
            if let Some(context) = self.context.as_ref() {
                if let Err(error) = context.check() {
                    self.record_failure(error);
                }
            }
        }
        result
    }
}

/// Managed-only output-budget adapter. The compatibility path deliberately
/// stays on [`ContextCheckedSink`] and does not instantiate this wrapper.
struct OutputBudgetedSink<W> {
    inner: W,
    context: ExecutionContext,
    failure: Arc<Mutex<Option<ExecutionError>>>,
    output_reservation_failures: Arc<DiagnosticCounter>,
}

impl<W: Write> OutputBudgetedSink<W> {
    fn record_failure(&self, error: ExecutionError) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check_before_write(&self) -> std::io::Result<()> {
        self.context.check().map_err(|error| {
            let message = error.to_string();
            self.record_failure(error);
            // `Write::write_all` retries `Interrupted` indefinitely. Use a
            // terminal I/O classification; the shared failure slot below
            // restores the typed execution error after ZIP writing returns.
            std::io::Error::other(message)
        })
    }

    fn reserve_output(&self, requested: usize) -> std::io::Result<Option<Reservation>> {
        if requested == 0 {
            return Ok(None);
        }
        let requested = u64::try_from(requested).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "source-backed OPC output request exceeds u64",
            )
        })?;
        self.context
            .reserve(Resource::OutputBytes, requested)
            .map(Some)
            .map_err(|error| {
                if matches!(
                    &error,
                    ExecutionError::ResourceLimit(limit)
                        if limit.resource == Resource::OutputBytes
                ) {
                    self.output_reservation_failures.increment();
                }
                let message = error.to_string();
                self.record_failure(error);
                // The ZIP preservation layer communicates sink failures through
                // std::io::Error. The shared failure slot restores the typed
                // execution error after that adapter returns.
                std::io::Error::other(message)
            })
    }
}

impl<W: Write> Write for OutputBudgetedSink<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.check_before_write()?;
        let reservation = self.reserve_output(bytes.len())?;
        let result = self.inner.write(bytes);
        match result {
            Ok(written) if written > bytes.len() => {
                drop(reservation);
                Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!(
                        "source-backed OPC sink reported {written} bytes for a {}-byte write",
                        bytes.len()
                    ),
                ))
            },
            Ok(written) => {
                let written = u64::try_from(written).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "source-backed OPC sink progress exceeds u64",
                    )
                })?;
                if let Some(reservation) = reservation {
                    if !reservation.commit(written) {
                        return Err(std::io::Error::new(
                            std::io::ErrorKind::InvalidData,
                            "source-backed OPC output reservation underflow",
                        ));
                    }
                }
                Ok(usize::try_from(written).unwrap_or(usize::MAX))
            },
            Err(error) => {
                drop(reservation);
                Err(error)
            },
        }
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.check_before_write()?;
        self.inner.flush()
    }
}

impl<W: Write> Write for SourceCheckedSink<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.snapshot.ensure_current_io_if_monitored()?;
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.snapshot.ensure_current_io_if_monitored()?;
        self.inner.flush()
    }
}

struct Chunked<W> {
    inner: W,
}

impl<W: Write> Write for Chunked<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.inner
            .write(&bytes[..bytes.len().min(SOURCE_PUBLICATION_CHUNK_BYTES)])
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

#[derive(Debug)]
struct CatalogPart {
    partname: PackURI,
    content_type: String,
    relationships: Relationships,
    entry_id: EntryId,
}

/// Immutable metadata and deferred payload access for one OPC package part.
pub struct PartView<'package> {
    package: &'package SourceBackedPackage,
    index: usize,
}

impl PartView<'_> {
    /// The part's absolute OPC URI.
    #[must_use]
    pub fn partname(&self) -> &PackURI {
        &self.package.parts[self.index].partname
    }

    /// The content type declared by `[Content_Types].xml`.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.package.parts[self.index].content_type
    }

    /// The part's already-validated relationships.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.package.parts[self.index].relationships
    }

    /// Return the ZIP central-directory declaration for this part's
    /// uncompressed size without loading its payload.
    ///
    /// The value is attacker-controlled central-directory metadata and is not
    /// proof of the decoded length. [`PartView::data`] still verifies the
    /// decoded length and closes the central-directory/local-header TOCTOU
    /// boundary before publishing a payload. Source freshness and the
    /// caller's cancellation policy are checked before and after this lookup.
    pub fn declared_uncompressed_size(&self) -> Result<u64> {
        self.package.source.ensure_current()?;
        self.package
            .cache
            .check_context()
            .map_err(map_execution_error)?;
        let entry_id = self
            .package
            .parts
            .get(self.index)
            .ok_or_else(|| OpcError::PartNotFound(self.index.to_string()))?
            .entry_id;
        let declared = self
            .package
            .archive
            .metadata_for(entry_id)?
            .uncompressed_size();
        self.package.source.ensure_current()?;
        self.package
            .cache
            .check_context()
            .map_err(map_execution_error)?;
        Ok(declared)
    }

    /// Read this part's payload, using the package's bounded cache when able.
    pub fn data(&self) -> Result<PartData> {
        self.package.read_part(self.index)
    }

    /// Read this part's payload while recording the cold ZIP work in a
    /// caller-owned report.
    ///
    /// Cache hits and same-Part flight waiters return with no counter changes.
    /// Only the loader or allocation-bypass path that performs the ZIP read
    /// contributes to the report. A failed cold read retains the counters
    /// observed before the failure, while a later retry owns its own work.
    pub fn data_with_accounting(
        &self,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<PartData> {
        self.package
            .read_part_with_accounting(self.index, Some(accounting))
    }
}

/// Shared immutable bytes returned by [`PartView::data`].
#[derive(Clone, Debug)]
pub struct PartData {
    payload: CachedPayload,
}

impl PartData {
    /// Borrow the part payload.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.payload.bytes.as_slice()
    }

    /// Share an unmanaged payload allocation with another owner.
    ///
    /// Managed payloads retain a hierarchical memory reservation for the
    /// lifetime of this handle and cannot be detached as a bare `Arc`. Use
    /// [`Self::as_bytes`] or clone the [`PartData`] handle instead.
    pub fn into_arc(&self) -> Result<Arc<Vec<u8>>> {
        if self.payload.reservation.is_some() {
            return Err(OpcError::ManagedPartDataArcEscape);
        }
        Ok(Arc::clone(&self.payload.bytes))
    }

    /// Return whether both values pin the same payload allocation.
    ///
    /// This compares allocation identity only; equal bytes loaded separately
    /// return `false`.
    #[must_use]
    pub fn shares_allocation_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.payload.bytes, &other.payload.bytes)
    }
}

/// One payload allocation and, for managed packages, the reservation retained
/// by a cache entry or active same-Part flight.
#[derive(Clone, Debug)]
struct CachedPayload {
    bytes: Arc<Vec<u8>>,
    reservation: Option<Arc<Reservation>>,
    object_reservation: Option<Arc<Reservation>>,
}

impl CachedPayload {
    fn reserved_bytes(&self) -> u64 {
        self.reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount())
    }

    fn reserved_objects(&self) -> u64 {
        self.object_reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount())
    }
}

#[derive(Debug)]
struct CacheEntry {
    payload: CachedPayload,
    last_used: u64,
}

#[derive(Debug, Default)]
struct CacheState {
    entries: HashMap<EntryId, CacheEntry>,
    flights: HashMap<EntryId, Arc<LoadFlight>>,
    total_bytes: usize,
    clock: u64,
}

#[derive(Debug, Default)]
struct FlightState {
    complete: bool,
    payload: Option<CachedPayload>,
}

#[derive(Debug)]
struct LoadFlight {
    state: Mutex<FlightState>,
    completed: Condvar,
    reservation: Option<Arc<Reservation>>,
    flight_object_reservation: Option<Arc<Reservation>>,
    payload_object_reservation: Option<Arc<Reservation>>,
}

impl LoadFlight {
    fn new(
        reservation: Option<Arc<Reservation>>,
        flight_object_reservation: Option<Arc<Reservation>>,
        payload_object_reservation: Option<Arc<Reservation>>,
    ) -> Self {
        Self {
            state: Mutex::new(FlightState::default()),
            completed: Condvar::new(),
            reservation,
            flight_object_reservation,
            payload_object_reservation,
        }
    }

    fn reservation(&self) -> Option<Arc<Reservation>> {
        self.reservation.as_ref().map(Arc::clone)
    }

    fn payload_object_reservation(&self) -> Option<Arc<Reservation>> {
        self.payload_object_reservation.as_ref().map(Arc::clone)
    }

    fn wait(
        &self,
        context: Option<&ExecutionContext>,
    ) -> std::result::Result<Option<CachedPayload>, ExecutionError> {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        while !state.complete {
            if let Some(context) = context {
                context.check()?;
            }
            state = match self
                .completed
                .wait_timeout(state, Duration::from_millis(10))
            {
                Ok((state, _)) => state,
                Err(poisoned) => poisoned.into_inner().0,
            };
        }
        if let Some(context) = context {
            context.check()?;
        }
        Ok(state.payload.clone())
    }

    fn finish_success(&self, payload: CachedPayload) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.payload = Some(payload);
        state.complete = true;
        self.completed.notify_all();
    }

    fn finish_failure(&self) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.complete = true;
        self.completed.notify_all();
    }
}

enum CacheAccess {
    Hit(CachedPayload),
    Loader(Arc<LoadFlight>),
    Waiter(Arc<LoadFlight>),
    Bypass(LoadResources),
}

struct LoadResources {
    reservation: Option<Arc<Reservation>>,
    payload_object_reservation: Option<Arc<Reservation>>,
}

#[derive(Debug)]
struct PartCache {
    limits: SourceCacheLimits,
    state: Mutex<CacheState>,
    counters: CacheCounters,
    diagnostics: Arc<DiagnosticState>,
    budget: Option<ExecutionContext>,
    input_reservation_failures: Option<Arc<DiagnosticCounter>>,
    output_reservation_failures: Option<Arc<DiagnosticCounter>>,
}

impl PartCache {
    #[cfg(test)]
    fn new(limits: SourceCacheLimits) -> Self {
        Self::new_with_diagnostics(limits, Arc::new(DiagnosticState::default()))
    }

    fn new_with_diagnostics(limits: SourceCacheLimits, diagnostics: Arc<DiagnosticState>) -> Self {
        Self {
            limits,
            state: Mutex::new(CacheState::default()),
            counters: CacheCounters::new(Arc::clone(&diagnostics)),
            diagnostics,
            budget: None,
            input_reservation_failures: None,
            output_reservation_failures: None,
        }
    }

    fn new_managed(
        limits: SourceCacheLimits,
        context: ExecutionContext,
        diagnostics: Arc<DiagnosticState>,
        input_reservation_failures: Arc<DiagnosticCounter>,
        output_reservation_failures: Arc<DiagnosticCounter>,
    ) -> Self {
        Self {
            limits,
            state: Mutex::new(CacheState::default()),
            counters: CacheCounters::new(Arc::clone(&diagnostics)),
            diagnostics,
            budget: Some(context),
            input_reservation_failures: Some(input_reservation_failures),
            output_reservation_failures: Some(output_reservation_failures),
        }
    }

    fn is_managed(&self) -> bool {
        self.budget.is_some()
    }

    fn check_context(&self) -> std::result::Result<(), ExecutionError> {
        if let Some(context) = self.budget.as_ref() {
            context.check()?;
        }
        Ok(())
    }

    fn record_budget_reservation_failure(&self) {
        self.counters.budget_reservation_failures.increment();
    }

    fn reservation_failure_counter(&self) -> Option<&DiagnosticCounter> {
        self.budget
            .as_ref()
            .map(|_| &self.counters.budget_reservation_failures)
    }

    fn context(&self) -> Option<&ExecutionContext> {
        self.budget.as_ref()
    }

    fn enter(
        &self,
        entry_id: EntryId,
        declared_bytes: u64,
    ) -> std::result::Result<CacheAccess, ExecutionError> {
        self.check_context()?;
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        if let Some(entry) = state.entries.get_mut(&entry_id) {
            entry.last_used = clock;
            self.counters.hits.increment();
            return Ok(CacheAccess::Hit(entry.payload.clone()));
        }
        if let Some(flight) = state.flights.get(&entry_id) {
            self.counters.waiter_joins.increment();
            return Ok(CacheAccess::Waiter(Arc::clone(flight)));
        }

        let reservation = self.reserve_for_load(&mut state, declared_bytes)?;
        let payload_object_reservation = self.reserve_object_for_load()?;
        if state.flights.try_reserve(1).is_err() {
            self.charge_cold_work(declared_bytes)?;
            self.counters.cold_loads.increment();
            self.counters.allocation_bypasses.increment();
            return Ok(CacheAccess::Bypass(LoadResources {
                reservation,
                payload_object_reservation,
            }));
        }
        let flight_object_reservation = self.reserve_object_for_load()?;
        self.charge_cold_work(declared_bytes)?;
        let flight = Arc::new(LoadFlight::new(
            reservation,
            flight_object_reservation,
            payload_object_reservation,
        ));
        state.flights.insert(entry_id, Arc::clone(&flight));
        self.counters.cold_loads.increment();
        Ok(CacheAccess::Loader(flight))
    }

    fn reserve_object_for_load(
        &self,
    ) -> std::result::Result<Option<Arc<Reservation>>, ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(None);
        };
        match context.reserve(Resource::Objects, 1) {
            Ok(reservation) => Ok(Some(Arc::new(reservation))),
            Err(error) => {
                if matches!(error, ExecutionError::ResourceLimit(_)) {
                    self.record_budget_reservation_failure();
                }
                Err(error)
            },
        }
    }

    fn charge_cold_work(&self, declared_bytes: u64) -> std::result::Result<(), ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(());
        };
        // Work is cumulative. Charge the declared decompression output before
        // the archive reader starts; `load_part` later requires the actual
        // verified output length to equal this declaration, so no guessed or
        // second work charge is needed.
        context.consume(Resource::Work, declared_bytes)
    }

    fn reserve_for_load(
        &self,
        state: &mut CacheState,
        declared_bytes: u64,
    ) -> std::result::Result<Option<Arc<Reservation>>, ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(None);
        };

        context.check()?;
        self.make_room_for_load(state, declared_bytes);
        match context.reserve(Resource::Memory, declared_bytes) {
            Ok(reservation) => Ok(Some(Arc::new(reservation))),
            Err(first_error) => {
                if matches!(first_error, ExecutionError::ResourceLimit(_)) {
                    self.record_budget_reservation_failure();
                }
                // Cache retention is best effort. If a shared ancestor is
                // currently full, dropping all clean entries can make room
                // without ever exceeding that ancestor's limit.
                self.evict_all(state);
                match context.reserve(Resource::Memory, declared_bytes) {
                    Ok(reservation) => Ok(Some(Arc::new(reservation))),
                    Err(error) => {
                        if matches!(error, ExecutionError::ResourceLimit(_)) {
                            self.record_budget_reservation_failure();
                        }
                        Err(error)
                    },
                }
            },
        }
    }

    fn make_room_for_load(&self, state: &mut CacheState, declared_bytes: u64) {
        let weight = usize::try_from(declared_bytes).unwrap_or(usize::MAX);
        if weight > self.limits.max_bytes {
            self.evict_all(state);
            return;
        }
        while state.entries.len() >= self.limits.max_entries
            || state.total_bytes.saturating_add(weight) > self.limits.max_bytes
        {
            if !self.evict_oldest(state) {
                break;
            }
        }
    }

    fn evict_all(&self, state: &mut CacheState) {
        while self.evict_oldest(state) {}
    }

    fn evict_oldest(&self, state: &mut CacheState) -> bool {
        let Some((&oldest, _)) = state
            .entries
            .iter()
            .filter(|(_, entry)| !payload_is_externally_pinned(&entry.payload))
            .min_by_key(|(_, entry)| entry.last_used)
        else {
            return false;
        };
        if let Some(removed) = state.entries.remove(&oldest) {
            state.total_bytes = state
                .total_bytes
                .saturating_sub(removed.payload.bytes.len());
            self.counters.evictions.increment();
            true
        } else {
            false
        }
    }

    fn complete_success(
        &self,
        entry_id: EntryId,
        flight: &Arc<LoadFlight>,
        payload: CachedPayload,
    ) -> std::result::Result<(), ExecutionError> {
        self.check_context()?;
        let delivered = payload.clone();
        {
            let mut state = self
                .state
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            // This is the final cooperative cancellation point immediately
            // before publishing a clean value into the shared cache.
            self.check_context()?;
            self.counters.successful_loads.increment();
            self.record_retention(self.insert_locked(&mut state, entry_id, payload));
        }
        // Complete before removing the flight so an oversized, deliberately
        // uncached value still has no gap in which a late peer can start a
        // duplicate load instead of joining this successful delivery.
        flight.finish_success(delivered);
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        remove_flight(&mut state, entry_id, flight);
        Ok(())
    }

    fn complete_failure(&self, entry_id: EntryId, flight: &Arc<LoadFlight>) {
        self.counters.failed_loads.increment();
        // Publish failure to current waiters before allowing a new retrying
        // loader to install a replacement flight.
        flight.finish_failure();
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        remove_flight(&mut state, entry_id, flight);
    }

    fn complete_bypass_success(
        &self,
        entry_id: EntryId,
        payload: CachedPayload,
    ) -> std::result::Result<(), ExecutionError> {
        self.check_context()?;
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        // Keep the no-flight allocation-fallback path under the same
        // pre-publication cancellation contract as the normal flight path.
        self.check_context()?;
        self.counters.successful_loads.increment();
        self.record_retention(self.insert_locked(&mut state, entry_id, payload));
        Ok(())
    }

    fn complete_bypass_failure(&self) {
        self.counters.failed_loads.increment();
    }

    fn record_retention(&self, retention: CacheRetention) {
        match retention {
            CacheRetention::Retained => {},
            CacheRetention::Oversized => {
                self.counters.bypasses.increment();
                self.counters.oversized_bypasses.increment();
            },
            CacheRetention::Pinned => {
                self.counters.bypasses.increment();
            },
            CacheRetention::AllocationFailure => {
                self.counters.bypasses.increment();
                self.counters.allocation_bypasses.increment();
            },
        }
    }

    fn insert_locked(
        &self,
        state: &mut CacheState,
        entry_id: EntryId,
        payload: CachedPayload,
    ) -> CacheRetention {
        let weight = payload.bytes.len();
        if weight > self.limits.max_bytes {
            return CacheRetention::Oversized;
        }
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        while state.entries.len() >= self.limits.max_entries
            || state.total_bytes.saturating_add(weight) > self.limits.max_bytes
        {
            if !self.evict_oldest(state) {
                return CacheRetention::Pinned;
            }
        }
        if state.entries.try_reserve(1).is_err() {
            return CacheRetention::AllocationFailure;
        }
        if let Some(previous) = state.entries.insert(
            entry_id,
            CacheEntry {
                payload,
                last_used: clock,
            },
        ) {
            state.total_bytes = state
                .total_bytes
                .saturating_sub(previous.payload.bytes.len());
        }
        state.total_bytes = state.total_bytes.saturating_add(weight);
        CacheRetention::Retained
    }

    fn diagnostics(&self) -> SourceCacheDiagnostics {
        let state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        self.diagnostics_snapshot(&state)
    }

    fn try_diagnostics(
        &self,
    ) -> std::result::Result<SourceCacheDiagnostics, SourceCacheDiagnosticsError> {
        let state = self
            .state
            .lock()
            .map_err(|_| SourceCacheDiagnosticsError::StatePoisoned)?;
        if self.diagnostics.overflowed.load(Ordering::Acquire) {
            return Err(SourceCacheDiagnosticsError::CounterOverflow);
        }
        let snapshot = self.diagnostics_snapshot(&state);
        if self.diagnostics.overflowed.load(Ordering::Acquire) {
            return Err(SourceCacheDiagnosticsError::CounterOverflow);
        }
        Ok(snapshot)
    }

    fn diagnostics_snapshot(&self, state: &CacheState) -> SourceCacheDiagnostics {
        // A successful loader briefly owns the same reservation through its
        // cache entry, completion payload, and returned handles. Count only
        // unique reservation identities so a diagnostic snapshot cannot
        // report more retained cache bytes than the hierarchical budget.
        let mut budget_cache_reserved_bytes = 0_u64;
        let mut budget_cache_reserved_objects = 0_u64;
        for entry in state.entries.values() {
            self.add_diagnostic_gauge(
                &mut budget_cache_reserved_bytes,
                entry.payload.reserved_bytes(),
            );
            self.add_diagnostic_gauge(
                &mut budget_cache_reserved_objects,
                entry.payload.reserved_objects(),
            );
        }
        for flight in state.flights.values() {
            if let Some(reservation) = flight.reservation.as_ref() {
                let already_counted = state.entries.values().any(|entry| {
                    entry
                        .payload
                        .reservation
                        .as_ref()
                        .is_some_and(|existing| Arc::ptr_eq(existing, reservation))
                });
                if !already_counted {
                    self.add_diagnostic_gauge(
                        &mut budget_cache_reserved_bytes,
                        reservation.amount(),
                    );
                }
            }
            if let Some(object_reservation) = flight.payload_object_reservation.as_ref() {
                let already_counted = state.entries.values().any(|entry| {
                    entry
                        .payload
                        .object_reservation
                        .as_ref()
                        .is_some_and(|existing| Arc::ptr_eq(existing, object_reservation))
                });
                if !already_counted {
                    self.add_diagnostic_gauge(
                        &mut budget_cache_reserved_objects,
                        object_reservation.amount(),
                    );
                }
            }
            if let Some(object_reservation) = flight.flight_object_reservation.as_ref() {
                self.add_diagnostic_gauge(
                    &mut budget_cache_reserved_objects,
                    object_reservation.amount(),
                );
            }
        }
        let (budget_input_bytes_used, budget_input_bytes_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::InputBytes),
                    Some(context.budget().limit(Resource::InputBytes)),
                )
            });
        let (budget_work_used, budget_work_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::Work),
                    Some(context.budget().limit(Resource::Work)),
                )
            });
        let (budget_output_bytes_used, budget_output_bytes_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::OutputBytes),
                    Some(context.budget().limit(Resource::OutputBytes)),
                )
            });
        let (budget_objects_used, budget_objects_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::Objects),
                    Some(context.budget().limit(Resource::Objects)),
                )
            });
        let budget_reservation_failures = self
            .counters
            .budget_reservation_failures
            .load()
            .checked_add(
                self.input_reservation_failures
                    .as_ref()
                    .map_or(0, |counter| counter.load()),
            )
            .and_then(|total| {
                total.checked_add(
                    self.output_reservation_failures
                        .as_ref()
                        .map_or(0, |counter| counter.load()),
                )
            })
            .unwrap_or_else(|| {
                self.diagnostics.overflowed.store(true, Ordering::Release);
                u64::MAX
            });
        SourceCacheDiagnostics {
            hits: self.counters.hits.load(),
            cold_loads: self.counters.cold_loads.load(),
            waiter_joins: self.counters.waiter_joins.load(),
            successful_loads: self.counters.successful_loads.load(),
            failed_loads: self.counters.failed_loads.load(),
            evictions: self.counters.evictions.load(),
            bypasses: self.counters.bypasses.load(),
            oversized_bypasses: self.counters.oversized_bypasses.load(),
            allocation_bypasses: self.counters.allocation_bypasses.load(),
            retained_entries: state.entries.len(),
            retained_bytes: state.total_bytes,
            in_flight_loads: state.flights.len(),
            budget_managed: self.budget.is_some(),
            budget_reservation_failures,
            budget_memory_used: self
                .budget
                .as_ref()
                .map_or(0, |context| context.budget().used(Resource::Memory)),
            budget_cache_reserved_bytes,
            budget_memory_limit: self
                .budget
                .as_ref()
                .map(|context| context.budget().limit(Resource::Memory)),
            budget_input_bytes_used,
            budget_input_bytes_limit,
            budget_output_bytes_used,
            budget_output_bytes_limit,
            budget_work_used,
            budget_work_limit,
            budget_objects_used,
            budget_objects_limit,
            budget_catalog_reserved_objects: 0,
            budget_cache_reserved_objects,
        }
    }

    fn add_diagnostic_gauge(&self, total: &mut u64, amount: u64) {
        if let Some(next) = total.checked_add(amount) {
            *total = next;
        } else {
            *total = u64::MAX;
            self.diagnostics.overflowed.store(true, Ordering::Release);
        }
    }
}

#[derive(Clone, Copy)]
enum CacheRetention {
    Retained,
    Oversized,
    Pinned,
    AllocationFailure,
}

fn remove_flight(state: &mut CacheState, entry_id: EntryId, flight: &Arc<LoadFlight>) {
    if state
        .flights
        .get(&entry_id)
        .is_some_and(|current| Arc::ptr_eq(current, flight))
    {
        state.flights.remove(&entry_id);
    }
}

fn payload_is_externally_pinned(payload: &CachedPayload) -> bool {
    // Unmanaged `PartData::into_arc` can outlive its handle, while managed
    // handles retain a reservation. Check both identities before evicting an
    // entry so either form of caller ownership keeps the bytes pinned.
    Arc::strong_count(&payload.bytes) > 1
        || payload
            .reservation
            .as_ref()
            .is_some_and(|reservation| Arc::strong_count(reservation) > 1)
}

/// A structurally validated OPC package backed by an immutable positional source.
///
/// Opening reads and validates ZIP metadata, content types, and relationship
/// XML, but never reads ordinary part payloads. The ordinary view is immutable.
/// [`Self::write_part_overlays_to_stream`] is a narrow, consuming publisher for
/// a bounded same-topology Part replacement set that raw-copies every other
/// ZIP member; call [`Self::into_opc_package`] when a general owning mutable
/// package is needed.
pub struct SourceBackedPackage {
    source: SourceSnapshot,
    archive: IndexedArchive<SourceReader>,
    limits: ReadLimits,
    content_types_member: String,
    package_relationships: Relationships,
    parts: Vec<CatalogPart>,
    parts_by_name: HashMap<PackURI, usize>,
    non_part_members: Vec<NonPartMember>,
    cache: PartCache,
    catalog_object_reservation: Option<Arc<Reservation>>,
}

/// Validation-only open failure with exact ingress phase provenance.
pub(crate) struct ValidationOpenError {
    pub(crate) phase: ValidationCatalogPhase,
    pub(crate) error: OpcError,
}

impl SourceBackedPackage {
    /// Validation-only source open. Ordinary callers retain the existing open
    /// path; this variant adds phase provenance without changing its hot path.
    pub(crate) fn from_read_at_for_validation(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
    ) -> std::result::Result<Self, ValidationOpenError> {
        let phase = |phase, error| ValidationOpenError { phase, error };
        let version = source
            .version()
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let length = source
            .len()
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        limits
            .check(ReadResource::InputBytes, length, limits.max_input_bytes())
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let diagnostics = Arc::new(DiagnosticState::default());
        let snapshot = SourceSnapshot {
            source: Arc::clone(&source),
            version,
            length,
            monitor_reads: Arc::new(AtomicBool::new(false)),
            lineage: SourceLineage(Arc::new(())),
            context: None,
            execution_failure: None,
            input_reservation_failures: None,
            output_reservation_failures: None,
        };
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let archive = IndexedArchive::from_reader_with_limits(
            SourceReader {
                snapshot: snapshot.clone(),
            },
            length,
            limits.zip_limits(),
        )
        .map_err(OpcError::from)
        .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
            content_types_member,
        } = PackageReader::source_catalog_for_validation(&archive, limits).map_err(
            |ValidationCatalogError {
                 phase: stage,
                 error,
             }| phase(stage, error),
        )?;
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;

        let package_relationships = relationships_for_package(pkg_srels)
            .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
        let mut catalog_parts = Vec::new();
        catalog_parts
            .try_reserve_exact(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC catalog parts",
                source,
            })
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let mut parts_by_name = HashMap::new();
        parts_by_name
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC part lookup",
                source,
            })
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        for (index, part) in parts.into_iter().enumerate() {
            let relationships = relationships_for_part(&part.partname, part.srels)
                .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
            let entry_id = archive
                .entry_id(part.partname.membername())
                .ok_or_else(|| OpcError::PartNotFound(part.partname.to_string()))
                .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
            parts_by_name.insert(part.partname.clone(), index);
            catalog_parts.push(CatalogPart {
                partname: part.partname,
                content_type: part.content_type,
                relationships,
                entry_id,
            });
        }

        Ok(Self {
            source: snapshot,
            archive,
            limits,
            content_types_member,
            package_relationships,
            parts: catalog_parts,
            parts_by_name,
            non_part_members,
            cache: PartCache::new_with_diagnostics(SourceCacheLimits::default(), diagnostics),
            catalog_object_reservation: None,
        })
    }

    /// Open a source-backed package with the standard bounded read policy.
    ///
    /// The input vector is moved into the source owner. Ordinary part payloads
    /// remain deferred until a [`PartView::data`] request, just as with
    /// [`Self::from_read_at`].
    pub fn from_vec(data: Vec<u8>) -> Result<Self> {
        Self::from_read_at(Arc::new(OwnedSource::new(data)))
    }

    /// Open an owned source-backed package with an explicit bounded read
    /// policy.
    ///
    /// The input vector is moved into the source owner. Opening validates the
    /// ZIP catalog, content types, and relationships without materializing
    /// ordinary part payloads; a payload is decompressed only when its
    /// [`PartView::data`] is requested.
    pub fn from_vec_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(Arc::new(OwnedSource::new(data)), limits)
    }

    /// Open a source-backed package with the standard bounded read policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(
            source,
            ReadLimits::default(),
            SourceCacheLimits::default(),
        )
    }

    /// Open a source-backed package with an explicit bounded read policy.
    ///
    /// The source version is captured before indexing and checked after every
    /// mandatory open read.  A changed source is never silently accepted.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(
            source,
            limits,
            SourceCacheLimits::default(),
        )
    }

    /// Open a source-backed package with an explicit payload-cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(source, ReadLimits::default(), cache_limits)
    }

    /// Open a source-backed package whose lazy payload cache is charged to an
    /// explicit hierarchical execution budget.
    ///
    /// Compatibility constructors remain unmanaged and keep their existing
    /// behavior. This opt-in path checks cancellation before opening and
    /// reserves each Part's declared uncompressed size before reading its
    /// payload. The reservation is retained with the clean cache entry and
    /// active same-Part flight. A returned [`PartData`] is a budgeted handle;
    /// use [`PartData::as_bytes`] or clone that handle while consuming the
    /// payload. Its [`PartData::into_arc`] escape is rejected on this managed
    /// path so the reservation cannot be silently detached.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            limits,
            SourceCacheLimits::default(),
            context,
        )
    }

    /// Open a managed source-backed package with explicit read, cache, and
    /// hierarchical execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_inner(source, limits, cache_limits, Some(context))
    }

    /// Open a source-backed package with explicit read and cache policies.
    ///
    /// The source version is captured before indexing and checked after every
    /// mandatory open read. A changed source is never silently accepted.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_inner(source, limits, cache_limits, None)
    }

    fn from_read_at_inner(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: Option<ExecutionContext>,
    ) -> Result<Self> {
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let execution_failure = context
            .as_ref()
            .map(|_| Arc::new(Mutex::new(None::<ExecutionError>)));
        let diagnostics = Arc::new(DiagnosticState::default());
        let input_reservation_failures = context
            .as_ref()
            .map(|_| Arc::new(DiagnosticCounter::new(Arc::clone(&diagnostics))));
        let output_reservation_failures = context
            .as_ref()
            .map(|_| Arc::new(DiagnosticCounter::new(Arc::clone(&diagnostics))));
        let version = source.version()?;
        let length = source.len()?;
        limits.check(ReadResource::InputBytes, length, limits.max_input_bytes())?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let snapshot = SourceSnapshot {
            source: Arc::clone(&source),
            version,
            length,
            monitor_reads: Arc::new(AtomicBool::new(false)),
            lineage: SourceLineage(Arc::new(())),
            context: context.clone(),
            execution_failure: execution_failure.clone(),
            input_reservation_failures: input_reservation_failures.clone(),
            output_reservation_failures: output_reservation_failures.clone(),
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let archive = match IndexedArchive::from_reader_with_limits(
            SourceReader {
                snapshot: snapshot.clone(),
            },
            length,
            limits.zip_limits(),
        ) {
            Ok(archive) => archive,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&snapshot) {
                    return Err(map_execution_error(execution));
                }
                return Err(error.into());
            },
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        // The indexed archive and its source catalog remain owned by the
        // package for its whole lifetime. Reserve one object for the package
        // catalog owner and one for every physical member before parsing the
        // source catalog and projecting deferred part vectors; this is a
        // bounded retained charge, not a cumulative event. Payloads and
        // load flights use separate units.
        let catalog_object_reservation = if let Some(context) = context.as_ref() {
            let member_objects = u64::try_from(archive.len()).map_err(|_| {
                overlay_unavailable("source-backed OPC catalog member count overflows u64")
            })?;
            let object_count = member_objects.checked_add(1).ok_or_else(|| {
                overlay_unavailable("source-backed OPC catalog object count overflows u64")
            })?;
            Some(Arc::new(
                context
                    .reserve(Resource::Objects, object_count)
                    .map_err(map_execution_error)?,
            ))
        } else {
            None
        };
        let SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
            content_types_member,
        } = match PackageReader::source_catalog(&archive, limits) {
            Ok(catalog) => catalog,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&snapshot) {
                    return Err(map_execution_error(execution));
                }
                return Err(error);
            },
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let package_relationships =
            relationships_for_package_with_context(pkg_srels, context.as_ref())?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let mut catalog_parts = Vec::new();
        catalog_parts
            .try_reserve_exact(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC catalog parts",
                source,
            })?;
        let mut parts_by_name = HashMap::new();
        parts_by_name
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC part lookup",
                source,
            })?;
        for (index, part) in parts.into_iter().enumerate() {
            if let Some(context) = context.as_ref() {
                context.check().map_err(map_execution_error)?;
            }
            let relationships =
                relationships_for_part_with_context(&part.partname, part.srels, context.as_ref())?;
            if let Some(context) = context.as_ref() {
                context.check().map_err(map_execution_error)?;
            }
            let entry_id = archive
                .entry_id(part.partname.membername())
                .ok_or_else(|| OpcError::PartNotFound(part.partname.to_string()))?;
            parts_by_name.insert(part.partname.clone(), index);
            catalog_parts.push(CatalogPart {
                partname: part.partname,
                content_type: part.content_type,
                relationships,
                entry_id,
            });
        }
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }

        let cache = if let Some(context) = context {
            let input_reservation_failures = input_reservation_failures.ok_or_else(|| {
                overlay_unavailable("managed source input reservation counter is unavailable")
            })?;
            let output_reservation_failures = output_reservation_failures.ok_or_else(|| {
                overlay_unavailable("managed source output reservation counter is unavailable")
            })?;
            PartCache::new_managed(
                cache_limits,
                context,
                diagnostics,
                input_reservation_failures,
                output_reservation_failures,
            )
        } else {
            PartCache::new_with_diagnostics(cache_limits, diagnostics)
        };
        Ok(Self {
            source: snapshot,
            archive,
            limits,
            content_types_member,
            package_relationships,
            parts: catalog_parts,
            parts_by_name,
            non_part_members,
            cache,
            catalog_object_reservation,
        })
    }

    /// Package-level relationships parsed during opening.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.package_relationships
    }

    /// Return metadata-only views of every ordinary part.
    pub fn iter_parts(&self) -> impl Iterator<Item = PartView<'_>> {
        (0..self.parts.len()).map(|index| PartView {
            package: self,
            index,
        })
    }

    /// Return the exact ZIP member names retained by the indexed source.
    ///
    /// This metadata-only iterator is intended for low-level physical-name
    /// collision checks. It includes ordinary parts, relationship members,
    /// package metadata, and non-part members without reading payload bytes.
    pub fn physical_member_names(&self) -> impl ExactSizeIterator<Item = &str> {
        self.archive.file_names()
    }

    /// Whether any retained physical member declares traditional ZIP encryption.
    ///
    /// This is central-directory metadata captured during bounded open; it does
    /// not read or decrypt any member payload.
    #[must_use]
    pub fn has_encrypted_entries(&self) -> bool {
        self.archive.has_encrypted_entries()
    }

    /// Look up one ordinary part without reading its payload.
    pub fn part(&self, partname: &PackURI) -> Result<PartView<'_>> {
        self.source.ensure_current()?;
        self.part_index(partname)
            .map(|index| PartView {
                package: self,
                index,
            })
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Resolve a catalog part by OPC part-name equivalence.
    ///
    /// The exact hash lookup is the common path. OPC part names compare
    /// case-insensitively, however, and [`PackageReader`] admits only one
    /// spelling of an ASCII-case-equivalent name. Consequently a miss can be
    /// resolved by a bounded, allocation-free scan without introducing a
    /// second folded-name index into the source-backed catalog.
    fn part_index(&self, partname: &PackURI) -> Option<usize> {
        self.parts_by_name.get(partname).copied().or_else(|| {
            let wanted = partname.as_str();
            self.parts
                .iter()
                .position(|part| part.partname.as_str().eq_ignore_ascii_case(wanted))
        })
    }

    /// Return the unique main document part without reading its payload.
    pub fn main_document_part(&self) -> Result<PartView<'_>> {
        self.source.ensure_current()?;
        let mut matching = self.package_relationships.iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::OFFICE_DOCUMENT | relationship_type::STRICT_OFFICE_DOCUMENT
            )
        });
        let relationship = matching.next().ok_or_else(|| {
            OpcError::InvalidRelationship("main-document relationship is missing".to_string())
        })?;
        if matching.next().is_some() {
            return Err(OpcError::InvalidRelationship(
                "package has multiple main-document relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OpcError::InvalidRelationship(
                "main-document relationship cannot be external".to_string(),
            ));
        }
        let partname = relationship.target_partname()?;
        self.part(&partname)
    }

    /// ZIP items present in the source but not modelled as OPC parts.
    #[must_use]
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Return content-free payload-cache and managed-publication activity.
    ///
    /// See [`SourceCacheDiagnostics`] for the precise event definitions. This
    /// operation does not read part payloads or expose member identifiers. For
    /// managed packages it also reports cumulative direct-publication
    /// `OutputBytes` usage/limit and budget refusal counts; this includes
    /// [`SourceArtifact::write_to_stream`] and bounded overlay sinks, while
    /// unmanaged compatibility publication remains uncharged.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        let mut diagnostics = self.cache.diagnostics();
        diagnostics.budget_catalog_reserved_objects = self
            .catalog_object_reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount());
        diagnostics
    }

    /// Return a fail-closed payload-cache diagnostic snapshot.
    ///
    /// Unlike [`Self::cache_diagnostics`], this instrumentation-oriented
    /// variant reports a poisoned cache-state mutex or a checked counter
    /// overflow as an error. Ordinary compatibility callers can continue to
    /// use the infallible method; benchmark and telemetry code should use this
    /// method so an invalid snapshot cannot be mistaken for a valid interval.
    pub fn try_cache_diagnostics(
        &self,
    ) -> std::result::Result<SourceCacheDiagnostics, SourceCacheDiagnosticsError> {
        let mut diagnostics = self.cache.try_diagnostics()?;
        diagnostics.budget_catalog_reserved_objects = self
            .catalog_object_reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount());
        Ok(diagnostics)
    }

    /// Check the caller-supplied execution policy for a source-backed
    /// operation.
    ///
    /// Compatibility packages have no execution context and therefore always
    /// pass this check. Managed packages use the check to let a semantic
    /// facade honor cancellation even when its own parsed value is already
    /// retained and no new [`PartData`] read is necessary.
    pub fn check_execution(&self) -> Result<()> {
        self.cache.check_context().map_err(map_execution_error)
    }

    /// Return the exact source lineage captured by this package.
    ///
    /// The returned token is clone-cheap and cannot be constructed by a
    /// caller. It lets a semantic facade bind snapshots and patches to this
    /// opened package rather than to a merely equal [`SourceVersion`].
    #[must_use]
    pub fn source_lineage(&self) -> SourceLineage {
        self.source.lineage.clone()
    }

    /// Clone the caller-supplied execution context, if this package is on the
    /// managed path. Compatibility packages return `None`.
    #[must_use]
    pub fn execution_context(&self) -> Option<ExecutionContext> {
        self.cache.context().cloned()
    }

    /// Return the exact process-local source identity and revision captured at
    /// open after verifying that the positional source is still current.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.source.ensure_current()?;
        Ok(self.source.version)
    }

    /// Retain an O(1) handle to the exact immutable source artifact.
    #[must_use]
    pub fn source_artifact(&self) -> SourceArtifact {
        SourceArtifact {
            snapshot: self.source.clone(),
        }
    }

    /// Validate that the located ZIP ends at the exact source boundary.
    ///
    /// The offset was retained by the initial positional archive locator, so
    /// this check performs no source read, allocation, or directory rescan. It
    /// lets a format facade refuse trailing opaque bytes before returning a
    /// topology-changing semantic plan.
    pub fn validate_topology_source_boundary(&self) -> Result<()> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        if self.archive.archive_end_offset() != self.source.length {
            return Err(overlay_unavailable(
                "source ZIP archive has trailing bytes outside its located archive",
            ));
        }
        Ok(())
    }

    /// Fully materialize this immutable view into the existing mutable package type.
    ///
    /// This conversion refuses managed packages before any ordinary payload
    /// read. An owning package would detach copied payloads from the cache
    /// handles that retain the execution context's hierarchical reservations;
    /// use the source-backed view while that context is active. A materialized
    /// signed graph retains borrowed-ingress policy and must be explicitly
    /// stripped or resigned before ordinary publication.
    pub fn into_opc_package(self) -> Result<OpcPackage> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        if self.cache.is_managed() {
            return Err(OpcError::ManagedPackageMaterialization);
        }
        let mut package = self;
        let non_part_members = std::mem::take(&mut package.non_part_members);
        let result = package.materialize_opc_package(non_part_members);
        package.finish_stage(result)
    }

    /// Materialize this immutable view into an owning package without
    /// consuming the source-backed package.
    ///
    /// The borrowed conversion is intentionally unavailable for packages
    /// opened with an execution context. An owning package would detach the
    /// copied payloads from the cache handles that retain the context's
    /// hierarchical memory and object reservations. The refusal occurs after
    /// source and execution checks, but before any ordinary payload read.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::ManagedPackageMaterialization`] for a managed
    /// package, or the same source, limit, relationship, ZIP, and allocation
    /// errors as [`Self::into_opc_package`]. A source mutation or cancellation
    /// observed before, during, or after materialization rejects the result.
    /// A signed borrowed graph remains policy-gated after materialization.
    pub fn to_opc_package(&self) -> Result<OpcPackage> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        if self.cache.is_managed() {
            return Err(OpcError::ManagedPackageMaterialization);
        }
        let non_part_members = self.finish_stage(self.cloned_non_part_members())?;
        let result = self.materialize_opc_package(non_part_members);
        self.finish_stage(result)
    }

    fn cloned_non_part_members(&self) -> Result<Vec<NonPartMember>> {
        let mut members = Vec::new();
        let reserve = members
            .try_reserve_exact(self.non_part_members.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC non-part members",
                source,
            });
        self.finish_stage(reserve)?;
        for member in &self.non_part_members {
            let result = NonPartMember::new(member.name(), member.reason());
            members.push(self.finish_stage(result)?);
        }
        Ok(members)
    }

    fn materialize_opc_package(&self, non_part_members: Vec<NonPartMember>) -> Result<OpcPackage> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        let mut package = OpcPackage::new();
        let result = copy_relationships(&self.package_relationships, package.rels_mut());
        self.finish_stage(result)?;
        for index in 0..self.parts.len() {
            let bytes = self.finish_stage(self.read_part(index))?;
            let catalog_part = &self.parts[index];
            let part_result = PartFactory::load(
                catalog_part.partname.clone(),
                catalog_part.content_type.clone(),
                bytes.as_bytes().to_vec(),
            );
            let mut part = self.finish_stage(part_result)?;
            let result = copy_relationships(&catalog_part.relationships, part.rels_mut());
            self.finish_stage(result)?;
            let result = package.try_add_source_part(part);
            self.finish_stage(result)?;
        }
        package.set_non_part_members(non_part_members);
        package.mark_source_ingress_signature_policy();
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        Ok(package)
    }

    fn finish_stage<T>(&self, result: Result<T>) -> Result<T> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        result
    }

    /// Publish a bounded source-backed OPC topology plan to a sequential
    /// stream.
    ///
    /// Existing Part payloads may be replaced, new typed Parts may be added,
    /// and relationships may be added, replaced, or removed on the package or
    /// an existing or newly added Part. Unchanged members retain their exact
    /// source ZIP records. The content-types manifest retains its original
    /// bytes and member spelling; only required new `Override` elements are
    /// inserted immediately before the parsed `Types` closing tag. Existing
    /// relationship members are changed only when their source bytes are
    /// exactly the current canonical [`Relationships::to_xml`] form.
    ///
    /// An empty plan is an exact source copy, including signed packages and
    /// physical details unsupported by the rewrite preservation primitive.
    pub fn write_topology_to_stream<W: Write>(
        self,
        writer: W,
        plan: SourceTopologyPlan,
    ) -> Result<()> {
        if plan.is_empty() {
            return self.write_exact_source(writer);
        }
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;

        let SourceTopologyPlan {
            mut replacements,
            mut additions,
            mut relationships,
            relationship_bytes: _,
        } = plan;
        self.check_topology_progress()?;
        replacements
            .sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));
        additions
            .sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));
        relationships.sort_unstable_by(|left, right| {
            left.owner
                .as_str()
                .cmp(right.owner.as_str())
                .then_with(|| left.r_id.cmp(&right.r_id))
        });
        self.check_topology_progress()?;
        let (physical_members, _physical_member_memory) = self.build_physical_member_lookup()?;
        let content_types_key = folded_ascii_name(
            &self.content_types_member,
            "source-backed OPC folded physical member name",
        )?;
        let content_types_member_count = physical_members
            .by_folded_name
            .get(&content_types_key)
            .map_or(0, |info| info.count);
        if content_types_member_count != 1 {
            return Err(overlay_unavailable(
                "source must contain exactly one content-types member",
            ));
        }

        // Resolve replacements against the immutable source catalog. A
        // replacement is deliberately never allowed to target a Part added by
        // the same plan; callers must express that as the new Part payload.
        let mut pending_replacements = Vec::new();
        pending_replacements
            .try_reserve_exact(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology replacement targets",
                source,
            })?;
        for (index, replacement) in replacements.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            let target = self
                .part_index(&replacement.partname)
                .ok_or_else(|| OpcError::PartNotFound(replacement.partname.to_string()))?;
            pending_replacements.push(PendingOverlay {
                target,
                replacement: Arc::clone(&replacement.replacement),
            });
        }
        self.validate_overlay_limits(
            pending_replacements
                .iter()
                .map(|overlay| (overlay.target, overlay.replacement.len())),
        )?;

        // Build the source/new Part namespace with the same duplicate,
        // equivalent, and derived-name rules used by package ingestion.
        let namespace_capacity = self
            .parts
            .len()
            .checked_add(additions.len())
            .ok_or_else(|| overlay_unavailable("topology Part count overflows usize"))?;
        self.limits.check(
            ReadResource::Parts,
            namespace_capacity as u64,
            self.limits.max_parts() as u64,
        )?;
        let mut part_names = PartNameIndex::try_with_capacity(namespace_capacity)?;
        for (index, part) in self.parts.iter().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            part_names.insert(&part.partname)?;
        }
        for (index, addition) in additions.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if is_relationship_member_name(addition.partname.membername()) {
                return Err(overlay_unavailable(
                    "a topology Part cannot use a reserved relationships member name",
                ));
            }
            part_names.insert(&addition.partname)?;
            let physical_key = folded_ascii_name(
                addition.partname.membername(),
                "source-backed OPC folded physical member name",
            )?;
            if physical_members.by_folded_name.contains_key(&physical_key) {
                return Err(overlay_unavailable(format!(
                    "new Part '{}' collides with a physical source member",
                    addition.partname
                )));
            }
        }

        // Keep one fallibly allocated, ASCII-folded namespace lookup for
        // relationship owner/target canonicalization. This prevents a
        // case-equivalent owner spelling from splitting groups and avoids a
        // relationship-by-relationship scan of the complete Part catalog.
        let mut canonical_part_names = HashMap::new();
        canonical_part_names
            .try_reserve(namespace_capacity)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology canonical Part lookup",
                source,
            })?;
        for (index, part) in self.parts.iter().enumerate() {
            let key = folded_part_name(&part.partname)?;
            if canonical_part_names.insert(key, index).is_some() {
                return Err(overlay_unavailable(
                    "source Part names are not uniquely ASCII-case-resolvable",
                ));
            }
        }
        let additions_offset = self.parts.len();
        for (index, addition) in additions.iter().enumerate() {
            let key = folded_part_name(&addition.partname)?;
            let encoded = additions_offset
                .checked_add(index)
                .ok_or_else(|| overlay_unavailable("topology Part lookup index overflows"))?;
            if canonical_part_names.insert(key, encoded).is_some() {
                return Err(overlay_unavailable(
                    "topology Part names are not uniquely ASCII-case-resolvable",
                ));
            }
        }

        // Validate all relationship owners and internal targets before
        // constructing any generated XML. Internal targets must resolve to a
        // physical existing or newly added Part; external targets remain
        // caller-owned URI references and are never resolved through the Part
        // graph.
        for (index, relationship) in relationships.iter_mut().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if relationship.owner.as_str() != PACKAGE_URI {
                let owner_key = folded_part_name(&relationship.owner)?;
                let Some(owner_index) = canonical_part_names.get(&owner_key).copied() else {
                    return Err(OpcError::PartNotFound(format!(
                        "relationship owner '{}'",
                        relationship.owner
                    )));
                };
                relationship.owner = if owner_index < additions_offset {
                    self.parts[owner_index].partname.clone()
                } else {
                    additions[owner_index - additions_offset].partname.clone()
                };
            }
            let (TopologyRelationshipOperation::Add { target, .. }
            | TopologyRelationshipOperation::Replace { target, .. }) = &mut relationship.operation
            else {
                continue;
            };
            let SourceRelationshipTarget::Internal(target) = target else {
                continue;
            };
            if target.as_str() == PACKAGE_URI {
                return Err(OpcError::InvalidRelationship(
                    "internal relationship target cannot be the package root".to_string(),
                ));
            }
            let target_key = folded_part_name(target)?;
            let Some(target_index) = canonical_part_names.get(&target_key).copied() else {
                return Err(OpcError::PartNotFound(format!(
                    "relationship target '{}'",
                    target
                )));
            };
            *target = if target_index < additions_offset {
                self.parts[target_index].partname.clone()
            } else {
                additions[target_index - additions_offset].partname.clone()
            };
        }
        relationships.sort_unstable_by(|left, right| {
            left.owner
                .as_str()
                .cmp(right.owner.as_str())
                .then_with(|| left.r_id.cmp(&right.r_id))
        });
        self.check_topology_progress()?;

        // Group relationship changes by owner and produce canonical XML.
        // The member itself is appended only when the source has no such
        // member; an existing member is regenerated in place after a strict
        // canonical-source check.
        let mut relationship_publications = Vec::new();
        relationship_publications
            .try_reserve_exact(relationships.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology relationship publications",
                source,
            })?;
        let mut group_start = 0usize;
        while group_start < relationships.len() {
            self.check_topology_progress()?;
            let owner = relationships[group_start].owner.clone();
            let mut group_end = group_start + 1;
            while group_end < relationships.len() && relationships[group_end].owner == owner {
                group_end += 1;
            }

            let owner_index = if owner.as_str() == PACKAGE_URI {
                None
            } else {
                self.part_index(&owner)
            };
            let (mut owner_relationships, owner_base) = if owner.as_str() == PACKAGE_URI {
                (self.package_relationships.clone(), "/".to_string())
            } else if let Some(index) = owner_index {
                (
                    self.parts[index].relationships.clone(),
                    self.parts[index].partname.base_uri().to_string(),
                )
            } else {
                (
                    Relationships::for_source(&owner),
                    owner.base_uri().to_string(),
                )
            };
            let mut group_changed = false;
            for (index, relationship) in relationships[group_start..group_end].iter().enumerate() {
                if index & 0x3f == 0 {
                    self.check_topology_progress()?;
                }
                match &relationship.operation {
                    TopologyRelationshipOperation::Add { reltype, target } => {
                        if owner_relationships.get(&relationship.r_id).is_some() {
                            return Err(OpcError::DuplicateRelationshipId(
                                relationship.r_id.clone(),
                            ));
                        }
                        let (target_ref, target_mode) =
                            Self::topology_relationship_target_ref(target, &owner_base)?;
                        self.validate_topology_relationship_field_limits(
                            &relationship.r_id,
                            reltype,
                            &target_ref,
                        )?;
                        owner_relationships.try_add_relationship(
                            reltype.clone(),
                            target_ref,
                            relationship.r_id.clone(),
                            target_mode,
                        )?;
                        group_changed = true;
                    },
                    TopologyRelationshipOperation::Replace {
                        reltype,
                        target,
                        required_mode,
                    } => {
                        let existing =
                            owner_relationships.get(&relationship.r_id).ok_or_else(|| {
                                OpcError::RelationshipNotFound(format!(
                                    "relationship '{}' was not found",
                                    relationship.r_id
                                ))
                            })?;
                        if required_mode.is_some_and(|mode| existing.target_mode() != mode) {
                            return Err(OpcError::InvalidRelationship(format!(
                                "relationship '{}' does not have the required target mode",
                                relationship.r_id
                            )));
                        }
                        let (target_ref, target_mode) =
                            Self::topology_relationship_target_ref(target, &owner_base)?;
                        self.validate_topology_relationship_field_limits(
                            &relationship.r_id,
                            reltype,
                            &target_ref,
                        )?;
                        if existing.reltype() == reltype
                            && existing.target_ref() == target_ref
                            && existing.target_mode() == target_mode
                        {
                            continue;
                        }
                        let removed = owner_relationships.remove(&relationship.r_id);
                        debug_assert!(removed.is_some());
                        owner_relationships.try_add_relationship(
                            reltype.clone(),
                            target_ref,
                            relationship.r_id.clone(),
                            target_mode,
                        )?;
                        group_changed = true;
                    },
                    TopologyRelationshipOperation::Remove { required_mode } => {
                        let existing =
                            owner_relationships.get(&relationship.r_id).ok_or_else(|| {
                                OpcError::RelationshipNotFound(format!(
                                    "relationship '{}' was not found",
                                    relationship.r_id
                                ))
                            })?;
                        if required_mode.is_some_and(|mode| existing.target_mode() != mode) {
                            return Err(OpcError::InvalidRelationship(format!(
                                "relationship '{}' does not have the required target mode",
                                relationship.r_id
                            )));
                        }
                        let removed = owner_relationships.remove(&relationship.r_id);
                        debug_assert!(removed.is_some());
                        group_changed = true;
                    },
                }
            }
            if !group_changed {
                group_start = group_end;
                continue;
            }
            let relationship_count = owner_relationships.len();
            let relationship_uri = owner.rels_uri().map_err(OpcError::InvalidPackUri)?;
            let member_name = relationship_uri.membername().to_owned();
            self.limits.check(
                ReadResource::ArchiveMemberNameBytes,
                member_name.len() as u64,
                self.limits.max_archive_member_name_bytes(),
            )?;
            if additions.iter().any(|addition| {
                addition
                    .partname
                    .membername()
                    .eq_ignore_ascii_case(&member_name)
            }) {
                return Err(overlay_unavailable(
                    "a generated relationships member collides with a new Part",
                ));
            }
            let existing_entry =
                self.source_entry_id_case_insensitive(&member_name, &physical_members)?;
            let xml = owner_relationships.to_xml().into_bytes();
            self.limits.check(
                ReadResource::RelationshipXmlBytes,
                xml.len() as u64,
                self.limits.max_relationship_xml_bytes() as u64,
            )?;
            validate_overlay_xml(relationship_uri.as_str(), &xml)?;
            if let Some(entry_id) = existing_entry {
                let original = self.archive.read_entry(entry_id).map_err(|error| {
                    if let Some(execution) = take_source_execution_failure(&self.source) {
                        map_execution_error(execution)
                    } else {
                        OpcError::from(error)
                    }
                })?;
                self.source.ensure_current()?;
                let canonical = if owner.as_str() == PACKAGE_URI {
                    self.package_relationships.to_xml()
                } else if let Some(index) = owner_index {
                    self.parts[index].relationships.to_xml()
                } else {
                    Relationships::for_source(&owner).to_xml()
                };
                if original.as_slice() != canonical.as_bytes() {
                    return Err(overlay_unavailable(format!(
                        "existing relationships member '{}' is not canonical",
                        member_name
                    )));
                }
            }
            relationship_publications.push(TopologyRelationshipPublication {
                member_name,
                xml: Arc::new(xml),
                existing_entry,
                relationship_count,
            });
            group_start = group_end;
        }

        relationship_publications
            .sort_unstable_by(|left, right| left.member_name.cmp(&right.member_name));
        if relationship_publications.windows(2).any(|pair| {
            pair[0]
                .member_name
                .eq_ignore_ascii_case(&pair[1].member_name)
        }) {
            return Err(overlay_unavailable(
                "multiple topology relationship owners resolve to one member",
            ));
        }
        let new_relationship_member_count = relationship_publications
            .iter()
            .filter(|publication| publication.existing_entry.is_none())
            .count();
        let new_archive_members = additions
            .len()
            .checked_add(new_relationship_member_count)
            .ok_or_else(|| overlay_unavailable("topology archive member count overflows usize"))?;
        self.limits.check(
            ReadResource::ArchiveMembers,
            self.archive
                .len()
                .checked_add(new_archive_members)
                .ok_or_else(|| {
                    overlay_unavailable("topology archive member count overflows usize")
                })? as u64,
            self.limits.max_archive_members() as u64,
        )?;
        let mut existing_relationship_members = 0usize;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            if is_relationship_member_name(name) {
                existing_relationship_members = existing_relationship_members
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("relationship member count overflows"))?;
            }
        }
        self.limits.check(
            ReadResource::RelationshipParts,
            existing_relationship_members
                .checked_add(new_relationship_member_count)
                .ok_or_else(|| overlay_unavailable("relationship member count overflows usize"))?
                as u64,
            self.limits.max_relationship_parts() as u64,
        )?;
        let mut relationship_count = self.package_relationships.len();
        for (index, part) in self.parts.iter().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            relationship_count = relationship_count
                .checked_add(part.relationships.len())
                .ok_or_else(|| overlay_unavailable("relationship count overflows usize"))?;
        }
        relationship_count = relationship_count
            .checked_add(relationships.len())
            .ok_or_else(|| overlay_unavailable("relationship count overflows usize"))?;
        self.limits.check(
            ReadResource::TotalRelationships,
            relationship_count as u64,
            self.limits.max_total_relationships() as u64,
        )?;
        for (index, publication) in relationship_publications.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            self.limits.check(
                ReadResource::RelationshipsPerPart,
                publication.relationship_count as u64,
                self.limits.max_relationships_per_part() as u64,
            )?;
        }

        // Preserve the raw manifest and add only missing overrides. The source
        // catalog deliberately does not retain this potentially large stream:
        // publication re-reads the exact source member after freshness checks,
        // so managed opens do not carry an uncharged manifest allocation.
        let (
            content_types_xml,
            _content_types_reservation,
            _source_content_types_memory,
            source_content_types,
        ) = if additions.is_empty() {
            (None, None, None, None)
        } else {
            let (xml, reservation) = self.read_content_types_xml()?;
            let parsed_memory = self.reserve_topology_memory(xml.len() as u64)?;
            let map = ContentTypeMap::from_xml(&xml, self.limits)?;
            (Some(xml), reservation, parsed_memory, Some(map))
        };
        let mut required_content_type_overrides = Vec::new();
        required_content_type_overrides
            .try_reserve_exact(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology content-type overrides",
                source,
            })?;
        for (index, addition) in additions.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            let source_content_types = source_content_types.as_ref().ok_or_else(|| {
                overlay_unavailable("content-types catalog is unavailable for a new Part")
            })?;
            if let Some(existing) = source_content_types.override_for(&addition.partname) {
                if existing.as_str() != addition.content_type.as_str() {
                    return Err(overlay_unavailable(format!(
                        "existing content-type override for '{}' conflicts",
                        addition.partname
                    )));
                }
                continue;
            }
            if source_content_types
                .get(&addition.partname)
                .is_ok_and(|existing| existing == addition.content_type.as_str())
            {
                continue;
            }
            required_content_type_overrides.push((
                addition.partname.clone(),
                addition.content_type.as_str().to_string(),
            ));
        }
        required_content_type_overrides
            .sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
        self.check_topology_progress()?;
        let (content_types_replacement, _generated_content_types_memory) =
            if required_content_type_overrides.is_empty() {
                (None, None)
            } else {
                let (xml, generated_memory_reservation) = content_types_with_overrides(
                    content_types_xml.as_deref().ok_or_else(|| {
                        overlay_unavailable("content-types source is unavailable for overrides")
                    })?,
                    &required_content_type_overrides,
                    self.limits,
                    self.cache.context(),
                    self.cache.reservation_failure_counter(),
                )?;
                (Some(Arc::new(xml)), generated_memory_reservation)
            };

        // Check bounded byte resources before auditing authored XML so a
        // payload that violates both limits reports the deterministic policy
        // refusal without first parsing attacker-controlled data.
        self.validate_topology_limits(
            &pending_replacements,
            &additions,
            &relationship_publications,
            content_types_replacement.as_deref().map(Vec::as_slice),
        )?;

        // Materialize only changed existing payloads. Exact no-op replacements
        // still preserve the source member byte-for-byte.
        let mut changed = Vec::new();
        changed
            .try_reserve_exact(
                pending_replacements
                    .len()
                    .checked_add(relationship_publications.len())
                    .and_then(|count| {
                        count.checked_add(usize::from(content_types_replacement.is_some()))
                    })
                    .ok_or_else(|| {
                        overlay_unavailable("topology changed-member count overflows usize")
                    })?,
            )
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology changed members",
                source,
            })?;
        for replacement in &pending_replacements {
            self.check_topology_progress()?;
            let original = self.read_part(replacement.target)?;
            // Compare bytes before any XML audit. A caller may intentionally
            // publish an exact no-op replacement for a source whose Part
            // payload is malformed; the contract is byte-preserving in that
            // case and must not turn the no-op into a parser failure.
            if original.as_bytes() == replacement.replacement.as_slice() {
                continue;
            }
            let part = &self.parts[replacement.target];
            if xml_minifier::audit::package::is_xml_part(part.partname.as_str(), &part.content_type)
            {
                validate_overlay_xml(part.partname.as_str(), original.as_bytes())?;
                validate_overlay_xml(part.partname.as_str(), &replacement.replacement)?;
            }
            changed.push(ChangedOverlay {
                target: ChangedOverlayTarget::Part(replacement.target),
                replacement: Arc::clone(&replacement.replacement),
            });
        }
        for (index, publication) in relationship_publications.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if publication.existing_entry.is_some() {
                changed.push(ChangedOverlay {
                    target: ChangedOverlayTarget::Member(publication.member_name.clone()),
                    replacement: Arc::clone(&publication.xml),
                });
            }
        }
        if let Some(content_types) = &content_types_replacement {
            changed.push(ChangedOverlay {
                target: ChangedOverlayTarget::Member(self.content_types_member.clone()),
                replacement: Arc::clone(content_types),
            });
        }

        if changed.is_empty()
            && additions.is_empty()
            && relationship_publications.is_empty()
            && content_types_replacement.is_none()
        {
            return self.write_exact_source(writer);
        }
        if !self.non_part_members.is_empty() {
            return Err(overlay_unavailable(
                "topology publication refuses non-Part or opaque physical members",
            ));
        }
        if self.archive.has_encrypted_entries() {
            return Err(overlay_unavailable(
                "topology publication refuses encrypted ZIP members",
            ));
        }
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        // Added members are deterministic: Parts first, then relationship
        // members sorted by their owner-derived member name.
        let mut appended = Vec::new();
        appended
            .try_reserve_exact(
                additions
                    .len()
                    .checked_add(new_relationship_member_count)
                    .ok_or_else(|| overlay_unavailable("topology append count overflows usize"))?,
            )
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology appended members",
                source,
            })?;
        for (index, addition) in additions.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if xml_minifier::audit::package::is_xml_part(
                addition.partname.as_str(),
                addition.content_type.as_str(),
            ) {
                validate_overlay_xml(addition.partname.as_str(), &addition.payload)?;
            }
            appended.push(
                soapberry_zip::RegeneratedEntry::new_shared(
                    addition.partname.membername(),
                    Arc::clone(&addition.payload),
                )
                .compression_method(soapberry_zip::CompressionMethod::Deflate),
            );
        }
        for (index, publication) in relationship_publications.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if publication.existing_entry.is_none() {
                appended.push(
                    soapberry_zip::RegeneratedEntry::new_shared(
                        &publication.member_name,
                        Arc::clone(&publication.xml),
                    )
                    .compression_method(soapberry_zip::CompressionMethod::Deflate),
                );
            }
        }
        self.write_changed_overlays_with_appended(writer, &changed, &appended)
    }

    fn topology_relationship_target_ref(
        target: &SourceRelationshipTarget,
        owner_base: &str,
    ) -> Result<(String, TargetMode)> {
        let target_ref = match target {
            SourceRelationshipTarget::Internal(target) => {
                let target_ref = target.relative_ref(owner_base);
                if target_ref.is_empty()
                    || target_ref.chars().any(char::is_control)
                    || target_ref.contains(['?', '#'])
                {
                    return Err(OpcError::InvalidRelationship(
                        "internal relationship target is not a valid relative reference"
                            .to_string(),
                    ));
                }
                target_ref
            },
            SourceRelationshipTarget::External(target) => target.clone(),
        };
        Ok((target_ref, target.mode()))
    }

    fn validate_topology_relationship_field_limits(
        &self,
        r_id: &str,
        reltype: &str,
        target_ref: &str,
    ) -> Result<()> {
        for value in [r_id, reltype, target_ref] {
            self.limits.check(
                ReadResource::XmlAttributeBytes,
                value.len() as u64,
                self.limits.max_xml_attribute_bytes() as u64,
            )?;
            self.limits.check(
                ReadResource::XmlAttributeBytes,
                escaped_xml_attribute_len(value)? as u64,
                self.limits.max_xml_attribute_bytes() as u64,
            )?;
        }
        self.limits.check(
            ReadResource::RelationshipTargetBytes,
            target_ref.len() as u64,
            self.limits.max_relationship_target_bytes() as u64,
        )
    }

    /// Replace one existing ordinary Part and publish to a sequential stream.
    ///
    /// This is an explicit low-level OPC operation. The Part URI, content
    /// type, relationships, package catalog, and physical member topology are
    /// immutable; only the selected payload may change. Every other ZIP member
    /// is raw-copied from the positional source. Unsupported physical layouts
    /// are refused before output instead of silently materializing the package.
    ///
    /// An exact payload no-op copies the complete source artifact byte for
    /// byte, including signatures and unsupported physical details. A real
    /// change to a signed package is refused because this operation accepts no
    /// signature-stripping or resigning policy.
    ///
    /// # Errors
    ///
    /// Returns a typed source, limit, Part, signature, XML-publication, ZIP, or
    /// sink error. If a non-atomic sink accepts bytes before failing, the error
    /// is [`OpcError::IncompleteOutput`].
    pub fn write_part_overlay_to_stream<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Vec<u8>,
    ) -> Result<()> {
        self.write_single_part_overlay_to_stream(writer, partname, replacement, Arc::new, None)
    }

    /// Replace one existing ordinary Part with caller-owned shared bytes and
    /// publish to a sequential stream.
    ///
    /// The [`Arc<Vec<u8>>`] is retained by the bounded publication plan until
    /// the selected ZIP member has been regenerated. This permits a caller
    /// that already owns a shared immutable payload to hand it over without
    /// copying the payload bytes. The same source, limit, validation,
    /// signature, cancellation, and sink semantics as
    /// [`Self::write_part_overlay_to_stream`] apply.
    pub fn write_part_overlay_shared_to_stream<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Arc<Vec<u8>>,
    ) -> Result<()> {
        self.write_single_part_overlay_to_stream(
            writer,
            partname,
            replacement,
            std::convert::identity,
            None,
        )
    }

    /// Replace one existing ordinary Part and publish it while recording
    /// cold ZIP work and accepted output in a caller-owned report.
    ///
    /// The selected source Part is loaded with [`PartView::data_with_accounting`]
    /// semantics: cache hits and same-Part flight waiters contribute no ZIP
    /// counters. Exact payload no-ops use the raw source publication path and
    /// therefore report the exact accepted source/output bytes. Changed
    /// payloads use the preservation path and report raw unchanged plus
    /// generated Store/Deflate payload counters separately from total output.
    /// Ordinary batch overlays, topology publication, `PartWriter`, eager
    /// package writes, and parallel paths remain intentionally unaccounted.
    pub fn write_part_overlay_to_stream_with_accounting<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Vec<u8>,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<()> {
        self.write_single_part_overlay_to_stream(
            writer,
            partname,
            replacement,
            Arc::new,
            Some(accounting),
        )
    }

    /// Publish one existing Part without constructing the general replacement
    /// vector and its intermediate overlay plans.
    ///
    /// The one-Part entry points are used by the format-owned source editors
    /// for the common single-slide/single-cell publication case. Keeping this
    /// path on the same validation and preservation machinery as the bounded
    /// multi-Part path avoids changing semantics while removing short-lived
    /// one-element vectors from the publication setup.
    fn write_single_part_overlay_to_stream<W: Write, P, F>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: P,
        into_shared: F,
        mut accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<()>
    where
        F: FnOnce(P) -> Arc<Vec<u8>>,
    {
        let target = self
            .parts_by_name
            .get(partname)
            .copied()
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
        let replacement = into_shared(replacement);
        self.validate_overlay_limits(std::iter::once((target, replacement.len())))?;

        // Reading the original before any XML audit preserves the exact
        // no-op contract: malformed but byte-identical source payloads still
        // reproduce the source artifact without being parsed.
        let original = match accounting.as_deref_mut() {
            Some(accounting) => self.read_part_with_accounting(target, Some(accounting))?,
            None => self.read_part(target)?,
        };
        if original.as_bytes() == replacement.as_slice() {
            return match accounting {
                Some(accounting) => self.write_exact_source_with_accounting(writer, accounting),
                None => self.write_exact_source(writer),
            };
        }
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        let target_part = &self.parts[target];
        if xml_minifier::audit::package::is_xml_part(
            target_part.partname.as_str(),
            &target_part.content_type,
        ) {
            validate_overlay_xml(target_part.partname.as_str(), original.as_bytes())?;
            validate_overlay_xml(target_part.partname.as_str(), &replacement)?;
        }

        let changed = [ChangedOverlay {
            target: ChangedOverlayTarget::Part(target),
            replacement,
        }];
        match accounting {
            Some(accounting) => self.write_changed_overlays_with_appended_accounting(
                writer,
                &changed,
                &[],
                accounting,
            ),
            None => self.write_changed_overlays(writer, &changed),
        }
    }

    /// Replace one ordinary Part while removing a bounded set of external
    /// relationships owned by that Part.
    ///
    /// This is an explicit, forward-only topology-changing publisher intended
    /// for sanitizers which have already proved that every removed relationship
    /// reference is inside the replacement payload. Relationship IDs must be
    /// unique, must exist on `partname`, and must identify external targets.
    /// The selected Part and its relationships member are regenerated; every
    /// other physical ZIP member is raw-copied.
    ///
    /// # Errors
    ///
    /// Returns before output for a missing, duplicate, internal, oversized, or
    /// signed selection. Sink failures after output begins are reported as
    /// [`OpcError::IncompleteOutput`].
    pub fn write_part_overlay_with_external_relationship_removals_to_stream<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Vec<u8>,
        mut removed_relationship_ids: Vec<String>,
    ) -> Result<()> {
        if removed_relationship_ids.len() > MAX_SOURCE_RELATIONSHIP_REMOVALS {
            return Err(overlay_unavailable(format!(
                "relationship removal set exceeds the {MAX_SOURCE_RELATIONSHIP_REMOVALS}-relationship bound"
            )));
        }
        if removed_relationship_ids.is_empty() {
            return self.write_part_overlay_to_stream(writer, partname, replacement);
        }
        removed_relationship_ids.sort_unstable();
        if let Some(duplicate) = removed_relationship_ids
            .windows(2)
            .find(|pair| pair[0] == pair[1])
        {
            return Err(OpcError::DuplicateRelationshipId(duplicate[0].clone()));
        }

        let target = self
            .parts_by_name
            .get(partname)
            .copied()
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
        let target_part = &self.parts[target];
        for id in &removed_relationship_ids {
            let relationship = target_part.relationships.get(id).ok_or_else(|| {
                OpcError::RelationshipNotFound(format!("relationship '{id}' was not found"))
            })?;
            if !relationship.is_external() {
                return Err(OpcError::InvalidRelationship(format!(
                    "relationship '{id}' is not external"
                )));
            }
        }
        let relationship_xml =
            relationship_xml_without(&target_part.relationships, &removed_relationship_ids)?;
        self.limits.check(
            ReadResource::RelationshipXmlBytes,
            relationship_xml.len() as u64,
            self.limits.max_relationship_xml_bytes() as u64,
        )?;

        let relationship_uri = partname.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let relationship_member = relationship_uri.membername().to_owned();
        let relationship_entry = self.archive.entry_id(&relationship_member).ok_or_else(|| {
            OpcError::RelationshipNotFound(format!(
                "relationships member '{}' was not found",
                relationship_uri.as_str()
            ))
        })?;
        self.validate_overlay_limits(std::iter::once((target, replacement.len())))?;
        self.validate_part_and_relationship_overlay_limits(
            target,
            replacement.len(),
            relationship_entry,
            relationship_xml.len(),
        )?;
        let original_part = self.read_part(target)?;
        let original_relationships = match self.archive.read_entry(relationship_entry) {
            Ok(bytes) => bytes,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&self.source) {
                    return Err(map_execution_error(execution));
                }
                return Err(error.into());
            },
        };
        self.source.ensure_current()?;
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        if xml_minifier::audit::package::is_xml_part(
            target_part.partname.as_str(),
            &target_part.content_type,
        ) {
            validate_overlay_xml(target_part.partname.as_str(), original_part.as_bytes())?;
            validate_overlay_xml(target_part.partname.as_str(), &replacement)?;
        }
        validate_overlay_xml(relationship_uri.as_str(), &original_relationships)?;
        validate_overlay_xml(relationship_uri.as_str(), &relationship_xml)?;

        let changed = [
            ChangedOverlay {
                target: ChangedOverlayTarget::Part(target),
                replacement: Arc::new(replacement),
            },
            ChangedOverlay {
                target: ChangedOverlayTarget::Member(relationship_member),
                replacement: Arc::new(relationship_xml),
            },
        ];
        self.write_changed_overlays(writer, &changed)
    }

    /// Replace several existing Parts while removing bounded sets of
    /// external relationships owned by those Parts.
    ///
    /// The payload and each selected `.rels` member are regenerated together;
    /// every other ZIP member is raw-copied. Relationship IDs must be unique
    /// and must identify external relationships. The caller owns the XML
    /// proof that removed IDs are no longer referenced by replacement
    /// payloads. A real change to a signed source is refused before output.
    pub fn write_part_overlays_with_external_relationship_removals_to_stream<W: Write>(
        self,
        writer: W,
        mut overlays: Vec<(PackURI, Vec<u8>, Vec<String>)>,
    ) -> Result<()> {
        if overlays.len() > MAX_SOURCE_OVERLAY_PARTS {
            return Err(overlay_unavailable(format!(
                "replacement set exceeds the {MAX_SOURCE_OVERLAY_PARTS}-Part bound"
            )));
        }
        if overlays.is_empty() {
            return self.write_exact_source(writer);
        }
        overlays.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
        if let Some(duplicate) = overlays.windows(2).find(|pair| pair[0].0 == pair[1].0) {
            return Err(OpcError::DuplicatePartName(duplicate[0].0.to_string()));
        }

        let mut part_overlays = Vec::new();
        part_overlays
            .try_reserve_exact(overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC relationship overlay parts",
                source,
            })?;
        let mut relationship_overlays = Vec::new();
        relationship_overlays
            .try_reserve_exact(overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC relationship overlay members",
                source,
            })?;
        let mut removed_relationship_total = 0usize;

        for (partname, replacement, mut removed_relationship_ids) in overlays {
            if removed_relationship_ids.len() > MAX_SOURCE_RELATIONSHIP_REMOVALS {
                return Err(overlay_unavailable(format!(
                    "relationship removal set exceeds the {MAX_SOURCE_RELATIONSHIP_REMOVALS}-relationship bound"
                )));
            }
            removed_relationship_ids.sort_unstable();
            if let Some(duplicate) = removed_relationship_ids
                .windows(2)
                .find(|pair| pair[0] == pair[1])
            {
                return Err(OpcError::DuplicateRelationshipId(duplicate[0].clone()));
            }
            removed_relationship_total = removed_relationship_total
                .checked_add(removed_relationship_ids.len())
                .ok_or_else(|| overlay_unavailable("relationship removal count overflows usize"))?;
            if removed_relationship_total > MAX_SOURCE_RELATIONSHIP_REMOVALS {
                return Err(overlay_unavailable(format!(
                    "aggregate relationship removal set exceeds the {MAX_SOURCE_RELATIONSHIP_REMOVALS}-relationship bound"
                )));
            }
            let target = self
                .parts_by_name
                .get(&partname)
                .copied()
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
            let target_part = &self.parts[target];
            if removed_relationship_ids.is_empty() {
                part_overlays.push(PendingOverlay {
                    target,
                    replacement: Arc::new(replacement),
                });
                continue;
            }
            for id in &removed_relationship_ids {
                let relationship = target_part.relationships.get(id).ok_or_else(|| {
                    OpcError::RelationshipNotFound(format!("relationship '{id}' was not found"))
                })?;
                if !relationship.is_external() {
                    return Err(OpcError::InvalidRelationship(format!(
                        "relationship '{id}' is not external"
                    )));
                }
            }
            let relationship_xml =
                relationship_xml_without(&target_part.relationships, &removed_relationship_ids)?;
            self.limits.check(
                ReadResource::RelationshipXmlBytes,
                relationship_xml.len() as u64,
                self.limits.max_relationship_xml_bytes() as u64,
            )?;
            let relationship_uri = partname.rels_uri().map_err(OpcError::InvalidPackUri)?;
            let relationship_member = relationship_uri.membername().to_owned();
            let relationship_entry =
                self.archive.entry_id(&relationship_member).ok_or_else(|| {
                    OpcError::RelationshipNotFound(format!(
                        "relationships member '{}' was not found",
                        relationship_uri.as_str()
                    ))
                })?;
            part_overlays.push(PendingOverlay {
                target,
                replacement: Arc::new(replacement),
            });
            relationship_overlays.push((
                target,
                relationship_entry,
                relationship_member,
                relationship_xml,
            ));
        }

        for (index, left) in relationship_overlays.iter().enumerate() {
            for right in relationship_overlays.iter().skip(index + 1) {
                if left.2 == right.2 {
                    return Err(overlay_unavailable(
                        "multiple relationship overlays resolve to one ZIP member",
                    ));
                }
            }
            for part in &part_overlays {
                if self.parts[part.target].partname.membername() == left.2 {
                    return Err(overlay_unavailable(
                        "a relationship overlay collides with a Part ZIP member",
                    ));
                }
            }
        }

        let mut payload_lengths = Vec::new();
        payload_lengths
            .try_reserve_exact(part_overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC relationship payload limits",
                source,
            })?;
        for overlay in &part_overlays {
            payload_lengths.push((overlay.target, overlay.replacement.len()));
        }
        let mut changed = Vec::new();
        changed
            .try_reserve_exact(
                part_overlays
                    .len()
                    .checked_add(relationship_overlays.len())
                    .and_then(|count| count.checked_mul(2))
                    .ok_or_else(|| {
                        overlay_unavailable("relationship overlay count overflows usize")
                    })?,
            )
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC relationship changed overlays",
                source,
            })?;
        let mut relationship_entries = Vec::new();
        relationship_entries
            .try_reserve_exact(relationship_overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC relationship overlay limits",
                source,
            })?;

        for (target, relationship_entry, _, relationship_xml) in &relationship_overlays {
            let original_relationships = match self.archive.read_entry(*relationship_entry) {
                Ok(bytes) => bytes,
                Err(error) => {
                    if let Some(execution) = take_source_execution_failure(&self.source) {
                        return Err(map_execution_error(execution));
                    }
                    return Err(error.into());
                },
            };
            self.source.ensure_current()?;
            let relationship_uri = self.parts[*target]
                .partname
                .rels_uri()
                .map_err(OpcError::InvalidPackUri)?;
            validate_overlay_xml(relationship_uri.as_str(), original_relationships.as_slice())?;
            validate_overlay_xml(relationship_uri.as_str(), relationship_xml)?;
            relationship_entries.push((*relationship_entry, relationship_xml.len()));
        }
        self.validate_combined_relationship_overlay_limits(
            &payload_lengths,
            &relationship_entries,
        )?;

        for (_, _, relationship_member, relationship_xml) in relationship_overlays {
            changed.push(ChangedOverlay {
                target: ChangedOverlayTarget::Member(relationship_member),
                replacement: Arc::new(relationship_xml),
            });
        }

        for overlay in part_overlays {
            let original = self.read_part(overlay.target)?;
            self.source.ensure_current()?;
            let target_part = &self.parts[overlay.target];
            if xml_minifier::audit::package::is_xml_part(
                target_part.partname.as_str(),
                &target_part.content_type,
            ) {
                validate_overlay_xml(target_part.partname.as_str(), original.as_bytes())?;
                validate_overlay_xml(target_part.partname.as_str(), &overlay.replacement)?;
            }
            if original.as_bytes() != overlay.replacement.as_slice() {
                changed.push(ChangedOverlay {
                    target: ChangedOverlayTarget::Part(overlay.target),
                    replacement: Arc::clone(&overlay.replacement),
                });
            }
        }

        if changed.is_empty() {
            return self.write_exact_source(writer);
        }
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        self.write_changed_overlays(writer, &changed)
    }

    /// Replace a bounded set of existing ordinary Parts and publish to a
    /// sequential stream.
    ///
    /// The replacement set is sorted and checked for duplicate Part URIs. Its
    /// maximum size is 64. Part URIs, content types, relationships, the package
    /// catalog, and physical member topology are immutable. Every unselected
    /// ZIP member and every selected exact no-op member is raw-copied.
    ///
    /// If every replacement is byte-identical to its source payload, the
    /// complete source artifact is copied byte for byte, including signatures
    /// and unsupported physical details. A real change to a signed package is
    /// refused because this operation accepts no signature policy.
    ///
    /// All selected payloads, aggregate limits, signatures, changed XML, and
    /// the preservation plan are validated before the first output byte.
    ///
    /// # Errors
    ///
    /// Returns a typed source, limit, duplicate-Part, Part, signature,
    /// XML-publication, ZIP, or sink error. If a non-atomic sink accepts bytes
    /// before failing, the error is [`OpcError::IncompleteOutput`].
    pub fn write_part_overlays_to_stream<W: Write>(
        self,
        writer: W,
        replacements: Vec<(PackURI, Vec<u8>)>,
    ) -> Result<()> {
        self.write_part_overlays_impl(writer, replacements, Arc::new)
    }

    /// Replace a bounded set of existing ordinary Parts with caller-owned
    /// shared payloads and publish to a sequential stream.
    ///
    /// This is the shared-ownership counterpart to
    /// [`Self::write_part_overlays_to_stream`]. Each [`Arc<Vec<u8>>`] is moved
    /// into the opaque publication plan; no payload-byte copy is performed.
    /// The replacement set is sorted and checked for duplicate Part URIs. Its
    /// maximum size is 64. Part URIs, content types, relationships, the
    /// package catalog, and physical member topology are immutable. Every
    /// unselected ZIP member and every selected exact no-op member is
    /// raw-copied.
    pub fn write_part_overlays_shared_to_stream<W: Write>(
        self,
        writer: W,
        replacements: Vec<(PackURI, Arc<Vec<u8>>)>,
    ) -> Result<()> {
        self.write_part_overlays_impl(writer, replacements, std::convert::identity)
    }

    fn write_part_overlays_impl<W: Write, P, F>(
        self,
        writer: W,
        mut replacements: Vec<(PackURI, P)>,
        mut into_shared: F,
    ) -> Result<()>
    where
        F: FnMut(P) -> Arc<Vec<u8>>,
    {
        if replacements.len() > MAX_SOURCE_OVERLAY_PARTS {
            return Err(overlay_unavailable(format!(
                "replacement set exceeds the {MAX_SOURCE_OVERLAY_PARTS}-Part bound"
            )));
        }
        if replacements.is_empty() {
            return self.write_exact_source(writer);
        }
        replacements.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
        if let Some(duplicate) = replacements.windows(2).find(|pair| pair[0].0 == pair[1].0) {
            return Err(OpcError::DuplicatePartName(duplicate[0].0.to_string()));
        }

        let mut overlays = Vec::new();
        overlays
            .try_reserve_exact(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC replacement plan",
                source,
            })?;
        for (partname, replacement) in replacements {
            let target = self
                .parts_by_name
                .get(&partname)
                .copied()
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
            overlays.push(PendingOverlay {
                target,
                replacement: into_shared(replacement),
            });
        }
        self.validate_overlay_limits(
            overlays
                .iter()
                .map(|overlay| (overlay.target, overlay.replacement.len())),
        )?;

        let mut changed = Vec::new();
        changed
            .try_reserve_exact(overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC changed replacement plan",
                source,
            })?;
        for overlay in overlays {
            // Reading every selected closure proves its local framing,
            // compression, declared size, and CRC before output.
            let original = self.read_part(overlay.target)?;
            if original.as_bytes() != overlay.replacement.as_slice() {
                changed.push((overlay, original));
            }
        }
        if changed.is_empty() {
            return self.write_exact_source(writer);
        }
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(changed.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC changed payloads",
                source,
            })?;
        for (overlay, original) in changed {
            let target_part = &self.parts[overlay.target];
            if xml_minifier::audit::package::is_xml_part(
                target_part.partname.as_str(),
                &target_part.content_type,
            ) {
                validate_overlay_xml(target_part.partname.as_str(), original.as_bytes())?;
                validate_overlay_xml(target_part.partname.as_str(), &overlay.replacement)?;
            }
            replacements.push(ChangedOverlay {
                target: ChangedOverlayTarget::Part(overlay.target),
                replacement: Arc::clone(&overlay.replacement),
            });
        }
        self.write_changed_overlays(writer, &replacements)
    }

    fn validate_overlay_limits<I>(&self, overlays: I) -> Result<()>
    where
        I: Iterator<Item = (usize, usize)> + Clone,
    {
        for (_, replacement_len) in overlays.clone() {
            let replacement_bytes =
                u64::try_from(replacement_len).map_err(|_| OpcError::ReadLimit {
                    resource: ReadResource::PartBytes,
                    actual: u64::MAX,
                    maximum: self.limits.max_part_bytes(),
                })?;
            self.limits.check(
                ReadResource::PartBytes,
                replacement_bytes,
                self.limits.max_part_bytes(),
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                replacement_bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
        }

        let mut part_total = 0_u64;
        let mut archive_total = 0_u64;
        for part in &self.parts {
            let bytes = self
                .archive
                .metadata_for(part.entry_id)?
                .uncompressed_size();
            part_total = checked_overlay_total(
                part_total,
                bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
        }
        for name in self.archive.file_names() {
            archive_total = checked_overlay_total(
                archive_total,
                self.archive.metadata(name)?.uncompressed_size(),
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        let mut adjusted_parts = part_total;
        let mut adjusted_archive = archive_total;
        for (target, replacement_len) in overlays {
            let target_bytes = self
                .archive
                .metadata_for(self.parts[target].entry_id)?
                .uncompressed_size();
            let replacement_bytes = replacement_len as u64;
            adjusted_parts = adjusted_overlay_total(
                adjusted_parts,
                target_bytes,
                replacement_bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
            adjusted_archive = adjusted_overlay_total(
                adjusted_archive,
                target_bytes,
                replacement_bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        self.limits.check(
            ReadResource::TotalPartBytes,
            adjusted_parts,
            self.limits.max_total_part_bytes(),
        )?;
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            adjusted_archive,
            self.limits.max_archive_total_bytes(),
        )?;
        Ok(())
    }

    fn validate_part_and_relationship_overlay_limits(
        &self,
        target: usize,
        replacement_len: usize,
        relationship_entry: EntryId,
        relationship_len: usize,
    ) -> Result<()> {
        let relationship_bytes = relationship_len as u64;
        self.limits.check(
            ReadResource::ArchiveEntryBytes,
            relationship_bytes,
            self.limits.max_archive_entry_bytes(),
        )?;
        self.limits.check(
            ReadResource::RelationshipXmlBytes,
            relationship_bytes,
            self.limits.max_relationship_xml_bytes() as u64,
        )?;

        let mut archive_total = 0_u64;
        let mut relationship_total = 0_u64;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            let bytes = self.archive.metadata(name)?.uncompressed_size();
            archive_total = checked_overlay_total(
                archive_total,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            if is_relationship_member_name(name) {
                relationship_total = checked_overlay_total(
                    relationship_total,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
            }
        }
        let original_part = self
            .archive
            .metadata_for(self.parts[target].entry_id)?
            .uncompressed_size();
        let original_relationship = self
            .archive
            .metadata_for(relationship_entry)?
            .uncompressed_size();
        let adjusted_archive = adjusted_overlay_total(
            archive_total,
            original_part,
            replacement_len as u64,
            ReadResource::ArchiveTotalBytes,
            self.limits.max_archive_total_bytes(),
        )?;
        let adjusted_archive = adjusted_overlay_total(
            adjusted_archive,
            original_relationship,
            relationship_bytes,
            ReadResource::ArchiveTotalBytes,
            self.limits.max_archive_total_bytes(),
        )?;
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            adjusted_archive,
            self.limits.max_archive_total_bytes(),
        )?;
        let adjusted_relationships = adjusted_overlay_total(
            relationship_total,
            original_relationship,
            relationship_bytes,
            ReadResource::TotalRelationshipXmlBytes,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            adjusted_relationships,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )
    }

    fn validate_combined_relationship_overlay_limits(
        &self,
        part_overlays: &[(usize, usize)],
        relationship_overlays: &[(EntryId, usize)],
    ) -> Result<()> {
        for (_, replacement_len) in part_overlays {
            let replacement_bytes = *replacement_len as u64;
            self.limits.check(
                ReadResource::PartBytes,
                replacement_bytes,
                self.limits.max_part_bytes(),
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                replacement_bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
        }
        for (_, replacement_len) in relationship_overlays {
            let replacement_bytes = *replacement_len as u64;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                replacement_bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
            self.limits.check(
                ReadResource::RelationshipXmlBytes,
                replacement_bytes,
                self.limits.max_relationship_xml_bytes() as u64,
            )?;
        }
        let mut archive_total = 0_u64;
        let mut part_total = 0_u64;
        let mut relationship_total = 0_u64;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            let bytes = self.archive.metadata(name)?.uncompressed_size();
            archive_total = checked_overlay_total(
                archive_total,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            if is_relationship_member_name(name) {
                relationship_total = checked_overlay_total(
                    relationship_total,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
            }
        }
        for (index, part) in self.parts.iter().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            part_total = checked_overlay_total(
                part_total,
                self.archive
                    .metadata_for(part.entry_id)?
                    .uncompressed_size(),
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
        }
        let mut adjusted_archive = archive_total;
        let mut adjusted_parts = part_total;
        let mut adjusted_relationships = relationship_total;
        for (target, replacement_len) in part_overlays {
            let original = self
                .archive
                .metadata_for(self.parts[*target].entry_id)?
                .uncompressed_size();
            adjusted_archive = adjusted_overlay_total(
                adjusted_archive,
                original,
                *replacement_len as u64,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            adjusted_parts = adjusted_overlay_total(
                adjusted_parts,
                original,
                *replacement_len as u64,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
        }
        for (entry, replacement_len) in relationship_overlays {
            let original = self.archive.metadata_for(*entry)?.uncompressed_size();
            adjusted_archive = adjusted_overlay_total(
                adjusted_archive,
                original,
                *replacement_len as u64,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            adjusted_relationships = adjusted_overlay_total(
                adjusted_relationships,
                original,
                *replacement_len as u64,
                ReadResource::TotalRelationshipXmlBytes,
                self.limits.max_total_relationship_xml_bytes() as u64,
            )?;
        }
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            adjusted_archive,
            self.limits.max_archive_total_bytes(),
        )?;
        self.limits.check(
            ReadResource::TotalPartBytes,
            adjusted_parts,
            self.limits.max_total_part_bytes(),
        )?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            adjusted_relationships,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )
    }

    fn validate_topology_limits(
        &self,
        replacements: &[PendingOverlay],
        additions: &[TopologyPartAddition],
        relationship_publications: &[TopologyRelationshipPublication],
        content_types_replacement: Option<&[u8]>,
    ) -> Result<()> {
        let mut part_total = 0_u64;
        let mut archive_total = 0_u64;
        let mut relationship_total = 0_u64;
        let mut relationship_event_total = 0_u64;
        let mut source_relationship_events = HashMap::new();
        source_relationship_events
            .try_reserve(relationship_publications.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC topology relationship event counts",
                source,
            })?;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            let bytes = self.archive.metadata(name)?.uncompressed_size();
            archive_total = checked_overlay_total(
                archive_total,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            if is_relationship_member_name(name) {
                relationship_total = checked_overlay_total(
                    relationship_total,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
                let entry = self
                    .archive
                    .entry_id(name)
                    .ok_or_else(|| OpcError::PartNotFound(name.to_string()))?;
                let xml = self.archive.read_entry(entry).map_err(|error| {
                    if let Some(execution) = take_source_execution_failure(&self.source) {
                        map_execution_error(execution)
                    } else {
                        OpcError::from(error)
                    }
                })?;
                self.source.ensure_current()?;
                let events = relationship_xml_event_count(&xml, self.limits)?;
                relationship_event_total = checked_overlay_total(
                    relationship_event_total,
                    events,
                    ReadResource::TotalRelationshipXmlEvents,
                    self.limits.max_total_relationship_xml_events() as u64,
                )?;
                if relationship_publications
                    .iter()
                    .any(|publication| publication.existing_entry == Some(entry))
                {
                    source_relationship_events.insert(entry, events);
                }
            }
        }
        for (index, part) in self.parts.iter().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            part_total = checked_overlay_total(
                part_total,
                self.archive
                    .metadata_for(part.entry_id)?
                    .uncompressed_size(),
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
        }
        for (index, replacement) in replacements.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            let replacement_bytes = u64::try_from(replacement.replacement.len())
                .map_err(|_| overlay_unavailable("topology replacement length overflows u64"))?;
            self.limits.check(
                ReadResource::PartBytes,
                replacement_bytes,
                self.limits.max_part_bytes(),
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                replacement_bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
            let original = self
                .archive
                .metadata_for(self.parts[replacement.target].entry_id)?
                .uncompressed_size();
            part_total = adjusted_overlay_total(
                part_total,
                original,
                replacement_bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
            archive_total = adjusted_overlay_total(
                archive_total,
                original,
                replacement_bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        for (index, addition) in additions.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            let bytes = u64::try_from(addition.payload.len())
                .map_err(|_| overlay_unavailable("topology Part length overflows u64"))?;
            self.limits.check(
                ReadResource::ArchiveMemberNameBytes,
                addition.partname.membername().len() as u64,
                self.limits.max_archive_member_name_bytes(),
            )?;
            self.limits
                .check(ReadResource::PartBytes, bytes, self.limits.max_part_bytes())?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
            part_total = checked_overlay_total(
                part_total,
                bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
            archive_total = checked_overlay_total(
                archive_total,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        if let Some(bytes) = content_types_replacement {
            let bytes = u64::try_from(bytes.len())
                .map_err(|_| overlay_unavailable("content-types replacement overflows u64"))?;
            self.limits.check(
                ReadResource::ContentTypesBytes,
                bytes,
                self.limits.max_content_types_bytes() as u64,
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
            let entry = self
                .archive
                .entry_id(&self.content_types_member)
                .ok_or_else(|| OpcError::PartNotFound(self.content_types_member.clone()))?;
            let original = self.archive.metadata_for(entry)?.uncompressed_size();
            archive_total = adjusted_overlay_total(
                archive_total,
                original,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        for (index, publication) in relationship_publications.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            let bytes = u64::try_from(publication.xml.len())
                .map_err(|_| overlay_unavailable("relationship XML length overflows u64"))?;
            self.limits.check(
                ReadResource::ArchiveMemberNameBytes,
                publication.member_name.len() as u64,
                self.limits.max_archive_member_name_bytes(),
            )?;
            let events = relationship_xml_event_count(&publication.xml, self.limits)?;
            self.limits.check(
                ReadResource::RelationshipXmlBytes,
                bytes,
                self.limits.max_relationship_xml_bytes() as u64,
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
            if let Some(entry) = publication.existing_entry {
                let original = self.archive.metadata_for(entry)?.uncompressed_size();
                archive_total = adjusted_overlay_total(
                    archive_total,
                    original,
                    bytes,
                    ReadResource::ArchiveTotalBytes,
                    self.limits.max_archive_total_bytes(),
                )?;
                relationship_total = adjusted_overlay_total(
                    relationship_total,
                    original,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
                let original_events = source_relationship_events.get(&entry).ok_or_else(|| {
                    overlay_unavailable("source relationship event count is unavailable")
                })?;
                relationship_event_total = adjusted_overlay_total(
                    relationship_event_total,
                    *original_events,
                    events,
                    ReadResource::TotalRelationshipXmlEvents,
                    self.limits.max_total_relationship_xml_events() as u64,
                )?;
            } else {
                archive_total = checked_overlay_total(
                    archive_total,
                    bytes,
                    ReadResource::ArchiveTotalBytes,
                    self.limits.max_archive_total_bytes(),
                )?;
                relationship_total = checked_overlay_total(
                    relationship_total,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
                relationship_event_total = checked_overlay_total(
                    relationship_event_total,
                    events,
                    ReadResource::TotalRelationshipXmlEvents,
                    self.limits.max_total_relationship_xml_events() as u64,
                )?;
            }
        }
        self.limits.check(
            ReadResource::TotalPartBytes,
            part_total,
            self.limits.max_total_part_bytes(),
        )?;
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            archive_total,
            self.limits.max_archive_total_bytes(),
        )?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            relationship_total,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlEvents,
            relationship_event_total,
            self.limits.max_total_relationship_xml_events() as u64,
        )?;

        let mut output_bound = self.source.length;
        for (index, replacement) in replacements.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            output_bound = output_bound
                .checked_add((replacement.replacement.len() as u64).saturating_mul(2))
                .ok_or_else(|| overlay_unavailable("topology output bound overflows u64"))?;
        }
        for (index, addition) in additions.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            output_bound = output_bound
                .checked_add(addition.payload.len() as u64)
                .and_then(|value| value.checked_add(4096))
                .ok_or_else(|| overlay_unavailable("topology output bound overflows u64"))?;
        }
        for (index, publication) in relationship_publications.iter().enumerate() {
            if index & 0x3f == 0 {
                self.check_topology_progress()?;
            }
            if publication.existing_entry.is_none() {
                output_bound = output_bound
                    .checked_add(publication.xml.len() as u64)
                    .and_then(|value| value.checked_add(4096))
                    .ok_or_else(|| overlay_unavailable("topology output bound overflows u64"))?;
            }
        }
        if let Some(bytes) = content_types_replacement {
            output_bound = output_bound
                .checked_add((bytes.len() as u64).saturating_mul(2))
                .ok_or_else(|| overlay_unavailable("topology output bound overflows u64"))?;
        }
        if output_bound > u64::from(u32::MAX) {
            return Err(overlay_unavailable(
                "topology publication may require ZIP64 output",
            ));
        }
        Ok(())
    }

    fn has_signature_infrastructure(&self) -> bool {
        self.package_relationships
            .iter()
            .any(is_signature_relationship_or_target)
            || self.parts.iter().any(|part| {
                is_signature_path(part.partname.as_str())
                    || is_signature_content_type(&part.content_type)
                    || part
                        .relationships
                        .iter()
                        .any(is_signature_relationship_or_target)
            })
    }

    /// Read the source content-types member at publication time.
    ///
    /// The source catalog parses this stream during open but does not retain a
    /// second managed-memory charge for it. Topology publication needs the
    /// original lexical bytes, so it re-reads the one catalogued member after
    /// checking source freshness and the managed execution policy.
    fn read_content_types_xml(&self) -> Result<(Vec<u8>, Option<Arc<Reservation>>)> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        let entry = self
            .archive
            .entry_id(&self.content_types_member)
            .ok_or_else(|| OpcError::PartNotFound(self.content_types_member.clone()))?;
        let metadata = self.archive.metadata_for(entry)?;
        let declared_bytes = metadata.uncompressed_size();
        self.limits.check(
            ReadResource::ContentTypesBytes,
            declared_bytes,
            self.limits.max_content_types_bytes() as u64,
        )?;
        let memory_reservation = self.reserve_topology_memory(declared_bytes)?;
        let bytes = self.archive.read_entry(entry).map_err(|error| {
            if let Some(execution) = take_source_execution_failure(&self.source) {
                map_execution_error(execution)
            } else {
                OpcError::from(error)
            }
        })?;
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        self.limits.check(
            ReadResource::ContentTypesBytes,
            bytes.len() as u64,
            self.limits.max_content_types_bytes() as u64,
        )?;
        if bytes.len() as u64 != declared_bytes {
            return Err(OpcError::ZipError(format!(
                "source-backed OPC content-types member declared {declared_bytes} uncompressed bytes but read {}",
                bytes.len()
            )));
        }
        Ok((bytes, memory_reservation))
    }

    fn build_physical_member_lookup(
        &self,
    ) -> Result<(PhysicalMemberLookup, Option<Arc<Reservation>>)> {
        let mut total_name_bytes = 0_u64;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            total_name_bytes = total_name_bytes
                .checked_add(name.len() as u64)
                .ok_or_else(|| overlay_unavailable("physical member name bytes overflow"))?;
        }
        let member_count = self.archive.len() as u64;
        // Folded String allocations and the hash table are retained for the
        // duration of publication. Charge a conservative bound before
        // allocating either: two copies of name bytes plus fixed per-entry
        // table/String metadata.
        let memory_bound = total_name_bytes
            .checked_mul(2)
            .and_then(|bytes| bytes.checked_add(member_count.checked_mul(128)?))
            .ok_or_else(|| overlay_unavailable("physical member lookup memory overflows"))?;
        let memory_reservation = self.reserve_topology_memory(memory_bound)?;
        let mut by_folded_name: HashMap<String, PhysicalMemberInfo> = HashMap::new();
        by_folded_name
            .try_reserve(self.archive.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC physical member lookup",
                source,
            })?;
        for (index, name) in self.archive.file_names().enumerate() {
            if index & 0xff == 0 {
                self.check_topology_progress()?;
            }
            let key = folded_ascii_name(name, "source-backed OPC folded physical member name")?;
            let entry_id = self.archive.entry_id(name).ok_or_else(|| {
                overlay_unavailable("indexed physical member disappeared during topology lookup")
            })?;
            if let Some(info) = by_folded_name.get_mut(&key) {
                info.count = info
                    .count
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("physical member count overflows"))?;
            } else {
                by_folded_name.insert(key, PhysicalMemberInfo { entry_id, count: 1 });
            }
        }
        self.check_topology_progress()?;
        Ok((PhysicalMemberLookup { by_folded_name }, memory_reservation))
    }

    fn source_entry_id_case_insensitive(
        &self,
        member_name: &str,
        physical_members: &PhysicalMemberLookup,
    ) -> Result<Option<EntryId>> {
        if let Some(entry) = self.archive.entry_id(member_name) {
            return Ok(Some(entry));
        }
        let key = folded_ascii_name(member_name, "source-backed OPC folded physical member name")?;
        Ok(physical_members
            .by_folded_name
            .get(&key)
            .map(|info| info.entry_id))
    }

    fn check_topology_progress(&self) -> Result<()> {
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)
    }

    fn reserve_topology_memory(&self, bytes: u64) -> Result<Option<Arc<Reservation>>> {
        let Some(context) = self.cache.context() else {
            return Ok(None);
        };
        let reservation = context.reserve(Resource::Memory, bytes).map_err(|error| {
            if matches!(error, ExecutionError::ResourceLimit(_)) {
                self.cache.record_budget_reservation_failure();
            }
            map_execution_error(error)
        })?;
        Ok(Some(Arc::new(reservation)))
    }

    fn write_exact_source<W: Write>(self, writer: W) -> Result<()> {
        write_exact_snapshot(&self.source, writer, self.cache.context())
    }

    fn write_exact_source_with_accounting<W: Write>(
        self,
        writer: W,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<()> {
        write_exact_snapshot_with_accounting(&self.source, writer, self.cache.context(), accounting)
    }

    fn write_changed_overlays<W: Write>(
        self,
        writer: W,
        replacements: &[ChangedOverlay],
    ) -> Result<()> {
        self.write_changed_overlays_with_appended(writer, replacements, &[])
    }

    fn write_changed_overlays_with_appended<W: Write>(
        self,
        writer: W,
        replacements: &[ChangedOverlay],
        appended: &[soapberry_zip::RegeneratedEntry],
    ) -> Result<()> {
        self.write_changed_overlays_with_appended_inner(writer, replacements, appended, None)
    }

    fn write_changed_overlays_with_appended_accounting<W: Write>(
        self,
        writer: W,
        replacements: &[ChangedOverlay],
        appended: &[soapberry_zip::RegeneratedEntry],
        accounting: &mut OpcOperationAccounting,
    ) -> Result<()> {
        self.write_changed_overlays_with_appended_inner(
            writer,
            replacements,
            appended,
            Some(accounting),
        )
    }

    fn write_changed_overlays_with_appended_inner<W: Write>(
        self,
        writer: W,
        replacements: &[ChangedOverlay],
        appended: &[soapberry_zip::RegeneratedEntry],
        mut accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<()> {
        self.source.monitor_publication();
        self.source.ensure_current()?;
        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC preservation index",
                source,
            })?;
        scratch.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let index = match self.archive.preservation_index(&mut scratch) {
            Ok(index) => index,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&self.source) {
                    return Err(map_execution_error(execution));
                }
                self.source.ensure_current()?;
                return Err(overlay_unavailable(error.to_string()));
            },
        };
        self.source.ensure_current()?;
        if index.archive_end_offset() != self.source.length {
            return Err(overlay_unavailable(
                "source ZIP archive has trailing bytes outside its located archive",
            ));
        }

        // Build a sorted keyed lookup once. The previous per-entry linear
        // search made publication O(archive_members * replacements), which is
        // avoidable even though the replacement set is bounded.
        let mut replacement_lookup: Vec<(&[u8], &ChangedOverlay)> = Vec::new();
        replacement_lookup
            .try_reserve_exact(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC replacement lookup",
                source,
            })?;
        for replacement in replacements {
            let target_name = replacement_member_name(replacement, &self.parts);
            replacement_lookup.push((target_name.as_bytes(), replacement));
        }
        replacement_lookup.sort_unstable_by(|left, right| left.0.cmp(right.0));
        if replacement_lookup
            .windows(2)
            .any(|pair| pair[0].0 == pair[1].0)
        {
            return Err(overlay_unavailable(
                "multiple changed overlays resolve to one source member",
            ));
        }
        let mut replacement_entry_counts: Vec<usize> = Vec::new();
        replacement_entry_counts
            .try_reserve_exact(replacement_lookup.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC replacement entry counts",
                source,
            })?;
        replacement_entry_counts.resize(replacement_lookup.len(), 0);
        for (entry_index, entry) in index.entries().iter().enumerate() {
            if entry_index & 0xff == 0 {
                self.source.ensure_current()?;
                self.cache.check_context().map_err(map_execution_error)?;
            }
            if let Ok(replacement_index) = replacement_lookup
                .binary_search_by(|candidate| candidate.0.cmp(entry.raw_name_bytes()))
            {
                replacement_entry_counts[replacement_index] = replacement_entry_counts
                    [replacement_index]
                    .checked_add(1)
                    .ok_or_else(|| {
                        overlay_unavailable("replacement entry count overflows usize")
                    })?;
            }
        }
        let mut replacement_bytes = 0_u64;
        for (replacement_index, (_, replacement)) in replacement_lookup.iter().enumerate() {
            if replacement_entry_counts[replacement_index] != 1 {
                return Err(overlay_unavailable(
                    "selected Part does not have one canonical UTF-8 source member",
                ));
            }
            replacement_bytes = replacement_bytes
                .checked_add(replacement.replacement.len() as u64)
                .ok_or_else(|| overlay_unavailable("replacement byte total overflows u64"))?;
        }
        let appended_bytes = (appended.len() as u64)
            .checked_mul(4096)
            .ok_or_else(|| overlay_unavailable("appended member size overflows u64"))?;
        let conservative_output_bound = self
            .source
            .length
            .checked_add(replacement_bytes.saturating_mul(2))
            .and_then(|bytes| bytes.checked_add(appended_bytes))
            .and_then(|bytes| bytes.checked_add(SOURCE_PUBLICATION_CHUNK_BYTES as u64));
        if conservative_output_bound.is_none_or(|bytes| bytes > u64::from(u32::MAX)) {
            return Err(overlay_unavailable(
                "selected Part replacement may require ZIP64 output",
            ));
        }

        let mut plan = soapberry_zip::PreservationPlan::new();
        plan.try_reserve_exact(index.entries().len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC preservation actions",
                source,
            })?;
        plan.try_reserve_appended(appended.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC preservation appended actions",
                source,
            })?;
        for (entry_index, entry) in index.entries().iter().enumerate() {
            if entry_index & 0xff == 0 {
                self.source.ensure_current()?;
                self.cache.check_context().map_err(map_execution_error)?;
            }
            if let Ok(replacement_index) = replacement_lookup
                .binary_search_by(|candidate| candidate.0.cmp(entry.raw_name_bytes()))
            {
                let replacement = replacement_lookup[replacement_index].1;
                let target_name = replacement_member_name(replacement, &self.parts);
                plan.push(soapberry_zip::PreservationAction::Regenerate {
                    id: entry.id(),
                    entry: soapberry_zip::RegeneratedEntry::new_shared(
                        target_name,
                        Arc::clone(&replacement.replacement),
                    )
                    .compression_method(entry.compression_method()),
                });
            } else {
                plan.push(soapberry_zip::PreservationAction::Copy(entry.id()));
            }
        }
        for entry in appended {
            plan.try_append(entry.clone())
                .map_err(|source| OpcError::Allocation {
                    resource: "source-backed OPC preservation appended member",
                    source,
                })?;
        }

        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        let mut written = 0_u64;
        let mut zip_accounting = LowLevelZipOperationAccounting::default();
        let mut accounting_error = None;
        let has_accounting = accounting.is_some();
        let execution_failure = Arc::new(Mutex::new(None));
        let result = if let Some(context) = self.cache.context() {
            let Some(output_reservation_failures) =
                self.source.output_reservation_failures.as_ref()
            else {
                return Err(overlay_unavailable(
                    "managed source output reservation counter is unavailable",
                ));
            };
            let counted = match accounting.as_deref_mut() {
                Some(report) => Counted::with_accounting(
                    writer,
                    &mut written,
                    report,
                    &mut accounting_error,
                    false,
                ),
                None => Counted::new(writer, &mut written),
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: self.source.clone(),
            };
            let cooperative = ContextCheckedSink {
                inner: checked,
                context: Some(context.clone()),
                failure: Arc::clone(&execution_failure),
            };
            let budgeted = OutputBudgetedSink {
                inner: cooperative,
                context: context.clone(),
                failure: Arc::clone(&execution_failure),
                output_reservation_failures: output_reservation_failures.clone(),
            };
            if has_accounting {
                match index.write_to_with_accounting(
                    &plan,
                    Chunked { inner: budgeted },
                    &mut zip_accounting,
                ) {
                    Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                    Err(error) => Err(map_preservation_error(error)),
                }
            } else {
                match index.write_to(&plan, Chunked { inner: budgeted }) {
                    Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                    Err(error) => Err(map_preservation_error(error)),
                }
            }
        } else {
            let counted = match accounting.as_deref_mut() {
                Some(report) => Counted::with_accounting(
                    writer,
                    &mut written,
                    report,
                    &mut accounting_error,
                    false,
                ),
                None => Counted::new(writer, &mut written),
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: self.source.clone(),
            };
            let cooperative = ContextCheckedSink {
                inner: checked,
                context: None,
                failure: Arc::clone(&execution_failure),
            };
            if has_accounting {
                match index.write_to_with_accounting(
                    &plan,
                    Chunked { inner: cooperative },
                    &mut zip_accounting,
                ) {
                    Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                    Err(error) => Err(map_preservation_error(error)),
                }
            } else {
                match index.write_to(&plan, Chunked { inner: cooperative }) {
                    Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                    Err(error) => Err(map_preservation_error(error)),
                }
            }
        };
        let result = execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = take_source_execution_failure(&self.source)
            .map_or(result, |error| Err(map_execution_error(error)));
        let merge_result = accounting.map_or(Ok(()), |report| report.merge_zip(&zip_accounting));
        // Merge the operation-local ZIP report before the final source
        // decision. `finish_source_publication` must remain authoritative for
        // SourceChanged and IncompleteOutput, even when accounting also fails.
        let result = finish_source_publication(result, &self.source, written);
        match result {
            Err(error) => Err(error),
            Ok(()) => accounting_error.map_or(merge_result, Err),
        }
    }

    fn read_part(&self, index: usize) -> Result<PartData> {
        self.read_part_with_accounting(index, None)
    }

    fn read_part_with_accounting(
        &self,
        index: usize,
        accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<PartData> {
        self.cache.check_context().map_err(map_execution_error)?;
        let entry_id = self
            .parts
            .get(index)
            .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?
            .entry_id;
        let declared_bytes = if self.cache.is_managed() {
            let declared = self.archive.metadata_for(entry_id)?.uncompressed_size();
            self.limits.check(
                ReadResource::PartBytes,
                declared,
                self.limits.max_part_bytes(),
            )?;
            Some(declared)
        } else {
            None
        };
        loop {
            self.source.ensure_current()?;
            match self
                .cache
                .enter(entry_id, declared_bytes.unwrap_or_default())
                .map_err(map_execution_error)?
            {
                CacheAccess::Hit(bytes) => {
                    self.source.ensure_current()?;
                    self.cache.check_context().map_err(map_execution_error)?;
                    return Ok(PartData { payload: bytes });
                },
                CacheAccess::Waiter(flight) => {
                    if let Some(payload) = flight
                        .wait(self.cache.context())
                        .map_err(map_execution_error)?
                    {
                        self.source.ensure_current()?;
                        self.cache.check_context().map_err(map_execution_error)?;
                        return Ok(PartData { payload });
                    }
                    // The loader may have failed; in that case the flight is
                    // removed and this caller retries rather than observing a
                    // retained error. This also re-checks source freshness.
                },
                CacheAccess::Loader(flight) => {
                    return self.load_part_with_accounting(
                        index,
                        entry_id,
                        declared_bytes,
                        Some(flight),
                        None,
                        accounting,
                    );
                },
                CacheAccess::Bypass(reservation) => {
                    return self.load_part_with_accounting(
                        index,
                        entry_id,
                        declared_bytes,
                        None,
                        Some(reservation),
                        accounting,
                    );
                },
            }
        }
    }

    fn load_part_with_accounting(
        &self,
        index: usize,
        entry_id: EntryId,
        declared_bytes: Option<u64>,
        flight: Option<Arc<LoadFlight>>,
        bypass_resources: Option<LoadResources>,
        mut accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<PartData> {
        let mut zip_accounting = LowLevelZipOperationAccounting::default();
        let result = (|| {
            let part = self
                .parts
                .get(index)
                .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?;
            let bytes = match match accounting.as_deref_mut() {
                Some(_) => self
                    .archive
                    .read_entry_with_accounting(part.entry_id, &mut zip_accounting),
                None => self.archive.read_entry(part.entry_id),
            } {
                Ok(bytes) => bytes,
                Err(error) => {
                    if let Some(execution) = take_source_execution_failure(&self.source) {
                        return Err(map_execution_error(execution));
                    }
                    return Err(error.into());
                },
            };
            // The decompressor has finished and no payload has been
            // published yet. Cancellation here discards the cold result.
            self.cache.check_context().map_err(map_execution_error)?;
            if let Some(declared) = declared_bytes {
                if bytes.len() as u64 != declared {
                    return Err(OpcError::ZipError(format!(
                        "source-backed OPC Part declared {declared} uncompressed bytes but read {}",
                        bytes.len()
                    )));
                }
            }
            self.source.ensure_current()?;
            self.limits.check(
                ReadResource::PartBytes,
                bytes.len() as u64,
                self.limits.max_part_bytes(),
            )?;
            // Check immediately before publishing. If the source changed
            // during the cold read, no stale payload enters the cache.
            self.source.ensure_current()?;
            self.cache.check_context().map_err(map_execution_error)?;
            Ok(Arc::new(bytes))
        })();
        let reservation = flight
            .as_ref()
            .and_then(|flight| flight.reservation())
            .or_else(|| {
                bypass_resources
                    .as_ref()
                    .and_then(|resources| resources.reservation.as_ref().map(Arc::clone))
            });
        let object_reservation = flight
            .as_ref()
            .and_then(|flight| flight.payload_object_reservation())
            .or_else(|| {
                bypass_resources.as_ref().and_then(|resources| {
                    resources
                        .payload_object_reservation
                        .as_ref()
                        .map(Arc::clone)
                })
            });
        let merge_result = accounting.map_or(Ok(()), |report| report.merge_zip(&zip_accounting));
        match (flight, result) {
            (Some(flight), Ok(bytes)) => {
                let payload = CachedPayload {
                    bytes,
                    reservation,
                    object_reservation,
                };
                // The low-level report has been merged above. Make source
                // freshness the last decision before a cold value becomes
                // visible through the cache or its same-Part flight.
                if let Err(error) = self.source.ensure_current() {
                    self.cache.complete_failure(entry_id, &flight);
                    return Err(error);
                }
                if let Err(error) = self
                    .cache
                    .complete_success(entry_id, &flight, payload.clone())
                {
                    self.cache.complete_failure(entry_id, &flight);
                    return Err(map_execution_error(error));
                }
                if let Err(error) = merge_result {
                    return Err(error);
                }
                Ok(PartData { payload })
            },
            (Some(flight), Err(error)) => {
                self.cache.complete_failure(entry_id, &flight);
                Err(error)
            },
            (None, Ok(bytes)) => {
                let payload = CachedPayload {
                    bytes,
                    reservation,
                    object_reservation,
                };
                // Keep the no-flight path on the same stale-source boundary
                // as the coordinated cold-loader path.
                if let Err(error) = self.source.ensure_current() {
                    self.cache.complete_bypass_failure();
                    return Err(error);
                }
                if let Err(error) = self
                    .cache
                    .complete_bypass_success(entry_id, payload.clone())
                {
                    self.cache.complete_bypass_failure();
                    return Err(map_execution_error(error));
                }
                merge_result?;
                Ok(PartData { payload })
            },
            (None, Err(error)) => {
                self.cache.complete_bypass_failure();
                Err(error)
            },
        }
    }
}

fn replacement_member_name<'a>(
    replacement: &'a ChangedOverlay,
    parts: &'a [CatalogPart],
) -> &'a str {
    match &replacement.target {
        ChangedOverlayTarget::Part(index) => parts[*index].partname.membername(),
        ChangedOverlayTarget::Member(name) => name,
    }
}

fn is_relationship_member_name(member_name: &str) -> bool {
    let Some((directory, filename)) = member_name.rsplit_once('/') else {
        return false;
    };
    let has_rels_extension = filename
        .rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"));
    has_rels_extension && (directory == "_rels" || directory.ends_with("/_rels"))
}

fn folded_part_name(partname: &PackURI) -> Result<String> {
    folded_ascii_name(
        partname.as_str(),
        "source-backed OPC folded topology Part name",
    )
}

fn folded_ascii_name(value: &str, resource: &'static str) -> Result<String> {
    let mut folded = String::new();
    folded
        .try_reserve(value.len())
        .map_err(|source| OpcError::Allocation { resource, source })?;
    folded.push_str(value);
    folded.make_ascii_lowercase();
    Ok(folded)
}

fn content_types_with_overrides(
    source: &[u8],
    overrides: &[(PackURI, String)],
    limits: ReadLimits,
    context: Option<&ExecutionContext>,
    reservation_failures: Option<&DiagnosticCounter>,
) -> Result<(Vec<u8>, Option<Arc<Reservation>>)> {
    if overrides.is_empty() {
        return Ok((source.to_vec(), None));
    }
    // The source catalog accepts only a normal `Types` root. Keep the
    // publication insertion point deliberately narrow: self-closing roots and
    // prefixed roots are refused because there is no safe lexical location for
    // unprefixed generated `Override` elements.
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_close_start = None;
    loop {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| overlay_unavailable("content-types XML position overflows usize"))?;
        let (_, event) = reader.read_resolved_event()?;
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| overlay_unavailable("content-types XML position overflows usize"))?;
        if event_end < event_start || event_end > source.len() {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types XML event range is invalid".to_string(),
            ));
        }
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("content-types XML depth overflows"))?;
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest("unmatched closing element".to_string())
                })?;
                if depth == 0 && element.local_name().as_ref() == b"Types" {
                    if element.name().as_ref() != b"Types" {
                        return Err(OpcError::InvalidContentTypesManifest(
                            "prefixed Types roots are unsupported for topology publication"
                                .to_string(),
                        ));
                    }
                    root_close_start = Some(event_start);
                    break;
                }
            },
            Event::Empty(element) if depth == 0 && element.local_name().as_ref() == b"Types" => {
                return Err(OpcError::InvalidContentTypesManifest(
                    "self-closing Types roots are unsupported for topology publication".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let root_close_start = root_close_start.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("missing Types closing tag".to_string())
    })?;
    const ELEMENT_OVERHEAD: usize = 38;
    const MAX_ESCAPE_EXPANSION: usize = 6;
    let inserted_capacity = overrides
        .iter()
        .try_fold(0usize, |total, (partname, content_type)| {
            let value_bytes = partname
                .as_str()
                .len()
                .checked_add(content_type.len())?
                .checked_mul(MAX_ESCAPE_EXPANSION)?;
            total
                .checked_add(ELEMENT_OVERHEAD)?
                .checked_add(value_bytes)
        })
        .ok_or_else(|| overlay_unavailable("content-types override capacity overflows usize"))?;
    let total_capacity = source
        .len()
        .checked_add(inserted_capacity)
        .ok_or_else(|| overlay_unavailable("content-types XML size overflows usize"))?;
    // Charge the generated insertion buffer, output buffer, and the parsed
    // output map before any of those allocations. The map's retained strings
    // and hash tables are bounded by one additional serialized-XML-sized
    // working reservation; the source-map reservation is held by the caller.
    let generated_memory_reservation = if let Some(context) = context {
        let parser_bytes = u64::try_from(total_capacity)
            .map_err(|_| overlay_unavailable("content-types XML size overflows u64"))?;
        let combined = parser_bytes
            .checked_mul(2)
            .and_then(|bytes| bytes.checked_add(inserted_capacity as u64))
            .ok_or_else(|| overlay_unavailable("content-types memory charge overflows u64"))?;
        let reservation = context
            .reserve(Resource::Memory, combined)
            .map_err(|error| {
                if matches!(error, ExecutionError::ResourceLimit(_)) {
                    if let Some(counter) = reservation_failures {
                        counter.increment();
                    }
                }
                map_execution_error(error)
            })?;
        Some(Arc::new(reservation))
    } else {
        None
    };
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(inserted_capacity)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC content-types overrides",
            source,
        })?;
    for (index, (partname, content_type)) in overrides.iter().enumerate() {
        if index & 0x3f == 0 {
            if let Some(context) = context {
                context.check().map_err(map_execution_error)?;
            }
        }
        append_relationship_xml_bytes(&mut inserted, b"<Override PartName=\"")?;
        push_xml_escaped(&mut inserted, partname.as_str())?;
        append_relationship_xml_bytes(&mut inserted, b"\" ContentType=\"")?;
        push_xml_escaped(&mut inserted, content_type)?;
        append_relationship_xml_bytes(&mut inserted, b"\"/>")?;
    }
    let total = source
        .len()
        .checked_add(inserted.len())
        .ok_or_else(|| overlay_unavailable("content-types XML size overflows usize"))?;
    limits.check(
        ReadResource::ContentTypesBytes,
        total as u64,
        limits.max_content_types_bytes() as u64,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC content-types XML",
            source,
        })?;
    output.extend_from_slice(&source[..root_close_start]);
    output.extend_from_slice(&inserted);
    output.extend_from_slice(&source[root_close_start..]);
    ContentTypeMap::from_xml(&output, limits)?;
    Ok((output, generated_memory_reservation))
}

fn relationship_xml_without(
    relationships: &Relationships,
    removed_ids: &[String],
) -> Result<Vec<u8>> {
    const HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#;
    const FOOTER: &str = "</Relationships>";
    const ELEMENT_OVERHEAD: usize = 80;
    const MAX_ESCAPE_EXPANSION: usize = 6;

    let retained_count = relationships
        .iter()
        .filter(|relationship| {
            removed_ids
                .binary_search_by(|id| id.as_str().cmp(relationship.r_id()))
                .is_err()
        })
        .count();
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(retained_count)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed relationship removal order",
            source,
        })?;
    for relationship in relationships.iter().filter(|relationship| {
        removed_ids
            .binary_search_by(|id| id.as_str().cmp(relationship.r_id()))
            .is_err()
    }) {
        retained.push(relationship);
    }
    retained.sort_unstable_by_key(|relationship| relationship.r_id());

    let capacity = retained
        .iter()
        .try_fold(HEADER.len() + FOOTER.len(), |total, relationship| {
            let value_bytes = relationship
                .r_id()
                .len()
                .checked_add(relationship.reltype().len())?
                .checked_add(relationship.target_ref().len())?;
            total
                .checked_add(ELEMENT_OVERHEAD)?
                .checked_add(value_bytes.checked_mul(MAX_ESCAPE_EXPANSION)?)
        })
        .ok_or_else(|| {
            overlay_unavailable("relationship removal serialization capacity overflows usize")
        })?;
    let mut xml = Vec::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed relationship removal XML",
            source,
        })?;
    append_relationship_xml_bytes(&mut xml, HEADER.as_bytes())?;
    for relationship in retained {
        append_relationship_xml_bytes(&mut xml, br#"<Relationship Id=""#)?;
        push_xml_escaped(&mut xml, relationship.r_id())?;
        append_relationship_xml_bytes(&mut xml, br#"" Type=""#)?;
        push_xml_escaped(&mut xml, relationship.reltype())?;
        append_relationship_xml_bytes(&mut xml, br#"" Target=""#)?;
        push_xml_escaped(&mut xml, relationship.target_ref())?;
        if relationship.target_mode() == TargetMode::External {
            append_relationship_xml_bytes(&mut xml, br#"" TargetMode="External"/>"#)?;
        } else {
            append_relationship_xml_bytes(&mut xml, br#""/>"#)?;
        }
    }
    append_relationship_xml_bytes(&mut xml, FOOTER.as_bytes())?;
    Ok(xml)
}

fn append_relationship_xml_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed relationship removal XML",
            source,
        })?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn push_xml_escaped(output: &mut Vec<u8>, value: &str) -> Result<()> {
    for byte in value.as_bytes() {
        let escaped = match byte {
            b'&' => Some(b"&amp;".as_slice()),
            b'<' => Some(b"&lt;".as_slice()),
            b'>' => Some(b"&gt;".as_slice()),
            b'"' => Some(b"&quot;".as_slice()),
            b'\'' => Some(b"&apos;".as_slice()),
            _ => None,
        };
        if let Some(escaped) = escaped {
            append_relationship_xml_bytes(output, escaped)?;
        } else {
            append_relationship_xml_bytes(output, std::slice::from_ref(byte))?;
        }
    }
    Ok(())
}

fn relationships_for_package(
    serialized: impl IntoIterator<Item = SerializedRelationship>,
) -> Result<Relationships> {
    relationships_for_package_with_context(serialized, None)
}

fn relationships_for_package_with_context(
    serialized: impl IntoIterator<Item = SerializedRelationship>,
    context: Option<&ExecutionContext>,
) -> Result<Relationships> {
    let mut relationships = Relationships::new(PACKAGE_URI.to_string());
    for relationship in serialized {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
        relationships.try_add_relationship(
            relationship.reltype,
            relationship.target_ref,
            relationship.r_id,
            relationship.target_mode,
        )?;
    }
    Ok(relationships)
}

fn relationships_for_part(
    partname: &PackURI,
    serialized: impl IntoIterator<Item = SerializedRelationship>,
) -> Result<Relationships> {
    relationships_for_part_with_context(partname, serialized, None)
}

fn relationships_for_part_with_context(
    partname: &PackURI,
    serialized: impl IntoIterator<Item = SerializedRelationship>,
    context: Option<&ExecutionContext>,
) -> Result<Relationships> {
    let mut relationships = Relationships::for_source(partname);
    for relationship in serialized {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
        relationships.try_add_relationship(
            relationship.reltype,
            relationship.target_ref,
            relationship.r_id,
            relationship.target_mode,
        )?;
    }
    Ok(relationships)
}

fn copy_relationships(from: &Relationships, to: &mut Relationships) -> Result<()> {
    for relationship in from.iter() {
        to.try_add_relationship(
            relationship.reltype().to_string(),
            relationship.target_ref().to_string(),
            relationship.r_id().to_string(),
            relationship.target_mode(),
        )?;
    }
    Ok(())
}

fn validate_overlay_xml(part: &str, bytes: &[u8]) -> Result<()> {
    xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default())
        .map(|_report| ())
        .map_err(|source| OpcError::XmlPublication {
            part: part.to_string(),
            source,
        })
}

fn checked_overlay_total(
    current: u64,
    bytes: u64,
    resource: ReadResource,
    maximum: u64,
) -> Result<u64> {
    current.checked_add(bytes).ok_or(OpcError::ReadLimit {
        resource,
        actual: u64::MAX,
        maximum,
    })
}

fn adjusted_overlay_total(
    current: u64,
    removed: u64,
    added: u64,
    resource: ReadResource,
    maximum: u64,
) -> Result<u64> {
    current
        .checked_sub(removed)
        .and_then(|remaining| remaining.checked_add(added))
        .ok_or(OpcError::ReadLimit {
            resource,
            actual: u64::MAX,
            maximum,
        })
}

fn overlay_unavailable(reason: impl Into<String>) -> OpcError {
    OpcError::SourceBackedOverlayUnavailable {
        reason: reason.into(),
    }
}

fn map_preservation_error(error: soapberry_zip::Error) -> OpcError {
    match error.into_kind() {
        soapberry_zip::ErrorKind::IO(error) | soapberry_zip::ErrorKind::Io(error) => {
            OpcError::IoError(error)
        },
        kind => OpcError::ZipError(kind.to_string()),
    }
}

fn accounting_overflow(counter: &'static str) -> OpcError {
    OpcError::OperationAccountingOverflow { counter }
}

fn record_accounting_error(slot: &mut Option<OpcError>, error: OpcError) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn map_execution_error(error: ExecutionError) -> OpcError {
    match error {
        ExecutionError::Cancelled => OpcError::Cancelled,
        error => OpcError::Execution(error),
    }
}

fn execution_io_error(error: ExecutionError) -> std::io::Error {
    std::io::Error::other(error.to_string())
}

fn record_execution_failure(failure: &Arc<Mutex<Option<ExecutionError>>>, error: ExecutionError) {
    let mut slot = failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn record_source_execution_failure(snapshot: &SourceSnapshot, error: ExecutionError) {
    if let Some(failure) = snapshot.execution_failure.as_ref() {
        record_execution_failure(failure, error);
    }
}

fn record_input_reservation_failure(snapshot: &SourceSnapshot) {
    if let Some(counter) = snapshot.input_reservation_failures.as_ref() {
        counter.increment();
    }
}

fn take_source_execution_failure(snapshot: &SourceSnapshot) -> Option<ExecutionError> {
    snapshot.execution_failure.as_ref().and_then(|failure| {
        failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
    })
}

fn finish_source_publication(
    result: Result<()>,
    source: &SourceSnapshot,
    written: u64,
) -> Result<()> {
    let freshness = source.ensure_current();
    let result = match (result, freshness) {
        (_, Err(error @ OpcError::SourceChanged { .. })) => Err(error),
        (Err(error), _) => Err(error),
        (Ok(()), Err(error)) => Err(error),
        (Ok(()), Ok(())) => Ok(()),
    };
    match result {
        Err(source) if written != 0 => Err(OpcError::IncompleteOutput {
            written,
            source: Box::new(source),
        }),
        other => other,
    }
}

fn write_exact_snapshot<W: Write>(
    source: &SourceSnapshot,
    writer: W,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    source.monitor_publication();
    source.ensure_current()?;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC publication buffer",
            source,
        })?;
    buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
    let mut written = 0_u64;
    let result = if let Some(context) = context {
        let Some(output_reservation_failures) = source.output_reservation_failures.as_ref() else {
            return Err(overlay_unavailable(
                "managed source output reservation counter is unavailable",
            ));
        };
        let execution_failure = Arc::new(Mutex::new(None));
        let counted = Counted::new(writer, &mut written);
        let checked = SourceCheckedSink {
            inner: counted,
            snapshot: source.clone(),
        };
        let cooperative = ContextCheckedSink {
            inner: checked,
            context: Some(context.clone()),
            failure: Arc::clone(&execution_failure),
        };
        let mut sink = OutputBudgetedSink {
            inner: cooperative,
            context: context.clone(),
            failure: Arc::clone(&execution_failure),
            output_reservation_failures: output_reservation_failures.clone(),
        };
        let mut offset = 0_u64;
        let result = (|| {
            while offset < source.length {
                context.check().map_err(map_execution_error)?;
                let remaining = usize::try_from((source.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
                let read = read_source_at_with_context(
                    source,
                    Some(context),
                    offset,
                    &mut buffer[..remaining],
                    "publication",
                )?;
                if read == 0 {
                    return Err(OpcError::IoError(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "source-backed OPC source ended during publication",
                    )));
                }
                context.check().map_err(map_execution_error)?;
                sink.write_all(&buffer[..read])?;
                offset = offset
                    .checked_add(read as u64)
                    .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
            }
            context.check().map_err(map_execution_error)?;
            sink.flush()?;
            Ok(())
        })();
        execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)))
    } else {
        let counted = Counted::new(writer, &mut written);
        let mut sink = SourceCheckedSink {
            inner: counted,
            snapshot: source.clone(),
        };
        let mut offset = 0_u64;
        (|| {
            while offset < source.length {
                let remaining = usize::try_from((source.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
                let read = read_source_at_with_context(
                    source,
                    None,
                    offset,
                    &mut buffer[..remaining],
                    "publication",
                )?;
                if read == 0 {
                    return Err(OpcError::IoError(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "source-backed OPC source ended during publication",
                    )));
                }
                sink.write_all(&buffer[..read])?;
                offset = offset
                    .checked_add(read as u64)
                    .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
            }
            sink.flush()?;
            Ok(())
        })()
    };
    finish_source_publication(result, source, written)
}

fn write_exact_snapshot_with_accounting<W: Write>(
    source: &SourceSnapshot,
    writer: W,
    context: Option<&ExecutionContext>,
    accounting: &mut OpcOperationAccounting,
) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    source.monitor_publication();
    source.ensure_current()?;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC publication buffer",
            source,
        })?;
    buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
    let mut written = 0_u64;
    let mut accounting_error = None;
    let result = if let Some(context) = context {
        let Some(output_reservation_failures) = source.output_reservation_failures.as_ref() else {
            return Err(overlay_unavailable(
                "managed source output reservation counter is unavailable",
            ));
        };
        let execution_failure = Arc::new(Mutex::new(None));
        let counted = Counted::with_accounting(
            writer,
            &mut written,
            accounting,
            &mut accounting_error,
            true,
        );
        let checked = SourceCheckedSink {
            inner: counted,
            snapshot: source.clone(),
        };
        let cooperative = ContextCheckedSink {
            inner: checked,
            context: Some(context.clone()),
            failure: Arc::clone(&execution_failure),
        };
        let mut sink = OutputBudgetedSink {
            inner: cooperative,
            context: context.clone(),
            failure: Arc::clone(&execution_failure),
            output_reservation_failures: output_reservation_failures.clone(),
        };
        let mut offset = 0_u64;
        let result = (|| {
            while offset < source.length {
                context.check().map_err(map_execution_error)?;
                let remaining = usize::try_from((source.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
                let read = read_source_at_with_context(
                    source,
                    Some(context),
                    offset,
                    &mut buffer[..remaining],
                    "publication",
                )?;
                if read == 0 {
                    return Err(OpcError::IoError(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "source-backed OPC source ended during publication",
                    )));
                }
                context.check().map_err(map_execution_error)?;
                sink.write_all(&buffer[..read])?;
                offset = offset
                    .checked_add(read as u64)
                    .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
            }
            context.check().map_err(map_execution_error)?;
            sink.flush()?;
            Ok(())
        })();
        execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)))
    } else {
        let counted = Counted::with_accounting(
            writer,
            &mut written,
            accounting,
            &mut accounting_error,
            true,
        );
        let mut sink = SourceCheckedSink {
            inner: counted,
            snapshot: source.clone(),
        };
        let mut offset = 0_u64;
        (|| {
            while offset < source.length {
                let remaining = usize::try_from((source.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
                let read = read_source_at_with_context(
                    source,
                    None,
                    offset,
                    &mut buffer[..remaining],
                    "publication",
                )?;
                if read == 0 {
                    return Err(OpcError::IoError(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "source-backed OPC source ended during publication",
                    )));
                }
                sink.write_all(&buffer[..read])?;
                offset = offset
                    .checked_add(read as u64)
                    .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
            }
            sink.flush()?;
            Ok(())
        })()
    };
    let result = take_source_execution_failure(source)
        .map_or(result, |error| Err(map_execution_error(error)));
    let result = finish_source_publication(result, source, written);
    match result {
        Err(error) => Err(error),
        Ok(()) => accounting_error.map_or(Ok(()), Err),
    }
}

/// Reserve a bounded positional-read window, perform exactly one source read,
/// and commit only the bytes actually accepted. The retry loop is important
/// for short-read adapters: a caller with one input byte remaining must still
/// be allowed to read one byte even when the adapter initially receives a
/// larger output buffer. A reservation is held only across the physical read,
/// so cumulative [`Resource::InputBytes`] usage is exact under retries and
/// never leaks on an I/O failure.
fn read_source_at_with_context(
    snapshot: &SourceSnapshot,
    context: Option<&ExecutionContext>,
    offset: u64,
    output: &mut [u8],
    operation: &str,
) -> Result<usize> {
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    snapshot.ensure_current_io_if_monitored()?;
    let requested = output.len();
    let (read_output, reservation) = if let Some(context) = context {
        let mut attempt = u64::try_from(requested)
            .map_err(|_| overlay_unavailable("source read length overflows u64"))?;
        loop {
            match context.reserve(Resource::InputBytes, attempt) {
                Ok(reservation) => {
                    let length = usize::try_from(attempt)
                        .map_err(|_| overlay_unavailable("source read length overflows usize"))?;
                    break (&mut output[..length], Some(reservation));
                },
                Err(error) => {
                    let Some(limit) = (match &error {
                        ExecutionError::ResourceLimit(limit) => Some(limit),
                        ExecutionError::Cancelled
                        | ExecutionError::WorkersExceedInFlightTasks { .. }
                        | ExecutionError::ParallelThresholdExceedsInFlightBytes { .. } => None,
                        _ => None,
                    }) else {
                        return Err(map_execution_error(error));
                    };
                    let previous = limit.observed.saturating_sub(attempt);
                    let available = limit.limit.saturating_sub(previous);
                    let next = available.min(attempt.saturating_sub(1));
                    if next == 0 {
                        record_input_reservation_failure(snapshot);
                        return Err(map_execution_error(error));
                    }
                    attempt = next;
                },
            }
        }
    } else {
        (output, None)
    };
    let read = match snapshot.source.read_at(offset, read_output) {
        Ok(read) => read,
        Err(error) => return Err(OpcError::IoError(error)),
    };
    validate_source_read_count(read, read_output.len(), operation)?;
    if let Some(reservation) = reservation {
        if !reservation.commit(read as u64) {
            return Err(overlay_unavailable(
                "source-backed OPC input reservation underflow",
            ));
        }
    }
    if let Some(context) = context {
        context.check().map_err(|error| {
            record_source_execution_failure(snapshot, error.clone());
            map_execution_error(error)
        })?;
    }
    snapshot
        .ensure_current_io_if_monitored()
        .map_err(OpcError::IoError)?;
    Ok(read)
}

fn validate_source_read_count(read: usize, requested: usize, operation: &str) -> Result<()> {
    if read <= requested {
        return Ok(());
    }
    Err(OpcError::IoError(std::io::Error::new(
        std::io::ErrorKind::InvalidData,
        format!(
            "source-backed OPC source reported {read} bytes for a {requested}-byte {operation} read"
        ),
    )))
}

fn is_signature_relationship(kind: &str) -> bool {
    [
        relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
    ]
    .iter()
    .any(|candidate| kind.eq_ignore_ascii_case(candidate))
}

fn is_signature_relationship_or_target(relationship: &crate::Relationship) -> bool {
    if is_signature_relationship(relationship.reltype()) {
        return true;
    }
    if relationship.is_external() {
        return false;
    }
    let target_path = relationship.target_path();
    if is_signature_path(target_path) {
        return true;
    }
    if !target_may_be_signature_path(target_path) {
        return false;
    }
    relationship
        .target_partname()
        .map_or(true, |target| is_signature_path(target.as_str()))
}

fn target_may_be_signature_path(path: &str) -> bool {
    path.split('/')
        .any(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
}

fn is_signature_path(path: &str) -> bool {
    const DIRECTORY: &[u8] = b"/_xmlsignatures/";
    path.as_bytes()
        .get(..DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(DIRECTORY))
}

fn is_signature_content_type(value: &str) -> bool {
    [
        content_type::OPC_DIGITAL_SIGNATURE_ORIGIN,
        content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
    ]
    .iter()
    .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic on failure by design"
    )]

    use super::*;
    use std::collections::HashMap;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Barrier;
    use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
    use std::time::Duration;

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };

    struct CountingSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
        versions: AtomicUsize,
        read_bytes: AtomicU64,
        read_ranges: Mutex<Vec<(u64, usize)>>,
        max_read: usize,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
                versions: AtomicUsize::new(0),
                read_bytes: AtomicU64::new(0),
                read_ranges: Mutex::new(Vec::new()),
                max_read: usize::MAX,
            }
        }

        fn chunked(bytes: Vec<u8>, max_read: usize) -> Self {
            let mut source = Self::new(bytes);
            source.max_read = max_read.max(1);
            source
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }

        fn read_ranges(&self) -> Vec<(u64, usize)> {
            self.read_ranges
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner())
                .clone()
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let count = count.min(self.max_read);
            self.read_ranges
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner())
                .push((offset as u64, count));
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            self.read_bytes.fetch_add(count as u64, Ordering::SeqCst);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            self.versions.fetch_add(1, Ordering::SeqCst);
            Ok(SourceVersion::new(42, self.revision.load(Ordering::SeqCst)))
        }
    }

    struct CancelOnHitVersionSource {
        bytes: Vec<u8>,
        cancellation_source: CancellationSource,
        skip_versions: AtomicUsize,
        armed: AtomicBool,
    }

    impl CancelOnHitVersionSource {
        fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
            Self {
                bytes,
                cancellation_source,
                skip_versions: AtomicUsize::new(0),
                armed: AtomicBool::new(false),
            }
        }

        fn arm_after_cache_enter(&self) {
            // The part lookup and `read_part` perform three freshness checks
            // before a hit's post-entry check can run.
            self.skip_versions.store(3, Ordering::SeqCst);
            self.armed.store(true, Ordering::SeqCst);
        }
    }

    impl ReadAt for CancelOnHitVersionSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            if self.armed.load(Ordering::SeqCst)
                && self
                    .skip_versions
                    .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |remaining| {
                        remaining.checked_sub(1)
                    })
                    .is_err()
            {
                self.armed.store(false, Ordering::SeqCst);
                self.cancellation_source.cancel();
            }
            Ok(SourceVersion::new(43, 0))
        }
    }

    struct OverReportingSource {
        bytes: Vec<u8>,
        overreport: AtomicBool,
    }

    impl OverReportingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                overreport: AtomicBool::new(false),
            }
        }
    }

    impl ReadAt for OverReportingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            if self.overreport.load(Ordering::SeqCst) {
                return Ok(output.len().saturating_add(1));
            }
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(93, 0))
        }
    }

    struct OverReportingSink {
        calls: usize,
        accepted: usize,
    }

    impl Write for OverReportingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.calls = self.calls.saturating_add(1);
            Ok(bytes.len().saturating_add(1))
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    struct SlowPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        payload_reads: AtomicUsize,
    }

    impl SlowPayloadSource {
        fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
            Self {
                bytes,
                payload_offset,
                payload_reads: AtomicUsize::new(0),
            }
        }
    }

    impl ReadAt for SlowPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset == self.payload_offset {
                self.payload_reads.fetch_add(1, Ordering::SeqCst);
                // Keep the cold load in flight long enough for the peer to
                // enter the same part concurrently.
                std::thread::sleep(Duration::from_millis(100));
            }
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(77, 0))
        }
    }

    struct CancelDuringPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        cancellation_source: CancellationSource,
        armed: AtomicBool,
    }

    struct CancelDuringOpenSource {
        bytes: Vec<u8>,
        cancellation_source: CancellationSource,
        reads: AtomicUsize,
        cancel_after: usize,
    }

    impl CancelDuringOpenSource {
        fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
            Self {
                bytes,
                cancellation_source,
                reads: AtomicUsize::new(0),
                cancel_after: 1,
            }
        }
    }

    impl ReadAt for CancelDuringOpenSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if self.reads.fetch_add(1, Ordering::SeqCst) + 1 >= self.cancel_after {
                self.cancellation_source.cancel();
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(95, 0))
        }
    }

    impl CancelDuringPayloadSource {
        fn new(
            bytes: Vec<u8>,
            payload_offset: usize,
            cancellation_source: CancellationSource,
        ) -> Self {
            Self {
                bytes,
                payload_offset,
                cancellation_source,
                armed: AtomicBool::new(true),
            }
        }
    }

    impl ReadAt for CancelDuringPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if offset == self.payload_offset && self.armed.swap(false, Ordering::SeqCst) {
                // The bytes have been decompressed into the loader's private
                // allocation, but the publication checks must reject them.
                self.cancellation_source.cancel();
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(79, 0))
        }
    }

    struct ChangeDuringPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        revision: AtomicU64,
        armed: AtomicBool,
    }

    impl ChangeDuringPayloadSource {
        fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
            Self {
                bytes,
                payload_offset,
                revision: AtomicU64::new(0),
                armed: AtomicBool::new(false),
            }
        }
    }

    impl ReadAt for ChangeDuringPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if offset == self.payload_offset && self.armed.swap(false, Ordering::SeqCst) {
                self.revision.fetch_add(1, Ordering::SeqCst);
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(88, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn managed_context_with_cancellation(
        memory: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let (budget, cancellation_source, context) =
            managed_context_with_resources(memory, u64::MAX, u64::MAX, u64::MAX);
        (budget, cancellation_source, context)
    }

    fn managed_context_with_all_resources(
        memory: u64,
        input_bytes: u64,
        output_bytes: u64,
        objects: u64,
        work: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "opc-source-cache-test",
            Limits::new(memory, input_bytes, output_bytes, objects, u64::MAX, work),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    fn managed_context_with_resources(
        memory: u64,
        input_bytes: u64,
        objects: u64,
        work: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        managed_context_with_all_resources(memory, input_bytes, u64::MAX, objects, work)
    }

    fn managed_context_with_output(
        output_bytes: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        managed_context_with_all_resources(4096, u64::MAX, output_bytes, u64::MAX, u64::MAX)
    }

    fn managed_context(memory: u64) -> (Budget, ExecutionContext) {
        let (budget, cancellation_source, context) = managed_context_with_cancellation(memory);
        drop(cancellation_source);
        (budget, context)
    }

    fn archive_bytes(root_relationships: &[u8], document: &[u8], include_junk: bool) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        if include_junk {
            writer.write_stored("scratch.bin", b"not a part").unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn owned_bytes_open_keeps_ordinary_payloads_deferred() {
        let package = SourceBackedPackage::from_vec(archive_bytes(
            root_relationships(),
            b"<document/>",
            true,
        ))
        .unwrap();

        let opened = package.cache_diagnostics();
        assert_eq!(opened.cold_loads, 0);
        assert_eq!(opened.retained_entries, 0);

        let document = PackURI::new("/word/document.xml").unwrap();
        let data = package.part(&document).unwrap().data().unwrap();
        assert_eq!(data.as_bytes(), b"<document/>");

        let loaded = package.cache_diagnostics();
        assert_eq!(loaded.cold_loads, 1);
        assert_eq!(loaded.successful_loads, 1);
        assert_eq!(loaded.retained_entries, 1);
    }

    #[test]
    fn part_view_declared_size_is_metadata_only() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<document/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let before = package.cache_diagnostics();
        let document = PackURI::new("/word/document.xml").unwrap();
        let view = package.part(&document).unwrap();

        assert_eq!(
            view.declared_uncompressed_size().unwrap(),
            b"<document/>".len() as u64
        );
        let after = package.cache_diagnostics();
        assert_eq!(after.cold_loads, before.cold_loads);
        assert_eq!(after.retained_entries, before.retained_entries);
    }

    #[test]
    fn part_view_declared_size_checks_source_and_cancellation() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<document/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let view = package.part(&document).unwrap();
        source.changed();
        assert!(matches!(
            view.declared_uncompressed_size(),
            Err(OpcError::SourceChanged { .. })
        ));

        let (_budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(archive_bytes(
                root_relationships(),
                b"<document/>",
                false,
            ))),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let view = package.part(&document).unwrap();
        cancellation_source.cancel();
        assert!(matches!(
            view.declared_uncompressed_size(),
            Err(OpcError::Cancelled)
        ));
    }

    #[test]
    fn topology_add_part_inserts_lossless_content_type_override() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_add_part(
            PackURI::new("/custom/new.bin").unwrap(),
            "application/octet-stream",
            b"new payload".to_vec(),
        )
        .unwrap();
        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();

        let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
        let content_types = archive.read("[Content_Types].xml").unwrap();
        let source_content_types = archive_bytes(root_relationships(), b"<before/>", false);
        assert!(content_types
            .windows(b"<Override PartName=\"/custom/new.bin\" ContentType=\"application/octet-stream\"/>".len())
            .any(|window| window
                == b"<Override PartName=\"/custom/new.bin\" ContentType=\"application/octet-stream\"/>"));
        assert_ne!(output, source_content_types);
        assert_eq!(archive.read("custom/new.bin").unwrap(), b"new payload");
        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(
            reopened
                .iter_parts()
                .find(|part| part.partname().as_str() == "/custom/new.bin")
                .unwrap()
                .blob(),
            b"new payload"
        );
        assert_eq!(archive.read("custom/orphan.xml").unwrap(), b"<orphan/>");
    }

    #[test]
    fn content_type_override_insertion_uses_the_parsed_root_close_span() {
        let source = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><!-- text that looks like </Types> --><?keep?></Types>"#;
        let partname = PackURI::new("/custom/new.bin").unwrap();
        let (output, _memory_reservation) = content_types_with_overrides(
            source,
            &[(partname, "application/octet-stream".to_string())],
            ReadLimits::default(),
            None,
            None,
        )
        .unwrap();
        assert_eq!(
            output,
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><!-- text that looks like </Types> --><?keep?><Override PartName="/custom/new.bin" ContentType="application/octet-stream"/></Types>"#
        );
        ContentTypeMap::from_xml(&output, ReadLimits::default()).unwrap();
    }

    #[test]
    fn topology_adds_package_relationship_to_new_part() {
        let root = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer.write_stored("_rels/.rels", root).unwrap();
        writer
            .write_stored("word/document.xml", b"<before/>")
            .unwrap();
        let source = writer.finish_to_bytes().unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let mut plan = SourceTopologyPlan::new();
        let target = PackURI::new("/custom/new.bin").unwrap();
        plan.try_add_part(
            target.clone(),
            "application/octet-stream",
            b"new payload".to_vec(),
        )
        .unwrap();
        plan.try_add_internal_relationship(
            PackURI::new("/").unwrap(),
            "rId2",
            "urn:litchi:test",
            target,
        )
        .unwrap();
        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        let reopened =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(output))).unwrap();
        assert_eq!(
            reopened.rels().get("rId2").unwrap().target_ref(),
            "custom/new.bin"
        );
        assert_eq!(
            reopened
                .part(&PackURI::new("/custom/new.bin").unwrap())
                .unwrap()
                .data()
                .unwrap()
                .as_bytes(),
            b"new payload"
        );
    }

    fn large_archive_bytes(
        root_relationships: &[u8],
        document: &[u8],
        junk_bytes: usize,
    ) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        let junk = vec![0xA5; junk_bytes];
        writer.write_stored("scratch.bin", &junk).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn archive_with_part_names(root_target: &str, part_names: &[&str]) -> Vec<u8> {
        let root_relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{root_target}"/></Relationships>"#
        );
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships.as_bytes())
            .unwrap();
        for (index, part_name) in part_names.iter().enumerate() {
            writer
                .write_stored(
                    part_name,
                    if index == 0 {
                        &b"document"[..]
                    } else {
                        &b"other"[..]
                    },
                )
                .unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn archive_with_ordered_xml_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        assert!(!parts.is_empty());
        let root_relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{}"/></Relationships>"#,
            parts[0].0
        );
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships.as_bytes())
            .unwrap();
        for (name, payload) in parts {
            writer.write_stored(name, payload).unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn archive_with_document_relationships(document_relationships: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<before/>")
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", document_relationships)
            .unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        writer.write_stored("scratch.bin", b"untouched").unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn document_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://remove.invalid/" TargetMode="External"/><Relationship Id="rInternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../custom/orphan.xml"/></Relationships>"#
    }

    fn topology_relationship_archive(document_relationships: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<before/>")
            .unwrap();
        writer
            .write_stored("word/target.xml", b"<target/>")
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", document_relationships)
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn canonical_document_relationships(entries: &[(&str, &str, &str, TargetMode)]) -> Vec<u8> {
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut relationships = Relationships::for_source(&owner);
        for (r_id, reltype, target, mode) in entries {
            relationships
                .try_add_relationship(
                    (*reltype).to_string(),
                    (*target).to_string(),
                    (*r_id).to_string(),
                    *mode,
                )
                .unwrap();
        }
        relationships.to_xml().into_bytes()
    }

    #[test]
    fn topology_adds_external_relationship_with_exact_uri_reference() {
        let relationships = canonical_document_relationships(&[]);
        let source = topology_relationship_archive(&relationships);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_add_external_relationship(
            owner.clone(),
            "rLink",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            "https://example.invalid/path?q=one#bookmark",
        )
        .unwrap();

        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        let reopened =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(output))).unwrap();
        let part = reopened.part(&owner).unwrap();
        let relationship = part.rels().get("rLink").unwrap();
        assert_eq!(relationship.target_mode(), TargetMode::External);
        assert_eq!(
            relationship.target_ref(),
            "https://example.invalid/path?q=one#bookmark"
        );
    }

    #[test]
    fn topology_replaces_and_removes_external_relationships() {
        const HYPERLINK: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        let relationships = canonical_document_relationships(&[
            (
                "rReplace",
                HYPERLINK,
                "https://before.invalid/",
                TargetMode::External,
            ),
            (
                "rRemove",
                HYPERLINK,
                "https://remove.invalid/",
                TargetMode::External,
            ),
        ]);
        let source = topology_relationship_archive(&relationships);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_external_relationship(
            owner.clone(),
            "rReplace",
            HYPERLINK,
            "https://after.invalid/?q=two#fragment",
        )
        .unwrap();
        plan.try_remove_external_relationship(owner.clone(), "rRemove")
            .unwrap();

        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        let reopened =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(output))).unwrap();
        let part = reopened.part(&owner).unwrap();
        let relationships = part.rels();
        assert_eq!(
            relationships.get("rReplace").unwrap().target_ref(),
            "https://after.invalid/?q=two#fragment"
        );
        assert!(relationships.get("rRemove").is_none());
    }

    #[test]
    fn topology_exact_external_replacement_preserves_noncanonical_source() {
        const HYPERLINK: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship TargetMode="External" Target="https://same.invalid/?q=1#same" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Id="rSame"/></Relationships>"#;
        let source = topology_relationship_archive(relationships);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source.clone())))
                .unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_external_relationship(
            PackURI::new("/word/document.xml").unwrap(),
            "rSame",
            HYPERLINK,
            "https://same.invalid/?q=1#same",
        )
        .unwrap();

        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        assert_eq!(output, source);
    }

    #[test]
    fn topology_external_wrappers_refuse_internal_relationships_before_output() {
        const HYPERLINK: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        let relationships = canonical_document_relationships(&[(
            "rInternal",
            HYPERLINK,
            "document.xml",
            TargetMode::Internal,
        )]);
        let source = topology_relationship_archive(&relationships);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_external_relationship(
            PackURI::new("/word/document.xml").unwrap(),
            "rInternal",
            HYPERLINK,
            "https://example.invalid/",
        )
        .unwrap();

        let mut output = Vec::new();
        let error = package
            .write_topology_to_stream(&mut output, plan)
            .unwrap_err();
        assert!(matches!(error, OpcError::InvalidRelationship(_)));
        assert!(output.is_empty());
    }

    #[test]
    fn topology_internal_replace_requires_and_preserves_internal_mode() {
        const TEST_RELATIONSHIP: &str = "urn:litchi:test";
        let relationships = canonical_document_relationships(&[
            (
                "rInternal",
                TEST_RELATIONSHIP,
                "document.xml",
                TargetMode::Internal,
            ),
            (
                "rExternal",
                TEST_RELATIONSHIP,
                "https://example.invalid/",
                TargetMode::External,
            ),
        ]);
        let source = topology_relationship_archive(&relationships);
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_internal_relationship(
            owner.clone(),
            "rInternal",
            TEST_RELATIONSHIP,
            PackURI::new("/word/target.xml").unwrap(),
        )
        .unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source.clone())))
                .unwrap();
        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        let reopened =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(output))).unwrap();
        let part = reopened.part(&owner).unwrap();
        let relationship = part.rels().get("rInternal").unwrap();
        assert_eq!(relationship.target_mode(), TargetMode::Internal);
        assert_eq!(
            relationship.target_partname().unwrap(),
            PackURI::new("/word/target.xml").unwrap()
        );

        let mut mismatch = SourceTopologyPlan::new();
        mismatch
            .try_replace_internal_relationship(
                owner,
                "rExternal",
                TEST_RELATIONSHIP,
                PackURI::new("/word/target.xml").unwrap(),
            )
            .unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let mut mismatch_output = Vec::new();
        let error = package
            .write_topology_to_stream(&mut mismatch_output, mismatch)
            .unwrap_err();
        assert!(matches!(error, OpcError::InvalidRelationship(_)));
        assert!(mismatch_output.is_empty());
    }

    #[test]
    fn topology_relationship_publication_honors_configured_reopen_limits() {
        const HYPERLINK: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        let relationships = canonical_document_relationships(&[]);
        let source = topology_relationship_archive(&relationships);
        let owner = PackURI::new("/word/document.xml").unwrap();

        let target_limits = ReadLimits::builder()
            .max_relationship_target_bytes(24)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source.clone())),
            target_limits,
        )
        .unwrap();
        let mut target_plan = SourceTopologyPlan::new();
        target_plan
            .try_add_external_relationship(
                owner.clone(),
                "rTargetLimit",
                HYPERLINK,
                "https://example.invalid/too-long",
            )
            .unwrap();
        let mut target_output = Vec::new();
        assert!(matches!(
            package.write_topology_to_stream(&mut target_output, target_plan),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipTargetBytes,
                ..
            })
        ));
        assert!(target_output.is_empty());

        let attribute_limits = ReadLimits::builder()
            .max_xml_attribute_bytes(128)
            .unwrap()
            .max_relationship_target_bytes(128)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source.clone())),
            attribute_limits,
        )
        .unwrap();
        let mut attribute_plan = SourceTopologyPlan::new();
        attribute_plan
            .try_add_external_relationship(
                owner.clone(),
                "rAttributeLimit",
                HYPERLINK,
                format!("https://example.invalid/?{}", "&".repeat(30)),
            )
            .unwrap();
        let mut attribute_output = Vec::new();
        assert!(matches!(
            package.write_topology_to_stream(&mut attribute_output, attribute_plan),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlAttributeBytes,
                ..
            })
        ));
        assert!(attribute_output.is_empty());

        let name_limits = ReadLimits::builder()
            .max_archive_member_name_bytes(32)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source.clone())),
            name_limits,
        )
        .unwrap();
        let mut name_plan = SourceTopologyPlan::new();
        name_plan
            .try_add_part(
                PackURI::new(format!("/custom/{}.xml", "x".repeat(40))).unwrap(),
                "application/xml",
                b"<new/>".to_vec(),
            )
            .unwrap();
        let mut name_output = Vec::new();
        assert!(matches!(
            package.write_topology_to_stream(&mut name_output, name_plan),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMemberNameBytes,
                ..
            })
        ));
        assert!(name_output.is_empty());

        let default_limits = ReadLimits::default();
        let source_events = relationship_xml_event_count(root_relationships(), default_limits)
            .unwrap()
            .checked_add(relationship_xml_event_count(&relationships, default_limits).unwrap())
            .unwrap();
        let event_limits = ReadLimits::builder()
            .max_xml_events(5)
            .unwrap()
            .max_total_relationship_xml_events(source_events as usize)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source)),
            event_limits,
        )
        .unwrap();
        let mut event_plan = SourceTopologyPlan::new();
        event_plan
            .try_add_external_relationship(
                owner,
                "rEventLimit",
                HYPERLINK,
                "https://example.invalid/",
            )
            .unwrap();
        let mut event_output = Vec::new();
        assert!(matches!(
            package.write_topology_to_stream(&mut event_output, event_plan),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlEvents,
                ..
            })
        ));
        assert!(event_output.is_empty());
    }

    fn relationship_batch(prefix: &str, count: usize) -> (Vec<String>, Vec<u8>) {
        let mut ids = Vec::new();
        let mut xml = String::from(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for index in 0..count {
            let id = format!("r{prefix}{index}");
            ids.push(id.clone());
            xml.push_str(&format!(
                r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://{prefix}{index}.invalid/" TargetMode="External"/>"#
            ));
        }
        xml.push_str("</Relationships>");
        (ids, xml.into_bytes())
    }

    fn archive_with_relationship_batches(
        first: usize,
        second: usize,
    ) -> (Vec<u8>, Vec<String>, Vec<String>) {
        let (first_ids, first_rels) = relationship_batch("first", first);
        let (second_ids, second_rels) = relationship_batch("second", second);
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<before/>")
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", &first_rels)
            .unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<before/>")
            .unwrap();
        writer
            .write_stored("custom/_rels/orphan.xml.rels", &second_rels)
            .unwrap();
        (writer.finish_to_bytes().unwrap(), first_ids, second_ids)
    }

    fn source_archive_total(package: &SourceBackedPackage) -> u64 {
        package
            .archive
            .file_names()
            .map(|name| package.archive.metadata(name).unwrap().uncompressed_size())
            .sum()
    }

    #[test]
    fn source_artifact_paths_reject_overreported_read_counts_without_output() {
        let source = Arc::new(OverReportingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let package = SourceBackedPackage::from_read_at(read_at).unwrap();
        let artifact = package.source_artifact();
        source.overreport.store(true, Ordering::SeqCst);

        assert!(matches!(
            artifact.fingerprint(),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));

        let mut output = Vec::new();
        assert!(matches!(
            artifact.write_to_stream(&mut output),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn source_artifact_copy_rejects_overreporting_sink_without_false_progress() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"<before/>", false),
        )))
        .unwrap();
        let artifact = package.source_artifact();
        let mut sink = OverReportingSink {
            calls: 0,
            accepted: 0,
        };

        assert!(matches!(
            artifact.write_to_stream(&mut sink),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));
        assert_eq!(sink.calls, 1);
        assert_eq!(sink.accepted, 0);
    }

    fn root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
    }

    fn signed_root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="signature/origin.xml"/></Relationships>"#
    }

    fn signed_archive(document: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", signed_root_relationships())
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer
            .write_stored("signature/origin.xml", b"<origin/>")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[derive(Debug)]
    struct RawRecord {
        local: Vec<u8>,
        central: Vec<u8>,
    }

    fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
        let archive = soapberry_zip::ZipArchive::from_slice(data)
            .unwrap()
            .into_zip_archive();
        let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
        let index = soapberry_zip::PreservationIndex::new(&archive, &mut scratch).unwrap();
        index
            .entries()
            .iter()
            .map(|entry| {
                let local = entry.local_span();
                let central = entry.central_record();
                (
                    entry.raw_name_bytes().to_vec(),
                    RawRecord {
                        local: data[local.start as usize..local.end as usize].to_vec(),
                        central: data[central.start as usize..central.end as usize].to_vec(),
                    },
                )
            })
            .collect()
    }

    fn central_without_local_offset(bytes: &[u8]) -> Vec<u8> {
        let mut bytes = bytes.to_vec();
        bytes[42..46].fill(0);
        bytes
    }

    struct MutatingSink {
        source: Arc<CountingSource>,
        bytes: Vec<u8>,
        change_after: usize,
        changed: bool,
    }

    impl Write for MutatingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            if !self.changed && self.bytes.len() >= self.change_after {
                self.source.changed();
                self.changed = true;
            }
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    struct BoundedFailingSink {
        accepted: usize,
        limit: usize,
        largest_write: usize,
    }

    struct CancelAfterWriteSink {
        cancellation_source: CancellationSource,
        bytes: Vec<u8>,
        cancelled: bool,
    }

    struct ShortInterruptedSink {
        bytes: Vec<u8>,
        maximum_write: usize,
        interrupted: bool,
    }

    struct FlushFailingSink {
        bytes: Vec<u8>,
    }

    impl Write for CancelAfterWriteSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            if !self.cancelled {
                self.cancelled = true;
                self.cancellation_source.cancel();
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    impl Write for BoundedFailingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.largest_write = self.largest_write.max(bytes.len());
            let remaining = self.limit.saturating_sub(self.accepted);
            if remaining == 0 {
                return Err(std::io::Error::other("injected sink failure"));
            }
            let written = remaining.min(bytes.len());
            self.accepted += written;
            Ok(written)
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    impl Write for ShortInterruptedSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            if !self.interrupted {
                self.interrupted = true;
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Interrupted,
                    "injected retryable interruption",
                ));
            }
            if bytes.is_empty() {
                return Ok(0);
            }
            let written = self.maximum_write.min(bytes.len());
            self.bytes.extend_from_slice(&bytes[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    impl Write for FlushFailingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Err(std::io::Error::other("injected flush failure"))
        }
    }

    #[test]
    fn mandatory_xml_is_opened_but_ordinary_payload_corruption_is_deferred() {
        let malformed = archive_bytes(b"<Relationships", b"document", false);
        let malformed_source = Arc::new(CountingSource::new(malformed));
        assert!(matches!(
            SourceBackedPackage::from_read_at(malformed_source),
            Err(OpcError::QuickXmlError(_))
        ));

        const DOCUMENT: &[u8] = b"source-backed deferred corruption";
        let mut corrupt = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = corrupt
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        corrupt[position] ^= 0xff;
        let source = Arc::new(CountingSource::new(corrupt));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let main = package.main_document_part().unwrap();
        assert!(matches!(main.data(), Err(OpcError::ZipError(_))));
    }

    #[test]
    fn source_backed_catalog_open_does_not_read_ordinary_payloads() {
        const DOCUMENT: &[u8] = b"source-backed cold catalog payload sentinel";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_start = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .expect("stored payload must be present in the archive");
        let payload_end = payload_start + DOCUMENT.len();
        let source = Arc::new(CountingSource::new(bytes));

        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        assert_eq!(package.cache_diagnostics().cold_loads, 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(package.iter_parts().count(), 2);
        assert!(source.read_bytes.load(Ordering::SeqCst) > 0);
        assert!(source.read_ranges().into_iter().all(|(offset, length)| {
            let start = usize::try_from(offset).expect("test source offset fits usize");
            let end = start + length;
            end <= payload_start || start >= payload_end
        }));
    }

    #[test]
    fn source_backed_catalog_matches_eager_part_admission() {
        let bytes = archive_bytes(root_relationships(), b"catalog parity", true);
        let eager = OpcPackage::from_bytes(&bytes).unwrap();
        let source =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

        let mut eager_parts: Vec<_> = eager
            .iter_parts()
            .map(|part| {
                let mut relationships: Vec<_> = part
                    .rels()
                    .iter()
                    .map(|relationship| {
                        (
                            relationship.r_id().to_owned(),
                            relationship.reltype().to_owned(),
                            relationship.target_ref().to_owned(),
                            relationship.target_mode(),
                        )
                    })
                    .collect();
                relationships.sort_by(|left, right| left.0.cmp(&right.0));
                (
                    part.partname().to_string(),
                    part.content_type().to_owned(),
                    relationships,
                )
            })
            .collect();
        let mut deferred_parts: Vec<_> = source
            .iter_parts()
            .map(|part| {
                let mut relationships: Vec<_> = part
                    .rels()
                    .iter()
                    .map(|relationship| {
                        (
                            relationship.r_id().to_owned(),
                            relationship.reltype().to_owned(),
                            relationship.target_ref().to_owned(),
                            relationship.target_mode(),
                        )
                    })
                    .collect();
                relationships.sort_by(|left, right| left.0.cmp(&right.0));
                (
                    part.partname().to_string(),
                    part.content_type().to_owned(),
                    relationships,
                )
            })
            .collect();
        eager_parts.sort_by(|left, right| left.0.cmp(&right.0));
        deferred_parts.sort_by(|left, right| left.0.cmp(&right.0));

        assert_eq!(eager_parts, deferred_parts);
        assert_eq!(eager.non_part_members(), source.non_part_members());
        assert_eq!(eager.rels().len(), source.rels().len());
        assert_eq!(
            eager.main_document_part().unwrap().partname(),
            source.main_document_part().unwrap().partname()
        );
    }

    #[test]
    fn source_backed_catalog_matches_eager_relationship_typing_and_non_parts() {
        let bytes = archive_with_document_relationships(document_relationships());
        let eager = OpcPackage::from_bytes(&bytes).unwrap();
        let source =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

        assert_eq!(eager.non_part_members(), source.non_part_members());
        assert_eq!(source.non_part_members().len(), 1);
        assert_eq!(source.non_part_members()[0].name(), "scratch.bin");
        assert!(
            source
                .physical_member_names()
                .any(|name| name == "word/_rels/document.xml.rels")
        );
        assert!(
            source
                .iter_parts()
                .all(|part| part.partname().as_str() != "/word/_rels/document.xml.rels")
        );
        assert_eq!(
            source
                .part(&PackURI::new("/word/document.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            eager
                .get_part(&PackURI::new("/word/document.xml").unwrap())
                .unwrap()
                .rels()
                .len()
        );
    }

    #[test]
    fn source_backed_catalog_enforces_part_limits_during_ordered_admission() {
        let bytes = archive_with_ordered_xml_parts(&[
            ("word/document.xml", b"document"),
            ("custom/orphan.xml", b"orphan"),
        ]);

        let per_part = ReadLimits::builder()
            .max_part_bytes(b"document".len() as u64 - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPackage::from_read_at_with_limits(
                Arc::new(CountingSource::new(bytes.clone())),
                per_part,
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                ..
            })
        ));

        let aggregate = ReadLimits::builder()
            .max_total_part_bytes((b"document".len() + b"orphan".len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPackage::from_read_at_with_limits(
                Arc::new(CountingSource::new(bytes.clone())),
                aggregate,
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                ..
            })
        ));

        let part_count = ReadLimits::builder().max_parts(1).unwrap().build().unwrap();
        assert!(matches!(
            SourceBackedPackage::from_read_at_with_limits(
                Arc::new(CountingSource::new(bytes)),
                part_count,
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                ..
            })
        ));
    }

    #[test]
    fn source_backed_part_limit_precedes_later_derived_or_equivalent_conflicts() {
        let oversized = b"oversized first part";
        let cases = [
            (
                [
                    ("word/oversized.xml", &oversized[..]),
                    ("word/container.xml", &b"container"[..]),
                    ("word/container.xml/child.xml", &b"child"[..]),
                ],
                true,
            ),
            (
                [
                    ("word/oversized.xml", &oversized[..]),
                    ("word/equivalent.xml", &b"first"[..]),
                    ("WORD/EQUIVALENT.XML", &b"second"[..]),
                ],
                false,
            ),
        ];
        let limits = ReadLimits::builder()
            .max_part_bytes(oversized.len() as u64 - 1)
            .unwrap()
            .build()
            .unwrap();

        for (parts, derived) in cases {
            let bytes = archive_with_ordered_xml_parts(&parts);
            assert!(matches!(
                SourceBackedPackage::from_read_at_with_limits(
                    Arc::new(CountingSource::new(bytes.clone())),
                    limits,
                ),
                Err(OpcError::ReadLimit {
                    resource: ReadResource::PartBytes,
                    ..
                })
            ));
            let eager_error = match OpcPackage::from_bytes_with_limits(&bytes, limits) {
                Ok(_) => panic!("eager package unexpectedly accepted a name conflict"),
                Err(error) => error,
            };
            if derived {
                assert!(matches!(eager_error, OpcError::DerivedPartNames { .. }));
            } else {
                assert!(matches!(eager_error, OpcError::EquivalentPartNames { .. }));
            }
        }
    }

    #[test]
    fn source_backed_part_limit_precedes_later_max_parts_error() {
        let oversized = b"oversized first part";
        let bytes = archive_with_ordered_xml_parts(&[
            ("word/oversized.xml", oversized),
            ("word/later.xml", b"later"),
        ]);
        let limits = ReadLimits::builder()
            .max_parts(1)
            .unwrap()
            .max_part_bytes(oversized.len() as u64 - 1)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            SourceBackedPackage::from_read_at_with_limits(
                Arc::new(CountingSource::new(bytes.clone())),
                limits,
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                ..
            })
        ));
        assert!(matches!(
            OpcPackage::from_bytes_with_limits(&bytes, limits),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                ..
            })
        ));
    }

    #[test]
    fn part_lookup_prefers_exact_names_and_matches_ascii_case_insensitive_targets() {
        let exact_name = PackURI::new("/word/document.xml").unwrap();
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_with_part_names("word/document.xml", &["word/document.xml"]),
        )))
        .unwrap();

        assert_eq!(package.part(&exact_name).unwrap().partname(), &exact_name);
        let case_variant = PackURI::new("/WORD/DOCUMENT.XML").unwrap();
        assert_eq!(package.part(&case_variant).unwrap().partname(), &exact_name);

        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_with_part_names("WORD/DOCUMENT.XML", &["word/document.xml"]),
        )))
        .unwrap();
        assert_eq!(
            package.main_document_part().unwrap().partname(),
            &exact_name
        );
    }

    #[test]
    fn source_catalog_still_rejects_case_equivalent_part_names() {
        let bytes = archive_with_part_names(
            "word/document.xml",
            &["word/document.xml", "WORD/DOCUMENT.XML"],
        );
        let Err(error) = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(bytes)))
        else {
            panic!("case-equivalent parts must remain ambiguous");
        };
        assert!(matches!(error, OpcError::EquivalentPartNames { .. }));
    }

    #[test]
    fn cache_hits_pin_payloads_and_failures_are_not_retained() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cached payload",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let part = package.main_document_part().unwrap();
        let first = part.data().unwrap();
        let after_first = source.reads.load(Ordering::SeqCst);
        let second = part.data().unwrap();
        assert_eq!(source.reads.load(Ordering::SeqCst), after_first);
        assert!(first.shares_allocation_with(&second));
        assert_eq!(
            package.cache_diagnostics(),
            SourceCacheDiagnostics {
                hits: 1,
                cold_loads: 1,
                successful_loads: 1,
                retained_entries: 1,
                retained_bytes: b"cached payload".len(),
                ..SourceCacheDiagnostics::default()
            }
        );

        const DOCUMENT: &[u8] = b"never cache a failed read";
        let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[position] ^= 0xff;
        let corrupt_source = Arc::new(CountingSource::new(bytes));
        let corrupt_package = SourceBackedPackage::from_read_at(corrupt_source.clone()).unwrap();
        let corrupt_part = corrupt_package.main_document_part().unwrap();
        assert!(corrupt_part.data().is_err());
        let after_failure = corrupt_source.reads.load(Ordering::SeqCst);
        assert!(corrupt_part.data().is_err());
        assert!(corrupt_source.reads.load(Ordering::SeqCst) > after_failure);
        let diagnostics = corrupt_package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 2);
        assert_eq!(diagnostics.failed_loads, 2);
        assert_eq!(diagnostics.retained_entries, 0);
    }

    #[test]
    fn managed_cache_reserves_declared_payload_and_releases_on_package_drop() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed payload",
            false,
        )));
        let (budget, context) = managed_context(1024);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        assert_eq!(budget.used(Resource::Memory), 0);

        let data = package.main_document_part().unwrap().data().unwrap();
        assert_eq!(data.as_bytes(), b"managed payload");
        assert_eq!(
            budget.used(Resource::Memory),
            b"managed payload".len() as u64
        );
        let diagnostics = package.cache_diagnostics();
        assert!(diagnostics.budget_managed);
        assert_eq!(diagnostics.budget_reservation_failures, 0);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            b"managed payload".len() as u64
        );
        assert_eq!(
            diagnostics.budget_memory_used,
            b"managed payload".len() as u64
        );

        drop(data);
        // The clean cache entry owns the reservation until the package drops.
        assert_eq!(
            budget.used(Resource::Memory),
            b"managed payload".len() as u64
        );
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_input_bytes_are_exact_for_chunked_reads_and_cache_hits_are_free() {
        const DOCUMENT: &[u8] = b"short-read physical input accounting";
        let source = Arc::new(CountingSource::chunked(
            archive_bytes(root_relationships(), DOCUMENT, false),
            3,
        ));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        assert_eq!(
            budget.used(Resource::InputBytes),
            source.read_bytes.load(Ordering::SeqCst)
        );
        let first = package.main_document_part().unwrap().data().unwrap();
        let after_cold = source.read_bytes.load(Ordering::SeqCst);
        assert_eq!(budget.used(Resource::InputBytes), after_cold);
        let second = package.main_document_part().unwrap().data().unwrap();
        assert!(first.shares_allocation_with(&second));
        assert_eq!(source.read_bytes.load(Ordering::SeqCst), after_cold);
        assert_eq!(
            package.cache_diagnostics().budget_input_bytes_used,
            after_cold
        );
        drop(second);
        drop(first);
        drop(package);
        // InputBytes and Work are cumulative; only retained resources release
        // when the package and its handles are dropped.
        assert_eq!(budget.used(Resource::InputBytes), after_cold);
    }

    #[test]
    fn managed_input_budget_refusal_counts_only_the_terminal_read_reservation() {
        const DOCUMENT: &[u8] = b"input reservation refusal accounting";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context.clone(),
        )
        .unwrap();
        let input_before = budget.used(Resource::InputBytes);
        context
            .consume(Resource::InputBytes, u64::MAX - input_before)
            .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::InputBytes
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::InputBytes), u64::MAX);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_work_one_under_refuses_payload_before_physical_io() {
        const DOCUMENT: &[u8] = b"work preflight one under";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, (DOCUMENT.len() - 1) as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Work
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Work), 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_object_one_under_refuses_payload_before_physical_io() {
        const DOCUMENT: &[u8] = b"object preflight one under";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        // archive_bytes(false) contains four non-directory members and one
        // package-level catalog owner is retained by SourceBackedPackage.
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, 5, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Objects
        ));
        // The catalog reservation is retained by the package; the one-under
        // payload-object preflight happens before any ordinary payload read.
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Objects), 5);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_failed_cold_load_consumes_work_and_input_but_releases_retained_objects() {
        const DOCUMENT: &[u8] = b"managed failed cold-load accounting";
        let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[position] ^= 0xff;
        let source = Arc::new(CountingSource::new(bytes));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.read_bytes.load(Ordering::SeqCst);
        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::ZipError(_))
        ));
        let reads_after = source.read_bytes.load(Ordering::SeqCst);
        assert!(reads_after > reads_before);
        assert_eq!(budget.used(Resource::InputBytes), reads_after);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_cancellation_between_cache_admission_and_source_read_is_typed() {
        const DOCUMENT: &[u8] = b"cancel before managed source read";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let part = package.main_document_part().unwrap();
        let index = part.index;
        let entry_id = package.parts[index].entry_id;
        let declared = package
            .archive
            .metadata_for(entry_id)
            .unwrap()
            .uncompressed_size();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let flight = match package.cache.enter(entry_id, declared).unwrap() {
            CacheAccess::Loader(flight) => flight,
            CacheAccess::Hit(_) | CacheAccess::Waiter(_) | CacheAccess::Bypass(_) => {
                panic!("fresh managed Part must become the loader")
            },
        };
        // The loader has charged its bounded cold-load resources, but has not
        // entered SourceReader yet. This is the exact race boundary where a
        // cancellation must survive ZIP's std::io::Error conversion.
        cancellation_source.cancel();
        let error = package
            .load_part_with_accounting(index, entry_id, Some(declared), Some(flight), None, None)
            .unwrap_err();
        assert!(matches!(error, OpcError::Cancelled));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_cancellation_before_preservation_source_read_is_typed() {
        const DOCUMENT: &[u8] = b"<before/>";
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let source = Arc::new(CancelOnHitVersionSource::new(
            archive_bytes(root_relationships(), DOCUMENT, false),
            cancellation_source,
        ));
        let budget = Budget::root(
            "opc-source-cache-preservation-cancel-test",
            Limits::new(4096, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(4096).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let target = package.main_document_part().unwrap().partname().clone();
        source.arm_after_cache_enter();
        let mut output = Vec::new();
        let error = package
            .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
            .unwrap_err();
        assert!(matches!(error, OpcError::Cancelled));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_constructor_honors_pre_cancellation_without_source_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cancelled before open",
            false,
        )));
        let reads_before = source.reads.load(Ordering::SeqCst);
        let budget = Budget::root(
            "opc-source-cache-cancel-test",
            Limits::new(1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        cancellation_source.cancel();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(1024).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget, cancellation, execution_limits);

        assert!(matches!(
            SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(OpcError::Cancelled)
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    }

    #[test]
    fn managed_cache_hit_honors_cancellation_without_releasing_cached_budget() {
        const DOCUMENT: &[u8] = b"managed cancellation hit";
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let source = Arc::new(CancelOnHitVersionSource::new(
            archive_bytes(root_relationships(), DOCUMENT, false),
            cancellation_source,
        ));
        let budget = Budget::root(
            "opc-source-cache-hit-cancel-test",
            Limits::new(
                DOCUMENT.len() as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let first = package.main_document_part().unwrap().data().unwrap();
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        source.arm_after_cache_enter();

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Cancelled)
        ));
        assert_eq!(package.cache_diagnostics().hits, 1);
        // Cancellation rejects the handle request, but never steals the
        // clean entry's reservation from the package-owned cache.
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        drop(first);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_cache_eviction_releases_unpinned_reservation_before_next_read() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        )));
        let (budget, context) = managed_context(64);
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                SourceCacheLimits::new(9, 2).unwrap(),
                context,
            )
            .unwrap();
        let first_name = package.parts[0].partname.clone();
        let second_name = package.parts[1].partname.clone();
        let first = package.part(&first_name).unwrap().data().unwrap();
        assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
        drop(first);

        let second = package.part(&second_name).unwrap().data().unwrap();
        assert_eq!(second.as_bytes(), b"<orphan/>");
        assert_eq!(budget.used(Resource::Memory), b"<orphan/>".len() as u64);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.evictions, 1);
        assert_eq!(diagnostics.retained_bytes, b"<orphan/>".len());
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            b"<orphan/>".len() as u64
        );
        drop(second);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_cache_does_not_evict_externally_pinned_entry_and_bypasses_retention() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        )));
        let (budget, context) = managed_context(64);
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                SourceCacheLimits::new(b"document".len(), 1).unwrap(),
                context,
            )
            .unwrap();
        let first_name = package.parts[0].partname.clone();
        let second_name = package.parts[1].partname.clone();
        let first = package.part(&first_name).unwrap().data().unwrap();
        let first_error = first.into_arc().expect_err("managed Arc escape must fail");
        assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
        assert_eq!(
            package.cache_diagnostics().retained_bytes,
            b"document".len()
        );
        let second = package.part(&second_name).unwrap().data().unwrap();

        assert!(matches!(first_error, OpcError::ManagedPartDataArcEscape));
        assert_eq!(first.as_bytes(), b"document");
        assert_eq!(second.as_bytes(), b"<orphan/>");
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.retained_bytes, b"document".len());
        assert_eq!(diagnostics.bypasses, 1);
        assert_eq!(
            budget.used(Resource::Memory),
            (b"document".len() + b"<orphan/>".len()) as u64
        );
        assert_eq!(budget.used(Resource::Objects), 7);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
        assert!(
            package
                .cache
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&package.parts[0].entry_id)
        );

        drop(second);
        drop(first);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_budget_rejects_before_payload_io_and_reports_content_free_failure() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"payload too large",
            false,
        )));
        let (budget, context) = managed_context((b"payload too large".len() - 1) as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Memory), 0);
        let diagnostics = package.cache_diagnostics();
        assert!(diagnostics.budget_managed);
        assert_eq!(diagnostics.budget_reservation_failures, 2);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.retained_bytes, 0);
    }

    #[test]
    fn managed_cache_respects_hierarchical_parent_memory_limit() {
        const DOCUMENT: &[u8] = b"hierarchical budget payload";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let root = Budget::root(
            "opc-source-cache-root",
            Limits::new(
                (DOCUMENT.len() - 1) as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let child = root.child(
            "opc-source-cache-child",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        drop(cancellation_source);
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(child, cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(root.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 2);
    }

    #[test]
    fn managed_sibling_caches_compete_for_parent_memory_before_payload_io() {
        const DOCUMENT: &[u8] = b"sibling parent budget payload";
        let root = Budget::root(
            "opc-source-cache-sibling-root",
            Limits::new(
                DOCUMENT.len() as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let first_budget = root.child(
            "opc-source-cache-sibling-first",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let second_budget = root.child(
            "opc-source-cache-sibling-second",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let (first_cancellation_source, first_cancellation) = CancellationSource::pair();
        let (second_cancellation_source, second_cancellation) = CancellationSource::pair();
        let first_context =
            ExecutionContext::new(first_budget, first_cancellation, execution_limits);
        let second_context =
            ExecutionContext::new(second_budget, second_cancellation, execution_limits);
        drop(first_cancellation_source);
        drop(second_cancellation_source);
        let first_source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let second_source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let first_package = SourceBackedPackage::from_read_at_with_execution_context(
            first_source,
            ReadLimits::default(),
            first_context,
        )
        .unwrap();
        let second_package = SourceBackedPackage::from_read_at_with_execution_context(
            second_source.clone(),
            ReadLimits::default(),
            second_context,
        )
        .unwrap();
        let reads_before = second_source.reads.load(Ordering::SeqCst);
        let first = first_package.main_document_part().unwrap().data().unwrap();
        assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
        drop(first);

        assert!(matches!(
            second_package.main_document_part().unwrap().data(),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(second_source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
        assert_eq!(
            second_package
                .cache_diagnostics()
                .budget_reservation_failures,
            2
        );
        drop(second_package);
        drop(first_package);
        assert_eq!(root.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_same_part_waiters_share_one_reservation_and_flight() {
        const DOCUMENT: &[u8] = b"managed single-flight payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let (budget, context) = managed_context(DOCUMENT.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package.main_document_part().unwrap().data().unwrap()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package.main_document_part().unwrap().data().unwrap()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            let diagnostics = package.cache_diagnostics();
            assert_eq!(diagnostics.in_flight_loads, 1);
            assert_eq!(
                diagnostics.budget_cache_reserved_bytes,
                DOCUMENT.len() as u64
            );
            assert_eq!(diagnostics.budget_cache_reserved_objects, 2);
            assert_eq!(budget.used(Resource::Objects), 7);
            (first_task.join().unwrap(), second_task.join().unwrap())
        });
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        assert!(first.shares_allocation_with(&second));
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(diagnostics.successful_loads, 1);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            DOCUMENT.len() as u64
        );
        assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
        drop(first);
        drop(second);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_waiter_cancellation_does_not_block_or_publish_loader_payload() {
        const DOCUMENT: &[u8] = b"managed waiter cancellation payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(DOCUMENT.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package.main_document_part().unwrap().data()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package.main_document_part().unwrap().data()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
            cancellation_source.cancel();
            (first_task.join().unwrap(), second_task.join().unwrap())
        });

        assert!(matches!(first, Err(OpcError::Cancelled)));
        assert!(matches!(second, Err(OpcError::Cancelled)));
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 0);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert!(budget.used(Resource::InputBytes) > 0);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_source_change_drops_reservation_and_does_not_retain_payload() {
        const DOCUMENT: &[u8] = b"source changes during managed read";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
        let (budget, context) = managed_context(1024);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        source.armed.store(true, Ordering::SeqCst);

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert!(budget.used(Resource::InputBytes) > 0);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    }

    #[test]
    fn managed_cancellation_after_decompression_prevents_publication() {
        const DOCUMENT: &[u8] = b"managed prepublication cancellation payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(DOCUMENT.len() as u64);
        let source = Arc::new(CancelDuringPayloadSource::new(
            bytes,
            payload_offset,
            cancellation_source,
        ));
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Cancelled)
        ));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 0);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_open_checks_cancellation_after_indexing_before_catalog_publication() {
        let (budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
        let source = Arc::new(CancelDuringOpenSource::new(
            archive_bytes(root_relationships(), b"open cancellation", false),
            cancellation_source,
        ));
        assert!(matches!(
            SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(OpcError::Cancelled)
        ));
        assert!(source.reads.load(Ordering::SeqCst) > 0);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn concurrent_cold_reads_share_one_archive_load_and_one_arc() {
        const DOCUMENT: &[u8] = b"single-flight source-backed payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package
                    .main_document_part()
                    .unwrap()
                    .data()
                    .unwrap()
                    .into_arc()
                    .unwrap()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package
                    .main_document_part()
                    .unwrap()
                    .data()
                    .unwrap()
                    .into_arc()
                    .unwrap()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
            (first_task.join().unwrap(), second_task.join().unwrap())
        });
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        assert!(Arc::ptr_eq(&first, &second));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(diagnostics.successful_loads, 1);
        assert_eq!(diagnostics.in_flight_loads, 0);
    }

    #[test]
    fn cache_evicts_by_byte_weight_and_entry_count_and_rejects_oversized_values() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"document", false),
        )))
        .unwrap();
        let first_id = package.parts[0].entry_id;
        let second_id = package.parts[1].entry_id;
        let cache = PartCache::new(SourceCacheLimits::new(3, 3).unwrap());
        let first = Arc::new(vec![1, 2]);
        cache
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::clone(&first),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(Arc::ptr_eq(
            &cache.state.lock().unwrap().entries[&first_id].payload.bytes,
            &first
        ));
        drop(first);
        cache
            .complete_bypass_success(
                second_id,
                CachedPayload {
                    bytes: Arc::new(vec![3, 4]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
        assert!(cache.state.lock().unwrap().entries.contains_key(&second_id));

        let entry_limited = PartCache::new(SourceCacheLimits::new(10, 1).unwrap());
        entry_limited
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::new(vec![1, 2]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        entry_limited
            .complete_bypass_success(
                second_id,
                CachedPayload {
                    bytes: Arc::new(vec![3, 4]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(
            !entry_limited
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&first_id)
        );
        assert!(
            entry_limited
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&second_id)
        );

        cache
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::new(vec![0, 0, 0, 0]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
        assert_eq!(cache.diagnostics().evictions, 1);
        assert_eq!(cache.diagnostics().oversized_bypasses, 1);
    }

    #[test]
    fn cache_limits_reject_zero_bounds() {
        assert_eq!(
            SourceCacheLimits::new(0, 1),
            Err(SourceCacheLimitError::ZeroMaximumBytes)
        );
        assert_eq!(
            SourceCacheLimits::new(1, 0),
            Err(SourceCacheLimitError::ZeroMaximumEntries)
        );
    }

    #[test]
    fn checked_counter_delta_rejects_a_regression() {
        let before = SourceCacheDiagnostics {
            hits: 4,
            ..SourceCacheDiagnostics::default()
        };
        let after = SourceCacheDiagnostics {
            hits: 3,
            ..SourceCacheDiagnostics::default()
        };

        assert_eq!(
            SourceCacheDiagnostics::checked_counter_delta(before, after),
            Err(SourceCacheDiagnosticsError::CounterMovedBackwards { counter: "hits" })
        );
    }

    #[test]
    fn checked_counter_delta_preserves_event_counts_without_gauge_subtraction() {
        let before = SourceCacheDiagnostics {
            hits: 2,
            retained_entries: 1,
            retained_bytes: 10,
            ..SourceCacheDiagnostics::default()
        };
        let after = SourceCacheDiagnostics {
            hits: 5,
            retained_entries: 0,
            retained_bytes: 0,
            ..SourceCacheDiagnostics::default()
        };

        assert_eq!(
            SourceCacheDiagnostics::checked_counter_delta(before, after)
                .expect("valid counter interval"),
            SourceCacheCounterDelta {
                hits: 3,
                ..SourceCacheCounterDelta::default()
            }
        );
    }

    #[test]
    fn checked_counter_overflow_fails_closed_snapshot() {
        let cache = PartCache::new(SourceCacheLimits::new(3, 3).unwrap());
        cache.counters.hits.value.store(u64::MAX, Ordering::Relaxed);
        cache.counters.hits.increment();

        assert_eq!(
            cache.try_diagnostics(),
            Err(SourceCacheDiagnosticsError::CounterOverflow)
        );
        assert_eq!(cache.counters.hits.load(), u64::MAX);
    }

    #[test]
    fn try_cache_diagnostics_matches_compatibility_snapshot_when_healthy() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"healthy diagnostics", false),
        )))
        .unwrap();

        assert_eq!(
            package
                .try_cache_diagnostics()
                .expect("healthy diagnostic snapshot"),
            package.cache_diagnostics()
        );
    }

    #[test]
    fn managed_input_reservation_counter_overflow_fails_closed_through_wrapper() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed input diagnostics",
            false,
        )));
        let (_budget, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let counter = package
            .cache
            .input_reservation_failures
            .as_ref()
            .expect("managed input reservation counter");
        counter.value.store(u64::MAX, Ordering::Relaxed);
        counter.increment();

        assert_eq!(
            package.try_cache_diagnostics(),
            Err(SourceCacheDiagnosticsError::CounterOverflow)
        );
    }

    #[test]
    fn managed_output_reservation_counter_overflow_fails_closed_through_wrapper() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed output diagnostics",
            false,
        )));
        let (_budget, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let counter = package
            .cache
            .output_reservation_failures
            .as_ref()
            .expect("managed output reservation counter");
        counter.value.store(u64::MAX, Ordering::Relaxed);
        counter.increment();

        assert_eq!(
            package.try_cache_diagnostics(),
            Err(SourceCacheDiagnosticsError::CounterOverflow)
        );
    }

    #[test]
    fn aggregate_reservation_counter_overflow_fails_closed_through_wrapper() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed aggregate diagnostics",
            false,
        )));
        let (_budget, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        package
            .cache
            .counters
            .budget_reservation_failures
            .value
            .store(u64::MAX, Ordering::Relaxed);
        package
            .cache
            .input_reservation_failures
            .as_ref()
            .expect("managed input reservation counter")
            .value
            .store(1, Ordering::Relaxed);

        assert_eq!(
            package.try_cache_diagnostics(),
            Err(SourceCacheDiagnosticsError::CounterOverflow)
        );
    }

    #[test]
    fn retained_gauge_aggregation_overflow_fails_closed_through_wrapper() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_with_part_names(
                "word/document.xml",
                &["word/document.xml", "custom/other.xml"],
            ),
        )))
        .unwrap();
        assert_eq!(package.parts.len(), 2);

        let bytes_budget_a = Budget::root(
            "diagnostic-gauge-bytes-a",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let bytes_budget_b = Budget::root(
            "diagnostic-gauge-bytes-b",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let objects_budget_a = Budget::root(
            "diagnostic-gauge-objects-a",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let objects_budget_b = Budget::root(
            "diagnostic-gauge-objects-b",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let bytes_a = Arc::new(
            bytes_budget_a
                .reserve(Resource::Memory, u64::MAX)
                .expect("maximum diagnostic byte reservation"),
        );
        let bytes_b = Arc::new(
            bytes_budget_b
                .reserve(Resource::Memory, u64::MAX)
                .expect("maximum diagnostic byte reservation"),
        );
        let objects_a = Arc::new(
            objects_budget_a
                .reserve(Resource::Objects, u64::MAX)
                .expect("maximum diagnostic object reservation"),
        );
        let objects_b = Arc::new(
            objects_budget_b
                .reserve(Resource::Objects, u64::MAX)
                .expect("maximum diagnostic object reservation"),
        );
        let mut state = package.cache.state.lock().expect("unpoisoned cache state");
        for (index, (entry_id, bytes, objects)) in [
            (package.parts[0].entry_id, bytes_a, objects_a),
            (package.parts[1].entry_id, bytes_b, objects_b),
        ]
        .into_iter()
        .enumerate()
        {
            state.entries.insert(
                entry_id,
                CacheEntry {
                    payload: CachedPayload {
                        bytes: Arc::new(Vec::new()),
                        reservation: Some(bytes),
                        object_reservation: Some(objects),
                    },
                    last_used: u64::try_from(index).expect("bounded test index"),
                },
            );
        }
        drop(state);

        assert_eq!(
            package.try_cache_diagnostics(),
            Err(SourceCacheDiagnosticsError::CounterOverflow)
        );
    }

    #[test]
    fn poisoned_cache_state_fails_closed_snapshot() {
        let cache = Arc::new(PartCache::new(SourceCacheLimits::new(3, 3).unwrap()));
        let poisoner = Arc::clone(&cache);
        let join = std::thread::spawn(move || {
            let _state = poisoner.state.lock().expect("unpoisoned cache state");
            panic!("test cache-state poison");
        });
        assert!(join.join().is_err());

        assert_eq!(
            cache.try_diagnostics(),
            Err(SourceCacheDiagnosticsError::StatePoisoned)
        );
    }

    #[test]
    fn poisoned_cache_state_fails_closed_through_wrapper() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"poisoned diagnostics", false),
        )))
        .unwrap();
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let _state = package.cache.state.lock().expect("unpoisoned cache state");
            panic!("test cache-state poison through package wrapper");
        }));
        assert!(result.is_err());

        assert_eq!(
            package.try_cache_diagnostics(),
            Err(SourceCacheDiagnosticsError::StatePoisoned)
        );
    }

    #[test]
    fn catalog_entry_id_resolves_to_the_same_payload_as_its_member_name() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"entry identity", false),
        )))
        .unwrap();
        let part = package.main_document_part().unwrap();
        let catalog = package
            .parts
            .iter()
            .find(|catalog| catalog.partname == *part.partname())
            .unwrap();
        assert_eq!(
            package.archive.read_entry(catalog.entry_id).unwrap(),
            package.archive.read(catalog.partname.membername()).unwrap()
        );
    }

    #[test]
    fn source_changes_reject_metadata_cache_and_conversion_access() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"stable payload",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let part = package.main_document_part().unwrap();
        part.data().unwrap();
        source.changed();
        assert!(matches!(
            package.part(part.partname()),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(matches!(part.data(), Err(OpcError::SourceChanged { .. })));
        assert!(matches!(
            package.into_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
    }

    #[test]
    fn catalog_reports_non_parts_and_conversion_matches_loaded_parts() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        assert_eq!(package.iter_parts().count(), 2);
        assert_eq!(package.non_part_members().len(), 1);
        assert_eq!(package.non_part_members()[0].name(), "scratch.bin");
        assert_eq!(
            package
                .main_document_part()
                .unwrap()
                .data()
                .unwrap()
                .as_bytes(),
            b"document"
        );
        let owned = package.into_opc_package().unwrap();
        assert_eq!(owned.part_count(), 2);
        assert_eq!(owned.non_part_members().len(), 1);
        assert_eq!(owned.main_document_part().unwrap().blob(), b"document");
    }

    #[test]
    fn borrowed_materialization_preserves_source_and_matches_eager_graph() {
        let bytes = archive_bytes(root_relationships(), b"document", true);
        let source = Arc::new(CountingSource::new(bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let owned = package.to_opc_package().unwrap();
        let eager = OpcPackage::from_bytes(&bytes).unwrap();

        assert_eq!(owned.part_count(), eager.part_count());
        assert_eq!(owned.non_part_members(), eager.non_part_members());
        assert_eq!(owned.rels().iter().count(), eager.rels().iter().count());
        for partname in [
            PackURI::new("/word/document.xml").unwrap(),
            PackURI::new("/custom/orphan.xml").unwrap(),
        ] {
            assert_eq!(
                owned.get_part(&partname).unwrap().blob(),
                eager.get_part(&partname).unwrap().blob()
            );
        }

        // The borrowed conversion must leave the source-backed catalog usable;
        // this second read is a cache hit and therefore does not touch the
        // positional source again.
        let reads_after_materialization = source.reads.load(Ordering::SeqCst);
        assert_eq!(
            package
                .main_document_part()
                .unwrap()
                .data()
                .unwrap()
                .as_bytes(),
            b"document"
        );
        assert_eq!(
            source.reads.load(Ordering::SeqCst),
            reads_after_materialization
        );
    }

    #[test]
    fn borrowed_materialization_rejects_managed_package_before_payload_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed document",
            true,
        )));
        let (budget, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let read_bytes_before = source.read_bytes.load(Ordering::SeqCst);

        assert!(matches!(
            package.to_opc_package(),
            Err(OpcError::ManagedPackageMaterialization)
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(source.read_bytes.load(Ordering::SeqCst), read_bytes_before);
        assert!(package.main_document_part().is_ok());
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn consuming_materialization_rejects_managed_package_before_payload_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed consuming document",
            true,
        )));
        let (budget, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let read_bytes_before = source.read_bytes.load(Ordering::SeqCst);

        assert!(matches!(
            package.into_opc_package(),
            Err(OpcError::ManagedPackageMaterialization)
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(source.read_bytes.load(Ordering::SeqCst), read_bytes_before);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn borrowed_materialization_rejects_stale_source_before_payload_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"stable document",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        source.changed();

        assert!(matches!(
            package.to_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    }

    #[test]
    fn borrowed_materialization_rejects_source_changed_during_payload_and_keeps_no_partial_result()
    {
        const DOCUMENT: &[u8] = b"document changes during materialization";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, true);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        source.armed.store(true, Ordering::SeqCst);

        assert!(matches!(
            package.to_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(package.main_document_part().is_err());
    }

    #[test]
    fn borrowed_materialization_rejects_cancellation_before_payload_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cancelled document",
            false,
        )));
        let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        cancellation_source.cancel();

        assert!(matches!(package.to_opc_package(), Err(OpcError::Cancelled)));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn borrowed_materialization_reports_malformed_payload_without_partial_result() {
        const DOCUMENT: &[u8] = b"payload whose CRC is invalid";
        let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[payload_offset] ^= 0xff;
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

        assert!(matches!(
            package.to_opc_package(),
            Err(OpcError::ZipError(_))
        ));
        assert!(package.main_document_part().is_ok());
        assert!(package.main_document_part().unwrap().data().is_err());
    }

    #[test]
    fn borrowed_materialization_preserves_opaque_xml_like_eager_unmarshal() {
        const DOCUMENT: &[u8] = b"<document> <child/></document>";
        let source_bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let source_package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let materialized = source_package.to_opc_package().unwrap();
        let eager = OpcPackage::from_bytes(&source_bytes).unwrap();

        let mut materialized_bytes = Vec::new();
        materialized.to_stream(&mut materialized_bytes).unwrap();
        let mut eager_bytes = Vec::new();
        eager.to_stream(&mut eager_bytes).unwrap();
        let materialized_reader =
            crate::phys_pkg::OwnedPhysPkgReader::from_bytes(materialized_bytes).unwrap();
        let eager_reader = crate::phys_pkg::OwnedPhysPkgReader::from_bytes(eager_bytes).unwrap();
        assert_eq!(
            materialized_reader
                .read_member("word/document.xml")
                .unwrap(),
            DOCUMENT
        );
        assert_eq!(
            materialized_reader
                .read_member("word/document.xml")
                .unwrap(),
            eager_reader.read_member("word/document.xml").unwrap()
        );
    }

    #[test]
    fn borrowed_signed_materialization_retains_publication_policy() {
        let source = signed_archive(b"<signed/>");
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source))).unwrap();
        let materialized = package.to_opc_package().unwrap();
        assert!(materialized.is_signed());

        let mut output = Vec::new();
        assert!(matches!(
            materialized.to_stream(&mut output),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_raw_copies_every_unselected_member_and_reopens() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let source_raw = raw_records(&source_bytes);
        let source = Arc::new(CountingSource::new(source_bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(reopened.get_part(&target).unwrap().blob(), b"<after/>");
        assert_eq!(reopened.part_count(), 2);
        assert_eq!(reopened.non_part_members().len(), 1);
        assert_eq!(reopened.non_part_members()[0].name(), "scratch.bin");
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if name == b"word/document.xml" {
                assert_ne!(output_raw[&name].local, source_record.local);
            } else {
                assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
                assert_eq!(
                    central_without_local_offset(&output_raw[&name].central),
                    central_without_local_offset(&source_record.central),
                    "{name:?}"
                );
            }
        }
    }

    #[test]
    fn shared_single_overlay_matches_vec_changed_and_signed_noop_publication() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut vec_output = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
            .unwrap()
            .write_part_overlay_to_stream(&mut vec_output, &target, b"<after/>".to_vec())
            .unwrap();
        let mut shared_output = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes)))
            .unwrap()
            .write_part_overlay_shared_to_stream(
                &mut shared_output,
                &target,
                Arc::new(b"<after/>".to_vec()),
            )
            .unwrap();
        assert_eq!(shared_output, vec_output);

        let signed_bytes = signed_archive(b"<signed/>");
        let signed_target = PackURI::new("/word/document.xml").unwrap();
        let mut vec_noop = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(signed_bytes.clone())))
            .unwrap()
            .write_part_overlay_to_stream(&mut vec_noop, &signed_target, b"<signed/>".to_vec())
            .unwrap();
        let mut shared_noop = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(signed_bytes.clone())))
            .unwrap()
            .write_part_overlay_shared_to_stream(
                &mut shared_noop,
                &signed_target,
                Arc::new(b"<signed/>".to_vec()),
            )
            .unwrap();
        assert_eq!(vec_noop, signed_bytes);
        assert_eq!(shared_noop, vec_noop);
    }

    #[test]
    fn shared_multi_overlay_matches_vec_and_reopens() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();
        let mut vec_output = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
            .unwrap()
            .write_part_overlays_to_stream(
                &mut vec_output,
                vec![
                    (orphan.clone(), b"<orphan-after/>".to_vec()),
                    (document.clone(), b"<document-after/>".to_vec()),
                ],
            )
            .unwrap();

        let mut shared_output = Vec::new();
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes)))
            .unwrap()
            .write_part_overlays_shared_to_stream(
                &mut shared_output,
                vec![
                    (orphan.clone(), Arc::new(b"<orphan-after/>".to_vec())),
                    (document.clone(), Arc::new(b"<document-after/>".to_vec())),
                ],
            )
            .unwrap();
        assert_eq!(shared_output, vec_output);
        let reopened = OpcPackage::from_bytes(&shared_output).unwrap();
        assert_eq!(
            reopened.get_part(&document).unwrap().blob(),
            b"<document-after/>"
        );
        assert_eq!(
            reopened.get_part(&orphan).unwrap().blob(),
            b"<orphan-after/>"
        );
    }

    #[test]
    fn shared_overlay_preserves_duplicate_and_limit_refusals_before_output() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let document = PackURI::new("/word/document.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_shared_to_stream(
                &mut output,
                vec![
                    (document.clone(), Arc::new(b"<first/>".to_vec())),
                    (document.clone(), Arc::new(b"<second/>".to_vec())),
                ],
            ),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(output.is_empty());

        let limits = ReadLimits::builder()
            .max_part_bytes(10)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            limits,
        )
        .unwrap();
        assert!(matches!(
            package.write_part_overlay_shared_to_stream(
                &mut output,
                &document,
                Arc::new(vec![b'x'; 11]),
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 11,
                maximum: 10,
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn shared_overlay_reports_partial_sink_failure_with_bounded_writes() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut sink = BoundedFailingSink {
            accepted: 0,
            limit: 100,
            largest_write: 0,
        };
        let error = package
            .write_part_overlay_shared_to_stream(&mut sink, &target, Arc::new(b"<after/>".to_vec()))
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
            other => panic!("unexpected sink error: {other:?}"),
        }
        assert!(sink.largest_write <= SOURCE_PUBLICATION_CHUNK_BYTES);
    }

    #[test]
    fn multi_part_overlay_changes_only_selected_raw_members_and_reopens() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let source_raw = raw_records(&source_bytes);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlays_to_stream(
                &mut output,
                vec![
                    (orphan.clone(), b"<orphan-after/>".to_vec()),
                    (document.clone(), b"<document-after/>".to_vec()),
                ],
            )
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(
            reopened.get_part(&document).unwrap().blob(),
            b"<document-after/>"
        );
        assert_eq!(
            reopened.get_part(&orphan).unwrap().blob(),
            b"<orphan-after/>"
        );
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if matches!(name.as_slice(), b"word/document.xml" | b"custom/orphan.xml") {
                assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
            } else {
                assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
                assert_eq!(
                    central_without_local_offset(&output_raw[&name].central),
                    central_without_local_offset(&source_record.central),
                    "{name:?}"
                );
            }
        }
    }

    #[test]
    fn multi_part_relationship_removal_has_exact_and_one_over_aggregate_bounds() {
        let exact_each = MAX_SOURCE_RELATIONSHIP_REMOVALS / 2;
        let (source_bytes, first_ids, second_ids) =
            archive_with_relationship_batches(exact_each, exact_each);
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlays_with_external_relationship_removals_to_stream(
                &mut output,
                vec![
                    (document.clone(), b"<after/>".to_vec(), first_ids),
                    (orphan.clone(), b"<after/>".to_vec(), second_ids),
                ],
            )
            .unwrap();
        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert!(reopened.get_part(&document).unwrap().rels().is_empty());
        assert!(reopened.get_part(&orphan).unwrap().rels().is_empty());

        let (source_bytes, first_ids, second_ids) =
            archive_with_relationship_batches(exact_each, exact_each + 1);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        output.clear();
        assert!(matches!(
            package.write_part_overlays_with_external_relationship_removals_to_stream(
                &mut output,
                vec![
                    (document, b"<after/>".to_vec(), first_ids),
                    (orphan, b"<after/>".to_vec(), second_ids),
                ],
            ),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn multi_part_relationship_removal_checks_combined_part_and_archive_totals() {
        let (source_bytes, first_ids, second_ids) = archive_with_relationship_batches(1, 1);
        let source_package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let source_total = source_archive_total(&source_package);
        drop(source_package);
        let large = format!("<after>{}</after>", "x".repeat(4096)).into_bytes();
        let overlays = vec![
            (
                PackURI::new("/word/document.xml").unwrap(),
                large.clone(),
                first_ids.clone(),
            ),
            (
                PackURI::new("/custom/orphan.xml").unwrap(),
                large.clone(),
                second_ids.clone(),
            ),
        ];
        let part_limits = ReadLimits::builder()
            .max_total_part_bytes(19)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes.clone())),
            part_limits,
        )
        .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_with_external_relationship_removals_to_stream(
                &mut output,
                overlays.clone(),
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                ..
            })
        ));
        assert!(output.is_empty());
        let limits = ReadLimits::builder()
            .max_archive_total_bytes(source_total + 1)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes.clone())),
            limits,
        )
        .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_with_external_relationship_removals_to_stream(
                &mut output,
                overlays,
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveTotalBytes,
                ..
            })
        ));
        assert!(output.is_empty());

        let (source_bytes, first_ids, second_ids) = archive_with_relationship_batches(1, 1);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut unrestricted = Vec::new();
        package
            .write_part_overlays_with_external_relationship_removals_to_stream(
                &mut unrestricted,
                vec![
                    (
                        PackURI::new("/word/document.xml").unwrap(),
                        large.clone(),
                        first_ids.clone(),
                    ),
                    (
                        PackURI::new("/custom/orphan.xml").unwrap(),
                        large.clone(),
                        second_ids.clone(),
                    ),
                ],
            )
            .unwrap();
        let published =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(unrestricted))).unwrap();
        let published_total = source_archive_total(&published);
        assert!(published_total > source_total);
        drop(published);
        let exact_limits = ReadLimits::builder()
            .max_archive_total_bytes(published_total)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes.clone())),
            exact_limits,
        )
        .unwrap();
        let mut exact_output = Vec::new();
        package
            .write_part_overlays_with_external_relationship_removals_to_stream(
                &mut exact_output,
                vec![
                    (
                        PackURI::new("/word/document.xml").unwrap(),
                        large.clone(),
                        first_ids.clone(),
                    ),
                    (
                        PackURI::new("/custom/orphan.xml").unwrap(),
                        large.clone(),
                        second_ids.clone(),
                    ),
                ],
            )
            .unwrap();
        assert!(!exact_output.is_empty());

        let under_limits = ReadLimits::builder()
            .max_archive_total_bytes(published_total - 1)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            under_limits,
        )
        .unwrap();
        let mut under_output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_with_external_relationship_removals_to_stream(
                &mut under_output,
                vec![
                    (
                        PackURI::new("/word/document.xml").unwrap(),
                        large.clone(),
                        first_ids,
                    ),
                    (
                        PackURI::new("/custom/orphan.xml").unwrap(),
                        large,
                        second_ids,
                    ),
                ],
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveTotalBytes,
                ..
            })
        ));
        assert!(under_output.is_empty());
    }

    #[test]
    fn external_relationship_removal_overlay_changes_only_owner_and_rels_member() {
        let source_bytes = archive_with_document_relationships(document_relationships());
        let source_raw = raw_records(&source_bytes);
        let document = PackURI::new("/word/document.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                b"<after/>".to_vec(),
                vec!["rExternal".to_owned()],
            )
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        let main = reopened.get_part(&document).unwrap();
        assert_eq!(main.blob(), b"<after/>");
        assert!(main.rels().get("rExternal").is_none());
        assert_eq!(
            main.rels()
                .get("rInternal")
                .unwrap()
                .target_partname()
                .unwrap(),
            PackURI::new("/custom/orphan.xml").unwrap()
        );
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if matches!(
                name.as_slice(),
                b"word/document.xml" | b"word/_rels/document.xml.rels"
            ) {
                assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
            } else {
                assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
                assert_eq!(
                    central_without_local_offset(&output_raw[&name].central),
                    central_without_local_offset(&source_record.central),
                    "{name:?}"
                );
            }
        }
    }

    #[test]
    fn external_relationship_removal_overlay_refuses_ids_and_limits_before_output() {
        let source_bytes = archive_with_document_relationships(document_relationships());
        let document = PackURI::new("/word/document.xml").unwrap();
        for (ids, expected) in [
            (vec!["missing".to_owned()], "missing external relationship"),
            (
                vec!["rExternal".to_owned(), "rExternal".to_owned()],
                "duplicate relationship",
            ),
            (vec!["rInternal".to_owned()], "internal relationship"),
        ] {
            let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
                source_bytes.clone(),
            )))
            .unwrap();
            let mut output = Vec::new();
            let error = package
                .write_part_overlay_with_external_relationship_removals_to_stream(
                    &mut output,
                    &document,
                    b"<after/>".to_vec(),
                    ids,
                )
                .unwrap_err();
            match expected {
                "missing external relationship" => {
                    assert!(matches!(error, OpcError::RelationshipNotFound(_)));
                },
                "duplicate relationship" => {
                    assert!(matches!(error, OpcError::DuplicateRelationshipId(_)));
                },
                "internal relationship" => {
                    assert!(matches!(error, OpcError::InvalidRelationship(_)));
                },
                _ => unreachable!(),
            }
            assert!(output.is_empty());
        }

        let exact_bound_ids = (0..MAX_SOURCE_RELATIONSHIP_REMOVALS)
            .map(|index| format!("missing-{index}"))
            .collect();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                b"<after/>".to_vec(),
                exact_bound_ids,
            ),
            Err(OpcError::RelationshipNotFound(_))
        ));
        assert!(output.is_empty());

        let over_bound_ids = (0..=MAX_SOURCE_RELATIONSHIP_REMOVALS)
            .map(|index| format!("missing-{index}"))
            .collect();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                b"<after/>".to_vec(),
                over_bound_ids,
            ),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());

        let limits = ReadLimits::builder()
            .max_part_bytes(10)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            limits,
        )
        .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                vec![b'x'; 11],
                vec!["rExternal".to_owned()],
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 11,
                maximum: 10,
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn multi_part_overlay_checks_set_bounds_duplicates_and_aggregate_limits() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();

        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document.clone(), b"<first/>".to_vec()),
                    (document.clone(), b"<second/>".to_vec()),
                ],
            ),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(output.is_empty());

        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let oversized_set = (0..=MAX_SOURCE_OVERLAY_PARTS)
            .map(|_| (document.clone(), b"<changed/>".to_vec()))
            .collect();
        assert!(matches!(
            package.write_part_overlays_to_stream(&mut output, oversized_set),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());

        let limits = ReadLimits::builder()
            .max_part_bytes(20)
            .unwrap()
            .max_total_part_bytes(21)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            limits,
        )
        .unwrap();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document, b"<document/>".to_vec()),
                    (orphan, b"<orphan-2/>".to_vec()),
                ],
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                ..
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn multi_part_overlay_all_noop_preserves_signed_source_identity() {
        let source_bytes = signed_archive(b"<signed/>");
        let document = PackURI::new("/word/document.xml").unwrap();
        let signature = PackURI::new("/signature/origin.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document.clone(), b"<signed/>".to_vec()),
                    (signature.clone(), b"<origin/>".to_vec()),
                ],
            )
            .unwrap();
        assert_eq!(output, source_bytes);

        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            signed_archive(b"<signed/>"),
        )))
        .unwrap();
        output.clear();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document, b"<changed/>".to_vec()),
                    (signature, b"<origin/>".to_vec()),
                ],
            ),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_exact_noop_preserves_every_source_byte() {
        let source_bytes = archive_bytes(root_relationships(), b"malformed but unchanged", true);
        let source = Arc::new(CountingSource::new(source_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"malformed but unchanged".to_vec())
            .unwrap();
        assert_eq!(output, source_bytes);
    }

    #[test]
    fn unmanaged_exact_source_publication_remains_unbudgeted() {
        let source_bytes = archive_bytes(root_relationships(), b"unmanaged exact source", true);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        artifact.write_to_stream(&mut output).unwrap();

        assert_eq!(output, source_bytes);
        let diagnostics = package.cache_diagnostics();
        assert!(!diagnostics.budget_managed);
        assert_eq!(diagnostics.budget_output_bytes_used, 0);
        assert_eq!(diagnostics.budget_output_bytes_limit, None);
        assert_eq!(diagnostics.budget_reservation_failures, 0);
    }

    #[test]
    fn exact_source_publication_checks_one_fewer_version_per_copy_chunk() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"exact source publication",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let versions_before = source.versions.load(Ordering::SeqCst);
        let mut output = Vec::new();

        package
            .write_part_overlays_to_stream(&mut output, Vec::new())
            .unwrap();

        // This fixture fits in one publication chunk. The initial check,
        // read pre/post checks, sink write/flush checks, and final check are
        // the complete freshness contract; the removed post-read check would
        // make this seven instead of six observations.
        assert_eq!(source.versions.load(Ordering::SeqCst) - versions_before, 6);
    }

    #[test]
    fn one_part_overlay_refuses_unsupported_physical_layout_before_output() {
        let mut source_bytes = b"unsupported ZIP prelude".to_vec();
        source_bytes.extend_from_slice(&archive_bytes(root_relationships(), b"<before/>", true));
        let source = Arc::new(CountingSource::new(source_bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_rejects_invalid_xml_and_signed_changes_before_output() {
        let target = PackURI::new("/word/document.xml").unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<broken".to_vec()),
            Err(OpcError::XmlPublication { .. })
        ));
        assert!(output.is_empty());

        let signed_bytes = signed_archive(b"<signed/>");
        let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(signed).unwrap();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<changed/>".to_vec()),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(output.is_empty());

        let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(signed).unwrap();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<signed/>".to_vec())
            .unwrap();
        assert_eq!(output, signed_bytes);
    }

    #[test]
    fn one_part_overlay_enforces_replacement_limits_without_output() {
        let limits = ReadLimits::builder()
            .max_part_bytes(10)
            .unwrap()
            .max_total_part_bytes(19)
            .unwrap()
            .build()
            .unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at_with_limits(source, limits).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, vec![b'x'; 11]),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 11,
                maximum: 10
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_reports_source_changes_before_and_during_output() {
        let target = PackURI::new("/word/document.xml").unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        source.changed();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(output.is_empty());

        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let mut sink = MutatingSink {
            source,
            bytes: Vec::new(),
            change_after: 1,
            changed: false,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert!(written > 0);
                assert!(matches!(*source, OpcError::SourceChanged { .. }));
            },
            other => panic!("unexpected source-change error: {other:?}"),
        }
    }

    #[test]
    fn one_part_overlay_bounds_writes_and_reports_partial_sink_failure() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut sink = BoundedFailingSink {
            accepted: 0,
            limit: 100,
            largest_write: 0,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
            other => panic!("unexpected sink error: {other:?}"),
        }
        assert!(sink.largest_write <= SOURCE_PUBLICATION_CHUNK_BYTES);
    }

    #[test]
    fn managed_publication_checks_before_output_and_between_copy_chunks() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let target = PackURI::new("/word/document.xml").unwrap();

        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(source_bytes.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        cancellation_source.cancel();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec()),
            Err(OpcError::Cancelled)
        ));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);

        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(source_bytes.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes)),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut sink = CancelAfterWriteSink {
            cancellation_source,
            bytes: Vec::new(),
            cancelled: false,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert!(written > 0);
                assert!(matches!(*source, OpcError::Cancelled));
            },
            other => panic!("unexpected managed cancellation error: {other:?}"),
        }
        assert!(!sink.bytes.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::OutputBytes), sink.bytes.len() as u64);
    }

    #[test]
    fn managed_exact_noop_overlay_charges_physical_output_bytes() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let target = PackURI::new("/word/document.xml").unwrap();
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec())
            .unwrap();

        assert_eq!(output, source_bytes);
        assert_eq!(budget.used(Resource::OutputBytes), output.len() as u64);
    }

    #[test]
    fn managed_exact_output_limit_accepts_exact_small_artifact() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let (budget, _cancellation_source, context) =
            managed_context_with_output(source_bytes.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        artifact.write_to_stream(&mut output).unwrap();

        assert_eq!(output, source_bytes);
        assert_eq!(
            budget.used(Resource::OutputBytes),
            source_bytes.len() as u64
        );
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.budget_reservation_failures, 0);
        assert_eq!(
            diagnostics.budget_output_bytes_used,
            source_bytes.len() as u64
        );
        assert_eq!(
            diagnostics.budget_output_bytes_limit,
            Some(source_bytes.len() as u64)
        );
    }

    #[test]
    fn managed_exact_output_limit_one_under_refuses_before_small_artifact_output() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let (budget, _cancellation_source, context) =
            managed_context_with_output((source_bytes.len() as u64).saturating_sub(1));
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes)),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        let error = artifact.write_to_stream(&mut output).unwrap_err();

        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::OutputBytes
        ));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::OutputBytes), 0);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.budget_reservation_failures, 1);
        assert_eq!(diagnostics.budget_output_bytes_used, 0);
    }

    #[test]
    fn managed_large_exact_publication_refuses_mid_stream_with_exact_partial_charge() {
        let source_bytes = large_archive_bytes(root_relationships(), b"<before/>", 150_000);
        assert!(source_bytes.len() > SOURCE_PUBLICATION_CHUNK_BYTES * 2);
        let output_limit = SOURCE_PUBLICATION_CHUNK_BYTES as u64 + 100;
        let (budget, _cancellation_source, context) = managed_context_with_output(output_limit);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes)),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        let error = artifact.write_to_stream(&mut output).unwrap_err();

        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert_eq!(written, SOURCE_PUBLICATION_CHUNK_BYTES as u64);
                match *source {
                    OpcError::Execution(ExecutionError::ResourceLimit(limit)) => {
                        assert_eq!(limit.resource, Resource::OutputBytes);
                    },
                    other => panic!("unexpected mid-stream source error: {other:?}"),
                }
            },
            other => panic!("unexpected mid-stream output error: {other:?}"),
        }
        assert_eq!(output.len(), SOURCE_PUBLICATION_CHUNK_BYTES);
        assert_eq!(
            budget.used(Resource::OutputBytes),
            SOURCE_PUBLICATION_CHUNK_BYTES as u64
        );
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
    }

    #[test]
    fn managed_output_charge_reaches_hierarchical_parent_and_child_budgets() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let output_bytes = source_bytes.len() as u64;
        let root = Budget::root(
            "opc-source-output-root",
            Limits::new(4096, u64::MAX, output_bytes, u64::MAX, u64::MAX, u64::MAX),
        );
        let child = root.child(
            "opc-source-output-child",
            Limits::new(4096, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(4096).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(child.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        artifact.write_to_stream(&mut output).unwrap();

        assert_eq!(output, source_bytes);
        assert_eq!(root.used(Resource::OutputBytes), output_bytes);
        assert_eq!(child.used(Resource::OutputBytes), output_bytes);
        drop(cancellation_source);
    }

    #[test]
    fn managed_cumulative_output_retry_exhaustion_refuses_second_exact_publication() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let output_bytes = source_bytes.len() as u64;
        let (budget, _cancellation_source, context) = managed_context_with_output(output_bytes);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut first = Vec::new();
        artifact.write_to_stream(&mut first).unwrap();
        let mut retry = Vec::new();
        let error = artifact.write_to_stream(&mut retry).unwrap_err();

        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::OutputBytes
        ));
        assert_eq!(first, source_bytes);
        assert!(retry.is_empty());
        assert_eq!(budget.used(Resource::OutputBytes), output_bytes);
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
    }

    #[test]
    fn managed_flush_failure_reports_exact_output_charge_as_incomplete() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut sink = FlushFailingSink { bytes: Vec::new() };
        let error = artifact.write_to_stream(&mut sink).unwrap_err();

        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert_eq!(written, source_bytes.len() as u64);
                assert!(matches!(
                    *source,
                    OpcError::IoError(error)
                        if error.kind() == std::io::ErrorKind::Other
                ));
            },
            other => panic!("unexpected flush error: {other:?}"),
        }
        assert_eq!(sink.bytes, source_bytes);
        assert_eq!(
            budget.used(Resource::OutputBytes),
            source_bytes.len() as u64
        );
    }

    #[test]
    fn managed_changed_overlay_charges_exact_accepted_output_bytes() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let target = PackURI::new("/word/document.xml").unwrap();
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes)),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(reopened.get_part(&target).unwrap().blob(), b"<after/>");
        assert_eq!(budget.used(Resource::OutputBytes), output.len() as u64);
    }

    #[test]
    fn managed_exact_noop_refuses_the_first_output_write_at_zero_limit() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let target = PackURI::new("/word/document.xml").unwrap();
        let (budget, _cancellation_source, context) = managed_context_with_output(0);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut output = Vec::new();
        let error = package
            .write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec())
            .unwrap_err();

        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::OutputBytes
        ));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::OutputBytes), 0);
    }

    #[test]
    fn managed_partial_sink_failure_commits_only_accepted_output_bytes() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let target = PackURI::new("/word/document.xml").unwrap();
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut sink = BoundedFailingSink {
            accepted: 0,
            limit: 100,
            largest_write: 0,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();

        match error {
            OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
            other => panic!("unexpected sink error: {other:?}"),
        }
        assert_eq!(budget.used(Resource::OutputBytes), 100);
    }

    #[test]
    fn managed_source_change_after_accepted_output_preserves_partial_accounting() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let target = PackURI::new("/word/document.xml").unwrap();
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut sink = MutatingSink {
            source,
            bytes: Vec::new(),
            change_after: 0,
            changed: false,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();

        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert_eq!(written, sink.bytes.len() as u64);
                assert!(matches!(*source, OpcError::SourceChanged { .. }));
            },
            other => panic!("unexpected source-change error: {other:?}"),
        }
        assert_eq!(budget.used(Resource::OutputBytes), sink.bytes.len() as u64);
    }

    #[test]
    fn managed_short_interrupted_source_copy_charges_only_accepted_bytes() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut sink = ShortInterruptedSink {
            bytes: Vec::new(),
            maximum_write: 7,
            interrupted: false,
        };
        artifact.write_to_stream(&mut sink).unwrap();

        assert_eq!(sink.bytes, source_bytes);
        assert_eq!(
            budget.used(Resource::OutputBytes),
            source_bytes.len() as u64
        );
        let diagnostics = package.cache_diagnostics();
        assert_eq!(
            diagnostics.budget_output_bytes_used,
            source_bytes.len() as u64
        );
        assert_eq!(diagnostics.budget_output_bytes_limit, Some(u64::MAX));
    }

    #[test]
    fn managed_overreporting_sink_does_not_commit_output_budget() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut sink = OverReportingSink {
            calls: 0,
            accepted: 0,
        };

        assert!(matches!(
            artifact.write_to_stream(&mut sink),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));
        assert_eq!(sink.calls, 1);
        assert_eq!(sink.accepted, 0);
        assert_eq!(budget.used(Resource::OutputBytes), 0);
    }
}
