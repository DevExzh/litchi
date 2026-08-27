//! Operation-scoped resource limits for XLSB external links.
//!
//! [`ExternalLinkLimits`] describes one governed parse, eager-construction, or
//! write API operation that accepts or stores the policy.  It is not an RSS
//! limit, a process-wide limit, a global singleton, or a limit shared by
//! concurrent operations.  A caller creates a fresh operation budget for each
//! operation and charges a resource before retaining the corresponding bytes,
//! records, or semantic objects.  Semantic clones requested later by the
//! caller are outside that operation budget.

use super::{
    Error, MAX_COLLECTION_ITEMS, MAX_LINK_PART_BYTES, MAX_UNKNOWN_BYTES, MAX_UNKNOWN_RECORDS,
    MAX_WIDE_STRING_UNITS, MAX_XLSB_EXTERNAL_CACHED_VALUES, Result,
};
use std::fmt;
use thiserror::Error as ThisError;

/// A resource governed by an [`ExternalLinkLimits`] policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ExternalLinkResource {
    /// Bytes retained for one external-link part.
    PartBytes,
    /// Bytes retained across all external-link parts in the operation.
    TotalPartBytes,
    /// Opaque bytes retained for one external-link part.
    OpaqueBytes,
    /// Opaque bytes retained across the operation.
    TotalOpaqueBytes,
    /// Records retained across the operation.
    Records,
    /// `BrtExtern` workbook-cache records retained across the operation.
    ///
    /// DDE/OLE `SUP_NAME_VALUE` records are ordinary external-link records,
    /// not workbook-cache records for this counter.
    CacheRecords,
    /// Opaque records retained across the operation.
    OpaqueRecords,
    /// External links retained across the operation.
    Links,
    /// Semantic entries retained across the operation, including workbook
    /// defined names and DDE/OLE items.
    Items,
    /// Cache matrices retained across the operation.
    Matrices,
    /// Cache cells retained across the operation.
    Cells,
    /// UTF-16 units in one string.
    Utf16Units,
    /// UTF-16 units retained across the operation.
    TotalUtf16Units,
    /// Decoded semantic bytes retained across the operation.
    DecodedSemanticBytes,
    /// Semantic and provenance objects retained across the operation.
    RetainedObjects,
}

impl fmt::Display for ExternalLinkResource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PartBytes => "external-link part bytes",
            Self::TotalPartBytes => "total external-link part bytes",
            Self::OpaqueBytes => "external-link opaque bytes",
            Self::TotalOpaqueBytes => "total external-link opaque bytes",
            Self::Records => "external-link records",
            Self::CacheRecords => "external-link workbook-cache records",
            Self::OpaqueRecords => "external-link opaque records",
            Self::Links => "external links",
            Self::Items => "external-link items",
            Self::Matrices => "external-link matrices",
            Self::Cells => "external-link cells",
            Self::Utf16Units => "external-link string UTF-16 units",
            Self::TotalUtf16Units => "total external-link UTF-16 units",
            Self::DecodedSemanticBytes => "decoded external-link semantic bytes",
            Self::RetainedObjects => "retained external-link objects",
        })
    }
}

/// A caller-supplied external-link limit that cannot be represented by the
/// selected policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum ExternalLinkLimitsError {
    /// A per-part ceiling exceeds the corresponding physical wire ceiling.
    #[error("invalid {resource} limit {value}; maximum is {maximum}")]
    HardMaximum {
        /// Resource whose requested ceiling was rejected.
        resource: ExternalLinkResource,
        /// Requested ceiling.
        value: usize,
        /// Physical maximum supported by this policy.
        maximum: usize,
    },
    /// A per-part ceiling exceeds its aggregate counterpart.
    #[error("{resource} per-part limit {value} exceeds its aggregate limit {maximum}")]
    PerPartExceedsAggregate {
        /// Per-part resource whose requested ceiling was rejected.
        resource: ExternalLinkResource,
        /// Requested per-part ceiling.
        value: usize,
        /// Configured aggregate ceiling.
        maximum: usize,
    },
}

impl ExternalLinkLimitsError {
    /// Resource whose limit was rejected.
    #[must_use]
    pub const fn resource(self) -> ExternalLinkResource {
        match self {
            Self::HardMaximum { resource, .. } | Self::PerPartExceedsAggregate { resource, .. } => {
                resource
            },
        }
    }

    /// Requested limit value.
    #[must_use]
    pub const fn value(self) -> usize {
        match self {
            Self::HardMaximum { value, .. } | Self::PerPartExceedsAggregate { value, .. } => value,
        }
    }

    /// Maximum against which the request was checked.
    #[must_use]
    pub const fn maximum(self) -> usize {
        match self {
            Self::HardMaximum { maximum, .. } | Self::PerPartExceedsAggregate { maximum, .. } => {
                maximum
            },
        }
    }
}

// The operands are fixed protocol ceilings and this expression is evaluated
// at compile time.  Their sum is safely representable on every supported
// target (which already needs to represent `MAX_XLSB_EXTERNAL_CACHED_VALUES`).
const DEFAULT_RETAINED_OBJECTS: usize = MAX_COLLECTION_ITEMS * 6 + MAX_XLSB_EXTERNAL_CACHED_VALUES;

/// Finite resource ceilings for one explicit external-link parse,
/// eager-construction, or write operation that accepts or stores this policy.
///
/// The policy is copied into one mutable operation budget per governed
/// operation.
/// It does not observe or cap process RSS, allocations made by other
/// operations, concurrent operations, or caller-requested semantic clones
/// made after the governed operation completes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes the resource-policy fields explicit."
)]
pub struct ExternalLinkLimits {
    max_part_bytes: usize,
    max_total_part_bytes: usize,
    max_opaque_bytes: usize,
    max_total_opaque_bytes: usize,
    max_utf16_units: usize,
    max_total_utf16_units: usize,
    max_records: usize,
    max_cache_records: usize,
    max_opaque_records: usize,
    max_links: usize,
    max_items: usize,
    max_matrices: usize,
    max_cells: usize,
    max_decoded_semantic_bytes: usize,
    max_retained_objects: usize,
}

impl ExternalLinkLimits {
    /// The largest allowed per-part external-link payload.
    pub const MAX_PART_BYTES: usize = MAX_LINK_PART_BYTES;
    /// The largest allowed per-part opaque payload.
    pub const MAX_OPAQUE_BYTES: usize = MAX_UNKNOWN_BYTES;
    /// The largest allowed one-string UTF-16 length.
    pub const MAX_UTF16_UNITS: usize = MAX_WIDE_STRING_UNITS;

    /// The standard finite profile for one external-link operation.
    pub const DEFAULT: Self = Self {
        max_part_bytes: MAX_LINK_PART_BYTES,
        max_total_part_bytes: MAX_LINK_PART_BYTES,
        max_opaque_bytes: MAX_UNKNOWN_BYTES,
        max_total_opaque_bytes: MAX_UNKNOWN_BYTES,
        max_utf16_units: MAX_WIDE_STRING_UNITS,
        max_total_utf16_units: MAX_LINK_PART_BYTES / 2,
        max_records: 1_048_576,
        max_cache_records: 1_048_576,
        max_opaque_records: MAX_UNKNOWN_RECORDS,
        max_links: MAX_COLLECTION_ITEMS,
        max_items: MAX_COLLECTION_ITEMS,
        max_matrices: MAX_COLLECTION_ITEMS,
        max_cells: 1_048_576,
        max_decoded_semantic_bytes: 3 * (MAX_LINK_PART_BYTES / 2),
        max_retained_objects: DEFAULT_RETAINED_OBJECTS,
    };

    /// Start building a policy from [`Self::DEFAULT`].
    #[must_use]
    pub const fn builder() -> ExternalLinkLimitsBuilder {
        ExternalLinkLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum bytes retained for one external-link part.
    #[must_use]
    pub const fn max_part_bytes(self) -> usize {
        self.max_part_bytes
    }

    /// Maximum bytes retained across external-link parts.
    #[must_use]
    pub const fn max_total_part_bytes(self) -> usize {
        self.max_total_part_bytes
    }

    /// Maximum opaque bytes retained for one external-link part.
    #[must_use]
    pub const fn max_opaque_bytes(self) -> usize {
        self.max_opaque_bytes
    }

    /// Maximum opaque bytes retained across the operation.
    #[must_use]
    pub const fn max_total_opaque_bytes(self) -> usize {
        self.max_total_opaque_bytes
    }

    /// Maximum UTF-16 units in one string.
    #[must_use]
    pub const fn max_utf16_units(self) -> usize {
        self.max_utf16_units
    }

    /// Maximum UTF-16 units retained across the operation.
    #[must_use]
    pub const fn max_total_utf16_units(self) -> usize {
        self.max_total_utf16_units
    }

    /// Maximum records retained across the operation.
    #[must_use]
    pub const fn max_records(self) -> usize {
        self.max_records
    }

    /// Maximum `BrtExtern` workbook-cache records retained across the
    /// operation. DDE/OLE `SUP_NAME_VALUE` records are excluded.
    #[must_use]
    pub const fn max_cache_records(self) -> usize {
        self.max_cache_records
    }

    /// Maximum opaque records retained across the operation.
    #[must_use]
    pub const fn max_opaque_records(self) -> usize {
        self.max_opaque_records
    }

    /// Maximum external links retained across the operation.
    #[must_use]
    pub const fn max_links(self) -> usize {
        self.max_links
    }

    /// Maximum semantic entries retained across the operation, including
    /// workbook defined names and DDE/OLE items.
    #[must_use]
    pub const fn max_items(self) -> usize {
        self.max_items
    }

    /// Maximum cache matrices retained across the operation.
    #[must_use]
    pub const fn max_matrices(self) -> usize {
        self.max_matrices
    }

    /// Maximum cache cells retained across the operation.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }

    /// Maximum decoded semantic bytes retained across the operation.
    #[must_use]
    pub const fn max_decoded_semantic_bytes(self) -> usize {
        self.max_decoded_semantic_bytes
    }

    /// Maximum semantic and provenance objects retained across the operation.
    #[must_use]
    pub const fn max_retained_objects(self) -> usize {
        self.max_retained_objects
    }

    /// Alias for [`Self::max_records`].
    #[must_use]
    pub const fn max_total_records(self) -> usize {
        self.max_records()
    }

    /// Alias for [`Self::max_cache_records`].
    #[must_use]
    pub const fn max_total_cache_records(self) -> usize {
        self.max_cache_records()
    }

    /// Alias for [`Self::max_opaque_records`].
    #[must_use]
    pub const fn max_total_opaque_records(self) -> usize {
        self.max_opaque_records()
    }

    /// Alias for [`Self::max_links`].
    #[must_use]
    pub const fn max_total_links(self) -> usize {
        self.max_links()
    }

    /// Alias for [`Self::max_items`].
    #[must_use]
    pub const fn max_total_items(self) -> usize {
        self.max_items()
    }

    /// Alias for [`Self::max_matrices`].
    #[must_use]
    pub const fn max_total_matrices(self) -> usize {
        self.max_matrices()
    }

    /// Alias for [`Self::max_cells`].
    #[must_use]
    pub const fn max_total_cells(self) -> usize {
        self.max_cells()
    }

    /// Alias for [`Self::max_decoded_semantic_bytes`].
    #[must_use]
    pub const fn max_total_decoded_semantic_bytes(self) -> usize {
        self.max_decoded_semantic_bytes()
    }

    /// Alias for [`Self::max_retained_objects`].
    #[must_use]
    pub const fn max_total_retained_objects(self) -> usize {
        self.max_retained_objects()
    }

    /// Create a fresh mutable operation budget.
    #[must_use]
    pub(crate) const fn budget(self) -> Budget {
        Budget::new(self)
    }

    fn validate(self) -> std::result::Result<Self, ExternalLinkLimitsError> {
        validate_per_part(
            ExternalLinkResource::PartBytes,
            self.max_part_bytes,
            MAX_LINK_PART_BYTES,
        )?;
        validate_per_part(
            ExternalLinkResource::OpaqueBytes,
            self.max_opaque_bytes,
            MAX_UNKNOWN_BYTES,
        )?;
        validate_per_part(
            ExternalLinkResource::Utf16Units,
            self.max_utf16_units,
            MAX_WIDE_STRING_UNITS,
        )?;
        if self.max_part_bytes > self.max_total_part_bytes {
            return Err(ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::PartBytes,
                value: self.max_part_bytes,
                maximum: self.max_total_part_bytes,
            });
        }
        if self.max_opaque_bytes > self.max_total_opaque_bytes {
            return Err(ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::OpaqueBytes,
                value: self.max_opaque_bytes,
                maximum: self.max_total_opaque_bytes,
            });
        }
        if self.max_utf16_units > self.max_total_utf16_units {
            return Err(ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::Utf16Units,
                value: self.max_utf16_units,
                maximum: self.max_total_utf16_units,
            });
        }
        Ok(self)
    }
}

impl Default for ExternalLinkLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Builder for [`ExternalLinkLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalLinkLimitsBuilder {
    limits: ExternalLinkLimits,
}

impl ExternalLinkLimitsBuilder {
    /// Finalize the checked policy.
    ///
    /// # Errors
    ///
    /// Returns an error if a per-part ceiling exceeds its wire hard maximum
    /// or its aggregate counterpart.
    pub fn build(self) -> std::result::Result<ExternalLinkLimits, ExternalLinkLimitsError> {
        self.limits.validate()
    }

    /// Set the per-part byte ceiling.
    ///
    /// The hard maximum [`ExternalLinkLimits::MAX_PART_BYTES`] is checked
    /// when [`Self::build`] is called.
    pub fn max_part_bytes(mut self, value: usize) -> Self {
        self.limits.max_part_bytes = value;
        self
    }

    /// Set the aggregate part-byte ceiling.
    pub fn max_total_part_bytes(mut self, value: usize) -> Self {
        self.limits.max_total_part_bytes = value;
        self
    }

    /// Set the per-part opaque-byte ceiling.
    ///
    /// The hard maximum [`ExternalLinkLimits::MAX_OPAQUE_BYTES`] is checked
    /// when [`Self::build`] is called.
    pub fn max_opaque_bytes(mut self, value: usize) -> Self {
        self.limits.max_opaque_bytes = value;
        self
    }

    /// Set the aggregate opaque-byte ceiling.
    pub fn max_total_opaque_bytes(mut self, value: usize) -> Self {
        self.limits.max_total_opaque_bytes = value;
        self
    }

    /// Set the maximum UTF-16 length of one string.
    ///
    /// The hard maximum [`ExternalLinkLimits::MAX_UTF16_UNITS`] is checked
    /// when [`Self::build`] is called.
    pub fn max_utf16_units(mut self, value: usize) -> Self {
        self.limits.max_utf16_units = value;
        self
    }

    /// Set the aggregate UTF-16-unit ceiling.
    pub fn max_total_utf16_units(mut self, value: usize) -> Self {
        self.limits.max_total_utf16_units = value;
        self
    }

    /// Set the aggregate record ceiling.
    pub fn max_records(mut self, value: usize) -> Self {
        self.limits.max_records = value;
        self
    }

    /// Set the aggregate cache-record ceiling.
    pub fn max_cache_records(mut self, value: usize) -> Self {
        self.limits.max_cache_records = value;
        self
    }

    /// Set the aggregate opaque-record ceiling.
    pub fn max_opaque_records(mut self, value: usize) -> Self {
        self.limits.max_opaque_records = value;
        self
    }

    /// Set the aggregate external-link ceiling.
    pub fn max_links(mut self, value: usize) -> Self {
        self.limits.max_links = value;
        self
    }

    /// Set the aggregate semantic-entry ceiling, including workbook defined
    /// names and DDE/OLE items.
    pub fn max_items(mut self, value: usize) -> Self {
        self.limits.max_items = value;
        self
    }

    /// Set the aggregate cache-matrix ceiling.
    pub fn max_matrices(mut self, value: usize) -> Self {
        self.limits.max_matrices = value;
        self
    }

    /// Set the aggregate cache-cell ceiling.
    pub fn max_cells(mut self, value: usize) -> Self {
        self.limits.max_cells = value;
        self
    }

    /// Set the aggregate decoded semantic-byte ceiling.
    pub fn max_decoded_semantic_bytes(mut self, value: usize) -> Self {
        self.limits.max_decoded_semantic_bytes = value;
        self
    }

    /// Set the aggregate retained-object ceiling.
    pub fn max_retained_objects(mut self, value: usize) -> Self {
        self.limits.max_retained_objects = value;
        self
    }

    /// Alias for [`Self::max_records`].
    pub fn max_total_records(self, value: usize) -> Self {
        self.max_records(value)
    }

    /// Alias for [`Self::max_cache_records`].
    pub fn max_total_cache_records(self, value: usize) -> Self {
        self.max_cache_records(value)
    }

    /// Alias for [`Self::max_opaque_records`].
    pub fn max_total_opaque_records(self, value: usize) -> Self {
        self.max_opaque_records(value)
    }

    /// Alias for [`Self::max_links`].
    pub fn max_total_links(self, value: usize) -> Self {
        self.max_links(value)
    }

    /// Alias for [`Self::max_items`], covering defined names and DDE/OLE items.
    pub fn max_total_items(self, value: usize) -> Self {
        self.max_items(value)
    }

    /// Alias for [`Self::max_matrices`].
    pub fn max_total_matrices(self, value: usize) -> Self {
        self.max_matrices(value)
    }

    /// Alias for [`Self::max_cells`].
    pub fn max_total_cells(self, value: usize) -> Self {
        self.max_cells(value)
    }

    /// Alias for [`Self::max_decoded_semantic_bytes`].
    pub fn max_total_decoded_semantic_bytes(self, value: usize) -> Self {
        self.max_decoded_semantic_bytes(value)
    }

    /// Alias for [`Self::max_retained_objects`].
    pub fn max_total_retained_objects(self, value: usize) -> Self {
        self.max_retained_objects(value)
    }
}

fn validate_per_part(
    resource: ExternalLinkResource,
    value: usize,
    maximum: usize,
) -> std::result::Result<(), ExternalLinkLimitsError> {
    if value > maximum {
        return Err(ExternalLinkLimitsError::HardMaximum {
            resource,
            value,
            maximum,
        });
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, Default)]
struct Usage {
    total_part_bytes: usize,
    total_opaque_bytes: usize,
    local_opaque_bytes: usize,
    records: usize,
    cache_records: usize,
    opaque_records: usize,
    links: usize,
    items: usize,
    matrices: usize,
    cells: usize,
    total_utf16_units: usize,
    decoded_semantic_bytes: usize,
    retained_objects: usize,
}

/// Mutable counters for one governed external-link operation.
///
/// The budget has no process-global state and is not safe or intended to be
/// shared between concurrent operations.  Callers should charge it before an
/// allocation or before publishing a retained semantic value.  Semantic
/// clones requested by callers after the operation are outside this budget.
#[derive(Debug)]
pub(crate) struct Budget {
    limits: ExternalLinkLimits,
    usage: Usage,
}

impl Budget {
    /// Create an empty budget for one operation.
    pub(crate) const fn new(limits: ExternalLinkLimits) -> Self {
        Self {
            limits,
            usage: Usage {
                total_part_bytes: 0,
                total_opaque_bytes: 0,
                local_opaque_bytes: 0,
                records: 0,
                cache_records: 0,
                opaque_records: 0,
                links: 0,
                items: 0,
                matrices: 0,
                cells: 0,
                total_utf16_units: 0,
                decoded_semantic_bytes: 0,
                retained_objects: 0,
            },
        }
    }

    /// Check a prospective number of links without consuming link budget.
    ///
    /// This is intended for callers that must reserve an external-link
    /// vector before parsing or publishing each link.  The actual link still
    /// has to be charged by [`Self::begin_link_part`].
    pub(crate) fn preflight_links(&self, count: usize) -> Result<()> {
        let _ = checked_add(
            self.usage.links,
            count,
            ExternalLinkResource::Links,
            self.limits.max_links,
        )?;
        Ok(())
    }

    /// Validate one deferred part against the current operation without
    /// consuming it. The parser still charges the actual decoded length.
    pub(crate) fn preflight_link_part(&self, data_len: usize) -> Result<()> {
        let _ = checked_add(
            0,
            data_len,
            ExternalLinkResource::PartBytes,
            self.limits.max_part_bytes,
        )?;
        let _ = checked_add(
            self.usage.total_part_bytes,
            data_len,
            ExternalLinkResource::TotalPartBytes,
            self.limits.max_total_part_bytes,
        )?;
        let _ = checked_add(
            self.usage.links,
            1,
            ExternalLinkResource::Links,
            self.limits.max_links,
        )?;
        Ok(())
    }

    /// Charge one link part before retaining its bytes.
    pub(crate) fn begin_link_part(&mut self, data_len: usize) -> Result<()> {
        let local_part_bytes = checked_add(
            0,
            data_len,
            ExternalLinkResource::PartBytes,
            self.limits.max_part_bytes,
        )?;
        let total_part_bytes = checked_add(
            self.usage.total_part_bytes,
            data_len,
            ExternalLinkResource::TotalPartBytes,
            self.limits.max_total_part_bytes,
        )?;
        let links = checked_add(
            self.usage.links,
            1,
            ExternalLinkResource::Links,
            self.limits.max_links,
        )?;
        let _ = local_part_bytes;
        self.usage.total_part_bytes = total_part_bytes;
        self.usage.links = links;
        self.usage.local_opaque_bytes = 0;
        Ok(())
    }

    /// Charge one parsed or emitted record.
    pub(crate) fn record(&mut self, cache_region: bool) -> Result<()> {
        let records = checked_add(
            self.usage.records,
            1,
            ExternalLinkResource::Records,
            self.limits.max_records,
        )?;
        let cache_records = if cache_region {
            Some(checked_add(
                self.usage.cache_records,
                1,
                ExternalLinkResource::CacheRecords,
                self.limits.max_cache_records,
            )?)
        } else {
            None
        };
        self.usage.records = records;
        if let Some(cache_records) = cache_records {
            self.usage.cache_records = cache_records;
        }
        Ok(())
    }

    /// Charge opaque records and bytes from the current local region.
    pub(crate) fn opaque(&mut self, count: usize, bytes: usize) -> Result<()> {
        let local_opaque_bytes = checked_add(
            self.usage.local_opaque_bytes,
            bytes,
            ExternalLinkResource::OpaqueBytes,
            self.limits.max_opaque_bytes,
        )?;
        let total_opaque_bytes = checked_add(
            self.usage.total_opaque_bytes,
            bytes,
            ExternalLinkResource::TotalOpaqueBytes,
            self.limits.max_total_opaque_bytes,
        )?;
        let opaque_records = checked_add(
            self.usage.opaque_records,
            count,
            ExternalLinkResource::OpaqueRecords,
            self.limits.max_opaque_records,
        )?;
        self.usage.local_opaque_bytes = local_opaque_bytes;
        self.usage.total_opaque_bytes = total_opaque_bytes;
        self.usage.opaque_records = opaque_records;
        Ok(())
    }

    /// Charge retained semantic entries, including workbook defined names
    /// and DDE/OLE items.
    pub(crate) fn items(&mut self, count: usize) -> Result<()> {
        self.usage.items = checked_add(
            self.usage.items,
            count,
            ExternalLinkResource::Items,
            self.limits.max_items,
        )?;
        Ok(())
    }

    /// Charge retained cache matrices.
    pub(crate) fn matrix(&mut self, count: usize) -> Result<()> {
        self.usage.matrices = checked_add(
            self.usage.matrices,
            count,
            ExternalLinkResource::Matrices,
            self.limits.max_matrices,
        )?;
        Ok(())
    }

    /// Charge retained cache cells.
    pub(crate) fn cells(&mut self, count: usize) -> Result<()> {
        self.usage.cells = checked_add(
            self.usage.cells,
            count,
            ExternalLinkResource::Cells,
            self.limits.max_cells,
        )?;
        Ok(())
    }

    /// Charge one string's UTF-16 units and exact UTF-8 semantic bytes.
    pub(crate) fn string(&mut self, units: usize, exact_utf8_bytes: usize) -> Result<()> {
        let string_units = checked_add(
            0,
            units,
            ExternalLinkResource::Utf16Units,
            self.limits.max_utf16_units,
        )?;
        let total_utf16_units = checked_add(
            self.usage.total_utf16_units,
            units,
            ExternalLinkResource::TotalUtf16Units,
            self.limits.max_total_utf16_units,
        )?;
        let decoded_semantic_bytes = checked_add(
            self.usage.decoded_semantic_bytes,
            exact_utf8_bytes,
            ExternalLinkResource::DecodedSemanticBytes,
            self.limits.max_decoded_semantic_bytes,
        )?;
        let _ = string_units;
        self.usage.total_utf16_units = total_utf16_units;
        self.usage.decoded_semantic_bytes = decoded_semantic_bytes;
        Ok(())
    }

    /// Charge decoded formula-token or other semantic bytes.
    pub(crate) fn token_bytes(&mut self, bytes: usize) -> Result<()> {
        self.usage.decoded_semantic_bytes = checked_add(
            self.usage.decoded_semantic_bytes,
            bytes,
            ExternalLinkResource::DecodedSemanticBytes,
            self.limits.max_decoded_semantic_bytes,
        )?;
        Ok(())
    }

    /// Charge semantic and provenance objects before retaining them.
    pub(crate) fn retained_objects(&mut self, count: usize) -> Result<()> {
        self.usage.retained_objects = checked_add(
            self.usage.retained_objects,
            count,
            ExternalLinkResource::RetainedObjects,
            self.limits.max_retained_objects,
        )?;
        Ok(())
    }
}

fn checked_add(
    current: usize,
    amount: usize,
    resource: ExternalLinkResource,
    maximum: usize,
) -> Result<usize> {
    let Some(actual) = current.checked_add(amount) else {
        return Err(Error::LimitExceeded {
            resource,
            actual: usize::MAX,
            maximum,
        });
    };
    if actual > maximum {
        return Err(Error::LimitExceeded {
            resource,
            actual,
            maximum,
        });
    }
    Ok(actual)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_limits_are_valid() {
        let limits = ExternalLinkLimits::builder().build().unwrap();
        assert_eq!(limits, ExternalLinkLimits::DEFAULT);
        assert_eq!(limits.max_part_bytes(), MAX_LINK_PART_BYTES);
        assert_eq!(limits.max_total_part_bytes(), MAX_LINK_PART_BYTES);
        assert_eq!(limits.max_opaque_bytes(), MAX_UNKNOWN_BYTES);
        assert_eq!(limits.max_total_opaque_bytes(), MAX_UNKNOWN_BYTES);
        assert_eq!(limits.max_utf16_units(), MAX_WIDE_STRING_UNITS);
        assert_eq!(limits.max_total_utf16_units(), MAX_LINK_PART_BYTES / 2);
        assert_eq!(limits.max_records(), 1_048_576);
        assert_eq!(limits.max_cache_records(), 1_048_576);
        assert_eq!(limits.max_opaque_records(), MAX_UNKNOWN_RECORDS);
        assert_eq!(limits.max_links(), MAX_COLLECTION_ITEMS);
        assert_eq!(limits.max_items(), MAX_COLLECTION_ITEMS);
        assert_eq!(limits.max_matrices(), MAX_COLLECTION_ITEMS);
        assert_eq!(limits.max_cells(), 1_048_576);
        assert_eq!(
            limits.max_decoded_semantic_bytes(),
            3 * (MAX_LINK_PART_BYTES / 2)
        );
        assert_eq!(
            limits.max_retained_objects(),
            MAX_COLLECTION_ITEMS * 6 + MAX_XLSB_EXTERNAL_CACHED_VALUES
        );
        assert_eq!(limits.max_total_records(), limits.max_records());
        assert_eq!(limits.max_total_cache_records(), limits.max_cache_records());
        assert_eq!(
            limits.max_total_opaque_records(),
            limits.max_opaque_records()
        );
        assert_eq!(limits.max_total_links(), limits.max_links());
        assert_eq!(limits.max_total_items(), limits.max_items());
        assert_eq!(limits.max_total_matrices(), limits.max_matrices());
        assert_eq!(limits.max_total_cells(), limits.max_cells());
        assert_eq!(
            limits.max_total_decoded_semantic_bytes(),
            limits.max_decoded_semantic_bytes()
        );
        assert_eq!(
            limits.max_total_retained_objects(),
            limits.max_retained_objects()
        );
    }

    #[test]
    fn per_part_hard_maxima_are_refused() {
        let error = ExternalLinkLimits::builder()
            .max_part_bytes(MAX_LINK_PART_BYTES + 1)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::HardMaximum {
                resource: ExternalLinkResource::PartBytes,
                value: MAX_LINK_PART_BYTES + 1,
                maximum: MAX_LINK_PART_BYTES,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::PartBytes);
        assert_eq!(error.value(), MAX_LINK_PART_BYTES + 1);
        assert_eq!(error.maximum(), MAX_LINK_PART_BYTES);

        let error = ExternalLinkLimits::builder()
            .max_opaque_bytes(MAX_UNKNOWN_BYTES + 1)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::HardMaximum {
                resource: ExternalLinkResource::OpaqueBytes,
                value: MAX_UNKNOWN_BYTES + 1,
                maximum: MAX_UNKNOWN_BYTES,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::OpaqueBytes);
        assert_eq!(error.value(), MAX_UNKNOWN_BYTES + 1);
        assert_eq!(error.maximum(), MAX_UNKNOWN_BYTES);

        let error = ExternalLinkLimits::builder()
            .max_utf16_units(MAX_WIDE_STRING_UNITS + 1)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::HardMaximum {
                resource: ExternalLinkResource::Utf16Units,
                value: MAX_WIDE_STRING_UNITS + 1,
                maximum: MAX_WIDE_STRING_UNITS,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::Utf16Units);
        assert_eq!(error.value(), MAX_WIDE_STRING_UNITS + 1);
        assert_eq!(error.maximum(), MAX_WIDE_STRING_UNITS);
    }

    #[test]
    fn per_part_limits_cannot_exceed_aggregate_limits() {
        let error = ExternalLinkLimits::builder()
            .max_total_part_bytes(0)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::PartBytes,
                value: MAX_LINK_PART_BYTES,
                maximum: 0,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::PartBytes);
        assert_eq!(error.value(), MAX_LINK_PART_BYTES);
        assert_eq!(error.maximum(), 0);

        let error = ExternalLinkLimits::builder()
            .max_total_opaque_bytes(0)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::OpaqueBytes,
                value: MAX_UNKNOWN_BYTES,
                maximum: 0,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::OpaqueBytes);
        assert_eq!(error.value(), MAX_UNKNOWN_BYTES);
        assert_eq!(error.maximum(), 0);

        let error = ExternalLinkLimits::builder()
            .max_total_utf16_units(0)
            .build()
            .unwrap_err();
        assert_eq!(
            error,
            ExternalLinkLimitsError::PerPartExceedsAggregate {
                resource: ExternalLinkResource::Utf16Units,
                value: MAX_WIDE_STRING_UNITS,
                maximum: 0,
            }
        );
        assert_eq!(error.resource(), ExternalLinkResource::Utf16Units);
        assert_eq!(error.value(), MAX_WIDE_STRING_UNITS);
        assert_eq!(error.maximum(), 0);
    }

    #[test]
    fn zero_budgets_are_allowed() {
        let limits = ExternalLinkLimits::builder()
            .max_part_bytes(0)
            .max_total_part_bytes(0)
            .max_opaque_bytes(0)
            .max_total_opaque_bytes(0)
            .max_utf16_units(0)
            .max_total_utf16_units(0)
            .max_records(0)
            .max_cache_records(0)
            .max_opaque_records(0)
            .max_links(0)
            .max_items(0)
            .max_matrices(0)
            .max_cells(0)
            .max_decoded_semantic_bytes(0)
            .max_retained_objects(0)
            .build()
            .unwrap();
        assert_eq!(limits.max_retained_objects(), 0);
    }

    #[test]
    fn failed_charge_is_atomic() {
        let limits = ExternalLinkLimits::builder()
            .max_part_bytes(2)
            .max_total_part_bytes(2)
            .max_links(2)
            .build()
            .unwrap();
        let mut budget = Budget::new(limits);
        budget.begin_link_part(2).unwrap();

        let error = budget.begin_link_part(1).unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                resource: ExternalLinkResource::TotalPartBytes,
                actual: 3,
                maximum: 2,
            }
        ));
        assert_eq!(budget.usage.total_part_bytes, 2);
        assert_eq!(budget.usage.links, 1);
    }

    #[test]
    fn parts_and_links_are_accounted_across_the_operation() {
        let limits = ExternalLinkLimits::builder()
            .max_part_bytes(2)
            .max_total_part_bytes(3)
            .max_links(2)
            .build()
            .unwrap();
        let mut budget = Budget::new(limits);
        budget.begin_link_part(2).unwrap();
        budget.begin_link_part(1).unwrap();
        assert_eq!(budget.usage.total_part_bytes, 3);
        assert_eq!(budget.usage.links, 2);

        let error = budget.preflight_links(1).unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                resource: ExternalLinkResource::Links,
                actual: 3,
                maximum: 2,
            }
        ));
        assert_eq!(budget.usage.total_part_bytes, 3);
        assert_eq!(budget.usage.links, 2);

        let error = budget.begin_link_part(0).unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                resource: ExternalLinkResource::Links,
                actual: 3,
                maximum: 2,
            }
        ));
    }
}
