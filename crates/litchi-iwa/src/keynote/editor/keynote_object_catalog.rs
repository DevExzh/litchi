//! A bounded, payload-free object catalog for Keynote editor operations.
//!
//! `ObjectGraph` remains the compatibility graph used by the existing native
//! editor.  This catalog is deliberately separate: it records only the
//! physical slot of each object and compact facts about its messages.  The
//! parsed [`Archive`] passed to a callback is borrowed for the duration of
//! that callback and is never retained by the catalog.

use std::mem::size_of;
use std::sync::Arc;

use prost::Message;

use crate::archive::{Archive, ArchiveObject};
use crate::{Error, IWorkPackage};

const DEFAULT_MAX_ARCHIVES: usize = 1_024;
const DEFAULT_MAX_ARCHIVE_READS: usize = 2_048;
const DEFAULT_MAX_OBJECTS: usize = 1_000_000;
const DEFAULT_MAX_MESSAGES: usize = 4_000_000;
const DEFAULT_MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;
const DEFAULT_MAX_REFERENCE_EDGES: usize = 8_000_000;
const DEFAULT_MAX_RETAINED_BYTES: usize = 256 * 1024 * 1024;
const DEFAULT_MAX_SEMANTIC_DECODES: usize = 1_000_000;

// These are hard ceilings for the private operation profile.  A caller may
// select a lower profile, but cannot turn a malformed package into an
// effectively unbounded catalog by supplying `usize::MAX`.
const ABSOLUTE_MAX_ARCHIVES: usize = 16_384;
const ABSOLUTE_MAX_ARCHIVE_READS: usize = 32_768;
const ABSOLUTE_MAX_OBJECTS: usize = 4_000_000;
const ABSOLUTE_MAX_MESSAGES: usize = 16_000_000;
const ABSOLUTE_MAX_PAYLOAD_BYTES: usize = 2 * 1024 * 1024 * 1024;
const ABSOLUTE_MAX_REFERENCE_EDGES: usize = 32_000_000;
const ABSOLUTE_MAX_RETAINED_BYTES: usize = 2 * 1024 * 1024 * 1024;
const ABSOLUTE_MAX_SEMANTIC_DECODES: usize = 8_000_000;

/// A resource axis owned by one Keynote catalog operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum KeynoteObjectCatalogLimitKind {
    Archives,
    ArchiveReads,
    Objects,
    Messages,
    PayloadBytes,
    ReferenceEdges,
    RetainedBytes,
    SemanticDecodes,
}

impl std::fmt::Display for KeynoteObjectCatalogLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Self::Archives => "archives",
            Self::ArchiveReads => "archive reads",
            Self::Objects => "objects",
            Self::Messages => "messages",
            Self::PayloadBytes => "payload bytes",
            Self::ReferenceEdges => "reference edges",
            Self::RetainedBytes => "retained bytes",
            Self::SemanticDecodes => "semantic decodes",
        };
        formatter.write_str(name)
    }
}

/// Internal error contract for the bounded Keynote catalog.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub(super) enum KeynoteObjectCatalogError {
    #[error("Keynote object catalog {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: KeynoteObjectCatalogLimitKind,
        observed: usize,
        maximum: usize,
    },
    #[error("invalid Keynote object catalog {kind} limit {value}; expected 1..={maximum}")]
    InvalidLimit {
        kind: KeynoteObjectCatalogLimitKind,
        value: usize,
        maximum: usize,
    },
    #[error("Keynote object catalog allocation failed for {resource}: {amount}")]
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    #[error("invalid Keynote object catalog source: {0}")]
    InvalidSource(String),
    #[error("Keynote object catalog source operation failed: {0}")]
    Source(String),
}

type CatalogResult<T> = std::result::Result<T, KeynoteObjectCatalogError>;

/// Map a private catalog failure at an existing editor API boundary.
///
/// The catalog keeps its typed resource axes internally so callers can make
/// a policy decision before publication.  The legacy editor error type has no
/// equivalent typed limit variants, so this adapter deliberately maps the
/// private error to its parse-error surface at the boundary.
pub(super) fn map_catalog_error(error: KeynoteObjectCatalogError) -> Error {
    Error::ParseError(error.to_string())
}

/// Finite operation limits for [`KeynoteObjectCatalog`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct KeynoteObjectCatalogLimits {
    pub(super) max_archives: usize,
    pub(super) max_archive_reads: usize,
    pub(super) max_objects: usize,
    pub(super) max_messages: usize,
    pub(super) max_payload_bytes: usize,
    pub(super) max_reference_edges: usize,
    pub(super) max_retained_bytes: usize,
    pub(super) max_semantic_decodes: usize,
}

impl Default for KeynoteObjectCatalogLimits {
    fn default() -> Self {
        Self {
            max_archives: DEFAULT_MAX_ARCHIVES,
            max_archive_reads: DEFAULT_MAX_ARCHIVE_READS,
            max_objects: DEFAULT_MAX_OBJECTS,
            max_messages: DEFAULT_MAX_MESSAGES,
            max_payload_bytes: DEFAULT_MAX_PAYLOAD_BYTES,
            max_reference_edges: DEFAULT_MAX_REFERENCE_EDGES,
            max_retained_bytes: DEFAULT_MAX_RETAINED_BYTES,
            max_semantic_decodes: DEFAULT_MAX_SEMANTIC_DECODES,
        }
    }
}

impl KeynoteObjectCatalogLimits {
    fn validate(self) -> CatalogResult<()> {
        validate_limit(
            KeynoteObjectCatalogLimitKind::Archives,
            self.max_archives,
            ABSOLUTE_MAX_ARCHIVES,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::ArchiveReads,
            self.max_archive_reads,
            ABSOLUTE_MAX_ARCHIVE_READS,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::Objects,
            self.max_objects,
            ABSOLUTE_MAX_OBJECTS,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::Messages,
            self.max_messages,
            ABSOLUTE_MAX_MESSAGES,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::PayloadBytes,
            self.max_payload_bytes,
            ABSOLUTE_MAX_PAYLOAD_BYTES,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::ReferenceEdges,
            self.max_reference_edges,
            ABSOLUTE_MAX_REFERENCE_EDGES,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            self.max_retained_bytes,
            ABSOLUTE_MAX_RETAINED_BYTES,
        )?;
        validate_limit(
            KeynoteObjectCatalogLimitKind::SemanticDecodes,
            self.max_semantic_decodes,
            ABSOLUTE_MAX_SEMANTIC_DECODES,
        )?;
        Ok(())
    }
}

fn validate_limit(
    kind: KeynoteObjectCatalogLimitKind,
    value: usize,
    maximum: usize,
) -> CatalogResult<()> {
    if value == 0 || value > maximum {
        return Err(KeynoteObjectCatalogError::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

/// Measurements from one catalog build and any borrowed semantic reads.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct KeynoteObjectCatalogStats {
    pub(super) archive_reads: usize,
    pub(super) archives_scanned: usize,
    pub(super) objects_indexed: usize,
    pub(super) messages_indexed: usize,
    pub(super) payload_bytes: usize,
    pub(super) reference_edges: usize,
    pub(super) semantic_decodes: usize,
    pub(super) peak_live_archives: usize,
    /// Logical bytes occupied by retained names and compact descriptors.
    /// Payload bytes are deliberately absent from this value.
    pub(super) retained_bytes: usize,
    /// Payload bytes are borrowed only during callbacks and never retained.
    pub(super) retained_payload_bytes: usize,
}

/// The physical location of one object in the package's IWA members.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct KeynoteObjectSlot {
    pub(super) archive_index: u32,
    pub(super) object_index: u32,
}

/// A compact message fact.  No payload bytes are retained.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct KeynoteMessageDescriptor {
    pub(super) type_: u32,
    pub(super) data_length: usize,
}

/// A compact object fact and source slot.  Message descriptors are stored in
/// one contiguous range starting at `message_start`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct KeynoteObjectDescriptor {
    pub(super) identifier: u64,
    pub(super) slot: KeynoteObjectSlot,
    pub(super) message_start: u32,
    pub(super) message_count: u32,
    pub(super) payload_bytes: usize,
    pub(super) reference_edges: usize,
}

#[derive(Debug, Clone, Copy, Default)]
struct ArchivePlan {
    objects: usize,
    messages: usize,
    payload_bytes: usize,
    reference_edges: usize,
}

/// A bounded Keynote object catalog that owns no parsed archive or payload.
///
/// The catalog is tied to the immutable package revision that produced it.
/// Callers must discard it after mutating a package clone and should use
/// [`Self::ensure_current`] before a borrowed lookup.
#[derive(Debug, Clone)]
pub(super) struct KeynoteObjectCatalog {
    package_revision: u64,
    limits: KeynoteObjectCatalogLimits,
    archive_names: Arc<[Box<str>]>,
    objects: Vec<KeynoteObjectDescriptor>,
    messages: Vec<KeynoteMessageDescriptor>,
    stats: KeynoteObjectCatalogStats,
}

impl KeynoteObjectCatalog {
    pub(super) fn build(package: &IWorkPackage) -> CatalogResult<Self> {
        Self::build_with_limits(package, KeynoteObjectCatalogLimits::default())
    }

    pub(super) fn build_with_limits(
        package: &IWorkPackage,
        limits: KeynoteObjectCatalogLimits,
    ) -> CatalogResult<Self> {
        limits.validate()?;

        let archive_count = package.iwa_entry_names().count();
        check_limit(
            KeynoteObjectCatalogLimitKind::Archives,
            archive_count,
            limits.max_archives,
        )?;
        // Every member is read once while the index is built.  Reject an
        // insufficient read budget before allocating names or entering the
        // archive cache; later borrowed lookups consume the same axis.
        check_limit(
            KeynoteObjectCatalogLimitKind::ArchiveReads,
            archive_count,
            limits.max_archive_reads,
        )?;
        let name_bytes = package.iwa_entry_names().try_fold(0usize, |total, name| {
            total
                .checked_add(name.len())
                .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                    kind: KeynoteObjectCatalogLimitKind::RetainedBytes,
                    observed: usize::MAX,
                    maximum: limits.max_retained_bytes,
                })
        })?;
        let names_retained = checked_mul(
            archive_count,
            size_of::<Box<str>>(),
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            limits.max_retained_bytes,
        )?
        .checked_add(name_bytes)
        .ok_or(KeynoteObjectCatalogError::LimitExceeded {
            kind: KeynoteObjectCatalogLimitKind::RetainedBytes,
            observed: usize::MAX,
            maximum: limits.max_retained_bytes,
        })?;
        check_limit(
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            names_retained,
            limits.max_retained_bytes,
        )?;

        let mut archive_names = Vec::new();
        archive_names
            .try_reserve_exact(archive_count)
            .map_err(|_| KeynoteObjectCatalogError::Allocation {
                resource: "Keynote catalog archive names",
                amount: archive_count,
            })?;
        for name in package.iwa_entry_names() {
            archive_names.push(fallible_boxed_str(name)?);
        }

        let mut catalog = Self {
            package_revision: package.mutation_revision(),
            limits,
            archive_names: Arc::from(archive_names.into_boxed_slice()),
            objects: Vec::new(),
            messages: Vec::new(),
            stats: KeynoteObjectCatalogStats {
                retained_bytes: names_retained,
                ..KeynoteObjectCatalogStats::default()
            },
        };

        let archive_names = Arc::clone(&catalog.archive_names);
        for (archive_index, archive_name) in package.iwa_entry_names().enumerate() {
            let archive_index = u32::try_from(archive_index).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource("archive index exceeds u32".to_owned())
            })?;
            // Keep the name borrowed from the package/Arc local while the
            // catalog is mutably updated; no temporary owned member name is
            // needed for a scan.
            debug_assert_eq!(
                archive_names
                    .get(usize::try_from(archive_index).unwrap_or(usize::MAX))
                    .map(|name| name.as_ref()),
                Some(archive_name)
            );
            catalog.scan_archive(package, archive_name, |catalog, archive| {
                catalog.index_archive(archive_index, archive)
            })?;
        }

        catalog
            .objects
            .sort_unstable_by_key(|object| object.identifier);
        for pair in catalog.objects.windows(2) {
            if pair[0].identifier == pair[1].identifier {
                return Err(KeynoteObjectCatalogError::InvalidSource(format!(
                    "object identifier {} appears more than once",
                    pair[0].identifier
                )));
            }
        }
        Ok(catalog)
    }

    fn index_archive(&mut self, archive_index: u32, archive: &Archive) -> CatalogResult<()> {
        let mut plan = ArchivePlan {
            objects: archive.objects.len(),
            ..ArchivePlan::default()
        };
        for (object_index, object) in archive.objects.iter().enumerate() {
            if object.archive_info.identifier.is_none() {
                return Err(KeynoteObjectCatalogError::InvalidSource(format!(
                    "object {object_index} in archive {archive_index} has no identifier"
                )));
            }
            if object.archive_info.identifier == Some(0) {
                return Err(KeynoteObjectCatalogError::InvalidSource(format!(
                    "object {object_index} in archive {archive_index} has zero identifier"
                )));
            }
            plan.messages = plan.messages.checked_add(object.messages.len()).ok_or(
                KeynoteObjectCatalogError::LimitExceeded {
                    kind: KeynoteObjectCatalogLimitKind::Messages,
                    observed: usize::MAX,
                    maximum: self.limits.max_messages,
                },
            )?;
            for message in &object.messages {
                plan.payload_bytes = plan.payload_bytes.checked_add(message.data.len()).ok_or(
                    KeynoteObjectCatalogError::LimitExceeded {
                        kind: KeynoteObjectCatalogLimitKind::PayloadBytes,
                        observed: usize::MAX,
                        maximum: self.limits.max_payload_bytes,
                    },
                )?;
            }
            plan.reference_edges = plan
                .reference_edges
                .checked_add(object_reference_edges(object)?)
                .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                    kind: KeynoteObjectCatalogLimitKind::ReferenceEdges,
                    observed: usize::MAX,
                    maximum: self.limits.max_reference_edges,
                })?;
            u32::try_from(object_index).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource("object slot index exceeds u32".to_owned())
            })?;
        }

        let retained_objects = checked_mul(
            plan.objects,
            size_of::<KeynoteObjectDescriptor>(),
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            self.limits.max_retained_bytes,
        )?;
        let retained_messages = checked_mul(
            plan.messages,
            size_of::<KeynoteMessageDescriptor>(),
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            self.limits.max_retained_bytes,
        )?;
        let retained_delta = retained_objects.checked_add(retained_messages).ok_or(
            KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::RetainedBytes,
                observed: usize::MAX,
                maximum: self.limits.max_retained_bytes,
            },
        )?;

        self.check_usage(
            KeynoteObjectCatalogLimitKind::Objects,
            self.stats.objects_indexed,
            plan.objects,
            self.limits.max_objects,
        )?;
        self.check_usage(
            KeynoteObjectCatalogLimitKind::Messages,
            self.stats.messages_indexed,
            plan.messages,
            self.limits.max_messages,
        )?;
        self.check_usage(
            KeynoteObjectCatalogLimitKind::PayloadBytes,
            self.stats.payload_bytes,
            plan.payload_bytes,
            self.limits.max_payload_bytes,
        )?;
        self.check_usage(
            KeynoteObjectCatalogLimitKind::ReferenceEdges,
            self.stats.reference_edges,
            plan.reference_edges,
            self.limits.max_reference_edges,
        )?;
        self.check_usage(
            KeynoteObjectCatalogLimitKind::RetainedBytes,
            self.stats.retained_bytes,
            retained_delta,
            self.limits.max_retained_bytes,
        )?;

        // The usage checks above are intentionally before these allocations.
        self.objects.try_reserve_exact(plan.objects).map_err(|_| {
            KeynoteObjectCatalogError::Allocation {
                resource: "Keynote catalog object descriptors",
                amount: plan.objects,
            }
        })?;
        self.messages
            .try_reserve_exact(plan.messages)
            .map_err(|_| KeynoteObjectCatalogError::Allocation {
                resource: "Keynote catalog message descriptors",
                amount: plan.messages,
            })?;

        for (object_index, object) in archive.objects.iter().enumerate() {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                KeynoteObjectCatalogError::InvalidSource(
                    "object identifier disappeared during catalog scan".to_owned(),
                )
            })?;
            if identifier == 0 {
                return Err(KeynoteObjectCatalogError::InvalidSource(
                    "object identifier is zero during catalog indexing".to_owned(),
                ));
            }
            let object_index = u32::try_from(object_index).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource("object slot index exceeds u32".to_owned())
            })?;
            let message_count = u32::try_from(object.messages.len()).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource("message count exceeds u32".to_owned())
            })?;
            let message_start = u32::try_from(self.messages.len()).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource(
                    "message descriptor index exceeds u32".to_owned(),
                )
            })?;
            for message in &object.messages {
                self.messages.push(KeynoteMessageDescriptor {
                    type_: message.type_,
                    data_length: message.data.len(),
                });
            }
            self.objects.push(KeynoteObjectDescriptor {
                identifier,
                slot: KeynoteObjectSlot {
                    archive_index,
                    object_index,
                },
                message_start,
                message_count,
                payload_bytes: object
                    .messages
                    .iter()
                    .try_fold(0usize, |total, message| {
                        total.checked_add(message.data.len())
                    })
                    .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                        kind: KeynoteObjectCatalogLimitKind::PayloadBytes,
                        observed: usize::MAX,
                        maximum: self.limits.max_payload_bytes,
                    })?,
                reference_edges: object_reference_edges(object)?,
            });
        }
        self.stats.objects_indexed = self.stats.objects_indexed.checked_add(plan.objects).ok_or(
            KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::Objects,
                observed: usize::MAX,
                maximum: self.limits.max_objects,
            },
        )?;
        self.stats.messages_indexed = self
            .stats
            .messages_indexed
            .checked_add(plan.messages)
            .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::Messages,
                observed: usize::MAX,
                maximum: self.limits.max_messages,
            })?;
        self.stats.payload_bytes = self
            .stats
            .payload_bytes
            .checked_add(plan.payload_bytes)
            .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::PayloadBytes,
                observed: usize::MAX,
                maximum: self.limits.max_payload_bytes,
            })?;
        self.stats.reference_edges = self
            .stats
            .reference_edges
            .checked_add(plan.reference_edges)
            .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::ReferenceEdges,
                observed: usize::MAX,
                maximum: self.limits.max_reference_edges,
            })?;
        self.stats.retained_bytes = self
            .stats
            .retained_bytes
            .checked_add(retained_delta)
            .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::RetainedBytes,
                observed: usize::MAX,
                maximum: self.limits.max_retained_bytes,
            })?;
        Ok(())
    }

    fn scan_archive<T, F>(
        &mut self,
        package: &IWorkPackage,
        archive_name: &str,
        read: F,
    ) -> CatalogResult<T>
    where
        F: FnOnce(&mut Self, &Archive) -> CatalogResult<T>,
    {
        self.check_usage(
            KeynoteObjectCatalogLimitKind::Archives,
            self.stats.archives_scanned,
            1,
            self.limits.max_archives,
        )?;
        self.stats.archives_scanned = self.stats.archives_scanned.checked_add(1).ok_or(
            KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::Archives,
                observed: usize::MAX,
                maximum: self.limits.max_archives,
            },
        )?;
        self.with_archive(package, archive_name, read)
    }

    fn with_archive<T, F>(
        &mut self,
        package: &IWorkPackage,
        archive_name: &str,
        read: F,
    ) -> CatalogResult<T>
    where
        F: FnOnce(&mut Self, &Archive) -> CatalogResult<T>,
    {
        self.ensure_current(package)?;
        self.check_usage(
            KeynoteObjectCatalogLimitKind::ArchiveReads,
            self.stats.archive_reads,
            1,
            self.limits.max_archive_reads,
        )?;
        self.stats.archive_reads = self.stats.archive_reads.checked_add(1).ok_or(
            KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::ArchiveReads,
                observed: usize::MAX,
                maximum: self.limits.max_archive_reads,
            },
        )?;
        self.stats.peak_live_archives = self.stats.peak_live_archives.max(1);
        let mut callback_error = None;
        let result = package.with_parsed_archive(archive_name, |archive| {
            read(self, archive).map_err(|error| {
                let boundary_error = map_catalog_error(error.clone());
                callback_error = Some(error);
                boundary_error
            })
        });
        match result {
            Ok(value) => Ok(value),
            Err(error) => callback_error.take().map_or_else(
                || Err(KeynoteObjectCatalogError::Source(error.to_string())),
                Err,
            ),
        }
    }

    fn check_usage(
        &self,
        kind: KeynoteObjectCatalogLimitKind,
        current: usize,
        additional: usize,
        maximum: usize,
    ) -> CatalogResult<()> {
        let observed =
            current
                .checked_add(additional)
                .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                    kind,
                    observed: usize::MAX,
                    maximum,
                })?;
        check_limit(kind, observed, maximum)
    }

    fn record_semantic_decode(&mut self) -> CatalogResult<()> {
        self.check_usage(
            KeynoteObjectCatalogLimitKind::SemanticDecodes,
            self.stats.semantic_decodes,
            1,
            self.limits.max_semantic_decodes,
        )?;
        self.stats.semantic_decodes = self.stats.semantic_decodes.checked_add(1).ok_or(
            KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::SemanticDecodes,
                observed: usize::MAX,
                maximum: self.limits.max_semantic_decodes,
            },
        )?;
        Ok(())
    }

    /// Reject use of a catalog after the package revision has changed.
    pub(super) fn ensure_current(&self, package: &IWorkPackage) -> CatalogResult<()> {
        if package.mutation_revision() != self.package_revision {
            return Err(KeynoteObjectCatalogError::InvalidSource(
                "catalog is stale after package mutation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Find the archive member containing an object without reading it.
    pub(super) fn archive_name(&self, identifier: u64) -> CatalogResult<&str> {
        let object = self.object_descriptor(identifier)?;
        self.archive_name_at(object.slot.archive_index)
    }

    fn archive_name_at(&self, archive_index: u32) -> CatalogResult<&str> {
        self.archive_names
            .get(usize::try_from(archive_index).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource(
                    "archive index does not fit usize".to_owned(),
                )
            })?)
            .map(|name| name.as_ref())
            .ok_or_else(|| {
                KeynoteObjectCatalogError::InvalidSource(format!(
                    "archive slot {archive_index} is missing"
                ))
            })
    }

    pub(super) fn object_descriptor(
        &self,
        identifier: u64,
    ) -> CatalogResult<&KeynoteObjectDescriptor> {
        self.objects
            .binary_search_by_key(&identifier, |object| object.identifier)
            .ok()
            .and_then(|index| self.objects.get(index))
            .ok_or_else(|| {
                KeynoteObjectCatalogError::InvalidSource(format!("object {identifier} is missing"))
            })
    }

    /// Iterate object identifiers in deterministic package order.
    ///
    /// The descriptors are sorted once at build completion, so this lookup
    /// allocates nothing and does not consume any catalog resource axis.
    pub(super) fn object_identifiers(&self) -> impl Iterator<Item = u64> + '_ {
        self.objects.iter().map(|object| object.identifier)
    }

    /// Return the number of indexed objects without allocating an identifier
    /// list.  The paired [`Self::object_identifier_at`] method lets callers
    /// walk deterministic catalog order while retaining their own scratch
    /// budget.
    pub(super) const fn object_count(&self) -> usize {
        self.objects.len()
    }

    /// Return one object identifier in deterministic catalog order.
    pub(super) fn object_identifier_at(&self, index: usize) -> CatalogResult<u64> {
        self.objects
            .get(index)
            .map(|object| object.identifier)
            .ok_or_else(|| {
                KeynoteObjectCatalogError::InvalidSource(format!(
                    "object catalog index {index} is out of bounds"
                ))
            })
    }

    pub(super) fn message_descriptors(
        &self,
        identifier: u64,
    ) -> CatalogResult<&[KeynoteMessageDescriptor]> {
        let object = self.object_descriptor(identifier)?;
        let start = usize::try_from(object.message_start).map_err(|_| {
            KeynoteObjectCatalogError::InvalidSource(
                "message descriptor start does not fit usize".to_owned(),
            )
        })?;
        let count = usize::try_from(object.message_count).map_err(|_| {
            KeynoteObjectCatalogError::InvalidSource(
                "message descriptor count does not fit usize".to_owned(),
            )
        })?;
        let end = start.checked_add(count).ok_or_else(|| {
            KeynoteObjectCatalogError::InvalidSource(
                "message descriptor range overflows usize".to_owned(),
            )
        })?;
        self.messages.get(start..end).ok_or_else(|| {
            KeynoteObjectCatalogError::InvalidSource(
                "message descriptor range is out of bounds".to_owned(),
            )
        })
    }

    pub(super) fn message_count(&self, identifier: u64) -> CatalogResult<usize> {
        Ok(self.message_descriptors(identifier)?.len())
    }

    pub(super) fn message_type_count(
        &self,
        identifier: u64,
        message_type: u32,
    ) -> CatalogResult<usize> {
        Ok(self
            .message_descriptors(identifier)?
            .iter()
            .filter(|message| message.type_ == message_type)
            .count())
    }

    pub(super) fn message_types(
        &self,
        identifier: u64,
    ) -> CatalogResult<impl Iterator<Item = u32> + '_> {
        Ok(self
            .message_descriptors(identifier)?
            .iter()
            .map(|message| message.type_))
    }

    /// Borrow one complete archive object for a bounded callback.
    ///
    /// The callback must not retain the object or any payload reference. One
    /// semantic decode and one archive read are charged before callback use.
    pub(super) fn with_object<T, F>(
        &mut self,
        package: &IWorkPackage,
        identifier: u64,
        read: F,
    ) -> CatalogResult<T>
    where
        F: FnOnce(&ArchiveObject) -> CatalogResult<T>,
    {
        self.ensure_current(package)?;
        let descriptor = *self.object_descriptor(identifier)?;
        self.record_semantic_decode()?;
        let archive_names = Arc::clone(&self.archive_names);
        let archive_name = archive_names
            .get(usize::try_from(descriptor.slot.archive_index).map_err(|_| {
                KeynoteObjectCatalogError::InvalidSource(
                    "archive index does not fit usize".to_owned(),
                )
            })?)
            .map(|name| name.as_ref())
            .ok_or_else(|| {
                KeynoteObjectCatalogError::InvalidSource(
                    "object archive slot is missing".to_owned(),
                )
            })?;
        self.with_archive(package, archive_name, |_catalog, archive| {
            let object = archive
                .objects
                .get(usize::try_from(descriptor.slot.object_index).map_err(|_| {
                    KeynoteObjectCatalogError::InvalidSource(
                        "object index does not fit usize".to_owned(),
                    )
                })?)
                .ok_or_else(|| {
                    KeynoteObjectCatalogError::InvalidSource(
                        "object slot is out of bounds".to_owned(),
                    )
                })?;
            if object.archive_info.identifier != Some(identifier)
                || object.messages.len()
                    != usize::try_from(descriptor.message_count).map_err(|_| {
                        KeynoteObjectCatalogError::InvalidSource(
                            "message count does not fit usize".to_owned(),
                        )
                    })?
            {
                return Err(KeynoteObjectCatalogError::InvalidSource(
                    "object slot no longer matches its catalog descriptor".to_owned(),
                ));
            }
            read(object)
        })
    }

    /// Borrow exactly one payload of `message_type` for a bounded callback.
    pub(super) fn with_message_data_type<T, F>(
        &mut self,
        package: &IWorkPackage,
        identifier: u64,
        message_type: u32,
        type_name: &str,
        read: F,
    ) -> CatalogResult<T>
    where
        F: FnOnce(&[u8]) -> CatalogResult<T>,
    {
        let count = self.message_type_count(identifier, message_type)?;
        if count == 0 {
            return Err(KeynoteObjectCatalogError::InvalidSource(format!(
                "object {identifier} has no {type_name} payload"
            )));
        }
        if count != 1 {
            return Err(KeynoteObjectCatalogError::InvalidSource(format!(
                "object {identifier} repeats its {type_name} payload"
            )));
        }
        self.with_object(package, identifier, |object| {
            let message = object
                .messages
                .iter()
                .find(|message| message.type_ == message_type)
                .ok_or_else(|| {
                    KeynoteObjectCatalogError::InvalidSource(
                        "message descriptor no longer matches its source".to_owned(),
                    )
                })?;
            read(message.data.as_slice())
        })
    }

    /// Decode exactly one selected message while retaining no source payload.
    pub(super) fn decode_type<T: Message + Default>(
        &mut self,
        package: &IWorkPackage,
        identifier: u64,
        message_type: u32,
        type_name: &str,
    ) -> CatalogResult<T> {
        self.with_message_data_type(package, identifier, message_type, type_name, |data| {
            T::decode(data).map_err(|error| {
                KeynoteObjectCatalogError::InvalidSource(format!(
                    "object {identifier} has malformed {type_name} payload: {error}"
                ))
            })
        })
    }

    #[cfg(test)]
    pub(super) fn stats(&self) -> KeynoteObjectCatalogStats {
        self.stats
    }
}

fn object_reference_edges(object: &ArchiveObject) -> CatalogResult<usize> {
    let mut total = 0usize;
    for message in &object.archive_info.message_infos {
        total = total
            .checked_add(message.object_references.len())
            .and_then(|total| total.checked_add(message.data_references.len()))
            .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::ReferenceEdges,
                observed: usize::MAX,
                maximum: usize::MAX,
            })?;
        for field in &message.field_infos {
            total = total
                .checked_add(field.object_references.len())
                .and_then(|total| total.checked_add(field.data_references.len()))
                .ok_or(KeynoteObjectCatalogError::LimitExceeded {
                    kind: KeynoteObjectCatalogLimitKind::ReferenceEdges,
                    observed: usize::MAX,
                    maximum: usize::MAX,
                })?;
        }
    }
    Ok(total)
}

fn checked_mul(
    left: usize,
    right: usize,
    kind: KeynoteObjectCatalogLimitKind,
    maximum: usize,
) -> CatalogResult<usize> {
    left.checked_mul(right)
        .ok_or(KeynoteObjectCatalogError::LimitExceeded {
            kind,
            observed: usize::MAX,
            maximum,
        })
}

fn check_limit(
    kind: KeynoteObjectCatalogLimitKind,
    observed: usize,
    maximum: usize,
) -> CatalogResult<()> {
    if observed > maximum {
        return Err(KeynoteObjectCatalogError::LimitExceeded {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn fallible_boxed_str(value: &str) -> CatalogResult<Box<str>> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_| KeynoteObjectCatalogError::Allocation {
            resource: "Keynote catalog archive name",
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};

    fn package_with_archives(archives: &[(&str, Vec<ArchiveObject>)]) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        for (name, objects) in archives {
            package
                .replace_archive(
                    name,
                    &Archive {
                        objects: objects.clone(),
                    },
                )
                .expect("synthetic archive is valid");
        }
        package
    }

    fn object(identifier: u64, messages: Vec<(u32, Vec<u8>)>) -> ArchiveObject {
        ArchiveObject::new(
            identifier,
            messages
                .into_iter()
                .map(|(type_, data)| RawMessage { type_, data })
                .collect(),
        )
        .expect("synthetic object is valid")
    }

    fn one_object_package() -> IWorkPackage {
        package_with_archives(&[(
            "Index/Document.iwa",
            vec![object(42, vec![(2_011, vec![1, 2, 3])])],
        )])
    }

    fn assert_limit(error: CatalogResult<()>, kind: KeynoteObjectCatalogLimitKind) {
        assert!(matches!(
            error,
            Err(KeynoteObjectCatalogError::LimitExceeded { kind: actual, .. }) if actual == kind
        ));
    }

    #[test]
    fn catalog_retains_slots_and_facts_but_not_payloads() {
        let package = one_object_package();
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");

        assert_eq!(
            catalog.archive_name(42).expect("archive name"),
            "Index/Document.iwa"
        );
        assert_eq!(catalog.message_count(42).expect("message count"), 1);
        assert_eq!(
            catalog.message_type_count(42, 2_011).expect("type count"),
            1
        );
        assert_eq!(catalog.object_identifiers().collect::<Vec<_>>(), vec![42]);
        assert_eq!(
            catalog
                .message_types(42)
                .expect("message types")
                .collect::<Vec<_>>(),
            vec![2_011]
        );
        let descriptor = *catalog.object_descriptor(42).expect("object descriptor");
        assert_eq!(descriptor.payload_bytes, 3);
        assert_eq!(catalog.stats().payload_bytes, 3);
        assert_eq!(catalog.stats().peak_live_archives, 1);
        assert!(catalog.stats().retained_bytes > 0);
        assert_eq!(catalog.stats().retained_payload_bytes, 0);
        assert_eq!(catalog.stats().archives_scanned, 1);
        assert_eq!(catalog.stats().archive_reads, 1);

        catalog
            .with_object(&package, 42, |object| {
                assert_eq!(object.messages[0].data, [1, 2, 3]);
                Ok(())
            })
            .expect("borrow object");
        catalog
            .with_message_data_type(&package, 42, 2_011, "shape", |data| {
                assert_eq!(data, [1, 2, 3]);
                Ok(())
            })
            .expect("borrow message");
        assert_eq!(catalog.stats().archives_scanned, 1);
        assert_eq!(catalog.stats().archive_reads, 3);
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn catalog_rejects_duplicate_object_identifiers_across_members() {
        let package = package_with_archives(&[
            ("Index/Document.iwa", vec![object(42, vec![(1, vec![])])]),
            ("Index/Slide-1.iwa", vec![object(42, vec![(2, vec![])])]),
        ]);
        assert!(matches!(
            KeynoteObjectCatalog::build(&package),
            Err(KeynoteObjectCatalogError::InvalidSource(message))
                if message.contains("appears more than once")
        ));
    }

    #[test]
    fn catalog_indexed_object_access_is_allocation_free() {
        let package = multi_object_package();
        let catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert_eq!(catalog.object_count(), 3);
        assert_eq!(catalog.object_identifier_at(0).expect("first object"), 101);
        assert_eq!(catalog.object_identifier_at(2).expect("last object"), 103);
        assert!(matches!(
            catalog.object_identifier_at(3),
            Err(KeynoteObjectCatalogError::InvalidSource(message))
                if message.contains("out of bounds")
        ));
    }

    #[test]
    fn catalog_rejects_each_finite_build_axis_before_descriptor_reserve() {
        let package = package_with_archives(&[
            (
                "Index/One.iwa",
                vec![object(1, vec![(1, vec![1, 2]), (2, vec![3])])],
            ),
            ("Index/Two.iwa", vec![object(2, vec![(3, vec![4])])]),
        ]);

        let defaults = KeynoteObjectCatalogLimits::default();
        let mut limits = defaults;
        limits.max_archives = 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Archives,
        );

        let mut limits = defaults;
        limits.max_archive_reads = 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::ArchiveReads,
        );

        let mut limits = defaults;
        limits.max_objects = 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Objects,
        );

        let mut limits = defaults;
        limits.max_messages = 2;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Messages,
        );

        let mut limits = defaults;
        limits.max_payload_bytes = 3;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::PayloadBytes,
        );

        let mut limits = defaults;
        limits.max_reference_edges = 1;
        let mut referenced = object(1, vec![(1, vec![1])]);
        referenced.archive_info.message_infos[0]
            .object_references
            .extend([2, 3]);
        let package = package_with_archives(&[("Index/Document.iwa", vec![referenced])]);
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::ReferenceEdges,
        );

        let mut limits = defaults;
        limits.max_retained_bytes = 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::RetainedBytes,
        );
    }

    #[test]
    fn catalog_charges_borrowed_semantic_reads_and_rejects_stale_revisions() {
        let package = one_object_package();
        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_semantic_decodes = 1;
        let mut catalog =
            KeynoteObjectCatalog::build_with_limits(&package, limits).expect("catalog");
        catalog
            .with_message_data_type(&package, 42, 2_011, "shape", |_| Ok(()))
            .expect("first semantic read");
        assert!(matches!(
            catalog.with_message_data_type(&package, 42, 2_011, "shape", |_| Ok(())),
            Err(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::SemanticDecodes,
                ..
            })
        ));

        let mut changed = package.clone();
        changed
            .update_archive("Index/Document.iwa", |archive| {
                archive.objects[0].messages[0].data.push(9);
                Ok(())
            })
            .expect("mutation");
        assert!(matches!(
            catalog.ensure_current(&changed),
            Err(KeynoteObjectCatalogError::InvalidSource(message))
                if message.contains("stale")
        ));
    }

    #[test]
    fn catalog_validates_nonzero_limits_and_missing_objects() {
        let package = one_object_package();
        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_objects = 0;
        assert!(matches!(
            KeynoteObjectCatalog::build_with_limits(&package, limits),
            Err(KeynoteObjectCatalogError::InvalidLimit {
                kind: KeynoteObjectCatalogLimitKind::Objects,
                value: 0,
                ..
            })
        ));

        let catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(matches!(
            catalog.object_descriptor(99),
            Err(KeynoteObjectCatalogError::InvalidSource(message))
                if message.contains("missing")
        ));
    }

    fn multi_object_package() -> IWorkPackage {
        let mut first = object(101, vec![(6_001, vec![0x01, 0x02]), (6_002, vec![0x03])]);
        first.archive_info.message_infos[0]
            .object_references
            .extend([102, 103]);
        let second = object(
            102,
            vec![(6_001, vec![0x04, 0x05, 0x06]), (6_003, vec![0x07])],
        );
        let third = object(103, vec![(6_001, vec![0x08, 0x09])]);
        package_with_archives(&[
            ("Index/CalculationEngine.iwa", vec![first, second]),
            ("Index/Slide-1.iwa", vec![third]),
        ])
    }

    #[test]
    fn catalog_multi_object_facts_match_legacy_graph_without_rescanning() {
        let package = multi_object_package();
        let before = package.to_bytes().expect("package bytes");
        let graph =
            crate::keynote::editor::slide_graph::ObjectGraph::read(&package).expect("legacy graph");
        let catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        let built = catalog.stats();

        assert_eq!(built.archives_scanned, 2);
        assert_eq!(built.archive_reads, 2);
        assert_eq!(built.objects_indexed, 3);
        assert_eq!(built.messages_indexed, 5);
        assert_eq!(built.payload_bytes, 9);
        assert_eq!(built.reference_edges, 2);
        assert_eq!(built.peak_live_archives, 1);
        assert_eq!(built.retained_payload_bytes, 0);

        for identifier in [101, 102, 103] {
            let legacy_messages = graph
                .objects
                .get(&identifier)
                .expect("legacy object descriptor");
            assert_eq!(
                catalog.archive_name(identifier).expect("catalog archive"),
                graph.archive_name(identifier).expect("legacy archive")
            );
            let descriptors = catalog
                .message_descriptors(identifier)
                .expect("catalog message descriptors");
            assert_eq!(descriptors.len(), legacy_messages.len());
            for (descriptor, legacy) in descriptors.iter().zip(legacy_messages) {
                assert_eq!(descriptor.type_, legacy.type_);
                assert_eq!(descriptor.data_length, legacy.data.len());
            }
        }

        // Listing facts are served entirely from the compact index: the
        // package-wide archive scan is performed once, regardless of the
        // number of table-like objects queried.
        assert_eq!(catalog.stats().archives_scanned, built.archives_scanned);
        assert_eq!(catalog.stats().archive_reads, built.archive_reads);
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn catalog_rejects_zero_and_missing_object_identifiers_atomically() {
        let zero_package = package_with_archives(&[(
            "Index/Document.iwa",
            vec![object(0, vec![(2_011, vec![1, 2, 3])])],
        )]);
        let zero_before = zero_package.to_bytes().expect("zero package bytes");
        assert!(matches!(
            KeynoteObjectCatalog::build(&zero_package),
            Err(KeynoteObjectCatalogError::InvalidSource(message))
                if message.contains("zero identifier")
        ));
        assert_eq!(
            zero_package.to_bytes().expect("zero package bytes"),
            zero_before
        );

        let package = one_object_package();
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        let mut missing = object(43, vec![(2_011, vec![1])]);
        missing.archive_info.identifier = None;
        let error = catalog
            .index_archive(
                0,
                &Archive {
                    objects: vec![missing],
                },
            )
            .expect_err("missing identifier");
        assert!(matches!(
            error,
            KeynoteObjectCatalogError::InvalidSource(message)
                if message.contains("has no identifier")
        ));
        assert!(catalog.object_descriptor(43).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn catalog_rejects_multi_object_max_minus_one_limits_before_publication() {
        let package = multi_object_package();
        let before = package.to_bytes().expect("package bytes");
        let baseline = KeynoteObjectCatalog::build(&package)
            .expect("baseline catalog")
            .stats();
        assert!(baseline.archives_scanned > 1);
        assert!(baseline.objects_indexed > 1);
        assert!(baseline.messages_indexed > 1);
        assert!(baseline.payload_bytes > 1);
        assert!(baseline.reference_edges > 1);
        assert!(baseline.retained_bytes > 1);

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_archives = baseline.archives_scanned - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Archives,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_archive_reads = baseline.archive_reads - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::ArchiveReads,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_objects = baseline.objects_indexed - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Objects,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_messages = baseline.messages_indexed - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::Messages,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_payload_bytes = baseline.payload_bytes - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::PayloadBytes,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_reference_edges = baseline.reference_edges - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::ReferenceEdges,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_retained_bytes = baseline.retained_bytes - 1;
        assert_limit(
            KeynoteObjectCatalog::build_with_limits(&package, limits).map(|_| ()),
            KeynoteObjectCatalogLimitKind::RetainedBytes,
        );

        let mut limits = KeynoteObjectCatalogLimits::default();
        limits.max_semantic_decodes = 1;
        let mut catalog = KeynoteObjectCatalog::build_with_limits(&package, limits)
            .expect("semantic-read catalog");
        catalog
            .with_message_data_type(&package, 101, 6_001, "table model", |_| Ok(()))
            .expect("first semantic read");
        assert!(matches!(
            catalog.with_message_data_type(&package, 102, 6_001, "table model", |_| Ok(())),
            Err(KeynoteObjectCatalogError::LimitExceeded {
                kind: KeynoteObjectCatalogLimitKind::SemanticDecodes,
                ..
            })
        ));

        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }
}
