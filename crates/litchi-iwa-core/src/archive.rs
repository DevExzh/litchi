//! Bounded framing for decompressed IWA archive streams.
//!
//! An IWA stream is a sequence of objects. Each object contains a varint
//! length, a `TSP.ArchiveInfo` protobuf, and the payloads described by that
//! header's `TSP.MessageInfo` entries. This module owns only that physical
//! framing. It does not resolve message types or decode application payloads.

#![allow(
    clippy::cast_possible_truncation,
    reason = "Varint encoding masks each narrowed byte to its seven-bit wire value."
)]
#![allow(
    clippy::map_err_ignore,
    reason = "The public physical-layer error API intentionally normalizes conversion and allocation failures."
)]
#![allow(
    clippy::missing_errors_doc,
    reason = "The module's Result-returning methods share the crate-level physical framing error contract."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "These names mirror the neutral IWA wire vocabulary and remain source-compatible with the extracted layer."
)]
#![allow(
    clippy::shadow_reuse,
    reason = "Validated limit profiles deliberately retain the public parameter name."
)]

use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_common::wire::{
    WireDescent, WireField, WirePreflight, WireVisit, parse_wire_fields, preflight_wire_tree,
};
use litchi_iwa_common::{Error as WireError, LimitKind as WireLimitKind, WireLimits};
use litchi_iwa_protos::archive_codec;
use smallvec::SmallVec;

use crate::{Error, HeaderKind, HeaderOperation, LimitKind, Limits, Result};

mod identity_clone;
mod ingress;
mod inventory;
pub use ingress::ArchiveIngressRequirements;
pub use inventory::{
    ArchiveCloneInventory, ArchiveCloneInventoryChunk, ArchiveCloneInventoryError,
    ArchiveCloneInventoryKind, ArchiveCloneInventoryResult, ArchiveCloneInventoryVisitor,
    ArchiveEncodingPreparationChunk, ArchiveEncodingPreparationError,
    ArchiveEncodingPreparationKind, ArchiveEncodingPreparationResult,
    ArchiveEncodingPreparationVisitor,
};

use identity_clone::{
    checked_replacement_lengths, clone_archive_info_with_identity_remap,
    clone_raw_messages_with_limits, preflight_clone_archive_info_header_length,
    prepare_identity_remap, retained_clone_source_header, rewrite_clone_archive_info_header,
    validate_clone_arguments, validate_clone_header_scratch, validate_clone_input_scratch,
    validate_clone_metadata_semantics, validate_clone_object_size,
    validate_clone_self_reference_mapping,
};
use inventory::admit_archive_encoding_object;

const MAX_VARINT_BYTES: usize = 10;
const DEFAULT_MESSAGE_VERSIONS: [u32; 3] = [1, 0, 5];
const MESSAGE_VERSIONS_INLINE_CAPACITY: usize = DEFAULT_MESSAGE_VERSIONS.len();
const MINIMAL_MESSAGE_INFO_MEMORY: usize = size_of::<MessageInfo>();

#[cfg(test)]
thread_local! {
    static MESSAGE_INFO_CONSTRUCTIONS: Cell<usize> = const { Cell::new(0) };
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HeaderNode {
    ArchiveInfo,
    MessageInfo,
    FieldInfo,
    FieldPath,
}

/// Exact resource totals collected while a retained ArchiveInfo header is
/// admitted.  The raw header is immutable private provenance after parsing;
/// retaining these totals lets later callers apply a different `Limits`
/// profile without walking the same wire tree again.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct HeaderAdmission {
    preflight: WirePreflight,
    message_infos: usize,
    decoded_memory: usize,
    metadata_items: usize,
}

/// Immutable source and admission provenance retained for a non-canonical
/// archive header. The object keeps only an `Arc` to this record so clones
/// share the source bytes and their exact admission totals.
#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceProvenance {
    /// Original bytes followed by the canonical bytes in one allocation.
    headers: Box<[u8]>,
    original_len: usize,
    canonical_len: usize,
    header_admission: Option<HeaderAdmission>,
    archive_info_admission: Option<ArchiveInfoAdmission>,
    admitted_message_bytes: Option<usize>,
    admitted_metadata_items: Option<usize>,
}

impl SourceProvenance {
    fn from_headers(
        original: &[u8],
        canonical: &[u8],
        header_admission: Option<HeaderAdmission>,
        archive_info_admission: Option<ArchiveInfoAdmission>,
        limits: Option<Limits>,
    ) -> Result<Arc<Self>> {
        let total_len = original
            .len()
            .checked_add(canonical.len())
            .ok_or_else(|| Error::allocation("IWA retained header provenance", usize::MAX))?;
        let mut headers = Vec::new();
        headers
            .try_reserve_exact(total_len)
            .map_err(|_| Error::allocation("IWA retained header provenance", total_len))?;
        headers.extend_from_slice(original);
        headers.extend_from_slice(canonical);
        Ok(Arc::new(Self {
            headers: headers.into_boxed_slice(),
            original_len: original.len(),
            canonical_len: canonical.len(),
            header_admission,
            archive_info_admission,
            admitted_message_bytes: limits.map(|limits| limits.max_message_bytes()),
            admitted_metadata_items: limits.map(|limits| limits.max_metadata_items()),
        }))
    }

    fn original_header(&self) -> &[u8] {
        &self.headers[..self.original_len]
    }

    fn original_canonical_header(&self) -> &[u8] {
        &self.headers[self.original_len..self.original_len + self.canonical_len]
    }

    #[cfg(test)]
    fn original_header_mut(&mut self) -> &mut [u8] {
        &mut self.headers[..self.original_len]
    }

    #[cfg(test)]
    fn original_canonical_header_mut(&mut self) -> &mut [u8] {
        &mut self.headers[self.original_len..self.original_len + self.canonical_len]
    }
}

fn retained_source_provenance(
    original: &[u8],
    canonical: &[u8],
    retain: bool,
) -> Result<Option<Arc<SourceProvenance>>> {
    retain
        .then(|| SourceProvenance::from_headers(original, canonical, None, None, None))
        .transpose()
}

/// Exact resource totals for the decoded neutral ArchiveInfo projection.
/// These are separate from the wire totals because constructor limits account
/// owned metadata while parser limits also account unknown source fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ArchiveInfoAdmission {
    header_length: usize,
    header_fields: usize,
    header_nesting: usize,
    header_memory: usize,
    metadata_items: usize,
    message_count: usize,
}

/// Admission totals accumulated while a message collection is built.
///
/// The collection owns its certificate only while it controls all mutable
/// access to the contained records. Converting an ordinary `Vec<MessageInfo>`
/// therefore produces an uncertified collection and keeps the raw constructor
/// path fully checked.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MessageInfoCollectionAdmission {
    encoded_messages_length: usize,
    header_fields: usize,
    max_message_nesting: usize,
    header_memory: usize,
    minimal_spill_memory: usize,
    metadata_items: usize,
    message_count: usize,
}

impl MessageInfoCollectionAdmission {
    const MINIMAL_ONLY_MARKER: usize = usize::MAX;

    fn empty() -> Self {
        Self {
            encoded_messages_length: 0,
            header_fields: 0,
            max_message_nesting: Self::MINIMAL_ONLY_MARKER,
            header_memory: 0,
            minimal_spill_memory: 0,
            metadata_items: 0,
            message_count: 0,
        }
    }

    #[inline]
    fn is_minimal_only(self) -> bool {
        self.max_message_nesting == Self::MINIMAL_ONLY_MARKER
    }

    fn materialize_minimal_totals(&mut self) -> Result<()> {
        if !self.is_minimal_only() {
            return Ok(());
        }
        self.header_fields = self
            .message_count
            .checked_mul(4)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        self.header_memory = self
            .message_count
            .checked_mul(MINIMAL_MESSAGE_INFO_MEMORY)
            .and_then(|value| value.checked_add(self.minimal_spill_memory))
            .ok_or_else(|| Error::invalid_archive(0, "neutral metadata memory overflow"))?;
        self.metadata_items = self
            .message_count
            .checked_mul(1 + DEFAULT_MESSAGE_VERSIONS.len())
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
        self.max_message_nesting = 0;
        Ok(())
    }

    fn archive_info_with_options(
        self,
        identifier: Option<u64>,
        should_merge: Option<bool>,
    ) -> Result<ArchiveInfoAdmission> {
        let mut header_length = 0usize;
        if let Some(identifier) = identifier {
            header_length = varint_field_length(1, identifier)?;
        }
        header_length = checked_encoded_add(header_length, self.encoded_messages_length)?;
        if let Some(should_merge) = should_merge {
            header_length = checked_encoded_add(
                header_length,
                varint_field_length(3, u64::from(should_merge))?,
            )?;
        }
        let (header_fields, header_memory, metadata_items, max_message_nesting) = if self
            .is_minimal_only()
        {
            let per_message_fields = self
                .message_count
                .checked_mul(4)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
            let header_memory = self
                .message_count
                .checked_mul(MINIMAL_MESSAGE_INFO_MEMORY)
                .and_then(|value| value.checked_add(self.minimal_spill_memory))
                .ok_or_else(|| Error::invalid_archive(0, "neutral metadata memory overflow"))?;
            let metadata_items = self
                .message_count
                .checked_mul(1 + DEFAULT_MESSAGE_VERSIONS.len())
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
            (per_message_fields, header_memory, metadata_items, 0)
        } else {
            (
                self.header_fields,
                self.header_memory,
                self.metadata_items,
                self.max_message_nesting,
            )
        };
        let mut header_fields = header_fields;
        header_fields = header_fields
            .checked_add(usize::from(identifier.is_some()))
            .and_then(|value| value.checked_add(usize::from(should_merge.is_some())))
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        let header_nesting = if self.message_count == 0 {
            0
        } else {
            1usize
                .checked_add(max_message_nesting)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo nesting overflow"))?
        };
        let header_memory = size_of::<ArchiveInfo>()
            .checked_add(header_memory)
            .ok_or_else(|| Error::invalid_archive(0, "neutral metadata memory overflow"))?;
        Ok(ArchiveInfoAdmission {
            header_length,
            header_fields,
            header_nesting,
            header_memory,
            metadata_items,
            message_count: self.message_count,
        })
    }

    fn archive_info(self, identifier: u64) -> Result<ArchiveInfoAdmission> {
        self.archive_info_with_options(Some(identifier), None)
    }

    /// Add a conventional message under the default profile.
    ///
    /// The fixed shape totals are materialized once, when the collection is
    /// published or when a nonminimal record follows.  While the collection
    /// remains minimal, the only aggregate that still depends on each public
    /// record is its exact encoded header length.  Keep that ledger update
    /// separate from the strict-profile path so the common route does not
    /// repeat fixed memory, field and metadata arithmetic for every push.
    #[inline]
    fn push_minimal_default(
        &mut self,
        type_: u32,
        length: u32,
        version_memory: usize,
        limits: Limits,
    ) -> Result<()> {
        let message_count = self
            .message_count
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo message count overflow"))?;
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }
        let encoded_length = minimal_message_info_field_length(type_, length)?;
        self.encoded_messages_length =
            checked_encoded_add(self.encoded_messages_length, encoded_length)?;
        if self.is_minimal_only() {
            self.minimal_spill_memory =
                checked_memory_add(self.minimal_spill_memory, version_memory)?;
        } else {
            self.header_fields = self
                .header_fields
                .checked_add(4)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
            self.header_memory = checked_memory_add(self.header_memory, version_memory)?;
            self.header_memory =
                checked_memory_add(self.header_memory, MINIMAL_MESSAGE_INFO_MEMORY)?;
            self.metadata_items = self
                .metadata_items
                .checked_add(1 + DEFAULT_MESSAGE_VERSIONS.len())
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
        }
        self.message_count = message_count;
        Ok(())
    }

    #[inline]
    fn push_minimal(
        &mut self,
        type_: u32,
        length: u32,
        version_memory: usize,
        limits: Limits,
    ) -> Result<()> {
        let message_count = self
            .message_count
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo message count overflow"))?;
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }

        let encoded_length = minimal_message_info_encoded_length(type_, length)?;
        let header_memory = checked_memory_add(MINIMAL_MESSAGE_INFO_MEMORY, version_memory)?;
        if !limits.is_default_profile() {
            check_constructor_header(
                encoded_length,
                3,
                0,
                header_memory,
                DEFAULT_MESSAGE_VERSIONS.len(),
                limits,
            )?;
        }
        self.encoded_messages_length = checked_encoded_add(
            self.encoded_messages_length,
            length_delimited_field_length(2, encoded_length)?,
        )?;
        if !self.is_minimal_only() {
            self.header_fields = self
                .header_fields
                .checked_add(4)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
            self.header_memory = checked_memory_add(self.header_memory, header_memory)?;
            self.metadata_items = self
                .metadata_items
                .checked_add(1 + DEFAULT_MESSAGE_VERSIONS.len())
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
        } else {
            self.minimal_spill_memory =
                checked_memory_add(self.minimal_spill_memory, version_memory)?;
        }
        self.message_count = message_count;
        Ok(())
    }

    fn push(&mut self, admission: MessageInfoAdmission, limits: Limits) -> Result<()> {
        let message_count = self
            .message_count
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo message count overflow"))?;
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }
        self.materialize_minimal_totals()?;
        self.encoded_messages_length = checked_encoded_add(
            self.encoded_messages_length,
            length_delimited_field_length(2, admission.encoded_length)?,
        )?;
        self.header_fields = self
            .header_fields
            .checked_add(
                1usize
                    .checked_add(admission.header_fields)
                    .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?,
            )
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        self.max_message_nesting = self.max_message_nesting.max(admission.header_nesting);
        self.header_memory = checked_memory_add(self.header_memory, admission.header_memory)?;
        self.metadata_items = self
            .metadata_items
            .checked_add(1)
            .and_then(|value| value.checked_add(admission.metadata_items))
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
        self.message_count = message_count;
        Ok(())
    }
}

/// Exact admission totals for one decoded or authoring-built MessageInfo.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MessageInfoAdmission {
    encoded_length: usize,
    header_fields: usize,
    header_nesting: usize,
    header_memory: usize,
    metadata_items: usize,
}

impl ArchiveInfoAdmission {
    fn check(self, limits: Limits) -> Result<()> {
        if self.message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                self.message_count,
                limits.max_messages_per_object(),
            ));
        }
        check_constructor_header(
            self.header_length,
            self.header_fields,
            self.header_nesting,
            self.header_memory,
            self.metadata_items,
            limits,
        )
    }
}

impl HeaderAdmission {
    fn check(self, header_length: usize, limits: Limits) -> Result<()> {
        check_header_length(header_length, limits)?;

        let scanned_bytes = limits
            .max_header_bytes()
            .saturating_mul(4)
            .min(WireLimits::MAX_INPUT_BYTES);
        if self.preflight.scanned_bytes() > scanned_bytes {
            return Err(limit(
                LimitKind::HeaderBytes,
                self.preflight.scanned_bytes(),
                scanned_bytes,
            ));
        }
        if self.preflight.fields() > limits.max_header_fields() {
            return Err(limit(
                LimitKind::HeaderFields,
                self.preflight.fields(),
                limits.max_header_fields(),
            ));
        }
        if self.preflight.max_depth() > limits.max_header_nesting() {
            return Err(limit(
                LimitKind::HeaderNesting,
                self.preflight.max_depth(),
                limits.max_header_nesting(),
            ));
        }
        if self.message_infos > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                self.message_infos,
                limits.max_messages_per_object(),
            ));
        }
        if self.decoded_memory > limits.max_header_memory_bytes() {
            return Err(limit(
                LimitKind::HeaderMemoryBytes,
                self.decoded_memory,
                limits.max_header_memory_bytes(),
            ));
        }
        if self.metadata_items > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                self.metadata_items,
                limits.max_metadata_items(),
            ));
        }
        Ok(())
    }
}

/// Schema-neutral path to a field nested inside an archive payload.
///
/// An empty component list is valid on the wire. Required `FieldInfo.path`
/// presence is enforced while decoding, before this value is published.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct FieldPath {
    /// Field numbers from the payload root to the selected field.
    pub path: Vec<u32>,
}

impl FieldPath {
    /// Construct a field path from its ordered field numbers.
    ///
    /// The path is a raw metadata record, so its limits are required even
    /// though the vector is already owned by the caller.  Empty paths and
    /// zero-valued components are valid wire values.
    #[inline]
    pub fn new(path: Vec<u32>, limits: Limits) -> Result<Self> {
        let value = Self { path };
        // The exact admission is out of line and reports only an error, so
        // the admitted record is built once, directly in the caller's slot.
        if limits.is_default_profile() {
            if value.path.len() >= DEFAULT_FIELD_PATH_COMPONENTS {
                admit_field_path_default(&value)?;
            }
        } else {
            admit_field_path(&value, limits)?;
        }
        Ok(value)
    }

    /// Borrow the ordered field numbers.
    #[must_use]
    pub fn as_slice(&self) -> &[u32] {
        &self.path
    }
}

/// Kind of value described by field-level archive metadata.
///
/// `Unrecognized` preserves future or producer-specific closed-enum values
/// exactly instead of coercing them to a known default.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FieldType {
    Value,
    ObjectReference,
    DataReference,
    Message,
    Unrecognized(i32),
}

/// Data-reference removal policy for a header-preserving message replacement.
///
/// `Selected` removes every occurrence of the supplied identifiers from the
/// target `MessageInfo` and all of its nested `FieldInfo` values. `All` clears
/// every data reference in that same scope. Duplicate selected identifiers are
/// harmless, and `Selected(&[])` is equivalent to `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataReferencePruning<'a> {
    /// Preserve all data references.
    None,
    /// Remove every occurrence of these data-reference identifiers.
    Selected(&'a [u64]),
    /// Remove every data-reference identifier.
    All,
}

/// Reference kind carried by physical `ArchiveInfo` metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ArchiveReferenceKind {
    /// An archive-object identifier.
    Object,
    /// A package data identifier.
    Data,
}

/// Location of one reference occurrence inside an archive object's metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ArchiveReferenceScope {
    /// The aggregate list on one `MessageInfo`.
    Message,
    /// A list on one nested `FieldInfo`.
    Field { field_index: usize },
}

/// One source-authoritative physical metadata reference occurrence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ArchiveReferenceOccurrence {
    pub object_identifier: u64,
    pub message_index: usize,
    pub scope: ArchiveReferenceScope,
    pub kind: ArchiveReferenceKind,
    pub referenced_identifier: u64,
}

/// Visitor for bounded archive-reference inspection.
pub trait ArchiveReferenceVisitor {
    fn visit_reference(&mut self, occurrence: ArchiveReferenceOccurrence) -> Result<()>;
}

/// Completeness policy for an archive-reference census.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArchiveReferencePolicy {
    /// Report every known reference while retaining unknown metadata.
    KnownReferences,
    /// Refuse any unknown `ArchiveInfo`, `MessageInfo`, or `FieldInfo` field.
    ///
    /// This is the deletion-grade policy: an opaque future field could carry
    /// an owner edge that the current schema cannot attribute.
    RejectUnknownMetadata,
}

impl DataReferencePruning<'_> {
    const fn is_none(self) -> bool {
        matches!(self, Self::None | Self::Selected([]))
    }
}

/// Exact before-to-after object-reference state for one nested `FieldInfo`.
///
/// The ordinal and path jointly authorize the target occurrence. Both lists
/// are complete decoded lists, not edit fragments.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FieldObjectReferenceTransition<'a> {
    /// Zero-based `FieldInfo` occurrence in the target `MessageInfo`.
    pub field_info_index: usize,
    /// Exact decoded path required at `field_info_index`.
    pub expected_path: &'a [u32],
    /// Complete source object-reference list.
    pub before: &'a [u64],
    /// Complete requested object-reference list.
    pub after: &'a [u64],
}

/// Exact aggregate and selected-field object-reference state transition.
///
/// All lists must contain unique nonzero identifiers. Each `after` list is a
/// permutation of identifiers retained from its corresponding `before` list,
/// followed by identifiers newly introduced to that list. Field-local values
/// must be members of the aggregate list in the same direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ObjectReferenceTransition<'a> {
    /// Complete source `MessageInfo.object_references` list.
    pub aggregate_before: &'a [u64],
    /// Complete requested `MessageInfo.object_references` list.
    pub aggregate_after: &'a [u64],
    /// Exact selected nested `FieldInfo` transitions.
    pub fields: &'a [FieldObjectReferenceTransition<'a>],
}

/// Minimal field-local attribution of object references.
///
/// Paths in a batch must be distinct. References must be nonempty, unique,
/// nonzero members of the corresponding aggregate transition state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CanonicalObjectReferenceField<'a> {
    /// Complete semantic field path.
    pub path: &'a [u32],
    /// Complete object-reference list for this field.
    pub references: &'a [u64],
}

/// Append canonical object-reference fields or remove their exact suffix.
///
/// Insertion requires every path to be absent. Removal requires the final
/// `FieldInfo` occurrences to match the supplied order and the exact minimal
/// wire encoding emitted by insertion, including type and reference presence.
/// Unknown or noncanonical fields cannot be discarded through this operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CanonicalObjectReferenceFields<'a> {
    /// Append the fields after all existing `FieldInfo` occurrences.
    Insert(&'a [CanonicalObjectReferenceField<'a>]),
    /// Remove the exact canonical suffix, retaining every preceding field.
    Remove(&'a [CanonicalObjectReferenceField<'a>]),
}

impl<'a> CanonicalObjectReferenceFields<'a> {
    const fn definitions(self) -> &'a [CanonicalObjectReferenceField<'a>] {
        match self {
            Self::Insert(fields) | Self::Remove(fields) => fields,
        }
    }
}

/// Exact before-to-after data-reference state for one nested `FieldInfo`.
///
/// The ordinal and path jointly authorize the target occurrence. Unlike the
/// historical permutation helper, both lists may have different cardinality;
/// this is needed by media owners when an item is inserted, replaced, or
/// removed. The caller still owns the format-level proof that the selected
/// field is the semantic owner of the aggregate list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FieldDataReferenceTransition<'a> {
    /// Zero-based `FieldInfo` occurrence in the target `MessageInfo`.
    pub field_info_index: usize,
    /// Exact decoded path required at `field_info_index`.
    pub expected_path: &'a [u32],
    /// Complete source data-reference list.
    pub before: &'a [u64],
    /// Complete requested data-reference list.
    pub after: &'a [u64],
}

/// Exact aggregate and selected-field data-reference state transition.
///
/// Lists are positional and may change cardinality. Every value is required
/// to be nonzero. Unknown and unrelated archive metadata remains raw-byte
/// authoritative; only the aggregate field and explicitly selected
/// `FieldInfo` occurrences are rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataReferenceTransition<'a> {
    /// Complete source `MessageInfo.data_references` list.
    pub aggregate_before: &'a [u64],
    /// Complete requested `MessageInfo.data_references` list.
    pub aggregate_after: &'a [u64],
    /// Exact selected nested `FieldInfo` transitions.
    pub fields: &'a [FieldDataReferenceTransition<'a>],
}

/// Direction of a canonical field-level data-reference metadata change.
///
/// This is intentionally separate from [`DataReferenceTransition`].  The
/// ordinary transition edits existing `FieldInfo` occurrences addressed by
/// their source ordinal.  A format owner may instead need to materialize the
/// first field-local attribution for a legacy header that only carries the
/// aggregate `MessageInfo.data_references` list.  Such a field is created or
/// removed as one canonical record, never guessed from an arbitrary existing
/// ordinal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CanonicalFieldDataReferenceOperation {
    /// Append one canonical `FieldInfo` carrying the target aggregate list.
    Insert,
    /// Remove the previously materialized canonical `FieldInfo`.
    Remove,
}

/// Exact aggregate transition that creates or removes one canonical
/// data-reference `FieldInfo`.
///
/// `field_path` is the caller-supplied semantic path.  On insertion the
/// source must not already contain that path and the canonical record is
/// appended after every source `FieldInfo`; its type is
/// [`FieldType::DataReference`] and its data references equal
/// `aggregate_after`.  On removal the source record must be the final
/// `FieldInfo`, have exactly the canonical bytes produced by insertion, and
/// carry `aggregate_before`; `aggregate_after` may contain any references
/// retained by the aggregate transition.  These strict directional rules
/// make `inverse()`-style physical replay byte-exact even when the containing
/// header has unknown or noncanonical bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CanonicalFieldDataReferenceTransition<'a> {
    /// Complete source `MessageInfo.data_references` list.
    pub aggregate_before: &'a [u64],
    /// Complete requested `MessageInfo.data_references` list.
    pub aggregate_after: &'a [u64],
    /// Exact path to materialize or remove.
    pub field_path: &'a [u32],
    /// Whether to append or remove the canonical field record.
    pub operation: CanonicalFieldDataReferenceOperation,
}

impl FieldType {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::Value,
            1 => Self::ObjectReference,
            2 => Self::DataReference,
            3 => Self::Message,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::Value => 0,
            Self::ObjectReference => 1,
            Self::DataReference => 2,
            Self::Message => 3,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Preservation policy for an unknown payload field.
///
/// `Unrecognized` retains future closed-enum values. In particular,
/// `NotSupported` is the known native value `-1`, not an unknown value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum UnknownFieldRule {
    IgnoreAndPreserveUntilModified,
    IgnoreAndPreserve,
    MustUnderstand,
    NotSupported,
    Unrecognized(i32),
}

impl UnknownFieldRule {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::IgnoreAndPreserveUntilModified,
            1 => Self::IgnoreAndPreserve,
            2 => Self::MustUnderstand,
            -1 => Self::NotSupported,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::IgnoreAndPreserveUntilModified => 0,
            Self::IgnoreAndPreserve => 1,
            Self::MustUnderstand => 2,
            Self::NotSupported => -1,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Preservation policy for a recognized payload field introduced by a newer
/// producer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum KnownFieldRule {
    None,
    PreserveNewerValueUntilModified,
    PreserveNewerValue,
    Unrecognized(i32),
}

impl KnownFieldRule {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::None,
            1 => Self::PreserveNewerValueUntilModified,
            2 => Self::PreserveNewerValue,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::PreserveNewerValueUntilModified => 1,
            Self::PreserveNewerValue => 2,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Field-level reference and compatibility metadata for one archive payload.
///
/// Optional enum fields retain presence separately from their schema default;
/// this is required for canonical-byte parity after a semantic edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldInfo {
    pub path: FieldPath,
    pub r#type: Option<FieldType>,
    pub unknown_field_rule: Option<UnknownFieldRule>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub known_field_rule: Option<KnownFieldRule>,
    pub known_field_version: Vec<u32>,
    pub known_field_feature_identifier: Option<String>,
}

impl FieldInfo {
    /// Construct metadata for a field path with every optional wire field
    /// absent.
    #[inline]
    pub fn new(path: FieldPath, limits: Limits) -> Result<Self> {
        // As in `FieldPath::new`, only the error leaves the cold path.
        if limits.is_default_profile() {
            if path.path.len() >= DEFAULT_FIELD_PATH_COMPONENTS {
                admit_minimal_field_info_default(&path)?;
            }
        } else {
            admit_minimal_field_info(&path, limits)?;
        }
        Ok(Self::minimal(path))
    }

    #[inline(always)]
    fn minimal(path: FieldPath) -> Self {
        Self {
            path,
            ..Self::default()
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_type(&self) -> FieldType {
        match self.r#type {
            Some(value) => value,
            None => FieldType::Value,
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_unknown_field_rule(&self) -> UnknownFieldRule {
        match self.unknown_field_rule {
            Some(value) => value,
            None => UnknownFieldRule::IgnoreAndPreserveUntilModified,
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_known_field_rule(&self) -> KnownFieldRule {
        match self.known_field_rule {
            Some(value) => value,
            None => KnownFieldRule::None,
        }
    }
}

/// Metadata for one object in an IWA component.
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveInfo {
    /// Unique identifier across the document package.
    pub identifier: Option<u64>,
    /// Metadata for each payload immediately following this header.
    pub message_infos: Vec<MessageInfo>,
    /// Whether this archive should be merged with an existing object.
    pub should_merge: Option<bool>,
}

impl ArchiveInfo {
    /// Decode one bounded `TSP.ArchiveInfo` protobuf from a reader.
    ///
    /// Because protobuf messages are not self-delimiting, callers must pass a
    /// reader bounded to exactly one header when additional bytes follow it.
    pub fn parse<R: Read>(reader: &mut R, limits: Limits) -> Result<Self> {
        let data = read_bounded(reader, limits.max_header_bytes(), "ArchiveInfo")?;
        Self::decode(&data, limits)
    }

    /// Decode one `TSP.ArchiveInfo` header under explicit limits.
    pub fn decode(data: &[u8], limits: Limits) -> Result<Self> {
        check_header_length(data.len(), limits)?;
        let preflight = preflight_header(data, HeaderKind::ArchiveInfo, limits)?;
        Self::decode_with_preflight(data, limits, preflight)
    }

    fn decode_with_preflight(
        data: &[u8],
        limits: Limits,
        preflight: WirePreflight,
    ) -> Result<Self> {
        Self::decode_with_preflight_and_admission(data, limits, preflight)
            .map(|(archive_info, _admission)| archive_info)
    }

    fn decode_with_preflight_and_admission(
        data: &[u8],
        limits: Limits,
        preflight: WirePreflight,
    ) -> Result<(Self, ArchiveInfoAdmission)> {
        let decoded =
            archive_codec::decode_archive_info(data, buffa_decode_options(preflight, limits))
                .map_err(|error| {
                    Error::header_codec(
                        HeaderKind::ArchiveInfo,
                        HeaderOperation::Decode,
                        error.to_string(),
                    )
                })?;
        let (archive_info, admission) = archive_info_from_codec_with_admission(decoded, limits)?;
        Ok((archive_info, admission))
    }

    /// Construct metadata for an object from its raw payload metadata.
    ///
    /// An identifier of zero is preserved; identifier policy remains a later
    /// archive validation concern. The supplied limits admit the complete
    /// neutral header before this record is published. A
    /// [`MessageInfoCollection`] carries totals from its bounded construction
    /// path, while an ordinary `Vec<MessageInfo>` remains a fully checked raw
    /// input.
    #[inline]
    pub fn new(
        identifier: u64,
        message_infos: impl Into<MessageInfoCollection>,
        limits: Limits,
    ) -> Result<Self> {
        let message_infos = message_infos.into();
        if !limits.is_default_profile() {
            return Self::from_collection(identifier, message_infos, limits);
        }
        if match (&message_infos.admission, message_infos.limits) {
            // Every record entered through a minimal default-profile
            // push: its payload length was checked and the message count
            // stayed within the per-object ceiling, so the compile-time
            // bounds of `DEFAULT_SHAPE_MESSAGE_ENTRY_BYTES` cover the
            // whole header.
            (Some(admission), CollectionProfile::Default) => admission.is_minimal_only(),
            // An uncertified collection (an ordinary `Vec`) is checked
            // record by record.
            (None, _) => default_archive_info_bounds_are_safe(&message_infos.message_infos),
            (Some(_), _) => false,
        } {
            return Ok(Self {
                identifier: Some(identifier),
                message_infos: message_infos.message_infos,
                should_merge: None,
            });
        }
        // Rebuild the default profile inside the out-of-line path so a
        // caller with a constant profile does not materialize it per call.
        Self::from_collection_default(identifier, message_infos)
    }

    #[inline(never)]
    fn from_collection_default(
        identifier: u64,
        message_infos: MessageInfoCollection,
    ) -> Result<Self> {
        Self::from_collection(identifier, message_infos, Limits::default())
    }

    #[inline(never)]
    fn from_collection(
        identifier: u64,
        message_infos: MessageInfoCollection,
        limits: Limits,
    ) -> Result<Self> {
        let (message_infos, collection_admission, collection_limits) = message_infos.into_parts();
        let value = Self {
            identifier: Some(identifier),
            message_infos,
            should_merge: None,
        };
        if let (Some(collection_admission), Some(collection_limits)) =
            (collection_admission, collection_limits)
            && same_profile(collection_limits, limits)
        {
            let admission = collection_admission.archive_info(identifier)?;
            admission.check(limits)?;
            return Ok(value);
        }
        // `new` already answered a default-profile uncertified collection
        // whose records all have the certified shape. Strict profiles and
        // other shapes retain the exact admission path.
        check_archive_info_constructor(&value, limits)?;
        Ok(value)
    }

    fn validate(&self, limits: Limits) -> Result<()> {
        check_archive_info_constructor(self, limits).map(|_| ())
    }
}

/// Whether two limit profiles are equal, answering the common default case
/// with one bit test per side: an unmarked profile equals `Limits::default()`.
#[inline(always)]
fn same_profile(left: Limits, right: Limits) -> bool {
    (left.is_default_profile() && right.is_default_profile()) || left == right
}

/// Upper bound of one encoded `ArchiveInfo.message_infos` entry whose only
/// populated fields are `type`, at most three inline `versions` values and
/// `length`: entry tag and length prefix (2), type (1 + 5), packed versions
/// (1 + 1 + 3 * 5) and length (1 + 5).
const DEFAULT_SHAPE_MESSAGE_ENTRY_BYTES: usize = 2 + 6 + 17 + 6;

// The default-profile certificate below is sound only while these products
// stay inside the default ceilings. Changing a ceiling, the default profile
// or the record layout must keep them true (they are checked at compile time).
const _: () = {
    let messages = Limits::MAX_MESSAGES_PER_OBJECT;
    // Identifier field (11 bytes) plus every message entry.
    assert!(11 + messages * DEFAULT_SHAPE_MESSAGE_ENTRY_BYTES <= Limits::MAX_HEADER_BYTES);
    // Identifier plus one entry, type, packed versions and length per message.
    assert!(messages * 4 < Limits::MAX_HEADER_FIELDS);
    // One message record plus three version values per message.
    assert!(messages * 4 <= Limits::MAX_METADATA_ITEMS);
    // Inline versions own no heap memory.
    assert!(
        size_of::<ArchiveInfo>() + messages * size_of::<MessageInfo>()
            <= Limits::DEFAULT_HEADER_MEMORY_BYTES
    );
    // ArchiveInfo -> MessageInfo.
    assert!(1 <= Limits::DEFAULT_HEADER_NESTING);
};

/// Constant-work-per-record admission certificate for the default profile.
///
/// A record qualifies when every repeated or optional field other than
/// `type`, `versions` and `length` is empty, `versions` is stored inline
/// (at most three values, no heap allocation) and `length` is within the
/// message-byte ceiling. The admitted resources of such a record do not
/// depend on its version or type values, and the compile-time assertions
/// above prove that any number of qualifying records up to the per-object
/// ceiling fits every default budget. The check reads each public field
/// once and combines them without a branch per field. Any record outside
/// this shape, and every non-default profile, takes the exact admission
/// path with its original error precedence.
#[inline]
fn default_archive_info_bounds_are_safe(message_infos: &[MessageInfo]) -> bool {
    if message_infos.len() > Limits::MAX_MESSAGES_PER_OBJECT {
        return false;
    }
    let mut outside = 0usize;
    for message_info in message_infos {
        outside |= message_info.field_infos.len()
            | message_info.object_references.len()
            | message_info.data_references.len()
            | usize::from(message_info.base_message_index.is_some())
            | message_info.diff_merge_version.len()
            | usize::from(message_info.diff_field_path.is_some())
            | message_info.fields_to_remove.len()
            | message_info.diff_read_version.len()
            | usize::from(message_info.versions.spilled())
            | usize::from(u64::from(message_info.length) > Limits::MAX_MESSAGE_BYTES as u64);
    }
    outside == 0
}

/// Metadata for one payload in an archive object.
#[derive(Debug, Clone, PartialEq)]
pub struct MessageInfo {
    pub type_: u32,
    /// Version metadata stored inline for the conventional three-value shape,
    /// with a heap-backed fallback for arbitrary sequences.
    pub versions: SmallVec<[u32; MESSAGE_VERSIONS_INLINE_CAPACITY]>,
    pub length: u32,
    pub field_infos: Vec<FieldInfo>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub base_message_index: Option<u32>,
    pub diff_merge_version: Vec<u32>,
    pub diff_field_path: Option<FieldPath>,
    pub fields_to_remove: Vec<FieldPath>,
    pub diff_read_version: Vec<u32>,
}

impl MessageInfo {
    /// Construct the conventional metadata for a newly created payload.
    ///
    /// `length` is checked against the message-byte budget before it is
    /// narrowed to the wire `u32` representation. The conventional version
    /// sequence is stored inline after the header and metadata budgets are
    /// admitted.
    #[inline]
    pub fn new(type_: u32, length: usize, limits: Limits) -> Result<Self> {
        let default_limits = limits.is_default_profile();
        check_message_length(length, limits)?;
        let length = u32::try_from(length)
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        if !default_limits {
            check_minimal_message_info_constructor(type_, length, limits)?;
        }
        let versions = SmallVec::from_buf(DEFAULT_MESSAGE_VERSIONS);
        Ok(Self {
            type_,
            versions,
            length,
            field_infos: Vec::new(),
            object_references: Vec::new(),
            data_references: Vec::new(),
            base_message_index: None,
            diff_merge_version: Vec::new(),
            diff_field_path: None,
            fields_to_remove: Vec::new(),
            diff_read_version: Vec::new(),
        })
    }

    /// Build the conventional record after its admission has been decided.
    ///
    /// Infallible and always inlined, so collection and object constructors
    /// write the record directly into its destination slot.
    #[inline(always)]
    fn conventional(type_: u32, length: u32) -> Self {
        #[cfg(test)]
        MESSAGE_INFO_CONSTRUCTIONS.with(|count| {
            count.set(count.get().saturating_add(1));
        });
        Self {
            type_,
            versions: SmallVec::from_buf(DEFAULT_MESSAGE_VERSIONS),
            length,
            field_infos: Vec::new(),
            object_references: Vec::new(),
            data_references: Vec::new(),
            base_message_index: None,
            diff_merge_version: Vec::new(),
            diff_field_path: None,
            fields_to_remove: Vec::new(),
            diff_read_version: Vec::new(),
        }
    }

    /// Decode one `TSP.MessageInfo` protobuf under explicit limits.
    pub fn decode(data: &[u8], limits: Limits) -> Result<Self> {
        check_header_length(data.len(), limits)?;
        let preflight = preflight_header(data, HeaderKind::MessageInfo, limits)?;
        let decoded =
            archive_codec::decode_message_info(data, buffa_decode_options(preflight, limits))
                .map_err(|error| {
                    Error::header_codec(
                        HeaderKind::MessageInfo,
                        HeaderOperation::Decode,
                        error.to_string(),
                    )
                })?;
        let message_info = message_info_from_codec(decoded)?;
        message_info.validate(limits)?;
        Ok(message_info)
    }

    /// Decode one bounded `TSP.MessageInfo` protobuf from a reader.
    ///
    /// Because protobuf messages are not self-delimiting, callers must pass a
    /// reader bounded to exactly one metadata message when additional bytes
    /// follow it.
    pub fn parse<R: Read>(reader: &mut R, limits: Limits) -> Result<Self> {
        let data = read_bounded(reader, limits.max_header_bytes(), "MessageInfo")?;
        Self::decode(&data, limits)
    }

    fn validate(&self, limits: Limits) -> Result<()> {
        let length = usize::try_from(self.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        if length > limits.max_message_bytes() {
            return Err(limit(
                LimitKind::MessageBytes,
                length,
                limits.max_message_bytes(),
            ));
        }
        let metadata_items = self.metadata_item_count()?;
        if metadata_items > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                metadata_items,
                limits.max_metadata_items(),
            ));
        }
        Ok(())
    }

    fn metadata_item_count(&self) -> Result<usize> {
        let mut count = 0;
        add_count(&mut count, self.versions.len())?;
        add_count(&mut count, self.field_infos.len())?;
        add_count(&mut count, self.object_references.len())?;
        add_count(&mut count, self.data_references.len())?;
        add_count(&mut count, self.diff_merge_version.len())?;
        add_count(&mut count, usize::from(self.diff_field_path.is_some()))?;
        add_count(&mut count, self.fields_to_remove.len())?;
        add_count(&mut count, self.diff_read_version.len())?;
        for field_info in &self.field_infos {
            add_count(&mut count, 1)?;
            add_count(&mut count, field_info.path.path.len())?;
            add_count(&mut count, field_info.object_references.len())?;
            add_count(&mut count, field_info.data_references.len())?;
            add_count(&mut count, field_info.known_field_version.len())?;
            add_count(
                &mut count,
                usize::from(field_info.known_field_feature_identifier.is_some()),
            )?;
        }
        if let Some(path) = &self.diff_field_path {
            add_count(&mut count, path.path.len())?;
        }
        for path in &self.fields_to_remove {
            add_count(&mut count, path.path.len())?;
        }
        Ok(count)
    }
}

/// A message metadata collection that retains admission work performed while
/// its contents are assembled.
///
/// `ArchiveInfo::new` consumes this collection and can publish the aggregate
/// record without scanning every message again. `as_mut_slice`, `get_mut`,
/// and the structural mutation methods invalidate the certificate before
/// exposing mutable state, so callers retain the same public mutation
/// capability as the raw `Vec<MessageInfo>` constructor path.
#[derive(Debug, Clone)]
pub struct MessageInfoCollection {
    message_infos: Vec<MessageInfo>,
    admission: Option<MessageInfoCollectionAdmission>,
    limits: CollectionProfile,
}

/// The admission profile a collection was started with.
///
/// The default profile is stored as a unit variant: per-record admission
/// then reads constant ceilings instead of loading a stored `Limits` value,
/// and starting a collection does not copy the profile.
#[derive(Debug, Clone, Copy)]
enum CollectionProfile {
    /// An ordinary `Vec<MessageInfo>` converted without a profile.
    Unbounded,
    Default,
    Custom(Limits),
}

impl CollectionProfile {
    #[inline(always)]
    fn new(limits: Limits) -> Self {
        if limits.is_default_profile() {
            Self::Default
        } else {
            Self::Custom(limits)
        }
    }

    #[inline(always)]
    fn limits(self) -> Option<Limits> {
        match self {
            Self::Unbounded => None,
            Self::Default => Some(Limits::default()),
            Self::Custom(limits) => Some(limits),
        }
    }
}

impl From<Vec<MessageInfo>> for MessageInfoCollection {
    fn from(message_infos: Vec<MessageInfo>) -> Self {
        Self {
            message_infos,
            admission: None,
            limits: CollectionProfile::Unbounded,
        }
    }
}

impl MessageInfoCollection {
    /// Start a bounded message collection.
    pub fn new(limits: Limits) -> Result<Self> {
        Self::with_capacity(0, limits)
    }

    /// Start a bounded message collection with reserved element capacity.
    pub fn with_capacity(capacity: usize, limits: Limits) -> Result<Self> {
        if capacity > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                capacity,
                limits.max_messages_per_object(),
            ));
        }
        let mut message_infos = Vec::new();
        message_infos
            .try_reserve_exact(capacity)
            .map_err(|_| Error::allocation("IWA message metadata", capacity))?;
        Ok(Self {
            message_infos,
            admission: Some(MessageInfoCollectionAdmission::empty()),
            limits: CollectionProfile::new(limits),
        })
    }

    /// Add an already assembled message and admit its complete neutral shape.
    pub fn push(&mut self, message_info: MessageInfo) -> Result<()> {
        let limits = self
            .limits
            .limits()
            .ok_or_else(|| Error::invalid_archive(0, "message collection has no limits"))?;
        let message_count = self
            .message_infos
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo message count overflow"))?;
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }
        if is_minimal_message_info(&message_info) {
            // `MessageInfo::new` admits this conventional shape before it
            // publishes the record. Recompute the current public length and
            // use the same fixed-shape totals only after checking the exact
            // current shape above; a mutation of any public nested field
            // falls through to the general admission path.
            let length = usize::try_from(message_info.length)
                .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
            check_message_length(length, limits)?;
            if let Some(collection_admission) = self.admission.as_mut() {
                let version_memory = message_versions_memory(&message_info.versions)?;
                if limits.is_default_profile() {
                    collection_admission.push_minimal_default(
                        message_info.type_,
                        message_info.length,
                        version_memory,
                        limits,
                    )?;
                } else {
                    collection_admission.push_minimal(
                        message_info.type_,
                        message_info.length,
                        version_memory,
                        limits,
                    )?;
                }
            }
        } else {
            let admission = message_info_admission(&message_info, limits)?;
            if let Some(collection_admission) = self.admission.as_mut() {
                collection_admission.push(admission, limits)?;
            }
        }
        self.message_infos.push(message_info);
        Ok(())
    }

    /// Add conventional metadata while accumulating its admission totals.
    ///
    /// This is the single-pass authoring route for the common
    /// `MessageInfo::new` shape: the fixed version sequence stays inline, and
    /// the collection certificate records the exact aggregate totals.
    pub fn push_new(&mut self, type_: u32, length: usize) -> Result<()> {
        let limits = self
            .limits
            .limits()
            .ok_or_else(|| Error::invalid_archive(0, "message collection has no limits"))?;
        let message_count = self
            .message_infos
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo message count overflow"))?;
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }
        check_message_length(length, limits)?;
        let length = u32::try_from(length)
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        if let Some(collection_admission) = self.admission.as_mut() {
            // Conventional versions are inline and own no heap memory.
            if limits.is_default_profile() {
                collection_admission.push_minimal_default(type_, length, 0, limits)?;
            } else {
                collection_admission.push_minimal(type_, length, 0, limits)?;
            }
        }
        self.message_infos
            .push(MessageInfo::conventional(type_, length));
        Ok(())
    }

    /// Borrow the collection without invalidating its admission certificate.
    #[must_use]
    pub fn as_slice(&self) -> &[MessageInfo] {
        &self.message_infos
    }

    /// Borrow one message without invalidating its admission certificate.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&MessageInfo> {
        self.message_infos.get(index)
    }

    /// Borrow the collection mutably and discard its admission certificate.
    pub fn as_mut_slice(&mut self) -> &mut [MessageInfo] {
        self.admission = None;
        &mut self.message_infos
    }

    /// Borrow one message mutably and discard its admission certificate.
    pub fn get_mut(&mut self, index: usize) -> Option<&mut MessageInfo> {
        self.admission = None;
        self.message_infos.get_mut(index)
    }

    /// Remove one message and discard the certificate.
    pub fn remove(&mut self, index: usize) -> MessageInfo {
        self.admission = None;
        self.message_infos.remove(index)
    }

    /// Remove every message and restart an empty certified collection when its
    /// original bounded profile is still available.
    pub fn clear(&mut self) {
        self.message_infos.clear();
        self.admission = self
            .limits
            .limits()
            .map(|_| MessageInfoCollectionAdmission::empty());
    }

    /// Return the number of messages currently held.
    #[must_use]
    pub fn len(&self) -> usize {
        self.message_infos.len()
    }

    /// Return whether no messages are currently held.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.message_infos.is_empty()
    }

    /// Iterate over messages without invalidating the certificate.
    pub fn iter(&self) -> std::slice::Iter<'_, MessageInfo> {
        self.message_infos.iter()
    }

    fn into_parts(
        self,
    ) -> (
        Vec<MessageInfo>,
        Option<MessageInfoCollectionAdmission>,
        Option<Limits>,
    ) {
        (self.message_infos, self.admission, self.limits.limits())
    }
}

/// Check the bounded resources represented by a neutral record before it is
/// published. These checks account the owned neutral model, not the private
/// decoder's compatibility scratch; parser ingress continues to use
/// `preflight_header` for that separate accounting.
///
/// Collection and string charges are logical element bytes (`len()`), as in
/// [`ArchiveIngressRequirements`]. A caller can hand a constructor a `Vec`
/// with allocator slack, but that pre-existing capacity is neither a wire
/// resource nor an allocation performed by this constructor. Charging it
/// would make the raw record contract depend on allocator growth policy and
/// reject records that the bounded parser admits under the same profile.
fn check_constructor_header(
    header_length: usize,
    header_fields: usize,
    header_nesting: usize,
    header_memory: usize,
    metadata_items: usize,
    limits: Limits,
) -> Result<()> {
    check_header_length(header_length, limits)?;
    if header_fields > limits.max_header_fields() {
        return Err(limit(
            LimitKind::HeaderFields,
            header_fields,
            limits.max_header_fields(),
        ));
    }
    if header_nesting > limits.max_header_nesting() {
        return Err(limit(
            LimitKind::HeaderNesting,
            header_nesting,
            limits.max_header_nesting(),
        ));
    }
    if header_memory > limits.max_header_memory_bytes() {
        return Err(limit(
            LimitKind::HeaderMemoryBytes,
            header_memory,
            limits.max_header_memory_bytes(),
        ));
    }
    if metadata_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            metadata_items,
            limits.max_metadata_items(),
        ));
    }
    Ok(())
}

/// Constant-time admission certificate for the default profile.
///
/// Under the default profile, a path with fewer than this many components
/// (one metadata item is reserved for an enclosing minimal `FieldInfo`)
/// satisfies every derived budget, as the compile-time assertions below
/// prove: its packed components, length prefix and wrapper fit the header
/// byte budget, its neutral memory fits the memory budget, and it adds one
/// field and at most one nesting level. Every other input, and every
/// non-default profile, takes the exact constructor admission with its
/// original error precedence.
const DEFAULT_FIELD_PATH_COMPONENTS: usize = Limits::MAX_METADATA_ITEMS;

const _: () = {
    let components = DEFAULT_FIELD_PATH_COMPONENTS;
    // Five varint bytes per u32 component, a ten-byte length prefix and a
    // one-byte tag, wrapped once more in a FieldInfo entry.
    assert!(components * 5 + 11 + 11 <= Limits::MAX_HEADER_BYTES);
    assert!(size_of::<FieldInfo>() + components * 4 <= Limits::DEFAULT_HEADER_MEMORY_BYTES);
    assert!(2 <= Limits::MAX_HEADER_FIELDS);
    assert!(1 <= Limits::DEFAULT_HEADER_NESTING);
};

// Exact constructor admission outside the constant-time default
// certificate. The default profile is rebuilt inside the cold path so a
// caller with a constant profile does not materialize it per call.
#[cold]
#[inline(never)]
fn admit_field_path_default(path: &FieldPath) -> Result<()> {
    check_field_path_constructor(path, Limits::default())
}

#[cold]
#[inline(never)]
fn admit_field_path(path: &FieldPath, limits: Limits) -> Result<()> {
    check_field_path_constructor(path, limits)
}

#[cold]
#[inline(never)]
fn admit_minimal_field_info_default(path: &FieldPath) -> Result<()> {
    check_minimal_field_info_constructor(path, Limits::default())
}

#[cold]
#[inline(never)]
fn admit_minimal_field_info(path: &FieldPath, limits: Limits) -> Result<()> {
    check_minimal_field_info_constructor(path, limits)
}

fn check_field_path_constructor(path: &FieldPath, limits: Limits) -> Result<()> {
    let metadata_items = path.path.len();
    if limits.is_default_profile() && metadata_items <= Limits::MAX_METADATA_ITEMS {
        // See `DEFAULT_FIELD_PATH_COMPONENTS`; the bound also holds at
        // exactly the metadata ceiling for a standalone path.
        return Ok(());
    }
    let header_length = admitted_path_header_length(path, false, limits)?;
    let header_fields = usize::from(!path.path.is_empty());
    let header_memory = field_path_memory(path)?;
    check_constructor_header(
        header_length,
        header_fields,
        0,
        header_memory,
        metadata_items,
        limits,
    )
}

// A u32 component occupies at most five varint bytes. When that upper
// bound fits, admission does not need to inspect the components. Tight
// ceilings take the exact path, preserving both acceptance and error counts.
#[inline]
fn admitted_path_header_length(path: &FieldPath, nested: bool, limits: Limits) -> Result<usize> {
    let upper = if path.path.is_empty() {
        Some(0)
    } else {
        path.path
            .len()
            .checked_mul(5)
            .and_then(|payload| payload.checked_add(1 + u64_varint_length(payload as u64)))
    };
    let upper = upper.and_then(|length| {
        if nested {
            length.checked_add(1 + u64_varint_length(length as u64))
        } else {
            Some(length)
        }
    });
    if let Some(upper) = upper
        && upper <= limits.max_header_bytes()
    {
        return Ok(upper);
    }
    let exact = encoded_field_path_length(path)?;
    if nested {
        length_delimited_field_length(1, exact)
    } else {
        Ok(exact)
    }
}

fn check_minimal_field_info_constructor(path: &FieldPath, limits: Limits) -> Result<()> {
    let metadata_items = path
        .path
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
    if limits.is_default_profile() && metadata_items <= Limits::MAX_METADATA_ITEMS {
        // The nested FieldPath adds only one fixed metadata item and one
        // length-delimited wrapper; the same default ceilings proved above
        // therefore cover the minimal FieldInfo shape as well.
        return Ok(());
    }
    let header_length = admitted_path_header_length(path, true, limits)?;
    let header_fields = 1 + usize::from(!path.path.is_empty());
    let header_memory = checked_memory_add(size_of::<FieldInfo>(), field_path_bytes(path)?)?;
    check_constructor_header(
        header_length,
        header_fields,
        1,
        header_memory,
        metadata_items,
        limits,
    )
}

fn check_minimal_message_info_constructor(type_: u32, length: u32, limits: Limits) -> Result<()> {
    let header_length = minimal_message_info_encoded_length(type_, length)?;
    let header_memory = minimal_message_info_memory()?;
    check_constructor_header(
        header_length,
        3,
        0,
        header_memory,
        DEFAULT_MESSAGE_VERSIONS.len(),
        limits,
    )
}

/// Preflight the conventional metadata that [`ArchiveObject::new`] will
/// construct from raw payloads.
///
/// The caller owns the payload vector already, but the message metadata is a
/// new allocation. Keep the complete neutral
/// header admission before reserving either collection.  The loop also
/// computes the canonical header length, so the successful constructor does
/// not need to walk the freshly allocated metadata a second time.
fn preflight_archive_object_constructor(
    identifier: u64,
    messages: &[RawMessage],
    limits: Limits,
) -> Result<usize> {
    if messages.len() > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            messages.len(),
            limits.max_messages_per_object(),
        ));
    }

    let mut metadata_items = messages.len();
    let mut header_fields = 1usize
        .checked_add(messages.len())
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    let header_nesting = usize::from(!messages.is_empty());
    let mut header_memory = size_of::<ArchiveInfo>();
    let mut header_length = varint_field_length(1, identifier)?;
    let message_fields = 2usize
        .checked_add(usize::from(!DEFAULT_MESSAGE_VERSIONS.is_empty()))
        .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    let message_memory = minimal_message_info_memory()?;

    for message in messages {
        let length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        // MessageInfo::validate runs before ArchiveInfo's aggregate checks.
        // Preserve that error precedence while avoiding repeated version
        // storage work performed by MessageInfo::conventional.
        if DEFAULT_MESSAGE_VERSIONS.len() > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                DEFAULT_MESSAGE_VERSIONS.len(),
                limits.max_metadata_items(),
            ));
        }

        let message_length = minimal_message_info_encoded_length(message.type_, length)?;
        header_length = checked_encoded_add(
            header_length,
            length_delimited_field_length(2, message_length)?,
        )?;
        header_fields = header_fields
            .checked_add(message_fields)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        header_memory = checked_memory_add(header_memory, message_memory)?;
        add_limited(
            &mut metadata_items,
            DEFAULT_MESSAGE_VERSIONS.len(),
            LimitKind::MetadataItems,
            limits.max_metadata_items(),
        )?;
    }

    check_constructor_header(
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        limits,
    )?;
    Ok(header_length)
}

fn minimal_message_info_memory() -> Result<usize> {
    Ok(MINIMAL_MESSAGE_INFO_MEMORY)
}

/// Account the separately owned version buffer without charging the inline
/// three-value representation twice. A spilled `SmallVec` owns the complete
/// sequence in its heap buffer; the neutral budget intentionally follows the
/// existing logical-length contract rather than allocator capacity.
fn message_versions_memory(
    versions: &SmallVec<[u32; MESSAGE_VERSIONS_INLINE_CAPACITY]>,
) -> Result<usize> {
    if versions.spilled() {
        versions
            .len()
            .checked_mul(size_of::<u32>())
            .ok_or_else(|| Error::invalid_archive(0, "message metadata memory overflow"))
    } else {
        Ok(0)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct FieldInfoAdmission {
    encoded_length: usize,
    header_fields: usize,
    header_memory: usize,
    metadata_items: usize,
}

fn field_info_admission(info: &FieldInfo) -> Result<FieldInfoAdmission> {
    let path_length = encoded_field_path_length(&info.path)?;
    let mut encoded_length = length_delimited_field_length(1, path_length)?;
    if let Some(value) = info.r#type {
        encoded_length = checked_encoded_add(
            encoded_length,
            varint_field_length(2, enum_wire_value(value.raw_value()))?,
        )?;
    }
    if let Some(value) = info.unknown_field_rule {
        encoded_length = checked_encoded_add(
            encoded_length,
            varint_field_length(3, enum_wire_value(value.raw_value()))?,
        )?;
    }
    if let Some(value) = info.known_field_rule {
        encoded_length = checked_encoded_add(
            encoded_length,
            varint_field_length(6, enum_wire_value(value.raw_value()))?,
        )?;
    }
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u64_field_length(4, &info.object_references)?,
    )?;
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u64_field_length(5, &info.data_references)?,
    )?;
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u32_field_length(7, &info.known_field_version)?,
    )?;
    if let Some(feature_identifier) = &info.known_field_feature_identifier {
        encoded_length = checked_encoded_add(
            encoded_length,
            length_delimited_field_length(8, feature_identifier.len())?,
        )?;
    }

    let mut header_fields = 1usize;
    for present in [
        info.r#type.is_some(),
        info.unknown_field_rule.is_some(),
        !info.object_references.is_empty(),
        !info.data_references.is_empty(),
        info.known_field_rule.is_some(),
        !info.known_field_version.is_empty(),
        info.known_field_feature_identifier.is_some(),
    ] {
        header_fields = header_fields
            .checked_add(usize::from(present))
            .ok_or_else(|| Error::invalid_archive(0, "field count overflow"))?;
    }
    header_fields = header_fields
        .checked_add(field_path_header_field_count(&info.path))
        .ok_or_else(|| Error::invalid_archive(0, "field count overflow"))?;

    let mut metadata_items = 1usize;
    for value in [
        info.path.path.len(),
        info.object_references.len(),
        info.data_references.len(),
        info.known_field_version.len(),
        usize::from(info.known_field_feature_identifier.is_some()),
    ] {
        metadata_items = metadata_items
            .checked_add(value)
            .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
    }

    let mut header_memory = size_of::<FieldInfo>();
    header_memory = checked_memory_add(header_memory, field_path_bytes(&info.path)?)?;
    header_memory = checked_memory_add(
        header_memory,
        info.object_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    header_memory = checked_memory_add(
        header_memory,
        info.data_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    header_memory = checked_memory_add(
        header_memory,
        info.known_field_version
            .len()
            .checked_mul(size_of::<u32>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    header_memory = checked_memory_add(
        header_memory,
        info.known_field_feature_identifier
            .as_deref()
            .map_or(0, str::len),
    )?;

    Ok(FieldInfoAdmission {
        encoded_length,
        header_fields,
        header_memory,
        metadata_items,
    })
}

fn message_info_admission(info: &MessageInfo, limits: Limits) -> Result<MessageInfoAdmission> {
    message_info_admission_with_values(info, info.type_, info.length, limits)
}

fn message_info_admission_with_values(
    info: &MessageInfo,
    type_: u32,
    message_length: u32,
    limits: Limits,
) -> Result<MessageInfoAdmission> {
    let length = usize::try_from(message_length)
        .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
    check_message_length(length, limits)?;

    let mut encoded_length = varint_field_length(1, u64::from(type_))?;
    encoded_length =
        checked_encoded_add(encoded_length, packed_u32_field_length(2, &info.versions)?)?;
    encoded_length = checked_encoded_add(
        encoded_length,
        varint_field_length(3, u64::from(message_length))?,
    )?;
    let mut header_fields = 2usize;
    for present in [
        !info.versions.is_empty(),
        !info.object_references.is_empty(),
        !info.data_references.is_empty(),
        info.base_message_index.is_some(),
        !info.diff_merge_version.is_empty(),
        info.diff_field_path.is_some(),
        !info.diff_read_version.is_empty(),
    ] {
        header_fields = header_fields
            .checked_add(usize::from(present))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    }
    header_fields = header_fields
        .checked_add(info.field_infos.len())
        .and_then(|value| value.checked_add(info.fields_to_remove.len()))
        .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;

    let mut metadata_items = 0usize;
    for value in [
        info.versions.len(),
        info.field_infos.len(),
        info.object_references.len(),
        info.data_references.len(),
        info.diff_merge_version.len(),
        usize::from(info.diff_field_path.is_some()),
        info.fields_to_remove.len(),
        info.diff_read_version.len(),
    ] {
        metadata_items = metadata_items
            .checked_add(value)
            .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
    }

    let mut header_memory = checked_memory_add(
        size_of::<MessageInfo>(),
        message_versions_memory(&info.versions)?,
    )?;
    for (length, element_size) in [
        (info.object_references.len(), size_of::<u64>()),
        (info.data_references.len(), size_of::<u64>()),
        (info.diff_merge_version.len(), size_of::<u32>()),
        (info.diff_read_version.len(), size_of::<u32>()),
    ] {
        header_memory = checked_memory_add(
            header_memory,
            length
                .checked_mul(element_size)
                .ok_or_else(|| Error::invalid_archive(0, "message metadata memory overflow"))?,
        )?;
    }

    for field_info in &info.field_infos {
        let admission = field_info_admission(field_info)?;
        encoded_length = checked_encoded_add(
            encoded_length,
            length_delimited_field_length(4, admission.encoded_length)?,
        )?;
        header_fields = header_fields
            .checked_add(admission.header_fields)
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
        metadata_items = metadata_items
            .checked_add(admission.metadata_items)
            .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
        header_memory = checked_memory_add(header_memory, admission.header_memory)?;
    }
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u64_field_length(5, &info.object_references)?,
    )?;
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u64_field_length(6, &info.data_references)?,
    )?;
    if let Some(base_message_index) = info.base_message_index {
        encoded_length = checked_encoded_add(
            encoded_length,
            varint_field_length(7, u64::from(base_message_index))?,
        )?;
    }
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u32_field_length(8, &info.diff_merge_version)?,
    )?;
    if let Some(path) = &info.diff_field_path {
        let path_length = encoded_field_path_length(path)?;
        encoded_length = checked_encoded_add(
            encoded_length,
            length_delimited_field_length(9, path_length)?,
        )?;
        header_fields = header_fields
            .checked_add(field_path_header_field_count(path))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
        metadata_items = metadata_items
            .checked_add(path.path.len())
            .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
        header_memory = checked_memory_add(header_memory, field_path_bytes(path)?)?;
    }
    for path in &info.fields_to_remove {
        let path_length = encoded_field_path_length(path)?;
        encoded_length = checked_encoded_add(
            encoded_length,
            length_delimited_field_length(10, path_length)?,
        )?;
        header_fields = header_fields
            .checked_add(field_path_header_field_count(path))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
        metadata_items = metadata_items
            .checked_add(path.path.len())
            .ok_or_else(|| Error::invalid_archive(0, "metadata item count overflow"))?;
        header_memory = checked_memory_add(header_memory, field_path_bytes(path)?)?;
    }
    header_memory = checked_memory_add(
        header_memory,
        info.fields_to_remove
            .len()
            .checked_mul(size_of::<FieldPath>())
            .ok_or_else(|| Error::invalid_archive(0, "message metadata memory overflow"))?,
    )?;
    encoded_length = checked_encoded_add(
        encoded_length,
        packed_u32_field_length(11, &info.diff_read_version)?,
    )?;
    let header_nesting = if info.field_infos.is_empty() {
        usize::from(info.diff_field_path.is_some() || !info.fields_to_remove.is_empty())
    } else {
        2
    };
    if metadata_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            metadata_items,
            limits.max_metadata_items(),
        ));
    }
    Ok(MessageInfoAdmission {
        encoded_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
    })
}

fn check_archive_info_constructor(
    info: &ArchiveInfo,
    limits: Limits,
) -> Result<ArchiveInfoAdmission> {
    if info.identifier.is_none() {
        return Err(Error::invalid_archive(
            0,
            "object is missing its archive identifier",
        ));
    }
    let message_count = info.message_infos.len();
    if message_count > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            message_count,
            limits.max_messages_per_object(),
        ));
    }
    if info.should_merge.is_none() && info.message_infos.iter().all(is_minimal_message_info) {
        if limits.is_default_profile() {
            return check_minimal_archive_info_default(info, limits);
        }
        return check_minimal_archive_info_constructor(info, limits);
    }
    let mut metadata_items = message_count;
    let mut header_fields = usize::from(info.identifier.is_some());
    header_fields = header_fields
        .checked_add(message_count)
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    header_fields = header_fields
        .checked_add(usize::from(info.should_merge.is_some()))
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    let mut header_nesting = usize::from(message_count != 0);
    let mut header_memory = size_of::<ArchiveInfo>();
    let mut encoded_messages_length = 0usize;
    for message_info in &info.message_infos {
        let admission = message_info_admission(message_info, limits)?;
        encoded_messages_length = checked_encoded_add(
            encoded_messages_length,
            length_delimited_field_length(2, admission.encoded_length)?,
        )?;
        add_limited(
            &mut metadata_items,
            admission.metadata_items,
            LimitKind::MetadataItems,
            limits.max_metadata_items(),
        )?;
        header_fields = header_fields
            .checked_add(admission.header_fields)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        header_nesting = header_nesting.max(
            1usize
                .checked_add(admission.header_nesting)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo nesting overflow"))?,
        );
        header_memory = checked_memory_add(header_memory, admission.header_memory)?;
    }
    let mut header_length = varint_field_length(
        1,
        info.identifier
            .expect("archive identifier was checked above"),
    )?;
    header_length = checked_encoded_add(header_length, encoded_messages_length)?;
    if let Some(should_merge) = info.should_merge {
        header_length = checked_encoded_add(
            header_length,
            varint_field_length(3, u64::from(should_merge))?,
        )?;
    }
    check_constructor_header(
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        limits,
    )?;
    Ok(ArchiveInfoAdmission {
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        message_count,
    })
}

/// Admit the exact header that serialization will emit after replacing each
/// MessageInfo's wire type and payload length from its current RawMessage.
/// This mirrors `Archive::to_bytes` without cloning the metadata projection.
fn check_archive_info_constructor_for_messages(
    info: &ArchiveInfo,
    messages: &[RawMessage],
    limits: Limits,
) -> Result<ArchiveInfoAdmission> {
    if info.identifier.is_none() {
        return Err(Error::invalid_archive(
            0,
            "object is missing its archive identifier",
        ));
    }
    if info.message_infos.len() != messages.len() {
        return Err(Error::invalid_archive(
            0,
            "message metadata and payload counts differ",
        ));
    }
    let message_count = info.message_infos.len();
    if message_count > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            message_count,
            limits.max_messages_per_object(),
        ));
    }
    let mut metadata_items = message_count;
    let mut header_fields = usize::from(info.identifier.is_some());
    header_fields = header_fields
        .checked_add(message_count)
        .and_then(|value| value.checked_add(usize::from(info.should_merge.is_some())))
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    let mut header_nesting = usize::from(message_count != 0);
    let mut header_memory = size_of::<ArchiveInfo>();
    let mut encoded_messages_length = 0usize;
    for (message_info, message) in info.message_infos.iter().zip(messages) {
        let message_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        let admission = message_info_admission_with_values(
            message_info,
            message.type_,
            message_length,
            limits,
        )?;
        encoded_messages_length = checked_encoded_add(
            encoded_messages_length,
            length_delimited_field_length(2, admission.encoded_length)?,
        )?;
        add_limited(
            &mut metadata_items,
            admission.metadata_items,
            LimitKind::MetadataItems,
            limits.max_metadata_items(),
        )?;
        header_fields = header_fields
            .checked_add(admission.header_fields)
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
        header_nesting = header_nesting.max(
            1usize
                .checked_add(admission.header_nesting)
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo nesting overflow"))?,
        );
        header_memory = checked_memory_add(header_memory, admission.header_memory)?;
    }
    let mut header_length = varint_field_length(
        1,
        info.identifier
            .expect("archive identifier was checked above"),
    )?;
    header_length = checked_encoded_add(header_length, encoded_messages_length)?;
    if let Some(should_merge) = info.should_merge {
        header_length = checked_encoded_add(
            header_length,
            varint_field_length(3, u64::from(should_merge))?,
        )?;
    }
    check_constructor_header(
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        limits,
    )?;
    Ok(ArchiveInfoAdmission {
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        message_count,
    })
}

/// Admit the common default-profile constructor without repeating the
/// general limit checks whose maxima are fixed hard ceilings. The arithmetic
/// still computes the exact neutral header totals returned to retained
/// provenance and uses checked operations so an oversized caller-owned vector
/// cannot bypass the contract.
#[inline]
fn check_minimal_archive_info_default(
    info: &ArchiveInfo,
    limits: Limits,
) -> Result<ArchiveInfoAdmission> {
    let identifier = info
        .identifier
        .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
    let message_count = info.message_infos.len();
    let header_fields = 1usize
        .checked_add(message_count)
        .and_then(|value| value.checked_add(message_count.checked_mul(3)?))
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    let header_nesting = usize::from(message_count != 0);
    let mut header_memory = size_of::<ArchiveInfo>();
    let mut header_length = varint_field_length(1, identifier)?;
    let mut metadata_items = message_count;
    for message_info in &info.message_infos {
        header_memory = checked_memory_add(
            header_memory,
            checked_memory_add(
                minimal_message_info_memory()?,
                message_versions_memory(&message_info.versions)?,
            )?,
        )?;
        let message_length = usize::try_from(message_info.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        check_message_length(message_length, limits)?;
        metadata_items = metadata_items
            .checked_add(DEFAULT_MESSAGE_VERSIONS.len())
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo metadata count overflow"))?;
        let message_length =
            minimal_message_info_encoded_length(message_info.type_, message_info.length)?;
        let encoded = length_delimited_field_length(2, message_length)?;
        header_length = checked_encoded_add(header_length, encoded)?;
    }
    Ok(ArchiveInfoAdmission {
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        message_count,
    })
}

#[inline]
fn is_minimal_message_info(info: &MessageInfo) -> bool {
    !info.versions.spilled()
        && info.versions.as_slice() == DEFAULT_MESSAGE_VERSIONS
        && info.field_infos.is_empty()
        && info.object_references.is_empty()
        && info.data_references.is_empty()
        && info.base_message_index.is_none()
        && info.diff_merge_version.is_empty()
        && info.diff_field_path.is_none()
        && info.fields_to_remove.is_empty()
        && info.diff_read_version.is_empty()
}

/// Admit the conventional records produced by [`MessageInfo::new`] without
/// entering the general metadata encoder. The generic path remains the
/// authority for any caller-mutated `MessageInfo`.
#[inline]
fn check_minimal_archive_info_constructor(
    info: &ArchiveInfo,
    limits: Limits,
) -> Result<ArchiveInfoAdmission> {
    let identifier = info
        .identifier
        .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
    let message_count = info.message_infos.len();
    let mut metadata_items = message_count;
    let header_fields = 1usize
        .checked_add(message_count)
        .and_then(|value| value.checked_add(message_count.checked_mul(3)?))
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo field count overflow"))?;
    let header_nesting = usize::from(message_count != 0);
    let mut header_memory = size_of::<ArchiveInfo>();
    let mut header_length = varint_field_length(1, identifier)?;
    for message_info in &info.message_infos {
        header_memory = checked_memory_add(
            header_memory,
            checked_memory_add(
                minimal_message_info_memory()?,
                message_versions_memory(&message_info.versions)?,
            )?,
        )?;
        let message_length = usize::try_from(message_info.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        check_message_length(message_length, limits)?;
        add_limited(
            &mut metadata_items,
            DEFAULT_MESSAGE_VERSIONS.len(),
            LimitKind::MetadataItems,
            limits.max_metadata_items(),
        )?;
        let message_length =
            minimal_message_info_encoded_length(message_info.type_, message_info.length)?;
        header_length = checked_encoded_add(
            header_length,
            length_delimited_field_length(2, message_length)?,
        )?;
    }
    check_constructor_header(
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        limits,
    )?;
    Ok(ArchiveInfoAdmission {
        header_length,
        header_fields,
        header_nesting,
        header_memory,
        metadata_items,
        message_count,
    })
}

#[inline]
fn minimal_message_info_encoded_length(type_: u32, length: u32) -> Result<usize> {
    Ok(7 + u64_varint_length(u64::from(type_)) + u64_varint_length(u64::from(length)))
}

#[inline]
fn minimal_message_info_field_length(type_: u32, length: u32) -> Result<usize> {
    length_delimited_field_length(2, minimal_message_info_encoded_length(type_, length)?)
}

fn field_path_memory(path: &FieldPath) -> Result<usize> {
    let bytes = field_path_bytes(path)?;
    checked_memory_add(size_of::<FieldPath>(), bytes)
}

fn field_path_bytes(path: &FieldPath) -> Result<usize> {
    path.path
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or_else(|| Error::invalid_archive(0, "field path memory overflow"))
}

#[cfg(test)]
fn field_info_memory(info: &FieldInfo) -> Result<usize> {
    let mut memory = size_of::<FieldInfo>();
    // `FieldInfo.path` is inline in the record; only its owned component
    // buffer is additional retained memory.
    memory = checked_memory_add(memory, field_path_bytes(&info.path)?)?;
    memory = checked_memory_add(
        memory,
        info.object_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    memory = checked_memory_add(
        memory,
        info.data_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    memory = checked_memory_add(
        memory,
        info.known_field_version
            .len()
            .checked_mul(size_of::<u32>())
            .ok_or_else(|| Error::invalid_archive(0, "field metadata memory overflow"))?,
    )?;
    checked_memory_add(
        memory,
        info.known_field_feature_identifier
            .as_deref()
            .map_or(0, str::len),
    )
}

#[cfg(test)]
fn message_info_memory(info: &MessageInfo) -> Result<usize> {
    let mut memory = checked_memory_add(
        size_of::<MessageInfo>(),
        message_versions_memory(&info.versions)?,
    )?;
    for (length, element_size) in [
        (info.object_references.len(), size_of::<u64>()),
        (info.data_references.len(), size_of::<u64>()),
        (info.diff_merge_version.len(), size_of::<u32>()),
        (info.diff_read_version.len(), size_of::<u32>()),
    ] {
        memory = checked_memory_add(
            memory,
            length
                .checked_mul(element_size)
                .ok_or_else(|| Error::invalid_archive(0, "message metadata memory overflow"))?,
        )?;
    }
    for field_info in &info.field_infos {
        memory = checked_memory_add(memory, field_info_memory(field_info)?)?;
    }
    if let Some(path) = &info.diff_field_path {
        // `diff_field_path` is an inline `Option<FieldPath>`; account only
        // for the path's separately owned component buffer here.
        memory = checked_memory_add(memory, field_path_bytes(path)?)?;
    }
    memory = checked_memory_add(
        memory,
        info.fields_to_remove
            .len()
            .checked_mul(size_of::<FieldPath>())
            .ok_or_else(|| Error::invalid_archive(0, "message metadata memory overflow"))?,
    )?;
    for path in &info.fields_to_remove {
        memory = checked_memory_add(memory, field_path_bytes(path)?)?;
    }
    Ok(memory)
}

fn field_path_header_field_count(path: &FieldPath) -> usize {
    usize::from(!path.path.is_empty())
}

#[cfg(test)]
fn field_info_header_field_count(info: &FieldInfo) -> Result<usize> {
    let mut fields = 1usize; // required nested FieldPath
    for present in [
        info.r#type.is_some(),
        info.unknown_field_rule.is_some(),
        !info.object_references.is_empty(),
        !info.data_references.is_empty(),
        info.known_field_rule.is_some(),
        !info.known_field_version.is_empty(),
        info.known_field_feature_identifier.is_some(),
    ] {
        fields = fields
            .checked_add(usize::from(present))
            .ok_or_else(|| Error::invalid_archive(0, "field count overflow"))?;
    }
    fields
        .checked_add(field_path_header_field_count(&info.path))
        .ok_or_else(|| Error::invalid_archive(0, "field count overflow"))
}

#[cfg(test)]
fn message_info_header_field_count(info: &MessageInfo) -> Result<usize> {
    let mut fields = 2usize; // required type and length
    for present in [
        !info.versions.is_empty(),
        !info.object_references.is_empty(),
        !info.data_references.is_empty(),
        info.base_message_index.is_some(),
        !info.diff_merge_version.is_empty(),
        info.diff_field_path.is_some(),
        !info.diff_read_version.is_empty(),
    ] {
        fields = fields
            .checked_add(usize::from(present))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    }
    fields = fields
        .checked_add(info.field_infos.len())
        .and_then(|value| value.checked_add(info.fields_to_remove.len()))
        .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    for field_info in &info.field_infos {
        fields = fields
            .checked_add(field_info_header_field_count(field_info)?)
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    }
    if let Some(path) = &info.diff_field_path {
        fields = fields
            .checked_add(field_path_header_field_count(path))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    }
    for path in &info.fields_to_remove {
        fields = fields
            .checked_add(field_path_header_field_count(path))
            .ok_or_else(|| Error::invalid_archive(0, "message field count overflow"))?;
    }
    Ok(fields)
}

#[cfg(test)]
fn message_info_header_nesting(info: &MessageInfo) -> Result<usize> {
    let nested = if info.field_infos.is_empty() {
        usize::from(info.diff_field_path.is_some() || !info.fields_to_remove.is_empty())
    } else {
        2
    };
    Ok(nested)
}

fn checked_memory_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| Error::invalid_archive(0, "neutral metadata memory overflow"))
}

/// Raw payload data for one archive message.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawMessage {
    pub type_: u32,
    pub data: Vec<u8>,
}

/// One object within a decompressed IWA archive stream.
#[derive(Debug, Clone)]
pub struct ArchiveObject {
    pub archive_info: ArchiveInfo,
    pub messages: Vec<RawMessage>,
    /// Offset of the object's length prefix in the parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub header_offset: u64,
    /// Size of the length prefix and encoded `ArchiveInfo` header in the
    /// parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub header_length: u64,
    /// Offset of the first payload in the parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub data_offset: u64,
    /// Total bytes occupied by the object's payloads in the parsed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations. Current payload
    /// size is available by summing [`Self::messages`].
    pub data_length: u64,
    /// Immutable source bytes and admission totals for a non-canonical
    /// source header. Canonical objects have no retained provenance. The
    /// allocation is shared by clones and is cleared by public mutations.
    source_provenance: Option<Arc<SourceProvenance>>,
}

impl PartialEq for ArchiveObject {
    fn eq(&self, other: &Self) -> bool {
        self.archive_info == other.archive_info
            && self.messages == other.messages
            && self.header_offset == other.header_offset
            && self.header_length == other.header_length
            && self.data_offset == other.data_offset
            && self.data_length == other.data_length
            && self.source_headers() == other.source_headers()
    }
}

impl ArchiveObject {
    /// Construct an object under explicit resource limits.
    pub fn new(identifier: u64, messages: Vec<RawMessage>, limits: Limits) -> Result<Self> {
        let header_length = preflight_archive_object_constructor(identifier, &messages, limits)?;
        let mut message_infos = Vec::new();
        message_infos
            .try_reserve_exact(messages.len())
            .map_err(|_| Error::allocation("IWA message metadata", messages.len()))?;
        for message in &messages {
            let length = u32::try_from(message.data.len())
                .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            message_infos.push(MessageInfo::conventional(message.type_, length));
        }
        let archive_info = ArchiveInfo {
            identifier: Some(identifier),
            message_infos,
            should_merge: None,
        };
        let object = Self {
            archive_info,
            messages,
            header_offset: 0,
            header_length: 0,
            data_offset: 0,
            data_length: 0,
            source_provenance: None,
        };
        // `preflight_archive_object_constructor` has already admitted the
        // complete neutral header before metadata allocation. Validate only
        // the payload/object framing here; calling the public validator would
        // walk that header a second time on this hot path.
        object.validate_payloads_with_header_length(limits, header_length)?;
        Ok(object)
    }

    /// Clone an object while changing its identity and remapping its known
    /// archive-object references.
    ///
    /// This is the physical primitive used by format owners when they copy a
    /// graph node into a new component. `object_remap` is a strict, finite
    /// old-to-new map. Its entries are applied to aggregate and nested
    /// [`FieldInfo`] object references; data references and payload bytes are
    /// intentionally left untouched. The source object identifier may appear
    /// in the map, but when it does its target must equal `new_identifier`.
    ///
    /// Replacement messages must have the same arity and message types as the
    /// source. Their bytes are copied into independently owned buffers and
    /// their lengths are reflected in the cloned metadata. Header rewriting
    /// operates on validated raw spans, so unknown fields, duplicate
    /// occurrences, non-canonical keys, and length encodings remain intact.
    /// Package-wide identifier collisions, opaque reference ownership, payload
    /// wire references, and component-registry updates remain the format
    /// owner's responsibility; this method only rewrites the known metadata
    /// projection described above.
    ///
    /// The source is never mutated. A failed validation or allocation leaves
    /// it unchanged because all output state is assembled before publication.
    pub fn clone_with_identity_remap(
        &self,
        new_identifier: u64,
        object_remap: &[(u64, u64)],
        replacement_messages: &[RawMessage],
        limits: Limits,
    ) -> Result<Self> {
        self.validate(limits)?;
        let source_identifier = self
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        validate_clone_arguments(
            source_identifier,
            new_identifier,
            object_remap,
            replacement_messages,
            &self.messages,
            limits,
        )?;
        validate_clone_metadata_semantics(&self.archive_info, source_identifier, new_identifier)?;
        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let source_header =
            retained_clone_source_header(self, canonical_before.as_slice(), limits)?;
        let source_preflight = preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        validate_clone_input_scratch(
            object_remap.len(),
            replacement_messages.len(),
            source_preflight.fields(),
            limits,
        )?;
        let remap =
            prepare_identity_remap(source_identifier, new_identifier, object_remap, limits)?;
        validate_clone_self_reference_mapping(
            &self.archive_info,
            source_identifier,
            new_identifier,
            &remap,
        )?;
        let replacement_lengths = checked_replacement_lengths(replacement_messages, limits)?;

        let rewritten_header_length = preflight_clone_archive_info_header_length(
            source_header,
            &self.archive_info,
            &self.messages,
            replacement_messages,
            new_identifier,
            &replacement_lengths,
            &remap,
            limits,
        )?;
        validate_clone_object_size(rewritten_header_length, replacement_messages, limits)?;
        validate_clone_header_scratch(
            canonical_before.len(),
            source_header.len(),
            rewritten_header_length,
            &self.archive_info,
            remap.len(),
            replacement_lengths.len(),
            source_preflight.fields(),
            limits,
        )?;
        let rewritten_header = rewrite_clone_archive_info_header(
            source_header,
            &self.archive_info,
            &self.messages,
            replacement_messages,
            new_identifier,
            &replacement_lengths,
            &remap,
            limits,
        )?;
        if rewritten_header.len() != rewritten_header_length {
            return Err(Error::invalid_archive(
                0,
                "clone ArchiveInfo rewrite length differs from preflight",
            ));
        }
        let archive_info = clone_archive_info_with_identity_remap(
            &self.archive_info,
            new_identifier,
            &remap,
            &replacement_lengths,
            limits,
        )?;
        let messages = clone_raw_messages_with_limits(replacement_messages, limits)?;
        let decoded = ArchiveInfo::decode(&rewritten_header, limits)?;
        if decoded != archive_info {
            return Err(Error::invalid_archive(
                0,
                "cloned ArchiveInfo raw rewrite differs from its neutral projection",
            ));
        }
        let canonical_after = encode_archive_info(&archive_info, limits)?;
        let retained_header = if rewritten_header == canonical_after {
            None
        } else {
            Some(rewritten_header.into_boxed_slice())
        };
        let retained_canonical = retained_header
            .as_ref()
            .map(|_| canonical_after.into_boxed_slice());
        let cloned = Self {
            archive_info,
            messages,
            header_offset: 0,
            header_length: 0,
            data_offset: 0,
            data_length: 0,
            source_provenance: match (retained_header, retained_canonical) {
                (Some(original), Some(canonical)) => Some(SourceProvenance::from_headers(
                    &original, &canonical, None, None, None,
                )?),
                (None, None) => None,
                _ => unreachable!("retained header pair is constructed together"),
            },
        };
        cloned.validate(limits)?;
        Ok(cloned)
    }

    /// Validate this object under explicit resource limits.
    pub fn validate(&self, limits: Limits) -> Result<()> {
        if self.archive_info.identifier.is_none() {
            return Err(Error::invalid_archive(
                0,
                "object is missing its archive identifier",
            ));
        }
        if self.archive_info.message_infos.len() != self.messages.len() {
            return Err(Error::invalid_archive(
                0,
                "message metadata and payload counts differ",
            ));
        }
        let cached_header_length = self.validate_cached_archive_info(limits)?;
        if cached_header_length.is_none() {
            self.archive_info.validate(limits)?;
        }
        self.validate_payloads(limits, cached_header_length)
    }

    fn source_headers(&self) -> Option<(&[u8], &[u8])> {
        self.source_provenance.as_deref().map(|provenance| {
            (
                provenance.original_header(),
                provenance.original_canonical_header(),
            )
        })
    }

    fn source_header(&self) -> Option<&[u8]> {
        self.source_provenance
            .as_deref()
            .map(SourceProvenance::original_header)
    }

    #[cfg(test)]
    fn source_canonical_header(&self) -> Option<&[u8]> {
        self.source_provenance
            .as_deref()
            .map(SourceProvenance::original_canonical_header)
    }

    #[cfg(test)]
    fn set_test_source_headers(
        &mut self,
        original: Option<Box<[u8]>>,
        canonical: Option<Box<[u8]>>,
    ) {
        self.source_provenance = match (original, canonical) {
            (Some(original), Some(canonical)) => Some(
                SourceProvenance::from_headers(&original, &canonical, None, None, None)
                    .expect("test source header allocation"),
            ),
            (None, None) => None,
            _ => panic!("test source headers must be supplied as a pair"),
        };
    }

    #[cfg(test)]
    fn clear_test_source_headers(&mut self) {
        self.source_provenance = None;
    }

    #[cfg(test)]
    fn test_original_header_mut(&mut self) -> Option<&mut [u8]> {
        self.source_provenance
            .as_mut()
            .map(Arc::make_mut)
            .map(SourceProvenance::original_header_mut)
    }

    #[cfg(test)]
    fn test_original_canonical_header_mut(&mut self) -> Option<&mut [u8]> {
        self.source_provenance
            .as_mut()
            .map(Arc::make_mut)
            .map(SourceProvenance::original_canonical_header_mut)
    }

    /// Reuse parsed neutral admission only after proving that the current
    /// metadata still emits the paired canonical header. Profiles that tighten
    /// sequence-sensitive message or metadata budgets use the full validator,
    /// preserving its first-failure precedence.
    fn validate_cached_archive_info(&self, limits: Limits) -> Result<Option<usize>> {
        let Some(provenance) = self.source_provenance.as_deref() else {
            return Ok(None);
        };
        let Some(admission) = provenance.archive_info_admission else {
            return Ok(None);
        };
        if limits.max_message_bytes()
            < provenance
                .admitted_message_bytes
                .expect("parsed source provenance has message limit")
            || limits.max_metadata_items()
                < provenance
                    .admitted_metadata_items
                    .expect("parsed source provenance has metadata limit")
        {
            return Ok(None);
        }
        let original_canonical = provenance.original_canonical_header();
        let mut probe = CanonicalHeaderProbe::new(Some(original_canonical));
        if emit_canonical_archive_info(&mut probe, self).is_err() {
            return Ok(None);
        }
        let (header_length, canonical_matches) = probe.finish();
        if !canonical_matches || header_length != admission.header_length {
            return Ok(None);
        }
        admission.check(limits)?;
        Ok(Some(provenance.original_header().len()))
    }

    fn validate_payloads(&self, limits: Limits, cached_header_length: Option<usize>) -> Result<()> {
        let header_length = if let Some(header_length) = cached_header_length {
            // Admission precedes the wire-length conversion used to choose
            // the retained header, including for mutated oversized payloads.
            for message in &self.messages {
                check_message_length(message.data.len(), limits)?;
            }
            admit_retained_header(
                self.source_header()
                    .expect("cached retained header is present"),
                limits,
                self.source_provenance
                    .as_deref()
                    .and_then(|provenance| provenance.header_admission),
            )?;
            header_length
        } else if self.source_header().is_some() {
            retained_header_length_for_object(self, limits)?.1
        } else {
            archive_info_encoded_len(&self.archive_info)?
        };
        self.validate_payloads_with_header_length(limits, header_length)
    }

    fn validate_payloads_with_header_length(
        &self,
        limits: Limits,
        header_length: usize,
    ) -> Result<()> {
        check_header_length(header_length, limits)?;
        let prefix_length = varint_len(header_length)?;
        let mut payload_length = 0;
        for message in &self.messages {
            check_message_length(message.data.len(), limits)?;
            add_limited(
                &mut payload_length,
                message.data.len(),
                LimitKind::ObjectBytes,
                limits.max_object_bytes(),
            )?;
        }
        let object_length = prefix_length
            .checked_add(header_length)
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        Ok(())
    }

    /// Return the type identifier of the first payload, if present.
    #[must_use]
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }

    /// Compare the source-authoritative contents of two physical objects.
    ///
    /// The comparison is allocation-free and ignores only the two offsets
    /// whose values are relative to the containing component. It includes
    /// decoded archive/message metadata, raw message bytes, retained raw and
    /// canonical `ArchiveInfo` headers, and the recorded header/payload
    /// framing lengths. In particular, objects with different length-prefix
    /// widths are not aliases even when their decoded metadata and payloads
    /// match.
    #[must_use]
    pub fn same_content_ignoring_offsets(&self, other: &Self) -> bool {
        self.archive_info == other.archive_info
            && self.messages == other.messages
            && self.header_length == other.header_length
            && self.data_length == other.data_length
            && self.source_headers() == other.source_headers()
    }

    /// Stream every aggregate and field-level object/data reference.
    ///
    /// The retained raw header is preflighted under the same rule used for
    /// publication, then cross-checked against the neutral projection before
    /// callbacks begin. This keeps producer-authored framing authoritative
    /// without exposing raw header bytes to format owners.
    pub fn inspect_references(
        &self,
        visitor: &mut impl ArchiveReferenceVisitor,
        limits: Limits,
    ) -> Result<usize> {
        self.inspect_references_with_policy(
            visitor,
            ArchiveReferencePolicy::KnownReferences,
            limits,
        )
    }

    /// Stream physical metadata references with an explicit completeness policy.
    pub fn inspect_references_with_policy(
        &self,
        visitor: &mut impl ArchiveReferenceVisitor,
        policy: ArchiveReferencePolicy,
        limits: Limits,
    ) -> Result<usize> {
        self.validate(limits)?;
        let canonical = encode_archive_info(&self.archive_info, limits)?;
        let source_header = match self.source_headers() {
            Some((original, original_canonical)) if original_canonical == canonical.as_slice() => {
                original
            },
            _ => canonical.as_slice(),
        };
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        if matches!(policy, ArchiveReferencePolicy::RejectUnknownMetadata)
            && header_has_unknown_reference_metadata(source_header, limits)?
        {
            return Err(Error::invalid_archive(
                0,
                "archive reference census encountered unknown metadata",
            ));
        }
        let decoded = ArchiveInfo::decode(source_header, limits)?;
        if decoded != self.archive_info {
            return Err(Error::invalid_archive(
                0,
                "source ArchiveInfo does not match its neutral projection",
            ));
        }
        let object_identifier = decoded
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object identifier is missing"))?;
        let mut occurrences = 0usize;
        for (message_index, message) in decoded.message_infos.iter().enumerate() {
            for (kind, identifiers) in [
                (
                    ArchiveReferenceKind::Object,
                    message.object_references.as_slice(),
                ),
                (
                    ArchiveReferenceKind::Data,
                    message.data_references.as_slice(),
                ),
            ] {
                for referenced_identifier in identifiers {
                    occurrences = occurrences.checked_add(1).ok_or_else(|| {
                        Error::invalid_archive(message_index, "reference count overflow")
                    })?;
                    if occurrences > limits.max_metadata_items() {
                        return Err(limit(
                            LimitKind::MetadataItems,
                            occurrences,
                            limits.max_metadata_items(),
                        ));
                    }
                    visitor.visit_reference(ArchiveReferenceOccurrence {
                        object_identifier,
                        message_index,
                        scope: ArchiveReferenceScope::Message,
                        kind,
                        referenced_identifier: *referenced_identifier,
                    })?;
                }
            }
            for (field_index, field) in message.field_infos.iter().enumerate() {
                for (kind, identifiers) in [
                    (
                        ArchiveReferenceKind::Object,
                        field.object_references.as_slice(),
                    ),
                    (ArchiveReferenceKind::Data, field.data_references.as_slice()),
                ] {
                    for referenced_identifier in identifiers {
                        occurrences = occurrences.checked_add(1).ok_or_else(|| {
                            Error::invalid_archive(message_index, "reference count overflow")
                        })?;
                        if occurrences > limits.max_metadata_items() {
                            return Err(limit(
                                LimitKind::MetadataItems,
                                occurrences,
                                limits.max_metadata_items(),
                            ));
                        }
                        visitor.visit_reference(ArchiveReferenceOccurrence {
                            object_identifier,
                            message_index,
                            scope: ArchiveReferenceScope::Field { field_index },
                            kind,
                            referenced_identifier: *referenced_identifier,
                        })?;
                    }
                }
            }
        }
        Ok(occurrences)
    }

    /// Replace one payload under explicit resource limits.
    pub fn replace_message(
        &mut self,
        index: usize,
        message: RawMessage,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.validate(limits)?;
        let message_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;
        let old_type = message_info.type_;
        let old_length = message_info.length;
        let old = std::mem::replace(
            self.messages
                .get_mut(index)
                .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?,
            message,
        );
        let info = &mut self.archive_info.message_infos[index];
        info.type_ = self.messages[index].type_;
        info.length = length;
        if let Err(error) = self.validate(limits) {
            drop(std::mem::replace(&mut self.messages[index], old));
            let restored_info = &mut self.archive_info.message_infos[index];
            restored_info.type_ = old_type;
            restored_info.length = old_length;
            return Err(error);
        }
        Ok(old)
    }

    /// Replace one payload while preserving untouched `ArchiveInfo` bytes.
    ///
    /// Unlike [`Self::replace_message`], this variant surgically updates only
    /// the effective `MessageInfo.type` and `MessageInfo.length` scalar values
    /// plus an enclosing length prefix when its width must change. Unknown
    /// fields, duplicate metadata, non-canonical keys and every unrelated raw
    /// header byte remain untouched. The mutation is atomic: validation or
    /// allocation failure leaves the object unchanged.
    pub fn replace_message_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        // Re-apply the caller's current limit profile to retained raw bytes.
        // An object may have been opened under a broader profile.
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_in_header(
            source_header,
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            limits,
        )?;

        let old_type = current_info.type_;
        let old_length = current_info.length;
        let old = self.replace_message(index, message, limits)?;
        let verification = (|| {
            let decoded = ArchiveInfo::decode(&rewritten_header, limits)?;
            if decoded != self.archive_info {
                return Err(Error::invalid_archive(
                    index,
                    "raw ArchiveInfo rewrite changed unrelated metadata",
                ));
            }
            let canonical = encode_archive_info(&self.archive_info, limits)?;
            let published_header_length = if retained_source_header && rewritten_header != canonical
            {
                rewritten_header.len()
            } else {
                canonical.len()
            };
            validate_raw_object_size(self, published_header_length, limits)?;
            Ok(canonical)
        })();
        let canonical_after = match verification {
            Ok(canonical) => canonical,
            Err(error) => {
                restore_replaced_message(self, index, old, old_type, old_length);
                return Err(error);
            },
        };

        let source_provenance = if retained_source_header && rewritten_header != canonical_after {
            Some(SourceProvenance::from_headers(
                &rewritten_header,
                &canonical_after,
                None,
                None,
                None,
            )?)
        } else {
            None
        };
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Replace one payload and prune selected object references while
    /// preserving every untouched `ArchiveInfo` byte.
    ///
    /// The identifiers are removed from the selected `MessageInfo`'s
    /// top-level object references and from the object references of every
    /// nested `FieldInfo`. Packed and unpacked uint64 fields may be mixed;
    /// retained values keep their original encoding and order. Duplicate
    /// identifiers in `identifiers` are harmless. An empty slice delegates to
    /// [`Self::replace_message_preserving_header`] exactly.
    ///
    /// This physical primitive does not prove graph exclusivity. The caller
    /// must establish that every supplied identifier has no surviving
    /// payload or opaque-field ownership and that removing it from *all* of
    /// the selected message's reference metadata is semantically valid.
    pub fn replace_message_pruning_object_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        identifiers: &[u64],
        limits: Limits,
    ) -> Result<RawMessage> {
        self.replace_message_pruning_references_preserving_header(
            index,
            message,
            identifiers,
            DataReferencePruning::None,
            limits,
        )
    }

    /// Replace one payload and prune selected object references plus selected
    /// or all data references while preserving every untouched header byte.
    ///
    /// Both reference kinds are pruned from the target `MessageInfo` and all
    /// nested `FieldInfo` values. Packed and unpacked occurrences may be
    /// mixed, and retained values keep their original order and encoding.
    /// Duplicate selected identifiers are harmless. Empty object identifiers
    /// together with [`DataReferencePruning::None`] or
    /// [`DataReferencePruning::Selected`] containing an empty slice delegate
    /// exactly to [`Self::replace_message_preserving_header`].
    ///
    /// The caller must prove that removing every supplied reference is valid
    /// for the surrounding package graph; this physical primitive establishes
    /// only bounded wire and neutral-metadata correctness.
    pub fn replace_message_pruning_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        object_identifiers: &[u64],
        data_references: DataReferencePruning<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        if object_identifiers.is_empty() && data_references.is_none() {
            return self.replace_message_preserving_header(index, message, limits);
        }

        let data_identifier_count = match data_references {
            DataReferencePruning::Selected(identifiers) => identifiers.len(),
            DataReferencePruning::None | DataReferencePruning::All => 0,
        };
        let selector_count = object_identifiers
            .len()
            .checked_add(data_identifier_count)
            .ok_or_else(|| Error::invalid_archive(index, "reference selector count overflow"))?;
        if selector_count > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                selector_count,
                limits.max_metadata_items(),
            ));
        }
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let mut object_removals = HashSet::new();
        object_removals
            .try_reserve(object_identifiers.len())
            .map_err(|_| {
                Error::allocation("IWA object-reference removal set", object_identifiers.len())
            })?;
        object_removals.extend(object_identifiers.iter().copied());

        let mut selected_data_removals = HashSet::new();
        if let DataReferencePruning::Selected(identifiers) = data_references {
            selected_data_removals
                .try_reserve(identifiers.len())
                .map_err(|_| {
                    Error::allocation("IWA data-reference removal set", identifiers.len())
                })?;
            selected_data_removals.extend(identifiers.iter().copied());
        }
        let data_removals = match data_references {
            DataReferencePruning::None | DataReferencePruning::Selected([]) => {
                ReferenceRemovals::None
            },
            DataReferencePruning::Selected(_) => {
                ReferenceRemovals::Selected(&selected_data_removals)
            },
            DataReferencePruning::All => ReferenceRemovals::All,
        };

        let object_removals = if object_identifiers.is_empty() {
            ReferenceRemovals::None
        } else {
            ReferenceRemovals::Selected(&object_removals)
        };

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        // The retained header may have been admitted by a broader profile.
        // Revalidate its complete nested wire tree under the caller's limits.
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_prune_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            object_removals,
            data_removals,
            limits,
        )?;

        let rewritten_info = ArchiveInfo::decode(&rewritten_header, limits)?;
        verify_pruned_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            object_removals,
            data_removals,
        )?;
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let source_provenance = if retain_rewritten_header {
            Some(SourceProvenance::from_headers(
                &rewritten_header,
                &canonical_after,
                None,
                None,
                None,
            )?)
        } else {
            None
        };

        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Replace one payload and reorder a selected subsequence of its aggregate
    /// object references while preserving every other header byte.
    ///
    /// Every identifier in `reordered_selected_identifiers` must be unique and
    /// occur exactly once in the target `MessageInfo`'s current aggregate
    /// object-reference list. The selected source slots receive those
    /// identifiers in the supplied order; all nonselected aggregate
    /// occurrences retain their positions and raw encodings. Nested
    /// `FieldInfo` metadata is neither inspected nor edited. An empty or
    /// already-current selected order performs the same header-preserving
    /// replacement without a reference edit.
    ///
    /// The complete expected aggregate list is borrowed directly from `self`
    /// and checked against the retained raw header, avoiding a caller-side
    /// clone while making stale raw provenance fail atomically. This physical
    /// primitive does not decide which object identifiers are semantically
    /// reorderable; the caller retains that format-level responsibility.
    pub fn replace_message_reordering_object_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        reordered_selected_identifiers: &[u64],
        limits: Limits,
    ) -> Result<RawMessage> {
        if reordered_selected_identifiers.len() > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                reordered_selected_identifiers.len(),
                limits.max_metadata_items(),
            ));
        }
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let expected_object_references = current_info.object_references.as_slice();
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let mut selected = HashMap::new();
        selected
            .try_reserve(reordered_selected_identifiers.len())
            .map_err(|_| {
                Error::allocation(
                    "IWA reordered object-reference selector",
                    reordered_selected_identifiers.len(),
                )
            })?;
        for identifier in reordered_selected_identifiers {
            if selected.insert(*identifier, false).is_some() {
                return Err(Error::invalid_archive(
                    index,
                    "reordered object-reference identifiers are not unique",
                ));
            }
        }

        for identifier in expected_object_references {
            if let Some(seen) = selected.get_mut(identifier) {
                if *seen {
                    return Err(Error::invalid_archive(
                        index,
                        "selected object reference occurs more than once",
                    ));
                }
                *seen = true;
            }
        }
        if selected.values().any(|seen| !seen) {
            return Err(Error::invalid_archive(
                index,
                "selected object reference is missing from MessageInfo",
            ));
        }
        let order_is_current = expected_object_references
            .iter()
            .filter(|identifier| selected.contains_key(identifier))
            .copied()
            .eq(reordered_selected_identifiers.iter().copied());
        if order_is_current {
            drop(selected);
            return self.replace_message_preserving_header(index, message, limits);
        }

        let mut reordered_object_references = Vec::new();
        reordered_object_references
            .try_reserve_exact(expected_object_references.len())
            .map_err(|_| {
                Error::allocation(
                    "IWA reordered object-reference list",
                    expected_object_references.len(),
                )
            })?;
        let mut replacements = reordered_selected_identifiers.iter().copied();
        for identifier in expected_object_references {
            if selected.contains_key(identifier) {
                reordered_object_references.push(replacements.next().ok_or_else(|| {
                    Error::invalid_archive(index, "selected object-reference count changed")
                })?);
            } else {
                reordered_object_references.push(*identifier);
            }
        }
        if replacements.next().is_some() {
            return Err(Error::invalid_archive(
                index,
                "selected object-reference count changed",
            ));
        }
        drop(selected);

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_reorder_object_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            expected_object_references,
            &reordered_object_references,
            limits,
        )?;
        drop(canonical_before);

        let rewritten_info = ArchiveInfo::decode(&rewritten_header, limits)?;
        verify_reordered_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            expected_object_references,
            &reordered_object_references,
        )?;
        drop(reordered_object_references);
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let source_provenance = retained_source_provenance(
            &rewritten_header,
            &canonical_after,
            retain_rewritten_header,
        )?;

        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Replace one payload and reorder data references under explicit limits.
    ///
    /// Validation, raw-wire rewriting, neutral projection, and enclosing
    /// object sizing complete before mutation. Unknown fields, field order,
    /// non-canonical encodings, and unrelated metadata remain byte-for-byte
    /// unchanged. For an actual reorder, the uniquely matching nested
    /// `FieldInfo` data-reference list follows the aggregate reorder;
    /// nonmatching nested lists are retained. If more than one nested list
    /// matches, an actual reorder fails atomically because the physical layer
    /// cannot select an owner. A byte-identical request delegates to the
    /// header-preserving replacement path without selecting a nested owner.
    #[allow(
        clippy::too_many_arguments,
        reason = "the physical rewrite keeps complete source and requested reference state explicit"
    )]
    /// Replace one payload and reorder the complete aggregate data-reference
    /// list while preserving every untouched `ArchiveInfo` byte.
    ///
    /// This physical primitive is intended for format owners whose repeated
    /// payload records and `MessageInfo.data_references` list have one-to-one
    /// positional semantics. The caller remains responsible for proving that
    /// the data references are safe to reorder. At most one nested
    /// `FieldInfo` may have a complete data-reference list matching the
    /// aggregate list when the order changes; that uniquely matching field is
    /// reordered with it. Multiple matching fields are rejected as ambiguous
    /// for an actual reorder, while a byte-identical no-op remains a header-
    /// preserving replacement. Unrelated nested metadata remains untouched.
    pub fn replace_message_reordering_data_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        reordered_data_references: &[u64],
        limits: Limits,
    ) -> Result<RawMessage> {
        if reordered_data_references.len() > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                reordered_data_references.len(),
                limits.max_metadata_items(),
            ));
        }
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let expected_data_references = current_info.data_references.as_slice();
        validate_reordered_data_reference_multiset(
            expected_data_references,
            reordered_data_references,
            index,
        )?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;
        if expected_data_references == reordered_data_references {
            return self.replace_message_preserving_header(index, message, limits);
        }
        let matching_field_info_count = current_info
            .field_infos
            .iter()
            .filter(|field| field.data_references.as_slice() == expected_data_references)
            .count();
        if matching_field_info_count > 1 {
            return Err(Error::invalid_archive(
                index,
                "data-reference reorder has ambiguous matching FieldInfo metadata",
            ));
        }

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_reorder_data_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            expected_data_references,
            reordered_data_references,
            limits,
        )?;
        drop(canonical_before);

        let rewritten_info = ArchiveInfo::decode(&rewritten_header, limits)?;
        verify_reordered_data_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            expected_data_references,
            reordered_data_references,
        )?;
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let source_provenance = retained_source_provenance(
            &rewritten_header,
            &canonical_after,
            retain_rewritten_header,
        )?;

        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Replace one payload and apply an exact object-reference transition under
    /// explicit resource limits.
    ///
    /// Both directional states are fully authorized before raw rewriting. The
    /// aggregate and every selected field list must be unique and nonzero;
    /// retained identifiers may be reordered, removals may occur anywhere,
    /// and newly introduced identifiers must form the final suffix. Selected
    /// `FieldInfo` occurrences are authorized by ordinal, path, and complete
    /// before/after lists. Untouched field lists retain their exact order and
    /// duplicate occurrences, but every retained value must remain a nonzero
    /// member of both aggregate states. All work and allocations are bounded
    /// by the header and metadata limits, and failure leaves the object
    /// unchanged.
    pub fn replace_message_transitioning_object_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: ObjectReferenceTransition<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.replace_message_transitioning_object_references_internal(
            index, message, transition, None, limits,
        )
    }

    /// Apply an exact aggregate/field transition and canonical field changes
    /// under explicit header, metadata, payload, and object limits.
    ///
    /// Existing selected fields must precede a removed suffix. Both the
    /// decoded metadata and retained raw bytes are checked before publication.
    /// On failure the original payload and archive header remain unchanged.
    pub fn replace_message_transitioning_object_references_with_canonical_fields_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: ObjectReferenceTransition<'_>,
        fields: CanonicalObjectReferenceFields<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.replace_message_transitioning_object_references_internal(
            index,
            message,
            transition,
            Some(fields),
            limits,
        )
    }

    fn replace_message_transitioning_object_references_internal(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: ObjectReferenceTransition<'_>,
        canonical_fields: Option<CanonicalObjectReferenceFields<'_>>,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;
        let prepared = prepare_object_reference_transition(
            current_info,
            transition,
            canonical_fields,
            limits,
            index,
        )?;
        preflight_transitioned_metadata_items(&self.archive_info, index, &prepared, limits)?;

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_transition_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            &prepared,
            limits,
        )?;
        drop(canonical_before);

        let rewritten_info = ArchiveInfo::decode(&rewritten_header, limits)?;
        verify_transitioned_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            &prepared,
        )?;
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let source_provenance = retained_source_provenance(
            &rewritten_header,
            &canonical_after,
            retain_rewritten_header,
        )?;

        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Replace one payload and apply a changed-cardinality data-reference
    /// transition under explicit resource limits.
    ///
    /// The operation is atomic. It validates the retained source header,
    /// rewrites only aggregate field 6 and explicitly selected nested
    /// `FieldInfo` field 5 records, decodes and compares the resulting neutral
    /// metadata, and publishes the new raw-header provenance only after every
    /// check succeeds. Values may be inserted, removed, or replaced at any
    /// position; zero identifiers and stale source authorizations are refused.
    pub fn replace_message_transitioning_data_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: DataReferenceTransition<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.replace_message_transitioning_data_references_preserving_header_internal(
            index, message, transition, None, limits,
        )
    }

    /// Replace one payload while inserting or removing one canonical
    /// data-reference `FieldInfo` for a source whose semantic owner currently
    /// exists only in the aggregate `MessageInfo.data_references` list.
    ///
    /// Insertion appends a minimal canonical field with the supplied path and
    /// the complete target aggregate list. Removal is the exact inverse: it
    /// accepts only the final canonical field produced by insertion. The
    /// target may restore any authorized aggregate-only source state. The
    /// ordinary transition method remains the right choice for existing field records or for
    /// aggregate-only removal that does not need to remove metadata. All
    /// validation, raw-header preservation, and publication remain atomic.
    pub fn replace_message_transitioning_data_references_with_canonical_field_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: CanonicalFieldDataReferenceTransition<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        let aggregate = DataReferenceTransition {
            aggregate_before: transition.aggregate_before,
            aggregate_after: transition.aggregate_after,
            fields: &[],
        };
        self.replace_message_transitioning_data_references_preserving_header_internal(
            index,
            message,
            aggregate,
            Some(transition),
            limits,
        )
    }

    fn replace_message_transitioning_data_references_preserving_header_internal(
        &mut self,
        index: usize,
        message: RawMessage,
        transition: DataReferenceTransition<'_>,
        canonical_field: Option<CanonicalFieldDataReferenceTransition<'_>>,
        limits: Limits,
    ) -> Result<RawMessage> {
        self.validate(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        if current_info.data_references.as_slice() != transition.aggregate_before {
            return Err(Error::invalid_archive(
                index,
                "aggregate source data references differ from transition authorization",
            ));
        }
        validate_data_reference_transition_lists(&transition, limits, index)?;
        let canonical_plan = canonical_field
            .map(|field| {
                prepare_canonical_field_data_reference_transition(
                    current_info,
                    transition,
                    field,
                    limits,
                    index,
                )
            })
            .transpose()?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let mut selected_fields = Vec::new();
        selected_fields
            .try_reserve_exact(current_info.field_infos.len())
            .map_err(|_| {
                Error::allocation(
                    "IWA FieldInfo data-reference transitions",
                    current_info.field_infos.len(),
                )
            })?;
        selected_fields.resize(current_info.field_infos.len(), None);
        for authorization in transition.fields {
            let current_field = current_info
                .field_infos
                .get(authorization.field_info_index)
                .ok_or_else(|| {
                    Error::invalid_archive(index, "FieldInfo transition index is out of bounds")
                })?;
            if current_field.path.path.as_slice() != authorization.expected_path
                || current_field.data_references.as_slice() != authorization.before
            {
                return Err(Error::invalid_archive(
                    index,
                    "FieldInfo source state differs from transition authorization",
                ));
            }
            if selected_fields[authorization.field_info_index]
                .replace(authorization)
                .is_some()
            {
                return Err(Error::invalid_archive(
                    index,
                    "FieldInfo data-reference transition occurrence is selected more than once",
                ));
            }
        }
        for (field_index, current_field) in current_info.field_infos.iter().enumerate() {
            let selected = selected_fields.get(field_index).and_then(Option::as_ref);
            let removed = canonical_plan.is_some_and(|plan| {
                matches!(
                    plan,
                    CanonicalFieldDataReferencePlan::Remove {
                        field_info_index,
                        ..
                    } if field_info_index == field_index
                )
            });
            if !removed
                && selected.is_none()
                && current_field
                    .data_references
                    .iter()
                    .any(|identifier| !transition.aggregate_after.contains(identifier))
            {
                return Err(Error::invalid_archive(
                    index,
                    "unselected FieldInfo retains a removed data reference",
                ));
            }
            if let Some(selected) = selected
                && selected
                    .after
                    .iter()
                    .any(|identifier| !transition.aggregate_after.contains(identifier))
            {
                return Err(Error::invalid_archive(
                    index,
                    "selected FieldInfo target is outside aggregate data references",
                ));
            }
        }

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match self.source_headers() {
            Some((original, canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_transition_data_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            transition,
            &selected_fields,
            canonical_plan,
            limits,
        )?;
        let rewritten_info = ArchiveInfo::decode(&rewritten_header, limits)?;
        verify_transitioned_data_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            transition,
            canonical_plan,
        )?;
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let source_provenance = retained_source_provenance(
            &rewritten_header,
            &canonical_after,
            retain_rewritten_header,
        )?;
        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.source_provenance = source_provenance;
        Ok(old)
    }

    /// Append one payload under explicit resource limits.
    pub fn push_message(&mut self, message: RawMessage, limits: Limits) -> Result<()> {
        self.validate(limits)?;
        let next_count = self
            .messages
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "message count overflow"))?;
        if next_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                next_count,
                limits.max_messages_per_object(),
            ));
        }
        check_message_length(message.data.len(), limits)?;
        let length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        let message_info = MessageInfo::conventional(message.type_, length);
        self.archive_info
            .message_infos
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA message metadata", 1))?;
        self.messages
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA object messages", 1))?;
        self.archive_info.message_infos.push(message_info);
        self.messages.push(message);
        if let Err(error) = self.validate(limits) {
            self.archive_info.message_infos.pop();
            self.messages.pop();
            return Err(error);
        }
        Ok(())
    }

    /// Remove one payload and its metadata, returning the payload if present.
    pub fn remove_message(&mut self, index: usize) -> Option<RawMessage> {
        if index >= self.messages.len() || index >= self.archive_info.message_infos.len() {
            return None;
        }
        self.archive_info.message_infos.remove(index);
        Some(self.messages.remove(index))
    }
}

/// A parsed, mutable decompressed IWA component archive.
#[derive(Debug, Clone, PartialEq)]
pub struct Archive {
    pub objects: Vec<ArchiveObject>,
}

/// Source-derived resource requirements for one prepared archive encoding.
///
/// The byte fields are cumulative ledger quantities: a caller that retains
/// the output and charges all temporary projection buffers can use this value
/// before execution.  No allocator-dependent capacity estimate is included;
/// every collection quantity is derived from the immutable neutral metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveEncodingRequirements {
    encoded_bytes: usize,
    object_count: usize,
    message_count: usize,
    header_bytes: usize,
    payload_bytes: usize,
    work_bytes: usize,
    scratch_bytes: usize,
    retained_bytes: usize,
    identifier_index_entries: usize,
    allocation_events: usize,
}

impl ArchiveEncodingRequirements {
    #[must_use]
    pub const fn encoded_bytes(self) -> usize {
        self.encoded_bytes
    }

    #[must_use]
    pub const fn object_count(self) -> usize {
        self.object_count
    }

    #[must_use]
    pub const fn message_count(self) -> usize {
        self.message_count
    }

    #[must_use]
    pub const fn header_bytes(self) -> usize {
        self.header_bytes
    }

    #[must_use]
    pub const fn payload_bytes(self) -> usize {
        self.payload_bytes
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn identifier_index_entries(self) -> usize {
        self.identifier_index_entries
    }

    /// Output, identifier-index, projection and header allocations counted by
    /// the prepared execution path.  A zero-sized archive performs no heap
    /// allocation.
    #[must_use]
    pub const fn allocation_events(self) -> usize {
        self.allocation_events
    }
}

/// Allocation-free archive serialization admission bound to an immutable
/// archive and one validated limit profile.
#[derive(Debug, Clone, Copy)]
pub struct ArchiveEncodingPlan<'a> {
    archive: &'a Archive,
    limits: Limits,
    requirements: ArchiveEncodingRequirements,
}

impl ArchiveEncodingPlan<'_> {
    #[must_use]
    pub const fn requirements(&self) -> ArchiveEncodingRequirements {
        self.requirements
    }

    /// Execute the prepared route.  The borrowed archive cannot change while
    /// this plan exists, so the source-derived requirements cannot become
    /// stale between admission and publication.
    pub fn execute(self) -> Result<Vec<u8>> {
        encode_prepared_archive(self.archive, self.limits, self.requirements)
    }
}

impl Archive {
    /// Construct an empty archive.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            objects: Vec::new(),
        }
    }

    /// Parse a decompressed IWA stream under explicit resource limits.
    pub fn parse(data: &[u8], limits: Limits) -> Result<Self> {
        Self::parse_internal_with_limits(data, limits, None)
    }

    /// Parse only objects with one archive identifier from a decompressed IWA
    /// stream under explicit resource limits.
    ///
    /// Every object in the source is still walked: its length framing,
    /// `ArchiveInfo`, message lengths, aggregate message count, and payload
    /// extent are validated against `limits`. Only payload bytes belonging to
    /// objects whose identifier equals `identifier` are copied into the
    /// returned archive. This is intended for bounded structural probes such
    /// as application detection, where unrelated objects must remain opaque.
    ///
    /// The returned archive is a selected view, not a lossless replacement for
    /// the source stream. Callers that may edit or serialize a component must
    /// use [`Self::parse`] so every payload remains available for preservation.
    pub fn parse_objects_with_identifier(
        data: &[u8],
        identifier: u64,
        limits: Limits,
    ) -> Result<Self> {
        Self::parse_internal_with_limits(data, limits, Some(identifier))
    }

    /// Prepare an allocation-free, source-bound serialization route.
    ///
    /// Preparation validates identifiers, metadata, framing and retained raw
    /// headers without materializing compatibility projections or output.
    /// Identifier presence is checked during preparation; duplicate identity
    /// is checked by the exact fallibly reserved index immediately before
    /// execution so preparation itself remains allocation-free.
    /// `visitor` receives source-derived collection chunks immediately before
    /// each bounded metadata, retained-header, payload, and projection-sizing
    /// traversal. Its associated error is preserved in
    /// [`ArchiveEncodingPreparationError::Admission`]. The returned
    /// requirements describe execution-only projection and output quantities;
    /// callers must not charge preparation chunks a second time from that
    /// report.
    pub fn prepare_encoding<V>(
        &self,
        limits: Limits,
        visitor: &mut V,
    ) -> ArchiveEncodingPreparationResult<ArchiveEncodingPlan<'_>, V::Error>
    where
        V: ArchiveEncodingPreparationVisitor,
    {
        let requirements = prepare_archive_encoding_requirements(&self.objects, limits, visitor)?;
        Ok(ArchiveEncodingPlan {
            archive: self,
            limits,
            requirements,
        })
    }

    fn parse_internal_with_limits(
        data: &[u8],
        limits: Limits,
        selected_identifier: Option<u64>,
    ) -> Result<Self> {
        if data.len() > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                data.len(),
                limits.max_archive_bytes(),
            ));
        }

        // Archive byte length is not an object-count hint: a single message
        // may consume almost the entire input. Reserving one slot per input
        // byte would turn a large object into a second, avoidable allocation
        // spike. Grow the object list only as validated objects are found.
        let mut objects = Vec::new();
        // A selected parse does not retain discarded objects, but it must
        // retain their identities long enough to preserve the archive-wide
        // uniqueness invariant enforced by a full parse.
        let mut discarded_identifiers = selected_identifier.map(|_| HashSet::new());
        let mut cursor = 0usize;
        let mut object_count = 0usize;
        let mut total_messages = 0usize;
        while cursor < data.len() {
            let object_start = cursor;
            if object_count >= limits.max_objects() {
                return Err(limit(
                    LimitKind::Objects,
                    object_count + 1,
                    limits.max_objects(),
                ));
            }
            object_count += 1;
            let (header_length_u64, prefix_length) = decode_varint(&data[cursor..])?;
            cursor = cursor
                .checked_add(prefix_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header offset overflow"))?;
            let header_length = usize::try_from(header_length_u64)
                .map_err(|_| Error::invalid_archive(object_start, "header length exceeds usize"))?;
            check_header_length(header_length, limits)?;
            let header_end = cursor
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header range overflow"))?;
            let header = data.get(cursor..header_end).ok_or_else(|| {
                Error::invalid_archive(object_start, "truncated ArchiveInfo header")
            })?;
            let header_admission =
                preflight_header_with_admission(header, HeaderKind::ArchiveInfo, limits)?;
            let (archive_info, archive_info_admission) =
                ArchiveInfo::decode_with_preflight_and_admission(
                    header,
                    limits,
                    header_admission.preflight,
                )?;
            if let Some(identifiers) = discarded_identifiers.as_mut() {
                let identifier = archive_info.identifier.ok_or_else(|| {
                    Error::invalid_archive(object_start, "object is missing its archive identifier")
                })?;
                identifiers
                    .try_reserve(1)
                    .map_err(|_| Error::allocation("IWA archive identifiers", 1))?;
                if !identifiers.insert(identifier) {
                    return Err(Error::invalid_archive(
                        object_start,
                        "duplicate object identifier",
                    ));
                }
            }
            let selected = selected_identifier
                .is_none_or(|identifier| archive_info.identifier == Some(identifier));
            let source_provenance = if selected {
                let canonical_header = encode_archive_info(&archive_info, limits)?;
                if header == canonical_header.as_slice() {
                    None
                } else {
                    Some(SourceProvenance::from_headers(
                        header,
                        &canonical_header,
                        Some(header_admission),
                        Some(archive_info_admission),
                        Some(limits),
                    )?)
                }
            } else {
                None
            };
            cursor = header_end;
            add_limited(
                &mut total_messages,
                archive_info.message_infos.len(),
                LimitKind::Messages,
                limits.max_messages(),
            )?;

            let object_prefix = prefix_length
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "object prefix overflow"))?;
            if object_prefix > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_prefix,
                    limits.max_object_bytes(),
                ));
            }
            let mut payload_length = 0usize;
            for message_info in &archive_info.message_infos {
                let length = usize::try_from(message_info.length)
                    .map_err(|_| Error::invalid_archive(cursor, "message length exceeds usize"))?;
                add_limited(
                    &mut payload_length,
                    length,
                    LimitKind::ObjectBytes,
                    limits.max_object_bytes() - object_prefix,
                )?;
            }
            let payload_end = cursor
                .checked_add(payload_length)
                .ok_or_else(|| Error::invalid_archive(cursor, "payload range overflow"))?;
            if payload_end > data.len() {
                return Err(Error::invalid_archive(cursor, "truncated message payload"));
            }

            let payload_start = cursor;
            if selected {
                let mut messages = Vec::new();
                messages
                    .try_reserve_exact(archive_info.message_infos.len())
                    .map_err(|_| {
                        Error::allocation("IWA object messages", archive_info.message_infos.len())
                    })?;
                for message_info in &archive_info.message_infos {
                    let length = usize::try_from(message_info.length).map_err(|_| {
                        Error::invalid_archive(cursor, "message length exceeds usize")
                    })?;
                    let end = cursor
                        .checked_add(length)
                        .ok_or_else(|| Error::invalid_archive(cursor, "message range overflow"))?;
                    let message_data = data.get(cursor..end).ok_or_else(|| {
                        Error::invalid_archive(cursor, "truncated message payload")
                    })?;
                    let mut owned = Vec::new();
                    owned
                        .try_reserve_exact(length)
                        .map_err(|_| Error::allocation("IWA message payload", length))?;
                    owned.extend_from_slice(message_data);
                    messages.push(RawMessage {
                        type_: message_info.type_,
                        data: owned,
                    });
                    cursor = end;
                }
                objects.push(ArchiveObject {
                    archive_info,
                    messages,
                    header_offset: u64::try_from(object_start)
                        .map_err(|_| Error::invalid_archive(object_start, "offset exceeds u64"))?,
                    header_length: u64::try_from(prefix_length + header_length)
                        .map_err(|_| Error::invalid_archive(object_start, "header exceeds u64"))?,
                    data_offset: u64::try_from(payload_start)
                        .map_err(|_| Error::invalid_archive(payload_start, "offset exceeds u64"))?,
                    data_length: u64::try_from(payload_length).map_err(|_| {
                        Error::invalid_archive(payload_start, "payload exceeds u64")
                    })?,
                    source_provenance,
                });
            } else {
                cursor = payload_end;
            }
        }
        let archive = Self { objects };
        archive.validate(limits)?;
        Ok(archive)
    }

    /// Validate canonical object framing against the decompressed source.
    ///
    /// This opt-in check requires every object-header length prefix to use the
    /// shortest protobuf varint encoding. It also requires the parsed source
    /// provenance to describe one exact, contiguous partition of
    /// `decompressed_source`: each header starts where the preceding object's
    /// payload ends, each `header_length` includes exactly its prefix and raw
    /// `ArchiveInfo`, each `data_offset` immediately follows that header, and
    /// the final payload ends at the end of the supplied source.
    ///
    /// Header contents are neither decoded nor re-encoded, so unknown fields,
    /// field ordering, and noncanonical encodings *inside* an `ArchiveInfo`
    /// remain accepted and untouched. [`Self::parse`] deliberately remains
    /// permissive about overlong object length prefixes; format mutations can
    /// call this method when they require an unambiguous byte-splice boundary.
    ///
    /// The check performs no allocation and runs in O(objects + messages)
    /// time, visiting metadata only to verify each recorded payload extent.
    pub fn validate_canonical_object_framing(
        &self,
        decompressed_source: &[u8],
        limits: Limits,
    ) -> Result<()> {
        if decompressed_source.len() > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                decompressed_source.len(),
                limits.max_archive_bytes(),
            ));
        }
        if self.objects.len() > limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                self.objects.len(),
                limits.max_objects(),
            ));
        }
        let source_length = u64::try_from(decompressed_source.len())
            .map_err(|_| Error::invalid_archive(0, "source length exceeds u64"))?;
        let mut expected_object_offset = 0u64;
        let mut message_count = 0usize;

        for (object_index, object) in self.objects.iter().enumerate() {
            let object_messages = object.archive_info.message_infos.len();
            if object_messages > limits.max_messages_per_object() {
                return Err(limit(
                    LimitKind::MessagesPerObject,
                    object_messages,
                    limits.max_messages_per_object(),
                ));
            }
            message_count = message_count
                .checked_add(object_messages)
                .ok_or_else(|| Error::invalid_archive(object_index, "message count overflow"))?;
            if message_count > limits.max_messages() {
                return Err(limit(
                    LimitKind::Messages,
                    message_count,
                    limits.max_messages(),
                ));
            }
            if object.header_offset != expected_object_offset {
                return Err(Error::invalid_archive(
                    object_index,
                    "object header offset does not follow the preceding payload",
                ));
            }
            let header_offset = usize::try_from(object.header_offset).map_err(|_| {
                Error::invalid_archive(object_index, "object header offset exceeds usize")
            })?;
            let remaining = decompressed_source.get(header_offset..).ok_or_else(|| {
                Error::invalid_archive(header_offset, "object header offset exceeds source")
            })?;
            let (raw_header_length, prefix_length) = decode_varint(remaining)?;
            let header_length = usize::try_from(raw_header_length).map_err(|_| {
                Error::invalid_archive(header_offset, "header length exceeds usize")
            })?;
            if header_length > limits.max_header_bytes() {
                return Err(limit(
                    LimitKind::HeaderBytes,
                    header_length,
                    limits.max_header_bytes(),
                ));
            }
            // A decoded unsigned varint is overlong exactly when its final
            // seven-bit group is zero and it has more than one byte.
            if prefix_length > 1 && remaining[prefix_length - 1] == 0 {
                return Err(Error::invalid_archive(
                    header_offset,
                    "object header length prefix is not canonical",
                ));
            }

            let framed_header_length = prefix_length
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(header_offset, "header range overflow"))?;
            let data_offset = header_offset
                .checked_add(framed_header_length)
                .ok_or_else(|| Error::invalid_archive(header_offset, "header range overflow"))?;
            if data_offset > decompressed_source.len() {
                return Err(Error::invalid_archive(
                    header_offset,
                    "truncated ArchiveInfo header",
                ));
            }
            let framed_header_length_u64 = u64::try_from(framed_header_length).map_err(|_| {
                Error::invalid_archive(header_offset, "framed header length exceeds u64")
            })?;
            let data_offset_u64 = u64::try_from(data_offset)
                .map_err(|_| Error::invalid_archive(header_offset, "data offset exceeds u64"))?;
            if object.header_length != framed_header_length_u64 {
                return Err(Error::invalid_archive(
                    header_offset,
                    "recorded header length does not match source framing",
                ));
            }
            if object.data_offset != data_offset_u64 {
                return Err(Error::invalid_archive(
                    header_offset,
                    "recorded data offset does not follow the source header",
                ));
            }

            let mut metadata_payload_length = 0u64;
            for message_info in &object.archive_info.message_infos {
                let message_length = usize::try_from(message_info.length).map_err(|_| {
                    Error::invalid_archive(data_offset, "message length exceeds usize")
                })?;
                if message_length > limits.max_message_bytes() {
                    return Err(limit(
                        LimitKind::MessageBytes,
                        message_length,
                        limits.max_message_bytes(),
                    ));
                }
                // The admitted count is at most MAX_MESSAGES_PER_OBJECT
                // (100,000), and each length is u32: their sum fits u64.
                metadata_payload_length += u64::from(message_info.length);
            }
            if object.data_length != metadata_payload_length {
                return Err(Error::invalid_archive(
                    data_offset,
                    "recorded data length does not match message metadata",
                ));
            }
            // data_offset was matched to a source offset no greater than
            // MAX_ARCHIVE_BYTES. Adding the admitted u32-length sum fits u64.
            let data_end = object.data_offset + metadata_payload_length;
            let object_length = usize::try_from(data_end - object.header_offset).map_err(|_| {
                Error::invalid_archive(header_offset, "object length exceeds usize")
            })?;
            if object_length > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_length,
                    limits.max_object_bytes(),
                ));
            }
            if data_end > source_length {
                return Err(Error::invalid_archive(
                    data_offset,
                    "truncated message payload",
                ));
            }
            expected_object_offset = data_end;
        }

        if expected_object_offset != source_length {
            let offset = usize::try_from(expected_object_offset).unwrap_or(usize::MAX);
            return Err(Error::invalid_archive(
                offset,
                "source contains bytes not described by archive objects",
            ));
        }
        Ok(())
    }

    /// Serialize this archive under explicit resource limits.
    pub fn to_bytes(&self, limits: Limits) -> Result<Vec<u8>> {
        let expected_length = self.encoded_len(limits)?;
        validate_object_set(self.objects.iter(), limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(expected_length)
            .map_err(|_| Error::allocation("IWA archive output", expected_length))?;
        for object in &self.objects {
            let mut info = object.archive_info.clone();
            for (message_info, message) in info.message_infos.iter_mut().zip(&object.messages) {
                message_info.type_ = message.type_;
                message_info.length = u32::try_from(message.data.len())
                    .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            }
            let canonical_header = encode_archive_info(&info, limits)?;
            let header = match object.source_headers() {
                Some((original, canonical)) if canonical == canonical_header.as_slice() => original,
                _ => canonical_header.as_slice(),
            };
            let header_length = header.len();
            let prefix_length = varint_len(header_length)?;
            let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
                total
                    .checked_add(message.data.len())
                    .ok_or_else(|| Error::invalid_archive(0, "payload length overflow"))
            })?;
            let object_length = prefix_length
                .checked_add(header_length)
                .and_then(|length| length.checked_add(payload_length))
                .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
            if object_length > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_length,
                    limits.max_object_bytes(),
                ));
            }
            let next_length = output
                .len()
                .checked_add(object_length)
                .ok_or_else(|| Error::invalid_archive(0, "archive length overflow"))?;
            if next_length > limits.max_archive_bytes() {
                return Err(limit(
                    LimitKind::ArchiveBytes,
                    next_length,
                    limits.max_archive_bytes(),
                ));
            }
            output
                .try_reserve_exact(object_length)
                .map_err(|_| Error::allocation("IWA archive output", object_length))?;
            let mut prefix = [0u8; MAX_VARINT_BYTES];
            let prefix = encode_varint(
                u64::try_from(header_length)
                    .map_err(|_| Error::invalid_archive(0, "header length exceeds u64"))?,
                &mut prefix,
            );
            output.extend_from_slice(prefix);
            output.extend_from_slice(header);
            for message in &object.messages {
                output.extend_from_slice(&message.data);
            }
        }
        debug_assert_eq!(output.len(), expected_length);
        Ok(output)
    }

    /// Return the exact decompressed IWA length produced under `limits`.
    ///
    /// The calculation validates object/message/header/archive ceilings,
    /// checked arithmetic, duplicate identifiers, and retained raw-header
    /// selection without allocating an encoded header or output buffer.
    /// Archives of more than 32 objects reserve one fallible identifier index
    /// so duplicate detection stays `O(n log n)`.
    pub fn encoded_len(&self, limits: Limits) -> Result<usize> {
        encoded_archive_len_preflight(&self.objects, limits)
    }

    /// Validate all objects under explicit resource limits.
    pub fn validate(&self, limits: Limits) -> Result<()> {
        validate_object_set(self.objects.iter(), limits)
    }

    /// Insert an object under explicit resource limits.
    pub fn insert_object(&mut self, object: ArchiveObject, limits: Limits) -> Result<()> {
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if self.object(identifier).is_some() {
            return Err(Error::invalid_archive(0, "duplicate object identifier"));
        }
        validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
        self.objects
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA archive objects", 1))?;
        self.objects.push(object);
        Ok(())
    }

    /// Append a batch of new objects under explicit resource limits.
    ///
    /// This is the atomic batch counterpart to [`Self::insert_object`]. It
    /// validates the projected full object set—including object, message,
    /// metadata, header, and archive byte ceilings—then reserves exactly the
    /// additional object slots before performing an infallible append. An
    /// empty batch is an exact no-op.
    pub fn append_objects(&mut self, objects: Vec<ArchiveObject>, limits: Limits) -> Result<()> {
        if objects.is_empty() {
            return Ok(());
        }
        if objects
            .iter()
            .any(|object| object.archive_info.identifier == Some(0))
        {
            return Err(Error::invalid_archive(
                0,
                "appended object identifier is zero",
            ));
        }
        validate_object_set(self.objects.iter().chain(&objects), limits)?;
        validate_archive_wire_size(self.objects.iter().chain(&objects), limits)?;
        self.objects
            .try_reserve_exact(objects.len())
            .map_err(|_| Error::allocation("IWA archive objects", objects.len()))?;
        self.objects.extend(objects);
        Ok(())
    }

    /// Find an object by its archive identifier.
    #[must_use]
    pub fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    /// Find an object mutably by its archive identifier.
    #[must_use]
    pub fn object_mut(&mut self, identifier: u64) -> Option<&mut ArchiveObject> {
        self.objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    /// Insert or replace an object under explicit resource limits.
    pub fn upsert_object(
        &mut self,
        object: ArchiveObject,
        limits: Limits,
    ) -> Result<Option<ArchiveObject>> {
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if let Some(index) = self
            .objects
            .iter()
            .position(|current| current.archive_info.identifier == Some(identifier))
        {
            validate_object_set(
                self.objects.iter().enumerate().map(
                    |(current, item)| {
                        if current == index { &object } else { item }
                    },
                ),
                limits,
            )?;
            return Ok(Some(std::mem::replace(&mut self.objects[index], object)));
        }

        validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
        self.objects
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA archive objects", 1))?;
        self.objects.push(object);
        Ok(None)
    }

    /// Remove an object by its archive identifier.
    pub fn remove_object(&mut self, identifier: u64) -> Option<ArchiveObject> {
        let index = self
            .objects
            .iter()
            .position(|object| object.archive_info.identifier == Some(identifier))?;
        Some(self.objects.remove(index))
    }

    /// Remove exactly one object after validating the complete archive.
    ///
    /// Zero, missing, and duplicate identifiers are rejected. Validation and
    /// target resolution complete before mutation, so every error preserves
    /// the archive exactly. Graph ownership remains the caller's concern.
    pub fn remove_object_checked(
        &mut self,
        identifier: u64,
        limits: Limits,
    ) -> Result<ArchiveObject> {
        if identifier == 0 {
            return Err(Error::invalid_archive(0, "object identifier is zero"));
        }
        self.validate(limits)?;
        let mut selected = None;
        for (index, object) in self.objects.iter().enumerate() {
            if object.archive_info.identifier != Some(identifier) {
                continue;
            }
            if selected.replace(index).is_some() {
                return Err(Error::invalid_archive(
                    index,
                    "object identifier is duplicated",
                ));
            }
        }
        let index = selected.ok_or_else(|| Error::invalid_archive(0, "object is missing"))?;
        Ok(self.objects.remove(index))
    }
}

impl Default for Archive {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug)]
enum HeaderFieldRewrite {
    Retain,
    Remove,
    Varint(u64),
    LengthDelimited(Vec<u8>),
    /// Complete encoded field bytes, including key and value.
    Raw(Vec<u8>),
}

#[derive(Clone, Copy)]
enum ReferenceRemovals<'a> {
    None,
    Selected(&'a HashSet<u64>),
    All,
}

struct PreparedReferenceList<'a> {
    before: &'a [u64],
    retained_after: &'a [u64],
    appended_after: &'a [u64],
    after_members: HashSet<u64>,
}

struct PreparedFieldTransition<'a> {
    authorization: &'a FieldObjectReferenceTransition<'a>,
    references: PreparedReferenceList<'a>,
}

struct PreparedObjectReferenceTransition<'a> {
    aggregate: PreparedReferenceList<'a>,
    fields: Vec<Option<PreparedFieldTransition<'a>>>,
    canonical_fields: Option<CanonicalObjectReferenceFields<'a>>,
    retained_field_count: usize,
}

#[derive(Debug, Clone, Copy)]
enum CanonicalFieldDataReferencePlan<'a> {
    Insert {
        path: &'a [u32],
    },
    Remove {
        field_info_index: usize,
        path: &'a [u32],
    },
}

impl ReferenceRemovals<'_> {
    const fn is_none(self) -> bool {
        matches!(self, Self::None)
    }

    fn contains(self, identifier: u64) -> bool {
        match self {
            Self::None => false,
            Self::Selected(identifiers) => identifiers.contains(&identifier),
            Self::All => true,
        }
    }
}

fn prepare_reference_list<'a>(
    before: &'a [u64],
    after: &'a [u64],
    limits: Limits,
    message_index: usize,
    context: &'static str,
) -> Result<PreparedReferenceList<'a>> {
    let authorization_items = before
        .len()
        .checked_add(after.len())
        .ok_or_else(|| Error::invalid_archive(message_index, "reference transition overflow"))?;
    if authorization_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            authorization_items,
            limits.max_metadata_items(),
        ));
    }
    let mut before_members = HashSet::new();
    before_members
        .try_reserve(before.len())
        .map_err(|_| Error::allocation(context, before.len()))?;
    if before
        .iter()
        .any(|identifier| *identifier == 0 || !before_members.insert(*identifier))
    {
        return Err(Error::invalid_archive(
            message_index,
            "source reference transition list contains zero or duplicates",
        ));
    }
    let mut after_members = HashSet::new();
    after_members
        .try_reserve(after.len())
        .map_err(|_| Error::allocation(context, after.len()))?;
    if after
        .iter()
        .any(|identifier| *identifier == 0 || !after_members.insert(*identifier))
    {
        return Err(Error::invalid_archive(
            message_index,
            "target reference transition list contains zero or duplicates",
        ));
    }
    let retained_count = after
        .iter()
        .take_while(|identifier| before_members.contains(identifier))
        .count();
    if after.get(retained_count..).is_some_and(|suffix| {
        suffix
            .iter()
            .any(|identifier| before_members.contains(identifier))
    }) {
        return Err(Error::invalid_archive(
            message_index,
            "new references do not form the final transition suffix",
        ));
    }
    let retained_before_count = before
        .iter()
        .filter(|identifier| after_members.contains(identifier))
        .count();
    if retained_before_count != retained_count {
        return Err(Error::invalid_archive(
            message_index,
            "retained reference transition cardinality changed",
        ));
    }
    let retained_after = after
        .get(..retained_count)
        .ok_or_else(|| Error::invalid_archive(message_index, "retained reference range failed"))?;
    let appended_after = after
        .get(retained_count..)
        .ok_or_else(|| Error::invalid_archive(message_index, "appended reference range failed"))?;
    Ok(PreparedReferenceList {
        before,
        retained_after,
        appended_after,
        after_members,
    })
}

fn prepare_object_reference_transition<'a>(
    current: &MessageInfo,
    transition: ObjectReferenceTransition<'a>,
    canonical_fields: Option<CanonicalObjectReferenceFields<'a>>,
    limits: Limits,
    message_index: usize,
) -> Result<PreparedObjectReferenceTransition<'a>> {
    if current.object_references.as_slice() != transition.aggregate_before {
        return Err(Error::invalid_archive(
            message_index,
            "aggregate source references differ from transition authorization",
        ));
    }
    let aggregate = prepare_reference_list(
        transition.aggregate_before,
        transition.aggregate_after,
        limits,
        message_index,
        "IWA aggregate reference transition",
    )?;
    let mut authorization_items = transition.fields.len();
    if let Some(canonical) = canonical_fields {
        for field in canonical.definitions() {
            authorization_items = authorization_items
                .checked_add(1)
                .and_then(|count| count.checked_add(field.path.len()))
                .and_then(|count| count.checked_add(field.references.len()))
                .ok_or_else(|| {
                    Error::invalid_archive(
                        message_index,
                        "canonical FieldInfo authorization overflow",
                    )
                })?;
        }
    }
    for field in transition.fields {
        authorization_items = authorization_items
            .checked_add(field.expected_path.len())
            .and_then(|count| count.checked_add(field.before.len()))
            .and_then(|count| count.checked_add(field.after.len()))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "FieldInfo transition authorization overflow")
            })?;
    }
    if authorization_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            authorization_items,
            limits.max_metadata_items(),
        ));
    }
    let retained_field_count = prepare_canonical_object_reference_fields(
        current,
        &transition,
        canonical_fields,
        message_index,
    )?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(current.field_infos.len())
        .map_err(|_| {
            Error::allocation(
                "IWA FieldInfo reference transitions",
                current.field_infos.len(),
            )
        })?;
    fields.resize_with(current.field_infos.len(), || None);
    for authorization in transition.fields {
        if authorization.field_info_index >= retained_field_count {
            return Err(Error::invalid_archive(
                message_index,
                "selected FieldInfo overlaps the removed canonical suffix",
            ));
        }
        let current_field = current
            .field_infos
            .get(authorization.field_info_index)
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "FieldInfo transition index is out of bounds")
            })?;
        if current_field.path.as_slice() != authorization.expected_path
            || current_field.object_references.as_slice() != authorization.before
        {
            return Err(Error::invalid_archive(
                message_index,
                "FieldInfo source state differs from transition authorization",
            ));
        }
        let references = prepare_reference_list(
            authorization.before,
            authorization.after,
            limits,
            message_index,
            "IWA FieldInfo reference transition",
        )?;
        let slot = fields
            .get_mut(authorization.field_info_index)
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "FieldInfo transition index is out of bounds")
            })?;
        if slot
            .replace(PreparedFieldTransition {
                authorization,
                references,
            })
            .is_some()
        {
            return Err(Error::invalid_archive(
                message_index,
                "FieldInfo transition occurrence is selected more than once",
            ));
        }
    }
    for (field_index, current_field) in current
        .field_infos
        .iter()
        .take(retained_field_count)
        .enumerate()
    {
        let selected = fields.get(field_index).and_then(Option::as_ref);
        let (before, after) = selected.map_or(
            (
                current_field.object_references.as_slice(),
                current_field.object_references.as_slice(),
            ),
            |field| (field.authorization.before, field.authorization.after),
        );
        if selected.is_none() {
            for identifier in before {
                if *identifier == 0
                    || !transition.aggregate_before.contains(identifier)
                    || !transition.aggregate_after.contains(identifier)
                {
                    return Err(Error::invalid_archive(
                        message_index,
                        "retained FieldInfo references are not aggregate members",
                    ));
                }
            }
            continue;
        }
        let mut local = HashSet::new();
        local
            .try_reserve(before.len().max(after.len()))
            .map_err(|_| {
                Error::allocation(
                    "IWA FieldInfo local references",
                    before.len().max(after.len()),
                )
            })?;
        for identifier in before {
            if *identifier == 0
                || !local.insert(*identifier)
                || !transition.aggregate_before.contains(identifier)
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "FieldInfo source references are not unique aggregate members",
                ));
            }
        }
        local.clear();
        for identifier in after {
            if *identifier == 0
                || !local.insert(*identifier)
                || !transition.aggregate_after.contains(identifier)
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "FieldInfo target references are not unique aggregate members",
                ));
            }
        }
    }
    Ok(PreparedObjectReferenceTransition {
        aggregate,
        fields,
        canonical_fields,
        retained_field_count,
    })
}

fn canonical_object_field_matches(
    field: &FieldInfo,
    expected: CanonicalObjectReferenceField<'_>,
) -> bool {
    field.path.as_slice() == expected.path
        && field.r#type == Some(FieldType::ObjectReference)
        && field.object_references.as_slice() == expected.references
        && field.data_references.is_empty()
        && field.unknown_field_rule.is_none()
        && field.known_field_rule.is_none()
        && field.known_field_version.is_empty()
        && field.known_field_feature_identifier.is_none()
}

fn prepare_canonical_object_reference_fields(
    current: &MessageInfo,
    transition: &ObjectReferenceTransition<'_>,
    canonical: Option<CanonicalObjectReferenceFields<'_>>,
    message_index: usize,
) -> Result<usize> {
    let Some(canonical) = canonical else {
        return Ok(current.field_infos.len());
    };
    let definitions = canonical.definitions();
    let retained = match canonical {
        CanonicalObjectReferenceFields::Insert(_) => current.field_infos.len(),
        CanonicalObjectReferenceFields::Remove(_) => current
            .field_infos
            .len()
            .checked_sub(definitions.len())
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "canonical FieldInfo suffix is missing")
            })?,
    };
    let aggregate = match canonical {
        CanonicalObjectReferenceFields::Insert(_) => transition.aggregate_after,
        CanonicalObjectReferenceFields::Remove(_) => transition.aggregate_before,
    };
    let mut members = HashSet::new();
    members
        .try_reserve(aggregate.len())
        .map_err(|_| Error::allocation("IWA canonical aggregate references", aggregate.len()))?;
    members.extend(aggregate.iter().copied());
    let mut paths = HashSet::new();
    paths
        .try_reserve(definitions.len())
        .map_err(|_| Error::allocation("IWA canonical FieldInfo paths", definitions.len()))?;
    let mut references = HashSet::new();
    for (offset, definition) in definitions.iter().copied().enumerate() {
        if !paths.insert(definition.path) || definition.references.is_empty() {
            return Err(Error::invalid_archive(
                message_index,
                "canonical FieldInfo paths must be distinct and references nonempty",
            ));
        }
        references.clear();
        references
            .try_reserve(definition.references.len())
            .map_err(|_| {
                Error::allocation(
                    "IWA canonical field references",
                    definition.references.len(),
                )
            })?;
        for identifier in definition.references {
            if *identifier == 0 || !references.insert(*identifier) || !members.contains(identifier)
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "canonical FieldInfo references must be unique nonzero aggregate members",
                ));
            }
        }
        if matches!(canonical, CanonicalObjectReferenceFields::Remove(_))
            && !current
                .field_infos
                .get(retained + offset)
                .is_some_and(|field| canonical_object_field_matches(field, definition))
        {
            return Err(Error::invalid_archive(
                message_index,
                "FieldInfo suffix differs from canonical removal authorization",
            ));
        }
    }
    for field in current.field_infos.iter().take(retained) {
        if paths.contains(field.path.as_slice()) {
            return Err(Error::invalid_archive(
                message_index,
                "canonical FieldInfo path already exists outside the removed suffix",
            ));
        }
    }
    Ok(retained)
}

fn preflight_transitioned_metadata_items(
    archive: &ArchiveInfo,
    message_index: usize,
    transition: &PreparedObjectReferenceTransition<'_>,
    limits: Limits,
) -> Result<()> {
    let mut total = archive.message_infos.len();
    for (index, info) in archive.message_infos.iter().enumerate() {
        let mut count = info.metadata_item_count()?;
        if index == message_index {
            count = count
                .checked_sub(transition.aggregate.before.len())
                .and_then(|value| value.checked_add(transition.aggregate.retained_after.len()))
                .and_then(|value| value.checked_add(transition.aggregate.appended_after.len()))
                .ok_or_else(|| {
                    Error::invalid_archive(index, "target metadata cardinality overflow")
                })?;
            for field in transition.fields.iter().flatten() {
                count = count
                    .checked_sub(field.references.before.len())
                    .and_then(|value| value.checked_add(field.references.retained_after.len()))
                    .and_then(|value| value.checked_add(field.references.appended_after.len()))
                    .ok_or_else(|| {
                        Error::invalid_archive(index, "target field cardinality overflow")
                    })?;
            }
            if let Some(canonical) = transition.canonical_fields {
                for field in canonical.definitions() {
                    let items = 2usize
                        .checked_add(field.path.len())
                        .and_then(|value| value.checked_add(field.references.len()))
                        .ok_or_else(|| {
                            Error::invalid_archive(index, "canonical metadata cardinality overflow")
                        })?;
                    count = match canonical {
                        CanonicalObjectReferenceFields::Insert(_) => count.checked_add(items),
                        CanonicalObjectReferenceFields::Remove(_) => count.checked_sub(items),
                    }
                    .ok_or_else(|| {
                        Error::invalid_archive(index, "canonical metadata cardinality overflow")
                    })?;
                }
            }
        }
        add_limited(
            &mut total,
            count,
            LimitKind::MetadataItems,
            limits.max_metadata_items(),
        )?;
    }
    Ok(())
}

fn rewrite_message_metadata_in_header(
    source: &[u8],
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    limits: Limits,
) -> Result<Vec<u8>> {
    if current_type == replacement_type && current_length == replacement_length {
        return try_copy_bytes(source, "IWA preserved ArchiveInfo header");
    }

    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let message_field = fields
        .iter()
        .copied()
        .filter(|field| field.number() == 2 && field.wire_type() == 2)
        .nth(message_index)
        .ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "message metadata is missing from ArchiveInfo",
            )
        })?;
    let message_source = message_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let rewritten_message = rewrite_effective_message_scalars(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        wire_limits,
        message_index,
    )?;
    let message_start = message_field.payload_start();
    let message_end = message_field.end();

    let length_prefix = source
        .get(message_field.key_end()..message_start)
        .ok_or_else(|| Error::invalid_archive(message_index, "message prefix range is invalid"))?;
    let mut encoded_length = [0u8; MAX_VARINT_BYTES];
    let encoded_length = encode_varint_with_width(
        u64::try_from(rewritten_message.len()).map_err(|_| {
            Error::invalid_archive(message_index, "message metadata length exceeds u64")
        })?,
        length_prefix.len(),
        &mut encoded_length,
    );
    let output_length = source
        .len()
        .checked_sub(message_source.len())
        .and_then(|length| length.checked_sub(length_prefix.len()))
        .and_then(|length| length.checked_add(encoded_length.len()))
        .and_then(|length| length.checked_add(rewritten_message.len()))
        .ok_or_else(|| Error::invalid_archive(message_index, "ArchiveInfo rewrite overflow"))?;
    check_header_length(output_length, limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA preserved ArchiveInfo header", output_length))?;
    output.extend_from_slice(
        source
            .get(..message_field.key_end())
            .ok_or_else(|| Error::invalid_archive(message_index, "message key range is invalid"))?,
    );
    output.extend_from_slice(encoded_length);
    output.extend_from_slice(&rewritten_message);
    output.extend_from_slice(source.get(message_end..).ok_or_else(|| {
        Error::invalid_archive(message_index, "message metadata range is invalid")
    })?);
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "ArchiveInfo rewrite length mismatch",
        ));
    }
    Ok(output)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the helper keeps both old and new required MessageInfo scalars explicit"
)]
fn rewrite_effective_message_scalars(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    wire_limits: WireLimits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = (current_type != replacement_type)
        .then(|| {
            fields
                .iter()
                .rposition(|field| field.number() == 1 && field.wire_type() == 0)
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "MessageInfo type field is missing")
                })
        })
        .transpose()?;
    let length_field = (current_length != replacement_length)
        .then(|| {
            fields
                .iter()
                .rposition(|field| field.number() == 3 && field.wire_type() == 0)
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "MessageInfo length field is missing")
                })
        })
        .transpose()?;

    let mut output_length = source.len();
    for (field_index, value) in [
        type_field.map(|index| (index, u64::from(replacement_type))),
        length_field.map(|index| (index, u64::from(replacement_length))),
    ]
    .into_iter()
    .flatten()
    {
        let payload_length = fields[field_index]
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
            .len();
        let encoded_length = encoded_varint_width(value, payload_length);
        output_length = output_length
            .checked_sub(payload_length)
            .and_then(|length| length.checked_add(encoded_length))
            .ok_or_else(|| Error::invalid_archive(message_index, "MessageInfo rewrite overflow"))?;
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA preserved MessageInfo header", output_length))?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        let replacement = if type_field == Some(field_index) {
            Some(u64::from(replacement_type))
        } else if length_field == Some(field_index) {
            Some(u64::from(replacement_length))
        } else {
            None
        };
        if let Some(value) = replacement {
            output.extend_from_slice(
                field
                    .key(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?,
            );
            let preferred_width = field
                .payload(source)
                .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
                .len();
            let mut encoded = [0u8; MAX_VARINT_BYTES];
            output.extend_from_slice(encode_varint_with_width(
                value,
                preferred_width,
                &mut encoded,
            ));
        } else {
            output.extend_from_slice(
                field
                    .raw(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?,
            );
        }
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "MessageInfo rewrite length mismatch",
        ));
    }
    Ok(output)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies source cardinality and both old and new required scalars"
)]
fn rewrite_message_metadata_and_prune_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let mut message_count = 0usize;
    let mut target = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        if message_count == message_index {
            target = Some((field_index, field));
        }
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
    }
    if message_count != expected_message_count {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    let (target_index, target_field) = target.ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "message metadata is missing from ArchiveInfo",
        )
    })?;
    let message_source = target_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let Some(rewritten_message) = rewrite_effective_message_scalars_and_references(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        object_removals,
        data_removals,
        wire_limits,
        limits,
        message_index,
    )?
    else {
        return try_copy_bytes(source, "IWA preserved ArchiveInfo header");
    };

    let mut rewrites = retained_field_rewrites(fields.len())?;
    assign_header_field_rewrite(
        &mut rewrites,
        target_index,
        HeaderFieldRewrite::LengthDelimited(rewritten_message),
        message_index,
    )?;
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        message_index,
        "IWA pruned ArchiveInfo header",
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies source cardinality, both scalars, and the authorized reference order"
)]
fn rewrite_message_metadata_and_reorder_object_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    expected_object_references: &[u64],
    reordered_object_references: &[u64],
    limits: Limits,
) -> Result<Vec<u8>> {
    rewrite_message_metadata_and_reorder_references_in_header(
        source,
        expected_message_count,
        message_index,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        5,
        false,
        expected_object_references,
        reordered_object_references,
        limits,
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite keeps complete source and requested data-reference state explicit"
)]
fn rewrite_message_metadata_and_reorder_data_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    expected_data_references: &[u64],
    reordered_data_references: &[u64],
    limits: Limits,
) -> Result<Vec<u8>> {
    rewrite_message_metadata_and_reorder_references_in_header(
        source,
        expected_message_count,
        message_index,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        6,
        true,
        expected_data_references,
        reordered_data_references,
        limits,
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies source cardinality, both scalars, and the authorized reference order"
)]
fn rewrite_message_metadata_and_reorder_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    reference_field_number: u32,
    reorder_matching_field_data_references: bool,
    expected_references: &[u64],
    reordered_references: &[u64],
    limits: Limits,
) -> Result<Vec<u8>> {
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let mut message_count = 0usize;
    let mut target = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        if message_count == message_index {
            target = Some((field_index, field));
        }
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
    }
    if message_count != expected_message_count {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    let (target_index, target_field) = target.ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "message metadata is missing from ArchiveInfo",
        )
    })?;
    let message_source = target_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let rewritten_message = rewrite_effective_message_scalars_and_reference_order(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        reference_field_number,
        reorder_matching_field_data_references,
        expected_references,
        reordered_references,
        wire_limits,
        limits,
        message_index,
    )?;

    let mut rewrites = retained_field_rewrites(fields.len())?;
    assign_header_field_rewrite(
        &mut rewrites,
        target_index,
        HeaderFieldRewrite::LengthDelimited(rewritten_message),
        message_index,
    )?;
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        message_index,
        "IWA reordered ArchiveInfo header",
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies source cardinality and both required message scalars"
)]
fn rewrite_message_metadata_and_transition_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    transition: &PreparedObjectReferenceTransition<'_>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let mut message_count = 0usize;
    let mut target = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        if message_count == message_index {
            target = Some((field_index, field));
        }
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
    }
    if message_count != expected_message_count {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    let (target_index, target_field) = target.ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "message metadata is missing from ArchiveInfo",
        )
    })?;
    let message_source = target_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let rewritten_message = rewrite_effective_message_scalars_and_reference_transition(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        transition,
        wire_limits,
        limits,
        message_index,
    )?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    assign_header_field_rewrite(
        &mut rewrites,
        target_index,
        HeaderFieldRewrite::LengthDelimited(rewritten_message),
        message_index,
    )?;
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        message_index,
        "IWA transitioned ArchiveInfo header",
    )
}

fn validate_data_reference_transition_lists(
    transition: &DataReferenceTransition<'_>,
    limits: Limits,
    message_index: usize,
) -> Result<()> {
    let aggregate_items = transition
        .aggregate_before
        .len()
        .checked_add(transition.aggregate_after.len())
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "data-reference transition overflow")
        })?;
    if aggregate_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            aggregate_items,
            limits.max_metadata_items(),
        ));
    }
    if transition
        .aggregate_before
        .iter()
        .chain(transition.aggregate_after)
        .any(|identifier| *identifier == 0)
    {
        return Err(Error::invalid_archive(
            message_index,
            "data-reference transition contains a zero identifier",
        ));
    }
    let mut authorization_items = transition.fields.len();
    for field in transition.fields {
        authorization_items = authorization_items
            .checked_add(field.expected_path.len())
            .and_then(|count| count.checked_add(field.before.len()))
            .and_then(|count| count.checked_add(field.after.len()))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "FieldInfo transition authorization overflow")
            })?;
        if field
            .before
            .iter()
            .chain(field.after)
            .any(|identifier| *identifier == 0)
        {
            return Err(Error::invalid_archive(
                message_index,
                "FieldInfo data-reference transition contains a zero identifier",
            ));
        }
        if field
            .before
            .iter()
            .any(|identifier| !transition.aggregate_before.contains(identifier))
            || field
                .after
                .iter()
                .any(|identifier| !transition.aggregate_after.contains(identifier))
        {
            return Err(Error::invalid_archive(
                message_index,
                "FieldInfo transition references are outside the aggregate state",
            ));
        }
    }
    if authorization_items > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            authorization_items,
            limits.max_metadata_items(),
        ));
    }
    Ok(())
}

fn prepare_canonical_field_data_reference_transition<'a>(
    current: &MessageInfo,
    transition: DataReferenceTransition<'a>,
    canonical: CanonicalFieldDataReferenceTransition<'a>,
    limits: Limits,
    message_index: usize,
) -> Result<CanonicalFieldDataReferencePlan<'a>> {
    if !transition.fields.is_empty() {
        return Err(Error::invalid_archive(
            message_index,
            "canonical FieldInfo transition cannot select existing fields",
        ));
    }
    if canonical.field_path.len() > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            canonical.field_path.len(),
            limits.max_metadata_items(),
        ));
    }
    match canonical.operation {
        CanonicalFieldDataReferenceOperation::Insert => {
            if current
                .field_infos
                .iter()
                .any(|field| field.path.as_slice() == canonical.field_path)
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "canonical FieldInfo path already exists",
                ));
            }
            if canonical.aggregate_after.is_empty() {
                return Err(Error::invalid_archive(
                    message_index,
                    "canonical FieldInfo insertion requires target data references",
                ));
            }
            Ok(CanonicalFieldDataReferencePlan::Insert {
                path: canonical.field_path,
            })
        },
        CanonicalFieldDataReferenceOperation::Remove => {
            let mut matching = None;
            for (field_info_index, field) in current.field_infos.iter().enumerate() {
                if field.path.as_slice() != canonical.field_path {
                    continue;
                }
                if matching.replace(field_info_index).is_some() {
                    return Err(Error::invalid_archive(
                        message_index,
                        "canonical FieldInfo path is ambiguous",
                    ));
                }
            }
            let field_info_index = matching.ok_or_else(|| {
                Error::invalid_archive(message_index, "canonical FieldInfo is missing")
            })?;
            if field_info_index != current.field_infos.len().saturating_sub(1) {
                return Err(Error::invalid_archive(
                    message_index,
                    "canonical FieldInfo is not the final field",
                ));
            }
            let field = current
                .field_infos
                .get(field_info_index)
                .ok_or_else(|| Error::invalid_archive(message_index, "FieldInfo is missing"))?;
            if field.r#type != Some(FieldType::DataReference)
                || !field.object_references.is_empty()
                || field.unknown_field_rule.is_some()
                || field.known_field_rule.is_some()
                || !field.known_field_version.is_empty()
                || field.known_field_feature_identifier.is_some()
                || field.data_references.as_slice() != canonical.aggregate_before
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "FieldInfo is not the canonical data-reference record",
                ));
            }
            Ok(CanonicalFieldDataReferencePlan::Remove {
                field_info_index,
                path: canonical.field_path,
            })
        },
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "The raw header rewrite keeps both complete data-reference states and selected FieldInfo occurrences explicit."
)]
fn rewrite_message_metadata_and_transition_data_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    transition: DataReferenceTransition<'_>,
    selected_fields: &[Option<&FieldDataReferenceTransition<'_>>],
    canonical_plan: Option<CanonicalFieldDataReferencePlan<'_>>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let mut message_count = 0usize;
    let mut target = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        if message_count == message_index {
            target = Some((field_index, field));
        }
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
    }
    if message_count != expected_message_count {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    let (target_index, target_field) = target.ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "message metadata is missing from ArchiveInfo",
        )
    })?;
    let message_source = target_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let rewritten_message = rewrite_effective_message_scalars_and_data_reference_transition(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        transition,
        selected_fields,
        canonical_plan,
        wire_limits,
        limits,
        message_index,
    )?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    assign_header_field_rewrite(
        &mut rewrites,
        target_index,
        HeaderFieldRewrite::LengthDelimited(rewritten_message),
        message_index,
    )?;
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        message_index,
        "IWA transitioned data-reference ArchiveInfo header",
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "The raw MessageInfo rewrite keeps scalar and complete data-reference authorization explicit."
)]
fn rewrite_effective_message_scalars_and_data_reference_transition(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    transition: DataReferenceTransition<'_>,
    selected_fields: &[Option<&FieldDataReferenceTransition<'_>>],
    canonical_plan: Option<CanonicalFieldDataReferencePlan<'_>>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(current_type),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_type != replacement_type {
        assign_header_field_rewrite(
            &mut rewrites,
            type_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_type)),
            message_index,
        )?;
    }
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    let mut field_info_index = 0usize;
    let mut aggregate_seen = false;
    for (field_index, field) in fields.iter().copied().enumerate() {
        match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let selected = selected_fields
                    .get(field_info_index)
                    .and_then(Option::as_ref);
                let remove = canonical_plan.is_some_and(|plan| {
                    matches!(
                        plan,
                        CanonicalFieldDataReferencePlan::Remove {
                            field_info_index: removed_index,
                            ..
                        } if removed_index == field_info_index
                    )
                });
                if remove {
                    let raw = field
                        .raw(source)
                        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                    let (path, references) = match canonical_plan {
                        Some(CanonicalFieldDataReferencePlan::Remove { path, .. }) => {
                            (path, transition.aggregate_before)
                        },
                        _ => unreachable!("canonical FieldInfo removal plan disappeared"),
                    };
                    let expected = encode_canonical_data_reference_field_info(
                        path,
                        references,
                        limits,
                        message_index,
                    )?;
                    if raw != expected.as_slice() {
                        return Err(Error::invalid_archive(
                            message_index,
                            "canonical FieldInfo source bytes differ from insertion output",
                        ));
                    }
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::Remove,
                        message_index,
                    )?;
                } else if let Some(selected) = selected {
                    let payload = field
                        .payload(source)
                        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                    let rewritten = rewrite_field_info_data_reference_transition(
                        payload,
                        selected.before,
                        selected.after,
                        wire_limits,
                        limits,
                        message_index,
                    )?;
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::LengthDelimited(rewritten),
                        message_index,
                    )?;
                }
                field_info_index = field_info_index.checked_add(1).ok_or_else(|| {
                    Error::invalid_archive(message_index, "FieldInfo count overflow")
                })?;
            },
            6 => {
                if !aggregate_seen {
                    let payload = field
                        .payload(source)
                        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                    let rewritten = rewrite_repeated_data_reference_fields(
                        source,
                        field,
                        payload,
                        transition.aggregate_before,
                        transition.aggregate_after,
                        wire_limits,
                        limits,
                        message_index,
                        "IWA transitioned MessageInfo data references",
                    )?;
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        rewritten,
                        message_index,
                    )?;
                    aggregate_seen = true;
                } else {
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::Remove,
                        message_index,
                    )?;
                }
            },
            _ => {},
        }
    }
    if field_info_index != selected_fields.len() {
        return Err(Error::invalid_archive(
            message_index,
            "raw and authorized FieldInfo counts differ",
        ));
    }
    if !aggregate_seen && !transition.aggregate_after.is_empty() {
        let mut rewritten = assemble_header_field_rewrites(
            source,
            &fields,
            &rewrites,
            HeaderKind::MessageInfo,
            limits,
            message_index,
            "IWA transitioned MessageInfo header",
        )?;
        append_packed_data_references_to_message_info(
            &mut rewritten,
            transition.aggregate_after,
            limits,
            message_index,
        )?;
        return append_canonical_field_info_if_needed(
            rewritten,
            canonical_plan,
            transition.aggregate_after,
            limits,
            message_index,
        );
    }
    let rewritten = assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA transitioned MessageInfo header",
    )?;
    append_canonical_field_info_if_needed(
        rewritten,
        canonical_plan,
        transition.aggregate_after,
        limits,
        message_index,
    )
}

fn append_canonical_field_info_if_needed(
    mut message: Vec<u8>,
    canonical_plan: Option<CanonicalFieldDataReferencePlan<'_>>,
    references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let Some(CanonicalFieldDataReferencePlan::Insert { path }) = canonical_plan else {
        return Ok(message);
    };
    let encoded =
        encode_canonical_data_reference_field_info(path, references, limits, message_index)?;
    let output_length = message
        .len()
        .checked_add(encoded.len())
        .ok_or_else(|| Error::invalid_archive(message_index, "FieldInfo append overflow"))?;
    check_header_length(output_length, limits)?;
    message
        .try_reserve_exact(encoded.len())
        .map_err(|_| Error::allocation("IWA canonical FieldInfo", encoded.len()))?;
    message.extend_from_slice(&encoded);
    Ok(message)
}

/// Encode the complete outer `MessageInfo.field_infos` record for the one
/// canonical field introduced by the aggregate-only compatibility path.
///
/// This writer is intentionally small and schema-explicit.  It emits the
/// required path presence, `FieldType::DataReference`, and one packed data
/// reference list, with no optional or unknown fields.  The caller can then
/// append the record to an otherwise source-authoritative MessageInfo without
/// re-encoding unrelated metadata.
fn encode_canonical_data_reference_field_info(
    path: &[u32],
    references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    encode_canonical_reference_field_info(
        path,
        references,
        FieldType::DataReference,
        limits,
        message_index,
    )
}

fn encode_canonical_object_reference_field_info(
    field: CanonicalObjectReferenceField<'_>,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    encode_canonical_reference_field_info(
        field.path,
        field.references,
        FieldType::ObjectReference,
        limits,
        message_index,
    )
}

fn encode_canonical_reference_field_info(
    path: &[u32],
    references: &[u64],
    kind: FieldType,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let (type_value, reference_field_number) = match kind {
        FieldType::ObjectReference => (1u8, 4u32),
        FieldType::DataReference => (2u8, 5u32),
        _ => {
            return Err(Error::invalid_archive(
                message_index,
                "canonical field requires a reference type",
            ));
        },
    };
    let path_values_length = packed_u32_length(path)?;
    let path_message_length = packed_u32_field_length(1, path)?;
    let path_field_length = length_delimited_field_length(1, path_message_length)?;
    let type_field_length = varint_field_length(2, u64::from(type_value))?;
    let references_field_length = packed_u64_field_length(reference_field_number, references)?;
    let payload_length = path_field_length
        .checked_add(type_field_length)
        .and_then(|length| length.checked_add(references_field_length))
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "canonical FieldInfo length overflow")
        })?;
    let output_length = length_delimited_field_length(4, payload_length)?;
    check_header_length(output_length, limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA canonical FieldInfo", output_length))?;
    append_wire_key_and_length(&mut output, 4, payload_length, message_index)?;
    append_wire_key_and_length(&mut output, 1, path_message_length, message_index)?;
    let mut bytes = [0u8; MAX_VARINT_BYTES];
    if !path.is_empty() {
        append_wire_key_and_length(&mut output, 1, path_values_length, message_index)?;
        for value in path {
            output.extend_from_slice(encode_varint(u64::from(*value), &mut bytes));
        }
    }
    output.extend_from_slice(&[0x10, type_value]);
    if !references.is_empty() {
        append_wire_key_and_length(
            &mut output,
            reference_field_number,
            packed_u64_length(references)?,
            message_index,
        )?;
        for value in references {
            output.extend_from_slice(encode_varint(*value, &mut bytes));
        }
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "canonical FieldInfo length mismatch",
        ));
    }
    Ok(output)
}

fn append_wire_key_and_length(
    output: &mut Vec<u8>,
    field_number: u32,
    payload_length: usize,
    message_index: usize,
) -> Result<()> {
    let key = u64::from(field_number)
        .checked_shl(3)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| Error::invalid_archive(message_index, "wire field key overflow"))?;
    let mut bytes = [0u8; MAX_VARINT_BYTES];
    output.extend_from_slice(encode_varint(key, &mut bytes));
    output.extend_from_slice(encode_varint(
        u64::try_from(payload_length)
            .map_err(|_| Error::invalid_archive(message_index, "wire field length exceeds u64"))?,
        &mut bytes,
    ));
    Ok(())
}

fn rewrite_field_info_data_reference_transition(
    source: &[u8],
    before: &[u64],
    after: &[u64],
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    let mut first = None;
    let mut observed = Vec::new();
    observed
        .try_reserve_exact(before.len())
        .map_err(|_| Error::allocation("IWA FieldInfo data-reference validation", before.len()))?;
    for (index, field) in fields.iter().copied().enumerate() {
        if field.number() == 5 {
            if first.is_none() {
                first = Some((index, field));
            }
            let payload = field
                .payload(source)
                .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
            append_decoded_reference_values(
                payload,
                field.wire_type(),
                &mut observed,
                message_index,
                "FieldInfo data-reference field has an ambiguous wire type",
            )?;
            assign_header_field_rewrite(
                &mut rewrites,
                index,
                HeaderFieldRewrite::Remove,
                message_index,
            )?;
        }
    }
    if observed != before {
        return Err(Error::invalid_archive(
            message_index,
            "raw FieldInfo data references differ from transition authorization",
        ));
    }
    if let Some((index, _field)) = first {
        if !after.is_empty() {
            let encoded = encode_packed_reference_field(
                5,
                after,
                limits,
                message_index,
                "IWA transitioned FieldInfo data references",
            )?;
            assign_header_field_rewrite(
                &mut rewrites,
                index,
                HeaderFieldRewrite::Raw(encoded),
                message_index,
            )?;
        }
        assemble_header_field_rewrites(
            source,
            &fields,
            &rewrites,
            HeaderKind::MessageInfo,
            limits,
            message_index,
            "IWA transitioned FieldInfo header",
        )
    } else if after.is_empty() {
        assemble_header_field_rewrites(
            source,
            &fields,
            &rewrites,
            HeaderKind::MessageInfo,
            limits,
            message_index,
            "IWA transitioned FieldInfo header",
        )
    } else {
        let mut rewritten = assemble_header_field_rewrites(
            source,
            &fields,
            &rewrites,
            HeaderKind::MessageInfo,
            limits,
            message_index,
            "IWA transitioned FieldInfo header",
        )?;
        let encoded = encode_packed_reference_field(
            5,
            after,
            limits,
            message_index,
            "IWA transitioned FieldInfo data references",
        )?;
        rewritten.try_reserve_exact(encoded.len()).map_err(|_| {
            Error::allocation("IWA transitioned FieldInfo data references", encoded.len())
        })?;
        rewritten.extend_from_slice(&encoded);
        check_header_length(rewritten.len(), limits)?;
        Ok(rewritten)
    }
}

fn rewrite_repeated_data_reference_fields(
    source: &[u8],
    first_field: WireField,
    first_payload: &[u8],
    before: &[u64],
    after: &[u64],
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
    allocation_context: &'static str,
) -> Result<HeaderFieldRewrite> {
    let _ = first_payload;
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut observed = Vec::new();
    observed
        .try_reserve_exact(before.len())
        .map_err(|_| Error::allocation(allocation_context, before.len()))?;
    let mut first_start = None;
    for field in fields.iter().copied().filter(|field| field.number() == 6) {
        if first_start.is_none() {
            first_start = Some(field.start());
        }
        let payload = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
        append_decoded_reference_values(
            payload,
            field.wire_type(),
            &mut observed,
            message_index,
            "MessageInfo data-reference field has an ambiguous wire type",
        )?;
    }
    if first_start != Some(first_field.start()) || observed != before {
        return Err(Error::invalid_archive(
            message_index,
            "raw MessageInfo data references differ from transition authorization",
        ));
    }
    if after.is_empty() {
        Ok(HeaderFieldRewrite::Remove)
    } else {
        Ok(HeaderFieldRewrite::Raw(encode_packed_reference_field(
            6,
            after,
            limits,
            message_index,
            allocation_context,
        )?))
    }
}

fn append_decoded_reference_values(
    payload: &[u8],
    wire_type: u8,
    output: &mut Vec<u64>,
    message_index: usize,
    ambiguous_reason: &'static str,
) -> Result<()> {
    match wire_type {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| {
                    Error::invalid_archive(message_index, "malformed unpacked data reference")
                })?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "unpacked data reference has trailing bytes",
                ));
            }
            output.push(value);
        },
        2 => {
            let mut cursor = 0usize;
            while cursor < payload.len() {
                let remaining = payload.get(cursor..).ok_or_else(|| {
                    Error::invalid_archive(message_index, "packed data-reference range is invalid")
                })?;
                let (value, encoded_length) =
                    litchi_iwa_common::decode_varint_from_bytes(remaining).map_err(|_| {
                        Error::invalid_archive(message_index, "malformed packed data reference")
                    })?;
                output.push(value);
                cursor = cursor.checked_add(encoded_length).ok_or_else(|| {
                    Error::invalid_archive(message_index, "packed data-reference range overflow")
                })?;
            }
        },
        _ => return Err(Error::invalid_archive(message_index, ambiguous_reason)),
    }
    Ok(())
}

fn encode_reference_payload(
    references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let mut length = 0usize;
    for identifier in references {
        length = length
            .checked_add(u64_varint_length(*identifier))
            .ok_or_else(|| Error::invalid_archive(message_index, "reference payload overflow"))?;
    }
    if length > limits.max_header_bytes() {
        return Err(limit(
            LimitKind::HeaderBytes,
            length,
            limits.max_header_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| Error::allocation("IWA transitioned data-reference payload", length))?;
    let mut bytes = [0u8; MAX_VARINT_BYTES];
    for identifier in references {
        output.extend_from_slice(encode_varint(*identifier, &mut bytes));
    }
    Ok(output)
}

fn append_packed_data_references_to_message_info(
    message: &mut Vec<u8>,
    references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<()> {
    append_packed_data_reference_field(message, 6, references, limits, message_index)
}

fn append_packed_data_reference_field(
    message: &mut Vec<u8>,
    field_number: u32,
    references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<()> {
    if references.is_empty() {
        return Ok(());
    }
    let encoded = encode_packed_reference_field(
        field_number,
        references,
        limits,
        message_index,
        "IWA appended data references",
    )?;
    let extra = encoded.len();
    let output_len = message
        .len()
        .checked_add(extra)
        .ok_or_else(|| Error::invalid_archive(message_index, "reference append overflow"))?;
    check_header_length(output_len, limits)?;
    message
        .try_reserve_exact(extra)
        .map_err(|_| Error::allocation("IWA appended data references", extra))?;
    message.extend_from_slice(&encoded);
    Ok(())
}

fn encode_packed_reference_field(
    field_number: u32,
    references: &[u64],
    limits: Limits,
    message_index: usize,
    allocation_context: &'static str,
) -> Result<Vec<u8>> {
    let payload = encode_reference_payload(references, limits, message_index)?;
    let key_value = u64::from(field_number)
        .checked_shl(3)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| Error::invalid_archive(message_index, "reference field key overflow"))?;
    let mut key_bytes = [0u8; MAX_VARINT_BYTES];
    let mut length_bytes = [0u8; MAX_VARINT_BYTES];
    let key = encode_varint(key_value, &mut key_bytes);
    let length = encode_varint(
        u64::try_from(payload.len()).map_err(|_| {
            Error::invalid_archive(message_index, "reference payload length exceeds u64")
        })?,
        &mut length_bytes,
    );
    let output_length = key
        .len()
        .checked_add(length.len())
        .and_then(|value| value.checked_add(payload.len()))
        .ok_or_else(|| Error::invalid_archive(message_index, "reference field overflow"))?;
    check_header_length(output_length, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation(allocation_context, output_length))?;
    output.extend_from_slice(key);
    output.extend_from_slice(length);
    output.extend_from_slice(&payload);
    Ok(output)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies both required scalars and exact nested reference state"
)]
fn rewrite_effective_message_scalars_and_reference_transition(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    transition: &PreparedObjectReferenceTransition<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(current_type),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_type != replacement_type {
        assign_header_field_rewrite(
            &mut rewrites,
            type_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_type)),
            message_index,
        )?;
    }
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    let mut aggregate_source_index = 0usize;
    let mut aggregate_retained_index = 0usize;
    let mut field_info_index = 0usize;
    for (field_index, field) in fields.iter().copied().enumerate() {
        match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let selected = transition.fields.get(field_info_index).ok_or_else(|| {
                    Error::invalid_archive(
                        message_index,
                        "raw FieldInfo list exceeds transition authorization",
                    )
                })?;
                if field_info_index >= transition.retained_field_count {
                    let Some(CanonicalObjectReferenceFields::Remove(definitions)) =
                        transition.canonical_fields
                    else {
                        return Err(Error::invalid_archive(
                            message_index,
                            "canonical FieldInfo removal authorization is missing",
                        ));
                    };
                    let definition = definitions
                        .get(field_info_index - transition.retained_field_count)
                        .ok_or_else(|| {
                            Error::invalid_archive(
                                message_index,
                                "canonical FieldInfo removal index is out of bounds",
                            )
                        })?;
                    let expected = encode_canonical_object_reference_field_info(
                        *definition,
                        limits,
                        message_index,
                    )?;
                    let raw = field
                        .raw(source)
                        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                    if raw != expected.as_slice() {
                        return Err(Error::invalid_archive(
                            message_index,
                            "canonical object FieldInfo source bytes differ from insertion output",
                        ));
                    }
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::Remove,
                        message_index,
                    )?;
                } else if let Some(selected) = selected {
                    let payload = field
                        .payload(source)
                        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                    let rewritten = rewrite_field_info_reference_transition(
                        payload,
                        &selected.references,
                        wire_limits,
                        limits,
                        message_index,
                    )?;
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::LengthDelimited(rewritten),
                        message_index,
                    )?;
                }
                field_info_index = field_info_index.checked_add(1).ok_or_else(|| {
                    Error::invalid_archive(message_index, "FieldInfo count overflow")
                })?;
            },
            5 => {
                let rewrite = rewrite_reference_transition_field(
                    source,
                    field,
                    &transition.aggregate,
                    &mut aggregate_source_index,
                    &mut aggregate_retained_index,
                    message_index,
                )?;
                assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
            },
            _ => {},
        }
    }
    verify_transition_cursor(
        &transition.aggregate,
        aggregate_source_index,
        aggregate_retained_index,
        message_index,
    )?;
    if field_info_index != transition.fields.len() {
        return Err(Error::invalid_archive(
            message_index,
            "raw and authorized FieldInfo counts differ",
        ));
    }
    let rewritten = assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA transitioned MessageInfo header",
    )?;
    let rewritten = append_packed_object_references_to_message_info(
        rewritten,
        transition.aggregate.appended_after,
        limits,
        message_index,
    )?;
    append_canonical_object_reference_fields(
        rewritten,
        transition.canonical_fields,
        limits,
        message_index,
    )
}

fn append_canonical_object_reference_fields(
    mut message: Vec<u8>,
    canonical: Option<CanonicalObjectReferenceFields<'_>>,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let Some(CanonicalObjectReferenceFields::Insert(definitions)) = canonical else {
        return Ok(message);
    };
    let mut additional = 0usize;
    for field in definitions {
        let path_message = packed_u32_field_length(1, field.path)?;
        let path_field = length_delimited_field_length(1, path_message)?;
        let references = packed_u64_field_length(4, field.references)?;
        let payload = path_field
            .checked_add(2)
            .and_then(|value| value.checked_add(references))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "canonical FieldInfo payload length overflow")
            })?;
        let record = length_delimited_field_length(4, payload)?;
        additional = additional.checked_add(record).ok_or_else(|| {
            Error::invalid_archive(message_index, "canonical FieldInfo batch length overflow")
        })?;
    }
    let length = message.len().checked_add(additional).ok_or_else(|| {
        Error::invalid_archive(message_index, "canonical MessageInfo length overflow")
    })?;
    check_header_length(length, limits)?;
    message
        .try_reserve_exact(additional)
        .map_err(|_| Error::allocation("IWA canonical object fields", additional))?;
    for field in definitions {
        let encoded = encode_canonical_object_reference_field_info(*field, limits, message_index)?;
        message.extend_from_slice(&encoded);
    }
    if message.len() != length {
        return Err(Error::invalid_archive(
            message_index,
            "canonical object FieldInfo batch length mismatch",
        ));
    }
    Ok(message)
}

fn rewrite_field_info_reference_transition(
    source: &[u8],
    transition: &PreparedReferenceList<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    let mut source_index = 0usize;
    let mut retained_index = 0usize;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 4 {
            continue;
        }
        let rewrite = rewrite_reference_transition_field(
            source,
            field,
            transition,
            &mut source_index,
            &mut retained_index,
            message_index,
        )?;
        assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
    }
    verify_transition_cursor(transition, source_index, retained_index, message_index)?;
    let rewritten = assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA transitioned FieldInfo header",
    )?;
    append_packed_object_references_to_field_info(
        &rewritten,
        transition.appended_after,
        limits,
        message_index,
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies both scalars and the complete authorized reference order"
)]
fn rewrite_effective_message_scalars_and_reference_order(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    reference_field_number: u32,
    reorder_matching_field_data_references: bool,
    expected_references: &[u64],
    reordered_references: &[u64],
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    if expected_references.len() != reordered_references.len() {
        return Err(Error::invalid_archive(
            message_index,
            "reordered reference count changed",
        ));
    }
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(current_type),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;

    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_type != replacement_type {
        assign_header_field_rewrite(
            &mut rewrites,
            type_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_type)),
            message_index,
        )?;
    }
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    let mut reference_index = 0usize;
    let mut matching_field_info_count = 0usize;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if reorder_matching_field_data_references && field.number() == 4 {
            let payload = field
                .payload(source)
                .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
            let rewrite = rewrite_field_info_data_reference_order(
                payload,
                expected_references,
                reordered_references,
                wire_limits,
                limits,
                message_index,
            )?;
            if matches!(rewrite, HeaderFieldRewrite::LengthDelimited(_)) {
                matching_field_info_count =
                    matching_field_info_count.checked_add(1).ok_or_else(|| {
                        Error::invalid_archive(message_index, "matching FieldInfo count overflow")
                    })?;
                if matching_field_info_count > 1 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "data-reference reorder has ambiguous matching FieldInfo metadata",
                    ));
                }
            }
            assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
        } else if field.number() == reference_field_number {
            let rewrite = rewrite_object_reference_order_field(
                source,
                field,
                expected_references,
                reordered_references,
                &mut reference_index,
                message_index,
            )?;
            assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
        }
    }
    if reference_index != expected_references.len() {
        return Err(Error::invalid_archive(
            message_index,
            "raw and authorized reference counts differ",
        ));
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA reordered MessageInfo header",
    )
}

fn rewrite_field_info_data_reference_order(
    source: &[u8],
    expected: &[u64],
    reordered: &[u64],
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut actual_count = 0usize;
    let mut matches_expected = true;
    // Count packed and unpacked values before retaining any decoded value, so
    // a hostile FieldInfo cannot grow an intermediate collection past the
    // caller's metadata budget.
    let mut observe = |value| {
        if expected.get(actual_count).copied() != Some(value) {
            matches_expected = false;
        }
        actual_count = actual_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "FieldInfo data-reference count overflow")
        })?;
        if actual_count > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                actual_count,
                limits.max_metadata_items(),
            ));
        }
        Ok::<(), Error>(())
    };
    for field in fields.iter().copied().filter(|field| field.number() == 5) {
        let payload = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
        match field.wire_type() {
            0 => {
                let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                    .map_err(|_| {
                        Error::invalid_archive(message_index, "malformed FieldInfo data reference")
                    })?;
                if encoded_length != payload.len() {
                    return Err(Error::invalid_archive(
                        message_index,
                        "FieldInfo data reference has trailing bytes",
                    ));
                }
                observe(value)?;
            },
            2 => {
                let mut remaining = payload;
                while !remaining.is_empty() {
                    let (value, encoded_length) =
                        litchi_iwa_common::decode_varint_from_bytes(remaining).map_err(|_| {
                            Error::invalid_archive(
                                message_index,
                                "malformed packed FieldInfo data reference",
                            )
                        })?;
                    observe(value)?;
                    remaining = remaining.get(encoded_length..).ok_or_else(|| {
                        Error::invalid_archive(
                            message_index,
                            "packed FieldInfo data-reference range is invalid",
                        )
                    })?;
                }
            },
            _ => {
                return Err(Error::invalid_archive(
                    message_index,
                    "FieldInfo data-reference field has an ambiguous wire type",
                ));
            },
        }
    }
    if !matches_expected || actual_count != expected.len() {
        return Ok(HeaderFieldRewrite::Retain);
    }
    let mut reference_index = 0usize;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 5 {
            continue;
        }
        let rewrite = rewrite_object_reference_order_field(
            source,
            field,
            expected,
            reordered,
            &mut reference_index,
            message_index,
        )?;
        assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
    }
    if reference_index != expected.len() {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral FieldInfo data-reference counts differ",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(
        assemble_header_field_rewrites(
            source,
            &fields,
            &rewrites,
            HeaderKind::MessageInfo,
            limits,
            message_index,
            "IWA reordered FieldInfo data references",
        )?,
    ))
}

fn append_packed_object_references_to_field_info(
    source: &[u8],
    appended_object_references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let message = try_copy_bytes(source, "IWA preserved FieldInfo header")?;
    append_packed_reference_field(
        message,
        4,
        appended_object_references,
        limits,
        message_index,
        "IWA appended packed FieldInfo object references",
    )
}

fn append_packed_object_references_to_message_info(
    message: Vec<u8>,
    appended_object_references: &[u64],
    limits: Limits,
    message_index: usize,
) -> Result<Vec<u8>> {
    append_packed_reference_field(
        message,
        5,
        appended_object_references,
        limits,
        message_index,
        "IWA appended packed MessageInfo object references",
    )
}

fn append_packed_reference_field(
    mut message: Vec<u8>,
    field_number: u32,
    appended_object_references: &[u64],
    limits: Limits,
    message_index: usize,
    allocation_context: &'static str,
) -> Result<Vec<u8>> {
    if appended_object_references.is_empty() {
        return Ok(message);
    }
    let mut payload_length = 0usize;
    for identifier in appended_object_references {
        payload_length = payload_length
            .checked_add(encoded_varint_width(*identifier, 0))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "appended reference length overflow")
            })?;
    }
    let prefix_length = varint_len(payload_length)?;
    let key = u64::from(field_number)
        .checked_mul(8)
        .and_then(|key| key.checked_add(2))
        .ok_or_else(|| Error::invalid_archive(message_index, "reference field key overflow"))?;
    let key_length = encoded_varint_width(key, 0);
    let field_length = key_length
        .checked_add(prefix_length)
        .and_then(|length| length.checked_add(payload_length))
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "appended reference field overflow")
        })?;
    let output_length = message
        .len()
        .checked_add(field_length)
        .ok_or_else(|| Error::invalid_archive(message_index, "MessageInfo append overflow"))?;
    check_header_length(output_length, limits)?;
    message
        .try_reserve_exact(field_length)
        .map_err(|_| Error::allocation(allocation_context, field_length))?;
    let mut encoded = [0u8; MAX_VARINT_BYTES];
    message.extend_from_slice(encode_varint(key, &mut encoded));
    message.extend_from_slice(encode_varint(
        u64::try_from(payload_length).map_err(|_| {
            Error::invalid_archive(message_index, "appended reference length exceeds u64")
        })?,
        &mut encoded,
    ));
    for identifier in appended_object_references {
        message.extend_from_slice(encode_varint(*identifier, &mut encoded));
    }
    if message.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "appended reference rewrite length mismatch",
        ));
    }
    Ok(message)
}

fn verify_transition_cursor(
    transition: &PreparedReferenceList<'_>,
    source_index: usize,
    retained_index: usize,
    message_index: usize,
) -> Result<()> {
    if source_index != transition.before.len() || retained_index != transition.retained_after.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw reference list differs from transition authorization",
        ));
    }
    Ok(())
}

fn transition_reference_at(
    transition: &PreparedReferenceList<'_>,
    source_index: &mut usize,
    retained_index: &mut usize,
    raw_value: u64,
    message_index: usize,
) -> Result<Option<u64>> {
    let expected = transition
        .before
        .get(*source_index)
        .copied()
        .ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "raw reference list exceeds transition authorization",
            )
        })?;
    if expected != raw_value {
        return Err(Error::invalid_archive(
            message_index,
            "raw reference order differs from transition authorization",
        ));
    }
    *source_index = source_index
        .checked_add(1)
        .ok_or_else(|| Error::invalid_archive(message_index, "reference index overflow"))?;
    if !transition.after_members.contains(&raw_value) {
        return Ok(None);
    }
    let replacement = transition
        .retained_after
        .get(*retained_index)
        .copied()
        .ok_or_else(|| {
            Error::invalid_archive(message_index, "retained transition list is truncated")
        })?;
    *retained_index = retained_index.checked_add(1).ok_or_else(|| {
        Error::invalid_archive(message_index, "retained reference index overflow")
    })?;
    Ok(Some(replacement))
}

fn rewrite_reference_transition_field(
    source: &[u8],
    field: WireField,
    transition: &PreparedReferenceList<'_>,
    source_index: &mut usize,
    retained_index: &mut usize,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| {
                    Error::invalid_archive(message_index, "malformed unpacked object reference")
                })?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "unpacked object reference has trailing bytes",
                ));
            }
            Ok(
                match transition_reference_at(
                    transition,
                    source_index,
                    retained_index,
                    value,
                    message_index,
                )? {
                    None => HeaderFieldRewrite::Remove,
                    Some(replacement) if replacement == value => HeaderFieldRewrite::Retain,
                    Some(replacement) => HeaderFieldRewrite::Varint(replacement),
                },
            )
        },
        2 => rewrite_packed_reference_transition(
            payload,
            transition,
            source_index,
            retained_index,
            message_index,
        ),
        _ => Err(Error::invalid_archive(
            message_index,
            "object-reference field has an ambiguous wire type",
        )),
    }
}

fn rewrite_packed_reference_transition(
    payload: &[u8],
    transition: &PreparedReferenceList<'_>,
    source_index: &mut usize,
    retained_index: &mut usize,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let start_source_index = *source_index;
    let start_retained_index = *retained_index;
    let mut cursor = 0usize;
    let mut output_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(
            payload.get(cursor..).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range is invalid")
            })?,
        )
        .map_err(|_| Error::invalid_archive(message_index, "malformed packed object reference"))?;
        match transition_reference_at(
            transition,
            source_index,
            retained_index,
            value,
            message_index,
        )? {
            None => changed = true,
            Some(replacement) => {
                output_length = output_length
                    .checked_add(encoded_varint_width(replacement, encoded_length))
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "packed transition length overflow")
                    })?;
                changed |= replacement != value;
            },
        }
        cursor = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
    }
    if !changed {
        return Ok(HeaderFieldRewrite::Retain);
    }

    let mut output = Vec::new();
    output.try_reserve_exact(output_length).map_err(|_| {
        Error::allocation("IWA transitioned packed object references", output_length)
    })?;
    cursor = 0;
    let mut replay_source_index = start_source_index;
    let mut replay_retained_index = start_retained_index;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| {
                Error::invalid_archive(message_index, "malformed packed object reference")
            })?;
        if let Some(replacement) = transition_reference_at(
            transition,
            &mut replay_source_index,
            &mut replay_retained_index,
            value,
            message_index,
        )? {
            let end = cursor.checked_add(encoded_length).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range overflow")
            })?;
            if replacement == value {
                output.extend_from_slice(payload.get(cursor..end).ok_or_else(|| {
                    Error::invalid_archive(message_index, "packed reference range is invalid")
                })?);
            } else {
                let mut encoded = [0u8; MAX_VARINT_BYTES];
                output.extend_from_slice(encode_varint_with_width(
                    replacement,
                    encoded_length,
                    &mut encoded,
                ));
            }
        }
        cursor = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
    }
    if output.len() != output_length
        || replay_source_index != *source_index
        || replay_retained_index != *retained_index
    {
        return Err(Error::invalid_archive(
            message_index,
            "packed reference transition replay diverged",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(output))
}

fn rewrite_object_reference_order_field(
    source: &[u8],
    field: WireField,
    expected: &[u64],
    reordered: &[u64],
    reference_index: &mut usize,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| {
                    Error::invalid_archive(message_index, "malformed unpacked object reference")
                })?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "unpacked object reference has trailing bytes",
                ));
            }
            let replacement = authorized_reordered_reference(
                expected,
                reordered,
                *reference_index,
                value,
                message_index,
            )?;
            *reference_index = reference_index.checked_add(1).ok_or_else(|| {
                Error::invalid_archive(message_index, "object-reference count overflow")
            })?;
            Ok(if value == replacement {
                HeaderFieldRewrite::Retain
            } else {
                HeaderFieldRewrite::Varint(replacement)
            })
        },
        2 => rewrite_packed_object_reference_order(
            payload,
            expected,
            reordered,
            reference_index,
            message_index,
        ),
        _ => Err(Error::invalid_archive(
            message_index,
            "object-reference field has an ambiguous wire type",
        )),
    }
}

fn rewrite_packed_object_reference_order(
    payload: &[u8],
    expected: &[u64],
    reordered: &[u64],
    reference_index: &mut usize,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let start_index = *reference_index;
    let mut cursor = 0usize;
    let mut output_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(
            payload.get(cursor..).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range is invalid")
            })?,
        )
        .map_err(|_| Error::invalid_archive(message_index, "malformed packed object reference"))?;
        let replacement = authorized_reordered_reference(
            expected,
            reordered,
            *reference_index,
            value,
            message_index,
        )?;
        output_length = output_length
            .checked_add(encoded_varint_width(replacement, encoded_length))
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference length overflow")
            })?;
        changed |= value != replacement;
        cursor = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        *reference_index = reference_index.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "object-reference count overflow")
        })?;
    }
    if !changed {
        return Ok(HeaderFieldRewrite::Retain);
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA reordered packed object references", output_length))?;
    cursor = 0;
    let mut current_index = start_index;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| {
                Error::invalid_archive(message_index, "malformed packed object reference")
            })?;
        let replacement = authorized_reordered_reference(
            expected,
            reordered,
            current_index,
            value,
            message_index,
        )?;
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        if value == replacement {
            output.extend_from_slice(payload.get(cursor..end).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range is invalid")
            })?);
        } else {
            let mut encoded = [0u8; MAX_VARINT_BYTES];
            output.extend_from_slice(encode_varint_with_width(
                replacement,
                encoded_length,
                &mut encoded,
            ));
        }
        cursor = end;
        current_index = current_index.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "object-reference count overflow")
        })?;
    }
    if output.len() != output_length || current_index != *reference_index {
        return Err(Error::invalid_archive(
            message_index,
            "packed object-reference rewrite length mismatch",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(output))
}

fn authorized_reordered_reference(
    expected: &[u64],
    reordered: &[u64],
    index: usize,
    raw_value: u64,
    message_index: usize,
) -> Result<u64> {
    let expected_value = expected.get(index).copied().ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "raw object-reference list exceeds authorization",
        )
    })?;
    if raw_value != expected_value {
        return Err(Error::invalid_archive(
            message_index,
            "raw object-reference order differs from authorization",
        ));
    }
    reordered.get(index).copied().ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "reordered object-reference list is truncated",
        )
    })
}

fn validate_reordered_data_reference_multiset(
    expected: &[u64],
    reordered: &[u64],
    message_index: usize,
) -> Result<()> {
    if expected.len() != reordered.len() {
        return Err(Error::invalid_archive(
            message_index,
            "reordered data-reference count changed",
        ));
    }

    let mut remaining = HashMap::new();
    remaining
        .try_reserve(expected.len())
        .map_err(|_| Error::allocation("IWA data-reference permutation", expected.len()))?;
    for identifier in expected {
        if *identifier == 0 {
            return Err(Error::invalid_archive(
                message_index,
                "data-reference identifiers must be nonzero",
            ));
        }
        let count = remaining.entry(*identifier).or_insert(0usize);
        *count = count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "data-reference multiplicity overflow")
        })?;
    }

    for identifier in reordered {
        if *identifier == 0 {
            return Err(Error::invalid_archive(
                message_index,
                "data-reference identifiers must be nonzero",
            ));
        }
        let count = remaining.get_mut(identifier).ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "reordered data-reference identifier is not in MessageInfo",
            )
        })?;
        if *count == 0 {
            return Err(Error::invalid_archive(
                message_index,
                "reordered data-reference multiplicity changed",
            ));
        }
        *count -= 1;
    }
    if remaining.values().any(|count| *count != 0) {
        return Err(Error::invalid_archive(
            message_index,
            "reordered data-reference multiplicity changed",
        ));
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies both old and new required MessageInfo scalars"
)]
fn rewrite_effective_message_scalars_and_references(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<Vec<u8>>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(current_type),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;

    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_type != replacement_type {
        assign_header_field_rewrite(
            &mut rewrites,
            type_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_type)),
            message_index,
        )?;
    }
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    for (field_index, field) in fields.iter().copied().enumerate() {
        match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let field_info = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                if let Some(rewritten) = rewrite_field_info_references(
                    field_info,
                    object_removals,
                    data_removals,
                    wire_limits,
                    limits,
                    message_index,
                )? {
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::LengthDelimited(rewritten),
                        message_index,
                    )?;
                }
            },
            5 if !object_removals.is_none() => {
                let rewrite =
                    rewrite_reference_field(source, field, object_removals, message_index)?;
                assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
            },
            6 if !data_removals.is_none() => {
                let rewrite = rewrite_reference_field(source, field, data_removals, message_index)?;
                assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
            },
            _ => {},
        }
    }

    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA pruned MessageInfo header",
    )
    .map(Some)
}

fn rewrite_field_info_references(
    source: &[u8],
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<Vec<u8>>> {
    let fields = parse_wire_fields(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        let removals = match field.number() {
            4 if !object_removals.is_none() => Some(object_removals),
            5 if !data_removals.is_none() => Some(data_removals),
            _ => None,
        };
        if let Some(removals) = removals {
            let rewrite = rewrite_reference_field(source, field, removals, message_index)?;
            assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
        }
    }
    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA pruned FieldInfo header",
    )
    .map(Some)
}

fn effective_required_varint_field(
    source: &[u8],
    fields: &[WireField],
    number: u32,
    expected: u64,
    message_index: usize,
    missing_reason: &'static str,
    mismatch_reason: &'static str,
) -> Result<usize> {
    let mut effective = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Error::invalid_archive(
                message_index,
                "MessageInfo contains an ambiguous required scalar",
            ));
        }
        let payload = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
            .map_err(|_| Error::invalid_archive(message_index, "malformed required scalar"))?;
        if encoded_length != payload.len() {
            return Err(Error::invalid_archive(
                message_index,
                "required scalar has trailing bytes",
            ));
        }
        effective = Some((field_index, value));
    }
    let (field_index, value) =
        effective.ok_or_else(|| Error::invalid_archive(message_index, missing_reason))?;
    if value != expected {
        return Err(Error::invalid_archive(message_index, mismatch_reason));
    }
    Ok(field_index)
}

fn rewrite_reference_field(
    source: &[u8],
    field: WireField,
    removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| {
                    Error::invalid_archive(message_index, "malformed unpacked reference")
                })?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "unpacked reference has trailing bytes",
                ));
            }
            Ok(if removals.contains(value) {
                HeaderFieldRewrite::Remove
            } else {
                HeaderFieldRewrite::Retain
            })
        },
        2 => rewrite_packed_references(payload, removals, message_index),
        _ => Err(Error::invalid_archive(
            message_index,
            "reference field has an ambiguous wire type",
        )),
    }
}

fn rewrite_packed_references(
    payload: &[u8],
    removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let mut cursor = 0usize;
    let mut retained_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let (value, encoded_length) =
            litchi_iwa_common::decode_varint_from_bytes(payload.get(cursor..).ok_or_else(
                || Error::invalid_archive(message_index, "packed reference range is invalid"),
            )?)
            .map_err(|_| Error::invalid_archive(message_index, "malformed packed reference"))?;
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        if removals.contains(value) {
            changed = true;
        } else {
            retained_length = retained_length.checked_add(encoded_length).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference length overflow")
            })?;
        }
        cursor = end;
    }
    if !changed {
        return Ok(HeaderFieldRewrite::Retain);
    }

    let mut retained = Vec::new();
    retained
        .try_reserve_exact(retained_length)
        .map_err(|_| Error::allocation("IWA retained packed object references", retained_length))?;
    cursor = 0;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| Error::invalid_archive(message_index, "malformed packed reference"))?;
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        if !removals.contains(value) {
            retained.extend_from_slice(payload.get(cursor..end).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range is invalid")
            })?);
        }
        cursor = end;
    }
    if retained.len() != retained_length {
        return Err(Error::invalid_archive(
            message_index,
            "packed reference rewrite length mismatch",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(retained))
}

fn retained_field_rewrites(length: usize) -> Result<Vec<HeaderFieldRewrite>> {
    let mut rewrites = Vec::new();
    rewrites
        .try_reserve_exact(length)
        .map_err(|_| Error::allocation("IWA header field rewrites", length))?;
    rewrites.extend((0..length).map(|_| HeaderFieldRewrite::Retain));
    Ok(rewrites)
}

fn assign_header_field_rewrite(
    rewrites: &mut [HeaderFieldRewrite],
    index: usize,
    rewrite: HeaderFieldRewrite,
    message_index: usize,
) -> Result<()> {
    let slot = rewrites
        .get_mut(index)
        .ok_or_else(|| Error::invalid_archive(message_index, "header rewrite index is invalid"))?;
    *slot = rewrite;
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "the assembler retains source-bound fields under format-level limits and diagnostics"
)]
fn assemble_header_field_rewrites(
    source: &[u8],
    fields: &[WireField],
    rewrites: &[HeaderFieldRewrite],
    header: HeaderKind,
    limits: Limits,
    message_index: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    if fields.len() != rewrites.len() {
        return Err(Error::invalid_archive(
            message_index,
            "header field rewrite count mismatch",
        ));
    }
    let mut output_length = 0usize;
    for (field, rewrite) in fields.iter().copied().zip(rewrites) {
        let field_length = match rewrite {
            HeaderFieldRewrite::Retain => field
                .raw(source)
                .map_err(|error| map_wire_error(error, header))?
                .len(),
            HeaderFieldRewrite::Remove => 0,
            HeaderFieldRewrite::Varint(value) => {
                let key_length = field
                    .key(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let preferred_width = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                key_length
                    .checked_add(encoded_varint_width(*value, preferred_width))
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "header rewrite overflow")
                    })?
            },
            HeaderFieldRewrite::LengthDelimited(payload) => {
                let key_length = field
                    .key(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let prefix_width = field
                    .payload_start()
                    .checked_sub(field.key_end())
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "field prefix range is invalid")
                    })?;
                let encoded_length = u64::try_from(payload.len()).map_err(|_| {
                    Error::invalid_archive(message_index, "field payload length exceeds u64")
                })?;
                key_length
                    .checked_add(encoded_varint_width(encoded_length, prefix_width))
                    .and_then(|length| length.checked_add(payload.len()))
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "header rewrite overflow")
                    })?
            },
            HeaderFieldRewrite::Raw(raw) => raw.len(),
        };
        output_length = output_length
            .checked_add(field_length)
            .ok_or_else(|| Error::invalid_archive(message_index, "header rewrite overflow"))?;
    }
    check_header_length(output_length, limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation(resource, output_length))?;
    for (field, rewrite) in fields.iter().copied().zip(rewrites) {
        match rewrite {
            HeaderFieldRewrite::Retain => output.extend_from_slice(
                field
                    .raw(source)
                    .map_err(|error| map_wire_error(error, header))?,
            ),
            HeaderFieldRewrite::Remove => {},
            HeaderFieldRewrite::Varint(value) => {
                output.extend_from_slice(
                    field
                        .key(source)
                        .map_err(|error| map_wire_error(error, header))?,
                );
                let preferred_width = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let mut encoded = [0u8; MAX_VARINT_BYTES];
                output.extend_from_slice(encode_varint_with_width(
                    *value,
                    preferred_width,
                    &mut encoded,
                ));
            },
            HeaderFieldRewrite::LengthDelimited(payload) => {
                output.extend_from_slice(
                    field
                        .key(source)
                        .map_err(|error| map_wire_error(error, header))?,
                );
                let prefix_width = field
                    .payload_start()
                    .checked_sub(field.key_end())
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "field prefix range is invalid")
                    })?;
                let encoded_length = u64::try_from(payload.len()).map_err(|_| {
                    Error::invalid_archive(message_index, "field payload length exceeds u64")
                })?;
                let mut encoded = [0u8; MAX_VARINT_BYTES];
                output.extend_from_slice(encode_varint_with_width(
                    encoded_length,
                    prefix_width,
                    &mut encoded,
                ));
                output.extend_from_slice(payload);
            },
            HeaderFieldRewrite::Raw(raw) => output.extend_from_slice(raw),
        }
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "header rewrite length mismatch",
        ));
    }
    Ok(output)
}

fn encoded_varint_width(value: u64, preferred_width: usize) -> usize {
    let canonical = if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    };
    canonical.max(preferred_width.min(MAX_VARINT_BYTES))
}

fn encode_varint_with_width(
    mut value: u64,
    preferred_width: usize,
    output: &mut [u8; MAX_VARINT_BYTES],
) -> &[u8] {
    let width = encoded_varint_width(value, preferred_width);
    for (index, byte) in output.iter_mut().take(width).enumerate() {
        *byte = u8::try_from(value & 0x7f).unwrap_or_default();
        value >>= 7;
        if index + 1 != width {
            *byte |= 0x80;
        }
    }
    &output[..width]
}

fn header_wire_limits(limits: Limits) -> Result<WireLimits> {
    let scanned_bytes = limits
        .max_header_bytes()
        .saturating_mul(4)
        .min(WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(scanned_bytes)
        .and_then(|profile| profile.with_fields(limits.max_header_fields()))
        .and_then(|profile| profile.with_nesting(limits.max_header_nesting()))
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))
}

fn try_copy_bytes(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn validate_raw_object_size(
    object: &ArchiveObject,
    header_length: usize,
    limits: Limits,
) -> Result<()> {
    check_header_length(header_length, limits)?;
    let mut object_length = varint_len(header_length)?
        .checked_add(header_length)
        .ok_or_else(|| Error::invalid_archive(0, "object prefix overflow"))?;
    for message in &object.messages {
        object_length = object_length
            .checked_add(message.data.len())
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
    }
    Ok(())
}

fn validate_raw_object_size_with_replacement(
    object: &ArchiveObject,
    message_index: usize,
    replacement_length: usize,
    header_length: usize,
    limits: Limits,
) -> Result<()> {
    if message_index >= object.messages.len() {
        return Err(Error::invalid_archive(
            message_index,
            "message index is out of bounds",
        ));
    }
    check_header_length(header_length, limits)?;
    let mut object_length = varint_len(header_length)?
        .checked_add(header_length)
        .ok_or_else(|| Error::invalid_archive(message_index, "object prefix overflow"))?;
    for (index, message) in object.messages.iter().enumerate() {
        let payload_length = if index == message_index {
            replacement_length
        } else {
            message.data.len()
        };
        object_length = object_length
            .checked_add(payload_length)
            .ok_or_else(|| Error::invalid_archive(message_index, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
    }
    Ok(())
}

fn verify_pruned_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo rewrite changed unrelated metadata",
        ));
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw ArchiveInfo rewrite changed another MessageInfo",
            ));
        }
    }
    verify_pruned_message_info(
        before_target,
        after_target,
        replacement_type,
        replacement_length,
        object_removals,
        data_removals,
        message_index,
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "directional verification keeps the authorized before and requested after orders explicit"
)]
fn verify_reordered_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    expected_object_references: &[u64],
    reordered_object_references: &[u64],
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo reorder changed unrelated metadata",
        ));
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw ArchiveInfo reorder changed another MessageInfo",
            ));
        }
    }
    if before_target.object_references != expected_object_references
        || after_target.object_references != reordered_object_references
        || after_target.type_ != replacement_type
        || after_target.length != replacement_length
        || before_target.versions != after_target.versions
        || before_target.field_infos != after_target.field_infos
        || before_target.data_references != after_target.data_references
        || before_target.base_message_index != after_target.base_message_index
        || before_target.diff_merge_version != after_target.diff_merge_version
        || before_target.diff_field_path != after_target.diff_field_path
        || before_target.fields_to_remove != after_target.fields_to_remove
        || before_target.diff_read_version != after_target.diff_read_version
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the authorized object-reference reorder",
        ));
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "directional verification keeps the authorized before and requested data-reference orders explicit"
)]
fn verify_reordered_data_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    expected_data_references: &[u64],
    reordered_data_references: &[u64],
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo data-reference reorder changed unrelated metadata",
        ));
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw ArchiveInfo data-reference reorder changed another MessageInfo",
            ));
        }
    }
    if before_target.data_references != expected_data_references
        || after_target.data_references != reordered_data_references
        || after_target.type_ != replacement_type
        || after_target.length != replacement_length
        || before_target.versions != after_target.versions
        || before_target.object_references != after_target.object_references
        || before_target.base_message_index != after_target.base_message_index
        || before_target.diff_merge_version != after_target.diff_merge_version
        || before_target.diff_field_path != after_target.diff_field_path
        || before_target.fields_to_remove != after_target.fields_to_remove
        || before_target.diff_read_version != after_target.diff_read_version
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the authorized data-reference reorder",
        ));
    }
    if before_target.field_infos.len() != after_target.field_infos.len() {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten FieldInfo count differs during data-reference reorder",
        ));
    }
    for (before_field, after_field) in before_target
        .field_infos
        .iter()
        .zip(&after_target.field_infos)
    {
        let expected_field_data_references =
            if before_field.data_references == expected_data_references {
                reordered_data_references
            } else {
                before_field.data_references.as_slice()
            };
        if before_field.path != after_field.path
            || before_field.r#type != after_field.r#type
            || before_field.unknown_field_rule != after_field.unknown_field_rule
            || before_field.known_field_rule != after_field.known_field_rule
            || before_field.known_field_version != after_field.known_field_version
            || before_field.known_field_feature_identifier
                != after_field.known_field_feature_identifier
            || before_field.object_references != after_field.object_references
            || after_field.data_references.as_slice() != expected_field_data_references
        {
            return Err(Error::invalid_archive(
                message_index,
                "rewritten FieldInfo does not match the authorized data-reference reorder",
            ));
        }
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "directional verification keeps both authorized states and message scalars explicit"
)]
fn verify_transitioned_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    transition: &PreparedObjectReferenceTransition<'_>,
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo transition changed unrelated metadata",
        ));
    }
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw transition changed another MessageInfo",
            ));
        }
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    let projected_aggregate = transition
        .aggregate
        .retained_after
        .iter()
        .chain(transition.aggregate.appended_after)
        .copied();
    let expected_field_count = match transition.canonical_fields {
        Some(CanonicalObjectReferenceFields::Insert(fields)) => transition
            .retained_field_count
            .checked_add(fields.len())
            .ok_or_else(|| {
                Error::invalid_archive(message_index, "canonical field count overflow")
            })?,
        _ => transition.retained_field_count,
    };
    if before_target.object_references.as_slice() != transition.aggregate.before
        || !projected_aggregate.eq(after_target.object_references.iter().copied())
        || after_target.type_ != replacement_type
        || after_target.length != replacement_length
        || before_target.versions != after_target.versions
        || after_target.field_infos.len() != expected_field_count
        || before_target.data_references != after_target.data_references
        || before_target.base_message_index != after_target.base_message_index
        || before_target.diff_merge_version != after_target.diff_merge_version
        || before_target.diff_field_path != after_target.diff_field_path
        || before_target.fields_to_remove != after_target.fields_to_remove
        || before_target.diff_read_version != after_target.diff_read_version
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the authorized reference transition",
        ));
    }
    for (field_index, (before_field, after_field)) in before_target
        .field_infos
        .iter()
        .zip(&after_target.field_infos)
        .take(transition.retained_field_count)
        .enumerate()
    {
        let selected = transition.fields.get(field_index).and_then(Option::as_ref);
        let before_references = selected
            .map_or(before_field.object_references.as_slice(), |selected| {
                selected.references.before
            });
        let after_matches = selected.map_or_else(
            || before_field.object_references == after_field.object_references,
            |selected| {
                selected
                    .references
                    .retained_after
                    .iter()
                    .chain(selected.references.appended_after)
                    .copied()
                    .eq(after_field.object_references.iter().copied())
            },
        );
        if before_field.object_references.as_slice() != before_references
            || !after_matches
            || before_field.path != after_field.path
            || before_field.r#type != after_field.r#type
            || before_field.unknown_field_rule != after_field.unknown_field_rule
            || before_field.known_field_rule != after_field.known_field_rule
            || before_field.known_field_version != after_field.known_field_version
            || before_field.known_field_feature_identifier
                != after_field.known_field_feature_identifier
            || before_field.data_references != after_field.data_references
        {
            return Err(Error::invalid_archive(
                message_index,
                "rewritten FieldInfo does not match the authorized reference transition",
            ));
        }
    }
    if let Some(CanonicalObjectReferenceFields::Insert(fields)) = transition.canonical_fields {
        for (field, expected) in after_target
            .field_infos
            .iter()
            .skip(transition.retained_field_count)
            .zip(fields)
        {
            if !canonical_object_field_matches(field, *expected) {
                return Err(Error::invalid_archive(
                    message_index,
                    "rewritten canonical object FieldInfo differs from authorization",
                ));
            }
        }
    }
    Ok(())
}

/// Verify a changed-cardinality data-reference transition against the decoded
/// archive metadata. This check is deliberately stricter than a protobuf
/// round-trip: only the requested aggregate list, selected nested fields, and
/// replacement message scalars may differ.
fn verify_transitioned_data_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    transition: DataReferenceTransition<'_>,
    canonical_plan: Option<CanonicalFieldDataReferencePlan<'_>>,
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo data-reference transition changed unrelated metadata",
        ));
    }
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw data-reference transition changed another MessageInfo",
            ));
        }
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    let expected_field_info_count = match canonical_plan {
        Some(CanonicalFieldDataReferencePlan::Insert { .. }) => before_target
            .field_infos
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(message_index, "FieldInfo count overflow"))?,
        Some(CanonicalFieldDataReferencePlan::Remove { .. }) => before_target
            .field_infos
            .len()
            .checked_sub(1)
            .ok_or_else(|| Error::invalid_archive(message_index, "FieldInfo count underflow"))?,
        None => before_target.field_infos.len(),
    };
    if before_target.data_references.as_slice() != transition.aggregate_before
        || after_target.data_references.as_slice() != transition.aggregate_after
        || after_target.type_ != replacement_type
        || after_target.length != replacement_length
        || before_target.versions != after_target.versions
        || before_target.object_references != after_target.object_references
        || before_target.base_message_index != after_target.base_message_index
        || before_target.diff_merge_version != after_target.diff_merge_version
        || before_target.diff_field_path != after_target.diff_field_path
        || before_target.fields_to_remove != after_target.fields_to_remove
        || before_target.diff_read_version != after_target.diff_read_version
        || after_target.field_infos.len() != expected_field_info_count
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the authorized data-reference transition",
        ));
    }
    if let Some(CanonicalFieldDataReferencePlan::Insert { path }) = canonical_plan {
        let field = after_target.field_infos.last().ok_or_else(|| {
            Error::invalid_archive(message_index, "canonical FieldInfo insertion is missing")
        })?;
        if field.path.as_slice() != path
            || field.r#type != Some(FieldType::DataReference)
            || field.unknown_field_rule.is_some()
            || !field.object_references.is_empty()
            || field.data_references.as_slice() != transition.aggregate_after
            || field.known_field_rule.is_some()
            || !field.known_field_version.is_empty()
            || field.known_field_feature_identifier.is_some()
        {
            return Err(Error::invalid_archive(
                message_index,
                "inserted FieldInfo is not canonical",
            ));
        }
    }
    if let Some(CanonicalFieldDataReferencePlan::Remove {
        field_info_index, ..
    }) = canonical_plan
    {
        if field_info_index >= before_target.field_infos.len() {
            return Err(Error::invalid_archive(
                message_index,
                "removed canonical FieldInfo index is invalid",
            ));
        }
        for (before_index, before_field) in before_target.field_infos.iter().enumerate() {
            if before_index == field_info_index {
                continue;
            }
            let after_index = if before_index < field_info_index {
                before_index
            } else {
                before_index - 1
            };
            let after_field = after_target.field_infos.get(after_index).ok_or_else(|| {
                Error::invalid_archive(message_index, "remaining FieldInfo is missing")
            })?;
            if before_field != after_field {
                return Err(Error::invalid_archive(
                    message_index,
                    "removing canonical FieldInfo changed another field",
                ));
            }
        }
    } else {
        for (field_index, (before_field, after_field)) in before_target
            .field_infos
            .iter()
            .zip(&after_target.field_infos)
            .enumerate()
        {
            let selected = transition
                .fields
                .iter()
                .find(|authorization| authorization.field_info_index == field_index);
            if let Some(authorization) = selected {
                if before_field.path.path.as_slice() != authorization.expected_path
                    || before_field.data_references.as_slice() != authorization.before
                    || after_field.data_references.as_slice() != authorization.after
                {
                    return Err(Error::invalid_archive(
                        message_index,
                        "selected FieldInfo does not match the authorized data-reference transition",
                    ));
                }
            } else if before_field.data_references != after_field.data_references {
                return Err(Error::invalid_archive(
                    message_index,
                    "unselected FieldInfo data references changed during transition",
                ));
            }
            if before_field.path != after_field.path
                || before_field.r#type != after_field.r#type
                || before_field.unknown_field_rule != after_field.unknown_field_rule
                || before_field.known_field_rule != after_field.known_field_rule
                || before_field.known_field_version != after_field.known_field_version
                || before_field.known_field_feature_identifier
                    != after_field.known_field_feature_identifier
                || before_field.object_references != after_field.object_references
            {
                return Err(Error::invalid_archive(
                    message_index,
                    "rewritten FieldInfo metadata changed outside data references",
                ));
            }
        }
    }
    Ok(())
}

fn verify_pruned_message_info(
    before: &MessageInfo,
    after: &MessageInfo,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<()> {
    if after.type_ != replacement_type
        || after.length != replacement_length
        || before.versions != after.versions
        || before.base_message_index != after.base_message_index
        || before.diff_merge_version != after.diff_merge_version
        || before.diff_field_path != after.diff_field_path
        || before.fields_to_remove != after.fields_to_remove
        || before.diff_read_version != after.diff_read_version
        || before.field_infos.len() != after.field_infos.len()
        || !references_match_after_pruning(
            &before.object_references,
            &after.object_references,
            object_removals,
        )
        || !references_match_after_pruning(
            &before.data_references,
            &after.data_references,
            data_removals,
        )
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the requested edit",
        ));
    }

    for (before_field, after_field) in before.field_infos.iter().zip(&after.field_infos) {
        if before_field.path != after_field.path
            || before_field.r#type != after_field.r#type
            || before_field.unknown_field_rule != after_field.unknown_field_rule
            || before_field.known_field_rule != after_field.known_field_rule
            || before_field.known_field_version != after_field.known_field_version
            || before_field.known_field_feature_identifier
                != after_field.known_field_feature_identifier
            || !references_match_after_pruning(
                &before_field.object_references,
                &after_field.object_references,
                object_removals,
            )
            || !references_match_after_pruning(
                &before_field.data_references,
                &after_field.data_references,
                data_removals,
            )
        {
            return Err(Error::invalid_archive(
                message_index,
                "rewritten FieldInfo does not match the requested reference pruning",
            ));
        }
    }
    Ok(())
}

fn references_match_after_pruning(
    before: &[u64],
    after: &[u64],
    removals: ReferenceRemovals<'_>,
) -> bool {
    before
        .iter()
        .copied()
        .filter(|identifier| !removals.contains(*identifier))
        .eq(after.iter().copied())
}

fn restore_replaced_message(
    object: &mut ArchiveObject,
    index: usize,
    old: RawMessage,
    old_type: u32,
    old_length: u32,
) {
    drop(std::mem::replace(&mut object.messages[index], old));
    let info = &mut object.archive_info.message_infos[index];
    info.type_ = old_type;
    info.length = old_length;
}

fn archive_info_from_codec_with_admission(
    value: archive_codec::ArchiveInfo,
    limits: Limits,
) -> Result<(ArchiveInfo, ArchiveInfoAdmission)> {
    if value.identifier.is_none() {
        return Err(Error::invalid_archive(
            0,
            "object is missing its archive identifier",
        ));
    }
    if value.message_infos.len() > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            value.message_infos.len(),
            limits.max_messages_per_object(),
        ));
    }
    let mut message_infos = Vec::new();
    message_infos
        .try_reserve_exact(value.message_infos.len())
        .map_err(|_| {
            Error::allocation("IWA neutral message metadata", value.message_infos.len())
        })?;
    let mut collection_admission = MessageInfoCollectionAdmission::empty();
    for message_info in value.message_infos {
        let message_info = message_info_from_codec(message_info)?;
        let admission = message_info_admission(&message_info, limits)?;
        collection_admission.push(admission, limits)?;
        message_infos.push(message_info);
    }
    let admission =
        collection_admission.archive_info_with_options(value.identifier, value.should_merge)?;
    admission.check(limits)?;
    Ok((
        ArchiveInfo {
            identifier: value.identifier,
            message_infos,
            should_merge: value.should_merge,
        },
        admission,
    ))
}

fn archive_info_to_codec(value: &ArchiveInfo) -> Result<archive_codec::ArchiveInfo> {
    let mut message_infos = Vec::new();
    message_infos
        .try_reserve_exact(value.message_infos.len())
        .map_err(|_| {
            Error::allocation(
                "IWA compatibility message metadata",
                value.message_infos.len(),
            )
        })?;
    for message_info in &value.message_infos {
        message_infos.push(message_info_to_codec(message_info)?);
    }
    Ok(archive_codec::ArchiveInfo {
        identifier: value.identifier,
        message_infos,
        should_merge: value.should_merge,
    })
}

fn message_info_from_codec(value: archive_codec::MessageInfo) -> Result<MessageInfo> {
    let mut field_infos = Vec::new();
    field_infos
        .try_reserve_exact(value.field_infos.len())
        .map_err(|_| Error::allocation("IWA neutral field metadata", value.field_infos.len()))?;
    for field_info in value.field_infos {
        field_infos.push(field_info_from_codec(field_info));
    }

    let mut fields_to_remove = Vec::new();
    fields_to_remove
        .try_reserve_exact(value.fields_to_remove.len())
        .map_err(|_| {
            Error::allocation(
                "IWA neutral removed field paths",
                value.fields_to_remove.len(),
            )
        })?;
    for field_path in value.fields_to_remove {
        fields_to_remove.push(field_path_from_codec(field_path));
    }

    Ok(MessageInfo {
        type_: value.r#type,
        versions: SmallVec::from_vec(value.version),
        length: value.length,
        field_infos,
        object_references: value.object_references,
        data_references: value.data_references,
        base_message_index: value.base_message_index,
        diff_merge_version: value.diff_merge_version,
        diff_field_path: value.diff_field_path.map(field_path_from_codec),
        fields_to_remove,
        diff_read_version: value.diff_read_version,
    })
}

fn message_info_to_codec(value: &MessageInfo) -> Result<archive_codec::MessageInfo> {
    let mut field_infos = Vec::new();
    field_infos
        .try_reserve_exact(value.field_infos.len())
        .map_err(|_| {
            Error::allocation("IWA compatibility field metadata", value.field_infos.len())
        })?;
    for field_info in &value.field_infos {
        field_infos.push(field_info_to_codec(field_info)?);
    }

    let mut fields_to_remove = Vec::new();
    fields_to_remove
        .try_reserve_exact(value.fields_to_remove.len())
        .map_err(|_| {
            Error::allocation(
                "IWA compatibility removed field paths",
                value.fields_to_remove.len(),
            )
        })?;
    for field_path in &value.fields_to_remove {
        fields_to_remove.push(field_path_to_codec(field_path)?);
    }

    Ok(archive_codec::MessageInfo {
        r#type: value.type_,
        version: try_copy_slice(&value.versions, "IWA compatibility message versions")?,
        length: value.length,
        field_infos,
        object_references: try_copy_slice(
            &value.object_references,
            "IWA compatibility object references",
        )?,
        data_references: try_copy_slice(
            &value.data_references,
            "IWA compatibility data references",
        )?,
        base_message_index: value.base_message_index,
        diff_merge_version: try_copy_slice(
            &value.diff_merge_version,
            "IWA compatibility diff merge versions",
        )?,
        diff_field_path: value
            .diff_field_path
            .as_ref()
            .map(field_path_to_codec)
            .transpose()?,
        fields_to_remove,
        diff_read_version: try_copy_slice(
            &value.diff_read_version,
            "IWA compatibility diff read versions",
        )?,
    })
}

fn field_info_from_codec(value: archive_codec::FieldInfo) -> FieldInfo {
    FieldInfo {
        path: field_path_from_codec(value.path),
        r#type: value.r#type.map(FieldType::from_raw),
        unknown_field_rule: value.unknown_field_rule.map(UnknownFieldRule::from_raw),
        object_references: value.object_references,
        data_references: value.data_references,
        known_field_rule: value.known_field_rule.map(KnownFieldRule::from_raw),
        known_field_version: value.known_field_version,
        known_field_feature_identifier: value.known_field_feature_identifier,
    }
}

fn field_info_to_codec(value: &FieldInfo) -> Result<archive_codec::FieldInfo> {
    Ok(archive_codec::FieldInfo {
        path: field_path_to_codec(&value.path)?,
        r#type: value.r#type.map(FieldType::raw_value),
        unknown_field_rule: value.unknown_field_rule.map(UnknownFieldRule::raw_value),
        object_references: try_copy_slice(
            &value.object_references,
            "IWA compatibility field object references",
        )?,
        data_references: try_copy_slice(
            &value.data_references,
            "IWA compatibility field data references",
        )?,
        known_field_rule: value.known_field_rule.map(KnownFieldRule::raw_value),
        known_field_version: try_copy_slice(
            &value.known_field_version,
            "IWA compatibility known field versions",
        )?,
        known_field_feature_identifier: value
            .known_field_feature_identifier
            .as_deref()
            .map(|identifier| {
                try_copy_string(identifier, "IWA compatibility field feature identifier")
            })
            .transpose()?,
    })
}

fn field_path_from_codec(value: archive_codec::FieldPath) -> FieldPath {
    FieldPath { path: value.path }
}

fn field_path_to_codec(value: &FieldPath) -> Result<archive_codec::FieldPath> {
    Ok(archive_codec::FieldPath {
        path: try_copy_slice(&value.path, "IWA compatibility field path")?,
    })
}

fn try_copy_slice<T: Copy>(source: &[T], resource: &'static str) -> Result<Vec<T>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn try_copy_versions(
    source: &[u32],
    resource: &'static str,
) -> Result<SmallVec<[u32; MESSAGE_VERSIONS_INLINE_CAPACITY]>> {
    let mut output = SmallVec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn try_copy_string(source: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.push_str(source);
    Ok(output)
}

fn check_header_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_header_bytes() {
        return Err(limit(
            LimitKind::HeaderBytes,
            length,
            limits.max_header_bytes(),
        ));
    }
    Ok(())
}

fn encode_archive_info(info: &ArchiveInfo, limits: Limits) -> Result<Vec<u8>> {
    let encoded_info = archive_info_to_codec(info)?;
    encode_archive_info_codec(&encoded_info, limits)
}

fn encode_archive_info_for_object(object: &ArchiveObject, limits: Limits) -> Result<Vec<u8>> {
    let mut encoded_info = archive_info_to_codec(&object.archive_info)?;
    if encoded_info.message_infos.len() != object.messages.len() {
        return Err(Error::invalid_archive(
            0,
            "message metadata and payload counts differ",
        ));
    }
    for (message_info, message) in encoded_info.message_infos.iter_mut().zip(&object.messages) {
        message_info.r#type = message.type_;
        message_info.length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
    }
    encode_archive_info_codec(&encoded_info, limits)
}

fn encode_archive_info_codec(
    encoded_info: &archive_codec::ArchiveInfo,
    limits: Limits,
) -> Result<Vec<u8>> {
    let header_length = usize::try_from(
        archive_codec::archive_info_encoded_len(encoded_info).map_err(|error| {
            Error::header_codec(
                HeaderKind::ArchiveInfo,
                HeaderOperation::Encode,
                error.to_string(),
            )
        })?,
    )
    .map_err(|_| Error::invalid_archive(0, "ArchiveInfo header length exceeds usize"))?;
    check_header_length(header_length, limits)?;
    let mut header = Vec::new();
    header
        .try_reserve_exact(header_length)
        .map_err(|_| Error::allocation("IWA ArchiveInfo header", header_length))?;
    let maximum = u32::try_from(limits.max_header_bytes())
        .map_err(|_| Error::invalid_archive(0, "header byte limit exceeds u32"))?;
    archive_codec::encode_archive_info(encoded_info, maximum, &mut header).map_err(|error| {
        Error::header_codec(
            HeaderKind::ArchiveInfo,
            HeaderOperation::Encode,
            error.to_string(),
        )
    })?;
    Ok(header)
}

fn archive_info_encoded_len(info: &ArchiveInfo) -> Result<usize> {
    let mut total = 0usize;
    if let Some(identifier) = info.identifier {
        total = checked_encoded_add(total, varint_field_length(1, identifier)?)?;
    }
    for message_info in &info.message_infos {
        let message_length = usize::try_from(message_info.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        let payload_length =
            encoded_message_info_length(message_info, message_info.type_, message_length)?;
        total = checked_encoded_add(total, length_delimited_field_length(2, payload_length)?)?;
    }
    if let Some(should_merge) = info.should_merge {
        total = checked_encoded_add(total, varint_field_length(3, u64::from(should_merge))?)?;
    }
    Ok(total)
}

/// A no-allocation sink for the canonical ArchiveInfo encoding.
///
/// When `target` is absent this only counts bytes. When present it also
/// compares the bytes emitted by the neutral projection with the retained
/// canonical header, which is the source-authoritative test used by
/// `to_bytes` to decide whether an original raw header remains
/// valid after an in-memory mutation.
struct CanonicalHeaderProbe<'a> {
    target: Option<&'a [u8]>,
    offset: usize,
    matches: bool,
}

impl<'a> CanonicalHeaderProbe<'a> {
    fn new(target: Option<&'a [u8]>) -> Self {
        Self {
            target,
            offset: 0,
            matches: target.is_some(),
        }
    }

    fn bytes(&mut self, bytes: &[u8]) -> Result<()> {
        let end = self
            .offset
            .checked_add(bytes.len())
            .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo encoded length overflow"))?;
        if self.matches {
            let equal = self
                .target
                .and_then(|target| target.get(self.offset..end))
                .is_some_and(|target| short_bytes_equal(target, bytes));
            if !equal {
                self.matches = false;
            }
        }
        self.offset = end;
        Ok(())
    }

    fn is_comparing(&self) -> bool {
        self.target.is_some() && self.matches
    }

    /// Read a canonical length from the target without doing a second walk of
    /// the current metadata. The target is generated by the same canonical
    /// encoder, so comparing the target's exact varint spelling through
    /// `varint` below still rejects noncanonical or truncated targets.
    fn target_varint(&self) -> Option<u64> {
        let target = self.target?;
        let bytes = target.get(self.offset..)?;
        let mut value = 0u64;
        for (index, byte) in bytes.iter().copied().enumerate() {
            if index >= MAX_VARINT_BYTES {
                return None;
            }
            let shift = index * 7;
            if index == MAX_VARINT_BYTES - 1 && byte > 1 {
                return None;
            }
            value |= u64::from(byte & 0x7f) << shift;
            if byte & 0x80 == 0 {
                return Some(value);
            }
        }
        None
    }

    fn resolve_length(&mut self, computed: Option<usize>) -> Result<usize> {
        if let Some(length) = computed {
            return Ok(length);
        }
        if self.matches {
            let length = self
                .target_varint()
                .and_then(|length| usize::try_from(length).ok());
            if let Some(length) = length {
                return Ok(length);
            }
            self.matches = false;
        }
        // A comparison that already diverged only needs a placeholder wide
        // enough to keep counting while the current payload is emitted. The
        // length_delimited_field caller replaces this prefix width with the
        // current payload width after the closure returns.
        Ok(0)
    }

    fn varint(&mut self, value: u64) -> Result<()> {
        let mut encoded = [0u8; MAX_VARINT_BYTES];
        self.bytes(encode_varint(value, &mut encoded))
    }

    fn field_key(&mut self, number: u32, wire_type: u8) -> Result<()> {
        self.varint((u64::from(number) << 3) | u64::from(wire_type))
    }

    fn varint_field(&mut self, number: u32, value: u64) -> Result<()> {
        self.field_key(number, 0)?;
        self.varint(value)
    }

    fn length_delimited_field<F>(
        &mut self,
        number: u32,
        payload_length: Option<usize>,
        write_payload: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut Self) -> Result<()>,
    {
        self.field_key(number, 2)?;
        let length_start = self.offset;
        let payload_length = self.resolve_length(payload_length)?;
        self.varint(
            u64::try_from(payload_length)
                .map_err(|_| Error::invalid_archive(0, "ArchiveInfo field length exceeds u64"))?,
        )?;
        let payload_start = self.offset;
        write_payload(self)?;
        let emitted_payload_length = self.offset.checked_sub(payload_start).ok_or_else(|| {
            Error::invalid_archive(0, "ArchiveInfo nested encoded length overflow")
        })?;
        if emitted_payload_length != payload_length {
            self.matches = false;
            let current_length_prefix = varint_len(emitted_payload_length)?;
            self.offset = length_start
                .checked_add(current_length_prefix)
                .and_then(|offset| offset.checked_add(emitted_payload_length))
                .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo encoded length overflow"))?;
        }
        Ok(())
    }

    fn finish(self) -> (usize, bool) {
        let matches = self
            .target
            .is_some_and(|target| self.matches && self.offset == target.len());
        (self.offset, matches)
    }
}

/// Position of the first object whose identifier already occurred earlier.
///
/// This is the object at which the ordered preflight reports a duplicate
/// identifier. Only the prefix the preflight can reach (the object ceiling)
/// is considered, and objects without an identifier are skipped because the
/// preflight refuses them at their own position first. Small archives use an
/// allocation-free pairwise scan; larger ones sort one fallibly reserved
/// `(identifier, position)` vector instead of comparing every earlier object.
fn first_duplicate_identifier(objects: &[ArchiveObject], limits: Limits) -> Result<Option<usize>> {
    const PAIRWISE_OBJECTS: usize = 32;
    let reachable = &objects[..objects.len().min(limits.max_objects())];
    if reachable.len() <= PAIRWISE_OBJECTS {
        return Ok(reachable.iter().enumerate().position(|(index, object)| {
            object.archive_info.identifier.is_some_and(|identifier| {
                reachable[..index]
                    .iter()
                    .any(|previous| previous.archive_info.identifier == Some(identifier))
            })
        }));
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(reachable.len())
        .map_err(|_| Error::allocation("IWA object identifier index", reachable.len()))?;
    entries.extend(
        reachable
            .iter()
            .enumerate()
            .filter_map(|(index, object)| object.archive_info.identifier.map(|id| (id, index))),
    );
    entries.sort_unstable();
    Ok(entries
        .windows(2)
        .filter(|pair| pair[0].0 == pair[1].0)
        .map(|pair| pair[1].1)
        .min())
}

/// Compare two equal-length slices of probe output.
///
/// The probe compares one varint or key (at most ten bytes) or one packed
/// payload at a time. A library `memcmp` call costs a data-dependent branch
/// per call on such short inputs; an OR-reduction of byte differences is
/// branch-free for them and falls back to the slice comparison otherwise.
#[inline]
fn short_bytes_equal(left: &[u8], right: &[u8]) -> bool {
    if left.len() != right.len() {
        return false;
    }
    if left.len() > 16 {
        return left == right;
    }
    left.iter()
        .zip(right)
        .fold(0u8, |difference, (left, right)| difference | (left ^ right))
        == 0
}

fn encoded_archive_len_preflight(objects: &[ArchiveObject], limits: Limits) -> Result<usize> {
    let first_duplicate = first_duplicate_identifier(objects, limits)?;
    encoded_archive_len_preflight_with_duplicate(objects, limits, first_duplicate)
}

fn encoded_archive_len_preflight_with_duplicate(
    objects: &[ArchiveObject],
    limits: Limits,
    first_duplicate: Option<usize>,
) -> Result<usize> {
    let mut archive_length = 0usize;
    let mut total_messages = 0usize;

    for (object_index, object) in objects.iter().enumerate() {
        if object_index >= limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                object_index.saturating_add(1),
                limits.max_objects(),
            ));
        }
        object.archive_info.identifier.ok_or_else(|| {
            Error::invalid_archive(object_index, "object is missing its archive identifier")
        })?;
        if first_duplicate == Some(object_index) {
            return Err(Error::invalid_archive(
                object_index,
                "duplicate object identifier",
            ));
        }
        if object.archive_info.message_infos.len() != object.messages.len() {
            return Err(Error::invalid_archive(
                object_index,
                "message metadata and payload counts differ",
            ));
        }
        object.archive_info.validate(limits)?;
        add_limited(
            &mut total_messages,
            object.messages.len(),
            LimitKind::Messages,
            limits.max_messages(),
        )?;

        let (canonical_header_length, header_length) =
            retained_header_length_for_object(object, limits)?;
        check_header_length(canonical_header_length, limits)?;
        check_header_length(header_length, limits)?;

        let mut payload_length = 0usize;
        for message in &object.messages {
            check_message_length(message.data.len(), limits)?;
            add_limited(
                &mut payload_length,
                message.data.len(),
                LimitKind::ObjectBytes,
                limits.max_object_bytes(),
            )?;
        }
        let object_length = varint_len(header_length)?
            .checked_add(header_length)
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(object_index, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        add_limited(
            &mut archive_length,
            object_length,
            LimitKind::ArchiveBytes,
            limits.max_archive_bytes(),
        )?;
    }
    Ok(archive_length)
}

fn prepare_archive_encoding_requirements<V>(
    objects: &[ArchiveObject],
    limits: Limits,
    visitor: &mut V,
) -> ArchiveEncodingPreparationResult<ArchiveEncodingRequirements, V::Error>
where
    V: ArchiveEncodingPreparationVisitor,
{
    // Duplicate identity is checked by the exact fallibly reserved index in
    // `encode_prepared_archive`. Preparation intentionally does not allocate
    // an index just to run a second census; all other framing and metadata
    // checks remain allocation-free here.
    let mut archive_length = 0usize;
    let mut header_bytes = 0usize;
    let mut payload_bytes = 0usize;
    let mut work_bytes = 0usize;
    let mut scratch_bytes = 0usize;
    let mut message_count = 0usize;
    let mut total_messages = 0usize;
    let mut allocation_events = 0usize;

    for (object_index, object) in objects.iter().enumerate() {
        if object_index >= limits.max_objects() {
            return Err(ArchiveEncodingPreparationError::Core(limit(
                LimitKind::Objects,
                object_index.saturating_add(1),
                limits.max_objects(),
            )));
        }
        object.archive_info.identifier.ok_or_else(|| {
            Error::invalid_archive(object_index, "object is missing its archive identifier")
        })?;
        if object.archive_info.message_infos.len() != object.messages.len() {
            return Err(ArchiveEncodingPreparationError::Core(
                Error::invalid_archive(object_index, "message metadata and payload counts differ"),
            ));
        }
        let preparation = admit_archive_encoding_object(object, visitor)?;
        visitor
            .admit(ArchiveEncodingPreparationChunk {
                kind: ArchiveEncodingPreparationKind::SemanticHeader,
                elements: 1,
                traversal_bytes: preparation.metadata_traversal_bytes,
            })
            .map_err(ArchiveEncodingPreparationError::Admission)?;
        let admission = check_archive_info_constructor_for_messages(
            &object.archive_info,
            &object.messages,
            limits,
        )?;
        add_limited(
            &mut total_messages,
            object.messages.len(),
            LimitKind::Messages,
            limits.max_messages(),
        )?;
        visitor
            .admit(ArchiveEncodingPreparationChunk {
                kind: ArchiveEncodingPreparationKind::RetainedHeader,
                elements: 1,
                traversal_bytes: admission.header_length,
            })
            .map_err(ArchiveEncodingPreparationError::Admission)?;
        let (canonical_header_length, header_length) =
            retained_header_length_for_object(object, limits)?;
        if admission.header_length != canonical_header_length {
            return Err(ArchiveEncodingPreparationError::Core(
                Error::invalid_archive(
                    object_index,
                    "prepared ArchiveInfo length disagrees with canonical emission",
                ),
            ));
        }
        check_header_length(canonical_header_length, limits)?;
        check_header_length(header_length, limits)?;
        header_bytes = header_bytes
            .checked_add(header_length)
            .ok_or_else(|| Error::invalid_archive(object_index, "header byte count overflow"))?;

        visitor
            .admit(ArchiveEncodingPreparationChunk {
                kind: ArchiveEncodingPreparationKind::PayloadValidation,
                elements: object.messages.len(),
                traversal_bytes: preparation.payload_validation_bytes,
            })
            .map_err(ArchiveEncodingPreparationError::Admission)?;
        let mut object_payload_bytes = 0usize;
        for message in &object.messages {
            check_message_length(message.data.len(), limits)?;
            object_payload_bytes = object_payload_bytes
                .checked_add(message.data.len())
                .ok_or_else(|| {
                    Error::invalid_archive(object_index, "payload byte count overflow")
                })?;
        }
        payload_bytes = payload_bytes
            .checked_add(object_payload_bytes)
            .ok_or_else(|| Error::invalid_archive(object_index, "payload byte count overflow"))?;
        message_count = message_count
            .checked_add(object.messages.len())
            .ok_or_else(|| Error::invalid_archive(object_index, "message count overflow"))?;

        let object_length = varint_len(header_length)?
            .checked_add(header_length)
            .and_then(|length| length.checked_add(object_payload_bytes))
            .ok_or_else(|| Error::invalid_archive(object_index, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(ArchiveEncodingPreparationError::Core(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            )));
        }
        add_limited(
            &mut archive_length,
            object_length,
            LimitKind::ArchiveBytes,
            limits.max_archive_bytes(),
        )?;

        visitor
            .admit(ArchiveEncodingPreparationChunk {
                kind: ArchiveEncodingPreparationKind::CompatibilityProjectionSizing,
                elements: 1,
                traversal_bytes: preparation.metadata_traversal_bytes,
            })
            .map_err(ArchiveEncodingPreparationError::Admission)?;
        let compatibility_bytes = archive_info_codec_memory(&object.archive_info)?;
        let object_scratch = compatibility_bytes
            .checked_add(canonical_header_length)
            .ok_or_else(|| Error::invalid_archive(object_index, "encoding scratch overflow"))?;
        scratch_bytes = scratch_bytes
            .checked_add(object_scratch)
            .ok_or_else(|| Error::invalid_archive(object_index, "encoding scratch overflow"))?;
        work_bytes = work_bytes
            .checked_add(admission.header_memory)
            .and_then(|value| value.checked_add(header_length))
            .and_then(|value| value.checked_add(object_payload_bytes))
            .ok_or_else(|| Error::invalid_archive(object_index, "encoding work overflow"))?;
        visitor
            .admit(ArchiveEncodingPreparationChunk {
                kind: ArchiveEncodingPreparationKind::CompatibilityAllocationSizing,
                elements: 1,
                traversal_bytes: preparation.metadata_traversal_bytes,
            })
            .map_err(ArchiveEncodingPreparationError::Admission)?;
        let compatibility_allocations = archive_info_codec_allocations(&object.archive_info)?;
        allocation_events = allocation_events
            .checked_add(compatibility_allocations)
            .and_then(|value| value.checked_add(usize::from(canonical_header_length != 0)))
            .ok_or_else(|| Error::invalid_archive(object_index, "allocation count overflow"))?;
    }

    let encoded_bytes = archive_length;
    let retained_bytes = encoded_bytes;
    allocation_events = allocation_events
        .checked_add(usize::from(encoded_bytes != 0))
        .ok_or_else(|| Error::invalid_archive(0, "allocation count overflow"))?;
    let identifier_index_entries = objects.len();
    if identifier_index_entries != 0 {
        allocation_events = allocation_events
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "allocation count overflow"))?;
    }
    let identifier_index_bytes = identifier_index_entries
        .checked_mul(size_of::<(u64, usize)>())
        .ok_or_else(|| Error::invalid_archive(0, "identifier index size overflow"))?;
    scratch_bytes = scratch_bytes
        .checked_add(identifier_index_bytes)
        .ok_or_else(|| Error::invalid_archive(0, "encoding scratch overflow"))?;
    work_bytes = work_bytes
        .checked_add(identifier_index_entries)
        .ok_or_else(|| Error::invalid_archive(0, "encoding work overflow"))?;
    Ok(ArchiveEncodingRequirements {
        encoded_bytes,
        object_count: objects.len(),
        message_count,
        header_bytes,
        payload_bytes,
        work_bytes,
        scratch_bytes,
        retained_bytes,
        identifier_index_entries,
        allocation_events,
    })
}

fn encode_prepared_archive(
    archive: &Archive,
    limits: Limits,
    requirements: ArchiveEncodingRequirements,
) -> Result<Vec<u8>> {
    if archive.objects.len() != requirements.object_count {
        return Err(Error::invalid_archive(
            0,
            "prepared archive object count changed",
        ));
    }
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(requirements.identifier_index_entries)
        .map_err(|_| {
            Error::allocation(
                "prepared IWA archive identifiers",
                requirements.identifier_index_entries * size_of::<(u64, usize)>(),
            )
        })?;
    for (object_index, object) in archive.objects.iter().enumerate() {
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::invalid_archive(object_index, "object is missing its archive identifier")
        })?;
        identifiers.push((identifier, object_index));
    }
    identifiers.sort_unstable_by_key(|(identifier, _)| *identifier);
    if let Some(pair) = identifiers.windows(2).find(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::invalid_archive(
            pair[1].1,
            "duplicate object identifier",
        ));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(requirements.encoded_bytes)
        .map_err(|_| {
            Error::allocation("IWA prepared archive output", requirements.encoded_bytes)
        })?;
    for (object_index, object) in archive.objects.iter().enumerate() {
        let canonical_header = encode_archive_info_for_object(object, limits)?;
        let header = match object.source_headers() {
            Some((original, canonical)) if canonical == canonical_header.as_slice() => original,
            _ => canonical_header.as_slice(),
        };
        let header_length = header.len();
        let prefix_length = varint_len(header_length)?;
        let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
            total
                .checked_add(message.data.len())
                .ok_or_else(|| Error::invalid_archive(object_index, "payload length overflow"))
        })?;
        let object_length = prefix_length
            .checked_add(header_length)
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(object_index, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        let next_length = output
            .len()
            .checked_add(object_length)
            .ok_or_else(|| Error::invalid_archive(object_index, "archive length overflow"))?;
        if next_length > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                next_length,
                limits.max_archive_bytes(),
            ));
        }
        let mut prefix = [0u8; MAX_VARINT_BYTES];
        let prefix = encode_varint(
            u64::try_from(header_length)
                .map_err(|_| Error::invalid_archive(object_index, "header length exceeds u64"))?,
            &mut prefix,
        );
        output.extend_from_slice(prefix);
        output.extend_from_slice(header);
        for message in &object.messages {
            output.extend_from_slice(&message.data);
        }
    }
    if output.len() != requirements.encoded_bytes {
        return Err(Error::invalid_archive(
            0,
            "prepared archive output length changed",
        ));
    }
    Ok(output)
}

fn add_memory_bytes(total: &mut usize, length: usize, element_size: usize) -> Result<()> {
    let bytes = length
        .checked_mul(element_size)
        .ok_or_else(|| Error::invalid_archive(0, "encoding allocation size overflow"))?;
    *total = checked_memory_add(*total, bytes)?;
    Ok(())
}

fn add_allocation_count(total: &mut usize, present: bool) -> Result<()> {
    if present {
        *total = total
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "encoding allocation count overflow"))?;
    }
    Ok(())
}

fn field_info_codec_memory(info: &FieldInfo) -> Result<usize> {
    let mut total = size_of::<archive_codec::FieldInfo>();
    add_memory_bytes(&mut total, info.path.path.len(), size_of::<u32>())?;
    add_memory_bytes(&mut total, info.object_references.len(), size_of::<u64>())?;
    add_memory_bytes(&mut total, info.data_references.len(), size_of::<u64>())?;
    add_memory_bytes(&mut total, info.known_field_version.len(), size_of::<u32>())?;
    total = checked_memory_add(
        total,
        info.known_field_feature_identifier
            .as_deref()
            .map_or(0, str::len),
    )?;
    Ok(total)
}

fn message_info_codec_memory(info: &MessageInfo) -> Result<usize> {
    let mut total = size_of::<archive_codec::MessageInfo>();
    add_memory_bytes(&mut total, info.versions.len(), size_of::<u32>())?;
    for field_info in &info.field_infos {
        total = checked_memory_add(total, field_info_codec_memory(field_info)?)?;
    }
    add_memory_bytes(&mut total, info.object_references.len(), size_of::<u64>())?;
    add_memory_bytes(&mut total, info.data_references.len(), size_of::<u64>())?;
    add_memory_bytes(&mut total, info.diff_merge_version.len(), size_of::<u32>())?;
    if let Some(path) = &info.diff_field_path {
        add_memory_bytes(&mut total, path.path.len(), size_of::<u32>())?;
    }
    add_memory_bytes(
        &mut total,
        info.fields_to_remove.len(),
        size_of::<archive_codec::FieldPath>(),
    )?;
    for path in &info.fields_to_remove {
        add_memory_bytes(&mut total, path.path.len(), size_of::<u32>())?;
    }
    add_memory_bytes(&mut total, info.diff_read_version.len(), size_of::<u32>())?;
    Ok(total)
}

fn archive_info_codec_memory(info: &ArchiveInfo) -> Result<usize> {
    let mut total = size_of::<archive_codec::ArchiveInfo>();
    for message_info in &info.message_infos {
        total = checked_memory_add(total, message_info_codec_memory(message_info)?)?;
    }
    Ok(total)
}

fn field_info_codec_allocations(info: &FieldInfo) -> Result<usize> {
    let mut total = 0usize;
    add_allocation_count(&mut total, !info.path.path.is_empty())?;
    add_allocation_count(&mut total, !info.object_references.is_empty())?;
    add_allocation_count(&mut total, !info.data_references.is_empty())?;
    add_allocation_count(&mut total, !info.known_field_version.is_empty())?;
    add_allocation_count(
        &mut total,
        info.known_field_feature_identifier
            .as_deref()
            .is_some_and(|value| !value.is_empty()),
    )?;
    Ok(total)
}

fn message_info_codec_allocations(info: &MessageInfo) -> Result<usize> {
    let mut total = 0usize;
    add_allocation_count(&mut total, !info.versions.is_empty())?;
    add_allocation_count(&mut total, !info.field_infos.is_empty())?;
    for field_info in &info.field_infos {
        total = total
            .checked_add(field_info_codec_allocations(field_info)?)
            .ok_or_else(|| Error::invalid_archive(0, "encoding allocation count overflow"))?;
    }
    add_allocation_count(&mut total, !info.object_references.is_empty())?;
    add_allocation_count(&mut total, !info.data_references.is_empty())?;
    add_allocation_count(&mut total, !info.diff_merge_version.is_empty())?;
    if let Some(path) = &info.diff_field_path {
        add_allocation_count(&mut total, !path.path.is_empty())?;
    }
    add_allocation_count(&mut total, !info.fields_to_remove.is_empty())?;
    for path in &info.fields_to_remove {
        add_allocation_count(&mut total, !path.path.is_empty())?;
    }
    add_allocation_count(&mut total, !info.diff_read_version.is_empty())?;
    Ok(total)
}

fn archive_info_codec_allocations(info: &ArchiveInfo) -> Result<usize> {
    let mut total = 0usize;
    add_allocation_count(&mut total, !info.message_infos.is_empty())?;
    for message_info in &info.message_infos {
        total = total
            .checked_add(message_info_codec_allocations(message_info)?)
            .ok_or_else(|| Error::invalid_archive(0, "encoding allocation count overflow"))?;
    }
    Ok(total)
}

fn canonical_archive_info_length_for_object(object: &ArchiveObject) -> Result<usize> {
    let mut probe = CanonicalHeaderProbe::new(None);
    emit_canonical_archive_info(&mut probe, object)?;
    Ok(probe.finish().0)
}

fn retained_header_length_for_object(
    object: &ArchiveObject,
    limits: Limits,
) -> Result<(usize, usize)> {
    let Some((original, original_canonical)) = object.source_headers() else {
        let canonical_length = canonical_archive_info_length_for_object(object)?;
        return Ok((canonical_length, canonical_length));
    };

    let mut probe = CanonicalHeaderProbe::new(Some(original_canonical));
    emit_canonical_archive_info(&mut probe, object)?;
    let (emitted_length, canonical_matches) = probe.finish();
    if canonical_matches {
        admit_retained_header(
            original,
            limits,
            object
                .source_provenance
                .as_deref()
                .and_then(|provenance| provenance.header_admission),
        )?;
        Ok((emitted_length, original.len()))
    } else {
        Ok((emitted_length, emitted_length))
    }
}

fn admit_retained_header(
    original: &[u8],
    limits: Limits,
    admission: Option<HeaderAdmission>,
) -> Result<()> {
    check_header_length(original.len(), limits)?;
    if let Some(admission) = admission {
        if admission.check(original.len(), limits).is_ok() {
            return Ok(());
        }
        // The cached totals prove a refusal but not its observation: fresh
        // ingress stops at the first charge that exceeds a budget and reports
        // that count. Replay the same bounded preflight (cold path) so a
        // retained header refuses with exactly the error its parse would.
        return refuse_retained_header(original, admission, limits);
    }
    // This private buffer was parsed or rewritten through the header codec.
    // Every field needs at least a tag and a value/length byte, and the known
    // schema descends at most ArchiveInfo -> MessageInfo -> FieldInfo ->
    // FieldPath. Unknown length-delimited fields are opaque to that codec.
    let fields_bound = original.len() / 2;
    // Bound every charge in field_memory_charge by its disjoint wire bytes:
    // message allocations use their tag/length bytes, while packed scalars
    // and opaque fields use their payload bytes. The deliberately generous
    // coefficient avoids a second wire walk for ordinary retained headers.
    // Tight profiles fall back to the exact preflight rather than refusing
    // on this upper bound or reporting a fabricated observed count.
    const MEMORY_PER_BYTE: usize = 2
        * (size_of::<MessageInfo>()
            + size_of::<archive_codec::FieldInfo>()
            + size_of::<FieldInfo>()
            + size_of::<archive_codec::FieldPath>()
            + size_of::<FieldPath>()
            + size_of::<&[u8]>()
            + size_of::<[usize; 6]>()
            + size_of::<u64>())
        + 1;
    if fields_bound <= limits.max_header_fields()
        && limits.max_header_nesting() >= 3
        && fields_bound <= limits.max_messages_per_object()
        && original
            .len()
            .checked_mul(2)
            .is_some_and(|items| items <= limits.max_metadata_items())
        && original
            .len()
            .checked_mul(MEMORY_PER_BYTE)
            .is_some_and(|bytes| bytes <= limits.max_header_memory_bytes())
    {
        return Ok(());
    }
    preflight_header(original, HeaderKind::ArchiveInfo, limits)?;
    Ok(())
}

#[cold]
#[inline(never)]
fn refuse_retained_header(
    original: &[u8],
    admission: HeaderAdmission,
    limits: Limits,
) -> Result<()> {
    preflight_header_with_admission(original, HeaderKind::ArchiveInfo, limits)?;
    // Unreachable for a consistent cache; keep the cached refusal as the
    // fail-closed answer rather than admitting the header.
    admission.check(original.len(), limits)
}

fn emit_canonical_archive_info(
    probe: &mut CanonicalHeaderProbe<'_>,
    object: &ArchiveObject,
) -> Result<()> {
    let info = &object.archive_info;
    if let Some(identifier) = info.identifier {
        probe.varint_field(1, identifier)?;
    }
    for (message_info, message) in info.message_infos.iter().zip(&object.messages) {
        let payload_length = if probe.is_comparing() {
            None
        } else {
            Some(encoded_message_info_length(
                message_info,
                message.type_,
                message.data.len(),
            )?)
        };
        probe.length_delimited_field(2, payload_length, |probe| {
            emit_canonical_message_info(probe, message_info, message.type_, message.data.len())
        })?;
    }
    if let Some(should_merge) = info.should_merge {
        probe.varint_field(3, u64::from(should_merge))?;
    }
    Ok(())
}

fn emit_canonical_message_info(
    probe: &mut CanonicalHeaderProbe<'_>,
    info: &MessageInfo,
    type_: u32,
    length: usize,
) -> Result<()> {
    probe.varint_field(1, u64::from(type_))?;
    emit_packed_u32(probe, 2, &info.versions)?;
    probe.varint_field(
        3,
        u64::from(
            u32::try_from(length)
                .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?,
        ),
    )?;
    for field_info in &info.field_infos {
        let payload_length = if probe.is_comparing() {
            None
        } else {
            Some(encoded_field_info_length(field_info)?)
        };
        probe.length_delimited_field(4, payload_length, |probe| {
            emit_canonical_field_info(probe, field_info)
        })?;
    }
    emit_packed_u64(probe, 5, &info.object_references)?;
    emit_packed_u64(probe, 6, &info.data_references)?;
    if let Some(base_message_index) = info.base_message_index {
        probe.varint_field(7, u64::from(base_message_index))?;
    }
    emit_packed_u32(probe, 8, &info.diff_merge_version)?;
    if let Some(path) = &info.diff_field_path {
        let payload_length = if probe.is_comparing() {
            None
        } else {
            Some(encoded_field_path_length(path)?)
        };
        probe.length_delimited_field(9, payload_length, |probe| {
            emit_canonical_field_path(probe, path)
        })?;
    }
    for path in &info.fields_to_remove {
        let payload_length = if probe.is_comparing() {
            None
        } else {
            Some(encoded_field_path_length(path)?)
        };
        probe.length_delimited_field(10, payload_length, |probe| {
            emit_canonical_field_path(probe, path)
        })?;
    }
    emit_packed_u32(probe, 11, &info.diff_read_version)
}

fn emit_canonical_field_info(probe: &mut CanonicalHeaderProbe<'_>, info: &FieldInfo) -> Result<()> {
    let payload_length = if probe.is_comparing() {
        None
    } else {
        Some(encoded_field_path_length(&info.path)?)
    };
    probe.length_delimited_field(1, payload_length, |probe| {
        emit_canonical_field_path(probe, &info.path)
    })?;
    if let Some(value) = info.r#type {
        emit_enum_or_unknown(
            probe,
            2,
            value.raw_value(),
            matches!(value, FieldType::Unrecognized(_)),
        )?;
    }
    if let Some(value) = info.unknown_field_rule {
        emit_enum_or_unknown(
            probe,
            3,
            value.raw_value(),
            matches!(value, UnknownFieldRule::Unrecognized(_)),
        )?;
    }
    emit_packed_u64(probe, 4, &info.object_references)?;
    emit_packed_u64(probe, 5, &info.data_references)?;
    if let Some(value) = info.known_field_rule {
        emit_enum_or_unknown(
            probe,
            6,
            value.raw_value(),
            matches!(value, KnownFieldRule::Unrecognized(_)),
        )?;
    }
    emit_packed_u32(probe, 7, &info.known_field_version)?;
    if let Some(feature_identifier) = &info.known_field_feature_identifier {
        probe.length_delimited_field(8, Some(feature_identifier.len()), |probe| {
            probe.bytes(feature_identifier.as_bytes())
        })?;
    }
    if let Some(value) = info.r#type {
        if matches!(value, FieldType::Unrecognized(_)) {
            probe.varint_field(2, enum_wire_value(value.raw_value()))?;
        }
    }
    if let Some(value) = info.unknown_field_rule {
        if matches!(value, UnknownFieldRule::Unrecognized(_)) {
            probe.varint_field(3, enum_wire_value(value.raw_value()))?;
        }
    }
    if let Some(value) = info.known_field_rule {
        if matches!(value, KnownFieldRule::Unrecognized(_)) {
            probe.varint_field(6, enum_wire_value(value.raw_value()))?;
        }
    }
    Ok(())
}

fn emit_canonical_field_path(probe: &mut CanonicalHeaderProbe<'_>, path: &FieldPath) -> Result<()> {
    emit_packed_u32(probe, 1, &path.path)
}

fn emit_enum_or_unknown(
    probe: &mut CanonicalHeaderProbe<'_>,
    number: u32,
    value: i32,
    unknown: bool,
) -> Result<()> {
    if unknown {
        return Ok(());
    }
    probe.varint_field(number, enum_wire_value(value))
}

fn emit_packed_u32(
    probe: &mut CanonicalHeaderProbe<'_>,
    number: u32,
    values: &[u32],
) -> Result<()> {
    if values.is_empty() {
        return Ok(());
    }
    let payload_length = if probe.is_comparing() {
        None
    } else {
        Some(packed_u32_length(values)?)
    };
    probe.length_delimited_field(number, payload_length, |probe| {
        for value in values {
            probe.varint(u64::from(*value))?;
        }
        Ok(())
    })
}

fn emit_packed_u64(
    probe: &mut CanonicalHeaderProbe<'_>,
    number: u32,
    values: &[u64],
) -> Result<()> {
    if values.is_empty() {
        return Ok(());
    }
    let payload_length = if probe.is_comparing() {
        None
    } else {
        Some(packed_u64_length(values)?)
    };
    probe.length_delimited_field(number, payload_length, |probe| {
        for value in values {
            probe.varint(*value)?;
        }
        Ok(())
    })
}

fn encoded_message_info_length(info: &MessageInfo, type_: u32, length: usize) -> Result<usize> {
    let mut total = varint_field_length(1, u64::from(type_))?;
    total = checked_encoded_add(total, packed_u32_field_length(2, &info.versions)?)?;
    let length = u32::try_from(length)
        .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
    total = checked_encoded_add(total, varint_field_length(3, u64::from(length))?)?;
    for field_info in &info.field_infos {
        let payload_length = encoded_field_info_length(field_info)?;
        total = checked_encoded_add(total, length_delimited_field_length(4, payload_length)?)?;
    }
    total = checked_encoded_add(total, packed_u64_field_length(5, &info.object_references)?)?;
    total = checked_encoded_add(total, packed_u64_field_length(6, &info.data_references)?)?;
    if let Some(value) = info.base_message_index {
        total = checked_encoded_add(total, varint_field_length(7, u64::from(value))?)?;
    }
    total = checked_encoded_add(total, packed_u32_field_length(8, &info.diff_merge_version)?)?;
    if let Some(path) = &info.diff_field_path {
        total = checked_encoded_add(
            total,
            length_delimited_field_length(9, encoded_field_path_length(path)?)?,
        )?;
    }
    for path in &info.fields_to_remove {
        total = checked_encoded_add(
            total,
            length_delimited_field_length(10, encoded_field_path_length(path)?)?,
        )?;
    }
    checked_encoded_add(total, packed_u32_field_length(11, &info.diff_read_version)?)
}

fn encoded_field_info_length(info: &FieldInfo) -> Result<usize> {
    let mut total = length_delimited_field_length(1, encoded_field_path_length(&info.path)?)?;
    if let Some(value) = info.r#type {
        if !matches!(value, FieldType::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(2, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    if let Some(value) = info.unknown_field_rule {
        if !matches!(value, UnknownFieldRule::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(3, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    total = checked_encoded_add(total, packed_u64_field_length(4, &info.object_references)?)?;
    total = checked_encoded_add(total, packed_u64_field_length(5, &info.data_references)?)?;
    if let Some(value) = info.known_field_rule {
        if !matches!(value, KnownFieldRule::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(6, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    total = checked_encoded_add(
        total,
        packed_u32_field_length(7, &info.known_field_version)?,
    )?;
    if let Some(feature_identifier) = &info.known_field_feature_identifier {
        total = checked_encoded_add(
            total,
            length_delimited_field_length(8, feature_identifier.len())?,
        )?;
    }
    if let Some(value) = info.r#type {
        if matches!(value, FieldType::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(2, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    if let Some(value) = info.unknown_field_rule {
        if matches!(value, UnknownFieldRule::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(3, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    if let Some(value) = info.known_field_rule {
        if matches!(value, KnownFieldRule::Unrecognized(_)) {
            total = checked_encoded_add(
                total,
                varint_field_length(6, enum_wire_value(value.raw_value()))?,
            )?;
        }
    }
    Ok(total)
}

fn encoded_field_path_length(path: &FieldPath) -> Result<usize> {
    packed_u32_field_length(1, &path.path)
}

fn packed_u32_field_length(number: u32, values: &[u32]) -> Result<usize> {
    if values.is_empty() {
        return Ok(0);
    }
    length_delimited_field_length(number, packed_u32_length(values)?)
}

fn packed_u64_field_length(number: u32, values: &[u64]) -> Result<usize> {
    if values.is_empty() {
        return Ok(0);
    }
    length_delimited_field_length(number, packed_u64_length(values)?)
}

fn packed_u32_length(values: &[u32]) -> Result<usize> {
    values.iter().try_fold(0usize, |total, value| {
        checked_encoded_add(total, u64_varint_length(u64::from(*value)))
    })
}

fn packed_u64_length(values: &[u64]) -> Result<usize> {
    values.iter().try_fold(0usize, |total, value| {
        checked_encoded_add(total, u64_varint_length(*value))
    })
}

fn varint_field_length(number: u32, value: u64) -> Result<usize> {
    checked_encoded_add(
        u64_varint_length(u64::from(number) << 3),
        u64_varint_length(value),
    )
}

fn length_delimited_field_length(number: u32, payload_length: usize) -> Result<usize> {
    let payload_length = u64::try_from(payload_length)
        .map_err(|_| Error::invalid_archive(0, "ArchiveInfo field length exceeds u64"))?;
    let total = checked_encoded_add(
        u64_varint_length((u64::from(number) << 3) | 2),
        u64_varint_length(payload_length),
    )?;
    checked_encoded_add(
        total,
        usize::try_from(payload_length)
            .map_err(|_| Error::invalid_archive(0, "ArchiveInfo field length exceeds usize"))?,
    )
}

fn checked_encoded_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| Error::invalid_archive(0, "ArchiveInfo encoded length overflow"))
}

const fn u64_varint_length(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

const fn enum_wire_value(value: i32) -> u64 {
    u64::from_ne_bytes((value as i64).to_ne_bytes())
}

fn buffa_decode_options(preflight: WirePreflight, limits: Limits) -> archive_codec::DecodeOptions {
    archive_codec::DecodeOptions::new(
        limits.max_header_bytes(),
        preflight.fields(),
        limits.max_header_memory_bytes(),
        u32::try_from(limits.max_header_nesting()).unwrap_or(u32::MAX),
    )
}

fn preflight_header(data: &[u8], header: HeaderKind, limits: Limits) -> Result<WirePreflight> {
    preflight_header_with_admission(data, header, limits).map(|admission| admission.preflight)
}

fn preflight_header_with_admission(
    data: &[u8],
    header: HeaderKind,
    limits: Limits,
) -> Result<HeaderAdmission> {
    let root = match header {
        HeaderKind::ArchiveInfo => HeaderNode::ArchiveInfo,
        HeaderKind::MessageInfo => HeaderNode::MessageInfo,
    };
    let scanned_bytes = limits
        .max_header_bytes()
        .saturating_mul(4)
        .min(WireLimits::MAX_INPUT_BYTES);
    let wire_limits = WireLimits::default()
        .with_input_bytes(scanned_bytes)
        .and_then(|profile| profile.with_fields(limits.max_header_fields()))
        .and_then(|profile| profile.with_nesting(limits.max_header_nesting()))
        .map_err(|error| map_wire_error(error, header))?;
    let decoded_memory = Cell::new(0usize);
    let metadata_items = Cell::new(0usize);
    let message_infos = Cell::new(0usize);

    let preflight = preflight_wire_tree(data, wire_limits, |visit| {
        let node = node_at_path(root, visit.path());
        if node == Some(HeaderNode::ArchiveInfo)
            && visit.field().number() == 2
            && visit.field().wire_type() == 2
        {
            message_infos.set(message_infos.get().saturating_add(1));
        }
        decoded_memory.set(
            decoded_memory
                .get()
                .saturating_add(field_memory_charge(node, visit)),
        );
        metadata_items.set(
            metadata_items
                .get()
                .saturating_add(metadata_item_charge(node, visit)?),
        );
        Ok(descent_for(node, visit.field().number()))
    })
    .map_err(|error| map_wire_error(error, header))?;

    if message_infos.get() > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            message_infos.get(),
            limits.max_messages_per_object(),
        ));
    }
    if decoded_memory.get() > limits.max_header_memory_bytes() {
        return Err(limit(
            LimitKind::HeaderMemoryBytes,
            decoded_memory.get(),
            limits.max_header_memory_bytes(),
        ));
    }
    if metadata_items.get() > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            metadata_items.get(),
            limits.max_metadata_items(),
        ));
    }
    Ok(HeaderAdmission {
        preflight,
        message_infos: message_infos.get(),
        decoded_memory: decoded_memory.get(),
        metadata_items: metadata_items.get(),
    })
}

fn header_has_unknown_reference_metadata(data: &[u8], limits: Limits) -> Result<bool> {
    let wire_limits = header_wire_limits(limits)?;
    let mut unknown = false;
    preflight_wire_tree(data, wire_limits, |visit| {
        let node = node_at_path(HeaderNode::ArchiveInfo, visit.path());
        let field = visit.field();
        unknown |= !known_header_field(node, field.number());
        Ok(descent_for(node, field.number()))
    })
    .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    Ok(unknown)
}

const fn known_header_field(node: Option<HeaderNode>, field: u32) -> bool {
    match node {
        Some(HeaderNode::ArchiveInfo) => matches!(field, 1..=3),
        Some(HeaderNode::MessageInfo) => matches!(field, 1..=11),
        Some(HeaderNode::FieldInfo) => matches!(field, 1..=8),
        Some(HeaderNode::FieldPath) => field == 1,
        None => false,
    }
}

fn node_at_path(root: HeaderNode, path: &[u32]) -> Option<HeaderNode> {
    path.iter()
        .try_fold(root, |node, field| match (node, field) {
            (HeaderNode::ArchiveInfo, 2) => Some(HeaderNode::MessageInfo),
            (HeaderNode::MessageInfo, 4) => Some(HeaderNode::FieldInfo),
            (HeaderNode::MessageInfo, 9 | 10) | (HeaderNode::FieldInfo, 1) => {
                Some(HeaderNode::FieldPath)
            },
            _ => None,
        })
}

const fn descent_for(node: Option<HeaderNode>, field: u32) -> WireDescent {
    if matches!(
        (node, field),
        (Some(HeaderNode::ArchiveInfo), 2)
            | (Some(HeaderNode::MessageInfo), 4 | 9 | 10)
            | (Some(HeaderNode::FieldInfo), 1)
    ) {
        WireDescent::Descend
    } else {
        WireDescent::Skip
    }
}

fn field_memory_charge(node: Option<HeaderNode>, visit: WireVisit<'_, '_>) -> usize {
    let field = visit.field();
    let number = field.number();
    let wire_type = field.wire_type();
    let scalar = |width: usize| match wire_type {
        0 => width.saturating_mul(2),
        2 => field
            .payload()
            .len()
            .saturating_mul(width)
            .saturating_mul(2),
        _ => 0,
    };
    let message = |width: usize| {
        if wire_type == 2 {
            width.saturating_mul(2).saturating_add(size_of::<&[u8]>())
        } else {
            0
        }
    };
    let projected_message = |compatibility_width: usize, neutral_width: usize| {
        if wire_type == 2 {
            message(compatibility_width).saturating_add(neutral_width)
        } else {
            0
        }
    };

    match (node, number) {
        (Some(HeaderNode::ArchiveInfo), 2) => message(size_of::<MessageInfo>()),
        (Some(HeaderNode::MessageInfo), 2 | 8 | 11) => scalar(size_of::<u32>()),
        (Some(HeaderNode::MessageInfo), 5 | 6) | (Some(HeaderNode::FieldInfo), 4 | 5) => {
            scalar(size_of::<u64>())
        },
        (Some(HeaderNode::MessageInfo), 4) => projected_message(
            size_of::<archive_codec::FieldInfo>(),
            size_of::<FieldInfo>(),
        ),
        (Some(HeaderNode::MessageInfo), 9) | (Some(HeaderNode::FieldInfo), 1) => {
            message(size_of::<archive_codec::FieldPath>())
        },
        (Some(HeaderNode::MessageInfo), 10) => projected_message(
            size_of::<archive_codec::FieldPath>(),
            size_of::<FieldPath>(),
        ),
        (Some(HeaderNode::FieldInfo), 7) | (Some(HeaderNode::FieldPath), 1) => {
            scalar(size_of::<u32>())
        },
        (Some(HeaderNode::FieldInfo), 8) if wire_type == 2 => field.payload().len(),
        (Some(HeaderNode::FieldInfo), 2 | 3 | 6) => field_info_enum_memory_charge(number, field),
        (Some(HeaderNode::ArchiveInfo), 1 | 3) | (Some(HeaderNode::MessageInfo), 1 | 3 | 7) => 0,
        _ => unknown_field_memory_charge(field),
    }
}

fn field_info_enum_memory_charge(
    number: u32,
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> usize {
    let value = (field.wire_type() == 0)
        .then(|| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
        .flatten()
        .map(|(value, _encoded_length)| value as i32);
    let known = matches!(
        (number, value),
        (2, Some(0..=3)) | (3, Some(-1..=2)) | (6, Some(0..=2))
    );
    if known {
        0
    } else {
        unknown_field_memory_charge(field)
    }
}

fn unknown_field_memory_charge(field: litchi_iwa_common::wire::WireFieldView<'_>) -> usize {
    field
        .raw()
        .len()
        .saturating_add(2 * size_of::<[usize; 6]>())
}

fn metadata_item_charge(
    node: Option<HeaderNode>,
    visit: WireVisit<'_, '_>,
) -> litchi_iwa_common::Result<usize> {
    let field = visit.field();
    let repeated = || repeated_varint_count(field);
    match (node, field.number()) {
        // ArchiveInfo validation includes one item for each MessageInfo.
        (Some(HeaderNode::MessageInfo), 2 | 5 | 6 | 8 | 11)
        | (Some(HeaderNode::FieldInfo), 4 | 5 | 7)
        | (Some(HeaderNode::FieldPath), 1) => repeated(),
        // A FieldInfo is counted once in the repeated collection and once
        // again when its nested metadata is traversed by validation.
        (Some(HeaderNode::MessageInfo), 4) if field.wire_type() == 2 => Ok(2),
        (Some(HeaderNode::ArchiveInfo), 2)
        | (Some(HeaderNode::MessageInfo), 9 | 10)
        | (Some(HeaderNode::FieldInfo), 8)
            if field.wire_type() == 2 =>
        {
            Ok(1)
        },
        _ => Ok(0),
    }
}

fn repeated_varint_count(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> litchi_iwa_common::Result<usize> {
    match field.wire_type() {
        0 => Ok(1),
        2 => {
            let mut remaining = field.payload();
            let mut count = 0usize;
            while !remaining.is_empty() {
                let (_, length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
                    .map_err(|error| WireError::InvalidFormat(error.to_string()))?;
                remaining = &remaining[length..];
                count = count.saturating_add(1);
            }
            Ok(count)
        },
        _ => Ok(0),
    }
}

fn map_wire_error(error: WireError, header: HeaderKind) -> Error {
    match error {
        WireError::LimitExceeded {
            kind: WireLimitKind::Fields,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderFields, observed, maximum),
        WireError::LimitExceeded {
            kind: WireLimitKind::Nesting,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderNesting, observed, maximum),
        WireError::LimitExceeded {
            kind: WireLimitKind::InputBytes,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderBytes, observed, maximum),
        WireError::Allocation { resource, amount } => Error::allocation(resource, amount),
        other @ (WireError::InvalidFormat(_)
        | WireError::LimitExceeded { .. }
        | WireError::InvalidLimit { .. }) => {
            Error::header_codec(header, HeaderOperation::Decode, other.to_string())
        },
        _ => Error::header_codec(
            header,
            HeaderOperation::Decode,
            "shared wire decoder rejected header",
        ),
    }
}

fn read_bounded<R: Read>(
    reader: &mut R,
    maximum: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    let mut buffer = [0u8; 8192];
    loop {
        let remaining = maximum.saturating_sub(data.len());
        let read_size = remaining.min(buffer.len()).max(1);
        let read = reader.read(&mut buffer[..read_size])?;
        if read == 0 {
            return Ok(data);
        }
        if data.len() == maximum {
            return Err(limit(
                LimitKind::HeaderBytes,
                maximum.saturating_add(read),
                maximum,
            ));
        }
        data.try_reserve(read)
            .map_err(|_| Error::allocation(resource, read))?;
        data.extend_from_slice(&buffer[..read]);
    }
}

fn validate_object_set<'a, I>(objects: I, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = &'a ArchiveObject>,
{
    let objects = objects.into_iter();
    let mut identifiers = HashSet::new();
    let initial_capacity = objects.size_hint().0.min(limits.max_objects());
    identifiers
        .try_reserve(initial_capacity)
        .map_err(|_| Error::allocation("IWA archive identifiers", initial_capacity))?;
    let mut object_count = 0usize;
    let mut total_messages = 0usize;
    let mut total_payload = 0usize;
    for object in objects {
        object_count = object_count
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "object count overflow"))?;
        if object_count > limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                object_count,
                limits.max_objects(),
            ));
        }
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if !identifiers.insert(identifier) {
            return Err(Error::invalid_archive(0, "duplicate object identifier"));
        }
        object.validate(limits)?;
        add_limited(
            &mut total_messages,
            object.messages.len(),
            LimitKind::Messages,
            limits.max_messages(),
        )?;
        for message in &object.messages {
            add_limited(
                &mut total_payload,
                message.data.len(),
                LimitKind::ArchiveBytes,
                limits.max_archive_bytes(),
            )?;
        }
    }
    Ok(())
}

fn validate_archive_wire_size<'a, I>(objects: I, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = &'a ArchiveObject>,
{
    let mut archive_length = 0usize;
    for object in objects {
        let canonical_header = encode_archive_info(&object.archive_info, limits)?;
        let header = match object.source_headers() {
            Some((original, canonical)) if canonical == canonical_header.as_slice() => {
                admit_retained_header(
                    original,
                    limits,
                    object
                        .source_provenance
                        .as_deref()
                        .and_then(|provenance| provenance.header_admission),
                )?;
                original
            },
            _ => canonical_header.as_slice(),
        };
        check_header_length(header.len(), limits)?;
        let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
            total
                .checked_add(message.data.len())
                .ok_or_else(|| Error::invalid_archive(0, "payload length overflow"))
        })?;
        let object_length = varint_len(header.len())?
            .checked_add(header.len())
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        add_limited(
            &mut archive_length,
            object_length,
            LimitKind::ArchiveBytes,
            limits.max_archive_bytes(),
        )?;
    }
    Ok(())
}

fn check_message_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_message_bytes() {
        return Err(limit(
            LimitKind::MessageBytes,
            length,
            limits.max_message_bytes(),
        ));
    }
    Ok(())
}

fn add_count(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "metadata count overflow"))?;
    Ok(())
}

fn add_limited(total: &mut usize, amount: usize, kind: LimitKind, maximum: usize) -> Result<()> {
    let observed = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "resource count overflow"))?;
    if observed > maximum {
        return Err(limit(kind, observed, maximum));
    }
    *total = observed;
    Ok(())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

fn decode_varint(data: &[u8]) -> Result<(u64, usize)> {
    let mut value = 0u64;
    for (index, byte) in data.iter().copied().take(MAX_VARINT_BYTES).enumerate() {
        if index == MAX_VARINT_BYTES - 1 && byte > 1 {
            return Err(Error::invalid_archive(
                index,
                "archive length varint overflow",
            ));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, index + 1));
        }
    }
    if data.len() < MAX_VARINT_BYTES {
        Err(Error::invalid_archive(
            data.len(),
            "truncated archive length varint",
        ))
    } else {
        Err(Error::invalid_archive(
            MAX_VARINT_BYTES - 1,
            "archive length varint overflow",
        ))
    }
}

fn varint_len(value: usize) -> Result<usize> {
    let value =
        u64::try_from(value).map_err(|_| Error::invalid_archive(0, "length exceeds u64"))?;
    Ok(if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    })
}

fn encode_varint(mut value: u64, output: &mut [u8; MAX_VARINT_BYTES]) -> &[u8] {
    let mut length = 0;
    while value >= 0x80 {
        output[length] = u8::try_from(value & 0x7f).unwrap_or_default() | 0x80;
        value >>= 7;
        length += 1;
    }
    output[length] = u8::try_from(value).unwrap_or_default();
    &output[..=length]
}

#[cfg(test)]
mod tests {
    use super::{
        Archive, ArchiveEncodingPreparationChunk, ArchiveEncodingPreparationError, ArchiveInfo,
        ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
        ArchiveReferenceScope, ArchiveReferenceVisitor, CanonicalFieldDataReferenceOperation,
        CanonicalFieldDataReferenceTransition, CanonicalObjectReferenceField,
        CanonicalObjectReferenceFields, DEFAULT_MESSAGE_VERSIONS, DataReferencePruning,
        DataReferenceTransition, Error, FieldDataReferenceTransition, FieldInfo,
        FieldObjectReferenceTransition, FieldPath, FieldType, KnownFieldRule,
        MESSAGE_INFO_CONSTRUCTIONS, MessageInfo, MessageInfoCollection, ObjectReferenceTransition,
        RawMessage, UnknownFieldRule, archive_info_encoded_len, archive_info_to_codec,
        check_archive_info_constructor, default_archive_info_bounds_are_safe, encode_archive_info,
        encode_varint, encode_varint_with_width, encoded_field_info_length,
        encoded_field_path_length, encoded_message_info_length, field_info_header_field_count,
        field_info_memory, field_path_memory, message_info_header_field_count,
        message_info_header_nesting, message_info_memory, minimal_message_info_encoded_length,
        varint_len,
    };
    use crate::{LimitKind, Limits, Result};
    use litchi_iwa_protos::archive_codec;

    struct ReferencePruningFixture {
        source: Vec<u8>,
        header: Vec<u8>,
        expected_header: Vec<u8>,
    }

    struct CombinedReferencePruningFixture {
        source: Vec<u8>,
        selected_header: Vec<u8>,
        all_data_header: Vec<u8>,
    }

    struct ReferenceReorderFixture {
        source: Vec<u8>,
        header: Vec<u8>,
        reordered_header: Vec<u8>,
    }

    struct ReferenceTransitionFixture {
        source: Vec<u8>,
        transitioned_header: Vec<u8>,
    }

    fn prepared_archive(
        archive: &Archive,
        limits: Limits,
    ) -> Result<super::ArchiveEncodingPlan<'_>> {
        let mut visitor = |_chunk: ArchiveEncodingPreparationChunk| Ok::<(), Error>(());
        archive
            .prepare_encoding(limits, &mut visitor)
            .map_err(|error| match error {
                ArchiveEncodingPreparationError::Core(error)
                | ArchiveEncodingPreparationError::Admission(error) => error,
                ArchiveEncodingPreparationError::ArithmeticOverflow { resource } => {
                    Error::invalid_archive(0, resource)
                },
            })
    }

    #[derive(Default)]
    struct ReferenceFacts(Vec<ArchiveReferenceOccurrence>);

    impl ArchiveReferenceVisitor for ReferenceFacts {
        fn visit_reference(&mut self, occurrence: ArchiveReferenceOccurrence) -> Result<()> {
            self.0.push(occurrence);
            Ok(())
        }
    }

    #[test]
    fn reference_inspection_streams_aggregate_and_field_occurrences() -> Result<()> {
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: vec![1],
            }],
            Default::default(),
        )?;
        let message = &mut object.archive_info.message_infos[0];
        message.object_references = vec![10, 10];
        message.data_references = vec![20];
        let mut field = FieldInfo::new(
            FieldPath::new(vec![1, 2], Limits::default()).unwrap(),
            Limits::default(),
        )
        .unwrap();
        field.object_references = vec![11];
        field.data_references = vec![21, 21];
        message.field_infos.push(field);

        let mut facts = ReferenceFacts::default();
        assert_eq!(
            object.inspect_references(&mut facts, Default::default())?,
            6
        );
        assert_eq!(facts.0.len(), 6);
        assert_eq!(facts.0[0].object_identifier, 42);
        assert_eq!(facts.0[0].scope, ArchiveReferenceScope::Message);
        assert_eq!(facts.0[0].kind, ArchiveReferenceKind::Object);
        assert_eq!(facts.0[0].referenced_identifier, 10);
        assert_eq!(
            facts.0[3].scope,
            ArchiveReferenceScope::Field { field_index: 0 }
        );
        assert_eq!(facts.0[3].referenced_identifier, 11);
        assert_eq!(facts.0[4].kind, ArchiveReferenceKind::Data);
        Ok(())
    }

    #[test]
    fn reference_inspection_enforces_metadata_limit_before_late_callbacks() -> Result<()> {
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: vec![1],
            }],
            Default::default(),
        )?;
        object.archive_info.message_infos[0].object_references = vec![10, 11, 12];
        let limits = Limits::default().with_metadata_items(3)?;
        let mut facts = ReferenceFacts::default();
        let error = object.inspect_references(&mut facts, limits).unwrap_err();
        assert!(matches!(error, Error::Limit { .. }));
        assert!(facts.0.is_empty());
        Ok(())
    }

    #[test]
    fn checked_object_removal_is_exact_and_atomic() -> Result<()> {
        let mut archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 7,
                        data: vec![1],
                    }],
                    Default::default(),
                )?,
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: 8,
                        data: vec![2],
                    }],
                    Default::default(),
                )?,
            ],
        };
        let source = archive.clone();
        assert!(
            archive
                .remove_object_checked(0, Default::default())
                .is_err()
        );
        assert_eq!(archive, source);
        assert!(
            archive
                .remove_object_checked(3, Default::default())
                .is_err()
        );
        assert_eq!(archive, source);

        let removed = archive.remove_object_checked(1, Default::default())?;
        assert_eq!(removed.archive_info.identifier, Some(1));
        assert_eq!(archive.objects.len(), 1);
        assert_eq!(archive.objects[0].archive_info.identifier, Some(2));
        Ok(())
    }

    #[test]
    fn deletion_grade_reference_inspection_rejects_unknown_header_metadata() -> Result<()> {
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: vec![1],
            }],
            Default::default(),
        )?;
        object.archive_info.message_infos[0].object_references = vec![10];
        let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
        let mut raw = canonical.clone();
        raw.extend_from_slice(&[0xb8, 0x3e, 0x01]);
        object.set_test_source_headers(
            Some(raw.into_boxed_slice()),
            Some(canonical.into_boxed_slice()),
        );

        let mut permissive = ReferenceFacts::default();
        assert_eq!(
            object.inspect_references(&mut permissive, Default::default())?,
            1
        );
        let mut strict = ReferenceFacts::default();
        let error = object
            .inspect_references_with_policy(
                &mut strict,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                Limits::default(),
            )
            .unwrap_err();
        assert!(matches!(error, Error::InvalidArchive { .. }));
        assert!(strict.0.is_empty());
        Ok(())
    }

    #[test]
    fn canonical_headers_use_compact_retention_path() -> Result<()> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![3, 4],
                }],
                Default::default(),
            )?],
        };
        let encoded = archive.to_bytes(Default::default())?;
        let parsed = Archive::parse(&encoded, Default::default())?;
        let [object] = parsed.objects.as_slice() else {
            return Err(Error::invalid_archive(
                0,
                "test archive did not contain exactly one object",
            ));
        };

        assert!(object.source_header().is_none());
        assert!(object.source_canonical_header().is_none());
        assert_eq!(parsed.to_bytes(Default::default())?, encoded);
        Ok(())
    }

    #[test]
    fn encoded_len_matches_serialization_for_full_metadata() -> Result<()> {
        let mut object = ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: 0x1234,
                data: vec![3, 4, 5, 6],
            }],
            Default::default(),
        )?;
        object.archive_info.should_merge = Some(false);
        let message = &mut object.archive_info.message_infos[0];
        message.versions = vec![0, 127, 128, u32::MAX].into();
        message.object_references = vec![1, 128, u64::MAX];
        message.data_references = vec![2, 129];
        message.base_message_index = Some(u32::MAX);
        message.diff_merge_version = vec![3, 130];
        message.diff_field_path =
            Some(FieldPath::new(vec![1, 128, u32::MAX], Limits::default()).unwrap());
        message.fields_to_remove = vec![
            FieldPath::new(vec![4, 5], Limits::default()).unwrap(),
            FieldPath::default(),
        ];
        message.diff_read_version = vec![6, 7];

        let mut field = FieldInfo::new(
            FieldPath::new(vec![9, 10], Limits::default()).unwrap(),
            Limits::default(),
        )
        .unwrap();
        field.r#type = Some(FieldType::Unrecognized(99));
        field.unknown_field_rule = Some(UnknownFieldRule::NotSupported);
        field.object_references = vec![11, 12];
        field.data_references = vec![13];
        field.known_field_rule = Some(KnownFieldRule::Unrecognized(-8));
        field.known_field_version = vec![14, 15];
        field.known_field_feature_identifier = Some("feature".to_owned());
        message.field_infos.push(field);

        let archive = Archive {
            objects: vec![object],
        };
        let encoded = archive.to_bytes(Default::default())?;
        assert_eq!(
            prepared_archive(&archive, Default::default())?.execute()?,
            encoded
        );
        assert_eq!(archive.encoded_len(Default::default())?, encoded.len());
        assert_eq!(archive.encoded_len(Limits::default())?, encoded.len());
        let encoded_info = archive_info_to_codec(&archive.objects[0].archive_info)?;
        let codec_length = archive_codec::archive_info_encoded_len(&encoded_info)
            .map_err(|_| Error::invalid_archive(0, "codec length parity failed"))
            .and_then(|length| {
                usize::try_from(length)
                    .map_err(|_| Error::invalid_archive(0, "codec length exceeds usize"))
            })?;
        assert_eq!(
            archive_info_encoded_len(&archive.objects[0].archive_info)?,
            codec_length
        );
        Ok(())
    }

    #[test]
    fn prepared_encoding_matches_legacy_for_large_archive() -> Result<()> {
        let mut objects = Vec::new();
        for identifier in 1..=33u64 {
            objects.push(ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: identifier as u32,
                    data: vec![identifier as u8, 0x7f],
                }],
                Limits::default(),
            )?);
        }
        let archive = Archive { objects };
        let legacy = archive.to_bytes(Limits::default())?;
        let plan = prepared_archive(&archive, Limits::default())?;
        let requirements = plan.requirements();
        assert_eq!(requirements.object_count(), 33);
        assert_eq!(requirements.message_count(), 33);
        assert_eq!(requirements.identifier_index_entries(), 33);
        assert_eq!(requirements.encoded_bytes(), legacy.len());
        assert!(requirements.scratch_bytes() > 33 * size_of::<u64>());
        assert_eq!(plan.execute()?, legacy);
        Ok(())
    }

    #[test]
    fn prepared_encoding_preserves_unknown_retained_headers() -> Result<()> {
        let mut object = ArchiveObject::new(
            901,
            vec![RawMessage {
                type_: 7,
                data: vec![1, 2, 3],
            }],
            Limits::default(),
        )?;
        let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
        let mut raw = canonical.clone();
        raw.extend_from_slice(&[0xb5, 0x0c, 0x01, 0x02, 0x03, 0x04]);
        object.set_test_source_headers(
            Some(raw.clone().into_boxed_slice()),
            Some(canonical.into_boxed_slice()),
        );
        let archive = Archive {
            objects: vec![object],
        };
        let legacy = archive.to_bytes(Limits::default())?;
        let prepared = prepared_archive(&archive, Limits::default())?.execute()?;
        assert_eq!(prepared, legacy);
        let reopened = Archive::parse(&prepared, Limits::default())?;
        assert_eq!(reopened.to_bytes(Limits::default())?, prepared);
        let (header, _) = split_test_archive(&prepared)?;
        assert!(header.starts_with(&raw));
        Ok(())
    }

    #[test]
    fn prepared_encoding_sizes_normalized_message_scalars() -> Result<()> {
        let mut object = ArchiveObject::new(
            902,
            vec![RawMessage {
                type_: 0x1234,
                data: vec![0x5a; 128],
            }],
            Limits::default(),
        )?;
        // The serializer intentionally takes these two scalars from the
        // current payload. Keep stale metadata crossing the 127/128 wire
        // boundary to prove admission sizes the normalized header without
        // cloning a compatibility projection.
        object.archive_info.message_infos[0].type_ = 1;
        object.archive_info.message_infos[0].length = 127;
        let archive = Archive {
            objects: vec![object],
        };
        let legacy = archive.to_bytes(Limits::default())?;
        let prepared = prepared_archive(&archive, Limits::default())?.execute()?;
        assert_eq!(prepared, legacy);
        assert_eq!(
            Archive::parse(&prepared, Limits::default())?.to_bytes(Limits::default())?,
            prepared
        );
        Ok(())
    }

    #[test]
    fn prepared_encoding_rejects_duplicate_and_tight_limits() -> Result<()> {
        let mut duplicate_objects = Vec::new();
        for identifier in 1..=33u64 {
            duplicate_objects.push(ArchiveObject::new(
                if identifier == 33 { 1 } else { identifier },
                vec![RawMessage {
                    type_: 1,
                    data: vec![0xaa],
                }],
                Limits::default(),
            )?);
        }
        let duplicate = Archive {
            objects: duplicate_objects,
        };
        let duplicate_plan = prepared_archive(&duplicate, Limits::default())?;
        assert!(matches!(
            duplicate_plan.execute(),
            Err(Error::InvalidArchive {
                reason: "duplicate object identifier",
                ..
            })
        ));

        let two = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 1,
                        data: vec![1],
                    }],
                    Limits::default(),
                )?,
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: 1,
                        data: vec![2],
                    }],
                    Limits::default(),
                )?,
            ],
        };
        let one_object = Limits::default().with_objects(1)?;
        assert!(matches!(
            prepared_archive(&two, one_object),
            Err(Error::Limit {
                kind: LimitKind::Objects,
                observed: 2,
                maximum: 1,
            })
        ));
        assert!(matches!(
            Limits::default().with_archive_bytes(0),
            Err(Error::InvalidLimits { .. })
        ));
        Ok(())
    }

    #[test]
    fn prepared_encoding_honors_work_admission_before_metadata() -> Result<()> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                11,
                vec![RawMessage {
                    type_: 1,
                    data: vec![1, 2],
                }],
                Limits::default(),
            )?],
        };
        let mut phases = 0usize;
        let error = archive
            .prepare_encoding(
                Limits::default(),
                &mut |chunk: ArchiveEncodingPreparationChunk| {
                    phases = phases.checked_add(1).unwrap();
                    Err(Error::allocation(
                        "test preparation ledger",
                        chunk.traversal_bytes(),
                    ))
                },
            )
            .expect_err("preparation work refusal");
        assert_eq!(phases, 1);
        assert!(matches!(
            error,
            ArchiveEncodingPreparationError::Admission(Error::Allocation { .. })
        ));

        #[derive(Debug)]
        struct CallerBudgetRefusal;
        let typed = archive.prepare_encoding(
            Limits::default(),
            &mut |_chunk: ArchiveEncodingPreparationChunk| Err(CallerBudgetRefusal),
        );
        assert!(matches!(
            typed,
            Err(ArchiveEncodingPreparationError::Admission(
                CallerBudgetRefusal
            ))
        ));
        Ok(())
    }

    #[test]
    fn constructor_path_bounds_preserve_exact_varint_boundary_admission() -> Result<()> {
        let defaults = Limits::default();
        for value in [0, 127, 128, 16_383, 16_384, u32::MAX] {
            for count in [0, 1, 25, 26, 127, 128] {
                let path = FieldPath {
                    path: vec![value; count],
                };
                let exact = encoded_field_path_length(&path)?;
                let loose = FieldPath::new(path.path.clone(), defaults)?;
                assert_eq!(loose, path);
                if exact > 0 {
                    assert_eq!(
                        FieldPath::new(path.path.clone(), defaults.with_header_bytes(exact)?)?,
                        path,
                    );
                    let error =
                        FieldPath::new(path.path.clone(), defaults.with_header_bytes(exact - 1)?)
                            .expect_err("one-under path bytes");
                    assert!(matches!(error, Error::Limit {
                        kind: LimitKind::HeaderBytes, observed, maximum,
                    } if observed == exact && maximum == exact - 1));
                }
                let field = FieldInfo::new(path.clone(), defaults)?;
                let exact = encoded_field_info_length(&field)?;
                assert_eq!(
                    FieldInfo::new(path.clone(), defaults.with_header_bytes(exact)?)?,
                    field,
                );
                let error = FieldInfo::new(path, defaults.with_header_bytes(exact - 1)?)
                    .expect_err("one-under field bytes");
                assert!(matches!(error, Error::Limit {
                    kind: LimitKind::HeaderBytes, observed, maximum,
                } if observed == exact && maximum == exact - 1));
            }
        }
        Ok(())
    }

    #[test]
    fn default_archive_info_admission_checks_public_message_length_mutations() -> Result<()> {
        let limits = Limits::default();
        let mut message = MessageInfo::new(7, 1, limits)?;
        message.length = u32::try_from(Limits::MAX_MESSAGE_BYTES + 1)
            .expect("the hard message limit fits in the wire u32");
        let info = ArchiveInfo {
            identifier: Some(0),
            message_infos: vec![message],
            should_merge: None,
        };

        let error = ArchiveInfo::new(0, info.message_infos.clone(), limits)
            .expect_err("constructor admission must inspect public message length");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                observed,
                maximum,
            } if observed == Limits::MAX_MESSAGE_BYTES + 1
                && maximum == Limits::MAX_MESSAGE_BYTES
        ));

        let encoded = encode_archive_info(&info, limits)?;
        let error = ArchiveInfo::decode(&encoded, limits)
            .expect_err("decoded minimal metadata must enforce message length");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                observed,
                maximum,
            } if observed == Limits::MAX_MESSAGE_BYTES + 1
                && maximum == Limits::MAX_MESSAGE_BYTES
        ));
        Ok(())
    }

    #[test]
    fn duplicate_identifier_position_matches_the_ordered_scan() -> Result<()> {
        let limits = Limits::default();
        let object = |identifier| {
            ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 1,
                    data: vec![1],
                }],
                limits,
            )
        };
        // Small (pairwise) and large (sorted) archives; the reported object is
        // always the first one whose identifier already occurred.
        for count in [4u64, 40, 200] {
            let mut objects = (1..=count).map(object).collect::<Result<Vec<_>>>()?;
            assert!(
                Archive {
                    objects: objects.clone()
                }
                .encoded_len(limits)
                .is_ok()
            );
            let last = usize::try_from(count).expect("small count") - 1;
            objects[last] = object(2)?;
            objects[last - 1] = object(3)?;
            let expected = (0..objects.len())
                .find(|&index| {
                    objects[..index].iter().any(|previous| {
                        previous.archive_info.identifier == objects[index].archive_info.identifier
                    })
                })
                .expect("fixture has a duplicate");
            let error = Archive { objects }.encoded_len(limits).unwrap_err();
            assert!(
                matches!(
                    error,
                    Error::InvalidArchive {
                        offset,
                        reason: "duplicate object identifier",
                    } if offset == expected
                ),
                "{count}: {error:?}"
            );
        }
        // A missing identifier before the first duplicate is reported first.
        let mut objects = (1..=40).map(object).collect::<Result<Vec<_>>>()?;
        objects[39] = object(1)?;
        objects[5].archive_info.identifier = None;
        assert!(matches!(
            Archive { objects }.encoded_len(limits),
            Err(Error::InvalidArchive {
                offset: 5,
                reason: "object is missing its archive identifier",
            })
        ));
        Ok(())
    }

    #[test]
    fn default_path_certificates_match_exact_admission_at_the_metadata_ceiling() -> Result<()> {
        let limits = Limits::default();
        let ceiling = Limits::MAX_METADATA_ITEMS;
        // Fast certificate, its first excluded length (exact path, admitted)
        // and the first refused length for a standalone path.
        for components in [ceiling - 1, ceiling] {
            let path = FieldPath::new(vec![u32::MAX; components], limits)?;
            assert_eq!(path.path.len(), components);
            super::check_field_path_constructor(&path, limits)?;
        }
        let refused = FieldPath::new(vec![0; ceiling + 1], limits)
            .expect_err("one component over the metadata ceiling");
        assert!(matches!(
            refused,
            Error::Limit { kind: LimitKind::MetadataItems, observed, maximum }
                if observed == ceiling + 1 && maximum == ceiling
        ));

        // A minimal FieldInfo reserves one metadata item for itself.
        let admitted = FieldPath {
            path: vec![u32::MAX; ceiling - 1],
        };
        super::check_minimal_field_info_constructor(&admitted, limits)?;
        let info = FieldInfo::new(admitted, limits)?;
        assert_eq!(info.path.path.len(), ceiling - 1);
        let refused = FieldInfo::new(
            FieldPath {
                path: vec![0; ceiling],
            },
            limits,
        )
        .expect_err("FieldInfo at the path ceiling exceeds the metadata budget");
        assert!(matches!(
            refused,
            Error::Limit { kind: LimitKind::MetadataItems, observed, maximum }
                if observed == ceiling + 1 && maximum == ceiling
        ));

        // A stricter profile never takes the default certificate.
        let strict = limits.with_metadata_items(2)?;
        assert!(FieldPath::new(vec![1, 2], strict).is_ok());
        assert!(matches!(
            FieldPath::new(vec![1, 2, 3], strict),
            Err(Error::Limit {
                kind: LimitKind::MetadataItems,
                observed: 3,
                maximum: 2
            })
        ));
        assert!(matches!(
            FieldInfo::new(FieldPath { path: vec![1, 2] }, strict),
            Err(Error::Limit {
                kind: LimitKind::MetadataItems,
                observed: 3,
                maximum: 2
            })
        ));
        Ok(())
    }

    #[test]
    fn default_archive_info_certificate_agrees_with_exact_admission() -> Result<()> {
        let limits = Limits::default();
        let mut records = Vec::new();
        for versions in [
            &[][..],
            &[0][..],
            &[u32::MAX, u32::MAX][..],
            &[u32::MAX; 3][..],
        ] {
            let mut message = MessageInfo::new(u32::MAX, Limits::MAX_MESSAGE_BYTES, limits)?;
            message.versions = smallvec::SmallVec::from_slice(versions);
            assert!(!message.versions.spilled());
            records.push(message);
        }
        assert!(default_archive_info_bounds_are_safe(&records));
        let exact = ArchiveInfo {
            identifier: Some(u64::MAX),
            message_infos: records.clone(),
            should_merge: None,
        };
        check_archive_info_constructor(&exact, limits)?;
        let info = ArchiveInfo::new(u64::MAX, records.clone(), limits)?;
        assert_eq!(info.message_infos, records);

        // Each field outside the certified shape takes the exact path.
        let base = records[0].clone();
        let mut outside = Vec::new();
        let mut message = base.clone();
        message.length = u32::try_from(Limits::MAX_MESSAGE_BYTES + 1).expect("fits u32");
        outside.push(message);
        let mut message = base.clone();
        message.versions = smallvec::SmallVec::from_slice(&[1, 2, 3, 4]);
        outside.push(message);
        let mut message = base.clone();
        message.field_infos.push(FieldInfo::default());
        outside.push(message);
        let mut message = base.clone();
        message.object_references.push(1);
        outside.push(message);
        let mut message = base.clone();
        message.data_references.push(1);
        outside.push(message);
        let mut message = base.clone();
        message.base_message_index = Some(0);
        outside.push(message);
        let mut message = base.clone();
        message.diff_merge_version.push(1);
        outside.push(message);
        let mut message = base.clone();
        message.diff_field_path = Some(FieldPath::default());
        outside.push(message);
        let mut message = base.clone();
        message.fields_to_remove.push(FieldPath::default());
        outside.push(message);
        let mut message = base;
        message.diff_read_version.push(1);
        outside.push(message);
        for message in outside {
            assert!(!default_archive_info_bounds_are_safe(std::slice::from_ref(
                &message
            )));
        }
        assert!(matches!(
            ArchiveInfo::new(
                1,
                vec![{
                    let mut message = records[0].clone();
                    message.length =
                        u32::try_from(Limits::MAX_MESSAGE_BYTES + 1).expect("fits u32");
                    message
                }],
                limits,
            ),
            Err(Error::Limit {
                kind: LimitKind::MessageBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn default_archive_info_admission_checks_public_message_version_mutations() -> Result<()> {
        let limits = Limits::default();
        let mut message = MessageInfo::new(7, 1, limits)?;
        message.versions[0] = u32::MAX;
        // Inline version values do not change the admitted resources, so the
        // value-independent certificate and the exact path must agree.
        assert!(default_archive_info_bounds_are_safe(std::slice::from_ref(
            &message
        )));
        let exact = ArchiveInfo {
            identifier: Some(0),
            message_infos: vec![message.clone()],
            should_merge: None,
        };
        check_archive_info_constructor(&exact, limits)?;
        let mut grown = message.clone();
        grown.versions.push(9);
        assert!(grown.versions.spilled());
        assert!(!default_archive_info_bounds_are_safe(std::slice::from_ref(
            &grown
        )));

        let info = ArchiveInfo::new(0, vec![message], limits)?;
        assert_eq!(info.message_infos[0].versions.as_slice(), [u32::MAX, 0, 5]);
        Ok(())
    }

    #[test]
    fn spilled_default_versions_do_not_use_the_minimal_collection_ledger() -> Result<()> {
        let defaults = Limits::default();
        let spilled_message = || {
            let mut message = MessageInfo::new(7, 1, defaults).expect("default metadata");
            message.versions.reserve(1);
            message
        };
        let message = spilled_message();
        assert!(message.versions.spilled());
        assert_eq!(message.versions.as_slice(), DEFAULT_MESSAGE_VERSIONS);

        let raw = ArchiveInfo {
            identifier: Some(42),
            message_infos: vec![message],
            should_merge: None,
        };
        let expected = check_archive_info_constructor(&raw, defaults)?;
        let limits = defaults.with_header_memory_bytes(expected.header_memory)?;
        let mut collection = MessageInfoCollection::new(limits)?;
        collection.push(spilled_message())?;
        let admitted = ArchiveInfo::new(42, collection, limits)?;
        assert!(admitted.message_infos[0].versions.spilled());
        assert_eq!(
            admitted.message_infos[0].versions.as_slice(),
            DEFAULT_MESSAGE_VERSIONS
        );
        Ok(())
    }

    #[test]
    fn message_info_collection_admission_is_invalidated_by_public_mutation() -> Result<()> {
        let limits = Limits::default();
        let mut collection = MessageInfoCollection::with_capacity(2, limits)?;
        collection.push_new(7, 1)?;
        collection.push_new(8, 2)?;
        let archive = ArchiveInfo::new(42, collection, limits)?;
        assert_eq!(archive.message_infos.len(), 2);

        let mut collection = MessageInfoCollection::new(limits)?;
        collection.push_new(7, 1)?;
        collection
            .get_mut(0)
            .expect("the pushed message exists")
            .length = u32::try_from(Limits::MAX_MESSAGE_BYTES + 1)
            .expect("the hard message limit fits in the wire u32");
        let error = ArchiveInfo::new(42, collection, limits)
            .expect_err("a public mutation must force exact admission");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                observed,
                maximum,
            } if observed == Limits::MAX_MESSAGE_BYTES + 1
                && maximum == Limits::MAX_MESSAGE_BYTES
        ));

        let mut collection = MessageInfoCollection::new(limits)?;
        collection.push_new(7, 1)?;
        collection.as_mut_slice()[0].versions.clear();
        let archive = ArchiveInfo::new(42, collection, limits)?;
        assert!(archive.message_infos[0].versions.is_empty());
        Ok(())
    }

    #[test]
    fn message_info_collection_admission_materializes_minimal_prefix() -> Result<()> {
        let defaults = Limits::default();
        let mut nonminimal = MessageInfo::new(8, 2, defaults)?;
        nonminimal.versions.clear();
        let raw = ArchiveInfo {
            identifier: Some(42),
            message_infos: vec![MessageInfo::new(7, 1, defaults)?, nonminimal.clone()],
            should_merge: None,
        };
        let expected = check_archive_info_constructor(&raw, defaults)?;
        let limits = defaults.with_header_fields(expected.header_fields - 1)?;
        let mut collection = MessageInfoCollection::with_capacity(2, limits)?;
        collection.push_new(7, 1)?;
        collection.push(nonminimal)?;
        let error = ArchiveInfo::new(42, collection, limits)
            .expect_err("the minimal prefix must remain in aggregate field admission");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                observed,
                maximum,
            } if observed == expected.header_fields && maximum == expected.header_fields - 1
        ));
        Ok(())
    }

    #[test]
    fn limits_default_profile_marker_tracks_real_changes() -> Result<()> {
        let defaults = Limits::default();
        assert!(defaults.is_default_profile());
        assert!(defaults.with_header_nesting(16)?.is_default_profile());
        assert!(!defaults.with_header_fields(1)?.is_default_profile());
        Ok(())
    }

    #[test]
    fn raw_metadata_constructors_enforce_each_header_budget_at_the_boundary() -> Result<()> {
        let defaults = Limits::default();
        assert!(FieldPath::new(Vec::new(), defaults).is_ok());
        let path = FieldPath::new(vec![0, 7], defaults)?;
        let path_length = 1 + 1 + 2;
        assert_eq!(encoded_field_path_length(&path)?, path_length);
        assert!(FieldPath::new(vec![0, 7], defaults.with_header_bytes(path_length)?,).is_ok());
        let error = FieldPath::new(
            vec![0, 7],
            defaults.with_header_bytes(path_length.saturating_sub(1))?,
        )
        .expect_err("one byte below a raw path header must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderBytes,
                ..
            }
        ));
        let path_memory = size_of::<FieldPath>() + 2 * size_of::<u32>();
        assert_eq!(field_path_memory(&path)?, path_memory);
        assert!(
            FieldPath::new(vec![0, 7], defaults.with_header_memory_bytes(path_memory)?,).is_ok()
        );
        let error = FieldPath::new(
            vec![0, 7],
            defaults.with_header_memory_bytes(path_memory.saturating_sub(1))?,
        )
        .expect_err("one byte below neutral path memory must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderMemoryBytes,
                ..
            }
        ));
        let path_metadata = 2;
        assert_eq!(path.path.len(), path_metadata);
        assert!(FieldPath::new(vec![0, 7], defaults.with_metadata_items(path_metadata)?,).is_ok());
        let error = FieldPath::new(
            vec![0, 7],
            defaults.with_metadata_items(path_metadata.saturating_sub(1))?,
        )
        .expect_err("one path item below the metadata budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataItems,
                ..
            }
        ));

        let field = FieldInfo::new(path, defaults)?;
        let field_fields = 2;
        assert_eq!(field_info_header_field_count(&field)?, field_fields);
        assert!(
            FieldInfo::new(
                field.path.clone(),
                defaults.with_header_fields(field_fields)?,
            )
            .is_ok()
        );
        let error = FieldInfo::new(
            field.path.clone(),
            defaults.with_header_fields(field_fields.saturating_sub(1))?,
        )
        .expect_err("one field below a raw FieldInfo header must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            }
        ));
        let field_length = 1 + 1 + path_length;
        assert_eq!(encoded_field_info_length(&field)?, field_length);
        assert!(
            FieldInfo::new(
                field.path.clone(),
                defaults.with_header_bytes(field_length)?,
            )
            .is_ok()
        );
        let error = FieldInfo::new(
            field.path.clone(),
            defaults.with_header_bytes(field_length.saturating_sub(1))?,
        )
        .expect_err("one byte below a raw FieldInfo header must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderBytes,
                ..
            }
        ));
        let field_memory = size_of::<FieldInfo>() + 2 * size_of::<u32>();
        assert_eq!(field_info_memory(&field)?, field_memory);
        assert!(
            FieldInfo::new(
                field.path.clone(),
                defaults.with_header_memory_bytes(field_memory)?,
            )
            .is_ok()
        );
        let error = FieldInfo::new(
            field.path.clone(),
            defaults.with_header_memory_bytes(field_memory.saturating_sub(1))?,
        )
        .expect_err("one byte below neutral FieldInfo memory must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderMemoryBytes,
                ..
            }
        ));
        let field_metadata = 1 + 2;
        assert!(
            FieldInfo::new(
                field.path.clone(),
                defaults.with_metadata_items(field_metadata)?,
            )
            .is_ok()
        );
        let error = FieldInfo::new(
            field.path.clone(),
            defaults.with_metadata_items(field_metadata.saturating_sub(1))?,
        )
        .expect_err("one FieldInfo item below the metadata budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataItems,
                ..
            }
        ));

        let mut message = MessageInfo::new(7, 1, defaults)?;
        let message_length = 2 + 5 + 2;
        assert_eq!(minimal_message_info_encoded_length(7, 1)?, message_length);
        assert!(MessageInfo::new(7, 1, defaults.with_header_bytes(message_length)?,).is_ok());
        let error = MessageInfo::new(
            7,
            1,
            defaults.with_header_bytes(message_length.saturating_sub(1))?,
        )
        .expect_err("one byte below a raw MessageInfo header must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderBytes,
                ..
            }
        ));
        let message_memory = size_of::<MessageInfo>();
        assert_eq!(message_info_memory(&message)?, message_memory);
        assert!(
            MessageInfo::new(7, 1, defaults.with_header_memory_bytes(message_memory)?,).is_ok()
        );
        let error = MessageInfo::new(
            7,
            1,
            defaults.with_header_memory_bytes(message_memory.saturating_sub(1))?,
        )
        .expect_err("one byte below neutral MessageInfo memory must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderMemoryBytes,
                ..
            }
        ));
        let mut spilled_message = message.clone();
        spilled_message.versions.push(9);
        assert!(spilled_message.versions.spilled());
        let spilled_memory = size_of::<MessageInfo>() + 4 * size_of::<u32>();
        assert_eq!(message_info_memory(&spilled_message)?, spilled_memory);
        assert!(MessageInfo::new(7, 1, defaults.with_header_fields(3)?).is_ok());
        let error = MessageInfo::new(7, 1, defaults.with_header_fields(2)?)
            .expect_err("one MessageInfo field below the header budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            }
        ));
        assert!(
            MessageInfo::new(
                7,
                1,
                defaults.with_metadata_items(DEFAULT_MESSAGE_VERSIONS.len())?,
            )
            .is_ok()
        );
        let error = MessageInfo::new(
            7,
            1,
            defaults.with_metadata_items(DEFAULT_MESSAGE_VERSIONS.len() - 1)?,
        )
        .expect_err("one MessageInfo item below the metadata budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataItems,
                ..
            }
        ));

        let mut path_message = MessageInfo::new(7, 1, defaults)?;
        path_message.diff_field_path = Some(FieldPath::new(vec![1], defaults)?);
        path_message.fields_to_remove.push(FieldPath::default());
        // type + versions + length, then diff path's outer and inner fields,
        // then the empty removed path's outer field only.
        assert_eq!(message_info_header_field_count(&path_message)?, 6);
        let error = ArchiveInfo::new(0, vec![path_message], defaults.with_header_fields(5)?)
            .expect_err("nested diff paths must count their own wire fields");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            }
        ));

        message.field_infos.push(field.clone());
        let nested = ArchiveInfo::new(0, vec![message.clone()], defaults)?;
        let archive_nesting = message_info_header_nesting(&message)? + 1;
        assert!(
            ArchiveInfo::new(
                0,
                vec![message.clone()],
                defaults.with_header_nesting(archive_nesting)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(
            0,
            vec![message.clone()],
            defaults.with_header_nesting(archive_nesting.saturating_sub(1))?,
        )
        .expect_err("one nesting level below a raw ArchiveInfo must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderNesting,
                ..
            }
        ));

        let archive_metadata =
            nested
                .message_infos
                .iter()
                .try_fold(nested.message_infos.len(), |total, info| {
                    total
                        .checked_add(info.metadata_item_count()?)
                        .ok_or_else(|| Error::invalid_archive(0, "test metadata count overflow"))
                })?;
        assert!(
            ArchiveInfo::new(
                0,
                nested.message_infos.clone(),
                defaults.with_metadata_items(archive_metadata)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(
            0,
            nested.message_infos.clone(),
            defaults.with_metadata_items(archive_metadata.saturating_sub(1))?,
        )
        .expect_err("one metadata item below a raw ArchiveInfo must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataItems,
                ..
            }
        ));

        let archive_header = 2 + 1 + 1 + 17;
        assert_eq!(archive_info_encoded_len(&nested)?, archive_header);
        assert!(
            ArchiveInfo::new(
                0,
                nested.message_infos.clone(),
                defaults.with_header_bytes(archive_header)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(
            0,
            nested.message_infos.clone(),
            defaults.with_header_bytes(archive_header.saturating_sub(1))?,
        )
        .expect_err("one byte below a raw ArchiveInfo must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderBytes,
                ..
            }
        ));
        let archive_fields = 1 + 1 + 6;
        assert_eq!(
            message_info_header_field_count(&nested.message_infos[0])?,
            6
        );
        assert!(
            ArchiveInfo::new(
                0,
                nested.message_infos.clone(),
                defaults.with_header_fields(archive_fields)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(
            0,
            nested.message_infos.clone(),
            defaults.with_header_fields(archive_fields.saturating_sub(1))?,
        )
        .expect_err("one ArchiveInfo field below the header budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            }
        ));

        let two_messages = vec![
            MessageInfo::new(7, 1, defaults)?,
            MessageInfo::new(8, 1, defaults)?,
        ];
        assert!(
            ArchiveInfo::new(
                0,
                two_messages.clone(),
                defaults.with_messages_per_object(2)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(0, two_messages, defaults.with_messages_per_object(1)?)
            .expect_err("one message above the per-object raw budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessagesPerObject,
                ..
            }
        ));

        let expected_field_memory = size_of::<FieldInfo>() + 2 * size_of::<u32>();
        let expected_message_memory = size_of::<MessageInfo>() + expected_field_memory;
        let archive_memory = size_of::<ArchiveInfo>() + expected_message_memory;
        assert_eq!(
            field_info_memory(&nested.message_infos[0].field_infos[0])?,
            expected_field_memory
        );
        assert_eq!(
            message_info_memory(&nested.message_infos[0])?,
            expected_message_memory
        );
        assert!(
            ArchiveInfo::new(
                0,
                vec![message.clone()],
                defaults.with_header_memory_bytes(archive_memory)?,
            )
            .is_ok()
        );
        let error = ArchiveInfo::new(
            0,
            vec![message.clone()],
            defaults.with_header_memory_bytes(archive_memory.saturating_sub(1))?,
        )
        .expect_err("one byte below neutral metadata memory must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderMemoryBytes,
                ..
            }
        ));

        let error = MessageInfo::new(7, 2, defaults.with_message_bytes(1)?)
            .expect_err("payload above the raw message budget must be refused");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                ..
            }
        ));
        let error = MessageInfo::new(7, usize::try_from(u32::MAX).unwrap() + 1, defaults)
            .expect_err("payload lengths above the configured budget must be refused first");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                ..
            }
        ));
        let error =
            encoded_message_info_length(&message, 7, usize::try_from(u32::MAX).unwrap() + 1)
                .expect_err("the wire length conversion must reject values above u32");
        assert!(matches!(error, Error::InvalidArchive { .. }));

        Ok(())
    }

    #[test]
    fn archive_object_preflights_header_before_metadata_allocation() -> Result<()> {
        let message_set = || {
            vec![
                RawMessage {
                    type_: 7,
                    data: Vec::new(),
                },
                RawMessage {
                    type_: 8,
                    data: Vec::new(),
                },
            ]
        };
        let constructions = || MESSAGE_INFO_CONSTRUCTIONS.with(|count| count.get());

        let assert_rejected_before_metadata = |limits: Limits, expected: LimitKind| {
            let before = constructions();
            let error = ArchiveObject::new(0, message_set(), limits)
                .expect_err("aggregate header limits must reject before metadata construction");
            assert!(matches!(
                error,
                Error::Limit { kind, .. } if kind == expected
            ));
            assert_eq!(
                constructions(),
                before,
                "MessageInfo::conventional must not run for a rejected aggregate header"
            );
        };

        assert_rejected_before_metadata(
            Limits::default().with_header_fields(1)?,
            LimitKind::HeaderFields,
        );
        assert_rejected_before_metadata(
            Limits::default().with_header_bytes(1)?,
            LimitKind::HeaderBytes,
        );
        assert_rejected_before_metadata(
            Limits::default().with_header_memory_bytes(size_of::<ArchiveInfo>())?,
            LimitKind::HeaderMemoryBytes,
        );

        let before = constructions();
        ArchiveObject::new(0, message_set(), Limits::default())
            .expect("an admitted aggregate header must construct metadata");
        assert_eq!(
            constructions() - before,
            2,
            "an admitted object must construct one MessageInfo per payload"
        );
        Ok(())
    }

    #[test]
    fn encoded_len_matches_retained_unknown_header_and_exact_limits() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let archive = Archive::parse(&fixture.source, Default::default())?;
        assert_eq!(
            archive.encoded_len(Default::default())?,
            fixture.source.len()
        );

        let exact_limits = Limits::default().with_archive_bytes(fixture.source.len())?;
        assert_eq!(archive.encoded_len(exact_limits)?, fixture.source.len());
        assert_eq!(archive.to_bytes(exact_limits)?, fixture.source);

        let below_limits = Limits::default().with_archive_bytes(fixture.source.len() - 1)?;
        let error = archive
            .encoded_len(below_limits)
            .expect_err("one byte below the exact archive length must fail");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::ArchiveBytes | LimitKind::ObjectBytes,
                observed,
                maximum
            } if observed == fixture.source.len() && maximum == fixture.source.len() - 1
        ));
        assert!(archive.to_bytes(below_limits).is_err());
        Ok(())
    }

    #[test]
    fn retained_admission_rechecks_public_mutations_and_strict_dimensions() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Limits::default())?;
        let admission =
            check_archive_info_constructor(&archive.objects[0].archive_info, Limits::default())?;
        let mut strict_profiles = Vec::new();
        if admission.header_length > 1 {
            strict_profiles.push((
                LimitKind::HeaderBytes,
                admission.header_length,
                Limits::default().with_header_bytes(admission.header_length - 1)?,
            ));
        }
        if admission.header_fields > 1 {
            strict_profiles.push((
                LimitKind::HeaderFields,
                admission.header_fields,
                Limits::default().with_header_fields(admission.header_fields - 1)?,
            ));
        }
        if admission.header_nesting > 1 {
            strict_profiles.push((
                LimitKind::HeaderNesting,
                admission.header_nesting,
                Limits::default().with_header_nesting(admission.header_nesting - 1)?,
            ));
        }
        if admission.header_memory > 1 {
            strict_profiles.push((
                LimitKind::HeaderMemoryBytes,
                admission.header_memory,
                Limits::default().with_header_memory_bytes(admission.header_memory - 1)?,
            ));
        }
        if admission.metadata_items > 1 {
            strict_profiles.push((
                LimitKind::MetadataItems,
                admission.metadata_items,
                Limits::default().with_metadata_items(admission.metadata_items - 1)?,
            ));
        }
        if admission.message_count > 1 {
            strict_profiles.push((
                LimitKind::MessagesPerObject,
                admission.message_count,
                Limits::default().with_messages_per_object(admission.message_count - 1)?,
            ));
        }
        for (kind, observed, limits) in strict_profiles {
            let error = archive.objects[0]
                .validate(limits)
                .expect_err("a retained header must reapply each stricter dimension");
            assert!(matches!(
                error,
                Error::Limit {
                    kind: actual,
                    observed: actual_observed,
                    maximum,
                } if actual == kind && actual_observed == observed && maximum == observed - 1
            ));
        }

        let object = &mut archive.objects[0];
        object.archive_info.message_infos[0].type_ += 1;
        object.validate(Limits::default())?;
        let mutated_admission =
            check_archive_info_constructor(&object.archive_info, Limits::default())?;
        let error = object
            .validate(Limits::default().with_header_fields(mutated_admission.header_fields - 1)?)
            .expect_err("metadata mutation must invalidate the cached canonical admission");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::HeaderFields,
                observed,
                maximum,
            } if observed == mutated_admission.header_fields
                && maximum == mutated_admission.header_fields - 1
        ));

        let message_limit = object.messages[0].data.len();
        object.messages[0].data.push(0);
        let error = object
            .validate(Limits::default().with_message_bytes(message_limit)?)
            .expect_err("payload mutation must still report MessageBytes");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MessageBytes,
                observed,
                maximum,
            } if observed == message_limit + 1 && maximum == message_limit
        ));
        Ok(())
    }

    #[test]
    fn same_content_ignoring_offsets_accepts_identical_objects_at_different_offsets() -> Result<()>
    {
        let source = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![3, 4],
                }],
                Default::default(),
            )?],
        }
        .to_bytes(Default::default())?;
        let object = Archive::parse(&source, Default::default())?.objects[0].clone();
        let mut alias = object.clone();
        alias.header_offset = 4096;
        alias.data_offset = alias.header_offset + alias.header_length;

        assert!(object.same_content_ignoring_offsets(&alias));
        assert_ne!(object.header_offset, alias.header_offset);
        assert_ne!(object.data_offset, alias.data_offset);
        Ok(())
    }

    #[test]
    fn same_content_ignoring_offsets_rejects_divergent_published_content() -> Result<()> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![3, 4],
            }],
            Default::default(),
        )?;
        let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
        let mut retained = canonical.clone();
        retained.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);
        object.set_test_source_headers(
            Some(retained.into_boxed_slice()),
            Some(canonical.into_boxed_slice()),
        );

        let mut payload = object.clone();
        payload.messages[0].data[0] ^= 1;
        assert!(!object.same_content_ignoring_offsets(&payload));

        let mut metadata = object.clone();
        metadata.archive_info.message_infos[0].type_ += 1;
        assert!(!object.same_content_ignoring_offsets(&metadata));

        let mut raw_header = object.clone();
        raw_header.test_original_header_mut().expect("raw header")[0] ^= 1;
        assert!(!object.same_content_ignoring_offsets(&raw_header));

        let mut canonical_header = object.clone();
        canonical_header
            .test_original_canonical_header_mut()
            .expect("canonical header")[0] ^= 1;
        assert!(!object.same_content_ignoring_offsets(&canonical_header));

        let mut raw_header_state = object.clone();
        raw_header_state.clear_test_source_headers();
        assert!(!object.same_content_ignoring_offsets(&raw_header_state));

        let mut framing = object.clone();
        framing.header_length += 1;
        assert!(!object.same_content_ignoring_offsets(&framing));

        let mut payload_framing = object.clone();
        payload_framing.data_length += 1;
        assert!(!object.same_content_ignoring_offsets(&payload_framing));
        Ok(())
    }

    #[test]
    fn same_content_ignoring_offsets_distinguishes_new_and_canonical_provenance() -> Result<()> {
        let fresh = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![3, 4],
            }],
            Default::default(),
        )?;
        let mut fresh_alias = fresh.clone();
        fresh_alias.header_offset = 128;
        fresh_alias.data_offset = 256;
        assert!(fresh.same_content_ignoring_offsets(&fresh_alias));

        let encoded = Archive {
            objects: vec![fresh.clone()],
        }
        .to_bytes(Default::default())?;
        let parsed = Archive::parse(&encoded, Default::default())?.objects[0].clone();
        assert!(!fresh.same_content_ignoring_offsets(&parsed));

        let mut parsed_alias = parsed.clone();
        parsed_alias.header_offset = 512;
        parsed_alias.data_offset = 768;
        assert!(parsed.same_content_ignoring_offsets(&parsed_alias));
        Ok(())
    }

    #[test]
    fn canonical_object_framing_checks_supplied_resource_limits() -> Result<()> {
        let source = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![
                        RawMessage {
                            type_: 2,
                            data: vec![3, 4],
                        },
                        RawMessage {
                            type_: 3,
                            data: vec![7, 8],
                        },
                    ],
                    Limits::default(),
                )?,
                ArchiveObject::new(
                    2,
                    vec![
                        RawMessage {
                            type_: 2,
                            data: vec![5, 6],
                        },
                        RawMessage {
                            type_: 3,
                            data: vec![9, 0],
                        },
                    ],
                    Limits::default(),
                )?,
            ],
        }
        .to_bytes(Limits::default())?;
        let archive = Archive::parse(&source, Limits::default())?;
        for (limits, expected_kind) in [
            (
                Limits::default().with_archive_bytes(source.len() - 1)?,
                LimitKind::ArchiveBytes,
            ),
            (Limits::default().with_objects(1)?, LimitKind::Objects),
            (Limits::default().with_messages(2)?, LimitKind::Messages),
            (
                Limits::default().with_messages_per_object(1)?,
                LimitKind::MessagesPerObject,
            ),
            (
                Limits::default().with_object_bytes(
                    usize::try_from(
                        archive.objects[0].header_length + archive.objects[0].data_length,
                    )
                    .expect("small fixture object")
                        - 1,
                )?,
                LimitKind::ObjectBytes,
            ),
            (
                Limits::default().with_message_bytes(1)?,
                LimitKind::MessageBytes,
            ),
            (
                Limits::default().with_header_bytes(1)?,
                LimitKind::HeaderBytes,
            ),
        ] {
            assert!(
                matches!(archive.validate_canonical_object_framing(&source, limits),
                Err(Error::Limit { kind, .. }) if kind == expected_kind)
            );
        }
        archive.validate_canonical_object_framing(&source, Limits::default())?;
        Ok(())
    }

    #[test]
    fn canonical_object_framing_accepts_unknown_raw_archive_info_bytes() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let archive = Archive::parse(&fixture.source, Default::default())?;
        archive.validate_canonical_object_framing(&fixture.source, Default::default())?;
        assert_eq!(archive.to_bytes(Default::default())?, fixture.source);
        assert_eq!(
            archive.objects[0].source_header(),
            Some(fixture.header.as_slice())
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_overlong_prefix_on_any_object() -> Result<()> {
        let canonical = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 2,
                        data: vec![3],
                    }],
                    Default::default(),
                )?,
                ArchiveObject::new(
                    4,
                    vec![RawMessage {
                        type_: 5,
                        data: vec![6, 7],
                    }],
                    Default::default(),
                )?,
            ],
        }
        .to_bytes(Default::default())?;
        let parsed = Archive::parse(&canonical, Default::default())?;
        parsed.validate_canonical_object_framing(&canonical, Default::default())?;
        let second_offset = usize::try_from(parsed.objects[1].header_offset)
            .map_err(|_| Error::invalid_archive(0, "test offset exceeds usize"))?;
        let (header_length, prefix_length) = super::decode_varint(
            canonical
                .get(second_offset..)
                .ok_or_else(|| Error::invalid_archive(0, "test object offset is invalid"))?,
        )?;
        assert_eq!(prefix_length, 1);
        assert!(header_length < 0x80);

        let mut overlong = Vec::new();
        overlong
            .try_reserve_exact(canonical.len() + 1)
            .map_err(|_| Error::allocation("test overlong archive", canonical.len() + 1))?;
        overlong.extend_from_slice(&canonical[..second_offset]);
        overlong.push(canonical[second_offset] | 0x80);
        overlong.push(0);
        overlong.extend_from_slice(&canonical[second_offset + 1..]);

        // Parsing remains deliberately permissive; only the opt-in framing
        // validator refuses the overlong second-object prefix.
        let permissive = Archive::parse(&overlong, Default::default())?;
        assert!(
            permissive
                .validate_canonical_object_framing(&overlong, Default::default())
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_truncated_source_regions() -> Result<()> {
        let source = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![3, 4, 5],
                }],
                Default::default(),
            )?],
        }
        .to_bytes(Default::default())?;
        let archive = Archive::parse(&source, Default::default())?;
        let data_offset = usize::try_from(archive.objects[0].data_offset)
            .map_err(|_| Error::invalid_archive(0, "test data offset exceeds usize"))?;

        let truncated_prefix = [source[0] | 0x80];
        assert!(
            archive
                .validate_canonical_object_framing(&truncated_prefix, Default::default())
                .is_err()
        );
        assert!(
            archive
                .validate_canonical_object_framing(&source[..data_offset - 1], Default::default())
                .is_err()
        );
        assert!(
            archive
                .validate_canonical_object_framing(&source[..source.len() - 1], Default::default())
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_inexact_offsets_and_extents() -> Result<()> {
        let source = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 2,
                        data: vec![3],
                    }],
                    Default::default(),
                )?,
                ArchiveObject::new(
                    4,
                    vec![RawMessage {
                        type_: 5,
                        data: vec![6, 7],
                    }],
                    Default::default(),
                )?,
            ],
        }
        .to_bytes(Default::default())?;
        let archive = Archive::parse(&source, Default::default())?;

        let mut wrong_header_offset = archive.clone();
        wrong_header_offset.objects[1].header_offset += 1;
        assert!(
            wrong_header_offset
                .validate_canonical_object_framing(&source, Default::default())
                .is_err()
        );

        let mut wrong_header_length = archive.clone();
        wrong_header_length.objects[0].header_length += 1;
        assert!(
            wrong_header_length
                .validate_canonical_object_framing(&source, Default::default())
                .is_err()
        );

        let mut wrong_data_offset = archive.clone();
        wrong_data_offset.objects[0].data_offset += 1;
        assert!(
            wrong_data_offset
                .validate_canonical_object_framing(&source, Default::default())
                .is_err()
        );

        let mut wrong_data_length = archive.clone();
        wrong_data_length.objects[0].data_length += 1;
        assert!(
            wrong_data_length
                .validate_canonical_object_framing(&source, Default::default())
                .is_err()
        );

        let mut trailing = source.clone();
        trailing.push(0);
        assert!(
            archive
                .validate_canonical_object_framing(&trailing, Default::default())
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_message_length_width_change_and_inverse_restore_exact_bytes() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![0x11; 127],
        };
        let source = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![original_message.clone()],
                Default::default(),
            )?],
        }
        .to_bytes(Default::default())?;
        let mut edited = Archive::parse(&source, Default::default())?;
        let replacement = RawMessage {
            type_: 2,
            data: vec![0x22; 128],
        };

        let old = edited.objects[0].replace_message_preserving_header(
            0,
            replacement.clone(),
            Default::default(),
        )?;
        assert_eq!(old, original_message);
        let expected = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![replacement.clone()],
                Default::default(),
            )?],
        }
        .to_bytes(Default::default())?;
        assert_eq!(edited.to_bytes(Default::default())?, expected);
        assert!(edited.objects[0].source_header().is_none());
        assert!(edited.objects[0].source_canonical_header().is_none());

        let replaced =
            edited.objects[0].replace_message_preserving_header(0, old, Default::default())?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes(Default::default())?, source);
        assert!(edited.objects[0].source_header().is_none());
        assert!(edited.objects[0].source_canonical_header().is_none());
        Ok(())
    }

    #[test]
    fn retained_probe_falls_back_after_packed_field_length_changes() -> Result<()> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 7,
                data: vec![3],
            }],
            Default::default(),
        )?;
        object.archive_info.message_infos[0].object_references = vec![10];
        let source = Archive {
            objects: vec![object],
        }
        .to_bytes(Default::default())?;

        let mut parsed = Archive::parse(&source, Default::default())?;
        parsed.objects[0].archive_info.message_infos[0].object_references = vec![10, 11];
        let encoded = parsed.to_bytes(Default::default())?;
        assert_ne!(encoded, source);

        let reparsed = Archive::parse(&encoded, Default::default())?;
        assert_eq!(
            reparsed.objects[0].archive_info.message_infos[0].object_references,
            [10, 11]
        );
        assert_eq!(reparsed.to_bytes(Default::default())?, encoded);
        Ok(())
    }

    #[test]
    fn canonical_last_reference_pruning_and_inverse_restore_exact_bytes() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![3, 4],
        };
        let mut original_object =
            ArchiveObject::new(1, vec![original_message.clone()], Default::default())?;
        let Some(original_info) = original_object.archive_info.message_infos.get_mut(0) else {
            return Err(Error::invalid_archive(
                0,
                "test object is missing message metadata",
            ));
        };
        original_info.object_references.push(77);
        let source_archive = Archive {
            objects: vec![original_object],
        };
        let source = source_archive.to_bytes(Default::default())?;
        let mut edited = Archive::parse(&source, Default::default())?;

        let replacement = RawMessage {
            type_: 9,
            data: vec![5, 6, 7],
        };
        let old = edited.objects[0].replace_message_pruning_object_references_preserving_header(
            0,
            replacement.clone(),
            &[77],
            Default::default(),
        )?;
        assert_eq!(old, original_message);
        assert!(edited.objects[0].source_header().is_none());
        assert!(edited.objects[0].source_canonical_header().is_none());

        let expected = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![replacement.clone()],
                Default::default(),
            )?],
        }
        .to_bytes(Default::default())?;
        assert_eq!(edited.to_bytes(Default::default())?, expected);

        let Some(edited_info) = edited.objects[0].archive_info.message_infos.get_mut(0) else {
            return Err(Error::invalid_archive(
                0,
                "edited test object is missing message metadata",
            ));
        };
        edited_info.object_references.push(77);
        let replaced =
            edited.objects[0].replace_message_preserving_header(0, old, Default::default())?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes(Default::default())?, source);
        Ok(())
    }

    #[test]
    fn reference_pruning_preserves_all_untouched_header_encodings() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let old = archive.objects[0].replace_message_pruning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[10, 12, 14, 10],
            Default::default(),
        )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);

        let encoded = archive.to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.expected_header);
        assert_eq!(payload, vec![0x5a; 130]);

        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.type_, 9);
        assert_eq!(info.length, 130);
        // The duplicate retained identifier and its two distinct raw
        // encodings survive in source order.
        assert_eq!(info.object_references, [13, 13]);
        assert_eq!(info.field_infos[0].object_references, [99, 100]);
        assert_eq!(info.field_infos[1].object_references, [77, 78]);
        assert_eq!(
            archive.objects[0].source_header(),
            Some(fixture.expected_header.as_slice())
        );

        let reparsed = Archive::parse(&encoded, Default::default())?;
        assert_eq!(
            reparsed.objects[0].archive_info,
            archive.objects[0].archive_info
        );
        assert_eq!(reparsed.to_bytes(Default::default())?, encoded);
        Ok(())
    }

    #[test]
    fn empty_reference_removal_delegates_to_preserving_replacement() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let parsed = Archive::parse(&fixture.source, Default::default())?;
        let mut delegated = parsed.objects[0].clone();
        let mut existing = parsed.objects[0].clone();
        let replacement = RawMessage {
            type_: 9,
            data: vec![0x44; 130],
        };

        let delegated_old = delegated.replace_message_pruning_object_references_preserving_header(
            0,
            replacement.clone(),
            &[],
            Default::default(),
        )?;
        let existing_old =
            existing.replace_message_preserving_header(0, replacement, Default::default())?;
        assert_eq!(delegated_old, existing_old);
        assert_eq!(delegated, existing);
        Ok(())
    }

    #[test]
    fn reference_pruning_failures_are_atomic() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let before_limit = archive.objects[0].clone();
        let limits = Limits::default()
            .with_header_bytes(fixture.header.len())?
            .with_object_bytes(fixture.source.len())?;
        let limit_error = archive.objects[0]
            .replace_message_pruning_object_references_preserving_header(
                0,
                RawMessage {
                    type_: 9,
                    data: vec![0x5a; 130],
                },
                &[10],
                limits,
            )
            .err();
        assert!(matches!(
            limit_error,
            Some(Error::Limit {
                kind: LimitKind::ObjectBytes,
                ..
            })
        ));
        assert_eq!(archive.objects[0], before_limit);

        for malformed_message in [
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x29, 1, 2, 3, 4, 5, 6, 7, 8,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x2a, 0x01, 0x80,
            ],
        ] {
            let mut object = ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![0xcc],
                }],
                Default::default(),
            )?;
            let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
            let mut raw_header = vec![0x08, 0x01];
            push_length_delimited(&mut raw_header, &[0x12], 1, &malformed_message)?;
            object.set_test_source_headers(
                Some(raw_header.into_boxed_slice()),
                Some(canonical.into_boxed_slice()),
            );
            let before = object.clone();

            let malformed_error = object
                .replace_message_pruning_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 2,
                        data: vec![0xdd],
                    },
                    &[1],
                    Default::default(),
                )
                .err();
            assert!(malformed_error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn selected_object_and_data_pruning_preserves_raw_occurrences() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let old = archive.objects[0].replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[10, 12, 10],
            DataReferencePruning::Selected(&[20, 22, 20]),
            Default::default(),
        )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);

        let encoded = archive.to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.selected_header);
        assert_eq!(payload, vec![0x5a; 130]);

        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, [13, 13, 13]);
        assert_eq!(info.data_references, [21, 21, 21, 21]);
        assert_eq!(info.field_infos[0].object_references, [100]);
        assert_eq!(info.field_infos[0].data_references, [21, 21]);
        assert_eq!(info.field_infos[1].object_references, [77]);
        assert_eq!(info.field_infos[1].data_references, [31, 32, 33]);
        assert_eq!(
            archive.objects[0].source_header(),
            Some(fixture.selected_header.as_slice())
        );

        // A second parse proves that the raw result and its typed projection
        // remain paired for future header-preserving edits.
        let reparsed = Archive::parse(&encoded, Default::default())?;
        assert_eq!(reparsed.to_bytes(Default::default())?, encoded);
        Ok(())
    }

    #[test]
    fn all_data_pruning_clears_aggregate_and_nested_references_only() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        archive.objects[0].replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[],
            DataReferencePruning::All,
            Default::default(),
        )?;

        let encoded = archive.to_bytes(Default::default())?;
        let (header, _) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.all_data_header);
        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, [10, 10, 13, 13, 12, 13, 10]);
        assert!(info.data_references.is_empty());
        assert_eq!(info.field_infos[0].object_references, [10, 12, 100, 10]);
        assert!(info.field_infos[0].data_references.is_empty());
        assert_eq!(info.field_infos[1].object_references, [77]);
        assert!(info.field_infos[1].data_references.is_empty());
        Ok(())
    }

    #[test]
    fn data_pruning_zero_and_noop_policies_preserve_existing_behavior() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let parsed = Archive::parse(&fixture.source, Default::default())?;
        let replacement = RawMessage {
            type_: 9,
            data: vec![0x44; 130],
        };
        let mut delegated = parsed.objects[0].clone();
        let mut preserving = parsed.objects[0].clone();
        let delegated_old = delegated.replace_message_pruning_references_preserving_header(
            0,
            replacement.clone(),
            &[],
            DataReferencePruning::Selected(&[]),
            Default::default(),
        )?;
        let preserving_old =
            preserving.replace_message_preserving_header(0, replacement, Default::default())?;
        assert_eq!(delegated_old, preserving_old);
        assert_eq!(delegated, preserving);

        let mut unmatched = parsed.objects[0].clone();
        let same = unmatched.messages[0].clone();
        unmatched.replace_message_pruning_references_preserving_header(
            0,
            same,
            &[u64::MAX],
            DataReferencePruning::Selected(&[u64::MAX]),
            Default::default(),
        )?;
        assert_eq!(unmatched, parsed.objects[0]);
        assert_eq!(
            Archive {
                objects: vec![unmatched],
            }
            .to_bytes(Default::default())?,
            fixture.source
        );

        let mut empty = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![3],
            }],
            Default::default(),
        )?;
        let empty_before = Archive {
            objects: vec![empty.clone()],
        }
        .to_bytes(Default::default())?;
        empty.replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 2,
                data: vec![3],
            },
            &[],
            DataReferencePruning::All,
            Default::default(),
        )?;
        assert_eq!(
            Archive {
                objects: vec![empty],
            }
            .to_bytes(Default::default())?,
            empty_before
        );
        Ok(())
    }

    #[test]
    fn canonical_data_pruning_can_restore_exact_bytes_after_metadata_restore() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![3, 4],
        };
        let mut object = ArchiveObject::new(1, vec![original_message.clone()], Default::default())?;
        let info = &mut object.archive_info.message_infos[0];
        info.data_references.extend([7, 8]);
        let mut field = FieldInfo::new(
            FieldPath::new(vec![1], Limits::default()).unwrap(),
            Limits::default(),
        )
        .unwrap();
        field.data_references.extend([9, 10]);
        info.field_infos.push(field);
        let source = Archive {
            objects: vec![object],
        }
        .to_bytes(Default::default())?;
        let mut edited = Archive::parse(&source, Default::default())?;
        let replacement = RawMessage {
            type_: 9,
            data: vec![5, 6, 7],
        };
        let old = edited.objects[0].replace_message_pruning_references_preserving_header(
            0,
            replacement.clone(),
            &[],
            DataReferencePruning::Selected(&[8, 10]),
            Default::default(),
        )?;
        assert_eq!(
            edited.objects[0].archive_info.message_infos[0].data_references,
            [7]
        );
        assert_eq!(
            edited.objects[0].archive_info.message_infos[0].field_infos[0].data_references,
            [9]
        );

        let edited_info = &mut edited.objects[0].archive_info.message_infos[0];
        edited_info.data_references.insert(1, 8);
        edited_info.field_infos[0].data_references.insert(1, 10);
        let replaced =
            edited.objects[0].replace_message_preserving_header(0, old, Default::default())?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes(Default::default())?, source);
        Ok(())
    }

    #[test]
    fn combined_reference_pruning_budgets_selectors_and_wire_work_atomically() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let before = archive.objects[0].clone();
        let selector_limits = Limits::default().with_metadata_items(2)?;
        let selector_error = archive.objects[0]
            .replace_message_pruning_references_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: vec![0xde, 0xad, 0xbe],
                },
                &[10, 12],
                DataReferencePruning::Selected(&[20]),
                selector_limits,
            )
            .err();
        assert!(matches!(
            selector_error,
            Some(Error::Limit {
                kind: LimitKind::MetadataItems,
                observed: 3,
                maximum: 2,
            })
        ));
        assert_eq!(archive.objects[0], before);

        let wire_limits = Limits::default().with_header_fields(1)?;
        let wire_error = archive.objects[0]
            .replace_message_pruning_references_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: vec![0xde, 0xad, 0xbe],
                },
                &[],
                DataReferencePruning::All,
                wire_limits,
            )
            .err();
        assert!(matches!(
            wire_error,
            Some(Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            })
        ));
        assert_eq!(archive.objects[0], before);
        Ok(())
    }

    #[test]
    fn ambiguous_or_malformed_data_reference_fields_are_rejected_atomically() -> Result<()> {
        for malformed_message in [
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x31, 1, 2, 3, 4, 5, 6, 7, 8,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x32, 0x01, 0x80,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x22, 0x0a, 0x0a, 0x00, 0x29,
                1, 2, 3, 4, 5, 6, 7, 8,
            ],
        ] {
            let mut object = ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![0xcc],
                }],
                Default::default(),
            )?;
            let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
            let mut raw_header = vec![0x08, 0x01];
            push_length_delimited(&mut raw_header, &[0x12], 1, &malformed_message)?;
            object.set_test_source_headers(
                Some(raw_header.into_boxed_slice()),
                Some(canonical.into_boxed_slice()),
            );
            let before = object.clone();

            let error = object
                .replace_message_pruning_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 2,
                        data: vec![0xdd],
                    },
                    &[],
                    DataReferencePruning::All,
                    Default::default(),
                )
                .err();
            assert!(error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn data_reference_reorder_accepts_arbitrary_nonzero_ids_and_rejects_new_or_zero_ids()
    -> Result<()> {
        let aggregate = [u64::MAX, 7, 42];
        let reordered = [42, u64::MAX, 7];
        let mut object = data_reorder_object(&aggregate, &[&aggregate])?;
        object.replace_message_reordering_data_references_preserving_header(
            0,
            RawMessage {
                type_: 2,
                data: vec![0xdd],
            },
            &reordered,
            Default::default(),
        )?;
        let info = &object.archive_info.message_infos[0];
        assert_eq!(info.data_references, reordered);
        assert_eq!(info.field_infos[0].data_references, reordered);

        let mut unknown_identifier = data_reorder_object(&aggregate, &[&aggregate])?;
        let before_unknown = unknown_identifier.clone();
        let error = unknown_identifier
            .replace_message_reordering_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 2,
                    data: vec![0xdd],
                },
                &[42, 999, u64::MAX],
                Default::default(),
            )
            .err();
        assert!(matches!(
            error,
            Some(Error::InvalidArchive {
                reason: "reordered data-reference identifier is not in MessageInfo",
                ..
            })
        ));
        assert_eq!(unknown_identifier, before_unknown);

        let mut zero_source = data_reorder_object(&[0, 7, 42], &[&[0, 7, 42]])?;
        let before_zero_source = zero_source.clone();
        let error = zero_source
            .replace_message_reordering_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 2,
                    data: vec![0xdd],
                },
                &[42, 0, 7],
                Default::default(),
            )
            .err();
        assert!(matches!(
            error,
            Some(Error::InvalidArchive {
                reason: "data-reference identifiers must be nonzero",
                ..
            })
        ));
        assert_eq!(zero_source, before_zero_source);

        let mut zero_selector = data_reorder_object(&aggregate, &[&aggregate])?;
        let before_zero_selector = zero_selector.clone();
        let error = zero_selector
            .replace_message_reordering_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 2,
                    data: vec![0xdd],
                },
                &[42, 0, u64::MAX],
                Default::default(),
            )
            .err();
        assert!(matches!(
            error,
            Some(Error::InvalidArchive {
                reason: "data-reference identifiers must be nonzero",
                ..
            })
        ));
        assert_eq!(zero_selector, before_zero_selector);
        Ok(())
    }

    #[test]
    fn data_reference_reorder_rejects_ambiguous_matching_field_infos_atomically() -> Result<()> {
        let aggregate = [10, 20, 30];
        let mut object = data_reorder_object(&aggregate, &[&aggregate, &aggregate, &[10, 30]])?;
        let before = object.clone();
        let error = object
            .replace_message_reordering_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 2,
                    data: vec![0xdd],
                },
                &[30, 10, 20],
                Default::default(),
            )
            .err();
        assert!(matches!(
            error,
            Some(Error::InvalidArchive {
                reason: "data-reference reorder has ambiguous matching FieldInfo metadata",
                ..
            })
        ));
        assert_eq!(object, before);
        Ok(())
    }

    #[test]
    fn data_reference_reorder_noop_preserves_source_with_ambiguous_field_infos() -> Result<()> {
        let aggregate = [10, 20, 30];
        let mut object = data_reorder_object(&aggregate, &[&aggregate, &aggregate, &[10, 30]])?;
        let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
        let unknown_archive_info = [0xb5, 0x0c, 1, 2, 3, 4];
        let mut raw_header = canonical.clone();
        raw_header.extend_from_slice(&unknown_archive_info);
        object.set_test_source_headers(
            Some(raw_header.clone().into_boxed_slice()),
            Some(canonical.into_boxed_slice()),
        );
        let source = Archive {
            objects: vec![object.clone()],
        }
        .to_bytes(Default::default())?;
        let replacement = object.messages[0].clone();

        let old = object.replace_message_reordering_data_references_preserving_header(
            0,
            replacement.clone(),
            &aggregate,
            Default::default(),
        )?;

        assert_eq!(old, replacement);
        assert_eq!(
            Archive {
                objects: vec![object.clone()],
            }
            .to_bytes(Default::default())?,
            source
        );
        assert_eq!(object.source_header(), Some(raw_header.as_slice()));
        Ok(())
    }

    #[test]
    fn data_reference_reorder_preserves_unknown_header_bytes() -> Result<()> {
        let aggregate = [11, 22, 33];
        let mut object = data_reorder_object(&aggregate, &[&aggregate, &[99]])?;
        let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
        let unknown_archive_info = [0xb5, 0x0c, 1, 2, 3, 4];
        let mut raw_header = canonical.clone();
        raw_header.extend_from_slice(&unknown_archive_info);
        object.set_test_source_headers(
            Some(raw_header.clone().into_boxed_slice()),
            Some(canonical.into_boxed_slice()),
        );

        object.replace_message_reordering_data_references_preserving_header(
            0,
            RawMessage {
                type_: 2,
                data: vec![0xdd],
            },
            &[33, 11, 22],
            Default::default(),
        )?;
        let encoded = Archive {
            objects: vec![object.clone()],
        }
        .to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(payload, [0xdd]);
        assert!(header.ends_with(&unknown_archive_info));
        assert_eq!(object.source_header(), Some(header));
        assert_eq!(
            object.archive_info.message_infos[0].field_infos[0].data_references,
            [33, 11, 22]
        );
        assert_eq!(
            object.archive_info.message_infos[0].field_infos[1].data_references,
            [99]
        );
        Ok(())
    }

    #[test]
    fn object_reference_reorder_preserves_raw_header_and_nested_metadata() -> Result<()> {
        let fixture = reference_reorder_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let reordered = [200, 50, 50, 10, 50];
        let replacement = RawMessage {
            type_: 9,
            data: vec![0xfa, 0xfb, 0xfc],
        };
        let old = archive.objects[0]
            .replace_message_reordering_object_references_preserving_header(
                0,
                replacement.clone(),
                &[200, 10],
                Default::default(),
            )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);

        let encoded = archive.to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.reordered_header);
        assert_eq!(payload, [0xfa, 0xfb, 0xfc]);
        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, reordered);
        assert_eq!(info.data_references, [77]);
        assert_eq!(info.field_infos[0].object_references, [10, 200]);
        assert_eq!(info.field_infos[0].data_references, [88]);
        assert_eq!(
            archive.objects[0].source_header(),
            Some(fixture.reordered_header.as_slice())
        );

        let replaced = archive.objects[0]
            .replace_message_reordering_object_references_preserving_header(
                0,
                old,
                &[10, 200],
                Default::default(),
            )?;
        assert_eq!(replaced, replacement);
        assert_eq!(archive.to_bytes(Default::default())?, fixture.source);
        assert_eq!(
            archive.objects[0].source_header(),
            Some(fixture.header.as_slice())
        );
        Ok(())
    }

    #[test]
    fn object_reference_reorder_noop_matches_preserving_replacement() -> Result<()> {
        let fixture = reference_reorder_fixture()?;
        let parsed = Archive::parse(&fixture.source, Default::default())?;
        let replacement = RawMessage {
            type_: 9,
            data: vec![0xfa, 0xfb, 0xfc],
        };
        let mut reordered = parsed.objects[0].clone();
        let mut preserving = parsed.objects[0].clone();
        let reordered_old = reordered
            .replace_message_reordering_object_references_preserving_header(
                0,
                replacement.clone(),
                &[10, 200],
                Default::default(),
            )?;
        let preserving_old =
            preserving.replace_message_preserving_header(0, replacement, Default::default())?;
        assert_eq!(reordered_old, preserving_old);
        assert_eq!(reordered, preserving);

        let mut empty = parsed.objects[0].clone();
        let same = empty.messages[0].clone();
        empty.replace_message_reordering_object_references_preserving_header(
            0,
            same,
            &[],
            Default::default(),
        )?;
        assert_eq!(empty, parsed.objects[0]);
        Ok(())
    }

    #[test]
    fn object_reference_reorder_authorization_failures_are_atomic() -> Result<()> {
        let fixture = reference_reorder_fixture()?;
        let archive = Archive::parse(&fixture.source, Default::default())?;
        for selected in [&[200, 200][..], &[999, 10][..], &[50, 10][..]] {
            let mut object = archive.objects[0].clone();
            let before = object.clone();
            let error = object
                .replace_message_reordering_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 9,
                        data: vec![0xfa, 0xfb, 0xfc],
                    },
                    selected,
                    Default::default(),
                )
                .err();
            assert!(error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn object_reference_reorder_limits_and_malformed_fields_are_atomic() -> Result<()> {
        let fixture = reference_reorder_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let before_limit = archive.objects[0].clone();
        let limits = Limits::default().with_metadata_items(1)?;
        let limit_error = archive.objects[0]
            .replace_message_reordering_object_references_preserving_header(
                0,
                RawMessage {
                    type_: 9,
                    data: vec![0xfa, 0xfb, 0xfc],
                },
                &[200, 10],
                limits,
            )
            .err();
        assert!(matches!(
            limit_error,
            Some(Error::Limit {
                kind: LimitKind::MetadataItems,
                ..
            })
        ));
        assert_eq!(archive.objects[0], before_limit);

        for malformed_message in [
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x29, 1, 2, 3, 4, 5, 6, 7, 8,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x2a, 0x01, 0x80,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x28, 0x0a, 0x28, 0x1e,
            ],
        ] {
            let mut object = ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![0xcc],
                }],
                Default::default(),
            )?;
            object.archive_info.message_infos[0].object_references = vec![10, 20];
            let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
            let mut raw_header = vec![0x08, 0x01];
            push_length_delimited(&mut raw_header, &[0x12], 1, &malformed_message)?;
            object.set_test_source_headers(
                Some(raw_header.into_boxed_slice()),
                Some(canonical.into_boxed_slice()),
            );
            let before = object.clone();

            let error = object
                .replace_message_reordering_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 2,
                        data: vec![0xdd],
                    },
                    &[20, 10],
                    Default::default(),
                )
                .err();
            assert!(error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn append_objects_is_atomic_ordered_and_bounded() -> Result<()> {
        let first = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![1],
            }],
            Default::default(),
        )?;
        let second = ArchiveObject::new(
            2,
            vec![RawMessage {
                type_: 3,
                data: vec![2, 2],
            }],
            Default::default(),
        )?;
        let third = ArchiveObject::new(
            3,
            vec![RawMessage {
                type_: 4,
                data: vec![3, 3, 3],
            }],
            Default::default(),
        )?;
        let mut archive = Archive {
            objects: vec![first.clone()],
        };
        archive.append_objects(vec![second.clone(), third.clone()], Default::default())?;
        assert_eq!(
            archive.objects,
            [first.clone(), second.clone(), third.clone()]
        );
        let encoded = archive.to_bytes(Default::default())?;
        assert_eq!(
            Archive::parse(&encoded, Default::default())?.to_bytes(Default::default())?,
            encoded
        );

        let mut empty = archive.clone();
        empty.append_objects(Vec::new(), Default::default())?;
        assert_eq!(empty, archive);

        let exact = Limits::default()
            .with_objects(3)?
            .with_messages(3)?
            .with_archive_bytes(encoded.len())?;
        let mut exact_archive = Archive {
            objects: vec![first.clone()],
        };
        exact_archive.append_objects(vec![second.clone(), third.clone()], exact)?;
        assert_eq!(exact_archive, archive);

        for limits in [
            Limits::default().with_objects(2)?,
            Limits::default().with_messages(2)?,
            Limits::default().with_archive_bytes(encoded.len() - 1)?,
        ] {
            let mut limited = Archive {
                objects: vec![first.clone()],
            };
            let before = limited.clone();
            let error = limited
                .append_objects(vec![second.clone(), third.clone()], limits)
                .err();
            assert!(matches!(error, Some(Error::Limit { .. })));
            assert_eq!(limited, before);
        }
        Ok(())
    }

    #[test]
    fn append_objects_rejects_zero_and_duplicate_identifiers_atomically() -> Result<()> {
        let existing = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![1],
            }],
            Default::default(),
        )?;
        for objects in [
            vec![ArchiveObject::new(0, Vec::new(), Default::default())?],
            vec![ArchiveObject::new(1, Vec::new(), Default::default())?],
            vec![
                ArchiveObject::new(2, Vec::new(), Default::default())?,
                ArchiveObject::new(2, Vec::new(), Default::default())?,
            ],
        ] {
            let mut archive = Archive {
                objects: vec![existing.clone()],
            };
            let before = archive.clone();
            assert!(archive.append_objects(objects, Default::default()).is_err());
            assert_eq!(archive, before);
        }
        Ok(())
    }

    #[test]
    fn exact_reference_transition_preserves_raw_spans_and_changes_all_selected_lists() -> Result<()>
    {
        let fixture = reference_transition_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let old = archive.objects[0]
            .replace_message_transitioning_object_references_preserving_header(
                0,
                RawMessage {
                    type_: 9,
                    data: vec![0xfa, 0xfb, 0xfc, 0xfd],
                },
                ObjectReferenceTransition {
                    aggregate_before: &[10, 20, 30],
                    aggregate_after: &[30, 20, 40],
                    fields: &[FieldObjectReferenceTransition {
                        field_info_index: 0,
                        expected_path: &[4, 1],
                        before: &[10, 20],
                        after: &[20, 40],
                    }],
                },
                Default::default(),
            )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);
        let encoded = archive.to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.transitioned_header);
        assert_eq!(payload, [0xfa, 0xfb, 0xfc, 0xfd]);
        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, [30, 20, 40]);
        assert_eq!(info.field_infos[0].object_references, [20, 40]);
        assert_eq!(info.field_infos[1].object_references, [30]);
        assert_eq!(
            Archive::parse(&encoded, Default::default())?.to_bytes(Default::default())?,
            encoded
        );
        Ok(())
    }

    #[test]
    fn exact_reference_transition_retains_duplicate_unselected_field_members() -> Result<()> {
        let fixture = reference_transition_fixture()?;
        let mut parsed = Archive::parse(&fixture.source, Default::default())?;
        parsed.objects[0].archive_info.message_infos[0].field_infos[1].object_references =
            vec![30, 30];
        let source = parsed.to_bytes(Default::default())?;
        let mut parsed = Archive::parse(&source, Default::default())?;
        parsed.objects[0].replace_message_transitioning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0xfa, 0xfb, 0xfc, 0xfd],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 40],
                fields: &[FieldObjectReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[4, 1],
                    before: &[10, 20],
                    after: &[20, 40],
                }],
            },
            Default::default(),
        )?;
        assert_eq!(
            parsed.objects[0].archive_info.message_infos[0].field_infos[1].object_references,
            [30, 30]
        );
        Ok(())
    }

    #[test]
    fn exact_reference_transition_hostile_authorization_is_atomic() -> Result<()> {
        let fixture = reference_transition_fixture()?;
        let parsed = Archive::parse(&fixture.source, Default::default())?;
        let transitions = [
            ObjectReferenceTransition {
                aggregate_before: &[10, 99, 30],
                aggregate_after: &[30, 99, 40],
                fields: &[],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[40, 30, 20],
                fields: &[],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 20],
                fields: &[],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 0],
                fields: &[],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 40],
                fields: &[FieldObjectReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[4, 9],
                    before: &[10, 20],
                    after: &[20, 40],
                }],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 40],
                fields: &[FieldObjectReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[4, 1],
                    before: &[10, 20],
                    after: &[20, 99],
                }],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 40],
                fields: &[],
            },
        ];
        for transition in transitions {
            let mut object = parsed.objects[0].clone();
            let before = object.clone();
            let error = object
                .replace_message_transitioning_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 9,
                        data: vec![0xfa, 0xfb, 0xfc, 0xfd],
                    },
                    transition,
                    Default::default(),
                )
                .err();
            assert!(error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn exact_reference_transition_limits_are_inclusive_and_atomic() -> Result<()> {
        let fixture = reference_transition_fixture()?;
        let parsed = Archive::parse(&fixture.source, Default::default())?;
        let transition = ObjectReferenceTransition {
            aggregate_before: &[10, 20, 30],
            aggregate_after: &[30, 20, 40],
            fields: &[FieldObjectReferenceTransition {
                field_info_index: 0,
                expected_path: &[4, 1],
                before: &[10, 20],
                after: &[20, 40],
            }],
        };
        let mut successful = parsed.objects[0].clone();
        successful.replace_message_transitioning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0xfa, 0xfb, 0xfc, 0xfd],
            },
            transition,
            Default::default(),
        )?;
        let successful_bytes = Archive {
            objects: vec![successful.clone()],
        }
        .to_bytes(Default::default())?;
        let (header, _) = split_test_archive(&successful_bytes)?;
        let metadata_items = successful.archive_info.message_infos.len()
            + successful
                .archive_info
                .message_infos
                .iter()
                .try_fold(0usize, |count, info| {
                    count
                        .checked_add(info.metadata_item_count()?)
                        .ok_or_else(|| Error::invalid_archive(0, "test metadata count overflow"))
                })?;
        let authorization_items = 3 + 3 + 1 + 2 + 2 + 2;
        let metadata_limit = metadata_items.max(authorization_items);
        let object_bytes = varint_len(header.len())?
            .checked_add(header.len())
            .and_then(|length| length.checked_add(4))
            .ok_or_else(|| Error::invalid_archive(0, "test object length overflow"))?;
        let exact_limits = Limits::default()
            .with_metadata_items(metadata_limit)?
            .with_header_bytes(header.len())?
            .with_object_bytes(object_bytes)?;
        let mut exact = parsed.objects[0].clone();
        exact.replace_message_transitioning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0xfa, 0xfb, 0xfc, 0xfd],
            },
            transition,
            exact_limits,
        )?;
        for limits in [
            Limits::default().with_metadata_items(metadata_limit - 1)?,
            Limits::default().with_header_bytes(header.len() - 1)?,
            Limits::default().with_object_bytes(object_bytes - 1)?,
        ] {
            let mut object = parsed.objects[0].clone();
            let before = object.clone();
            let error = object
                .replace_message_transitioning_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 9,
                        data: vec![0xfa, 0xfb, 0xfc, 0xfd],
                    },
                    transition,
                    limits,
                )
                .err();
            assert!(matches!(error, Some(Error::Limit { .. })));
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn exact_reference_transition_rejects_ambiguous_raw_reference_atomically() -> Result<()> {
        let fixture = reference_transition_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let object = &mut archive.objects[0];
        let mut malformed = object
            .source_header()
            .ok_or_else(|| Error::invalid_archive(0, "test raw header was not retained"))?
            .to_vec();
        let key = malformed
            .windows(3)
            .position(|window| window == [0x28, 0x8a, 0x00])
            .ok_or_else(|| Error::invalid_archive(0, "test reference key is missing"))?;
        malformed[key] = 0x29;
        let canonical = object
            .source_canonical_header()
            .map(|header| header.to_vec().into_boxed_slice());
        object.set_test_source_headers(Some(malformed.into_boxed_slice()), canonical);
        let before = object.clone();
        let error = object
            .replace_message_transitioning_object_references_preserving_header(
                0,
                RawMessage {
                    type_: 9,
                    data: vec![0xfa, 0xfb, 0xfc, 0xfd],
                },
                ObjectReferenceTransition {
                    aggregate_before: &[10, 20, 30],
                    aggregate_after: &[30, 20, 40],
                    fields: &[FieldObjectReferenceTransition {
                        field_info_index: 0,
                        expected_path: &[4, 1],
                        before: &[10, 20],
                        after: &[20, 40],
                    }],
                },
                Default::default(),
            )
            .err();
        assert!(error.is_some());
        assert_eq!(object, &before);
        Ok(())
    }

    fn data_reorder_object(
        aggregate: &[u64],
        field_data_references: &[&[u64]],
    ) -> Result<ArchiveObject> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![0xcc],
            }],
            Default::default(),
        )?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or_else(|| Error::invalid_archive(0, "test object is missing message metadata"))?;
        info.data_references.extend_from_slice(aggregate);
        for (index, references) in field_data_references.iter().enumerate() {
            let path = u32::try_from(index)
                .map_err(|_| Error::invalid_archive(0, "test FieldInfo path exceeds u32"))?;
            let mut field = FieldInfo::new(
                FieldPath::new(vec![path], Limits::default()).unwrap(),
                Limits::default(),
            )
            .unwrap();
            field.data_references.extend_from_slice(references);
            info.field_infos.push(field);
        }
        object.validate(Default::default())?;
        Ok(object)
    }

    fn reference_reorder_fixture() -> Result<ReferenceReorderFixture> {
        let nested_field = reference_reorder_nested_field()?;
        let message = reference_reorder_message(7, false, &nested_field)?;
        let reordered_message = reference_reorder_message(9, true, &nested_field)?;
        let header = reference_reorder_header(&message)?;
        let reordered_header = reference_reorder_header(&reordered_message)?;

        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(ReferenceReorderFixture {
            source,
            header,
            reordered_header,
        })
    }

    #[test]
    fn canonical_object_fields_preserve_unknown_header_and_remove_exact_suffix() -> Result<()> {
        let source = reference_transition_fixture()?.source;
        let mut archive = Archive::parse(&source, Default::default())?;
        let prior = archive.objects[0].archive_info.clone();
        let before = [10, 20, 30];
        let after = [10, 20, 30, 40, 50, 60, 70];
        let mut aggregate_only = archive.clone();
        let original_payload = aggregate_only.objects[0]
            .replace_message_transitioning_object_references_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: vec![1, 2, 3, 4],
                },
                ObjectReferenceTransition {
                    aggregate_before: &before,
                    aggregate_after: &after,
                    fields: &[],
                },
                Default::default(),
            )?;
        aggregate_only.objects[0]
            .replace_message_transitioning_object_references_preserving_header(
                0,
                original_payload,
                ObjectReferenceTransition {
                    aggregate_before: &after,
                    aggregate_after: &before,
                    fields: &[],
                },
                Default::default(),
            )?;
        let fields = [
            CanonicalObjectReferenceField {
                path: &[34],
                references: &[40],
            },
            CanonicalObjectReferenceField {
                path: &[35],
                references: &[50],
            },
        ];
        let old = archive.objects[0]
            .replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 7, data: vec![1, 2, 3, 4] }, ObjectReferenceTransition { aggregate_before: &before, aggregate_after: &after, fields: &[] }, CanonicalObjectReferenceFields::Insert(&fields), Default::default())?;
        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, after);
        assert_eq!(
            info.field_infos.len(),
            prior.message_infos[0].field_infos.len() + 2
        );
        assert_eq!(
            &info.field_infos[..2],
            prior.message_infos[0].field_infos.as_slice()
        );
        for (actual, expected) in info.field_infos[2..].iter().zip(fields) {
            assert!(super::canonical_object_field_matches(actual, expected));
        }
        // Reopen to ensure the inverse relies on persisted bytes, not a
        // remembered insertion action or an in-memory metadata cache.
        let mut reopened =
            Archive::parse(&archive.to_bytes(Default::default())?, Default::default())?;
        reopened.objects[0]
            .replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, old, ObjectReferenceTransition { aggregate_before: &after, aggregate_after: &before, fields: &[] }, CanonicalObjectReferenceFields::Remove(&fields), Default::default())?;
        assert_eq!(reopened.objects[0].archive_info, prior);
        // Ordinary aggregate transitions retain empty packed occurrences.
        // Adding and removing attribution must introduce no other raw delta.
        assert_eq!(
            reopened.to_bytes(Default::default())?,
            aggregate_only.to_bytes(Default::default())?
        );
        Ok(())
    }

    #[test]
    fn canonical_object_fields_reject_ambiguous_authorization_atomically() -> Result<()> {
        let parsed = Archive::parse(&reference_transition_fixture()?.source, Default::default())?;
        let duplicate_paths = [
            CanonicalObjectReferenceField {
                path: &[34],
                references: &[40],
            },
            CanonicalObjectReferenceField {
                path: &[34],
                references: &[50],
            },
        ];
        let bad_definitions: &[&[CanonicalObjectReferenceField<'_>]] = &[
            &duplicate_paths,
            &[CanonicalObjectReferenceField {
                path: &[4, 1],
                references: &[40],
            }],
            &[CanonicalObjectReferenceField {
                path: &[34],
                references: &[40, 40],
            }],
            &[CanonicalObjectReferenceField {
                path: &[34],
                references: &[99],
            }],
            &[CanonicalObjectReferenceField {
                path: &[34],
                references: &[0],
            }],
            &[CanonicalObjectReferenceField {
                path: &[34],
                references: &[],
            }],
        ];
        for definitions in bad_definitions {
            let mut object = parsed.objects[0].clone();
            let before = object.clone();
            assert!(object.replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 7, data: vec![1, 2, 3] }, ObjectReferenceTransition { aggregate_before: &[10, 20, 30], aggregate_after: &[10, 20, 30, 40, 50], fields: &[] }, CanonicalObjectReferenceFields::Insert(definitions), Default::default()).is_err());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn canonical_object_fields_admit_exact_metadata_and_reject_one_under() -> Result<()> {
        let parsed = Archive::parse(&reference_transition_fixture()?.source, Default::default())?;
        let fields = [CanonicalObjectReferenceField {
            path: &[34],
            references: &[40],
        }];
        let transition = ObjectReferenceTransition {
            aggregate_before: &[10, 20, 30],
            aggregate_after: &[10, 20, 30, 40],
            fields: &[],
        };
        let mut expected = parsed.objects[0].clone();
        expected.replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 7, data: vec![1, 2, 3] }, transition, CanonicalObjectReferenceFields::Insert(&fields), Default::default())?;
        let metadata_items = 1 + expected.archive_info.message_infos[0].metadata_item_count()?;
        for (maximum, success) in [(metadata_items, true), (metadata_items - 1, false)] {
            let mut object = parsed.objects[0].clone();
            let before = object.clone();
            let result = object.replace_message_transitioning_object_references_with_canonical_fields_preserving_header(
                0, RawMessage { type_: 7, data: vec![1, 2, 3] }, transition,
                CanonicalObjectReferenceFields::Insert(&fields), Limits::default().with_metadata_items(maximum)?,
            );
            if success {
                result?;
                assert_eq!(object, expected);
            } else {
                assert!(
                    matches!(result, Err(Error::Limit { kind: LimitKind::MetadataItems, observed, maximum: admitted }) if observed == metadata_items && admitted == maximum)
                );
                assert_eq!(object, before);
            }
        }
        Ok(())
    }

    #[test]
    fn canonical_object_fields_compose_with_selected_reference_transition() -> Result<()> {
        let fixture = reference_transition_fixture()?;
        let mut archive = Archive::parse(&fixture.source, Default::default())?;
        let definition = [CanonicalObjectReferenceField {
            path: &[34],
            references: &[50],
        }];
        archive.objects[0].replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 9, data: vec![1, 2, 3, 4] }, ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30], aggregate_after: &[30, 20, 40, 50],
                fields: &[FieldObjectReferenceTransition { field_info_index: 0, expected_path: &[4, 1], before: &[10, 20], after: &[20, 40] }],
            }, CanonicalObjectReferenceFields::Insert(&definition), Default::default())?;
        let mut expected = Archive::parse(&fixture.source, Default::default())?;
        expected.objects[0].replace_message_transitioning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![1, 2, 3, 4],
            },
            ObjectReferenceTransition {
                aggregate_before: &[10, 20, 30],
                aggregate_after: &[30, 20, 40, 50],
                fields: &[FieldObjectReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[4, 1],
                    before: &[10, 20],
                    after: &[20, 40],
                }],
            },
            Default::default(),
        )?;
        // Removing attribution without changing the aggregate proves that the
        // selected raw field (including its unknown bytes) matches the existing
        // independent transition path exactly.
        archive.objects[0].replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 9, data: vec![1, 2, 3, 4] }, ObjectReferenceTransition { aggregate_before: &[30, 20, 40, 50], aggregate_after: &[30, 20, 40, 50], fields: &[] }, CanonicalObjectReferenceFields::Remove(&definition), Default::default())?;
        assert_eq!(
            archive.to_bytes(Default::default())?,
            expected.to_bytes(Default::default())?
        );
        Ok(())
    }

    #[test]
    fn canonical_object_fields_bound_output_header_atomically() -> Result<()> {
        let archive = Archive::parse(&reference_transition_fixture()?.source, Default::default())?;
        let definition = [CanonicalObjectReferenceField {
            path: &[34],
            references: &[40],
        }];
        let transition = ObjectReferenceTransition {
            aggregate_before: &[10, 20, 30],
            aggregate_after: &[10, 20, 30, 40],
            fields: &[],
        };
        let mut expected = archive.objects[0].clone();
        expected.replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 7, data: vec![1, 2, 3] }, transition, CanonicalObjectReferenceFields::Insert(&definition), Default::default())?;
        let header_bytes = expected
            .source_header()
            .ok_or_else(|| Error::invalid_archive(0, "test missing raw header"))?
            .len();
        for maximum in [header_bytes, header_bytes - 1] {
            let mut candidate = archive.objects[0].clone();
            let original = candidate.clone();
            let result = candidate.replace_message_transitioning_object_references_with_canonical_fields_preserving_header(
                0, RawMessage { type_: 7, data: vec![1, 2, 3] }, transition, CanonicalObjectReferenceFields::Insert(&definition),
                Limits::default().with_header_bytes(maximum)?,
            );
            if maximum == header_bytes {
                result?;
                assert_eq!(candidate, expected);
            } else {
                assert!(matches!(
                    result,
                    Err(Error::Limit {
                        kind: LimitKind::HeaderBytes,
                        ..
                    })
                ));
                assert_eq!(candidate, original);
            }
        }
        Ok(())
    }

    #[test]
    fn canonical_object_field_removal_refuses_noncanonical_wire() -> Result<()> {
        let definitions = [CanonicalObjectReferenceField {
            path: &[34],
            references: &[40],
        }];
        let canonical = super::encode_canonical_object_reference_field_info(
            definitions[0],
            Limits::default(),
            0,
        )?;
        let mut message = vec![0x08, 7, 0x18, 3, 0x2a, 1, 40];
        // Same decoded FieldInfo, but its outer key has an overlong encoding.
        message.extend_from_slice(&[0xa2, 0]);
        message.extend_from_slice(&canonical[1..]);
        let mut header = vec![0x08, 1];
        push_length_delimited(&mut header, &[0x12], 1, &message)?;
        let mut source = Vec::new();
        push_test_varint(&mut source, header.len() as u64, 1);
        source.extend_from_slice(&header);
        source.extend_from_slice(&[1, 2, 3]);
        let mut parsed = Archive::parse(&source, Default::default())?;
        let original = parsed.clone();
        assert!(parsed.objects[0].replace_message_transitioning_object_references_with_canonical_fields_preserving_header(0, RawMessage { type_: 7, data: vec![4, 5, 6] }, ObjectReferenceTransition { aggregate_before: &[40], aggregate_after: &[], fields: &[] }, CanonicalObjectReferenceFields::Remove(&definitions), Default::default()).is_err());
        assert_eq!(parsed, original);
        assert_eq!(parsed.to_bytes(Default::default())?, source);
        Ok(())
    }

    fn reference_transition_fixture() -> Result<ReferenceTransitionFixture> {
        let source_first_field = reference_transition_field(&[4, 1], &[10, 20], false)?;
        let target_first_field = reference_transition_field(&[4, 1], &[20, 40], true)?;
        let second_field = reference_transition_field(&[4, 2], &[30], false)?;
        let source_message = reference_transition_message(
            7,
            3,
            &[10, 20, 30],
            &source_first_field,
            &second_field,
            false,
        )?;
        let target_message = reference_transition_message(
            9,
            4,
            &[30, 20, 40],
            &target_first_field,
            &second_field,
            true,
        )?;
        let header = reference_transition_header(1, &source_message)?;
        let transitioned_header = reference_transition_header(1, &target_message)?;
        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(ReferenceTransitionFixture {
            source,
            transitioned_header,
        })
    }

    fn reference_transition_field(
        path: &[u8],
        references: &[u64],
        transitioned: bool,
    ) -> Result<Vec<u8>> {
        let mut path_message = Vec::new();
        push_length_delimited(&mut path_message, &[0x0a], 2, path)?;
        let mut field = vec![0xa0, 0x01, 0x81, 0x00];
        push_length_delimited(&mut field, &[0x0a], 2, &path_message)?;
        if !transitioned {
            let mut packed = Vec::new();
            for reference in references {
                push_test_varint(&mut packed, *reference, usize::from(*reference == 10) + 1);
            }
            push_length_delimited(&mut field, &[0x22], 2, &packed)?;
        } else {
            push_length_delimited(&mut field, &[0x22], 2, &[0x14])?;
        }
        field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);
        if transitioned {
            push_length_delimited(&mut field, &[0x22], 1, &[0x28])?;
        }
        Ok(field)
    }

    fn reference_transition_message(
        type_: u64,
        length: u64,
        aggregate: &[u64],
        first_field: &[u8],
        second_field: &[u8],
        transitioned: bool,
    ) -> Result<Vec<u8>> {
        let mut message = vec![0xcd, 0x0c, 9, 8, 7, 6];
        message.push(0x08);
        push_test_varint(&mut message, type_, 2);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18]);
        push_test_varint(&mut message, length, 2);
        if transitioned {
            push_length_delimited(&mut message, &[0x2a], 2, &[0x1e, 0x14])?;
        } else {
            message.extend_from_slice(&[0x28, 0x8a, 0x00]);
            let mut packed = Vec::new();
            for reference in aggregate.iter().skip(1) {
                push_test_varint(&mut packed, *reference, 1);
            }
            push_length_delimited(&mut message, &[0x2a], 2, &packed)?;
        }
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, first_field)?;
        push_length_delimited(&mut message, &[0x22], 2, second_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;
        if transitioned {
            push_length_delimited(&mut message, &[0x2a], 1, &[0x28])?;
        }
        Ok(message)
    }

    fn reference_transition_header(identifier: u64, message: &[u8]) -> Result<Vec<u8>> {
        let mut header = vec![0x08];
        push_test_varint(&mut header, identifier, 2);
        header.extend_from_slice(&[0xd5, 0x0c, 1, 2, 3, 4]);
        push_length_delimited(&mut header, &[0x92, 0x00], 2, message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);
        Ok(header)
    }

    fn reference_reorder_nested_field() -> Result<Vec<u8>> {
        let mut field = vec![0xa0, 0x01, 0x81, 0x00];
        push_length_delimited(&mut field, &[0x0a], 2, &[])?;
        field.extend_from_slice(&[0x20, 0x8a, 0x00]);
        push_length_delimited(&mut field, &[0x22], 2, &[0xc8, 0x01])?;
        field.extend_from_slice(&[0x28, 0x58]);
        field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);
        Ok(field)
    }

    fn reference_reorder_message(
        type_: u64,
        reordered: bool,
        nested_field: &[u8],
    ) -> Result<Vec<u8>> {
        let mut message = Vec::new();
        message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00]);
        push_test_varint(&mut message, type_, 2);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00, 0x83, 0x00]);
        if reordered {
            message.extend_from_slice(&[0x28, 0xc8, 0x01]);
        } else {
            message.extend_from_slice(&[0x28, 0x8a, 0x00]);
        }
        message.extend_from_slice(&[0xa8, 0x00, 0x32]);
        if reordered {
            push_length_delimited(&mut message, &[0x2a], 2, &[0xb2, 0x00, 0x8a, 0x00, 0x32])?;
        } else {
            push_length_delimited(&mut message, &[0x2a], 2, &[0xb2, 0x00, 0xc8, 0x01, 0x32])?;
        }
        message.extend_from_slice(&[0x30, 0x4d]);
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, nested_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;
        Ok(message)
    }

    fn reference_reorder_header(message: &[u8]) -> Result<Vec<u8>> {
        let mut header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut header, &[0x92, 0x00], 2, message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);
        Ok(header)
    }

    fn combined_reference_pruning_fixture() -> Result<CombinedReferencePruningFixture> {
        let source_message = combined_reference_message(7, 3, false, 0)?;
        let selected_message = combined_reference_message(9, 130, true, 1)?;
        let all_data_message = combined_reference_message(9, 130, false, 2)?;
        let source_header = combined_reference_header(&source_message)?;
        let selected_header = combined_reference_header(&selected_message)?;
        let all_data_header = combined_reference_header(&all_data_message)?;

        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(source_header.len() as u64, &mut prefix));
        source.extend_from_slice(&source_header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(CombinedReferencePruningFixture {
            source,
            selected_header,
            all_data_header,
        })
    }

    fn combined_reference_header(message: &[u8]) -> Result<Vec<u8>> {
        let mut header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut header, &[0x92, 0x00], 2, message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);
        Ok(header)
    }

    // data_mode: 0 retains all data references, 1 prunes 20 and 22, and 2
    // prunes all data references. Conditional construction keeps the exact
    // expected raw bytes visible independently of the production rewriter.
    fn combined_reference_message(
        type_: u64,
        length: u64,
        prune_objects: bool,
        data_mode: u8,
    ) -> Result<Vec<u8>> {
        let first_field = combined_first_field(prune_objects, data_mode)?;
        let second_field = combined_second_field(data_mode)?;
        let mut message = Vec::new();
        message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00]);
        push_test_varint(&mut message, type_, 2);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00]);
        push_test_varint(&mut message, length, 2);

        if prune_objects {
            message.extend_from_slice(&[0x28, 0x0d, 0xa8, 0x00, 0x8d, 0x00]);
            push_length_delimited(&mut message, &[0x2a], 2, &[0x0d])?;
        } else {
            message.extend_from_slice(&[
                0x28, 0x8a, 0x00, 0xa8, 0x00, 0x8a, 0x00, 0x28, 0x0d, 0xa8, 0x00, 0x8d, 0x00,
            ]);
            push_length_delimited(&mut message, &[0x2a], 2, &[0x0c, 0x0d, 0x0a])?;
        }
        push_length_delimited(&mut message, &[0x2a], 2, &[])?;

        match data_mode {
            0 => {
                message.extend_from_slice(&[
                    0x30, 0x94, 0x00, 0xb0, 0x00, 0x94, 0x00, 0x30, 0x15, 0xb0, 0x00, 0x95, 0x00,
                ]);
                push_length_delimited(&mut message, &[0x32], 2, &[0x16, 0x15, 0x14, 0x15])?;
            },
            1 => {
                message.extend_from_slice(&[0x30, 0x15, 0xb0, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut message, &[0x32], 2, &[0x15, 0x15])?;
            },
            2 => push_length_delimited(&mut message, &[0x32], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        push_length_delimited(&mut message, &[0x32], 2, &[])?;
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, &first_field)?;
        push_length_delimited(&mut message, &[0x22], 2, &second_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;
        Ok(message)
    }

    fn combined_first_field(prune_objects: bool, data_mode: u8) -> Result<Vec<u8>> {
        let mut field = vec![0xa0, 0x01, 0x81, 0x00];
        push_length_delimited(&mut field, &[0x0a], 2, &[])?;
        if prune_objects {
            push_length_delimited(&mut field, &[0x22], 2, &[0xe4, 0x00])?;
        } else {
            field.extend_from_slice(&[0x20, 0x8a, 0x00]);
            push_length_delimited(&mut field, &[0x22], 2, &[0x0c, 0xe4, 0x00, 0x0a])?;
        }
        match data_mode {
            0 => {
                field.extend_from_slice(&[0x28, 0x94, 0x00, 0xa8, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut field, &[0x2a], 2, &[0x14, 0x96, 0x00, 0x15, 0x14])?;
            },
            1 => {
                field.extend_from_slice(&[0xa8, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut field, &[0x2a], 2, &[0x15])?;
            },
            2 => push_length_delimited(&mut field, &[0x2a], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        push_length_delimited(&mut field, &[0x2a], 2, &[])?;
        push_length_delimited(&mut field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;
        Ok(field)
    }

    fn combined_second_field(data_mode: u8) -> Result<Vec<u8>> {
        let mut field = Vec::new();
        push_length_delimited(&mut field, &[0x0a], 1, &[])?;
        field.extend_from_slice(&[0x20, 0x4d]);
        match data_mode {
            0 | 1 => {
                field.extend_from_slice(&[0x28, 0x1f]);
                push_length_delimited(&mut field, &[0xaa, 0x00], 2, &[0x20, 0x21])?;
            },
            2 => push_length_delimited(&mut field, &[0xaa, 0x00], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);
        Ok(field)
    }

    fn push_test_varint(output: &mut Vec<u8>, value: u64, width: usize) {
        let mut encoded = [0u8; 10];
        output.extend_from_slice(encode_varint_with_width(value, width, &mut encoded));
    }

    #[test]
    fn changed_cardinality_data_transition_supports_insert_remove_and_replace() -> Result<()> {
        let source = data_reference_transition_fixture(true)?;
        let original = Archive::parse(&source, Default::default())?;

        for (after, payload) in [
            (vec![10, 30, 20], data_reference_payload(&[10, 30, 20])?),
            (vec![20], data_reference_payload(&[20])?),
            (vec![40, 20], data_reference_payload(&[40, 20])?),
        ] {
            let mut edited = original.clone();
            let before = [10, 20];
            let field_before = [10, 20];
            let field_after = after.clone();
            let transition = DataReferenceTransition {
                aggregate_before: &before,
                aggregate_after: &after,
                fields: &[FieldDataReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[3],
                    before: &field_before,
                    after: &field_after,
                }],
            };
            edited.objects[0].replace_message_transitioning_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: payload,
                },
                transition,
                Default::default(),
            )?;
            assert_eq!(
                edited.objects[0].archive_info.message_infos[0].data_references,
                after
            );
            assert_eq!(
                edited.objects[0].archive_info.message_infos[0].field_infos[0].data_references,
                after
            );
            let encoded = edited.to_bytes(Default::default())?;
            let (header, _) = split_test_archive(&encoded)?;
            assert!(header.ends_with(&[0xb0, 0x3e, 0x07]));

            // The source header's unknown suffix is retained, while the
            // archive identifier prefix remains source-exact.
            assert_eq!(&header[..2], &[0x08, 0x2a]);
        }
        Ok(())
    }

    #[test]
    fn changed_cardinality_data_transition_inverse_restores_exact_source() -> Result<()> {
        let source = data_reference_transition_fixture(true)?;
        let mut edited = Archive::parse(&source, Default::default())?;
        let before = [10, 20];
        let after = [10, 30, 20];
        let transition = DataReferenceTransition {
            aggregate_before: &before,
            aggregate_after: &after,
            fields: &[FieldDataReferenceTransition {
                field_info_index: 0,
                expected_path: &[3],
                before: &before,
                after: &after,
            }],
        };
        let old = edited.objects[0]
            .replace_message_transitioning_data_references_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&after)?,
                },
                transition,
                Default::default(),
            )?;
        let inverse = DataReferenceTransition {
            aggregate_before: &after,
            aggregate_after: &before,
            fields: &[FieldDataReferenceTransition {
                field_info_index: 0,
                expected_path: &[3],
                before: &after,
                after: &before,
            }],
        };
        edited.objects[0].replace_message_transitioning_data_references_preserving_header(
            0,
            old,
            inverse,
            Default::default(),
        )?;
        assert_eq!(edited.to_bytes(Default::default())?, source);
        Ok(())
    }

    #[test]
    fn changed_cardinality_data_transition_rejects_stale_or_limited_requests_atomically()
    -> Result<()> {
        let source = data_reference_transition_fixture(false)?;
        let mut edited = Archive::parse(&source, Default::default())?;
        let original = edited.clone();
        let wrong_before = [99, 20];
        let after = [10, 30, 20];
        let stale = DataReferenceTransition {
            aggregate_before: &wrong_before,
            aggregate_after: &after,
            fields: &[FieldDataReferenceTransition {
                field_info_index: 0,
                expected_path: &[3],
                before: &wrong_before,
                after: &after,
            }],
        };
        assert!(
            edited.objects[0]
                .replace_message_transitioning_data_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 7,
                        data: data_reference_payload(&after)?,
                    },
                    stale,
                    Default::default(),
                )
                .is_err()
        );
        assert_eq!(edited, original);

        let before = [10, 20];
        let limited = DataReferenceTransition {
            aggregate_before: &before,
            aggregate_after: &after,
            fields: &[FieldDataReferenceTransition {
                field_info_index: 0,
                expected_path: &[3],
                before: &before,
                after: &after,
            }],
        };
        let limits = Limits::default().with_metadata_items(1)?;
        assert!(
            edited.objects[0]
                .replace_message_transitioning_data_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 7,
                        data: data_reference_payload(&after)?,
                    },
                    limited,
                    limits,
                )
                .is_err()
        );
        assert_eq!(edited, original);

        // A malformed retained raw header is rejected before the object is
        // touched. This exercises raw-wire admission independently from the
        // stale authorization check above.
        let mut malformed = original.clone();
        let malformed_canonical =
            encode_archive_info(&malformed.objects[0].archive_info, Limits::default())?;
        malformed.objects[0].set_test_source_headers(
            Some(Box::from([0x30, 0x80])),
            Some(malformed_canonical.into_boxed_slice()),
        );
        let malformed_before = malformed.clone();
        assert!(
            malformed.objects[0]
                .replace_message_transitioning_data_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 7,
                        data: data_reference_payload(&after)?,
                    },
                    limited,
                    Default::default(),
                )
                .is_err()
        );
        assert_eq!(malformed, malformed_before);
        Ok(())
    }

    #[test]
    fn canonical_field_transition_inserts_and_removes_exactly_with_unknown_header() -> Result<()> {
        let source = aggregate_only_data_reference_transition_fixture(true)?;
        let mut edited = Archive::parse(&source, Default::default())?;
        let before = [10, 20];
        let after = [10, 30, 20];
        let inserted = CanonicalFieldDataReferenceTransition {
            aggregate_before: &before,
            aggregate_after: &after,
            field_path: &[3],
            operation: CanonicalFieldDataReferenceOperation::Insert,
        };
        let old = edited.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&after)?,
                },
                inserted,
                Default::default(),
            )?;
        let info = &edited.objects[0].archive_info.message_infos[0];
        assert_eq!(info.data_references, after);
        assert_eq!(info.field_infos.len(), 1);
        assert_eq!(info.field_infos[0].path.as_slice(), [3]);
        assert_eq!(info.field_infos[0].r#type, Some(FieldType::DataReference));
        assert_eq!(info.field_infos[0].data_references, after);
        let encoded = edited.to_bytes(Default::default())?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert!(header.ends_with(&[0xb0, 0x3e, 0x07]));
        assert_eq!(payload, data_reference_payload(&after)?);

        let inverse = CanonicalFieldDataReferenceTransition {
            aggregate_before: &after,
            aggregate_after: &before,
            field_path: &[3],
            operation: CanonicalFieldDataReferenceOperation::Remove,
        };
        edited.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                old,
                inverse,
                Default::default(),
            )?;
        assert_eq!(edited.to_bytes(Default::default())?, source);
        Ok(())
    }

    #[test]
    fn canonical_field_transition_rejects_stale_and_existing_paths_atomically() -> Result<()> {
        let source = aggregate_only_data_reference_transition_fixture(true)?;
        let mut stale = Archive::parse(&source, Default::default())?;
        let original = stale.clone();
        let stale_before = [99, 20];
        let after = [99, 20, 30];
        let request = CanonicalFieldDataReferenceTransition {
            aggregate_before: &stale_before,
            aggregate_after: &after,
            field_path: &[3],
            operation: CanonicalFieldDataReferenceOperation::Insert,
        };
        assert!(stale.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&after)?,
                },
                request,
                Default::default(),
            )
            .is_err());
        assert_eq!(stale, original);

        let before = [10, 20];
        let inserted_after = [10, 30, 20];
        let mut existing = Archive::parse(&source, Default::default())?;
        existing.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&inserted_after)?,
                },
                CanonicalFieldDataReferenceTransition {
                    aggregate_before: &before,
                    aggregate_after: &inserted_after,
                    field_path: &[3],
                    operation: CanonicalFieldDataReferenceOperation::Insert,
                },
                Default::default(),
            )?;
        let original = existing.clone();
        let next = [10, 30, 20, 40];
        assert!(existing.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&next)?,
                },
                CanonicalFieldDataReferenceTransition {
                    aggregate_before: &inserted_after,
                    aggregate_after: &next,
                    field_path: &[3],
                    operation: CanonicalFieldDataReferenceOperation::Insert,
                },
                Default::default(),
            )
            .is_err());
        assert_eq!(existing, original);
        Ok(())
    }

    #[test]
    fn canonical_field_transition_rejects_noncanonical_remove_and_limits_atomically() -> Result<()>
    {
        let source = aggregate_only_data_reference_transition_fixture(true)?;
        let mut noncanonical = Archive::parse(&source, Default::default())?;
        let before = [10, 20];
        let after = [10, 30, 20];
        noncanonical.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&after)?,
                },
                CanonicalFieldDataReferenceTransition {
                    aggregate_before: &before,
                    aggregate_after: &after,
                    field_path: &[3],
                    operation: CanonicalFieldDataReferenceOperation::Insert,
                },
                Default::default(),
            )?;
        let raw = noncanonical.objects[0]
            .source_header()
            .ok_or_else(|| Error::invalid_archive(0, "canonical FieldInfo source is missing"))?;
        let expected =
            super::encode_canonical_data_reference_field_info(&[3], &after, Limits::default(), 0)?;
        let mut noncanonical_field = vec![0x10, 0x02];
        let path_message = [0x0a, 0x01, 0x03];
        push_length_delimited(&mut noncanonical_field, &[0x0a], 1, &path_message)?;
        push_length_delimited(&mut noncanonical_field, &[0x2a], 1, &[0x0a, 0x1e, 0x14])?;
        let mut noncanonical_record = Vec::new();
        push_length_delimited(&mut noncanonical_record, &[0x22], 1, &noncanonical_field)?;
        assert_eq!(noncanonical_record.len(), expected.len());
        let key = raw
            .windows(expected.len())
            .position(|window| window == expected.as_slice())
            .ok_or_else(|| Error::invalid_archive(0, "canonical FieldInfo is missing"))?;
        let mut altered = raw.to_vec();
        altered[key..key + expected.len()].copy_from_slice(&noncanonical_record);
        let canonical = noncanonical.objects[0]
            .source_canonical_header()
            .map(|header| header.to_vec().into_boxed_slice());
        noncanonical.objects[0]
            .set_test_source_headers(Some(altered.into_boxed_slice()), canonical);
        let original = noncanonical.clone();
        assert!(noncanonical.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&before)?,
                },
                CanonicalFieldDataReferenceTransition {
                    aggregate_before: &after,
                    aggregate_after: &before,
                    field_path: &[3],
                    operation: CanonicalFieldDataReferenceOperation::Remove,
                },
                Default::default(),
            )
            .is_err());
        assert_eq!(noncanonical, original);

        let mut limited = Archive::parse(&source, Default::default())?;
        let original = limited.clone();
        let limits = Limits::default().with_metadata_items(10)?;
        assert!(limited.objects[0]
            .replace_message_transitioning_data_references_with_canonical_field_preserving_header(
                0,
                RawMessage {
                    type_: 7,
                    data: data_reference_payload(&after)?,
                },
                CanonicalFieldDataReferenceTransition {
                    aggregate_before: &before,
                    aggregate_after: &after,
                    field_path: &[3],
                    operation: CanonicalFieldDataReferenceOperation::Insert,
                },
                limits,
            )
            .is_err());
        assert_eq!(limited, original);
        Ok(())
    }

    fn aggregate_only_data_reference_transition_fixture(
        with_unknown_header: bool,
    ) -> Result<Vec<u8>> {
        let before_payload = data_reference_payload(&[10, 20])?;
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: before_payload.clone(),
            }],
            Default::default(),
        )?;
        object.archive_info.message_infos[0].data_references = vec![10, 20];
        if !with_unknown_header {
            return Archive {
                objects: vec![object],
            }
            .to_bytes(Default::default());
        }

        let mut header = encode_archive_info(&object.archive_info, Limits::default())?;
        header.extend_from_slice(&[0xb0, 0x3e, 0x07]);
        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&before_payload);
        Ok(source)
    }

    fn data_reference_transition_fixture(with_unknown_header: bool) -> Result<Vec<u8>> {
        let before_payload = data_reference_payload(&[10, 20])?;
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: before_payload.clone(),
            }],
            Default::default(),
        )?;
        let info = &mut object.archive_info.message_infos[0];
        info.data_references = vec![10, 20];
        let mut field = FieldInfo::new(
            FieldPath::new(vec![3], Limits::default()).unwrap(),
            Limits::default(),
        )
        .unwrap();
        field.r#type = Some(FieldType::DataReference);
        field.data_references = vec![10, 20];
        info.field_infos.push(field);
        if !with_unknown_header {
            return Archive {
                objects: vec![object],
            }
            .to_bytes(Default::default());
        }

        let mut header = encode_archive_info(&object.archive_info, Limits::default())?;
        header.extend_from_slice(&[0xb0, 0x3e, 0x07]);
        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&before_payload);
        Ok(source)
    }

    fn data_reference_payload(references: &[u64]) -> Result<Vec<u8>> {
        let mut payload = Vec::new();
        for identifier in references {
            let mut reference = Vec::new();
            reference.extend_from_slice(&[0x08]);
            let mut bytes = [0u8; 10];
            reference.extend_from_slice(encode_varint(*identifier, &mut bytes));
            push_length_delimited(&mut payload, &[0x1a], 1, &reference)?;
        }
        Ok(payload)
    }

    fn reference_pruning_fixture() -> Result<ReferencePruningFixture> {
        let mut changed_field = Vec::new();
        changed_field.extend_from_slice(&[0xa0, 0x01, 0x81, 0x00]);
        changed_field.extend_from_slice(&[0x20, 0x8a, 0x00]);
        changed_field.extend_from_slice(&[0xa0, 0x00, 0xe3, 0x00]);
        push_length_delimited(&mut changed_field, &[0x22], 2, &[0x0c, 0xe4, 0x00, 0x0e])?;
        push_length_delimited(&mut changed_field, &[0x0a], 2, &[])?;
        push_length_delimited(&mut changed_field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;

        let mut expected_changed_field = Vec::new();
        expected_changed_field.extend_from_slice(&[0xa0, 0x01, 0x81, 0x00]);
        expected_changed_field.extend_from_slice(&[0xa0, 0x00, 0xe3, 0x00]);
        push_length_delimited(&mut expected_changed_field, &[0x22], 2, &[0xe4, 0x00])?;
        push_length_delimited(&mut expected_changed_field, &[0x0a], 2, &[])?;
        push_length_delimited(&mut expected_changed_field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;

        let mut untouched_field = Vec::new();
        push_length_delimited(&mut untouched_field, &[0x0a], 1, &[])?;
        push_length_delimited(&mut untouched_field, &[0xa2, 0x00], 2, &[0xcd, 0x00, 0x4e])?;
        untouched_field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);

        let mut message = Vec::new();
        message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00, 0x87, 0x00]);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00, 0x83, 0x00]);
        message.extend_from_slice(&[0x28, 0x8a, 0x00]);
        message.extend_from_slice(&[0xa8, 0x00, 0x8d, 0x00]);
        push_length_delimited(&mut message, &[0x2a], 2, &[0x0c, 0x8d, 0x00, 0x0e])?;
        push_length_delimited(&mut message, &[0x2a], 2, &[0x0a, 0x0c])?;
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, &changed_field)?;
        push_length_delimited(&mut message, &[0x22], 2, &untouched_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;

        let mut expected_message = Vec::new();
        expected_message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        expected_message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00, 0x89, 0x00]);
        push_length_delimited(&mut expected_message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        expected_message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00, 0x82, 0x01]);
        expected_message.extend_from_slice(&[0xa8, 0x00, 0x8d, 0x00]);
        push_length_delimited(&mut expected_message, &[0x2a], 2, &[0x8d, 0x00])?;
        push_length_delimited(&mut expected_message, &[0x2a], 2, &[])?;
        push_length_delimited(
            &mut expected_message,
            &[0xa2, 0x00],
            2,
            &expected_changed_field,
        )?;
        push_length_delimited(&mut expected_message, &[0x22], 2, &untouched_field)?;
        push_length_delimited(&mut expected_message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;

        let mut header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut header, &[0x92, 0x00], 2, &message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);

        let mut expected_header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut expected_header, &[0x92, 0x00], 2, &expected_message)?;
        expected_header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);

        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(ReferencePruningFixture {
            source,
            header,
            expected_header,
        })
    }

    fn push_length_delimited(
        output: &mut Vec<u8>,
        key: &[u8],
        prefix_width: usize,
        payload: &[u8],
    ) -> Result<()> {
        let additional = key
            .len()
            .checked_add(prefix_width)
            .and_then(|length| length.checked_add(payload.len()))
            .ok_or_else(|| Error::invalid_archive(0, "test field length overflow"))?;
        output
            .try_reserve(additional)
            .map_err(|_| Error::allocation("test length-delimited field", additional))?;
        output.extend_from_slice(key);
        let mut prefix = [0u8; 10];
        output.extend_from_slice(encode_varint_with_width(
            payload.len() as u64,
            prefix_width,
            &mut prefix,
        ));
        output.extend_from_slice(payload);
        Ok(())
    }

    fn split_test_archive(data: &[u8]) -> Result<(&[u8], &[u8])> {
        let (header_length, prefix_length) = super::decode_varint(data)?;
        let header_length = usize::try_from(header_length)
            .map_err(|_| Error::invalid_archive(0, "test header length exceeds usize"))?;
        let header_end = prefix_length
            .checked_add(header_length)
            .ok_or_else(|| Error::invalid_archive(0, "test header range overflow"))?;
        let header = data
            .get(prefix_length..header_end)
            .ok_or_else(|| Error::invalid_archive(0, "test header range is invalid"))?;
        let payload = data
            .get(header_end..)
            .ok_or_else(|| Error::invalid_archive(0, "test payload range is invalid"))?;
        Ok((header, payload))
    }
}
