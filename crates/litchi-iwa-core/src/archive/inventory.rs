//! Allocation-free accounting for cloning an already parsed physical archive.
//!
//! The inventory deliberately describes only the deep clone of the supplied
//! in-memory source. It does not describe decoding a new wire candidate,
//! encoding a changed header, or the temporary work required by an identity
//! remap. Format owners must admit those operations separately.

#![allow(
    clippy::missing_errors_doc,
    reason = "The visitor and archive methods share the physical-layer Result contract."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "The public names identify the clone inventory seam explicitly."
)]

use std::{fmt, mem::size_of, result::Result};

use super::{Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, MessageInfo, RawMessage};

/// One heap-backed collection visited while cloning a parsed archive.
///
/// `retained_bytes` counts the logical bytes retained by the clone for this
/// collection. `traversal_bytes` is the conservative source walk charge for
/// the collection; it is reported separately so a caller can charge clone
/// retention and clone traversal with different ledgers. For the physical
/// collections currently owned here, both values are equal. `elements` is the
/// number of logical elements in a vector, or `1` for an optional byte/string
/// allocation. `allocations` is the number of heap allocation events charged
/// for the collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveCloneInventoryChunk {
    kind: ArchiveCloneInventoryKind,
    elements: usize,
    retained_bytes: usize,
    traversal_bytes: usize,
    allocations: usize,
}

impl ArchiveCloneInventoryChunk {
    /// The physical collection being admitted.
    #[must_use]
    pub const fn kind(self) -> ArchiveCloneInventoryKind {
        self.kind
    }

    /// Number of logical elements in this collection.
    #[must_use]
    pub const fn elements(self) -> usize {
        self.elements
    }

    /// Heap bytes retained by the cloned collection.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Conservative bytes charged for walking this collection.
    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }

    /// Heap allocation events charged for the collection.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Physical collection kind reported before its nested values are traversed.
///
/// These are collection names, rather than protobuf wire-field counts. Raw
/// unknown header fields are retained in the two original-header boxes and
/// are therefore represented by [`Self::OriginalHeader`] and
/// [`Self::OriginalCanonicalHeader`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ArchiveCloneInventoryKind {
    /// The archive's top-level `Vec<ArchiveObject>`.
    Objects,
    /// An object's `Vec<RawMessage>` metadata slots.
    Messages,
    /// A raw message's payload `Vec<u8>`.
    MessagePayload,
    /// An archive header's `Vec<MessageInfo>`.
    MessageInfos,
    /// A message's version list.
    Versions,
    /// A message's `Vec<FieldInfo>`.
    FieldInfos,
    /// A message or field aggregate object-reference list.
    ObjectReferences,
    /// A message or field aggregate data-reference list.
    DataReferences,
    /// A message's diff-merge version list.
    DiffMergeVersions,
    /// A message's optional diff field path.
    DiffFieldPath,
    /// A message's removed-field path list.
    FieldsToRemove,
    /// A message's diff-read version list.
    DiffReadVersions,
    /// Any `FieldPath.path` vector, including removed-field paths.
    FieldPath,
    /// A field's known-field-version list.
    KnownFieldVersions,
    /// A field's optional known feature identifier string.
    FeatureIdentifier,
    /// The source-authoritative noncanonical ArchiveInfo header.
    OriginalHeader,
    /// The canonical ArchiveInfo header paired with [`Self::OriginalHeader`].
    OriginalCanonicalHeader,
}

/// Admission callback for an allocation-free archive clone inventory walk.
///
/// The callback runs after the collection's length has been checked and before
/// any element in that collection is inspected. A caller can therefore reject
/// a zero-budget or max-minus-one request before the walk reaches nested
/// metadata. The callback may update its own finite budget; the inventory
/// itself performs checked accounting and allocates no memory.
pub trait ArchiveCloneInventoryVisitor {
    /// Error returned when the caller's finite budget rejects a collection.
    type Error;

    /// Admit one collection before the inventory traverses it.
    fn admit(&mut self, chunk: ArchiveCloneInventoryChunk) -> Result<(), Self::Error>;
}

impl<F, E> ArchiveCloneInventoryVisitor for F
where
    F: FnMut(ArchiveCloneInventoryChunk) -> Result<(), E>,
{
    type Error = E;

    fn admit(&mut self, chunk: ArchiveCloneInventoryChunk) -> Result<(), E> {
        self(chunk)
    }
}

/// Failure from a clone inventory walk.
///
/// `Admission` preserves the format owner's typed finite-budget error. The
/// inventory's own arithmetic is allocation-free and can only fail with the
/// checked-overflow variant.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ArchiveCloneInventoryError<E> {
    /// The visitor rejected the next collection before traversal.
    Admission(E),
    /// A collection or aggregate counter exceeded `usize`.
    ArithmeticOverflow { resource: &'static str },
}

/// Result type for a clone inventory walk with a caller-defined admission
/// error.
pub type ArchiveCloneInventoryResult<T, E> = Result<T, ArchiveCloneInventoryError<E>>;

impl<E: fmt::Display> fmt::Display for ArchiveCloneInventoryError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Admission(error) => {
                write!(formatter, "clone inventory admission failed: {error}")
            },
            Self::ArithmeticOverflow { resource } => {
                write!(formatter, "clone inventory {resource} overflow")
            },
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for ArchiveCloneInventoryError<E> {}

/// Exact deep-clone accounting for one parsed archive or object.
///
/// The byte and allocation fields cover heap storage reachable from the
/// derived deep clone: vector backing stores, payload bytes, strings, and both
/// source-authoritative raw-header boxes. Inline Rust struct fields are not
/// included. `objects`, `messages`, `message_infos`, `field_infos`, and
/// `references` are structural counts, where `references` counts every
/// object and data reference occurrence in message and field metadata.
///
/// This report is a clone census for the supplied source. It is not a bound on
/// parsing or encoding a changed candidate, whose header and scratch
/// requirements must be admitted by the operation that creates it.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ArchiveCloneInventory {
    objects: usize,
    messages: usize,
    message_infos: usize,
    field_infos: usize,
    references: usize,
    allocations: usize,
    retained_bytes: usize,
    traversal_bytes: usize,
}

impl ArchiveCloneInventory {
    /// Number of cloned archive objects.
    #[must_use]
    pub const fn objects(self) -> usize {
        self.objects
    }

    /// Number of cloned raw messages.
    #[must_use]
    pub const fn messages(self) -> usize {
        self.messages
    }

    /// Number of cloned `MessageInfo` values.
    #[must_use]
    pub const fn message_infos(self) -> usize {
        self.message_infos
    }

    /// Number of cloned `FieldInfo` values.
    #[must_use]
    pub const fn field_infos(self) -> usize {
        self.field_infos
    }

    /// Number of object and data reference occurrences.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Number of heap allocation events in the deep clone.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Heap bytes retained by the deep clone.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Conservative bytes charged while walking clone-owned collections.
    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }

    fn admit_collection_with_error<E, V>(
        &mut self,
        visitor: &mut V,
        kind: ArchiveCloneInventoryKind,
        elements: usize,
        element_size: usize,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        let bytes = elements.checked_mul(element_size).ok_or(
            ArchiveCloneInventoryError::ArithmeticOverflow {
                resource: "collection size",
            },
        )?;
        let allocations = usize::from(elements != 0);
        visitor
            .admit(ArchiveCloneInventoryChunk {
                kind,
                elements,
                retained_bytes: bytes,
                traversal_bytes: bytes,
                allocations,
            })
            .map_err(ArchiveCloneInventoryError::Admission)?;
        self.add_bytes(bytes)?;
        self.allocations = self.allocations.checked_add(allocations).ok_or_else(|| {
            ArchiveCloneInventoryError::ArithmeticOverflow {
                resource: "allocation count",
            }
        })?;
        Ok(())
    }

    fn admit_optional_bytes<E, V>(
        &mut self,
        visitor: &mut V,
        kind: ArchiveCloneInventoryKind,
        bytes: usize,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        // A present Box<[u8]> or String is conservatively charged as one
        // allocation even when its length is zero. This keeps the report safe
        // across allocator-specific empty-slice representations.
        visitor
            .admit(ArchiveCloneInventoryChunk {
                kind,
                elements: 1,
                retained_bytes: bytes,
                traversal_bytes: bytes,
                allocations: 1,
            })
            .map_err(ArchiveCloneInventoryError::Admission)?;
        self.add_bytes(bytes)?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            ArchiveCloneInventoryError::ArithmeticOverflow {
                resource: "allocation count",
            }
        })?;
        Ok(())
    }

    fn add_bytes<E>(&mut self, bytes: usize) -> ArchiveCloneInventoryResult<(), E> {
        self.retained_bytes = self.retained_bytes.checked_add(bytes).ok_or(
            ArchiveCloneInventoryError::ArithmeticOverflow {
                resource: "retained bytes",
            },
        )?;
        self.traversal_bytes = self.traversal_bytes.checked_add(bytes).ok_or(
            ArchiveCloneInventoryError::ArithmeticOverflow {
                resource: "traversal bytes",
            },
        )?;
        Ok(())
    }

    fn add_count<E>(
        slot: &mut usize,
        amount: usize,
        resource: &'static str,
    ) -> ArchiveCloneInventoryResult<(), E> {
        *slot = slot
            .checked_add(amount)
            .ok_or(ArchiveCloneInventoryError::ArithmeticOverflow { resource })?;
        Ok(())
    }

    fn visit_object<E, V>(
        &mut self,
        object: &ArchiveObject,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        Self::add_count(&mut self.objects, 1, "object count")?;
        self.visit_archive_info(&object.archive_info, visitor)?;

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::Messages,
            object.messages.len(),
            size_of::<RawMessage>(),
        )?;
        Self::add_count(&mut self.messages, object.messages.len(), "message count")?;
        for message in &object.messages {
            self.admit_collection_with_error(
                visitor,
                ArchiveCloneInventoryKind::MessagePayload,
                message.data.len(),
                size_of::<u8>(),
            )?;
        }

        if let Some(header) = object.original_header.as_deref() {
            self.admit_optional_bytes(
                visitor,
                ArchiveCloneInventoryKind::OriginalHeader,
                header.len(),
            )?;
        }
        if let Some(header) = object.original_canonical_header.as_deref() {
            self.admit_optional_bytes(
                visitor,
                ArchiveCloneInventoryKind::OriginalCanonicalHeader,
                header.len(),
            )?;
        }
        Ok(())
    }

    fn visit_archive_info<E, V>(
        &mut self,
        info: &ArchiveInfo,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::MessageInfos,
            info.message_infos.len(),
            size_of::<MessageInfo>(),
        )?;
        Self::add_count(
            &mut self.message_infos,
            info.message_infos.len(),
            "MessageInfo count",
        )?;
        for message in &info.message_infos {
            self.visit_message_info(message, visitor)?;
        }
        Ok(())
    }

    fn visit_message_info<E, V>(
        &mut self,
        message: &MessageInfo,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::Versions,
            message.versions.len(),
            size_of::<u32>(),
        )?;

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::FieldInfos,
            message.field_infos.len(),
            size_of::<FieldInfo>(),
        )?;
        Self::add_count(
            &mut self.field_infos,
            message.field_infos.len(),
            "FieldInfo count",
        )?;
        for field in &message.field_infos {
            self.visit_field_info(field, visitor)?;
        }

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::ObjectReferences,
            message.object_references.len(),
            size_of::<u64>(),
        )?;
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::DataReferences,
            message.data_references.len(),
            size_of::<u64>(),
        )?;
        Self::add_count(
            &mut self.references,
            message.object_references.len(),
            "reference count",
        )?;
        Self::add_count(
            &mut self.references,
            message.data_references.len(),
            "reference count",
        )?;

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::DiffMergeVersions,
            message.diff_merge_version.len(),
            size_of::<u32>(),
        )?;
        if let Some(path) = message.diff_field_path.as_ref() {
            self.admit_collection_with_error(
                visitor,
                ArchiveCloneInventoryKind::DiffFieldPath,
                path.path.len(),
                size_of::<u32>(),
            )?;
            // The path's allocation is the same vector described by the
            // diff-field-path chunk; no second collection is traversed here.
        }

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::FieldsToRemove,
            message.fields_to_remove.len(),
            size_of::<FieldPath>(),
        )?;
        for path in &message.fields_to_remove {
            self.admit_collection_with_error(
                visitor,
                ArchiveCloneInventoryKind::FieldPath,
                path.path.len(),
                size_of::<u32>(),
            )?;
        }

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::DiffReadVersions,
            message.diff_read_version.len(),
            size_of::<u32>(),
        )?;
        Ok(())
    }

    fn visit_field_info<E, V>(
        &mut self,
        field: &FieldInfo,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<(), E>
    where
        V: ArchiveCloneInventoryVisitor<Error = E>,
    {
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::FieldPath,
            field.path.path.len(),
            size_of::<u32>(),
        )?;
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::ObjectReferences,
            field.object_references.len(),
            size_of::<u64>(),
        )?;
        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::DataReferences,
            field.data_references.len(),
            size_of::<u64>(),
        )?;
        Self::add_count(
            &mut self.references,
            field.object_references.len(),
            "reference count",
        )?;
        Self::add_count(
            &mut self.references,
            field.data_references.len(),
            "reference count",
        )?;

        self.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::KnownFieldVersions,
            field.known_field_version.len(),
            size_of::<u32>(),
        )?;
        if let Some(identifier) = field.known_field_feature_identifier.as_deref() {
            self.admit_optional_bytes(
                visitor,
                ArchiveCloneInventoryKind::FeatureIdentifier,
                identifier.len(),
            )?;
        }
        Ok(())
    }
}

impl ArchiveObject {
    /// Inventory this object's deep clone, admitting every heap collection
    /// before its nested values are traversed.
    pub fn clone_inventory_with_admission<V>(
        &self,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<ArchiveCloneInventory, V::Error>
    where
        V: ArchiveCloneInventoryVisitor,
    {
        let mut inventory = ArchiveCloneInventory::default();
        inventory.visit_object(self, visitor)?;
        Ok(inventory)
    }
}

impl Archive {
    /// Inventory this archive's deep clone, admitting every heap collection
    /// before its nested values are traversed.
    pub fn clone_inventory_with_admission<V>(
        &self,
        visitor: &mut V,
    ) -> ArchiveCloneInventoryResult<ArchiveCloneInventory, V::Error>
    where
        V: ArchiveCloneInventoryVisitor,
    {
        let mut inventory = ArchiveCloneInventory::default();
        inventory.admit_collection_with_error(
            visitor,
            ArchiveCloneInventoryKind::Objects,
            self.objects.len(),
            size_of::<ArchiveObject>(),
        )?;
        for object in &self.objects {
            inventory.visit_object(object, visitor)?;
        }
        Ok(inventory)
    }
}

#[cfg(test)]
mod tests {
    use super::{ArchiveCloneInventoryKind, ArchiveCloneInventoryVisitor};
    use crate::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, Result};
    use std::convert::Infallible;

    #[derive(Default)]
    struct Chunks(Vec<(ArchiveCloneInventoryKind, usize, usize, usize, usize)>);

    impl ArchiveCloneInventoryVisitor for Chunks {
        type Error = Infallible;

        fn admit(
            &mut self,
            chunk: super::ArchiveCloneInventoryChunk,
        ) -> std::result::Result<(), Infallible> {
            self.0.push((
                chunk.kind(),
                chunk.elements(),
                chunk.retained_bytes(),
                chunk.traversal_bytes(),
                chunk.allocations(),
            ));
            Ok(())
        }
    }

    #[test]
    fn inventory_counts_all_nested_clone_storage_and_references() -> Result<()> {
        let mut object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 7,
                data: vec![1, 2, 3],
            }],
        )?;
        let message = &mut object.archive_info.message_infos[0];
        message.versions = vec![1, 2];
        message.object_references = vec![10, 11];
        message.data_references = vec![20];
        message.diff_merge_version = vec![3];
        message.diff_field_path = Some(FieldPath::new(vec![4, 5]));
        message.fields_to_remove = vec![FieldPath::new(vec![6]), FieldPath::default()];
        message.diff_read_version = vec![7];
        let mut field = FieldInfo::new(vec![8, 9]);
        field.object_references = vec![12];
        field.data_references = vec![21, 22];
        field.known_field_version = vec![10];
        field.known_field_feature_identifier = Some("feature".to_owned());
        message.field_infos.push(field);

        let canonical =
            super::super::encode_archive_info(&object.archive_info, crate::Limits::default())?;
        let mut raw = canonical.clone();
        raw.extend_from_slice(&[0xb2, 0x06, 0x01, b'x']);
        let raw_len = raw.len();
        let canonical_len = canonical.len();
        object.original_header = Some(raw.into_boxed_slice());
        object.original_canonical_header = Some(canonical.into_boxed_slice());

        let archive = Archive {
            objects: vec![object],
        };
        let mut chunks = Chunks::default();
        let report = archive
            .clone_inventory_with_admission(&mut chunks)
            .expect("the no-op test visitor cannot reject");

        assert_eq!(report.objects(), 1);
        assert_eq!(report.messages(), 1);
        assert_eq!(report.message_infos(), 1);
        assert_eq!(report.field_infos(), 1);
        assert_eq!(report.references(), 6);
        assert!(report.allocations() > 0);
        assert!(report.retained_bytes() > 0);
        assert_eq!(report.retained_bytes(), report.traversal_bytes());
        assert!(
            chunks
                .0
                .iter()
                .any(|chunk| chunk.0 == ArchiveCloneInventoryKind::OriginalHeader)
        );
        assert!(
            chunks
                .0
                .iter()
                .any(|chunk| { chunk.0 == ArchiveCloneInventoryKind::OriginalCanonicalHeader })
        );
        assert_eq!(
            chunks
                .0
                .iter()
                .filter(|chunk| {
                    matches!(
                        chunk.0,
                        ArchiveCloneInventoryKind::OriginalHeader
                            | ArchiveCloneInventoryKind::OriginalCanonicalHeader
                    )
                })
                .map(|chunk| chunk.2)
                .sum::<usize>(),
            raw_len + canonical_len
        );
        Ok(())
    }

    #[test]
    fn admission_rejects_before_nested_metadata_is_inspected() -> Result<()> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![0; 64],
            }],
        )?;
        object.archive_info.message_infos[0].field_infos = (0..32)
            .map(|index| FieldInfo::new(vec![index; 128]))
            .collect();

        let mut calls = 0;
        let error = object
            .clone_inventory_with_admission(&mut |chunk: super::ArchiveCloneInventoryChunk| {
                calls += 1;
                if chunk.kind() == ArchiveCloneInventoryKind::MessageInfos {
                    return Err(crate::Error::limit(
                        crate::LimitKind::HeaderMemoryBytes,
                        chunk.retained_bytes().saturating_add(1),
                        chunk.retained_bytes(),
                    ));
                }
                Ok(())
            })
            .expect_err("the outer metadata allocation must be rejected");
        assert!(matches!(
            error,
            super::ArchiveCloneInventoryError::Admission(crate::Error::Limit { .. })
        ));
        assert_eq!(calls, 1);
        Ok(())
    }

    #[test]
    fn max_minus_one_rejects_before_the_next_nested_collection() -> Result<()> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![0; 32],
            }],
        )?;
        object.archive_info.message_infos[0].field_infos = (0..8)
            .map(|index| FieldInfo::new(vec![index; 16]))
            .collect();

        let mut exact_chunks = Chunks::default();
        let exact = object
            .clone_inventory_with_admission(&mut exact_chunks)
            .expect("the no-op test visitor cannot reject");
        let maximum = exact
            .retained_bytes()
            .checked_sub(1)
            .expect("the fixture must retain at least one byte");
        let mut retained = 0usize;
        let mut calls = 0usize;
        let error = object
            .clone_inventory_with_admission(&mut |chunk: super::ArchiveCloneInventoryChunk| {
                calls += 1;
                let next = retained
                    .checked_add(chunk.retained_bytes())
                    .expect("test budget arithmetic cannot overflow");
                if next > maximum {
                    return Err(crate::Error::limit(
                        crate::LimitKind::HeaderMemoryBytes,
                        next,
                        maximum,
                    ));
                }
                retained = next;
                Ok(())
            })
            .expect_err("one byte below the exact clone census must reject");
        assert!(matches!(
            error,
            super::ArchiveCloneInventoryError::Admission(crate::Error::Limit { .. })
        ));
        // The final collection is admitted before its bytes would be copied;
        // no inventory collection is visited after the rejection.
        assert_eq!(calls, exact_chunks.0.len());
        assert!(retained <= maximum);
        Ok(())
    }

    #[test]
    fn archive_and_object_reports_have_expected_root_collection_difference() -> Result<()> {
        let object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![3],
            }],
        )?;
        let mut object_chunks = Chunks::default();
        let object_report = object
            .clone_inventory_with_admission(&mut object_chunks)
            .expect("the no-op test visitor cannot reject");
        let mut archive_chunks = Chunks::default();
        let archive_report = Archive {
            objects: vec![object],
        }
        .clone_inventory_with_admission(&mut archive_chunks)
        .expect("the no-op test visitor cannot reject");
        assert_eq!(object_report.objects(), archive_report.objects());
        assert!(archive_report.retained_bytes() > object_report.retained_bytes());
        assert_eq!(
            archive_chunks.0.first().map(|chunk| chunk.0),
            Some(ArchiveCloneInventoryKind::Objects)
        );
        Ok(())
    }
}
