//! Source-checked snapshots and atomic edits for one PPT document structure.

use std::sync::Arc;

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

use super::codec;
use super::model::{CustomTableStylesPlacement, DocumentStructure, Limits, Master, Slide};
use super::validation;

/// Deterministic identity of one exact serialized document structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64 ^ bytes.len() as u64;
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Compact source fingerprint useful for optimistic owner checks.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// Kind of semantic operation staged by a document-structure transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChangeKind {
    SlidesReordered,
    SlideNonOutlineDataChanged,
    SlideInserted,
    SlideRemoved,
    MastersReordered,
    CustomTableStylesMoved,
    CustomTableStylesInserted,
    CustomTableStylesRemoved,
    CustomTableStylesReplaced,
}

/// One reversible semantic operation over the document's structural lists.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Change {
    /// Reorder slide groups while retaining each group's text and opaque atoms.
    SlidesReordered { before: Vec<u32>, after: Vec<u32> },
    /// Change whether one slide contains non-outline data while retaining every other flag.
    SlideNonOutlineDataChanged {
        position: usize,
        before: bool,
        after: bool,
    },
    /// Insert one complete slide-list group at a semantic position.
    SlideInserted { position: usize, group: Vec<Record> },
    /// Remove one complete slide-list group at a semantic position.
    SlideRemoved { position: usize, group: Vec<Record> },
    /// Reorder master reference groups.
    MastersReordered { before: Vec<u32>, after: Vec<u32> },
    /// Move an existing custom table-style atom around `EndDocumentAtom`.
    CustomTableStylesMoved {
        before: CustomTableStylesPlacement,
        after: CustomTableStylesPlacement,
    },
    /// Insert a custom table-style atom at the document tail.
    CustomTableStylesInserted {
        placement: CustomTableStylesPlacement,
        record: Record,
    },
    /// Remove a custom table-style atom from the document tail.
    CustomTableStylesRemoved {
        placement: CustomTableStylesPlacement,
        record: Record,
    },
    /// Replace the payload while retaining the atom's placement.
    CustomTableStylesReplaced {
        placement: CustomTableStylesPlacement,
        before: Record,
        after: Record,
    },
}

impl Change {
    /// Semantic category of this operation.
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        match self {
            Self::SlidesReordered { .. } => ChangeKind::SlidesReordered,
            Self::SlideNonOutlineDataChanged { .. } => ChangeKind::SlideNonOutlineDataChanged,
            Self::SlideInserted { .. } => ChangeKind::SlideInserted,
            Self::SlideRemoved { .. } => ChangeKind::SlideRemoved,
            Self::MastersReordered { .. } => ChangeKind::MastersReordered,
            Self::CustomTableStylesMoved { .. } => ChangeKind::CustomTableStylesMoved,
            Self::CustomTableStylesInserted { .. } => ChangeKind::CustomTableStylesInserted,
            Self::CustomTableStylesRemoved { .. } => ChangeKind::CustomTableStylesRemoved,
            Self::CustomTableStylesReplaced { .. } => ChangeKind::CustomTableStylesReplaced,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::SlidesReordered { before, after } => Self::SlidesReordered {
                before: after.clone(),
                after: before.clone(),
            },
            Self::SlideNonOutlineDataChanged {
                position,
                before,
                after,
            } => Self::SlideNonOutlineDataChanged {
                position: *position,
                before: *after,
                after: *before,
            },
            Self::SlideInserted { position, group } => Self::SlideRemoved {
                position: *position,
                group: group.clone(),
            },
            Self::SlideRemoved { position, group } => Self::SlideInserted {
                position: *position,
                group: group.clone(),
            },
            Self::MastersReordered { before, after } => Self::MastersReordered {
                before: after.clone(),
                after: before.clone(),
            },
            Self::CustomTableStylesMoved { before, after } => Self::CustomTableStylesMoved {
                before: *after,
                after: *before,
            },
            Self::CustomTableStylesInserted { placement, record } => {
                Self::CustomTableStylesRemoved {
                    placement: *placement,
                    record: record.clone(),
                }
            },
            Self::CustomTableStylesRemoved { placement, record } => {
                Self::CustomTableStylesInserted {
                    placement: *placement,
                    record: record.clone(),
                }
            },
            Self::CustomTableStylesReplaced {
                placement,
                before,
                after,
            } => Self::CustomTableStylesReplaced {
                placement: *placement,
                before: after.clone(),
                after: before.clone(),
            },
        }
    }
}

/// Immutable, source-preserving document-structure snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    root: Record,
    structure: DocumentStructure,
    revision: Revision,
    limits: Limits,
}

impl Snapshot {
    /// Parse exactly one complete `DocumentContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one document structure with explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let input = bytes.as_ref();
        if input.len() > validated_limits.max_bytes {
            return Err(Error::InvalidFormat(
                "document-structure snapshot exceeds its byte limit".into(),
            ));
        }
        let (root, consumed) = Record::parse_strict(input, 0)?;
        if consumed != input.len() {
            return Err(Error::Corrupted(
                "document-structure input contains trailing bytes".into(),
            ));
        }
        let structure = validation::validate_document(&root, validated_limits)?;
        let encoded = codec::encode_document(&root)?;
        if encoded != input {
            return Err(Error::Corrupted(
                "document-structure snapshot is not losslessly representable".into(),
            ));
        }
        Ok(Self::from_parts(
            Arc::from(input.to_vec().into_boxed_slice()),
            root,
            structure,
            validated_limits,
        ))
    }

    /// Capture a parsed record tree using the default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record(root: Record) -> Result<Self> {
        Self::from_record_with_limits(root, Limits::default())
    }

    /// Capture a parsed record tree with explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record_with_limits(mut root: Record, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        codec::synchronize(&mut root)?;
        let bytes = codec::encode_document(&root)?;
        if bytes.len() > validated_limits.max_bytes {
            return Err(Error::InvalidFormat(
                "document-structure snapshot exceeds its byte limit".into(),
            ));
        }
        Self::parse_with_limits(bytes, validated_limits)
    }

    fn from_parts(
        bytes: Arc<[u8]>,
        root: Record,
        structure: DocumentStructure,
        limits: Limits,
    ) -> Self {
        let revision = Revision::from_bytes(&bytes);
        Self {
            bytes,
            root,
            structure,
            revision,
            limits,
        }
    }

    /// Exact source or committed `DocumentContainer` bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Parsed record tree retained for advanced, read-only inspection.
    #[must_use]
    pub const fn record(&self) -> &Record {
        &self.root
    }

    /// Typed structural inventory in source order.
    #[must_use]
    pub const fn structure(&self) -> &DocumentStructure {
        &self.structure
    }

    /// Typed master references in source order.
    #[must_use]
    pub fn masters(&self) -> &[Master] {
        self.structure.masters()
    }

    /// Typed presentation-slide references in source order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        self.structure.slides()
    }

    /// Opaque top-level records retained outside the modeled structure.
    pub fn unknown_atoms(&self) -> impl Iterator<Item = &Record> {
        self.root
            .children
            .iter()
            .filter(|record| !validation::is_top_level_known(record))
    }

    /// Compact identity of the exact serialized snapshot.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Resource bounds retained for subsequent edits and patch application.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Start an isolated semantic edit over this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.root.clone(),
            changes: Vec::new(),
        }
    }
}

/// Isolated, failure-atomic semantic edit over one document structure.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Record,
    changes: Vec<Change>,
}

impl Transaction {
    /// Immutable source snapshot used for optimistic conflict checks.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Current transaction-local record tree.
    #[must_use]
    pub const fn record(&self) -> &Record {
        &self.candidate
    }

    /// Current transaction-local structural inventory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn structure(&self) -> Result<DocumentStructure> {
        validation::validate_document(&self.candidate, self.source.limits)
    }

    /// Current transaction-local master references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn masters(&self) -> Result<Vec<Master>> {
        Ok(self.structure()?.masters().to_vec())
    }

    /// Current transaction-local slide references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slides(&self) -> Result<Vec<Slide>> {
        Ok(self.structure()?.slides().to_vec())
    }

    /// Whether any staged operation changes the candidate tree.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.root
    }

    /// Semantic operations staged in call order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Move one slide from its current zero-based position to `destination`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_slide(&mut self, from: usize, destination: usize) -> Result<()> {
        let count = self.slides()?.len();
        if from >= count || destination >= count {
            return invalid("slide position is out of range");
        }
        let mut order: Vec<_> = (0..count).collect();
        let value = order.remove(from);
        order.insert(destination, value);
        self.reorder_slides(&order)
    }

    /// Move a slide selected by its semantic `slideId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_slide_id(&mut self, slide_id: u32, destination: usize) -> Result<()> {
        let index = self
            .slides()?
            .iter()
            .position(|slide| slide.slide_id() == slide_id)
            .ok_or_else(|| Error::InvalidFormat(format!("slide ID {slide_id} was not found")))?;
        self.move_slide(index, destination)
    }

    /// Changes whether one slide contains data other than placeholder text.
    ///
    /// This edits only `fNonOutlineData` (bit 2) of the fixed-width
    /// `SlidePersistAtom.flags` field, as defined by `[MS-PPT]` 2.4.14.5.
    /// It does not control whether the slide is hidden during a slide show;
    /// that state belongs to `SlideShowSlideInfoAtom.fHidden`.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is absent or the resulting document
    /// violates the structural invariants.
    pub fn set_slide_non_outline_data(&mut self, position: usize, present: bool) -> Result<()> {
        const NON_OUTLINE_DATA_FLAG: u32 = 1 << 2;

        let slide = self
            .slides()?
            .get(position)
            .copied()
            .ok_or_else(|| Error::InvalidFormat("slide position is out of range".into()))?;
        let before = slide.flags() & NON_OUTLINE_DATA_FLAG != 0;
        if before == present {
            return Ok(());
        }
        let list_index = validation::list_index(&self.candidate, 0)?
            .ok_or_else(|| Error::InvalidFormat("the presentation slide list is absent".into()))?;
        let mut candidate = self.candidate.clone();
        let atom = candidate.children[list_index]
            .children
            .get_mut(slide.child_index())
            .ok_or_else(|| Error::Corrupted("slide reference index is inconsistent".into()))?;
        if atom.record_type != RecordType::SlidePersistAtom || atom.data.len() != 20 {
            return Err(Error::Corrupted(
                "slide reference is not a fixed-width SlidePersistAtom".into(),
            ));
        }
        let mut flags = u32::from_le_bytes(
            atom.data[4..8]
                .try_into()
                .map_err(|_error| Error::Corrupted("slide flags are truncated".into()))?,
        );
        if present {
            flags |= NON_OUTLINE_DATA_FLAG;
        } else {
            flags &= !NON_OUTLINE_DATA_FLAG;
        }
        atom.data[4..8].copy_from_slice(&flags.to_le_bytes());
        self.publish_candidate(candidate)?;
        self.changes.push(Change::SlideNonOutlineDataChanged {
            position,
            before,
            after: present,
        });
        Ok(())
    }

    /// Returns one complete slide-list group without changing the candidate.
    ///
    /// The returned records begin with the selected `SlidePersistAtom` and
    /// include every following outline-text record owned by that slide.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is absent or the slide list is
    /// malformed.
    pub fn slide_group(&self, position: usize) -> Result<Vec<Record>> {
        let list_index = validation::list_index(&self.candidate, 0)?
            .ok_or_else(|| Error::InvalidFormat("the presentation slide list is absent".into()))?;
        grouped_children(&self.candidate.children[list_index])?
            .groups
            .get(position)
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("slide position is out of range".into()))
    }

    /// Inserts one complete slide-list group at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is outside the insertion range, the
    /// group does not begin with exactly one `SlidePersistAtom`, or the
    /// resulting document violates identifier and outline-text invariants.
    pub fn insert_slide_group(&mut self, position: usize, group: Vec<Record>) -> Result<()> {
        validate_slide_group(&group)?;
        let list_index = validation::list_index(&self.candidate, 0)?
            .ok_or_else(|| Error::InvalidFormat("the presentation slide list is absent".into()))?;
        let mut candidate = self.candidate.clone();
        let groups = grouped_children_allow_empty(&candidate.children[list_index])?;
        if position > groups.groups.len() {
            return invalid("slide insertion position is out of range");
        }
        let mut reordered = groups.groups;
        reordered.insert(position, group.clone());
        let mut children = groups.prefix;
        for owned_group in reordered {
            children.extend(owned_group);
        }
        candidate.children[list_index].children = children;
        self.publish_candidate(candidate)?;
        self.changes.push(Change::SlideInserted { position, group });
        Ok(())
    }

    /// Removes and returns one complete slide-list group.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is absent or publication would
    /// violate the document invariants.
    pub fn remove_slide(&mut self, position: usize) -> Result<Vec<Record>> {
        let list_index = validation::list_index(&self.candidate, 0)?
            .ok_or_else(|| Error::InvalidFormat("the presentation slide list is absent".into()))?;
        let mut candidate = self.candidate.clone();
        let groups = grouped_children(&candidate.children[list_index])?;
        if position >= groups.groups.len() {
            return invalid("slide position is out of range");
        }
        let mut reordered = groups.groups;
        let removed = reordered.remove(position);
        let mut children = groups.prefix;
        for owned_group in reordered {
            children.extend(owned_group);
        }
        candidate.children[list_index].children = children;
        self.publish_candidate(candidate)?;
        self.changes.push(Change::SlideRemoved {
            position,
            group: removed.clone(),
        });
        Ok(removed)
    }

    /// Reorder all slide groups using a checked permutation of current positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_slides(&mut self, order: &[usize]) -> Result<()> {
        let structure = self.structure()?;
        let before = structure
            .slides()
            .iter()
            .map(|slide| slide.slide_id())
            .collect::<Vec<_>>();
        let groups = self.reorder_list(0, order)?;
        if before == order.iter().map(|index| before[*index]).collect::<Vec<_>>() {
            return Ok(());
        }
        let after = order.iter().map(|index| before[*index]).collect();
        self.changes.push(Change::SlidesReordered { before, after });
        let _ = groups;
        Ok(())
    }

    /// Move one master from its current zero-based position to `destination`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_master(&mut self, from: usize, destination: usize) -> Result<()> {
        let count = self.masters()?.len();
        if from >= count || destination >= count {
            return invalid("master position is out of range");
        }
        let mut order: Vec<_> = (0..count).collect();
        let value = order.remove(from);
        order.insert(destination, value);
        self.reorder_masters(&order)
    }

    /// Move a master selected by its semantic `masterId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_master_id(&mut self, master_id: u32, destination: usize) -> Result<()> {
        let index = self
            .masters()?
            .iter()
            .position(|master| master.master_id() == master_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("master ID {master_id:#010x} was not found"))
            })?;
        self.move_master(index, destination)
    }

    /// Reorder all master groups using a checked permutation of current positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_masters(&mut self, order: &[usize]) -> Result<()> {
        let structure = self.structure()?;
        let before = structure
            .masters()
            .iter()
            .map(|master| master.master_id())
            .collect::<Vec<_>>();
        let groups = self.reorder_list(1, order)?;
        if before == order.iter().map(|index| before[*index]).collect::<Vec<_>>() {
            return Ok(());
        }
        let after = order.iter().map(|index| before[*index]).collect();
        self.changes
            .push(Change::MastersReordered { before, after });
        let _ = groups;
        Ok(())
    }

    /// Move an existing custom-table-style atom before or after the end atom.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_custom_table_styles(
        &mut self,
        placement: CustomTableStylesPlacement,
    ) -> Result<()> {
        let structure = self.structure()?;
        let before = structure.custom_table_styles.ok_or_else(|| {
            Error::InvalidFormat("the document has no custom table-style atom".into())
        })?;
        if before == placement {
            return Ok(());
        }
        let mut candidate = self.candidate.clone();
        let style_index = find_custom_styles(&candidate)?;
        let style = candidate.children.remove(style_index);
        let end_index = candidate
            .children
            .iter()
            .position(|record| record.record_type == RecordType::EndDocument)
            .ok_or_else(|| Error::Corrupted("document lost EndDocumentAtom during edit".into()))?;
        let insert_at = match placement {
            CustomTableStylesPlacement::BeforeEndDocument => end_index,
            CustomTableStylesPlacement::AfterEndDocument => end_index + 1,
        };
        candidate.children.insert(insert_at, style);
        self.publish_candidate(candidate)?;
        self.changes.push(Change::CustomTableStylesMoved {
            before,
            after: placement,
        });
        Ok(())
    }

    /// Insert one opaque custom-table-style atom at the document tail.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert_custom_table_styles(
        &mut self,
        record: Record,
        placement: CustomTableStylesPlacement,
    ) -> Result<()> {
        validate_custom_styles_record(&record)?;
        if self.structure()?.custom_table_styles.is_some() {
            return invalid("the document already has a custom table-style atom");
        }
        let mut candidate = self.candidate.clone();
        let end_index = candidate
            .children
            .iter()
            .position(|item| item.record_type == RecordType::EndDocument)
            .ok_or_else(|| Error::Corrupted("document has no EndDocumentAtom".into()))?;
        let index = match placement {
            CustomTableStylesPlacement::BeforeEndDocument => end_index,
            CustomTableStylesPlacement::AfterEndDocument => end_index + 1,
        };
        candidate.children.insert(index, record.clone());
        self.publish_candidate(candidate)?;
        self.changes
            .push(Change::CustomTableStylesInserted { placement, record });
        Ok(())
    }

    /// Remove the custom-table-style atom and return its opaque record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_custom_table_styles(&mut self) -> Result<Record> {
        let placement = self.structure()?.custom_table_styles.ok_or_else(|| {
            Error::InvalidFormat("the document has no custom table-style atom".into())
        })?;
        let mut candidate = self.candidate.clone();
        let index = find_custom_styles(&candidate)?;
        let record = candidate.children.remove(index);
        self.publish_candidate(candidate)?;
        self.changes.push(Change::CustomTableStylesRemoved {
            placement,
            record: record.clone(),
        });
        Ok(record)
    }

    /// Replace custom-table-style bytes without changing their placement.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_custom_table_styles(&mut self, record: Record) -> Result<Record> {
        validate_custom_styles_record(&record)?;
        let placement = self.structure()?.custom_table_styles.ok_or_else(|| {
            Error::InvalidFormat("the document has no custom table-style atom".into())
        })?;
        let mut candidate = self.candidate.clone();
        let index = find_custom_styles(&candidate)?;
        let before = candidate.children[index].clone();
        if before == record {
            return Ok(before);
        }
        candidate.children[index] = record.clone();
        self.publish_candidate(candidate)?;
        self.changes.push(Change::CustomTableStylesReplaced {
            placement,
            before: before.clone(),
            after: record,
        });
        Ok(before)
    }

    /// Capture the current candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.encoded_candidate()?;
        if bytes.as_slice() == self.source.bytes.as_ref() {
            return Ok(self.source.clone());
        }
        Snapshot::parse_with_limits(bytes, self.source.limits)
    }

    /// Validate and publish the candidate atomically with its reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        let bytes = self.encoded_candidate()?;
        let source = self.source;
        let snapshot = if bytes.as_slice() == source.bytes.as_ref() {
            source.clone()
        } else {
            Snapshot::parse_with_limits(bytes, source.limits)?
        };
        let changes = if snapshot.bytes == source.bytes {
            Vec::new()
        } else {
            self.changes
        };
        let patch = Patch {
            base: source.revision,
            target: snapshot.revision,
            before: source.bytes.clone(),
            after: snapshot.bytes.clone(),
            changes,
            limits: source.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Alias for move-owned writer terminology.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard all staged edits and recover the exact source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn reorder_list(&mut self, instance: u16, order: &[usize]) -> Result<Vec<Vec<Record>>> {
        let structure = self.structure()?;
        let expected = match instance {
            0 => structure.slides().len(),
            1 => structure.masters().len(),
            _ => unreachable!("document structure only reorders master or slide lists"),
        };
        validate_permutation(order, expected)?;
        let list_index = validation::list_index(&self.candidate, instance)?.ok_or_else(|| {
            Error::InvalidFormat("the requested structural list is absent".into())
        })?;
        let mut candidate = self.candidate.clone();
        let groups = grouped_children(&candidate.children[list_index])?;
        let mut children = groups.prefix;
        for index in order {
            children.extend(groups.groups[*index].iter().cloned());
        }
        candidate.children[list_index].children = children;
        codec::synchronize(&mut candidate)?;
        validation::validate_document(&candidate, self.source.limits)?;
        self.candidate = candidate;
        Ok(groups.groups)
    }

    fn publish_candidate(&mut self, mut candidate: Record) -> Result<()> {
        codec::synchronize(&mut candidate)?;
        validation::validate_document(&candidate, self.source.limits)?;
        self.candidate = candidate;
        Ok(())
    }

    fn encoded_candidate(&self) -> Result<Vec<u8>> {
        let mut candidate = self.candidate.clone();
        codec::synchronize(&mut candidate)?;
        validation::validate_document(&candidate, self.source.limits)?;
        let bytes = codec::encode_document(&candidate)?;
        if bytes.len() > self.source.limits.max_bytes {
            return Err(Error::InvalidFormat(
                "document-structure candidate exceeds its byte limit".into(),
            ));
        }
        Ok(bytes)
    }
}

/// Successful immutable target and its source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from the source to this target.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Undo this commit against its exact target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.undo(current)
    }

    /// Redo this commit against its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.redo(current)
    }

    /// Split the published target and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible patch bound to the exact source bytes that produced it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    changes: Vec<Change>,
    limits: Limits,
}

impl Patch {
    /// Source revision required for forward application.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Target revision produced by forward application.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Semantic operations represented by this patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Exact source bytes bound to this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact target bytes bound to this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Compatibility-free explicit aliases for byte-oriented callers.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before()
    }

    /// Exact target bytes bound to this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after()
    }

    /// Whether this patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Apply only to the exact source snapshot used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision != self.base || current.bytes.as_ref() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "document-structure patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(current.clone());
        }
        Snapshot::parse_with_limits(self.after.as_ref(), self.limits)
    }

    /// Apply the inverse to the exact committed target.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(current)
    }

    /// Reapply this patch to its exact source.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.apply(current)
    }

    /// Build a source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
            limits: self.limits,
        }
    }
}

#[derive(Debug)]
struct Groups {
    prefix: Vec<Record>,
    groups: Vec<Vec<Record>>,
}

fn grouped_children(list: &Record) -> Result<Groups> {
    let starts = list
        .children
        .iter()
        .enumerate()
        .filter_map(|(index, child)| {
            (child.record_type == RecordType::SlidePersistAtom).then_some(index)
        })
        .collect::<Vec<_>>();
    if starts.is_empty() {
        return invalid("the structural list contains no persist references");
    }
    let mut groups = Vec::with_capacity(starts.len());
    for (position, start) in starts.iter().copied().enumerate() {
        let end = starts
            .get(position + 1)
            .copied()
            .unwrap_or(list.children.len());
        groups.push(list.children[start..end].to_vec());
    }
    Ok(Groups {
        prefix: list.children[..starts[0]].to_vec(),
        groups,
    })
}

fn grouped_children_allow_empty(list: &Record) -> Result<Groups> {
    let has_slide = list
        .children
        .iter()
        .any(|child| child.record_type == RecordType::SlidePersistAtom);
    if has_slide {
        grouped_children(list)
    } else {
        Ok(Groups {
            prefix: list.children.clone(),
            groups: Vec::new(),
        })
    }
}

fn validate_slide_group(group: &[Record]) -> Result<()> {
    if group.first().map(|record| record.record_type) != Some(RecordType::SlidePersistAtom)
        || group
            .iter()
            .skip(1)
            .any(|record| record.record_type == RecordType::SlidePersistAtom)
    {
        return invalid("slide group must begin with exactly one SlidePersistAtom");
    }
    Ok(())
}

fn validate_permutation(order: &[usize], expected: usize) -> Result<()> {
    if order.len() != expected {
        return invalid("structural reorder must contain every current position exactly once");
    }
    let mut seen = vec![false; expected];
    for &index in order {
        let slot = seen.get_mut(index).ok_or_else(|| {
            Error::InvalidFormat("structural reorder position is out of range".into())
        })?;
        if *slot {
            return invalid("structural reorder contains a duplicate position");
        }
        *slot = true;
    }
    Ok(())
}

fn find_custom_styles(document: &Record) -> Result<usize> {
    document
        .children
        .iter()
        .position(|record| record.record_type == RecordType::RoundTripCustomTableStyles12Atom)
        .ok_or_else(|| Error::Corrupted("document has no custom table-style atom".into()))
}

fn validate_custom_styles_record(record: &Record) -> Result<()> {
    if record.record_type != RecordType::RoundTripCustomTableStyles12Atom
        || record.record_type_raw != RecordType::RoundTripCustomTableStyles12Atom.as_u16()
        || record.instance != 0
        || record.version > 0x000f
        || !record.children.is_empty()
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return invalid("custom table-style atom has an invalid record header or payload");
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
