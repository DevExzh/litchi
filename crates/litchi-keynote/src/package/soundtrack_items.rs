//! Physical owner for the semantic Keynote soundtrack-item lifecycle.
//!
//! The public vocabulary lives in [`crate::soundtrack::items`].  This module
//! is deliberately the only place where that vocabulary is connected to an
//! IWA object, PackageMetadata, or a ZIP member.  Every operation is planned
//! against an immutable source snapshot and publishes a new, reopened
//! [`Package`] only after the complete media closure has been checked again.

#![allow(
    clippy::too_many_lines,
    reason = "The transaction is kept in one source-authoritative adapter so its preflight and publication invariants stay adjacent."
)]

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_archive::{SourceCatalog, package::EntryInsertion};
use litchi_iwa_common::media::Type as MediaType;
use litchi_iwa_common::varint::{encode_varint_into, encoded_len};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{
    Archive, DataReferenceTransition, FieldDataReferenceTransition, RawMessage, SnappyStream,
};
use litchi_iwa_protos::package_metadata_media_codec as metadata_codec;
use sha1::{Digest, Sha1};

use crate::soundtrack::items::{
    AudioSource, Commit, Diagnostics, Edit, Error, Item, ItemHandle, ItemSelector, OperationKind,
    Patch, StagedOperation,
};

use super::soundtrack_physical::{
    self, Budget, PACKAGE_METADATA_MESSAGE_TYPE, SOUNDTRACK_MEDIA_FIELD, Selection, SelectionPolicy,
};
use super::{Package, PhysicalSource};

const SOUNDTRACK_MESSAGE_TYPE: u32 = soundtrack_physical::SOUNDTRACK_MESSAGE_TYPE;
const METADATA_COMPONENT: &str = soundtrack_physical::METADATA_COMPONENT;
const DATA_PREFIX: &str = "Data/";
const SHA1_BYTES: usize = 20;

/// A copied, validated DataInfo projection.  The bytes remain in the package
/// catalog; only the small identifying metadata is retained across planning.
#[derive(Debug, Clone, PartialEq, Eq)]
struct DataRecord {
    identifier: u64,
    digest: [u8; SHA1_BYTES],
    preferred_name: Box<str>,
    current_name: Box<str>,
    materialized_length: usize,
    unknown_fields: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct OwnerRecord {
    data_identifier: u64,
    object_identifier: u64,
    count: u32,
    selected: bool,
}

#[derive(Debug, Clone, Default)]
struct MetadataFacts {
    data: Vec<DataRecord>,
    owners: Vec<OwnerRecord>,
    selected_data_references: Vec<u64>,
    selected_component_identifier: Option<u64>,
    selected_component_locator: Option<Box<str>>,
    selected_component_unknown: bool,
}

#[derive(Debug, Clone)]
struct ResolvedItem {
    semantic: Item,
    identifier: u64,
    current_name: Box<str>,
    owner_count: u32,
}

#[derive(Debug)]
struct SourceView<'a> {
    selection: Selection<'a>,
    items: Vec<ResolvedItem>,
    metadata: MetadataFacts,
}

impl<'a> SourceView<'a> {
    fn semantic_items(&self) -> Result<Box<[Item]>, Error> {
        let mut items = Vec::new();
        items
            .try_reserve_exact(self.items.len())
            .map_err(|_| Error::Allocation {
                amount: self.items.len(),
            })?;
        for item in &self.items {
            items.push(item.semantic.clone());
        }
        Ok(items.into_boxed_slice())
    }
}

/// Scan one selected PackageMetadata payload with the lazy Buffa-backed
/// metadata codec.  The visitor copies only bounded semantic facts; media
/// payload bytes and native identifiers do not leave this module.
fn metadata_facts(
    package: &Package,
    catalog: &SourceCatalog,
    expected_locator: &str,
    soundtrack_identifier: u64,
) -> Result<MetadataFacts, Error> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let snappy_limits = package
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|_| Error::InvalidSource)?;
    let archive_limits = package
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| Error::InvalidSource)?;
    let mut payload = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if payload.replace(message.data.as_slice()).is_some() {
                return Err(Error::InvalidSource);
            }
            if object
                .archive_info
                .message_infos
                .get(index)
                .map(|info| info.type_)
                != Some(PACKAGE_METADATA_MESSAGE_TYPE)
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    let payload = payload.ok_or(Error::InvalidSource)?;

    let mut visitor = MetadataVisitor {
        expected_locator,
        soundtrack_identifier,
        facts: MetadataFacts::default(),
    };
    let options = metadata_codec::DecodeOptions::for_source(payload);
    metadata_codec::visit_package_metadata_media(payload, options, &mut visitor)
        .map_err(|_| Error::InvalidSource)?;
    if visitor.facts.selected_component_identifier.is_none()
        || visitor.facts.selected_component_locator.is_none()
    {
        return Err(Error::InvalidSource);
    }
    if visitor.facts.selected_component_unknown {
        return Err(Error::InvalidSource);
    }
    Ok(visitor.facts)
}

struct MetadataVisitor<'a> {
    expected_locator: &'a str,
    soundtrack_identifier: u64,
    facts: MetadataFacts,
}

impl metadata_codec::PackageMetadataMediaVisitor for MetadataVisitor<'_> {
    fn visit_data_info(
        &mut self,
        data_info: metadata_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        let digest = <[u8; SHA1_BYTES]>::try_from(data_info.digest())
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        let materialized_length = data_info
            .materialized_length()
            .and_then(|length| usize::try_from(length).ok())
            .unwrap_or(0);
        let current_name = data_info
            .file_name()
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| data_info.preferred_file_name());
        self.facts
            .data
            .try_reserve(1)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts.data.push(DataRecord {
            identifier: data_info.identifier(),
            digest,
            preferred_name: data_info.preferred_file_name().into(),
            current_name: current_name.into(),
            materialized_length,
            unknown_fields: data_info.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        if component.is_versioned() || component.effective_locator() != self.expected_locator {
            return Ok(());
        }
        if self.facts.selected_component_identifier.is_some() {
            // The codec has already validated the wire shape.  Keep this
            // adapter strict when two unversioned components claim the same
            // effective locator; the post-visit check turns the marker into a
            // semantic InvalidSource error without manufacturing a codec
            // error value in a callback.
            self.facts.selected_component_unknown = true;
            return Ok(());
        }
        self.facts.selected_component_identifier = Some(component.identifier());
        self.facts.selected_component_locator = Some(component.effective_locator().into());
        self.facts.selected_component_unknown = component.has_unknown_fields();
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
        data_reference: metadata_codec::ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        if component.is_versioned()
            || component.effective_locator() != self.expected_locator
            || Some(component.identifier()) != self.facts.selected_component_identifier
        {
            return Ok(());
        }
        self.facts
            .selected_data_references
            .try_reserve(1)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts
            .selected_data_references
            .push(data_reference.data_identifier());
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
        data_reference: metadata_codec::ComponentDataReferenceSnapshot<'_>,
        owner: metadata_codec::OwnerSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        let selected = !component.is_versioned()
            && component.effective_locator() == self.expected_locator
            && Some(component.identifier()) == self.facts.selected_component_identifier;
        self.facts
            .owners
            .try_reserve(1)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts.owners.push(OwnerRecord {
            data_identifier: data_reference.data_identifier(),
            object_identifier: owner.object_identifier(),
            count: owner.count(),
            selected: selected && owner.object_identifier() == self.soundtrack_identifier,
        });
        Ok(())
    }
}

/// Build a source view after strict root/show/soundtrack selection and a
/// complete selected media closure.  The returned value borrows only the
/// immutable package supplied by the caller.
fn source_view<'a>(
    package: &'a Package,
    require_metadata: bool,
) -> Result<Option<SourceView<'a>>, Error> {
    let mut budget = Budget::new(package).map_err(map_physical_error)?;
    let policy = if require_metadata {
        SelectionPolicy::Rewrite
    } else {
        SelectionPolicy::Read
    };
    let Some(selection) = soundtrack_physical::select_soundtrack(package, &mut budget, policy)
        .map_err(map_physical_error)?
    else {
        return Ok(None);
    };
    let catalog = soundtrack_physical::physical_source(package).map_err(map_physical_error)?;
    if !catalog.source_is_exact() && require_metadata {
        return Err(Error::UnsupportedSource);
    }
    let expected_locator = selection
        .soundtrack_component
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    let metadata = if require_metadata || !selection.media_ids().is_empty() {
        metadata_facts(
            package,
            catalog,
            expected_locator,
            selection.soundtrack_identifier,
        )?
    } else {
        MetadataFacts::default()
    };

    let mut records = Vec::new();
    records
        .try_reserve_exact(selection.item_count())
        .map_err(|_| Error::Allocation {
            amount: selection.item_count(),
        })?;
    let lineage = package_lineage(package);
    let mut occurrences = HashMap::<u64, usize>::new();
    occurrences
        .try_reserve(selection.media_ids().len())
        .map_err(|_| Error::Allocation {
            amount: selection.media_ids().len(),
        })?;
    for identifier in selection.media_ids() {
        let count = occurrences.entry(*identifier).or_insert(0);
        *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    for (position, identifier) in selection.media_ids().iter().copied().enumerate() {
        let data = metadata
            .data
            .iter()
            .find(|data| data.identifier == identifier)
            .ok_or(Error::InvalidSource)?;
        let materialized_length = data.materialized_length;
        let current_name = validate_existing_name(data)?;
        let full_name = format!("{DATA_PREFIX}{current_name}");
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == full_name)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque()
            || entry.data().len() != materialized_length
            || MediaType::from_extension(extension(current_name)) != MediaType::Audio
            || MediaType::from_bytes(entry.data()) != MediaType::Audio
        {
            return Err(Error::InvalidSource);
        }
        let digest = Sha1::digest(entry.data());
        if digest.as_slice() != data.digest.as_ref() {
            return Err(Error::InvalidSource);
        }
        let mut selected_owners = metadata
            .owners
            .iter()
            .filter(|owner| owner.selected && owner.data_identifier == identifier);
        let Some(selected_owner) = selected_owners.next() else {
            return Err(Error::InvalidSource);
        };
        if selected_owners.next().is_some()
            || selected_owner.count as usize != occurrences[&identifier]
        {
            return Err(Error::InvalidSource);
        }
        let owner_count = selected_owner.count;
        records.push(ResolvedItem {
            semantic: Item::new(
                lineage,
                Position::new(position),
                data.preferred_name.clone(),
                materialized_length,
            ),
            identifier,
            current_name: data.current_name.clone(),
            owner_count,
        });
    }
    // PackageMetadata stores a component's unique data-reference set, while
    // the soundtrack payload stores the ordered playback sequence and may
    // repeat an identifier. Require a canonical metadata set and prove every
    // selected occurrence resolves through it.
    for (index, identifier) in metadata.selected_data_references.iter().enumerate() {
        if metadata.selected_data_references[index + 1..]
            .iter()
            .any(|candidate| candidate == identifier)
        {
            return Err(Error::InvalidSource);
        }
    }
    if selection
        .media_ids()
        .iter()
        .any(|identifier| !metadata.selected_data_references.contains(identifier))
    {
        return Err(Error::InvalidSource);
    }
    Ok(Some(SourceView {
        selection,
        items: records,
        metadata,
    }))
}

fn validate_existing_name(data: &DataRecord) -> Result<&str, Error> {
    if data.unknown_fields
        || data.preferred_name.is_empty()
        || data.current_name.is_empty()
        || data.current_name.contains(['/', '\\', '\0'])
        || data.preferred_name.contains(['/', '\\', '\0'])
        || data.current_name.as_ref() == "."
        || data.current_name.as_ref() == ".."
        || MediaType::from_extension(extension(data.preferred_name.as_ref())) != MediaType::Audio
        || MediaType::from_extension(extension(data.current_name.as_ref())) != MediaType::Audio
    {
        return Err(Error::InvalidSource);
    }
    Ok(&data.current_name)
}

fn extension(name: &str) -> &str {
    name.rsplit_once('.').map_or("", |(_, extension)| extension)
}

fn package_lineage(package: &Package) -> [u8; SHA1_BYTES] {
    match &package.state.source {
        PhysicalSource::Package(catalog) => Sha1::digest(catalog.shared_source().as_ref()).into(),
        PhysicalSource::Semantic(_) => [0; SHA1_BYTES],
    }
}

fn map_physical_error(error: soundtrack_physical::Error) -> Error {
    match error {
        soundtrack_physical::Error::UnsupportedSource => Error::UnsupportedSource,
        soundtrack_physical::Error::SoundtrackNotFound => Error::SoundtrackNotFound,
        soundtrack_physical::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                soundtrack_physical::LimitKind::InputBytes => {
                    crate::soundtrack::items::LimitKind::InputBytes
                },
                soundtrack_physical::LimitKind::OutputBytes => {
                    crate::soundtrack::items::LimitKind::OutputBytes
                },
                soundtrack_physical::LimitKind::Entries => {
                    crate::soundtrack::items::LimitKind::Entries
                },
                soundtrack_physical::LimitKind::EntryBytes => {
                    crate::soundtrack::items::LimitKind::EntryBytes
                },
                soundtrack_physical::LimitKind::TotalBytes => {
                    crate::soundtrack::items::LimitKind::TotalBytes
                },
                soundtrack_physical::LimitKind::Items => crate::soundtrack::items::LimitKind::Items,
                soundtrack_physical::LimitKind::References => {
                    crate::soundtrack::items::LimitKind::References
                },
                soundtrack_physical::LimitKind::WireBytes => {
                    crate::soundtrack::items::LimitKind::WireBytes
                },
                soundtrack_physical::LimitKind::WireFields => {
                    crate::soundtrack::items::LimitKind::WireFields
                },
                soundtrack_physical::LimitKind::WireNesting => {
                    crate::soundtrack::items::LimitKind::WireNesting
                },
                soundtrack_physical::LimitKind::WireWork => {
                    crate::soundtrack::items::LimitKind::WireWork
                },
            },
            observed,
            maximum,
        },
        soundtrack_physical::Error::Allocation { amount } => Error::Allocation { amount },
        soundtrack_physical::Error::InvalidSource => Error::InvalidSource,
    }
}

fn operation_position(
    operation: &StagedOperation,
    before: &[ResolvedItem],
) -> Result<Position, Error> {
    match operation {
        StagedOperation::Add(_) => Ok(Position::from(before.len())),
        StagedOperation::Insert { position, .. } => {
            if position.get() > before.len() {
                return Err(Error::InsertionOutOfRange {
                    position: *position,
                    item_count: before.len(),
                });
            }
            Ok(*position)
        },
        StagedOperation::Replace { selector, .. } | StagedOperation::Remove { selector } => {
            resolve_selector(*selector, before)
        },
    }
}

fn resolve_selector(selector: ItemSelector, before: &[ResolvedItem]) -> Result<Position, Error> {
    let position = match selector {
        ItemSelector::Position(position) => position,
        ItemSelector::Handle(ItemHandle {
            lineage,
            occurrence,
        }) => {
            if lineage == [0; SHA1_BYTES]
                || before
                    .first()
                    .map(|item| item.semantic.handle().position())
                    .is_none()
            {
                return Err(Error::ItemHandleConflict);
            }
            // The caller's lineage is checked by the edit commit below.  This
            // branch only resolves the captured semantic occurrence.
            let _ = lineage;
            Position::new(occurrence)
        },
    };
    if position.get() >= before.len() {
        return Err(Error::SourcePositionNotFound { position });
    }
    Ok(position)
}

fn project_ids(
    before: &[ResolvedItem],
    operation: &StagedOperation,
    position: Position,
    new_identifier: Option<u64>,
) -> Result<Vec<u64>, Error> {
    let mut ids = Vec::new();
    let capacity = before
        .len()
        .checked_add(usize::from(new_identifier.is_some()))
        .ok_or(Error::InvalidSource)?;
    ids.try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: capacity })?;
    match operation {
        StagedOperation::Add(_) | StagedOperation::Insert { .. } => {
            let index = position.get();
            ids.extend(before[..index].iter().map(|item| item.identifier));
            ids.push(new_identifier.ok_or(Error::InvalidSource)?);
            ids.extend(before[index..].iter().map(|item| item.identifier));
        },
        StagedOperation::Replace { .. } => {
            for (index, item) in before.iter().enumerate() {
                ids.push(if index == position.get() {
                    new_identifier.ok_or(Error::InvalidSource)?
                } else {
                    item.identifier
                });
            }
        },
        StagedOperation::Remove { .. } => {
            ids.extend(
                before
                    .iter()
                    .enumerate()
                    .filter(|(index, _)| *index != position.get())
                    .map(|(_, item)| item.identifier),
            );
        },
    }
    Ok(ids)
}

fn allocate_identifier(catalog: &SourceCatalog, metadata: &MetadataFacts) -> Result<u64, Error> {
    // Data identifiers and object identifiers are independent native
    // namespaces. Keynote keeps media IDs near the existing DataInfo range;
    // allocating above the package object watermark can make TSPDataManager
    // reject the package even when every reference is otherwise consistent.
    let mut candidate = metadata
        .data
        .iter()
        .map(|data| data.identifier)
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    if candidate == 0 {
        return Err(Error::InvalidSource);
    }

    // Retain a package-wide collision guard: the next media-space value must
    // not alias any object or reference identifier already present in an IWA
    // header, even though those identifiers normally occupy another range.
    let mut used = HashSet::new();
    used.try_reserve(
        catalog
            .components()
            .iter()
            .map(|component| component.archive().objects.len())
            .sum::<usize>()
            .saturating_add(metadata.data.len()),
    )
    .map_err(|_| Error::Allocation { amount: 1 })?;
    for component in catalog.components().iter() {
        for object in &component.archive().objects {
            if let Some(identifier) = object.archive_info.identifier {
                used.insert(identifier);
            }
            for info in &object.archive_info.message_infos {
                for identifier in info.data_references.iter().chain(&info.object_references) {
                    used.insert(*identifier);
                }
                for field in &info.field_infos {
                    for identifier in field.data_references.iter().chain(&field.object_references) {
                        used.insert(*identifier);
                    }
                }
            }
        }
    }
    for data in &metadata.data {
        used.insert(data.identifier);
    }
    for identifier in &metadata.selected_data_references {
        used.insert(*identifier);
    }
    while used.contains(&candidate) {
        candidate = candidate.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if candidate == 0 {
        return Err(Error::InvalidSource);
    }
    Ok(candidate)
}

fn existing_data_for_digest<'a>(
    metadata: &'a MetadataFacts,
    digest: &[u8; SHA1_BYTES],
) -> Result<Option<&'a DataRecord>, Error> {
    let mut matching = metadata.data.iter().filter(|data| &data.digest == digest);
    let first = matching.next();
    if matching.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(first)
}

fn validate_reusable_data<'a>(
    catalog: &'a SourceCatalog,
    data: &DataRecord,
    audio: &AudioSource,
) -> Result<&'a [u8], Error> {
    let current_name = validate_existing_name(data)?;
    let full_name = format!("{DATA_PREFIX}{current_name}");
    let mut matching = catalog
        .package()
        .iter()
        .filter(|entry| entry.name() == full_name);
    let entry = matching.next().ok_or(Error::InvalidSource)?;
    if matching.next().is_some()
        || entry.is_opaque()
        || data.materialized_length != audio.byte_length()
        || entry.data() != audio.bytes()
        || MediaType::from_bytes(entry.data()) != MediaType::Audio
    {
        return Err(Error::InvalidSource);
    }
    Ok(entry.data())
}

fn generated_name(
    source_name: &str,
    identifier: u64,
    catalog: &SourceCatalog,
) -> Result<String, Error> {
    let (stem, extension) = source_name.rsplit_once('.').ok_or(Error::InvalidSource)?;
    let mut candidate = format!("{stem}-{identifier}.{extension}");
    if candidate.len() > crate::soundtrack::items::MAX_FILENAME_BYTES {
        candidate = format!("litchi-{identifier}.{extension}");
    }
    for attempt in 0_u64..1_024 {
        let name = if attempt == 0 {
            candidate.clone()
        } else {
            format!("{stem}-{identifier}-{attempt}.{extension}")
        };
        if name.len() > crate::soundtrack::items::MAX_FILENAME_BYTES {
            continue;
        }
        let full = format!("{DATA_PREFIX}{name}");
        if !catalog.package().iter().any(|entry| entry.name() == full) {
            return Ok(name);
        }
    }
    Err(Error::InvalidSource)
}

fn encode_reference_field(output: &mut Vec<u8>, identifier: u64) {
    let payload_length = encoded_len(8) + encoded_len(identifier);
    encode_varint_into(output, u64::from(SOUNDTRACK_MEDIA_FIELD << 3) | 2);
    encode_varint_into(output, u64::try_from(payload_length).unwrap_or(u64::MAX));
    encode_varint_into(output, 8);
    encode_varint_into(output, identifier);
}

fn rewrite_soundtrack_payload(
    source: &[u8],
    before_len: usize,
    position: Position,
    operation: OperationKind,
    new_identifier: Option<u64>,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(|_| Error::InvalidSource)?;
    let new_ref_len = new_identifier.map_or(0, |identifier| {
        encoded_len(u64::from(SOUNDTRACK_MEDIA_FIELD << 3) | 2)
            + encoded_len(
                u64::try_from(encoded_len(8) + encoded_len(identifier)).unwrap_or(u64::MAX),
            )
            + encoded_len(8)
            + encoded_len(identifier)
    });
    // Reserve an upper bound before rewriting.  A removal can only make the
    // payload smaller, while add/insert/replace may append one encoded
    // reference; the final wire-limit check below remains authoritative.
    let output_len = source
        .len()
        .checked_add(new_ref_len)
        .ok_or(Error::InvalidSource)?;
    if output_len > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: crate::soundtrack::items::LimitKind::OutputBytes,
            observed: output_len as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|_| Error::Allocation { amount: output_len })?;
    let mut seen = 0usize;
    let mut inserted = false;
    for field in view.fields() {
        if field.number() != SOUNDTRACK_MEDIA_FIELD {
            output.extend_from_slice(field.raw());
            continue;
        }
        if seen >= before_len {
            return Err(Error::InvalidSource);
        }
        if matches!(operation, OperationKind::Insert | OperationKind::Add)
            && !inserted
            && seen == position.get()
        {
            encode_reference_field(&mut output, new_identifier.ok_or(Error::InvalidSource)?);
            inserted = true;
        }
        let selected = seen == position.get();
        match operation {
            OperationKind::Replace if selected => {
                encode_reference_field(&mut output, new_identifier.ok_or(Error::InvalidSource)?);
            },
            OperationKind::Remove if selected => {},
            _ => output.extend_from_slice(field.raw()),
        }
        seen = seen.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if seen != before_len {
        return Err(Error::InvalidSource);
    }
    if matches!(operation, OperationKind::Insert | OperationKind::Add) && !inserted {
        encode_reference_field(&mut output, new_identifier.ok_or(Error::InvalidSource)?);
    }
    if output.len() > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: crate::soundtrack::items::LimitKind::OutputBytes,
            observed: output.len() as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    Ok(output)
}

fn rewrite_metadata_payload(
    package: &Package,
    catalog: &SourceCatalog,
    batch: metadata_codec::MediaRewriteBatch<'_>,
) -> Result<Vec<u8>, Error> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .limits()
            .snappy_limits()
            .map_err(|_| Error::InvalidSource)?,
    )
    .map_err(|_| Error::InvalidSource)?;
    let archive_limits = package
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| Error::InvalidSource)?;
    let mut selected = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                if selected
                    .replace((object_index, message_index, message.data.as_slice()))
                    .is_some()
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let (object_index, message_index, payload) = selected.ok_or(Error::InvalidSource)?;
    let options = metadata_codec::DecodeOptions::for_source(payload);
    let rewritten = metadata_codec::rewrite_package_metadata_media(payload, batch, options)
        .map_err(|_| Error::InvalidSource)?
        .into_bytes();
    let object = archive
        .objects
        .get_mut(object_index)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let serialized = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    SnappyStream::compress(&serialized).map_err(|_| Error::InvalidSource)
}

fn rewrite_package(
    source: &Package,
    view: &SourceView<'_>,
    operation: &StagedOperation,
) -> Result<(Package, Box<[Item]>), Error> {
    let catalog = soundtrack_physical::physical_source(source).map_err(map_physical_error)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let operation_kind = operation.kind();
    let position = operation_position(operation, &view.items)?;
    let old = match operation {
        StagedOperation::Replace { .. } | StagedOperation::Remove { .. } => {
            view.items.get(position.get())
        },
        _ => None,
    };
    let (source_audio, new_identifier, current_name, digest, reuse_existing) = match operation {
        StagedOperation::Add(audio)
        | StagedOperation::Insert { source: audio, .. }
        | StagedOperation::Replace { source: audio, .. } => {
            let digest: [u8; SHA1_BYTES] = Sha1::digest(audio.bytes()).into();
            if let Some(data) = existing_data_for_digest(&view.metadata, &digest)? {
                validate_reusable_data(catalog, data, audio)?;
                (Some(audio), Some(data.identifier), None, None, true)
            } else {
                let identifier = allocate_identifier(catalog, &view.metadata)?;
                let name = generated_name(audio.filename(), identifier, catalog)?;
                (
                    Some(audio),
                    Some(identifier),
                    Some(name),
                    Some(digest),
                    false,
                )
            }
        },
        StagedOperation::Remove { .. } => (None, None, None, None, false),
    };
    let before_ids = view
        .items
        .iter()
        .map(|item| item.identifier)
        .collect::<Vec<_>>();
    let after_ids = project_ids(view.items.as_slice(), operation, position, new_identifier)?;
    let wire_limits = source.wire_limits().map_err(|_| Error::InvalidSource)?;
    let rewritten_payload = rewrite_soundtrack_payload(
        view.selection.soundtrack_payload,
        before_ids.len(),
        position,
        operation_kind,
        new_identifier,
        wire_limits,
    )?;

    let archive_entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == view.selection.soundtrack_component)
        .ok_or(Error::InvalidSource)?;
    if archive_entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let stream = SnappyStream::decompress_with_limits(
        archive_entry.data(),
        source
            .limits()
            .snappy_limits()
            .map_err(|_| Error::InvalidSource)?,
    )
    .map_err(|_| Error::InvalidSource)?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| Error::InvalidSource)?;
    let object = archive
        .object_mut(view.selection.soundtrack_identifier)
        .ok_or(Error::InvalidSource)?;
    let matching_fields = object
        .archive_info
        .message_infos
        .get(view.selection.soundtrack_message_index)
        .ok_or(Error::InvalidSource)?
        .field_infos
        .iter()
        .enumerate()
        .filter_map(|(index, field)| {
            (field.path.as_slice() == [SOUNDTRACK_MEDIA_FIELD]
                && field.data_references == before_ids)
                .then_some(index)
        })
        .collect::<Vec<_>>();
    let replacement = RawMessage {
        type_: SOUNDTRACK_MESSAGE_TYPE,
        data: rewritten_payload,
    };
    match matching_fields.as_slice() {
        [field_index] => object
            .replace_message_transitioning_data_references_preserving_header_with_limits(
                view.selection.soundtrack_message_index,
                replacement,
                DataReferenceTransition {
                    aggregate_before: &before_ids,
                    aggregate_after: &after_ids,
                    fields: &[FieldDataReferenceTransition {
                        field_info_index: *field_index,
                        expected_path: &[SOUNDTRACK_MEDIA_FIELD],
                        before: &before_ids,
                        after: &after_ids,
                    }],
                },
                archive_limits,
            ),
        [] => object.replace_message_transitioning_data_references_preserving_header_with_limits(
            view.selection.soundtrack_message_index,
            replacement,
            DataReferenceTransition {
                aggregate_before: &before_ids,
                aggregate_after: &after_ids,
                fields: &[],
            },
            archive_limits,
        ),
        _ => return Err(Error::InvalidSource),
    }
    .map_err(|_| Error::InvalidSource)?;
    let serialized_soundtrack = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    let compressed_soundtrack =
        SnappyStream::compress(&serialized_soundtrack).map_err(|_| Error::InvalidSource)?;

    let mut data_additions = Vec::new();
    let mut owner_additions = Vec::new();
    let mut owner_removals = Vec::new();
    let mut owner_updates = Vec::new();
    let locator = view
        .metadata
        .selected_component_locator
        .as_deref()
        .ok_or(Error::InvalidSource)?;
    let component_identifier = view
        .metadata
        .selected_component_identifier
        .ok_or(Error::InvalidSource)?;
    if reuse_existing && old.is_none_or(|old| old.identifier != new_identifier.unwrap_or(0)) {
        let identifier = new_identifier.ok_or(Error::InvalidSource)?;
        let owner = view.metadata.owners.iter().find(|owner| {
            owner.selected
                && owner.data_identifier == identifier
                && owner.object_identifier == view.selection.soundtrack_identifier
        });
        if let Some(owner) = owner {
            owner_updates.push(metadata_codec::DataReferenceOwnerCountUpdate::new(
                metadata_codec::ComponentSelector::new(component_identifier, locator),
                identifier,
                view.selection.soundtrack_identifier,
                owner.count,
                owner.count.checked_add(1).ok_or(Error::InvalidSource)?,
            ));
        } else {
            owner_additions.push(metadata_codec::DataReferenceOwnerAddition::new(
                metadata_codec::ComponentSelector::new(component_identifier, locator),
                identifier,
                view.selection.soundtrack_identifier,
                1,
            ));
        }
    }
    if let (Some(audio), Some(identifier), Some(name), Some(digest)) = (
        source_audio,
        new_identifier,
        current_name.as_deref(),
        digest.as_ref(),
    ) {
        data_additions.push(
            metadata_codec::DataInfoAddition::new(identifier, digest, audio.filename())
                .with_file_name(name)
                .with_materialized_length(
                    u64::try_from(audio.byte_length()).map_err(|_| Error::InvalidSource)?,
                ),
        );
        owner_additions.push(metadata_codec::DataReferenceOwnerAddition::new(
            metadata_codec::ComponentSelector::new(component_identifier, locator),
            identifier,
            view.selection.soundtrack_identifier,
            1,
        ));
    }
    if let Some(old) = old
        && old.identifier != new_identifier.unwrap_or(0)
    {
        let selector = metadata_codec::ComponentSelector::new(component_identifier, locator);
        if old.owner_count > 1 {
            owner_updates.push(metadata_codec::DataReferenceOwnerCountUpdate::new(
                selector,
                old.identifier,
                view.selection.soundtrack_identifier,
                old.owner_count,
                old.owner_count - 1,
            ));
        } else {
            owner_removals.push(metadata_codec::DataReferenceOwnerRemoval::new(
                selector,
                old.identifier,
                view.selection.soundtrack_identifier,
                old.owner_count,
            ));
        }
    }
    let batch = metadata_codec::MediaRewriteBatch::new(
        &data_additions,
        &[],
        &owner_additions,
        &owner_removals,
    )
    .with_owner_count_updates(&owner_updates);
    let compressed_metadata = rewrite_metadata_payload(source, catalog, batch)?;

    let insertion_name = current_name
        .as_deref()
        .map(|name| format!("{DATA_PREFIX}{name}"));
    let mut insertions = Vec::new();
    if !reuse_existing && let (Some(audio), Some(name)) = (source_audio, insertion_name.as_deref())
    {
        insertions.push(EntryInsertion::new(name, audio.bytes()));
    }
    let insertion_refs = insertions;
    let edits = [
        EntryEdit::new(view.selection.soundtrack_component, &compressed_soundtrack),
        EntryEdit::new(METADATA_COMPONENT, &compressed_metadata),
    ];
    let output = catalog
        .package()
        .reassemble_with_changes_to_bytes(&insertion_refs, &edits, &[], catalog.limits())
        .map_err(|_| Error::InvalidSource)?;
    let target: Arc<[u8]> = output.into();
    let candidate = Package::from_source_with_options(target, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_view = source_view(&candidate, true)?.ok_or(Error::Verification)?;
    let candidate_items = candidate_view.semantic_items()?;
    let expected_items = expected_items_after(
        source,
        view,
        operation,
        position,
        source_audio,
        new_identifier,
    )?;
    if !semantic_items_equal(&candidate_items, &expected_items) {
        return Err(Error::Verification);
    }
    // Ensure the changed target can itself be read after reopening, and that
    // every failed path above leaves the source package untouched by design.
    Ok((candidate, candidate_items))
}

fn expected_items_after(
    source: &Package,
    view: &SourceView<'_>,
    operation: &StagedOperation,
    position: Position,
    audio: Option<&AudioSource>,
    identifier: Option<u64>,
) -> Result<Box<[Item]>, Error> {
    let lineage = package_lineage(source);
    let mut items = Vec::new();
    items
        .try_reserve_exact(
            view.items
                .len()
                .checked_add(usize::from(audio.is_some()))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(|_| Error::Allocation {
            amount: view.items.len(),
        })?;
    let canonical_filename = identifier
        .and_then(|identifier| {
            view.metadata
                .data
                .iter()
                .find(|data| data.identifier == identifier)
        })
        .map(|data| data.preferred_name.as_ref())
        .or_else(|| audio.map(AudioSource::filename))
        .unwrap_or("");
    let new_item = || {
        Item::new(
            lineage,
            position,
            canonical_filename,
            audio.map_or(0, AudioSource::byte_length),
        )
    };
    match operation {
        StagedOperation::Add(_) | StagedOperation::Insert { .. } => {
            let index = position.get();
            items.extend(view.items[..index].iter().map(|item| item.semantic.clone()));
            let mut inserted = new_item();
            inserted.handle = ItemHandle::new(lineage, index);
            items.push(inserted);
            items.extend(
                view.items[index..]
                    .iter()
                    .enumerate()
                    .map(|(offset, item)| {
                        Item::new(
                            lineage,
                            Position::new(index + offset + 1),
                            item.semantic.filename(),
                            item.semantic.byte_length(),
                        )
                    }),
            );
        },
        StagedOperation::Replace { .. } => {
            for (index_at, item) in view.items.iter().enumerate() {
                if index_at == position.get() {
                    let mut replacement = new_item();
                    replacement.handle = ItemHandle::new(lineage, index_at);
                    items.push(replacement);
                } else {
                    items.push(Item::new(
                        lineage,
                        Position::new(index_at),
                        item.semantic.filename(),
                        item.semantic.byte_length(),
                    ));
                }
            }
        },
        StagedOperation::Remove { .. } => {
            for (index_at, item) in view.items.iter().enumerate() {
                if index_at != position.get() {
                    let next = if index_at > position.get() {
                        index_at - 1
                    } else {
                        index_at
                    };
                    items.push(Item::new(
                        lineage,
                        Position::new(next),
                        item.semantic.filename(),
                        item.semantic.byte_length(),
                    ));
                }
            }
        },
    }
    let _ = identifier;
    Ok(items.into_boxed_slice())
}

fn semantic_items_equal(left: &[Item], right: &[Item]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| {
            left.position() == right.position()
                && left.filename() == right.filename()
                && left.byte_length() == right.byte_length()
        })
}

fn physical_source_bytes(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => {
            Ok(catalog.shared_source())
        },
        PhysicalSource::Package(_) => Err(Error::UnsupportedSource),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

impl Package {
    /// List the existing rooted soundtrack media collection.  `None` means no
    /// soundtrack object is rooted from the show; `Some(empty)` means an
    /// existing soundtrack object has no media references.
    pub fn soundtrack_items(&self) -> Result<Option<Box<[Item]>>, Error> {
        let Some(view) = source_view(self, false)? else {
            return Ok(None);
        };
        Ok(Some(view.semantic_items()?))
    }

    /// Begin one immutable soundtrack-item lifecycle edit.
    pub fn edit_soundtrack_items(&self) -> Result<Edit<'_>, Error> {
        let Some(view) = source_view(self, true)? else {
            return Err(Error::SoundtrackNotFound);
        };
        Ok(Edit::new_with_item_count(self, Some(view.items.len())))
    }

    /// Apply an exact-source soundtrack-item patch and reopen its target.
    pub fn apply_soundtrack_items(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_bytes = physical_source_bytes(self)?;
        if !patch.artifacts.authorizes_source(&source_bytes) {
            return Err(Error::PatchConflict);
        }
        let Some(view) = source_view(self, true)? else {
            return Err(Error::PatchConflict);
        };
        if !semantic_items_equal(&view.semantic_items()?, &patch.before) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let target = patch.artifacts.target();
        let candidate = Package::from_source_with_options(target, self.state.options)
            .map_err(|_| Error::PatchConflict)?;
        let Some(candidate_view) = source_view(&candidate, true)? else {
            return Err(Error::PatchConflict);
        };
        let candidate_items = candidate_view.semantic_items()?;
        if !semantic_items_equal(&candidate_items, &patch.after) {
            return Err(Error::PatchConflict);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(2),
        })
    }
}

impl Edit<'_> {
    /// Validate, publish, and reopen the staged one-operation transaction.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self.operation.ok_or(Error::NoStagedOperation)?;
        let Some(view) = source_view(self.source, true)? else {
            return Err(Error::SoundtrackNotFound);
        };
        if let StagedOperation::Replace {
            selector,
            source: _,
        } = &operation
            && let ItemSelector::Handle(handle) = selector
            && handle.lineage != package_lineage(self.source)
        {
            return Err(Error::ItemHandleConflict);
        }
        if let StagedOperation::Remove { selector } = &operation
            && let ItemSelector::Handle(handle) = selector
            && handle.lineage != package_lineage(self.source)
        {
            return Err(Error::ItemHandleConflict);
        }
        let position = operation_position(&operation, &view.items)?;
        if let StagedOperation::Replace { source, .. } = &operation {
            let current = view
                .items
                .get(position.get())
                .ok_or(Error::SourcePositionNotFound { position })?;
            let catalog =
                soundtrack_physical::physical_source(self.source).map_err(map_physical_error)?;
            let data_name = format!("{DATA_PREFIX}{}", current.current_name);
            if source.filename() == current.semantic.filename()
                && catalog
                    .package()
                    .iter()
                    .find(|entry| entry.name() == data_name)
                    .is_some_and(|entry| entry.data() == source.bytes())
            {
                let source_bytes = physical_source_bytes(self.source)?;
                let before = view.semantic_items()?;
                let patch = Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    before: before.clone(),
                    after: before,
                    operation: OperationKind::Replace,
                    inverse_operation: OperationKind::Replace,
                };
                return Ok(Commit {
                    package: self.source.snapshot(),
                    patch,
                    diagnostics: Diagnostics::unchanged(),
                });
            }
        }
        let source_bytes = physical_source_bytes(self.source)?;
        let (package, after) = rewrite_package(self.source, &view, &operation)?;
        let target = physical_source_bytes(&package)?;
        let before = view.semantic_items()?;
        let patch = Patch {
            artifacts: ExactArtifacts::new(source_bytes, target),
            before,
            after,
            operation: operation.kind(),
            inverse_operation: match operation.kind() {
                OperationKind::Add | OperationKind::Insert => OperationKind::Remove,
                OperationKind::Replace => OperationKind::Replace,
                OperationKind::Remove if position.get() + 1 == view.items.len() => {
                    OperationKind::Add
                },
                OperationKind::Remove => OperationKind::Insert,
            },
        };
        let is_noop = patch.is_noop();
        Ok(Commit {
            package: if is_noop {
                self.source.snapshot()
            } else {
                package
            },
            patch,
            diagnostics: if is_noop {
                Diagnostics::unchanged()
            } else {
                Diagnostics::published(2)
            },
        })
    }
}
