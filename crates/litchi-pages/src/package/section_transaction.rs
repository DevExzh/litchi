//! Shared exact-source machinery for settings stored on a Pages section.

use std::num::NonZeroU64;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};

use super::{
    MAX_SECTIONS, NativeSectionReference, Package, PackageError, SECTION_MESSAGE_TYPE,
    decode_body_storage, effective_text_limit, find_object, native_section_references, page_layout,
    root_references_with_limits, validate_section_table_wire_with_limits,
};
use crate::{
    SectionSelector,
    section::settings::{DependencyKind, Error, LimitKind, Path},
};

pub(super) const TEMPLATE_MESSAGE_TYPE: u32 = 10_143;
pub(super) const STORAGE_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Target {
    pub(super) position: Position,
    pub(super) identifier: NonZeroU64,
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct RewriteStats {
    pub(super) touched_components: usize,
    pub(super) deleted_previews: usize,
    pub(super) source_layout_state: Option<u64>,
    pub(super) target_layout_state: Option<u64>,
    pub(super) source_preview_count: usize,
    pub(super) target_preview_count: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(super) struct Usage {
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) references: usize,
    pub(super) transaction_work: usize,
    pub(super) output_allocations: usize,
    pub(super) candidate_reopens: usize,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct TransactionBudget {
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_transaction_work: usize,
    reserved_transaction_work: usize,
    usage: Usage,
}

impl TransactionBudget {
    pub(super) fn new(package: &Package) -> Result<Self, Error> {
        let physical = package.state.source.limits();
        let archive = physical
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let components = physical.max_entries().max(1);
        Ok(Self {
            maximum_fields: archive.max_header_fields().saturating_mul(components),
            maximum_work: physical.max_iwa_stream_bytes().saturating_mul(16),
            maximum_references: archive.max_metadata_items().saturating_mul(components),
            maximum_transaction_work: usize::try_from(physical.max_total_bytes())
                .unwrap_or(usize::MAX)
                .saturating_mul(32),
            reserved_transaction_work: 0,
            usage: Usage::default(),
        })
    }

    pub(super) const fn remaining_fields(self) -> usize {
        self.maximum_fields.saturating_sub(self.usage.fields)
    }

    pub(super) const fn remaining_work(self) -> usize {
        self.maximum_work.saturating_sub(self.usage.work)
    }

    pub(super) fn charge_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge(
            &mut self.usage.fields,
            amount,
            self.maximum_fields,
            LimitKind::WireFields,
            path,
        )
    }

    pub(super) fn charge_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge(
            &mut self.usage.work,
            amount,
            self.maximum_work,
            LimitKind::WireWork,
            path,
        )
    }

    pub(super) fn charge_references(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge(
            &mut self.usage.references,
            amount,
            self.maximum_references,
            LimitKind::References,
            path,
        )
    }

    pub(super) fn charge_transaction_work(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        let observed = self
            .usage
            .transaction_work
            .checked_add(amount)
            .unwrap_or(usize::MAX);
        if observed > self.maximum_transaction_work {
            return Err(Error::LimitExceeded {
                path,
                kind: LimitKind::TransactionWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.maximum_transaction_work),
            });
        }
        self.usage.transaction_work = observed;
        self.reserved_transaction_work = self.reserved_transaction_work.saturating_sub(amount);
        Ok(())
    }

    pub(super) fn reserve_transaction_work(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        let observed = self
            .usage
            .transaction_work
            .checked_add(self.reserved_transaction_work)
            .and_then(|value| value.checked_add(amount))
            .unwrap_or(usize::MAX);
        if observed > self.maximum_transaction_work {
            return Err(Error::LimitExceeded {
                path,
                kind: LimitKind::TransactionWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.maximum_transaction_work),
            });
        }
        self.reserved_transaction_work = self
            .reserved_transaction_work
            .checked_add(amount)
            .ok_or(Error::InvalidSource { path })?;
        Ok(())
    }

    pub(super) fn settle_transaction_reservation(&mut self) {
        self.reserved_transaction_work = 0;
    }

    #[cfg(test)]
    pub(super) const fn usage(self) -> Usage {
        self.usage
    }

    #[cfg(test)]
    pub(super) fn with_transaction_limit(package: &Package, maximum: usize) -> Result<Self, Error> {
        let mut budget = Self::new(package)?;
        budget.maximum_transaction_work = maximum;
        Ok(budget)
    }

    fn note_output_allocation(&mut self) {
        self.usage.output_allocations = self.usage.output_allocations.saturating_add(1);
    }

    fn note_candidate_reopen(&mut self) {
        self.usage.candidate_reopens = self.usage.candidate_reopens.saturating_add(1);
    }
}

fn charge(
    current: &mut usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
    path: Path,
) -> Result<(), Error> {
    let observed = current.checked_add(amount).unwrap_or(usize::MAX);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            path,
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        });
    }
    *current = observed;
    Ok(())
}

pub(super) fn resolve_position<'selector>(
    package: &Package,
    selector: impl Into<SectionSelector<'selector>>,
) -> Result<Position, Error> {
    let selector = selector.into();
    let selected = package
        .semantic_document()
        .select_section(selector)
        .map_err(|selection_error| match selection_error {
            crate::SelectorError::AmbiguousSectionName {
                first, duplicate, ..
            } => Error::AmbiguousSelector {
                first: Position::new(first),
                duplicate: Position::new(duplicate),
            },
        })?
        .ok_or(match selector {
            SectionSelector::Name(_) => Error::NameNotFound,
            SectionSelector::Position(position) => Error::PositionNotFound { position },
        })?;
    Ok(Position::new(selected.index()))
}

pub(super) fn resolve_target(
    package: &Package,
    position: Position,
    budget: &mut TransactionBudget,
) -> Result<Target, Error> {
    let path = Path::section(position);
    let components = package.state.source.components();
    let limits = package.state.source.limits();
    budget.charge_transaction_work(components.len(), Path::Package)?;
    let root = root_references_with_limits(components, limits).map_err(map_package_error)?;
    let body_identifier = root.body.ok_or(Error::UnsupportedSource {
        path: Path::Package,
    })?;
    let body = find_object(components, body_identifier.get()).ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    budget.charge_transaction_work(body.messages.len(), Path::Package)?;
    let max_text_bytes = effective_text_limit(limits);
    let (native, payload) = decode_body_storage(
        &body.messages,
        body_identifier,
        MAX_SECTIONS,
        max_text_bytes,
        limits,
    )
    .map_err(map_package_error)?;
    budget.charge_work(payload.len(), Path::Package)?;
    validate_section_table_wire_with_limits(payload, &native, body_identifier, limits)
        .map_err(map_package_error)?;
    let references = native_section_references(&native, root.initial_section, MAX_SECTIONS)
        .map_err(map_package_error)?;
    budget.charge_references(references.len(), Path::Package)?;
    let identifier = references
        .get(position.get())
        .copied()
        .map(|NativeSectionReference { identifier, .. }| identifier)
        .ok_or(Error::PositionNotFound { position })?;

    let mut found = None;
    for (component_index, component) in components.iter().enumerate() {
        budget.charge_transaction_work(component.archive().objects.len(), Path::Package)?;
        let Some((object_index, object)) = component
            .archive()
            .objects
            .iter()
            .enumerate()
            .find(|(_index, object)| object.archive_info.identifier == Some(identifier.get()))
        else {
            continue;
        };
        if found.is_some() {
            return Err(Error::InvalidSource { path });
        }
        let (message_index, message) = unique_message(object, SECTION_MESSAGE_TYPE, path)?;
        validate_selected_metadata(object, message_index, path)?;
        budget.charge_fields(
            message_field_count(&message.data, wire_limits(package)?)?,
            path,
        )?;
        budget.charge_work(message.data.len(), path)?;
        found = Some(Target {
            position,
            identifier,
            component_index,
            object_index,
            message_index,
        });
    }
    found.ok_or(Error::InvalidSource { path })
}

fn message_field_count(payload: &[u8], limits: WireLimits) -> Result<usize, Error> {
    WireView::parse_with_limits(payload, limits)
        .map(|view| view.len())
        .map_err(map_wire_error)
}

pub(super) fn selected_payload(package: &Package, target: Target) -> Result<&[u8], Error> {
    let path = Path::section(target.position);
    let component = package
        .state
        .source
        .components()
        .iter()
        .nth(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .filter(|object| object.archive_info.identifier == Some(target.identifier.get()))
        .ok_or(Error::InvalidSource { path })?;
    let message = object
        .messages
        .get(target.message_index)
        .filter(|message| message.type_ == SECTION_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource { path })?;
    Ok(&message.data)
}

pub(super) fn rewrite_package(
    source: &Package,
    target: Target,
    rewritten_payload: Vec<u8>,
    invalidate_layout: bool,
    budget: &mut TransactionBudget,
) -> Result<(Package, RewriteStats), Error> {
    let path = Path::section(target.position);
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .iter()
        .nth(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let component_name = owned(component.name())?;
    let source_layout_state = if invalidate_layout {
        page_layout::view_state_layout_identifier(source).map_err(map_page_layout_error)?
    } else {
        None
    };
    let source_preview_count = page_layout::preview_count(source);
    let view_location = if invalidate_layout {
        page_layout::view_state_location(source).map_err(map_page_layout_error)?
    } else {
        None
    };
    let shared_view_component = view_location
        .as_ref()
        .is_some_and(|location| location.component_name == component_name);

    let (mut archive, archive_limits) = editable_archive(source, &component_name)?;
    {
        let object = archive
            .object_mut(target.identifier.get())
            .ok_or(Error::InvalidSource { path })?;
        let message = object
            .messages
            .get(target.message_index)
            .filter(|message| message.type_ == SECTION_MESSAGE_TYPE)
            .ok_or(Error::InvalidSource { path })?;
        if message.data.as_slice() == rewritten_payload.as_slice() {
            return Err(Error::Verification { path });
        }
        object
            .replace_message_preserving_header_with_limits(
                target.message_index,
                RawMessage {
                    type_: SECTION_MESSAGE_TYPE,
                    data: rewritten_payload,
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }
    if shared_view_component && let Some(location) = view_location.as_ref() {
        page_layout::invalidate_view_state_in_archive(
            source,
            &mut archive,
            location,
            archive_limits,
        )
        .map_err(map_page_layout_error)?;
    }
    let section_compressed = compress_archive(archive, archive_limits)?;

    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(2)
        .map_err(|_error| Error::Allocation { amount: 2 })?;
    compressed.push((component_name, section_compressed));
    if !shared_view_component && let Some(location) = view_location.as_ref() {
        let (mut view_archive, view_limits) = editable_archive(source, &location.component_name)?;
        page_layout::invalidate_view_state_in_archive(
            source,
            &mut view_archive,
            location,
            view_limits,
        )
        .map_err(map_page_layout_error)?;
        compressed.push((
            location.component_name.clone(),
            compress_archive(view_archive, view_limits)?,
        ));
    }

    let edits: Vec<_> = compressed
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect();
    let deletions: Vec<_> = if invalidate_layout {
        page_layout::PREVIEW_ENTRY_NAMES
            .into_iter()
            .filter(|name| catalog.package().iter().any(|entry| entry.name() == *name))
            .collect()
    } else {
        Vec::new()
    };
    let compressed_len = compressed
        .iter()
        .try_fold(0usize, |total, (_, bytes)| total.checked_add(bytes.len()))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (package_bound, reassembly_work) =
        reassembly_cost(source.source_bytes().len(), compressed_len)?;
    budget.charge_transaction_work(reassembly_work, Path::Package)?;
    let post_output_work = package_bound
        .saturating_mul(2)
        .saturating_add(128)
        .saturating_add(source.stats().total_objects().saturating_mul(2))
        .saturating_add(source.state.source.components().len().saturating_mul(4));
    budget.reserve_transaction_work(post_output_work, Path::Package)?;
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deletions, catalog.limits())
        .map_err(map_archive_error)?;
    budget.note_output_allocation();
    if output.len() > package_bound {
        return Err(Error::Verification {
            path: Path::Package,
        });
    }
    budget.charge_transaction_work(
        output.len().saturating_mul(2).saturating_add(128),
        Path::Package,
    )?;
    budget.note_candidate_reopen();
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    let target_layout_state = if invalidate_layout {
        page_layout::view_state_layout_identifier(&candidate).map_err(map_page_layout_error)?
    } else {
        None
    };
    let target_preview_count = page_layout::preview_count(&candidate);
    if invalidate_layout && (target_layout_state.is_some() || target_preview_count != 0) {
        return Err(Error::Verification { path });
    }
    Ok((
        candidate,
        RewriteStats {
            touched_components: compressed.len(),
            deleted_previews: deletions.len(),
            source_layout_state,
            target_layout_state,
            source_preview_count,
            target_preview_count,
        },
    ))
}

pub(super) fn reassembly_cost(
    source_len: usize,
    compressed_len: usize,
) -> Result<(usize, usize), Error> {
    let package_bound = source_len
        .checked_add(compressed_len.saturating_mul(2))
        .and_then(|value| value.checked_add(1_024))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    Ok((package_bound, package_bound.saturating_mul(3)))
}

pub(super) fn editable_archive(
    package: &Package,
    component_name: &str,
) -> Result<(Archive, litchi_iwa_core::Limits), Error> {
    let source = &package.state.source;
    let entry = source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let limits = source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source.limits().snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    Ok((archive, limits))
}

pub(super) fn compress_archive(
    archive: Archive,
    limits: litchi_iwa_core::Limits,
) -> Result<Vec<u8>, Error> {
    let bytes = archive
        .to_bytes_with_limits(limits)
        .map_err(map_core_error)?;
    SnappyStream::compress(&bytes).map_err(map_core_error)
}

pub(super) fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
    path: Path,
) -> Result<(usize, &RawMessage), Error> {
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let selected = messages.next().ok_or(Error::InvalidSource { path })?;
    if messages.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(selected)
}

pub(super) fn validate_selected_metadata(
    object: &ArchiveObject,
    message_index: usize,
    path: Path,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

pub(super) fn strict_reference(
    payload: &[u8],
    limits: WireLimits,
    path: Path,
) -> Result<u64, Error> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut type_seen = false;
    let mut external_seen = false;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource { path });
                }
                identifier = Some(strict_varint(field.payload(), path)?);
            },
            2 => {
                if std::mem::replace(&mut type_seen, true) || field.wire_type() != 0 {
                    return Err(Error::InvalidSource { path });
                }
                strict_varint(field.payload(), path)?;
            },
            3 => {
                if std::mem::replace(&mut external_seen, true) || field.wire_type() != 0 {
                    return Err(Error::InvalidSource { path });
                }
                if strict_varint(field.payload(), path)? != 0 {
                    return Err(Error::InvalidSource { path });
                }
            },
            _ => return Err(Error::InvalidSource { path }),
        }
    }
    match identifier {
        Some(0) | None => Err(Error::InvalidSource { path }),
        Some(value) => Ok(value),
    }
}

fn strict_varint(payload: &[u8], path: Path) -> Result<u64, Error> {
    let (value, consumed) =
        decode_varint_from_bytes(payload).map_err(|_error| Error::InvalidSource { path })?;
    if consumed != payload.len() || encoded_len(value) != consumed {
        return Err(Error::InvalidSource { path });
    }
    Ok(value)
}

pub(super) fn wire_limits(package: &Package) -> Result<WireLimits, Error> {
    let archive = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    WireLimits::default()
        .with_input_bytes(archive.max_message_bytes())
        .and_then(|limits| limits.with_output_bytes(archive.max_message_bytes()))
        .map_err(map_wire_error)
}

fn owned(value: &str) -> Result<String, Error> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(output)
}

pub(super) fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

pub(super) fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::SectionNamesTooLarge { observed, limit } => Error::LimitExceeded {
            path: Path::Package,
            kind: LimitKind::RetainedBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::NotPages => Error::UnsupportedSource {
            path: Path::Package,
        },
        PackageError::PayloadLimit { observed, limit } => Error::LimitExceeded {
            path: Path::Package,
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => Error::LimitExceeded {
            path: Path::Package,
            kind: LimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

pub(super) fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            path: Path::Package,
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource {
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

pub(super) fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            path: Path::Package,
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::PayloadBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

pub(super) fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            path: Path::Package,
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireInputBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::WireFields,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> Error {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => Error::UnsupportedDependency {
            path: Path::Package,
            kind: DependencyKind::LayoutCache,
        },
        page_layout::PageLayoutError::Allocation { amount } => Error::Allocation { amount },
        page_layout::PageLayoutError::LimitExceeded {
            observed, maximum, ..
        } => Error::LimitExceeded {
            path: Path::Package,
            kind: LimitKind::TransactionWork,
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Verification => Error::Verification {
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

pub(super) fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::{Error, LimitKind, Package, TransactionBudget, resolve_target};
    use litchi_core::Position;
    use litchi_iwa_archive::{Limits, package};
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::{tp, tsp, tswp};
    use prost::Message as _;

    const DOCUMENT: u64 = 1;
    const BODY: u64 = 2;
    const SECTION: u64 = 3;
    const TEMPLATE: u64 = 4;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..tsp::Reference::default()
        }
    }

    fn object(
        identifier: u64,
        message_type: u32,
        data: Vec<u8>,
    ) -> Result<ArchiveObject, Box<dyn std::error::Error>> {
        Ok(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data,
            }],
        )?)
    }

    fn topology_package(object_count: usize) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        assert!(object_count >= 4);
        let mut document = object(
            DOCUMENT,
            10_000,
            tp::DocumentArchive {
                body_storage: Some(reference(BODY)),
                ..tp::DocumentArchive::default()
            }
            .encode_to_vec(),
        )?;
        document.archive_info.message_infos[0].object_references = vec![BODY];
        let mut body = object(
            BODY,
            2_001,
            tswp::StorageArchive {
                text: vec!["scale".to_owned()],
                table_section: Some(tswp::ObjectAttributeTable {
                    entries: vec![tswp::object_attribute_table::ObjectAttribute {
                        character_index: 0,
                        object: Some(reference(SECTION)),
                    }],
                }),
                ..tswp::StorageArchive::default()
            }
            .encode_to_vec(),
        )?;
        body.archive_info.message_infos[0].object_references = vec![SECTION];
        let mut section = object(
            SECTION,
            10_011,
            tp::SectionArchive {
                inherit_previous_header_footer: Some(true),
                section_template_first_page_different: Some(false),
                section_template_even_odd_pages_different: Some(false),
                section_start_kind: Some(0),
                section_page_number_kind: Some(0),
                section_page_number_start: Some(1),
                first_section_template_page: Some(reference(TEMPLATE)),
                even_section_template_page: Some(reference(TEMPLATE)),
                odd_section_template_page: Some(reference(TEMPLATE)),
                name: Some("Scale".to_owned()),
                section_template_first_page_hides_header_footer: Some(false),
                ..tp::SectionArchive::default()
            }
            .encode_to_vec(),
        )?;
        section.archive_info.message_infos[0].object_references = vec![TEMPLATE];
        let template = object(
            TEMPLATE,
            10_143,
            tp::SectionTemplateArchive::default().encode_to_vec(),
        )?;
        let mut objects = Vec::new();
        objects.try_reserve_exact(object_count)?;
        objects.extend([document, body, section, template]);
        for offset in 0..object_count - 4 {
            objects.push(object(
                10_000_u64.saturating_add(u64::try_from(offset)?),
                99_999,
                vec![0x08, 0x00],
            )?);
        }
        let compressed = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
        Ok(package::to_bytes(
            [("Index/Document.iwa", compressed.as_slice())],
            Limits::default(),
        )?)
    }

    fn topology_usage(object_count: usize) -> Result<super::Usage, Box<dyn std::error::Error>> {
        let bytes = topology_package(object_count)?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(package.stats().total_objects(), object_count);
        let mut after = package.section_settings(0usize)?;
        after.set_even_odd_pages_different(Some(true));
        Ok(super::super::section_settings::production_test_usage(
            &package, &after, None,
        )?)
    }

    #[test]
    fn production_budget_reports_observed_rooted_work() -> Result<(), Box<dyn std::error::Error>> {
        let source = include_bytes!("../../../../test-data/iwork/pages/basic.pages");
        let package = Package::from_bytes(source)?;
        let mut budget = TransactionBudget::new(&package)?;
        let _target = resolve_target(&package, Position::new(0), &mut budget)?;
        let usage = budget.usage();
        assert!(usage.fields != 0);
        assert!(usage.work != 0);
        assert!(usage.references != 0);
        assert!(usage.transaction_work != 0);
        Ok(())
    }

    #[test]
    fn production_rooted_settings_counters_scale_linearly_and_limit_before_output()
    -> Result<(), Box<dyn std::error::Error>> {
        let small = topology_usage(4_096)?;
        let large = topology_usage(8_192)?;
        assert_eq!(
            (
                small.fields,
                small.work,
                small.references,
                small.transaction_work
            ),
            (77, 564, 4, 292_154),
        );
        assert_eq!(
            (
                large.fields,
                large.work,
                large.references,
                large.transaction_work
            ),
            (77, 564, 4, 587_222),
        );
        for (small_counter, large_counter) in [
            (small.fields, large.fields),
            (small.work, large.work),
            (small.references, large.references),
            (small.transaction_work, large.transaction_work),
        ] {
            assert!(small_counter != 0);
            assert!(large_counter.saturating_mul(10) <= small_counter.saturating_mul(23) + 320);
        }
        assert_eq!(small.output_allocations, 1);
        assert_eq!(large.output_allocations, 1);
        assert_eq!(small.candidate_reopens, 1);
        assert_eq!(large.candidate_reopens, 1);

        let bytes = topology_package(8_192)?;
        let package = Package::from_bytes(&bytes)?;
        let mut after = package.section_settings(0usize)?;
        after.set_even_odd_pages_different(Some(true));
        let (attempt, attempted_usage) = super::super::section_settings::production_test_attempt(
            &package,
            &after,
            large.transaction_work.saturating_sub(1),
        );
        assert!(matches!(
            attempt,
            Err(Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                ..
            })
        ));
        assert_eq!(attempted_usage.output_allocations, 0);
        assert_eq!(attempted_usage.candidate_reopens, 0);
        Ok(())
    }
}
