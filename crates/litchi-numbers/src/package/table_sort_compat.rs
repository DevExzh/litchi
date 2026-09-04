//! Physical-only persisted-sort admission for legacy migration snapshots.
//!
//! Some historical `litchi-iwa` builder snapshots contain valid native table
//! models whose cell-storage projection is not accepted by the strict
//! semantic package reader.  The migration host may still need to configure
//! field 44 for those snapshots.  Keep that exception here, in the focused
//! Numbers owner: only the rooted sheet/table graph and the sort payload are
//! admitted, and the same source-preserving codec is used for the rewrite.

use std::sync::Arc;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    table_model_discovery_codec as model_codec, table_sort_order_codec as codec,
};

use super::{
    Components, Document, DocumentLimits, Index, Package, ReadOptions, SemanticLimits, State,
    table_headers,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::lock::State as LockState;
use crate::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};

use super::table_sort::{
    Error, LimitKind, Path, map_archive_error, map_codec_error, map_core_error, map_header_error,
};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;

type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, Clone, PartialEq, Eq)]
struct Target {
    native: table_headers::Target,
    columns: u32,
    before: Option<Order>,
}

impl Target {
    const fn path(&self) -> Path {
        Path::Table {
            sheet: self.native.sheet_position,
            table: self.native.table_position,
        }
    }
}

impl Package {
    /// Read field 44 without projecting legacy table cell storage.
    #[doc(hidden)]
    pub fn __table_sort_order_from_bytes_for_compatibility<'sheet, 'table>(
        bytes: &[u8],
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Option<Order>> {
        let source = physical_package(bytes)?;
        let target = resolve_target(&source, sheet.into(), table.into())?;
        Ok(target.before)
    }

    /// Rewrite field 44 without projecting legacy table cell storage.
    #[doc(hidden)]
    pub fn __set_table_sort_order_from_bytes_for_compatibility<'sheet, 'table>(
        bytes: &[u8],
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        after: Order,
    ) -> Result<Vec<u8>> {
        let source = physical_package(bytes)?;
        let target = resolve_target(&source, sheet.into(), table.into())?;
        validate_order_columns(&after, target.columns)?;
        if target.before.as_ref() == Some(&after) {
            return copy_bytes(source.state.source.as_ref());
        }
        if target.native.locked == LockState::Locked {
            return Err(Error::TableLocked {
                path: target.path(),
            });
        }
        let output = rewrite_bytes(&source, &target, Some(&after))?;
        verify_candidate(&source, &output, &target, Some(&after))?;
        Ok(output)
    }

    /// Remove field 44 without projecting legacy table cell storage.
    #[doc(hidden)]
    pub fn __clear_table_sort_order_from_bytes_for_compatibility<'sheet, 'table>(
        bytes: &[u8],
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Vec<u8>> {
        let source = physical_package(bytes)?;
        let target = resolve_target(&source, sheet.into(), table.into())?;
        if target.before.is_none() {
            return copy_bytes(source.state.source.as_ref());
        }
        if target.native.locked == LockState::Locked {
            return Err(Error::TableLocked {
                path: target.path(),
            });
        }
        let output = rewrite_bytes(&source, &target, None)?;
        verify_candidate(&source, &output, &target, None)?;
        Ok(output)
    }
}

fn physical_package(bytes: &[u8]) -> Result<Package> {
    let options = ReadOptions::default();
    let components = Components::from_bytes(bytes, options.archive()).map_err(map_package_error)?;
    super::validate_numbers_application(&components, options.archive())
        .map_err(map_package_error)?;
    let index = Index::from_components(&components, options.semantic().max_objects())
        .map_err(map_package_error)?;
    let source = components
        .physical()
        .ok_or(Error::UnsupportedSource)?
        .__source_owner();
    let document = Document::from_sheets_with_limits(Vec::new(), DocumentLimits::default())
        .map_err(|_| invalid_source())?;
    Ok(Package {
        state: Arc::new(State {
            source,
            components,
            index,
            document,
            options,
        }),
    })
}

fn resolve_target(
    source: &Package,
    sheet: SheetSelector<'_>,
    table: TableSelector<'_>,
) -> Result<Target> {
    let sheet_position = resolve_sheet_position(source, sheet)?;
    let table_position = resolve_table_position(source, sheet_position, table)?;
    let native =
        table_headers::resolve::resolve_target_physical(source, sheet_position, table_position)
            .map_err(map_header_error)?;
    if !matches!(
        native.message_type,
        TABLE_MODEL_MESSAGE_TYPE | LEGACY_TABLE_MODEL_MESSAGE_TYPE
    ) {
        return Err(invalid_source());
    }
    let model =
        table_headers::rewrite::selected_payload(source, native).map_err(map_header_error)?;
    let model_snapshot =
        model_codec::decode_table_model(model, model_codec::DecodeOptions::for_source(model))
            .map_err(map_model_codec_error)?;
    table_headers::ownership::validate_selected_ownership(source, native)
        .map_err(map_header_error)?;
    let columns = model_snapshot.number_of_columns();
    let before = decode_sort(model, columns)?;
    Ok(Target {
        native,
        columns,
        before,
    })
}

fn resolve_sheet_position(source: &Package, selector: SheetSelector<'_>) -> Result<usize> {
    let root = Package::root_sheet_order(&source.state.components, SemanticLimits::default())
        .map_err(|_| invalid_source())?;
    match selector {
        SheetSelector::Index(index) => root
            .sheet_references()
            .get(index)
            .map(|_| index)
            .ok_or(Error::SheetNotFound),
        SheetSelector::Name(name) => {
            let mut found = None;
            for index in 0..root.sheet_references().len() {
                let target = resolve_sheet_target(source, index)?;
                if sheet_name(source, target)? == name && found.replace(index).is_some() {
                    return Err(invalid_source());
                }
            }
            found.ok_or(Error::SheetNotFound)
        },
    }
}

fn resolve_table_position(
    source: &Package,
    sheet_position: usize,
    selector: TableSelector<'_>,
) -> Result<usize> {
    match selector {
        TableSelector::Index(index) => {
            table_headers::resolve::resolve_target_physical(source, sheet_position, index)
                .map_err(map_header_error)?;
            Ok(index)
        },
        TableSelector::Name(name) => {
            let mut found = None;
            let table_count = physical_table_count(source, sheet_position)?;
            for table_index in 0..table_count {
                let target = table_headers::resolve::resolve_target_physical(
                    source,
                    sheet_position,
                    table_index,
                )
                .map_err(map_header_error)?;
                if table_name(source, target)? == name && found.replace(table_index).is_some() {
                    return Err(invalid_source());
                }
            }
            found.ok_or(Error::TableNotFound)
        },
    }
}

fn table_name(source: &Package, target: table_headers::Target) -> Result<&str> {
    let message = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or_else(invalid_source)?;
    super::names::preflight_table_name(&message.data).map_err(|_| invalid_source())
}

fn physical_table_count(source: &Package, sheet_position: usize) -> Result<usize> {
    let sheet = resolve_sheet_target(source, sheet_position)?;
    let message = source
        .state
        .components
        .catalog()
        .get_index(sheet.component_index)
        .and_then(|component| component.archive().objects.get(sheet.object_index))
        .and_then(|object| object.messages.get(sheet.message_index))
        .ok_or_else(invalid_source)?;
    let payloads = table_headers::resolve::sheet_drawable_payloads(message.type_, &message.data)
        .map_err(map_header_error)?;
    let mut count = 0usize;
    for payload in payloads {
        let identifier = table_headers::resolve::local_reference_identifier(payload)
            .map_err(map_header_error)?;
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(|_| invalid_source())?
            .ok_or_else(invalid_source)?;
        if table_headers::resolve::unique_table_info(resolved)
            .map_err(map_header_error)?
            .is_some()
        {
            count = count.checked_add(1).ok_or_else(invalid_source)?;
        }
    }
    Ok(count)
}

fn resolve_sheet_target(source: &Package, sheet_position: usize) -> Result<SheetTarget> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or_else(invalid_source)?;
    let (_index, document_message) = table_headers::resolve::unique_message_index(
        &document_object.messages,
        super::DOCUMENT_MESSAGE_TYPE,
    )
    .map_err(map_header_error)?
    .ok_or_else(invalid_source)?;
    let sheet_payloads =
        table_headers::resolve::repeated_length_payloads(&document_message.data, 1)
            .map_err(map_header_error)?;
    let sheet_identifier = table_headers::resolve::local_reference_identifier(
        sheet_payloads
            .get(sheet_position)
            .ok_or(Error::SheetNotFound)?,
    )
    .map_err(map_header_error)?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(|_| invalid_source())?
        .ok_or_else(invalid_source)?;
    let message_index = table_headers::resolve::unique_sheet_message_index(resolved.messages)
        .map_err(map_header_error)?;
    let message = resolved
        .messages
        .get(message_index)
        .ok_or_else(invalid_source)?;
    Ok(SheetTarget {
        identifier: sheet_identifier,
        component_index: resolved.component_index,
        object_index: resolved.object_index,
        message_index,
        message_type: message.type_,
    })
}

#[derive(Debug, Clone, Copy)]
struct SheetTarget {
    identifier: u64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
}

fn sheet_name(source: &Package, target: SheetTarget) -> Result<&str> {
    let message = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or_else(invalid_source)?;
    let payload = match target.message_type {
        SHEET_MESSAGE_TYPE => message.data.as_slice(),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            table_headers::resolve::singular_length_payload(&message.data, 1)
                .map_err(map_header_error)?
        },
        _ => return Err(invalid_source()),
    };
    super::names::preflight_sheet_name(target.message_type, payload).map_err(|_| invalid_source())
}

fn decode_sort(source: &[u8], columns: u32) -> Result<Option<Order>> {
    let options = codec::DecodeOptions::for_source(source)
        .with_max_columns(usize::try_from(columns).unwrap_or(usize::MAX));
    let snapshot =
        codec::decode_table_model_sort_order(source, options).map_err(map_codec_error)?;
    snapshot.map(order_from_snapshot).transpose()
}

fn order_from_snapshot(snapshot: codec::SortOrderSnapshot) -> Result<Order> {
    let scope = match snapshot.scope() {
        codec::SortScope::EntireTable => Scope::EntireTable,
        codec::SortScope::SelectedRows => Scope::SelectedRows,
    };
    let mut rules = Vec::new();
    rules
        .try_reserve_exact(snapshot.rules().len())
        .map_err(|_| invalid_source())?;
    for rule in snapshot.rules() {
        let column = ColumnIndex::from_native(rule.column()).map_err(|_| invalid_source())?;
        let direction = match rule.direction() {
            codec::SortDirection::Ascending => Direction::Ascending,
            codec::SortDirection::Descending => Direction::Descending,
        };
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules).map_err(|_| invalid_source())
}

fn snapshot_from_order(order: &Order) -> Result<codec::SortOrderSnapshot> {
    let scope = match order.scope() {
        Scope::EntireTable => codec::SortScope::EntireTable,
        Scope::SelectedRows => codec::SortScope::SelectedRows,
    };
    let rules = order.rules().iter().map(|rule| {
        let direction = match rule.direction() {
            Direction::Ascending => codec::SortDirection::Ascending,
            Direction::Descending => codec::SortDirection::Descending,
        };
        codec::SortRule::new(rule.column().native_value(), direction)
    });
    codec::SortOrderSnapshot::new(scope, rules).map_err(map_codec_error)
}

fn rewrite_bytes(source: &Package, target: &Target, after: Option<&Order>) -> Result<Vec<u8>> {
    let catalog = table_headers::rewrite::physical_source(source).map_err(map_header_error)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.native.component_index)
        .ok_or_else(invalid_source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or_else(invalid_source)?;
    let physical_limits = catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    drop(stream);
    let object = archive
        .objects
        .get_mut(target.native.object_index)
        .ok_or_else(invalid_source)?;
    if object.archive_info.identifier != Some(target.native.model_identifier) {
        return Err(invalid_source());
    }
    table_headers::resolve::validate_message_metadata(object, target.native.message_index)
        .map_err(map_header_error)?;
    let desired = after.map(snapshot_from_order).transpose()?;
    let replacement = {
        let original = &object
            .messages
            .get(target.native.message_index)
            .ok_or_else(invalid_source)?
            .data;
        if decode_sort(original, target.columns)? != target.before {
            return Err(invalid_source());
        }
        let options = codec::DecodeOptions::for_source(original)
            .with_max_columns(usize::try_from(target.columns).unwrap_or(usize::MAX));
        codec::rewrite_table_model_sort_order(original, desired, options)
            .map_err(map_codec_error)?
            .into_bytes()
    };
    object
        .replace_message_preserving_header_with_limits(
            target.native.message_index,
            RawMessage {
                type_: target.native.message_type,
                data: replacement,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(rewritten.len()).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::EntryBytes,
            observed: u64::try_from(compressed_bound).unwrap_or(u64::MAX),
            maximum: u64::try_from(snappy_limits.max_compressed_stream()).unwrap_or(u64::MAX),
            path: Path::Package,
        });
    }
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    let edits = [EntryEdit::new(component.name(), compressed.as_slice())];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)
}

fn verify_candidate(
    source: &Package,
    output: &[u8],
    target: &Target,
    expected: Option<&Order>,
) -> Result<()> {
    let candidate = physical_package(output)?;
    let selected = resolve_target(
        &candidate,
        SheetSelector::Index(target.native.sheet_position),
        TableSelector::Index(target.native.table_position),
    )?;
    if selected.before.as_ref() != expected
        || selected.native.model_identifier != target.native.model_identifier
    {
        return Err(Error::Verification);
    }
    verify_unchanged_entries(source, &candidate, target, expected)
}

fn verify_unchanged_entries(
    source: &Package,
    candidate: &Package,
    target: &Target,
    expected: Option<&Order>,
) -> Result<()> {
    let source_catalog =
        table_headers::rewrite::physical_source(source).map_err(map_header_error)?;
    let candidate_catalog =
        table_headers::rewrite::physical_source(candidate).map_err(map_header_error)?;
    if source_catalog.package().iter().count() != candidate_catalog.package().iter().count() {
        return Err(Error::Verification);
    }
    let selected_name = source
        .state
        .components
        .catalog()
        .get_index(target.native.component_index)
        .ok_or(Error::Verification)?
        .name();
    for (before, after) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        if before.name() != after.name() {
            return Err(Error::Verification);
        }
        if before.name() == selected_name {
            if !table_headers::rewrite::selected_package_member_preserved(before, after) {
                return Err(Error::Verification);
            }
            continue;
        }
        if before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || !central_record_preserved_except_offset(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
            )
        {
            return Err(Error::Verification);
        }
    }
    verify_selected_component_objects(source, candidate, target, expected)
}

fn verify_selected_component_objects(
    source: &Package,
    candidate: &Package,
    target: &Target,
    expected: Option<&Order>,
) -> Result<()> {
    let source_component = source
        .state
        .components
        .catalog()
        .get_index(target.native.component_index)
        .ok_or(Error::Verification)?;
    let candidate_component = candidate
        .state
        .components
        .catalog()
        .get(source_component.name())
        .ok_or(Error::Verification)?;
    if source_component.archive().objects.len() != candidate_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (object_index, (before, after)) in source_component
        .archive()
        .objects
        .iter()
        .zip(&candidate_component.archive().objects)
        .enumerate()
    {
        if object_index != target.native.object_index {
            if !before.same_content_ignoring_offsets(after) {
                return Err(Error::Verification);
            }
            continue;
        }
        if before.archive_info.identifier != Some(target.native.model_identifier)
            || after.archive_info.identifier != Some(target.native.model_identifier)
            || before.archive_info.should_merge != after.archive_info.should_merge
            || before.messages.len() != after.messages.len()
            || before.archive_info.message_infos.len() != after.archive_info.message_infos.len()
        {
            return Err(Error::Verification);
        }
        for (message_index, (before_message, after_message)) in
            before.messages.iter().zip(&after.messages).enumerate()
        {
            let before_info = before
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification)?;
            let after_info = after
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification)?;
            if message_index == target.native.message_index {
                if before_message.type_ != after_message.type_
                    || !message_info_preserved_except_length(before_info, after_info)
                {
                    return Err(Error::Verification);
                }
                verify_selected_payload(
                    &before_message.data,
                    &after_message.data,
                    target.columns,
                    expected,
                )?;
            } else if before_message != after_message || before_info != after_info {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn verify_selected_payload(
    source: &[u8],
    candidate: &[u8],
    columns: u32,
    expected: Option<&Order>,
) -> Result<()> {
    let desired = expected.map(snapshot_from_order).transpose()?;
    let options = codec::DecodeOptions::for_source(source)
        .with_max_columns(usize::try_from(columns).unwrap_or(usize::MAX));
    let permitted =
        codec::rewrite_table_model_sort_order(source, desired, options).map_err(map_codec_error)?;
    if candidate != permitted.output() {
        return Err(Error::Verification);
    }
    Ok(())
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn central_record_preserved_except_offset(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn validate_order_columns(order: &Order, columns: u32) -> Result<()> {
    if order
        .rules()
        .iter()
        .any(|rule| u64::from(rule.column().native_value()) >= u64::from(columns))
    {
        return Err(invalid_source());
    }
    Ok(())
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
            amount: source.len(),
            path: Path::Package,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

const fn invalid_source() -> Error {
    Error::InvalidSource {
        path: Path::Package,
    }
}

fn map_package_error(error: super::Error) -> Error {
    if let Some(resource) = error.resource_error() {
        return match resource {
            super::ResourceError::Allocation { amount } => Error::Allocation {
                amount,
                path: Path::Package,
            },
            super::ResourceError::LimitExceeded {
                kind,
                observed,
                maximum,
            } => Error::LimitExceeded {
                kind: match kind {
                    super::PayloadLimitKind::InputBytes => LimitKind::WireBytes,
                    super::PayloadLimitKind::Fields => LimitKind::WireFields,
                    super::PayloadLimitKind::OutputBytes => LimitKind::WireOutputBytes,
                    super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                    super::PayloadLimitKind::RewriteWork => LimitKind::WireWork,
                    super::PayloadLimitKind::TableRows
                    | super::PayloadLimitKind::TableColumns
                    | super::PayloadLimitKind::TableCells
                    | super::PayloadLimitKind::MaterializedCells => LimitKind::TransactionWork,
                },
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
                path: Path::Package,
            },
        };
    }
    match error {
        super::Error::Archive(error) => map_archive_error(error),
        super::Error::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        super::Error::SemanticLimit {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

fn map_model_codec_error(error: model_codec::DecodeError) -> Error {
    let Some(limit) = error.resource_limit() else {
        return invalid_source();
    };
    let (kind, observed, maximum) = match limit {
        model_codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        model_codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        model_codec::DecodeLimit::Work { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        model_codec::DecodeLimit::Text { observed, maximum } => {
            (LimitKind::PayloadBytes, observed, maximum)
        },
        model_codec::DecodeLimit::Output { observed, maximum } => {
            (LimitKind::WireOutputBytes, observed, maximum)
        },
        model_codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::WireAllocations, observed, maximum)
        },
        model_codec::DecodeLimit::Retained { observed, maximum } => {
            (LimitKind::WireRetainedBytes, observed, maximum)
        },
        model_codec::DecodeLimit::Scratch { observed, maximum } => {
            (LimitKind::WireScratchBytes, observed, maximum)
        },
        model_codec::DecodeLimit::Nesting { observed, maximum } => {
            return Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path: Path::Package,
            };
        },
        _ => return invalid_source(),
    };
    Error::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path: Path::Package,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn one_rule_order() -> Order {
        Order::new([Rule::new(
            ColumnIndex::new(0).expect("column fits"),
            Direction::Ascending,
        )])
        .expect("one-rule order")
    }

    #[test]
    fn selected_payload_verification_allows_only_the_exact_field_44_rewrite() {
        // Field 1 is outside the persisted-sort envelope and must remain
        // byte-identical even when the requested sort semantics read back.
        // Keep enough opaque model bytes for the codec's finite
        // source-derived output budget to admit one newly-added rule.
        let source = [
            0x08, 0x07, 0x10, 0x01, 0x18, 0x01, 0x20, 0x01, 0x28, 0x01, 0x30, 0x01,
        ];
        let order = one_rule_order();
        let desired = snapshot_from_order(&order).expect("sort snapshot");
        let options = codec::DecodeOptions::for_source(&source).with_max_columns(1);
        let candidate = codec::rewrite_table_model_sort_order(&source, Some(desired), options)
            .expect("permitted rewrite")
            .into_bytes();

        verify_selected_payload(&source, &candidate, 1, Some(&order))
            .expect("exact field-44 rewrite");

        let mut collateral = candidate;
        collateral[1] = 0x08;
        assert_eq!(
            verify_selected_payload(&source, &collateral, 1, Some(&order)),
            Err(Error::Verification)
        );
    }
}
