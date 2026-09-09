//! Selector-first comment reads for legacy Numbers source graphs.
//!
//! The normal [`super::Package`] constructor projects every rooted table into
//! the semantic document.  A small set of historical producers emits valid
//! comment and table-storage graphs while omitting the optional style
//! envelopes required by that broad projection.  This module keeps the
//! migration exception narrow: it admits the physical package, resolves one
//! selector to one rooted table, and delegates cell/comment decoding to the
//! same lazy Buffa-backed reader used by the normal comment API.
//!
//! The package remains private to this adapter.  Native object identifiers,
//! archive routes, and protobuf snapshots never appear in the compatibility
//! API.  The source bytes remain authoritative and all archive, index, header,
//! cell-storage, and comment limits are applied before retained work grows.

use std::sync::Arc;

use litchi_iwa_archive::LimitKind as ArchiveLimitKind;
use litchi_iwa_common::{
    WireLimits,
    wire::{WireDescent, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{LimitKind as CoreLimitKind, RawMessage};
use litchi_iwa_protos::table_info_codec;

use super::super::{
    Components, Document, DocumentLimits, Error as PackageError, Index, Package, ReadOptions,
    SemanticLimitKind, State, table_headers, validate_numbers_application,
};
use super::{Comment, CommentReply, Error, Path, Target};
use crate::{SheetSelector, TableSelector, table::CellPosition};

const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;

type Result<T> = std::result::Result<T, Error>;

impl Package {
    /// Read one semantic comment through the bounded legacy-source adapter.
    ///
    /// This hidden entry point lets the migration host read a source whose
    /// focused table projection is incomplete. Callers still provide only
    /// semantic selectors and a checked cell position.
    #[doc(hidden)]
    pub fn __table_cell_comment_from_bytes_for_compatibility<'sheet, 'table>(
        bytes: &[u8],
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<Comment>> {
        Self::__table_cell_comment_from_bytes_for_compatibility_with_options(
            bytes,
            ReadOptions::default(),
            sheet,
            table,
            position,
        )
    }

    /// Read one semantic comment with caller-selected physical and semantic
    /// resource bounds.
    #[doc(hidden)]
    pub fn __table_cell_comment_from_bytes_for_compatibility_with_options<'sheet, 'table>(
        bytes: &[u8],
        options: ReadOptions,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<Comment>> {
        let source = compatibility_package(bytes, options)?;
        let located = resolve_compatibility_comment(&source, sheet.into(), table.into(), position)?;
        Ok(located.comment)
    }

    /// Read direct semantic replies through the bounded legacy-source adapter.
    #[doc(hidden)]
    pub fn __table_cell_comment_replies_from_bytes_for_compatibility<'sheet, 'table>(
        bytes: &[u8],
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Box<[CommentReply]>> {
        Self::__table_cell_comment_replies_from_bytes_for_compatibility_with_options(
            bytes,
            ReadOptions::default(),
            sheet,
            table,
            position,
        )
    }

    /// Read direct semantic replies with caller-selected physical and
    /// semantic resource bounds.
    #[doc(hidden)]
    pub fn __table_cell_comment_replies_from_bytes_for_compatibility_with_options<
        'sheet,
        'table,
    >(
        bytes: &[u8],
        options: ReadOptions,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Box<[CommentReply]>> {
        let source = compatibility_package(bytes, options)?;
        let located = resolve_compatibility_comment(&source, sheet.into(), table.into(), position)?;
        source.read_comment_replies_compatibility(located)
    }
}

/// Build only the physical source/index state required by a focused read.
///
/// In particular, this does not call `Package::from_bytes_with_options` and
/// does not materialize any sheet, table, row, or cell projection.  The empty
/// document is an implementation detail that lets the existing native
/// comment decoder share the package-owned source and bounds.
fn compatibility_package(bytes: &[u8], options: ReadOptions) -> Result<Package> {
    let components = Components::from_bytes(bytes, options.archive()).map_err(map_package_error)?;
    validate_numbers_application(&components, options.archive()).map_err(map_package_error)?;
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

fn resolve_compatibility_comment(
    source: &Package,
    sheet: SheetSelector<'_>,
    table: TableSelector<'_>,
    position: CellPosition,
) -> Result<super::Located> {
    let sheet_position = resolve_sheet_position(source, sheet)?;
    let table_count = bounded_table_count(source, sheet_position)?;
    let table_position = resolve_table_position(source, sheet_position, table, table_count)?;
    let native = table_headers::resolve::resolve_target(source, sheet_position, table_position)
        .map_err(map_header_error)?;
    let row = usize::try_from(position.row()).map_err(|_| invalid_source())?;
    let column = usize::try_from(position.column()).map_err(|_| invalid_source())?;
    let path = Path::Cell {
        sheet: sheet_position,
        table: table_position,
        row,
        column,
    };
    if position.row() >= native.rows || position.column() >= native.columns {
        return Err(Error::OutOfBounds { path });
    }
    super::resolve_comment_native_compatibility(
        source,
        Target {
            path,
            native,
            row,
            column,
            tile_size: super::DEFAULT_TILE_SIZE,
        },
    )
}

fn resolve_sheet_position(source: &Package, selector: SheetSelector<'_>) -> Result<usize> {
    let root = Package::root_sheet_order(&source.state.components, source.state.options.semantic())
        .map_err(map_package_error)?;
    match selector {
        SheetSelector::Index(index) => root
            .sheet_references()
            .get(index)
            .map(|_| index)
            .ok_or(Error::SheetNotFound),
        SheetSelector::Name(name) => {
            let mut found = None;
            for (index, reference) in root.sheet_references().iter().enumerate() {
                let target = resolve_sheet_target(source, reference.identifier())?;
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
    table_count: usize,
) -> Result<usize> {
    match selector {
        TableSelector::Index(index) => {
            if index >= table_count {
                return Err(Error::TableNotFound);
            }
            Ok(index)
        },
        TableSelector::Name(name) => {
            let mut found = None;
            let sheet_identifier = root_sheet_identifier(source, sheet_position)?;
            let sheet_target = resolve_sheet_target(source, sheet_identifier)?;
            let message = sheet_message(source, sheet_target)?;
            let mut table_position = 0usize;
            visit_drawable_payloads(message.type_, &message.data, |payload| {
                let Some(candidate) = table_name_for_payload(source, payload)? else {
                    return Ok(());
                };
                if table_position >= table_count {
                    return Err(invalid_source());
                }
                if candidate == name && found.replace(table_position).is_some() {
                    return Err(invalid_source());
                }
                table_position = table_position.checked_add(1).ok_or_else(|| {
                    table_limit_exceeded(source.state.options.semantic().max_tables())
                })?;
                Ok(())
            })?;
            if table_position != table_count {
                return Err(invalid_source());
            }
            found.ok_or(Error::TableNotFound)
        },
    }
}

/// Count only the selected sheet's rooted table drawables before any header
/// resolver is allowed to materialize its payload vector. The source-built
/// compatibility path has the same table ceiling for positional and name
/// selectors, and every inspected drawable consumes one reference unit.
fn bounded_table_count(source: &Package, sheet_position: usize) -> Result<usize> {
    let sheet_identifier = root_sheet_identifier(source, sheet_position)?;
    let target = resolve_sheet_target(source, sheet_identifier)?;
    let message = sheet_message(source, target)?;
    let maximum_tables = source.state.options.semantic().max_tables();
    let maximum_references = source.state.options.semantic().max_references();
    let mut references = 0usize;
    let mut tables = 0usize;
    visit_drawable_payloads(message.type_, &message.data, |payload| {
        references = references.checked_add(1).ok_or(Error::LimitExceeded {
            kind: super::LimitKind::References,
            observed: usize::MAX,
            maximum: maximum_references,
            path: Path::Package,
        })?;
        if references > maximum_references {
            return Err(Error::LimitExceeded {
                kind: super::LimitKind::References,
                observed: references,
                maximum: maximum_references,
                path: Path::Package,
            });
        }
        let identifier = super::super::names::preflight_local_reference(payload)
            .map_err(|error| super::map_wire_error(error, Path::Package))?;
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(map_package_error)?
            .ok_or_else(invalid_source)?;
        if table_headers::resolve::unique_table_info(resolved)
            .map_err(map_header_error)?
            .is_some()
        {
            tables = tables
                .checked_add(1)
                .ok_or_else(|| table_limit_exceeded(maximum_tables))?;
            if tables > maximum_tables {
                return Err(table_limit_exceeded(maximum_tables));
            }
        }
        Ok(())
    })?;
    Ok(tables)
}

fn visit_drawable_payloads(
    message_type: u32,
    source: &[u8],
    mut visitor: impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    let limits = drawable_wire_limits(source, message_type)?;
    let mut form_super_seen = false;
    let mut callback_error = None;
    let preflight = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        let is_form_super = message_type == FORM_BASED_SHEET_MESSAGE_TYPE
            && visit.path().is_empty()
            && field.number() == 1;
        if is_form_super {
            if form_super_seen || field.wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid form-based sheet super payload".into(),
                ));
            }
            field.validate_canonical_framing()?;
            form_super_seen = true;
            return Ok(WireDescent::Descend);
        }

        let is_drawable = match message_type {
            SHEET_MESSAGE_TYPE => visit.path().is_empty() && field.number() == 2,
            FORM_BASED_SHEET_MESSAGE_TYPE => visit.path() == [1] && field.number() == 2,
            _ => {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid sheet type".into(),
                ));
            },
        };
        if is_drawable {
            if field.wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid drawable reference".into(),
                ));
            }
            field.validate_canonical_framing()?;
            if let Err(error) = visitor(field.payload()) {
                callback_error = Some(error);
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "drawable visitor stopped the bounded scan".into(),
                ));
            }
        }
        Ok(WireDescent::Skip)
    });
    if let Some(error) = callback_error {
        return Err(error);
    }
    preflight.map_err(|error| super::map_wire_error(error, Path::Package))?;
    if message_type == FORM_BASED_SHEET_MESSAGE_TYPE && !form_super_seen {
        return Err(invalid_source());
    }
    Ok(())
}

fn drawable_wire_limits(source: &[u8], message_type: u32) -> Result<WireLimits> {
    let nesting = match message_type {
        SHEET_MESSAGE_TYPE => 1,
        FORM_BASED_SHEET_MESSAGE_TYPE => 1,
        _ => return Err(invalid_source()),
    };
    WireLimits::default()
        .with_input_bytes(
            source
                .len()
                .saturating_mul(2)
                .clamp(1, WireLimits::MAX_INPUT_BYTES),
        )
        .and_then(|limits| {
            limits.with_fields(
                source
                    .len()
                    .saturating_mul(2)
                    .clamp(1, WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| limits.with_nesting(nesting))
        .map_err(|error| super::map_wire_error(error, Path::Package))
}

#[derive(Debug, Clone, Copy)]
struct SheetTarget {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
}

fn root_sheet_identifier(source: &Package, sheet_position: usize) -> Result<u64> {
    let root = Package::root_sheet_order(&source.state.components, source.state.options.semantic())
        .map_err(map_package_error)?;
    root.sheet_references()
        .get(sheet_position)
        .map(|reference| reference.identifier())
        .ok_or(Error::SheetNotFound)
}

fn resolve_sheet_target(source: &Package, sheet_identifier: u64) -> Result<SheetTarget> {
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(map_package_error)?
        .ok_or_else(invalid_source)?;
    let message_index = table_headers::resolve::unique_sheet_message_index(resolved.messages)
        .map_err(map_header_error)?;
    let message = resolved
        .messages
        .get(message_index)
        .ok_or_else(invalid_source)?;
    Ok(SheetTarget {
        component_index: resolved.component_index,
        object_index: resolved.object_index,
        message_index,
        message_type: message.type_,
    })
}

fn sheet_message(source: &Package, target: SheetTarget) -> Result<&RawMessage> {
    source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or_else(invalid_source)
}

fn table_name_for_payload<'source>(
    source: &'source Package,
    payload: &[u8],
) -> Result<Option<&'source str>> {
    let identifier = super::super::names::preflight_local_reference(payload)
        .map_err(|error| super::map_wire_error(error, Path::Package))?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(map_package_error)?
        .ok_or_else(invalid_source)?;
    let Some((_, info_message)) =
        table_headers::resolve::unique_table_info(resolved).map_err(map_header_error)?
    else {
        return Ok(None);
    };
    let info = table_info_codec::decode_table_info(
        &info_message.data,
        super::super::table_info_decode_options(&info_message.data),
    )
    .map_err(map_table_info_codec_error)?;
    let model_identifier = info.table_model().identifier().get();
    let model = source
        .state
        .index
        .resolve_ref_id(&source.state.components, model_identifier)
        .map_err(map_package_error)?
        .ok_or_else(invalid_source)?;
    let (_, model_message) =
        table_headers::resolve::unique_table_model(model.messages).map_err(map_header_error)?;
    super::super::names::preflight_table_name(&model_message.data)
        .map(Some)
        .map_err(|_| invalid_source())
}

fn map_table_info_codec_error(error: table_info_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: super::LimitKind::WireFields,
            observed,
            maximum,
            path: Path::Package,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            table_info_codec::WireResourceLimit::Bytes { observed, maximum } => {
                Error::LimitExceeded {
                    kind: super::LimitKind::WireBytes,
                    observed: observed.unwrap_or(usize::MAX),
                    maximum: maximum.unwrap_or(usize::MAX),
                    path: Path::Package,
                }
            },
            table_info_codec::WireResourceLimit::Nesting { observed, maximum } => {
                Error::LimitExceeded {
                    kind: super::LimitKind::WireWork,
                    observed: observed
                        .map(|value| usize::try_from(value).unwrap_or(usize::MAX))
                        .unwrap_or(usize::MAX),
                    maximum: maximum
                        .map(|value| usize::try_from(value).unwrap_or(usize::MAX))
                        .unwrap_or(usize::MAX),
                    path: Path::Package,
                }
            },
            _ => invalid_source(),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: super::LimitKind::WireWork,
            observed,
            maximum,
            path: Path::Package,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation {
            amount,
            path: Path::Package,
        };
    }
    invalid_source()
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
    super::super::names::preflight_sheet_name(target.message_type, &message.data)
        .map_err(|_| invalid_source())
}

fn map_header_error(error: table_headers::Error) -> Error {
    match error {
        table_headers::Error::SheetNotFound => Error::SheetNotFound,
        table_headers::Error::TableNotFound => Error::TableNotFound,
        table_headers::Error::LimitExceeded {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: map_header_limit_kind(kind),
            observed: usize_from_u64(observed),
            maximum: usize_from_u64(maximum),
            path: Path::Package,
        },
        table_headers::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

const fn map_header_limit_kind(kind: table_headers::LimitKind) -> super::LimitKind {
    match kind {
        table_headers::LimitKind::InputBytes => super::LimitKind::InputBytes,
        table_headers::LimitKind::OutputBytes | table_headers::LimitKind::WireOutputBytes => {
            super::LimitKind::OutputBytes
        },
        table_headers::LimitKind::WireFields => super::LimitKind::WireFields,
        table_headers::LimitKind::WireBytes
        | table_headers::LimitKind::PayloadBytes
        | table_headers::LimitKind::TotalPayloadBytes
        | table_headers::LimitKind::EntryBytes
        | table_headers::LimitKind::TotalEntryBytes
        | table_headers::LimitKind::PackageBytes
        | table_headers::LimitKind::Entries => super::LimitKind::WireBytes,
        table_headers::LimitKind::PayloadReferences => super::LimitKind::References,
        table_headers::LimitKind::WireNesting | table_headers::LimitKind::WireWork => {
            super::LimitKind::WireWork
        },
        table_headers::LimitKind::PayloadObjects
        | table_headers::LimitKind::PayloadMessages
        | table_headers::LimitKind::PayloadItems
        | table_headers::LimitKind::TransactionWork => super::LimitKind::References,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Common(error) => super::map_wire_error(error, Path::Package),
        PackageError::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: super::LimitKind::InputBytes,
            observed: usize_from_u64(observed),
            maximum: usize_from_u64(maximum),
            path: Path::Package,
        },
        PackageError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: map_semantic_limit_kind(kind),
            observed,
            maximum,
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => super::LimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => super::LimitKind::OutputBytes,
                _ => super::LimitKind::WireBytes,
            },
            observed: usize_from_u64(observed),
            maximum: usize_from_u64(maximum),
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => invalid_source(),
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                CoreLimitKind::ArchiveBytes
                | CoreLimitKind::ObjectBytes
                | CoreLimitKind::MessageBytes
                | CoreLimitKind::HeaderBytes
                | CoreLimitKind::HeaderMemoryBytes
                | CoreLimitKind::SnappyChunkBytes => super::LimitKind::WireBytes,
                CoreLimitKind::Objects
                | CoreLimitKind::Messages
                | CoreLimitKind::MessagesPerObject
                | CoreLimitKind::HeaderFields
                | CoreLimitKind::MetadataItems
                | CoreLimitKind::SnappyFrames => super::LimitKind::References,
                CoreLimitKind::HeaderNesting => super::LimitKind::WireWork,
                CoreLimitKind::SnappyStreamBytes
                | CoreLimitKind::SnappyCompressedChunkBytes
                | CoreLimitKind::SnappyCompressedStreamBytes => super::LimitKind::WireBytes,
            },
            observed,
            maximum,
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

const fn map_semantic_limit_kind(kind: SemanticLimitKind) -> super::LimitKind {
    match kind {
        SemanticLimitKind::OutputTextBytes
        | SemanticLimitKind::TextBytes
        | SemanticLimitKind::FormulaWireBytes => super::LimitKind::TextBytes,
        SemanticLimitKind::References
        | SemanticLimitKind::Objects
        | SemanticLimitKind::Sheets
        | SemanticLimitKind::Tables
        | SemanticLimitKind::MaterializedCells
        | SemanticLimitKind::FormulaWork
        | SemanticLimitKind::FormulaDepth
        | SemanticLimitKind::FormulaRenderWork
        | SemanticLimitKind::FormulaRenderDepth => super::LimitKind::References,
    }
}

fn usize_from_u64(value: u64) -> usize {
    usize::try_from(value).unwrap_or(usize::MAX)
}

const fn invalid_source() -> Error {
    Error::InvalidSource {
        path: Path::Package,
    }
}

fn table_limit_exceeded(maximum: usize) -> Error {
    Error::LimitExceeded {
        kind: super::LimitKind::References,
        observed: maximum.saturating_add(1),
        maximum,
        path: Path::Package,
    }
}
