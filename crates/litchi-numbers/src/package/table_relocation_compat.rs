//! Physical-topology admission for the retiring `litchi-iwa` host.
//!
//! The normal [`Package`](super::Package) constructor deliberately projects
//! every rooted table into the focused semantic model. Historical host-built
//! workbooks can contain valid table storage that this projection does not yet
//! support. This module keeps the narrow migration path in the Numbers owner:
//! it resolves only rooted sheets, table identity, lock state, and parent
//! topology, then invokes the same relocation engine as the public transaction.

use std::sync::Arc;

use super::table_relocation::{
    Error, Operation, Path, SheetTarget, map_header_error, physical_source, resolve_sheet_target,
    rewrite_bytes_for_compatibility, root_preview_deletions, verify_locality,
    verify_physical_source_state, verify_physical_target_state,
};
use super::{
    Components, Document, DocumentLimits, Package, ReadOptions, SemanticLimits, State,
    table_headers,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::lock::State as LockState;

const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;

type Result<T> = std::result::Result<T, Error>;

impl Package {
    /// Relocate a table for the retiring compatibility host without projecting
    /// its cell storage.
    ///
    /// This is not a second mutation implementation. It performs physical-only
    /// selector admission and delegates rewriting plus source, target, and
    /// locality verification to the focused transaction engine.
    #[doc(hidden)]
    pub fn __move_table_from_bytes_for_compatibility<'source_sheet, 'table, 'destination_sheet>(
        bytes: &[u8],
        source_sheet: impl Into<SheetSelector<'source_sheet>>,
        table: impl Into<TableSelector<'table>>,
        destination_sheet: impl Into<SheetSelector<'destination_sheet>>,
    ) -> Result<Vec<u8>> {
        let source = physical_package(bytes)?;
        let source_sheet = resolve_sheet_position(&source, source_sheet.into())?;
        let source_table = resolve_table_position(&source, source_sheet, table.into())?;
        let destination_sheet = resolve_sheet_position(&source, destination_sheet.into())?;
        let path = Path::Table {
            source_sheet,
            table: source_table,
            destination_sheet,
        };
        let table =
            table_headers::resolve::resolve_target_physical(&source, source_sheet, source_table)
                .map_err(|_error| invalid_source(path))?;
        let source_target = resolve_sheet_target(&source, source_sheet, path)?;
        let destination_target = resolve_sheet_target(&source, destination_sheet, path)?;
        if table.sheet_identifier != source_target.identifier {
            return Err(invalid_source(path));
        }
        let destination_table = if source_target.identifier == destination_target.identifier {
            source_table
        } else {
            physical_table_count(&source, destination_sheet, path)?
        };
        let operation = Operation {
            source_sheet,
            source_table,
            destination_sheet,
            destination_table,
            source_sheet_identifier: source_target.identifier,
            destination_sheet_identifier: destination_target.identifier,
            drawable_identifier: table.drawable_identifier,
            model_identifier: table.model_identifier,
        };

        if operation.is_same_sheet() {
            return copy_bytes(source.state.source.as_ref(), path);
        }
        if table.locked == LockState::Locked {
            return Err(Error::TableLocked { path });
        }

        let catalog = physical_source(&source)?;
        if !catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let deleted_previews = root_preview_deletions(catalog)?;
        let source_parent = verify_physical_source_state(&source, operation)?;
        let output = rewrite_bytes_for_compatibility(
            &source,
            operation,
            table,
            source_target,
            destination_target,
            &deleted_previews,
        )?;
        let candidate = physical_package(&output)?;
        verify_physical_target_state(&candidate, operation, source_parent)?;
        verify_locality(&source, &candidate, operation)?;
        Ok(output)
    }
}

fn physical_package(bytes: &[u8]) -> Result<Package> {
    let options = ReadOptions::default();
    let components = Components::from_bytes(bytes, options.archive())
        .map_err(|_error| invalid_source(Path::Package))?;
    super::validate_numbers_application(&components, options.archive())
        .map_err(|_error| invalid_source(Path::Package))?;
    let index = super::Index::from_components(&components, options.semantic().max_objects())
        .map_err(|_error| invalid_source(Path::Package))?;
    let source = components
        .physical()
        .ok_or(Error::UnsupportedSource)?
        .__source_owner();
    let document = Document::from_sheets_with_limits(Vec::new(), DocumentLimits::default())
        .map_err(|_error| invalid_source(Path::Package))?;
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

fn resolve_sheet_position(source: &Package, selector: SheetSelector<'_>) -> Result<usize> {
    let root = Package::root_sheet_order(&source.state.components, SemanticLimits::default())
        .map_err(|_error| invalid_source(Path::Package))?;
    match selector {
        SheetSelector::Index(index) => root
            .sheet_references()
            .get(index)
            .map(|_reference| index)
            .ok_or(Error::SheetNotFound),
        SheetSelector::Name(name) => {
            let mut found = None;
            for index in 0..root.sheet_references().len() {
                let target = resolve_sheet_target(source, index, Path::Package)?;
                if sheet_name(source, target)? == name && found.replace(index).is_some() {
                    return Err(invalid_source(Path::Package));
                }
            }
            found.ok_or(Error::SheetNotFound)
        },
    }
}

fn sheet_name(source: &Package, target: SheetTarget) -> Result<&str> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or_else(|| invalid_source(Path::Package))?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or_else(|| invalid_source(Path::Package))?;
    let payload = match target.message_type {
        SHEET_MESSAGE_TYPE => message.data.as_slice(),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            table_headers::resolve::singular_length_payload(&message.data, 1)
                .map_err(|_error| invalid_source(Path::Package))?
        },
        _ => return Err(invalid_source(Path::Package)),
    };
    super::names::preflight_sheet_name(target.message_type, payload)
        .map_err(|_error| invalid_source(Path::Package))
}

fn resolve_table_position(
    source: &Package,
    sheet_position: usize,
    selector: TableSelector<'_>,
) -> Result<usize> {
    let path = Path::Table {
        source_sheet: sheet_position,
        table: 0,
        destination_sheet: sheet_position,
    };
    match selector {
        TableSelector::Index(index) => {
            table_headers::resolve::resolve_target_physical(source, sheet_position, index)
                .map_err(|error| map_header_error(error, path))?;
            Ok(index)
        },
        TableSelector::Name(name) => {
            let table_count = physical_table_count(source, sheet_position, path)?;
            let mut found = None;
            for index in 0..table_count {
                let target =
                    table_headers::resolve::resolve_target_physical(source, sheet_position, index)
                        .map_err(|error| map_header_error(error, path))?;
                if table_name(source, target, path)? == name && found.replace(index).is_some() {
                    return Err(invalid_source(path));
                }
            }
            found.ok_or(Error::TableNotFound)
        },
    }
}

fn table_name(source: &Package, target: table_headers::Target, path: Path) -> Result<&str> {
    let message = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or_else(|| invalid_source(path))?;
    super::names::preflight_table_name(&message.data).map_err(|_error| invalid_source(path))
}

fn physical_table_count(source: &Package, sheet_position: usize, path: Path) -> Result<usize> {
    let sheet = resolve_sheet_target(source, sheet_position, path)?;
    let message = source
        .state
        .components
        .catalog()
        .get_index(sheet.component_index)
        .and_then(|component| component.archive().objects.get(sheet.object_index))
        .and_then(|object| object.messages.get(sheet.message_index))
        .ok_or_else(|| invalid_source(path))?;
    let payloads = table_headers::resolve::sheet_drawable_payloads(message.type_, &message.data)
        .map_err(|_error| invalid_source(path))?;
    let mut count = 0usize;
    for payload in payloads {
        let identifier = table_headers::resolve::local_reference_identifier(payload)
            .map_err(|_error| invalid_source(path))?;
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(|_error| invalid_source(path))?
            .ok_or_else(|| invalid_source(path))?;
        if table_headers::resolve::unique_table_info(resolved)
            .map_err(|_error| invalid_source(path))?
            .is_some()
        {
            count = count.checked_add(1).ok_or_else(|| invalid_source(path))?;
        }
    }
    Ok(count)
}

fn copy_bytes(source: &[u8], path: Path) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            amount: source.len(),
            path,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

const fn invalid_source(path: Path) -> Error {
    Error::InvalidSource { path }
}
