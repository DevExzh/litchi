//! Numbers package ingress and metadata-only merged-cell selection.
//!
//! Physical sources and borrowed name-codec access belong here. The parent
//! module owns selector-first geometry reads and the shared merge-wire budget.

use std::{fmt, path::Path, sync::Arc};

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::numbers_names_codec;

use super::super::{Components, Index, ReadOptions, read_source, validate_numbers_application};
use super::{
    Budget, DOCUMENT_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE, SHEET_MESSAGE_TYPE,
    TableMergesError, TableMergesLimitKind, decode_model_merges, length_payload, limit,
    map_package_error, read_reference, require_declared_reference, resolve_model_payload,
    resolve_object, scan_fields, table_model_identifier, unique_message, unique_sheet_message,
    unique_table_info, unique_table_model, validate_message_metadata,
};
use crate::{SheetSelector, TableSelector, table::merge::Region};

/// A lazy, metadata-only Numbers merged-cell reader.
///
/// `MergeReader` retains the validated physical component catalog and compact
/// object index needed for one or more focused queries.  It deliberately does
/// not construct the rooted [`Document`](crate::Document), semantic tables,
/// or decoded cell/BNC values.  Each query resolves the requested sheet and
/// table through their visible names or checked positions, then decodes only
/// the selected table model's borrowed merge payload.
///
/// The reader is read-only and owns no public native identifiers or generated
/// protobuf values.  Use [`crate::Package::table_merges`] when a fully materialized
/// [`crate::Package`] is already available.
///
/// Physical package framing and object indexing are validated at construction.
/// Queries validate the metadata needed for selection and the selected merge
/// path; they do not validate unrelated cell contents. The physical component
/// catalog still retains those payload bytes. Cell and output-text
/// materialization limits do not apply because this reader produces neither.
///
/// ```no_run
/// use litchi_numbers::MergeReader;
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let reader = MergeReader::open("budget.numbers")?;
/// let regions = reader.table_merges("Summary", "Revenue")?;
/// for region in regions {
///     println!("{region:?}");
/// }
/// # Ok(())
/// # }
/// ```
pub struct MergeReader {
    state: Arc<MergeReaderState>,
}

struct MergeReaderState {
    components: Components,
    index: Index,
    options: ReadOptions,
}

impl fmt::Debug for MergeReader {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MergeReader")
            .field("options", &self.state.options)
            .finish_non_exhaustive()
    }
}

impl Clone for MergeReader {
    fn clone(&self) -> Self {
        Self {
            state: Arc::clone(&self.state),
        }
    }
}

impl MergeReader {
    /// Open a Numbers package for lazy merged-cell queries using default
    /// limits.
    ///
    /// The source is streamed through the Numbers bounded filesystem ingress
    /// before its physical components are decoded; no semantic document or
    /// cell projection is constructed.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the path cannot be read, physical ingress is
    /// over budget, or the package is not a valid Numbers source.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, TableMergesError> {
        Self::open_with_options(path, ReadOptions::default())
    }

    /// Open a Numbers package for lazy merged-cell queries under explicit
    /// physical and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes_with_options`],
    /// plus bounded filesystem ingress failures.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: ReadOptions,
    ) -> Result<Self, TableMergesError> {
        let bytes = read_source(path.as_ref(), options.archive()).map_err(map_package_error)?;
        let components =
            Components::from_owned_bytes(bytes, options.archive()).map_err(map_package_error)?;
        Self::from_components_with_options(components, options)
    }

    /// Parse a Numbers package for lazy merged-cell queries.
    ///
    /// Physical package and IWA limits use their defaults.  Unlike
    /// [`crate::Package::from_bytes`], this constructor stops after application
    /// validation and object indexing; it does not build a semantic document
    /// or materialize cells.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the source is not a Numbers package, its
    /// physical contents are malformed, or the compact object index exceeds
    /// the default semantic object ceiling.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, TableMergesError> {
        Self::from_bytes_with_options(bytes, ReadOptions::default())
    }

    /// Parse a Numbers package for lazy merged-cell queries under explicit
    /// physical and semantic limits.
    ///
    /// The selected semantic profile governs selector traversal, native
    /// metadata inspection, and the borrowed merge result.  No rooted cell or
    /// BNC projection is constructed.
    ///
    /// # Errors
    ///
    /// Returns a typed error when physical ingress, Numbers application
    /// validation, object indexing, or a selected limit fails.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: ReadOptions,
    ) -> Result<Self, TableMergesError> {
        let components =
            Components::from_bytes(bytes, options.archive()).map_err(map_package_error)?;
        Self::from_components_with_options(components, options)
    }

    /// Parse shared package bytes without copying the ZIP source allocation.
    ///
    /// This is useful when a coordinator already retains an immutable source
    /// allocation.  The reader still retains only the physical component
    /// catalog and compact object index required by merge queries.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes_with_options`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: ReadOptions,
    ) -> Result<Self, TableMergesError> {
        let components =
            Components::from_shared_bytes(bytes, options.archive()).map_err(map_package_error)?;
        Self::from_components_with_options(components, options)
    }

    /// Parse already-owned shared package bytes with default limits.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes`].
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self, TableMergesError> {
        Self::from_shared_bytes_with_options(bytes, ReadOptions::default())
    }

    fn from_components_with_options(
        components: Components,
        options: ReadOptions,
    ) -> Result<Self, TableMergesError> {
        validate_numbers_application(&components, options.archive()).map_err(map_package_error)?;
        let index = Index::from_components(&components, options.semantic().max_objects())
            .map_err(map_package_error)?;
        Ok(Self {
            state: Arc::new(MergeReaderState {
                components,
                index,
                options,
            }),
        })
    }

    /// Read all validated merged-cell regions for one selected table.
    ///
    /// Selectors use exact visible names, checked zero-based positions, or
    /// typed semantic positions. Queries decode the metadata needed to resolve
    /// selectors and the selected merge payload. They do not materialize cells
    /// or full semantic table objects.
    ///
    /// # Errors
    ///
    /// Returns [`TableMergesError::SheetNotFound`],
    /// [`TableMergesError::TableNotFound`], or
    /// [`TableMergesError::AmbiguousSelector`] for selector failures, and a
    /// typed source or resource error for malformed selected metadata.
    pub fn table_merges<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Vec<Region>, TableMergesError> {
        let mut budget = Budget::from_options(self.state.options)?;
        let (sheet_position, table_position) = select_metadata_positions(
            &self.state.components,
            &self.state.index,
            sheet.into(),
            table.into(),
            &mut budget,
        )?;
        let model = resolve_model_payload(
            &self.state.components,
            &self.state.index,
            sheet_position,
            table_position,
            &mut budget,
        )?;
        decode_model_merges(model, &mut budget)
    }
}

/// Resolve selectors against only the rooted document, sheet, and table
/// metadata needed by a focused query.  This deliberately mirrors the
/// semantic order used by [`crate::Package::table_merges`] without constructing the
/// package's cell-bearing document projection.
fn select_metadata_positions(
    components: &Components,
    index: &Index,
    sheet_selector: SheetSelector<'_>,
    table_selector: TableSelector<'_>,
    budget: &mut Budget,
) -> Result<(usize, usize), TableMergesError> {
    let document_archive = components
        .get_archive("Index/Document.iwa")
        .ok_or(TableMergesError::UnsupportedSource)?;
    budget.charge_work(document_archive.objects.len())?;
    let document = document_archive
        .object(1)
        .ok_or(TableMergesError::InvalidSource)?;
    if document.archive_info.identifier != Some(1) {
        return Err(TableMergesError::InvalidSource);
    }
    budget.charge_objects(1)?;
    let (document_index, document_message) =
        unique_message(document, DOCUMENT_MESSAGE_TYPE, budget)?
            .ok_or(TableMergesError::InvalidSource)?;
    validate_message_metadata(document, document_index, budget)?;

    let mut sheet_count = 0usize;
    let mut selected_sheet = None;
    scan_fields(&document_message.data, 0, budget, |field, budget| {
        if field.number() != 1 {
            return Ok(());
        }
        let sheet_payload = length_payload(field)?;
        let sheet_identifier = read_reference(sheet_payload, 1, budget)?;
        let sheet_position = sheet_count;
        sheet_count = sheet_count
            .checked_add(1)
            .ok_or(TableMergesError::InvalidSource)?;
        budget.charge_sheets(1)?;
        let inspect = match sheet_selector {
            SheetSelector::Name(_) => true,
            SheetSelector::Index(index) => index == sheet_position,
        };
        if !inspect {
            return Ok(());
        }
        let name = metadata_sheet_name(components, index, sheet_identifier, budget)?;
        if let SheetSelector::Name(expected) = sheet_selector {
            if name != expected {
                return Ok(());
            }
            if selected_sheet.is_some() {
                return Err(TableMergesError::AmbiguousSelector);
            }
        }
        require_declared_reference(document, document_index, sheet_identifier, &[1], budget)?;
        selected_sheet = Some((sheet_position, sheet_identifier));
        Ok(())
    })?;
    let (sheet_position, sheet_identifier) =
        selected_sheet.ok_or(TableMergesError::SheetNotFound)?;
    ensure_single_document_sheet_edge(&document_message.data, sheet_identifier, budget)?;

    let table_position =
        select_metadata_table(components, index, sheet_identifier, table_selector, budget)?;
    Ok((sheet_position, table_position))
}

fn name_decode_options(
    budget: &Budget,
) -> Result<numbers_names_codec::DecodeOptions, TableMergesError> {
    let input = budget.remaining_input()?;
    Ok(numbers_names_codec::DecodeOptions::new(
        input,
        budget.remaining_fields()?,
        budget.remaining_work()?,
        u32::try_from(budget.wire_max_nesting).map_err(|_| TableMergesError::InvalidSource)?,
    )
    .with_total_input_bytes(input))
}

fn map_names_error(error: numbers_names_codec::DecodeError) -> TableMergesError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return limit(TableMergesLimitKind::WireFields, observed, maximum);
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return limit(TableMergesLimitKind::WireWork, observed, maximum);
    }
    match error.wire_resource_limit() {
        Some(numbers_names_codec::WireResourceLimit::Bytes { observed, maximum }) => {
            limit(TableMergesLimitKind::WireBytes, observed, maximum)
        },
        Some(numbers_names_codec::WireResourceLimit::Nesting { observed, maximum }) => {
            TableMergesError::LimitExceeded {
                kind: TableMergesLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        _ => TableMergesError::InvalidSource,
    }
}

fn metadata_sheet_name<'source>(
    components: &'source Components,
    index: &'source Index,
    identifier: u64,
    budget: &mut Budget,
) -> Result<&'source str, TableMergesError> {
    let sheet = resolve_object(components, index, identifier, budget)?;
    let (message_index, message) = unique_sheet_message(sheet, budget)?;
    validate_message_metadata(sheet, message_index, budget)?;
    let options = name_decode_options(budget)?;
    let (snapshot, report) = match message.type_ {
        SHEET_MESSAGE_TYPE => {
            numbers_names_codec::decode_sheet_name_with_report(&message.data, options)
        },
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            numbers_names_codec::decode_form_sheet_name_with_report(&message.data, options)
        },
        _ => return Err(TableMergesError::InvalidSource),
    }
    .map_err(map_names_error)?;
    let name = snapshot.name();
    budget.charge_wire_report(report.input_bytes(), report.fields(), report.work())?;
    budget.charge_work(name.len().saturating_add(1))?;
    budget.charge_items(1)?;
    Ok(name)
}

fn select_metadata_table(
    components: &Components,
    index: &Index,
    sheet_identifier: u64,
    table_selector: TableSelector<'_>,
    budget: &mut Budget,
) -> Result<usize, TableMergesError> {
    let sheet = resolve_object(components, index, sheet_identifier, budget)?;
    let (sheet_message_index, sheet_message) = unique_sheet_message(sheet, budget)?;
    validate_message_metadata(sheet, sheet_message_index, budget)?;
    let mut semantic_table = 0usize;
    let mut selected = None;
    let mut selected_drawable = None;
    scan_sheet_drawables(
        sheet_message.type_,
        &sheet_message.data,
        budget,
        |drawable_identifier, budget| {
            if selected.is_some() && matches!(table_selector, TableSelector::Index(_)) {
                return Ok(());
            }
            let info_object = resolve_object(components, index, drawable_identifier, budget)?;
            let Some((info_index, info_message)) = unique_table_info(info_object, budget)? else {
                return Ok(());
            };
            validate_message_metadata(info_object, info_index, budget)?;
            if info_object.archive_info.identifier != Some(drawable_identifier) {
                return Err(TableMergesError::InvalidSource);
            }
            let table_position = semantic_table;
            semantic_table = semantic_table
                .checked_add(1)
                .ok_or(TableMergesError::InvalidSource)?;
            budget.charge_tables(1)?;
            match table_selector {
                TableSelector::Index(expected) => {
                    if table_position == expected {
                        selected = Some(table_position);
                        selected_drawable = Some(drawable_identifier);
                        require_declared_reference(
                            sheet,
                            sheet_message_index,
                            drawable_identifier,
                            sheet_reference_path(sheet_message.type_),
                            budget,
                        )?;
                    }
                },
                TableSelector::Name(expected) => {
                    let model = metadata_table_model(
                        components,
                        index,
                        sheet_identifier,
                        drawable_identifier,
                        info_object,
                        info_index,
                        info_message,
                        sheet,
                        sheet_message_index,
                        sheet_message.type_,
                        budget,
                    )?;
                    let options = name_decode_options(budget)?;
                    let (snapshot, report) =
                        numbers_names_codec::decode_table_names_with_report(model, options)
                            .map_err(map_names_error)?;
                    let name = snapshot.table_name();
                    budget.charge_wire_report(
                        report.input_bytes(),
                        report.fields(),
                        report.work(),
                    )?;
                    budget.charge_work(name.len().saturating_add(1))?;
                    budget.charge_items(1)?;
                    if name != expected {
                        return Ok(());
                    }
                    if selected.is_some() {
                        return Err(TableMergesError::AmbiguousSelector);
                    }
                    selected = Some(table_position);
                    selected_drawable = Some(drawable_identifier);
                },
            }
            Ok(())
        },
    )?;
    let selected = selected.ok_or(TableMergesError::TableNotFound)?;
    let selected_drawable = selected_drawable.ok_or(TableMergesError::InvalidSource)?;
    ensure_single_table_info_edge(
        sheet_message.type_,
        &sheet_message.data,
        selected_drawable,
        budget,
    )?;
    Ok(selected)
}

/// Prove that a selected sheet is rooted exactly once in the document payload.
///
/// Archive metadata also records this edge, but it can remain unchanged when a
/// malformed payload repeats a protobuf reference.  Re-scan the bounded
/// document message so selector results cannot depend on which duplicate was
/// visited first.
fn ensure_single_document_sheet_edge(
    source: &[u8],
    expected: u64,
    budget: &mut Budget,
) -> Result<(), TableMergesError> {
    let mut occurrences = 0usize;
    scan_fields(source, 0, budget, |field, budget| {
        if field.number() != 1 {
            return Ok(());
        }
        let identifier = read_reference(length_payload(field)?, 1, budget)?;
        if identifier == expected {
            occurrences = occurrences
                .checked_add(1)
                .ok_or(TableMergesError::InvalidSource)?;
        }
        Ok(())
    })?;
    if occurrences == 1 {
        Ok(())
    } else {
        Err(TableMergesError::InvalidSource)
    }
}

/// Prove that a selected table-info edge occurs exactly once in its sheet
/// payload.  The focused reader intentionally does not build the semantic
/// table vector, so this bounded metadata pass supplies the ownership check
/// that a full projection normally performs before selecting a table.
fn ensure_single_table_info_edge(
    message_type: u32,
    source: &[u8],
    expected: u64,
    budget: &mut Budget,
) -> Result<(), TableMergesError> {
    let mut occurrences = 0usize;
    scan_sheet_drawables(message_type, source, budget, |identifier, _budget| {
        if identifier == expected {
            occurrences = occurrences
                .checked_add(1)
                .ok_or(TableMergesError::InvalidSource)?;
        }
        Ok(())
    })?;
    if occurrences == 1 {
        Ok(())
    } else {
        Err(TableMergesError::InvalidSource)
    }
}

fn metadata_table_model<'source>(
    components: &'source Components,
    index: &'source Index,
    sheet_identifier: u64,
    drawable_identifier: u64,
    info_object: &'source ArchiveObject,
    info_index: usize,
    info_message: &'source RawMessage,
    sheet: &ArchiveObject,
    sheet_message_index: usize,
    sheet_message_type: u32,
    budget: &mut Budget,
) -> Result<&'source [u8], TableMergesError> {
    let model_identifier = table_model_identifier(&info_message.data, budget)?;
    require_declared_reference(info_object, info_index, model_identifier, &[2], budget)?;
    require_declared_reference(
        sheet,
        sheet_message_index,
        drawable_identifier,
        sheet_reference_path(sheet_message_type),
        budget,
    )?;
    if sheet_identifier == drawable_identifier
        || sheet_identifier == model_identifier
        || drawable_identifier == model_identifier
    {
        return Err(TableMergesError::InvalidSource);
    }
    let model_object = resolve_object(components, index, model_identifier, budget)?;
    let (model_index, model_message) = unique_table_model(model_object, budget)?;
    validate_message_metadata(model_object, model_index, budget)?;
    if model_object.archive_info.identifier != Some(model_identifier) {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(model_message.data.as_slice())
}

const fn sheet_reference_path(message_type: u32) -> &'static [u32] {
    match message_type {
        SHEET_MESSAGE_TYPE => &[2],
        FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
        _ => &[],
    }
}

fn scan_sheet_drawables<F>(
    message_type: u32,
    source: &[u8],
    budget: &mut Budget,
    mut visitor: F,
) -> Result<(), TableMergesError>
where
    F: FnMut(u64, &mut Budget) -> Result<(), TableMergesError>,
{
    match message_type {
        SHEET_MESSAGE_TYPE => scan_fields(source, 0, budget, |field, budget| {
            if field.number() != 2 {
                return Ok(());
            }
            let payload = length_payload(field)?;
            let identifier = read_reference(payload, 1, budget)?;
            visitor(identifier, budget)
        }),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            let mut super_payload = None;
            scan_fields(source, 0, budget, |field, _budget| {
                if field.number() != 1 {
                    return Ok(());
                }
                if super_payload.is_some() {
                    return Err(TableMergesError::InvalidSource);
                }
                super_payload = Some(length_payload(field)?);
                Ok(())
            })?;
            let super_payload = super_payload.ok_or(TableMergesError::InvalidSource)?;
            scan_fields(super_payload, 1, budget, |field, budget| {
                if field.number() != 2 {
                    return Ok(());
                }
                let payload = length_payload(field)?;
                let identifier = read_reference(payload, 2, budget)?;
                visitor(identifier, budget)
            })
        },
        _ => Err(TableMergesError::InvalidSource),
    }
}
