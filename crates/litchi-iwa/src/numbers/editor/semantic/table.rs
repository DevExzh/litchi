//! Table and cell editing semantics.

use std::io::{self, Write};

use super::*;
use crate::numbers::editor::selectors;
use crate::text::{Alignment, Indents, LineSpacing, Spacing};
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_common::table::cell::{BorderSide, Borders, layout::Layout};
use litchi_numbers::cell::comment::{
    CommentReplyIndex, transaction::Error as FocusedCommentReplyError,
};
use litchi_numbers::table::merge::Region;
use litchi_numbers::{Package as FocusedNumbersPackage, TableCellCommentError};

use crate::numbers::bnc::{BncCellView, CachedScalar, NumericCellType, StoredValue};
use litchi_numbers::cell::CellControl;

type FocusedControlError = litchi_numbers::cell::data_format::control::transaction::Error;

fn focused_control_error(error: FocusedControlError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-control operation failed: {error}"
    ))
}

fn focused_data_format_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell data-format operation failed: {error}"
    ))
}

fn focused_allocation_error(amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation {
        resource: "focused Numbers package output bytes",
        amount,
    })
}

fn parse_focused_source(source_bytes: &[u8]) -> Result<FocusedNumbersPackage> {
    FocusedNumbersPackage::from_bytes(source_bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers cell source validation failed: {error}"
        ))
    })
}

struct FocusedCellLocation {
    source: FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
    source_len: usize,
}

fn focused_cell_location(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedCellLocation> {
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!("invalid Numbers cell coordinate: {error}"))
        })?;
    let (source, source_len) = if let Some(source_bytes) = editor.package.exact_source_bytes() {
        let source_len = source_bytes.len();
        (parse_focused_source(source_bytes)?, source_len)
    } else {
        let source_bytes = editor.to_bytes()?;
        let source_len = source_bytes.len();
        (parse_focused_source(&source_bytes)?, source_len)
    };
    Ok(FocusedCellLocation {
        source,
        sheet,
        table,
        position,
        source_len,
    })
}

struct FalliblePackageBytes {
    bytes: Vec<u8>,
    allocation_failure: Option<usize>,
}

impl FalliblePackageBytes {
    fn with_capacity(capacity: usize) -> Result<Self> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|_| focused_allocation_error(capacity))?;
        Ok(Self {
            bytes,
            allocation_failure: None,
        })
    }
}

impl Write for FalliblePackageBytes {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.try_reserve(bytes.len()).is_err() {
            self.allocation_failure = Some(self.bytes.len().saturating_add(bytes.len()));
            return Err(io::Error::new(
                io::ErrorKind::OutOfMemory,
                "focused Numbers package output allocation failed",
            ));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn focused_package_to_editor(
    package: &FocusedNumbersPackage,
    source_len: usize,
) -> Result<NumbersEditor> {
    let mut output = FalliblePackageBytes::with_capacity(source_len)?;
    if let Err(error) = package.write_to(&mut output) {
        if let Some(amount) = output.allocation_failure {
            return Err(focused_allocation_error(amount));
        }
        return Err(Error::Io(error.into_io_error()));
    }
    NumbersEditor::from_bytes(&output.bytes)
}

fn focused_control_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<CellControl>> {
    let location = focused_cell_location(editor, table_id, row, column)?;
    let FocusedCellLocation {
        source,
        sheet,
        table,
        position,
        ..
    } = location;
    source
        .table_cell_control_format(sheet, table, position)
        .map_err(focused_control_error)
}

fn focused_set_data_format(
    source: &FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
    source_len: usize,
    format: &DataFormat,
) -> Result<Option<NumbersEditor>> {
    macro_rules! commit_set {
        ($edit:expr, $value:expr) => {{
            let commit = $edit
                .map_err(focused_data_format_error)?
                .set($value)
                .commit()
                .map_err(focused_data_format_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        }};
    }

    match format {
        DataFormat::Number(value) => commit_set!(
            source.edit_table_cell_number_format(sheet, table, position),
            *value
        ),
        DataFormat::Percentage(value) => commit_set!(
            source.edit_table_cell_percentage_format(sheet, table, position),
            *value
        ),
        DataFormat::Currency(value) => commit_set!(
            source.edit_table_cell_currency_format(sheet, table, position),
            *value
        ),
        DataFormat::Scientific(value) => commit_set!(
            source.edit_table_cell_scientific_format(sheet, table, position),
            *value
        ),
        DataFormat::Fraction(value) => commit_set!(
            source.edit_table_cell_fraction_format(sheet, table, position),
            *value
        ),
        DataFormat::DateTime(value) => commit_set!(
            source.edit_table_cell_date_time_format(sheet, table, position),
            value.clone()
        ),
        DataFormat::Duration(value) => commit_set!(
            source.edit_table_cell_duration_format(sheet, table, position),
            *value
        ),
        DataFormat::Text(value) => commit_set!(
            source.edit_table_cell_text_format(sheet, table, position),
            *value
        ),
        DataFormat::Custom(value) => {
            let commit = source
                .edit_table_cell_custom_format(sheet, table, position)
                .map_err(focused_data_format_error)?
                .set(value.clone())
                .commit()
                .map_err(focused_data_format_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        },
        DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::PopUpMenu(_) => {
            let control = CellControl::try_from(format.clone()).map_err(|_| {
                Error::InvalidFormat(
                    "focused Numbers control conversion rejected the requested format".to_owned(),
                )
            })?;
            let commit = source
                .edit_table_cell_control_format(sheet, table, position)
                .map_err(focused_control_error)?
                .set(control)
                .commit()
                .map_err(focused_control_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        },
        DataFormat::Automatic | DataFormat::NumeralSystem(_) => Err(Error::InvalidFormat(
            "focused Numbers owner does not publish an automatic or Numeral-System format"
                .to_owned(),
        )),
    }
}

fn focused_clear_data_format(
    source: &FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
    source_len: usize,
    format: &DataFormat,
) -> Result<Option<NumbersEditor>> {
    macro_rules! commit_clear {
        ($edit:expr) => {{
            let commit = $edit
                .map_err(focused_data_format_error)?
                .clear()
                .commit()
                .map_err(focused_data_format_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        }};
    }

    match format {
        DataFormat::Number(_) => {
            commit_clear!(source.edit_table_cell_number_format(sheet, table, position))
        },
        DataFormat::Percentage(_) => {
            commit_clear!(source.edit_table_cell_percentage_format(sheet, table, position))
        },
        DataFormat::Currency(_) => {
            commit_clear!(source.edit_table_cell_currency_format(sheet, table, position))
        },
        DataFormat::Scientific(_) => {
            commit_clear!(source.edit_table_cell_scientific_format(sheet, table, position))
        },
        DataFormat::Fraction(_) => {
            commit_clear!(source.edit_table_cell_fraction_format(sheet, table, position))
        },
        DataFormat::DateTime(_) => {
            commit_clear!(source.edit_table_cell_date_time_format(sheet, table, position))
        },
        DataFormat::Duration(_) => {
            commit_clear!(source.edit_table_cell_duration_format(sheet, table, position))
        },
        DataFormat::Text(_) => {
            commit_clear!(source.edit_table_cell_text_format(sheet, table, position))
        },
        DataFormat::Custom(_) => {
            let commit = source
                .edit_table_cell_custom_format(sheet, table, position)
                .map_err(focused_data_format_error)?
                .clear()
                .commit()
                .map_err(focused_data_format_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        },
        DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::PopUpMenu(_) => {
            let commit = source
                .edit_table_cell_control_format(sheet, table, position)
                .map_err(focused_control_error)?
                .clear()
                .commit()
                .map_err(focused_control_error)?;
            if commit.patch().is_noop() {
                Ok(None)
            } else {
                focused_package_to_editor(commit.package(), source_len).map(Some)
            }
        },
        DataFormat::Automatic | DataFormat::NumeralSystem(_) => Err(Error::InvalidFormat(
            "focused Numbers owner cannot clear an automatic or Numeral-System format".to_owned(),
        )),
    }
}

fn is_focused_control_data_format(format: &DataFormat) -> bool {
    matches!(
        format,
        DataFormat::Checkbox(_)
            | DataFormat::StarRating(_)
            | DataFormat::Slider(_)
            | DataFormat::Stepper(_)
            | DataFormat::PopUpMenu(_)
    )
}

fn is_focused_numeric_data_format(format: &DataFormat) -> bool {
    matches!(
        format,
        DataFormat::Number(_)
            | DataFormat::Percentage(_)
            | DataFormat::Currency(_)
            | DataFormat::Scientific(_)
            | DataFormat::Fraction(_)
    )
}

fn focused_cell_value(
    source: &FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
) -> Result<Option<litchi_numbers::cell::Value>> {
    source
        .table_cell(sheet, table, position)
        .map_err(focused_data_format_error)
        .map(|state| state.storage().value().cloned())
}

fn focused_cell_storage(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Vec<u8>>> {
    let location = locate_attached_cell(&editor.package, table_id, row, column)?;
    read_tile_cell(
        &editor.package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )
}

fn focused_cell_has_numeric_storage(cell: Option<&[u8]>) -> Result<bool> {
    let Some(cell) = cell else {
        return Ok(true);
    };
    let cell = BncCellView::parse(cell).map_err(focused_data_format_error)?;
    let numeric_type = cell.numeric_cell_type();
    let numeric_type = matches!(
        numeric_type,
        Some(NumericCellType::Number | NumericCellType::AlternateNumber)
    );
    let numeric_cache = matches!(cell.cached_scalar(), None | Some(CachedScalar::Number(_)));
    Ok(match cell.stored_value() {
        StoredValue::Empty => cell.numeric_cell_type().is_none() && cell.cached_scalar().is_none(),
        StoredValue::Number | StoredValue::Formula(_) => numeric_type && numeric_cache,
        StoredValue::Text(_)
        | StoredValue::RichText(_)
        | StoredValue::Date
        | StoredValue::Boolean
        | StoredValue::Duration
        | StoredValue::Error
        | StoredValue::Unsupported(_) => false,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FocusedDataFormatFamily {
    Number,
    Percentage,
    Currency,
    Scientific,
    Fraction,
    DateTime,
    Duration,
    Text,
    Control,
}

fn focused_data_format_family(format: &DataFormat) -> Option<FocusedDataFormatFamily> {
    match format {
        DataFormat::Number(_) => Some(FocusedDataFormatFamily::Number),
        DataFormat::Percentage(_) => Some(FocusedDataFormatFamily::Percentage),
        DataFormat::Currency(_) => Some(FocusedDataFormatFamily::Currency),
        DataFormat::Scientific(_) => Some(FocusedDataFormatFamily::Scientific),
        DataFormat::Fraction(_) => Some(FocusedDataFormatFamily::Fraction),
        DataFormat::DateTime(_) => Some(FocusedDataFormatFamily::DateTime),
        DataFormat::Duration(_) => Some(FocusedDataFormatFamily::Duration),
        DataFormat::Text(_) => Some(FocusedDataFormatFamily::Text),
        DataFormat::Custom(_) => None,
        DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::PopUpMenu(_) => Some(FocusedDataFormatFamily::Control),
        DataFormat::Automatic | DataFormat::NumeralSystem(_) => None,
    }
}

fn is_focused_data_format(format: &DataFormat) -> bool {
    focused_data_format_family(format).is_some()
}

fn is_focused_custom_data_format(format: &DataFormat) -> bool {
    matches!(format, DataFormat::Custom(_))
}

/// Distinguish the legacy registry route before selecting a focused owner.
/// Builder packages retain their registry in the TSA base and omit the TN
/// document's field 9. The compatibility reader has already qualified their
/// unique registry. Any present field 9 selects focused validation, including
/// malformed or duplicated references; a focused error never falls back.
fn has_focused_custom_registry_edge(editor: &NumbersEditor) -> Result<bool> {
    editor
        .package
        .with_parsed_archive("Index/Document.iwa", |archive| {
            let mut documents = archive
                .objects
                .iter()
                .filter(|object| object.archive_info.identifier == Some(1))
                .flat_map(|object| object.messages.iter())
                .filter(|message| message.type_ == 1);
            let document = documents.next().ok_or_else(|| {
                Error::InvalidFormat("Numbers Custom-format document is missing".to_owned())
            })?;
            if documents.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers Custom-format document is ambiguous".to_owned(),
                ));
            }
            let limits = litchi_iwa_common::WireLimits::default()
                .with_input_bytes(document.data.len().max(1))?
                .with_fields(document.data.len().max(1))?;
            let view =
                litchi_iwa_common::wire::WireView::parse_with_limits(&document.data, limits)?;
            Ok(view.fields().any(|field| field.number() == 9))
        })
}

fn same_custom_format_family(left: &Custom, right: &Custom) -> bool {
    matches!(
        (left, right),
        (Custom::Number(_), Custom::Number(_))
            | (Custom::Text(_), Custom::Text(_))
            | (Custom::DateTime(_), Custom::DateTime(_))
    )
}

/// Select the focused owner only for an operation represented by that owner.
///
/// Cross-family display-format changes remain a compatibility-host concern;
/// selecting an owner for those changes would turn a typed family refusal
/// into an accidental fallback.  The focused owner explicitly supports
/// transitions between the shared numeric storage families and from an
/// interactive control to one of those families by staging a typed clear and
/// target set. Other cross-family conversions remain compatibility-host work.
fn uses_focused_data_format_owner(current: &DataFormat, requested: &DataFormat) -> bool {
    // The focused Custom owner handles only an existing Custom entry in the
    // same family, or clearing an existing Custom entry.  Generic authoring
    // and cross-family conversions keep the compatibility writer because
    // they may need broader format-list and cell-storage semantics.
    match (current, requested) {
        (DataFormat::Custom(current), DataFormat::Custom(requested)) => {
            return same_custom_format_family(current, requested);
        },
        (DataFormat::Custom(_), DataFormat::Automatic) => return true,
        (_, DataFormat::Custom(_)) => return false,
        (DataFormat::Custom(_), _) => return false,
        _ => {},
    }
    if matches!(requested, DataFormat::Automatic) {
        return is_focused_data_format(current);
    }
    if matches!(requested, DataFormat::NumeralSystem(_)) {
        return false;
    }
    if is_focused_numeric_data_format(requested)
        && (is_focused_control_data_format(current) || is_focused_numeric_data_format(current))
    {
        return true;
    }
    if matches!(
        requested,
        DataFormat::Checkbox(_)
            | DataFormat::StarRating(_)
            | DataFormat::Slider(_)
            | DataFormat::Stepper(_)
            | DataFormat::PopUpMenu(_)
    ) {
        return matches!(current, DataFormat::Automatic) || is_focused_data_format(current);
    }
    matches!(current, DataFormat::Automatic)
        || focused_data_format_family(current) == focused_data_format_family(requested)
}

/// Return the parsed focused source when its native storage shape is eligible.
///
/// Currency and Scientific owners can preserve only empty, numeric, or
/// numeric-cache formula storage.  The compatibility writer has historically
/// handled the other values (including controls and nonnumeric formula
/// caches), so dispatching them to a focused owner would reject an existing
/// operation or require an invented coercion.  The raw storage check runs
/// before focused parsing for numeric families, so an unsupported shape does
/// not allocate a discarded focused package before the compatibility route
/// releases a control. Admitted sources return that parsed package to the
/// commit path; failures after owner selection remain terminal and never fall
/// back to the compatibility writer.
fn focused_data_format_owner_is_eligible(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    current: &DataFormat,
    requested: &DataFormat,
) -> Result<Option<FocusedCellLocation>> {
    if !uses_focused_data_format_owner(current, requested) {
        return Ok(None);
    }
    let needs_numeric_storage =
        matches!(current, DataFormat::Currency(_) | DataFormat::Scientific(_))
            || matches!(
                requested,
                DataFormat::Currency(_) | DataFormat::Scientific(_)
            );
    if needs_numeric_storage {
        let storage = focused_cell_storage(editor, table_id, row, column)?;
        if !focused_cell_has_numeric_storage(storage.as_deref())? {
            return Ok(None);
        }
    }
    Ok(Some(focused_cell_location(editor, table_id, row, column)?))
}

fn commit_exact_focused_data_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    current: &DataFormat,
    requested: &DataFormat,
) -> Result<NumbersEditor> {
    let location = focused_cell_location(editor, table_id, row, column)?;
    commit_exact_focused_data_format_with_location(
        editor, table_id, row, column, current, requested, location,
    )
}

fn commit_exact_focused_data_format_with_location(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    current: &DataFormat,
    requested: &DataFormat,
    location: FocusedCellLocation,
) -> Result<NumbersEditor> {
    let FocusedCellLocation {
        source,
        sheet,
        table,
        position,
        source_len,
    } = location;
    if !is_focused_data_format(current)
        && !is_focused_custom_data_format(current)
        && !matches!(current, DataFormat::Automatic)
    {
        return Err(Error::InvalidFormat(
            "focused Numbers owner does not admit the current cell format".to_owned(),
        ));
    }
    if matches!(requested, DataFormat::Automatic) {
        let candidate =
            focused_clear_data_format(&source, sheet, table, position, source_len, current)?;
        let Some(candidate) = candidate else {
            verify_focused_data_format(editor, table_id, row, column, &DataFormat::Automatic)?;
            return Ok(editor.clone());
        };
        verify_focused_data_format(&candidate, table_id, row, column, &DataFormat::Automatic)?;
        return Ok(candidate);
    }
    if let DataFormat::Custom(requested) = requested {
        let DataFormat::Custom(current) = current else {
            return Err(Error::InvalidFormat(
                "focused Numbers owner does not admit Custom-format authoring".to_owned(),
            ));
        };
        if !same_custom_format_family(current, requested) {
            return Err(Error::InvalidFormat(
                "focused Numbers owner does not admit a cross-family Custom-format conversion"
                    .to_owned(),
            ));
        }
        let requested_format = DataFormat::Custom(requested.clone());
        let candidate = focused_set_data_format(
            &source,
            sheet,
            table,
            position,
            source_len,
            &requested_format,
        )?;
        let Some(candidate) = candidate else {
            verify_focused_data_format(editor, table_id, row, column, &requested_format)?;
            return Ok(editor.clone());
        };
        verify_focused_data_format(&candidate, table_id, row, column, &requested_format)?;
        return Ok(candidate);
    }
    let Some(requested_family) = focused_data_format_family(requested) else {
        return Err(Error::InvalidFormat(
            "focused Numbers owner does not admit the requested cell format".to_owned(),
        ));
    };
    if matches!(requested_family, FocusedDataFormatFamily::Control)
        || matches!(current, DataFormat::Automatic)
        || focused_data_format_family(current) == Some(requested_family)
    {
        let candidate =
            focused_set_data_format(&source, sheet, table, position, source_len, requested)?;
        let Some(candidate) = candidate else {
            verify_focused_data_format(editor, table_id, row, column, requested)?;
            return Ok(editor.clone());
        };
        verify_focused_data_format(&candidate, table_id, row, column, requested)?;
        return Ok(candidate);
    }

    if is_focused_numeric_data_format(requested)
        && (is_focused_control_data_format(current) || is_focused_numeric_data_format(current))
    {
        return commit_focused_data_format_transition(
            table_id, row, column, &source, sheet, table, position, source_len, current, requested,
        );
    }

    Err(Error::InvalidFormat(
        "focused Numbers owner does not admit a cross-family cell-format conversion".to_owned(),
    ))
}

fn commit_focused_data_format_transition(
    table_id: u64,
    row: usize,
    column: usize,
    source: &FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
    source_len: usize,
    current: &DataFormat,
    requested: &DataFormat,
) -> Result<NumbersEditor> {
    let before_value = focused_cell_value(source, sheet, table, position)?;
    let Some(cleared) =
        focused_clear_data_format(source, sheet, table, position, source_len, current)?
    else {
        return Err(Error::InvalidFormat(
            "focused Numbers format transition clear unexpectedly produced a no-op".to_owned(),
        ));
    };
    let FocusedCellLocation {
        source: cleared_source,
        sheet: cleared_sheet,
        table: cleared_table,
        position: cleared_position,
        source_len: cleared_source_len,
    } = focused_cell_location(&cleared, table_id, row, column)?;
    let Some(candidate) = focused_set_data_format(
        &cleared_source,
        cleared_sheet,
        cleared_table,
        cleared_position,
        cleared_source_len,
        requested,
    )?
    else {
        return Err(Error::InvalidFormat(
            "focused Numbers format transition target unexpectedly produced a no-op".to_owned(),
        ));
    };
    verify_focused_data_format(&candidate, table_id, row, column, requested)?;
    let FocusedCellLocation {
        source: candidate_source,
        sheet: candidate_sheet,
        table: candidate_table,
        position: candidate_position,
        ..
    } = focused_cell_location(&candidate, table_id, row, column)?;
    if focused_cell_value(
        &candidate_source,
        candidate_sheet,
        candidate_table,
        candidate_position,
    )? != before_value
    {
        return Err(Error::InvalidFormat(
            "focused Numbers format transition changed the cell value".to_owned(),
        ));
    }
    Ok(candidate)
}

fn verify_focused_data_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    expected: &DataFormat,
) -> Result<()> {
    if editor.table_cell_data_format(table_id, row, column)? != *expected {
        return Err(Error::InvalidFormat(
            "focused Numbers table-cell data format failed package validation".to_owned(),
        ));
    }
    Ok(())
}

/// Apply the legacy native control writer to an editor-owned package.
///
/// Packages produced by `NumbersDocumentBuilder` intentionally do not carry
/// an exact-source owner for the interactive-control graph yet. Keep the
/// focused owner strict for exact packages, while allowing these in-memory
/// compatibility packages to acquire the same canonical control/list state
/// through the existing transactional writer.
fn commit_compatibility_control_format(
    editor: &mut NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: &DataFormat,
) -> Result<()> {
    let mut staged = editor.package.clone();
    cell_data_format::set_cell_data_format(&mut staged, table_id, row, column, format)?;
    staged.validate()?;
    let verified = NumbersEditor::from_package(staged)?;
    if cell_data_format::cell_data_format(&verified.package, table_id, row, column)? != *format {
        return Err(Error::InvalidFormat(
            "Numbers compatibility cell-control write failed package validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

enum FocusedCommentReplacement {
    Published(NumbersEditor),
    LegacyFallback,
}

fn focused_comment_error(error: TableCellCommentError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-comment replacement failed: {error}"
    ))
}

fn replace_cell_comment_with_focused_owner(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    text: &str,
) -> Result<FocusedCommentReplacement> {
    // A package built by the legacy editor has no exact physical owner for
    // this graph.  Select the compatibility host before probing the focused
    // owner; once an exact source is admitted, every structural/ownership
    // failure below remains fail-closed.
    if !editor.package.source_is_exact() {
        return Ok(FocusedCommentReplacement::LegacyFallback);
    }
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!("invalid Numbers comment coordinate: {error}"))
        })?;
    let source_bytes = editor.to_bytes()?;
    let source = FocusedNumbersPackage::from_bytes(&source_bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers comment source validation failed: {error}"
        ))
    })?;
    let commit = match source.set_table_cell_comment(sheet, table, position, text) {
        Ok(commit) => commit,
        Err(TableCellCommentError::CommentNotFound { .. })
        | Err(TableCellCommentError::UnsupportedDependency { .. }) => {
            return Ok(FocusedCommentReplacement::LegacyFallback);
        },
        Err(error) => return Err(focused_comment_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat("could not allocate focused Numbers comment candidate".to_owned())
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    let verified = NumbersEditor::from_bytes(&bytes)?;
    let observed =
        cell_comment_in_package(verified.package(), table_id, row, column)?.ok_or_else(|| {
            Error::InvalidFormat(
                "focused Numbers cell-comment replacement lost the selected comment".to_owned(),
            )
        })?;
    if observed.comment.text != text {
        return Err(Error::InvalidFormat(
            "focused Numbers cell-comment replacement failed legacy readback".to_owned(),
        ));
    }
    Ok(FocusedCommentReplacement::Published(verified))
}

struct FocusedCommentReplySource {
    source: FocusedNumbersPackage,
    sheet: litchi_numbers::SheetSelector<'static>,
    table: litchi_numbers::TableSelector<'static>,
    position: litchi_numbers::table::CellPosition,
    source_bytes: Vec<u8>,
}

enum FocusedCommentReplyPublication {
    Published {
        editor: NumbersEditor,
        reply_id: Option<u64>,
    },
    LegacyFallback,
}

fn focused_comment_reply_source(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedCommentReplySource> {
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    // Re-materialize the checked positional selectors at this boundary so
    // the legacy native ID is resolved exactly once before entering the
    // selector-first package owner.
    let sheet = litchi_numbers::SheetSelector::index(sheet.as_index().ok_or_else(|| {
        Error::InvalidFormat("focused Numbers comment-reply sheet was not positional".to_owned())
    })?);
    let table = litchi_numbers::TableSelector::index(table.as_index().ok_or_else(|| {
        Error::InvalidFormat("focused Numbers comment-reply table was not positional".to_owned())
    })?);
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!("invalid Numbers comment-reply coordinate: {error}"))
        })?;
    let source_bytes = editor.to_bytes()?;
    let source = FocusedNumbersPackage::from_bytes(&source_bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers comment-reply source validation failed: {error}"
        ))
    })?;
    Ok(FocusedCommentReplySource {
        source,
        sheet,
        table,
        position,
        source_bytes,
    })
}

fn focused_comment_reply_error(error: FocusedCommentReplyError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers comment-reply operation failed: {error}"
    ))
}

fn focused_comment_reply_can_fallback(error: FocusedCommentReplyError) -> bool {
    matches!(
        error,
        FocusedCommentReplyError::CommentNotFound { .. }
            | FocusedCommentReplyError::UnsupportedDependency { .. }
            | FocusedCommentReplyError::UnsupportedSource
    )
}

fn publish_focused_comment_reply_package(
    source_bytes: &[u8],
    package: &FocusedNumbersPackage,
) -> Result<NumbersEditor> {
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers comment-reply candidate".to_owned(),
        )
    })?;
    package
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_add_cell_comment_reply(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    text: &str,
) -> Result<FocusedCommentReplyPublication> {
    // Source-built editor packages are an explicitly supported compatibility
    // category.  Route them before focused validation so an exact source's
    // InvalidSource result can never be converted into a legacy write.
    if !editor.package.source_is_exact() {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    }
    let before = cell_comment_replies_in_package(editor.package(), table_id, row, column)?;
    let FocusedCommentReplySource {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?;
    let commit = match source.add_table_cell_comment_reply(sheet, table, position, text) {
        Ok(commit) => commit,
        Err(error) if focused_comment_reply_can_fallback(error) => {
            return Ok(FocusedCommentReplyPublication::LegacyFallback);
        },
        Err(error) => return Err(focused_comment_reply_error(error)),
    };
    let verified = publish_focused_comment_reply_package(&source_bytes, commit.package())?;
    let after = cell_comment_replies_in_package(verified.package(), table_id, row, column)?;
    if after.len() != before.len().saturating_add(1)
        || before.iter().zip(after.iter()).any(|(before, after)| {
            before.storage_id != after.storage_id || before.comment.text != after.comment.text
        })
        || after.last().is_none_or(|reply| reply.comment.text != text)
    {
        return Err(Error::InvalidFormat(
            "focused Numbers comment-reply append failed legacy readback".to_owned(),
        ));
    }
    Ok(FocusedCommentReplyPublication::Published {
        editor: verified,
        reply_id: after.last().map(|reply| reply.storage_id.get()),
    })
}

fn focused_set_cell_comment_reply(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: &str,
) -> Result<FocusedCommentReplyPublication> {
    if !editor.package.source_is_exact() {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    }
    let before = cell_comment_replies_in_package(editor.package(), table_id, row, column)?;
    let Some(ordinal) = before
        .iter()
        .position(|reply| reply.storage_id.get() == reply_storage_object_id)
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
    let index = CommentReplyIndex::try_from_usize(ordinal)
        .map_err(|_| Error::InvalidFormat("Numbers comment-reply ordinal overflow".to_owned()))?;
    let FocusedCommentReplySource {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?;
    let commit = match source.set_table_cell_comment_reply(sheet, table, position, index, text) {
        Ok(commit) => commit,
        Err(error) if focused_comment_reply_can_fallback(error) => {
            return Ok(FocusedCommentReplyPublication::LegacyFallback);
        },
        Err(error) => return Err(focused_comment_reply_error(error)),
    };
    let verified = publish_focused_comment_reply_package(&source_bytes, commit.package())?;
    let after = cell_comment_replies_in_package(verified.package(), table_id, row, column)?;
    if after.len() != before.len()
        || after.iter().enumerate().any(|(index, reply)| {
            if index == ordinal {
                reply.comment.text != text
            } else {
                reply.storage_id != before[index].storage_id
                    || reply.comment.text != before[index].comment.text
            }
        })
    {
        return Err(Error::InvalidFormat(
            "focused Numbers comment-reply replacement failed legacy readback".to_owned(),
        ));
    }
    Ok(FocusedCommentReplyPublication::Published {
        editor: verified,
        reply_id: after.get(ordinal).map(|reply| reply.storage_id.get()),
    })
}

fn focused_remove_cell_comment_reply(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<FocusedCommentReplyPublication> {
    if !editor.package.source_is_exact() {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    }
    let before = cell_comment_replies_in_package(editor.package(), table_id, row, column)?;
    let Some(ordinal) = before
        .iter()
        .position(|reply| reply.storage_id.get() == reply_storage_object_id)
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
    let index = CommentReplyIndex::try_from_usize(ordinal)
        .map_err(|_| Error::InvalidFormat("Numbers comment-reply ordinal overflow".to_owned()))?;
    let FocusedCommentReplySource {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?;
    let commit = match source.remove_table_cell_comment_reply(sheet, table, position, index) {
        Ok(commit) => commit,
        Err(error) if focused_comment_reply_can_fallback(error) => {
            return Ok(FocusedCommentReplyPublication::LegacyFallback);
        },
        Err(error) => return Err(focused_comment_reply_error(error)),
    };
    let verified = publish_focused_comment_reply_package(&source_bytes, commit.package())?;
    let after = cell_comment_replies_in_package(verified.package(), table_id, row, column)?;
    if after.len().saturating_add(1) != before.len()
        || after
            .iter()
            .any(|reply| reply.storage_id.get() == reply_storage_object_id)
        || after.iter().enumerate().any(|(index, reply)| {
            let before_index = if index < ordinal { index } else { index + 1 };
            reply.storage_id != before[before_index].storage_id
                || reply.comment.text != before[before_index].comment.text
        })
    {
        return Err(Error::InvalidFormat(
            "focused Numbers comment-reply removal failed legacy readback".to_owned(),
        ));
    }
    Ok(FocusedCommentReplyPublication::Published {
        editor: verified,
        reply_id: None,
    })
}

impl NumbersEditor {
    /// List absolute pivot categories backed by valid calculation-engine
    /// aggregate coordinates.
    pub fn pivot_categories(&self) -> Result<Vec<NumbersPivotCategoryInfo>> {
        let mut categories = formula_pivot_categories(&self.package)?
            .into_iter()
            .map(|(key, value)| NumbersPivotCategoryInfo {
                reference: FormulaPivotCategoryReference::new(
                    key.group_by_uid,
                    key.column_uid,
                    key.group_uid,
                    value.aggregate_type,
                    value.group_level,
                ),
                label: value.label,
            })
            .collect::<Vec<_>>();
        categories.sort_by(|left, right| {
            left.reference
                .group_by_uid
                .cmp(&right.reference.group_by_uid)
                .then_with(|| left.reference.column_uid.cmp(&right.reference.column_uid))
                .then_with(|| left.reference.group_level.cmp(&right.reference.group_level))
                .then_with(|| left.label.cmp(&right.label))
                .then_with(|| left.reference.group_uid.cmp(&right.reference.group_uid))
        });
        Ok(categories)
    }

    /// Read the explicit data format for one zero-based table cell.
    pub fn table_cell_data_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<DataFormat> {
        let format = cell_data_format::cell_data_format(&self.package, table_id, row, column)?;
        if CellControl::try_from(format.clone()).is_ok() {
            return focused_control_format(self, table_id, row, column)?.map_or_else(
                || {
                    Err(Error::InvalidFormat(
                        "focused Numbers cell-control read lost the selected format".to_owned(),
                    ))
                },
                |format| Ok(format.into_data_format()),
            );
        }
        Ok(format)
    }

    /// Create, replace, or reset one cell's typed data format transactionally.
    pub fn set_table_cell_data_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: DataFormat,
    ) -> Result<()> {
        let source_built = !self.package.source_is_exact();
        let current = cell_data_format::cell_data_format(&self.package, table_id, row, column)?;
        let legacy_custom_registry = !source_built
            && is_focused_custom_data_format(&current)
            && uses_focused_data_format_owner(&current, &format)
            && !has_focused_custom_registry_edge(self)?;
        if !source_built
            && current == format
            && !is_focused_data_format(&format)
            && (!is_focused_custom_data_format(&format) || legacy_custom_registry)
        {
            // Preserve exact-source no-op bytes for formats without an
            // eligible focused owner: Automatic, NumeralSystem, and legacy
            // Custom registries without the TN document edge.
            // Re-running the native writer here would needlessly allocate a
            // new package and can normalize inherited automatic metadata.
            return Ok(());
        }
        let focused_owner_context = if !source_built
            && current == format
            && is_focused_data_format(&format)
        {
            // A focused no-op still has to enter the focused transaction so
            // malformed family metadata keeps its terminal refusal
            // semantics and a valid no-op can retain the exact source. Shape
            // admission is intentionally skipped for this forced validation.
            Some(focused_cell_location(self, table_id, row, column)?)
        } else if !source_built && !legacy_custom_registry {
            focused_data_format_owner_is_eligible(self, table_id, row, column, &current, &format)?
        } else {
            None
        };
        if let Some(location) = focused_owner_context {
            *self = commit_exact_focused_data_format_with_location(
                self, table_id, row, column, &current, &format, location,
            )?;
            return Ok(());
        }

        // Unsupported exact-source control-to-scalar conversions still have
        // two owners. Release the focused control graph first, then let the
        // compatibility writer replace the unsupported display metadata in
        // the private snapshot. Supported numeric targets were selected by
        // the focused branch above and never reach this path.
        if !source_built
            && CellControl::try_from(current.clone()).is_ok()
            && !matches!(format, DataFormat::Automatic)
        {
            let mut staged = commit_exact_focused_data_format(
                self,
                table_id,
                row,
                column,
                &current,
                &DataFormat::Automatic,
            )?;
            cell_data_format::set_cell_data_format(
                &mut staged.package,
                table_id,
                row,
                column,
                &format,
            )?;
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != format {
                return Err(Error::InvalidFormat(
                    "Numbers table-cell data format failed package validation".to_owned(),
                ));
            }
            *self = verified;
            return Ok(());
        }

        // Generated packages and cross-family/unsupported exact-source
        // formats still use the native compatibility writer.  The focused
        // owner branches above are terminal once selected, so a focused error
        // cannot silently fall back to a raw-ID mutation.
        if source_built
            && (CellControl::try_from(format.clone()).is_ok()
                || CellControl::try_from(current.clone()).is_ok())
        {
            commit_compatibility_control_format(self, table_id, row, column, &format)?;
            return Ok(());
        }
        let mut staged = self.package.clone();
        cell_data_format::set_cell_data_format(&mut staged, table_id, row, column, &format)?;
        // Keep generated packages source-built through the compatibility
        // mutation. Reopening their bytes would incorrectly turn the
        // normalized builder output into an exact-source package and disable
        // the narrowly-scoped focused-owner fallback on the next operation.
        let verified = if source_built {
            staged.validate()?;
            Self::from_package(staged)?
        } else {
            Self::from_bytes(&staged.to_bytes()?)?
        };
        if verified.table_cell_data_format(table_id, row, column)? != format {
            return Err(Error::InvalidFormat(
                "Numbers table-cell data format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read a named custom Number, Date & Time, or Text format.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_custom_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Custom>> {
        cell_data_format::cell_custom_format(&self.package, table_id, row, column)
    }

    /// Create or replace a named custom format transactionally.
    pub fn set_table_cell_custom_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Custom,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from a named custom format.
    pub fn reset_table_cell_custom_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let current = cell_data_format::cell_data_format(&self.package, table_id, row, column)?;
        if !matches!(&current, DataFormat::Custom(_)) {
            return match current {
                DataFormat::Automatic => Ok(false),
                _ => Err(Error::InvalidFormat(
                    "Cannot reset Custom format from a non-Custom cell".to_owned(),
                )),
            };
        }

        // The focused owner is admitted only for an exact source carrying the
        // rooted TN.DocumentArchive field-9 registry edge.  Legacy builder
        // packages keep the registry under the TSA field-12 edge; preserving
        // that profile through the compatibility writer is required for both
        // in-memory builder snapshots and their reopened bytes.  A present
        // but malformed field-9 edge remains a focused terminal error.
        if self.package.source_is_exact() && has_focused_custom_registry_edge(self)? {
            let location = focused_data_format_owner_is_eligible(
                self,
                table_id,
                row,
                column,
                &current,
                &DataFormat::Automatic,
            )?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "focused Numbers owner rejected the existing Custom format".to_owned(),
                )
            })?;
            *self = commit_exact_focused_data_format_with_location(
                self,
                table_id,
                row,
                column,
                &current,
                &DataFormat::Automatic,
                location,
            )?;
            return Ok(true);
        }

        let mut staged = self.package.clone();
        // The Custom family was checked above; avoid decoding it again in
        // the compatibility convenience wrapper.
        let changed = cell_data_format::reset_cell_data_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != DataFormat::Automatic {
                return Err(Error::InvalidFormat(
                    "Numbers Custom-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit positional numeral-system format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_numeral_system_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<NumeralSystem>> {
        cell_data_format::cell_numeral_system_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit positional numeral-system format transactionally.
    pub fn set_table_cell_numeral_system_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: NumeralSystem,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Numeral System cell.
    pub fn reset_table_cell_numeral_system_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_numeral_system_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != DataFormat::Automatic {
                return Err(Error::InvalidFormat(
                    "Numbers numeral-system reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective text layout for one zero-based table cell.
    pub fn table_cell_layout(&self, table_id: u64, row: usize, column: usize) -> Result<Layout> {
        cell_layout::cell_layout(&self.package, table_id, row, column)
    }

    /// Create or replace local text-layout overrides for one table cell.
    pub fn set_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        layout: Layout,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_layout::set_cell_layout(&mut staged, table_id, row, column, layout)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_layout(table_id, row, column)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers table-cell layout failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local text-layout overrides and restore inherited cell values.
    pub fn reset_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_layout::reset_cell_layout(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective horizontal text alignment for one zero-based table cell.
    pub fn table_cell_text_alignment(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Alignment> {
        cell_paragraph_style::alignment(&self.package, table_id, row, column)
    }

    /// Create or replace a local horizontal text-alignment override.
    pub fn set_table_cell_text_alignment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        alignment: Alignment,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_alignment(&mut staged, table_id, row, column, alignment)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_alignment(table_id, row, column)? != alignment {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text alignment failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local horizontal alignment and restore the inherited table style.
    pub fn reset_table_cell_text_alignment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_alignment(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph line spacing for one table cell.
    pub fn table_cell_paragraph_line_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<LineSpacing> {
        cell_paragraph_style::line_spacing(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell paragraph line-spacing override.
    pub fn set_table_cell_paragraph_line_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: LineSpacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_line_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_line_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph line spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local line spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_line_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_line_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing for one table cell.
    pub fn table_cell_paragraph_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Spacing> {
        cell_paragraph_style::spacing(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell before/after paragraph spacing.
    pub fn set_table_cell_paragraph_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: Spacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local before/after spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the canonical list preset applied uniformly to a table cell.
    pub fn table_cell_paragraph_list(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphList> {
        cell_paragraph_list::paragraph_list(&self.package, table_id, row, column)
    }

    /// Promote a plain text cell when necessary and apply one native list preset.
    pub fn set_table_cell_paragraph_list(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        list: NumbersTableCellParagraphList,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list(&mut staged, table_id, row, column, list)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list(table_id, row, column)? != list {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the canonical None list preset for a table cell.
    pub fn reset_table_cell_paragraph_list(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_list::reset_paragraph_list(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read all paragraph-scoped list preset boundaries in a table cell.
    pub fn table_cell_paragraph_lists(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<NumbersTableCellParagraphListPlacement>> {
        cell_paragraph_list::paragraph_lists(&self.package, table_id, row, column)
    }

    /// Promote a plain cell when necessary and replace all list preset boundaries.
    pub fn set_table_cell_paragraph_lists(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        placements: &[NumbersTableCellParagraphListPlacement],
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_lists(&mut staged, table_id, row, column, placements)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let expected = cell_paragraph_list::paragraph_lists(&staged, table_id, row, column)?;
        if verified.table_cell_paragraph_lists(table_id, row, column)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph-list placements failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read every effective list-level boundary in a table cell.
    pub fn table_cell_paragraph_list_levels(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<NumbersTableCellParagraphListLevelPlacement>> {
        cell_paragraph_list::paragraph_list_levels(&self.package, table_id, row, column)
    }

    /// Set one validated paragraph's list level without changing later paragraphs.
    pub fn set_table_cell_paragraph_list_level(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        level: NumbersTableCellParagraphListLevel,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_level(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            level,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if cell_paragraph_list::paragraph_list_level(
            &verified.package,
            table_id,
            row,
            column,
            paragraph,
        )? != level
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list level failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_table_cell_paragraph_list_level(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_level(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read whether one table-cell paragraph continues or restarts list numbering.
    pub fn table_cell_paragraph_list_numbering(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListNumbering> {
        cell_paragraph_list::paragraph_list_numbering(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Continue or restart numbered-list sequencing at one table-cell paragraph.
    pub fn set_table_cell_paragraph_list_numbering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        numbering: NumbersTableCellParagraphListNumbering,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_numbering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            numbering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if cell_paragraph_list::paragraph_list_numbering(
            &verified.package,
            table_id,
            row,
            column,
            paragraph,
        )? != numbering
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list numbering failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one numbered table-cell paragraph's effective label format.
    pub fn table_cell_paragraph_list_number_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListNumberFormat> {
        cell_paragraph_list::paragraph_list_number_format(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered table-cell paragraph's locale-aware label format.
    pub fn set_table_cell_paragraph_list_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        format: NumbersTableCellParagraphListNumberFormat,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_format(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_format(table_id, row, column, paragraph)?
            != format
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number format failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_table_cell_paragraph_list_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_format(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read whether one numbered table-cell paragraph displays hierarchical numbering.
    pub fn table_cell_paragraph_list_number_tiering(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListNumberTiering> {
        cell_paragraph_list::paragraph_list_number_tiering(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Choose flat or hierarchical numbering for one table-cell list level.
    pub fn set_table_cell_paragraph_list_number_tiering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        tiering: NumbersTableCellParagraphListNumberTiering,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_tiering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            tiering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_tiering(table_id, row, column, paragraph)?
            != tiering
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number tiering failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore flat numbering for one table-cell list level.
    pub fn reset_table_cell_paragraph_list_number_tiering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_tiering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one numbered table-cell paragraph's number-label size.
    pub fn table_cell_paragraph_list_number_scale(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListNumberScale> {
        cell_paragraph_list::paragraph_list_number_scale(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered table-cell paragraph's number-label size.
    pub fn set_table_cell_paragraph_list_number_scale(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        scale: NumbersTableCellParagraphListNumberScale,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_scale(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            scale,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_scale(table_id, row, column, paragraph)?
            != scale
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number scale failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_table_cell_paragraph_list_number_scale(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_scale(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell paragraph's effective text-bullet marker.
    pub fn table_cell_paragraph_list_bullet(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListBullet> {
        cell_paragraph_list::paragraph_list_bullet(&self.package, table_id, row, column, paragraph)
    }

    /// Set one table-cell paragraph's text-bullet marker.
    pub fn set_table_cell_paragraph_list_bullet(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        bullet: &NumbersTableCellParagraphListBullet,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_bullet(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            bullet,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_bullet(table_id, row, column, paragraph)? != *bullet {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph text bullet failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one table-cell paragraph.
    pub fn reset_table_cell_paragraph_list_bullet(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_bullet(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell paragraph's effective bullet size and baseline.
    pub fn table_cell_paragraph_list_bullet_geometry(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListBulletGeometry> {
        cell_paragraph_list::paragraph_list_bullet_geometry(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell paragraph's bullet size and baseline.
    pub fn set_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        geometry: NumbersTableCellParagraphListBulletGeometry,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_bullet_geometry(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            geometry,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_bullet_geometry(table_id, row, column, paragraph)?
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph bullet geometry failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_bullet_geometry(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell list paragraph's label and text-gap indentation.
    pub fn table_cell_paragraph_list_indentation(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListIndentation> {
        cell_paragraph_list::paragraph_list_indentation(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell list paragraph's label and text-gap indentation.
    pub fn set_table_cell_paragraph_list_indentation(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        indentation: NumbersTableCellParagraphListIndentation,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_indentation(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            indentation,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_indentation(table_id, row, column, paragraph)?
            != indentation
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list indentation failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_table_cell_paragraph_list_indentation(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_indentation(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell list paragraph's effective label color.
    pub fn table_cell_paragraph_list_label_color(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<NumbersTableCellParagraphListLabelColor> {
        cell_paragraph_list::paragraph_list_label_color(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell list paragraph's bullet or number color.
    pub fn set_table_cell_paragraph_list_label_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        color: NumbersTableCellParagraphListLabelColor,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_label_color(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            color,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_label_color(table_id, row, column, paragraph)?
            != color
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-label color failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_table_cell_paragraph_list_label_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_label_color(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right paragraph indents.
    pub fn table_cell_paragraph_indents(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Indents> {
        cell_paragraph_style::indents(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell paragraph indents.
    pub fn set_table_cell_paragraph_indents(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        indents: Indents,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_indents(&mut staged, table_id, row, column, indents)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_indents(table_id, row, column)? != indents {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph indents failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local paragraph indents and restore the inherited table style.
    pub fn reset_table_cell_paragraph_indents(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_indents(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the ordered explicit ruler tab stops for one table cell.
    pub fn table_cell_paragraph_tab_stops(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphTabStops> {
        cell_paragraph_style::tab_stops(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell ruler tab stops.
    pub fn set_table_cell_paragraph_tab_stops(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        stops: NumbersTableCellParagraphTabStops,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_tab_stops(&mut staged, table_id, row, column, stops.clone())?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_tab_stops(table_id, row, column)? != stops {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph tab stops failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local ruler tab stops and restore the inherited table style.
    pub fn reset_table_cell_paragraph_tab_stops(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_tab_stops(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective background painted behind one table cell's text.
    pub fn table_cell_text_background(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextBackground> {
        cell_paragraph_style::background(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text-background override.
    pub fn set_table_cell_text_background(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        background: NumbersTableCellTextBackground,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_background(&mut staged, table_id, row, column, background)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_background(table_id, row, column)? != background {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text background failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text background and restore the inherited value.
    pub fn reset_table_cell_text_background(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_background(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of one table cell.
    pub fn table_cell_text_baseline_shift(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextBaselineShift> {
        cell_paragraph_style::baseline_shift(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell custom baseline displacement.
    pub fn set_table_cell_text_baseline_shift(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        shift: NumbersTableCellTextBaselineShift,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_baseline_shift(&mut staged, table_id, row, column, shift)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_baseline_shift(table_id, row, column)? != shift {
            return Err(Error::InvalidFormat(
                "Numbers table-cell baseline shift failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local baseline displacement and restore the inherited value.
    pub fn reset_table_cell_text_baseline_shift(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_baseline_shift(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective capitalization of one table cell.
    pub fn table_cell_text_capitalization(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextCapitalization> {
        cell_paragraph_style::capitalization(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell capitalization.
    pub fn set_table_cell_text_capitalization(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        capitalization: NumbersTableCellTextCapitalization,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_capitalization(
            &mut staged,
            table_id,
            row,
            column,
            capitalization,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_capitalization(table_id, row, column)? != capitalization {
            return Err(Error::InvalidFormat(
                "Numbers table-cell capitalization failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local capitalization and restore the inherited value.
    pub fn reset_table_cell_text_capitalization(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_capitalization(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of one table cell.
    pub fn table_cell_text_character_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextCharacterSpacing> {
        cell_paragraph_style::character_spacing(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell character spacing.
    pub fn set_table_cell_text_character_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: NumbersTableCellTextCharacterSpacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_character_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_character_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell character spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local character spacing and restore the inherited value.
    pub fn reset_table_cell_text_character_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_character_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective foreground text color of one table cell.
    pub fn table_cell_text_color(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextColor> {
        cell_paragraph_style::text_color(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell foreground text-color override.
    pub fn set_table_cell_text_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        color: NumbersTableCellTextColor,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_text_color(&mut staged, table_id, row, column, color)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_color(table_id, row, column)? != color {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text color failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text-color override and restore the inherited color.
    pub fn reset_table_cell_text_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_text_color(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective whole-cell underline and strikethrough formatting.
    pub fn table_cell_text_decorations(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextDecorations> {
        cell_paragraph_style::decorations(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell underline and strikethrough formatting.
    pub fn set_table_cell_text_decorations(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        decorations: NumbersTableCellTextDecorations,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_decorations(&mut staged, table_id, row, column, decorations)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_decorations(table_id, row, column)? != decorations {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text decorations failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local decorations and restore the inherited cell formatting.
    pub fn reset_table_cell_text_decorations(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_decorations(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of one table cell.
    pub fn table_cell_text_font(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextFont> {
        cell_paragraph_style::font(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell PostScript font override.
    pub fn set_table_cell_text_font(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        font: NumbersTableCellTextFont,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_font(&mut staged, table_id, row, column, font.clone())?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_font(table_id, row, column)? != font {
            return Err(Error::InvalidFormat(
                "Numbers table-cell font failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local font override and restore the inherited table font.
    pub fn reset_table_cell_text_font(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_font(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of one table cell.
    pub fn table_cell_text_ligatures(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextLigatures> {
        cell_paragraph_style::ligatures(&self.package, table_id, row, column)
    }

    /// Create or replace the whole-cell ligature policy.
    pub fn set_table_cell_text_ligatures(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        ligatures: NumbersTableCellTextLigatures,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_ligatures(&mut staged, table_id, row, column, ligatures)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_ligatures(table_id, row, column)? != ligatures {
            return Err(Error::InvalidFormat(
                "Numbers table-cell ligatures failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local ligature policy and restore the inherited value.
    pub fn reset_table_cell_text_ligatures(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_ligatures(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective outline of one table cell's text.
    pub fn table_cell_text_outline(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextOutline> {
        cell_paragraph_style::outline(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text outline.
    pub fn set_table_cell_text_outline(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        outline: NumbersTableCellTextOutline,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_outline(&mut staged, table_id, row, column, outline)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_outline(table_id, row, column)? != outline {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text outline failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text outline and restore the inherited value.
    pub fn reset_table_cell_text_outline(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_outline(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective normal, superscript, or subscript formatting.
    pub fn table_cell_text_script(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextScript> {
        cell_paragraph_style::script(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell baseline script formatting.
    pub fn set_table_cell_text_script(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        script: NumbersTableCellTextScript,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_script(&mut staged, table_id, row, column, script)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_script(table_id, row, column)? != script {
            return Err(Error::InvalidFormat(
                "Numbers table-cell baseline script failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local baseline script formatting and restore the inherited value.
    pub fn reset_table_cell_text_script(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_script(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective drop shadow of one table cell's text.
    pub fn table_cell_text_shadow(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextShadow> {
        cell_paragraph_style::shadow(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text drop shadow.
    pub fn set_table_cell_text_shadow(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        shadow: NumbersTableCellTextShadow,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_shadow(&mut staged, table_id, row, column, shadow)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_shadow(table_id, row, column)? != shadow {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text shadow failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text shadow and restore the inherited value.
    pub fn reset_table_cell_text_shadow(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_shadow(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective whole-cell point size, bold, and italic formatting.
    pub fn table_cell_text_style(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextStyle> {
        cell_paragraph_style::text_style(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell point size, bold, and italic formatting.
    pub fn set_table_cell_text_style(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        style: NumbersTableCellTextStyle,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_text_style(&mut staged, table_id, row, column, style)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_style(table_id, row, column)? != style {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text style failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local point size, bold, and italic formatting.
    pub fn reset_table_cell_text_style(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_text_style(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective fill for one zero-based table cell.
    pub fn table_cell_fill(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::shapes::ShapeFill> {
        cell_fill::cell_fill(&self.package, table_id, row, column)
    }

    /// Create or replace a local table-cell fill transactionally.
    pub fn set_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        fill: &crate::shapes::ShapeFill,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_fill::set_cell_fill(&mut staged, table_id, row, column, fill)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if &verified.table_cell_fill(table_id, row, column)? != fill {
            return Err(Error::InvalidFormat(
                "Numbers table-cell fill failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct fill override and restore the inherited table style.
    pub fn reset_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_fill::reset_cell_fill(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective explicit borders for one zero-based table cell.
    pub fn table_cell_borders(&self, table_id: u64, row: usize, column: usize) -> Result<Borders> {
        stroke_layers::cell_borders(&self.package, table_id, row, column)
    }

    /// Create or replace one explicit table-cell border transactionally.
    pub fn set_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
        stroke: Stroke,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, Some(stroke))
    }

    /// Explicitly clear one table-cell border transactionally.
    pub fn clear_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, None)
    }

    fn update_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
        stroke: Option<Stroke>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        stroke_layers::set_cell_border(&mut staged, table_id, row, column, side, stroke)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .table_cell_borders(table_id, row, column)?
            .get(side)
            != stroke
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell border failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// List every native merged-cell rectangle in one attached table.
    pub fn table_cell_merges(&self, table_id: u64) -> Result<Vec<Region>> {
        cell_merge::regions_in_package(&self.package, table_id)
    }

    /// Merge one non-overlapping rectangular cell region transactionally.
    pub fn merge_cells(&mut self, table_id: u64, region: Region) -> Result<()> {
        let mut staged = self.package.clone();
        cell_merge::merge_in_package(&mut staged, table_id, region)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if !verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell merge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove one exact merged-cell rectangle, returning whether it existed.
    pub fn unmerge_cells(&mut self, table_id: u64, region: Region) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_merge::unmerge_in_package(&mut staged, table_id, region)?;
        if !changed {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell unmerge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read the comment attached to a writable BNC cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment API; use litchi_numbers::Package::table_cell_comment with SheetSelector and TableSelector for the focused semantic root-comment read; unsupported graph creation and replies remain migration-host scope"
    )]
    pub fn cell_comment(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<TableCellComment>> {
        cell_comment_in_package(&self.package, table_id, row, column)
    }

    /// Create or replace a cell comment without changing the cell value or style.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment API; use litchi_numbers::Package::edit_table_cell_comment or set_table_cell_comment with SheetSelector and TableSelector for supported strict-graph creation and replacement; unsupported graph creation and reply identity remain migration-host scope"
    )]
    pub fn set_cell_comment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let text = text.into();
        match replace_cell_comment_with_focused_owner(self, table_id, row, column, &text)? {
            FocusedCommentReplacement::Published(verified) => {
                *self = verified;
                return Ok(());
            },
            FocusedCommentReplacement::LegacyFallback => {},
        }
        let mut staged = self.package.clone();
        set_cell_comment_in_package(&mut staged, table_id, row, column, text)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Inspect the conditional-highlight style set attached to a writable BNC cell.
    pub fn cell_conditional_highlighting(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<TableCellConditionalHighlightInfo>> {
        conditional_highlight::info_in_package(&self.package, table_id, row, column)
    }

    /// Read the supported ordered conditional-highlight rules attached to a cell.
    pub fn cell_conditional_highlight_rules(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Vec<Rule>>> {
        conditional_highlight::rules_in_package(&self.package, table_id, row, column)
    }

    /// Delete conditional highlighting from one cell without changing its value or base style.
    pub fn clear_cell_conditional_highlighting(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        conditional_highlight::clear_in_package(&mut staged, table_id, row, column)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .cell_conditional_highlighting(table_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a cell's conditional highlighting and return its storage identity.
    pub fn set_cell_conditional_highlighting(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        rules: &[Rule],
    ) -> Result<TableCellConditionalHighlightInfo> {
        let mut staged = self.package.clone();
        conditional_highlight::set_in_package(&mut staged, table_id, row, column, rules)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let actual = verified
            .cell_conditional_highlighting(table_id, row, column)?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-highlight creation failed validation".to_owned(),
                )
            })?;
        if actual.rule_count as usize != rules.len() {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight rule count failed validation".to_owned(),
            ));
        }
        if verified
            .cell_conditional_highlight_rules(table_id, row, column)?
            .as_deref()
            != Some(rules)
        {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight rules failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(actual)
    }

    /// Read the direct replies attached to a cell comment in stored order.
    ///
    /// This compatibility method retains native table, root-storage, and
    /// reply-storage identifiers in its return value. Prefer the selector-first
    /// `litchi_numbers::Package::table_cell_comment_replies` projection for an
    /// ID-free semantic read.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment-reply read; use litchi_numbers::Package::table_cell_comment_replies with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read; reply identity remains compatibility-host scope"
    )]
    pub fn cell_comment_replies(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<TableCellReply>> {
        cell_comment_replies_in_package(&self.package, table_id, row, column)
    }

    /// Append a direct reply to an existing cell comment.
    ///
    /// Supported exact graphs delegate to the selector-first package owner;
    /// this deprecated raw-ID surface remains as a compatibility fallback.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment-reply creation API; use litchi_numbers::Package::add_table_cell_comment_reply with SheetSelector, TableSelector, and CellPosition for selector-first writes; this method delegates supported graphs and retains a compatibility fallback"
    )]
    pub fn add_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<u64> {
        let text = text.into();
        match focused_add_cell_comment_reply(self, table_id, row, column, &text)? {
            FocusedCommentReplyPublication::Published { editor, reply_id } => {
                *self = editor;
                return reply_id.ok_or_else(|| {
                    Error::InvalidFormat(
                        "focused Numbers comment-reply append returned no identity".to_owned(),
                    )
                });
            },
            FocusedCommentReplyPublication::LegacyFallback => {},
        }
        let mut staged = self.package.clone();
        let reply_id = add_cell_comment_reply_in_package(&mut staged, table_id, row, column, text)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Replace one direct reply and return its new copy-on-write object ID.
    ///
    /// Supported exact graphs delegate to the selector-first package owner;
    /// this deprecated raw-ID surface remains as a compatibility fallback.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment-reply replacement API; use litchi_numbers::Package::set_table_cell_comment_reply with SheetSelector, TableSelector, CellPosition, and CommentReplyIndex for selector-first writes; this method delegates supported graphs and retains a compatibility fallback"
    )]
    pub fn set_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        let text = text.into();
        match focused_set_cell_comment_reply(
            self,
            table_id,
            row,
            column,
            reply_storage_object_id,
            &text,
        )? {
            FocusedCommentReplyPublication::Published { editor, reply_id } => {
                *self = editor;
                return reply_id.ok_or_else(|| {
                    Error::InvalidFormat(
                        "focused Numbers comment-reply replacement returned no identity".to_owned(),
                    )
                });
            },
            FocusedCommentReplyPublication::LegacyFallback => {},
        }
        let mut staged = self.package.clone();
        let reply_id = set_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
            text,
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Remove one direct reply from an existing cell comment.
    ///
    /// Supported exact graphs delegate to the selector-first package owner;
    /// this deprecated raw-ID surface remains as a compatibility fallback.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell-comment-reply removal API; use litchi_numbers::Package::remove_table_cell_comment_reply with SheetSelector, TableSelector, CellPosition, and CommentReplyIndex for selector-first writes; this method delegates supported graphs and retains a compatibility fallback"
    )]
    pub fn remove_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        match focused_remove_cell_comment_reply(
            self,
            table_id,
            row,
            column,
            reply_storage_object_id,
        )? {
            FocusedCommentReplyPublication::Published { editor, .. } => {
                *self = editor;
                return Ok(());
            },
            FocusedCommentReplyPublication::LegacyFallback => {},
        }
        let mut staged = self.package.clone();
        remove_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }
}

#[cfg(test)]
mod focused_custom_tests {

    use super::*;
    use crate::numbers::cell::CellValue;
    use crate::numbers::{NumbersDocumentBuilder, NumbersEditor};
    use litchi_numbers::cell::data_format::custom::{
        Custom, Name, Number as CustomNumber, NumberPattern, Text as CustomText,
    };
    use litchi_numbers::{CellPosition, SheetSelector, TableSelector};
    use prost::Message as _;

    fn custom_number(name: &str, pattern: &str) -> Custom {
        Custom::Number(CustomNumber::new(
            Name::try_new(name).expect("valid custom name"),
            NumberPattern::try_new(pattern).expect("valid custom pattern"),
        ))
    }

    #[test]
    fn exact_custom_same_family_replacement_clear_and_noop_use_focused_owner() {
        let source_bytes = include_bytes!(
            "../../../../../../test-data/iwork/synthetic/numbers/custom-focused.numbers"
        );
        let source = FocusedNumbersPackage::from_bytes(source_bytes).expect("focused source");
        let position = CellPosition::new(0, 0);
        let before_value = source
            .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
            .expect("original focused cell")
            .storage()
            .value()
            .cloned();
        let replacement = custom_number("Signed Integer", "#,##0;(#,##0)");
        let direct = source
            .edit_table_cell_custom_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .expect("focused custom edit")
            .set(replacement.clone())
            .commit()
            .expect("focused custom replacement");
        assert_eq!(
            direct
                .package()
                .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
                .expect("direct replacement cell")
                .storage()
                .value()
                .cloned(),
            before_value
        );

        let mut actual = NumbersEditor::from_bytes(source_bytes).expect("exact editor");
        // The focused fixture intentionally omits the appearance dependency
        // required by `NumbersEditor::tables`; its table-model object is still
        // the stable native table id used by the cell APIs.
        let table_id = 4;
        assert!(matches!(
            actual
                .table_cell_data_format(table_id, 0, 0)
                .expect("original format read"),
            DataFormat::Custom(Custom::Number(_))
        ));
        actual
            .set_table_cell_data_format(table_id, 0, 0, replacement.clone().into())
            .expect("host custom replacement");
        let host_replacement_bytes = actual.to_bytes().expect("host replacement bytes");
        let host_replacement =
            FocusedNumbersPackage::from_bytes(&host_replacement_bytes).expect("host focused");
        assert_eq!(
            actual
                .table_cell_data_format(table_id, 0, 0)
                .expect("replacement read"),
            DataFormat::Custom(replacement.clone())
        );
        assert_eq!(
            host_replacement
                .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
                .expect("host replacement cell")
                .storage()
                .value()
                .cloned(),
            before_value
        );
        assert_eq!(
            host_replacement
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    position,
                )
                .expect("host replacement custom"),
            Some(replacement.clone())
        );
        assert!(
            host_replacement
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    CellPosition::new(0, 1),
                )
                .expect("shared registry custom")
                .is_some()
        );
        let original_editor = NumbersEditor::from_bytes(source_bytes).expect("original editor");
        assert_eq!(
            actual.package().entry("Index/Unrelated.iwa"),
            original_editor.package().entry("Index/Unrelated.iwa")
        );
        assert_eq!(
            actual.package().entry("Data/data-format-sentinel.bin"),
            original_editor
                .package()
                .entry("Data/data-format-sentinel.bin")
        );

        let before_noop = actual.to_bytes().expect("replacement source bytes");
        let source_pointer = actual
            .package()
            .exact_source_bytes()
            .expect("focused replacement source")
            .as_ptr();
        actual
            .set_table_cell_data_format(table_id, 0, 0, replacement.into())
            .expect("host custom no-op");
        assert_eq!(actual.to_bytes().expect("host no-op bytes"), before_noop);
        assert_eq!(
            actual
                .package()
                .exact_source_bytes()
                .expect("focused no-op source")
                .as_ptr(),
            source_pointer
        );

        let replacement_source =
            FocusedNumbersPackage::from_bytes(&before_noop).expect("focused replacement source");
        let direct_clear = replacement_source
            .edit_table_cell_custom_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .expect("focused custom clear edit")
            .clear()
            .commit()
            .expect("focused custom clear");
        assert_eq!(
            direct_clear
                .package()
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    position,
                )
                .expect("direct clear custom"),
            None
        );
        actual
            .set_table_cell_data_format(table_id, 0, 0, DataFormat::Automatic)
            .expect("host custom clear");
        let host_clear = actual.to_bytes().expect("host clear bytes");
        let host_clear = FocusedNumbersPackage::from_bytes(&host_clear).expect("clear focused");
        assert_eq!(
            actual
                .table_cell_data_format(table_id, 0, 0)
                .expect("clear read"),
            DataFormat::Automatic
        );
        assert_eq!(
            host_clear
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    position,
                )
                .expect("host clear custom"),
            None
        );
        assert!(
            host_clear
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    CellPosition::new(0, 1),
                )
                .expect("shared registry after clear")
                .is_some()
        );
        assert_eq!(
            host_clear
                .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
                .expect("host clear cell")
                .storage()
                .value()
                .cloned(),
            before_value
        );
    }

    #[test]
    fn exact_custom_reset_uses_focused_owner_and_preserves_registry_value_and_locality() {
        let source_bytes = include_bytes!(
            "../../../../../../test-data/iwork/synthetic/numbers/custom-focused.numbers"
        );
        let position = CellPosition::new(0, 0);
        let source = FocusedNumbersPackage::from_bytes(source_bytes).expect("focused source");
        let before_value = source
            .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
            .expect("original focused cell")
            .storage()
            .value()
            .cloned();
        let original = NumbersEditor::from_bytes(source_bytes).expect("exact editor");

        let mut editor = original.clone();
        assert!(
            editor
                .reset_table_cell_custom_format(4, 0, 0)
                .expect("focused custom reset")
        );
        assert_eq!(
            editor
                .table_cell_data_format(4, 0, 0)
                .expect("focused reset format"),
            DataFormat::Automatic
        );
        assert_eq!(
            editor
                .table_cell_custom_format(4, 0, 0)
                .expect("focused reset custom"),
            None
        );
        let reset = FocusedNumbersPackage::from_bytes(&editor.to_bytes().expect("reset bytes"))
            .expect("focused reset package");
        assert_eq!(
            reset
                .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
                .expect("reset focused cell")
                .storage()
                .value()
                .cloned(),
            before_value
        );
        assert!(
            reset
                .table_cell_custom_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    CellPosition::new(0, 1),
                )
                .expect("shared registry after focused reset")
                .is_some()
        );
        assert_eq!(
            editor.package().entry("Index/Unrelated.iwa"),
            original.package().entry("Index/Unrelated.iwa")
        );
        assert_eq!(
            editor.package().entry("Data/data-format-sentinel.bin"),
            original.package().entry("Data/data-format-sentinel.bin")
        );

        let before_noop = editor.to_bytes().expect("automatic bytes");
        assert!(
            !editor
                .reset_table_cell_custom_format(4, 0, 0)
                .expect("automatic reset no-op")
        );
        assert_eq!(
            editor.to_bytes().expect("automatic no-op bytes"),
            before_noop
        );
    }

    #[test]
    fn builder_custom_reopen_keeps_compatibility_profile() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .expect("builder source");
        let table_id = editor.tables().expect("builder table")[0].id();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            0,
            0,
            CellValue::number(42.0).expect("finite test number"),
        )
        .expect("builder cell value");
        let initial = custom_number("Builder Integer", "#,##0");
        editor
            .set_table_cell_data_format(table_id, 0, 0, initial.clone().into())
            .expect("builder custom format");

        let mut source_built_reset = editor.clone();
        assert!(
            source_built_reset
                .reset_table_cell_custom_format(table_id, 0, 0)
                .expect("source-built custom reset")
        );
        assert_eq!(
            source_built_reset
                .table_cell_data_format(table_id, 0, 0)
                .expect("source-built reset format"),
            DataFormat::Automatic
        );

        let source_bytes = editor.to_bytes().expect("builder bytes");
        let mut reopened = NumbersEditor::from_bytes(&source_bytes).expect("reopened builder");
        assert_eq!(
            reopened
                .table_cell_custom_format(table_id, 0, 0)
                .expect("builder custom read"),
            Some(initial)
        );

        let replacement = custom_number("Builder Signed Integer", "#,##0;(#,##0)");
        reopened
            .set_table_cell_data_format(table_id, 0, 0, replacement.clone().into())
            .expect("builder custom replacement");
        assert_eq!(
            reopened
                .table_cell_custom_format(table_id, 0, 0)
                .expect("builder replacement read"),
            Some(replacement.clone())
        );
        let before_value =
            FocusedNumbersPackage::from_bytes(&reopened.to_bytes().expect("replacement bytes"))
                .expect("focused builder replacement")
                .table_cell(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    CellPosition::new(0, 0),
                )
                .expect("builder replacement cell")
                .storage()
                .value()
                .cloned();
        assert_eq!(
            before_value,
            Some(CellValue::number(42.0).expect("finite expected number"))
        );

        let before_noop = reopened.to_bytes().expect("builder no-op source");
        let source_pointer = reopened
            .package()
            .exact_source_bytes()
            .expect("builder no-op source bytes")
            .as_ptr();
        reopened
            .set_table_cell_data_format(table_id, 0, 0, replacement.clone().into())
            .expect("builder exact no-op");
        assert_eq!(
            reopened.to_bytes().expect("builder no-op bytes"),
            before_noop
        );
        assert_eq!(
            reopened
                .package()
                .exact_source_bytes()
                .expect("builder no-op retained source")
                .as_ptr(),
            source_pointer
        );

        let mut reset_clone = reopened.clone();
        assert!(
            reset_clone
                .reset_table_cell_custom_format(table_id, 0, 0)
                .expect("builder custom reset")
        );
        assert_eq!(
            reset_clone
                .table_cell_custom_format(table_id, 0, 0)
                .expect("builder reset read"),
            None
        );
        assert_eq!(
            reset_clone
                .table_cell_data_format(table_id, 0, 0)
                .expect("builder reset format"),
            DataFormat::Automatic
        );

        let mut non_custom = reopened.clone();
        non_custom
            .set_table_cell_data_format(table_id, 0, 0, DataFormat::Number(Number::default()))
            .expect("builder Number format");
        let non_custom_bytes = non_custom.to_bytes().expect("Number source bytes");
        let error = non_custom
            .reset_table_cell_custom_format(table_id, 0, 0)
            .expect_err("Number reset must reject non-Custom format");
        assert!(error.to_string().contains("non-Custom cell"));
        assert_eq!(
            non_custom.to_bytes().expect("Number source unchanged"),
            non_custom_bytes
        );

        reopened
            .set_table_cell_data_format(table_id, 0, 0, DataFormat::Automatic)
            .expect("builder custom clear");
        assert_eq!(
            reopened
                .table_cell_custom_format(table_id, 0, 0)
                .expect("builder clear read"),
            None
        );
        let cleared =
            FocusedNumbersPackage::from_bytes(&reopened.to_bytes().expect("builder clear bytes"))
                .expect("focused builder clear");
        assert_eq!(
            cleared
                .table_cell(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    CellPosition::new(0, 0),
                )
                .expect("builder clear cell")
                .storage()
                .value()
                .cloned(),
            Some(CellValue::number(42.0).expect("finite expected number"))
        );
    }

    #[test]
    fn focused_custom_malformed_refcount_refuses_after_admission() {
        let source_bytes = include_bytes!(
            "../../../../../../test-data/iwork/synthetic/numbers/custom-focused.numbers"
        );
        let mut malformed = NumbersEditor::from_bytes(source_bytes).expect("focused source");
        malformed
            .package
            .update_archive("Index/Tables.iwa", |archive| {
                let format_list_type =
                    litchi_iwa_protos::tst::table_data_list::ListType::Format as i32;
                let mut segment_ids = Vec::new();
                {
                    let object = archive.object_mut(5).expect("format sidecar");
                    let message = object
                        .messages
                        .iter_mut()
                        .find(|message| {
                            message.type_ == 6_005
                                && litchi_iwa_protos::tst::TableDataList::decode(
                                    message.data.as_slice(),
                                )
                                .map(|list| list.list_type == format_list_type)
                                .unwrap_or(false)
                        })
                        .expect("format list message");
                    let mut list =
                        litchi_iwa_protos::tst::TableDataList::decode(message.data.as_slice())
                            .expect("format list payload");
                    if let Some(entry) = list.entries.iter_mut().next() {
                        entry.refcount = 1;
                        message.data = list.encode_to_vec();
                        return Ok(());
                    }
                    segment_ids.extend(list.segments.iter().map(|reference| reference.identifier));
                }
                for segment_id in segment_ids {
                    let object = archive.object_mut(segment_id).expect("format segment");
                    let Some(message) = object.messages.iter_mut().find(|message| {
                        message.type_ == 6_005
                            && litchi_iwa_protos::tst::TableDataList::decode(
                                message.data.as_slice(),
                            )
                            .map(|list| list.list_type == format_list_type)
                            .unwrap_or(false)
                    }) else {
                        continue;
                    };
                    let mut list =
                        litchi_iwa_protos::tst::TableDataList::decode(message.data.as_slice())
                            .expect("format segment payload");
                    if let Some(entry) = list.entries.iter_mut().next() {
                        entry.refcount = 1;
                        message.data = list.encode_to_vec();
                        return Ok(());
                    }
                }
                Err(Error::InvalidFormat(
                    "format fixture has no mutable format-list entry".to_owned(),
                ))
            })
            .expect("malformed source mutation");
        let malformed_bytes = malformed.to_bytes().expect("malformed bytes");
        let mut actual = NumbersEditor::from_bytes(&malformed_bytes).expect("malformed editor");
        let current = actual
            .table_cell_data_format(4, 0, 0)
            .expect("malformed custom read");
        assert!(matches!(&current, &DataFormat::Custom(Custom::Number(_))));
        let before = actual.to_bytes().expect("malformed source baseline");
        let error = actual
            .set_table_cell_data_format(4, 0, 0, current.clone())
            .expect_err("malformed focused custom no-op must refuse");
        assert!(
            error
                .to_string()
                .contains("focused Numbers cell data-format")
        );
        assert_eq!(
            actual
                .to_bytes()
                .expect("malformed source unchanged after no-op"),
            before
        );
        let error = actual
            .reset_table_cell_custom_format(4, 0, 0)
            .expect_err("malformed focused custom reset must refuse");
        assert!(
            error
                .to_string()
                .contains("focused Numbers cell data-format")
        );
        assert_eq!(
            actual
                .to_bytes()
                .expect("malformed source unchanged after reset"),
            before
        );
        let error = actual
            .set_table_cell_data_format(4, 0, 0, custom_number("Rejected", "#,##0;(#,##0)").into())
            .expect_err("malformed focused graph must refuse");
        assert!(
            error
                .to_string()
                .contains("focused Numbers cell data-format")
        );
        assert_eq!(
            actual.to_bytes().expect("malformed source unchanged"),
            before
        );
    }

    #[test]
    fn malformed_present_registry_edge_cannot_select_legacy_custom_writer() {
        let source_bytes = include_bytes!(
            "../../../../../../test-data/iwork/synthetic/numbers/custom-focused.numbers"
        );
        let mut malformed = NumbersEditor::from_bytes(source_bytes).expect("focused source");
        malformed
            .package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive.object_mut(1).expect("document object");
                let message = object
                    .messages
                    .iter_mut()
                    .find(|message| message.type_ == 1)
                    .expect("document message");
                // A duplicate registry edge must remain on strict focused admission.
                message.data.extend_from_slice(&[0x4a, 0x02, 0x08, 0x01]);
                Ok(())
            })
            .expect("duplicate registry edge");
        let bytes = malformed.to_bytes().expect("malformed source bytes");
        let mut editor = NumbersEditor::from_bytes(&bytes).expect("discovery defers registry");
        assert!(has_focused_custom_registry_edge(&editor).expect("present edge"));
        let current = editor
            .table_cell_data_format(4, 0, 0)
            .expect("legacy custom read");
        for requested in [
            current,
            DataFormat::Automatic,
            custom_number("Rejected", "0.00").into(),
        ] {
            let error = editor
                .set_table_cell_data_format(4, 0, 0, requested)
                .expect_err("present malformed edge is terminal");
            assert!(
                error
                    .to_string()
                    .contains("focused Numbers cell data-format")
            );
            assert_eq!(editor.to_bytes().expect("unchanged source"), bytes);
        }
        let error = editor
            .reset_table_cell_custom_format(4, 0, 0)
            .expect_err("present malformed edge reset is terminal");
        assert!(
            error
                .to_string()
                .contains("focused Numbers cell data-format")
        );
        assert_eq!(
            editor.to_bytes().expect("unchanged source after reset"),
            bytes
        );
    }

    #[test]
    fn focused_custom_owner_selection_preserves_family_boundaries() {
        let number = custom_number("Signed Integer", "#,##0;(#,##0)");
        let other_number = custom_number("Grouped Integer", "#,##0");
        let text = Custom::Text(
            CustomText::try_new(Name::try_new("Text Prefix").unwrap(), "ID: ", "").unwrap(),
        );

        assert!(uses_focused_data_format_owner(
            &DataFormat::Custom(number.clone()),
            &DataFormat::Custom(other_number),
        ));
        assert!(uses_focused_data_format_owner(
            &DataFormat::Custom(number.clone()),
            &DataFormat::Automatic,
        ));
        assert!(!uses_focused_data_format_owner(
            &DataFormat::Custom(number.clone()),
            &DataFormat::Custom(text.clone()),
        ));
        assert!(!uses_focused_data_format_owner(
            &DataFormat::Automatic,
            &DataFormat::Custom(number.clone()),
        ));
        assert!(!uses_focused_data_format_owner(
            &DataFormat::Custom(number),
            &DataFormat::Number(Number::default()),
        ));
    }
}
