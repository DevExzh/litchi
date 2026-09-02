//! Table and cell editing semantics.

use super::*;
use crate::numbers::editor::selectors;
use crate::numbers::editor::table::cell::Borders;
use crate::text::{Alignment, Indents, LineSpacing, Spacing};
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_common::table::cell::{BorderSide, layout::Layout};
use litchi_numbers::cell::comment::{
    CommentReplyIndex, transaction::Error as FocusedCommentReplyError,
};
use litchi_numbers::table::merge::Region;
use litchi_numbers::{Package as FocusedNumbersPackage, TableCellCommentError};

use litchi_numbers::cell::CellControl;

type FocusedControlError = litchi_numbers::cell::data_format::control::transaction::Error;
type FocusedNumberFormatError = litchi_numbers::cell::data_format::number::transaction::Error;
type FocusedPercentageFormatError =
    litchi_numbers::cell::data_format::percentage::transaction::Error;
type FocusedCurrencyFormatError = litchi_numbers::cell::data_format::currency::transaction::Error;
type FocusedScientificFormatError =
    litchi_numbers::cell::data_format::scientific::transaction::Error;
type FocusedFractionFormatError = litchi_numbers::cell::data_format::fraction::transaction::Error;

fn focused_control_error(error: FocusedControlError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-control operation failed: {error}"
    ))
}

fn focused_control_location(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<(
    FocusedNumbersPackage,
    litchi_numbers::SheetSelector<'static>,
    litchi_numbers::TableSelector<'static>,
    litchi_numbers::table::CellPosition,
)> {
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!("invalid Numbers cell-control coordinate: {error}"))
        })?;
    let source_bytes = editor.to_bytes()?;
    let source = FocusedNumbersPackage::from_bytes(&source_bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers cell-control source validation failed: {error}"
        ))
    })?;
    Ok((source, sheet, table, position))
}

fn focused_control_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<CellControl>> {
    let (source, sheet, table, position) = focused_control_location(editor, table_id, row, column)?;
    source
        .table_cell_control_format(sheet, table, position)
        .map_err(focused_control_error)
}

fn commit_focused_control_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<CellControl>,
) -> Result<NumbersEditor> {
    let source_bytes = editor.to_bytes()?;
    let (source, sheet, table, position) = focused_control_location(editor, table_id, row, column)?;
    let edit = source
        .edit_table_cell_control_format(sheet, table, position)
        .map_err(focused_control_error)?;
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    }
    .map_err(focused_control_error)?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat("could not allocate focused Numbers cell-control candidate".to_owned())
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_number_format_error(error: FocusedNumberFormatError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-number-format operation failed: {error}"
    ))
}

enum FocusedNumberFormatLocation {
    Owner {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
    },
    LegacyFallback,
}

fn focused_number_format_location(
    editor: &NumbersEditor,
    source_bytes: &[u8],
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedNumberFormatLocation> {
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid Numbers cell-number-format coordinate: {error}"
            ))
        })?;
    let source = match FocusedNumbersPackage::from_bytes(source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_))
            if !editor.package.source_is_exact() =>
        {
            return Ok(FocusedNumberFormatLocation::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers cell-number-format source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedNumberFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    })
}

const fn focused_number_format_read_can_fallback(
    error: FocusedNumberFormatError,
    source_built: bool,
) -> bool {
    source_built
        && matches!(
            error,
            FocusedNumberFormatError::CellNotFound
                | FocusedNumberFormatError::UnsupportedDependency { .. }
                | FocusedNumberFormatError::UnsupportedSource
                | FocusedNumberFormatError::InvalidSource { .. }
        )
}

const fn focused_number_format_edit_can_fallback(
    error: FocusedNumberFormatError,
    source_built: bool,
) -> bool {
    // Exact native packages fail closed after every focused-owner rejection,
    // including cross-family edits. Synthetic source-built packages may use
    // the legacy codec until their builder graph is admitted by the owner.
    source_built
        && (matches!(error, FocusedNumberFormatError::WrongFormatFamily { .. })
            || focused_number_format_read_can_fallback(error, true))
}

#[cfg(test)]
mod number_format_fallback_policy_tests {
    use super::{
        FocusedNumberFormatError, focused_number_format_edit_can_fallback,
        focused_number_format_read_can_fallback,
    };
    use litchi_numbers::cell::data_format::number::transaction::Path;

    #[test]
    fn exact_sources_never_fallback_after_structural_admission_failure() {
        let structural = FocusedNumberFormatError::UnsupportedSource;
        assert!(!focused_number_format_read_can_fallback(structural, false));
        assert!(!focused_number_format_edit_can_fallback(structural, false));
        assert!(focused_number_format_read_can_fallback(structural, true));
        assert!(focused_number_format_edit_can_fallback(structural, true));

        let family = FocusedNumberFormatError::WrongFormatFamily {
            path: Path::Package,
        };
        assert!(!focused_number_format_read_can_fallback(family, false));
        assert!(!focused_number_format_edit_can_fallback(family, false));
        assert!(focused_number_format_edit_can_fallback(family, true));
    }
}

#[cfg(test)]
mod percentage_format_fallback_policy_tests {
    use super::{
        FocusedPercentageFormatError, focused_percentage_format_edit_can_fallback,
        focused_percentage_format_read_can_fallback,
    };
    use litchi_numbers::cell::data_format::percentage::transaction::Path;

    #[test]
    fn exact_sources_never_fallback_after_structural_admission_failure() {
        let structural = FocusedPercentageFormatError::UnsupportedSource;
        assert!(!focused_percentage_format_read_can_fallback(
            structural, false
        ));
        assert!(!focused_percentage_format_edit_can_fallback(
            structural, false, true
        ));
        assert!(focused_percentage_format_read_can_fallback(
            structural, true
        ));
        assert!(focused_percentage_format_edit_can_fallback(
            structural, true, true
        ));

        let family = FocusedPercentageFormatError::WrongFormatFamily {
            path: Path::Package,
        };
        assert!(!focused_percentage_format_read_can_fallback(family, false));
        assert!(!focused_percentage_format_edit_can_fallback(
            family, false, true
        ));
        assert!(focused_percentage_format_edit_can_fallback(
            family, true, true
        ));
        assert!(!focused_percentage_format_edit_can_fallback(
            family, true, false
        ));
    }
}

fn focused_number_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Number>> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_number_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedNumberFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return cell_data_format::cell_number_format(&editor.package, table_id, row, column);
    };
    match source.table_cell_number_format(sheet, table, position) {
        Ok(format) => Ok(format),
        Err(error) if focused_number_format_read_can_fallback(error, source_built) => {
            cell_data_format::cell_number_format(&editor.package, table_id, row, column)
        },
        Err(error) => Err(focused_number_format_error(error)),
    }
}

fn commit_legacy_number_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Number>,
) -> Result<NumbersEditor> {
    let mut staged = editor.package.clone();
    match format {
        Some(format) => {
            cell_data_format::set_cell_number_format(&mut staged, table_id, row, column, format)?;
        },
        None => {
            cell_data_format::reset_cell_number_format(&mut staged, table_id, row, column)?;
        },
    }
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    let observed = cell_data_format::cell_number_format(&verified.package, table_id, row, column)?;
    if observed != format {
        return Err(Error::InvalidFormat(
            "Numbers table-cell number-format failed legacy package validation".to_owned(),
        ));
    }
    Ok(verified)
}

fn commit_focused_number_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Number>,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_number_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedNumberFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return commit_legacy_number_format(editor, table_id, row, column, format);
    };
    let edit = match source.edit_table_cell_number_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error) if focused_number_format_edit_can_fallback(error, source_built) => {
            return commit_legacy_number_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_number_format_error(error)),
    };
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    };
    let commit = match commit {
        Ok(commit) => commit,
        Err(error) if focused_number_format_edit_can_fallback(error, source_built) => {
            return commit_legacy_number_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_number_format_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers cell-number-format candidate".to_owned(),
        )
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_percentage_format_error(error: FocusedPercentageFormatError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-percentage-format operation failed: {error}"
    ))
}

enum FocusedPercentageFormatLocation {
    Owner {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
    },
    LegacyFallback,
}

fn focused_percentage_format_location(
    editor: &NumbersEditor,
    source_bytes: &[u8],
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedPercentageFormatLocation> {
    let (sheet, table): (
        litchi_numbers::SheetSelector<'static>,
        litchi_numbers::TableSelector<'static>,
    ) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid Numbers cell-percentage-format coordinate: {error}"
            ))
        })?;
    let source = match FocusedNumbersPackage::from_bytes(source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_))
            if !editor.package.source_is_exact() =>
        {
            return Ok(FocusedPercentageFormatLocation::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers cell-percentage-format source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedPercentageFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    })
}

const fn focused_percentage_format_read_can_fallback(
    error: FocusedPercentageFormatError,
    source_built: bool,
) -> bool {
    source_built
        && matches!(
            error,
            FocusedPercentageFormatError::CellNotFound
                | FocusedPercentageFormatError::UnsupportedDependency { .. }
                | FocusedPercentageFormatError::UnsupportedSource
                | FocusedPercentageFormatError::InvalidSource { .. }
        )
}

const fn focused_percentage_format_edit_can_fallback(
    error: FocusedPercentageFormatError,
    source_built: bool,
    allow_family_replacement: bool,
) -> bool {
    // Exact native packages fail closed after every focused-owner rejection.
    // Cross-family replacement remains a source-built compatibility behavior;
    // it cannot bypass the exact package's lock, budget, and locality owner.
    source_built
        && ((allow_family_replacement
            && matches!(
                error,
                FocusedPercentageFormatError::WrongFormatFamily { .. }
            ))
            || focused_percentage_format_read_can_fallback(error, true))
}

fn focused_percentage_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Percentage>> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location =
        focused_percentage_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedPercentageFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return cell_data_format::cell_percentage_format(&editor.package, table_id, row, column);
    };
    match source.table_cell_percentage_format(sheet, table, position) {
        Ok(format) => Ok(format),
        Err(error) if focused_percentage_format_read_can_fallback(error, source_built) => {
            cell_data_format::cell_percentage_format(&editor.package, table_id, row, column)
        },
        Err(error) => Err(focused_percentage_format_error(error)),
    }
}

fn commit_legacy_percentage_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Percentage>,
) -> Result<NumbersEditor> {
    let mut staged = editor.package.clone();
    match format {
        Some(format) => {
            let data_format = DataFormat::Percentage(format);
            cell_data_format::set_cell_data_format(
                &mut staged,
                table_id,
                row,
                column,
                &data_format,
            )?;
        },
        None => {
            cell_data_format::reset_cell_percentage_format(&mut staged, table_id, row, column)?;
        },
    }
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    let observed =
        cell_data_format::cell_percentage_format(&verified.package, table_id, row, column)?;
    if observed != format {
        return Err(Error::InvalidFormat(
            "Numbers table-cell percentage-format failed legacy package validation".to_owned(),
        ));
    }
    Ok(verified)
}

fn commit_focused_percentage_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Percentage>,
    allow_family_replacement: bool,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location =
        focused_percentage_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedPercentageFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return commit_legacy_percentage_format(editor, table_id, row, column, format);
    };
    let edit = match source.edit_table_cell_percentage_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error)
            if focused_percentage_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_percentage_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_percentage_format_error(error)),
    };
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    };
    let commit = match commit {
        Ok(commit) => commit,
        Err(error)
            if focused_percentage_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_percentage_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_percentage_format_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers cell-percentage-format candidate".to_owned(),
        )
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_currency_format_error(error: FocusedCurrencyFormatError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-currency-format operation failed: {error}"
    ))
}

enum FocusedCurrencyFormatLocation {
    Owner {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
    },
    LegacyFallback,
}

fn focused_currency_format_location(
    editor: &NumbersEditor,
    source_bytes: &[u8],
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedCurrencyFormatLocation> {
    let (sheet, table): (
        litchi_numbers::SheetSelector<'static>,
        litchi_numbers::TableSelector<'static>,
    ) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid Numbers cell-currency-format coordinate: {error}"
            ))
        })?;
    let source = match FocusedNumbersPackage::from_bytes(source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_))
            if !editor.package.source_is_exact() =>
        {
            return Ok(FocusedCurrencyFormatLocation::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers cell-currency-format source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedCurrencyFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    })
}

const fn focused_currency_format_read_can_fallback(
    error: FocusedCurrencyFormatError,
    source_built: bool,
) -> bool {
    source_built
        && matches!(
            error,
            FocusedCurrencyFormatError::CellNotFound
                | FocusedCurrencyFormatError::UnsupportedDependency { .. }
                | FocusedCurrencyFormatError::UnsupportedSource
                | FocusedCurrencyFormatError::InvalidSource { .. }
        )
}

const fn focused_currency_format_edit_can_fallback(
    error: FocusedCurrencyFormatError,
    source_built: bool,
    allow_family_replacement: bool,
) -> bool {
    // Exact native packages fail closed after every focused-owner rejection.
    // Cross-family replacement remains a source-built compatibility behavior;
    // it cannot bypass the exact package's lock, budget, and locality owner.
    source_built
        && ((allow_family_replacement
            && matches!(error, FocusedCurrencyFormatError::WrongFormatFamily { .. }))
            || focused_currency_format_read_can_fallback(error, true))
}

#[cfg(test)]
mod currency_format_fallback_policy_tests {
    use super::{
        FocusedCurrencyFormatError, focused_currency_format_edit_can_fallback,
        focused_currency_format_read_can_fallback,
    };
    use litchi_numbers::cell::data_format::currency::transaction::Path;

    #[test]
    fn exact_sources_never_fallback_after_structural_admission_failure() {
        let structural = FocusedCurrencyFormatError::UnsupportedSource;
        assert!(!focused_currency_format_read_can_fallback(
            structural, false
        ));
        assert!(!focused_currency_format_edit_can_fallback(
            structural, false, true
        ));
        assert!(focused_currency_format_read_can_fallback(structural, true));
        assert!(focused_currency_format_edit_can_fallback(
            structural, true, true
        ));

        let family = FocusedCurrencyFormatError::WrongFormatFamily {
            path: Path::Package,
        };
        assert!(!focused_currency_format_read_can_fallback(family, false));
        assert!(!focused_currency_format_edit_can_fallback(
            family, false, true
        ));
        assert!(focused_currency_format_edit_can_fallback(
            family, true, true
        ));
        assert!(!focused_currency_format_edit_can_fallback(
            family, true, false
        ));
    }
}

fn focused_currency_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Currency>> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_currency_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedCurrencyFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return cell_data_format::cell_currency_format(&editor.package, table_id, row, column);
    };
    match source.table_cell_currency_format(sheet, table, position) {
        Ok(format) => Ok(format),
        Err(error) if focused_currency_format_read_can_fallback(error, source_built) => {
            cell_data_format::cell_currency_format(&editor.package, table_id, row, column)
        },
        Err(error) => Err(focused_currency_format_error(error)),
    }
}

fn commit_legacy_currency_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Currency>,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let mut staged = editor.package.clone();
    match format {
        Some(format) => {
            let data_format = DataFormat::Currency(format);
            cell_data_format::set_cell_data_format(
                &mut staged,
                table_id,
                row,
                column,
                &data_format,
            )?;
        },
        None => {
            cell_data_format::reset_cell_currency_format(&mut staged, table_id, row, column)?;
        },
    }
    // A generated package must stay source-built across this compatibility
    // mutation. Reopening its normalized bytes would manufacture exact-source
    // provenance and disable the same bounded fallback on the next operation.
    let verified = if source_built {
        staged.validate()?;
        NumbersEditor::from_package(staged)?
    } else {
        NumbersEditor::from_bytes(&staged.to_bytes()?)?
    };
    let observed =
        cell_data_format::cell_currency_format(&verified.package, table_id, row, column)?;
    if observed != format {
        return Err(Error::InvalidFormat(
            "Numbers table-cell currency-format failed legacy package validation".to_owned(),
        ));
    }
    Ok(verified)
}

fn commit_focused_currency_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Currency>,
    allow_family_replacement: bool,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_currency_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedCurrencyFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return commit_legacy_currency_format(editor, table_id, row, column, format);
    };
    let edit = match source.edit_table_cell_currency_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error)
            if focused_currency_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_currency_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_currency_format_error(error)),
    };
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    };
    let commit = match commit {
        Ok(commit) => commit,
        Err(error)
            if focused_currency_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_currency_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_currency_format_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers cell-currency-format candidate".to_owned(),
        )
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_scientific_format_error(error: FocusedScientificFormatError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-scientific-format operation failed: {error}"
    ))
}

enum FocusedScientificFormatLocation {
    Owner {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
    },
    LegacyFallback,
}

fn focused_scientific_format_location(
    editor: &NumbersEditor,
    source_bytes: &[u8],
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedScientificFormatLocation> {
    let (sheet, table): (
        litchi_numbers::SheetSelector<'static>,
        litchi_numbers::TableSelector<'static>,
    ) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid Numbers cell-scientific-format coordinate: {error}"
            ))
        })?;
    let source = match FocusedNumbersPackage::from_bytes(source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_))
            if !editor.package.source_is_exact() =>
        {
            return Ok(FocusedScientificFormatLocation::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers cell-scientific-format source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedScientificFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    })
}

const fn focused_scientific_format_read_can_fallback(
    error: FocusedScientificFormatError,
    source_built: bool,
) -> bool {
    source_built
        && matches!(
            error,
            FocusedScientificFormatError::CellNotFound
                | FocusedScientificFormatError::UnsupportedDependency { .. }
                | FocusedScientificFormatError::UnsupportedSource
                | FocusedScientificFormatError::InvalidSource { .. }
        )
}

const fn focused_scientific_format_edit_can_fallback(
    error: FocusedScientificFormatError,
    source_built: bool,
    allow_family_replacement: bool,
) -> bool {
    // Exact native packages fail closed after every focused-owner rejection.
    // Cross-family replacement remains a source-built compatibility behavior;
    // it cannot bypass the exact package's lock, budget, and locality owner.
    source_built
        && ((allow_family_replacement
            && matches!(
                error,
                FocusedScientificFormatError::WrongFormatFamily { .. }
            ))
            || focused_scientific_format_read_can_fallback(error, true))
}

#[cfg(test)]
mod scientific_format_fallback_policy_tests {
    use super::{
        FocusedScientificFormatError, focused_scientific_format_edit_can_fallback,
        focused_scientific_format_read_can_fallback,
    };
    use litchi_numbers::cell::data_format::scientific::transaction::Path;

    #[test]
    fn exact_sources_never_fallback_after_structural_admission_failure() {
        let structural = FocusedScientificFormatError::UnsupportedSource;
        assert!(!focused_scientific_format_read_can_fallback(
            structural, false
        ));
        assert!(!focused_scientific_format_edit_can_fallback(
            structural, false, true
        ));
        assert!(focused_scientific_format_read_can_fallback(
            structural, true
        ));
        assert!(focused_scientific_format_edit_can_fallback(
            structural, true, true
        ));

        let family = FocusedScientificFormatError::WrongFormatFamily {
            path: Path::Package,
        };
        assert!(!focused_scientific_format_read_can_fallback(family, false));
        assert!(!focused_scientific_format_edit_can_fallback(
            family, false, true
        ));
        assert!(focused_scientific_format_edit_can_fallback(
            family, true, true
        ));
        assert!(!focused_scientific_format_edit_can_fallback(
            family, true, false
        ));
    }
}

fn focused_scientific_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Scientific>> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location =
        focused_scientific_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedScientificFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return cell_data_format::cell_scientific_format(&editor.package, table_id, row, column);
    };
    match source.table_cell_scientific_format(sheet, table, position) {
        Ok(format) => Ok(format),
        Err(error) if focused_scientific_format_read_can_fallback(error, source_built) => {
            cell_data_format::cell_scientific_format(&editor.package, table_id, row, column)
        },
        Err(error) => Err(focused_scientific_format_error(error)),
    }
}

fn commit_legacy_scientific_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Scientific>,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let mut staged = editor.package.clone();
    match format {
        Some(format) => {
            let data_format = DataFormat::Scientific(format);
            cell_data_format::set_cell_data_format(
                &mut staged,
                table_id,
                row,
                column,
                &data_format,
            )?;
        },
        None => {
            cell_data_format::reset_cell_scientific_format(&mut staged, table_id, row, column)?;
        },
    }
    // A generated package must stay source-built across this compatibility
    // mutation. Reopening its normalized bytes would manufacture exact-source
    // provenance and disable the same bounded fallback on the next operation.
    let verified = if source_built {
        staged.validate()?;
        NumbersEditor::from_package(staged)?
    } else {
        NumbersEditor::from_bytes(&staged.to_bytes()?)?
    };
    let observed =
        cell_data_format::cell_scientific_format(&verified.package, table_id, row, column)?;
    if observed != format {
        return Err(Error::InvalidFormat(
            "Numbers table-cell scientific-format failed legacy package validation".to_owned(),
        ));
    }
    Ok(verified)
}

fn commit_focused_scientific_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Scientific>,
    allow_family_replacement: bool,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location =
        focused_scientific_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedScientificFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return commit_legacy_scientific_format(editor, table_id, row, column, format);
    };
    let edit = match source.edit_table_cell_scientific_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error)
            if focused_scientific_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_scientific_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_scientific_format_error(error)),
    };
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    };
    let commit = match commit {
        Ok(commit) => commit,
        Err(error)
            if focused_scientific_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_scientific_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_scientific_format_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers cell-scientific-format candidate".to_owned(),
        )
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
}

fn focused_fraction_format_error(error: FocusedFractionFormatError) -> Error {
    Error::InvalidFormat(format!(
        "focused Numbers cell-fraction-format operation failed: {error}"
    ))
}

enum FocusedFractionFormatLocation {
    Owner {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
    },
    LegacyFallback,
}

fn focused_fraction_format_location(
    editor: &NumbersEditor,
    source_bytes: &[u8],
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<FocusedFractionFormatLocation> {
    let (sheet, table): (
        litchi_numbers::SheetSelector<'static>,
        litchi_numbers::TableSelector<'static>,
    ) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid Numbers cell-fraction-format coordinate: {error}"
            ))
        })?;
    let source = match FocusedNumbersPackage::from_bytes(source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_))
            if !editor.package.source_is_exact() =>
        {
            return Ok(FocusedFractionFormatLocation::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers cell-fraction-format source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedFractionFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    })
}

const fn focused_fraction_format_read_can_fallback(
    error: FocusedFractionFormatError,
    source_built: bool,
) -> bool {
    source_built
        && matches!(
            error,
            FocusedFractionFormatError::CellNotFound
                | FocusedFractionFormatError::UnsupportedDependency { .. }
                | FocusedFractionFormatError::UnsupportedSource
                | FocusedFractionFormatError::InvalidSource { .. }
        )
}

const fn focused_fraction_format_edit_can_fallback(
    error: FocusedFractionFormatError,
    source_built: bool,
    allow_family_replacement: bool,
) -> bool {
    // Exact native packages fail closed after every focused-owner rejection.
    // Cross-family replacement remains a source-built compatibility behavior;
    // it cannot bypass the exact package's lock, budget, and locality owner.
    source_built
        && ((allow_family_replacement
            && matches!(error, FocusedFractionFormatError::WrongFormatFamily { .. }))
            || focused_fraction_format_read_can_fallback(error, true))
}

#[cfg(test)]
mod fraction_format_fallback_policy_tests {
    use super::{
        FocusedFractionFormatError, focused_fraction_format_edit_can_fallback,
        focused_fraction_format_read_can_fallback,
    };
    use litchi_numbers::cell::data_format::fraction::transaction::Path;

    #[test]
    fn exact_sources_never_fallback_after_structural_admission_failure() {
        let structural = FocusedFractionFormatError::UnsupportedSource;
        assert!(!focused_fraction_format_read_can_fallback(
            structural, false
        ));
        assert!(!focused_fraction_format_edit_can_fallback(
            structural, false, true
        ));
        assert!(focused_fraction_format_read_can_fallback(structural, true));
        assert!(focused_fraction_format_edit_can_fallback(
            structural, true, true
        ));

        let family = FocusedFractionFormatError::WrongFormatFamily {
            path: Path::Package,
        };
        assert!(!focused_fraction_format_read_can_fallback(family, false));
        assert!(!focused_fraction_format_edit_can_fallback(
            family, false, true
        ));
        assert!(focused_fraction_format_edit_can_fallback(
            family, true, true
        ));
        assert!(!focused_fraction_format_edit_can_fallback(
            family, true, false
        ));
    }
}

fn focused_fraction_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Fraction>> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_fraction_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedFractionFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return cell_data_format::cell_fraction_format(&editor.package, table_id, row, column);
    };
    match source.table_cell_fraction_format(sheet, table, position) {
        Ok(format) => Ok(format),
        Err(error) if focused_fraction_format_read_can_fallback(error, source_built) => {
            cell_data_format::cell_fraction_format(&editor.package, table_id, row, column)
        },
        Err(error) => Err(focused_fraction_format_error(error)),
    }
}

fn commit_legacy_fraction_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Fraction>,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let mut staged = editor.package.clone();
    match format {
        Some(format) => {
            let data_format = DataFormat::Fraction(format);
            cell_data_format::set_cell_data_format(
                &mut staged,
                table_id,
                row,
                column,
                &data_format,
            )?;
        },
        None => {
            cell_data_format::reset_cell_fraction_format(&mut staged, table_id, row, column)?;
        },
    }
    // Keep generated packages source-built through the compatibility
    // mutation. Reopening their normalized bytes would incorrectly turn the
    // builder output into an exact-source package and disable this bounded
    // fallback on the next operation.
    let verified = if source_built {
        staged.validate()?;
        NumbersEditor::from_package(staged)?
    } else {
        NumbersEditor::from_bytes(&staged.to_bytes()?)?
    };
    let observed =
        cell_data_format::cell_fraction_format(&verified.package, table_id, row, column)?;
    if observed != format {
        return Err(Error::InvalidFormat(
            "Numbers table-cell fraction-format failed legacy package validation".to_owned(),
        ));
    }
    Ok(verified)
}

fn commit_focused_fraction_format(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
    format: Option<Fraction>,
    allow_family_replacement: bool,
) -> Result<NumbersEditor> {
    let source_built = !editor.package.source_is_exact();
    let source_bytes = editor.to_bytes()?;
    let location = focused_fraction_format_location(editor, &source_bytes, table_id, row, column)?;
    let FocusedFractionFormatLocation::Owner {
        source,
        sheet,
        table,
        position,
    } = location
    else {
        return commit_legacy_fraction_format(editor, table_id, row, column, format);
    };
    let edit = match source.edit_table_cell_fraction_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error)
            if focused_fraction_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_fraction_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_fraction_format_error(error)),
    };
    let commit = match format {
        Some(format) => edit.set(format).commit(),
        None => edit.clear().commit(),
    };
    let commit = match commit {
        Ok(commit) => commit,
        Err(error)
            if focused_fraction_format_edit_can_fallback(
                error,
                source_built,
                allow_family_replacement,
            ) =>
        {
            return commit_legacy_fraction_format(editor, table_id, row, column, format);
        },
        Err(error) => return Err(focused_fraction_format_error(error)),
    };
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(source_bytes.len()).map_err(|_| {
        Error::InvalidFormat(
            "could not allocate focused Numbers cell-fraction-format candidate".to_owned(),
        )
    })?;
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|error| Error::Io(error.into_io_error()))?;
    NumbersEditor::from_bytes(&bytes)
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
    let (sheet, table) = selectors::focused_table_location(editor, table_id)?;
    let position =
        litchi_numbers::table::CellPosition::try_from_usize(row, column).map_err(|error| {
            Error::InvalidFormat(format!("invalid Numbers comment coordinate: {error}"))
        })?;
    let source_bytes = editor.to_bytes()?;
    let source = match FocusedNumbersPackage::from_bytes(&source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_)) => {
            return Ok(FocusedCommentReplacement::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers comment source validation failed: {error}"
            )));
        },
    };
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

enum FocusedCommentReplySource {
    Ready {
        source: FocusedNumbersPackage,
        sheet: litchi_numbers::SheetSelector<'static>,
        table: litchi_numbers::TableSelector<'static>,
        position: litchi_numbers::table::CellPosition,
        source_bytes: Vec<u8>,
    },
    LegacyFallback,
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
    let source = match FocusedNumbersPackage::from_bytes(&source_bytes) {
        Ok(source) => source,
        Err(litchi_numbers::PackageError::InvalidFormat(_)) => {
            return Ok(FocusedCommentReplySource::LegacyFallback);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Numbers comment-reply source validation failed: {error}"
            )));
        },
    };
    Ok(FocusedCommentReplySource::Ready {
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
    let before = match cell_comment_replies_in_package(editor.package(), table_id, row, column) {
        Ok(replies) => replies,
        Err(_) => return Ok(FocusedCommentReplyPublication::LegacyFallback),
    };
    let FocusedCommentReplySource::Ready {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
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
    let before = match cell_comment_replies_in_package(editor.package(), table_id, row, column) {
        Ok(replies) => replies,
        Err(_) => return Ok(FocusedCommentReplyPublication::LegacyFallback),
    };
    let Some(ordinal) = before
        .iter()
        .position(|reply| reply.storage_id.get() == reply_storage_object_id)
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
    let index = CommentReplyIndex::try_from_usize(ordinal)
        .map_err(|_| Error::InvalidFormat("Numbers comment-reply ordinal overflow".to_owned()))?;
    let FocusedCommentReplySource::Ready {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
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
    let before = match cell_comment_replies_in_package(editor.package(), table_id, row, column) {
        Ok(replies) => replies,
        Err(_) => return Ok(FocusedCommentReplyPublication::LegacyFallback),
    };
    let Some(ordinal) = before
        .iter()
        .position(|reply| reply.storage_id.get() == reply_storage_object_id)
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
    let index = CommentReplyIndex::try_from_usize(ordinal)
        .map_err(|_| Error::InvalidFormat("Numbers comment-reply ordinal overflow".to_owned()))?;
    let FocusedCommentReplySource::Ready {
        source,
        sheet,
        table,
        position,
        source_bytes,
    } = focused_comment_reply_source(editor, table_id, row, column)?
    else {
        return Ok(FocusedCommentReplyPublication::LegacyFallback);
    };
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
        if let Ok(control) = CellControl::try_from(format.clone()) {
            *self = commit_focused_control_format(self, table_id, row, column, Some(control))?;
            return Ok(());
        }
        if CellControl::try_from(current).is_ok() {
            let mut staged = commit_focused_control_format(self, table_id, row, column, None)?;
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

    /// Read an explicit decimal-number format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Number-format API; use litchi_numbers::Package::table_cell_number_format with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read"
    )]
    pub fn table_cell_number_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Number>> {
        focused_number_format(self, table_id, row, column)
    }

    /// Create or replace an explicit decimal-number format transactionally.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Number-format API; use litchi_numbers::Package::edit_table_cell_number_format with SheetSelector, TableSelector, and CellPosition for selector-first writes"
    )]
    pub fn set_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Number,
    ) -> Result<()> {
        *self = commit_focused_number_format(self, table_id, row, column, Some(format))?;
        Ok(())
    }

    /// Restore iWork's automatic data format for one table cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Number-format API; use litchi_numbers::Package::edit_table_cell_number_format with SheetSelector, TableSelector, and CellPosition to clear the explicit format"
    )]
    pub fn reset_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        if focused_number_format(self, table_id, row, column)?.is_none() {
            return Ok(false);
        }
        let verified = commit_focused_number_format(self, table_id, row, column, None)?;
        if focused_number_format(&verified, table_id, row, column)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table-cell number-format reset failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read an explicit Text format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_text_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Text>> {
        cell_data_format::cell_text_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Text format transactionally.
    pub fn set_table_cell_text_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, Text.into())
    }

    /// Restore Automatic from an explicit Text cell.
    pub fn reset_table_cell_text_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_data_format::reset_cell_text_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != DataFormat::Automatic {
                return Err(Error::InvalidFormat(
                    "Numbers Text-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
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
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_custom_format(&mut staged, table_id, row, column)?;
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

    /// Read an explicit currency format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Currency-format API; use litchi_numbers::Package::table_cell_currency_format with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read"
    )]
    pub fn table_cell_currency_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Currency>> {
        focused_currency_format(self, table_id, row, column)
    }

    /// Create or replace an explicit currency format transactionally.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Currency-format API; use litchi_numbers::Package::edit_table_cell_currency_format with SheetSelector, TableSelector, and CellPosition for selector-first writes"
    )]
    pub fn set_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Currency,
    ) -> Result<()> {
        *self = commit_focused_currency_format(self, table_id, row, column, Some(format), true)?;
        Ok(())
    }

    /// Restore Automatic from an explicit Currency cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Currency-format API; use litchi_numbers::Package::edit_table_cell_currency_format with SheetSelector, TableSelector, and CellPosition to clear the explicit Currency format"
    )]
    pub fn reset_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        if focused_currency_format(self, table_id, row, column)?.is_none() {
            return Ok(false);
        }
        let verified = commit_focused_currency_format(self, table_id, row, column, None, false)?;
        if focused_currency_format(&verified, table_id, row, column)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table-cell currency-format reset failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read an explicit percentage format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Percentage-format API; use litchi_numbers::Package::table_cell_percentage_format with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read"
    )]
    pub fn table_cell_percentage_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Percentage>> {
        focused_percentage_format(self, table_id, row, column)
    }

    /// Create or replace an explicit percentage format transactionally.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Percentage-format API; use litchi_numbers::Package::edit_table_cell_percentage_format with SheetSelector, TableSelector, and CellPosition for selector-first writes"
    )]
    pub fn set_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Percentage,
    ) -> Result<()> {
        *self = commit_focused_percentage_format(self, table_id, row, column, Some(format), true)?;
        Ok(())
    }

    /// Restore iWork's automatic format from an explicit Percentage cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Percentage-format API; use litchi_numbers::Package::edit_table_cell_percentage_format with SheetSelector, TableSelector, and CellPosition to clear the explicit format"
    )]
    pub fn reset_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        if focused_percentage_format(self, table_id, row, column)?.is_none() {
            return Ok(false);
        }
        let verified = commit_focused_percentage_format(self, table_id, row, column, None, false)?;
        if focused_percentage_format(&verified, table_id, row, column)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table-cell percentage-format reset failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read an explicit scientific-notation format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Scientific-format API; use litchi_numbers::Package::table_cell_scientific_format with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read"
    )]
    pub fn table_cell_scientific_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Scientific>> {
        focused_scientific_format(self, table_id, row, column)
    }

    /// Create or replace an explicit scientific-notation format transactionally.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Scientific-format API; use litchi_numbers::Package::edit_table_cell_scientific_format with SheetSelector, TableSelector, and CellPosition for selector-first writes"
    )]
    pub fn set_table_cell_scientific_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Scientific,
    ) -> Result<()> {
        *self = commit_focused_scientific_format(self, table_id, row, column, Some(format), true)?;
        Ok(())
    }

    /// Restore Automatic from an explicit Scientific cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Scientific-format API; use litchi_numbers::Package::edit_table_cell_scientific_format with SheetSelector, TableSelector, and CellPosition to clear the explicit Scientific format"
    )]
    pub fn reset_table_cell_scientific_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        if focused_scientific_format(self, table_id, row, column)?.is_none() {
            return Ok(false);
        }
        let verified = commit_focused_scientific_format(self, table_id, row, column, None, false)?;
        if focused_scientific_format(&verified, table_id, row, column)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table-cell scientific-format reset failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read an explicit mixed-fraction format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Fraction-format API; use litchi_numbers::Package::table_cell_fraction_format with SheetSelector, TableSelector, and CellPosition for the selector-first semantic read"
    )]
    pub fn table_cell_fraction_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Fraction>> {
        focused_fraction_format(self, table_id, row, column)
    }

    /// Create or replace an explicit mixed-fraction format transactionally.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Fraction-format API; use litchi_numbers::Package::edit_table_cell_fraction_format with SheetSelector, TableSelector, and CellPosition for selector-first writes"
    )]
    pub fn set_table_cell_fraction_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Fraction,
    ) -> Result<()> {
        *self = commit_focused_fraction_format(self, table_id, row, column, Some(format), true)?;
        Ok(())
    }

    /// Restore Automatic from an explicit Fraction cell.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Numbers cell Fraction-format API; use litchi_numbers::Package::edit_table_cell_fraction_format with SheetSelector, TableSelector, and CellPosition to clear the explicit Fraction format"
    )]
    pub fn reset_table_cell_fraction_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        if focused_fraction_format(self, table_id, row, column)?.is_none() {
            return Ok(false);
        }
        let verified = commit_focused_fraction_format(self, table_id, row, column, None, false)?;
        if focused_fraction_format(&verified, table_id, row, column)?.is_some() {
            return Err(Error::InvalidFormat(
                "Numbers table-cell fraction-format reset failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
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

    /// Read an explicit Date & Time format for one table cell.
    ///
    /// `None` means the Date value uses iWork's automatic data format.
    pub fn table_cell_date_time_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<DateTime>> {
        cell_data_format::cell_date_time_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Date & Time format transactionally.
    pub fn set_table_cell_date_time_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: DateTime,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Date & Time cell.
    pub fn reset_table_cell_date_time_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_date_time_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != DataFormat::Automatic {
                return Err(Error::InvalidFormat(
                    "Numbers Date & Time reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Duration format for one table cell.
    ///
    /// `None` means the Duration value uses iWork's automatic data format.
    pub fn table_cell_duration_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Duration>> {
        cell_data_format::cell_duration_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Duration format transactionally.
    pub fn set_table_cell_duration_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Duration,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Duration cell.
    pub fn reset_table_cell_duration_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_duration_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)? != DataFormat::Automatic {
                return Err(Error::InvalidFormat(
                    "Numbers Duration reset failed package validation".to_owned(),
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
