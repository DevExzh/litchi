//! Immutable XLSB defined-name writes (record 0261).
//!
//! `add("/", namedrange, {name, ref, scope?})` sets a defined name that
//! refers to a worksheet range (`Sheet1!$A$1:$B$5`), `set("/namedrange[N]",
//! {ref})` replaces its reference, and `remove("/namedrange[N]")` removes
//! it. Each stages one native structural edit
//! ([`litchi_xlsb::cell_values::structure::Edit::set_defined_name`]), which
//! splices only the workbook stream's `BrtName` and extern-sheet records.
//! Names are appended and only the last name can be removed, so no name
//! reference in a formula is renumbered; other removals are typed refusals.

use litchi_core::selector::{CellRange, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsb::formula::Range;
use litchi_xlsb::sheet::Selector as NativeSelector;

use super::Nav;
use super::Xlsb;
use super::node::Item;
use super::spec::{Kind, NAMED_RANGE_WRITE_TABLE, NamedRangeSetter};
use super::write::{Commit, NativeStructureOptions, native_error, package_from_editor};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::props::Props;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::Scalar;

#[derive(Debug, Default)]
struct NameAssignments {
    name: Option<String>,
    reference: Option<String>,
    scope: Option<String>,
}

fn admit(operation: Operation, props: &Props<'_>) -> Result<NameAssignments> {
    let mut assignments = NameAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = NAMED_RANGE_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let entry = *entry;
        entry
            .prepare(operation, value, |setter, value| {
                let Some(Scalar::Text(text)) = value else {
                    return Err(Error::new(Code::InvalidValue, "the property must be text")
                        .with_property(entry.spec.name));
                };
                let slot = match setter {
                    NamedRangeSetter::Name => &mut assignments.name,
                    NamedRangeSetter::Ref => &mut assignments.reference,
                    NamedRangeSetter::Scope => &mut assignments.scope,
                };
                if slot.is_some() {
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "the same XLSB defined-name property was supplied more than once",
                    )
                    .with_property(entry.spec.name));
                }
                *slot = Some(text.to_owned());
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    Ok(assignments)
}

/// Resolve `Sheet!$A$1:$B$2` to a worksheet's catalog position and range.
fn parse_reference(file: &Xlsb, text: &str) -> Result<(usize, Range)> {
    let invalid = || {
        Error::new(
            Code::InvalidValue,
            "a defined-name reference is a sheet-qualified range such as Sheet1!$A$1:$B$2",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("ref")
    };
    let text = text.strip_prefix('=').unwrap_or(text).trim();
    let (sheet, area) = text.rsplit_once('!').ok_or_else(invalid)?;
    let sheet = match sheet
        .strip_prefix('\'')
        .and_then(|rest| rest.strip_suffix('\''))
    {
        Some(quoted) => quoted.replace("''", "'"),
        None => sheet.to_owned(),
    };
    let position = sheet_position(file, &sheet, "ref")?;
    let area: String = area.chars().filter(|character| *character != '$').collect();
    let range = CellRange::parse(&area).ok_or_else(invalid)?;
    let (first, last) = (range.start(), range.end());
    let range = Range::new(
        first.row().min(last.row()),
        first.row().max(last.row()),
        first.column().min(last.column()),
        first.column().max(last.column()),
    )
    .map_err(|_| invalid())?;
    Ok((position, range))
}

fn sheet_position(file: &Xlsb, name: &str, property: &'static str) -> Result<usize> {
    let catalog = file.catalog()?;
    let tab = catalog.select(name)?.ok_or_else(|| {
        Error::new(Code::InvalidValue, "the defined name names a missing sheet")
            .with_format(FileFormat::Xlsb)
            .with_property(property)
    })?;
    if !tab.kind().is_worksheet() {
        return Err(
            Error::new(Code::InvalidValue, "a defined name refers to a worksheet")
                .with_format(FileFormat::Xlsb)
                .with_property(property),
        );
    }
    Ok(tab.position().get())
}

fn edit(file: &Xlsb, limits: &SemanticLimits) -> Result<litchi_xlsb::cell_values::structure::Edit> {
    let package = package_from_editor(file.editor_for_write()?);
    let native_limits = litchi_xlsb::cell_values::TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)
}

fn name_error(error: litchi_xlsb::package::PackageError) -> Error {
    match error {
        litchi_xlsb::package::PackageError::InvalidFormula(_) => Error::new(
            Code::InvalidValue,
            "the XLSB defined name is not a valid name",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("name"),
        litchi_xlsb::package::PackageError::UnsupportedFeature(_) => Error::new(
            Code::UnsupportedCapability,
            "only the workbook's last defined name can be removed",
        )
        .with_format(FileFormat::Xlsb)
        .with_suggestion("removing an earlier name would renumber every formula's name references"),
        error => native_error(error),
    }
}

fn commit(
    file: &Xlsb,
    edit: litchi_xlsb::cell_values::structure::Edit,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let native = edit.commit().map_err(name_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok(commit)
}

fn named_range_path(index: usize) -> Path<'static> {
    Path::from_steps([Step::new("namedrange").with(Predicate::Index(Position::new(index)))])
}

/// `add` of a `namedrange` under the workbook.
pub(super) fn add_named_range(
    file: &Xlsb,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let assignments = admit(Operation::Add, node.props())?;
    let missing = |property: &'static str| {
        Error::new(
            Code::InvalidArgument,
            "a defined name needs `name` and `ref`",
        )
        .with_format(FileFormat::Xlsb)
        .with_property(property)
    };
    let name = assignments.name.ok_or_else(|| missing("name"))?;
    let reference = assignments.reference.ok_or_else(|| missing("ref"))?;
    let (sheet, range) = parse_reference(file, &reference)?;
    let scope = assignments
        .scope
        .filter(|scope| !scope.eq_ignore_ascii_case("workbook"))
        .map(|scope| sheet_position(file, &scope, "scope"))
        .transpose()?;
    let entries = file.defined_name_entries()?;
    let index = entries
        .iter()
        .position(|entry| entry.scope == scope && entry.name.to_uppercase() == name.to_uppercase())
        .unwrap_or(entries.len());
    let mut edit = edit(file, limits)?;
    edit.set_defined_name(
        &name,
        scope.map(NativeSelector::from),
        NativeSelector::from(sheet),
        range,
    )
    .map_err(name_error)?
    .ok_or_else(internal_missing_tab)?;
    Ok((commit(file, edit, limits)?, named_range_path(index)))
}

fn internal_missing_tab() -> Error {
    Error::new(Code::PathNotFound, "the selected XLSB tab is absent").with_format(FileFormat::Xlsb)
}

/// The defined name a target path selects, if it selects one.
fn selected_name(
    file: &Xlsb,
    target: &TargetArg<'_>,
) -> Result<Option<litchi_xlsb::named_ranges::Entry>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    if path
        .path()
        .steps()
        .last()
        .is_none_or(|step| step.name() != "namedrange")
    {
        return Ok(None);
    }
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.element() != Kind::NamedRange {
        return Ok(None);
    }
    let Item::NamedRange(index) = node.item() else {
        return Ok(None);
    };
    Ok(file.defined_name_entries()?.into_iter().nth(*index))
}

/// `set` of a defined name's `ref`; `None` for targets that are not names.
pub(super) fn try_set_named_range(
    file: &Xlsb,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let Some(entry) = selected_name(file, target)? else {
        return Ok(None);
    };
    let assignments = admit(Operation::Set, props)?;
    let Some(reference) = assignments.reference else {
        return Err(
            Error::new(Code::InvalidArgument, "a defined-name set needs `ref`")
                .with_format(FileFormat::Xlsb)
                .with_property("ref"),
        );
    };
    let (sheet, range) = parse_reference(file, &reference)?;
    let mut edit = edit(file, limits)?;
    edit.set_defined_name(
        &entry.name,
        entry.scope.map(NativeSelector::from),
        NativeSelector::from(sheet),
        range,
    )
    .map_err(name_error)?
    .ok_or_else(internal_missing_tab)?;
    commit(file, edit, limits).map(Some)
}

/// `remove` of a defined name; `None` for targets that are not names.
pub(super) fn try_remove_named_range(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let Some(entry) = selected_name(file, target)? else {
        return Ok(None);
    };
    let mut edit = edit(file, limits)?;
    edit.remove_defined_name(&entry.name, entry.scope.map(NativeSelector::from))
        .map_err(name_error)?
        .ok_or_else(internal_missing_tab)?;
    commit(file, edit, limits).map(Some)
}
