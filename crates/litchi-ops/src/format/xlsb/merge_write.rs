//! Immutable XLSB merged-range writes (record 0261).
//!
//! `set(range, {merge: true})` merges exactly that range and
//! `{merge: false}` (or null) removes exactly that merge. Every selected
//! range stages in one native structural edit
//! ([`litchi_xlsb::cell_values::structure::Edit::merge_cells`]), which
//! rewrites only the worksheet's `BrtMergeCell` block. Cell records are not
//! changed. A range overlapping a different merge is refused, so a caller
//! cannot merge across, or remove, a merge it did not name.

use litchi_core::FileFormat;
use litchi_xlsb::merged_cells::MergedCell;

use super::Nav;
use super::Xlsb;
use super::node::{Item, Node};
use super::spec::{Kind, MERGE_WRITE_TABLE};
use super::write::{Commit, NativeStructureOptions, native_error, package_from_editor};
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::Scalar;

/// One selected range: its complete-catalog sheet position and cells.
#[derive(Debug, Clone, PartialEq, Eq)]
struct MergeTarget {
    sheet: usize,
    range: MergedCell,
}

/// The merge route: `Ok(None)` when no `merge` property is supplied, so
/// another write family admits the request.
pub(super) fn set_merges_with_limits(
    file: &Xlsb,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let mut merge = None;
    for (key, value) in props.iter() {
        let Some((_, entry)) = MERGE_WRITE_TABLE.lookup(key.as_str()) else {
            continue;
        };
        if props.len() != 1 {
            return Err(Error::new(
                Code::InvalidArgument,
                "an XLSB merge cannot be combined with other properties",
            )
            .with_format(FileFormat::Xlsb)
            .with_property("merge")
            .with_suggestion("set `merge` in its own command"));
        }
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |_, value| {
                merge = Some(match value {
                    Some(Scalar::Bool(value)) => value,
                    None => false,
                    _ => {
                        return Err(Error::new(Code::InvalidValue, "merge must be boolean")
                            .with_property("merge"));
                    },
                });
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    let Some(merge) = merge else {
        return Ok(None);
    };
    let targets = resolve_ranges(file, target, limits)?;
    let editor = file.editor_for_write()?;
    let package = package_from_editor(editor);
    let native_limits = litchi_xlsb::cell_values::TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets().saturating_mul(2),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?;
    for target in targets {
        let staged = if merge {
            edit.merge_cells(target.sheet, target.range)
        } else {
            edit.unmerge_cells(target.sheet, target.range)
        };
        staged.map_err(merge_error)?.ok_or_else(|| {
            Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                .with_format(FileFormat::Xlsb)
        })?;
    }
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok(Some(commit))
}

fn resolve_ranges(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<MergeTarget>> {
    let mut targets = Vec::new();
    let mut push = |node: Node<'_>| -> Result<()> {
        if node.element() != Kind::Range {
            return Err(Error::new(
                Code::UnsupportedElement,
                "an XLSB merge targets a range such as /Sheet1/A1:B2",
            )
            .with_format(FileFormat::Xlsb)
            .with_property("merge"));
        }
        let Item::Range(range) = node.item() else {
            return Err(
                Error::new(Code::InvalidPath, "the XLSB range reference is invalid")
                    .with_format(FileFormat::Xlsb),
            );
        };
        let start = range.rect.start();
        let (end_row, end_column) = range.rect.end();
        let merged = MergedCell::new(
            start.row().get(),
            end_row.saturating_sub(1),
            start.column().get(),
            end_column.saturating_sub(1),
        );
        if merged.row_first == merged.row_last && merged.col_first == merged.col_last {
            return Err(Error::new(
                Code::InvalidValue,
                "a merged range must span more than one cell",
            )
            .with_format(FileFormat::Xlsb)
            .with_property("merge"));
        }
        if targets.len() >= limits.max_expanded_targets() {
            return Err(crate::transaction::Error::Limit {
                kind: crate::transaction::LimitKind::ExpandedTargets,
                observed: targets.len().checked_add(1).unwrap_or(usize::MAX),
                maximum: limits.max_expanded_targets(),
            }
            .into());
        }
        targets
            .try_reserve(1)
            .map_err(|_| Error::new(Code::InternalError, "could not reserve XLSB merge targets"))?;
        let target = MergeTarget {
            sheet: node.sheet_position(),
            range: merged,
        };
        if !targets.contains(&target) {
            targets.push(target);
        }
        Ok(())
    };
    if let Some(path) = target.path() {
        push(crate::path::resolve::get(&Nav::new(file), path)?)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsb));
        }
        for node in walk {
            push(node?)?;
        }
        if targets.is_empty() {
            return Err(Error::new(
                Code::ZeroMatches,
                "the mutation selector matched no element",
            )
            .with_format(FileFormat::Xlsb));
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsb),
        );
    }
    file.check_source_for_replay()?;
    Ok(targets)
}

/// Preserve typed refusal classification while attaching merge guidance.
fn merge_error(error: litchi_xlsb::package::PackageError) -> Error {
    match error {
        error @ litchi_xlsb::package::PackageError::Refused(_) => native_error(error)
            .with_property("merge")
            .with_suggestion("name an existing merge exactly, or a range clear of other merges"),
        litchi_xlsb::package::PackageError::UnsupportedFeature(_)
        | litchi_xlsb::package::PackageError::MergedCell(_)
        | litchi_xlsb::package::PackageError::InvalidFormat(_) => Error::new(
            Code::InvalidArgument,
            "the XLSB range overlaps a different merge or is not a valid merge",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("merge")
        .with_suggestion("name an existing merge exactly, or a range clear of other merges"),
        error => native_error(error),
    }
}
