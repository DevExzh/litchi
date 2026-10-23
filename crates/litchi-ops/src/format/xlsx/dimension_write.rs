//! Bounded XLSX row and column property writes.
//!
//! This route owns only the four dimension properties that have a stable
//! unified vocabulary and a typed native editor: row `height`, `hidden`,
//! `outlineLevel` and `collapsed`, and the corresponding column `width`,
//! `hidden`, `outlineLevel` and `collapsed`.  It resolves a row or column
//! selector against the immutable owner and stages every match in one native
//! workbook transaction.  The native leaf remains responsible for worksheet
//! protection, markup-compatibility fences and the exact SpreadsheetML
//! rewrite.

use std::collections::HashSet;
use std::collections::TryReserveError;

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, HeightAt, RowAt, WidthAt};

use super::node::Node;
use super::spec::{COLUMN_WRITE_TABLE, ColumnSetter, Kind, ROW_WRITE_TABLE, RowSetter};
use super::write::{Commit, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::props::{Key, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

const WRITE_ENTRY_OVERHEAD: usize = 64;

#[derive(Debug, Default)]
struct RowAssignments {
    height: Option<Option<f64>>,
    hidden: Option<Option<bool>>,
    outline_level: Option<Option<i32>>,
    collapsed: Option<Option<bool>>,
}

#[derive(Debug, Default)]
struct ColumnAssignments {
    width: Option<Option<f64>>,
    hidden: Option<Option<bool>>,
    outline_level: Option<Option<i32>>,
    collapsed: Option<Option<bool>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
enum DimensionKind {
    Row,
    Column,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
struct DimensionTarget {
    kind: DimensionKind,
    sheet: usize,
    index: u32,
}

struct ResolvedDimensions {
    kind: DimensionKind,
    targets: Vec<DimensionTarget>,
}

enum DimensionAssignments {
    Row(RowAssignments),
    Column(ColumnAssignments),
}

/// Try the row/column property route.
///
/// `Ok(None)` means that the target belongs to another mutation family.  A
/// selector whose first match is a cell, sheet or other unrelated element is
/// left for that family's adapter; once a row or column match is observed,
/// later mixed matches are refused so a partial dimension write cannot fall
/// through to a different editor.
pub(crate) fn try_set_dimensions(
    file: &Xlsx,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    if direct_cell_or_range_path(target) {
        return Ok(None);
    }

    file.check_source_for_replay()?;
    let result = (|| {
        let Some(resolved) = resolve_targets(file, target, limits)? else {
            return Ok(None);
        };
        let (charge, assignments) = match resolved.kind {
            DimensionKind::Row => {
                let assignments = admit_rows(props)?;
                (
                    row_charge(&assignments, limits)?,
                    DimensionAssignments::Row(assignments),
                )
            },
            DimensionKind::Column => {
                let assignments = admit_columns(props)?;
                (
                    column_charge(&assignments, limits)?,
                    DimensionAssignments::Column(assignments),
                )
            },
        };
        check_target_count(resolved.targets.len(), charge, limits)?;

        let source = file.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        match assignments {
            DimensionAssignments::Row(assignments) => {
                apply_rows(&mut edit, &resolved.targets, &assignments)?;
            },
            DimensionAssignments::Column(assignments) => {
                apply_columns(&mut edit, &resolved.targets, &assignments)?;
            },
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(file, source, native, limits.max_snapshot_bytes()).map(Some)
    })();
    file.check_source_for_replay()?;
    result
}

impl Xlsx {
    /// Assign row or column properties through one source-fenced native edit.
    pub(crate) fn try_set_dimensions(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        try_set_dimensions(self, target, props, limits)
    }
}

fn resolve_targets(
    file: &Xlsx,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<ResolvedDimensions>> {
    if direct_cell_or_range_path(target) {
        return Ok(None);
    }

    let nav = Nav::new(file);
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&nav, path)?;
        let Some((kind, target)) = node_target(node)? else {
            return Ok(None);
        };
        return Ok(Some(ResolvedDimensions {
            kind,
            targets: vec![target],
        }));
    }

    let Some(query) = target.query() else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsx),
        );
    };
    let walk = crate::query::run(nav, query)?;
    if !walk.plan().diagnostics().is_empty() {
        return Err(Error::new(
            Code::InvalidSelector,
            "the mutation selector contains an unsupported predicate",
        )
        .with_format(FileFormat::Xlsx));
    }

    let mut kind = None;
    let mut targets = Vec::new();
    let mut seen = HashSet::new();
    let mut matched = false;
    for node in walk {
        let node = node?;
        matched = true;
        let Some((node_kind, target)) = node_target(node)? else {
            // Do not walk an unrelated selector merely to prove that it has
            // no row/column match. Its own adapter is the authority for that
            // family and will apply its own expansion budget.
            if kind.is_none() {
                return Ok(None);
            }
            return Err(mixed_selector());
        };
        if let Some(previous) = kind {
            if previous != node_kind {
                return Err(mixed_selector());
            }
        } else {
            kind = Some(node_kind);
        }
        append_target(&mut targets, &mut seen, target, limits)?;
    }
    file.check_source_for_replay()?;
    if !matched {
        return Err(Error::new(
            Code::ZeroMatches,
            "the mutation selector matched no element",
        )
        .with_format(FileFormat::Xlsx));
    }
    let Some(kind) = kind else {
        return Ok(None);
    };
    targets.sort_unstable();
    Ok(Some(ResolvedDimensions { kind, targets }))
}

fn node_target(node: Node<'_>) -> Result<Option<(DimensionKind, DimensionTarget)>> {
    let kind = match node.element() {
        Kind::Row => DimensionKind::Row,
        Kind::Column => DimensionKind::Column,
        _ => return Ok(None),
    };
    let number = node.number().ok_or_else(|| {
        Error::new(
            Code::InvalidPath,
            "the resolved XLSX dimension has no coordinate",
        )
        .with_format(FileFormat::Xlsx)
    })?;
    let index = number.checked_sub(1).ok_or_else(|| {
        Error::new(
            Code::InvalidPath,
            "the resolved XLSX dimension coordinate is invalid",
        )
        .with_format(FileFormat::Xlsx)
    })?;
    Ok(Some((
        kind,
        DimensionTarget {
            kind,
            sheet: node.sheet_position(),
            index,
        },
    )))
}

fn append_target(
    targets: &mut Vec<DimensionTarget>,
    seen: &mut HashSet<DimensionTarget>,
    target: DimensionTarget,
    limits: &SemanticLimits,
) -> Result<()> {
    if seen.contains(&target) {
        return Ok(());
    }
    let observed = targets
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if observed > limits.max_expanded_targets() {
        return Err(expanded_limit(observed, limits.max_expanded_targets()));
    }
    targets
        .try_reserve(1)
        .map_err(|source| allocation("XLSX dimension write targets", source))?;
    seen.try_reserve(1)
        .map_err(|source| allocation("XLSX dimension write target identities", source))?;
    seen.insert(target);
    targets.push(target);
    Ok(())
}

fn direct_cell_or_range_path(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        path.path().steps().len() > 1
            && path
                .path()
                .steps()
                .last()
                .is_some_and(|step| step.cell_range().is_some())
    })
}

fn mixed_selector() -> Error {
    Error::new(
        Code::UnsupportedElement,
        "an XLSX dimension selector cannot mix rows, columns or other elements",
    )
    .with_format(FileFormat::Xlsx)
}

fn check_target_count(
    count: usize,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    let staged = count.saturating_mul(per_target_charge);
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn row_charge(assignments: &RowAssignments, limits: &SemanticLimits) -> Result<usize> {
    // Charge the logical staged representation, rather than a native XML
    // rewrite whose size depends on source spelling and compaction.  Resets
    // still consume their field slot, so null cannot evade the budget.
    let payload = assignments.height.is_some() as usize * size_of::<f64>()
        + assignments.hidden.is_some() as usize * size_of::<bool>()
        + assignments.outline_level.is_some() as usize * size_of::<i32>()
        + assignments.collapsed.is_some() as usize * size_of::<bool>();
    payload
        .checked_add(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
}

fn column_charge(assignments: &ColumnAssignments, limits: &SemanticLimits) -> Result<usize> {
    let payload = assignments.width.is_some() as usize * size_of::<f64>()
        + assignments.hidden.is_some() as usize * size_of::<bool>()
        + assignments.outline_level.is_some() as usize * size_of::<i32>()
        + assignments.collapsed.is_some() as usize * size_of::<bool>();
    payload
        .checked_add(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
}

fn admit_rows(props: &Props<'_>) -> Result<RowAssignments> {
    let mut assignments = RowAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = ROW_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                match setter {
                    RowSetter::Height => {
                        assign_once(&mut assignments.height, entry.spec.name)?;
                        assignments.height = Some(match value {
                            None => None,
                            Some(Scalar::Length(value)) => Some(value.points()),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    RowSetter::Hidden => {
                        assign_once(&mut assignments.hidden, entry.spec.name)?;
                        assignments.hidden = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    RowSetter::Outline => {
                        assign_once(&mut assignments.outline_level, entry.spec.name)?;
                        assignments.outline_level = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => Some(
                                i32::try_from(value)
                                    .map_err(|_| invalid_property(entry.spec.name))?,
                            ),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    RowSetter::Collapsed => {
                        assign_once(&mut assignments.collapsed, entry.spec.name)?;
                        assignments.collapsed = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                }
                Ok(())
            })
            .map_err(|error| property_error(error, key))?;
    }
    Ok(assignments)
}

fn admit_columns(props: &Props<'_>) -> Result<ColumnAssignments> {
    let mut assignments = ColumnAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = COLUMN_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                match setter {
                    ColumnSetter::Width => {
                        assign_once(&mut assignments.width, entry.spec.name)?;
                        assignments.width = Some(match value {
                            None => None,
                            Some(Scalar::Number(value)) => Some(value),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    ColumnSetter::Hidden => {
                        assign_once(&mut assignments.hidden, entry.spec.name)?;
                        assignments.hidden = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    ColumnSetter::Outline => {
                        assign_once(&mut assignments.outline_level, entry.spec.name)?;
                        assignments.outline_level = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => Some(
                                i32::try_from(value)
                                    .map_err(|_| invalid_property(entry.spec.name))?,
                            ),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                    ColumnSetter::Collapsed => {
                        assign_once(&mut assignments.collapsed, entry.spec.name)?;
                        assignments.collapsed = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => return Err(invalid_property(entry.spec.name)),
                        });
                    },
                }
                Ok(())
            })
            .map_err(|error| property_error(error, key))?;
    }
    Ok(assignments)
}

fn assign_once<T>(slot: &mut Option<T>, key: &'static str) -> Result<()> {
    if slot.is_some() {
        Err(Error::new(
            Code::InvalidArgument,
            "the same XLSX dimension property was supplied more than once",
        )
        .with_format(FileFormat::Xlsx)
        .with_property(key))
    } else {
        Ok(())
    }
}

fn apply_rows(
    edit: &mut xlsx::workbook::Edit,
    targets: &[DimensionTarget],
    assignments: &RowAssignments,
) -> Result<()> {
    let mut position = 0;
    while position < targets.len() {
        let sheet_position = targets[position].sheet;
        let mut sheet = edit
            .sheet(sheet_position)
            .map_err(native_error)?
            .ok_or_else(|| out_of_range("the XLSX worksheet position is out of range"))?;
        while let Some(target) = targets
            .get(position)
            .filter(|target| target.sheet == sheet_position)
        {
            let mut row = sheet.row(RowAt::from(target.index)).map_err(native_error)?;
            if let Some(height) = assignments.height {
                match height {
                    Some(height) => {
                        row.height(HeightAt::from(height))
                            .map_err(|error| native_property_error(error, "height"))?;
                    },
                    None => {
                        row.reset_height();
                    },
                }
            }
            if let Some(hidden) = assignments.hidden {
                match hidden {
                    Some(true) => {
                        row.hide();
                    },
                    Some(false) | None => {
                        row.show();
                    },
                }
            }
            if let Some(outline) = assignments.outline_level {
                row.outline(outline.unwrap_or(0))
                    .map_err(|error| native_property_error(error, "outlineLevel"))?;
            }
            if let Some(collapsed) = assignments.collapsed {
                match collapsed {
                    Some(true) => {
                        row.collapse();
                    },
                    Some(false) | None => {
                        row.expand();
                    },
                }
            }
            position += 1;
        }
    }
    Ok(())
}

fn apply_columns(
    edit: &mut xlsx::workbook::Edit,
    targets: &[DimensionTarget],
    assignments: &ColumnAssignments,
) -> Result<()> {
    let mut position = 0;
    while position < targets.len() {
        let sheet_position = targets[position].sheet;
        let mut sheet = edit
            .sheet(sheet_position)
            .map_err(native_error)?
            .ok_or_else(|| out_of_range("the XLSX worksheet position is out of range"))?;
        while let Some(target) = targets
            .get(position)
            .filter(|target| target.sheet == sheet_position)
        {
            let mut column = sheet
                .column(xlsx::ColumnAt::from(target.index))
                .map_err(native_error)?;
            if let Some(width) = assignments.width {
                match width {
                    Some(width) => {
                        column
                            .width(WidthAt::from(width))
                            .map_err(|error| native_property_error(error, "width"))?;
                    },
                    None => {
                        column.reset_width();
                    },
                }
            }
            if let Some(hidden) = assignments.hidden {
                match hidden {
                    Some(true) => {
                        column.hide();
                    },
                    Some(false) | None => {
                        column.show();
                    },
                }
            }
            if let Some(outline) = assignments.outline_level {
                column
                    .outline(outline.unwrap_or(0))
                    .map_err(|error| native_property_error(error, "outlineLevel"))?;
            }
            if let Some(collapsed) = assignments.collapsed {
                match collapsed {
                    Some(true) => {
                        column.collapse();
                    },
                    Some(false) | None => {
                        column.expand();
                    },
                }
            }
            position += 1;
        }
    }
    Ok(())
}

fn invalid_property(key: &'static str) -> Error {
    Error::new(
        Code::InvalidValue,
        "the XLSX dimension property has the wrong type",
    )
    .with_format(FileFormat::Xlsx)
    .with_property(key)
}

fn property_error(error: Error, key: &Key<'_>) -> Error {
    error
        .with_format(FileFormat::Xlsx)
        .with_property(key.as_str().to_owned())
}

fn native_property_error(error: xlsx::Error, key: &'static str) -> Error {
    native_error(error)
        .caused_by_argument(Code::InvalidValue)
        .with_property(key)
}

fn native_error(error: xlsx::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xlsx)
}

fn out_of_range(message: &'static str) -> Error {
    Error::new(Code::OutOfRange, message).with_format(FileFormat::Xlsx)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::from(xlsx::Error::Allocation { resource, source }).with_format(FileFormat::Xlsx)
}
