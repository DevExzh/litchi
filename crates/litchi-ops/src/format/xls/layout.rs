//! Unified XLS row and column structure.
//!
//! Row and column insertion and deletion delegate to the native axis shift
//! of the primary `Package::edit`.  The leaf shifts the worksheet's cells,
//! row blocks, selections, merges, hyperlinks and drawing anchors in one
//! closure and refuses workbooks whose formulas, names or tables it cannot
//! rewrite.  Deletion is limited to empty rows and columns because only an
//! empty span has a durable semantic inverse.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xls::cell_values::Selector;

use super::spec::Kind;
use super::write::{WriteCommit, finish_commit, native_edit_error};
use super::{Nav, Xls};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

const BIFF8_ROWS: usize = 65_536;
const BIFF8_COLUMNS: usize = 256;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Axis {
    Row,
    Column,
}

impl Axis {
    fn of(kind: &str) -> Option<Self> {
        if kind.eq_ignore_ascii_case("row") {
            Some(Self::Row)
        } else if kind.eq_ignore_ascii_case("col") || kind.eq_ignore_ascii_case("column") {
            Some(Self::Column)
        } else {
            None
        }
    }

    const fn size(self) -> usize {
        match self {
            Self::Row => BIFF8_ROWS,
            Self::Column => BIFF8_COLUMNS,
        }
    }

    const fn kind(self) -> Kind {
        match self {
            Self::Row => Kind::Row,
            Self::Column => Kind::Column,
        }
    }
}

fn refusal(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Xls)
}

fn admit(limits: &SemanticLimits) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(Error::from(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: 1,
            maximum: limits.max_operations(),
        })
        .with_format(FileFormat::Xls));
    }
    Ok(())
}

fn axis_path(sheet: usize, axis: Axis, index: usize) -> Path<'static> {
    let sheet = Step::new("sheet").with(Predicate::Index(Position::new(sheet)));
    let step = match axis {
        Axis::Row => Step::new("row").with(Predicate::Index(Position::new(index))),
        Axis::Column => crate::node::PathStep::column("col", index as u32).to_step(),
    };
    Path::from_steps([sheet, step])
}

/// Resolve a `/sheet/row[N]` or `/sheet/col[X]` path to its zero-based index.
fn resolve_axis(file: &Xls, path: &PathArg<'_>, axis: Axis) -> Result<(usize, usize)> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.element() != axis.kind() {
        return Err(refusal(
            Code::UnsupportedType,
            "the XLS axis placement must name a row or column of the same kind",
        ));
    }
    let number = node
        .number()
        .ok_or_else(|| refusal(Code::InvalidPath, "the XLS axis path has no index"))?;
    Ok((node.sheet_position(), number as usize - 1))
}

/// Insert one empty row or column; later rows or columns shift by one.
pub(crate) fn add_axis(
    file: &Xls,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, Path<'static>)> {
    let Some(axis) = Axis::of(node.kind()) else {
        return Err(refusal(
            Code::UnsupportedElement,
            "XLS adds sheet tabs under the workbook and rows or columns under a sheet",
        ));
    };
    file.admit_edit_for_write()?;
    if !node.props().is_empty() {
        return Err(refusal(
            Code::UnsupportedCapability,
            "XLS row and column insertion takes no properties; set them after inserting",
        ));
    }
    let parent = crate::path::resolve::get(&Nav::new(file), parent)?;
    if parent.element() != Kind::Sheet {
        return Err(refusal(
            Code::UnsupportedType,
            "an XLS row or column can only be added directly under a sheet",
        ));
    }
    let sheet = parent.sheet_position();
    let index = match node.placement() {
        None | Some(Placement::End) => {
            return Err(refusal(
                Code::MissingArgument,
                "XLS row and column insertion requires an explicit position",
            ));
        },
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => {
            let (anchor_sheet, index) = resolve_axis(file, path, axis)?;
            same_sheet(sheet, anchor_sheet)?;
            index
        },
        Some(Placement::After(path)) => {
            let (anchor_sheet, index) = resolve_axis(file, path, axis)?;
            same_sheet(sheet, anchor_sheet)?;
            index.saturating_add(1)
        },
        Some(Placement::Into(_, Some(position))) => position.get(),
        Some(Placement::Into(_, None)) => {
            return Err(refusal(
                Code::MissingArgument,
                "XLS row and column insertion requires an explicit position",
            ));
        },
    };
    if index >= axis.size() {
        return Err(refusal(
            Code::OutOfRange,
            "the XLS row or column insertion position is outside the BIFF8 grid",
        ));
    }
    admit(limits)?;
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 1)?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    let selector = Selector::Position(Position::new(sheet));
    match axis {
        Axis::Row => edit.insert_rows(selector, index as u16, 1),
        Axis::Column => edit.insert_columns(selector, index as u8, 1),
    }
    .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    Ok((
        finish_commit(file, native, 1, limits)?,
        axis_path(sheet, axis, index),
    ))
}

fn same_sheet(parent: usize, anchor: usize) -> Result<()> {
    if parent == anchor {
        Ok(())
    } else {
        Err(refusal(
            Code::InvalidArgument,
            "the XLS axis placement names another sheet",
        ))
    }
}

/// Delete one empty row or column selected by a direct axis path.  Returns
/// `None` when the target is not a row or column path.
pub(crate) fn remove_axis(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<WriteCommit>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    let axis = match node.element() {
        Kind::Row => Axis::Row,
        Kind::Column => Axis::Column,
        _ => return Ok(None),
    };
    let index = node
        .number()
        .ok_or_else(|| refusal(Code::InvalidPath, "the XLS axis path has no index"))?
        as usize
        - 1;
    let sheet = node.sheet_position();
    admit(limits)?;
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 1)?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    let selector = Selector::Position(Position::new(sheet));
    match axis {
        Axis::Row => edit.delete_rows(selector, index as u16, 1),
        Axis::Column => edit.delete_columns(selector, index as u8, 1),
    }
    .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits).map(Some)
}

/// One resolved layout target.
enum LayoutTarget {
    Row { sheet: usize, row: u16 },
    Column { sheet: usize, column: u8 },
}

/// Admitted row or column layout assignments.
#[derive(Default)]
struct LayoutAssignments {
    /// Row height in twips, or column width in 1/256 character.
    size: Option<u16>,
    hidden: Option<bool>,
}

fn resolve_layout_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<(Axis, Vec<LayoutTarget>)>> {
    let mut nodes = Vec::new();
    if let Some(path) = target.path() {
        // Only each node's kind, number and sheet are read below, so a path
        // resolves its coordinates without the worksheet grid (record 0310).
        nodes.push(crate::path::resolve::get(&Nav::coordinates(file), path)?);
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(refusal(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            let node = node?;
            if nodes.len() >= limits.max_expanded_targets() {
                return Err(Error::from(TransactionError::Limit {
                    kind: LimitKind::ExpandedTargets,
                    observed: nodes.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: limits.max_expanded_targets(),
                })
                .with_format(FileFormat::Xls));
            }
            nodes.push(node);
        }
    } else {
        return Ok(None);
    }
    let Some(first) = nodes.first() else {
        return Ok(None);
    };
    let axis = match first.element() {
        Kind::Row => Axis::Row,
        Kind::Column => Axis::Column,
        _ => return Ok(None),
    };
    let mut targets = Vec::new();
    for node in nodes {
        if node.element() != axis.kind() {
            return Err(refusal(
                Code::InvalidArgument,
                "an XLS layout selector must match only rows or only columns",
            ));
        }
        let index = node
            .number()
            .and_then(|number| number.checked_sub(1))
            .ok_or_else(|| refusal(Code::InvalidPath, "the XLS axis path has no index"))?;
        let sheet = node.sheet_position();
        targets.push(match axis {
            Axis::Row => LayoutTarget::Row {
                sheet,
                row: u16::try_from(index)
                    .map_err(|_| refusal(Code::OutOfRange, "the XLS row is outside BIFF8"))?,
            },
            Axis::Column => LayoutTarget::Column {
                sheet,
                column: u8::try_from(index)
                    .map_err(|_| refusal(Code::OutOfRange, "the XLS column is outside BIFF8"))?,
            },
        });
    }
    file.check_source()?;
    Ok(Some((axis, targets)))
}

fn admit_layout(axis: Axis, props: &crate::props::Props<'_>) -> Result<LayoutAssignments> {
    use super::spec::{COLUMN_WRITE_TABLE, ColumnSetter, ROW_WRITE_TABLE, RowSetter};
    use crate::registry::Operation;
    use crate::value::Scalar;

    let mut assignments = LayoutAssignments::default();
    let duplicate = |name: &'static str| {
        refusal(
            Code::InvalidArgument,
            "the same XLS layout property was supplied more than once",
        )
        .with_property(name)
    };
    for (key, value) in props.iter() {
        let with_key = |error: Error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        };
        match axis {
            Axis::Row => {
                let (_, entry) = ROW_WRITE_TABLE.resolve(key.as_str()).map_err(with_key)?;
                let entry = *entry;
                entry
                    .prepare(Operation::Set, value, |setter, value| match (setter, value) {
                        (RowSetter::Height, Some(Scalar::Length(length))) => {
                            let twips = (length.points() * 20.0).round();
                            if !(1.0..=8_180.0).contains(&twips) {
                                return Err(refusal(
                                    Code::InvalidValue,
                                    "an XLS row height must be greater than zero and at most 409 points",
                                )
                                .with_property("height"));
                            }
                            if assignments.size.replace(twips as u16).is_some() {
                                return Err(duplicate("height"));
                            }
                            Ok(())
                        },
                        (RowSetter::Hidden, Some(Scalar::Bool(hidden))) => {
                            if assignments.hidden.replace(hidden).is_some() {
                                return Err(duplicate("hidden"));
                            }
                            Ok(())
                        },
                        _ => Err(refusal(Code::InvalidValue, "invalid XLS row layout value")),
                    })
                    .map_err(with_key)?;
            },
            Axis::Column => {
                let (_, entry) = COLUMN_WRITE_TABLE.resolve(key.as_str()).map_err(with_key)?;
                let entry = *entry;
                entry
                    .prepare(Operation::Set, value, |setter, value| {
                        match (setter, value) {
                            (ColumnSetter::Width, Some(Scalar::Number(width))) => {
                                let units = (width * 256.0).round();
                                if assignments.size.replace(units as u16).is_some() {
                                    return Err(duplicate("width"));
                                }
                                Ok(())
                            },
                            (ColumnSetter::Width, Some(Scalar::Int(width)))
                                if (0..=255).contains(&width) =>
                            {
                                if assignments.size.replace((width * 256) as u16).is_some() {
                                    return Err(duplicate("width"));
                                }
                                Ok(())
                            },
                            (ColumnSetter::Hidden, Some(Scalar::Bool(hidden))) => {
                                if assignments.hidden.replace(hidden).is_some() {
                                    return Err(duplicate("hidden"));
                                }
                                Ok(())
                            },
                            _ => Err(refusal(
                                Code::InvalidValue,
                                "invalid XLS column layout value",
                            )),
                        }
                    })
                    .map_err(with_key)?;
            },
        }
    }
    Ok(assignments)
}

/// Set row heights, column widths and hidden flags selected by a row or
/// column path or a homogeneous row/column query.  Returns `None` when the
/// target is not a row or column.
pub(crate) fn set_layout(
    file: &Xls,
    target: &TargetArg<'_>,
    props: &crate::props::Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<WriteCommit>> {
    let Some((axis, targets)) = resolve_layout_targets(file, target, limits)? else {
        return Ok(None);
    };
    let assignments = admit_layout(axis, props)?;
    if assignments.size.is_none() && assignments.hidden.is_none() || targets.is_empty() {
        let native = file
            .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
            .edit()
            .map_err(native_edit_error)?
            .commit()
            .map_err(native_edit_error)?;
        return finish_commit(file, native, 0, limits).map(Some);
    }
    admit(limits)?;
    let staged = targets.len().saturating_mul(16);
    if staged > limits.max_staged_bytes() {
        return Err(Error::from(TransactionError::Limit {
            kind: LimitKind::StagedBytes,
            observed: staged,
            maximum: limits.max_staged_bytes(),
        })
        .with_format(FileFormat::Xls));
    }
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), staged)?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    for target in targets {
        match target {
            LayoutTarget::Row { sheet, row } => {
                if let Some(twips) = assignments.size {
                    edit.set_row_height(Selector::Position(Position::new(sheet)), row, twips)
                        .map_err(native_edit_error)?;
                }
                if let Some(hidden) = assignments.hidden {
                    edit.set_row_hidden(Selector::Position(Position::new(sheet)), row, hidden)
                        .map_err(native_edit_error)?;
                }
            },
            LayoutTarget::Column { sheet, column } => {
                if let Some(width) = assignments.size {
                    edit.set_column_width(Selector::Position(Position::new(sheet)), column, width)
                        .map_err(native_edit_error)?;
                }
                if let Some(hidden) = assignments.hidden {
                    edit.set_column_hidden(
                        Selector::Position(Position::new(sheet)),
                        column,
                        hidden,
                    )
                    .map_err(native_edit_error)?;
                }
            },
        }
    }
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits).map(Some)
}
