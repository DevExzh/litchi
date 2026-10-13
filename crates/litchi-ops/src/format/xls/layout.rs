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
    file.admit_edit()?;
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
    let mut edit = file.editable_package()?.edit().map_err(native_edit_error)?;
    let selector = Selector::Position(sheet);
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
    let mut edit = file.editable_package()?.edit().map_err(native_edit_error)?;
    let selector = Selector::Position(sheet);
    match axis {
        Axis::Row => edit.delete_rows(selector, index as u16, 1),
        Axis::Column => edit.delete_columns(selector, index as u8, 1),
    }
    .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits).map(Some)
}
