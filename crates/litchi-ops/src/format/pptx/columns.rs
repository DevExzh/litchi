//! Table grid columns (record 0321, L-PPTX-R8).
//!
//! `/slide[N]/table[M]/col[C]` reads one `a:gridCol` (`width`). `add` under
//! a table inserts a column (`width`, default the average of the existing
//! columns; seed `text` for its cells) at a start, end or position
//! placement; `set width` resizes it; `remove` deletes it with its cell in
//! every row. The frame's width follows the sum of the column widths, as
//! OfficeCLI keeps it. Tables with cells merged across columns refuse
//! insertion and removal.
//!
//! Source owners are promoted to an eager package, as for theme writes.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};

use super::node::{Item, TreeItem};
use super::spec::{COLUMN_WRITE_TABLE, ColumnSetter, Kind};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options};
use super::{Editable, Nav, Pptx};
use crate::node::Value;
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{Length, Scalar};
use crate::{Code, NewNode, Placement, Props, Result};

/// One grid column of a table.
#[derive(Clone)]
pub(super) struct ColumnItem {
    pub(super) table: TreeItem,
    pub(super) widths: Arc<[i64]>,
    pub(super) index: u32,
}

/// The grid widths of a table frame.
pub(super) fn widths(table: &TreeItem) -> Result<Arc<[i64]>> {
    Ok(Arc::from(table.handle.table()?.grid_columns()?))
}

/// The columns of a table frame.
pub(super) fn children(table: &TreeItem) -> Result<Vec<Item>> {
    let widths = widths(table)?;
    Ok((0..widths.len())
        .map(|index| {
            Item::Column(Box::new(ColumnItem {
                table: table.clone(),
                widths: Arc::clone(&widths),
                index: u32::try_from(index).unwrap_or(u32::MAX),
            }))
        })
        .collect())
}

/// The value of the column property at `index`.
pub(super) fn value(item: &ColumnItem, index: usize) -> Option<Value<'static>> {
    match super::spec::ColumnKey::at(index)? {
        super::spec::ColumnKey::Width => item
            .widths
            .get(item.index as usize)
            .map(|width| Value::Length(Length::from_emu(*width))),
    }
}

/// The table frame and slide of a write target: a top-level slide table.
fn slide_table(table: &TreeItem) -> Result<usize> {
    let Some(slide) = table.container.slide() else {
        return Err(error(
            Code::UnsupportedCapability,
            "master and layout tables are not editable",
        ));
    };
    if table.handle.parent().is_some() {
        return Err(error(
            Code::UnsupportedCapability,
            "table column writes require a top-level slide table",
        ));
    }
    Ok(slide as usize)
}

enum Change {
    Width {
        column: usize,
        width: i64,
    },
    Insert {
        position: usize,
        width: Option<i64>,
        text: String,
    },
    Remove {
        column: usize,
    },
}

impl Pptx {
    fn change_columns(
        &self,
        table: &TreeItem,
        change: Change,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let slide = slide_table(table)?;
        let shape = table.handle.position().get();
        let source = if matches!(&change, Change::Width { .. }) {
            self.editable_owner_for_write()?
        } else {
            self.editable_owner()?
        };
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        match change {
            Change::Width { column, width } => {
                native.set_table_column_width(slide, shape, Position::new(column), width)?;
            },
            Change::Insert {
                position,
                width,
                text,
            } => native.insert_table_column(slide, shape, Position::new(position), width, &text)?,
            Change::Remove { column } => {
                native.remove_table_column(slide, shape, Position::new(column))?;
            },
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        ))
    }

    /// `add(/slide[N]/table[M], col, {width, text})`.
    pub(super) fn add_column(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Tree(table) = resolved.item() else {
            return Err(error(
                Code::UnsupportedType,
                "columns require a table parent",
            ));
        };
        if table.kind != Kind::Table {
            return Err(error(
                Code::UnsupportedType,
                "columns require a table parent",
            ));
        }
        let table = table.clone();
        drop(resolved);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let count = widths(&table)?.len();
        let position = match node.placement() {
            None | Some(Placement::End) => count,
            Some(Placement::Start) => 0,
            Some(Placement::Position(position)) => position.get(),
            Some(_) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "column insertion supports start, end and position placement",
                ));
            },
        };
        if position > count {
            return Err(error(
                Code::OutOfRange,
                "the column position is past the last column",
            ));
        }
        let mut width = None;
        let mut text = String::new();
        let mut seen = 0_u8;
        for (key, value) in node.props().iter() {
            let (index, entry) = COLUMN_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u8 << index;
            if seen & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a column property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            seen |= bit;
            entry.prepare(Operation::Add, value, |setter, value| {
                match (setter, value) {
                    (ColumnSetter::Width, Some(Scalar::Length(length))) => {
                        width = Some(length.emu());
                    },
                    (ColumnSetter::Text, Some(Scalar::Text(value))) => {
                        bound(
                            LimitKind::StagedBytes,
                            value.len().saturating_mul(text_cells(&table)?),
                            limits.max_staged_bytes(),
                        )?;
                        super::write::validate_text(value)?;
                        value.clone_into(&mut text);
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let slide = slide_table(&table)?;
        let commit = self.change_columns(
            &table,
            Change::Insert {
                position,
                width,
                text,
            },
            limits,
        )?;
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("table").with(Predicate::Index(Position::new(table.ordinal as usize))),
            Step::new("col").with(Predicate::Index(Position::new(position))),
        ]);
        Ok((commit, path))
    }

    /// A `set` on a column path; `None` for other targets.
    pub(super) fn try_set_column(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.column_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut width = None;
        for (key, value) in props.iter() {
            let (_, entry) = COLUMN_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            entry.prepare(Operation::Set, value, |setter, value| {
                if let (ColumnSetter::Width, Some(Scalar::Length(length))) = (setter, value) {
                    width = Some(length.emu());
                }
                Ok(())
            })?;
        }
        let Some(width) = width else {
            return Err(error(Code::MissingProperty, "nothing to set").with_property("width"));
        };
        self.change_columns(
            &item.table,
            Change::Width {
                column: item.index as usize,
                width,
            },
            limits,
        )
        .map(Some)
    }

    /// A `remove` of a column path; `None` for other targets.
    pub(super) fn try_remove_column(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.column_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        self.change_columns(
            &item.table,
            Change::Remove {
                column: item.index as usize,
            },
            limits,
        )
        .map(Some)
    }

    /// The column a path target names.
    fn column_target(&self, target: &TargetArg<'_>) -> Result<Option<ColumnItem>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            step.name().eq_ignore_ascii_case("col") || step.name().eq_ignore_ascii_case("column")
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        Ok(match node.item() {
            Item::Column(item) => Some(ColumnItem::clone(item)),
            _ => None,
        })
    }
}

/// The number of rows of a table: the cells a column's seed text fills.
fn text_cells(table: &TreeItem) -> Result<usize> {
    Ok(super::text::table_shape(&table.handle)?.0.max(1))
}
