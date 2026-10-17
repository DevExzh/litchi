//! Table rows (record 0333).
//!
//! `/slide[N]/table[M]/tr[R]` reads one `a:tr` (`height`). `add` under a
//! table inserts a row (`height`, default the first row's; OfficeCLI's
//! `cols` check; seed texts `c1`…`c20`) at a start, end or position
//! placement or beside a sibling row; `set height` resizes it; `remove`
//! deletes it. The frame's height follows the sum of the row heights, as
//! OfficeCLI keeps it, and vertical merges stay consistent: a row is not
//! inserted inside one, and removing a row shrinks the merges it is part of.
//!
//! Source owners are promoted to an eager package, as column writes are.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::Path;

use super::node::{Item, TableItem, TreeItem};
use super::spec::{Kind, ROW_WRITE_TABLE, RowSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, validate_text};
use super::{Editable, Nav, Pptx};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Props, Result};

/// The slide of a top-level slide table.
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
            "table row writes require a top-level slide table",
        ));
    }
    Ok(slide as usize)
}

enum Change {
    Insert {
        position: usize,
        height: Option<i64>,
        texts: Vec<String>,
    },
    Height {
        row: usize,
        height: i64,
    },
    Remove {
        row: usize,
    },
}

impl Pptx {
    fn change_rows(&self, table: &TreeItem, change: Change) -> Result<Commit> {
        let slide = slide_table(table)?;
        let shape = table.handle.position().get();
        let source = self.editable_owner()?;
        let base = super::smartart::eager(&source)?;
        let mut native = base.edit()?;
        match change {
            Change::Insert {
                position,
                height,
                texts,
            } => {
                let texts: Vec<&str> = texts.iter().map(String::as_str).collect();
                native.insert_table_row(slide, shape, Position::new(position), height, &texts)?;
            },
            Change::Height { row, height } => {
                native.set_table_row_height(slide, shape, Position::new(row), height)?;
            },
            Change::Remove { row } => native.remove_table_row(slide, shape, Position::new(row))?,
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Commit {
            after,
            source,
            target: Editable::Eager(target),
            proof: Proof::Eager(patch),
            operations: 1,
        })
    }

    /// `add(/slide[N]/table[M], tr, {height, cols, c1…c20})`.
    pub(super) fn add_row(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let nav = Nav::new(self);
        let resolved = crate::path::resolve::get(&nav, &parent)?;
        let Item::Tree(table) = resolved.item() else {
            return Err(error(Code::UnsupportedType, "rows require a table parent"));
        };
        if table.kind != Kind::Table {
            return Err(error(Code::UnsupportedType, "rows require a table parent"));
        }
        let table = table.clone();
        let parent_path = crate::node::Project::path(&resolved).to_string();
        drop(resolved);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let rows = table.table_grid()?.map_or(0, |grid| grid.len());
        // The grid is read only when a key checks or seeds columns.
        let columns = if node.props().iter().any(|(key, _)| {
            key.as_str()
                .get(..1)
                .is_some_and(|first| first.eq_ignore_ascii_case("c"))
        }) {
            super::columns::widths(&table)?.len()
        } else {
            0
        };
        let position = match node.placement() {
            None | Some(Placement::End) => rows,
            Some(Placement::Start) => 0,
            Some(Placement::Position(position)) => position.get().min(rows),
            Some(Placement::Before(path) | Placement::After(path)) => {
                let sibling = crate::path::resolve::get(&nav, path)?;
                let Item::Row(TableItem {
                    table: sibling_table,
                    row,
                    ..
                }) = sibling.item()
                else {
                    return Err(error(
                        Code::InvalidArgument,
                        "a row goes beside a row of the same table",
                    ));
                };
                if sibling_table.container.slide() != table.container.slide()
                    || sibling_table.handle.position() != table.handle.position()
                {
                    return Err(error(
                        Code::InvalidArgument,
                        "a row goes beside a row of the same table",
                    ));
                }
                *row as usize + usize::from(matches!(node.placement(), Some(Placement::After(_))))
            },
            Some(Placement::Into(..)) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a row is placed at the start, the end, a position or beside a sibling",
                ));
            },
        };
        let mut height = None;
        let mut texts: Vec<String> = Vec::new();
        let mut staged = 0_usize;
        let mut seen = 0_u32;
        for (key, value) in node.props().iter() {
            let (index, entry) = ROW_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if seen & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a row property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            seen |= bit;
            entry.prepare(Operation::Add, value, |setter, value| {
                match (setter, value) {
                    (RowSetter::Height, Some(Scalar::Length(length))) => {
                        height = Some(length.emu());
                    },
                    (RowSetter::Cols, Some(Scalar::Int(cols))) => {
                        if usize::try_from(cols).ok() != Some(columns) {
                            return Err(error(
                                Code::InvalidValue,
                                "cols does not match the table's grid; a row has one cell per grid column",
                            )
                            .with_property("cols")
                            .with_suggestion("omit cols, or add a column first"));
                        }
                    },
                    (RowSetter::Cell(column), Some(Scalar::Text(text))) => {
                        let column = usize::from(column);
                        if column >= columns {
                            return Err(error(
                                Code::OutOfRange,
                                "the seed text names a column past the table's grid",
                            )
                            .with_property(entry.spec.name));
                        }
                        staged = staged.saturating_add(text.len());
                        bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                        validate_text(text).map_err(|e| e.with_property(entry.spec.name))?;
                        if texts.len() <= column {
                            texts.resize(column + 1, String::new());
                        }
                        text.clone_into(&mut texts[column]);
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let commit = self.change_rows(
            &table,
            Change::Insert {
                position,
                height,
                texts,
            },
        )?;
        let text = format!("{parent_path}/tr[{}]", position + 1);
        let path = Path::parse(&text)
            .map(Path::into_owned)
            .map_err(|_| error(Code::InternalError, "the row path does not parse"))?;
        Ok((commit, path))
    }

    /// A `set` on a row path; `None` for other targets.
    pub(super) fn try_set_row(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.row_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut height = None;
        for (key, value) in props.iter() {
            let (_, entry) = ROW_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            entry.prepare(Operation::Set, value, |setter, value| {
                if let (RowSetter::Height, Some(Scalar::Length(length))) = (setter, value) {
                    height = Some(length.emu());
                }
                Ok(())
            })?;
        }
        let Some(height) = height else {
            return Err(error(Code::MissingProperty, "nothing to set").with_property("height"));
        };
        self.change_rows(
            &item.table,
            Change::Height {
                row: item.row as usize,
                height,
            },
        )
        .map(Some)
    }

    /// A `remove` of a row path; `None` for other targets.
    pub(super) fn try_remove_row(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.row_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        self.change_rows(
            &item.table,
            Change::Remove {
                row: item.row as usize,
            },
        )
        .map(Some)
    }

    /// The row a path target names.
    fn row_target(&self, target: &TargetArg<'_>) -> Result<Option<TableItem>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            step.name().eq_ignore_ascii_case("tr") || step.name().eq_ignore_ascii_case("row")
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        Ok(match node.item() {
            Item::Row(item) => Some(item.clone()),
            _ => None,
        })
    }
}
