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
//! `add(/slide[N], table, {rows, cols, data, x, y, width, height, name})`
//! inserts a table (OfficeCLI's defaults: three by three, `data` rows split
//! by `;` and cells by `,`, padded to the widest row).
//!
//! `set(/slide[N]/table[M], {name, firstRow, lastRow, firstCol, lastCol,
//! bandedRows, bandedCols})` renames the frame and changes the table's
//! style switches (record 0342).
//!
//! Source owners are promoted to an eager package, as column writes are.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::Path;

use super::node::{Item, TableItem, TreeItem};
use super::spec::{
    Kind, ROW_WRITE_TABLE, RowSetter, TABLE_SWITCHES, TABLE_WRITE_TABLE, TableKey, TableSetter,
};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options, validate_text};
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
    fn change_rows(
        &self,
        table: &TreeItem,
        change: Change,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let slide = slide_table(table)?;
        let shape = table.handle.position().get();
        let source = if matches!(&change, Change::Height { .. }) {
            self.editable_owner_for_write()?
        } else {
            self.editable_owner()?
        };
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
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
        self.commit_native(source, native)
    }

    /// Publish a native table edit over `source`.
    fn commit_native(&self, source: Editable, native: litchi_pptx::opened::Edit) -> Result<Commit> {
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

    /// A table-level `set` (record 0342): the frame's `name` and the
    /// table's style switches. `None` for other targets.
    pub(super) fn try_set_table(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path
            .path()
            .steps()
            .last()
            .is_some_and(|step| step.name().eq_ignore_ascii_case("table"))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let table = match node.item() {
            Item::Tree(tree) if tree.kind == Kind::Table => tree.clone(),
            _ => return Ok(None),
        };
        drop(node);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut name = None;
        let mut switches = [None::<bool>; TABLE_SWITCHES.len()];
        let mut seen = 0_u32;
        for (key, value) in props.iter() {
            let (index, entry) = TABLE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if seen & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a table property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            seen |= bit;
            entry.prepare(Operation::Set, value, |setter, value| {
                match (setter, value) {
                    (TableSetter::Name, Some(Scalar::Text(value))) => {
                        validate_text(value).map_err(|e| e.with_property("name"))?;
                        name = Some(value.to_owned());
                    },
                    (TableSetter::Switch(switch), Some(Scalar::Bool(on))) => {
                        switches[usize::from(switch)] = Some(on);
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        if name.is_none() && switches.iter().all(Option::is_none) {
            return Err(error(Code::MissingProperty, "nothing to set"));
        }
        let slide = slide_table(&table)?;
        let shape = table.handle.position().get();
        let source = self.editable_owner_for_write()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        if let Some(name) = &name {
            native.set_shape_name(slide, shape, name)?;
        }
        if switches.iter().any(Option::is_some) {
            native.set_table_properties(slide, shape, |properties| {
                for (key, on) in TABLE_SWITCHES.iter().zip(switches) {
                    let Some(on) = on else {
                        continue;
                    };
                    let slot = match key {
                        TableKey::FirstRow => &mut properties.first_row,
                        TableKey::LastRow => &mut properties.last_row,
                        TableKey::FirstCol => &mut properties.first_col,
                        TableKey::LastCol => &mut properties.last_col,
                        TableKey::BandedRows => &mut properties.band_row,
                        _ => &mut properties.band_col,
                    };
                    *slot = on;
                }
            })?;
        }
        self.commit_native(source, native).map(Some)
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
            limits,
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
            limits,
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
            limits,
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

/// OfficeCLI's table frame defaults (EMU).
const TABLE_X: i64 = 457_200;
const TABLE_Y: i64 = 1_600_200;
const TABLE_WIDTH: i64 = 8_229_600;
const TABLE_ROW_HEIGHT: i64 = 370_840;

impl Pptx {
    /// `add(/slide[N], table, {rows, cols, data, x, y, width, height, name})`
    /// (record 0333).
    pub(super) fn add_table(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = resolved.item() else {
            return Err(error(
                Code::UnsupportedType,
                "tables require a slide parent",
            ));
        };
        let slide = slide.position as usize;
        drop(resolved);
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "table insertion supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut rows = None;
        let mut cols = None;
        let mut data: Option<String> = None;
        let mut name = None;
        let mut frame = [None; 4];
        let mut seen = 0_u32;
        for (key, value) in node.props().iter() {
            let (index, entry) = TABLE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if seen & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a table property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            seen |= bit;
            entry.prepare(Operation::Add, value, |setter, value| {
                match (setter, value) {
                    (TableSetter::Rows, Some(Scalar::Int(value))) => rows = Some(value),
                    (TableSetter::Cols, Some(Scalar::Int(value))) => cols = Some(value),
                    (TableSetter::Data, Some(Scalar::Text(value))) => {
                        bound(
                            LimitKind::StagedBytes,
                            value.len(),
                            limits.max_staged_bytes(),
                        )?;
                        validate_text(value).map_err(|e| e.with_property("data"))?;
                        data = Some(value.to_owned());
                    },
                    (TableSetter::Name, Some(Scalar::Text(value))) => {
                        validate_text(value).map_err(|e| e.with_property("name"))?;
                        name = Some(value.to_owned());
                    },
                    (setter, Some(Scalar::Length(length))) => {
                        let index = match setter {
                            TableSetter::X => 0,
                            TableSetter::Y => 1,
                            TableSetter::Width => 2,
                            _ => 3,
                        };
                        frame[index] = Some(length.emu());
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        // `data` fixes the grid, which `rows`/`cols` may only widen
        // (OfficeCLI); without data the grid is `rows` × `cols`, three by
        // three by default.
        let mut cells: Vec<Vec<String>> = match &data {
            Some(data) => data
                .split(';')
                .map(|row| row.split(',').map(|cell| cell.trim().to_owned()).collect())
                .collect(),
            None => Vec::new(),
        };
        let wanted = |value: Option<i64>| value.and_then(|value| usize::try_from(value).ok());
        let (row_count, col_count) = if data.is_some() {
            let width = cells.iter().map(Vec::len).max().unwrap_or(0);
            (
                cells.len().max(wanted(rows).unwrap_or(0)),
                width.max(wanted(cols).unwrap_or(0)),
            )
        } else {
            (wanted(rows).unwrap_or(3), wanted(cols).unwrap_or(3))
        };
        if row_count == 0 || col_count == 0 {
            return Err(error(
                Code::InvalidValue,
                "a table has at least one row and one column",
            )
            .with_property("data"));
        }
        cells.resize_with(row_count, Vec::new);
        for row in &mut cells {
            row.resize(col_count, String::new());
        }
        let bounds = (
            frame[0].unwrap_or(TABLE_X),
            frame[1].unwrap_or(TABLE_Y),
            frame[2].unwrap_or(TABLE_WIDTH),
            frame[3].unwrap_or(TABLE_ROW_HEIGHT.saturating_mul(row_count as i64)),
        );
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        let ordinal = scene
            .roots()
            .filter(|shape| super::steps::tree_kind(shape) == Some(Kind::Table))
            .count();
        let name = name.unwrap_or_else(|| format!("Table {}", ordinal + 1));
        let source = self.editable_owner()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        edit.add_table(slide, &cells, bounds)?;
        // The new table is the last shape of the slide in pre-order.
        edit.set_shape_name(slide, scene.len(), &name)?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::parse(&format!("/slide[{}]/table[{}]", slide + 1, ordinal + 1))
            .map(Path::into_owned)
            .map_err(|_| error(Code::InternalError, "the table path does not parse"))?;
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }
}
