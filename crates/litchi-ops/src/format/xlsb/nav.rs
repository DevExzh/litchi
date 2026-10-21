//! XLSB path resolution and query inspection (record 0010 §4.1, §4.7, §5.3).
//!
//! Workbook tabs are resolved in complete catalog order. Worksheet payload
//! paths then use the retained sparse grid bridge, whose coordinates are
//! checked before any leaf lookup. Source-backed reads are fenced before and
//! after every public navigation, inspection, and child-iterator operation.

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRange, CellRef, Position};
use litchi_xlsb::sheet::Rect;

use super::children::{Children, column_rect, coordinate, range};
use super::header;
use super::node::{CellSource, Item, Node, Sheet, defined_names, display_text, raw_value};
use super::owner::SheetHandle;
use super::spec::{Kind, SPEC};
use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, Value};
use crate::path::navigate::{Available, ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::{HeaderBinding, HeaderKey, Inspect};

/// The XLSB adapter: path resolution and query inspection over one handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Xlsb,
    /// Direct mutation paths only need the checked coordinate and worksheet
    /// position.  Keep the normal navigation path read-capable while letting
    /// the write resolver defer the sparse-grid parse until a node is
    /// actually inspected.
    cell_target: bool,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Xlsb) -> Self {
        Self {
            file,
            cell_target: false,
        }
    }

    pub(crate) const fn for_cell_target(file: &'f Xlsb) -> Self {
        Self {
            file,
            cell_target: true,
        }
    }

    pub(super) fn file(self) -> &'f Xlsb {
        self.file
    }

    fn found(node: Node<'f>) -> Lookup<Node<'f>, GapSpec> {
        Lookup::Found(node)
    }

    /// Resolve a complete-catalog sheet selector. Unlike the legacy
    /// worksheet ordinal, this includes chart, dialog, macro and unbound tabs.
    fn sheet(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let file = self.file;
        let handle = match select {
            Select::Name { name, .. } => SheetHandle::named(file, name)?,
            Select::Ordinal(position) => SheetHandle::at(file, position.get())?,
            Select::None => SheetHandle::at(file, 0)?,
            Select::Last => tab_count(file)?
                .checked_sub(1)
                .map(|position| SheetHandle::at(file, position))
                .transpose()?
                .flatten(),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match handle {
            Some(handle) => {
                Self::found(Node::new_sheet(file, Sheet::new(handle.position(), handle)))
            },
            None => Lookup::Missing,
        })
    }

    fn named_range(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let names = defined_names(self.file)?;
        let index = match select {
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => names.len().checked_sub(1),
            Select::Name { name, case } => {
                unique_position(names.iter().map(|held| case.matches(held.as_str(), name)))?
            },
            Select::Identity(identities) => unique_position(names.iter().map(|held| {
                identities
                    .iter()
                    .all(|(identity, value)| identity.case.matches(held.as_str(), value))
            }))?,
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match index.filter(|index| *index < names.len()) {
            Some(index) => Self::found(Node::new_named_range(self.file, index)),
            None => Lookup::Missing,
        })
    }

    /// Resolve a workbook-scope table in workbook tab order.
    fn workbook_table(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let file = self.file;
        let mut sheets = Vec::new();
        for position in 0..tab_count(file)? {
            if let Some(handle) = SheetHandle::at(file, position)? {
                if handle.is_grid() {
                    sheets.push((position, handle.tables()?));
                }
            }
        }
        let total = sheets.iter().map(|(_, tables)| tables.len()).sum::<usize>();
        let target = match select {
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => total.checked_sub(1),
            Select::Identity(identities) => {
                let flat = sheets.iter().flat_map(|(_, tables)| tables.iter());
                table_position(flat, |table| matches_identities(table, &identities))?
            },
            _ => return Ok(Lookup::NotApplicable),
        };
        let Some(mut target) = target else {
            return Ok(Lookup::Missing);
        };
        for (position, tables) in sheets {
            if target < tables.len() {
                let part = tables.into_iter().nth(target);
                return Ok(match part {
                    Some(part) => Self::found(Node::new_table(file, position, target, part)),
                    None => Lookup::Missing,
                });
            }
            target -= tables.len();
        }
        Ok(Lookup::Missing)
    }

    fn in_sheet(
        self,
        sheet: &Sheet,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let file = self.file;
        let position = sheet.position;
        if super::features::is_feature(kind) {
            return super::features::lookup(file, sheet, kind, select);
        }
        if kind == Kind::Chart {
            return super::chart::lookup(file, sheet, select);
        }
        if !sheet.handle.is_grid()
            && matches!(kind, Kind::Cell | Kind::Range | Kind::Row | Kind::Column)
        {
            return Err(Error::new(
                Code::UnsupportedType,
                "the selected workbook tab has no worksheet grid",
            )
            .with_format(FileFormat::Xlsb));
        }
        Ok(match (kind, select) {
            (Kind::Cell, Select::Cell(at)) => {
                let at = address(at)?;
                if self.cell_target {
                    Self::found(Node::new_cell_target(
                        file,
                        position,
                        at,
                        sheet.handle.clone(),
                    ))
                } else {
                    let grid = sheet.grid()?.clone();
                    Self::found(Node::new_cell(file, position, at, CellSource::Grid(grid)))
                }
            },
            (Kind::Range, Select::Range(cells)) => {
                let rect = rect(cells)?;
                let grid = sheet.grid()?.clone();
                Self::found(Node::new_range(file, position, grid, rect))
            },
            (Kind::Row, Select::Number(number)) => {
                let index = index(number)?;
                super::children::row_rect(index)?;
                let grid = sheet.grid()?.clone();
                Self::found(Node::new_row(file, position, grid, index))
            },
            (Kind::Row, Select::Last) => {
                let grid = sheet.grid()?;
                match last_row(grid)? {
                    Some(index) => Self::found(Node::new_row(file, position, grid.clone(), index)),
                    None => Lookup::Missing,
                }
            },
            (Kind::Row, Select::None) => {
                let grid = sheet.grid()?;
                match first_row(grid)? {
                    Some(index) => Self::found(Node::new_row(file, position, grid.clone(), index)),
                    None => Lookup::Missing,
                }
            },
            (Kind::Column, Select::Number(number)) => {
                let index = index(number)?;
                column_rect(index)?;
                let grid = sheet.grid()?.clone();
                Self::found(Node::new_column(file, position, grid, index))
            },
            (Kind::Column, Select::Last) => {
                let grid = sheet.grid()?;
                match last_column(grid)? {
                    Some(index) => {
                        Self::found(Node::new_column(file, position, grid.clone(), index))
                    },
                    None => Lookup::Missing,
                }
            },
            (Kind::Column, Select::None) => {
                let grid = sheet.grid()?;
                match first_column(grid)? {
                    Some(index) => {
                        Self::found(Node::new_column(file, position, grid.clone(), index))
                    },
                    None => Lookup::Missing,
                }
            },
            (Kind::Table, select) => {
                let tables = sheet.handle.tables()?;
                let index = match select {
                    Select::Ordinal(position) => Some(position.get()),
                    Select::None => Some(0),
                    Select::Last => tables.len().checked_sub(1),
                    Select::Identity(identities) => table_position(tables.iter(), |table| {
                        matches_identities(table, &identities)
                    })?,
                    _ => return Ok(Lookup::NotApplicable),
                };
                match index
                    .and_then(|index| tables.into_iter().nth(index).map(|part| (index, part)))
                {
                    Some((index, part)) => {
                        Self::found(Node::new_table(file, position, index, part))
                    },
                    None => Lookup::Missing,
                }
            },
            _ => Lookup::NotApplicable,
        })
    }
}

/// Count complete workbook tabs, preserving source freshness through the
/// caller's fence.
pub(super) fn tab_count(file: &Xlsb) -> Result<usize> {
    match &file.0.owner {
        XlsbOwner::Source(workbook) => Ok(workbook.workbook().sheet_count()?),
        XlsbOwner::Eager(package) => Ok(super::projection(package)?.sheets().len()),
    }
}

/// Fence a public operation around both its source read and result creation.
pub(super) fn finish<T>(file: &Xlsb, operation: impl FnOnce() -> Result<T>) -> Result<T> {
    // Inside a resolver or query step, whose check after the step covers
    // this read, the read runs alone (record 0262).
    if crate::path::fence::is_active(fence_key(file)) {
        return operation();
    }
    file.0.owner.ensure_current()?;
    let value = operation();
    let freshness = file.0.owner.ensure_current();
    match freshness {
        Err(error) => Err(error),
        Ok(()) => value,
    }
}

/// The handle's key for [`crate::path::fence::Scope`].
fn fence_key(file: &Xlsb) -> usize {
    std::ptr::from_ref(&file.0.owner).addr()
}

fn table_position<'t>(
    tables: impl Iterator<Item = &'t litchi_xlsb::package::table::Table> + Clone,
    matches: impl Fn(&TableNames<'_>) -> bool,
) -> Result<Option<usize>> {
    match unique_position(
        tables
            .clone()
            .map(|part| matches(&TableNames::display(part))),
    )? {
        Some(position) => Ok(Some(position)),
        None => unique_position(tables.map(|part| matches(&TableNames::name(part)))),
    }
}

fn unique_position(matches: impl Iterator<Item = bool>) -> Result<Option<usize>> {
    let mut found = None;
    for (position, matches) in matches.enumerate() {
        if matches {
            if found.is_some() {
                return Err(Error::new(
                    Code::InvalidPath,
                    "semantic name selects multiple candidates",
                ));
            }
            found = Some(position);
        }
    }
    Ok(found)
}

struct TableNames<'a> {
    name: Option<&'a str>,
}

impl<'a> TableNames<'a> {
    fn display(table: &'a litchi_xlsb::package::table::Table) -> Self {
        Self {
            name: table.display_name.as_deref(),
        }
    }

    fn name(table: &'a litchi_xlsb::package::table::Table) -> Self {
        Self {
            name: table.name.as_deref(),
        }
    }
}

fn matches_identities(
    names: &TableNames<'_>,
    identities: &crate::path::navigate::Identities<'_>,
) -> bool {
    identities.iter().all(|(identity, value)| {
        names
            .name
            .is_some_and(|name| identity.case.matches(name, value))
    })
}

fn address(at: CellRef) -> Result<(u32, u32)> {
    litchi_xlsb::sheet::At::from((at.row(), at.column()))
        .resolve()
        .map(|_| (at.row(), at.column()))
        .map_err(coordinate)
}

fn rect(cells: CellRange) -> Result<Rect> {
    let (start, end) = (cells.start(), cells.end());
    Rect::at(
        start.row().min(end.row()),
        start.column().min(end.column()),
        start.row().max(end.row()).checked_add(1).ok_or_else(|| {
            Error::new(Code::OutOfRange, "worksheet range row is outside the grid")
        })?,
        start
            .column()
            .max(end.column())
            .checked_add(1)
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "worksheet range column is outside the grid",
                )
            })?,
    )
    .map_err(range)
}

fn index(number: usize) -> Result<u32> {
    number
        .checked_sub(1)
        .and_then(|index| u32::try_from(index).ok())
        .ok_or_else(|| {
            Error::new(
                Code::InvalidPath,
                "worksheet row or column number must be at least one",
            )
        })
}

fn last_row(grid: &super::node::Grid) -> Result<Option<u32>> {
    Ok(grid.row_indices()?.into_iter().max())
}

fn first_row(grid: &super::node::Grid) -> Result<Option<u32>> {
    Ok(grid.row_indices()?.into_iter().min())
}

fn last_column(grid: &super::node::Grid) -> Result<Option<u32>> {
    Ok(grid
        .column_ranges()?
        .into_iter()
        .map(|(_, last)| last)
        .max())
}

fn first_column(grid: &super::node::Grid) -> Result<Option<u32>> {
    Ok(grid
        .column_ranges()?
        .into_iter()
        .map(|(first, _)| first)
        .min())
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Xlsb;

    fn check_source(&self) -> Result<()> {
        self.file.0.owner.ensure_current()
    }

    fn fence_key(&self) -> usize {
        fence_key(self.file)
    }

    fn root(&self) -> Result<Node<'f>> {
        finish(self.file, || Ok(Node::new_workbook(self.file)))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        finish(self.file, || {
            // The shared workbook spec is the vocabulary baseline; this local
            // spec supplies the XLSB capability gaps and property table.
            if let Some(gap) = SPEC.elements[usize::from(step.kind.ordinal())].gap {
                return Ok(Lookup::Gap(gap));
            }
            match (parent.item(), step.kind) {
                (Item::Workbook, Kind::Sheet) => self.sheet(step.select),
                (Item::Workbook, Kind::NamedRange) => self.named_range(step.select),
                (Item::Workbook, Kind::Table) => self.workbook_table(step.select),
                (Item::Sheet(sheet), kind) => self.in_sheet(sheet, kind, step.select),
                (Item::Chart(chart), Kind::Series | Kind::Axis) => {
                    super::chart::child(self.file, parent, chart, step.kind, step.select)
                },
                _ => Ok(Lookup::NotApplicable),
            }
        })
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        finish(self.file, || Children::walk(parent, kinds))
    }

    fn available(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        finish(self.file, || {
            match (parent.item(), step.kind) {
                (Item::Workbook, Kind::Sheet) => out.kind("sheet", tab_count(self.file)?),
                (Item::Workbook, Kind::NamedRange) => {
                    out.kind("namedrange", defined_names(self.file)?.len())
                },
                (Item::Sheet(sheet), Kind::Table) => {
                    out.kind("table", sheet.handle.tables()?.len())
                },
                (Item::Sheet(sheet), Kind::Chart) => {
                    out.kind("chart", super::chart::list(sheet)?.len());
                },
                (Item::Chart(chart), Kind::Series) => {
                    out.kind("series", chart.summary()?.series.len());
                },
                (Item::Chart(chart), Kind::Axis) => {
                    out.kind("axis", chart.summary()?.axes.len());
                },
                (Item::Sheet(sheet), kind) if super::features::is_feature(kind) => {
                    out.kind(
                        super::features::step_name(kind),
                        super::features::list(self.file, sheet, kind)?.len(),
                    );
                },
                (Item::Sheet(sheet), Kind::Row | Kind::Column) => {
                    if let Some(extent) = sheet.grid()?.extents()? {
                        let (rows, columns) = extent.end();
                        if step.kind == Kind::Row {
                            out.kind("row", rows as usize);
                        } else {
                            out.kind("col", columns as usize);
                        }
                    }
                },
                _ => {},
            }
            Ok(())
        })
    }
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Sheet];

    fn kind(&self, node: &Node<'f>) -> Kind {
        node.element()
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.kind()
    }

    fn early_children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        finish(self.file, || Children::walk_early(parent, kinds))
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        finish(self.file, || {
            Ok(match node.item() {
                Item::Cell(_) => node.cell_text()?.map(|text| Cow::Owned(text.into_owned())),
                Item::Row(row) => Some(Cow::Owned(joined_text(
                    &row.grid,
                    super::children::row_rect(row.index)?,
                )?)),
                Item::Column(column) => Some(Cow::Owned(joined_text(
                    &column.grid,
                    column_rect(column.index)?,
                )?)),
                Item::Sheet(_) => node.sheet_name()?.map(Cow::Owned),
                Item::Table(table) => Some(Cow::Borrowed(table_label(table))),
                Item::NamedRange(index) => defined_names(self.file)?
                    .get(*index)
                    .map(|name| Cow::Owned(name.clone())),
                Item::Feature(feature) => feature.text(),
                Item::Chart(chart) => chart.title()?,
                Item::Workbook | Item::Range(_) | Item::Series(_) | Item::Axis(_) => None,
            })
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        finish(self.file, || node.props())
    }

    fn number(&self, node: &Node<'f>) -> Option<usize> {
        node.number().map(|number| number as usize)
    }

    fn cell(&self, node: &Node<'f>) -> Option<CellRef> {
        node.cell_ref()
    }

    fn cell_at(&self, node: &Node<'f>) -> Option<(u32, u32)> {
        // What `cell_ref` parses back from the A1 label, which every
        // coordinate inside the sheet grid does.
        node.cell()
            .filter(|&(column, row)| column < 16_384 && row < 1_048_576)
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        match node.last_step() {
            Some(step) => fmt::Display::fmt(&step, out),
            None => fmt::Display::fmt(&crate::node::PathStep::at(kind.name(), ordinal), out),
        }
    }

    fn bind_header(&self, scope: &Node<'f>, keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        finish(self.file, || header::bind(*self, scope, keys))
    }

    fn names_column(&self, scope: &Node<'f>, key: &str) -> Result<bool> {
        finish(self.file, || header::names_column(*self, scope, key))
    }

    fn row_cell<'n>(
        &'n self,
        row: &'n Node<'f>,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        finish(self.file, || {
            let Item::Row(item) = row.item() else {
                return Ok(None);
            };
            let Some(column) = column.checked_sub(1) else {
                return Ok(None);
            };
            let Some(cell) = item.grid.record((item.index, column))? else {
                return Ok(None);
            };
            Ok(Some(if raw {
                raw_value(cell)
            } else {
                Value::Text(Cow::Owned(display_text(cell).into_owned()))
            }))
        })
    }
}

fn table_label(table: &super::node::TableItem) -> &str {
    table
        .part
        .display_name
        .as_deref()
        .or(table.part.name.as_deref())
        .unwrap_or("")
}

fn joined_text(grid: &super::node::Grid, rect: Rect) -> Result<String> {
    let mut out = String::new();
    for (index, (_, cell)) in grid.cells(rect)?.enumerate() {
        if index > 0 {
            out.push(' ');
        }
        out.push_str(&display_text(cell));
    }
    Ok(out)
}
