//! The XLSX adapter of path resolution and queries (record 0010 §4.1,
//! §4.7, §5.3).
//!
//! [`Nav`] is a copyable reference to one [`Xlsx`] handle. Resolution asks it
//! for one child per path step and runs only that step's leaf read:
//!
//! | Step | Leaf read |
//! |---|---|
//! | `/<name>`, `/sheet[N]`, `[last()]` | the catalog: names match with the leaf's own canonical key (NFD plus case folding), so `/dAtA` is `Data` |
//! | `/<sheet>/A1` | the leaf's cell route (`SourceWorksheet::cell`, which streams a part too large to retain) |
//! | `/<sheet>/A1:C5`, `row[N]`, `col[B]` | the worksheet's parsed grid, materialized once per payload by the leaf's cache and pinned by the node |
//! | `/<sheet>/table[N]`, `table[@name=X]`, `/table[@name=X]` | the worksheet's table parts |
//! | `/namedrange[N]`, `namedrange[NAME]`, `[@name=X]` | the defined names retained with the catalog |
//!
//! Every coordinate inside the grid resolves, a missing one included
//! (record 0010 §4.7). Elements Litchi cannot read yet answer their spec
//! gap (`unsupported_capability`).

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRange, CellRef, Position};
use litchi_xlsx::{Address, Rect};

use super::Xlsx;
use super::children::{Children, column_rect, coordinate, range};
use super::header;
use super::node::{
    CellSource, Item, Node, Sheet, SheetHandle, ViewSource, defined_names, display_text, raw_value,
};
use super::spec::{Kind, SPEC};
use crate::Result;
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, Value};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::{HeaderBinding, HeaderKey, Inspect};

/// The XLSX adapter: path resolution and queries over one handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Xlsx,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Xlsx) -> Self {
        Self { file }
    }

    pub(super) fn file(self) -> &'f Xlsx {
        self.file
    }

    fn found(node: Node<'f>) -> Lookup<Node<'f>, GapSpec> {
        Lookup::Found(node)
    }

    /// A sheet of the workbook.
    fn sheet(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let file = self.file;
        let handle = match select {
            // The leaf's canonical key is the authority for sheet names
            // (record 0010 §4.2); the vocabulary's `Caseless` rule defers to
            // it.
            Select::Name { name, .. } => SheetHandle::named(file, name)?,
            Select::Ordinal(position) => SheetHandle::at(file, position.get())?,
            Select::None => SheetHandle::at(file, 0)?,
            Select::Last => match file.0.owner.len().checked_sub(1) {
                Some(position) => SheetHandle::at(file, position)?,
                None => None,
            },
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match handle {
            Some(handle) => Self::found(Node::new_sheet(
                file,
                Sheet::new(file, handle.position(), handle),
            )),
            None => Lookup::Missing,
        })
    }

    /// A defined name.
    fn named_range(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let names = defined_names(self.file);
        let index = match select {
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => names.len().checked_sub(1),
            Select::Name { name, case } => {
                unique_position(names.iter().map(|held| case.matches(&held.name, name)))?
            },
            Select::Identity(identities) => unique_position(names.iter().map(|held| {
                identities
                    .iter()
                    .all(|(identity, value)| identity.case.matches(&held.name, value))
            }))?,
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match index.filter(|index| *index < names.len()) {
            Some(index) => Self::found(Node::new_named_range(self.file, index)),
            None => Lookup::Missing,
        })
    }

    /// A table addressed from the workbook: the Nth table in workbook order,
    /// or the table with a display name, else a name, across worksheets.
    fn workbook_table(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let file = self.file;
        let mut sheets = Vec::new();
        for position in 0..file.0.owner.len() {
            if let Some(handle) = SheetHandle::at(file, position)? {
                if handle.is_grid() {
                    sheets.push((position, handle.tables()?));
                }
            }
        }
        let total: usize = sheets.iter().map(|(_, tables)| tables.len()).sum();
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

    /// An element of a worksheet.
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
        Ok(match (kind, select) {
            (Kind::Cell, Select::Cell(at)) => {
                let at = address(at)?;
                let view = sheet.handle.cell(at)?;
                // The leaf reports a merged follower as covered even when a
                // record is stored there; OfficeCLI, and this adapter's
                // listings, read the stored record.
                let stored_follower = matches!(view, litchi_xlsx::SourceCellView::Covered(_))
                    && sheet.grid()?.record(at)?.is_some();
                let source = if stored_follower {
                    CellSource::Grid(sheet.grid()?.clone())
                } else {
                    CellSource::View(Box::new(ViewSource {
                        handle: sheet.handle.clone(),
                        view,
                    }))
                };
                Self::found(Node::new_cell(file, position, at, source))
            },
            (Kind::Range, Select::Range(cells)) => {
                let rect = rect(cells)?;
                let grid = sheet.grid()?.clone();
                Self::found(Node::new_range(file, position, grid, rect))
            },
            (Kind::Row, Select::Number(number)) => {
                let grid = sheet.grid()?.clone();
                Self::found(Node::new_row(file, position, grid, index(number)?))
            },
            (Kind::Row, Select::Last | Select::None) => {
                let grid = sheet.grid()?;
                let mut rows = grid.rows()?;
                let row = if matches!(select, Select::Last) {
                    rows.next_back()
                } else {
                    rows.next()
                };
                match row.map(|row| row.index().get()) {
                    Some(index) => Self::found(Node::new_row(file, position, grid.clone(), index)),
                    None => Lookup::Missing,
                }
            },
            (Kind::Column, Select::Number(number)) => {
                let grid = sheet.grid()?.clone();
                let index = index(number)?;
                // Rejects columns past the grid as the leaf does.
                column_rect(index)?;
                Self::found(Node::new_column(file, position, grid, index))
            },
            (Kind::Column, Select::Last | Select::None) => {
                let grid = sheet.grid()?;
                let mut columns = grid.columns()?;
                let column = if matches!(select, Select::Last) {
                    columns.next_back()
                } else {
                    columns.next()
                };
                match column.map(|column| column.index().get()) {
                    Some(index) => {
                        Self::found(Node::new_column(file, position, grid.clone(), index))
                    },
                    None => Lookup::Missing,
                }
            },
            (Kind::Chart, select) => {
                // A positional step reads only its own chart (`None` past
                // the last); `last()` counts the sheet's charts first.
                let index = match select {
                    Select::Ordinal(ordinal) => Some(ordinal.get()),
                    Select::None => Some(0),
                    Select::Last => file.sheet_chart_count(position)?.checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                match index {
                    Some(index) => match file.sheet_chart(position, index)? {
                        Some(part) => Self::found(Node::new_chart(file, position, index, part)),
                        None => Lookup::Missing,
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

/// The position of the first table whose display name matches, else of the
/// first whose name matches, as the leaf's `table(name)` looks tables up.
fn table_position<'t>(
    tables: impl Iterator<Item = &'t litchi_xlsx::workbook::worksheet::TablePart> + Clone,
    matches: impl Fn(&Names<'_>) -> bool,
) -> Result<Option<usize>> {
    match unique_position(
        tables
            .clone()
            .map(|part| matches(&Names::display(part.table()))),
    )? {
        Some(position) => Ok(Some(position)),
        None => unique_position(tables.map(|part| matches(&Names::name(part.table())))),
    }
}

fn unique_position(matches: impl Iterator<Item = bool>) -> Result<Option<usize>> {
    let mut found = None;
    for (position, matches) in matches.enumerate() {
        if matches {
            if found.is_some() {
                return Err(crate::Error::new(
                    crate::Code::InvalidPath,
                    "semantic name selects multiple candidates",
                ));
            }
            found = Some(position);
        }
    }
    Ok(found)
}

/// One name of a table, for identity matching.
struct Names<'a> {
    name: &'a str,
}

impl<'a> Names<'a> {
    fn display(table: &'a litchi_xlsx::Table) -> Self {
        Self {
            name: &table.display_name,
        }
    }

    fn name(table: &'a litchi_xlsx::Table) -> Self {
        Self { name: &table.name }
    }
}

fn matches_identities(
    names: &Names<'_>,
    identities: &crate::path::navigate::Identities<'_>,
) -> bool {
    identities
        .iter()
        .all(|(identity, value)| identity.case.matches(names.name, value))
}

/// A resolved A1 cell as a grid coordinate.
fn address(at: CellRef) -> Result<Address> {
    Address::at(at.row(), at.column()).map_err(coordinate)
}

/// A resolved A1 range as a normalized rectangle: corners in either order.
pub(super) fn rect(cells: CellRange) -> Result<Rect> {
    let (start, end) = (cells.start(), cells.end());
    let first = Address::at(start.row().min(end.row()), start.column().min(end.column()))
        .map_err(coordinate)?;
    // Validate the inclusive end before converting to exclusive coordinates.
    // A checked XLSX address is bounded well below u32::MAX.
    let last = Address::at(start.row().max(end.row()), start.column().max(end.column()))
        .map_err(coordinate)?;
    Rect::new(first, last.row().get() + 1, last.column().get() + 1).map_err(range)
}

/// A 1-based row or column number as a zero-based index.
fn index(number: usize) -> Result<u32> {
    number
        .checked_sub(1)
        .and_then(|index| u32::try_from(index).ok())
        .ok_or_else(|| {
            crate::Error::new(
                crate::Code::InvalidPath,
                format!("{number} is not a worksheet row or column number"),
            )
        })
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Xlsx;

    fn check_source(&self) -> Result<()> {
        self.file.0.owner.ensure_current()
    }

    fn fence_key(&self) -> usize {
        self.file.0.owner.fence_key()
    }

    fn root(&self) -> Result<Node<'f>> {
        self.file
            .0
            .owner
            .finish(|| Ok(Node::new_workbook(self.file)))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        self.file.0.owner.finish(|| {
            // The XLSX table, not the shared workbook one: XLSX reads the
            // chart elements the shared table lists as gaps (record 0270).
            if let Some(gap) = SPEC.elements[step.kind as usize].gap {
                return Ok(Lookup::Gap(gap));
            }
            match (parent.item(), step.kind) {
                (Item::Workbook, Kind::Sheet) => self.sheet(step.select),
                (Item::Workbook, Kind::NamedRange) => self.named_range(step.select),
                (Item::Workbook, Kind::Table) => self.workbook_table(step.select),
                (Item::Sheet(sheet), kind) => self.in_sheet(sheet, kind, step.select),
                (Item::Chart(chart), Kind::Series | Kind::Axis) => {
                    chart_child(self.file, parent, chart, step.kind, step.select)
                },
                (Item::Cell(_), Kind::Run) => super::runs::lookup(self.file, parent, step.select),
                _ => Ok(Lookup::NotApplicable),
            }
        })
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        self.file.0.owner.finish(|| Children::walk(parent, kinds))
    }

    fn available(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        self.file.0.owner.finish(|| {
            let file = self.file;
            match (parent.item(), step.kind) {
                (Item::Workbook, Kind::Sheet) => {
                    out.kind("sheet", file.0.owner.len());
                },
                (Item::Workbook, Kind::NamedRange) => {
                    out.kind("namedrange", defined_names(file).len());
                },
                (Item::Sheet(sheet), Kind::Table) => {
                    out.kind("table", sheet.handle.tables()?.len());
                },
                (Item::Sheet(sheet), Kind::Chart) => {
                    out.kind("chart", file.sheet_chart_count(sheet.position)?);
                },
                (Item::Sheet(sheet), kind) if super::features::is_feature(kind) => {
                    out.kind(
                        super::features::step_name(kind),
                        super::features::list(file, sheet, kind)?.len(),
                    );
                },
                (Item::Chart(chart), Kind::Series) => {
                    out.kind("series", chart.summary()?.series.len());
                },
                (Item::Chart(chart), Kind::Axis) => {
                    out.kind("axis", chart.summary()?.axes.len());
                },
                (Item::Sheet(sheet), Kind::Row | Kind::Column) => {
                    if let Some(extent) = sheet.grid()?.extents()?.stored() {
                        let (rows, columns) = (extent.end().0, extent.end().1);
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
        self.file
            .0
            .owner
            .finish(|| Children::walk_early(parent, kinds))
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        self.file.0.owner.finish(|| {
            Ok(match node.item() {
                Item::Cell(_) => node.cell_text()?.map(Cow::Borrowed),
                // OfficeCLI matches a row or column on the text of the cells it
                // spans (`searchText`).
                Item::Row(row) => Some(Cow::Owned(joined_text(
                    &row.grid,
                    super::children::row_rect(row.index)?,
                )?)),
                Item::Column(column) => Some(Cow::Owned(joined_text(
                    &column.grid,
                    column_rect(column.index)?,
                )?)),
                Item::Sheet(_) => node.sheet_name()?.map(Cow::Borrowed),
                Item::Table(table) => Some(Cow::Borrowed(table.part.table().display_name.as_str())),
                Item::NamedRange(index) => defined_names(self.file)
                    .get(*index)
                    .map(|name| Cow::Borrowed(name.name.as_str())),
                Item::Chart(chart) => chart.summary()?.title.as_deref().map(Cow::Borrowed),
                Item::Feature(feature) => feature.text(),
                Item::Run(run) => run.run().map(|run| Cow::Borrowed(run.text())),
                Item::Workbook | Item::Range(_) | Item::Series(_) | Item::Axis(_) => None,
            })
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        self.file.0.owner.finish(|| node.props())
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
            None => fmt::Display::fmt(&crate::node::PathStep::at(kind.spec().name, ordinal), out),
        }
    }

    fn bind_header(&self, scope: &Node<'f>, keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        self.file
            .0
            .owner
            .finish(|| header::bind(*self, scope, keys))
    }

    fn names_column(&self, scope: &Node<'f>, key: &str) -> Result<bool> {
        self.file
            .0
            .owner
            .finish(|| header::names_column(*self, scope, key))
    }

    fn row_cell<'n>(
        &'n self,
        row: &'n Node<'f>,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        self.file.0.owner.finish(|| {
            let Item::Row(item) = row.item() else {
                return Ok(None);
            };
            let Some(column) = column.checked_sub(1) else {
                return Ok(None);
            };
            let at = Address::at(item.index, column).map_err(coordinate)?;
            Ok(item.grid.record(at)?.map(|cell| {
                if raw {
                    raw_value(cell)
                } else {
                    Value::Text(Cow::Borrowed(display_text(cell)))
                }
            }))
        })
    }
}

/// The display text of every record in `rect`, joined with spaces.
fn joined_text(grid: &super::node::Grid, rect: Rect) -> Result<String> {
    let mut out = String::new();
    for (index, (_, cell)) in grid.cells(rect)?.enumerate() {
        if index > 0 {
            out.push(' ');
        }
        out.push_str(display_text(cell));
    }
    Ok(out)
}

/// A worksheet chart's series (by position) or axis (by position or
/// `@role`).
fn chart_child<'f>(
    file: &'f Xlsx,
    parent: &Node<'f>,
    chart: &super::node::ChartItem,
    kind: Kind,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let summary = chart.summary()?;
    let count = if kind == Kind::Series {
        summary.series.len()
    } else {
        summary.axes.len()
    };
    let index = match select {
        Select::Ordinal(ordinal) => Some(ordinal.get()),
        Select::None => Some(0),
        Select::Last => count.checked_sub(1),
        Select::Identity(identities) if kind == Kind::Axis => {
            let mut found = None;
            for (identity, value) in identities.iter() {
                if !identity.key.eq_ignore_ascii_case("role") {
                    return Ok(Lookup::NotApplicable);
                }
                found = crate::format::chart::Role::parse(value)
                    .and_then(|role| summary.axes.iter().position(|axis| axis.role == role));
            }
            found
        },
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index.filter(|index| *index < count) {
        Some(index) => Lookup::Found(Node::new_chart_part(
            file,
            parent.sheet_position(),
            kind,
            chart.with_child(index),
        )),
        None => Lookup::Missing,
    })
}
