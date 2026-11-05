//! Path resolution and query inspection for the legacy XLS workbook nodes.

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{CellRange, CellRef, Step};

use super::Xls;
use super::children::Children;
use super::node::{CellSource, Item, Node, Sheet, display_value, raw_value};
use super::owner::{Grid, SheetHandle};
use super::spec::{Kind, SPEC};
use crate::error::{Code, Error, Result};
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep, Value};
use crate::path::arg::{Origin, PathArg};
use crate::path::fence::Scope;
use crate::path::navigate::{Available, ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::resolve;
use crate::query::inspect::{HeaderBinding, HeaderKey, Inspect};

#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Xls,
    coordinates_only: bool,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Xls) -> Self {
        Self {
            file,
            coordinates_only: false,
        }
    }

    /// Resolve direct structural coordinates without decoding their payload.
    /// The native edit inventory supplies cell facts for a coalesced run; a
    /// row or column named by its number carries no grid (record 0310: the
    /// layout writer reads only its sheet and index, and building the grid
    /// read every cell of the worksheet per call on a positional owner).
    pub(super) const fn coordinates(file: &'f Xls) -> Self {
        Self {
            file,
            coordinates_only: true,
        }
    }

    fn found(node: Node<'f>) -> Lookup<Node<'f>, GapSpec> {
        Lookup::Found(node)
    }

    fn sheet(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let handle = match select {
            Select::Name { name, case } => SheetHandle::named(self.file, name, case)?,
            Select::Ordinal(position) => SheetHandle::at(self.file, position.get())?,
            Select::None => SheetHandle::at(self.file, 0)?,
            Select::Last => tab_count(self.file)?
                .checked_sub(1)
                .map(|position| SheetHandle::at(self.file, position))
                .transpose()?
                .flatten(),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(handle.map_or(Lookup::Missing, |handle| {
            Self::found(Node::new_sheet(
                self.file,
                Sheet::new(handle.position(), handle),
            ))
        }))
    }

    /// A defined name: the Nth in workbook order, or by unique name.
    fn named_range(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let names = super::node::defined_names(self.file)?;
        let index = match select {
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => names.len().checked_sub(1),
            Select::Name { name, case } => {
                let mut found = None;
                for (index, held) in names.iter().enumerate() {
                    if case.matches(&held.name, name) {
                        if found.is_some() {
                            return Err(Error::new(
                                Code::InvalidPath,
                                "the defined-name selector matches more than one name",
                            )
                            .with_format(FileFormat::Xls));
                        }
                        found = Some(index);
                    }
                }
                found
            },
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match index.filter(|index| *index < names.len()) {
            Some(index) => Self::found(Node::new_named_range(self.file, index)),
            None => Lookup::Missing,
        })
    }

    fn in_sheet(
        self,
        sheet: &Sheet<'f>,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let position = sheet.position;
        if super::features::is_feature(kind) {
            return super::features::lookup(self.file, sheet, kind, select);
        }
        if !sheet.handle.is_grid()
            && matches!(kind, Kind::Cell | Kind::Range | Kind::Row | Kind::Column)
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the selected workbook tab has no worksheet grid",
            )
            .with_format(FileFormat::Xls));
        }
        let within_biff = match (kind, select) {
            (Kind::Cell, Select::Cell(at)) => at.row() < 65_536 && at.column() < 256,
            (Kind::Range, Select::Range(range)) => [range.start(), range.end()]
                .iter()
                .all(|at| at.row() < 65_536 && at.column() < 256),
            (Kind::Row, Select::Number(number)) => (1..=65_536).contains(&number),
            (Kind::Column, Select::Number(number)) => (1..=256).contains(&number),
            _ => true,
        };
        if !within_biff {
            return Err(Error::new(
                Code::InvalidPath,
                "XLS coordinates must lie within A1:IV65536",
            )
            .with_format(FileFormat::Xls));
        }
        if let (Kind::Cell, Select::Cell(at)) = (kind, select) {
            let at = (at.row(), at.column());
            return Ok(Self::found(Node::new_cell(
                self.file,
                position,
                at,
                if self.coordinates_only {
                    CellSource::Stub
                } else {
                    CellSource::Direct(sheet.handle.cell(at)?)
                },
            )));
        }
        if self.coordinates_only {
            match (kind, select) {
                (Kind::Row, Select::Number(number)) => {
                    let index = index(number)?;
                    return Ok(Self::found(Node::new_row(
                        self.file,
                        position,
                        Grid::coordinates_only(),
                        index,
                    )));
                },
                (Kind::Column, Select::Number(number)) => {
                    let index = index(number)?;
                    return Ok(Self::found(Node::new_column(
                        self.file,
                        position,
                        Grid::coordinates_only(),
                        index,
                    )));
                },
                _ => {},
            }
        }
        let grid = sheet.grid()?.clone();
        Ok(match (kind, select) {
            (Kind::Range, Select::Range(range)) => {
                Self::found(Node::new_range(self.file, position, grid, rect(range)?))
            },
            (Kind::Row, Select::Number(number)) => {
                let index = index(number)?;
                Self::found(Node::new_row(self.file, position, grid, index))
            },
            (Kind::Row, Select::Last) => grid
                .row_indices()
                .into_iter()
                .max()
                .map(|index| Self::found(Node::new_row(self.file, position, grid, index)))
                .unwrap_or(Lookup::Missing),
            (Kind::Row, Select::None) => grid
                .row_indices()
                .into_iter()
                .next()
                .map(|index| Self::found(Node::new_row(self.file, position, grid, index)))
                .unwrap_or(Lookup::Missing),
            (Kind::Column, Select::Number(number)) => {
                let index = index(number)?;
                Self::found(Node::new_column(self.file, position, grid, index))
            },
            (Kind::Column, Select::Last) => grid
                .column_ranges()
                .into_iter()
                .map(|(_, last)| last)
                .max()
                .map(|index| Self::found(Node::new_column(self.file, position, grid, index)))
                .unwrap_or(Lookup::Missing),
            (Kind::Column, Select::None) => grid
                .column_ranges()
                .into_iter()
                .map(|(first, _)| first)
                .next()
                .map(|index| Self::found(Node::new_column(self.file, position, grid, index)))
                .unwrap_or(Lookup::Missing),
            _ => Lookup::NotApplicable,
        })
    }
}

/// The sheet step of a handle's last cell `get` and the worksheet it
/// resolved to (record 0315, record 0310 G8).
///
/// A text-path `get` of a cell (`/sheet[1]/A7`, `/Data/B2`) resolves its
/// sheet step through the vocabulary and the worksheet catalog on every
/// call: most of its cost over the native cell read. The step's worksheet
/// is a fact of the handle's immutable snapshot, so the handle keeps the
/// last one; the next `get` whose sheet step is the same step (equal name
/// and predicates) starts from that worksheet. A write publishes a new
/// state with an empty cache ([`super::XlsState`]). The cache is a
/// convenience only: when it is busy, empty or holds another step, `get`
/// takes the general resolver, and every result and error is the
/// resolver's.
#[derive(Default)]
pub(crate) struct SheetStep(std::sync::Mutex<Option<(Step<'static>, usize)>>);

impl SheetStep {
    /// The catalog position `step` resolved to, when it is the cached step.
    fn position(&self, step: &Step<'_>) -> Option<usize> {
        let held = self.0.try_lock().ok()?;
        match &*held {
            Some((cached, position)) if cached == step => Some(*position),
            _ => None,
        }
    }

    /// Caches `step` as resolving to the worksheet at `position`.
    fn remember(&self, step: &Step<'_>, position: usize) {
        let Ok(mut held) = self.0.try_lock() else {
            return;
        };
        if !matches!(&*held, Some((cached, at)) if cached == step && *at == position) {
            *held = Some((step.clone().into_owned(), position));
        }
    }
}

/// `File::get` of an XLS path: the general resolver, except that a cell
/// path whose sheet step the handle cached (record 0315) starts from the
/// cached worksheet. The fast route makes the resolver's reads in the
/// resolver's order inside the same fence — the root, the worksheet by its
/// catalog position, the cell — and closes with the same source check; on
/// any other outcome it takes the resolver, so errors are the resolver's.
pub(crate) fn get<'f>(file: &'f Xls, path: &PathArg<'_>) -> Result<Node<'f>> {
    let nav = Nav::new(file);
    let cell = cell_path(path);
    if let Some((step, at)) = cell
        && let Some(position) = file.sheet_step().position(step)
        && let Some(node) = cached_get(nav, path, position, at)
    {
        return Ok(node);
    }
    let node = resolve::get(&nav, path)?;
    if let Some((step, _)) = cell
        && matches!(node.item(), Item::Cell(_))
    {
        file.sheet_step().remember(step, node.sheet_position());
    }
    Ok(node)
}

/// The sheet step and the cell of an absolute two-step cell path
/// (`/<sheet step>/<cell>`, typed or text, not raw).
fn cell_path<'a>(path: &'a PathArg<'_>) -> Option<(&'a Step<'a>, CellRef)> {
    if !path.is_absolute() || path.raw || !matches!(path.origin(), Origin::Typed | Origin::Text(_))
    {
        return None;
    }
    let [sheet, cell] = path.path().steps() else {
        return None;
    };
    if !cell.predicates().is_empty() || cell.name().contains(':') {
        return None;
    }
    CellRange::parse(cell.name()).map(|range| (sheet, range.start()))
}

/// The cell at `at` of the worksheet at catalog `position`, resolved as
/// [`resolve::get`] would: path admission, then the root, the worksheet and
/// the cell inside one fence, then the closing source check. `None` for
/// anything but a found cell and a current source.
fn cached_get<'f>(
    nav: Nav<'f>,
    path: &PathArg<'_>,
    position: usize,
    at: CellRef,
) -> Option<Node<'f>> {
    crate::path::arg::validate(path).ok()?;
    let found = {
        let _scope = Scope::enter(nav.fence_key());
        cached_steps(nav, position, at)
    };
    nav.check_source().ok()?;
    found.ok().flatten()
}

fn cached_steps<'f>(nav: Nav<'f>, position: usize, at: CellRef) -> Result<Option<Node<'f>>> {
    let root = nav.root()?;
    let sheet = Resolved {
        kind: Kind::Sheet,
        select: Select::Ordinal(Position::new(position)),
    };
    let Lookup::Found(sheet) = nav.child(&root, &sheet)? else {
        return Ok(None);
    };
    let cell = Resolved {
        kind: Kind::Cell,
        select: Select::Cell(at),
    };
    Ok(match nav.child(&sheet, &cell)? {
        Lookup::Found(cell) => Some(cell),
        _ => None,
    })
}

pub(super) fn tab_count(file: &Xls) -> Result<usize> {
    file.owner().workbook().sheet_count().map_err(Into::into)
}

/// Run one navigation step.
///
/// The generic resolver fences a whole path resolution and the query walker
/// fences every advance with [`Navigate::check_source`], and every
/// `Navigate`/`Inspect` call below runs inside one of those fences.  A step
/// therefore adds no source-version probe of its own: on a positional owner
/// each probe is a file-metadata system call.
pub(super) fn finish<T>(_file: &Xls, operation: impl FnOnce() -> Result<T>) -> Result<T> {
    operation()
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Xls;
    const ABSENT: &'static [&'static str] = &[
        "table",
        "chart",
        "series",
        "axis",
        "pivottable",
        "slicer",
        "sparkline",
        "mergeCell",
        "ole",
        "run",
    ];

    fn check_source(&self) -> Result<()> {
        self.file.check_source()
    }

    fn fence_key(&self) -> usize {
        self.file.fence_key()
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
            if let Some(gap) = SPEC.elements[usize::from(step.kind.ordinal())].gap {
                return Ok(Lookup::Gap(gap));
            }
            match (parent.item(), step.kind) {
                (Item::Workbook, Kind::Sheet) => self.sheet(step.select),
                (Item::Workbook, Kind::NamedRange) => self.named_range(step.select),
                (Item::Sheet(sheet), kind) => self.in_sheet(sheet, kind, step.select),
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
                    out.kind("namedrange", super::node::defined_names(self.file)?.len());
                },
                (Item::Sheet(sheet), Kind::Row) => {
                    out.kind("row", sheet.grid()?.row_indices().len())
                },
                (Item::Sheet(sheet), Kind::Column) => {
                    let count = sheet
                        .grid()?
                        .column_ranges()
                        .into_iter()
                        .map(|(first, last)| {
                            usize::try_from(last - first + 1).unwrap_or(usize::MAX)
                        })
                        .sum();
                    out.kind("col", count);
                },
                (Item::Sheet(sheet), kind) if super::features::is_feature(kind) => {
                    out.kind(
                        super::features::step_name(kind),
                        super::features::list(self.file, sheet, kind)?.len(),
                    );
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
                Item::Row(row) => Some(Cow::Owned(joined_text(&row.grid, row_rect(row.index)))),
                Item::Column(column) => Some(Cow::Owned(joined_text(
                    &column.grid,
                    column_rect(column.index),
                ))),
                Item::Sheet(_) => node.sheet_name()?.map(Cow::Owned),
                Item::NamedRange(_) => node.named_range_text()?.map(Cow::Owned),
                Item::Feature(feature) => feature.text().map(|text| Cow::Owned(text.into_owned())),
                Item::Workbook | Item::Range(_) => None,
            })
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        finish(self.file, || node.props_unfenced())
    }

    fn number(&self, node: &Node<'f>) -> Option<usize> {
        node.number().map(|number| number as usize)
    }

    fn cell(&self, node: &Node<'f>) -> Option<CellRef> {
        node.cell_ref()
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
            None => fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out),
        }
    }

    fn bind_header(&self, _scope: &Node<'f>, _keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        Err(Error::new(
            Code::UnsupportedCapability,
            "header-row predicates are not supported for legacy XLS",
        )
        .with_format(FileFormat::Xls))
    }

    fn names_column(&self, _scope: &Node<'f>, _key: &str) -> Result<bool> {
        Ok(false)
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
            let Some(cell) = item.grid.record((item.index, column)) else {
                return Ok(None);
            };
            Ok(Some(if raw {
                raw_value(cell)
            } else {
                Value::Text(Cow::Owned(display_value(&cell.value).into_owned()))
            }))
        })
    }
}

fn rect(range: CellRange) -> Result<(u32, u32, u32, u32)> {
    let start = range.start();
    let end = range.end();
    Ok((
        start.row().min(end.row()),
        start.column().min(end.column()),
        start.row().max(end.row()),
        start.column().max(end.column()),
    ))
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
            .with_format(FileFormat::Xls)
        })
}

fn row_rect(index: u32) -> (u32, u32, u32, u32) {
    (index, 0, index, u32::MAX)
}

fn column_rect(index: u32) -> (u32, u32, u32, u32) {
    (0, index, u32::MAX, index)
}

fn joined_text(grid: &Grid, rect: (u32, u32, u32, u32)) -> String {
    let mut out = String::new();
    for (index, (_, cell)) in grid.cells_in(rect).enumerate() {
        if index > 0 {
            out.push(' ');
        }
        out.push_str(&display_value(&cell.value));
    }
    out
}
