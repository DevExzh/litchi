//! The read owner bridge for the XLSB unified adapter.
//!
//! This module is deliberately a small semantic boundary over the two leaf
//! owners.  The catalog is built from the validated leaf catalog and keeps
//! only authored names and semantic tab facts.  Physical relationship IDs and
//! part names stay in the leaf owner.  Worksheet payloads are obtained through
//! [`litchi_xlsb::sheet::Grid`] or, for the source-backed pivot/slicer/timeline
//! boundary, through the explicitly requested indexed cached-value route.

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use caseless::Caseless as _;
use litchi_core::{FileFormat, Position};
use litchi_xlsb::package::styles_table::CellFormat;
use litchi_xlsb::package::table::Table;
use litchi_xlsb::sheet::{Grid, Kind, Selector, State};
use litchi_xlsb::{Cell, SourceBackedWorksheet};
use unicode_normalization::UnicodeNormalization;

use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};

/// One semantic tab in complete workbook order.
///
/// The name is intentionally available only through the semantic accessor;
/// [`Debug`] omits it because authored names are document content.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct Tab {
    name: String,
    position: Position,
    worksheet_position: Option<Position>,
    kind: Kind,
    state: State,
}

impl Tab {
    fn new(
        name: String,
        position: Position,
        worksheet_position: Option<Position>,
        kind: Kind,
        state: State,
    ) -> Self {
        Self {
            name,
            position,
            worksheet_position,
            kind,
            state,
        }
    }

    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    pub(crate) const fn position(&self) -> Position {
        self.position
    }

    pub(crate) const fn worksheet_position(&self) -> Option<Position> {
        self.worksheet_position
    }

    pub(crate) const fn kind(&self) -> Kind {
        self.kind
    }

    pub(crate) const fn state(&self) -> State {
        self.state
    }
}

impl fmt::Debug for Tab {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Tab")
            .field("position", &self.position)
            .field("worksheet_position", &self.worksheet_position)
            .field("kind", &self.kind)
            .field("state", &self.state)
            .finish()
    }
}

/// The complete semantic workbook catalog and its canonical name index.
///
/// The index uses the same NFD/full-case-fold/NFD key as the shared workbook
/// path vocabulary. Caseless collisions are retained as an ambiguous index
/// entry: positional selection remains valid, while a matching name selector
/// is refused instead of silently choosing one authored tab.
#[derive(Clone)]
pub(crate) struct Catalog {
    tabs: Box<[Tab]>,
    /// `None` marks a caseless collision. Positional selection remains valid;
    /// only a name selector for that key is refused as ambiguous.
    names: HashMap<String, Option<usize>>,
}

impl Catalog {
    fn from_tabs(tabs: Vec<Tab>) -> Result<Self> {
        let mut names = HashMap::new();
        names
            .try_reserve(tabs.len())
            .map_err(|_| internal("could not reserve the XLSB semantic tab name index"))?;
        for (position, tab) in tabs.iter().enumerate() {
            let key = caseless_key(tab.name());
            match names.entry(key) {
                std::collections::hash_map::Entry::Vacant(entry) => {
                    entry.insert(Some(position));
                },
                std::collections::hash_map::Entry::Occupied(mut entry) => {
                    entry.insert(None);
                },
            }
        }
        Ok(Self {
            tabs: tabs.into_boxed_slice(),
            names,
        })
    }

    pub(crate) fn len(&self) -> usize {
        self.tabs.len()
    }

    pub(crate) fn get(&self, position: usize) -> Option<&Tab> {
        self.tabs.get(position)
    }

    pub(crate) fn select<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<&Tab>> {
        let position = match selector.into() {
            Selector::Position(position) => position.get(),
            Selector::Name(name) => match self.names.get(&caseless_key(name.as_ref())) {
                None => return Ok(None),
                Some(Some(position)) => *position,
                Some(None) => {
                    return Err(Error::new(
                        Code::InvalidPath,
                        "XLSB tab name is ambiguous under Unicode caseless matching",
                    )
                    .with_format(FileFormat::Xlsb));
                },
            },
            _ => {
                return Err(Error::new(
                    Code::InvalidSelector,
                    "the XLSB selector form is not supported by the semantic catalog",
                )
                .with_format(FileFormat::Xlsb));
            },
        };
        Ok(self.tabs.get(position))
    }
}

impl fmt::Debug for Catalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Catalog")
            .field("tab_count", &self.tabs.len())
            .finish_non_exhaustive()
    }
}

/// A cell format with the records it names, cloned from the owner's
/// stylesheet: the font, fill and border when they are not the workbook's
/// default records, and the number-format code when it is not `General`.
#[derive(Clone)]
pub(crate) struct ResolvedStyle {
    pub(crate) xf: CellFormat,
    pub(crate) font: Option<litchi_xlsb::styles::Font>,
    pub(crate) fill: Option<litchi_xlsb::styles::Fill>,
    pub(crate) border: Option<litchi_xlsb::package::styles::Border>,
    pub(crate) number_format: Option<String>,
}

impl ResolvedStyle {
    fn resolve(xf: CellFormat, styles: &litchi_xlsb::package::styles_table::StylesTable) -> Self {
        let record = |id: u32| usize::try_from(id).ok().filter(|index| *index != 0);
        let font = record(xf.font_id).and_then(|index| styles.get_font(index).cloned());
        let fill = record(xf.fill_id).and_then(|index| styles.get_fill(index).cloned());
        let border = record(xf.border_id).and_then(|index| styles.get_border(index).cloned());
        let number_format = (xf.num_fmt_id != 0)
            .then(|| styles.get_num_fmt(xf.num_fmt_id).map(str::to_owned))
            .flatten();
        Self {
            xf,
            font,
            fill,
            border,
            number_format,
        }
    }
}

/// A selected tab.  Eager handles retain the cheap facade clone and semantic
/// facts; source handles additionally retain the leaf's source-bound tab
/// handle.  Neither variant parses a worksheet during selection.
#[derive(Clone)]
pub(crate) enum SheetHandle {
    Eager {
        file: Xlsb,
        tab: Tab,
    },
    Source {
        tab: Tab,
        sheet: SourceBackedWorksheet,
    },
}

/// A selected worksheet grid and whether it contains only scalar cached
/// values from the source binary index. The latter is a semantic boundary:
/// formula, style, and worksheet-adjunct properties must be refused by the
/// node layer when this flag is set.
#[derive(Clone)]
pub(crate) struct GridRead {
    pub(crate) leaf: Grid,
    pub(crate) cached_only: bool,
}

impl SheetHandle {
    /// Select a complete-catalog position.  A missing position is `Ok(None)`;
    /// a selected non-worksheet is retained and refused only by grid reads.
    pub(crate) fn at(file: &Xlsb, position: usize) -> Result<Option<Self>> {
        let catalog = file.catalog()?;
        let Some(tab) = catalog.get(position).cloned() else {
            return Ok(None);
        };
        Ok(match &file.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .sheet_by_index(position)?
                .map(|sheet| Self::Source { tab, sheet }),
            XlsbOwner::Eager(_) => Some(Self::Eager {
                file: file.clone(),
                tab,
            }),
        })
    }

    /// Select a complete-catalog name using the owner-wide Unicode caseless
    /// key index.  This does not defer to the XLSB leaf's legacy exact match.
    pub(crate) fn named(file: &Xlsb, name: &str) -> Result<Option<Self>> {
        let catalog = file.catalog()?;
        let Some(tab) = catalog.select(name)? else {
            return Ok(None);
        };
        Self::at(file, tab.position().get())
    }

    pub(crate) fn position(&self) -> Position {
        self.tab().position()
    }

    pub(crate) fn name(&self) -> &str {
        self.tab().name()
    }

    pub(crate) fn kind(&self) -> Kind {
        self.tab().kind()
    }

    pub(crate) fn state(&self) -> State {
        self.tab().state()
    }

    pub(crate) fn worksheet_position(&self) -> Option<Position> {
        self.tab().worksheet_position()
    }

    pub(crate) fn is_grid(&self) -> bool {
        self.kind().is_worksheet()
    }

    /// Materialize the selected worksheet through the leaf's weighted grid.
    /// Source-backed unsupported materialization falls back only to the
    /// explicit scalar cached-value projection; malformed, stale, and limit
    /// failures are returned unchanged.
    pub(crate) fn grid_read(&self) -> Result<GridRead> {
        if !self.is_grid() {
            return Err(unsupported_grid());
        }
        match self {
            Self::Eager { file, tab } => match &file.0.owner {
                XlsbOwner::Eager(package) => super::projection(package)?
                    .grid(Selector::Position(tab.position()))?
                    .map(|leaf| GridRead {
                        leaf,
                        cached_only: false,
                    })
                    .ok_or_else(unsupported_grid),
                XlsbOwner::Source(_) => Err(internal("XLSB eager handle lost its owner state")),
            },
            Self::Source { sheet, .. } => match sheet.grid() {
                Ok(leaf) => Ok(GridRead {
                    leaf,
                    cached_only: false,
                }),
                Err(error) => {
                    let error: Error = error.into();
                    if error.code() != Code::UnsupportedCapability {
                        return Err(error);
                    }
                    sheet
                        .cached_grid()
                        .map(|leaf| GridRead {
                            leaf,
                            cached_only: true,
                        })
                        .map_err(Into::into)
                },
            },
        }
    }

    /// Return this worksheet's structured tables.  The leaf remains the sole
    /// parser and owns any source-side semantic cache.
    pub(crate) fn tables(&self) -> Result<Vec<Table>> {
        let worksheet_position = self
            .worksheet_position()
            .ok_or_else(unsupported_tables)?
            .get();
        match self {
            Self::Eager { file, .. } => match &file.0.owner {
                XlsbOwner::Eager(package) => {
                    let mut tables = Vec::new();
                    for (index, table) in super::projection(package)?.structured_tables() {
                        if *index == worksheet_position {
                            tables.push(table.clone());
                        }
                    }
                    Ok(tables)
                },
                XlsbOwner::Source(_) => Err(internal("XLSB eager handle lost its owner state")),
            },
            Self::Source { sheet, .. } => sheet
                .tables()
                .map(|tables| tables.iter().cloned().collect())
                .map_err(Into::into),
        }
    }

    /// Resolve a materialized cell's style, with its font, fill, border and
    /// number-format records, through the owner-scoped leaf stylesheet.
    /// Style parsing remains lazy on source owners.
    pub(crate) fn style_for_cell(&self, cell: &Cell) -> Result<Option<ResolvedStyle>> {
        match self {
            Self::Eager { file, .. } => match &file.0.owner {
                XlsbOwner::Eager(package) => {
                    let workbook = super::projection(package)?;
                    Ok(workbook
                        .style_for_cell(cell)
                        .map(|xf| ResolvedStyle::resolve(xf.clone(), workbook.styles())))
                },
                XlsbOwner::Source(_) => Err(internal("XLSB eager handle lost its owner state")),
            },
            Self::Source { sheet, .. } => {
                let Some(xf) = sheet.style_for_cell(cell)? else {
                    return Ok(None);
                };
                let styles = sheet.styles()?;
                Ok(Some(ResolvedStyle::resolve(xf, &styles)))
            },
        }
    }

    fn tab(&self) -> &Tab {
        match self {
            Self::Eager { tab, .. } | Self::Source { tab, .. } => tab,
        }
    }
}

impl fmt::Debug for SheetHandle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SheetHandle")
            .field("position", &self.position())
            .field("kind", &self.kind())
            .field("state", &self.state())
            .finish_non_exhaustive()
    }
}

impl XlsbOwner {
    /// Bracket a read with source and execution freshness fences.  A source
    /// change after the operation wins over the operation's own result.
    /// Inside a resolver or query step for this owner's handle, whose bracket
    /// already checks after the whole step, `read` runs alone (record 0262,
    /// as `nav::finish` and `XlsxOwner::finish` do; record 0310: node property
    /// and text reads paid two extra checks per visited cell).
    pub(crate) fn finish<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        if crate::path::fence::is_active(std::ptr::from_ref(self).addr()) {
            return read();
        }
        self.ensure_current()?;
        let result = read();
        self.ensure_current()?;
        result
    }

    /// Build the complete semantic catalog from the validated leaf catalog.
    /// This method is separate so `XlsbState` can retain one `OnceLock<Arc<_>>`
    /// without moving catalog policy into the node layer.
    pub(crate) fn build_catalog(&self) -> Result<Catalog> {
        match self {
            Self::Eager(package) => {
                let workbook = super::projection(package)?;
                let mut tabs = Vec::new();
                tabs.try_reserve_exact(workbook.sheets().len())
                    .map_err(|_| internal("could not reserve the XLSB semantic tab catalog"))?;
                for tab in workbook.sheets() {
                    tabs.push(Tab::new(
                        tab.name().to_owned(),
                        tab.position(),
                        tab.worksheet_position(),
                        tab.kind(),
                        tab.state(),
                    ));
                }
                Catalog::from_tabs(tabs)
            },
            Self::Source(workbook) => {
                let sheets = workbook.workbook().sheets()?;
                let mut tabs = Vec::new();
                tabs.try_reserve_exact(sheets.len())
                    .map_err(|_| internal("could not reserve the XLSB semantic tab catalog"))?;
                let mut worksheet_position = 0usize;
                for sheet in sheets {
                    let kind = sheet.kind()?;
                    let worksheet = kind.is_worksheet().then(|| {
                        let position = Position::new(worksheet_position);
                        worksheet_position += 1;
                        position
                    });
                    tabs.push(Tab::new(
                        sheet.name()?.to_owned(),
                        Position::new(sheet.workbook_position()?),
                        worksheet,
                        kind,
                        sheet.state()?,
                    ));
                }
                Catalog::from_tabs(tabs)
            },
        }
    }
}

impl Xlsb {
    /// Return the complete semantic catalog. The state owns one lazy Arc so
    /// repeated selectors do not rebuild the tab list or its Unicode key
    /// index. Cached reads still run source freshness fences.
    pub(crate) fn catalog(&self) -> Result<Arc<Catalog>> {
        self.0.owner.ensure_current()?;
        let catalog = if let Some(catalog) = self.0.catalog.get() {
            Arc::clone(catalog)
        } else {
            let candidate = Arc::new(self.0.owner.finish(|| self.0.owner.build_catalog())?);
            Arc::clone(self.0.catalog.get_or_init(|| candidate))
        };
        self.0.owner.ensure_current()?;
        Ok(catalog)
    }

    pub(crate) fn tab_count(&self) -> Result<usize> {
        self.catalog().map(|catalog| catalog.len())
    }

    pub(crate) fn active_catalog_position(&self) -> Result<Option<Position>> {
        self.0.owner.finish(|| match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .active_catalog_position()
                .map(|position| position.map(Position::new))
                .map_err(Into::into),
            XlsbOwner::Eager(package) => Ok(super::projection(package)?
                .active_catalog_position()
                .map(Position::new)),
        })
    }

    /// The workbook's defined names with their scope and formula text, in
    /// record order (the order of [`Self::defined_names`]).
    pub(crate) fn defined_name_entries(&self) -> Result<Vec<litchi_xlsb::named_ranges::Entry>> {
        self.0.owner.finish(|| match &self.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .defined_name_entries()
                .map_err(Into::into),
            XlsbOwner::Eager(package) => super::projection(package)?
                .defined_name_entries()
                .map_err(Into::into),
        })
    }

    /// Formula-context defined names retained by the workbook catalog.
    pub(crate) fn defined_names(&self) -> Result<Arc<[String]>> {
        self.0.owner.ensure_current()?;
        let names = if let Some(names) = self.0.defined_names.get() {
            Arc::clone(names)
        } else {
            let candidate = self.0.owner.finish(|| match &self.0.owner {
                XlsbOwner::Source(workbook) => {
                    workbook.workbook().defined_names().map_err(Into::into)
                },
                XlsbOwner::Eager(package) => {
                    Ok(super::projection(package)?.defined_names().to_vec().into())
                },
            })?;
            Arc::clone(self.0.defined_names.get_or_init(|| candidate))
        };
        self.0.owner.ensure_current()?;
        Ok(names)
    }

    pub(crate) fn finish<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        self.0.owner.finish(read)
    }
}

fn caseless_key(value: &str) -> String {
    value.nfd().default_case_fold().nfd().collect()
}

fn internal(message: &'static str) -> Error {
    Error::new(Code::InternalError, message).with_format(FileFormat::Xlsb)
}

fn unsupported_grid() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "the selected XLSB tab does not expose a worksheet grid",
    )
    .with_format(FileFormat::Xlsb)
}

fn unsupported_tables() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "the selected XLSB tab does not own worksheet tables",
    )
    .with_format(FileFormat::Xlsb)
}
