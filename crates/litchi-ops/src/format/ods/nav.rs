//! Path resolution and query inspection for ODS workbook nodes (record
//! 0277).

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{CellRange, CellRef};

use super::Ods;
use super::children::Children;
use super::node::{Item, Node, annotations, raw_value, row_ends, sheet_index};
use super::spec::{Kind, SPEC};
use crate::error::{Code, Error, Result};
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep, Value};
use crate::path::navigate::{Available, ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::Inspect;

/// The largest zero-based row and column an ODS grid addresses (LibreOffice
/// Calc's 1,048,576 rows by 16,384 columns).
const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

/// The ODS read adapter over one handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Ods,
    /// Look coordinates up through the snapshot's cached row-run ends
    /// (record 0287), for callers that resolve many coordinates.
    indexed_rows: bool,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Ods) -> Self {
        Self {
            file,
            indexed_rows: false,
        }
    }

    /// A navigator that finds coordinates by binary search over the cached
    /// row runs; it resolves exactly what [`Self::new`] does.
    pub(crate) const fn with_indexed_rows(file: &'f Ods) -> Self {
        Self {
            file,
            indexed_rows: true,
        }
    }

    fn sheet(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let sheets = self.file.worksheets()?;
        let position = match select {
            Select::Name { name, case } => {
                let mut found = None;
                for (position, sheet) in sheets.iter().enumerate() {
                    if case.matches(&sheet.name, name) {
                        if found.is_some() {
                            return Err(Error::new(
                                Code::InvalidPath,
                                "the worksheet selector matches more than one sheet",
                            )
                            .with_format(FileFormat::Ods));
                        }
                        found = Some(position);
                    }
                }
                found
            },
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => sheets.len().checked_sub(1),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(
            match position.and_then(|position| sheets.get(position).map(|sheet| (position, sheet)))
            {
                Some((position, sheet)) => Lookup::Found(Node::new_sheet(
                    self.file,
                    u32::try_from(position).unwrap_or(u32::MAX),
                    sheet,
                )),
                None => Lookup::Missing,
            },
        )
    }

    /// A named range or expression: the Nth in document order, or by unique
    /// name.
    fn named_range(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let definitions = self.file.definitions()?;
        let index = match select {
            Select::Ordinal(position) => Some(position.get()),
            Select::None => Some(0),
            Select::Last => definitions.len().checked_sub(1),
            Select::Name { name, case } => {
                let mut found = None;
                for (index, definition) in definitions.iter().enumerate() {
                    if case.matches(definition.name(), name) {
                        if found.is_some() {
                            return Err(Error::new(
                                Code::InvalidPath,
                                "the named-range selector matches more than one definition",
                            )
                            .with_format(FileFormat::Ods));
                        }
                        found = Some(index);
                    }
                }
                found
            },
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match index.filter(|index| *index < definitions.len()) {
            Some(index) => Lookup::Found(Node::new_named_range(self.file, index)),
            None => Lookup::Missing,
        })
    }

    fn in_sheet(
        self,
        parent: &Node<'f>,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let position = parent.sheet_position();
        let within = match (kind, select) {
            (Kind::Cell, Select::Cell(at)) => at.row() <= MAX_ROW && at.column() <= MAX_COLUMN,
            (Kind::Range, Select::Range(range)) => [range.start(), range.end()]
                .iter()
                .all(|at| at.row() <= MAX_ROW && at.column() <= MAX_COLUMN),
            (Kind::Row, Select::Number(number)) => (1..=MAX_ROW as usize + 1).contains(&number),
            (Kind::Column, Select::Number(number)) => {
                (1..=MAX_COLUMN as usize + 1).contains(&number)
            },
            _ => true,
        };
        if !within {
            return Err(Error::new(
                Code::InvalidPath,
                "ODS coordinates must lie within A1:XFD1048576",
            )
            .with_format(FileFormat::Ods));
        }
        let sheet = parent.sheet()?;
        Ok(match (kind, select) {
            (Kind::Cell, Select::Cell(at)) => {
                let (row, column) = (at.row(), at.column());
                let cell = if self.indexed_rows {
                    row_ends(self.file, position)?.cell(sheet, row as usize, column as usize)
                } else {
                    sheet.cell(row as usize, column as usize)
                };
                Lookup::Found(Node::new_cell(self.file, position, (row, column), cell))
            },
            (Kind::Range, Select::Range(range)) => {
                Lookup::Found(Node::new_range(self.file, position, rect(range)))
            },
            (Kind::Row, Select::Number(number)) => {
                Lookup::Found(Node::new_row(self.file, position, index(number)?))
            },
            (Kind::Row, Select::Last | Select::None) => {
                let rows = sheet_index(self.file, position)?.row_entries();
                let row = if matches!(select, Select::Last) {
                    rows.last()
                } else {
                    rows.first()
                };
                row.map_or(Lookup::Missing, |(row, _)| {
                    Lookup::Found(Node::new_row(self.file, position, *row))
                })
            },
            (Kind::Column, Select::Number(number)) => {
                Lookup::Found(Node::new_column(self.file, position, index(number)?))
            },
            (Kind::Column, Select::Last | Select::None) => {
                let columns = sheet_index(self.file, position)?.columns();
                let column = if matches!(select, Select::Last) {
                    columns.last()
                } else {
                    columns.first()
                };
                column.map_or(Lookup::Missing, |column| {
                    Lookup::Found(Node::new_column(self.file, position, *column))
                })
            },
            (Kind::Comment, select) => self.comment(position, &sheet.name, select)?,
            (kind, select) if super::features::is_feature(kind) => {
                super::features::lookup(self.file, position, kind, select)?
            },
            (Kind::MergeCell, select) => {
                let merges = sheet_index(self.file, position)?.merges();
                let ordinal = match select {
                    Select::Ordinal(position) => Some(position.get()),
                    Select::None => Some(0),
                    Select::Last => merges.len().checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                match ordinal.and_then(|ordinal| merges.get(ordinal).map(|span| (ordinal, span))) {
                    Some((ordinal, span)) => {
                        Lookup::Found(Node::new_merge(self.file, position, *span, ordinal))
                    },
                    None => Lookup::Missing,
                }
            },
            _ => Lookup::NotApplicable,
        })
    }

    /// The Nth annotation of a sheet.
    fn comment(
        self,
        position: u32,
        name: &str,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let entries = annotations(self.file)?.entries();
        let mut on_sheet = entries
            .iter()
            .enumerate()
            .filter(|(_, entry)| entry.cell().sheet() == name);
        let found = match select {
            Select::Ordinal(ordinal) => on_sheet
                .nth(ordinal.get())
                .map(|(index, _)| (index, ordinal.get())),
            Select::None => on_sheet.next().map(|(index, _)| (index, 0)),
            Select::Last => on_sheet
                .enumerate()
                .last()
                .map(|(ordinal, (index, _))| (index, ordinal)),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match found {
            Some((index, ordinal)) => {
                Lookup::Found(Node::new_comment(self.file, position, index, ordinal))
            },
            None => Lookup::Missing,
        })
    }
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Ods;
    const ABSENT: &'static [&'static str] = &["slicer"];

    fn check_source(&self) -> Result<()> {
        self.file.check_source()
    }

    fn root(&self) -> Result<Node<'f>> {
        Ok(Node::new_workbook(self.file))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        // OpenDocument has no slicer element: the step names nothing.
        if step.kind == Kind::Slicer {
            return Ok(Lookup::NotApplicable);
        }
        if let Some(gap) = SPEC.elements[usize::from(step.kind.ordinal())].gap {
            return Ok(Lookup::Gap(gap));
        }
        match (parent.item(), step.kind) {
            (Item::Workbook, Kind::Sheet) => self.sheet(step.select),
            (Item::Workbook, Kind::NamedRange) => self.named_range(step.select),
            (Item::Sheet(_), kind) => self.in_sheet(parent, kind, step.select),
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        Children::walk(parent, kinds, self.file)
    }

    fn available(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        match (parent.item(), step.kind) {
            (Item::Workbook, Kind::Sheet) => out.kind("sheet", self.file.worksheet_count()?),
            (Item::Workbook, Kind::NamedRange) => {
                out.kind("namedrange", self.file.definitions()?.len());
            },
            (Item::Sheet(_), Kind::Row) => out.kind("row", parent.index()?.row_count()),
            (Item::Sheet(_), Kind::Column) => out.kind("col", parent.index()?.columns().len()),
            (Item::Sheet(_), Kind::MergeCell) => {
                out.kind("mergeCell", parent.index()?.merges().len());
            },
            (Item::Sheet(sheet), Kind::Comment) => {
                let count = annotations(self.file)?
                    .entries()
                    .iter()
                    .filter(|entry| entry.cell().sheet() == sheet.name)
                    .count();
                out.kind("comment", count);
            },
            (Item::Sheet(_), kind) if super::features::is_feature(kind) => {
                out.kind(
                    super::features::step_name(kind),
                    super::features::count(self.file, parent.sheet_position(), kind)?,
                );
            },
            _ => {},
        }
        Ok(())
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

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        Ok(match node.item() {
            Item::Cell(_) => node.cell_text(),
            Item::Row(row) => Some(Cow::Owned(super::node::joined_text(
                node,
                (row.index, 0, row.index, u32::MAX),
            )?)),
            Item::Column(column) => Some(Cow::Owned(super::node::joined_text(
                node,
                (0, column.index, u32::MAX, column.index),
            )?)),
            Item::Sheet(sheet) => Some(Cow::Borrowed(sheet.name.as_str())),
            Item::NamedRange(_) => node.named_range_text()?.map(Cow::Borrowed),
            Item::Comment(_) => node
                .comment()?
                .map(|entry| Cow::Owned(entry.annotation().text())),
            Item::Merge(merge) => Some(Cow::Owned(super::node::joined_text(node, merge.rect)?)),
            Item::Feature(feature) => feature.text().map(|text| Cow::Owned(text.into_owned())),
            Item::Workbook | Item::Range(_) => None,
        })
    }

    fn style<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        Ok(match node.item() {
            Item::Cell(cell) => cell
                .cell
                .and_then(|cell| cell.style_name.as_deref())
                .map(Cow::Borrowed),
            _ => None,
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        node.props_unfenced()
    }

    fn number(&self, node: &Node<'f>) -> Option<usize> {
        node.number().map(|number| number as usize)
    }

    fn cell(&self, node: &Node<'f>) -> Option<CellRef> {
        node.cell_ref()
    }

    fn cell_at(&self, node: &Node<'f>) -> Option<(u32, u32)> {
        node.cell()
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

    fn row_cell<'n>(
        &'n self,
        row: &'n Node<'f>,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        let Item::Row(item) = row.item() else {
            return Ok(None);
        };
        let Some(column) = column.checked_sub(1) else {
            return Ok(None);
        };
        let sheet = row.sheet()?;
        let Some(entry) = row
            .index()?
            .row_cells(item.index)
            .iter()
            .find(|entry| entry.column == column)
        else {
            return Ok(None);
        };
        let Some(cell) = entry.cell(sheet) else {
            return Ok(None);
        };
        Ok(Some(if raw {
            raw_value(cell)
        } else {
            Value::Text(Cow::Borrowed(cell.text.as_str()))
        }))
    }
}

fn rect(range: CellRange) -> (u32, u32, u32, u32) {
    let start = range.start();
    let end = range.end();
    (
        start.row().min(end.row()),
        start.column().min(end.column()),
        start.row().max(end.row()),
        start.column().max(end.column()),
    )
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
            .with_format(FileFormat::Ods)
        })
}
