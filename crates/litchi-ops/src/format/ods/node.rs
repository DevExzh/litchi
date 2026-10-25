//! Typed workbook nodes over the ODS leaf's typed worksheets (record 0277).
//!
//! A node borrows the handle's owner: a cell node holds the leaf
//! [`litchi_ods::Cell`] run it reads, a sheet node the
//! [`litchi_ods::Sheet`], so a walk copies no cell payload.

use std::borrow::Cow;
use std::fmt;
use std::sync::OnceLock;

use litchi_core::Position;
use litchi_core::selector::CellRef;
use litchi_ods::model::names::{Definition, Scope};
use litchi_ods::{Cell, CellValue, Merge, Sheet};

use super::Ods;
use super::children::Children;
use super::index::{SheetIndex, Span};
use super::spec::{
    CELL_PROPS, COMMENT_PROPS, CellKey, CommentKey, Kind, MERGE_PROPS, NAMED_RANGE_PROPS,
    NamedRangeKey, RANGE_PROPS, SHEET_PROPS, WORKBOOK_PROPS, WorkbookKey,
};
use crate::error::{Code, Error, Result};
use crate::format::workbook::Label;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};

/// Inclusive zero-based `(first_row, first_column, last_row, last_column)`.
pub(super) type Rect = (u32, u32, u32, u32);

/// A workbook node: an element of one ODS snapshot.
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Ods,
    item: Item<'f>,
    /// The position of the node's worksheet; its path is written from it
    /// and the item when asked for.
    sheet: u32,
}

#[derive(Clone)]
pub(super) enum Item<'f> {
    Workbook,
    Sheet(&'f Sheet),
    Cell(CellItem<'f>),
    Range(RangeItem),
    Row(LineItem),
    Column(LineItem),
    /// The definition at this index of the workbook's named ranges and
    /// expressions.
    NamedRange(usize),
    Comment(CommentItem),
    Merge(MergeItem),
    /// A worksheet feature: a hyperlink, picture, shape, conditional
    /// format, validation, auto-filter or page break (record 0327).
    Feature(std::sync::Arc<super::features::Feature>),
}

#[derive(Clone)]
pub(super) struct CellItem<'f> {
    /// Zero-based `(row, column)`.
    pub(super) at: (u32, u32),
    /// The run covering the coordinate, when one is stored.
    pub(super) cell: Option<&'f Cell>,
    /// The A1 label, written on first use (record 0287: a walk that never
    /// reads it writes none).
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct RangeItem {
    pub(super) rect: Rect,
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct LineItem {
    pub(super) index: u32,
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct CommentItem {
    /// The annotation's index in the snapshot's annotation owner.
    pub(super) entry: usize,
    /// Its ordinal among its sheet's annotations.
    pub(super) ordinal: usize,
}

#[derive(Clone)]
pub(super) struct MergeItem {
    pub(super) rect: Rect,
    pub(super) ordinal: usize,
    label: OnceLock<Label>,
}

impl<'f> Node<'f> {
    pub(super) fn new_workbook(file: &'f Ods) -> Self {
        Self {
            file,
            item: Item::Workbook,
            sheet: 0,
        }
    }

    pub(super) fn new_sheet(file: &'f Ods, position: u32, sheet: &'f Sheet) -> Self {
        Self {
            file,
            item: Item::Sheet(sheet),
            sheet: position,
        }
    }

    pub(super) fn new_cell(
        file: &'f Ods,
        sheet: u32,
        (row, column): (u32, u32),
        cell: Option<&'f Cell>,
    ) -> Self {
        Self {
            file,
            item: Item::Cell(CellItem {
                at: (row, column),
                cell,
                label: OnceLock::new(),
            }),
            sheet,
        }
    }

    pub(super) fn new_range(file: &'f Ods, sheet: u32, rect: Rect) -> Self {
        Self {
            file,
            item: Item::Range(RangeItem {
                rect,
                label: OnceLock::new(),
            }),
            sheet,
        }
    }

    pub(super) fn new_row(file: &'f Ods, sheet: u32, index: u32) -> Self {
        Self {
            file,
            item: Item::Row(LineItem {
                index,
                label: OnceLock::new(),
            }),
            sheet,
        }
    }

    pub(super) fn new_column(file: &'f Ods, sheet: u32, index: u32) -> Self {
        Self {
            file,
            item: Item::Column(LineItem {
                index,
                label: OnceLock::new(),
            }),
            sheet,
        }
    }

    pub(super) fn new_named_range(file: &'f Ods, index: usize) -> Self {
        Self {
            file,
            item: Item::NamedRange(index),
            sheet: 0,
        }
    }

    pub(super) fn new_comment(file: &'f Ods, sheet: u32, entry: usize, ordinal: usize) -> Self {
        Self {
            file,
            item: Item::Comment(CommentItem { entry, ordinal }),
            sheet,
        }
    }

    pub(super) fn new_merge(file: &'f Ods, sheet: u32, span: Span, ordinal: usize) -> Self {
        Self {
            file,
            item: Item::Merge(MergeItem {
                rect: (span.row, span.column, span.last_row, span.last_column),
                ordinal,
                label: OnceLock::new(),
            }),
            sheet,
        }
    }

    /// An element of the worksheet at `sheet`, for a sibling module.
    pub(super) fn in_sheet_item(file: &'f Ods, sheet: u32, item: Item<'f>) -> Self {
        Self { file, item, sheet }
    }

    pub(super) fn file(&self) -> &'f Ods {
        self.file
    }

    pub(super) fn item(&self) -> &Item<'f> {
        &self.item
    }

    /// The position of the node's worksheet (0 for the workbook and named
    /// ranges).
    pub(super) fn sheet_position(&self) -> u32 {
        self.sheet
    }

    /// The node's worksheet.
    pub(super) fn sheet(&self) -> Result<&'f Sheet> {
        if let Item::Sheet(sheet) = self.item {
            return Ok(sheet);
        }
        self.file
            .worksheets()?
            .get(self.sheet as usize)
            .ok_or_else(missing_sheet)
    }

    /// The content index of the node's worksheet.
    pub(super) fn index(&self) -> Result<&'f SheetIndex> {
        sheet_index(self.file, self.sheet)
    }

    /// The A1 label of a cell, range, row, column or merge node, written on
    /// first use.
    fn label(&self) -> Option<&Label> {
        Some(match &self.item {
            Item::Cell(cell) => cell.label.get_or_init(|| Label::cell(cell.at.1, cell.at.0)),
            Item::Range(range) => range.label.get_or_init(|| {
                let (first_row, first_column, last_row, last_column) = range.rect;
                Label::range((first_column, first_row), (last_column, last_row))
            }),
            Item::Row(row) => row.label.get_or_init(|| Label::row(row.index)),
            Item::Column(column) => column.label.get_or_init(|| Label::column(column.index)),
            Item::Merge(merge) => merge_label(merge),
            _ => return None,
        })
    }

    /// The node's own path step under its parent, written from the item.
    pub(super) fn last_step(&self) -> Option<PathStep<'static>> {
        Some(match &self.item {
            Item::Workbook => return None,
            Item::Sheet(_) => PathStep::at("sheet", Position::new(self.sheet as usize)),
            Item::Cell(cell) => PathStep::cell(cell.at.1, cell.at.0),
            Item::Range(range) => {
                let (first_row, first_column, last_row, last_column) = range.rect;
                PathStep::range((first_column, first_row), (last_column, last_row))
            },
            Item::Row(row) => PathStep::at("row", Position::new(row.index as usize)),
            Item::Column(column) => PathStep::column("col", column.index),
            Item::NamedRange(index) => PathStep::at("namedrange", Position::new(*index)),
            Item::Comment(comment) => PathStep::at("comment", Position::new(comment.ordinal)),
            Item::Merge(merge) => PathStep::at("mergeCell", Position::new(merge.ordinal)),
            Item::Feature(feature) => feature.step(),
        })
    }

    /// The node's element kind.
    pub(super) fn element(&self) -> Kind {
        match self.item {
            Item::Workbook => Kind::Workbook,
            Item::Sheet(_) => Kind::Sheet,
            Item::Cell(_) => Kind::Cell,
            Item::Range(_) => Kind::Range,
            Item::Row(_) => Kind::Row,
            Item::Column(_) => Kind::Column,
            Item::NamedRange(_) => Kind::NamedRange,
            Item::Comment(_) => Kind::Comment,
            Item::Merge(_) => Kind::MergeCell,
            Item::Feature(ref feature) => feature.kind,
        }
    }

    /// The OfficeCLI type of the node.
    #[must_use]
    pub fn kind(&self) -> &'static str {
        super::spec::SPEC.elements[self.element() as usize].name
    }

    /// The name of the node's worksheet; `None` for the workbook and named
    /// ranges.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when a filesystem-backed source changed.
    pub fn sheet_name(&self) -> Result<Option<&'f str>> {
        match self.item {
            Item::Workbook | Item::NamedRange(_) => Ok(None),
            _ => self.file.finish(|| Ok(Some(self.sheet()?.name.as_str()))),
        }
    }

    /// The zero-based `(column, row)` of a cell node.
    #[must_use]
    pub fn cell(&self) -> Option<(u32, u32)> {
        match &self.item {
            Item::Cell(cell) => Some((cell.at.1, cell.at.0)),
            _ => None,
        }
    }

    /// The 1-based number of a row or column node.
    #[must_use]
    pub fn number(&self) -> Option<u32> {
        match &self.item {
            Item::Row(line) | Item::Column(line) => line.index.checked_add(1),
            _ => None,
        }
    }

    pub(super) fn cell_ref(&self) -> Option<CellRef> {
        match &self.item {
            Item::Cell(_) => self
                .label()
                .and_then(|label| CellRef::parse(label.as_str())),
            _ => None,
        }
    }

    /// A cell's displayed text: the stored text, empty when no run is
    /// stored.
    pub(super) fn cell_text(&self) -> Option<Cow<'f, str>> {
        match &self.item {
            Item::Cell(cell) => Some(Cow::Borrowed(cell.cell.map_or("", |cell| &cell.text))),
            _ => None,
        }
    }

    /// The text of a named range: its reference or expression.
    pub(super) fn named_range_text(&self) -> Result<Option<&'f str>> {
        let Item::NamedRange(index) = self.item else {
            return Ok(None);
        };
        Ok(self
            .file
            .definitions()?
            .get(index)
            .map(definition_reference))
    }

    /// An annotation's anchor and value.
    pub(super) fn comment(&self) -> Result<Option<&'f litchi_ods::annotations::Entry>> {
        let Item::Comment(comment) = &self.item else {
            return Ok(None);
        };
        Ok(annotations(self.file)?.entries().get(comment.entry))
    }

    pub(super) fn props(&self) -> Result<Format<'f>> {
        self.file.finish(|| self.props_unfenced())
    }

    pub(super) fn props_unfenced(&self) -> Result<Format<'f>> {
        let (table, props) = match &self.item {
            Item::Workbook => (WORKBOOK_PROPS, Props::Workbook(self.file.metadata_ref()?)),
            Item::Sheet(sheet) => (SHEET_PROPS, Props::Sheet(&sheet.name)),
            Item::Cell(cell) => (
                CELL_PROPS,
                Props::Cell {
                    cell: cell.cell,
                    merge: self.cell_merge(cell)?,
                },
            ),
            Item::Range(range) => (RANGE_PROPS, Props::Range(self.is_merged(range.rect)?)),
            Item::Row(_) | Item::Column(_) => return Ok(Format::empty()),
            Item::NamedRange(index) => {
                let Some(definition) = self.file.definitions()?.get(*index) else {
                    return Ok(Format::empty());
                };
                (NAMED_RANGE_PROPS, Props::Name(definition))
            },
            Item::Comment(_) => {
                let Some(entry) = self.comment()? else {
                    return Ok(Format::empty());
                };
                let cell = entry.cell();
                let label = Label::cell(
                    u32::try_from(cell.column()).unwrap_or(u32::MAX),
                    u32::try_from(cell.row()).unwrap_or(u32::MAX),
                );
                let sheet = self.sheet()?.name.as_str();
                (
                    COMMENT_PROPS,
                    Props::Comment {
                        entry,
                        label,
                        sheet,
                    },
                )
            },
            Item::Merge(merge) => (MERGE_PROPS, Props::Merge(*merge_label(merge))),
            Item::Feature(feature) => (
                feature.table(),
                Props::Feature(std::sync::Arc::clone(feature)),
            ),
        };
        Ok(Format::new(table, Source::Ods(props)))
    }

    /// The merged span a cell belongs to: its own span when it anchors one,
    /// else the span covering it.
    fn cell_merge(&self, cell: &CellItem<'f>) -> Result<Option<(Rect, bool)>> {
        let Some(stored) = cell.cell else {
            return Ok(None);
        };
        let (row, column) = cell.at;
        match stored.merge {
            Merge::None => Ok(None),
            Merge::Span { rows, columns } => Ok(Some((
                (
                    row,
                    column,
                    row.saturating_add(u32::try_from(rows.get() - 1).unwrap_or(u32::MAX)),
                    column.saturating_add(u32::try_from(columns.get() - 1).unwrap_or(u32::MAX)),
                ),
                true,
            ))),
            Merge::Covered => Ok(self
                .index()?
                .merges()
                .iter()
                .find(|span| {
                    span.row <= row
                        && row <= span.last_row
                        && span.column <= column
                        && column <= span.last_column
                })
                .map(|span| {
                    (
                        (span.row, span.column, span.last_row, span.last_column),
                        false,
                    )
                })),
        }
    }

    /// Whether `rect` is exactly one merged span.
    fn is_merged(&self, rect: Rect) -> Result<bool> {
        let (first_row, first_column, last_row, last_column) = rect;
        let Some(cell) = self
            .sheet()?
            .cell(first_row as usize, first_column as usize)
        else {
            return Ok(false);
        };
        Ok(match cell.merge {
            Merge::Span { rows, columns } => {
                rows.get() as u64 == u64::from(last_row - first_row) + 1
                    && columns.get() as u64 == u64::from(last_column - first_column) + 1
            },
            _ => false,
        })
    }
}

/// The reference text of a definition: a range's address, an expression's
/// formula.
pub(super) fn definition_reference(definition: &Definition) -> &str {
    match definition {
        Definition::Range(range) => &range.cell_range_address,
        Definition::Expression(expression) => &expression.expression,
    }
}

/// The snapshot's annotation owner, parsed once.
pub(super) fn annotations(file: &Ods) -> Result<&litchi_ods::annotations::Snapshot> {
    let limits = file.odf_options().xml_limits_value();
    file.cache().annotations(|| file.content_xml(), &limits)
}

/// The content index of the worksheet at `position`.
pub(super) fn sheet_index(file: &Ods, position: u32) -> Result<&SheetIndex> {
    let sheets = file.worksheets()?;
    file.cache().sheet(sheets, position as usize)
}

/// The row-run ends of the worksheet at `position`.
pub(super) fn row_ends(file: &Ods, position: u32) -> Result<&super::index::RowEnds> {
    let sheets = file.worksheets()?;
    file.cache().rows(sheets, position as usize)
}

pub(super) fn missing_sheet() -> Error {
    Error::new(
        Code::PathNotFound,
        "the worksheet position is outside the workbook",
    )
    .with_format(litchi_core::FileFormat::Ods)
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ods::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

/// The property source of one node.
pub(crate) enum Props<'a> {
    Workbook(&'a litchi_core::Metadata),
    Sheet(&'a str),
    Cell {
        cell: Option<&'a Cell>,
        merge: Option<(Rect, bool)>,
    },
    Range(bool),
    Name(&'a Definition),
    Comment {
        entry: &'a litchi_ods::annotations::Entry,
        label: Label,
        sheet: &'a str,
    },
    Merge(Label),
    /// A worksheet feature (record 0327); its values are owned.
    Feature(std::sync::Arc<super::features::Feature>),
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Workbook(_) => "Workbook",
            Self::Sheet(_) => "Sheet",
            Self::Cell { .. } => "Cell",
            Self::Range(_) => "Range",
            Self::Name(_) => "Name",
            Self::Comment { .. } => "Comment",
            Self::Merge(_) => "Merge",
            Self::Feature(_) => "Feature",
        };
        formatter.debug_tuple("Props").field(&name).finish()
    }
}

impl<'a> Props<'a> {
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match self {
            Self::Workbook(metadata) => workbook_value(metadata, WorkbookKey::at(index)?),
            Self::Sheet(name) => (index == 0).then_some(Value::Text(Cow::Borrowed(*name))),
            Self::Cell { cell, merge } => cell_value(*cell, *merge, CellKey::at(index)?),
            Self::Range(merged) => (index == 0 && *merged).then_some(Value::Bool(true)),
            Self::Name(definition) => match NamedRangeKey::at(index)? {
                NamedRangeKey::Name => Some(Value::Text(Cow::Borrowed(definition.name()))),
                NamedRangeKey::Ref => {
                    Some(Value::Text(Cow::Borrowed(definition_reference(definition))))
                },
                NamedRangeKey::Scope => Some(Value::Text(match definition.scope() {
                    Scope::Global => Cow::Borrowed("workbook"),
                    Scope::Sheet(name) => Cow::Borrowed(name.as_str()),
                })),
            },
            Self::Comment {
                entry,
                label,
                sheet,
            } => {
                let annotation = entry.annotation();
                match CommentKey::at(index)? {
                    CommentKey::Ref => Some(Value::Text(Cow::Owned(label.as_str().to_owned()))),
                    CommentKey::Author => annotation
                        .creator()
                        .filter(|author| !author.is_empty())
                        .map(|author| Value::Text(Cow::Owned(author))),
                    CommentKey::Text => Some(Value::Text(Cow::Owned(annotation.text()))),
                    CommentKey::Date => annotation
                        .date()
                        .filter(|date| !date.is_empty())
                        .map(|date| Value::Text(Cow::Owned(date))),
                    CommentKey::AnchoredTo => Some(Value::Text(Cow::Owned(format!(
                        "/{sheet}/{}",
                        label.as_str()
                    )))),
                }
            },
            Self::Merge(label) => {
                (index == 0).then(|| Value::Text(Cow::Owned(label.as_str().to_owned())))
            },
            Self::Feature(feature) => feature.value(index).map(Value::into_owned),
        }
    }
}

fn text(value: &Option<String>) -> Option<Value<'_>> {
    value
        .as_deref()
        .filter(|value| !value.is_empty())
        .map(|value| Value::Text(Cow::Borrowed(value)))
}

fn workbook_value(metadata: &litchi_core::Metadata, key: WorkbookKey) -> Option<Value<'_>> {
    match key {
        WorkbookKey::Title => text(&metadata.title),
        WorkbookKey::Author => text(&metadata.author),
        WorkbookKey::Subject => text(&metadata.subject),
        WorkbookKey::Keywords => Value::owned_list(&metadata.keywords),
        WorkbookKey::Description => text(&metadata.description),
        WorkbookKey::Category => text(&metadata.category),
        WorkbookKey::LastModifiedBy => text(&metadata.last_modified_by),
        WorkbookKey::RevisionNumber => text(&metadata.revision),
        WorkbookKey::Created => metadata
            .created
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|value| Value::Text(Cow::Owned(value.to_string()))),
        WorkbookKey::Modified => metadata
            .modified
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|value| Value::Text(Cow::Owned(value.to_string()))),
        _ => None,
    }
}

/// The `type` token of an ODF cell value; `None` for an empty value.
pub(super) fn cell_type(value: &CellValue) -> Option<&'static str> {
    Some(match value {
        CellValue::Empty => return None,
        CellValue::Text(_) => "String",
        CellValue::Number(_) => "Number",
        CellValue::Currency { .. } => "Currency",
        CellValue::Percentage(_) => "Percentage",
        CellValue::Boolean(_) => "Boolean",
        CellValue::Date(_) => "Date",
        CellValue::Time(_) => "Time",
        CellValue::Unknown { .. } => "Unknown",
    })
}

fn cell_value<'a>(
    cell: Option<&'a Cell>,
    merge: Option<(Rect, bool)>,
    key: CellKey,
) -> Option<Value<'a>> {
    let Some(cell) = cell else {
        return matches!(key, CellKey::Empty).then_some(Value::Bool(true));
    };
    match key {
        CellKey::Type => cell_type(&cell.value).map(Value::Token),
        CellKey::Formula => cell
            .formula
            .as_deref()
            .map(|formula| Value::Text(Cow::Borrowed(formula))),
        CellKey::CachedValue => (cell.formula.is_some() && !cell.text.is_empty())
            .then_some(Value::Text(Cow::Borrowed(cell.text.as_str()))),
        CellKey::Empty => (matches!(cell.value, CellValue::Empty)
            && cell.text.is_empty()
            && cell.formula.is_none())
        .then_some(Value::Bool(true)),
        CellKey::Merge => merge.map(|((first_row, first_column, last_row, last_column), _)| {
            Value::Text(Cow::Owned(
                Label::range((first_column, first_row), (last_column, last_row))
                    .as_str()
                    .to_owned(),
            ))
        }),
        CellKey::MergeAnchor => merge
            .filter(|(_, anchor)| *anchor)
            .map(|_| Value::Bool(true)),
        CellKey::Link => cell
            .hyperlinks
            .first()
            .map(|link| Value::Text(Cow::Borrowed(link.href.as_str()))),
    }
}

/// The raw stored value of a cell, for numeric comparisons.
pub(super) fn raw_value(cell: &Cell) -> Value<'_> {
    match &cell.value {
        CellValue::Number(value)
        | CellValue::Percentage(value)
        | CellValue::Currency { value, .. } => {
            Value::number(*value).unwrap_or_else(|| Value::Text(Cow::Borrowed(&cell.text)))
        },
        _ => Value::Text(Cow::Borrowed(&cell.text)),
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> litchi_core::FileFormat {
        litchi_core::FileFormat::Ods
    }

    fn path(&self) -> NodePath<'_> {
        match (&self.item, self.last_step()) {
            (_, None) => NodePath::ROOT,
            (Item::Sheet(_) | Item::NamedRange(_), Some(step)) => NodePath::short(step, None),
            (_, Some(step)) => NodePath::short(
                PathStep::at("sheet", Position::new(self.sheet as usize)),
                Some(step),
            ),
        }
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.finish(|| {
            Ok(match &self.item {
                Item::Cell(_) => self.cell_text(),
                Item::Row(row) => Some(Cow::Owned(joined_text(
                    self,
                    (row.index, 0, row.index, u32::MAX),
                )?)),
                Item::Column(column) => Some(Cow::Owned(joined_text(
                    self,
                    (0, column.index, u32::MAX, column.index),
                )?)),
                Item::Sheet(sheet) => Some(Cow::Borrowed(sheet.name.as_str())),
                Item::NamedRange(_) => self.named_range_text()?.map(Cow::Borrowed),
                Item::Comment(_) => self
                    .comment()?
                    .map(|entry| Cow::Owned(entry.annotation().text())),
                Item::Merge(merge) => Some(Cow::Owned(joined_text(self, merge.rect)?)),
                Item::Feature(feature) => feature.text(),
                Item::Workbook | Item::Range(_) => None,
            })
        })
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.finish(|| {
            Ok(match &self.item {
                Item::Workbook | Item::Comment(_) => None,
                Item::Sheet(sheet) => Some(Preview::Text(Cow::Borrowed(sheet.name.as_str()))),
                Item::Cell(_)
                | Item::Range(_)
                | Item::Row(_)
                | Item::Column(_)
                | Item::Merge(_) => self
                    .label()
                    .map(|label| Preview::Text(Cow::Borrowed(label.as_str()))),
                Item::NamedRange(_) => self
                    .named_range_text()?
                    .map(|text| Preview::Text(Cow::Borrowed(text))),
                Item::Feature(feature) => feature.preview(),
            })
        })
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(match &self.item {
            Item::Cell(cell) => cell
                .cell
                .and_then(|cell| cell.style_name.as_deref())
                .map(Cow::Borrowed),
            _ => None,
        })
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        self.file.finish(|| {
            let count = match &self.item {
                Item::Workbook => self.file.worksheet_count()?,
                Item::Sheet(_) | Item::Row(_) | Item::Range(_) => Children::count(self, options)?,
                _ => 0,
            };
            Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
        })
    }

    fn format(&self) -> Result<Format<'_>> {
        self.props()
    }

    fn children(&self, options: &ViewOptions) -> ViewChildren<'_, 'f> {
        if let Err(error) = self.file.check_source() {
            return ViewChildren::failed(error);
        }
        match Children::view(self, options) {
            Ok(children) => ViewChildren::ods(children),
            Err(error) => ViewChildren::failed(error),
        }
    }
}

/// The A1 label of a merge span, written on first use.
fn merge_label(merge: &MergeItem) -> &Label {
    merge.label.get_or_init(|| {
        let (row, column, last_row, last_column) = merge.rect;
        Label::range((column, row), (last_column, last_row))
    })
}

/// The displayed text of the content cells in `rect`, space-separated.
pub(super) fn joined_text(node: &Node<'_>, rect: Rect) -> Result<String> {
    let sheet = node.sheet()?;
    let mut out = String::new();
    for (index, entry) in node.index()?.cells_in(rect).enumerate() {
        if index > 0 {
            out.push(' ');
        }
        if let Some(cell) = entry.cell(sheet) {
            out.push_str(&cell.text);
        }
    }
    Ok(out)
}
