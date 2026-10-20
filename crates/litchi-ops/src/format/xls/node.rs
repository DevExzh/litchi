//! Typed workbook nodes backed by the native legacy XLS Package projection.

use std::borrow::Cow;
use std::fmt;
use std::sync::OnceLock;

use litchi_core::Position;
use litchi_core::selector::CellRef;
use litchi_core::sheet::CellValue;

use super::Xls;
use super::children::Children;
use super::owner::{CellData, ColumnLayout, Grid, LayoutLookup as _, RowLayout, SheetHandle};
use super::spec::{
    CELL_PROPS, COLUMN_PROPS, CellKey, ColumnKey, Kind, NAMED_RANGE_PROPS, NamedRangeKey,
    ROW_PROPS, RowKey, SHEET_PROPS, WORKBOOK_PROPS, WorkbookKey,
};
use crate::error::Result;
use crate::format::workbook::Label;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};

/// A workbook node. The node keeps only a checked path and a small native
/// owner handle; worksheet payload is pinned lazily by `Sheet::grid`.
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Xls,
    item: Item<'f>,
    /// The position of the node's worksheet; its path is written from it
    /// and the item when asked for (record 0262), so walks move a smaller
    /// node.
    sheet: usize,
}

#[derive(Clone)]
pub(super) enum Item<'f> {
    Workbook,
    Sheet(Sheet<'f>),
    Cell(CellItem),
    Range(RangeItem),
    Row(RowItem),
    Column(ColumnItem),
    /// The defined name at this index of the workbook's non-macro names.
    NamedRange(usize),
    /// A worksheet feature: a note, hyperlink, picture, shape, conditional
    /// format, validation, auto-filter or page break (record 0327).
    Feature(std::sync::Arc<super::features::Feature>),
}

/// The workbook's non-macro defined names, through the eager reader.
///
/// The positional owner does not parse `Lbl` records, so a source-backed
/// workbook materializes its eager reader once (cached by the native owner)
/// when a path, query or view first reaches a defined name.
pub(super) fn defined_names(file: &Xls) -> Result<&[litchi_xls::DefinedName]> {
    Ok(file.owner().reader_ref()?.defined_names())
}

#[derive(Clone)]
pub(super) struct Sheet<'f> {
    pub(super) position: usize,
    pub(super) handle: SheetHandle<'f>,
    grid: OnceLock<Grid>,
}

impl<'f> Sheet<'f> {
    pub(super) fn new(position: usize, handle: SheetHandle<'f>) -> Self {
        Self {
            position,
            handle,
            grid: OnceLock::new(),
        }
    }

    pub(super) fn grid(&self) -> Result<&Grid> {
        if let Some(grid) = self.grid.get() {
            return Ok(grid);
        }
        let grid = self.handle.grid()?;
        Ok(self.grid.get_or_init(|| grid))
    }
}

#[derive(Clone)]
pub(super) struct CellItem {
    pub(super) at: (u32, u32),
    pub(super) source: CellSource,
    pub(super) label: Label,
}

#[derive(Clone)]
pub(super) enum CellSource {
    Grid(Grid),
    Direct(Option<CellData>),
    Stub,
}

#[derive(Clone)]
pub(super) struct RangeItem {
    pub(super) grid: Grid,
    pub(super) rect: Rect,
    pub(super) label: Label,
}

#[derive(Clone)]
pub(super) struct RowItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    pub(super) label: Label,
}

#[derive(Clone)]
pub(super) struct ColumnItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    pub(super) label: Label,
}

/// Inclusive zero-based rectangle used by the sparse grid bridge.
pub(super) type Rect = (u32, u32, u32, u32);

impl<'f> Node<'f> {
    pub(super) fn new_workbook(file: &'f Xls) -> Self {
        Self {
            file,
            item: Item::Workbook,
            sheet: 0,
        }
    }

    pub(super) fn new_sheet(file: &'f Xls, sheet: Sheet<'f>) -> Self {
        let position = sheet.position;
        Self {
            file,
            item: Item::Sheet(sheet),
            sheet: position,
        }
    }

    /// An element of the worksheet at `sheet`, for a sibling module.
    pub(super) fn in_sheet_item(file: &'f Xls, sheet: usize, item: Item<'f>) -> Self {
        Self::in_sheet(file, sheet, item)
    }

    fn in_sheet(file: &'f Xls, sheet: usize, item: Item<'f>) -> Self {
        Self { file, item, sheet }
    }

    pub(super) fn new_cell(
        file: &'f Xls,
        sheet: usize,
        at: (u32, u32),
        source: CellSource,
    ) -> Self {
        let (row, column) = at;
        Self::in_sheet(
            file,
            sheet,
            Item::Cell(CellItem {
                at,
                source,
                label: Label::cell(column, row),
            }),
        )
    }

    pub(super) fn new_range(file: &'f Xls, sheet: usize, grid: Grid, rect: Rect) -> Self {
        let (first_row, first_column, last_row, last_column) = rect;
        Self::in_sheet(
            file,
            sheet,
            Item::Range(RangeItem {
                grid,
                rect,
                label: Label::range((first_column, first_row), (last_column, last_row)),
            }),
        )
    }

    pub(super) fn new_named_range(file: &'f Xls, index: usize) -> Self {
        Self {
            file,
            item: Item::NamedRange(index),
            sheet: 0,
        }
    }

    /// The name's reference without its leading `=`, else the name.
    pub(super) fn named_range_text(&self) -> Result<Option<String>> {
        let Item::NamedRange(index) = &self.item else {
            return Ok(None);
        };
        Ok(defined_names(self.file)?.get(*index).map(|name| {
            match name
                .formula
                .as_deref()
                .map(|formula| formula.strip_prefix('=').unwrap_or(formula))
            {
                Some(reference) if !reference.is_empty() => reference.to_owned(),
                _ => name.name.clone(),
            }
        }))
    }

    pub(super) fn new_row(file: &'f Xls, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::in_sheet(
            file,
            sheet,
            Item::Row(RowItem {
                grid,
                index,
                label: Label::row(index),
            }),
        )
    }

    pub(super) fn new_column(file: &'f Xls, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::in_sheet(
            file,
            sheet,
            Item::Column(ColumnItem {
                grid,
                index,
                label: Label::column(index),
            }),
        )
    }

    pub(super) fn file(&self) -> &'f Xls {
        self.file
    }

    pub(super) fn item(&self) -> &Item<'f> {
        &self.item
    }

    pub(super) fn sheet_position(&self) -> usize {
        self.sheet
    }

    /// The node's own path step under its parent, written from the item.
    pub(super) fn last_step(&self) -> Option<PathStep<'static>> {
        Some(match &self.item {
            Item::Workbook => return None,
            Item::Sheet(_) => PathStep::at("sheet", Position::new(self.sheet)),
            Item::Cell(cell) => {
                let (row, column) = cell.at;
                PathStep::cell(column, row)
            },
            Item::Range(range) => {
                let (first_row, first_column, last_row, last_column) = range.rect;
                PathStep::range((first_column, first_row), (last_column, last_row))
            },
            Item::Row(row) => PathStep::at("row", Position::new(row.index as usize)),
            Item::Column(column) => PathStep::column("col", column.index),
            Item::NamedRange(index) => PathStep::at("namedrange", Position::new(*index)),
            Item::Feature(feature) => feature.step(),
        })
    }

    pub(super) fn element(&self) -> Kind {
        match self.item {
            Item::Workbook => Kind::Workbook,
            Item::Sheet(_) => Kind::Sheet,
            Item::Cell(_) => Kind::Cell,
            Item::Range(_) => Kind::Range,
            Item::Row(_) => Kind::Row,
            Item::Column(_) => Kind::Column,
            Item::NamedRange(_) => Kind::NamedRange,
            Item::Feature(ref feature) => feature.kind,
        }
    }

    #[must_use]
    pub fn kind(&self) -> &'static str {
        super::spec::SPEC.elements[self.element() as usize].name
    }

    pub fn sheet_name(&self) -> Result<Option<String>> {
        self.file.finish(|| match &self.item {
            Item::Workbook => Ok(None),
            Item::Sheet(sheet) => Ok(Some(sheet.handle.name().to_owned())),
            _ => SheetHandle::at(self.file, self.sheet)
                .map(|handle| handle.map(|handle| handle.name().to_owned())),
        })
    }

    #[must_use]
    pub fn cell(&self) -> Option<(u32, u32)> {
        match &self.item {
            Item::Cell(cell) => Some((cell.at.1, cell.at.0)),
            _ => None,
        }
    }

    #[must_use]
    pub fn number(&self) -> Option<u32> {
        match &self.item {
            Item::Row(row) => row.index.checked_add(1),
            Item::Column(column) => column.index.checked_add(1),
            _ => None,
        }
    }

    pub(super) fn cell_ref(&self) -> Option<CellRef> {
        match &self.item {
            Item::Cell(cell) => CellRef::parse(cell.label.as_str()),
            _ => None,
        }
    }

    fn sheet_handle(&self) -> Result<Option<SheetHandle<'f>>> {
        SheetHandle::at(self.file, self.sheet)
    }

    pub(super) fn cell_state(&self) -> Result<CellState<'_>> {
        let Item::Cell(cell) = &self.item else {
            return Ok(CellState::Missing);
        };
        match &cell.source {
            CellSource::Stub => Ok(CellState::Stub),
            CellSource::Direct(cell) => {
                Ok(cell.as_ref().map_or(CellState::Missing, CellState::Stored))
            },
            CellSource::Grid(grid) => Ok(grid
                .record(cell.at)
                .map_or(CellState::Missing, CellState::Stored)),
        }
    }

    pub(super) fn cell_text(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(match self.cell_state()? {
            CellState::Stored(cell) => Some(display_value(&cell.value)),
            CellState::Missing | CellState::Stub => Some(Cow::Borrowed("")),
        })
    }

    pub(super) fn props(&self) -> Result<Format<'_>> {
        self.file.finish(|| self.props_unfenced())
    }

    pub(super) fn props_unfenced(&self) -> Result<Format<'_>> {
        let (table, props) = match &self.item {
            Item::Workbook => {
                let metadata = self.file.owner().metadata()?;
                (
                    WORKBOOK_PROPS,
                    Props::Workbook {
                        metadata: Box::new(metadata),
                        date_1904: self.file.date_1904()?,
                    },
                )
            },
            Item::Sheet(sheet) => (
                SHEET_PROPS,
                Props::Sheet {
                    name: sheet.handle.name().to_owned(),
                    visibility: sheet.handle.visibility(),
                },
            ),
            Item::Cell(_) => {
                let state = self.cell_state()?;
                let styles = match state {
                    CellState::Stored(_) => Some(styles(self.file)?),
                    CellState::Missing | CellState::Stub => None,
                };
                (CELL_PROPS, Props::Cell(CellProps { state, styles }))
            },
            Item::Range(_) => return Ok(Format::empty()),
            Item::Row(row) => (
                ROW_PROPS,
                Props::Row(self.sheet_handle()?.and_then_row(row.index)?),
            ),
            Item::Column(column) => (
                COLUMN_PROPS,
                Props::Column(self.sheet_handle()?.and_then_column(column.index)?),
            ),
            Item::NamedRange(index) => {
                let names = defined_names(self.file)?;
                let Some(name) = names.get(*index) else {
                    return Ok(Format::empty());
                };
                let scope = match name.scope {
                    litchi_xls::NameScope::Workbook => None,
                    litchi_xls::NameScope::Worksheet(position) => {
                        SheetHandle::at(self.file, position)?
                            .map(|handle| Cow::Owned(handle.name().to_owned()))
                    },
                };
                (NAMED_RANGE_PROPS, Props::NamedRange { name, scope })
            },
            Item::Feature(feature) => {
                let sheet = SheetHandle::at(self.file, self.sheet)?
                    .map(|handle| handle.name().to_owned())
                    .unwrap_or_default();
                (feature.table(), Props::Feature(feature, sheet))
            },
        };
        Ok(Format::new(table, Source::Xls(props)))
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("xls::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

#[derive(Clone, Copy)]
pub(super) enum CellState<'a> {
    Stored(&'a CellData),
    Missing,
    Stub,
}

pub(crate) enum Props<'a> {
    Workbook {
        metadata: Box<litchi_core::Metadata>,
        date_1904: bool,
    },
    Sheet {
        name: String,
        visibility: litchi_xls::SheetVisibility,
    },
    Cell(CellProps<'a>),
    Row(Option<RowLayout>),
    Column(Option<ColumnLayout>),
    NamedRange {
        name: &'a litchi_xls::DefinedName,
        scope: Option<Cow<'a, str>>,
    },
    /// A worksheet feature and its worksheet's name (record 0327).
    Feature(&'a super::features::Feature, String),
}

pub(crate) struct CellProps<'a> {
    state: CellState<'a>,
    styles: Option<super::style::Styles<'a>>,
}

/// The formatting tables of a file's cells.
fn styles(file: &Xls) -> Result<super::style::Styles<'_>> {
    Ok(match file.owner().source_workbook() {
        Some(source) => super::style::Styles::Source {
            formatting: source.formatting()?,
            tables: source.style_tables()?,
        },
        None => super::style::Styles::Eager(file.owner().reader_ref()?),
    })
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Workbook { .. } => "Workbook",
            Self::Sheet { .. } => "Sheet",
            Self::Cell(_) => "Cell",
            Self::Row(_) => "Row",
            Self::Column(_) => "Column",
            Self::NamedRange { .. } => "NamedRange",
            Self::Feature(..) => "Feature",
        };
        formatter.debug_tuple("Props").field(&name).finish()
    }
}

impl<'a> Props<'a> {
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match self {
            Self::Workbook {
                metadata,
                date_1904,
            } => workbook_value(metadata, *date_1904, WorkbookKey::at(index)?),
            Self::Sheet { name, visibility } => match super::spec::SheetKey::at(index)? {
                super::spec::SheetKey::Name => Some(Value::Text(Cow::Owned(name.clone()))),
                super::spec::SheetKey::Hidden => matches!(
                    visibility,
                    litchi_xls::SheetVisibility::Hidden | litchi_xls::SheetVisibility::VeryHidden
                )
                .then_some(Value::Bool(true)),
                super::spec::SheetKey::Visibility => match visibility {
                    litchi_xls::SheetVisibility::Visible => None,
                    litchi_xls::SheetVisibility::Hidden => Some(Value::Token("hidden")),
                    litchi_xls::SheetVisibility::VeryHidden => Some(Value::Token("veryHidden")),
                },
            },
            Self::Cell(cell) => cell.get(CellKey::at(index)?),
            Self::Feature(feature, sheet) => feature.value(index, sheet),
            Self::Row(row) => row.and_then(|layout| row_value(layout, RowKey::at(index)?)),
            Self::Column(column) => {
                column.and_then(|layout| column_value(layout, ColumnKey::at(index)?))
            },
            Self::NamedRange { name, scope } => match NamedRangeKey::at(index)? {
                NamedRangeKey::Name => Some(Value::Text(Cow::Borrowed(&name.name))),
                NamedRangeKey::Ref => name
                    .formula
                    .as_deref()
                    .map(|formula| formula.strip_prefix('=').unwrap_or(formula))
                    .filter(|reference| !reference.is_empty())
                    .map(|reference| Value::Text(Cow::Borrowed(reference))),
                NamedRangeKey::Scope => match name.scope {
                    litchi_xls::NameScope::Workbook => Some(Value::Text(Cow::Borrowed("workbook"))),
                    litchi_xls::NameScope::Worksheet(_) => scope.clone().map(Value::Text),
                },
                NamedRangeKey::Comment => name
                    .comment
                    .as_deref()
                    .filter(|comment| !comment.is_empty())
                    .map(|comment| Value::Text(Cow::Borrowed(comment))),
                NamedRangeKey::Volatile => name.function.then_some(Value::Bool(true)),
            },
        }
    }
}

fn workbook_value(
    metadata: &litchi_core::Metadata,
    date_1904: bool,
    key: WorkbookKey,
) -> Option<Value<'static>> {
    match key {
        WorkbookKey::Title => metadata.title.as_ref().map(|value| text(value)),
        WorkbookKey::Author => metadata.author.as_ref().map(|value| text(value)),
        WorkbookKey::Subject => metadata.subject.as_ref().map(|value| text(value)),
        WorkbookKey::Keywords => Value::owned_list(&metadata.keywords),
        WorkbookKey::Description => metadata.description.as_ref().map(|value| text(value)),
        WorkbookKey::Category => metadata.category.as_ref().map(|value| text(value)),
        WorkbookKey::LastModifiedBy => metadata.last_modified_by.as_ref().map(|value| text(value)),
        WorkbookKey::RevisionNumber => metadata.revision.as_ref().map(|value| text(value)),
        WorkbookKey::Created => metadata
            .created
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|value| Value::Text(Cow::Owned(value.to_string()))),
        WorkbookKey::Modified => metadata
            .modified
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|value| Value::Text(Cow::Owned(value.to_string()))),
        WorkbookKey::Date1904 => date_1904.then_some(Value::Bool(true)),
        WorkbookKey::ActiveTab
        | WorkbookKey::ExtendedTemplate
        | WorkbookKey::ExtendedManager
        | WorkbookKey::ExtendedCompany
        | WorkbookKey::ExtendedApplication
        | WorkbookKey::ExtendedPages
        | WorkbookKey::ExtendedWords
        | WorkbookKey::ExtendedCharacters
        | WorkbookKey::ExtendedTotalTime => None,
    }
}

fn text(value: &str) -> Value<'static> {
    Value::Text(Cow::Owned(value.to_owned()))
}

impl<'a> CellProps<'a> {
    fn get(&self, key: CellKey) -> Option<Value<'a>> {
        let value = match self.state {
            CellState::Stored(cell) => &cell.value,
            CellState::Stub => {
                return matches!(key, CellKey::Empty).then_some(Value::Bool(true));
            },
            CellState::Missing => return None,
        };
        match key {
            CellKey::Type => Some(Value::Token(cell_type(value))),
            CellKey::Formula => match self.state {
                CellState::Stored(cell) => cell.formula_text.as_deref().map(|formula| {
                    Value::Text(Cow::Borrowed(formula.strip_prefix('=').unwrap_or(formula)))
                }),
                _ => None,
            },
            CellKey::CachedValue => match self.state {
                CellState::Stored(cell) if cell.formula => {
                    let cached = display_value(value);
                    (!cached.is_empty()).then_some(Value::Text(cached))
                },
                _ => None,
            },
            CellKey::Empty => matches!(value, CellValue::Empty).then_some(Value::Bool(true)),
            key => match (self.state, &self.styles) {
                (CellState::Stored(cell), Some(styles)) => {
                    super::style::value(styles, cell.xf, key)
                },
                _ => None,
            },
        }
    }
}

fn row_value(layout: RowLayout, key: RowKey) -> Option<Value<'static>> {
    match key {
        RowKey::Height => Some(Value::Length(crate::value::Length::from_emu(
            i64::from(layout.height_twips) * 635,
        ))),
        RowKey::Hidden => layout.hidden.then_some(Value::Bool(true)),
        RowKey::OutlineLevel => {
            (layout.outline_level > 0).then_some(Value::Int(i64::from(layout.outline_level)))
        },
        RowKey::Collapsed => layout.collapsed.then_some(Value::Bool(true)),
        RowKey::CustomHeight => layout.custom_height.then_some(Value::Bool(true)),
    }
}

fn column_value(layout: ColumnLayout, key: ColumnKey) -> Option<Value<'static>> {
    match key {
        ColumnKey::Width => Value::number(f64::from(layout.width_256ths) / 256.0),
        ColumnKey::Hidden => layout.hidden.then_some(Value::Bool(true)),
        ColumnKey::CustomWidth => layout.user_set.then_some(Value::Bool(true)),
        ColumnKey::OutlineLevel => {
            (layout.outline_level > 0).then_some(Value::Int(i64::from(layout.outline_level)))
        },
        ColumnKey::Collapsed => layout.collapsed.then_some(Value::Bool(true)),
    }
}

pub(super) fn display_value(value: &CellValue) -> Cow<'_, str> {
    match value {
        CellValue::Empty => Cow::Borrowed(""),
        CellValue::Bool(value) => Cow::Borrowed(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => Cow::Owned(value.to_string()),
        CellValue::Float(value) | CellValue::DateTime(value) => Cow::Owned(value.to_string()),
        CellValue::String(value) | CellValue::Error(value) => Cow::Borrowed(value),
        CellValue::Formula { cached_value, .. } => cached_value
            .as_deref()
            .map_or(Cow::Borrowed(""), display_value),
    }
}

pub(super) fn cell_type(value: &CellValue) -> &'static str {
    match value {
        CellValue::Bool(_) => "Boolean",
        CellValue::String(_) => "SharedString",
        CellValue::DateTime(_) => "Date",
        CellValue::Error(_) => "Error",
        CellValue::Formula { cached_value, .. } => {
            cached_value.as_deref().map_or("Number", cell_type)
        },
        CellValue::Empty | CellValue::Int(_) | CellValue::Float(_) => "Number",
    }
}

pub(super) fn raw_value(cell: &CellData) -> Value<'static> {
    match &cell.value {
        CellValue::Int(value) => Value::Int(*value),
        CellValue::Float(value) | CellValue::DateTime(value) => {
            Value::number(*value).unwrap_or_else(|| Value::Text(Cow::Owned(value.to_string())))
        },
        _ => Value::Text(Cow::Owned(display_value(&cell.value).into_owned())),
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> litchi_core::FileFormat {
        litchi_core::FileFormat::Xls
    }

    fn path(&self) -> NodePath<'_> {
        match (&self.item, self.last_step()) {
            (_, None) => NodePath::ROOT,
            (Item::Sheet(_) | Item::NamedRange(_), Some(step)) => NodePath::short(step, None),
            (_, Some(step)) => {
                NodePath::short(PathStep::at("sheet", Position::new(self.sheet)), Some(step))
            },
        }
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.finish(|| {
            Ok(match &self.item {
                Item::Cell(_) => self.cell_text()?,
                Item::Row(row) => Some(Cow::Owned(joined_text(&row.grid, row_rect(row.index)?)?)),
                Item::Column(column) => Some(Cow::Owned(joined_text(
                    &column.grid,
                    column_rect(column.index)?,
                )?)),
                Item::Sheet(_) => self.sheet_name()?.map(Cow::Owned),
                Item::NamedRange(_) => self.named_range_text()?.map(Cow::Owned),
                Item::Feature(feature) => feature.text(),
                Item::Workbook | Item::Range(_) => None,
            })
        })
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.finish(|| {
            Ok(match &self.item {
                Item::Workbook => None,
                Item::Sheet(sheet) => Some(Preview::Text(Cow::Borrowed(sheet.handle.name()))),
                Item::Cell(cell) => Some(Preview::Text(Cow::Borrowed(cell.label.as_str()))),
                Item::Range(range) => Some(Preview::Text(Cow::Borrowed(range.label.as_str()))),
                Item::Row(row) => Some(Preview::Text(Cow::Borrowed(row.label.as_str()))),
                Item::Column(column) => Some(Preview::Text(Cow::Borrowed(column.label.as_str()))),
                Item::NamedRange(_) => self
                    .named_range_text()?
                    .map(|text| Preview::Text(Cow::Owned(text))),
                Item::Feature(feature) => feature.preview(),
            })
        })
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        self.file.finish(|| {
            let count = match &self.item {
                Item::Workbook => super::nav::tab_count(self.file)?,
                Item::NamedRange(_) | Item::Feature(_) => 0,
                Item::Cell(_) | Item::Range(_) | Item::Column(_) => {
                    if matches!(self.item, Item::Range(_)) {
                        Children::count(self, options)?
                    } else {
                        0
                    }
                },
                Item::Sheet(_) | Item::Row(_) => Children::count(self, options)?,
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
            Ok(children) => ViewChildren::xls(children),
            Err(error) => ViewChildren::failed(error),
        }
    }
}

fn row_rect(index: u32) -> Result<Rect> {
    Ok((index, 0, index, u32::MAX))
}

fn column_rect(index: u32) -> Result<Rect> {
    Ok((0, index, u32::MAX, index))
}

fn joined_text(grid: &Grid, rect: Rect) -> Result<String> {
    let mut out = String::new();
    for (index, (_, cell)) in grid.cells_in(rect).enumerate() {
        if index > 0 {
            out.push(' ');
        }
        out.push_str(&display_value(&cell.value));
    }
    Ok(out)
}
