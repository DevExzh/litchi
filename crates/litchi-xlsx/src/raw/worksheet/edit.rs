//! Minimal worksheet XML surgery for ordinary cell and row-property transactions.
//!
//! The scanner records exact byte ranges and regenerates only touched rows and
//! cells. Untouched XML, unknown worksheet children, extension payloads, and
//! lexical choices outside those narrow ranges remain byte-for-byte identical.

use std::collections::{BTreeMap, HashMap};

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::{COLUMNS, Cell as Address, Column, ROWS, Rect, Row};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{optional_bool, optional_u32, parse_a1, parse_one_based_row, required_u32};
use crate::cell::{Content, Value};
use crate::column::{Assignments, Width};
use crate::error::{ColumnEditBlock, EditBlock, Error, Result, RowEditBlock, invalid};
use crate::outline::Outline;
use crate::raw::namespace::is_spreadsheetml_name;
use crate::raw::strings::encode_spreadsheet_text;
use crate::row::Height;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Payload {
    Set(Content),
    /// Ensure an explicit empty cell record exists.
    Clear,
    /// Clear only when another effect or the base snapshot retains the cell.
    ClearIfPresent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StyleEffect {
    Set(u32),
    Reset,
}

/// Orthogonal effects on one cell record.
///
/// `Remove` owns the whole record. An `Update` may independently change its
/// payload and local style, allowing proven-disjoint effects to be joined.
#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Action {
    Update {
        payload: Option<Payload>,
        style: Option<StyleEffect>,
    },
    Remove,
}

impl Action {
    pub(crate) fn set(content: Content) -> Self {
        Self::Update {
            payload: Some(Payload::Set(content)),
            style: None,
        }
    }

    pub(crate) const fn clear(create: bool) -> Self {
        Self::Update {
            payload: Some(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
            style: None,
        }
    }

    pub(crate) const fn style(key: u32) -> Self {
        Self::Update {
            payload: None,
            style: Some(StyleEffect::Set(key)),
        }
    }

    pub(crate) const fn reset_style() -> Self {
        Self::Update {
            payload: None,
            style: Some(StyleEffect::Reset),
        }
    }

    pub(crate) const fn payload(&self) -> Option<&Payload> {
        match self {
            Self::Update { payload, .. } => payload.as_ref(),
            Self::Remove => None,
        }
    }

    pub(crate) const fn creates_missing(&self) -> bool {
        match self {
            Self::Update { payload, style } => {
                matches!(payload, Some(Payload::Set(_) | Payload::Clear))
                    || matches!(style, Some(StyleEffect::Set(_)))
            },
            Self::Remove => false,
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Remove, _) | (_, Self::Remove) => true,
            (
                Self::Update {
                    payload: left_payload,
                    style: left_style,
                },
                Self::Update {
                    payload: right_payload,
                    style: right_style,
                },
            ) => {
                (left_payload.is_some() && right_payload.is_some())
                    || (left_style.is_some() && right_style.is_some())
            },
        }
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if let (
            Self::Update { payload, style },
            Self::Update {
                payload: other_payload,
                style: other_style,
            },
        ) = (self, other)
        {
            // `Edit::join` proves these facets disjoint before moving either
            // map. The conditional assignments keep this primitive total and
            // panic-free if an internal caller ever violates that contract.
            if payload.is_none() {
                *payload = other_payload;
            }
            if style.is_none() {
                *style = other_style;
            }
        }
    }

    pub(crate) fn set_payload(&mut self, effect: Payload) {
        *self = match std::mem::replace(self, Self::Remove) {
            Self::Update { style, .. } => Self::Update {
                payload: Some(effect),
                style,
            },
            Self::Remove => Self::Update {
                payload: Some(effect),
                style: None,
            },
        };
    }

    pub(crate) fn set_style(&mut self, effect: StyleEffect) {
        *self = match std::mem::replace(self, Self::Remove) {
            Self::Update { payload, .. } => Self::Update {
                payload,
                style: Some(effect),
            },
            Self::Remove => Self::Update {
                payload: None,
                style: Some(effect),
            },
        };
    }
}

/// One checked height mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HeightEffect {
    Set(Height),
    Reset,
}

/// Orthogonal effects on one stored row record.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct RowAction {
    pub(crate) hidden: Option<bool>,
    pub(crate) height: Option<HeightEffect>,
    pub(crate) outline: Option<Outline>,
    pub(crate) collapsed: Option<bool>,
    pub(crate) thick_top: Option<bool>,
    pub(crate) thick_bottom: Option<bool>,
    pub(crate) phonetic: Option<bool>,
}

impl RowAction {
    #[cfg(test)]
    pub(crate) const fn hide() -> Self {
        Self {
            hidden: Some(true),
            height: None,
            outline: None,
            collapsed: None,
            thick_top: None,
            thick_bottom: None,
            phonetic: None,
        }
    }

    #[cfg(test)]
    pub(crate) const fn show() -> Self {
        Self {
            hidden: Some(false),
            height: None,
            outline: None,
            collapsed: None,
            thick_top: None,
            thick_bottom: None,
            phonetic: None,
        }
    }

    pub(crate) const fn materializes(self) -> bool {
        matches!(self.hidden, Some(true))
            || matches!(self.height, Some(HeightEffect::Set(_)))
            || matches!(self.outline, Some(level) if level.get() != 0)
            || matches!(self.collapsed, Some(true))
            || matches!(self.thick_top, Some(true))
            || matches!(self.thick_bottom, Some(true))
            || matches!(self.phonetic, Some(true))
    }

    pub(crate) const fn overlaps(self, other: Self) -> bool {
        (self.hidden.is_some() && other.hidden.is_some())
            || (self.height.is_some() && other.height.is_some())
            || (self.outline.is_some() && other.outline.is_some())
            || (self.collapsed.is_some() && other.collapsed.is_some())
            || (self.thick_top.is_some() && other.thick_top.is_some())
            || (self.thick_bottom.is_some() && other.thick_bottom.is_some())
            || (self.phonetic.is_some() && other.phonetic.is_some())
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if self.hidden.is_none() {
            self.hidden = other.hidden;
        }
        if self.height.is_none() {
            self.height = other.height;
        }
        if self.outline.is_none() {
            self.outline = other.outline;
        }
        if self.collapsed.is_none() {
            self.collapsed = other.collapsed;
        }
        if self.thick_top.is_none() {
            self.thick_top = other.thick_top;
        }
        if self.thick_bottom.is_none() {
            self.thick_bottom = other.thick_bottom;
        }
        if self.phonetic.is_none() {
            self.phonetic = other.phonetic;
        }
    }
}

/// One checked width mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum WidthEffect {
    Set(Width),
    Reset,
}

/// Orthogonal effects on one effective column-property record.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct ColumnAction {
    pub(crate) hidden: Option<bool>,
    pub(crate) width: Option<WidthEffect>,
    pub(crate) best_fit: Option<bool>,
    pub(crate) outline: Option<Outline>,
    pub(crate) collapsed: Option<bool>,
    pub(crate) phonetic: Option<bool>,
}

impl ColumnAction {
    #[cfg(test)]
    pub(crate) const fn hide() -> Self {
        Self {
            hidden: Some(true),
            width: None,
            best_fit: None,
            outline: None,
            collapsed: None,
            phonetic: None,
        }
    }

    #[cfg(test)]
    pub(crate) const fn show() -> Self {
        Self {
            hidden: Some(false),
            width: None,
            best_fit: None,
            outline: None,
            collapsed: None,
            phonetic: None,
        }
    }

    pub(crate) const fn materializes(self) -> bool {
        matches!(self.hidden, Some(true))
            || matches!(self.width, Some(WidthEffect::Set(_)))
            || matches!(self.best_fit, Some(true))
            || matches!(self.outline, Some(level) if level.get() != 0)
            || matches!(self.collapsed, Some(true))
            || matches!(self.phonetic, Some(true))
    }

    pub(crate) const fn overlaps(self, other: Self) -> bool {
        (self.hidden.is_some() && other.hidden.is_some())
            || (self.width.is_some() && other.width.is_some())
            || (self.best_fit.is_some() && other.best_fit.is_some())
            || (self.outline.is_some() && other.outline.is_some())
            || (self.collapsed.is_some() && other.collapsed.is_some())
            || (self.phonetic.is_some() && other.phonetic.is_some())
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if self.hidden.is_none() {
            self.hidden = other.hidden;
        }
        if self.width.is_none() {
            self.width = other.width;
        }
        if self.best_fit.is_none() {
            self.best_fit = other.best_fit;
        }
        if self.outline.is_none() {
            self.outline = other.outline;
        }
        if self.collapsed.is_none() {
            self.collapsed = other.collapsed;
        }
        if self.phonetic.is_none() {
            self.phonetic = other.phonetic;
        }
    }
}

/// Move-only worksheet rewrite plan with orthogonal cell and row facets.
#[derive(Debug, Default)]
pub(crate) struct Plan {
    pub(crate) cells: BTreeMap<Address, Action>,
    pub(crate) rows: BTreeMap<Row, RowAction>,
    pub(crate) columns: BTreeMap<Column, ColumnAction>,
}

impl Plan {
    pub(crate) fn cells(cells: BTreeMap<Address, Action>) -> Self {
        Self {
            cells,
            rows: BTreeMap::new(),
            columns: BTreeMap::new(),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.cells.is_empty() && self.rows.is_empty() && self.columns.is_empty()
    }
}

impl From<BTreeMap<Address, Action>> for Plan {
    fn from(value: BTreeMap<Address, Action>) -> Self {
        Self::cells(value)
    }
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug, Clone)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug)]
struct CellSlot {
    address: Address,
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    primary: Box<[Span]>,
    mce_payload: bool,
    empty: bool,
}

#[derive(Debug)]
struct RowSlot {
    number: u32,
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    cells: Box<[CellSlot]>,
    empty: bool,
}

#[derive(Debug)]
struct ColumnSlot {
    first: Column,
    last: Column,
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    payload: bool,
    empty: bool,
}

#[derive(Debug)]
struct ColumnsSlot {
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    columns: Box<[ColumnSlot]>,
    payload: bool,
    empty: bool,
}

#[derive(Debug)]
struct SheetData {
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    rows: Box<[RowSlot]>,
    empty: bool,
}

#[derive(Debug)]
struct DimensionTag {
    span: Span,
    tag: Tag,
    empty: bool,
    declared: Rect,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Worksheet,
    Columns,
    Column,
    SheetData,
    Row,
    Cell,
    Primary,
    Other,
}

#[derive(Debug, Clone, Copy)]
struct Frame {
    kind: FrameKind,
    start: usize,
}

#[derive(Debug)]
struct PendingCell {
    address: Address,
    start: usize,
    tag_end: usize,
    tag: Tag,
    primary: Vec<Span>,
    mce_payload: bool,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    last_column: u32,
    start: usize,
    tag_end: usize,
    tag: Tag,
    cells: Vec<CellSlot>,
}

#[derive(Debug)]
struct PendingColumn {
    first: Column,
    last: Column,
    start: usize,
    tag_end: usize,
    tag: Tag,
    payload: bool,
}

#[derive(Debug)]
struct PendingColumns {
    start: usize,
    tag_end: usize,
    tag: Tag,
    columns: Vec<ColumnSlot>,
    payload: bool,
}

#[derive(Debug)]
struct PendingSheetData {
    start: usize,
    tag_end: usize,
    tag: Tag,
    rows: Vec<RowSlot>,
}

#[derive(Debug, Clone, Copy)]
struct SelectionRange {
    first_row: u32,
    first_column: u32,
    last_row: u32,
    last_column: u32,
}

impl SelectionRange {
    fn cell_or_area(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        if last.contains(':') {
            return Err(invalid(format!("invalid cell range '{value}'")));
        }
        let first = Address::from_a1(first)?;
        let last = Address::from_a1(last)?;
        if first.row() > last.row() || first.column() > last.column() {
            return Err(invalid(format!("reversed cell range '{value}'")));
        }
        Ok(Self {
            first_row: first.row().get(),
            first_column: first.column().get(),
            last_row: last.row().get(),
            last_column: last.column().get(),
        })
    }

    fn selection(value: &str) -> Result<Self> {
        if let Ok(range) = Self::cell_or_area(value) {
            return Ok(range);
        }
        let (first, last) = value
            .split_once(':')
            .ok_or_else(|| invalid(format!("invalid selection range '{value}'")))?;
        if first
            .bytes()
            .all(|byte| byte == b'$' || byte.is_ascii_alphabetic())
            && last
                .bytes()
                .all(|byte| byte == b'$' || byte.is_ascii_alphabetic())
        {
            let first = column(first)?;
            let last = column(last)?;
            if first > last {
                return Err(invalid(format!("reversed column range '{value}'")));
            }
            return Ok(Self {
                first_row: 0,
                first_column: first,
                last_row: ROWS - 1,
                last_column: last,
            });
        }
        if first
            .bytes()
            .all(|byte| byte == b'$' || byte.is_ascii_digit())
            && last
                .bytes()
                .all(|byte| byte == b'$' || byte.is_ascii_digit())
        {
            let first = row(first)?;
            let last = row(last)?;
            if first > last {
                return Err(invalid(format!("reversed row range '{value}'")));
            }
            return Ok(Self {
                first_row: first,
                first_column: 0,
                last_row: last,
                last_column: COLUMNS - 1,
            });
        }
        Err(invalid(format!("invalid selection range '{value}'")))
    }

    fn contains(self, address: Address) -> bool {
        (self.first_row..=self.last_row).contains(&address.row().get())
            && (self.first_column..=self.last_column).contains(&address.column().get())
    }

    fn starts_at(self, address: Address) -> bool {
        self.first_row == address.row().get() && self.first_column == address.column().get()
    }
}

#[derive(Debug)]
struct FormulaStorage {
    address: Address,
    kind: Box<str>,
    index: Option<u32>,
    range: Option<SelectionRange>,
}

#[derive(Debug)]
struct Layout {
    sheet_data: SheetData,
    columns: Option<ColumnsSlot>,
    dimension: Option<DimensionTag>,
    protected: bool,
    merged: Box<[SelectionRange]>,
    validations: Box<[SelectionRange]>,
    extended_validation: bool,
    formula_ranges: Box<[SelectionRange]>,
}

#[derive(Debug, Default)]
struct Scanner {
    sheet_data: Option<SheetData>,
    columns: Option<ColumnsSlot>,
    dimension: Option<DimensionTag>,
    pending_sheet_data: Option<PendingSheetData>,
    pending_columns: Option<PendingColumns>,
    column: Option<PendingColumn>,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    previous_row: u32,
    protected: bool,
    merged: Vec<SelectionRange>,
    validations: Vec<SelectionRange>,
    extended_validation: bool,
    formulas: Vec<FormulaStorage>,
}

pub(crate) fn rewrite(content: &[u8], sheet: &str, plan: impl Into<Plan>) -> Result<Vec<u8>> {
    let plan = plan.into();
    if plan.is_empty() {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    validate_actions(&layout, sheet, &plan.cells)?;
    validate_row_actions(&layout, sheet, &plan.rows)?;
    validate_column_actions(&layout, sheet, &plan.columns)?;
    let dimension = expanded_dimension(&layout, &plan.cells);

    let effects = plan
        .cells
        .len()
        .checked_add(plan.rows.len())
        .and_then(|count| count.checked_add(plan.columns.len()))
        .ok_or_else(|| invalid("worksheet edit effect count overflow"))?;
    let extra = effects
        .checked_mul(128)
        .and_then(|value| content.len().checked_add(value))
        .ok_or_else(|| invalid("worksheet edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve(extra)
        .map_err(|error| invalid(format!("cannot reserve worksheet edit output: {error}")))?;
    let Plan {
        cells,
        rows,
        columns,
    } = plan;
    let mut cursor = 0usize;
    if let Some((tag, range)) = dimension {
        output.extend_from_slice(&content[cursor..tag.span.start]);
        write_tag(
            &mut output,
            &tag.tag,
            tag.empty,
            &["ref"],
            &[("ref", range.a1())],
        );
        cursor = tag.span.end;
    }
    if !columns.is_empty() {
        match layout.columns.as_ref() {
            Some(stored) => {
                output.extend_from_slice(&content[cursor..stored.span.start]);
                write_columns(&mut output, content, stored, columns, sheet)?;
                cursor = stored.span.end;
            },
            None => {
                output.extend_from_slice(&content[cursor..layout.sheet_data.span.start]);
                write_new_columns(&mut output, &layout.sheet_data.tag.name, columns);
                cursor = layout.sheet_data.span.start;
            },
        }
    }
    output.extend_from_slice(&content[cursor..layout.sheet_data.span.start]);
    if cells.is_empty() && rows.is_empty() {
        output
            .extend_from_slice(&content[layout.sheet_data.span.start..layout.sheet_data.span.end]);
    } else {
        write_sheet_data(&mut output, content, &layout.sheet_data, cells, rows)?;
    }
    output.extend_from_slice(&content[layout.sheet_data.span.end..]);
    Ok(output)
}

fn validate_column_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Column, ColumnAction>,
) -> Result<()> {
    if layout.protected
        && let Some(column) = actions.keys().next()
    {
        return Err(Error::ColumnEditBlocked {
            sheet: sheet.to_owned(),
            column: *column,
            reason: ColumnEditBlock::ProtectedSheet,
        });
    }
    Ok(())
}

fn validate_row_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Row, RowAction>,
) -> Result<()> {
    if layout.protected
        && let Some(row) = actions.keys().next()
    {
        return Err(Error::RowEditBlocked {
            sheet: sheet.to_owned(),
            row: *row,
            reason: RowEditBlock::ProtectedSheet,
        });
    }
    Ok(())
}

fn validate_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Address, Action>,
) -> Result<()> {
    for (address, action) in actions {
        let blocked = if layout.protected {
            Some(EditBlock::ProtectedSheet)
        } else if layout.extended_validation
            || layout
                .validations
                .iter()
                .any(|range| range.contains(*address))
        {
            Some(EditBlock::DataValidation)
        } else if layout
            .formula_ranges
            .iter()
            .any(|range| range.contains(*address))
        {
            Some(EditBlock::GroupFormula)
        } else if layout
            .merged
            .iter()
            .any(|range| range.contains(*address) && !range.starts_at(*address))
        {
            Some(EditBlock::CoveredMerge)
        } else if cell_slot(&layout.sheet_data, *address).is_some_and(|cell| cell.mce_payload) {
            Some(EditBlock::MarkupCompatibility)
        } else {
            None
        };
        if let Some(reason) = blocked {
            return Err(Error::EditBlocked {
                sheet: sheet.to_owned(),
                address: *address,
                reason,
            });
        }
        if let Some(Payload::Set(content)) = action.payload() {
            content.validate_for_write()?;
        }
    }
    Ok(())
}

fn cell_slot(sheet_data: &SheetData, address: Address) -> Option<&CellSlot> {
    let row = sheet_data
        .rows
        .binary_search_by_key(&(address.row().get() + 1), |row| row.number)
        .ok()
        .and_then(|index| sheet_data.rows.get(index))?;
    row.cells
        .binary_search_by_key(&address, |cell| cell.address)
        .ok()
        .and_then(|index| row.cells.get(index))
}

fn scan(content: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(content);
    let mut scanner = Scanner::default();
    let mut stack = Vec::<Frame>::new();

    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                let kind = scanner.start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    event_start,
                    event_end,
                )?;
                stack.push(Frame {
                    kind,
                    start: event_start,
                });
            },
            Event::Empty(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                scanner.empty(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet edit scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("worksheet edit scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    fn start(
        &mut self,
        parent: Option<FrameKind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        end: usize,
    ) -> Result<FrameKind> {
        self.scan_guard(namespace, element, decoder)?;
        if parent.is_none() && is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
            return Ok(FrameKind::Worksheet);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"dimension")
        {
            self.record_dimension(element, decoder, Span { start, end }, false)?;
            return Ok(FrameKind::Other);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            if self.pending_columns.is_some() || self.columns.is_some() {
                return Err(invalid("worksheet has duplicate cols during edit"));
            }
            if self.sheet_data.is_some() || self.pending_sheet_data.is_some() {
                return Err(invalid(
                    "worksheet cols appears after sheetData during edit",
                ));
            }
            self.pending_columns = Some(PendingColumns {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                columns: Vec::new(),
                payload: false,
            });
            return Ok(FrameKind::Columns);
        }
        if parent == Some(FrameKind::Columns)
            && is_spreadsheetml_name(namespace, element.name(), b"col")
        {
            let (first, last) = column_range(element, decoder)?;
            self.column = Some(PendingColumn {
                first,
                last,
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                payload: false,
            });
            return Ok(FrameKind::Column);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            if self.pending_sheet_data.is_some() || self.sheet_data.is_some() {
                return Err(invalid("worksheet has duplicate sheetData during edit"));
            }
            self.pending_sheet_data = Some(PendingSheetData {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                rows: Vec::new(),
            });
            return Ok(FrameKind::SheetData);
        }
        if parent == Some(FrameKind::SheetData)
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            self.start_row(element, decoder, start, end)?;
            return Ok(FrameKind::Row);
        }
        if parent == Some(FrameKind::Row) && is_spreadsheetml_name(namespace, element.name(), b"c")
        {
            self.start_cell(element, decoder, start, end)?;
            return Ok(FrameKind::Cell);
        }
        if parent == Some(FrameKind::Cell)
            && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
            && is_spreadsheetml_name(
                namespace,
                element.name(),
                element.name().local_name().as_ref(),
            )
        {
            if element.name().local_name().as_ref() == b"f" {
                self.scan_formula(element, decoder)?;
            }
            return Ok(FrameKind::Primary);
        }
        if self.cell.is_some()
            && (is_mce_name(namespace, element, b"AlternateContent")
                || (parent == Some(FrameKind::Cell)
                    && !is_spreadsheetml_name(namespace, element.name(), b"extLst")))
        {
            // Unknown direct cell children may carry future value semantics.
            // Keeping them beside a replacement payload could silently create
            // two competing representations, so the ordinary editor refuses.
            if let Some(cell) = self.cell.as_mut() {
                cell.mce_payload = true;
            }
        }
        if parent == Some(FrameKind::Column)
            && let Some(column) = self.column.as_mut()
        {
            column.payload = true;
        }
        if parent == Some(FrameKind::Columns)
            && let Some(columns) = self.pending_columns.as_mut()
        {
            columns.payload = true;
        }
        Ok(FrameKind::Other)
    }

    fn empty(
        &mut self,
        parent: Option<FrameKind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        span: Span,
    ) -> Result<()> {
        self.scan_guard(namespace, element, decoder)?;
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"dimension")
        {
            self.record_dimension(element, decoder, span, true)?;
            return Ok(());
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            if self.pending_columns.is_some() || self.columns.is_some() {
                return Err(invalid("worksheet has duplicate cols during edit"));
            }
            if self.sheet_data.is_some() || self.pending_sheet_data.is_some() {
                return Err(invalid(
                    "worksheet cols appears after sheetData during edit",
                ));
            }
            self.columns = Some(ColumnsSlot {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                columns: Box::new([]),
                payload: false,
                empty: true,
            });
            return Ok(());
        }
        if parent == Some(FrameKind::Columns)
            && is_spreadsheetml_name(namespace, element.name(), b"col")
        {
            let (first, last) = column_range(element, decoder)?;
            self.pending_columns
                .as_mut()
                .ok_or_else(|| invalid("empty col outside cols"))?
                .columns
                .push(ColumnSlot {
                    first,
                    last,
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    payload: false,
                    empty: true,
                });
            return Ok(());
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            if self.pending_sheet_data.is_some() || self.sheet_data.is_some() {
                return Err(invalid("worksheet has duplicate sheetData during edit"));
            }
            self.sheet_data = Some(SheetData {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                rows: Box::new([]),
                empty: true,
            });
            return Ok(());
        }
        if parent == Some(FrameKind::SheetData)
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            let (number, _) = self.row_position(element, decoder)?;
            let row = RowSlot {
                number,
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                cells: Box::new([]),
                empty: true,
            };
            self.pending_sheet_data
                .as_mut()
                .ok_or_else(|| invalid("empty row outside sheetData"))?
                .rows
                .push(row);
            return Ok(());
        }
        if parent == Some(FrameKind::Row) && is_spreadsheetml_name(namespace, element.name(), b"c")
        {
            let address = self.cell_address(element, decoder)?;
            self.row
                .as_mut()
                .ok_or_else(|| invalid("empty cell outside row"))?
                .cells
                .push(CellSlot {
                    address,
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    primary: Box::new([]),
                    mce_payload: false,
                    empty: true,
                });
            return Ok(());
        }
        if parent == Some(FrameKind::Cell)
            && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
            && is_spreadsheetml_name(
                namespace,
                element.name(),
                element.name().local_name().as_ref(),
            )
        {
            if element.name().local_name().as_ref() == b"f" {
                self.scan_formula(element, decoder)?;
            }
            self.cell
                .as_mut()
                .ok_or_else(|| invalid("empty cell payload outside cell"))?
                .primary
                .push(span);
        } else if self.cell.is_some()
            && (is_mce_name(namespace, element, b"AlternateContent")
                || (parent == Some(FrameKind::Cell)
                    && !is_spreadsheetml_name(namespace, element.name(), b"extLst")))
            && let Some(cell) = self.cell.as_mut()
        {
            cell.mce_payload = true;
        }
        if parent == Some(FrameKind::Column)
            && let Some(column) = self.column.as_mut()
        {
            column.payload = true;
        }
        if parent == Some(FrameKind::Columns)
            && let Some(columns) = self.pending_columns.as_mut()
        {
            columns.payload = true;
        }
        Ok(())
    }

    fn finish(&mut self, frame: Frame, close_start: usize, end: usize) -> Result<()> {
        match frame.kind {
            FrameKind::Column => {
                let column = self
                    .column
                    .take()
                    .ok_or_else(|| invalid("col close without edit state"))?;
                self.pending_columns
                    .as_mut()
                    .ok_or_else(|| invalid("col closed outside cols"))?
                    .columns
                    .push(ColumnSlot {
                        first: column.first,
                        last: column.last,
                        span: Span {
                            start: column.start,
                            end,
                        },
                        tag_end: column.tag_end,
                        close_start,
                        tag: column.tag,
                        payload: column.payload,
                        empty: false,
                    });
            },
            FrameKind::Columns => {
                let columns = self
                    .pending_columns
                    .take()
                    .ok_or_else(|| invalid("cols close without edit state"))?;
                if columns.columns.is_empty() {
                    return Err(invalid("worksheet cols contains no col during edit"));
                }
                self.columns = Some(ColumnsSlot {
                    span: Span {
                        start: columns.start,
                        end,
                    },
                    tag_end: columns.tag_end,
                    close_start,
                    tag: columns.tag,
                    columns: columns.columns.into_boxed_slice(),
                    payload: columns.payload,
                    empty: false,
                });
            },
            FrameKind::Primary => {
                self.cell
                    .as_mut()
                    .ok_or_else(|| invalid("cell payload closed outside a cell"))?
                    .primary
                    .push(Span {
                        start: frame.start,
                        end,
                    });
            },
            FrameKind::Cell => {
                let cell = self
                    .cell
                    .take()
                    .ok_or_else(|| invalid("cell close without edit state"))?;
                self.row
                    .as_mut()
                    .ok_or_else(|| invalid("cell closed outside a row"))?
                    .cells
                    .push(CellSlot {
                        address: cell.address,
                        span: Span {
                            start: cell.start,
                            end,
                        },
                        tag_end: cell.tag_end,
                        close_start,
                        tag: cell.tag,
                        primary: cell.primary.into_boxed_slice(),
                        mce_payload: cell.mce_payload,
                        empty: false,
                    });
            },
            FrameKind::Row => {
                let row = self
                    .row
                    .take()
                    .ok_or_else(|| invalid("row close without edit state"))?;
                if row
                    .cells
                    .windows(2)
                    .any(|pair| pair[0].address >= pair[1].address)
                {
                    return Err(invalid(
                        "cell edits require strictly increasing cell references within each row",
                    ));
                }
                self.pending_sheet_data
                    .as_mut()
                    .ok_or_else(|| invalid("row closed outside sheetData"))?
                    .rows
                    .push(RowSlot {
                        number: row.number,
                        span: Span {
                            start: row.start,
                            end,
                        },
                        tag_end: row.tag_end,
                        close_start,
                        tag: row.tag,
                        cells: row.cells.into_boxed_slice(),
                        empty: false,
                    });
            },
            FrameKind::SheetData => {
                let data = self
                    .pending_sheet_data
                    .take()
                    .ok_or_else(|| invalid("sheetData close without edit state"))?;
                self.sheet_data = Some(SheetData {
                    span: Span {
                        start: data.start,
                        end,
                    },
                    tag_end: data.tag_end,
                    close_start,
                    tag: data.tag,
                    rows: data.rows.into_boxed_slice(),
                    empty: false,
                });
            },
            _ => {},
        }
        Ok(())
    }

    fn start_row(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let (number, last_column) = self.row_position(element, decoder)?;
        self.row = Some(PendingRow {
            number,
            last_column,
            start,
            tag_end: end,
            tag: tag(element, decoder)?,
            cells: Vec::new(),
        });
        Ok(())
    }

    fn row_position(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<(u32, u32)> {
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => self
                .previous_row
                .checked_add(1)
                .filter(|number| *number <= ROWS)
                .ok_or_else(|| invalid("inferred edit row exceeds the grid"))?,
        };
        if self.previous_row != 0 && number <= self.previous_row {
            return Err(invalid(
                "cell edits require strictly increasing worksheet rows",
            ));
        }
        self.previous_row = number;
        Ok((number, 0))
    }

    fn start_cell(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let address = self.cell_address(element, decoder)?;
        self.cell = Some(PendingCell {
            address,
            start,
            tag_end: end,
            tag: tag(element, decoder)?,
            primary: Vec::new(),
            mce_payload: false,
        });
        Ok(())
    }

    fn cell_address(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<Address> {
        let row = self
            .row
            .as_ref()
            .ok_or_else(|| invalid("cell outside edit row"))?
            .number;
        let column = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(reference) => {
                let (reference_row, column) = parse_a1(&reference)?;
                if reference_row != row {
                    return Err(invalid(format!(
                        "cell reference '{reference}' does not belong to row {row}"
                    )));
                }
                column
            },
            None => self
                .row
                .as_ref()
                .and_then(|row| row.last_column.checked_add(1))
                .filter(|column| *column <= COLUMNS)
                .ok_or_else(|| invalid("inferred edit column exceeds the grid"))?,
        };
        let pending = self
            .row
            .as_mut()
            .ok_or_else(|| invalid("cell outside edit row"))?;
        pending.last_column = column;
        Address::at(row - 1, column - 1).map_err(Into::into)
    }

    fn scan_formula(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let kind = unqualified_attribute_value(element, b"t", decoder)?
            .unwrap_or_else(|| "normal".to_owned());
        if !matches!(kind.as_str(), "shared" | "array" | "dataTable") {
            return Ok(());
        }
        let address = self
            .cell
            .as_ref()
            .ok_or_else(|| invalid("formula outside edit cell"))?
            .address;
        let range = unqualified_attribute_value(element, b"ref", decoder)?
            .map(|value| SelectionRange::cell_or_area(&value))
            .transpose()?;
        let index = if kind == "shared" {
            optional_u32(element, b"si", decoder, "shared formula index")?
        } else {
            None
        };
        self.formulas.push(FormulaStorage {
            address,
            kind: kind.into_boxed_str(),
            index,
            range,
        });
        Ok(())
    }

    fn scan_guard(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        if is_spreadsheetml_name(namespace, element.name(), b"sheetProtection") {
            self.protected |= optional_bool(element, b"sheet", decoder, "sheet protection flag")?
                .unwrap_or(false);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCell") {
            let value = unqualified_attribute_value(element, b"ref", decoder)?
                .ok_or_else(|| invalid("mergeCell is missing ref during edit"))?;
            self.merged.push(SelectionRange::cell_or_area(&value)?);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"dataValidation") {
            let value = unqualified_attribute_value(element, b"sqref", decoder)?
                .ok_or_else(|| invalid("dataValidation is missing sqref during edit"))?;
            for token in value.split_whitespace() {
                self.validations.push(SelectionRange::selection(token)?);
            }
        }
        if element.name().local_name().as_ref() == b"dataValidation"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == X14)
        {
            self.extended_validation = true;
        }
        Ok(())
    }

    fn record_dimension(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        span: Span,
        empty: bool,
    ) -> Result<()> {
        if self.dimension.is_some() {
            return Err(invalid(
                "worksheet has duplicate dimension elements during edit",
            ));
        }
        let reference = unqualified_attribute_value(element, b"ref", decoder)?
            .ok_or_else(|| invalid("worksheet dimension is missing ref during edit"))?;
        let declared = Rect::from_a1(&reference).map_err(|error| {
            invalid(format!(
                "invalid worksheet dimension '{reference}' during edit: {error}"
            ))
        })?;
        self.dimension = Some(DimensionTag {
            span,
            tag: tag(element, decoder)?,
            empty,
            declared,
        });
        Ok(())
    }

    fn finish_layout(self) -> Result<Layout> {
        let sheet_data = self
            .sheet_data
            .ok_or_else(|| invalid("worksheet cell edits require a direct sheetData element"))?;
        if self
            .columns
            .as_ref()
            .is_some_and(|columns| columns.columns.is_empty())
        {
            return Err(invalid("worksheet cols contains no col during edit"));
        }
        if self
            .columns
            .as_ref()
            .is_some_and(|columns| columns.span.start >= sheet_data.span.start)
        {
            return Err(invalid("worksheet cols must precede sheetData during edit"));
        }
        if self
            .dimension
            .as_ref()
            .is_some_and(|dimension| dimension.span.start >= sheet_data.span.start)
        {
            return Err(invalid(
                "worksheet dimension must precede sheetData during cell edits",
            ));
        }
        let mut formula_ranges = Vec::new();
        let mut shared = HashMap::<u32, SelectionRange>::new();
        for formula in &self.formulas {
            match formula.kind.as_ref() {
                "array" | "dataTable" => {
                    formula_ranges.push(formula.range.unwrap_or(SelectionRange {
                        first_row: formula.address.row().get(),
                        first_column: formula.address.column().get(),
                        last_row: formula.address.row().get(),
                        last_column: formula.address.column().get(),
                    }))
                },
                "shared" => {
                    if let (Some(index), Some(range)) = (formula.index, formula.range) {
                        shared.insert(index, range);
                    }
                },
                _ => {},
            }
        }
        for formula in &self.formulas {
            if formula.kind.as_ref() == "shared" {
                formula_ranges.push(
                    formula
                        .index
                        .and_then(|index| shared.get(&index).copied())
                        .unwrap_or(SelectionRange {
                            first_row: formula.address.row().get(),
                            first_column: formula.address.column().get(),
                            last_row: formula.address.row().get(),
                            last_column: formula.address.column().get(),
                        }),
                );
            }
        }
        Ok(Layout {
            sheet_data,
            columns: self.columns,
            dimension: self.dimension,
            protected: self.protected,
            merged: self.merged.into_boxed_slice(),
            validations: self.validations.into_boxed_slice(),
            extended_validation: self.extended_validation,
            formula_ranges: formula_ranges.into_boxed_slice(),
        })
    }
}

#[derive(Debug, Default)]
struct CellBounds(Option<Rect>);

impl CellBounds {
    fn push(&mut self, address: Address) {
        let cell = Rect::single(address);
        self.0 = Some(self.0.map_or(cell, |range| range.union(cell)));
    }
}

fn expanded_dimension<'a>(
    layout: &'a Layout,
    actions: &BTreeMap<Address, Action>,
) -> Option<(&'a DimensionTag, Rect)> {
    let dimension = layout.dimension.as_ref()?;
    let mut bounds = CellBounds::default();
    for row in &layout.sheet_data.rows {
        for cell in &row.cells {
            if !matches!(actions.get(&cell.address), Some(Action::Remove)) {
                bounds.push(cell.address);
            }
        }
    }
    for (address, action) in actions {
        if action.creates_missing() {
            bounds.push(*address);
        }
    }
    let result = bounds.0?;
    let expanded = dimension.declared.union(result);
    (expanded != dimension.declared).then_some((dimension, expanded))
}

#[derive(Debug, Clone, Copy)]
enum ColumnPiece {
    Keep(Column, Column),
    Edit(Column, Column, ColumnAction),
}

fn write_columns(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnsSlot,
    actions: BTreeMap<Column, ColumnAction>,
    sheet: &str,
) -> Result<()> {
    let mut owners = Assignments::new()?;
    for (index, column) in stored.columns.iter().enumerate() {
        owners.assign(column.first, column.last, index);
    }
    let mut by_owner = HashMap::<usize, BTreeMap<Column, ColumnAction>>::new();
    let mut implicit = BTreeMap::new();
    for (column, action) in actions {
        if let Some(owner) = owners.get(column) {
            by_owner.entry(owner).or_default().insert(column, action);
        } else if action.materializes() {
            implicit.insert(column, action);
        }
    }

    if stored.payload
        && let Some(column) = implicit.keys().next()
    {
        return Err(Error::ColumnEditBlocked {
            sheet: sheet.to_owned(),
            column: *column,
            reason: ColumnEditBlock::MarkupCompatibility,
        });
    }

    if stored.empty {
        return Err(invalid("worksheet cols contains no col during edit"));
    }
    output.extend_from_slice(&source[stored.span.start..stored.tag_end]);
    let mut cursor = stored.tag_end;
    for (index, column) in stored.columns.iter().enumerate() {
        output.extend_from_slice(&source[cursor..column.span.start]);
        if let Some(edits) = by_owner.remove(&index) {
            let pieces = column_pieces(column, &edits)?;
            if column.payload && pieces.len() > 1 {
                let edited = edits.keys().next().copied().unwrap_or(column.first);
                return Err(Error::ColumnEditBlocked {
                    sheet: sheet.to_owned(),
                    column: edited,
                    reason: ColumnEditBlock::MarkupCompatibility,
                });
            }
            for piece in pieces {
                write_column_piece(output, source, column, piece);
            }
        } else {
            output.extend_from_slice(&source[column.span.start..column.span.end]);
        }
        cursor = column.span.end;
    }
    output.extend_from_slice(&source[cursor..stored.close_start]);
    write_column_actions(output, &stored.tag.name, implicit);
    output.extend_from_slice(&source[stored.close_start..stored.span.end]);
    Ok(())
}

fn column_pieces(
    stored: &ColumnSlot,
    edits: &BTreeMap<Column, ColumnAction>,
) -> Result<Vec<ColumnPiece>> {
    let capacity = edits
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("column edit split count overflow"))?;
    let mut pieces = Vec::new();
    pieces
        .try_reserve_exact(capacity)
        .map_err(|error| invalid(format!("cannot reserve column edit splits: {error}")))?;
    let mut next = stored.first.get();
    for (column, action) in edits {
        if column.get() > next {
            pieces.push(ColumnPiece::Keep(
                Column::new(next)?,
                Column::new(column.get() - 1)?,
            ));
        }
        if let Some(ColumnPiece::Edit(_, last, previous)) = pieces.last_mut()
            && previous == action
            && last.next() == Some(*column)
        {
            *last = *column;
        } else {
            pieces.push(ColumnPiece::Edit(*column, *column, *action));
        }
        next = column.get().saturating_add(1);
    }
    if next <= stored.last.get() {
        pieces.push(ColumnPiece::Keep(Column::new(next)?, stored.last));
    }
    Ok(pieces)
}

fn write_column_piece(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnSlot,
    piece: ColumnPiece,
) {
    let (first, last, action) = match piece {
        ColumnPiece::Keep(first, last) => (first, last, None),
        ColumnPiece::Edit(first, last, action) => (first, last, Some(action)),
    };
    let mut removed = vec!["min", "max"];
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    if let Some(action) = action {
        column_effect_attributes(action, &mut removed, &mut appended);
    }
    write_tag(output, &stored.tag, stored.empty, &removed, &appended);
    if !stored.empty {
        output.extend_from_slice(&source[stored.tag_end..stored.close_start]);
        write_close(output, &stored.tag.name);
    }
}

fn write_new_columns(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    if !actions.values().any(|action| action.materializes()) {
        return;
    }
    let name = sibling_name(sheet_data_name, "cols");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    write_tag(output, &tag, false, &[], &[]);
    write_column_actions(output, &name, actions);
    write_close(output, &name);
}

fn write_column_actions(
    output: &mut Vec<u8>,
    columns_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    let name = sibling_name(columns_name, "col");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut pending: Option<(Column, Column, ColumnAction)> = None;
    for (column, action) in actions {
        if !action.materializes() {
            continue;
        }
        match pending {
            Some((first, last, previous)) if previous == action && last.next() == Some(column) => {
                pending = Some((first, column, action));
            },
            Some((first, last, previous)) => {
                write_new_column(output, &tag, first, last, previous);
                pending = Some((column, column, action));
            },
            None => pending = Some((column, column, action)),
        }
    }
    if let Some((first, last, action)) = pending {
        write_new_column(output, &tag, first, last, action);
    }
}

fn write_new_column(
    output: &mut Vec<u8>,
    tag: &Tag,
    first: Column,
    last: Column,
    action: ColumnAction,
) {
    let mut removed = Vec::new();
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    column_effect_attributes(action, &mut removed, &mut appended);
    write_tag(output, tag, true, &removed, &appended);
}

fn column_effect_attributes(
    action: ColumnAction,
    removed: &mut Vec<&'static str>,
    appended: &mut Vec<(&'static str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(width) = action.width {
        removed.extend(["width", "customWidth"]);
        if let WidthEffect::Set(width) = width {
            appended.push(("width", width.get().to_string()));
            appended.push(("customWidth", "1".to_owned()));
        }
    }
    if let Some(best_fit) = action.best_fit {
        removed.push("bestFit");
        if best_fit {
            appended.push(("bestFit", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    if let Some(collapsed) = action.collapsed {
        removed.push("collapsed");
        if collapsed {
            appended.push(("collapsed", "1".to_owned()));
        }
    }
    if let Some(phonetic) = action.phonetic {
        removed.push("phonetic");
        if phonetic {
            appended.push(("phonetic", "1".to_owned()));
        }
    }
}

fn write_sheet_data(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &SheetData,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<Row, RowAction>,
) -> Result<()> {
    let mut by_row = BTreeMap::<u32, RowEdits>::new();
    for (address, action) in cells {
        by_row
            .entry(address.row().get() + 1)
            .or_default()
            .cells
            .insert(address, action);
    }
    for (row, action) in rows {
        by_row.entry(row.get() + 1).or_default().row = Some(action);
    }

    if data.empty {
        write_tag(output, &data.tag, false, &[], &[]);
        for (number, edits) in by_row {
            write_new_row(output, &data.tag.name, number, &edits)?;
        }
        write_close(output, &data.tag.name);
        return Ok(());
    }

    output.extend_from_slice(&source[data.span.start..data.tag_end]);
    let mut cursor = data.tag_end;
    let mut pending = by_row.into_iter().peekable();
    for row in &data.rows {
        output.extend_from_slice(&source[cursor..row.span.start]);
        while pending
            .peek()
            .is_some_and(|(number, _)| *number < row.number)
        {
            if let Some((number, edits)) = pending.next() {
                write_new_row(output, &data.tag.name, number, &edits)?;
            }
        }
        if pending
            .peek()
            .is_some_and(|(number, _)| *number == row.number)
        {
            let (_, edits) = pending
                .next()
                .ok_or_else(|| invalid("worksheet row edit ordering was lost"))?;
            write_row(output, source, row, &edits)?;
        } else {
            output.extend_from_slice(&source[row.span.start..row.span.end]);
        }
        cursor = row.span.end;
    }
    output.extend_from_slice(&source[cursor..data.close_start]);
    for (number, edits) in pending {
        write_new_row(output, &data.tag.name, number, &edits)?;
    }
    output.extend_from_slice(&source[data.close_start..data.span.end]);
    Ok(())
}

#[derive(Debug, Default)]
struct RowEdits {
    cells: BTreeMap<Address, Action>,
    row: Option<RowAction>,
}

fn write_row(output: &mut Vec<u8>, source: &[u8], row: &RowSlot, edits: &RowEdits) -> Result<()> {
    let actions = &edits.cells;
    let membership_changed = actions.iter().any(|(address, action)| {
        let exists = row
            .cells
            .binary_search_by_key(address, |cell| cell.address)
            .is_ok();
        (!exists && action.creates_missing()) || (exists && matches!(action, Action::Remove))
    });

    if row.empty {
        let creates_cell = actions.values().any(Action::creates_missing);
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if creates_cell {
            removed.extend(["spans", "r"]);
            appended.push(("r", row.number.to_string()));
        }
        if let Some(action) = edits.row {
            row_effect_attributes(action, &mut removed, &mut appended);
        }
        write_tag(output, &row.tag, !creates_cell, &removed, &appended);
        if !creates_cell {
            return Ok(());
        }
        for (address, action) in actions {
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        write_close(output, &row.tag.name);
        return Ok(());
    }

    if membership_changed || edits.row.is_some() {
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if membership_changed {
            removed.push("spans");
        }
        if let Some(action) = edits.row {
            row_effect_attributes(action, &mut removed, &mut appended);
        }
        write_tag(output, &row.tag, false, &removed, &appended);
    } else {
        output.extend_from_slice(&source[row.span.start..row.tag_end]);
    }
    let mut cursor = row.tag_end;
    let mut pending = actions.iter().peekable();
    for cell in &row.cells {
        output.extend_from_slice(&source[cursor..cell.span.start]);
        while pending
            .peek()
            .is_some_and(|(address, _)| **address < cell.address)
        {
            let (address, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        if pending
            .peek()
            .is_some_and(|(address, _)| **address == cell.address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            match action {
                Action::Update { .. } => write_cell(output, source, cell, action)?,
                Action::Remove => {},
            }
        } else {
            output.extend_from_slice(&source[cell.span.start..cell.span.end]);
        }
        cursor = cell.span.end;
    }
    output.extend_from_slice(&source[cursor..row.close_start]);
    for (address, action) in pending {
        write_new_action(output, &row.tag.name, *address, action)?;
    }
    output.extend_from_slice(&source[row.close_start..row.span.end]);
    Ok(())
}

fn write_new_row(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    number: u32,
    edits: &RowEdits,
) -> Result<()> {
    let creates_cell = edits.cells.values().any(Action::creates_missing);
    let materializes = edits.row.is_some_and(RowAction::materializes);
    if !creates_cell && !materializes {
        return Ok(());
    }
    let name = sibling_name(sheet_data_name, "row");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", number.to_string())];
    let mut removed = Vec::new();
    if let Some(action) = edits.row {
        row_effect_attributes(action, &mut removed, &mut appended);
    }
    write_tag(output, &tag, !creates_cell, &removed, &appended);
    if !creates_cell {
        return Ok(());
    }
    for (address, action) in &edits.cells {
        write_new_action(output, &name, *address, action)?;
    }
    write_close(output, &name);
    Ok(())
}

fn row_effect_attributes(
    action: RowAction,
    removed: &mut Vec<&'static str>,
    appended: &mut Vec<(&'static str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(height) = action.height {
        removed.extend(["ht", "customHeight"]);
        if let HeightEffect::Set(height) = height {
            appended.push(("ht", height.get().to_string()));
            appended.push(("customHeight", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    for (value, name) in [
        (action.collapsed, "collapsed"),
        (action.thick_top, "thickTop"),
        (action.thick_bottom, "thickBot"),
        (action.phonetic, "ph"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
}

fn write_cell(output: &mut Vec<u8>, source: &[u8], cell: &CellSlot, action: &Action) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Err(invalid("cannot rewrite a removed cell"));
    };
    let content = match payload.as_ref() {
        Some(Payload::Set(content)) => Some(content),
        Some(Payload::Clear | Payload::ClearIfPresent) | None => None,
    };
    let cell_type = content.and_then(content_type);
    let mut removed = vec!["r"];
    if payload.is_some() {
        removed.push("t");
    }
    if style.is_some() {
        removed.push("s");
    }
    let mut appended = vec![("r", cell.address.a1())];
    if let Some(cell_type) = cell_type {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let remains_empty = cell.empty && payload.is_none();
    write_tag(output, &cell.tag, remains_empty, &removed, &appended);
    if remains_empty {
        return Ok(());
    }
    if let Some(content) = content {
        write_content(output, &cell.tag.name, content)?;
    }
    if !cell.empty {
        if payload.is_some() {
            copy_without(
                output,
                source,
                cell.tag_end,
                cell.close_start,
                &cell.primary,
            );
        } else {
            output.extend_from_slice(&source[cell.tag_end..cell.close_start]);
        }
    }
    write_close(output, &cell.tag.name);
    Ok(())
}

fn write_new_action(
    output: &mut Vec<u8>,
    row_name: &str,
    address: Address,
    action: &Action,
) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Ok(());
    };
    if !action.creates_missing() {
        return Ok(());
    }
    let content = match payload.as_ref() {
        Some(Payload::Set(content)) => Some(content),
        Some(Payload::Clear | Payload::ClearIfPresent) | None => None,
    };
    let name = sibling_name(row_name, "c");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", address.a1())];
    if let Some(cell_type) = content.and_then(content_type) {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let empty = content.is_none();
    write_tag(output, &tag, empty, &[], &appended);
    if let Some(content) = content {
        write_content(output, &name, content)?;
        write_close(output, &name);
    }
    Ok(())
}

fn content_type(content: &Content) -> Option<&'static str> {
    match content {
        Content::Value(Value::Bool(_)) => Some("b"),
        Content::Value(Value::Text(_)) => Some("inlineStr"),
        Content::Value(Value::Date(_)) => Some("d"),
        Content::Value(Value::Error(_)) => Some("e"),
        Content::Value(Value::Number(_)) | Content::Formula(_) => None,
    }
}

fn write_content(output: &mut Vec<u8>, cell_name: &str, content: &Content) -> Result<()> {
    match content {
        Content::Value(Value::Bool(value)) => {
            write_text_element(output, cell_name, "v", if *value { "1" } else { "0" });
        },
        Content::Value(Value::Number(value)) => {
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Value(Value::Text(value)) => {
            let inline = sibling_name(cell_name, "is");
            let text = sibling_name(cell_name, "t");
            output.extend_from_slice(b"<");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b"><");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b" xml:space=\"preserve\">");
            output.extend_from_slice(escape_xml(&encode_spreadsheet_text(value)).as_bytes());
            output.extend_from_slice(b"</");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b"></");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b">");
        },
        Content::Value(Value::Date(value)) => {
            require_xml_text(value)?;
            write_text_element(output, cell_name, "v", &escape_xml(value));
        },
        Content::Value(Value::Error(value)) => {
            require_xml_text(value.as_str())?;
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Formula(formula) => {
            require_xml_text(formula.text())?;
            write_text_element(output, cell_name, "f", &escape_xml(formula.text()));
        },
    }
    Ok(())
}

fn write_text_element(output: &mut Vec<u8>, cell_name: &str, local: &str, value: &str) {
    let name = sibling_name(cell_name, local);
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
    output.extend_from_slice(value.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn require_xml_text(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
            || ('\u{20}'..='\u{D7FF}').contains(&character)
            || ('\u{E000}'..='\u{FFFD}').contains(&character)
            || ('\u{10000}'..='\u{10FFFF}').contains(&character)
    }) {
        Ok(())
    } else {
        Err(invalid(
            "cell content contains a character forbidden by XML 1.0",
        ))
    }
}

fn copy_without(output: &mut Vec<u8>, source: &[u8], start: usize, end: usize, removed: &[Span]) {
    let mut cursor = start;
    for span in removed {
        output.extend_from_slice(&source[cursor..span.start]);
        cursor = span.end;
    }
    output.extend_from_slice(&source[cursor..end]);
}

fn write_tag(
    output: &mut Vec<u8>,
    tag: &Tag,
    empty: bool,
    removed: &[&str],
    appended: &[(&str, String)],
) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if removed.iter().any(|name| *name == attribute.name.as_ref()) {
            continue;
        }
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(&attribute.value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    for (name, value) in appended {
        output.extend_from_slice(b" ");
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    if empty {
        output.extend_from_slice(b"/>");
    } else {
        output.extend_from_slice(b">");
    }
}

fn write_close(output: &mut Vec<u8>, name: &str) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| invalid(format!("worksheet element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("worksheet attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn column_range(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Column, Column)> {
    let min = required_u32(element, b"min", decoder, "worksheet column minimum")?;
    let max = required_u32(element, b"max", decoder, "worksheet column maximum")?;
    if min == 0 || min > max || max > COLUMNS {
        return Err(invalid(format!(
            "invalid worksheet column range '{min}:{max}' during edit"
        )));
    }
    Ok((Column::new(min - 1)?, Column::new(max - 1)?))
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn column(value: &str) -> Result<u32> {
    let value = value.trim_start_matches('$');
    let mut column = 0u32;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid(format!("invalid column reference '{value}'")));
        }
        column = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid(format!("column reference '{value}' overflows")))?;
    }
    Column::new(
        column
            .checked_sub(1)
            .ok_or_else(|| invalid(format!("invalid column reference '{value}'")))?,
    )
    .map(Column::get)
    .map_err(Into::into)
}

fn row(value: &str) -> Result<u32> {
    let value = value.trim_start_matches('$');
    let row = value
        .parse::<u32>()
        .ok()
        .and_then(|row| row.checked_sub(1))
        .ok_or_else(|| invalid(format!("invalid row reference '{value}'")))?;
    Row::new(row).map(Row::get).map_err(Into::into)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("worksheet XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Cell;
    use crate::raw::worksheet;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn minimally_rewrites_set_clear_remove_and_new_rows() {
        let xml = format!(
            r#"<?xml version="1.0"?><x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension ref="A1:C1" z:hint="kept"/><x:sheetData data="kept">
  <x:row r="1" spans="1:4" z:row="kept"><x:c r="A1" s="2" t="s" z:cell="kept"><x:v>0</x:v><x:extLst><z:data/></x:extLst></x:c><x:c r="C1"><x:v>3</x:v></x:c></x:row>
  <x:row r="5"><x:c r="D5" s="4"/></x:row>
</x:sheetData><x:extLst><z:untouched value="yes"/></x:extLst></x:worksheet>"#
        );
        let mut actions = BTreeMap::new();
        actions.insert(
            Address::from_a1("A1").unwrap(),
            Action::set("new & text".into()),
        );
        actions.insert(Address::from_a1("B1").unwrap(), Action::set(42_i32.into()));
        actions.insert(Address::from_a1("C1").unwrap(), Action::Remove);
        actions.insert(Address::from_a1("D5").unwrap(), Action::clear(true));
        actions.insert(Address::from_a1("A3").unwrap(), Action::set(true.into()));

        let edited = rewrite(xml.as_bytes(), "Data", actions).unwrap();
        let edited = std::str::from_utf8(&edited).unwrap();
        assert!(edited.contains(r#"z:cell="kept""#));
        assert!(edited.contains(r#"<x:dimension z:hint="kept" ref="A1:D5"/>"#));
        assert!(edited.contains("<x:extLst><z:data/></x:extLst>"));
        assert!(edited.contains(
            r#"<x:c s="2" z:cell="kept" r="A1" t="inlineStr"><x:is><x:t xml:space="preserve">new &amp; text</x:t></x:is>"#
        ));
        assert!(edited.contains(r#"<x:c r="B1"><x:v>42</x:v></x:c>"#));
        assert!(!edited.contains(r#"r="C1""#));
        assert!(edited.contains(r#"<x:row r="3"><x:c r="A3" t="b"><x:v>1</x:v></x:c></x:row>"#));
        assert!(edited.contains(r#"<x:c s="4" r="D5"></x:c>"#));
        assert!(edited.contains(r#"<x:extLst><z:untouched value="yes"/></x:extLst>"#));
        assert!(!edited.contains("spans="));

        let store = worksheet::parse(edited.as_bytes(), || Ok(None)).unwrap();
        assert!(matches!(
            store.get(Address::from_a1("A1").unwrap()),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "new & text"
        ));
        assert!(store.get(Address::from_a1("C1").unwrap()).is_none());
        assert!(matches!(
            store.get(Address::from_a1("D5").unwrap()),
            Some(Cell::Empty)
        ));
    }

    #[test]
    fn dimension_expansion_never_narrows_producer_bounds() {
        let empty =
            format!(r#"<worksheet xmlns="{S}"><dimension ref="A1"/><sheetData/></worksheet>"#);
        let created = rewrite(
            empty.as_bytes(),
            "Data",
            BTreeMap::from([(
                Address::from_a1("C3").expect("address"),
                Action::set(1_i32.into()),
            )]),
        )
        .expect("create C3");
        assert!(
            std::str::from_utf8(&created)
                .expect("UTF-8")
                .contains(r#"<dimension ref="A1:C3"/>"#)
        );

        let populated = format!(
            r#"<worksheet xmlns="{S}"><dimension ref="A1:C3"/><sheetData><row r="3"><c r="C3"><v>1</v></c></row></sheetData></worksheet>"#
        );
        let removed = rewrite(
            populated.as_bytes(),
            "Data",
            BTreeMap::from([(Address::from_a1("C3").expect("address"), Action::Remove)]),
        )
        .expect("remove C3");
        assert!(
            std::str::from_utf8(&removed)
                .expect("UTF-8")
                .contains(r#"<dimension ref="A1:C3"/>"#)
        );

        let absent = format!(r#"<worksheet xmlns="{S}"><sheetData/></worksheet>"#);
        let edited = rewrite(
            absent.as_bytes(),
            "Data",
            BTreeMap::from([(
                Address::from_a1("B2").expect("address"),
                Action::set(1_i32.into()),
            )]),
        )
        .expect("edit without producer dimension");
        assert!(
            !std::str::from_utf8(&edited)
                .expect("UTF-8")
                .contains("dimension")
        );
    }

    #[test]
    fn row_visibility_surgery_is_sparse_lossless_and_composes_with_cells() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension ref="A1:A4"/><x:sheetData><x:row r="1" hidden="1" z:keep="yes"><x:c r="A1"><x:v>1</x:v></x:c></x:row><x:row r="2" hidden="0" z:empty="keep"/><x:row r="4"><x:c r="A4"><x:v>4</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
        let plan = Plan {
            cells: BTreeMap::from([(
                Address::from_a1("A4").expect("A4"),
                Action::set(40_i32.into()),
            )]),
            rows: BTreeMap::from([
                (Row::new(0).expect("row 1"), RowAction::show()),
                (Row::new(1).expect("row 2"), RowAction::hide()),
                (Row::new(2).expect("row 3"), RowAction::hide()),
                (Row::new(3).expect("row 4"), RowAction::hide()),
            ]),
            columns: BTreeMap::new(),
        };

        let edited = rewrite(xml.as_bytes(), "Data", plan).expect("visibility edit");
        let edited = std::str::from_utf8(&edited).expect("UTF-8");
        assert!(edited.contains(r#"<x:row r="1" z:keep="yes">"#));
        assert!(edited.contains(r#"<x:row r="2" z:empty="keep" hidden="1"/>"#));
        assert!(edited.contains(r#"<x:row r="3" hidden="1"/>"#));
        assert!(edited.contains(r#"<x:row r="4" hidden="1"><x:c r="A4"><x:v>40</x:v>"#));
        assert!(edited.contains(r#"<x:dimension ref="A1:A4"/>"#));

        let store = worksheet::parse(edited.as_bytes(), || Ok(None)).expect("reparse rows");
        assert!(!store.row(Row::new(0).expect("row 1")).hidden());
        assert!(store.row(Row::new(1).expect("row 2")).hidden());
        assert!(store.row(Row::new(2).expect("row 3")).hidden());
        assert!(store.row(Row::new(3).expect("row 4")).hidden());
        assert!(matches!(
            store.get(Address::from_a1("A4").expect("A4")),
            Some(Cell::Value(Value::Number(value))) if value.as_str() == "40"
        ));
    }

    #[test]
    fn row_layout_facets_preserve_unedited_state_and_materialize_sparsely() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData><x:row r="2" s="1" customFormat="1" ht="20" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" thickTop="1" thickBot="1" ph="1" z:keep="yes"><x:c r="A2"><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
        let edited = rewrite(
            xml.as_bytes(),
            "Data",
            Plan {
                cells: BTreeMap::new(),
                rows: BTreeMap::from([
                    (
                        Row::new(1).expect("row 2"),
                        RowAction {
                            hidden: Some(false),
                            height: Some(HeightEffect::Reset),
                            outline: Some(Outline::new(3).expect("outline")),
                            collapsed: Some(false),
                            thick_top: Some(false),
                            phonetic: Some(false),
                            ..RowAction::default()
                        },
                    ),
                    (
                        Row::new(2).expect("row 3"),
                        RowAction {
                            height: Some(HeightEffect::Set(Height::new(25.0).expect("height"))),
                            outline: Some(Outline::new(1).expect("outline")),
                            collapsed: Some(true),
                            thick_bottom: Some(true),
                            phonetic: Some(true),
                            ..RowAction::default()
                        },
                    ),
                    (
                        Row::new(3).expect("row 4"),
                        RowAction {
                            hidden: Some(false),
                            height: Some(HeightEffect::Reset),
                            ..RowAction::default()
                        },
                    ),
                ]),
                columns: BTreeMap::new(),
            },
        )
        .expect("row layout rewrite");
        let text = std::str::from_utf8(&edited).expect("UTF-8");
        assert!(text.contains(concat!(
            r#"<x:row r="2" s="1" customFormat="1" thickBot="1" z:keep="yes" "#,
            r#"outlineLevel="3">"#
        )));
        assert!(text.contains(concat!(
            r#"<x:row r="3" ht="25" customHeight="1" outlineLevel="1" "#,
            r#"collapsed="1" thickBot="1" ph="1"/>"#
        )));
        assert!(!text.contains(r#"r="4""#));

        let store = worksheet::parse(&edited, || Ok(None)).expect("reparse row layout");
        let second = store.row(Row::new(1).expect("row 2"));
        assert_eq!(second.height(), None);
        assert!(!second.custom_height());
        assert!(!second.hidden());
        assert_eq!(second.outline().get(), 3);
        assert!(!second.collapsed());
        assert!(!second.thick_top());
        assert!(second.thick_bottom());
        assert!(!second.phonetic());
        assert!(second.custom_format());
        assert_eq!(
            store.row_entry(second.index()).unwrap().properties.style,
            Some(1)
        );
        let third = store.row(Row::new(2).expect("row 3"));
        assert_eq!(third.height().map(Height::get), Some(25.0));
        assert!(third.custom_height());
        assert_eq!(third.outline().get(), 1);
        assert!(third.collapsed());
        assert!(third.thick_bottom());
        assert!(third.phonetic());
        assert!(!store.row(Row::new(3).expect("row 4")).stored());
    }

    #[test]
    fn protected_sheet_blocks_row_visibility_before_rewrite() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#
        );
        let result = rewrite(
            xml.as_bytes(),
            "Data",
            Plan {
                cells: BTreeMap::new(),
                rows: BTreeMap::from([(Row::new(0).expect("row 1"), RowAction::hide())]),
                columns: BTreeMap::new(),
            },
        );
        assert!(matches!(
            result,
            Err(Error::RowEditBlocked {
                reason: RowEditBlock::ProtectedSheet,
                ..
            })
        ));
    }

    #[test]
    fn column_visibility_splits_effective_owners_and_preserves_other_properties() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:cols><x:col min="2" max="4" width="20" hidden="1" z:keep="yes"/><x:col min="3" max="3" width="10"/></x:cols><x:sheetData z:untouched="yes"/></x:worksheet>"#
        );
        let edited = rewrite(
            xml.as_bytes(),
            "Data",
            Plan {
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([
                    (Column::new(1).expect("B"), ColumnAction::show()),
                    (Column::new(2).expect("C"), ColumnAction::hide()),
                    (Column::new(4).expect("E"), ColumnAction::hide()),
                ]),
            },
        )
        .expect("column rewrite");
        let text = std::str::from_utf8(&edited).expect("UTF-8");
        assert!(text.contains(r#"<x:col width="20" z:keep="yes" min="2" max="2"/>"#));
        assert!(text.contains(r#"<x:col width="20" hidden="1" z:keep="yes" min="3" max="4"/>"#));
        assert!(text.contains(r#"<x:col width="10" min="3" max="3" hidden="1"/>"#));
        assert!(text.contains(r#"<x:col min="5" max="5" hidden="1"/>"#));
        assert!(text.contains(r#"<x:sheetData z:untouched="yes"/>"#));

        let store = worksheet::parse(&edited, || Ok(None)).expect("reparse columns");
        let b = store.column(Column::new(1).expect("B"));
        assert!(!b.hidden());
        assert_eq!(b.width().map(crate::column::Width::get), Some(20.0));
        let c = store.column(Column::new(2).expect("C"));
        assert!(c.hidden());
        assert_eq!(c.width().map(crate::column::Width::get), Some(10.0));
        assert!(store.column(Column::new(3).expect("D")).hidden());
        assert!(store.column(Column::new(4).expect("E")).hidden());
    }

    #[test]
    fn column_layout_facets_split_compactly_and_preserve_unedited_attributes() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:cols><x:col min="2" max="4" width="20" style="1" hidden="1" bestFit="1" customWidth="1" phonetic="1" outlineLevel="2" collapsed="1" z:keep="yes"/></x:cols><x:sheetData z:untouched="yes"/></x:worksheet>"#
        );
        let edited = rewrite(
            xml.as_bytes(),
            "Data",
            Plan {
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([
                    (
                        Column::new(1).expect("B"),
                        ColumnAction {
                            width: Some(WidthEffect::Reset),
                            best_fit: Some(false),
                            outline: Some(Outline::NONE),
                            collapsed: Some(false),
                            phonetic: Some(false),
                            ..ColumnAction::default()
                        },
                    ),
                    (
                        Column::new(2).expect("C"),
                        ColumnAction {
                            hidden: Some(false),
                            width: Some(WidthEffect::Set(Width::new(12.5).expect("width"))),
                            outline: Some(Outline::new(3).expect("outline")),
                            ..ColumnAction::default()
                        },
                    ),
                    (
                        Column::new(4).expect("E"),
                        ColumnAction {
                            width: Some(WidthEffect::Set(Width::new(15.0).expect("width"))),
                            best_fit: Some(true),
                            outline: Some(Outline::new(1).expect("outline")),
                            collapsed: Some(true),
                            phonetic: Some(true),
                            ..ColumnAction::default()
                        },
                    ),
                ]),
            },
        )
        .expect("column layout rewrite");
        let text = std::str::from_utf8(&edited).expect("UTF-8");
        assert!(text.contains(r#"style="1" hidden="1" z:keep="yes" min="2" max="2""#));
        assert!(text.contains(concat!(
            r#"style="1" bestFit="1" phonetic="1" collapsed="1" z:keep="yes" "#,
            r#"min="3" max="3" width="12.5" customWidth="1" outlineLevel="3""#
        )));
        assert!(text.contains(concat!(
            r#"<x:col min="5" max="5" width="15" customWidth="1" bestFit="1" "#,
            r#"outlineLevel="1" collapsed="1" phonetic="1"/>"#
        )));
        assert!(text.contains(r#"<x:sheetData z:untouched="yes"/>"#));

        let store = worksheet::parse(&edited, || Ok(None)).expect("reparse layout");
        let b = store.column(Column::new(1).expect("B"));
        assert_eq!(b.width(), None);
        assert!(b.hidden());
        assert!(!b.best_fit());
        assert_eq!(b.outline(), Outline::NONE);
        assert!(!b.collapsed());
        assert!(!b.phonetic());
        assert_eq!(
            store
                .column_entry(b.index())
                .map(|entry| entry.properties.style),
            Some(Some(1))
        );
        let c = store.column(Column::new(2).expect("C"));
        assert_eq!(c.width().map(Width::get), Some(12.5));
        assert!(!c.hidden());
        assert!(c.best_fit());
        assert_eq!(c.outline().get(), 3);
        assert!(c.collapsed());
        assert!(c.phonetic());
        let e = store.column(Column::new(4).expect("E"));
        assert_eq!(e.width().map(Width::get), Some(15.0));
        assert!(e.best_fit());
        assert_eq!(e.outline().get(), 1);
        assert!(e.collapsed());
        assert!(e.phonetic());
    }

    #[test]
    fn column_visibility_inserts_sparse_cols_and_blocks_unsafe_splits() {
        let plain = format!(r#"<x:worksheet xmlns:x="{S}"><x:sheetData/></x:worksheet>"#);
        let inserted = rewrite(
            plain.as_bytes(),
            "Data",
            Plan {
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([
                    (Column::new(1).expect("B"), ColumnAction::hide()),
                    (Column::new(2).expect("C"), ColumnAction::hide()),
                ]),
            },
        )
        .expect("insert cols");
        assert!(
            std::str::from_utf8(&inserted)
                .expect("UTF-8")
                .contains(r#"<x:cols><x:col min="2" max="3" hidden="1"/></x:cols><x:sheetData/>"#)
        );

        let extended = format!(
            r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><cols><col min="1" max="2"><z:future/></col></cols><sheetData/></worksheet>"#
        );
        assert!(matches!(
            rewrite(
                extended.as_bytes(),
                "Data",
                Plan {
                    cells: BTreeMap::new(),
                    rows: BTreeMap::new(),
                    columns: BTreeMap::from([(Column::new(0).expect("A"), ColumnAction::hide(),)]),
                },
            ),
            Err(Error::ColumnEditBlocked {
                reason: ColumnEditBlock::MarkupCompatibility,
                ..
            })
        ));

        let extended_columns = format!(
            r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><cols><col min="1" max="1"/><z:future/></cols><sheetData/></worksheet>"#
        );
        assert!(matches!(
            rewrite(
                extended_columns.as_bytes(),
                "Data",
                Plan {
                    cells: BTreeMap::new(),
                    rows: BTreeMap::new(),
                    columns: BTreeMap::from([(Column::new(1).expect("B"), ColumnAction::hide(),)]),
                },
            ),
            Err(Error::ColumnEditBlocked {
                reason: ColumnEditBlock::MarkupCompatibility,
                ..
            })
        ));

        let protected = format!(
            r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#
        );
        assert!(matches!(
            rewrite(
                protected.as_bytes(),
                "Data",
                Plan {
                    cells: BTreeMap::new(),
                    rows: BTreeMap::new(),
                    columns: BTreeMap::from([(Column::new(0).expect("A"), ColumnAction::hide(),)]),
                },
            ),
            Err(Error::ColumnEditBlocked {
                reason: ColumnEditBlock::ProtectedSheet,
                ..
            })
        ));
    }

    #[test]
    fn style_effects_preserve_payload_and_compose_with_value_effects() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData><x:row r="1"><x:c r="A1" s="1" z:keep="yes"><x:v>5</x:v></x:c><x:c r="B1" s="1"/></x:row></x:sheetData></x:worksheet>"#
        );
        let mut combined = Action::set(7_i32.into());
        combined.set_style(StyleEffect::Set(3));
        let actions = BTreeMap::from([
            (Address::from_a1("A1").unwrap(), Action::style(2)),
            (Address::from_a1("B1").unwrap(), Action::reset_style()),
            (Address::from_a1("C1").unwrap(), Action::style(3)),
            (Address::from_a1("D1").unwrap(), combined),
        ]);

        let edited = rewrite(xml.as_bytes(), "Data", actions).unwrap();
        let edited = std::str::from_utf8(&edited).unwrap();
        assert!(edited.contains(r#"z:keep="yes" r="A1" s="2"><x:v>5</x:v>"#));
        assert!(edited.contains(r#"<x:c r="B1"/>"#));
        assert!(edited.contains(r#"<x:c r="C1" s="3"/>"#));
        assert!(edited.contains(r#"<x:c r="D1" s="3"><x:v>7</x:v></x:c>"#));

        let store = worksheet::parse(edited.as_bytes(), || Ok(None)).unwrap();
        assert_eq!(
            store.entry(Address::from_a1("A1").unwrap()).unwrap().style,
            Some(2)
        );
        assert_eq!(
            store.entry(Address::from_a1("B1").unwrap()).unwrap().style,
            None
        );
        assert!(matches!(
            store.get(Address::from_a1("C1").unwrap()),
            Some(Cell::Empty)
        ));
    }

    #[test]
    fn blocks_dependencies_instead_of_guessing() {
        let cases = [
            (
                format!(
                    r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><sheetProtection sheet="1"/></worksheet>"#
                ),
                "A1",
                EditBlock::ProtectedSheet,
            ),
            (
                format!(
                    r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
                ),
                "B2",
                EditBlock::CoveredMerge,
            ),
            (
                format!(
                    r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B2">A1:B2*2</f></c></row></sheetData></worksheet>"#
                ),
                "B2",
                EditBlock::GroupFormula,
            ),
            (
                format!(
                    r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><dataValidations count="1"><dataValidation sqref="A1:B2"/></dataValidations></worksheet>"#
                ),
                "B2",
                EditBlock::DataValidation,
            ),
            (
                format!(
                    r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><sheetData><row r="1"><c r="A1"><z:value/></c></row></sheetData></worksheet>"#
                ),
                "A1",
                EditBlock::MarkupCompatibility,
            ),
        ];
        for (xml, address, expected) in cases {
            let address = Address::from_a1(address).unwrap();
            let actions = BTreeMap::from([(address, Action::set(1_i32.into()))]);
            assert!(matches!(
                rewrite(xml.as_bytes(), "Data", actions),
                Err(Error::EditBlocked { reason, .. }) if reason == expected
            ));
        }
    }
}
