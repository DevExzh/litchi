//! Inert table filter/sort snapshots. Known fields are read semantically;
//! complete source collections retain ignored bits and unmodeled records.

use super::Range;
use crate::package::{
    error::{Error, Result},
    walker::{RecordWalker, malformed},
};
use crate::raw::{Cursor, Record, kind};
use std::{collections::BTreeSet, sync::Arc};

/// A table's saved `AutoFilter`. The immutable source collection is retained
/// for table publication; this model does not evaluate or apply criteria.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Filter {
    owner: Range,
    range: Range,
    columns: Vec<FilterColumn>,
    unmodeled: bool,
    source: Arc<Vec<u8>>,
}

impl Filter {
    /// The saved filter range, including its header row.
    #[must_use]
    pub const fn range(&self) -> Range {
        self.range
    }
    /// Column definitions in source order; indices are relative to the range.
    #[must_use]
    pub fn columns(&self) -> &[FilterColumn] {
        &self.columns
    }
    /// Whether any criteria or extension records lack a semantic projection.
    #[must_use]
    pub fn has_unmodeled_records(&self) -> bool {
        self.unmodeled || self.columns.iter().any(FilterColumn::has_unmodeled_records)
    }
    /// Exact encoded source records, including their begin/end records.
    #[must_use]
    pub fn source_records(&self) -> &[u8] {
        &self.source
    }
    pub(super) fn check_owner(&self, owner: Range) -> Result<()> {
        check_owner(self.owner, owner)
    }
}

/// Saved properties and known criteria of one relative filter column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FilterColumn {
    index: u32,
    hide_button: bool,
    button_in_next_column: bool,
    values: Option<ValueFilter>,
    unmodeled: bool,
}

impl FilterColumn {
    /// Zero-based column offset from the filter range's first column.
    #[must_use]
    pub const fn index(&self) -> u32 {
        self.index
    }
    /// Saved `fHideArrow` flag; no user interface is changed.
    #[must_use]
    pub const fn hide_button(&self) -> bool {
        self.hide_button
    }
    /// Saved `fNoBtn` flag: the button appears in the next column.
    #[must_use]
    pub const fn button_in_next_column(&self) -> bool {
        self.button_in_next_column
    }
    /// String criteria, when a `BrtBeginFilters` collection is present.
    #[must_use]
    pub const fn values(&self) -> Option<&ValueFilter> {
        self.values.as_ref()
    }
    /// Whether non-string criteria or extension records are retained opaquely.
    #[must_use]
    pub fn has_unmodeled_records(&self) -> bool {
        self.unmodeled
            || self
                .values
                .as_ref()
                .is_some_and(ValueFilter::has_unmodeled_records)
    }
}

/// Known string criteria of `BrtBeginFilters`. Date-group and future criteria
/// remain in `source_records` and set `has_unmodeled_records`; this is not an
/// executable predicate or a claim that every criterion is modeled.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueFilter {
    include_blank: bool,
    values: Vec<String>,
    unmodeled: bool,
}

impl ValueFilter {
    /// Whether blank values are included by the saved filter.
    #[must_use]
    pub const fn include_blank(&self) -> bool {
        self.include_blank
    }
    /// Exact saved strings in record order, without case folding or evaluation.
    #[must_use]
    pub fn values(&self) -> &[String] {
        &self.values
    }
    /// Whether additional criteria or extension records lack typed semantics.
    #[must_use]
    pub const fn has_unmodeled_records(&self) -> bool {
        self.unmodeled
    }
}

/// Saved classic sort properties. Ranges are retained literally: native Excel
/// can store a condition range wider than its sort-state range. No effective
/// range is synthesized and no rows or columns are reordered.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortState {
    owner: Range,
    range: Range,
    by_columns: bool,
    case_sensitive: bool,
    alternate_method: bool,
    conditions: Vec<SortCondition>,
    unmodeled: bool,
    source: Arc<Vec<u8>>,
}

impl SortState {
    /// The literal saved sort-state range.
    #[must_use]
    pub const fn range(&self) -> Range {
        self.range
    }
    /// The saved `fCol` flag; the specification ignores it under an `AutoFilter`.
    #[must_use]
    pub const fn by_columns(&self) -> bool {
        self.by_columns
    }
    /// Whether case-sensitive sorting was requested.
    #[must_use]
    pub const fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }
    /// Whether a method other than character order was requested.
    #[must_use]
    pub const fn alternate_method(&self) -> bool {
        self.alternate_method
    }
    /// Classic sort conditions in source order.
    #[must_use]
    pub fn conditions(&self) -> &[SortCondition] {
        &self.conditions
    }
    /// Whether future sort conditions or extension records remain opaque.
    #[must_use]
    pub const fn has_unmodeled_records(&self) -> bool {
        self.unmodeled
    }
    /// Exact encoded source collection, including begin/end records.
    #[must_use]
    pub fn source_records(&self) -> &[u8] {
        &self.source
    }
    pub(super) fn check_owner(&self, owner: Range) -> Result<()> {
        check_owner(self.owner, owner)
    }
}

/// What a classic sort condition compares. Formatting and icon identifiers
/// are inert saved references, not resolved display values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortOn {
    /// Cell values, optionally using a custom list.
    Value,
    /// Cell fill color identified by a differential-format index.
    CellColor(u32),
    /// Font color identified by a differential-format index.
    FontColor(u32),
    /// Icon-set and icon indices from `CFFlag`, retained without rendering.
    Icon { set: i32, icon: i32 },
}

/// One saved classic `BrtBeginSortCond` condition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortCondition {
    range: Range,
    descending: bool,
    sort_on: SortOn,
    custom_list: Option<String>,
}

impl SortCondition {
    /// Literal saved row/column range, without clipping it to the sort state.
    #[must_use]
    pub const fn range(&self) -> Range {
        self.range
    }
    /// Descending values, or matching colors/icons placed last.
    #[must_use]
    pub const fn descending(&self) -> bool {
        self.descending
    }
    /// Saved comparison kind and any inert formatting reference.
    #[must_use]
    pub const fn sort_on(&self) -> SortOn {
        self.sort_on
    }
    /// Comma-delimited custom value order; ignored for non-value sorts.
    #[must_use]
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
}

fn check_owner(before: Range, after: Range) -> Result<()> {
    if before != after {
        return Err(malformed(
            "table filter/sort metadata",
            "table range changed without rebinding retained metadata",
        ));
    }
    Ok(())
}

fn range(cursor: &mut Cursor<'_>) -> Result<Range> {
    Ok(Range {
        first_row: cursor.read_u32()?,
        last_row: cursor.read_u32()?,
        first_column: cursor.read_u32()?,
        last_column: cursor.read_u32()?,
    })
}

fn push<T>(values: &mut Vec<T>, value: T) -> Result<()> {
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "table filter/sort metadata",
        source,
    })?;
    values.push(value);
    Ok(())
}

fn retain(data: &[u8], start: usize, end: usize) -> Result<Arc<Vec<u8>>> {
    let bytes = data.get(start..end).ok_or_else(|| {
        malformed(
            "table filter/sort metadata",
            "source collection range is invalid",
        )
    })?;
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation {
            resource: "table filter/sort source",
            source,
        })?;
    copy.extend_from_slice(bytes);
    Ok(Arc::new(copy))
}

fn finish_end(walker: &mut RecordWalker<'_, '_>, record: Record<'_>) -> Result<()> {
    walker.finish(&Cursor::new(
        record.payload(),
        "table filter/sort end",
        walker.limits(),
    ))
}

fn skip(walker: &mut RecordWalker<'_, '_>, record: Record<'_>) -> Result<()> {
    if matches!(
        record.kind(),
        kind::END_LIST
            | kind::END_A_FILTER
            | kind::END_FILTER_COLUMN
            | kind::END_FILTERS
            | kind::END_SORT_STATE
            | kind::END_SORT_COND
    ) {
        return Err(malformed(
            "table filter/sort metadata",
            "unexpected collection end",
        ));
    }
    walker.skip_unhandled(record.kind(), "table filter/sort metadata")
}

pub(super) fn read_filter(
    data: &[u8],
    walker: &mut RecordWalker<'_, '_>,
    first: Record<'_>,
    owner: Range,
) -> Result<Filter> {
    let mut cursor = Cursor::new(first.payload(), "BrtBeginAFilter", walker.limits());
    let range = range(&mut cursor)?;
    walker.finish(&cursor)?;
    if range.first_row > range.last_row
        || range.last_row > 1_048_575
        || range.first_column > range.last_column
        || range.last_column > 16_383
    {
        return Err(malformed(
            "BrtBeginAFilter",
            "range is outside the worksheet grid",
        ));
    }
    let mut columns = Vec::new();
    let mut seen = BTreeSet::new();
    let mut unmodeled = false;
    loop {
        let record = walker.required("BrtEndAFilter")?;
        match record.kind() {
            kind::END_A_FILTER => {
                finish_end(walker, record)?;
                return Ok(Filter {
                    owner,
                    range,
                    columns,
                    unmodeled,
                    source: retain(data, first.offset(), walker.offset())?,
                });
            },
            kind::BEGIN_FILTER_COLUMN => {
                let column = read_column(walker, record)?;
                if column.index > range.last_column - range.first_column
                    || !seen.insert(column.index)
                {
                    return Err(malformed(
                        "BrtBeginFilterColumn",
                        "column is repeated or outside the filter range",
                    ));
                }
                push(&mut columns, column)?;
            },
            _ => {
                unmodeled = true;
                skip(walker, record)?;
            },
        }
    }
}

fn read_column(walker: &mut RecordWalker<'_, '_>, first: Record<'_>) -> Result<FilterColumn> {
    let mut cursor = Cursor::new(first.payload(), "BrtBeginFilterColumn", walker.limits());
    let index = cursor.read_u32()?;
    let flags = cursor.read_u16()?;
    walker.finish(&cursor)?;
    let mut column = FilterColumn {
        index,
        hide_button: flags & 1 != 0,
        button_in_next_column: flags & 2 != 0,
        values: None,
        unmodeled: false,
    };
    loop {
        let record = walker.required("BrtEndFilterColumn")?;
        match record.kind() {
            kind::END_FILTER_COLUMN => {
                finish_end(walker, record)?;
                return Ok(column);
            },
            kind::BEGIN_FILTERS => {
                if column.values.is_some() {
                    return Err(malformed("BrtBeginFilterColumn", "duplicate values filter"));
                }
                column.values = Some(read_values(walker, record)?);
            },
            _ => {
                column.unmodeled = true;
                skip(walker, record)?;
            },
        }
    }
}

fn read_values(walker: &mut RecordWalker<'_, '_>, first: Record<'_>) -> Result<ValueFilter> {
    let mut cursor = Cursor::new(first.payload(), "BrtBeginFilters", walker.limits());
    let include_blank = cursor.read_bool32()?;
    cursor.skip(4)?; // Undefined, ignored field; retained in the source collection.
    walker.finish(&cursor)?;
    let mut filter = ValueFilter {
        include_blank,
        values: Vec::new(),
        unmodeled: false,
    };
    loop {
        let record = walker.required("BrtEndFilters")?;
        match record.kind() {
            kind::END_FILTERS => {
                finish_end(walker, record)?;
                return Ok(filter);
            },
            kind::FILTER => {
                let mut cursor = Cursor::new(record.payload(), "BrtFilter", walker.limits());
                let value = cursor.read_wide_string()?;
                walker.finish(&cursor)?;
                if !(1..=255).contains(&value.encode_utf16().count()) {
                    return Err(malformed(
                        "BrtFilter",
                        "criteria string length is outside 1..=255",
                    ));
                }
                push(&mut filter.values, value)?;
            },
            _ => {
                filter.unmodeled = true;
                skip(walker, record)?;
            },
        }
    }
}

pub(super) fn read_sort(
    data: &[u8],
    walker: &mut RecordWalker<'_, '_>,
    first: Record<'_>,
    owner: Range,
) -> Result<SortState> {
    let mut cursor = Cursor::new(first.payload(), "BrtBeginSortState", walker.limits());
    let flags = cursor.read_u16()?;
    let range = range(&mut cursor)?;
    walker.finish(&cursor)?;
    let mut conditions = Vec::new();
    let mut unmodeled = false;
    loop {
        let record = walker.required("BrtEndSortState")?;
        match record.kind() {
            kind::END_SORT_STATE => {
                finish_end(walker, record)?;
                return Ok(SortState {
                    owner,
                    range,
                    by_columns: flags & 1 != 0,
                    case_sensitive: flags & 2 != 0,
                    alternate_method: flags & 4 != 0,
                    conditions,
                    unmodeled,
                    source: retain(data, first.offset(), walker.offset())?,
                });
            },
            kind::BEGIN_SORT_COND => {
                let mut cursor = Cursor::new(record.payload(), "BrtBeginSortCond", walker.limits());
                let flags = cursor.read_u16()?;
                let range = self::range(&mut cursor)?;
                let first = cursor.read_u32()?;
                let second = cursor.read_u32()?;
                let sort_on = match (flags >> 1) & 15 {
                    0 => SortOn::Value,
                    1 => SortOn::CellColor(first),
                    2 => SortOn::FontColor(first),
                    3 => SortOn::Icon {
                        set: i32::from_le_bytes(first.to_le_bytes()),
                        icon: i32::from_le_bytes(second.to_le_bytes()),
                    },
                    _ => return Err(malformed("BrtBeginSortCond", "unknown comparison kind")),
                };
                let custom = cursor.read_nullable_wide_string()?;
                walker.finish(&cursor)?;
                let custom_list = if sort_on == SortOn::Value {
                    custom
                } else {
                    None
                };
                push(
                    &mut conditions,
                    SortCondition {
                        range,
                        descending: flags & 1 != 0,
                        sort_on,
                        custom_list,
                    },
                )?;
                loop {
                    let child = walker.required("BrtEndSortCond")?;
                    if child.kind() == kind::END_SORT_COND {
                        finish_end(walker, child)?;
                        break;
                    }
                    unmodeled = true;
                    skip(walker, child)?;
                }
            },
            _ => {
                unmodeled = true;
                skip(walker, record)?;
            },
        }
    }
}
