//! PivotTable aggregate and workbook/package integration.

use super::codec::*;
use super::model::*;
use crate::error::Result;

/// Complete pivot table definition aggregated from multiple SX* records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTable {
    /// View definition (SXVIEW).
    pub view: PivotViewDef,
    /// Source type (SXVS).
    pub source_type: PivotSourceType,
    /// Field definitions (SXVD records, in order).
    pub fields: Vec<PivotViewField>,
    /// All items across all fields (SXVI records, in order).
    pub items: Vec<PivotViewItem>,
    /// Data field definitions (SXDI records).
    pub data_items: Vec<PivotDataItem>,
    /// Page field entries (SXPI records).
    pub page_entries: Vec<PageFieldEntry>,
    /// Explicit row-axis field ordering (first SXIVD).
    pub row_fields: Vec<PivotAxisField>,
    /// Explicit column-axis field ordering (second SXIVD).
    pub column_fields: Vec<PivotAxisField>,
    /// Visible row layout lines (first SXLI).
    pub row_lines: Vec<PivotLayoutLine>,
    /// Visible column layout lines (second SXLI).
    pub column_lines: Vec<PivotLayoutLine>,
    /// Legacy view extension (SXEX).
    pub extension: Option<PivotViewExtension>,
    /// Query/Pivot producer tag.
    pub query_tag: Option<PivotQueryTag>,
    /// Excel 9+ layout extension.
    pub view_ex9: Option<PivotViewEx9>,
    /// Losslessly preserved view-scoped SXADDL records.
    pub additional_extensions: Vec<PivotAdditionalExtension>,
}

impl PivotTable {
    /// Create a new pivot table from its view definition.
    pub fn new(view: PivotViewDef) -> Self {
        Self {
            source_type: PivotSourceType::Worksheet,
            fields: Vec::with_capacity(view.field_count as usize),
            items: Vec::new(),
            data_items: Vec::with_capacity(view.data_field_count as usize),
            page_entries: Vec::with_capacity(view.page_field_count as usize),
            row_fields: Vec::with_capacity(view.row_field_count as usize),
            column_fields: Vec::with_capacity(view.col_field_count as usize),
            row_lines: Vec::with_capacity(view.data_row_count as usize),
            column_lines: Vec::with_capacity(view.data_col_count as usize),
            extension: None,
            query_tag: None,
            view_ex9: None,
            additional_extensions: Vec::new(),
            view,
        }
    }

    pub const fn cache_index(&self) -> u16 {
        self.view.cache_index
    }

    /// Returns the cache field addressed by a view-field ordinal.
    pub fn cache_field<'a>(
        &self,
        cache: &'a PivotCache,
        field_index: u16,
    ) -> Option<&'a PivotCacheField> {
        cache.fields().get(usize::from(field_index))
    }
}

struct PivotTableBuild {
    table: PivotTable,
    row_axis_seen: bool,
    column_axis_seen: bool,
    page_seen: bool,
    row_lines_seen: bool,
    column_lines_seen: bool,
    sxaddl_field_cursor: usize,
    sxaddl_field_open: bool,
    extension_bytes: usize,
}

impl PivotTableBuild {
    fn new(view: PivotViewDef) -> Self {
        Self {
            table: PivotTable::new(view),
            row_axis_seen: false,
            column_axis_seen: false,
            page_seen: false,
            row_lines_seen: false,
            column_lines_seen: false,
            sxaddl_field_cursor: 0,
            sxaddl_field_open: false,
            extension_bytes: 0,
        }
    }

    fn fields_complete(&self) -> bool {
        self.table.fields.len() == usize::from(self.table.view.field_count)
            && self.table.fields.iter().all(|field| {
                field.items.len() == usize::from(field.item_count) && field.extension.is_some()
            })
    }

    fn axes_complete(&self) -> bool {
        (self.table.view.row_field_count == 0 || self.row_axis_seen)
            && (self.table.view.col_field_count == 0 || self.column_axis_seen)
    }

    fn page_complete(&self) -> bool {
        self.table.view.page_field_count == 0 || self.page_seen
    }
    fn data_complete(&self) -> bool {
        self.table.data_items.len() == usize::from(self.table.view.data_field_count)
    }
    fn lines_complete(&self) -> bool {
        (self.table.view.data_row_count == 0 || self.row_lines_seen)
            && (self.table.view.data_col_count == 0 || self.column_lines_seen)
    }

    fn require_fields(&self, record_type: u16) -> Result<()> {
        if self.fields_complete() {
            Ok(())
        } else {
            Err(cache_invalid(
                record_type,
                "record appears before all SXVD/SXVI/SXVDEx field groups",
            ))
        }
    }

    fn add_extension_bytes(&mut self, record_type: u16, count: usize) -> Result<()> {
        self.extension_bytes = self
            .extension_bytes
            .checked_add(count)
            .ok_or_else(|| cache_invalid(record_type, "PivotTable extension size overflow"))?;
        if self.extension_bytes > MAX_PIVOT_EXTENSION_BYTES {
            return Err(cache_invalid(
                record_type,
                "PivotTable extensions exceed resource bound",
            ));
        }
        Ok(())
    }

    fn feed(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        match record_type {
            SXVD_TYPE => {
                if let Some(previous) = self.table.fields.last()
                    && (previous.extension.is_none()
                        || previous.items.len() != usize::from(previous.item_count))
                {
                    return Err(cache_invalid(
                        record_type,
                        "SXVD starts before the previous field group is complete",
                    ));
                }
                if self.table.fields.len() == usize::from(self.table.view.field_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVD count exceeds SXVIEW field count",
                    ));
                }
                self.table.fields.push(parse_sxvd(data)?);
            },
            SXVI_TYPE => {
                let field = self
                    .table
                    .fields
                    .last_mut()
                    .ok_or_else(|| cache_invalid(record_type, "SXVI appears without SXVD"))?;
                if field.extension.is_some() {
                    return Err(cache_invalid(record_type, "SXVI appears after SXVDEx"));
                }
                if field.items.len() == usize::from(field.item_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVI count exceeds SXVD item count",
                    ));
                }
                let item = parse_sxvi(data)?;
                field.items.push(item.clone());
                self.table.items.push(item);
                if self.table.items.len() > MAX_PIVOT_ITEMS {
                    return Err(cache_invalid(
                        record_type,
                        "PivotTable items exceed resource bound",
                    ));
                }
            },
            SXVDEX_TYPE => {
                let extension = parse_sxvdex(data)?;
                let field = self
                    .table
                    .fields
                    .last_mut()
                    .ok_or_else(|| cache_invalid(record_type, "SXVDEx appears without SXVD"))?;
                if field.items.len() != usize::from(field.item_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXVDEx appears before all declared SXVI items",
                    ));
                }
                if field.extension.replace(extension).is_some() {
                    return Err(cache_invalid(record_type, "duplicate SXVDEx"));
                }
            },
            SXIVD_TYPE => {
                self.require_fields(record_type)?;
                let fields = parse_sxivd(data)?;
                if self.table.view.row_field_count != 0 && !self.row_axis_seen {
                    if fields.len() != usize::from(self.table.view.row_field_count) {
                        return Err(cache_invalid(
                            record_type,
                            "row SXIVD count does not match SXVIEW",
                        ));
                    }
                    self.table.row_fields = fields;
                    self.row_axis_seen = true;
                } else if self.table.view.col_field_count != 0 && !self.column_axis_seen {
                    if fields.len() != usize::from(self.table.view.col_field_count) {
                        return Err(cache_invalid(
                            record_type,
                            "column SXIVD count does not match SXVIEW",
                        ));
                    }
                    self.table.column_fields = fields;
                    self.column_axis_seen = true;
                } else {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order SXIVD",
                    ));
                }
            },
            SXPI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() {
                    return Err(cache_invalid(record_type, "SXPI appears before SXIVD axes"));
                }
                if self.page_seen || self.table.view.page_field_count == 0 {
                    return Err(cache_invalid(record_type, "duplicate or unexpected SXPI"));
                }
                let entries = parse_sxpi(data)?;
                if entries.len() != usize::from(self.table.view.page_field_count) {
                    return Err(cache_invalid(
                        record_type,
                        "SXPI count does not match SXVIEW",
                    ));
                }
                self.table.page_entries = entries;
                self.page_seen = true;
            },
            SXDI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() || !self.page_complete() {
                    return Err(cache_invalid(
                        record_type,
                        "SXDI appears before axes/page fields",
                    ));
                }
                if self.table.data_items.len() == usize::from(self.table.view.data_field_count) {
                    return Err(cache_invalid(record_type, "SXDI count exceeds SXVIEW"));
                }
                self.table.data_items.push(parse_sxdi(data)?);
            },
            SXLI_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete() || !self.page_complete() || !self.data_complete() {
                    return Err(cache_invalid(
                        record_type,
                        "SXLI appears before axis/page/data records",
                    ));
                }
                if self.table.view.data_row_count != 0 && !self.row_lines_seen {
                    self.table.row_lines = parse_sxli(
                        data,
                        usize::from(self.table.view.data_row_count),
                        usize::from(self.table.view.field_count).saturating_add(1),
                    )?;
                    self.row_lines_seen = true;
                } else if self.table.view.data_col_count != 0 && !self.column_lines_seen {
                    self.table.column_lines = parse_sxli(
                        data,
                        usize::from(self.table.view.data_col_count),
                        usize::from(self.table.view.field_count).saturating_add(1),
                    )?;
                    self.column_lines_seen = true;
                } else {
                    return Err(cache_invalid(record_type, "duplicate or out-of-order SXLI"));
                }
            },
            SXEX_TYPE => {
                self.require_fields(record_type)?;
                if !self.axes_complete()
                    || !self.page_complete()
                    || !self.data_complete()
                    || !self.lines_complete()
                {
                    return Err(cache_invalid(
                        record_type,
                        "SXEX appears before the core PivotTable view is complete",
                    ));
                }
                if self.table.extension.replace(parse_sxex(data)?).is_some() {
                    return Err(cache_invalid(record_type, "duplicate SXEX"));
                }
            },
            QSI_SX_TAG_TYPE => {
                if self.table.extension.is_none() {
                    return Err(cache_invalid(record_type, "QsiSxTag appears before SXEX"));
                }
                if self.table.query_tag.is_some()
                    || self.table.view_ex9.is_some()
                    || !self.table.additional_extensions.is_empty()
                {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order QsiSxTag",
                    ));
                }
                self.add_extension_bytes(record_type, data.len())?;
                let tag = parse_qsi_sx_tag(data)?;
                if tag.table_name != self.table.view.name {
                    return Err(cache_invalid(
                        record_type,
                        "QsiSxTag table name does not match SXVIEW",
                    ));
                }
                self.table.query_tag = Some(tag);
            },
            SXVIEWEX9_TYPE => {
                if self.table.query_tag.is_none()
                    || self.table.view_ex9.is_some()
                    || !self.table.additional_extensions.is_empty()
                {
                    return Err(cache_invalid(
                        record_type,
                        "duplicate or out-of-order SXVIEWEX9",
                    ));
                }
                self.add_extension_bytes(record_type, data.len())?;
                self.table.view_ex9 = Some(parse_sxviewex9(data)?);
            },
            SXADDL_TYPE => {
                if self.table.extension.is_none() {
                    return Err(cache_invalid(record_type, "SXADDL appears before SXEX"));
                }
                self.add_extension_bytes(record_type, data.len())?;
                let extension = parse_sxaddl(data)?;
                if extension.class == 0x17 {
                    if self.sxaddl_field_cursor >= self.table.fields.len() {
                        return Err(cache_invalid(
                            record_type,
                            "field SXADDL ordinal exceeds SXVD count",
                        ));
                    }
                    if !self.sxaddl_field_open && extension.kind != 0x00 {
                        return Err(cache_invalid(
                            record_type,
                            "field SXADDL group does not start with a name record",
                        ));
                    }
                    self.sxaddl_field_open = extension.kind != 0xFF;
                    self.table.fields[self.sxaddl_field_cursor]
                        .additional_extensions
                        .push(extension);
                    if !self.sxaddl_field_open {
                        self.sxaddl_field_cursor += 1;
                    }
                } else {
                    if self.sxaddl_field_open {
                        return Err(cache_invalid(
                            record_type,
                            "view SXADDL interrupts a field extension group",
                        ));
                    }
                    self.table.additional_extensions.push(extension);
                }
            },
            _ => return Err(cache_invalid(record_type, "unexpected PivotTable record")),
        }
        Ok(())
    }

    fn finish(self) -> Result<PivotTable> {
        if !self.fields_complete()
            || !self.axes_complete()
            || !self.page_complete()
            || !self.data_complete()
            || !self.lines_complete()
            || self.table.extension.is_none()
            || self.sxaddl_field_open
        {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "incomplete or unterminated PivotTable worksheet-view record set",
            ));
        }
        Ok(self.table)
    }
}

/// Ordered worksheet PivotTable record collector.
pub(crate) struct PivotTableCollector {
    current: Option<PivotTableBuild>,
    completed: Vec<PivotTable>,
}

impl PivotTableCollector {
    pub(crate) fn new() -> Self {
        Self {
            current: None,
            completed: Vec::new(),
        }
    }

    fn push_current(&mut self) -> Result<()> {
        let Some(build) = self.current.take() else {
            return Ok(());
        };
        let table = build.finish()?;
        if self.completed.len() == MAX_PIVOT_VIEWS_PER_SHEET {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "PivotTable view count exceeds resource bound",
            ));
        }
        if self
            .completed
            .iter()
            .any(|prior| ranges_overlap(&prior.view, &table.view))
        {
            return Err(cache_invalid(
                SXVIEW_TYPE,
                "overlapping PivotTable output ranges",
            ));
        }
        self.completed.push(table);
        Ok(())
    }

    /// Returns true when the record belongs to the PivotTable aggregate.
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<bool> {
        let pivot_record = is_worksheet_view_record(record_type);
        if record_type == SXVIEW_TYPE {
            self.push_current()?;
            self.current = Some(PivotTableBuild::new(parse_sxview(data)?));
            return Ok(true);
        }
        if pivot_record {
            if self.current.is_none()
                && record_type == QSI_SX_TAG_TYPE
                && parse_qsi_sx_tag(data)?.table_type != 1
            {
                return Ok(false);
            }
            let build = self.current.as_mut().ok_or_else(|| {
                cache_invalid(record_type, "orphan PivotTable worksheet-view record")
            })?;
            build.feed(record_type, data)?;
            return Ok(true);
        }
        if self
            .current
            .as_ref()
            .is_some_and(|build| build.table.extension.is_none())
        {
            self.push_current()?;
        }
        Ok(false)
    }

    pub(crate) fn finish(mut self) -> Result<Vec<PivotTable>> {
        self.push_current()?;
        Ok(self.completed)
    }
}

fn ranges_overlap(left: &PivotViewDef, right: &PivotViewDef) -> bool {
    left.first_row <= right.last_row
        && right.first_row <= left.last_row
        && left.first_col <= right.last_col
        && right.first_col <= left.last_col
}

pub(crate) fn validate_pivot_cache_links(
    worksheets: &[crate::worksheet::Worksheet],
    caches: &[PivotCache],
    cache_stream_ids: &[u16],
) -> Result<()> {
    for worksheet in worksheets {
        for table in worksheet.pivot_tables() {
            let stream_id = *cache_stream_ids
                .get(usize::from(table.view.cache_index))
                .ok_or_else(|| {
                    cache_invalid(SXVIEW_TYPE, "SXVIEW global cache index is out of range")
                })?;
            let cache = caches
                .iter()
                .find(|cache| cache.stream_id() == stream_id)
                .ok_or_else(|| {
                    cache_invalid(
                        SXVIEW_TYPE,
                        "SXStreamID has no matching PivotCache storage stream",
                    )
                })?;
            if table.fields.len() != cache.fields().len() {
                return Err(cache_invalid(
                    SXVIEW_TYPE,
                    "SXVIEW field count does not match linked PivotCache",
                ));
            }
            validate_axis_fields(table, &table.row_fields, PivotAxis::Row)?;
            validate_axis_fields(table, &table.column_fields, PivotAxis::Column)?;
            for (index, field) in table.fields.iter().enumerate() {
                let cache_field = &cache.fields()[index];
                let visible_items = cache_field
                    .grouping()
                    .map(PivotCacheGrouping::group_items)
                    .unwrap_or(cache_field.items());
                for item in &field.items {
                    if item.item_type.code() == 0
                        && usize::from(item.cache_index) >= visible_items.len()
                    {
                        return Err(cache_invalid(
                            SXVI_TYPE,
                            "SXVI cache item ordinal is out of range",
                        ));
                    }
                }
                if let Some(extension) = &field.extension {
                    for ordinal in [
                        extension.auto_sort_data_index,
                        extension.auto_show_data_index,
                    ]
                    .into_iter()
                    .flatten()
                    {
                        if usize::from(ordinal) >= table.data_items.len() {
                            return Err(cache_invalid(
                                SXVDEX_TYPE,
                                "SXVDEx data-item ordinal is out of range",
                            ));
                        }
                    }
                }
            }
            for page in &table.page_entries {
                let field = table
                    .fields
                    .get(usize::from(page.field_index))
                    .ok_or_else(|| {
                        cache_invalid(SXPI_TYPE, "SXPI field ordinal is out of range")
                    })?;
                if field.axis != PivotAxis::Page
                    || !matches!(page.selection(), PivotPageSelection::All)
                        && usize::from(page.item_index) >= field.items.len()
                {
                    return Err(cache_invalid(
                        SXPI_TYPE,
                        "SXPI item ordinal or axis is invalid",
                    ));
                }
            }
            for item in &table.data_items {
                if usize::from(item.source_field_index) >= cache.fields().len() {
                    return Err(cache_invalid(
                        SXDI_TYPE,
                        "SXDI source field ordinal is out of range",
                    ));
                }
                if item.display_format != 0 {
                    let base = table
                        .fields
                        .get(usize::from(item.base_field_index))
                        .ok_or_else(|| {
                            cache_invalid(SXDI_TYPE, "SXDI base field ordinal is out of range")
                        })?;
                    if usize::from(item.base_item_index) >= base.items.len() {
                        return Err(cache_invalid(
                            SXDI_TYPE,
                            "SXDI base item ordinal is out of range",
                        ));
                    }
                }
            }
        }
    }
    Ok(())
}

fn validate_axis_fields(
    table: &PivotTable,
    fields: &[PivotAxisField],
    axis: PivotAxis,
) -> Result<()> {
    let mut data_layout_seen = false;
    for entry in fields {
        match *entry {
            PivotAxisField::Field(index) => {
                let field = table.fields.get(usize::from(index)).ok_or_else(|| {
                    cache_invalid(SXIVD_TYPE, "SXIVD field ordinal is out of range")
                })?;
                if field.axis != axis {
                    return Err(cache_invalid(
                        SXIVD_TYPE,
                        "SXIVD field ordinal references the wrong axis",
                    ));
                }
            },
            PivotAxisField::DataLayout => {
                if data_layout_seen
                    || table.view.data_field_count <= 1
                    || table.view.data_axis != axis
                {
                    return Err(cache_invalid(
                        SXIVD_TYPE,
                        "invalid or duplicate SXIVD data-layout field",
                    ));
                }
                data_layout_seen = true;
            },
        }
    }
    Ok(())
}
