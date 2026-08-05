//! Transactional rewrite of existing BIFF8 worksheet PivotTable views.

use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};

use super::Workbook;
use super::error::{Error, Result};
use super::pivot_table::*;
use super::writer::biff;

const BOUNDSHEET: u16 = 0x0085;
const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const MAX_RECORD: usize = 8_224;

/// An owned, transactional editor for PivotTable worksheet views in an XLS file.
pub struct PivotViewEditor {
    original: Vec<u8>,
    original_tables: Vec<Vec<PivotTable>>,
    tables: Vec<Vec<PivotTable>>,
    original_caches: Vec<PivotCache>,
    caches: Vec<PivotCache>,
    cache_stream_ids: Vec<u16>,
}

impl PivotViewEditor {
    /// Parses an owned XLS file. Calling [`finish`](Self::finish) without edits returns these bytes unchanged.
    pub fn new(bytes: Vec<u8>) -> Result<Self> {
        let workbook = Workbook::new(Cursor::new(bytes.clone()))?;
        let worksheet_count = workbook
            .sheets()
            .iter()
            .filter_map(|sheet| sheet.parsed_worksheet_index())
            .max()
            .map_or(0, |value| value + 1);
        let mut tables = vec![Vec::new(); worksheet_count];
        for sheet in workbook.sheets() {
            if let Some(index) = sheet.parsed_worksheet_index() {
                tables[index] = workbook.xls_worksheet(index)?.pivot_tables().to_vec();
            }
        }
        let caches = workbook.pivot_caches().to_vec();
        let cache_stream_ids = workbook.pivot_cache_stream_ids().to_vec();
        Ok(Self {
            original: bytes,
            original_tables: tables.clone(),
            tables,
            original_caches: caches.clone(),
            caches,
            cache_stream_ids,
        })
    }

    pub fn worksheet_count(&self) -> usize {
        self.tables.len()
    }
    pub fn pivot_tables(&self, worksheet: usize) -> Result<&[PivotTable]> {
        self.tables
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub fn add(&mut self, worksheet: usize, table: PivotTable) -> Result<()> {
        self.transaction(|tables, _| {
            let target = tables
                .get_mut(worksheet)
                .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
            if target
                .iter()
                .any(|value| value.view.name == table.view.name)
            {
                return Err(invalid(
                    SXVIEW_TYPE,
                    "duplicate PivotTable name on worksheet",
                ));
            }
            target.push(table);
            Ok(())
        })
    }

    pub fn replace_by_index(
        &mut self,
        worksheet: usize,
        index: usize,
        table: PivotTable,
    ) -> Result<PivotTable> {
        let mut removed = None;
        self.transaction(|tables, _| {
            let target = table_at_mut(tables, worksheet, index)?;
            removed = Some(std::mem::replace(target, table));
            Ok(())
        })?;
        Ok(removed.expect("transaction ran"))
    }

    pub fn replace_by_name(
        &mut self,
        worksheet: usize,
        name: &str,
        table: PivotTable,
    ) -> Result<PivotTable> {
        let index = self.index_by_name(worksheet, name)?;
        self.replace_by_index(worksheet, index, table)
    }

    pub fn update_by_index<F>(&mut self, worksheet: usize, index: usize, update: F) -> Result<()>
    where
        F: FnOnce(&mut PivotTable),
    {
        self.transaction(|tables, _| {
            update(table_at_mut(tables, worksheet, index)?);
            Ok(())
        })
    }

    pub fn update_by_name<F>(&mut self, worksheet: usize, name: &str, update: F) -> Result<()>
    where
        F: FnOnce(&mut PivotTable),
    {
        let index = self.index_by_name(worksheet, name)?;
        self.update_by_index(worksheet, index, update)
    }

    pub fn remove_by_index(&mut self, worksheet: usize, index: usize) -> Result<PivotTable> {
        let mut removed = None;
        self.transaction(|tables, _| {
            let target = tables
                .get_mut(worksheet)
                .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
            if index >= target.len() {
                return Err(invalid(SXVIEW_TYPE, "PivotTable index is out of range"));
            }
            removed = Some(target.remove(index));
            Ok(())
        })?;
        Ok(removed.expect("transaction ran"))
    }

    pub fn remove_by_name(&mut self, worksheet: usize, name: &str) -> Result<PivotTable> {
        let index = self.index_by_name(worksheet, name)?;
        self.remove_by_index(worksheet, index)
    }

    pub fn move_by_index(
        &mut self,
        from_sheet: usize,
        index: usize,
        to_sheet: usize,
        to_index: usize,
    ) -> Result<()> {
        self.transaction(|tables, _| {
            if from_sheet >= tables.len() || to_sheet >= tables.len() {
                return Err(Error::WorksheetNotFound(
                    "PivotTable move worksheet".to_string(),
                ));
            }
            if index >= tables[from_sheet].len() || to_index > tables[to_sheet].len() {
                return Err(invalid(
                    SXVIEW_TYPE,
                    "PivotTable move index is out of range",
                ));
            }
            let table = tables[from_sheet].remove(index);
            let adjusted = if from_sheet == to_sheet && to_index > index {
                to_index - 1
            } else {
                to_index
            };
            tables[to_sheet].insert(adjusted, table);
            Ok(())
        })
    }

    pub fn move_by_name(
        &mut self,
        from_sheet: usize,
        name: &str,
        to_sheet: usize,
        to_index: usize,
    ) -> Result<()> {
        let index = self.index_by_name(from_sheet, name)?;
        self.move_by_index(from_sheet, index, to_sheet, to_index)
    }

    pub fn reassign_cache_by_index(
        &mut self,
        worksheet: usize,
        index: usize,
        cache_index: u16,
    ) -> Result<()> {
        self.update_by_index(worksheet, index, |table| {
            table.view.cache_index = cache_index
        })
    }

    pub fn reassign_cache_by_name(
        &mut self,
        worksheet: usize,
        name: &str,
        cache_index: u16,
    ) -> Result<()> {
        let index = self.index_by_name(worksheet, name)?;
        self.reassign_cache_by_index(worksheet, index, cache_index)
    }

    /// Replaces grouping on a cache field and repairs the base-field child link atomically.
    pub fn update_cache_grouping(
        &mut self,
        cache_index: u16,
        field_index: u16,
        grouping: Option<PivotCacheGrouping>,
    ) -> Result<()> {
        let stream_id = *self
            .cache_stream_ids
            .get(usize::from(cache_index))
            .ok_or_else(|| invalid(SXVIEW_TYPE, "global cache index is out of range"))?;
        self.transaction(|_, caches| {
            let cache = caches
                .iter_mut()
                .find(|cache| cache.stream_id() == stream_id)
                .ok_or_else(|| invalid(SXVIEW_TYPE, "SXStreamID has no cache stream"))?;
            let fields = cache.fields_mut();
            let field_pos = usize::from(field_index);
            if field_pos >= fields.len() {
                return Err(invalid(0x00C7, "cache field index is out of range"));
            }
            if let Some(PivotCacheGrouping::Discrete(value)) = fields[field_pos].grouping() {
                let base = usize::from(value.base_field_index);
                if base < fields.len() {
                    fields[base].set_group_parent(None);
                }
            }
            if let Some(PivotCacheGrouping::Discrete(value)) = &grouping {
                let base = usize::from(value.base_field_index);
                if base >= fields.len() || base == field_pos {
                    return Err(invalid(0x00D9, "discrete grouping base field is invalid"));
                }
                fields[base].set_group_parent(Some(field_index));
            }
            fields[field_pos].replace_grouping(grouping);
            Ok(())
        })
    }

    /// Serializes all staged edits. Validation and serialization complete before any output is returned.
    pub fn finish(self) -> Result<Vec<u8>> {
        if self.tables == self.original_tables && self.caches == self.original_caches {
            return Ok(self.original);
        }
        let dirty_sheets = self
            .tables
            .iter()
            .zip(&self.original_tables)
            .enumerate()
            .filter_map(|(index, (left, right))| (left != right).then_some(index))
            .collect::<HashSet<_>>();
        let dirty_caches = self
            .caches
            .iter()
            .filter(|cache| {
                self.original_caches
                    .iter()
                    .find(|old| old.stream_id() == cache.stream_id())
                    != Some(*cache)
            })
            .map(PivotCache::stream_id)
            .collect::<HashSet<_>>();
        let mut input = OleFile::open(Cursor::new(self.original))?;
        let paths = input.list_streams();
        let workbook_path = paths
            .iter()
            .find(|path| {
                path.len() == 1
                    && (path[0].eq_ignore_ascii_case("Workbook")
                        || path[0].eq_ignore_ascii_case("Book"))
            })
            .cloned()
            .ok_or_else(|| Error::InvalidData("Workbook stream not found".to_string()))?;
        let refs = workbook_path.iter().map(String::as_str).collect::<Vec<_>>();
        let workbook_stream = input.open_stream(&refs)?;
        let rewritten_workbook =
            rewrite_workbook_stream(&workbook_stream, &self.tables, &dirty_sheets)?;
        let regenerated = regenerate_caches(&self.caches, &dirty_caches)?;
        let mut writer = OleWriter::new();
        for path in paths {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = if path == workbook_path {
                rewritten_workbook.clone()
            } else if path.len() == 2 && path[0].eq_ignore_ascii_case("_SX_DB_CUR") {
                u16::from_str_radix(&path[1], 16)
                    .ok()
                    .and_then(|id| regenerated.get(&id).cloned())
                    .unwrap_or(input.open_stream(&refs)?)
            } else {
                input.open_stream(&refs)?
            };
            writer.create_stream(&refs, &data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    fn index_by_name(&self, worksheet: usize, name: &str) -> Result<usize> {
        self.pivot_tables(worksheet)?
            .iter()
            .position(|table| table.view.name == name)
            .ok_or_else(|| invalid(SXVIEW_TYPE, format!("PivotTable {name:?} not found")))
    }

    fn transaction<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Vec<Vec<PivotTable>>, &mut Vec<PivotCache>) -> Result<()>,
    {
        let mut tables = self.tables.clone();
        let mut caches = self.caches.clone();
        edit(&mut tables, &mut caches)?;
        normalize_tables(&mut tables)?;
        validate_tables(&tables, &caches, &self.cache_stream_ids)?;
        self.tables = tables;
        self.caches = caches;
        Ok(())
    }
}

fn table_at_mut(
    tables: &mut [Vec<PivotTable>],
    worksheet: usize,
    index: usize,
) -> Result<&mut PivotTable> {
    tables
        .get_mut(worksheet)
        .and_then(|tables| tables.get_mut(index))
        .ok_or_else(|| invalid(SXVIEW_TYPE, "PivotTable index is out of range"))
}

fn normalize_tables(sheets: &mut [Vec<PivotTable>]) -> Result<()> {
    for tables in sheets {
        for table in tables {
            table.view.field_count = u16::try_from(table.fields.len())
                .map_err(|_| invalid(SXVIEW_TYPE, "field count overflow"))?;
            for field in &mut table.fields {
                field.item_count = u16::try_from(field.items.len())
                    .map_err(|_| invalid(SXVD_TYPE, "item count overflow"))?;
            }
            table.items = table
                .fields
                .iter()
                .flat_map(|field| field.items.iter().cloned())
                .collect();
            table.view.row_field_count = u16::try_from(table.row_fields.len())
                .map_err(|_| invalid(SXIVD_TYPE, "row axis count overflow"))?;
            table.view.col_field_count = u16::try_from(table.column_fields.len())
                .map_err(|_| invalid(SXIVD_TYPE, "column axis count overflow"))?;
            table.view.page_field_count = u16::try_from(table.page_entries.len())
                .map_err(|_| invalid(SXPI_TYPE, "page count overflow"))?;
            table.view.data_field_count = u16::try_from(table.data_items.len())
                .map_err(|_| invalid(SXDI_TYPE, "data count overflow"))?;
            table.view.data_row_count = u16::try_from(table.row_lines.len())
                .map_err(|_| invalid(SXLI_TYPE, "row line count overflow"))?;
            table.view.data_col_count = u16::try_from(table.column_lines.len())
                .map_err(|_| invalid(SXLI_TYPE, "column line count overflow"))?;
            if let Some(tag) = &mut table.query_tag {
                tag.table_name.clone_from(&table.view.name);
            }
        }
    }
    Ok(())
}

fn validate_tables(
    sheets: &[Vec<PivotTable>],
    caches: &[PivotCache],
    stream_ids: &[u16],
) -> Result<()> {
    let mut total = 0usize;
    for tables in sheets {
        total = total
            .checked_add(tables.len())
            .ok_or_else(|| invalid(SXVIEW_TYPE, "view count overflow"))?;
        if total > 1_024 {
            return Err(invalid(SXVIEW_TYPE, "view count exceeds resource bound"));
        }
        for (index, table) in tables.iter().enumerate() {
            if table.view.first_row > table.view.last_row
                || table.view.first_col > table.view.last_col
            {
                return Err(invalid(SXVIEW_TYPE, "reversed view range"));
            }
            if tables[..index].iter().any(|prior| {
                prior.view.first_row <= table.view.last_row
                    && table.view.first_row <= prior.view.last_row
                    && prior.view.first_col <= table.view.last_col
                    && table.view.first_col <= prior.view.last_col
            }) {
                return Err(invalid(SXVIEW_TYPE, "overlapping PivotTable views"));
            }
            let stream = *stream_ids
                .get(usize::from(table.view.cache_index))
                .ok_or_else(|| invalid(SXVIEW_TYPE, "cache ordinal is out of range"))?;
            let cache = caches
                .iter()
                .find(|cache| cache.stream_id() == stream)
                .ok_or_else(|| invalid(SXVIEW_TYPE, "cache stream is missing"))?;
            if cache.fields().len() != table.fields.len() {
                return Err(invalid(SXVIEW_TYPE, "view/cache field count mismatch"));
            }
            for axis in table.row_fields.iter().chain(&table.column_fields) {
                if let PivotAxisField::Field(value) = axis
                    && usize::from(*value) >= table.fields.len()
                {
                    return Err(invalid(SXIVD_TYPE, "axis field index is out of range"));
                }
            }
            for page in &table.page_entries {
                if usize::from(page.field_index) >= table.fields.len() {
                    return Err(invalid(SXPI_TYPE, "page field index is out of range"));
                }
            }
            for data in &table.data_items {
                if usize::from(data.source_field_index) >= cache.fields().len() {
                    return Err(invalid(
                        SXDI_TYPE,
                        "data source field index is out of range",
                    ));
                }
            }
        }
    }
    Ok(())
}

fn rewrite_workbook_stream(
    input: &[u8],
    tables: &[Vec<PivotTable>],
    dirty: &HashSet<usize>,
) -> Result<Vec<u8>> {
    let records = record_ranges(input)?;
    let mut bindings = Vec::new();
    let mut worksheet_index = 0usize;
    for (start, _, kind, body_start, body_end) in &records {
        if *kind == BOUNDSHEET {
            let body = &input[*body_start..*body_end];
            if body.len() < 6 {
                return Err(invalid(BOUNDSHEET, "short BoundSheet"));
            }
            let position = usize::try_from(u32::from_le_bytes(
                body[..4].try_into().expect("length checked"),
            ))
            .map_err(|_| invalid(BOUNDSHEET, "offset overflow"))?;
            let sheet_type = body[5];
            let parsed = (sheet_type == 0).then(|| {
                let current = worksheet_index;
                worksheet_index += 1;
                current
            });
            bindings.push((*start + 4, position, parsed));
        }
    }
    if worksheet_index != tables.len() {
        return Err(invalid(BOUNDSHEET, "worksheet directory count mismatch"));
    }
    let mut starts = bindings
        .iter()
        .map(|(_, position, parsed)| (*position, *parsed))
        .collect::<Vec<_>>();
    starts.sort_by_key(|value| value.0);
    let globals_end = records
        .iter()
        .find_map(|(_, end, kind, _, _)| (*kind == EOF).then_some(*end))
        .ok_or_else(|| invalid(EOF, "workbook globals have no EOF"))?;
    let mut previous = None;
    for (position, _) in &starts {
        if *position < globals_end || *position >= input.len() {
            return Err(invalid(BOUNDSHEET, "sheet offset is outside substreams"));
        }
        if previous == Some(*position) {
            return Err(invalid(BOUNDSHEET, "duplicate sheet offset"));
        }
        let target_kind = records
            .iter()
            .find_map(|(start, _, kind, _, _)| (*start == *position).then_some(*kind))
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet offset is not a record boundary"))?;
        if target_kind != BOF {
            return Err(invalid(BOUNDSHEET, "sheet offset does not target a BOF"));
        }
        previous = Some(*position);
    }
    let first = starts.first().map_or(input.len(), |value| value.0);
    let globals = input
        .get(..first)
        .ok_or_else(|| invalid(BOUNDSHEET, "workbook globals range is invalid"))?;
    let mut output = globals.to_vec();
    let mut new_positions = HashMap::new();
    for (ordinal, (start, parsed)) in starts.iter().enumerate() {
        let end = starts.get(ordinal + 1).map_or(input.len(), |value| value.0);
        let source = input
            .get(*start..end)
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet substream range is invalid"))?;
        let source_records = record_ranges(source)?;
        if source_records.first().map(|record| record.2) != Some(BOF)
            || source_records.last().map(|record| record.2) != Some(EOF)
        {
            return Err(invalid(
                BOUNDSHEET,
                "sheet substream must be bounded by BOF and EOF",
            ));
        }
        new_positions.insert(*start, output.len());
        if let Some(sheet) = parsed.filter(|sheet| dirty.contains(sheet)) {
            output.extend_from_slice(&rewrite_worksheet(source, &tables[sheet])?);
        } else {
            output.extend_from_slice(source);
        }
    }
    for (payload_offset, old_position, _) in bindings {
        let new_position = *new_positions
            .get(&old_position)
            .ok_or_else(|| invalid(BOUNDSHEET, "BoundSheet target missing"))?;
        let payload_end = payload_offset
            .checked_add(4)
            .ok_or_else(|| invalid(BOUNDSHEET, "BoundSheet payload offset overflow"))?;
        let payload = output
            .get_mut(payload_offset..payload_end)
            .ok_or_else(|| invalid(BOUNDSHEET, "BoundSheet payload is outside globals"))?;
        payload.copy_from_slice(
            &u32::try_from(new_position)
                .map_err(|_| invalid(BOUNDSHEET, "new sheet offset overflow"))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_worksheet(input: &[u8], tables: &[PivotTable]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(input.len());
    let mut inserted = false;
    for (start, end, kind, body_start, body_end) in record_ranges(input)? {
        let remove = is_worksheet_view_record(kind)
            && (kind != QSI_SX_TAG_TYPE
                || parse_qsi_sx_tag(&input[body_start..body_end])
                    .is_ok_and(|tag| tag.table_type == 1));
        if kind == EOF && !inserted {
            for table in tables {
                output.extend_from_slice(&serialize_table(table)?);
            }
            inserted = true;
        }
        if !remove {
            output.extend_from_slice(&input[start..end]);
        }
    }
    if !inserted {
        return Err(invalid(EOF, "worksheet has no EOF"));
    }
    Ok(output)
}

fn serialize_table(table: &PivotTable) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    biff::write_sxview(
        &mut out,
        &biff::SxViewConfig {
            first_row: table.view.first_row,
            last_row: table.view.last_row,
            first_col: table.view.first_col,
            last_col: table.view.last_col,
            first_header_row: table.view.first_header_row,
            first_data_row: table.view.first_data_row,
            first_data_col: table.view.first_data_col,
            cache_index: table.view.cache_index,
            data_axis: table.view.data_axis.code(),
            data_position: table.view.data_position,
            field_count: table.view.field_count,
            row_field_count: table.view.row_field_count,
            col_field_count: table.view.col_field_count,
            page_field_count: table.view.page_field_count,
            data_field_count: table.view.data_field_count,
            data_row_count: table.view.data_row_count,
            data_col_count: table.view.data_col_count,
            flags: table.view.flags,
            auto_format_index: table.view.auto_format_index,
            name: &table.view.name,
            data_field_name: &table.view.data_field_name,
        },
    )?;
    for field in &table.fields {
        biff::write_sxvd(
            &mut out,
            &biff::SxVdConfig {
                axis: field.axis.code(),
                subtotal_count: field.subtotal_count,
                subtotal_flags: field.subtotal_flags,
                item_count: field.item_count,
                name: field.name.as_deref(),
            },
        )?;
        for item in &field.items {
            biff::write_sxvi(
                &mut out,
                &biff::SxViConfig {
                    item_type: item.item_type.code(),
                    flags: item.flags,
                    cache_index: item.cache_index,
                    name: item.name.as_deref(),
                },
            )?;
        }
        write_field_extension(
            &mut out,
            field
                .extension
                .as_ref()
                .ok_or_else(|| invalid(SXVDEX_TYPE, "missing field extension"))?,
        )?;
    }
    write_axis(&mut out, &table.row_fields)?;
    write_axis(&mut out, &table.column_fields)?;
    if !table.page_entries.is_empty() {
        let entries = table
            .page_entries
            .iter()
            .map(|value| (value.item_index, value.field_index, value.object_id))
            .collect::<Vec<_>>();
        biff::write_sxpi(&mut out, &entries)?;
    }
    for data in &table.data_items {
        biff::write_sxdi(
            &mut out,
            &biff::SxDiConfig {
                source_field_index: data.source_field_index,
                function: data.function.code(),
                display_format: data.display_format,
                base_field_index: data.base_field_index,
                base_item_index: data.base_item_index,
                num_format_index: data.num_format_index,
                name: &data.name,
            },
        )?;
    }
    write_lines(&mut out, &table.row_lines)?;
    write_lines(&mut out, &table.column_lines)?;
    write_view_extension(
        &mut out,
        table
            .extension
            .as_ref()
            .ok_or_else(|| invalid(SXEX_TYPE, "missing view extension"))?,
    )?;
    if let Some(tag) = &table.query_tag {
        write_query_tag(&mut out, tag)?;
    }
    if let Some(value) = &table.view_ex9 {
        write_view_ex9(&mut out, value)?;
    }
    let split = table
        .additional_extensions
        .iter()
        .position(|value| value.kind == 0x1E)
        .unwrap_or(table.additional_extensions.len());
    for value in &table.additional_extensions[..split] {
        write_addl(&mut out, value)?;
    }
    for field in &table.fields {
        for value in &field.additional_extensions {
            write_addl(&mut out, value)?;
        }
    }
    for value in &table.additional_extensions[split..] {
        write_addl(&mut out, value)?;
    }
    Ok(out)
}

fn write_axis(out: &mut Vec<u8>, axis: &[PivotAxisField]) -> Result<()> {
    if axis.is_empty() {
        return Ok(());
    }
    let raw = axis
        .iter()
        .map(|value| match value {
            PivotAxisField::Field(index) => *index,
            PivotAxisField::DataLayout => 0xFFFE,
        })
        .collect::<Vec<_>>();
    biff::write_sxivd(out, &raw)
}
fn write_lines(out: &mut Vec<u8>, lines: &[PivotLayoutLine]) -> Result<()> {
    if lines.is_empty() {
        return Ok(());
    }
    let mut body = Vec::new();
    for line in lines {
        body.extend_from_slice(&line.repeated_item_count.to_le_bytes());
        body.extend_from_slice(&line.item_type.to_le_bytes());
        body.extend_from_slice(&(line.item_indices.len() as u16).to_le_bytes());
        body.extend_from_slice(&line.custom_name_flags.to_le_bytes());
        for value in &line.item_indices {
            body.extend_from_slice(&value.to_le_bytes());
        }
    }
    record(out, SXLI_TYPE, &body)
}
fn write_field_extension(out: &mut Vec<u8>, value: &PivotViewFieldExtension) -> Result<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&value.flags.to_le_bytes());
    body.extend_from_slice(&value.auto_sort_data_index.unwrap_or(u16::MAX).to_le_bytes());
    body.extend_from_slice(&value.auto_show_data_index.unwrap_or(u16::MAX).to_le_bytes());
    body.extend_from_slice(&value.number_format_index.to_le_bytes());
    body.extend_from_slice(
        &value
            .subtotal_name
            .as_ref()
            .map_or(u16::MAX, |s| s.chars().count() as u16)
            .to_le_bytes(),
    );
    body.extend_from_slice(&value.reserved);
    if let Some(name) = &value.subtotal_name {
        body.extend_from_slice(&no_cch(name));
    }
    record(out, SXVDEX_TYPE, &body)
}
fn write_view_extension(out: &mut Vec<u8>, v: &PivotViewExtension) -> Result<()> {
    let strings = [
        &v.error_string,
        &v.null_string,
        &v.tag,
        &v.page_field_style,
        &v.table_style,
        &v.vacate_style,
    ];
    let mut body = Vec::new();
    body.extend_from_slice(&v.format_count.to_le_bytes());
    for s in &strings[..3] {
        body.extend_from_slice(
            &s.as_ref()
                .map_or(u16::MAX, |x| x.chars().count() as u16)
                .to_le_bytes(),
        );
    }
    body.extend_from_slice(&v.select_count.to_le_bytes());
    body.extend_from_slice(&v.page_rows.to_le_bytes());
    body.extend_from_slice(&v.page_cols.to_le_bytes());
    body.extend_from_slice(&v.flags.to_le_bytes());
    for s in &strings[3..] {
        body.extend_from_slice(
            &s.as_ref()
                .map_or(u16::MAX, |x| x.chars().count() as u16)
                .to_le_bytes(),
        );
    }
    for s in strings.into_iter().flatten() {
        body.extend_from_slice(&no_cch(s));
    }
    record(out, SXEX_TYPE, &body)
}
fn write_query_tag(out: &mut Vec<u8>, v: &PivotQueryTag) -> Result<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&QSI_SX_TAG_TYPE.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&v.table_type.to_le_bytes());
    body.extend_from_slice(&v.flags.to_le_bytes());
    body.extend_from_slice(&v.options.to_le_bytes());
    body.push(v.last_refresh_version);
    body.push(v.minimum_refresh_version);
    body.push(16);
    body.push(v.first_created_version);
    body.extend_from_slice(&full_string(&v.table_name));
    body.extend_from_slice(&v.trailing_payload);
    record(out, QSI_SX_TAG_TYPE, &body)
}
fn write_view_ex9(out: &mut Vec<u8>, v: &PivotViewEx9) -> Result<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&SXVIEWEX9_TYPE.to_le_bytes());
    body.extend_from_slice(&v.frt_flags.to_le_bytes());
    body.extend_from_slice(&v.report_flags.to_le_bytes());
    body.extend_from_slice(&v.view_flags.to_le_bytes());
    body.extend_from_slice(&v.auto_format_index.to_le_bytes());
    body.extend_from_slice(&full_string(&v.grand_total_name));
    record(out, SXVIEWEX9_TYPE, &body)
}
fn write_addl(out: &mut Vec<u8>, v: &PivotAdditionalExtension) -> Result<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&SXADDL_TYPE.to_le_bytes());
    body.extend_from_slice(&v.reserved.to_le_bytes());
    body.push(v.class);
    body.push(v.kind);
    body.extend_from_slice(&v.payload);
    record(out, SXADDL_TYPE, &body)
}
fn no_cch(value: &str) -> Vec<u8> {
    if value.chars().any(|c| c as u32 > 0xff) {
        let mut out = vec![1];
        for unit in value.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes())
        }
        out
    } else {
        let mut out = vec![0];
        out.extend(value.chars().map(|c| c as u8));
        out
    }
}
fn full_string(value: &str) -> Vec<u8> {
    let mut out = (value.chars().count() as u16).to_le_bytes().to_vec();
    out.extend_from_slice(&no_cch(value));
    out
}
fn record(out: &mut Vec<u8>, kind: u16, body: &[u8]) -> Result<()> {
    if body.len() > MAX_RECORD {
        return Err(invalid(kind, "record exceeds BIFF8 size"));
    }
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(body.len() as u16).to_le_bytes());
    out.extend_from_slice(body);
    Ok(())
}

fn regenerate_caches(caches: &[PivotCache], dirty: &HashSet<u16>) -> Result<HashMap<u16, Vec<u8>>> {
    let mut result = HashMap::new();
    for cache in caches.iter().filter(|c| dirty.contains(&c.stream_id())) {
        let source = cache
            .fields()
            .iter()
            .map(|f| !matches!(f.grouping(), Some(PivotCacheGrouping::Discrete(_))))
            .collect::<Vec<_>>();
        let numeric = cache
            .fields()
            .iter()
            .enumerate()
            .map(|(i, f)| {
                source[i]
                    && (matches!(f.grouping(), Some(PivotCacheGrouping::Numeric(_)))
                        || f.items().is_empty()
                            && cache
                                .rows()
                                .iter()
                                .all(|r| matches!(r.get(i), Some(PivotCacheItem::Number(_)))))
            })
            .collect::<Vec<_>>();
        let mut indices = Vec::new();
        let mut numbers = Vec::new();
        for row in cache.rows() {
            let mut ix = Vec::new();
            let mut nums = Vec::new();
            for (i, field) in cache.fields().iter().enumerate() {
                if !source[i] {
                    continue;
                }
                let value = row
                    .get(i)
                    .ok_or_else(|| invalid(0x00C8, "cache row is short"))?;
                if numeric[i] && field.items().is_empty() {
                    if let PivotCacheItem::Number(n) = value {
                        nums.push(*n)
                    } else {
                        return Err(invalid(0x00C8, "unsupported inline typed cache rewrite"));
                    }
                } else if !field.items().is_empty() {
                    ix.push(
                        u16::try_from(
                            field
                                .items()
                                .iter()
                                .position(|item| item == value)
                                .ok_or_else(|| invalid(0x00C8, "cache row item is not shared"))?,
                        )
                        .map_err(|_| invalid(0x00C8, "cache item ordinal overflow"))?,
                    )
                } else {
                    return Err(invalid(0x00C8, "unsupported inline typed cache rewrite"));
                }
            }
            indices.push(ix);
            numbers.push(nums)
        }
        let infos = cache
            .fields()
            .iter()
            .enumerate()
            .map(|(i, f)| biff::PivotCacheFieldInfo {
                name: f.name(),
                items: f.items(),
                is_numeric: numeric[i],
                unique_numeric_count: f.items().len() as u16,
                grouping: f.grouping(),
                group_child: f.group_parent(),
                is_source_field: source[i],
            })
            .collect::<Vec<_>>();
        let rows = indices
            .iter()
            .zip(&numbers)
            .map(|(i, n)| biff::PivotCacheSourceRow {
                item_indices: i,
                numeric_values: n,
            })
            .collect::<Vec<_>>();
        let bytes = biff::generate_pivot_cache_stream(&biff::PivotCacheStreamInfo {
            stream_id: cache.stream_id(),
            record_count: cache.record_count(),
            fields: &infos,
            source_rows: &rows,
        })?;
        result.insert(cache.stream_id(), bytes);
    }
    Ok(result)
}

#[allow(clippy::type_complexity)]
fn record_ranges(data: &[u8]) -> Result<Vec<(usize, usize, u16, usize, usize)>> {
    let mut out = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let header = data.get(offset..offset + 4).ok_or(Error::InvalidLength {
            expected: offset + 4,
            found: data.len(),
        })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset
            .checked_add(4 + len)
            .ok_or_else(|| invalid(kind, "record size overflow"))?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        out.push((offset, end, kind, offset + 4, end));
        offset = end;
    }
    Ok(out)
}
fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn raw_record(kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(body.len() + 4);
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(body.len() as u16).to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    fn bound_sheet(position: u32) -> Vec<u8> {
        let mut body = position.to_le_bytes().to_vec();
        body.extend_from_slice(&[0, 2]);
        raw_record(BOUNDSHEET, &body)
    }

    #[test]
    fn malformed_bound_sheet_offsets_are_rejected_before_slicing() {
        let mut out_of_range = raw_record(BOF, &[]);
        out_of_range.extend(bound_sheet(u32::MAX));
        out_of_range.extend(raw_record(EOF, &[]));
        assert!(rewrite_workbook_stream(&out_of_range, &[], &HashSet::new()).is_err());

        let mut duplicate = raw_record(BOF, &[]);
        duplicate.extend(bound_sheet(28));
        duplicate.extend(bound_sheet(28));
        duplicate.extend(raw_record(EOF, &[]));
        duplicate.extend(raw_record(BOF, &[]));
        duplicate.extend(raw_record(EOF, &[]));
        assert!(rewrite_workbook_stream(&duplicate, &[], &HashSet::new()).is_err());

        let mut inside_record = raw_record(BOF, &[]);
        inside_record.extend(bound_sheet(19));
        inside_record.extend(raw_record(EOF, &[]));
        inside_record.extend(raw_record(BOF, &[]));
        inside_record.extend(raw_record(EOF, &[]));
        assert!(rewrite_workbook_stream(&inside_record, &[], &HashSet::new()).is_err());
    }
}
