#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! Strict BIFF12 codecs for slicer caches and worksheet slicer views.

use super::model::{
    Cache, CrossFilter, Item, Native, PivotTable, SortOrder, Source, Table, View, Views,
};
use super::validation::{cache as validate_cache, views as validate_views};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Record, Records, Writer, kind};

const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_RECORDS: usize = 1_000_000;

fn invalid(typ: &str, value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.to_string(),
        val: value.into(),
    }
}

fn empty(record: &Record<'_>, typ: &str) -> Result<()> {
    if !record.payload().is_empty() {
        return Err(Error::InvalidLength {
            expected: 0,
            found: record.payload().len(),
        });
    }
    let _ = typ;
    Ok(())
}

fn parse_all(data: &[u8]) -> Result<Vec<Record<'_>>> {
    if data.len() > MAX_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_PART_BYTES,
            found: data.len(),
        });
    }
    let mut output = Vec::new();
    for record in Records::new(data) {
        if output.len() >= MAX_RECORDS {
            return Err(invalid("slicer BIFF12 stream", "record limit exceeded"));
        }
        output.push(record?);
    }
    Ok(output)
}

fn strings(payload: &[u8], first: &'static str, second: &'static str) -> Result<(String, String)> {
    let mut cursor = Cursor::new(payload, "slicer cache definition");
    let first_value = cursor.read_wide_string()?;
    let second_value = cursor.read_wide_string()?;
    cursor.finish()?;
    let _ = (first, second);
    Ok((first_value, second_value))
}

fn pivot_tables(payload: &[u8]) -> Result<Vec<PivotTable>> {
    let mut cursor = Cursor::new(payload, "slicer cache PivotTables");
    let count = usize::try_from(cursor.read_u32()?).map_err(|_| Error::InvalidLength {
        expected: usize::MAX,
        found: usize::MAX,
    })?;
    if count > super::model::MAX_PIVOT_TABLES {
        return Err(Error::InvalidLength {
            expected: super::model::MAX_PIVOT_TABLES,
            found: count,
        });
    }
    let mut output = Vec::with_capacity(count);
    for _ in 0..count {
        output.push(PivotTable {
            sheet_id: cursor.read_u32()?,
            name: cursor.read_wide_string()?,
        });
    }
    cursor.finish()?;
    Ok(output)
}

fn native(payload: &[u8], items_payload: &[u8]) -> Result<Native> {
    if payload.len() != 9 {
        return Err(Error::InvalidLength {
            expected: 9,
            found: payload.len(),
        });
    }
    let mut header = Cursor::new(payload, "BrtBeginSlicerCacheNative");
    let reserved = header.read_u32()?;
    if reserved != 0 {
        return Err(invalid(
            "BrtBeginSlicerCacheNative reserved1",
            reserved.to_string(),
        ));
    }
    let cache_id = header.read_u32()?;
    let flags = header.read_u8()?;
    if flags & 0xc0 != 0 {
        return Err(invalid(
            "BrtBeginSlicerCacheNative reserved2",
            flags.to_string(),
        ));
    }
    let mut items = Cursor::new(items_payload, "BrtSlicerCacheNativeItem");
    let count = usize::try_from(items.read_u32()?).map_err(|_| Error::InvalidLength {
        expected: usize::MAX,
        found: usize::MAX,
    })?;
    if count == 0 || count > super::model::MAX_ITEMS {
        return Err(Error::InvalidLength {
            expected: super::model::MAX_ITEMS,
            found: count,
        });
    }
    let bytes = count.checked_mul(5).ok_or(Error::InvalidLength {
        expected: usize::MAX,
        found: count,
    })?;
    if items.remaining() != bytes {
        return Err(Error::InvalidLength {
            expected: bytes + 4,
            found: items_payload.len(),
        });
    }
    let mut output = Native {
        cache_id,
        sort_order: SortOrder::from_wire(flags & 0x03, "BrtBeginSlicerCacheNative fSortOrder")?,
        cross_filter: CrossFilter::from_wire(
            (flags >> 2) & 0x03,
            "BrtBeginSlicerCacheNative fCrossFilter",
        )?,
        sort_using_custom_lists: flags & 0x10 != 0,
        show_all_items: flags & 0x20 != 0,
        items: Vec::with_capacity(count),
    };
    for _ in 0..count {
        let cache_index = items.read_u32()?;
        let item_flags = items.read_u8()?;
        if item_flags & 0xfc != 0 {
            return Err(invalid(
                "BrtSlicerCacheNativeItem reserved flags",
                item_flags.to_string(),
            ));
        }
        output.items.push(Item {
            cache_index,
            selected: item_flags & 0x01 != 0,
            no_data: item_flags & 0x02 != 0,
        });
    }
    items.finish()?;
    Ok(output)
}

fn table(payload: &[u8]) -> Result<Table> {
    if payload.len() != 13 {
        return Err(Error::InvalidLength {
            expected: 13,
            found: payload.len(),
        });
    }
    let mut cursor = Cursor::new(payload, "BrtBeginTableSlicerCache");
    let reserved = cursor.read_u32()?;
    if reserved != 0 {
        return Err(invalid(
            "BrtBeginTableSlicerCache FRTBlank",
            reserved.to_string(),
        ));
    }
    let column = cursor.read_u32()?;
    let table_id = cursor.read_u32()?;
    let flags = cursor.read_u8()?;
    if flags & 0xe0 != 0 {
        return Err(invalid(
            "BrtBeginTableSlicerCache reserved flags",
            flags.to_string(),
        ));
    }
    cursor.finish()?;
    Ok(Table {
        column,
        table_id,
        sort_order: SortOrder::from_wire(flags & 0x03, "BrtBeginTableSlicerCache fSortOrder")?,
        cross_filter: CrossFilter::from_wire(
            (flags >> 2) & 0x03,
            "BrtBeginTableSlicerCache iCrossFilter",
        )?,
        sort_using_custom_lists: flags & 0x10 != 0,
    })
}

/// Parse one slicer cache part.
pub fn parse_cache(data: &[u8]) -> Result<Cache> {
    let records = parse_all(data)?;
    let mut index = 0;
    let begin = records
        .get(index)
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtBeginSlicerCache".to_string()))?;
    if begin.kind() != kind::BEGIN_SLICER_CACHE {
        return Err(invalid(
            "slicer cache",
            format!(
                "expected {}, found {}",
                kind::BEGIN_SLICER_CACHE,
                begin.kind()
            ),
        ));
    }
    empty(begin, "BrtBeginSlicerCache")?;
    index += 1;
    let definition = records
        .get(index)
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtBeginSlicerCacheDef".to_string()))?;
    if definition.kind() != kind::BEGIN_SLICER_CACHE_DEF {
        return Err(invalid("slicer cache", "missing BrtBeginSlicerCacheDef"));
    }
    let (name, hierarchy) = strings(definition.payload(), "stName", "stHierarchy")?;
    index += 1;
    let mut pivot = Vec::new();
    if records
        .get(index)
        .is_some_and(|record| record.kind() == kind::SLICER_CACHE_PIVOT_TABLES)
    {
        pivot = pivot_tables(records[index].payload())?;
        index += 1;
    }

    let source = match records.get(index).map(|record| record.kind()) {
        Some(kind) if kind == kind::BEGIN_SLICER_CACHE_NATIVE => {
            let header = records[index].payload();
            index += 1;
            let item = records.get(index).ok_or_else(|| {
                Error::UnexpectedEndOfStream("BrtSlicerCacheNativeItem".to_string())
            })?;
            if item.kind() != kind::SLICER_CACHE_NATIVE_ITEM {
                return Err(invalid("slicer cache", "native source has no item record"));
            }
            let native = native(header, item.payload())?;
            index += 1;
            let end = records.get(index).ok_or_else(|| {
                Error::UnexpectedEndOfStream("BrtEndSlicerCacheNative".to_string())
            })?;
            if end.kind() != kind::END_SLICER_CACHE_NATIVE {
                return Err(invalid("slicer cache", "native source has no end record"));
            }
            empty(end, "BrtEndSlicerCacheNative")?;
            index += 1;
            Source::Native(native)
        },
        Some(kind) if kind == kind::BEGIN_TABLE_SLICER_CACHE => {
            let table = table(records[index].payload())?;
            index += 1;
            let end = records.get(index).ok_or_else(|| {
                Error::UnexpectedEndOfStream("BrtEndTableSlicerCache".to_string())
            })?;
            if end.kind() != kind::END_TABLE_SLICER_CACHE {
                return Err(invalid("slicer cache", "table source has no end record"));
            }
            empty(end, "BrtEndTableSlicerCache")?;
            index += 1;
            Source::Table(table)
        },
        Some(kind) if kind == kind::BEGIN_SLICER_CACHE_OLAP_IMPL => {
            return Err(Error::UnsupportedFeature(
                "OLAP slicer cache nested item ranges are outside the bounded XLSB slice"
                    .to_string(),
            ));
        },
        _ => return Err(invalid("slicer cache", "missing supported source")),
    };

    let definition_end = records
        .get(index)
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtEndSlicerCacheDef".to_string()))?;
    if definition_end.kind() != kind::END_SLICER_CACHE_DEF {
        return Err(invalid("slicer cache", "missing BrtEndSlicerCacheDef"));
    }
    empty(definition_end, "BrtEndSlicerCacheDef")?;
    index += 1;
    let end = records
        .get(index)
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtEndSlicerCache".to_string()))?;
    if end.kind() != kind::END_SLICER_CACHE {
        return Err(invalid("slicer cache", "missing BrtEndSlicerCache"));
    }
    empty(end, "BrtEndSlicerCache")?;
    index += 1;
    if index != records.len() {
        return Err(invalid(
            "slicer cache",
            "unexpected records after BrtEndSlicerCache",
        ));
    }
    let cache = Cache {
        name,
        hierarchy,
        pivot_tables: pivot,
        source,
    };
    validate_cache(&cache)?;
    Ok(cache)
}

fn write_wide(output: &mut Vec<u8>, value: &str) -> Result<()> {
    Writer::new(output).write_wide_string(value)?;
    Ok(())
}

fn write_cache_inner(cache: &Cache, output: &mut Vec<u8>) -> Result<()> {
    validate_cache(cache)?;
    let mut writer = Writer::new(output);
    writer.write_record(kind::BEGIN_SLICER_CACHE, &[])?;
    let mut definition = Vec::new();
    write_wide(&mut definition, &cache.name)?;
    write_wide(&mut definition, &cache.hierarchy)?;
    writer.write_record(kind::BEGIN_SLICER_CACHE_DEF, &definition)?;
    if !cache.pivot_tables.is_empty() {
        let mut payload = Vec::new();
        Writer::new(&mut payload).write_u32(cache.pivot_tables.len() as u32)?;
        for pivot in &cache.pivot_tables {
            let mut item = Vec::new();
            let mut item_writer = Writer::new(&mut item);
            item_writer.write_u32(pivot.sheet_id)?;
            item_writer.write_wide_string(&pivot.name)?;
            payload.extend_from_slice(&item);
        }
        writer.write_record(kind::SLICER_CACHE_PIVOT_TABLES, &payload)?;
    }
    match &cache.source {
        Source::Native(native) => {
            let flags = native.sort_order.wire()
                | (native.cross_filter.wire() << 2)
                | (u8::from(native.sort_using_custom_lists) << 4)
                | (u8::from(native.show_all_items) << 5);
            let mut header = Vec::with_capacity(9);
            let mut header_writer = Writer::new(&mut header);
            header_writer.write_u32(0)?;
            header_writer.write_u32(native.cache_id)?;
            header_writer.write_u8(flags)?;
            writer.write_record(kind::BEGIN_SLICER_CACHE_NATIVE, &header)?;
            let mut items = Vec::with_capacity(4 + native.items.len() * 5);
            let mut item_writer = Writer::new(&mut items);
            item_writer.write_u32(native.items.len() as u32)?;
            for item in &native.items {
                item_writer.write_u32(item.cache_index)?;
                item_writer.write_u8(u8::from(item.selected) | (u8::from(item.no_data) << 1))?;
            }
            writer.write_record(kind::SLICER_CACHE_NATIVE_ITEM, &items)?;
            writer.write_record(kind::END_SLICER_CACHE_NATIVE, &[])?;
        },
        Source::Table(table) => {
            let flags = table.sort_order.wire()
                | (table.cross_filter.wire() << 2)
                | (u8::from(table.sort_using_custom_lists) << 4);
            let mut payload = Vec::with_capacity(13);
            let mut payload_writer = Writer::new(&mut payload);
            payload_writer.write_u32(0)?;
            payload_writer.write_u32(table.column)?;
            payload_writer.write_u32(table.table_id)?;
            payload_writer.write_u8(flags)?;
            writer.write_record(kind::BEGIN_TABLE_SLICER_CACHE, &payload)?;
            writer.write_record(kind::END_TABLE_SLICER_CACHE, &[])?;
        },
        Source::Olap(_) => {
            return Err(Error::UnsupportedFeature(
                "OLAP slicer cache authoring is outside the bounded XLSB slice".to_string(),
            ));
        },
    }
    writer.write_record(kind::END_SLICER_CACHE_DEF, &[])?;
    writer.write_record(kind::END_SLICER_CACHE, &[])?;
    Ok(())
}

/// Encode one slicer cache part.
pub fn write_cache(cache: &Cache) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    write_cache_inner(cache, &mut output)?;
    Ok(output)
}

fn parse_view(payload: &[u8]) -> Result<View> {
    let mut cursor = Cursor::new(payload, "BrtBeginSlicer");
    let flags = cursor.read_u8()?;
    if flags & 0xf0 != 0 {
        return Err(invalid("BrtBeginSlicer unused flags", flags.to_string()));
    }
    let view = View {
        start_item: cursor.read_u32()?,
        column_count: cursor.read_u32()?,
        level: cursor.read_u32()?,
        row_height: cursor.read_u32()?,
        name: cursor.read_wide_string()?,
        cache_name: cursor.read_wide_string()?,
        caption_visible: flags & 0x01 != 0,
        caption: None,
        style: None,
        locked_position: flags & 0x08 != 0,
    };
    let mut view = view;
    if flags & 0x02 != 0 {
        view.caption = Some(cursor.read_wide_string()?);
    }
    if flags & 0x04 != 0 {
        view.style = Some(cursor.read_wide_string()?);
    }
    cursor.finish()?;
    Ok(view)
}

/// Parse one worksheet slicer views part.
pub fn parse_views(data: &[u8]) -> Result<Views> {
    let records = parse_all(data)?;
    if records.len() < 2 || records[0].kind() != kind::BEGIN_SLICERS {
        return Err(invalid("slicer views", "missing BrtBeginSlicers"));
    }
    empty(&records[0], "BrtBeginSlicers")?;
    let mut views = Views::new();
    let mut index = 1;
    while index < records.len() {
        if records[index].kind() == kind::END_SLICERS {
            empty(&records[index], "BrtEndSlicers")?;
            index += 1;
            break;
        }
        if records[index].kind() != kind::BEGIN_SLICER {
            return Err(invalid(
                "slicer views",
                "unexpected record in view collection",
            ));
        }
        let view = parse_view(records[index].payload())?;
        index += 1;
        let end = records
            .get(index)
            .ok_or_else(|| Error::UnexpectedEndOfStream("BrtEndSlicer".to_string()))?;
        if end.kind() != kind::END_SLICER {
            return Err(invalid("slicer views", "missing BrtEndSlicer"));
        }
        empty(end, "BrtEndSlicer")?;
        views.items.push(view);
        index += 1;
    }
    if index != records.len() {
        return Err(invalid(
            "slicer views",
            "unexpected records after BrtEndSlicers",
        ));
    }
    validate_views(&views)?;
    Ok(views)
}

/// Encode one worksheet slicer views part.
pub fn write_views(views: &Views) -> Result<Vec<u8>> {
    validate_views(views)?;
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_SLICERS, &[])?;
    for view in &views.items {
        let mut payload = Vec::new();
        let mut flags = u8::from(view.caption_visible) | (u8::from(view.caption.is_some()) << 1);
        flags |= u8::from(view.style.is_some()) << 2;
        flags |= u8::from(view.locked_position) << 3;
        let mut payload_writer = Writer::new(&mut payload);
        payload_writer.write_u8(flags)?;
        payload_writer.write_u32(view.start_item)?;
        payload_writer.write_u32(view.column_count)?;
        payload_writer.write_u32(view.level)?;
        payload_writer.write_u32(view.row_height)?;
        payload_writer.write_wide_string(&view.name)?;
        payload_writer.write_wide_string(&view.cache_name)?;
        if let Some(caption) = &view.caption {
            payload_writer.write_wide_string(caption)?;
        }
        if let Some(style) = &view.style {
            payload_writer.write_wide_string(style)?;
        }
        writer.write_record(kind::BEGIN_SLICER, &payload)?;
        writer.write_record(kind::END_SLICER, &[])?;
    }
    writer.write_record(kind::END_SLICERS, &[])?;
    Ok(output)
}
