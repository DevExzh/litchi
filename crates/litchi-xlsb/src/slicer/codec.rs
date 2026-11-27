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
use crate::raw::{Cursor, Encoder, Record, Records, kind};

const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_RECORDS: usize = 1_000_000;

fn invalid(typ: &'static str, value: &'static str) -> Error {
    Error::Unrecognized { typ, val: value }
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

fn parse_all(data: &[u8], limits: crate::raw::Limits) -> Result<Vec<Record<'_>>> {
    if data.len() > MAX_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_PART_BYTES,
            found: data.len(),
        });
    }
    let mut output = Vec::new();
    for record in Records::new(data, limits) {
        if output.len() >= MAX_RECORDS {
            return Err(invalid("slicer BIFF12 stream", "record limit exceeded"));
        }
        output.push(record?);
    }
    Ok(output)
}

fn strings(
    payload: &[u8],
    first: &'static str,
    second: &'static str,
    limits: crate::raw::Limits,
) -> Result<(String, String)> {
    let mut cursor = Cursor::new(payload, "slicer cache definition", limits);
    let first_value = cursor.read_wide_string()?;
    let second_value = cursor.read_wide_string()?;
    cursor.finish()?;
    let _ = (first, second);
    Ok((first_value, second_value))
}

fn pivot_tables(payload: &[u8], limits: crate::raw::Limits) -> Result<Vec<PivotTable>> {
    let mut cursor = Cursor::new(payload, "slicer cache PivotTables", limits);
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

fn native(payload: &[u8], items_payload: &[u8], limits: crate::raw::Limits) -> Result<Native> {
    if payload.len() != 9 {
        return Err(Error::InvalidLength {
            expected: 9,
            found: payload.len(),
        });
    }
    let mut header = Cursor::new(payload, "BrtBeginSlicerCacheNative", limits);
    let reserved = header.read_u32()?;
    if reserved != 0 {
        return Err(invalid(
            "BrtBeginSlicerCacheNative reserved1",
            "unexpected value",
        ));
    }
    let cache_id = header.read_u32()?;
    let flags = header.read_u8()?;
    if flags & 0xc0 != 0 {
        return Err(invalid(
            "BrtBeginSlicerCacheNative reserved2",
            "unexpected value",
        ));
    }
    let mut items = Cursor::new(items_payload, "BrtSlicerCacheNativeItem", limits);
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
                "unexpected value",
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

fn table(payload: &[u8], limits: crate::raw::Limits) -> Result<Table> {
    if payload.len() != 13 {
        return Err(Error::InvalidLength {
            expected: 13,
            found: payload.len(),
        });
    }
    let mut cursor = Cursor::new(payload, "BrtBeginTableSlicerCache", limits);
    let reserved = cursor.read_u32()?;
    if reserved != 0 {
        return Err(invalid(
            "BrtBeginTableSlicerCache FRTBlank",
            "unexpected value",
        ));
    }
    let column = cursor.read_u32()?;
    let table_id = cursor.read_u32()?;
    let flags = cursor.read_u8()?;
    if flags & 0xe0 != 0 {
        return Err(invalid(
            "BrtBeginTableSlicerCache reserved flags",
            "unexpected value",
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
pub fn parse_cache(data: &[u8], limits: crate::raw::Limits) -> Result<Cache> {
    let records = parse_all(data, limits)?;
    let mut index = 0;
    let begin = records
        .get(index)
        .ok_or(Error::UnexpectedEndOfStream("BrtBeginSlicerCache"))?;
    if begin.kind() != kind::BEGIN_SLICER_CACHE {
        return Err(invalid("slicer cache", "expected, found"));
    }
    empty(begin, "BrtBeginSlicerCache")?;
    index += 1;
    let mut alternate_content = Vec::new();
    if records
        .get(index)
        .is_some_and(|record| record.kind() == kind::AC_BEGIN)
    {
        let start = records[index].offset();
        index = cache_identity_end(&records, index)?;
        let end = records.get(index).map_or(data.len(), Record::offset);
        alternate_content.extend_from_slice(&data[start..end]);
    }
    let definition = records
        .get(index)
        .ok_or(Error::UnexpectedEndOfStream("BrtBeginSlicerCacheDef"))?;
    if definition.kind() != kind::BEGIN_SLICER_CACHE_DEF {
        return Err(invalid("slicer cache", "missing BrtBeginSlicerCacheDef"));
    }
    let (name, hierarchy) = strings(definition.payload(), "stName", "stHierarchy", limits)?;
    index += 1;
    let mut pivot = Vec::new();
    if records
        .get(index)
        .is_some_and(|record| record.kind() == kind::SLICER_CACHE_PIVOT_TABLES)
    {
        pivot = pivot_tables(records[index].payload(), limits)?;
        index += 1;
    }

    let source = match records.get(index).map(|record| record.kind()) {
        Some(kind) if kind == kind::BEGIN_SLICER_CACHE_NATIVE => {
            let header = records[index].payload();
            index += 1;
            let item = records
                .get(index)
                .ok_or(Error::UnexpectedEndOfStream("BrtSlicerCacheNativeItem"))?;
            if item.kind() != kind::SLICER_CACHE_NATIVE_ITEM {
                return Err(invalid("slicer cache", "native source has no item record"));
            }
            let native = native(header, item.payload(), limits)?;
            index += 1;
            let end = records
                .get(index)
                .ok_or(Error::UnexpectedEndOfStream("BrtEndSlicerCacheNative"))?;
            if end.kind() != kind::END_SLICER_CACHE_NATIVE {
                return Err(invalid("slicer cache", "native source has no end record"));
            }
            empty(end, "BrtEndSlicerCacheNative")?;
            index += 1;
            Source::Native(native)
        },
        Some(kind) if kind == kind::BEGIN_TABLE_SLICER_CACHE => {
            let table = table(records[index].payload(), limits)?;
            index += 1;
            let end = records
                .get(index)
                .ok_or(Error::UnexpectedEndOfStream("BrtEndTableSlicerCache"))?;
            if end.kind() != kind::END_TABLE_SLICER_CACHE {
                return Err(invalid("slicer cache", "table source has no end record"));
            }
            empty(end, "BrtEndTableSlicerCache")?;
            index += 1;
            Source::Table(table)
        },
        Some(kind) if kind == kind::BEGIN_SLICER_CACHE_OLAP_IMPL => {
            return Err(Error::UnsupportedFeature(
                "OLAP slicer cache nested item ranges are outside the bounded XLSB slice",
            ));
        },
        _ => return Err(invalid("slicer cache", "missing supported source")),
    };

    let definition_end = records
        .get(index)
        .ok_or(Error::UnexpectedEndOfStream("BrtEndSlicerCacheDef"))?;
    if definition_end.kind() != kind::END_SLICER_CACHE_DEF {
        return Err(invalid("slicer cache", "missing BrtEndSlicerCacheDef"));
    }
    empty(definition_end, "BrtEndSlicerCacheDef")?;
    index += 1;
    let end = records
        .get(index)
        .ok_or(Error::UnexpectedEndOfStream("BrtEndSlicerCache"))?;
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
        alternate_content,
        name,
        hierarchy,
        pivot_tables: pivot,
        source,
    };
    validate_cache(&cache)?;
    Ok(cache)
}

fn write_wide(output: &mut Vec<u8>, value: &str, limits: crate::raw::Limits) -> Result<()> {
    Encoder::new(output, limits).write_wide_string(value)?;
    Ok(())
}

fn write_cache_inner(
    cache: &Cache,
    output: &mut Vec<u8>,
    limits: crate::raw::Limits,
) -> Result<()> {
    validate_cache(cache)?;
    let mut writer = Encoder::new(output, limits);
    writer.write_record(kind::BEGIN_SLICER_CACHE, &[])?;
    if !cache.alternate_content.is_empty() {
        let records = parse_all(&cache.alternate_content, limits)?;
        if cache_identity_end(&records, 0)? != records.len() {
            return Err(invalid(
                "slicer cache identity",
                "records follow the identity block",
            ));
        }
        writer.write_encoded_records(&cache.alternate_content)?;
    }
    let mut definition = Vec::new();
    write_wide(&mut definition, &cache.name, limits)?;
    write_wide(&mut definition, &cache.hierarchy, limits)?;
    writer.write_record(kind::BEGIN_SLICER_CACHE_DEF, &definition)?;
    if !cache.pivot_tables.is_empty() {
        let mut payload = Vec::new();
        Encoder::new(&mut payload, limits).write_u32(cache.pivot_tables.len() as u32)?;
        for pivot in &cache.pivot_tables {
            let mut item = Vec::new();
            let mut item_writer = Encoder::new(&mut item, limits);
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
            let mut header_writer = Encoder::new(&mut header, limits);
            header_writer.write_u32(0)?;
            header_writer.write_u32(native.cache_id)?;
            header_writer.write_u8(flags)?;
            writer.write_record(kind::BEGIN_SLICER_CACHE_NATIVE, &header)?;
            let mut items = Vec::with_capacity(4 + native.items.len() * 5);
            let mut item_writer = Encoder::new(&mut items, limits);
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
            let mut payload_writer = Encoder::new(&mut payload, limits);
            payload_writer.write_u32(0)?;
            payload_writer.write_u32(table.column)?;
            payload_writer.write_u32(table.table_id)?;
            payload_writer.write_u8(flags)?;
            writer.write_record(kind::BEGIN_TABLE_SLICER_CACHE, &payload)?;
            writer.write_record(kind::END_TABLE_SLICER_CACHE, &[])?;
        },
        Source::Olap(_) => {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_slicers",
                "OLAP slicer cache authoring is outside this XLSB route",
            ));
        },
    }
    writer.write_record(kind::END_SLICER_CACHE_DEF, &[])?;
    writer.write_record(kind::END_SLICER_CACHE, &[])?;
    Ok(())
}

/// Encode one slicer cache part.
pub fn write_cache(cache: &Cache, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    write_cache_inner(cache, &mut output, limits)?;
    Ok(output)
}

fn parse_view(payload: &[u8], limits: crate::raw::Limits) -> Result<View> {
    let mut cursor = Cursor::new(payload, "BrtBeginSlicer", limits);
    let flags = cursor.read_u8()?;
    if flags & 0xf0 != 0 {
        return Err(invalid("BrtBeginSlicer unused flags", "unexpected value"));
    }
    let view = View {
        alternate_content: Vec::new(),
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
pub fn parse_views(data: &[u8], limits: crate::raw::Limits) -> Result<Views> {
    let records = parse_all(data, limits)?;
    if records.len() < 2 || records[0].kind() != kind::BEGIN_SLICERS {
        return Err(invalid("slicer views", "missing BrtBeginSlicers"));
    }
    empty(&records[0], "BrtBeginSlicers")?;
    let mut views = Views::new();
    let mut index = 1;
    let mut ended = false;
    while index < records.len() {
        if records[index].kind() == kind::END_SLICERS {
            empty(&records[index], "BrtEndSlicers")?;
            index += 1;
            ended = true;
            break;
        }
        let mut prefix = Vec::new();
        if records[index].kind() == kind::AC_BEGIN {
            let start = records[index].offset();
            index = alternate_end(&records, index)?;
            let stop = records.get(index).map_or(data.len(), Record::offset);
            prefix.extend_from_slice(&data[start..stop]);
        }
        if records
            .get(index)
            .is_none_or(|record| record.kind() != kind::BEGIN_SLICER)
        {
            return Err(invalid(
                "slicer views",
                "unexpected record in view collection",
            ));
        }
        let mut view = parse_view(records[index].payload(), limits)?;
        view.alternate_content = prefix;
        index += 1;
        let end = records
            .get(index)
            .ok_or(Error::UnexpectedEndOfStream("BrtEndSlicer"))?;
        if end.kind() != kind::END_SLICER {
            return Err(invalid("slicer views", "missing BrtEndSlicer"));
        }
        empty(end, "BrtEndSlicer")?;
        views.items.push(view);
        index += 1;
    }
    if !ended {
        return Err(Error::UnexpectedEndOfStream("BrtEndSlicers"));
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
pub fn write_views(views: &Views, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    validate_views(views)?;
    let prefix_bytes = views
        .items
        .iter()
        .try_fold(0usize, |total, view| {
            total.checked_add(view.alternate_content.len())
        })
        .ok_or_else(|| invalid("slicer views", "alternate prefix size overflow"))?;
    if prefix_bytes > MAX_PART_BYTES {
        return Err(invalid(
            "slicer views",
            "alternate prefixes exceed part limit",
        ));
    }
    let mut output = Vec::new();
    let mut writer = Encoder::new(&mut output, limits);
    writer.write_record(kind::BEGIN_SLICERS, &[])?;
    for view in &views.items {
        if !view.alternate_content.is_empty() {
            let records = parse_all(&view.alternate_content, limits)?;
            if alternate_end(&records, 0)? != records.len() {
                return Err(invalid(
                    "slicer alternate prefix",
                    "records follow the alternate block",
                ));
            }
            writer.write_encoded_records(&view.alternate_content)?;
        }
        let mut payload = Vec::new();
        let mut flags = u8::from(view.caption_visible) | (u8::from(view.caption.is_some()) << 1);
        flags |= u8::from(view.style.is_some()) << 2;
        flags |= u8::from(view.locked_position) << 3;
        let mut payload_writer = Encoder::new(&mut payload, limits);
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
    if output.len() > MAX_PART_BYTES {
        return Err(invalid("slicer views", "part exceeds byte limit"));
    }
    Ok(output)
}

fn cache_identity_end(records: &[Record<'_>], start: usize) -> Result<usize> {
    let end = alternate_end(records, start)?;
    // Native Excel emits the 16-byte UID as kind 0x0c00. The cache's ACUID
    // slot holds exactly that record, not an arbitrary alternate collection.
    if end != start + 3
        || records[start + 1].kind().get() != 0x0c00
        || records[start + 1].payload().len() != 16
    {
        return Err(invalid(
            "slicer cache identity",
            "expected one native UID record",
        ));
    }
    Ok(end)
}

fn alternate_end(records: &[Record<'_>], start: usize) -> Result<usize> {
    let begin = records
        .get(start)
        .ok_or(Error::UnexpectedEndOfStream("slicer alternate prefix"))?;
    if begin.kind() != kind::AC_BEGIN || begin.payload().len() < 2 {
        return Err(invalid("slicer alternate prefix", "invalid begin record"));
    }
    let count = usize::from(u16::from_le_bytes([begin.payload()[0], begin.payload()[1]]));
    if count == 0 || begin.payload().len() != 2 + count * 4 {
        return Err(invalid(
            "slicer alternate prefix",
            "invalid product version array",
        ));
    }
    let mut index = start + 1;
    loop {
        let record = records
            .get(index)
            .ok_or(Error::UnexpectedEndOfStream("slicer alternate prefix"))?;
        if record.kind() == kind::AC_END {
            empty(record, "BrtACEnd")?;
            if index == start + 1 {
                return Err(invalid("slicer alternate prefix", "empty alternate block"));
            }
            return Ok(index + 1);
        }
        if matches!(
            record.kind(),
            kind::AC_BEGIN
                | kind::FRT_BEGIN
                | kind::FRT_END
                | kind::BEGIN_SLICER
                | kind::END_SLICER
                | kind::BEGIN_SLICERS
                | kind::END_SLICERS
        ) {
            return Err(invalid(
                "slicer alternate prefix",
                "unterminated or nested block",
            ));
        }
        index += 1;
    }
}
