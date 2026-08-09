//! BIFF8 `PivotTable` and `PivotCache` codecs.
//!
//! This module owns SX* record constants, bounded binary parsing, and the
//! lossless record helpers used by the semantic aggregate.

use super::model::{
    PageFieldEntry, PivotAdditionalExtension, PivotAxis, PivotAxisField, PivotCache,
    PivotCacheDateGroupUnit, PivotCacheDateGrouping, PivotCacheDateTime,
    PivotCacheDiscreteGrouping, PivotCacheError, PivotCacheField, PivotCacheGrouping,
    PivotCacheItem, PivotCacheNumericGrouping, PivotDataItem, PivotFunction, PivotItemType,
    PivotLayoutLine, PivotQueryTag, PivotSourceType, PivotViewDef, PivotViewEx9,
    PivotViewExtension, PivotViewField, PivotViewFieldExtension, PivotViewItem,
};
use crate::error::{Error, Result};
use litchi_core::binary;

pub(crate) fn cache_invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn cache_records(data: &[u8]) -> Result<Vec<(u16, &[u8])>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let header = data.get(offset..offset + 4).ok_or(Error::InvalidLength {
            expected: offset + 4,
            found: data.len(),
        })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(len))
            .ok_or_else(|| cache_invalid(kind, "PivotCache record length overflow"))?;
        let body = data.get(offset + 4..end).ok_or(Error::InvalidLength {
            expected: end,
            found: data.len(),
        })?;
        records.push((kind, body));
        offset = end;
    }
    Ok(records)
}

fn parse_cache_string(data: &[u8], record_type: u16) -> Result<(String, usize)> {
    if data.len() < 3 {
        return Err(Error::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }
    let count = usize::from(u16::from_le_bytes([data[0], data[1]]));
    let wide = data[2] & 1 != 0;
    let byte_count = count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| cache_invalid(record_type, "PivotCache string length overflow"))?;
    let end = 3usize
        .checked_add(byte_count)
        .ok_or_else(|| cache_invalid(record_type, "PivotCache string length overflow"))?;
    let chars = data.get(3..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let value = if wide {
        let units = chars
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_error| cache_invalid(record_type, "invalid UTF-16 PivotCache string"))?
    } else {
        chars.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, end))
}

fn parse_cache_item(record_type: u16, data: &[u8]) -> Result<PivotCacheItem> {
    let item = match record_type {
        0x00C9 => {
            if data.len() != 8 {
                return Err(Error::InvalidLength {
                    expected: 8,
                    found: data.len(),
                });
            }
            PivotCacheItem::Number(f64::from_le_bytes(data.try_into().unwrap()))
        },
        0x00CA => {
            if data.len() != 2 {
                return Err(Error::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            match u16::from_le_bytes(data.try_into().unwrap()) {
                0 => PivotCacheItem::Boolean(false),
                1 => PivotCacheItem::Boolean(true),
                value => {
                    return Err(cache_invalid(
                        record_type,
                        format!("invalid SXBOOLEAN value {value}"),
                    ));
                },
            }
        },
        0x00CB => {
            if data.len() != 2 {
                return Err(Error::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            PivotCacheItem::Error(PivotCacheError::try_from(u16::from_le_bytes(
                data.try_into().unwrap(),
            ))?)
        },
        0x00CC => {
            if data.len() != 2 {
                return Err(Error::InvalidLength {
                    expected: 2,
                    found: data.len(),
                });
            }
            PivotCacheItem::Number(f64::from(i16::from_le_bytes(data.try_into().unwrap())))
        },
        0x00CD => {
            let (value, used) = parse_cache_string(data, record_type)?;
            if used != data.len() {
                return Err(cache_invalid(record_type, "trailing SXSTRING payload"));
            }
            PivotCacheItem::String(value)
        },
        0x00CE => {
            if data.len() != 8 {
                return Err(Error::InvalidLength {
                    expected: 8,
                    found: data.len(),
                });
            }
            PivotCacheItem::DateTime(PivotCacheDateTime::try_new(
                u16::from_le_bytes([data[0], data[1]]),
                u16::from_le_bytes([data[2], data[3]]),
                data[4],
                data[5],
                data[6],
                data[7],
            )?)
        },
        0x00CF => {
            if !data.is_empty() {
                return Err(Error::InvalidLength {
                    expected: 0,
                    found: data.len(),
                });
            }
            PivotCacheItem::Empty
        },
        _ => {
            return Err(cache_invalid(
                record_type,
                "unexpected PivotCache item record",
            ));
        },
    };
    item.validate()?;
    Ok(item)
}

/// Parse one `_SX_DB_CUR/nnnn` `PivotCache` stream.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
/// # Panics
///
/// Panics only if an internal BIFF invariant has been violated.
pub fn parse_pivot_cache_stream(data: &[u8]) -> Result<PivotCache> {
    let records = cache_records(data)?;
    let (sxdb_type, sxdb) = records
        .first()
        .ok_or_else(|| cache_invalid(0x00C6, "empty PivotCache stream"))?;
    if *sxdb_type != 0x00C6 || sxdb.len() < 18 {
        return Err(cache_invalid(*sxdb_type, "PivotCache must start with SXDB"));
    }
    let record_count = u32::from_le_bytes(sxdb[0..4].try_into().unwrap());
    let stream_id = u16::from_le_bytes(sxdb[4..6].try_into().unwrap());
    let cache_flags = u16::from_le_bytes(sxdb[6..8].try_into().unwrap());
    let standard_field_count = usize::from(u16::from_le_bytes(sxdb[10..12].try_into().unwrap()));
    let field_count = usize::from(u16::from_le_bytes(sxdb[12..14].try_into().unwrap()));
    if standard_field_count > field_count {
        return Err(cache_invalid(
            0x00C6,
            "standard PivotCache field count exceeds total count",
        ));
    }
    if stream_id == 0 {
        return Err(cache_invalid(
            0x00C6,
            "PivotCache stream ID must be nonzero",
        ));
    }
    let mut position = 1usize;
    if records
        .get(position)
        .is_some_and(|(kind, _)| *kind == 0x0122)
    {
        position += 1;
    }
    let mut fields = Vec::with_capacity(field_count);
    for _ in 0..field_count {
        let (kind, body) = records
            .get(position)
            .ok_or_else(|| cache_invalid(0x00C7, "missing SXFDB"))?;
        if *kind != 0x00C7 || body.len() < 17 {
            return Err(cache_invalid(*kind, "expected valid SXFDB"));
        }
        let flags = u16::from_le_bytes(body[0..2].try_into().unwrap());
        let raw_parent = u16::from_le_bytes(body[2..4].try_into().unwrap());
        let raw_base = u16::from_le_bytes(body[4..6].try_into().unwrap());
        let group_count = usize::from(u16::from_le_bytes(body[8..10].try_into().unwrap()));
        let base_count = usize::from(u16::from_le_bytes(body[10..12].try_into().unwrap()));
        let original_count = usize::from(u16::from_le_bytes(body[12..14].try_into().unwrap()));
        let (name, used) = parse_cache_string(&body[14..], 0x00C7)?;
        if 14 + used != body.len() {
            return Err(cache_invalid(0x00C7, "trailing SXFDB payload"));
        }
        position += 1;
        let (kind, body) = records
            .get(position)
            .ok_or_else(|| cache_invalid(0x01BB, "missing SXFDBTYPE"))?;
        if *kind != 0x01BB || body != &[0, 0] {
            return Err(cache_invalid(*kind, "invalid SXFDBTYPE"));
        }
        position += 1;
        let mut group_items = Vec::with_capacity(group_count);
        for _ in 0..group_count {
            let (kind, body) = records
                .get(position)
                .ok_or_else(|| cache_invalid(0x00C7, "missing PivotCache group item"))?;
            group_items.push(parse_cache_item(*kind, body)?);
            position += 1;
        }
        let grouping = if let Some((0x00D9, body)) = records.get(position) {
            if body.len() != base_count * 2 {
                return Err(cache_invalid(
                    0x00D9,
                    "SXGROUPINFO size does not match base-item count",
                ));
            }
            let item_to_group = body
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            if item_to_group
                .iter()
                .any(|index| usize::from(*index) >= group_items.len())
            {
                return Err(cache_invalid(
                    0x00D9,
                    "SXGROUPINFO group ordinal is out of range",
                ));
            }
            position += 1;
            Some(PivotCacheGrouping::Discrete(PivotCacheDiscreteGrouping {
                base_field_index: raw_base,
                group_items: group_items.clone(),
                item_to_group,
            }))
        } else if let Some((0x00D8, body)) = records.get(position) {
            if body.len() != 2 {
                return Err(Error::InvalidLength {
                    expected: 2,
                    found: body.len(),
                });
            }
            let group_flags = u16::from_le_bytes((*body).try_into().unwrap());
            let group_type = (group_flags >> 2) & 0xF;
            position += 1;
            let mut limits = Vec::with_capacity(3);
            for _ in 0..3 {
                let (kind, body) = records
                    .get(position)
                    .ok_or_else(|| cache_invalid(0x00D8, "missing SXNUMGROUP limit item"))?;
                limits.push(parse_cache_item(*kind, body)?);
                position += 1;
            }
            if group_type == 8 {
                let numbers = limits
                    .iter()
                    .map(|item| match item {
                        PivotCacheItem::Number(value) => Ok(*value),
                        _ => Err(cache_invalid(
                            0x00D8,
                            "numeric grouping limits must be numeric",
                        )),
                    })
                    .collect::<Result<Vec<_>>>()?;
                Some(PivotCacheGrouping::Numeric(PivotCacheNumericGrouping {
                    start: numbers[0],
                    end: numbers[1],
                    step: numbers[2],
                    auto_start: group_flags & 1 != 0,
                    auto_end: group_flags & 2 != 0,
                    group_items: group_items.clone(),
                }))
            } else {
                let start = match limits[0] {
                    PivotCacheItem::DateTime(value) => value,
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping start must be SXDATETIME",
                        ));
                    },
                };
                let end = match limits[1] {
                    PivotCacheItem::DateTime(value) => value,
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping end must be SXDATETIME",
                        ));
                    },
                };
                let step = match limits[2] {
                    PivotCacheItem::Number(value)
                        if value >= 1.0 && value <= f64::from(u16::MAX) && value.fract() == 0.0 =>
                    {
                        crate::utils::saturating_f64_to_u16(value)
                    },
                    _ => {
                        return Err(cache_invalid(
                            0x00D8,
                            "date grouping step must be a positive integer",
                        ));
                    },
                };
                Some(PivotCacheGrouping::Date(PivotCacheDateGrouping {
                    unit: PivotCacheDateGroupUnit::try_from(group_type)?,
                    start,
                    end,
                    step,
                    auto_start: group_flags & 1 != 0,
                    auto_end: group_flags & 2 != 0,
                    group_items: group_items.clone(),
                }))
            }
        } else {
            None
        };
        let mut items = Vec::with_capacity(original_count);
        for _ in 0..original_count {
            let (kind, body) = records
                .get(position)
                .ok_or_else(|| cache_invalid(0x00C7, "missing shared PivotCache item"))?;
            items.push(parse_cache_item(*kind, body)?);
            position += 1;
        }
        if (flags & 1 != 0) == (items.is_empty() && group_items.is_empty()) {
            return Err(cache_invalid(0x00C7, "SXFDB item flag/count mismatch"));
        }
        fields.push(PivotCacheField {
            name,
            flags,
            group_parent: (flags & 0x0008 != 0).then_some(raw_parent),
            group_base: matches!(grouping, Some(PivotCacheGrouping::Discrete(_)))
                .then_some(raw_base),
            items,
            grouping,
        });
    }
    let mut rows = Vec::new();
    while let Some((kind, body)) = records.get(position) {
        if *kind == 0x000A {
            position += 1;
            break;
        }
        if *kind != 0x00C8 {
            return Err(cache_invalid(*kind, "expected SXINDEXLIST or EOF"));
        }
        let mut body_offset = 0usize;
        let mut row = Vec::with_capacity(fields.len());
        position += 1;
        for field in fields.iter().take(standard_field_count) {
            if field.items.is_empty() {
                let (item_type, item_body) = records
                    .get(position)
                    .ok_or_else(|| cache_invalid(0x00C8, "missing inline PivotCache row item"))?;
                row.push(parse_cache_item(*item_type, item_body)?);
                position += 1;
            } else {
                let width = if field.flags & 0x0200 != 0 { 2 } else { 1 };
                let encoded =
                    body.get(body_offset..body_offset + width)
                        .ok_or(Error::InvalidLength {
                            expected: body_offset + width,
                            found: body.len(),
                        })?;
                let index = if width == 2 {
                    usize::from(u16::from_le_bytes(encoded.try_into().unwrap()))
                } else {
                    usize::from(encoded[0])
                };
                row.push(
                    field
                        .items
                        .get(index)
                        .ok_or_else(|| {
                            cache_invalid(0x00C8, "SXINDEXLIST item index out of range")
                        })?
                        .clone(),
                );
                body_offset += width;
            }
        }
        if body_offset != body.len() {
            return Err(cache_invalid(
                0x00C8,
                "SXINDEXLIST payload size does not match fields",
            ));
        }
        rows.push(row);
    }
    if position != records.len() {
        return Err(cache_invalid(
            0x000A,
            "trailing PivotCache records after EOF",
        ));
    }
    if cache_flags & 1 != 0 && rows.len() != record_count as usize {
        return Err(cache_invalid(0x00C6, "saved PivotCache row count mismatch"));
    }
    Ok(PivotCache {
        stream_id,
        flags: cache_flags,
        record_count,
        fields,
        rows,
    })
}

/// SXVIEW record type.
pub const SXVIEW_TYPE: u16 = 0x00B0;
/// SXVD (View Fields) record type.
pub const SXVD_TYPE: u16 = 0x00B1;
/// SXVI (View Item) record type.
pub const SXVI_TYPE: u16 = 0x00B2;
/// SXPI (Page Item) record type.
pub const SXPI_TYPE: u16 = 0x00B6;
/// SXDI (Data Item) record type.
pub const SXDI_TYPE: u16 = 0x00C5;
/// SXVS (View Source) record type.
pub const SXVS_TYPE: u16 = 0x00E3;
pub const SXIVD_TYPE: u16 = 0x00B4;
pub const SXLI_TYPE: u16 = 0x00B5;
pub const SXEX_TYPE: u16 = 0x00F1;
pub const SXVDEX_TYPE: u16 = 0x0100;
pub const QSI_SX_TAG_TYPE: u16 = 0x0802;
pub const SXVIEWEX9_TYPE: u16 = 0x0810;
pub const SXADDL_TYPE: u16 = 0x0864;

const DATA_LAYOUT_FIELD: u16 = 0xFFFE;
pub(crate) const MAX_PIVOT_VIEWS_PER_SHEET: usize = 1_024;
pub(crate) const MAX_PIVOT_FIELDS: usize = 4_096;
pub(crate) const MAX_PIVOT_ITEMS: usize = 1_048_576;
pub(crate) const MAX_PIVOT_EXTENSION_BYTES: usize = 1_048_576;

pub(crate) const fn is_worksheet_view_record(record_type: u16) -> bool {
    matches!(
        record_type,
        SXVIEW_TYPE
            | SXVD_TYPE
            | SXVI_TYPE
            | SXIVD_TYPE
            | SXLI_TYPE
            | SXPI_TYPE
            | SXDI_TYPE
            | SXEX_TYPE
            | SXVDEX_TYPE
            | QSI_SX_TAG_TYPE
            | SXVIEWEX9_TYPE
            | SXADDL_TYPE
    )
}

/// Parse an SXVIEW record.
///
/// Layout (Apache POI `ViewDefinitionRecord`):
/// ```text
///  0  u16  rwFirst
///  2  u16  rwLast
///  4  u16  colFirst
///  6  u16  colLast
///  8  u16  rwFirstHead
/// 10  u16  rwFirstData
/// 12  u16  colFirstData
/// 14  u16  cDimRw       (row field count)
/// 16  u16  cDimCol
/// 18  u16  cDimPg
/// 20  u16  cDimData
/// 22  u16  cRw          (data row count)
/// 24  u16  cDim         (total field count)
/// 26  u16  cItm         (unused)
/// 28  u16  cITMData     (unused)
/// 30  u16  sxaxis4Data
/// 32  u16  ipos4Data
/// 34  u16  cchName      (length of name)
/// 36  u16  cchData      (length of data field name)
/// 38  var  name (XLUnicodeStringNoCch)
///     var  dataField (XLUnicodeStringNoCch)
/// ```
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxview(data: &[u8]) -> Result<PivotViewDef> {
    if data.len() < 44 {
        return Err(Error::InvalidLength {
            expected: 44,
            found: data.len(),
        });
    }

    let first_row = binary::read_u16_le_at(data, 0)?;
    let last_row = binary::read_u16_le_at(data, 2)?;
    let first_col = binary::read_u16_le_at(data, 4)?;
    let last_col = binary::read_u16_le_at(data, 6)?;
    let first_header_row = binary::read_u16_le_at(data, 8)?;
    let first_data_row = binary::read_u16_le_at(data, 10)?;
    let first_data_col = binary::read_u16_le_at(data, 12)?;
    let cache_index = binary::read_u16_le_at(data, 14)?;
    if binary::read_u16_le_at(data, 16)? != 0 {
        return Err(cache_invalid(SXVIEW_TYPE, "nonzero SXVIEW reserved field"));
    }
    let data_axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 18)?)?;
    let data_position = binary::read_u16_le_at(data, 20)?;
    let field_count = binary::read_u16_le_at(data, 22)?;
    let row_field_count = binary::read_u16_le_at(data, 24)?;
    let col_field_count = binary::read_u16_le_at(data, 26)?;
    let page_field_count = binary::read_u16_le_at(data, 28)?;
    let data_field_count = binary::read_u16_le_at(data, 30)?;
    let data_row_count = binary::read_u16_le_at(data, 32)?;
    let data_col_count = binary::read_u16_le_at(data, 34)?;
    let flags = binary::read_u16_le_at(data, 36)?;
    let auto_format_index = binary::read_u16_le_at(data, 38)?;
    let cch_name = usize::from(binary::read_u16_le_at(data, 40)?);
    let cch_data = usize::from(binary::read_u16_le_at(data, 42)?);
    if usize::from(field_count) > MAX_PIVOT_FIELDS {
        return Err(cache_invalid(
            SXVIEW_TYPE,
            "PivotTable field count exceeds resource bound",
        ));
    }
    if first_row > last_row
        || first_col > last_col
        || first_header_row < first_row
        || first_header_row > last_row
        || first_data_row < first_row
        || first_data_row > last_row
        || first_data_col < first_col
        || first_data_col > last_col
    {
        return Err(cache_invalid(
            SXVIEW_TYPE,
            "invalid or reversed PivotTable output range",
        ));
    }

    let mut offset = 44;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;
    let data_field_name = read_xl_string_no_cch(data, &mut offset, cch_data)?;
    if offset != data.len() {
        return Err(cache_invalid(SXVIEW_TYPE, "trailing SXVIEW payload"));
    }

    Ok(PivotViewDef {
        first_row,
        last_row,
        first_col,
        last_col,
        first_header_row,
        first_data_row,
        first_data_col,
        cache_index,
        row_field_count,
        col_field_count,
        page_field_count,
        data_field_count,
        data_row_count,
        data_col_count,
        field_count,
        data_axis,
        data_position,
        flags,
        auto_format_index,
        name,
        data_field_name,
    })
}

/// Parse an SXVD record.
///
/// Layout:
/// ```text
///  0  u16  sxaxis   (axis)
///  2  u16  cSub     (subtotal count)
///  4  u16  grbitSub (subtotal flags)
///  6  u16  cItm     (item count)
///  8  u16  cchName  (0xFFFF = not present)
/// 10  var  name (XLUnicodeStringNoCch)  — only if cchName != 0xFFFF
/// ```
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxvd(data: &[u8]) -> Result<PivotViewField> {
    if data.len() < 10 {
        return Err(Error::InvalidLength {
            expected: 10,
            found: data.len(),
        });
    }

    let axis = PivotAxis::from_u16(binary::read_u16_le_at(data, 0)?)?;
    let subtotal_count = binary::read_u16_le_at(data, 2)?;
    let subtotal_flags = binary::read_u16_le_at(data, 4)?;
    let item_count = binary::read_u16_le_at(data, 6)?;
    let cch_name = binary::read_u16_le_at(data, 8)?;

    let mut offset = 10;
    let name = if cch_name == 0xFFFF {
        None
    } else {
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    };

    if offset != data.len() {
        return Err(cache_invalid(SXVD_TYPE, "trailing SXVD payload"));
    }
    Ok(PivotViewField {
        axis,
        subtotal_count,
        subtotal_flags,
        item_count,
        name,
        items: Vec::with_capacity(usize::from(item_count)),
        extension: None,
        additional_extensions: Vec::new(),
    })
}

/// Parse an SXVI record.
///
/// Layout:
/// ```text
///  0  u16  itmType
///  2  u16  grbitItem
///  4  u16  iCache
///  6  u16  cchName  (0xFFFF = not present)
///  8  var  name
/// ```
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxvi(data: &[u8]) -> Result<PivotViewItem> {
    if data.len() < 8 {
        return Err(Error::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }

    let item_type = PivotItemType::from_u16(binary::read_u16_le_at(data, 0)?);
    let flags = binary::read_u16_le_at(data, 2)?;
    let cache_index = binary::read_u16_le_at(data, 4)?;
    let cch_name = binary::read_u16_le_at(data, 6)?;

    let mut offset = 8;
    let name = if cch_name == 0xFFFF {
        None
    } else {
        Some(read_xl_string_no_cch(data, &mut offset, cch_name as usize)?)
    };

    if offset != data.len() {
        return Err(cache_invalid(SXVI_TYPE, "trailing SXVI payload"));
    }
    Ok(PivotViewItem {
        item_type,
        flags,
        cache_index,
        name,
    })
}

/// Parse an SXDI record.
///
/// Layout (POI `DataItemRecord`):
/// ```text
///  0  u16  isxvdData   (source field index)
///  2  u16  iiftab      (aggregation function)
///  4  u16  df          (display format)
///  6  u16  isxvd       (base field index)
///  8  u16  isxvi       (base item index)
/// 10  u16  ifmt        (number format)
/// 12  u16  cchName
/// 14  var  name
/// ```
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxdi(data: &[u8]) -> Result<PivotDataItem> {
    if data.len() < 14 {
        return Err(Error::InvalidLength {
            expected: 14,
            found: data.len(),
        });
    }

    let source_field_index = binary::read_u16_le_at(data, 0)?;
    let function = PivotFunction::from_u16(binary::read_u16_le_at(data, 2)?);
    let display_format = binary::read_u16_le_at(data, 4)?;
    let base_field_index = binary::read_u16_le_at(data, 6)?;
    let base_item_index = binary::read_u16_le_at(data, 8)?;
    let num_format_index = binary::read_u16_le_at(data, 10)?;
    let cch_name = binary::read_u16_le_at(data, 12)? as usize;

    let mut offset = 14;
    let name = read_xl_string_no_cch(data, &mut offset, cch_name)?;
    if offset != data.len() {
        return Err(cache_invalid(SXDI_TYPE, "trailing SXDI payload"));
    }

    Ok(PivotDataItem {
        source_field_index,
        function,
        display_format,
        base_field_index,
        base_item_index,
        num_format_index,
        name,
    })
}

/// Parse an SXVS record (2 bytes: source type).
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxvs(data: &[u8]) -> Result<PivotSourceType> {
    if data.len() != 2 {
        return Err(Error::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(PivotSourceType::from_u16(binary::read_u16_le_at(data, 0)?))
}

/// Parse an SXPI record.
///
/// Each entry is 6 bytes: `(isxvi: u16, isxvd: u16, idObj: u16)`.
/// The number of entries is `data.len() / 6`.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxpi(data: &[u8]) -> Result<Vec<PageFieldEntry>> {
    if !data.len().is_multiple_of(6) {
        return Err(cache_invalid(
            SXPI_TYPE,
            "SXPI length is not a multiple of six",
        ));
    }
    let entry_count = data.len() / 6;
    let mut entries = Vec::with_capacity(entry_count);

    for i in 0..entry_count {
        let offset = i * 6;
        entries.push(PageFieldEntry {
            field_index: binary::read_u16_le_at(data, offset)?,
            item_index: binary::read_u16_le_at(data, offset + 2)?,
            object_id: binary::read_u16_le_at(data, offset + 4)?,
        });
    }

    Ok(entries)
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxivd(data: &[u8]) -> Result<Vec<PivotAxisField>> {
    if !data.len().is_multiple_of(2) {
        return Err(cache_invalid(SXIVD_TYPE, "SXIVD length must be even"));
    }
    data.chunks_exact(2)
        .map(|bytes| match u16::from_le_bytes([bytes[0], bytes[1]]) {
            DATA_LAYOUT_FIELD => Ok(PivotAxisField::DataLayout),
            value if value != u16::MAX => Ok(PivotAxisField::Field(value)),
            _ => Err(cache_invalid(SXIVD_TYPE, "invalid SXIVD field ordinal")),
        })
        .collect()
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
/// # Panics
///
/// Panics only if an internal BIFF invariant has been violated.
pub fn parse_sxvdex(data: &[u8]) -> Result<PivotViewFieldExtension> {
    if data.len() < 20 {
        return Err(Error::InvalidLength {
            expected: 20,
            found: data.len(),
        });
    }
    let cch = binary::read_u16_le_at(data, 10)?;
    let mut offset = 20;
    let subtotal_name = if cch == u16::MAX {
        None
    } else {
        Some(read_xl_string_no_cch(data, &mut offset, usize::from(cch))?)
    };
    if offset != data.len() {
        return Err(cache_invalid(SXVDEX_TYPE, "trailing SXVDEx payload"));
    }
    Ok(PivotViewFieldExtension {
        flags: binary::read_u32_le_at(data, 0)?,
        auto_sort_data_index: match binary::read_u16_le_at(data, 4)? {
            u16::MAX => None,
            value => Some(value),
        },
        auto_show_data_index: match binary::read_u16_le_at(data, 6)? {
            u16::MAX => None,
            value => Some(value),
        },
        number_format_index: binary::read_u16_le_at(data, 8)?,
        subtotal_name,
        reserved: data[12..20].try_into().expect("fixed length checked"),
    })
}

pub(crate) fn parse_sxli(
    data: &[u8],
    expected_lines: usize,
    max_indices: usize,
) -> Result<Vec<PivotLayoutLine>> {
    if expected_lines == 0 || !data.len().is_multiple_of(expected_lines) {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI byte length is inconsistent with its declared line count",
        ));
    }
    let line_size = data.len() / expected_lines;
    if line_size < 8 || !(line_size - 8).is_multiple_of(2) {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI has an invalid fixed line size",
        ));
    }
    let index_count = (line_size - 8) / 2;
    if index_count > max_indices {
        return Err(cache_invalid(
            SXLI_TYPE,
            "SXLI item-index count exceeds the PivotTable field count",
        ));
    }
    let mut lines = Vec::with_capacity(expected_lines);
    for line in data.chunks_exact(line_size) {
        let declared_max = usize::from(u16::from_le_bytes([line[4], line[5]]));
        if declared_max > index_count && declared_max != usize::from(u16::MAX) {
            return Err(cache_invalid(
                SXLI_TYPE,
                "SXLI declared item ordinal exceeds its line payload",
            ));
        }
        let indices = line[8..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect();
        lines.push(PivotLayoutLine {
            repeated_item_count: u16::from_le_bytes([line[0], line[1]]),
            item_type: u16::from_le_bytes([line[2], line[3]]),
            custom_name_flags: u16::from_le_bytes([line[6], line[7]]),
            item_indices: indices,
        });
    }
    Ok(lines)
}

fn parse_optional_sx_string(data: &[u8], offset: &mut usize, cch: u16) -> Result<Option<String>> {
    if cch == u16::MAX {
        Ok(None)
    } else {
        read_xl_string_no_cch(data, offset, usize::from(cch)).map(Some)
    }
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxex(data: &[u8]) -> Result<PivotViewExtension> {
    if data.len() < 24 {
        return Err(Error::InvalidLength {
            expected: 24,
            found: data.len(),
        });
    }
    let lengths = [2usize, 4, 6, 18, 20, 22]
        .map(|offset| u16::from_le_bytes([data[offset], data[offset + 1]]));
    let mut offset = 24;
    let error_string = parse_optional_sx_string(data, &mut offset, lengths[0])?;
    let null_string = parse_optional_sx_string(data, &mut offset, lengths[1])?;
    let tag = parse_optional_sx_string(data, &mut offset, lengths[2])?;
    let page_field_style = parse_optional_sx_string(data, &mut offset, lengths[3])?;
    let table_style = parse_optional_sx_string(data, &mut offset, lengths[4])?;
    let vacate_style = parse_optional_sx_string(data, &mut offset, lengths[5])?;
    if offset != data.len() {
        return Err(cache_invalid(SXEX_TYPE, "trailing SXEX payload"));
    }
    Ok(PivotViewExtension {
        format_count: binary::read_u16_le_at(data, 0)?,
        select_count: binary::read_u16_le_at(data, 8)?,
        page_rows: binary::read_u16_le_at(data, 10)?,
        page_cols: binary::read_u16_le_at(data, 12)?,
        flags: binary::read_u32_le_at(data, 14)?,
        error_string,
        null_string,
        tag,
        page_field_style,
        table_style,
        vacate_style,
    })
}

fn read_xl_unicode_string(data: &[u8], offset: &mut usize) -> Result<String> {
    let cch_end = offset
        .checked_add(2)
        .ok_or_else(|| cache_invalid(QSI_SX_TAG_TYPE, "string offset overflow"))?;
    let cch_bytes = data.get(*offset..cch_end).ok_or(Error::InvalidLength {
        expected: cch_end,
        found: data.len(),
    })?;
    *offset = cch_end;
    read_xl_string_no_cch(
        data,
        offset,
        usize::from(u16::from_le_bytes([cch_bytes[0], cch_bytes[1]])),
    )
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_qsi_sx_tag(data: &[u8]) -> Result<PivotQueryTag> {
    if data.len() < 19 {
        return Err(Error::InvalidLength {
            expected: 19,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != QSI_SX_TAG_TYPE || data[14] != 16 {
        return Err(cache_invalid(
            QSI_SX_TAG_TYPE,
            "invalid QsiSxTag FRT header",
        ));
    }
    let mut offset = 16;
    let table_name = read_xl_unicode_string(data, &mut offset)?;
    Ok(PivotQueryTag {
        table_type: binary::read_u16_le_at(data, 4)?,
        flags: binary::read_u16_le_at(data, 6)?,
        options: binary::read_u32_le_at(data, 8)?,
        last_refresh_version: data[12],
        minimum_refresh_version: data[13],
        first_created_version: data[15],
        table_name,
        trailing_payload: data[offset..].to_vec(),
    })
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxviewex9(data: &[u8]) -> Result<PivotViewEx9> {
    if data.len() < 17 {
        return Err(Error::InvalidLength {
            expected: 17,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != SXVIEWEX9_TYPE {
        return Err(cache_invalid(
            SXVIEWEX9_TYPE,
            "invalid SXVIEWEX9 FRT header",
        ));
    }
    let mut offset = 14;
    let grand_total_name = read_xl_unicode_string(data, &mut offset)?;
    if offset != data.len() {
        return Err(cache_invalid(SXVIEWEX9_TYPE, "trailing SXVIEWEX9 payload"));
    }
    Ok(PivotViewEx9 {
        frt_flags: binary::read_u16_le_at(data, 2)?,
        report_flags: binary::read_u32_le_at(data, 4)?,
        view_flags: binary::read_u32_le_at(data, 8)?,
        auto_format_index: binary::read_u16_le_at(data, 12)?,
        grand_total_name,
    })
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_sxaddl(data: &[u8]) -> Result<PivotAdditionalExtension> {
    if data.len() < 6 {
        return Err(Error::InvalidLength {
            expected: 6,
            found: data.len(),
        });
    }
    if binary::read_u16_le_at(data, 0)? != SXADDL_TYPE {
        return Err(cache_invalid(SXADDL_TYPE, "invalid SXADDL FRT header"));
    }
    Ok(PivotAdditionalExtension {
        reserved: binary::read_u16_le_at(data, 2)?,
        class: data[4],
        kind: data[5],
        payload: data[6..].to_vec(),
    })
}

/// Read an `XLUnicodeStringNoCch`: 1-byte flags then `cch` chars.
fn read_xl_string_no_cch(data: &[u8], offset: &mut usize, cch: usize) -> Result<String> {
    if cch == 0 {
        let end = offset
            .checked_add(1)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        data.get(*offset..end).ok_or(Error::InvalidLength {
            expected: end,
            found: data.len(),
        })?;
        *offset = end;
        return Ok(String::new());
    }

    if *offset >= data.len() {
        return Err(Error::InvalidLength {
            expected: *offset + 1,
            found: data.len(),
        });
    }

    let flags = data[*offset];
    *offset += 1;
    let is_utf16 = flags & 0x01 != 0;

    if is_utf16 {
        let byte_len = cch
            .checked_mul(2)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string size overflow"))?;
        let end = offset
            .checked_add(byte_len)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let words: Vec<u16> = data[*offset..end]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        *offset = end;
        String::from_utf16(&words)
            .map_err(|e| Error::InvalidData(format!("Invalid UTF-16 in pivot string: {e}")))
    } else {
        let end = offset
            .checked_add(cch)
            .ok_or_else(|| cache_invalid(SXVIEW_TYPE, "pivot string offset overflow"))?;
        if end > data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        let s: String = data[*offset..end].iter().map(|&b| b as char).collect();
        *offset = end;
        Ok(s)
    }
}
