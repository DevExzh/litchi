//! Ordinary Feature11/Feature12 table parsing.

use crate::Result;
use crate::list_object::codec::binary::{parse_range, parse_string, u16_at, u32_at, validate_frt};
use crate::list_object::model::*;
use crate::list_object::{
    FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST, MAX_FEATURE_BYTES, invalid,
};

impl ListObject {
    pub(in crate::list_object) fn parse_feature(data: &[u8], rt: u16) -> Result<Self> {
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, FEATURE11_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, FEATURE11_RECORD_TYPE, "reserved2")? != 0
            || u16_at(data, 19, FEATURE11_RECORD_TYPE, "cref2")? != 1
            || u16_at(data, 25, FEATURE11_RECORD_TYPE, "reserved3")? != 0
            || parse_range(data, 27, FEATURE11_RECORD_TYPE)? != range
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid Feature11 fixed fields",
            ));
        }
        let base = 35;
        let source_type = u32_at(data, base, FEATURE11_RECORD_TYPE, "lt")?;
        if matches!(source_type, 1 | 2) {
            return Self::parse_source_feature(data, rt, source_type);
        }
        if source_type != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported table source type",
            ));
        }
        let id = ListObjectId::try_new(u32_at(data, base + 4, FEATURE11_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, base + 8, FEATURE11_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, base + 12, FEATURE11_RECORD_TYPE, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, FEATURE11_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, FEATURE11_RECORD_TYPE, "flags")?;
        let table_flags = crate::list_object::TableFlags::from_raw(flags);
        if flags & 0x0020_E320 != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported external table flags",
            ));
        }
        let (name, mut offset) = parse_string(data, base + 64, FEATURE11_RECORD_TYPE, "rgbName")?;
        let count = usize::from(u16_at(data, offset, FEATURE11_RECORD_TYPE, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) {
            return Err(invalid(FEATURE11_RECORD_TYPE, "invalid table column count"));
        }
        if flags & 0x0010_0000 != 0 {
            let (entry, next) = parse_string(data, offset, FEATURE11_RECORD_TYPE, "entryId")?;
            if entry != id.value().to_string() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "entryId mismatch"));
            }
            offset = next;
        }
        let mut columns = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid =
                ListColumnId::try_new(u32_at(data, start, FEATURE11_RECORD_TYPE, "idField")?)?;
            if u32_at(data, start + 4, FEATURE11_RECORD_TYPE, "lfdt")? != 0
                || u32_at(data, start + 8, FEATURE11_RECORD_TYPE, "lfxidt")? != 0
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external column data is unsupported",
                ));
            }
            let agg = u32_at(data, start + 16, FEATURE11_RECORD_TYPE, "cbFmtAgg")? as usize;
            let cflags = u32_at(data, start + 24, FEATURE11_RECORD_TYPE, "column flags")?;
            let insert = u32_at(data, start + 28, FEATURE11_RECORD_TYPE, "cbFmtInsert")? as usize;
            let aggregation = ListTotalAggregation::from_code(u32_at(
                data,
                start + 12,
                FEATURE11_RECORD_TYPE,
                "ilta",
            )?)?;
            // Feature11 permits calculated columns (fAutoCreateCalcCol) but not
            // Feature12-only XML/Web mappings or loaded total formulas/strings.
            let forbidden = 0x4c
                | if rt == FEATURE11_RECORD_TYPE {
                    0x580
                } else {
                    0
                };
            if cflags & forbidden != 0 || cflags & 0x100 != 0 {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external, array, or reserved column data is unsupported",
                ));
            }
            let (_, after_field) =
                parse_string(data, start + 36, FEATURE11_RECORD_TYPE, "field name")?;
            let (caption, after_caption) =
                parse_string(data, after_field, FEATURE11_RECORD_TYPE, "caption")?;
            offset = after_caption
                .checked_add(agg + insert)
                .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column size overflow"))?;
            if cflags & 1 != 0 {
                let n = u32_at(data, offset, FEATURE11_RECORD_TYPE, "AutoFilter size")? as usize;
                offset = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "AutoFilter size overflow"))?;
            }
            let total_formula = if cflags & 0x80 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if n == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let end = offset + 2 + n;
                let value = data
                    .get(offset + 2..end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = end;
                Some(value)
            } else {
                None
            };
            let total_string = if cflags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "total string")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            if cflags & 0x200 != 0 {
                offset = parse_string(data, offset, FEATURE11_RECORD_TYPE, "cached style")?.1;
            }
            if offset > data.len() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "truncated column data"));
            }
            let column = ListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
        }
        if offset != data.len() {
            return Err(invalid(FEATURE11_RECORD_TYPE, "trailing Feature11 data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: table_flags.auto_filter(),
            table_flags,
            comment: String::new(),
            feature_version: if rt == FEATURE12_RECORD_TYPE {
                ListObjectFeatureVersion::Feature12
            } else {
                ListObjectFeatureVersion::Feature11
            },
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
}
