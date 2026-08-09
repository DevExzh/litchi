//! Feature12 LTEXTERNALDATA table parsing.

use crate::Result;
use crate::list_object::codec::binary::{
    PendingFeature, parse_list_formula_extra_end, parse_range, parse_string, u16_at, u32_at,
    validate_frt,
};
use crate::list_object::model::{
    CachedDiskHeader, ExternalTableField, ExternalTableMetadata, ExternalTableVersion,
    ListColumnId, ListObject, ListObjectColumn, ListObjectFeatureVersion, ListObjectId,
    ListTotalAggregation, OpaqueListObjectFeature, validate_column_name, validate_name,
    validate_table_name,
};
use crate::list_object::{FEATURE12_RECORD_TYPE, ISF_LIST, MAX_FEATURE_BYTES, invalid};

impl ListObject {
    pub(in crate::list_object) fn parse_external_feature12(
        pending: PendingFeature,
    ) -> Result<Self> {
        let data = &pending.combined;
        let rt = FEATURE12_RECORD_TYPE;
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid external table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, rt, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, rt, "reserved2")? != 0
            || u16_at(data, 19, rt, "cref2")? != 1
            || u16_at(data, 25, rt, "reserved3")? != 0
            || parse_range(data, 27, rt)? != range
        {
            return Err(invalid(rt, "invalid Feature12 fixed fields"));
        }
        let declared = usize::try_from(u32_at(data, 21, rt, "cbFeatData")?)
            .map_err(|_error| invalid(rt, "cbFeatData overflows"))?;
        if declared != 0 && declared != data.len() - 35 {
            return Err(invalid(
                rt,
                "cbFeatData does not match external feature size",
            ));
        }
        let base = 35;
        if u32_at(data, base, rt, "lt")? != 3 {
            return Err(invalid(rt, "external parser requires LTEXTERNALDATA"));
        }
        let id = ListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, rt, "cbFSData")? != 64
            || u32_at(data, base + 44, rt, "lem")? != 0
        {
            return Err(invalid(
                rt,
                "invalid external TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let table_flags = crate::list_object::TableFlags::from_raw(flags);
        let version = ExternalTableVersion::from_code((flags >> 16) & 0xF)?;
        if flags & 0x0020_E7A0 != 0
            || flags & 0x4 != 0 && flags & 0x2 == 0
            || flags & 0x10 != 0 && flags & 0x8 == 0
            || flags & 0x2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid external table flags"));
        }
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "external field count must match table range"));
        }
        if flags & 0x0000_4000 != 0 {
            return Err(invalid(
                rt,
                "external table cannot load a SharePoint CSP name",
            ));
        }
        if flags & 0x0010_0000 != 0 {
            offset = parse_string(data, offset, rt, "entryId")?.1;
        }
        let mut columns = Vec::with_capacity(count);
        let mut fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = ListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            if u32_at(data, start + 4, rt, "lfdt")? != 0
                || u32_at(data, start + 8, rt, "lfxidt")? != 0
            {
                return Err(invalid(rt, "external field data types must be zero"));
            }
            let aggregation =
                ListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let aggregate_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_error| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_style = u32_at(data, start + 20, rt, "istnAgg")?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_error| invalid(rt, "insert format length overflows"))?;
            let insert_row_style = u32_at(data, start + 32, rt, "istnInsertRow")?;
            if field_flags & 0x4C != 0
                || field_flags & 0x40 != 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != ListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != ListTotalAggregation::None
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || (field_flags & 1 != 0) != (flags & 2 != 0)
            {
                return Err(invalid(rt, "invalid external field flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            validate_name(&source_name, "external source field name")?;
            let (caption, after_caption) = parse_string(data, after_source, rt, "strCaption")?;
            validate_column_name(&caption)?;
            let aggregate_end = after_caption
                .checked_add(aggregate_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..aggregate_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = aggregate_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(aggregate_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_error| invalid(rt, "AutoFilter length overflows"))?;
                if size > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + size)
                    .ok_or_else(|| invalid(rt, "AutoFilter length overflows"))?;
                let value = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                value
            } else {
                vec![0; 6]
            };
            let (total_formula, formula_extra) = if field_flags & 0x80 != 0 {
                let size = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if size == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let token_end = offset
                    .checked_add(2 + size)
                    .ok_or_else(|| invalid(rt, "total formula length overflows"))?;
                let tokens = data
                    .get(offset + 2..token_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = token_end;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            let query_field_id = u32_at(data, offset, rt, "qsif")?;
            if query_field_id == 0 {
                return Err(invalid(rt, "external qsif must be nonzero"));
            }
            offset += 4;
            let header_cache = if header == 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbdxfHdrDisk")?)
                    .map_err(|_error| invalid(rt, "header cache length overflows"))?;
                let format_end = offset
                    .checked_add(4 + size)
                    .ok_or_else(|| invalid(rt, "header cache length overflows"))?;
                data.get(offset..format_end)
                    .ok_or_else(|| invalid(rt, "truncated header cache"))?;
                let end = if field_flags & 0x200 != 0 {
                    parse_string(data, format_end, rt, "header style name")?.1
                } else {
                    format_end
                };
                let value = CachedDiskHeader::parse(
                    data[offset..end].to_vec(),
                    field_flags & 0x200 != 0,
                    rt,
                )?;
                offset = end;
                value
            } else {
                if field_flags & 0x200 != 0 {
                    return Err(invalid(rt, "header style name requires a CachedDiskHeader"));
                }
                CachedDiskHeader::empty()
            };
            let column = ListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
            fields.push(ExternalTableField {
                column_id: cid,
                source_name,
                query_field_id,
                aggregate_format,
                insert_row_format,
                auto_filter,
                formula_extra,
                header_cache,
                aggregate_style,
                insert_row_style,
                filter_hidden: field_flags & 2 != 0,
                total_array_formula: field_flags & 0x100 != 0,
                auto_create_calculated_column: field_flags & 0x800 != 0,
            });
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing external Feature12 data"));
        }
        let metadata = ExternalTableMetadata {
            version,
            build_number: u16_at(data, base + 24, rt, "rupBuild")?,
            fields,
        };
        metadata.validate()?;
        let opaque_feature = OpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
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
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: Some(metadata),
            source_metadata: None,
        })
    }
}
