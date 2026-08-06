//! Feature11/Feature12 table feature parsing.

use super::super::super::model::*;
use super::super::super::model::{validate_column_name, validate_name, validate_table_name};
use super::super::super::{
    FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST, MAX_FEATURE_BYTES, invalid,
};
use super::super::binary::{
    PendingFeature, parse_formula, parse_list_formula_extra_end, parse_range, parse_string,
    parse_web_info, u16_at, u32_at, validate_frt,
};
use crate::Result;

impl ListObject {
    pub(in crate::list_object) fn parse_opaque_feature12(pending: PendingFeature) -> Result<Self> {
        let data = &pending.combined;
        if data.len() < 108 {
            return Err(invalid(FEATURE12_RECORD_TYPE, "truncated opaque Feature12"));
        }
        validate_frt(data, FEATURE12_RECORD_TYPE, true)?;
        if u16_at(data, 12, FEATURE12_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 55, FEATURE12_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 fixed fields",
            ));
        }
        let range = parse_range(data, 4, FEATURE12_RECORD_TYPE)?;
        let id = ListObjectId::try_new(u32_at(data, 39, FEATURE12_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, 43, FEATURE12_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, 47, FEATURE12_RECORD_TYPE, "crwTotals")?;
        if header > 1 || totals > 1 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 row flags",
            ));
        }
        let flags = u32_at(data, 63, FEATURE12_RECORD_TYPE, "flags")?;
        let (name, _) = parse_string(data, 99, FEATURE12_RECORD_TYPE, "rgbName")?;
        let opaque_feature = OpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns: Vec::new(),
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
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
            .map_err(|_| invalid(rt, "cbFeatData overflows"))?;
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
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_style = u32_at(data, start + 20, rt, "istnAgg")?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert format length overflows"))?;
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
                    .map_err(|_| invalid(rt, "AutoFilter length overflows"))?;
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
                    .map_err(|_| invalid(rt, "header cache length overflows"))?;
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
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: Some(metadata),
            source_metadata: None,
        })
    }
    fn parse_source_feature(data: &[u8], rt: u16, lt: u32) -> Result<Self> {
        if !matches!(rt, FEATURE11_RECORD_TYPE | FEATURE12_RECORD_TYPE) {
            return Err(invalid(
                rt,
                "Web/XML source table requires Feature11 or Feature12",
            ));
        }
        let range = parse_range(data, 4, rt)?;
        let base = 35;
        let id = ListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1 || totals > 1 || u32_at(data, base + 20, rt, "cbFSData")? != 64 {
            return Err(invalid(rt, "invalid Web/XML TableFeatureType fixed fields"));
        }
        let build = u16_at(data, base + 24, rt, "rupBuild")?;
        let ignored_fixed_word = u16_at(data, base + 26, rt, "unused1")?;
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let version = ExternalTableVersion::from_code((flags >> 16) & 0xf)?;
        if flags & 0x0000_0480 != 0
            || flags & 4 != 0 && flags & 2 == 0
            || flags & 0x10 != 0 && flags & 8 == 0
            || flags & 2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid Web/XML table flags"));
        }
        let single = flags & 0x200 != 0;
        if single && (lt != 2 || header != 0 || totals != 0 || range.column_count() != 1) {
            return Err(invalid(rt, "invalid single-cell XML table"));
        }
        if lt == 2 && flags & 0x0020_e120 != 0 {
            return Err(invalid(rt, "Web-only flags occur on XML table"));
        }
        let mut ignored_fixed_tail = [0; 32];
        let (cache_position, cache_size, cache_characters, edit_mode, hash) = if lt == 1 {
            let mut hash = [0; 16];
            hash.copy_from_slice(
                data.get(base + 48..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated Web hash parameters"))?,
            );
            (
                u32_at(data, base + 32, rt, "cache position")?,
                u32_at(data, base + 36, rt, "cache size")?,
                u32_at(data, base + 40, rt, "cache characters")?,
                WebEditMode::from_code(u32_at(data, base + 44, rt, "edit mode")?)?,
                hash,
            )
        } else {
            ignored_fixed_tail.copy_from_slice(
                data.get(base + 32..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated XML fixed tail"))?,
            );
            if ignored_fixed_tail[12..16].iter().any(|byte| *byte != 0) {
                return Err(invalid(rt, "XML edit mode must be zero"));
            }
            (0, 0, 0, WebEditMode::Normal, [0; 16])
        };
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "Web/XML field count must match table range"));
        }
        let provider_name = if flags & 0x4000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "cSPName")?;
            offset = end;
            Some(v)
        } else {
            None
        };
        let entry_id = if flags & 0x0010_0000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "entryId")?;
            offset = end;
            if lt == 2 && v != id.value().to_string() {
                return Err(invalid(rt, "XML entryId does not match table id"));
            }
            Some(v)
        } else {
            None
        };
        let mut columns = Vec::with_capacity(count);
        let mut web_fields = Vec::with_capacity(count);
        let mut xml_fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = ListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            let web_type = u32_at(data, start + 4, rt, "lfdt")?;
            let xml_type = u32_at(data, start + 8, rt, "lfxidt")?;
            if (lt == 1 && xml_type != 0) || (lt == 2 && web_type != 0) {
                return Err(invalid(rt, "field data type does not match table source"));
            }
            let aggregation =
                ListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let agg_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert-row format length overflows"))?;
            if field_flags & 0x0000_0040 != 0
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != ListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != ListTotalAggregation::None
                || (field_flags & 1 != 0) != (flags & 2 != 0)
                || (lt == 1 && field_flags & 0x804 != 0)
                || (lt == 2 && field_flags & 8 != 0)
                || (rt == FEATURE11_RECORD_TYPE && field_flags & 0x480 != 0)
            {
                return Err(invalid(rt, "invalid source field condition flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            let (caption, after_caption) = if single {
                (source_name.clone(), after_source)
            } else {
                parse_string(data, after_source, rt, "strCaption")?
            };
            validate_column_name(&caption)?;
            let agg_end = after_caption
                .checked_add(agg_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..agg_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = agg_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(agg_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let n = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_| invalid(rt, "AutoFilter size overflows"))?;
                if n > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(rt, "AutoFilter size overflows"))?;
                let v = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                v
            } else {
                vec![0; 6]
            };
            let mapping = if field_flags & 4 != 0 {
                if u16_at(data, offset, rt, "iXmapMac")? != 1 {
                    return Err(invalid(rt, "XML mapped field must contain one map entry"));
                }
                let map_flags = u32_at(data, offset + 2, rt, "XMap flags")?;
                if map_flags & !6 != 0 || map_flags & 2 == 0 {
                    return Err(invalid(rt, "invalid XML map flags"));
                }
                let map_id = u32_at(data, offset + 6, rt, "XML map id")?;
                let (xpath, end) = parse_string(data, offset + 10, rt, "XPath")?;
                offset = end;
                Some(XmlColumnMapping::try_new(
                    map_id,
                    xpath,
                    map_flags & 4 != 0,
                )?)
            } else {
                None
            };
            let calculated_formula = if field_flags & 8 != 0 {
                Some(parse_formula(data, &mut offset, rt, "calculated formula")?)
            } else {
                None
            };
            let (total_formula, total_extra) = if field_flags & 0x80 != 0 {
                let tokens = parse_formula(data, &mut offset, rt, "total formula")?;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (v, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(v)
            } else {
                None
            };
            let web_kind = if lt == 1 {
                Some(WebColumnType::from_code(web_type)?)
            } else {
                None
            };
            let web_info = if let Some(kind) = web_kind {
                Some(parse_web_info(data, &mut offset, kind, rt)?)
            } else {
                None
            };
            if header == 0 && !single {
                let n = usize::try_from(u32_at(data, offset, rt, "cached header format size")?)
                    .map_err(|_| invalid(rt, "cached header size overflows"))?;
                let end = offset
                    .checked_add(4 + n)
                    .ok_or_else(|| invalid(rt, "cached header size overflows"))?;
                data.get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated cached header"))?;
                offset = if field_flags & 0x200 != 0 {
                    parse_string(data, end, rt, "cached header style")?.1
                } else {
                    end
                }
            } else if field_flags & 0x200 != 0 {
                return Err(invalid(rt, "cached header style lacks cached header"));
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
            if let (Some(kind), Some(info)) = (web_kind, web_info) {
                web_fields.push(WebTableField {
                    column_id: cid,
                    source_name,
                    data_type: kind,
                    info,
                    calculated_formula,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            } else {
                xml_fields.push(XmlTableField {
                    column_id: cid,
                    source_name,
                    data_type: XmlDataType::try_new(xml_type)?,
                    mapping,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            }
        }
        let source_metadata = if lt == 1 {
            let parse_ids = |data: &[u8], offset: &mut usize, label: &str| -> Result<Vec<u32>> {
                let count = usize::from(u16_at(data, *offset, rt, label)?);
                *offset += 2;
                let end = (*offset)
                    .checked_add(
                        count
                            .checked_mul(4)
                            .ok_or_else(|| invalid(rt, "source id count overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "source id count overflows"))?;
                let bytes = data
                    .get(*offset..end)
                    .ok_or_else(|| invalid(rt, format!("truncated {label}")))?;
                let ids = bytes
                    .chunks_exact(4)
                    .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
                    .collect();
                *offset = end;
                Ok(ids)
            };
            let deleted_row_ids = if flags & 0x20 != 0 {
                parse_ids(data, &mut offset, "deleted row ids")?
            } else {
                Vec::new()
            };
            let changed_row_ids = if flags & 0x8000 != 0 {
                parse_ids(data, &mut offset, "changed row ids")?
            } else {
                Vec::new()
            };
            let invalid_cells = if flags & 0x0020_0000 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "invalid cell count")?);
                offset += 2;
                let mut out = Vec::with_capacity(n);
                for _ in 0..n {
                    let row = u32_at(data, offset, rt, "invalid cell row")?;
                    let column =
                        ListColumnId::try_new(u32_at(data, offset + 4, rt, "invalid cell field")?)?;
                    offset += 8;
                    out.push(WebInvalidCell::new(row, column))
                }
                out
            } else {
                Vec::new()
            };
            ListObjectSourceMetadata::Web(WebTableMetadata {
                version,
                build_number: build,
                fields: web_fields,
                edit_mode,
                cache_position,
                cache_size,
                cache_characters,
                hash_parameters: hash,
                provider_name,
                entry_id,
                deleted_row_ids,
                changed_row_ids,
                invalid_cells,
                needs_commit: flags & 0x100 != 0,
                compressed_cache: flags & 0x2000 != 0,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
            })
        } else {
            ListObjectSourceMetadata::Xml(XmlTableMetadata {
                version,
                build_number: build,
                fields: xml_fields,
                entry_id,
                single_cell: single,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
                ignored_fixed_tail,
            })
        };
        let has_feature12_field = columns
            .iter()
            .any(|column| column.total_formula.is_some() || column.total_string.is_some());
        if rt == FEATURE12_RECORD_TYPE && header != 0 && !has_feature12_field {
            return Err(invalid(
                rt,
                "Feature12 Web/XML source lacks a Feature12-only property",
            ));
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing Web/XML feature data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
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
            source_metadata: Some(source_metadata),
        })
    }
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
            autofilter: flags & 2 != 0,
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
