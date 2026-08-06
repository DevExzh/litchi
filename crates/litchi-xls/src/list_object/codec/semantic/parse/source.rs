//! Web and XML source-table parsing for Feature11/Feature12.

use crate::Result;
use crate::list_object::codec::binary::{
    parse_formula, parse_list_formula_extra_end, parse_range, parse_string, parse_web_info, u16_at,
    u32_at,
};
use crate::list_object::model::*;
use crate::list_object::{FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, invalid};

impl ListObject {
    pub(super) fn parse_source_feature(data: &[u8], rt: u16, lt: u32) -> Result<Self> {
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
        let table_flags = crate::list_object::TableFlags::from_raw(flags);
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
            source_metadata: Some(source_metadata),
        })
    }
}
