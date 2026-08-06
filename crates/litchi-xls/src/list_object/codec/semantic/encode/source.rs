//! Web and XML source-table encoding for Feature11/Feature12.

use crate::Result;
use crate::list_object::codec::binary::{
    append_formula, append_frt, append_range, append_string, append_web_info, record,
};
use crate::list_object::model::*;
use crate::list_object::{
    CONTINUE_FRT11_RECORD_TYPE, FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST,
    MAX_CONTINUE_RGB, MAX_FEATURE_BYTES, MAX_PAYLOAD, invalid,
};

impl ListObject {
    pub(super) fn to_source_feature_record_bytes(
        &self,
        source: &ListObjectSourceMetadata,
    ) -> Result<Vec<Vec<u8>>> {
        let rt = match self.feature_version {
            ListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            ListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let (lt, version, build, single, fields_len): (u32, _, _, _, _) = match source {
            ListObjectSourceMetadata::Web(v) => {
                (1u32, v.version, v.build_number, false, v.fields.len())
            },
            ListObjectSourceMetadata::Xml(v) => (
                2u32,
                v.version,
                v.build_number,
                v.single_cell,
                v.fields.len(),
            ),
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&lt.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|c| c.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&build.to_le_bytes());
        let ignored_fixed_word = match source {
            ListObjectSourceMetadata::Web(v) => v.ignored_fixed_word,
            ListObjectSourceMetadata::Xml(v) => v.ignored_fixed_word,
        };
        feature.extend_from_slice(&ignored_fixed_word.to_le_bytes());
        let mut flags = (version.code() << 16)
            | (u32::from(self.autofilter) * 0x806)
            | (u32::from(self.has_totals) * 0x40)
            | (u32::from(single) << 9)
            | 0x0040_0000;
        match source {
            ListObjectSourceMetadata::Web(v) => {
                flags |= u32::from(!v.deleted_row_ids.is_empty()) << 5
                    | u32::from(v.needs_commit) << 8
                    | u32::from(v.compressed_cache) << 13
                    | u32::from(v.provider_name.is_some()) << 14
                    | u32::from(!v.changed_row_ids.is_empty()) << 15
                    | u32::from(v.entry_id.is_some()) << 20
                    | u32::from(!v.invalid_cells.is_empty()) << 21
                    | v.ignored_flags
            },
            ListObjectSourceMetadata::Xml(v) => {
                flags |= u32::from(v.entry_id.is_some()) << 20 | v.ignored_flags
            },
        };
        feature.extend_from_slice(&flags.to_le_bytes());
        match source {
            ListObjectSourceMetadata::Web(v) => {
                feature.extend_from_slice(&v.cache_position.to_le_bytes());
                feature.extend_from_slice(&v.cache_size.to_le_bytes());
                feature.extend_from_slice(&v.cache_characters.to_le_bytes());
                feature.extend_from_slice(&v.edit_mode.code().to_le_bytes());
                feature.extend_from_slice(&v.hash_parameters)
            },
            ListObjectSourceMetadata::Xml(v) => feature.extend_from_slice(&v.ignored_fixed_tail),
        };
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(fields_len as u16).to_le_bytes());
        match source {
            ListObjectSourceMetadata::Web(v) => {
                if let Some(name) = &v.provider_name {
                    append_string(&mut feature, name)
                }
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
            ListObjectSourceMetadata::Xml(v) => {
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
        }
        for (index, column) in self.columns.iter().enumerate() {
            let (web, xml) = match source {
                ListObjectSourceMetadata::Web(v) => (Some(&v.fields[index]), None),
                ListObjectSourceMetadata::Xml(v) => (None, Some(&v.fields[index])),
            };
            let (
                source_name,
                web_type,
                xml_type,
                mapped,
                calc,
                auto_filter,
                agg_fmt,
                insert_fmt,
                total_extra,
                ignored_flags,
            ) = if let Some(v) = web {
                (
                    &v.source_name,
                    v.data_type.code(),
                    0,
                    false,
                    v.calculated_formula.as_deref(),
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            } else {
                let v = xml.unwrap();
                (
                    &v.source_name,
                    0,
                    v.data_type.value(),
                    v.mapping.is_some(),
                    None,
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            };
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&web_type.to_le_bytes());
            feature.extend_from_slice(&xml_type.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&(agg_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let ff = u32::from(self.autofilter)
                | (u32::from(mapped) << 2)
                | (u32::from(calc.is_some()) << 3)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(!total_extra.is_empty()) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | ignored_flags;
            feature.extend_from_slice(&ff.to_le_bytes());
            feature.extend_from_slice(&(insert_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, source_name);
            if !single {
                append_string(&mut feature, &column.name)
            }
            feature.extend_from_slice(agg_fmt);
            feature.extend_from_slice(insert_fmt);
            if self.autofilter {
                feature.extend_from_slice(auto_filter)
            }
            if let Some(v) = xml.and_then(|v| v.mapping.as_ref()) {
                feature.extend_from_slice(&1u16.to_le_bytes());
                feature.extend_from_slice(&(2u32 | u32::from(v.can_be_single) << 2).to_le_bytes());
                feature.extend_from_slice(&v.map_id.get().to_le_bytes());
                append_string(&mut feature, v.xpath.as_str())
            }
            if let Some(tokens) = calc {
                append_formula(&mut feature, tokens)?
            }
            if let Some(tokens) = &column.total_formula {
                append_formula(&mut feature, tokens)?;
                feature.extend_from_slice(total_extra)
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value)
            }
            if let Some(v) = web {
                append_web_info(&mut feature, &v.info)?
            }
        }
        if let ListObjectSourceMetadata::Web(v) = source {
            if !v.deleted_row_ids.is_empty() {
                feature.extend_from_slice(&(v.deleted_row_ids.len() as u16).to_le_bytes());
                for id in &v.deleted_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.changed_row_ids.is_empty() {
                feature.extend_from_slice(&(v.changed_row_ids.len() as u16).to_le_bytes());
                for id in &v.changed_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.invalid_cells.is_empty() {
                feature.extend_from_slice(&(v.invalid_cells.len() as u16).to_le_bytes());
                for cell in &v.invalid_cells {
                    feature.extend_from_slice(&cell.row_id.to_le_bytes());
                    feature.extend_from_slice(&cell.column_id.value().to_le_bytes())
                }
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&(feature.len() as u32).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "Web/XML table feature exceeds aggregate resource bound",
            ));
        }
        let first = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first].to_vec())?];
        for chunk in payload[first..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?)
        }
        Ok(records)
    }
}
