//! Feature12 LTEXTERNALDATA table encoding.

use crate::Result;
use crate::list_object::codec::binary::{append_frt, append_range, append_string, record};
use crate::list_object::model::{ExternalTableMetadata, ListObject};
use crate::list_object::{
    CONTINUE_FRT11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST, MAX_CONTINUE_RGB,
    MAX_FEATURE_BYTES, MAX_PAYLOAD, invalid,
};

impl ListObject {
    pub(super) fn to_external_feature_record_bytes(
        &self,
        metadata: &ExternalTableMetadata,
    ) -> Result<Vec<Vec<u8>>> {
        let mut feature = Vec::new();
        feature.extend_from_slice(&3u32.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|column| column.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE12_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&metadata.build_number.to_le_bytes());
        feature.extend_from_slice(&0u16.to_le_bytes());
        let mut flags = metadata.version.code() << 16 | 0x0010_0000;
        if self.autofilter {
            flags |= 0x0000_0806;
        }
        if self.has_totals {
            flags |= 0x40;
        }
        feature.extend_from_slice(&flags.to_le_bytes());
        feature.extend_from_slice(&[0; 12]);
        feature.extend_from_slice(&0u32.to_le_bytes());
        feature.extend_from_slice(&[0; 16]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(
            &crate::utils::truncate_usize_to_u16(self.columns.len()).to_le_bytes(),
        );
        append_string(&mut feature, &self.id.value().to_string());
        for (column, field) in self.columns.iter().zip(&metadata.fields) {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(
                &crate::utils::truncate_usize_to_u32(field.aggregate_format.len()).to_le_bytes(),
            );
            feature.extend_from_slice(&field.aggregate_style.to_le_bytes());
            let mut field_flags = u32::from(self.autofilter)
                | (u32::from(field.filter_hidden) << 1)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(field.total_array_formula) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | (u32::from(field.auto_create_calculated_column) << 11);
            if !self.has_header && field.header_cache.style_name().is_some() {
                field_flags |= 0x200;
            }
            feature.extend_from_slice(&field_flags.to_le_bytes());
            feature.extend_from_slice(
                &crate::utils::truncate_usize_to_u32(field.insert_row_format.len()).to_le_bytes(),
            );
            feature.extend_from_slice(&field.insert_row_style.to_le_bytes());
            append_string(&mut feature, &field.source_name);
            append_string(&mut feature, &column.name);
            feature.extend_from_slice(&field.aggregate_format);
            feature.extend_from_slice(&field.insert_row_format);
            if self.autofilter {
                feature.extend_from_slice(&field.auto_filter);
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(
                    &crate::utils::truncate_usize_to_u16(tokens.len()).to_le_bytes(),
                );
                feature.extend_from_slice(tokens);
                feature.extend_from_slice(&field.formula_extra);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
            feature.extend_from_slice(&field.query_field_id.to_le_bytes());
            if !self.has_header {
                feature.extend_from_slice(field.header_cache.as_bytes());
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, FEATURE12_RECORD_TYPE, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload
            .extend_from_slice(&crate::utils::truncate_usize_to_u32(feature.len()).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(
            FEATURE12_RECORD_TYPE,
            payload[..first_len].to_vec(),
        )?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
}
