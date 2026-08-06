//! Ordinary Feature11/Feature12 table encoding.

use crate::Result;
use crate::list_object::codec::binary::{append_frt, append_range, append_string, record};
use crate::list_object::model::*;
use crate::list_object::{
    CONTINUE_FRT11_RECORD_TYPE, FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST,
    MAX_CONTINUE_RGB, MAX_FEATURE_BYTES, MAX_PAYLOAD, invalid,
};

impl ListObject {
    pub(super) fn to_ordinary_feature_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let rt = match self.feature_version {
            ListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            ListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&0u32.to_le_bytes());
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
        feature.extend_from_slice(&[0; 4]);
        let table_flags = self.table_flags.with_auto_filter(self.autofilter);
        feature.extend_from_slice(&table_flags.raw().to_le_bytes());
        feature.extend_from_slice(&[0; 32]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(self.columns.len() as u16).to_le_bytes());
        if table_flags.loads_entry_id() {
            append_string(&mut feature, &self.id.value().to_string());
        }
        for column in &self.columns {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let cflags = u32::from(table_flags.auto_filter())
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(column.total_string.is_some()) << 10);
            feature.extend_from_slice(&cflags.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, &column.name);
            append_string(&mut feature, &column.name);
            if self.autofilter {
                feature.extend_from_slice(&0u32.to_le_bytes());
                feature.extend_from_slice(&(column.id.value() as u16).to_le_bytes());
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
                feature.extend_from_slice(tokens);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first_len].to_vec())?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
}
