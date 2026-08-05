//! Package-facing List12 and following-record emission.

use super::super::model::ListObject;
use super::super::{AUTO_FILTER12_RECORD_TYPE, LIST12_RECORD_TYPE, invalid};
use super::binary::{append_frt, append_string, record};
use crate::Result;
use crate::autofilter12::write_table_autofilter12;

impl ListObject {
    pub(crate) fn to_list12_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let mut block = Vec::new();
        append_frt(&mut block, LIST12_RECORD_TYPE, None);
        block.extend_from_slice(&0u16.to_le_bytes());
        block.extend_from_slice(&self.id.value().to_le_bytes());
        for v in [0i32, -1, 0, -1, 0, -1, 0, 0, 0] {
            block.extend_from_slice(&v.to_le_bytes());
        }
        let style = self.style.as_ref().unwrap();
        let mut styled = Vec::new();
        append_frt(&mut styled, LIST12_RECORD_TYPE, None);
        styled.extend_from_slice(&1u16.to_le_bytes());
        styled.extend_from_slice(&self.id.value().to_le_bytes());
        let bits = u16::from(style.first)
            | u16::from(style.last) << 1
            | u16::from(style.row_stripes) << 2
            | u16::from(style.column_stripes) << 3
            | u16::from(style.default_style) << 6;
        styled.extend_from_slice(&bits.to_le_bytes());
        append_string(&mut styled, &style.name);
        let mut display = Vec::new();
        append_frt(&mut display, LIST12_RECORD_TYPE, None);
        display.extend_from_slice(&2u16.to_le_bytes());
        display.extend_from_slice(&self.id.value().to_le_bytes());
        append_string(&mut display, &self.name);
        append_string(&mut display, &self.comment);
        Ok(vec![
            record(LIST12_RECORD_TYPE, block)?,
            record(LIST12_RECORD_TYPE, styled)?,
            record(LIST12_RECORD_TYPE, display)?,
        ])
    }
    pub(crate) fn to_following_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let list12 = self.to_list12_record_bytes()?;
        let mut output = Vec::new();
        for (index, item) in list12.into_iter().enumerate() {
            output.push(item);
            if index == 0
                && let Some(filter) = &self.autofilter12_criteria
            {
                output.extend(write_table_autofilter12(filter, self.range, self.id)?);
            }
            for future in self
                .opaque_future_records
                .iter()
                .filter(|v| v.after_list12_count == index + 1)
            {
                output.push(record(future.record_type, future.payload.clone())?);
                for payload in &future.continuation_payloads {
                    output.push(record(
                        crate::sort_data::CONTINUE_FRT12_RECORD_TYPE,
                        payload.clone(),
                    )?);
                }
            }
        }
        if self
            .opaque_future_records
            .iter()
            .any(|v| v.after_list12_count == 0 || v.after_list12_count > 3)
        {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "opaque table future-record insertion point is invalid",
            ));
        }
        Ok(output)
    }
}
