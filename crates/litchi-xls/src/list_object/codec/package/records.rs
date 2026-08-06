//! List12 wire-record emission.

use super::super::super::LIST12_RECORD_TYPE;
use super::super::super::model::ListObject;
use super::super::binary::{append_frt, append_string, record};
use crate::Result;

impl ListObject {
    pub(crate) fn to_list12_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let mut block = Vec::new();
        append_frt(&mut block, LIST12_RECORD_TYPE, None);
        block.extend_from_slice(&0u16.to_le_bytes());
        block.extend_from_slice(&self.id.value().to_le_bytes());
        for value in [0i32, -1, 0, -1, 0, -1, 0, 0, 0] {
            block.extend_from_slice(&value.to_le_bytes());
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
}
