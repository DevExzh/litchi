//! List12 style and display metadata updates.

use super::super::super::model::{ListObject, ListObjectStyleOptions};
use super::super::super::{LIST12_RECORD_TYPE, invalid};
use super::super::binary::{parse_string, u16_at, u32_at, validate_frt};
use crate::Result;

impl ListObject {
    pub(in crate::list_object) fn apply_list12(&mut self, data: &[u8]) -> Result<u16> {
        if data.len() < 18 {
            return Err(invalid(LIST12_RECORD_TYPE, "truncated List12"));
        }
        validate_frt(data, LIST12_RECORD_TYPE, false)?;
        let kind = u16_at(data, 12, LIST12_RECORD_TYPE, "lsd")?;
        if u32_at(data, 14, LIST12_RECORD_TYPE, "idList")? != self.id.value() {
            return Err(invalid(LIST12_RECORD_TYPE, "List12 id mismatch"));
        }
        match kind {
            0 => {
                if data.len() < 54 {
                    return Err(invalid(LIST12_RECORD_TYPE, "truncated block-level List12"));
                }
            },
            1 => {
                let bits = u16_at(data, 18, LIST12_RECORD_TYPE, "style flags")?;
                let (name, end) = parse_string(data, 20, LIST12_RECORD_TYPE, "style name")?;
                if end != data.len() {
                    return Err(invalid(LIST12_RECORD_TYPE, "trailing style List12 data"));
                }
                self.style = Some(ListObjectStyleOptions {
                    name,
                    first: bits & 1 != 0,
                    last: bits & 2 != 0,
                    row_stripes: bits & 4 != 0,
                    column_stripes: bits & 8 != 0,
                    default_style: bits & 0x40 != 0,
                });
            },
            2 => {
                let (name, next) = parse_string(data, 18, LIST12_RECORD_TYPE, "display name")?;
                let (comment, end) = parse_string(data, next, LIST12_RECORD_TYPE, "comment")?;
                if end != data.len() || (!name.is_empty() && name != self.name) {
                    return Err(invalid(LIST12_RECORD_TYPE, "inconsistent display List12"));
                }
                self.comment = comment;
            },
            _ => return Err(invalid(LIST12_RECORD_TYPE, "reserved List12 type")),
        }
        Ok(kind)
    }
}
