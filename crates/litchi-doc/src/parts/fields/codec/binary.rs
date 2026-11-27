//! PLCF field-range materialization.

use super::super::model::{Field, FieldText, corrupted};
use crate::package::Result;
impl FieldText {
    pub(crate) fn from_field<F>(field: &Field, mut text_at_range: F) -> Result<Self>
    where
        F: FnMut(u32, u32) -> Result<String>,
    {
        let instruction_start = field
            .start_cp
            .checked_add(1)
            .ok_or_else(|| corrupted("field instruction start overflows"))?;
        let instruction_end = field.separator_cp.unwrap_or(field.end_cp);
        if instruction_start > instruction_end {
            return Err(corrupted(
                "field instruction range has its start after its end",
            ));
        }
        let instruction = text_at_range(instruction_start, instruction_end)?;
        let result = match field.separator_cp {
            Some(separator) => {
                let start = separator
                    .checked_add(1)
                    .ok_or_else(|| corrupted("field result start overflows"))?;
                if start > field.end_cp {
                    return Err(corrupted("field result range has its start after its end"));
                }
                Some(text_at_range(start, field.end_cp)?)
            },
            None => None,
        };

        Ok(Self {
            field: field.clone(),
            instruction,
            result,
        })
    }
}
