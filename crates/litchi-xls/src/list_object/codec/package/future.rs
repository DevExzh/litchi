//! Ordered emission of typed and opaque records following List12.

use super::super::super::model::ListObject;
use super::super::binary::record;
use crate::Result;
use crate::autofilter12::write_table_autofilter12;

impl ListObject {
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
                .filter(|value| value.after_list12_count == index + 1)
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
        super::validation::validate_future_insertion(self)?;
        Ok(output)
    }
}
