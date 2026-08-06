//! Bounded retention and materialization of table future records.

use super::super::codec::validate_frt_any;
use super::super::model::{ListObject, OpaqueListObjectFutureRecord};
use super::super::{AUTO_FILTER12_RECORD_TYPE, MAX_FEATURE_BYTES, MAX_PAYLOAD, invalid};
use crate::Result;
use crate::autofilter12::parse_table_autofilter12;

pub(super) struct PendingFuture {
    payload: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    after_list12_count: usize,
}

impl PendingFuture {
    pub(super) fn new(payload: Vec<u8>, after_list12_count: usize) -> Self {
        Self {
            payload,
            continuations: Vec::new(),
            after_list12_count,
        }
    }

    pub(super) fn push_continuation(&mut self, rt: u16, data: &[u8]) -> Result<()> {
        if self.payload.len() < 60 {
            return Err(invalid(
                rt,
                "ContinueFrt12 follows a truncated AutoFilter12 base",
            ));
        }
        if !(12..=MAX_PAYLOAD).contains(&data.len()) {
            return Err(invalid(rt, "invalid ContinueFrt12 length"));
        }
        validate_frt_any(data, rt)?;
        let total = self.payload.len()
            + self.continuations.iter().map(Vec::len).sum::<usize>()
            + data.len();
        if total > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "AutoFilter12 continuation chain exceeds resource bound",
            ));
        }
        self.continuations.push(data.to_vec());
        Ok(())
    }

    pub(super) fn finish(self, table: &mut ListObject) -> Result<()> {
        if let Some(filter) =
            parse_table_autofilter12(&self.payload, &self.continuations, table.range, table.id)?
        {
            table.autofilter12_criteria = Some(filter);
        } else {
            table
                .opaque_future_records
                .push(OpaqueListObjectFutureRecord {
                    record_type: AUTO_FILTER12_RECORD_TYPE,
                    payload: self.payload,
                    continuation_payloads: self.continuations,
                    after_list12_count: self.after_list12_count,
                });
        }
        Ok(())
    }
}
