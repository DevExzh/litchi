//! Worksheet-level BIFF record sequence assembly for list objects.

use super::codec::{
    PendingFeature, append_frt, record, u16_at, u32_at, validate_frt, validate_frt_any,
};
use super::model::*;
use super::{
    AUTO_FILTER12_RECORD_TYPE, CONTINUE_FRT11_RECORD_TYPE, FEAT_HDR11_RECORD_TYPE,
    FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, ISF_LIST, LIST12_RECORD_TYPE, MAX_FEATURE_BYTES,
    MAX_PAYLOAD, invalid,
};
use crate::Result;
use crate::autofilter12::parse_table_autofilter12;
use std::collections::HashSet;

struct PendingFuture {
    payload: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    after_list12_count: usize,
}
pub(crate) struct ListObjectCollector {
    header: Option<u32>,
    pending: Option<PendingFeature>,
    current: Option<ListObject>,
    pending_future: Option<PendingFuture>,
    kinds: HashSet<u16>,
    list12_count: usize,
    sort_continuations: usize,
    tables: Vec<ListObject>,
    ended: bool,
}
impl ListObjectCollector {
    pub(crate) fn new() -> Self {
        Self {
            header: None,
            pending: None,
            current: None,
            pending_future: None,
            kinds: HashSet::new(),
            list12_count: 0,
            sort_continuations: 0,
            tables: Vec::new(),
            ended: false,
        }
    }
    pub(crate) fn feed_record(&mut self, rt: u16, data: &[u8]) -> Result<()> {
        if self.header.is_none()
            && matches!(
                rt,
                AUTO_FILTER12_RECORD_TYPE
                    | crate::sort_data::SORT_DATA_RECORD_TYPE
                    | crate::sort_data::CONTINUE_FRT12_RECORD_TYPE
            )
        {
            return Ok(());
        }
        let family = matches!(
            rt,
            FEAT_HDR11_RECORD_TYPE
                | FEATURE11_RECORD_TYPE
                | FEATURE12_RECORD_TYPE
                | CONTINUE_FRT11_RECORD_TYPE
                | LIST12_RECORD_TYPE
                | AUTO_FILTER12_RECORD_TYPE
                | crate::sort_data::SORT_DATA_RECORD_TYPE
                | crate::sort_data::CONTINUE_FRT12_RECORD_TYPE
        );
        if !family {
            if self.header.is_some() {
                if self.sort_continuations != 0 {
                    return Err(invalid(rt, "incomplete table SortData continuation chain"));
                }
                self.materialize()?;
                self.finish_future()?;
                self.ended = true;
            }
            return Ok(());
        }
        if rt == FEAT_HDR11_RECORD_TYPE {
            // FeatHdr11 is shared by all Feat11 feature families. Leave non-list
            // discriminators to their dedicated collectors.
            if data.len() < 14 || u16_at(data, 12, rt, "isf")? != ISF_LIST {
                if self.header.is_some() {
                    self.ended = true;
                }
                return Ok(());
            }
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            if self.header.is_some() || data.len() != 29 {
                return Err(invalid(rt, "duplicate or malformed FeatHdr11"));
            }
            validate_frt(data, rt, false)?;
            if u16_at(data, 12, rt, "isf")? != ISF_LIST
                || data[14] != 1
                || u32_at(data, 15, rt, "reserved2")? != u32::MAX
                || u32_at(data, 19, rt, "reserved3")? != u32::MAX
                || u16_at(data, 27, rt, "reserved4")? != 0
            {
                return Err(invalid(rt, "invalid FeatHdr11 fields"));
            }
            self.header = Some(u32_at(data, 23, rt, "idListNext")?);
        } else if matches!(rt, FEATURE11_RECORD_TYPE | FEATURE12_RECORD_TYPE) {
            if data.len() < 14 || u16_at(data, 12, rt, "isf")? != ISF_LIST {
                if self.header.is_some() {
                    self.ended = true;
                }
                return Ok(());
            }
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            if self.header.is_none() {
                return Err(invalid(rt, "table feature without FeatHdr11"));
            }
            if self.sort_continuations != 0 {
                return Err(invalid(
                    rt,
                    "table feature interrupts SortData continuation chain",
                ));
            }
            self.materialize()?;
            self.finish_future()?;
            self.flush()?;
            if data.len() > MAX_PAYLOAD {
                return Err(invalid(rt, "base table feature exceeds BIFF record limit"));
            }
            self.pending = Some(PendingFeature {
                record_type: rt,
                base: data.to_vec(),
                continuations: Vec::new(),
                combined: data.to_vec(),
            });
            self.kinds.clear();
            self.list12_count = 0;
        } else if rt == CONTINUE_FRT11_RECORD_TYPE {
            let pending = self
                .pending
                .as_mut()
                .ok_or_else(|| invalid(rt, "orphan ContinueFrt11"))?;
            if pending.base.len() != MAX_PAYLOAD
                || pending
                    .continuations
                    .last()
                    .is_some_and(|v| v.len() != MAX_PAYLOAD)
            {
                return Err(invalid(
                    rt,
                    "ContinueFrt11 follows a non-full feature fragment",
                ));
            }
            if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                return Err(invalid(rt, "invalid ContinueFrt11 length"));
            }
            validate_frt(data, rt, false)?;
            pending.combined.extend_from_slice(&data[12..]);
            if pending.combined.len() > MAX_FEATURE_BYTES {
                return Err(invalid(
                    rt,
                    "table feature continuation chain exceeds resource bound",
                ));
            }
            pending.continuations.push(data.to_vec());
        } else if rt == LIST12_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            if self.ended {
                return Err(invalid(rt, "noncontiguous list FEAT11 family"));
            }
            let kind = self
                .current
                .as_mut()
                .ok_or_else(|| invalid(rt, "List12 without Feature11"))?
                .apply_list12(data)?;
            if !self.kinds.insert(kind) {
                return Err(invalid(rt, "duplicate List12 type"));
            }
            self.list12_count += 1;
        } else if rt == AUTO_FILTER12_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            if self.current.is_none() || self.list12_count == 0 {
                return Err(invalid(
                    rt,
                    "AutoFilter12 is not attached after a table List12",
                ));
            }
            if self.current.as_ref().is_some_and(|table| {
                table.autofilter12_criteria.is_some()
                    || table
                        .opaque_future_records
                        .iter()
                        .any(|future| future.record_type == AUTO_FILTER12_RECORD_TYPE)
            }) {
                return Err(invalid(rt, "duplicate AutoFilter12 for table"));
            }
            if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                return Err(invalid(rt, "invalid AutoFilter12 length"));
            }
            validate_frt_any(data, rt)?;
            self.pending_future = Some(PendingFuture {
                payload: data.to_vec(),
                continuations: Vec::new(),
                after_list12_count: self.list12_count,
            });
        } else if rt == crate::sort_data::SORT_DATA_RECORD_TYPE {
            self.materialize()?;
            self.finish_future()?;
            let table = self
                .current
                .as_ref()
                .ok_or_else(|| invalid(rt, "SortData without table feature"))?;
            if data.len() != 38
                || ((u16_at(data, 12, rt, "sort flags")? >> 3) & 0x7) != 1
                || u32_at(data, 34, rt, "sort parent id")? != table.id.value()
            {
                return Err(invalid(
                    rt,
                    "table SortData parent does not match Feature11/12",
                ));
            }
            self.sort_continuations = u32_at(data, 30, rt, "sort condition count")? as usize;
        } else if let Some(future) = self.pending_future.as_mut() {
            if future.payload.len() < 60 {
                return Err(invalid(
                    rt,
                    "ContinueFrt12 follows a truncated AutoFilter12 base",
                ));
            }
            if !(12..=MAX_PAYLOAD).contains(&data.len()) {
                return Err(invalid(rt, "invalid ContinueFrt12 length"));
            }
            validate_frt_any(data, rt)?;
            let total = future.payload.len()
                + future.continuations.iter().map(Vec::len).sum::<usize>()
                + data.len();
            if total > MAX_FEATURE_BYTES {
                return Err(invalid(
                    rt,
                    "AutoFilter12 continuation chain exceeds resource bound",
                ));
            }
            future.continuations.push(data.to_vec());
        } else if self.sort_continuations != 0 {
            self.sort_continuations -= 1;
        } else {
            return Err(invalid(rt, "orphan ContinueFrt12 in table feature family"));
        }
        Ok(())
    }
    fn materialize(&mut self) -> Result<()> {
        let Some(pending) = self.pending.take() else {
            return Ok(());
        };
        let source_type = u32_at(&pending.combined, 35, pending.record_type, "lt")?;
        self.current = Some(
            if pending.record_type == FEATURE12_RECORD_TYPE && source_type == 3 {
                ListObject::parse_external_feature12(pending)?
            } else if pending.record_type == FEATURE12_RECORD_TYPE && source_type > 3 {
                ListObject::parse_opaque_feature12(pending)?
            } else {
                ListObject::parse_feature(&pending.combined, pending.record_type)?
            },
        );
        Ok(())
    }
    fn finish_future(&mut self) -> Result<()> {
        if let Some(future) = self.pending_future.take() {
            let table = self
                .current
                .as_mut()
                .ok_or_else(|| invalid(AUTO_FILTER12_RECORD_TYPE, "detached AutoFilter12"))?;
            if let Some(filter) = parse_table_autofilter12(
                &future.payload,
                &future.continuations,
                table.range,
                table.id,
            )? {
                table.autofilter12_criteria = Some(filter);
            } else {
                table
                    .opaque_future_records
                    .push(OpaqueListObjectFutureRecord {
                        record_type: AUTO_FILTER12_RECORD_TYPE,
                        payload: future.payload,
                        continuation_payloads: future.continuations,
                        after_list12_count: future.after_list12_count,
                    });
            }
        }
        Ok(())
    }
    fn flush(&mut self) -> Result<()> {
        if let Some(table) = self.current.take() {
            table.validate()?;
            self.tables.push(table);
        }
        Ok(())
    }
    pub(crate) fn finish(mut self) -> Result<Vec<ListObject>> {
        if self.sort_continuations != 0 {
            return Err(invalid(
                crate::sort_data::SORT_DATA_RECORD_TYPE,
                "incomplete table SortData continuation chain",
            ));
        }
        self.materialize()?;
        self.finish_future()?;
        self.flush()?;
        let mut ids = HashSet::new();
        let mut names = HashSet::new();
        for table in &self.tables {
            if !ids.insert(table.id) || !names.insert(table.name.to_lowercase()) {
                return Err(invalid(FEATURE11_RECORD_TYPE, "duplicate table id or name"));
            }
        }
        Ok(self.tables)
    }
}
pub(crate) fn feature_header_record(tables: &[ListObject]) -> Result<Vec<u8>> {
    let next = tables
        .iter()
        .map(|t| t.id.value())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| invalid(FEAT_HDR11_RECORD_TYPE, "next table id overflows"))?;
    let mut p = Vec::new();
    append_frt(&mut p, FEAT_HDR11_RECORD_TYPE, None);
    p.extend_from_slice(&ISF_LIST.to_le_bytes());
    p.push(1);
    p.extend_from_slice(&u32::MAX.to_le_bytes());
    p.extend_from_slice(&u32::MAX.to_le_bytes());
    p.extend_from_slice(&next.to_le_bytes());
    p.extend_from_slice(&0u16.to_le_bytes());
    record(FEAT_HDR11_RECORD_TYPE, p)
}
