use litchi_odraw::shape::Shape;
use litchi_odraw::{Record, RecordKind};

use crate::package::{Error, Result};

pub(super) const CLIENT_DATA_RAW_KIND: u16 = 0xf011;
pub(super) const CLIENT_TEXTBOX_RAW_KIND: u16 = 0xf00d;
pub(super) const MAX_HOST_RECORDS: u32 = 1_000_000;

pub(super) fn host_record<'data>(
    shape: &Shape<'data>,
    kind: RecordKind,
) -> Result<Option<Record<'data>>> {
    match kind {
        RecordKind::ClientData => Ok(shape.client_data().cloned()),
        RecordKind::ClientTextbox => Ok(shape.textbox().cloned()),
        _ => shape.container().find(kind).map_err(Into::into),
    }
}

pub(super) fn validate_host_record(record: &Record<'_>, raw_kind: u16, name: &str) -> Result<()> {
    if record.version() != 0x0f
        || record.instance() != 0
        || record.raw_kind() != raw_kind
        || usize::try_from(record.len()).ok() != Some(record.data().len())
    {
        return Err(corrupted(&format!(
            "Invalid OfficeArt {name} record header"
        )));
    }
    Ok(())
}

pub(super) fn advance(offset: usize, consumed: usize, context: &str) -> Result<usize> {
    if consumed == 0 {
        return Err(corrupted(&format!("Zero-length PPT record in {context}")));
    }
    offset
        .checked_add(consumed)
        .ok_or_else(|| corrupted(&format!("{context} offset overflow")))
}

pub(super) fn visit_host_record(records: &mut u32) -> Result<()> {
    *records = records
        .checked_add(1)
        .ok_or_else(|| corrupted("Host payload record count overflow"))?;
    if *records > MAX_HOST_RECORDS {
        return Err(corrupted("Host payload exceeds the PPT record limit"));
    }
    Ok(())
}

pub(super) fn corrupted(message: &str) -> Error {
    Error::Corrupted(message.to_owned())
}

#[derive(Debug, Clone, Copy)]
pub(super) struct RawRecord<'data> {
    pub(super) kind: u16,
    pub(super) data: &'data [u8],
}

/// Strict, allocation-free iterator for records embedded in host payloads.
pub(super) struct RawRecords<'data> {
    data: &'data [u8],
    offset: usize,
    seen: u32,
    failed: bool,
}

impl<'data> RawRecords<'data> {
    pub(super) fn new(data: &'data [u8]) -> Self {
        Self {
            data,
            offset: 0,
            seen: 0,
            failed: false,
        }
    }
}

impl<'data> Iterator for RawRecords<'data> {
    type Item = Result<RawRecord<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed || self.offset == self.data.len() {
            return None;
        }
        self.seen = match self.seen.checked_add(1) {
            Some(seen) if seen <= MAX_HOST_RECORDS => seen,
            _ => {
                self.failed = true;
                return Some(Err(corrupted("Host payload exceeds the PPT record limit")));
            },
        };
        let Some(header_end) = self.offset.checked_add(8) else {
            self.failed = true;
            return Some(Err(corrupted("PPT record header offset overflow")));
        };
        let Some(header) = self.data.get(self.offset..header_end) else {
            self.failed = true;
            return Some(Err(corrupted(
                "Host payload ends with a truncated PPT record header",
            )));
        };
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let len = match usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ])) {
            Ok(len) => len,
            Err(_) => {
                self.failed = true;
                return Some(Err(corrupted(
                    "PPT record length cannot be represented on this platform",
                )));
            },
        };
        let Some(end) = header_end.checked_add(len) else {
            self.failed = true;
            return Some(Err(corrupted("PPT record payload offset overflow")));
        };
        let Some(data) = self.data.get(header_end..end) else {
            self.failed = true;
            return Some(Err(corrupted(
                "Host payload contains a truncated PPT record",
            )));
        };
        self.offset = end;
        Some(Ok(RawRecord { kind, data }))
    }
}
