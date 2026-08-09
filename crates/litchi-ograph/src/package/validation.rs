use std::io::Cursor;

use litchi_biff::{Kind, RecordRef, Records};
use litchi_cfb::consts::STGTY_STREAM;
use litchi_cfb::{DirectoryEntry, OleFile};

use crate::limits::as_u64;
use crate::{Error, Limits, Result};

use super::codec::{COMP_OBJ, OLE, WORKBOOK};
use super::semantic::Topology;

pub(super) const BOF: Kind = Kind::from_wire(0x0809);
pub(super) const EOF: Kind = Kind::from_wire(0x000A);
pub(super) const BOF_BYTES: usize = 16;
pub(super) const OGRAPH_VERSION: u16 = 0x0680;
pub(super) const GLOBALS: u16 = 0x0005;
pub(super) const CHART_SHEET: u16 = 0x8000;
pub(super) const OGRAPH_YEAR_1996: u16 = 0x07CC;
pub(super) const OGRAPH_YEAR_1997: u16 = 0x07CD;
pub(super) const REQUIRED_PLATFORM_FLAGS: u32 = 0x0000_0009;
pub(super) const FORBIDDEN_PLATFORM_FLAGS: u32 = 0x0000_0136;
pub(super) const RESERVED1: u32 = 0xFFF8_0000;
pub(super) const RESERVED2: u32 = 0xFFFF_F000;

#[derive(Debug, Clone, Copy)]
pub(super) struct ValidatedPackage {
    pub(super) topology: Topology,
    pub(super) workbook: WorkbookLayout,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct WorkbookLayout {
    pub(super) chart_start: usize,
    pub(super) chart_end: usize,
    pub(super) stream_end: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StreamState {
    GlobalsBof,
    Globals,
    ChartBof,
    Chart,
    Done,
}

impl WorkbookLayout {
    pub(super) fn check(self, len: usize) -> Result<()> {
        if self.stream_end != len
            || self.chart_start > self.chart_end
            || self.chart_end > self.stream_end
        {
            return workbook_error(0, "validated Workbook layout no longer matches its bytes");
        }
        Ok(())
    }
}

pub(super) fn validate(bytes: &[u8], limits: Limits) -> Result<ValidatedPackage> {
    let mut cfb = OleFile::open(Cursor::new(bytes))?;
    let entries = cfb.list_directory_entries(&[])?;
    check_limit("root entries", entries.len(), limits.max_streams)?;

    let mut workbook = None;
    let mut comp_obj = None;
    let mut ole = None;
    for entry in entries {
        validate_entry(entry, limits)?;
        let slot = match entry.name.as_str() {
            WORKBOOK => &mut workbook,
            COMP_OBJ => &mut comp_obj,
            OLE => &mut ole,
            _ => {
                return Err(Error::UnexpectedEntry {
                    name: entry.name.clone(),
                    entry_type: entry.entry_type,
                });
            },
        };
        if slot.replace(entry.size).is_some() {
            return Err(Error::DuplicateStream {
                name: entry.name.clone(),
            });
        }
    }

    let workbook_bytes = workbook.ok_or(Error::MissingStream { name: WORKBOOK })?;
    check_limit_u64("Workbook bytes", workbook_bytes, limits.max_workbook_bytes)?;
    let workbook = cfb.open_stream(&[WORKBOOK])?;
    check_limit("Workbook bytes", workbook.len(), limits.max_workbook_bytes)?;
    let workbook_layout = validate_workbook(&workbook, limits)?;

    Ok(ValidatedPackage {
        topology: Topology {
            workbook_bytes,
            comp_obj_bytes: comp_obj,
            ole_bytes: ole,
        },
        workbook: workbook_layout,
    })
}

fn validate_entry(entry: &DirectoryEntry, limits: Limits) -> Result<()> {
    if entry.entry_type != STGTY_STREAM {
        return Err(Error::UnexpectedEntry {
            name: entry.name.clone(),
            entry_type: entry.entry_type,
        });
    }
    check_limit_u64("stream bytes", entry.size, limits.max_stream_bytes)
}

pub(super) fn validate_workbook(bytes: &[u8], limits: Limits) -> Result<WorkbookLayout> {
    let mut state = StreamState::GlobalsBof;
    let mut chart_start = None;
    let mut chart_end = None;
    let mut chart_records = 0usize;
    for item in Records::with_limits(bytes, limits.biff)? {
        let record = item?;
        state = match state {
            StreamState::GlobalsBof => {
                validate_bof(record, GLOBALS)?;
                StreamState::Globals
            },
            StreamState::Globals => {
                if record.kind() == BOF {
                    return workbook_error(record.offset(), "nested BOF in globals substream");
                }
                if record.kind() == EOF {
                    validate_eof(record)?;
                    StreamState::ChartBof
                } else {
                    StreamState::Globals
                }
            },
            StreamState::ChartBof => {
                validate_bof(record, CHART_SHEET)?;
                chart_start = Some(record.offset());
                chart_records = 1;
                StreamState::Chart
            },
            StreamState::Chart => {
                chart_records = chart_records.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "chart record count",
                })?;
                check_limit(
                    "chart record count",
                    chart_records,
                    limits.max_chart_records,
                )?;
                if record.kind() == BOF {
                    return workbook_error(record.offset(), "nested BOF in chart substream");
                }
                if record.kind() == EOF {
                    validate_eof(record)?;
                    chart_end = Some(record.offset().checked_add(record.encoded().len()).ok_or(
                        Error::SizeOverflow {
                            resource: "chart substream",
                        },
                    )?);
                    StreamState::Done
                } else {
                    StreamState::Chart
                }
            },
            StreamState::Done => {
                return workbook_error(record.offset(), "records follow the chart substream EOF");
            },
        };
    }

    match state {
        StreamState::Done => {
            let chart_start = chart_start.ok_or(Error::InvalidWorkbook {
                offset: 0,
                reason: "missing chart substream start",
            })?;
            let chart_end = chart_end.ok_or(Error::InvalidWorkbook {
                offset: bytes.len(),
                reason: "missing chart substream end",
            })?;
            Ok(WorkbookLayout {
                chart_start,
                chart_end,
                stream_end: bytes.len(),
            })
        },
        StreamState::GlobalsBof => workbook_error(0, "missing globals substream BOF"),
        StreamState::Globals => workbook_error(bytes.len(), "missing globals substream EOF"),
        StreamState::ChartBof => workbook_error(bytes.len(), "missing chart substream BOF"),
        StreamState::Chart => workbook_error(bytes.len(), "missing chart substream EOF"),
    }
}

fn validate_bof(record: RecordRef<'_>, expected_doc_type: u16) -> Result<()> {
    if record.kind() != BOF {
        return workbook_error(record.offset(), "substream does not begin with BOF");
    }
    if record.payload().len() != BOF_BYTES {
        return workbook_error(record.offset(), "BOF payload is not 16 bytes");
    }
    let payload = record.payload();
    let version = le_u16(payload, 0).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF version is truncated",
    })?;
    let doc_type = le_u16(payload, 2).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF docType is truncated",
    })?;
    if version != OGRAPH_VERSION {
        return workbook_error(record.offset(), "BOF version is not 0x0680");
    }
    if doc_type != expected_doc_type {
        return if expected_doc_type == GLOBALS {
            workbook_error(record.offset(), "first BOF docType is not workbook globals")
        } else {
            workbook_error(record.offset(), "second BOF docType is not chart sheet")
        };
    }
    let year = le_u16(payload, 6).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF application year is truncated",
    })?;
    if !matches!(year, OGRAPH_YEAR_1996 | OGRAPH_YEAR_1997) {
        return workbook_error(record.offset(), "BOF application year is invalid");
    }
    let flags = le_u32(payload, 8).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF platform flags are truncated",
    })?;
    if flags & REQUIRED_PLATFORM_FLAGS != REQUIRED_PLATFORM_FLAGS
        || flags & FORBIDDEN_PLATFORM_FLAGS != 0
    {
        return workbook_error(record.offset(), "BOF platform flags are invalid");
    }
    if flags & RESERVED1 != 0 {
        return workbook_error(record.offset(), "BOF reserved1 bits are nonzero");
    }
    let highest = (flags >> 14) & 0xF;
    if !valid_version(highest) {
        return workbook_error(
            record.offset(),
            "BOF highest application version is invalid",
        );
    }
    let versions = le_u32(payload, 12).ok_or(Error::InvalidWorkbook {
        offset: record.offset(),
        reason: "BOF version flags are truncated",
    })?;
    if versions & 0xFF != 0x06 {
        return workbook_error(record.offset(), "BOF lowest BIFF version is not 0x06");
    }
    if versions & RESERVED2 != 0 {
        return workbook_error(record.offset(), "BOF reserved2 bits are nonzero");
    }
    let last = (versions >> 8) & 0xF;
    if !valid_version(last) || last > highest {
        return workbook_error(
            record.offset(),
            "BOF last-saved application version is invalid",
        );
    }
    Ok(())
}

fn validate_eof(record: RecordRef<'_>) -> Result<()> {
    if !record.payload().is_empty() {
        return workbook_error(record.offset(), "EOF record has a non-empty payload");
    }
    Ok(())
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let pair = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([*pair.first()?, *pair.get(1)?]))
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    let value = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([
        *value.first()?,
        *value.get(1)?,
        *value.get(2)?,
        *value.get(3)?,
    ]))
}

const fn valid_version(value: u32) -> bool {
    matches!(value, 0 | 1 | 2 | 3 | 4 | 6)
}

fn workbook_error<T>(offset: usize, reason: &'static str) -> Result<T> {
    Err(Error::InvalidWorkbook { offset, reason })
}

pub(super) fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed: as_u64(observed),
            maximum: as_u64(maximum),
        });
    }
    Ok(())
}

fn check_limit_u64(resource: &'static str, observed: u64, maximum: usize) -> Result<()> {
    let maximum = as_u64(maximum);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed,
            maximum,
        });
    }
    Ok(())
}
