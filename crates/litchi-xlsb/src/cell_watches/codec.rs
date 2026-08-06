//! BIFF12 payload and worksheet-stream conversion for cell watches.

use super::model::{Item, UnknownRecord, Watch};
use super::phonetic::{Alignment, Info, Type};
use super::validation;
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Kind, Limits, Records, Writer, kind};
use std::sync::Arc;

/// One source record span. The source allocation is shared by snapshots and
/// edits, so untouched records are not copied one-by-one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Span {
    pub(crate) kind: Kind,
    pub(crate) start: usize,
    pub(crate) payload_start: usize,
    pub(crate) end: usize,
}

impl Span {
    pub(crate) fn payload<'a>(self, source: &'a [u8]) -> &'a [u8] {
        &source[self.payload_start..self.end]
    }

    pub(crate) fn raw<'a>(self, source: &'a [u8]) -> &'a [u8] {
        &source[self.start..self.end]
    }
}

/// Parsed worksheet stream state consumed by the semantic snapshot.
#[derive(Debug, Clone)]
pub(crate) struct Parsed {
    pub(crate) source: Arc<[u8]>,
    pub(crate) records: Vec<Span>,
    pub(crate) watch_block: Option<(usize, usize)>,
    pub(crate) items: Vec<Item>,
    pub(crate) phonetic_index: Option<usize>,
    pub(crate) phonetic: Option<Info>,
    pub(crate) end_sheet_index: usize,
}

pub(crate) fn parse_watch(payload: &[u8]) -> Result<Watch> {
    if payload.len() != 8 {
        return Err(Error::InvalidLength {
            expected: 8,
            found: payload.len(),
        });
    }
    let mut cursor = Cursor::with_limits(payload, "BrtCellWatch", Limits::DEFAULT);
    let row = cursor.read_u32()?;
    let column = cursor.read_u32()?;
    cursor.finish()?;
    Watch::new(row, column)
}

pub(crate) fn write_watch(watch: Watch) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(8);
    let mut writer = Writer::new(&mut payload);
    writer.write_u32(watch.row())?;
    writer.write_u32(watch.column())?;
    Ok(payload)
}

pub(crate) fn parse_phonetic(payload: &[u8]) -> Result<Info> {
    if payload.len() != 10 {
        return Err(Error::InvalidLength {
            expected: 10,
            found: payload.len(),
        });
    }
    let mut cursor = Cursor::with_limits(payload, "BrtPhoneticInfo", Limits::DEFAULT);
    let font_index = cursor.read_u16()?;
    let phonetic_type = Type::from_wire(cursor.read_u32()?)?;
    let alignment = Alignment::from_wire(cursor.read_u32()?)?;
    cursor.finish()?;
    Ok(Info::new(font_index, phonetic_type, alignment))
}

pub(crate) fn write_phonetic(info: Info) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(10);
    let mut writer = Writer::new(&mut payload);
    writer.write_u16(info.font_index())?;
    writer.write_u32(info.phonetic_type().wire())?;
    writer.write_u32(info.alignment().wire())?;
    Ok(payload)
}

pub(crate) fn encode_record(record_kind: Kind, payload: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(payload.len().saturating_add(5));
    Writer::new(&mut output).write_record(record_kind, payload)?;
    Ok(output)
}

pub(crate) fn parse_stream(data: &[u8]) -> Result<Parsed> {
    let source: Arc<[u8]> = Arc::from(data);
    let mut iterator = Records::with_limits(&source, Limits::DEFAULT);
    let mut records = Vec::new();
    let mut watch_begin = None;
    let mut watch_end = None;
    let mut items = Vec::new();
    let mut phonetic_index = None;
    let mut phonetic = None;
    let mut end_sheet_data = None;
    let mut end_sheet_index = None;
    let mut opaque_records = 0usize;
    let mut opaque_payload = 0usize;

    while let Some(result) = iterator.next() {
        let record = result?;
        if records.len() >= super::MAX_RECORDS {
            return Err(Error::InvalidLength {
                expected: super::MAX_RECORDS,
                found: records.len().saturating_add(1),
            });
        }
        let end = iterator.offset();
        let payload_start =
            end.checked_sub(record.payload().len())
                .ok_or(Error::InvalidLength {
                    expected: record.payload().len(),
                    found: end,
                })?;
        let index = records.len();
        records.push(Span {
            kind: record.kind(),
            start: record.offset(),
            payload_start,
            end,
        });

        match record.kind() {
            kind::BEGIN_CELL_WATCHES => {
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                if end_sheet_data.is_none() {
                    return Err(validation::invalid(
                        "BrtBeginCellWatches",
                        "collection occurs before BrtEndSheetData",
                    ));
                }
                if watch_begin.replace(index).is_some() || watch_end.is_some() {
                    return Err(validation::invalid(
                        "BrtBeginCellWatches",
                        "duplicate or out-of-order collection",
                    ));
                }
            },
            kind::END_CELL_WATCHES => {
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                if watch_begin.is_none() || watch_end.is_some() {
                    return Err(validation::invalid(
                        "BrtEndCellWatches",
                        "record has no matching begin",
                    ));
                }
                watch_end = Some(index);
            },
            kind::CELL_WATCH => {
                if watch_begin.is_none() || watch_end.is_some() {
                    return Err(validation::invalid(
                        "BrtCellWatch",
                        "record occurs outside the watch collection",
                    ));
                }
                items.push(Item::Watch(parse_watch(record.payload())?));
            },
            kind::PHONETIC_INFO => {
                if end_sheet_data.is_none() {
                    return Err(validation::invalid(
                        "BrtPhoneticInfo",
                        "record occurs before BrtEndSheetData",
                    ));
                }
                if phonetic_index.replace(index).is_some() {
                    return Err(validation::invalid("BrtPhoneticInfo", "duplicate record"));
                }
                phonetic = Some(parse_phonetic(record.payload())?);
            },
            kind::END_SHEET_DATA => {
                if end_sheet_data.replace(index).is_some() {
                    return Err(validation::invalid("BrtEndSheetData", "duplicate record"));
                }
                if watch_begin.is_some() && watch_end.is_none() {
                    return Err(validation::invalid(
                        "BrtCellWatches",
                        "collection is not closed before BrtEndSheetData",
                    ));
                }
            },
            kind::END_SHEET => {
                if watch_begin.is_some() && watch_end.is_none() {
                    return Err(validation::invalid(
                        "BrtCellWatches",
                        "collection is not closed before BrtEndSheet",
                    ));
                }
                if end_sheet_index.replace(index).is_some() {
                    return Err(validation::invalid("BrtEndSheet", "duplicate record"));
                }
            },
            _ if watch_begin.is_some() && watch_end.is_none() => {
                opaque_records = opaque_records.saturating_add(1);
                opaque_payload = opaque_payload.checked_add(record.payload().len()).ok_or(
                    Error::InvalidLength {
                        expected: super::MAX_OPAQUE_PAYLOAD,
                        found: usize::MAX,
                    },
                )?;
                if opaque_records > super::MAX_OPAQUE_RECORDS
                    || opaque_payload > super::MAX_OPAQUE_PAYLOAD
                {
                    return Err(Error::InvalidLength {
                        expected: super::MAX_OPAQUE_RECORDS,
                        found: opaque_records,
                    });
                }
                items.push(Item::Unknown(index));
            },
            _ => {},
        }
    }

    let end_sheet_data = end_sheet_data
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtEndSheetData".to_string()))?;
    let end_sheet_index =
        end_sheet_index.ok_or_else(|| Error::UnexpectedEndOfStream("BrtEndSheet".to_string()))?;
    if end_sheet_index <= end_sheet_data {
        return Err(validation::invalid(
            "BrtEndSheet",
            "record occurs before BrtEndSheetData",
        ));
    }
    if watch_begin.is_some() != watch_end.is_some() {
        return Err(Error::UnexpectedEndOfStream(
            "BrtEndCellWatches".to_string(),
        ));
    }
    if let (Some(begin), Some(end)) = (watch_begin, watch_end) {
        if begin >= end {
            return Err(validation::invalid(
                "BrtCellWatches",
                "collection boundaries are inverted",
            ));
        }
    }
    let watches = super::model::Watches::from_validated(
        items
            .iter()
            .filter_map(|item| match item {
                Item::Watch(watch) => Some(*watch),
                Item::Unknown(_) => None,
            })
            .collect(),
    );
    validation::watches(&watches)?;
    Ok(Parsed {
        source,
        records,
        watch_block: watch_begin.zip(watch_end),
        items,
        phonetic_index,
        phonetic,
        end_sheet_index,
    })
}

pub(crate) fn opaque_record(parsed: &Parsed, index: usize) -> Result<UnknownRecord> {
    let span = parsed.records.get(index).ok_or_else(|| {
        validation::invalid("opaque BIFF12 record", "record index is out of bounds")
    })?;
    UnknownRecord::new(span.kind.get(), span.payload(&parsed.source))
}
