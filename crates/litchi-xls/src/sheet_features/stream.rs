//! Record-level views of a BIFF8 Workbook stream for the worksheet-feature
//! editor (record 0327).
//!
//! The editor never decodes cells. It splits the workbook globals and one
//! worksheet substream into raw records (an embedded chart substream is one
//! opaque unit), finds where each feature lives by its record type, and
//! splices whole records. A splice keeps every other byte and repairs the
//! absolute stream offsets that move with it:
//!
//! * `BoundSheet8.lbPlyPos` of every later substream (`[MS-XLS]` 2.4.28);
//! * every `Index` offset (`ibXF` and each `rgibRw`, 2.4.144) at or after the
//!   splice, in the edited worksheet and in every later one;
//! * the `ExtSST` bucket offsets (2.4.107) when a globals splice precedes the
//!   shared-string table.
//!
//! `DBCell` offsets are relative to their own record and never move.

use crate::{Error, Result};
use litchi_core::edit::RefusalKind;

pub(super) const BOF: u16 = 0x0809;
pub(super) const EOF: u16 = 0x000a;
pub(super) const CONTINUE: u16 = 0x003c;
pub(super) const BOUND_SHEET: u16 = 0x0085;
pub(super) const INDEX: u16 = 0x020b;
pub(super) const DIMENSIONS: u16 = 0x0200;
pub(super) const SST: u16 = 0x00fc;
pub(super) const EXT_SST: u16 = 0x00ff;
pub(super) const FILE_PASS: u16 = 0x002f;
pub(super) const MSO_DRAWING_GROUP: u16 = 0x00eb;
pub(super) const MSO_DRAWING: u16 = 0x00ec;
pub(super) const MSO_DRAWING_SELECTION: u16 = 0x00ed;
pub(super) const OBJ: u16 = 0x005d;
pub(super) const TXO: u16 = 0x01b6;
pub(super) const NOTE: u16 = 0x001c;
pub(super) const HLINK: u16 = 0x01b8;
pub(super) const HLINK_TOOLTIP: u16 = 0x0800;
pub(super) const COND_FMT: u16 = 0x01b0;
pub(super) const CF: u16 = 0x01b1;
pub(super) const COND_FMT12: u16 = 0x0879;
pub(super) const CF12: u16 = 0x087a;
pub(super) const CF_EX: u16 = 0x087b;
pub(super) const DVAL: u16 = 0x01b2;
pub(super) const DV: u16 = 0x01be;
pub(super) const AUTO_FILTER_INFO: u16 = 0x009d;
pub(super) const AUTO_FILTER: u16 = 0x009e;
pub(super) const AUTO_FILTER12: u16 = 0x087e;
pub(super) const CONTINUE_FRT12: u16 = 0x087f;
pub(super) const FILTER_MODE: u16 = 0x009b;
pub(super) const HORIZONTAL_PAGE_BREAKS: u16 = 0x001b;
pub(super) const VERTICAL_PAGE_BREAKS: u16 = 0x001a;
pub(super) const WS_BOOL: u16 = 0x0081;
pub(super) const SYNC: u16 = 0x0097;
pub(super) const LPR: u16 = 0x0098;
pub(super) const WINDOW2: u16 = 0x023e;
pub(super) const LBL: u16 = 0x0018;
pub(super) const EXTERN_SHEET: u16 = 0x0017;
pub(super) const SUP_BOOK: u16 = 0x01ae;
pub(super) const PALETTE: u16 = 0x0092;

/// The record types of a worksheet's cell table (`CELLTABLE`, 2.1.7.20.5).
const CELL_TABLE: [u16; 14] = [
    0x0208, // Row
    0x00d7, // DBCell
    0x0203, // Number
    0x027e, // RK
    0x00bd, // MulRk
    0x00be, // MulBlank
    0x0201, // Blank
    0x0205, // BoolErr
    0x00fd, // LabelSst
    0x0204, // Label
    0x00d6, // RString
    0x0006, // Formula
    0x0207, // String
    0x04bc, // ShrFmla
];
const ARRAY: u16 = 0x0221;
const TABLE: u16 = 0x0236;

/// The position of a record type after the cell table in the worksheet
/// substream grammar (`[MS-XLS]` 2.1.7.20.5); unknown types have none.
const fn rank(kind: u16) -> Option<u8> {
    Some(match kind {
        MSO_DRAWING | OBJ | TXO | MSO_DRAWING_SELECTION | 0x00e9 => 1,
        0x0866 => 2, // HFPicture
        NOTE => 3,
        0x00b0..=0x00b6 | 0x00c5 | 0x00c6 | 0x00f1 | 0x0100 | 0x0802 => 4, // PIVOTVIEW
        0x0050..=0x0052 | 0x01b5 => 5,                                     // DCON
        WINDOW2 | 0x088b | 0x00a0 | 0x0041 | 0x001d => 6,                  // WINDOW
        0x01aa | 0x01ab => 7,                                              // CUSTOMVIEW
        0x0090 => 8,                                                       // Sort
        0x0099 => 9,                                                       // DxGCol
        0x00e5 => 10,                                                      // MergeCells
        0x015f => 11,                                                      // LRng
        0x01ad | 0x0803..=0x0807 => 12,                                    // QUERYTABLE
        0x00ef => 13,                                                      // PhoneticInfo
        COND_FMT | CF | COND_FMT12 | CF12 | CF_EX => 14,
        HLINK | HLINK_TOOLTIP => 15,
        DVAL | DV => 16,
        0x01ba => 17,          // CodeName
        0x0801 => 18,          // WebPub
        0x086c => 19,          // CellWatch
        0x0862 => 20,          // SheetExt
        0x0867 | 0x0868 => 21, // FEAT
        0x0871 | 0x0872 | 0x0877 | 0x0878 | AUTO_FILTER12 | CONTINUE_FRT12 => 22,
        EOF => 99,
        _ => return None,
    })
}

/// One raw record: its type and its absolute byte range in the stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Raw {
    pub(super) kind: u16,
    pub(super) start: usize,
    pub(super) end: usize,
}

impl Raw {
    pub(super) fn payload<'s>(&self, stream: &'s [u8]) -> &'s [u8] {
        stream.get(self.start + 4..self.end).unwrap_or(&[])
    }

    pub(super) fn bytes<'s>(&self, stream: &'s [u8]) -> &'s [u8] {
        stream.get(self.start..self.end).unwrap_or(&[])
    }
}

fn header(stream: &[u8], at: usize) -> Result<(u16, usize)> {
    let bytes = stream
        .get(at..at + 4)
        .ok_or_else(|| Error::InvalidData("BIFF record header is truncated".into()))?;
    let kind = u16::from_le_bytes([bytes[0], bytes[1]]);
    let length = usize::from(u16::from_le_bytes([bytes[2], bytes[3]]));
    let end = at + 4 + length;
    if end > stream.len() {
        return Err(Error::InvalidData(
            "BIFF record payload is truncated".into(),
        ));
    }
    Ok((kind, end))
}

/// The records of one substream that starts with `BOF` at `start`, up to and
/// including its own `EOF`. An embedded substream (`BOF`…`EOF`, an embedded
/// chart after its drawing object) is one raw unit of kind `BOF`.
pub(super) fn substream(stream: &[u8], start: usize) -> Result<Vec<Raw>> {
    let (kind, end) = header(stream, start)?;
    if kind != BOF {
        return Err(Error::InvalidData(
            "a substream does not start with BOF".into(),
        ));
    }
    let mut records = vec![Raw { kind, start, end }];
    let mut at = end;
    loop {
        let (kind, end) = header(stream, at)?;
        if kind == BOF {
            // Skip the embedded substream as one unit.
            let first = at;
            let mut depth = 0_usize;
            let mut cursor = at;
            let last_end = loop {
                let (kind, end) = header(stream, cursor)?;
                match kind {
                    BOF => depth += 1,
                    EOF => {
                        depth -= 1;
                        if depth == 0 {
                            break end;
                        }
                    },
                    _ => {},
                }
                cursor = end;
            };
            records.push(Raw {
                kind: BOF,
                start: first,
                end: last_end,
            });
            at = last_end;
            continue;
        }
        records.push(Raw {
            kind,
            start: at,
            end,
        });
        at = end;
        if kind == EOF {
            return Ok(records);
        }
    }
}

/// One worksheet catalog entry of the globals.
#[derive(Debug, Clone)]
pub(super) struct Bound {
    /// The `BoundSheet8` record.
    pub(super) record: Raw,
    /// `lbPlyPos`: the substream's absolute position.
    pub(super) position: usize,
    /// `dt`: 0 worksheet or dialog, 1 macro, 2 chart, 6 VBA module.
    pub(super) kind: u8,
}

/// The workbook globals substream and its tab catalog.
#[derive(Debug, Clone)]
pub(super) struct Globals {
    pub(super) records: Vec<Raw>,
    pub(super) bounds: Vec<Bound>,
}

impl Globals {
    pub(super) fn parse(stream: &[u8]) -> Result<Self> {
        let records = substream(stream, 0)?;
        let mut bounds = Vec::new();
        for record in &records {
            if record.kind == FILE_PASS {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit sheet features",
                    "an encrypted XLS workbook's worksheet features are not edited",
                ));
            }
            if record.kind == BOUND_SHEET {
                let payload = record.payload(stream);
                if payload.len() < 6 {
                    return Err(Error::InvalidData("BoundSheet8 is truncated".into()));
                }
                let position =
                    u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]) as usize;
                bounds.push(Bound {
                    record: *record,
                    position,
                    kind: payload[5],
                });
            }
        }
        Ok(Self { records, bounds })
    }

    /// The end of the globals substream (one past its `EOF`).
    pub(super) fn end(&self) -> usize {
        self.records.last().map_or(0, |record| record.end)
    }

    /// The substream position of the catalog tab at `sheet`, when it is a
    /// worksheet.
    pub(super) fn worksheet(&self, sheet: usize) -> Result<usize> {
        let bound = self
            .bounds
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("sheet position {sheet}")))?;
        if bound.kind != 0 {
            return Err(Error::UnsupportedFeature(
                "worksheet features belong to ordinary worksheets".into(),
            ));
        }
        Ok(bound.position)
    }

    pub(super) fn first(&self, kind: u16) -> Option<usize> {
        self.records.iter().position(|record| record.kind == kind)
    }
}

/// One worksheet substream and the landmarks the feature families use.
#[derive(Debug, Clone)]
pub(super) struct Sheet {
    pub(super) records: Vec<Raw>,
    /// Index of the `Dimensions` record.
    pub(super) dimensions: usize,
    /// Index one past the last record of the cell table.
    pub(super) cells_end: usize,
}

impl Sheet {
    pub(super) fn parse(stream: &[u8], start: usize) -> Result<Self> {
        let records = substream(stream, start)?;
        let dimensions = records
            .iter()
            .position(|record| record.kind == DIMENSIONS)
            .ok_or_else(|| Error::InvalidData("a worksheet has no Dimensions record".into()))?;
        let mut cells_end = dimensions + 1;
        for (index, record) in records.iter().enumerate().skip(dimensions + 1) {
            if CELL_TABLE.contains(&record.kind)
                || record.kind == ARRAY
                || record.kind == TABLE
                || (record.kind == CONTINUE && index == cells_end)
            {
                cells_end = index + 1;
            } else if rank(record.kind).is_some() {
                break;
            }
        }
        Ok(Self {
            records,
            dimensions,
            cells_end,
        })
    }

    pub(super) fn eof(&self) -> usize {
        self.records.len().saturating_sub(1)
    }

    /// The records after the cell table.
    pub(super) fn tail(&self) -> impl Iterator<Item = (usize, &Raw)> {
        self.records.iter().enumerate().skip(self.cells_end)
    }

    /// The index of the first record after the cell table whose grammar
    /// position is after `position` (or that `stop` names), else the `EOF`.
    fn before_rank(&self, position: u8, stop: &[u16]) -> usize {
        self.tail()
            .find(|(_, record)| {
                stop.contains(&record.kind) || rank(record.kind).is_some_and(|rank| rank > position)
            })
            .map_or(self.eof(), |(index, _)| index)
    }

    /// The record index where a new drawing object block starts.
    pub(super) fn objects_insert(&self) -> usize {
        self.before_rank(1, &[])
    }

    /// The index where a new `Note` is inserted (after the existing ones).
    pub(super) fn notes_insert(&self) -> usize {
        self.before_rank(3, &[])
    }

    /// The index where a new hyperlink is inserted (after the existing ones).
    pub(super) fn hyperlinks_insert(&self) -> usize {
        self.before_rank(15, &[])
    }

    /// The index where a new conditional format is inserted: after the
    /// existing `CondFmt`/`CondFmt12` blocks, before any `CFEx`.
    pub(super) fn formats_insert(&self) -> usize {
        self.before_rank(14, &[CF_EX])
    }

    /// The index where a new `DVal` block is inserted.
    pub(super) fn validations_insert(&self) -> usize {
        self.before_rank(16, &[])
    }

    /// The index where a page-break record of `kind` is inserted when the
    /// worksheet has none: after `WsBool`, `Sync` and `LPr` (and, for
    /// vertical breaks, after the horizontal ones), before `Dimensions`.
    pub(super) fn breaks_insert(&self, kind: u16) -> Result<usize> {
        let mut after = vec![WS_BOOL, SYNC, LPR];
        if kind == VERTICAL_PAGE_BREAKS {
            after.push(HORIZONTAL_PAGE_BREAKS);
        }
        self.records[..self.dimensions]
            .iter()
            .rposition(|record| after.contains(&record.kind))
            .map(|index| index + 1)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit sheet features",
                    "a worksheet without WsBool takes no page breaks",
                )
            })
    }

    pub(super) fn find(&self, kind: u16) -> Option<usize> {
        self.records.iter().position(|record| record.kind == kind)
    }
}

/// Build one record with its header; payloads above 8,224 bytes are refused.
pub(super) fn record(kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .ok()
        .filter(|length| *length <= 8_224)
        .ok_or_else(|| Error::InvalidData("a BIFF record payload exceeds 8,224 bytes".into()))?;
    let mut out = Vec::with_capacity(payload.len() + 4);
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&length.to_le_bytes());
    out.extend_from_slice(payload);
    Ok(out)
}

/// Replace `stream[start..end]` with `replacement` and repair the absolute
/// offsets that move with it. `start..end` lies inside one substream.
pub(super) fn splice(
    stream: &mut Vec<u8>,
    start: usize,
    end: usize,
    replacement: &[u8],
) -> Result<()> {
    if start > end || end > stream.len() {
        return Err(Error::InvalidData(
            "a splice range is outside the stream".into(),
        ));
    }
    let delta = i64::try_from(replacement.len()).unwrap_or(i64::MAX)
        - i64::try_from(end - start).unwrap_or(i64::MAX);
    if delta == 0 && stream[start..end] == *replacement {
        return Ok(());
    }
    let globals = Globals::parse(stream)?;
    let in_globals = start < globals.end();
    // Absolute offsets at or after `end` move by `delta`.
    let moved = |offset: usize| offset >= end;
    let shift = |offset: usize| -> Result<u32> {
        let value = i64::try_from(offset).unwrap_or(i64::MAX) + delta;
        u32::try_from(value).map_err(|_| {
            Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit sheet features",
                "a stream offset moves outside 32 bits",
            )
        })
    };
    // Collect fixed-width field patches against the source positions first.
    let mut patches: Vec<(usize, u32)> = Vec::new();
    for bound in &globals.bounds {
        if moved(bound.position) {
            patches.push((bound.record.start + 4, shift(bound.position)?));
        }
        if bound.kind == 0 || bound.kind == 1 {
            // Index offsets of every worksheet substream; an edit inside a
            // worksheet moves only the offsets after it.
            if let Ok(records) = substream(stream, bound.position) {
                if let Some(index) = records.iter().take(8).find(|record| record.kind == INDEX) {
                    let payload = index.payload(stream);
                    if payload.len() >= 16 && (payload.len() - 16) % 4 == 0 {
                        for field in (12..payload.len()).step_by(4) {
                            let value = u32::from_le_bytes([
                                payload[field],
                                payload[field + 1],
                                payload[field + 2],
                                payload[field + 3],
                            ]) as usize;
                            if value != 0 && moved(value) {
                                patches.push((index.start + 4 + field, shift(value)?));
                            }
                        }
                    }
                }
            }
        }
    }
    if in_globals {
        if let Some(index) = globals.first(EXT_SST) {
            let record = globals.records[index];
            let payload = record.payload(stream);
            for field in (2..payload.len()).step_by(8) {
                if field + 4 > payload.len() {
                    break;
                }
                let value = u32::from_le_bytes([
                    payload[field],
                    payload[field + 1],
                    payload[field + 2],
                    payload[field + 3],
                ]) as usize;
                if moved(value) {
                    patches.push((record.start + 4 + field, shift(value)?));
                }
            }
        }
    }
    for (at, value) in patches {
        stream[at..at + 4].copy_from_slice(&value.to_le_bytes());
    }
    let mut rebuilt = Vec::new();
    rebuilt
        .try_reserve_exact(stream.len() - (end - start) + replacement.len())
        .map_err(|_| Error::Allocation("splicing an XLS Workbook stream"))?;
    rebuilt.extend_from_slice(&stream[..start]);
    rebuilt.extend_from_slice(replacement);
    rebuilt.extend_from_slice(&stream[end..]);
    *stream = rebuilt;
    Ok(())
}

pub(super) fn u16_at(bytes: &[u8], at: usize) -> u16 {
    bytes
        .get(at..at + 2)
        .map_or(0, |bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
}

pub(super) fn u32_at(bytes: &[u8], at: usize) -> u32 {
    bytes.get(at..at + 4).map_or(0, |bytes| {
        u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
    })
}
