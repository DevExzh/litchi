//! BIFF8 shared-workbook revision log stream (MS-XLS 2.1.7.14).
//!
//! A shared workbook stores its revision logs in a dedicated CFB stream named
//! `Revision Log`, separate from the `Workbook` stream. This module locates
//! that stream in a CFB directory listing and parses its record sequence into
//! a typed, inert model. Parsing never applies, rejects, or replays any of
//! the recorded revisions.
//!
//! Record sequence (MS-XLS 2.1.7.14 ABNF):
//!
//! ```text
//! REVISION = RRDInfo [FileLock] [UsrExcl] *(HEADER *(revision)) EOF
//! HEADER   = RRDHead [RRTabId]
//! ```

use super::revision_records as records;
use super::{Error, Result};
use records::{
    CONTINUE_RECORD_TYPE, EOF_RECORD_TYPE, FILE_LOCK_RECORD_TYPE, FileLock, NOTE_RECORD_TYPE,
    RECORD_HEADER_LEN, RR_AUTO_FMT_RECORD_TYPE, RR_FORMAT_RECORD_TYPE, RR_INSERT_SH_RECORD_TYPE,
    RR_TAB_ID_RECORD_TYPE, RRD_CHG_CELL_RECORD_TYPE, RRD_CONFLICT_RECORD_TYPE,
    RRD_DEF_NAME_RECORD_TYPE, RRD_HEAD_RECORD_TYPE, RRD_INFO_RECORD_TYPE,
    RRD_INS_DEL_BEGIN_RECORD_TYPE, RRD_INS_DEL_END_RECORD_TYPE, RRD_INS_DEL_RECORD_TYPE,
    RRD_MOVE_BEGIN_RECORD_TYPE, RRD_MOVE_END_RECORD_TYPE, RRD_MOVE_RECORD_TYPE,
    RRD_REN_SHEET_RECORD_TYPE, RRD_RST_ETXP_RECORD_TYPE, RRD_TQSIF_RECORD_TYPE,
    RRD_USER_VIEW_RECORD_TYPE, RrInsertSh, RrTabId, RrdChgCell, RrdConflict, RrdHead, RrdInfo,
    RrdInsDel, RrdMove, RrdRenSheet, RrdUserView, USR_EXCL_RECORD_TYPE, UsrExcl,
};

/// Name of the revision stream in the CFB root storage (MS-XLS 2.1.7.14).
pub const REVISION_LOG_STREAM_NAME: &str = "Revision Log";

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Locate the `Revision Log` stream in a CFB directory listing.
///
/// CFB storage member names compare case-insensitively (MS-CFB 2.6.4), and the
/// stream MUST live in the root storage. Returns the directory entry's actual
/// name so callers can open the stream exactly as stored.
#[must_use]
pub fn find_revision_log_stream(stream_paths: &[Vec<String>]) -> Option<&str> {
    stream_paths.iter().find_map(|path| {
        if path.len() == 1 && path[0].eq_ignore_ascii_case(REVISION_LOG_STREAM_NAME) {
            Some(path[0].as_str())
        } else {
            None
        }
    })
}

/// One framed BIFF record borrowed from the stream bytes.
struct FramedRecord<'a> {
    record_type: u16,
    payload: &'a [u8],
}

/// Split the stream into BIFF records, validating the record framing.
fn frame_records(data: &[u8]) -> Result<Vec<FramedRecord<'_>>> {
    let mut records_out = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let header = data
            .get(offset..offset + RECORD_HEADER_LEN)
            .ok_or_else(|| {
                Error::UnexpectedEndOfStream("truncated revision record header".to_string())
            })?;
        let record_type = u16::from_le_bytes([header[0], header[1]]);
        let length = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let payload = data
            .get(offset + RECORD_HEADER_LEN..offset + RECORD_HEADER_LEN + length)
            .ok_or_else(|| {
                invalid(
                    record_type,
                    format!("revision record payload of {length} bytes is truncated"),
                )
            })?;
        records_out.push(FramedRecord {
            record_type,
            payload,
        });
        offset += RECORD_HEADER_LEN + length;
    }
    Ok(records_out)
}

/// A revision record this module does not model in depth, preserved verbatim.
///
/// Covers `RRFormat`, `RRAutoFmt`, `RRDDefName`, `RRDTQSIF`, and `Note`
/// records; their grammars involve differential formats, parsed formulas, and
/// shape note substreams that are specified outside the revision records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueRevisionRecord {
    record_type: u16,
    payload: Vec<u8>,
}

impl OpaqueRevisionRecord {
    #[must_use]
    pub fn record_type(&self) -> u16 {
        self.record_type
    }
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }
}

/// A cell change or formatting change nested inside an insert/delete or move.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RevisionChange {
    /// An `RRDChgCell` revision with its Continue and `RRDRstEtxp` records.
    CellChange(Box<RrdChgCellRevision>),
    /// An `RRFormat` formatting revision, preserved raw.
    Format(OpaqueRevisionRecord),
}

/// An `RRDChgCell` record together with the records the ABNF attaches to it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdChgCellRevision {
    record: RrdChgCell,
    continue_payloads: Vec<Vec<u8>>,
    formatting_runs: Vec<Vec<u8>>,
}

impl RrdChgCellRevision {
    #[must_use]
    pub fn record(&self) -> &RrdChgCell {
        &self.record
    }
    /// Raw `Continue` payloads following the `RRDChgCell` record.
    #[must_use]
    pub fn continue_payloads(&self) -> &[Vec<u8>] {
        &self.continue_payloads
    }
    /// Raw `RRDRstEtxp` payloads; the count matches `cetxpRst`.
    #[must_use]
    pub fn formatting_runs(&self) -> &[Vec<u8>] {
        &self.formatting_runs
    }
}

/// An insertion or deletion of rows/columns with its nested changes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdInsDelRevision {
    record: RrdInsDel,
    deletion: bool,
    changes: Vec<RevisionChange>,
}

impl RrdInsDelRevision {
    #[must_use]
    pub fn record(&self) -> &RrdInsDel {
        &self.record
    }
    /// Whether the record was wrapped in `RRDInsDelBegin`/`RRDInsDelEnd`
    /// (the ABNF `DEL` production), as opposed to a bare insertion.
    #[must_use]
    pub fn is_deletion(&self) -> bool {
        self.deletion
    }
    #[must_use]
    pub fn changes(&self) -> &[RevisionChange] {
        &self.changes
    }
}

/// A cell-range move with its nested changes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdMoveRevision {
    record: RrdMove,
    changes: Vec<RevisionChange>,
}

impl RrdMoveRevision {
    #[must_use]
    pub fn record(&self) -> &RrdMove {
        &self.record
    }
    #[must_use]
    pub fn changes(&self) -> &[RevisionChange] {
        &self.changes
    }
}

/// One revision inside a revision header's set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Revision {
    RenSheet(RrdRenSheet),
    InsDel(RrdInsDelRevision),
    Move(RrdMoveRevision),
    InsertSheet(RrInsertSh),
    CellChange(Box<RrdChgCellRevision>),
    Conflict(RrdConflict),
    UserView(RrdUserView),
    /// `RRAutoFmt`, `RRDDefName`, `Note`, or `RRDTQSIF`, preserved raw.
    Opaque(OpaqueRevisionRecord),
}

/// One user's set of revisions: an `RRDHead`, its optional `RRTabId`, and
/// the revisions that follow it (MS-XLS 2.1.7.14 `HEADER` production).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionHeader {
    head: RrdHead,
    sheet_ids: Option<RrTabId>,
    revisions: Vec<Revision>,
}

impl RevisionHeader {
    #[must_use]
    pub fn head(&self) -> &RrdHead {
        &self.head
    }
    /// Sheet identifiers for this revision set; absent when the workbook has
    /// more than 4112 sheets (MS-XLS 2.4.241).
    #[must_use]
    pub fn sheet_ids(&self) -> Option<&RrTabId> {
        self.sheet_ids.as_ref()
    }
    #[must_use]
    pub fn revisions(&self) -> &[Revision] {
        &self.revisions
    }
}

/// The parsed `Revision Log` stream of a shared workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionLog {
    info: RrdInfo,
    file_lock: Option<FileLock>,
    exclusive_lock: Option<UsrExcl>,
    headers: Vec<RevisionHeader>,
}

impl RevisionLog {
    #[must_use]
    pub fn info(&self) -> &RrdInfo {
        &self.info
    }
    #[must_use]
    pub fn file_lock(&self) -> Option<&FileLock> {
        self.file_lock.as_ref()
    }
    #[must_use]
    pub fn exclusive_lock(&self) -> Option<&UsrExcl> {
        self.exclusive_lock.as_ref()
    }
    #[must_use]
    pub fn headers(&self) -> &[RevisionHeader] {
        &self.headers
    }
}

/// Consume the records following an `RRDChgCell`: any `Continue` records,
/// then exactly `cetxpRst` `RRDRstEtxp` records.
fn parse_chg_cell_revision(
    stream: &[FramedRecord<'_>],
    cursor: &mut usize,
) -> Result<RrdChgCellRevision> {
    let framed = &stream[*cursor];
    debug_assert_eq!(framed.record_type, RRD_CHG_CELL_RECORD_TYPE);
    let record = RrdChgCell::parse_payload(framed.payload)?;
    *cursor += 1;
    let mut continue_payloads = Vec::new();
    while stream
        .get(*cursor)
        .is_some_and(|next| next.record_type == CONTINUE_RECORD_TYPE)
    {
        continue_payloads.push(stream[*cursor].payload.to_vec());
        *cursor += 1;
    }
    let run_count = usize::from(record.formatting_run_count());
    let mut formatting_runs = Vec::with_capacity(run_count);
    for _ in 0..run_count {
        let Some(next) = stream.get(*cursor) else {
            return Err(invalid(
                RRD_RST_ETXP_RECORD_TYPE,
                "RRDChgCell is missing its RRDRstEtxp records",
            ));
        };
        if next.record_type != RRD_RST_ETXP_RECORD_TYPE {
            return Err(invalid(
                next.record_type,
                "expected an RRDRstEtxp record after RRDChgCell",
            ));
        }
        formatting_runs.push(next.payload.to_vec());
        *cursor += 1;
    }
    Ok(RrdChgCellRevision {
        record,
        continue_payloads,
        formatting_runs,
    })
}

/// Consume the `*(CHGCELL / FORMAT)` production following an `RRDInsDel` or
/// `RRDMove` record.
fn parse_nested_changes(
    stream: &[FramedRecord<'_>],
    cursor: &mut usize,
) -> Result<Vec<RevisionChange>> {
    let mut changes = Vec::new();
    while let Some(next) = stream.get(*cursor) {
        match next.record_type {
            RRD_CHG_CELL_RECORD_TYPE => {
                changes.push(RevisionChange::CellChange(Box::new(
                    parse_chg_cell_revision(stream, cursor)?,
                )));
            },
            RR_FORMAT_RECORD_TYPE => {
                changes.push(RevisionChange::Format(OpaqueRevisionRecord {
                    record_type: next.record_type,
                    payload: next.payload.to_vec(),
                }));
                *cursor += 1;
            },
            _ => break,
        }
    }
    Ok(changes)
}

fn expect_marker(
    stream: &[FramedRecord<'_>],
    cursor: &mut usize,
    expected: u16,
    name: &str,
) -> Result<()> {
    let Some(next) = stream.get(*cursor) else {
        return Err(invalid(expected, format!("missing {name} record")));
    };
    if next.record_type != expected {
        return Err(invalid(
            next.record_type,
            format!("expected {name}, found record 0x{:04X}", next.record_type),
        ));
    }
    records::validate_empty_marker(expected, next.payload, name)?;
    *cursor += 1;
    Ok(())
}

/// Parse one revision record at `stream[*cursor]`, advancing past it.
fn parse_revision(stream: &[FramedRecord<'_>], cursor: &mut usize) -> Result<Revision> {
    let framed = &stream[*cursor];
    match framed.record_type {
        RRD_REN_SHEET_RECORD_TYPE => {
            *cursor += 1;
            Ok(Revision::RenSheet(RrdRenSheet::parse_payload(
                framed.payload,
            )?))
        },
        RRD_INS_DEL_RECORD_TYPE => {
            *cursor += 1;
            let record = RrdInsDel::parse_payload(framed.payload)?;
            let changes = parse_nested_changes(stream, cursor)?;
            Ok(Revision::InsDel(RrdInsDelRevision {
                record,
                deletion: false,
                changes,
            }))
        },
        RRD_INS_DEL_BEGIN_RECORD_TYPE => {
            expect_marker(
                stream,
                cursor,
                RRD_INS_DEL_BEGIN_RECORD_TYPE,
                "RRDInsDelBegin",
            )?;
            let Some(next) = stream.get(*cursor) else {
                return Err(invalid(
                    RRD_INS_DEL_RECORD_TYPE,
                    "RRDInsDelBegin is not followed by RRDInsDel",
                ));
            };
            if next.record_type != RRD_INS_DEL_RECORD_TYPE {
                return Err(invalid(
                    next.record_type,
                    "RRDInsDelBegin is not followed by RRDInsDel",
                ));
            }
            let record = RrdInsDel::parse_payload(next.payload)?;
            *cursor += 1;
            let changes = parse_nested_changes(stream, cursor)?;
            expect_marker(stream, cursor, RRD_INS_DEL_END_RECORD_TYPE, "RRDInsDelEnd")?;
            Ok(Revision::InsDel(RrdInsDelRevision {
                record,
                deletion: true,
                changes,
            }))
        },
        RRD_MOVE_BEGIN_RECORD_TYPE => {
            expect_marker(stream, cursor, RRD_MOVE_BEGIN_RECORD_TYPE, "RRDMoveBegin")?;
            let Some(next) = stream.get(*cursor) else {
                return Err(invalid(
                    RRD_MOVE_RECORD_TYPE,
                    "RRDMoveBegin is not followed by RRDMove",
                ));
            };
            if next.record_type != RRD_MOVE_RECORD_TYPE {
                return Err(invalid(
                    next.record_type,
                    "RRDMoveBegin is not followed by RRDMove",
                ));
            }
            let record = RrdMove::parse_payload(next.payload)?;
            *cursor += 1;
            let changes = parse_nested_changes(stream, cursor)?;
            expect_marker(stream, cursor, RRD_MOVE_END_RECORD_TYPE, "RRDMoveEnd")?;
            Ok(Revision::Move(RrdMoveRevision { record, changes }))
        },
        RRD_CHG_CELL_RECORD_TYPE => Ok(Revision::CellChange(Box::new(parse_chg_cell_revision(
            stream, cursor,
        )?))),
        RRD_CONFLICT_RECORD_TYPE => {
            *cursor += 1;
            Ok(Revision::Conflict(RrdConflict::parse_payload(
                framed.payload,
            )?))
        },
        RR_INSERT_SH_RECORD_TYPE => {
            *cursor += 1;
            Ok(Revision::InsertSheet(RrInsertSh::parse_payload(
                framed.payload,
            )?))
        },
        RRD_USER_VIEW_RECORD_TYPE => {
            *cursor += 1;
            Ok(Revision::UserView(RrdUserView::parse_payload(
                framed.payload,
            )?))
        },
        RR_AUTO_FMT_RECORD_TYPE
        | RRD_DEF_NAME_RECORD_TYPE
        | NOTE_RECORD_TYPE
        | RRD_TQSIF_RECORD_TYPE => {
            *cursor += 1;
            Ok(Revision::Opaque(OpaqueRevisionRecord {
                record_type: framed.record_type,
                payload: framed.payload.to_vec(),
            }))
        },
        other => Err(invalid(
            other,
            format!("unexpected record 0x{other:04X} in a revision header"),
        )),
    }
}

/// Parse the `Revision Log` stream bytes into a typed, inert model.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse_revision_log_stream(data: &[u8]) -> Result<RevisionLog> {
    let stream = frame_records(data)?;
    if stream.len() < 2 {
        return Err(invalid(
            RRD_INFO_RECORD_TYPE,
            "revision log must contain at least RRDInfo and EOF",
        ));
    }
    if stream[0].record_type != RRD_INFO_RECORD_TYPE {
        return Err(invalid(
            stream[0].record_type,
            "revision log must start with RRDInfo",
        ));
    }
    let info = RrdInfo::parse_payload(stream[0].payload)?;
    let last = &stream[stream.len() - 1];
    if last.record_type != EOF_RECORD_TYPE {
        return Err(invalid(
            last.record_type,
            "revision log must end with an EOF record",
        ));
    }
    if !last.payload.is_empty() {
        return Err(invalid(
            EOF_RECORD_TYPE,
            "revision log EOF record has a payload",
        ));
    }

    let mut cursor = 1usize;
    let file_lock = if stream[cursor].record_type == FILE_LOCK_RECORD_TYPE {
        let lock = FileLock::parse_payload(stream[cursor].payload)?;
        cursor += 1;
        Some(lock)
    } else {
        None
    };
    let exclusive_lock = if stream[cursor].record_type == USR_EXCL_RECORD_TYPE {
        let lock = UsrExcl::parse_payload(stream[cursor].payload)?;
        cursor += 1;
        Some(lock)
    } else {
        None
    };

    let mut headers = Vec::new();
    while stream[cursor].record_type != EOF_RECORD_TYPE {
        if stream[cursor].record_type != RRD_HEAD_RECORD_TYPE {
            return Err(invalid(
                stream[cursor].record_type,
                "expected an RRDHead record in the revision log",
            ));
        }
        let head = RrdHead::parse_payload(stream[cursor].payload)?;
        cursor += 1;
        let sheet_ids = if stream[cursor].record_type == RR_TAB_ID_RECORD_TYPE {
            let ids = RrTabId::parse_payload(stream[cursor].payload)?;
            cursor += 1;
            Some(ids)
        } else {
            None
        };
        let mut revisions = Vec::new();
        while stream[cursor].record_type != RRD_HEAD_RECORD_TYPE
            && stream[cursor].record_type != EOF_RECORD_TYPE
        {
            revisions.push(parse_revision(&stream, &mut cursor)?);
        }
        headers.push(RevisionHeader {
            head,
            sheet_ids,
            revisions,
        });
    }
    if cursor != stream.len() - 1 {
        return Err(invalid(
            EOF_RECORD_TYPE,
            "an EOF record appears before the end of the revision log",
        ));
    }

    Ok(RevisionLog {
        info,
        file_lock,
        exclusive_lock,
        headers,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use records::RevisionType;

    /// Frame a BIFF record: `rt`, `cb`, payload.
    fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::with_capacity(4 + payload.len());
        data.extend_from_slice(&record_type.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn rrd(revt: u16, revid: i32, tabid: u16) -> [u8; 14] {
        let mut data = [0u8; 14];
        data[0..4].copy_from_slice(&26u32.to_le_bytes());
        data[4..8].copy_from_slice(&revid.to_le_bytes());
        data[8..10].copy_from_slice(&revt.to_le_bytes());
        data[12..14].copy_from_slice(&tabid.to_le_bytes());
        data
    }

    fn short_dtr() -> [u8; 8] {
        let mut data = [0u8; 8];
        data[0..2].copy_from_slice(&2023u16.to_le_bytes());
        data[2] = 6;
        data[3] = 30;
        data[4] = 9;
        data[5] = 15;
        data[6] = 0;
        data[7] = 5;
        data
    }

    fn string_field(field_len: usize, text: &str) -> Vec<u8> {
        let mut field = vec![0u8; field_len];
        field[1..1 + text.len()].copy_from_slice(text.as_bytes());
        field
    }

    fn rrd_info() -> Vec<u8> {
        let mut data = vec![0u8; 50];
        data[0..2].copy_from_slice(&8u16.to_le_bytes());
        data[4..6].copy_from_slice(&0x000Bu16.to_le_bytes());
        data[38..42].copy_from_slice(&99i32.to_le_bytes());
        data[42..46].copy_from_slice(&4u32.to_le_bytes());
        data[46..48].copy_from_slice(&45u16.to_le_bytes());
        record(RRD_INFO_RECORD_TYPE, &data)
    }

    fn file_lock() -> Vec<u8> {
        let mut data = vec![0u8; 162];
        data[0..4].copy_from_slice(&0x0001_0001u32.to_le_bytes());
        data[4..6].copy_from_slice(&3u16.to_le_bytes());
        data[7..10].copy_from_slice(b"Bob");
        record(FILE_LOCK_RECORD_TYPE, &data)
    }

    fn usr_excl() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&short_dtr());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&string_field(1 + 147, "Bob"));
        record(USR_EXCL_RECORD_TYPE, &data)
    }

    fn rrd_head(user: &str, guid_byte: u8) -> Vec<u8> {
        let mut data = Vec::with_capacity(158);
        let mut header = rrd(0x0020, 0, 0xFFFF);
        header[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        data.extend_from_slice(&header);
        data.extend_from_slice(&[guid_byte; 16]);
        data.extend_from_slice(&1200u16.to_le_bytes());
        data.extend_from_slice(&(user.len() as u16).to_le_bytes());
        data.extend_from_slice(&string_field(114, user));
        data.extend_from_slice(&short_dtr());
        data.extend_from_slice(&4i16.to_le_bytes());
        record(RRD_HEAD_RECORD_TYPE, &data)
    }

    fn rr_tab_id(ids: &[u16]) -> Vec<u8> {
        let mut data = Vec::new();
        for id in ids {
            data.extend_from_slice(&id.to_le_bytes());
        }
        record(RR_TAB_ID_RECORD_TYPE, &data)
    }

    fn ren_sheet() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(0x0009, 21, 1));
        data.extend_from_slice(&6u16.to_le_bytes());
        data.extend_from_slice(&string_field(255, "Budget"));
        data.extend_from_slice(&7u16.to_le_bytes());
        data.extend_from_slice(&string_field(255, "Budget2"));
        record(RRD_REN_SHEET_RECORD_TYPE, &data)
    }

    fn ins_del(revt: u16, flags: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(revt, 22, 1));
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&5u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        record(RRD_INS_DEL_RECORD_TYPE, &data)
    }

    fn chg_cell(revid: i32, formatting_runs: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(0x0008, revid, 1));
        data.extend_from_slice(&0x0002u32.to_le_bytes()); // vt = Xnum
        data.extend_from_slice(&12u16.to_le_bytes());
        data.extend_from_slice(&4u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&formatting_runs.to_le_bytes());
        data.extend_from_slice(&9.25f64.to_le_bytes());
        record(RRD_CHG_CELL_RECORD_TYPE, &data)
    }

    fn rrd_move() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(0x0004, 24, 1));
        for (first, last) in [(0u16, 1u16), (5, 6)] {
            data.extend_from_slice(&first.to_le_bytes());
            data.extend_from_slice(&last.to_le_bytes());
            data.extend_from_slice(&0u16.to_le_bytes());
            data.extend_from_slice(&2u16.to_le_bytes());
        }
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        record(RRD_MOVE_RECORD_TYPE, &data)
    }

    fn insert_sh() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(0x0005, 25, 2));
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&4u16.to_le_bytes());
        data.extend_from_slice(&string_field(256, "Q3 x"));
        record(RR_INSERT_SH_RECORD_TYPE, &data)
    }

    fn conflict() -> Vec<u8> {
        record(RRD_CONFLICT_RECORD_TYPE, &rrd(0x0025, 26, 0xFFFF))
    }

    fn user_view() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rrd(0x002C, 0, 0xFFFF));
        data.extend_from_slice(&[0x77; 16]);
        record(RRD_USER_VIEW_RECORD_TYPE, &data)
    }

    fn eof() -> Vec<u8> {
        record(EOF_RECORD_TYPE, &[])
    }

    /// A full revision log exercising every supported production.
    fn full_stream() -> Vec<u8> {
        let mut stream = Vec::new();
        stream.extend_from_slice(&rrd_info());
        stream.extend_from_slice(&file_lock());
        stream.extend_from_slice(&usr_excl());

        // First header: rename + bare insertion with a nested cell change and
        // an RRFormat, plus a conflict and an opaque AutoFormat revision.
        stream.extend_from_slice(&rrd_head("Alice", 0x01));
        stream.extend_from_slice(&rr_tab_id(&[1, 2, 3]));
        stream.extend_from_slice(&ren_sheet());
        stream.extend_from_slice(&ins_del(0x0000, 0));
        stream.extend_from_slice(&chg_cell(23, 1));
        stream.extend_from_slice(&record(CONTINUE_RECORD_TYPE, &[1, 2, 3]));
        stream.extend_from_slice(&record(RRD_RST_ETXP_RECORD_TYPE, &[0, 0, 4, 0]));
        stream.extend_from_slice(&record(RR_FORMAT_RECORD_TYPE, &[9, 9, 9]));
        stream.extend_from_slice(&conflict());
        stream.extend_from_slice(&record(RR_AUTO_FMT_RECORD_TYPE, &[1]));

        // Second header (no RRTabId): deletion group, move group, inserted
        // sheet, custom view, defined name, note, and query-table trash.
        stream.extend_from_slice(&rrd_head("Bob", 0x02));
        stream.extend_from_slice(&record(RRD_INS_DEL_BEGIN_RECORD_TYPE, &[]));
        stream.extend_from_slice(&ins_del(0x0002, 0));
        stream.extend_from_slice(&record(RRD_INS_DEL_END_RECORD_TYPE, &[]));
        stream.extend_from_slice(&record(RRD_MOVE_BEGIN_RECORD_TYPE, &[]));
        stream.extend_from_slice(&rrd_move());
        stream.extend_from_slice(&chg_cell(27, 0));
        stream.extend_from_slice(&record(RRD_MOVE_END_RECORD_TYPE, &[]));
        stream.extend_from_slice(&insert_sh());
        stream.extend_from_slice(&user_view());
        stream.extend_from_slice(&record(RRD_DEF_NAME_RECORD_TYPE, &[2, 2]));
        stream.extend_from_slice(&record(NOTE_RECORD_TYPE, &[3]));
        stream.extend_from_slice(&record(RRD_TQSIF_RECORD_TYPE, &[4]));
        stream.extend_from_slice(&chg_cell(28, 0));

        stream.extend_from_slice(&eof());
        stream
    }

    #[test]
    fn finds_revision_log_stream_case_insensitively_at_root() {
        let paths = vec![
            vec!["Workbook".to_string()],
            vec!["revision log".to_string()],
            vec!["_VBA_PROJECT_CUR".to_string(), "dir".to_string()],
            vec!["Nested".to_string(), "Revision Log".to_string()],
        ];
        assert_eq!(find_revision_log_stream(&paths), Some("revision log"));

        let without = vec![
            vec!["Workbook".to_string()],
            vec!["Nested".to_string(), "Revision Log".to_string()],
        ];
        assert_eq!(find_revision_log_stream(&without), None);
    }

    #[test]
    fn parses_full_revision_log_stream() {
        let log = parse_revision_log_stream(&full_stream()).unwrap();

        let info = log.info();
        assert!(info.is_shared());
        assert!(info.track_revisions());
        assert_eq!(info.revision_id(), 99);
        assert_eq!(info.history_interval_days(), 45);

        let lock = log.file_lock().unwrap();
        assert_eq!(lock.user_name(), "Bob");
        assert_eq!(lock.purpose(), records::FileLockPurpose::WritingUserInfo);
        let exclusive = log.exclusive_lock().unwrap();
        assert!(exclusive.is_exclusive());
        assert_eq!(exclusive.user_name(), "Bob");

        assert_eq!(log.headers().len(), 2);

        let first = &log.headers()[0];
        assert_eq!(first.head().user_name(), "Alice");
        assert_eq!(first.head().guid(), &[0x01; 16]);
        assert_eq!(first.sheet_ids().unwrap().sheet_ids(), &[1, 2, 3]);
        assert_eq!(first.revisions().len(), 4);
        match &first.revisions()[0] {
            Revision::RenSheet(sheet) => {
                assert_eq!(sheet.old_name(), "Budget");
                assert_eq!(sheet.new_name(), "Budget2");
            },
            other => panic!("expected RenSheet, got {other:?}"),
        }
        match &first.revisions()[1] {
            Revision::InsDel(ins_del) => {
                assert!(!ins_del.is_deletion());
                assert_eq!(
                    ins_del.record().header().revision_type(),
                    RevisionType::InsertRow
                );
                assert_eq!(ins_del.changes().len(), 2);
                match &ins_del.changes()[0] {
                    RevisionChange::CellChange(cell) => {
                        assert_eq!(cell.record().location().row(), 12);
                        assert_eq!(cell.continue_payloads(), &[vec![1, 2, 3]]);
                        assert_eq!(cell.formatting_runs(), &[vec![0, 0, 4, 0]]);
                    },
                    other => panic!("expected nested cell change, got {other:?}"),
                }
                match &ins_del.changes()[1] {
                    RevisionChange::Format(format) => {
                        assert_eq!(format.record_type(), RR_FORMAT_RECORD_TYPE);
                        assert_eq!(format.payload(), &[9, 9, 9]);
                    },
                    other => panic!("expected format change, got {other:?}"),
                }
            },
            other => panic!("expected InsDel, got {other:?}"),
        }
        assert!(matches!(first.revisions()[2], Revision::Conflict(_)));
        match &first.revisions()[3] {
            Revision::Opaque(opaque) => {
                assert_eq!(opaque.record_type(), RR_AUTO_FMT_RECORD_TYPE);
            },
            other => panic!("expected opaque AutoFormat, got {other:?}"),
        }

        let second = &log.headers()[1];
        assert_eq!(second.head().user_name(), "Bob");
        assert!(second.sheet_ids().is_none());
        assert_eq!(second.revisions().len(), 8);
        match &second.revisions()[0] {
            Revision::InsDel(ins_del) => {
                assert!(ins_del.is_deletion());
                assert_eq!(
                    ins_del.record().header().revision_type(),
                    RevisionType::DeleteRow
                );
            },
            other => panic!("expected deletion group, got {other:?}"),
        }
        match &second.revisions()[1] {
            Revision::Move(moved) => {
                assert_eq!(moved.record().source().first_row(), 0);
                assert_eq!(moved.record().destination().first_row(), 5);
                assert_eq!(moved.changes().len(), 1);
            },
            other => panic!("expected move group, got {other:?}"),
        }
        match &second.revisions()[2] {
            Revision::InsertSheet(sheet) => assert_eq!(sheet.name(), "Q3 x"),
            other => panic!("expected inserted sheet, got {other:?}"),
        }
        match &second.revisions()[3] {
            Revision::UserView(view) => {
                assert_eq!(view.header().revision_type(), RevisionType::DeleteView);
            },
            other => panic!("expected user view, got {other:?}"),
        }
        for (index, expected) in [
            (4, RRD_DEF_NAME_RECORD_TYPE),
            (5, NOTE_RECORD_TYPE),
            (6, RRD_TQSIF_RECORD_TYPE),
        ] {
            match &second.revisions()[index] {
                Revision::Opaque(opaque) => assert_eq!(opaque.record_type(), expected),
                other => panic!("expected opaque record, got {other:?}"),
            }
        }
        match &second.revisions()[7] {
            Revision::CellChange(cell) => {
                assert_eq!(cell.record().header().revision_id(), 28);
            },
            other => panic!("expected standalone cell change, got {other:?}"),
        }
    }

    #[test]
    fn parses_minimal_stream_without_optional_records() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&rrd_info());
        stream.extend_from_slice(&rrd_head("Carol", 0x03));
        stream.extend_from_slice(&eof());
        let log = parse_revision_log_stream(&stream).unwrap();
        assert!(log.file_lock().is_none());
        assert!(log.exclusive_lock().is_none());
        assert_eq!(log.headers().len(), 1);
        assert!(log.headers()[0].revisions().is_empty());
    }

    #[test]
    fn rejects_malformed_stream_structure() {
        // Empty stream.
        assert!(parse_revision_log_stream(&[]).is_err());
        // Not starting with RRDInfo.
        let mut stream = Vec::new();
        stream.extend_from_slice(&eof());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
        // Missing EOF.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Dan", 0x04));
        assert!(parse_revision_log_stream(&stream).is_err());
        // EOF with a payload.
        let mut stream = rrd_info();
        stream.extend_from_slice(&record(EOF_RECORD_TYPE, &[1]));
        assert!(parse_revision_log_stream(&stream).is_err());
        // EOF in the middle of the stream.
        let mut stream = rrd_info();
        stream.extend_from_slice(&eof());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
        // Unexpected record in a header.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Erin", 0x05));
        stream.extend_from_slice(&record(0x0201, &[0; 6])); // BLANK is not a revision
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
        // Revision outside any header.
        let mut stream = rrd_info();
        stream.extend_from_slice(&ren_sheet());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
        // Truncated record framing.
        let mut stream = rrd_info();
        stream.extend_from_slice(&eof());
        stream.truncate(stream.len() - 1);
        assert!(parse_revision_log_stream(&stream).is_err());
    }

    #[test]
    fn rejects_broken_collection_markers() {
        // RRDInsDelBegin without RRDInsDel.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Fred", 0x06));
        stream.extend_from_slice(&record(RRD_INS_DEL_BEGIN_RECORD_TYPE, &[]));
        stream.extend_from_slice(&ren_sheet());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());

        // Deletion group missing its end marker.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Fred", 0x06));
        stream.extend_from_slice(&record(RRD_INS_DEL_BEGIN_RECORD_TYPE, &[]));
        stream.extend_from_slice(&ins_del(0x0002, 0));
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());

        // Move group missing its end marker.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Fred", 0x06));
        stream.extend_from_slice(&record(RRD_MOVE_BEGIN_RECORD_TYPE, &[]));
        stream.extend_from_slice(&rrd_move());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());

        // Marker records carrying payloads.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Fred", 0x06));
        stream.extend_from_slice(&record(RRD_MOVE_BEGIN_RECORD_TYPE, &[1]));
        stream.extend_from_slice(&rrd_move());
        stream.extend_from_slice(&record(RRD_MOVE_END_RECORD_TYPE, &[]));
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
    }

    #[test]
    fn enforces_rrd_rst_etxp_count() {
        // cetxpRst says one formatting run, but none follows.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Grace", 0x07));
        stream.extend_from_slice(&chg_cell(30, 1));
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());

        // A different record where the RRDRstEtxp should be.
        let mut stream = rrd_info();
        stream.extend_from_slice(&rrd_head("Grace", 0x07));
        stream.extend_from_slice(&chg_cell(30, 1));
        stream.extend_from_slice(&ren_sheet());
        stream.extend_from_slice(&eof());
        assert!(parse_revision_log_stream(&stream).is_err());
    }
}
