//! Row heights, column widths and hidden flags for the opened-workbook edit.
//!
//! A row's layout lives in its `Row` record (`[MS-XLS]` 2.4.221: `miyRw`
//! height in twips, `fDyZero` hidden, `fUnsynced` custom height); a column's
//! lives in the `ColInfo` range that covers it (2.4.53: `coldx` width in
//! 1/256 character, `fHidden`, `fUserSet`).  Layout changes compose with cell
//! edits on the same worksheet: the row-block rebuild patches or creates
//! `Row` records after cell operations, and rewrites the `ColInfo` run that
//! follows `DefColWidth`.  Axis shifts, which renumber rows and columns, are
//! never staged together with a layout change of the same worksheet.

use super::{Error, Result};
use litchi_biff::Records;
use litchi_core::binary;
use litchi_core::edit::RefusalKind;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const ROW: u16 = 0x0208;
const COL_INFO: u16 = 0x007d;
const DEF_COL_WIDTH: u16 = 0x0055;
const DEFAULT_ROW_HEIGHT: u16 = 0x0225;
const ROW_HIDDEN: u16 = 0x0020;
const ROW_CUSTOM_HEIGHT: u16 = 0x0040;
const COLUMN_HIDDEN: u16 = 0x0001;
const COLUMN_USER_SET: u16 = 0x0002;
/// BIFF8 default row height (12.75 pt) when no `DefaultRowHeight` exists.
const FALLBACK_ROW_HEIGHT: u16 = 0x00ff;
/// The largest row height Excel accepts, 409 pt.
pub(crate) const MAX_ROW_TWIPS: u16 = 8_180;
/// The largest column width Excel accepts, 255 characters.
pub(crate) const MAX_COLUMN_WIDTH: u16 = 255 * 256;

/// One row's layout.  `present` is whether a `Row` record exists; an absent
/// row reports the worksheet default height.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct RowState {
    pub(super) present: bool,
    pub(super) height: u16,
    pub(super) custom: bool,
    pub(super) hidden: bool,
}

/// One column's layout.  `present` is whether a `ColInfo` range covers it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ColumnState {
    pub(super) present: bool,
    pub(super) width: u16,
    pub(super) user_set: bool,
    pub(super) hidden: bool,
}

/// One staged layout transition with its exact precondition.
#[derive(Debug, Clone)]
pub(super) enum LayoutChange {
    Row {
        sheet: usize,
        row: u16,
        before: RowState,
        after: RowState,
    },
    Column {
        sheet: usize,
        column: u8,
        before: ColumnState,
        after: ColumnState,
    },
}

impl LayoutChange {
    pub(super) const fn sheet(&self) -> usize {
        match self {
            Self::Row { sheet, .. } | Self::Column { sheet, .. } => *sheet,
        }
    }

    pub(super) const fn is_noop(&self) -> bool {
        match self {
            Self::Row { before, after, .. } => {
                before.present == after.present
                    && before.height == after.height
                    && before.custom == after.custom
                    && before.hidden == after.hidden
            },
            Self::Column { before, after, .. } => {
                before.present == after.present
                    && before.width == after.width
                    && before.user_set == after.user_set
                    && before.hidden == after.hidden
            },
        }
    }
}

struct Frame {
    kind: u16,
    start: usize,
    end: usize,
}

fn frames_until_eof(bytes: &[u8]) -> Result<Vec<Frame>> {
    let mut frames = Vec::new();
    for item in Records::new(bytes, litchi_biff::Limits::default())? {
        let record = item?;
        let start = record.offset();
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        let kind = record.kind().get();
        frames
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing worksheet layout records"))?;
        frames.push(Frame { kind, start, end });
        if kind == EOF {
            break;
        }
    }
    if frames.first().map(|frame| frame.kind) != Some(BOF) {
        return Err(Error::InvalidData("worksheet substream has no BOF".into()));
    }
    Ok(frames)
}

/// The worksheet's default row height in twips.
fn default_row_height(stream: &[u8], frames: &[Frame]) -> Result<u16> {
    match frames.iter().find(|frame| frame.kind == DEFAULT_ROW_HEIGHT) {
        Some(frame) if frame.end - frame.start >= 8 => {
            Ok(binary::read_u16_le_at(stream, frame.start + 6)? & 0x7fff)
        },
        _ => Ok(FALLBACK_ROW_HEIGHT),
    }
}

/// Decode a `Row` record's layout state (the record includes its header).
pub(super) fn row_record_state(record: &[u8]) -> Result<RowState> {
    if record.len() < 20 {
        return Err(Error::InvalidRecord {
            record_type: ROW,
            message: "Row record is truncated".into(),
        });
    }
    let height = binary::read_u16_le_at(record, 10)? & 0x7fff;
    let flags = binary::read_u16_le_at(record, 16)?;
    Ok(RowState {
        present: true,
        height,
        custom: flags & ROW_CUSTOM_HEIGHT != 0,
        hidden: flags & ROW_HIDDEN != 0,
    })
}

/// Read one row's state from a worksheet substream.
pub(super) fn row_state(stream: &[u8], row: u16) -> Result<RowState> {
    let frames = frames_until_eof(stream)?;
    for frame in frames.iter().filter(|frame| frame.kind == ROW) {
        let record = &stream[frame.start..frame.end];
        if binary::read_u16_le_at(record, 4)? == row {
            return row_record_state(record);
        }
    }
    Ok(RowState {
        present: false,
        height: default_row_height(stream, &frames)?,
        custom: false,
        hidden: false,
    })
}

/// Read one column's state from a worksheet substream.
pub(super) fn column_state(stream: &[u8], column: u8) -> Result<ColumnState> {
    let frames = frames_until_eof(stream)?;
    for frame in frames.iter().filter(|frame| frame.kind == COL_INFO) {
        let info = ColInfo::decode(&stream[frame.start..frame.end])?;
        if (info.first..=info.last).contains(&u16::from(column)) {
            return Ok(info.state());
        }
    }
    Ok(ColumnState {
        present: false,
        width: 0,
        user_set: false,
        hidden: false,
    })
}

/// Patch an existing `Row` record to `after`.
pub(super) fn patch_row_record(record: &mut [u8], after: RowState) -> Result<()> {
    if record.len() < 20 {
        return Err(Error::InvalidRecord {
            record_type: ROW,
            message: "Row record is truncated".into(),
        });
    }
    let height = (binary::read_u16_le_at(record, 10)? & 0x8000) | (after.height & 0x7fff);
    record[10..12].copy_from_slice(&height.to_le_bytes());
    let mut flags = binary::read_u16_le_at(record, 16)? & !(ROW_HIDDEN | ROW_CUSTOM_HEIGHT);
    if after.hidden {
        flags |= ROW_HIDDEN;
    }
    if after.custom {
        flags |= ROW_CUSTOM_HEIGHT;
    }
    record[16..18].copy_from_slice(&flags.to_le_bytes());
    Ok(())
}

/// A new empty `Row` record carrying `after`.
pub(super) fn new_row_record(row: u16, after: RowState) -> Result<Vec<u8>> {
    let mut record = Vec::new();
    crate::writer::biff::write_row(
        &mut record,
        u32::from(row),
        0,
        0,
        after.height,
        after.hidden,
    )?;
    patch_row_record(&mut record, after)?;
    Ok(record)
}

/// One decoded `ColInfo` range.  The undecoded tail (`ixfe`, outline and
/// reserved bits) is retained exactly when a range is split.
#[derive(Debug, Clone, PartialEq, Eq)]
struct ColInfo {
    first: u16,
    last: u16,
    width: u16,
    xf: u16,
    flags: u16,
    reserved: Vec<u8>,
}

impl ColInfo {
    fn decode(record: &[u8]) -> Result<Self> {
        if record.len() < 16 {
            return Err(Error::InvalidRecord {
                record_type: COL_INFO,
                message: "ColInfo record is truncated".into(),
            });
        }
        let first = binary::read_u16_le_at(record, 4)?;
        let last = binary::read_u16_le_at(record, 6)?;
        if first > last {
            return Err(Error::InvalidRecord {
                record_type: COL_INFO,
                message: "ColInfo range is reversed".into(),
            });
        }
        Ok(Self {
            first,
            last,
            width: binary::read_u16_le_at(record, 8)?,
            xf: binary::read_u16_le_at(record, 10)?,
            flags: binary::read_u16_le_at(record, 12)?,
            reserved: record[14..].to_vec(),
        })
    }

    fn state(&self) -> ColumnState {
        ColumnState {
            present: true,
            width: self.width,
            user_set: self.flags & COLUMN_USER_SET != 0,
            hidden: self.flags & COLUMN_HIDDEN != 0,
        }
    }

    fn with_state(&self, first: u16, last: u16, state: ColumnState) -> Self {
        let mut flags = self.flags & !(COLUMN_HIDDEN | COLUMN_USER_SET);
        if state.hidden {
            flags |= COLUMN_HIDDEN;
        }
        if state.user_set {
            flags |= COLUMN_USER_SET;
        }
        Self {
            first,
            last,
            width: state.width,
            xf: self.xf,
            flags,
            reserved: self.reserved.clone(),
        }
    }

    fn encode(&self, output: &mut Vec<u8>) -> Result<()> {
        let length = u16::try_from(10 + self.reserved.len())
            .map_err(|_error| Error::InvalidData("ColInfo payload exceeds u16".into()))?;
        output.extend_from_slice(&COL_INFO.to_le_bytes());
        output.extend_from_slice(&length.to_le_bytes());
        for value in [self.first, self.last, self.width, self.xf, self.flags] {
            output.extend_from_slice(&value.to_le_bytes());
        }
        output.extend_from_slice(&self.reserved);
        Ok(())
    }
}

/// Apply column layout changes to the records between `Index` and the row
/// table.  The `ColInfo` run must directly follow `DefColWidth`; ranges are
/// split around a changed column and re-encoded in ascending order.
pub(super) fn apply_columns(between: &[u8], changes: &[&LayoutChange]) -> Result<Vec<u8>> {
    let mut columns = Vec::new();
    for change in changes {
        if let LayoutChange::Column {
            column,
            before,
            after,
            ..
        } = change
        {
            columns.push((*column, *before, *after));
        }
    }
    if columns.is_empty() {
        return Ok(between.to_vec());
    }
    let mut frames = Vec::new();
    for item in Records::new(between, litchi_biff::Limits::default())? {
        let record = item?;
        let start = record.offset();
        frames.push(Frame {
            kind: record.kind().get(),
            start,
            end: start + 4 + record.payload().len(),
        });
    }
    let def = frames
        .iter()
        .position(|frame| frame.kind == DEF_COL_WIDTH)
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "worksheet has no DefColWidth record",
            )
        })?;
    let run_end = frames[def + 1..]
        .iter()
        .position(|frame| frame.kind != COL_INFO)
        .map_or(frames.len(), |offset| def + 1 + offset);
    if frames[run_end..].iter().any(|frame| frame.kind == COL_INFO)
        || frames[..def].iter().any(|frame| frame.kind == COL_INFO)
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "ColInfo records outside the DefColWidth run are not rewritten",
        ));
    }
    let mut infos = frames[def + 1..run_end]
        .iter()
        .map(|frame| ColInfo::decode(&between[frame.start..frame.end]))
        .collect::<Result<Vec<_>>>()?;
    for (column, before, after) in columns {
        let column = u16::from(column);
        let covering = infos
            .iter()
            .position(|info| (info.first..=info.last).contains(&column));
        let observed = covering.map_or(
            ColumnState {
                present: false,
                width: 0,
                user_set: false,
                hidden: false,
            },
            |index| infos[index].state(),
        );
        if observed != before {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "column layout precondition is stale",
            });
        }
        let mut replacement = Vec::new();
        match covering {
            Some(index) => {
                let info = infos.remove(index);
                if info.first < column {
                    replacement.push(info.with_state(info.first, column - 1, info.state()));
                }
                if after.present {
                    replacement.push(info.with_state(column, column, after));
                }
                if column < info.last {
                    replacement.push(info.with_state(column + 1, info.last, info.state()));
                }
            },
            None if after.present => {
                let template = ColInfo {
                    first: column,
                    last: column,
                    width: 0,
                    xf: 0x000f,
                    flags: 0,
                    reserved: vec![0, 0],
                };
                replacement.push(template.with_state(column, column, after));
            },
            None => {},
        }
        infos.extend(replacement);
        infos.sort_by_key(|info| info.first);
    }
    let mut output = Vec::new();
    output
        .try_reserve(between.len() + 64)
        .map_err(|_error| Error::Allocation("rewriting worksheet ColInfo records"))?;
    output.extend_from_slice(&between[..frames[def].end]);
    for info in &infos {
        info.encode(&mut output)?;
    }
    let tail_start = frames
        .get(run_end)
        .map_or(between.len(), |frame| frame.start);
    output.extend_from_slice(&between[tail_start..]);
    Ok(output)
}
