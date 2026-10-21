//! Complete-catalog worksheet tab structure for the opened-workbook edit.
//!
//! A tab-structure change rewrites the workbook globals' tab directory and
//! relocates sheet substreams.  Every workbook-global record that addresses a
//! tab by catalog position is remapped in the same closure:
//!
//! * `BoundSheet8` order and absolute substream positions (`[MS-XLS]` 2.4.28);
//! * `Window1` active tab, first visible tab and selected-tab count (2.4.346);
//! * the `RRTabId` revision identifiers (2.4.236);
//! * the self-referencing `SupBook` tab count (2.4.271);
//! * internal `ExternSheet` XTI ranges (2.4.106), so 3-D formula references in
//!   any substream keep naming the same tabs;
//! * sheet-scoped `Lbl` defined names (2.4.150);
//! * every worksheet `Index` absolute offset (2.4.144) and the `ExtSST`
//!   bucket offsets (2.4.107) that move with the rewritten globals.
//!
//! Changes that would leave a dependency dangling are refused before any byte
//! is published: removing a tab that an XTI range, a sheet-scoped name, a
//! pivot data connection or a drawing still uses; inserting inside a 3-D XTI
//! range; or reordering so that a 3-D range would become noncontiguous.

use super::{Error, Result};
use litchi_biff::Records;
use litchi_core::binary;
use litchi_core::edit::RefusalKind;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const BOUND_SHEET: u16 = 0x0085;
const WINDOW1: u16 = 0x003d;
const WINDOW2: u16 = 0x023e;
const RR_TAB_ID: u16 = 0x013d;
const SUP_BOOK: u16 = 0x01ae;
const EXTERN_SHEET: u16 = 0x0017;
const LBL: u16 = 0x0018;
const CONTINUE: u16 = 0x003c;
const SST: u16 = 0x00fc;
const EXT_SST: u16 = 0x00ff;
const INDEX: u16 = 0x020b;
const MSO_DRAWING: u16 = 0x00ec;
const OBJ: u16 = 0x005d;
const SX_VIEW: u16 = 0x00b0;
const DCON_REF: u16 = 0x0051;
const DCON_NAME: u16 = 0x0052;
const DCON_BIN: u16 = 0x01b5;
const WORKSHEET_SUBSTREAM: u16 = 0x0010;
const WINDOW2_SELECTED: u16 = 0x0200;
const WINDOW2_PAGED: u16 = 0x0400;

/// One staged complete-catalog tab-structure change.
///
/// The candidate Workbook stream is computed and verified while staging, so
/// a refused dependency never reaches commit and commit never repeats the
/// rewrite.
#[derive(Debug, Clone)]
pub(super) struct TabChange {
    pub(super) kind: TabChangeKind,
    /// Catalog names in source order, retained for semantic preconditions.
    pub(super) source_names: Vec<String>,
    /// Catalog names in candidate order, verified after publication.
    pub(super) target_names: Vec<String>,
    pub(super) workbook: Vec<u8>,
}

/// The semantic shape of a [`TabChange`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum TabChangeKind {
    /// New catalog position `i` holds source tab `order[i]`.
    Reorder { order: Vec<usize> },
    /// An empty worksheet inserted at `position`.
    Insert { position: usize, name: String },
    /// The ordinary worksheet at source `position` removed.
    Remove {
        position: usize,
        name: String,
        /// Content fingerprint of the removed tab's `BoundSheet8` body and
        /// substream; the exact artifact patch is its only restoring inverse.
        fingerprint: String,
    },
}

/// One `BoundSheet8` catalog entry with its substream range.
#[derive(Debug, Clone)]
pub(super) struct Tab {
    pub(super) name: String,
    /// `hsState`: 0 visible, 1 hidden, 2 very hidden.
    pub(super) state: u8,
    /// `dt`: 0 worksheet or dialog, 1 macro, 2 chart, 6 VBA module.
    pub(super) kind: u8,
    body_start: usize,
    body_end: usize,
    pub(super) start: usize,
    pub(super) end: usize,
}

struct RawRecord {
    kind: u16,
    start: usize,
    end: usize,
}

impl RawRecord {
    fn payload<'a>(&self, bytes: &'a [u8]) -> Result<&'a [u8]> {
        bytes
            .get(self.start + 4..self.end)
            .ok_or_else(|| Error::InvalidData("BIFF record payload is truncated".into()))
    }
}

fn raw_records(bytes: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    for item in Records::new(bytes, litchi_biff::Limits::default())? {
        let record = item?;
        let start = record.offset();
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing XLS tab records"))?;
        records.push(RawRecord {
            kind: record.kind().get(),
            start,
            end,
        });
    }
    Ok(records)
}

/// Parse the complete `BoundSheet8` catalog in tab order.
pub(super) fn catalog(workbook: &[u8]) -> Result<Vec<Tab>> {
    parse(workbook).map(|(_, tabs)| tabs)
}

fn parse(workbook: &[u8]) -> Result<(usize, Vec<Tab>)> {
    let mut tabs = Vec::new();
    let mut globals_eof = None;
    for record in raw_records_until_eof(workbook)? {
        if record.kind == EOF {
            globals_eof = Some(record.end);
            break;
        }
        if record.kind != BOUND_SHEET {
            continue;
        }
        let payload = record.payload(workbook)?;
        if payload.len() < 8 {
            return Err(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: "BoundSheet8 is truncated".into(),
            });
        }
        let position = usize::try_from(binary::read_u32_le_at(payload, 0)?)
            .map_err(|_error| Error::InvalidData("BoundSheet8 position exceeds usize".into()))?;
        tabs.try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing XLS tabs"))?;
        tabs.push(Tab {
            name: decode_name(payload)?,
            state: payload[4] & 0x03,
            kind: payload[5],
            body_start: record.start + 4,
            body_end: record.end,
            start: position,
            end: 0,
        });
    }
    let globals_eof = globals_eof.ok_or(Error::Eof("Workbook globals EOF"))?;
    if tabs.is_empty() {
        return Err(Error::InvalidRecord {
            record_type: BOUND_SHEET,
            message: "workbook globals contain no BoundSheet8 records".into(),
        });
    }
    let mut starts: Vec<usize> = tabs.iter().map(|tab| tab.start).collect();
    starts.sort_unstable();
    if starts.windows(2).any(|pair| pair[0] == pair[1])
        || starts[0] < globals_eof
        || starts[starts.len() - 1] >= workbook.len()
    {
        return Err(Error::InvalidRecord {
            record_type: BOUND_SHEET,
            message: "BoundSheet8 substream positions are invalid".into(),
        });
    }
    for tab in &mut tabs {
        let slot = starts.partition_point(|start| *start <= tab.start);
        tab.end = starts.get(slot).copied().unwrap_or(workbook.len());
        let kind = workbook
            .get(tab.start..tab.start + 2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]));
        if kind != Some(BOF) {
            return Err(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: "BoundSheet8 does not address a BOF record".into(),
            });
        }
    }
    Ok((starts[0], tabs))
}

/// Globals records up to and including the first EOF.
fn raw_records_until_eof(bytes: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    for item in Records::new(bytes, litchi_biff::Limits::default())? {
        let record = item?;
        let start = record.offset();
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        let kind = record.kind().get();
        records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing XLS globals"))?;
        records.push(RawRecord { kind, start, end });
        if kind == EOF {
            break;
        }
    }
    Ok(records)
}

fn decode_name(payload: &[u8]) -> Result<String> {
    let units = usize::from(payload[6]);
    let wide = payload[7] & 1 != 0;
    if wide {
        let bytes = payload
            .get(8..8 + units * 2)
            .ok_or_else(|| Error::InvalidData("wide BoundSheet8 name is truncated".into()))?;
        let utf16: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16(&utf16)
            .map_err(|_error| Error::InvalidData("BoundSheet8 name is not UTF-16".into()))
    } else {
        let bytes = payload
            .get(8..8 + units)
            .ok_or_else(|| Error::InvalidData("BoundSheet8 name is truncated".into()))?;
        Ok(bytes.iter().map(|byte| char::from(*byte)).collect())
    }
}

/// Encode a fresh worksheet `BoundSheet8` body with a zero position.
fn bound_sheet_body(name: &str) -> Result<Vec<u8>> {
    let units: Vec<u16> = name.encode_utf16().collect();
    let count = u8::try_from(units.len())
        .map_err(|_error| Error::InvalidData("worksheet name length exceeds u8".into()))?;
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let mut body = Vec::new();
    body.try_reserve_exact(8 + units.len() * 2)
        .map_err(|_error| Error::Allocation("encoding a BoundSheet8 record"))?;
    body.extend_from_slice(&[0, 0, 0, 0, 0, 0, count, u8::from(!compressed)]);
    for unit in units {
        if compressed {
            body.push(unit.to_le_bytes()[0]);
        } else {
            body.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(body)
}

fn record_bytes(kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .ok()
        .filter(|length| *length <= 8_224)
        .ok_or_else(|| Error::InvalidData("BIFF8 record payload exceeds 8,224 bytes".into()))?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(4 + payload.len())
        .map_err(|_error| Error::Allocation("encoding an XLS tab record"))?;
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(payload);
    Ok(bytes)
}

/// One tab of the candidate catalog.
struct Planned<'a> {
    /// Source catalog position, or `None` for an inserted worksheet.
    source: Option<usize>,
    body: Vec<u8>,
    stream: &'a [u8],
    /// Absolute offset the stream's `Index` fields were written against.
    stream_origin: usize,
    state: u8,
}

/// The source-independent template of an inserted empty worksheet.
pub(super) struct Template {
    pub(super) stream: Vec<u8>,
    pub(super) origin: usize,
}

/// Build the candidate Workbook stream for one tab-structure change.
pub(super) fn rewrite(
    workbook: &[u8],
    kind: &TabChangeKind,
    template: Option<&Template>,
) -> Result<Vec<u8>> {
    let (globals_end, tabs) = parse(workbook)?;
    let old_count = tabs.len();
    let mut plan: Vec<Planned<'_>> = Vec::new();
    plan.try_reserve_exact(old_count + 1)
        .map_err(|_error| Error::Allocation("planning XLS tab structure"))?;
    let source_planned = |index: usize| -> Result<Planned<'_>> {
        let tab = tabs
            .get(index)
            .ok_or_else(|| Error::InvalidData("tab permutation index is invalid".into()))?;
        let mut body = Vec::new();
        body.try_reserve_exact(tab.body_end - tab.body_start)
            .map_err(|_error| Error::Allocation("retaining a BoundSheet8 body"))?;
        body.extend_from_slice(&workbook[tab.body_start..tab.body_end]);
        Ok(Planned {
            source: Some(index),
            body,
            stream: &workbook[tab.start..tab.end],
            stream_origin: tab.start,
            state: tab.state,
        })
    };
    let mut insert_index = None;
    match kind {
        TabChangeKind::Reorder { order } => {
            if order.len() != old_count {
                return Err(Error::InvalidData(
                    "tab reorder must name every catalog tab exactly once".into(),
                ));
            }
            let mut seen = vec![false; old_count];
            for &index in order {
                if index >= old_count || std::mem::replace(&mut seen[index], true) {
                    return Err(Error::InvalidData(
                        "tab reorder contains an invalid or repeated tab".into(),
                    ));
                }
                plan.push(source_planned(index)?);
            }
        },
        TabChangeKind::Remove { position, .. } => {
            if *position >= old_count {
                return Err(Error::WorksheetNotFound("removed tab position".into()));
            }
            if old_count == 1 {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "a workbook must retain at least one tab",
                ));
            }
            for index in (0..old_count).filter(|index| index != position) {
                plan.push(source_planned(index)?);
            }
        },
        TabChangeKind::Insert { position, name } => {
            let template = template.ok_or_else(|| {
                Error::InvalidData("tab insertion requires an empty worksheet template".into())
            })?;
            if *position > old_count {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "worksheet insertion position is outside the tab catalog",
                ));
            }
            for index in 0..old_count {
                if index == *position {
                    plan.push(Planned {
                        source: None,
                        body: bound_sheet_body(name)?,
                        stream: &template.stream,
                        stream_origin: template.origin,
                        state: 0,
                    });
                }
                plan.push(source_planned(index)?);
            }
            if *position == old_count {
                plan.push(Planned {
                    source: None,
                    body: bound_sheet_body(name)?,
                    stream: &template.stream,
                    stream_origin: template.origin,
                    state: 0,
                });
            }
            insert_index = Some(*position);
        },
    }
    if plan.is_empty() || plan.len() > usize::from(u16::MAX) {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "tab catalog size is outside BIFF8 bounds",
        ));
    }
    let mut old_to_new = vec![None; old_count];
    for (new, planned) in plan.iter().enumerate() {
        if let Some(old) = planned.source {
            old_to_new[old] = Some(new);
        }
    }

    // A removed tab must not be the only visible tab.
    if !plan.iter().any(|planned| planned.state == 0) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "a workbook must retain a visible tab",
        ));
    }

    let globals = &workbook[..globals_end];
    let records = raw_records(globals)?;
    let window_target = window_target(globals, &records, &old_to_new, &plan)?;
    let mut output = Vec::new();
    output
        .try_reserve(workbook.len() + 512)
        .map_err(|_error| Error::Allocation("rewriting the XLS tab directory"))?;
    let mut bound_positions = Vec::new();
    bound_positions
        .try_reserve_exact(plan.len())
        .map_err(|_error| Error::Allocation("rewriting the XLS tab directory"))?;
    let mut bounds_written = false;
    let mut sst_shift: Option<(usize, usize)> = None;
    let mut ext_sst_output = None;
    let internal_books = internal_sup_books(globals, &records)?;
    for (ordinal, record) in records.iter().enumerate() {
        let original = &globals[record.start..record.end];
        let payload = record.payload(globals)?;
        let followed_by_continue = records
            .get(ordinal + 1)
            .is_some_and(|next| next.kind == CONTINUE);
        match record.kind {
            BOUND_SHEET => {
                if bounds_written {
                    continue;
                }
                bounds_written = true;
                for planned in &plan {
                    bound_positions.push(output.len() + 4);
                    let mut body = planned.body.clone();
                    body[..4].fill(0);
                    output.extend_from_slice(&record_bytes(BOUND_SHEET, &body)?);
                }
            },
            WINDOW1 => {
                let rewritten = remap_window1(payload, window_target, plan.len())?;
                output.extend_from_slice(&record_bytes(WINDOW1, &rewritten)?);
            },
            RR_TAB_ID => {
                if followed_by_continue {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "continued RRTabId records are not remapped",
                    ));
                }
                let rewritten = remap_rr_tab_ids(payload, old_count, &plan)?;
                output.extend_from_slice(&record_bytes(RR_TAB_ID, &rewritten)?);
            },
            SUP_BOOK if payload.len() == 4 && binary::read_u16_le_at(payload, 2)? == 0x0401 => {
                let mut rewritten = payload.to_vec();
                rewritten[..2].copy_from_slice(
                    &u16::try_from(plan.len())
                        .map_err(|_error| Error::InvalidData("tab count exceeds u16".into()))?
                        .to_le_bytes(),
                );
                output.extend_from_slice(&record_bytes(SUP_BOOK, &rewritten)?);
            },
            EXTERN_SHEET => {
                if followed_by_continue {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "continued ExternSheet records are not remapped",
                    ));
                }
                let rewritten =
                    remap_extern_sheet(payload, &internal_books, &old_to_new, insert_index)?;
                output.extend_from_slice(&record_bytes(EXTERN_SHEET, &rewritten)?);
            },
            LBL => {
                let mut rewritten = original.to_vec();
                remap_lbl(&mut rewritten[4..], &old_to_new)?;
                output.extend_from_slice(&rewritten);
            },
            SST => {
                sst_shift = Some((record.start, output.len()));
                output.extend_from_slice(original);
            },
            EXT_SST => {
                if followed_by_continue {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "continued ExtSST records are not remapped",
                    ));
                }
                ext_sst_output = Some(output.len());
                output.extend_from_slice(original);
            },
            _ => output.extend_from_slice(original),
        }
    }
    if !bounds_written {
        return Err(Error::InvalidRecord {
            record_type: BOUND_SHEET,
            message: "workbook globals contain no BoundSheet8 records".into(),
        });
    }
    if let (Some(at), Some((old, new))) = (ext_sst_output, sst_shift) {
        shift_ext_sst(&mut output, at, old, new)?;
    } else if ext_sst_output.is_some() {
        return Err(Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "ExtSST has no SST owner",
        ));
    }

    for (slot, planned) in plan.iter().enumerate() {
        let start = output.len();
        output.extend_from_slice(planned.stream);
        let position = u32::try_from(start)
            .map_err(|_error| Error::InvalidData("BoundSheet8 position exceeds u32".into()))?;
        output[bound_positions[slot]..bound_positions[slot] + 4]
            .copy_from_slice(&position.to_le_bytes());
        let delta = i64::try_from(start)
            .ok()
            .zip(i64::try_from(planned.stream_origin).ok())
            .map(|(new, old)| new - old)
            .ok_or_else(|| Error::InvalidData("substream relocation overflows".into()))?;
        adjust_index_offsets(&mut output[start..], delta)?;
        if Some(slot) == window_target.select_new_active {
            select_window2(&mut output[start..])?;
        }
    }
    Ok(output)
}

/// Window1 remapping decided before the globals are rewritten.
#[derive(Clone, Copy)]
struct WindowTarget {
    active: usize,
    first: usize,
    selected: u16,
    /// A candidate tab that becomes active because the old active tab was
    /// removed; its `Window2` gains the selected/paged flags.
    select_new_active: Option<usize>,
}

fn window_target(
    globals: &[u8],
    records: &[RawRecord],
    old_to_new: &[Option<usize>],
    plan: &[Planned<'_>],
) -> Result<WindowTarget> {
    let Some(window) = records.iter().find(|record| record.kind == WINDOW1) else {
        return Ok(WindowTarget {
            active: 0,
            first: 0,
            selected: 1,
            select_new_active: None,
        });
    };
    let payload = window.payload(globals)?;
    if payload.len() != 18 {
        return Err(Error::InvalidRecord {
            record_type: WINDOW1,
            message: "Window1 must contain 18 bytes".into(),
        });
    }
    let first_visible = plan
        .iter()
        .position(|planned| planned.state == 0)
        .unwrap_or(0);
    let old_active = usize::from(binary::read_u16_le_at(payload, 10)?);
    let old_first = usize::from(binary::read_u16_le_at(payload, 12)?);
    let old_selected = binary::read_u16_le_at(payload, 14)?;
    let (active, select_new_active, selected) = match old_to_new.get(old_active).copied().flatten()
    {
        Some(active) => (active, None, old_selected),
        None => (first_visible, Some(first_visible), 1),
    };
    let first = old_to_new
        .get(old_first)
        .copied()
        .flatten()
        .unwrap_or(0)
        .min(active);
    let selected = selected.clamp(
        1,
        u16::try_from(plan.len()).map_err(|_error| Error::InvalidData("tab count".into()))?,
    );
    Ok(WindowTarget {
        active,
        first,
        selected,
        select_new_active,
    })
}

fn remap_window1(payload: &[u8], target: WindowTarget, count: usize) -> Result<Vec<u8>> {
    let mut output = payload.to_vec();
    let to_u16 = |value: usize| {
        u16::try_from(value).map_err(|_error| Error::InvalidData("Window1 tab exceeds u16".into()))
    };
    if target.active >= count {
        return Err(Error::InvalidData(
            "Window1 active tab is out of range".into(),
        ));
    }
    output[10..12].copy_from_slice(&to_u16(target.active)?.to_le_bytes());
    output[12..14].copy_from_slice(&to_u16(target.first)?.to_le_bytes());
    output[14..16].copy_from_slice(&target.selected.to_le_bytes());
    Ok(output)
}

fn remap_rr_tab_ids(payload: &[u8], old_count: usize, plan: &[Planned<'_>]) -> Result<Vec<u8>> {
    if payload.len() != old_count * 2 {
        return Err(Error::InvalidRecord {
            record_type: RR_TAB_ID,
            message: "RRTabId count does not match the tab catalog".into(),
        });
    }
    let old: Vec<u16> = payload
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    let next = old
        .iter()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| Error::InvalidData("RRTabId identifier overflow".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(plan.len() * 2)
        .map_err(|_error| Error::Allocation("rewriting RRTabId"))?;
    for planned in plan {
        let value = match planned.source {
            Some(index) => old[index],
            None => next,
        };
        output.extend_from_slice(&value.to_le_bytes());
    }
    Ok(output)
}

fn internal_sup_books(globals: &[u8], records: &[RawRecord]) -> Result<Vec<u16>> {
    let mut internal = Vec::new();
    let mut ordinal = 0_u16;
    for record in records.iter().filter(|record| record.kind == SUP_BOOK) {
        let payload = record.payload(globals)?;
        if payload.len() == 4 && binary::read_u16_le_at(payload, 2)? == 0x0401 {
            internal.push(ordinal);
        }
        ordinal = ordinal
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("SupBook count overflow".into()))?;
    }
    Ok(internal)
}

fn remap_extern_sheet(
    payload: &[u8],
    internal: &[u16],
    old_to_new: &[Option<usize>],
    insert_index: Option<usize>,
) -> Result<Vec<u8>> {
    if payload.len() < 2 {
        return Err(Error::InvalidRecord {
            record_type: EXTERN_SHEET,
            message: "ExternSheet is truncated".into(),
        });
    }
    let count = usize::from(binary::read_u16_le_at(payload, 0)?);
    if payload.len() != 2 + count * 6 {
        return Err(Error::InvalidRecord {
            record_type: EXTERN_SHEET,
            message: "ExternSheet count does not match its payload".into(),
        });
    }
    let mut output = payload.to_vec();
    for index in 0..count {
        let offset = 2 + index * 6;
        if !internal.contains(&binary::read_u16_le_at(payload, offset)?) {
            continue;
        }
        let first = binary::read_u16_le_at(payload, offset + 2)?;
        let last = binary::read_u16_le_at(payload, offset + 4)?;
        // 0xFFFE addresses the workbook and 0xFFFF a deleted reference.
        if matches!(first, 0xfffe | 0xffff) || matches!(last, 0xfffe | 0xffff) {
            continue;
        }
        let first = usize::from(first);
        let last = usize::from(last);
        if first > last || last >= old_to_new.len() {
            return Err(Error::InvalidRecord {
                record_type: EXTERN_SHEET,
                message: "internal ExternSheet range is invalid".into(),
            });
        }
        if insert_index.is_some_and(|insert| first < insert && insert <= last) {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a worksheet cannot be inserted inside an existing 3-D reference range",
            ));
        }
        let mut minimum = usize::MAX;
        let mut maximum = 0;
        for mapped in &old_to_new[first..=last] {
            let new = mapped.ok_or_else(|| {
                Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "a formula or name still references the removed tab",
                )
            })?;
            minimum = minimum.min(new);
            maximum = maximum.max(new);
        }
        if maximum - minimum != last - first {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "the tab order would make a 3-D reference range noncontiguous",
            ));
        }
        let to_u16 = |value: usize| {
            u16::try_from(value)
                .map_err(|_error| Error::InvalidData("ExternSheet tab exceeds u16".into()))
        };
        output[offset + 2..offset + 4].copy_from_slice(&to_u16(minimum)?.to_le_bytes());
        output[offset + 4..offset + 6].copy_from_slice(&to_u16(maximum)?.to_le_bytes());
    }
    Ok(output)
}

fn remap_lbl(payload: &mut [u8], old_to_new: &[Option<usize>]) -> Result<()> {
    if payload.len() < 10 {
        return Err(Error::InvalidRecord {
            record_type: LBL,
            message: "Lbl is truncated".into(),
        });
    }
    let scope = usize::from(binary::read_u16_le_at(payload, 8)?);
    if scope == 0 {
        return Ok(());
    }
    let new = old_to_new
        .get(scope - 1)
        .copied()
        .flatten()
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "the removed tab owns a sheet-scoped name",
            )
        })?;
    let scope = u16::try_from(new + 1)
        .map_err(|_error| Error::InvalidData("Lbl tab scope exceeds u16".into()))?;
    payload[8..10].copy_from_slice(&scope.to_le_bytes());
    Ok(())
}

/// Shift `ExtSST` absolute bucket offsets by the SST record's relocation.
fn shift_ext_sst(output: &mut [u8], at: usize, old_sst: usize, new_sst: usize) -> Result<()> {
    if old_sst == new_sst {
        return Ok(());
    }
    let length = usize::from(binary::read_u16_le_at(output, at + 2)?);
    let payload_start = at + 4;
    if length < 2 || (length - 2) % 8 != 0 {
        return Err(Error::InvalidRecord {
            record_type: EXT_SST,
            message: "ExtSST payload is malformed".into(),
        });
    }
    let delta = i64::try_from(new_sst)
        .ok()
        .zip(i64::try_from(old_sst).ok())
        .map(|(new, old)| new - old)
        .ok_or_else(|| Error::InvalidData("SST relocation overflows".into()))?;
    for bucket in 0..(length - 2) / 8 {
        let offset = payload_start + 2 + bucket * 8;
        let before = binary::read_u32_le_at(output, offset)?;
        let after = i64::from(before)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "ExtSST offset relocation overflows",
                )
            })?;
        output[offset..offset + 4].copy_from_slice(&after.to_le_bytes());
    }
    Ok(())
}

/// Shift a relocated worksheet's `Index` absolute offsets (`ibXF` and every
/// nonzero DBCELL position).  Non-worksheet substreams have no `Index`.
fn adjust_index_offsets(stream: &mut [u8], delta: i64) -> Result<()> {
    if delta == 0 {
        return Ok(());
    }
    let records = raw_records_until_eof(stream)?;
    let Some(index) = records.iter().find(|record| record.kind == INDEX) else {
        return Ok(());
    };
    let payload_start = index.start + 4;
    let length = index.end - payload_start;
    if length < 16 || (length - 16) % 4 != 0 {
        return Err(Error::InvalidRecord {
            record_type: INDEX,
            message: "Index payload is malformed".into(),
        });
    }
    for relative in (12..length).step_by(4) {
        let offset = payload_start + relative;
        let before = binary::read_u32_le_at(stream, offset)?;
        if before == 0 {
            continue;
        }
        let after = i64::from(before)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "Index offset relocation overflows",
                )
            })?;
        stream[offset..offset + 4].copy_from_slice(&after.to_le_bytes());
    }
    Ok(())
}

fn window2_flags(stream: &mut [u8], set: u16, clear: u16) -> Result<()> {
    let records = raw_records_until_eof(stream)?;
    if let Some(window) = records.iter().find(|record| record.kind == WINDOW2) {
        let offset = window.start + 4;
        if window.end - offset < 2 {
            return Err(Error::InvalidRecord {
                record_type: WINDOW2,
                message: "Window2 is truncated".into(),
            });
        }
        let flags = (binary::read_u16_le_at(stream, offset)? | set) & !clear;
        stream[offset..offset + 2].copy_from_slice(&flags.to_le_bytes());
    }
    Ok(())
}

fn select_window2(stream: &mut [u8]) -> Result<()> {
    window2_flags(stream, WINDOW2_SELECTED | WINDOW2_PAGED, 0)
}

/// Refuse removing a tab whose dependencies this rewrite cannot close.
pub(super) fn certify_removal(workbook: &[u8], catalog: &[Tab], tab: &Tab) -> Result<()> {
    if tab.kind != 0 {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "only ordinary worksheet tabs can be removed",
        ));
    }
    let stream = &workbook[tab.start..tab.end];
    let records = raw_records_until_eof(stream)?;
    let substream = records
        .first()
        .filter(|record| record.kind == BOF)
        .map(|record| record.payload(stream))
        .transpose()?
        .map(|payload| binary::read_u16_le_at(payload, 2))
        .transpose()?;
    if substream != Some(WORKSHEET_SUBSTREAM) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "only ordinary worksheet tabs can be removed",
        ));
    }
    if records
        .iter()
        .any(|record| matches!(record.kind, MSO_DRAWING | OBJ | SX_VIEW))
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "worksheet removal cannot close drawing or pivot record",
        ));
    }
    // Pivot data connections name their source sheet by text anywhere in the
    // workbook; their closure is not rewritten here.
    // Each substream is scanned to its own EOF; bytes after a substream's EOF
    // are retained as opaque padding and are not BIFF records.
    let substreams = std::iter::once(workbook).chain(
        catalog
            .iter()
            .map(|other| &workbook[other.start..other.end]),
    );
    for substream in substreams {
        if raw_records_until_eof(substream)?
            .iter()
            .any(|record| matches!(record.kind, DCON_REF | DCON_NAME | DCON_BIN))
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "worksheet removal cannot close pivot data connection record",
            ));
        }
    }
    Ok(())
}

/// Prepare an empty worksheet substream from a freshly authored workbook.
///
/// The writer's single default worksheet is the template: its records are
/// exactly what the complete reader accepts for a new worksheet.  Its
/// `Window2` is deselected so the inserted tab never joins the selection.
pub(super) fn empty_worksheet_template() -> Result<Template> {
    let mut builder = crate::writer::Builder::new();
    builder.add_worksheet("Sheet1")?;
    let template_error = |error: crate::package::Error| match error {
        crate::package::Error::Reader(error) => error,
        _ => Error::InvalidData("the empty worksheet template could not be authored".into()),
    };
    let package = builder.build().map_err(template_error)?;
    let bytes = package.to_bytes().map_err(template_error)?;
    let snapshot =
        super::Snapshot::from_bytes(bytes, litchi_ole_common::object::Limits::default())?;
    let workbook = snapshot.workbook_stream();
    let (_, tabs) = parse(workbook)?;
    let tab = tabs
        .first()
        .ok_or_else(|| Error::InvalidData("worksheet template has no tab".into()))?;
    let mut stream = Vec::new();
    stream
        .try_reserve_exact(tab.end - tab.start)
        .map_err(|_error| Error::Allocation("retaining the worksheet template"))?;
    stream.extend_from_slice(&workbook[tab.start..tab.end]);
    window2_flags(&mut stream, 0, WINDOW2_SELECTED | WINDOW2_PAGED)?;
    Ok(Template {
        stream,
        origin: tab.start,
    })
}

/// Content fingerprint of a source tab's directory entry and substream.
pub(super) fn fingerprint(workbook: &[u8], tab: &Tab) -> String {
    let mut bytes = Vec::with_capacity(tab.body_end - tab.body_start + tab.end - tab.start);
    bytes.extend_from_slice(&workbook[tab.body_start + 4..tab.body_end]);
    bytes.extend_from_slice(&workbook[tab.start..tab.end]);
    litchi_core::patch::DiagnosticFingerprint::of(&bytes).as_hex()
}

/// Semantic patch target of a complete-catalog reorder.
pub(super) const SEMANTIC_TARGET: &str = "workbook/sheets";

/// Semantic patch target of one catalog position.
pub(super) fn semantic_tab_target(position: usize) -> String {
    format!("workbook/sheet/{position}")
}

pub(super) fn parse_semantic_tab_target(target: &str) -> Result<usize> {
    target
        .strip_prefix("workbook/sheet/")
        .and_then(|position| position.parse().ok())
        .ok_or_else(|| Error::InvalidData("tab patch target is malformed".into()))
}
