//! Row and cell snapshot writer.

use std::collections::BTreeMap;

use litchi_core::xml::escape_xml;
use litchi_sheet::{Cell as Address, Row};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::super::super::wire::{
    cell_tag, sibling_name, tag, write_attribute, write_cell_tag, write_close, write_tag,
};
use super::super::model::{
    CellSlot, CompactCellSlot, CompactRowSlot, CompactSheetData, RowSlot, SheetData, Span, Tag,
};
use crate::cell::{Content, Stored, Value};
use crate::error::{Result, invalid};
use crate::outline::Outline;
use crate::raw::strings::encode_spreadsheet_text;
use crate::raw::worksheet::edit::CompactDimensionTag;
use crate::raw::worksheet::edit::OmittedCells;
use crate::raw::worksheet::edit::model::{
    Action, DescentEffect, HeightEffect, Payload, RowAction, StyleEffect,
};

pub(crate) fn write_sheet_data(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &SheetData,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<Row, RowAction>,
    descent_name: &str,
) -> Result<()> {
    let mut by_row = BTreeMap::<u32, RowEdits>::new();
    for (address, action) in cells {
        by_row
            .entry(address.row().get() + 1)
            .or_default()
            .cells
            .insert(address, action);
    }
    for (row, action) in rows {
        by_row.entry(row.get() + 1).or_default().row = Some(action);
    }

    if data.empty {
        write_tag(output, &data.tag, false, &[], &[]);
        for (number, edits) in by_row {
            write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
        }
        write_close(output, &data.tag.name);
        return Ok(());
    }

    output.extend_from_slice(&source[data.span.start..data.tag_end]);
    let mut cursor = data.tag_end;
    let mut pending = by_row.into_iter().peekable();
    for row in &data.rows {
        output.extend_from_slice(&source[cursor..row.span.start]);
        while pending
            .peek()
            .is_some_and(|(number, _)| *number < row.number)
        {
            if let Some((number, edits)) = pending.next() {
                write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
            }
        }
        if pending
            .peek()
            .is_some_and(|(number, _)| *number == row.number)
        {
            let (_, edits) = pending
                .next()
                .ok_or_else(|| invalid("worksheet row edit ordering was lost"))?;
            write_row(output, source, row, &edits, descent_name)?;
        } else {
            output.extend_from_slice(&source[row.span.start..row.span.end]);
        }
        cursor = row.span.end;
    }
    output.extend_from_slice(&source[cursor..data.close_start]);
    for (number, edits) in pending {
        write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
    }
    output.extend_from_slice(&source[data.close_start..data.span.end]);
    Ok(())
}

/// Write the complete ordinary output and retain the actual output spans of
/// cell bodies that a reduced semantic readback may omit. This is a separate
/// value-only route; callers that need row, column, or default effects keep
/// using [`write_sheet_data`].
pub(crate) fn write_sheet_data_with_provenance(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &SheetData,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<Row, RowAction>,
    descent_name: &str,
) -> Result<Box<[OmittedCells]>> {
    let mut by_row = BTreeMap::<u32, RowEdits>::new();
    for (address, action) in cells {
        let number = address
            .row()
            .get()
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet edit row overflows u32"))?;
        by_row
            .entry(number)
            .or_default()
            .cells
            .insert(address, action);
    }
    for (row, action) in rows {
        let number = row
            .get()
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet edit row overflows u32"))?;
        by_row.entry(number).or_default().row = Some(action);
    }
    let mut omitted = Vec::new();
    let mut recording = omitted.try_reserve(data.rows.len()).is_ok();
    if !recording {
        omitted.clear();
    }

    if data.empty {
        write_tag(output, &data.tag, false, &[], &[]);
        for (number, edits) in by_row {
            write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
        }
        write_close(output, &data.tag.name);
        return Ok(if recording {
            omitted.into_boxed_slice()
        } else {
            Box::new([])
        });
    }

    output.extend_from_slice(&source[data.span.start..data.tag_end]);
    let mut cursor = data.tag_end;
    let mut pending = by_row.into_iter().peekable();
    for row in &data.rows {
        output.extend_from_slice(&source[cursor..row.span.start]);
        while pending
            .peek()
            .is_some_and(|(number, _)| *number < row.number)
        {
            if let Some((number, edits)) = pending.next() {
                write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
            }
        }
        if pending
            .peek()
            .is_some_and(|(number, _)| *number == row.number)
        {
            let (_, edits) = pending
                .next()
                .ok_or_else(|| invalid("worksheet row edit ordering was lost"))?;
            let membership_changed = edits.cells.iter().any(|(address, action)| {
                let exists = row
                    .cells
                    .binary_search_by_key(address, |cell| cell.address)
                    .is_ok();
                !exists || matches!(action, Action::Remove)
            });
            if membership_changed || edits.row.is_some() || row.empty {
                write_row(output, source, row, &edits, descent_name)?;
            } else {
                write_replacement_row(output, source, row, &edits, &mut omitted, &mut recording)?;
            }
        } else {
            output.extend_from_slice(&source[row.span.start..row.tag_end]);
            let body_start = output.len();
            output.extend_from_slice(&source[row.tag_end..row.close_start]);
            let body_end = output.len();
            record_omitted_row(&mut omitted, &mut recording, row, body_start, body_end)?;
            output.extend_from_slice(&source[row.close_start..row.span.end]);
        }
        cursor = row.span.end;
    }
    output.extend_from_slice(&source[cursor..data.close_start]);
    for (number, edits) in pending {
        write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
    }
    output.extend_from_slice(&source[data.close_start..data.span.end]);
    Ok(if recording {
        omitted.into_boxed_slice()
    } else {
        Box::new([])
    })
}

/// Rewrite an existing-row scalar plan from a compact source proof.
///
/// The proof route intentionally accepts only updates to existing cells. Any
/// membership, row, formula, or empty-sheet effect remains on the complete
/// scanner/writer path. Unchanged cell bytes are copied directly and only the
/// changed cell tag is materialized.
pub(crate) fn write_sheet_data_with_compact_provenance<'a>(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &CompactSheetData,
    entries: &[Stored],
    cells: &'a BTreeMap<Address, Action>,
) -> Result<Box<[OmittedCells]>> {
    let data_start = data.span.start();
    let data_end = data.span.end();
    let data_tag_end = data.tag_end as usize;
    let data_close_start = data.close_start as usize;
    if data.empty || cells.is_empty() {
        output.extend_from_slice(&source[data_start..data_end]);
        return Ok(Box::new([]));
    }

    let mut by_row = BTreeMap::<u32, BTreeMap<Address, &'a Action>>::new();
    for (address, action) in cells {
        let number = address
            .row()
            .get()
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet edit row overflows u32"))?;
        by_row.entry(number).or_default().insert(*address, action);
    }
    let mut omitted = Vec::new();
    let mut recording = omitted.try_reserve_exact(data.rows.len()).is_ok();
    if !recording {
        omitted.clear();
    }
    output.extend_from_slice(&source[data_start..data_tag_end]);
    let mut cursor = data_tag_end;
    let mut entry_index = 0usize;
    for row in &data.rows {
        let row_start = row.span.start();
        let row_end = row.span.end();
        let row_tag_end = row.tag_end as usize;
        let row_close_start = row.close_start as usize;
        output.extend_from_slice(&source[cursor..row_start]);
        if let Some(edits) = by_row.remove(&row.number) {
            write_compact_replacement_row(
                output,
                source,
                row,
                entries,
                &mut entry_index,
                edits,
                &mut omitted,
                &mut recording,
            )?;
        } else {
            let (first, last) = compact_row_addresses(entries, &mut entry_index, row)?;
            output.extend_from_slice(&source[row_start..row_tag_end]);
            let body_start = output.len();
            output.extend_from_slice(&source[row_tag_end..row_close_start]);
            let body_end = output.len();
            if let (Some(first), Some(last)) = (first, last) {
                record_omitted(
                    &mut omitted,
                    &mut recording,
                    first,
                    last,
                    body_start,
                    body_end,
                )?;
            }
            output.extend_from_slice(&source[row_close_start..row_end]);
        }
        cursor = row_end;
    }
    if !by_row.is_empty() {
        return Err(invalid("compact worksheet proof lost an edited row"));
    }
    if entry_index != entries.len() {
        return Err(invalid("compact worksheet proof lost a source cell"));
    }
    output.extend_from_slice(&source[cursor..data_close_start]);
    output.extend_from_slice(&source[data_close_start..data_end]);
    Ok(if recording {
        omitted.into_boxed_slice()
    } else {
        Box::new([])
    })
}

pub(crate) fn write_compact_dimension(
    output: &mut Vec<u8>,
    source: &[u8],
    dimension: &CompactDimensionTag,
    range: litchi_sheet::Rect,
) -> Result<()> {
    let start = dimension.span.start();
    let end = dimension.span.end();
    if start > end || end > source.len() {
        return Err(invalid(
            "compact worksheet dimension span is outside its source",
        ));
    }
    let mut reader = NsReader::from_reader(&source[start..end]);
    let event = reader
        .read_event()
        .map_err(|error| invalid(error.to_string()))?;
    let (element, empty) = match event {
        Event::Start(element) => (element, false),
        Event::Empty(element) => (element, true),
        _ => return Err(invalid("compact worksheet dimension has no opening tag")),
    };
    if empty != dimension.empty {
        return Err(invalid(
            "compact worksheet dimension form changed under its proof",
        ));
    }
    let captured = tag(&element, reader.decoder())?;
    write_tag(output, &captured, empty, &["ref"], &[("ref", range.a1())]);
    Ok(())
}

fn write_compact_replacement_row<'a>(
    output: &mut Vec<u8>,
    source: &[u8],
    row: &CompactRowSlot,
    entries: &[Stored],
    entry_index: &mut usize,
    edits: BTreeMap<Address, &'a Action>,
    omitted: &mut Vec<OmittedCells>,
    recording: &mut bool,
) -> Result<()> {
    if row.empty {
        return Err(invalid("compact worksheet proof selected an empty row"));
    }
    let row_start = row.span.start();
    let row_end = row.span.end();
    let row_tag_end = row.tag_end as usize;
    let row_close_start = row.close_start as usize;
    output.extend_from_slice(&source[row_start..row_tag_end]);
    let mut cursor = row_tag_end;
    let mut pending = edits.into_iter().peekable();
    let mut run = None::<(usize, Address, Address)>;
    for cell in &row.cells {
        let cell_start_offset = cell.span.start();
        let cell_end_offset = cell.span.end();
        let address = compact_next_address(entries, entry_index, row.number)?;
        output.extend_from_slice(&source[cursor..cell_start_offset]);
        let cell_start = output.len();
        if pending
            .peek()
            .is_some_and(|(pending, _)| *pending < address)
        {
            return Err(invalid("compact worksheet proof lost an edited cell"));
        }
        if pending
            .peek()
            .is_some_and(|(pending, _)| *pending == address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("compact worksheet edit ordering was lost"))?;
            if let Some((start, first, last)) = run.take() {
                record_omitted(omitted, recording, first, last, start, cell_start)?;
            }
            if matches!(
                action,
                Action::Remove
                    | Action::Update {
                        payload: Some(Payload::SharedFormula { .. }),
                        ..
                    }
            ) {
                return Err(invalid(
                    "compact worksheet proof does not support membership edits",
                ));
            }
            write_compact_cell(output, source, address, cell, action)?;
        } else {
            let start = run.map(|(start, _, _)| start).unwrap_or(cell_start);
            let first = run.map(|(_, first, _)| first).unwrap_or(address);
            run = Some((start, first, address));
            output.extend_from_slice(&source[cell_start_offset..cell_end_offset]);
        }
        cursor = cell_end_offset;
    }
    if pending.next().is_some() {
        return Err(invalid("compact worksheet proof lost an edited cell"));
    }
    output.extend_from_slice(&source[cursor..row_close_start]);
    if let Some((start, first, last)) = run {
        record_omitted(omitted, recording, first, last, start, output.len())?;
    }
    output.extend_from_slice(&source[row_close_start..row_end]);
    Ok(())
}

fn compact_next_address(
    entries: &[Stored],
    entry_index: &mut usize,
    row_number: u32,
) -> Result<Address> {
    let entry = entries
        .get(*entry_index)
        .ok_or_else(|| invalid("compact worksheet proof lost a source cell"))?;
    if entry.address.row().get().checked_add(1) != Some(row_number) {
        return Err(invalid(
            "compact worksheet proof source row ordering changed",
        ));
    }
    *entry_index = (*entry_index)
        .checked_add(1)
        .ok_or_else(|| invalid("compact worksheet proof source cell count overflow"))?;
    Ok(entry.address)
}

fn compact_row_addresses(
    entries: &[Stored],
    entry_index: &mut usize,
    row: &CompactRowSlot,
) -> Result<(Option<Address>, Option<Address>)> {
    let mut first = None;
    let mut last = None;
    for _cell in &row.cells {
        let address = compact_next_address(entries, entry_index, row.number)?;
        first.get_or_insert(address);
        last = Some(address);
    }
    Ok((first, last))
}

fn compact_absolute_offset(base: usize, offset: usize) -> Result<usize> {
    base.checked_add(offset)
        .ok_or_else(|| invalid("compact worksheet cell offset overflows usize"))
}

fn write_compact_cell(
    output: &mut Vec<u8>,
    source: &[u8],
    address: Address,
    compact: &CompactCellSlot,
    action: &Action,
) -> Result<()> {
    let start = compact.span.start();
    let end = compact.span.end();
    if start > end || end > source.len() {
        return Err(invalid("compact worksheet cell span is outside its source"));
    }
    let slot = compact_cell_slot(source, address, start, end)?;
    write_cell(output, source, &slot, action)
}

fn compact_cell_slot(
    source: &[u8],
    address: Address,
    start: usize,
    end: usize,
) -> Result<CellSlot> {
    let mut reader = NsReader::from_reader(&source[start..end]);
    reader.config_mut().check_end_names = true;
    let opening = reader
        .read_event()
        .map_err(|error| invalid(error.to_string()))?;
    let (tag, empty) = match opening {
        Event::Start(element) => (cell_tag(&element, reader.decoder())?, false),
        Event::Empty(element) => (cell_tag(&element, reader.decoder())?, true),
        _ => return Err(invalid("compact worksheet cell has no opening tag")),
    };
    let mut primary = None::<Span>;
    let mut depth = 1usize;
    let tag_end = compact_absolute_offset(
        start,
        usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("compact worksheet cell offset exceeds usize"))?,
    )?;
    if empty {
        return Ok(CellSlot {
            address,
            span: Span { start, end },
            tag_end,
            close_start: tag_end,
            tag,
            primary: Box::new([]),
            mce_payload: false,
            empty: true,
        });
    }
    let close_start;
    loop {
        let event_start = compact_absolute_offset(
            start,
            usize::try_from(reader.buffer_position())
                .map_err(|_| invalid("compact worksheet cell offset exceeds usize"))?,
        )?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?;
        let event_end = compact_absolute_offset(
            start,
            usize::try_from(reader.buffer_position())
                .map_err(|_| invalid("compact worksheet cell offset exceeds usize"))?,
        )?;
        match event {
            Event::Start(element) => {
                let direct_primary = depth == 1
                    && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is");
                if direct_primary {
                    if primary.is_some() {
                        return Err(invalid("compact worksheet cell has duplicate primary"));
                    }
                    primary = Some(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("compact worksheet cell depth overflow"))?;
            },
            Event::Empty(element) => {
                if depth == 1 && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
                {
                    if primary.is_some() {
                        return Err(invalid("compact worksheet cell has duplicate primary"));
                    }
                    primary = Some(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
            },
            Event::End(element) => {
                let previous = depth;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("compact worksheet cell has unmatched close"))?;
                if previous == 2
                    && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
                {
                    let Some(primary) = primary.as_mut() else {
                        return Err(invalid("compact worksheet cell primary close is unmatched"));
                    };
                    primary.end = event_end;
                }
                if previous == 1 {
                    if element.name().local_name().as_ref() != b"c" || depth != 0 {
                        return Err(invalid("compact worksheet cell close is invalid"));
                    }
                    close_start = event_start;
                    break;
                }
            },
            Event::Eof => {
                return Err(invalid("compact worksheet cell has no closing tag"));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_) => {},
        }
    }
    let primary: Box<[Span]> = match primary {
        Some(span) => vec![span].into_boxed_slice(),
        None => Box::new([]),
    };
    let slot = CellSlot {
        address,
        span: Span { start, end },
        tag_end,
        close_start,
        tag,
        primary,
        mce_payload: false,
        empty,
    };
    Ok(slot)
}

fn write_replacement_row(
    output: &mut Vec<u8>,
    source: &[u8],
    row: &RowSlot,
    edits: &RowEdits,
    omitted: &mut Vec<OmittedCells>,
    recording: &mut bool,
) -> Result<()> {
    output.extend_from_slice(&source[row.span.start..row.tag_end]);
    let mut cursor = row.tag_end;
    let mut pending = edits.cells.iter().peekable();
    let mut run = None::<(usize, Address, Address)>;
    for cell in &row.cells {
        output.extend_from_slice(&source[cursor..cell.span.start]);
        let cell_start = output.len();
        while pending
            .peek()
            .is_some_and(|(address, _)| **address < cell.address)
        {
            let (address, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            if let Some((start, first, last)) = run.take() {
                record_omitted(omitted, recording, first, last, start, cell_start)?;
            }
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        if pending
            .peek()
            .is_some_and(|(address, _)| **address == cell.address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            if let Some((start, first, last)) = run.take() {
                record_omitted(omitted, recording, first, last, start, cell_start)?;
            }
            match action {
                Action::Update { .. } => write_cell(output, source, cell, action)?,
                Action::Remove => {
                    return Err(invalid("replacement row unexpectedly removes a cell"));
                },
            }
        } else {
            let start = run.map(|(start, _, _)| start).unwrap_or(cell_start);
            let first = run.map(|(_, first, _)| first).unwrap_or(cell.address);
            run = Some((start, first, cell.address));
            output.extend_from_slice(&source[cell.span.start..cell.span.end]);
        }
        cursor = cell.span.end;
    }
    output.extend_from_slice(&source[cursor..row.close_start]);
    if let Some((start, first, last)) = run {
        record_omitted(omitted, recording, first, last, start, output.len())?;
    }
    output.extend_from_slice(&source[row.close_start..row.span.end]);
    Ok(())
}

fn record_omitted_row(
    omitted: &mut Vec<OmittedCells>,
    recording: &mut bool,
    row: &RowSlot,
    start: usize,
    end: usize,
) -> Result<()> {
    let Some(first) = row.cells.first() else {
        return Ok(());
    };
    let last = row
        .cells
        .last()
        .ok_or_else(|| invalid("worksheet row cell provenance lost its last cell"))?;
    record_omitted(omitted, recording, first.address, last.address, start, end)
}

fn record_omitted(
    omitted: &mut Vec<OmittedCells>,
    recording: &mut bool,
    first: Address,
    last: Address,
    start: usize,
    end: usize,
) -> Result<()> {
    if start >= end {
        return Ok(());
    }
    if first.row() != last.row() || first.column() > last.column() {
        return Err(invalid(
            "worksheet cell provenance has an invalid address range",
        ));
    }
    if !*recording {
        return Ok(());
    }
    if omitted.try_reserve(1).is_err() {
        *recording = false;
        omitted.clear();
        return Ok(());
    }
    omitted.push(OmittedCells {
        row: first.row().get(),
        first_column: first.column().get(),
        last_column: last.column().get(),
        start,
        end,
    });
    Ok(())
}

#[derive(Debug, Default)]
struct RowEdits {
    cells: BTreeMap<Address, Action>,
    row: Option<RowAction>,
}

fn write_row(
    output: &mut Vec<u8>,
    source: &[u8],
    row: &RowSlot,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let actions = &edits.cells;
    let membership_changed = actions.iter().any(|(address, action)| {
        let exists = row
            .cells
            .binary_search_by_key(address, |cell| cell.address)
            .is_ok();
        (!exists && action.creates_missing()) || (exists && matches!(action, Action::Remove))
    });

    if row.empty {
        let creates_cell = actions.values().any(Action::creates_missing);
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if creates_cell {
            removed.extend(["spans", "r"]);
            appended.push(("r", row.number.to_string()));
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, !creates_cell, &removed, &appended);
        if !creates_cell {
            return Ok(());
        }
        for (address, action) in actions {
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        write_close(output, &row.tag.name);
        return Ok(());
    }

    if membership_changed || edits.row.is_some() {
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if membership_changed {
            removed.push("spans");
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, false, &removed, &appended);
    } else {
        output.extend_from_slice(&source[row.span.start..row.tag_end]);
    }
    let mut cursor = row.tag_end;
    let mut pending = actions.iter().peekable();
    for cell in &row.cells {
        output.extend_from_slice(&source[cursor..cell.span.start]);
        while pending
            .peek()
            .is_some_and(|(address, _)| **address < cell.address)
        {
            let (address, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        if pending
            .peek()
            .is_some_and(|(address, _)| **address == cell.address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            match action {
                Action::Update { .. } => write_cell(output, source, cell, action)?,
                Action::Remove => {},
            }
        } else {
            output.extend_from_slice(&source[cell.span.start..cell.span.end]);
        }
        cursor = cell.span.end;
    }
    output.extend_from_slice(&source[cursor..row.close_start]);
    for (address, action) in pending {
        write_new_action(output, &row.tag.name, *address, action)?;
    }
    output.extend_from_slice(&source[row.close_start..row.span.end]);
    Ok(())
}

fn write_new_row(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    number: u32,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let creates_cell = edits.cells.values().any(Action::creates_missing);
    let materializes = edits.row.is_some_and(RowAction::materializes);
    if !creates_cell && !materializes {
        return Ok(());
    }
    let name = sibling_name(sheet_data_name, "row");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", number.to_string())];
    let mut removed = Vec::new();
    if let Some(action) = edits.row {
        row_effect_attributes(action, descent_name, &mut removed, &mut appended);
    }
    write_tag(output, &tag, !creates_cell, &removed, &appended);
    if !creates_cell {
        return Ok(());
    }
    for (address, action) in &edits.cells {
        write_new_action(output, &name, *address, action)?;
    }
    write_close(output, &name);
    Ok(())
}

fn row_effect_attributes<'a>(
    action: RowAction,
    descent_name: &'a str,
    removed: &mut Vec<&'a str>,
    appended: &mut Vec<(&'a str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(height) = action.height {
        removed.extend(["ht", "customHeight"]);
        if let HeightEffect::Set(height) = height {
            appended.push(("ht", height.get().to_string()));
            appended.push(("customHeight", "1".to_owned()));
        }
    }
    if let Some(descent) = action.descent {
        removed.push(descent_name);
        if let DescentEffect::Set(value) = descent {
            appended.push((descent_name, value.get().to_string()));
        }
    }
    if let Some(style) = action.style {
        removed.extend(["s", "customFormat"]);
        if let StyleEffect::Set(key) = style {
            appended.push(("s", key.to_string()));
            appended.push(("customFormat", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    for (value, name) in [
        (action.collapsed, "collapsed"),
        (action.thick_top, "thickTop"),
        (action.thick_bottom, "thickBot"),
        (action.phonetic, "ph"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
}

fn write_cell(output: &mut Vec<u8>, source: &[u8], cell: &CellSlot, action: &Action) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Err(invalid("cannot rewrite a removed cell"));
    };
    let content = payload.as_ref().filter(|payload| {
        matches!(
            payload,
            Payload::Set(_) | Payload::SharedString { .. } | Payload::SharedFormula { .. }
        )
    });
    let cell_type = content.and_then(payload_type);
    let mut removed = vec!["r"];
    if payload.is_some() {
        removed.push("t");
    }
    if style.is_some() {
        removed.push("s");
    }
    let mut appended = vec![("r", cell.address.a1())];
    if let Some(cell_type) = cell_type {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let remains_empty = cell.empty && payload.is_none();
    let cell_name = if let Some(tag) = cell.tag.as_ref() {
        write_tag(output, tag, remains_empty, &removed, &appended);
        tag.name.as_ref()
    } else {
        write_cell_tag(output, remains_empty, &appended);
        "c"
    };
    if remains_empty {
        return Ok(());
    }
    if let Some(content) = content {
        write_payload(output, cell_name, content)?;
    }
    if !cell.empty {
        if payload.is_some() {
            copy_without(
                output,
                source,
                cell.tag_end,
                cell.close_start,
                &cell.primary,
            );
        } else {
            output.extend_from_slice(&source[cell.tag_end..cell.close_start]);
        }
    }
    write_close(output, cell_name);
    Ok(())
}

fn write_new_action(
    output: &mut Vec<u8>,
    row_name: &str,
    address: Address,
    action: &Action,
) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Ok(());
    };
    if !action.creates_missing() {
        return Ok(());
    }
    let content = payload.as_ref().filter(|payload| {
        matches!(
            payload,
            Payload::Set(_) | Payload::SharedString { .. } | Payload::SharedFormula { .. }
        )
    });
    let name = sibling_name(row_name, "c");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", address.a1())];
    if let Some(cell_type) = content.and_then(payload_type) {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let empty = content.is_none();
    write_tag(output, &tag, empty, &[], &appended);
    if let Some(content) = content {
        write_payload(output, &name, content)?;
        write_close(output, &name);
    }
    Ok(())
}

fn content_type(content: &Content) -> Option<&'static str> {
    match content {
        Content::Value(Value::Bool(_)) => Some("b"),
        Content::Value(Value::Text(_)) => Some("inlineStr"),
        Content::Value(Value::Date(_)) => Some("d"),
        Content::Value(Value::Error(_)) => Some("e"),
        Content::Value(Value::Number(_)) | Content::Formula(_) => None,
    }
}

fn payload_type(payload: &Payload) -> Option<&'static str> {
    match payload {
        Payload::Set(content) => content_type(content),
        Payload::SharedString { .. } => Some("s"),
        Payload::SharedFormula { .. } => None,
        Payload::Clear | Payload::ClearIfPresent => None,
    }
}

fn write_payload(output: &mut Vec<u8>, cell_name: &str, payload: &Payload) -> Result<()> {
    match payload {
        Payload::Set(content) => write_content(output, cell_name, content),
        Payload::SharedString { index, .. } => {
            write_text_element(output, cell_name, "v", &index.to_string());
            Ok(())
        },
        Payload::SharedFormula {
            index,
            reference,
            formula,
        } => write_shared_formula(output, cell_name, *index, reference, formula.as_ref()),
        Payload::Clear | Payload::ClearIfPresent => Ok(()),
    }
}

fn write_shared_formula(
    output: &mut Vec<u8>,
    cell_name: &str,
    index: u32,
    reference: &str,
    formula: Option<&crate::formula::Formula>,
) -> Result<()> {
    require_xml_text(reference)?;
    if let Some(formula) = formula {
        require_xml_text(formula.text())?;
    }
    let name = sibling_name(cell_name, "f");
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    write_attribute(output, "t", "shared");
    if formula.is_some() {
        write_attribute(output, "ref", reference);
    }
    write_attribute(output, "si", &index.to_string());
    if let Some(formula) = formula {
        output.extend_from_slice(b">");
        output.extend_from_slice(escape_xml(formula.text()).as_bytes());
        write_close(output, &name);
    } else {
        output.extend_from_slice(b"/>");
    }
    Ok(())
}

fn write_content(output: &mut Vec<u8>, cell_name: &str, content: &Content) -> Result<()> {
    match content {
        Content::Value(Value::Bool(value)) => {
            write_text_element(output, cell_name, "v", if *value { "1" } else { "0" });
        },
        Content::Value(Value::Number(value)) => {
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Value(Value::Text(value)) => {
            let inline = sibling_name(cell_name, "is");
            let text = sibling_name(cell_name, "t");
            output.extend_from_slice(b"<");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b"><");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b" xml:space=\"preserve\">");
            output.extend_from_slice(escape_xml(&encode_spreadsheet_text(value)).as_bytes());
            output.extend_from_slice(b"</");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b"></");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b">");
        },
        Content::Value(Value::Date(value)) => {
            require_xml_text(value)?;
            write_text_element(output, cell_name, "v", &escape_xml(value));
        },
        Content::Value(Value::Error(value)) => {
            require_xml_text(value.as_str())?;
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Formula(formula) => {
            require_xml_text(formula.text())?;
            write_text_element(output, cell_name, "f", &escape_xml(formula.text()));
        },
    }
    Ok(())
}

fn write_text_element(output: &mut Vec<u8>, cell_name: &str, local: &str, value: &str) {
    let name = sibling_name(cell_name, local);
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
    output.extend_from_slice(value.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn require_xml_text(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
            || ('\u{20}'..='\u{D7FF}').contains(&character)
            || ('\u{E000}'..='\u{FFFD}').contains(&character)
            || ('\u{10000}'..='\u{10FFFF}').contains(&character)
    }) {
        Ok(())
    } else {
        Err(invalid(
            "cell content contains a character forbidden by XML 1.0",
        ))
    }
}

fn copy_without(output: &mut Vec<u8>, source: &[u8], start: usize, end: usize, removed: &[Span]) {
    let mut cursor = start;
    for span in removed {
        output.extend_from_slice(&source[cursor..span.start]);
        cursor = span.end;
    }
    output.extend_from_slice(&source[cursor..end]);
}
