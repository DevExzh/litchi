use super::{
    BLANK, BOOL_ERR, Error, FORMULA, LABEL_SST, MUL_RK, NUMBER, RK, Reference, Result, Storage,
    StructuralChange, StyleIndex, Value,
};
use crate::SourceLimits;
use crate::writer::row_blocks::{RowBlockLayoutPlan, RowBlockLayoutRow};
use litchi_biff::Records;
use litchi_core::binary;
use litchi_core::edit::RefusalKind;
use std::collections::BTreeMap;

const MAX_SST_PAYLOAD: usize = 8_224;

const BOUND_SHEET: u16 = 0x0085;
const XF: u16 = 0x00e0;
const EXT_SST: u16 = 0x00ff;
const CONTINUE: u16 = 0x003c;
const INDEX: u16 = 0x020b;
const ROW: u16 = 0x0208;
const DBCELL: u16 = 0x00d7;
const DIMENSIONS: u16 = 0x0200;
const DEF_COL_WIDTH: u16 = 0x0055;
const TABLE: u16 = 0x0236;
const SHARED_FORMULA: u16 = 0x04bc;
const ARRAY: u16 = 0x0221;
const STRING: u16 = 0x0207;
const LABEL: u16 = 0x0204;
const RSTRING: u16 = 0x00d6;
const MUL_BLANK: u16 = 0x00be;
const MERGED_CELLS: u16 = 0x00e5;
const HLINK: u16 = 0x01b8;
const HLINK_TOOLTIP: u16 = 0x0800;
const MSO_DRAWING: u16 = 0x00ec;
const OBJ: u16 = 0x005d;
const SELECTION: u16 = 0x001d;
const VERTICAL_PAGE_BREAKS: u16 = 0x001a;
const HORIZONTAL_PAGE_BREAKS: u16 = 0x001b;
const NOTE: u16 = 0x001c;
const COL_INFO: u16 = 0x007d;

#[derive(Clone, Copy)]
pub(super) enum AxisShift {
    Rows {
        start: u16,
        count: u16,
        insert: bool,
    },
    Columns {
        start: u8,
        count: u8,
        insert: bool,
    },
}

#[derive(Clone)]
struct RawRecord {
    kind: u16,
    start: usize,
    end: usize,
}

#[derive(Clone)]
struct BoundSheet {
    record: RawRecord,
    position: u32,
}

struct RowData {
    row_record: Vec<u8>,
    cell_records: Vec<Vec<u8>>,
    /// The source `Row` extents `(colMic, colMac)` when the producer wrote
    /// wider extents than the row's cells span (Excel and LibreOffice often
    /// write the sheet-wide extents). Column shifts move them like the cells;
    /// [`patch_row_extents`] keeps them while they still cover the cells, so
    /// a structural edit and its inverse leave such rows byte-identical.
    stored_extents: Option<(u16, u16)>,
}

pub(super) fn apply(
    mut workbook: Vec<u8>,
    source: &super::Snapshot,
    changes: &[StructuralChange],
    fixed_changes: &[super::Change],
    resources: &[super::ResourceChange],
    layout: &[&super::layout::LayoutChange],
    shared_strings: &[String],
) -> Result<Vec<u8>> {
    let rebuild_sst = needs_sst_rebuild(source, changes, fixed_changes, resources);
    let mut inserted_resources: Vec<_> = resources
        .iter()
        .filter(|resource| {
            resource_insert(resource)
                && !(rebuild_sst
                    && matches!(
                        resource,
                        super::ResourceChange::SharedString { .. }
                            | super::ResourceChange::RichSharedString { .. }
                    ))
        })
        .collect();
    inserted_resources.sort_by_cached_key(|resource| super::resource_target(resource));
    for resource in inserted_resources {
        apply_resource(&mut workbook, resource)?;
    }
    for change in changes {
        if let StructuralChange::RenameSheet {
            sheet,
            before,
            after,
        } = change
        {
            rename_sheet(&mut workbook, source, *sheet, before, after)?;
        }
    }

    let mut by_sheet: BTreeMap<usize, Vec<&StructuralChange>> = BTreeMap::new();
    for change in changes {
        if matches!(change, StructuralChange::RenameSheet { .. }) {
            continue;
        }
        let sheet = operation_sheet(change);
        by_sheet.entry(sheet).or_default().push(change);
    }
    for resource in resources {
        if let super::ResourceChange::FormulaCell { sheet, .. } = resource {
            by_sheet.entry(*sheet).or_default();
        }
    }
    for change in layout {
        by_sheet.entry(change.sheet()).or_default();
    }
    // Rewriting later streams first keeps earlier absolute offsets stable and
    // minimizes the number of BoundSheet position adjustments.
    for (sheet, operations) in by_sheet.into_iter().rev() {
        let sheet_layout: Vec<&super::layout::LayoutChange> = layout
            .iter()
            .copied()
            .filter(|change| change.sheet() == sheet)
            .collect();
        rewrite_worksheet(
            &mut workbook,
            source,
            sheet,
            &operations,
            resources,
            &sheet_layout,
            shared_strings,
        )?;
    }
    let mut removed_resources: Vec<_> = resources
        .iter()
        .filter(|resource| {
            !resource_insert(resource)
                && !(rebuild_sst
                    && matches!(
                        resource,
                        super::ResourceChange::SharedString { .. }
                            | super::ResourceChange::RichSharedString { .. }
                    ))
        })
        .collect();
    removed_resources
        .sort_by_cached_key(|resource| std::cmp::Reverse(super::resource_target(resource)));
    for resource in removed_resources {
        apply_resource(&mut workbook, resource)?;
    }
    if rebuild_sst {
        compact_shared_strings(&mut workbook, source, resources, shared_strings)?;
    }
    Ok(workbook)
}

fn needs_sst_rebuild(
    source: &super::Snapshot,
    changes: &[StructuralChange],
    fixed_changes: &[super::Change],
    resources: &[super::ResourceChange],
) -> bool {
    fixed_changes
        .iter()
        .filter(|change| super::change_is_effective(source, change))
        .any(|change| {
            let before = source
                .inner
                .sheets
                .get(change.sheet)
                .and_then(|sheet| sheet.entries.get(change.entry))
                .map(|entry| entry.cell.storage);
            before == Some(Storage::LabelSst) || change.storage == Storage::LabelSst
        })
        || resources.iter().any(|resource| {
            matches!(
                resource,
                super::ResourceChange::SharedString { .. }
                    | super::ResourceChange::RichSharedString { .. }
            )
        })
        || changes.iter().any(|change| {
            matches!(
                change,
                StructuralChange::Cell {
                    before: Some((Storage::LabelSst, ..)),
                    ..
                } | StructuralChange::Cell {
                    after: Some((Storage::LabelSst, ..)),
                    ..
                }
            )
        })
}

/// Rebuilds the BIFF8 shared-string closure from the final LabelSst records.
///
/// SST resources are an implementation detail of the transaction planner. A
/// text replacement can leave the old unique string unreferenced, while a
/// text-to-number or clear operation can leave no reference at all. Keeping
/// that dead entry would require inflating `cstTotal`, which would no longer
/// describe the workbook. This closure therefore removes dead entries,
/// remaps every final LabelSst index, and regenerates ExtSST when the source
/// carried that optional accelerator.
pub(super) fn compact_shared_strings(
    workbook: &mut Vec<u8>,
    source: &super::Snapshot,
    resources: &[super::ResourceChange],
    shared_strings: &[String],
) -> Result<()> {
    let globals = workbook_globals(workbook)?;
    let sst = globals.iter().find(|record| record.kind == super::SST);
    let (sst_start, family_end) = if let Some(sst) = sst {
        if sst.end - sst.start < 12 {
            return Err(Error::InvalidData("SST header is truncated".into()));
        }
        let sst_index = globals
            .iter()
            .position(|record| record.start == sst.start)
            .ok_or(Error::Internal("SST record disappeared from globals"))?;
        let mut family_end = sst.end;
        for record in &globals[sst_index + 1..] {
            if record.kind != CONTINUE {
                break;
            }
            family_end = record.end;
        }
        (sst.start, family_end)
    } else {
        if !source.inner.shared_strings.is_empty() {
            return Err(Error::InvalidData(
                "source shared strings have no SST owner".into(),
            ));
        }
        let eof = globals
            .last()
            .filter(|record| record.kind == 0x000a)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "workbook globals have no EOF owner",
                )
            })?;
        (eof.start, eof.start)
    };
    let ext_sst = globals
        .iter()
        .filter(|record| record.kind == EXT_SST)
        .try_fold(None, |found, record| {
            if found.is_some() {
                Err(Error::InvalidData("duplicate ExtSST record".into()))
            } else {
                Ok(Some(record.clone()))
            }
        })?;
    if let Some(ext_sst) = &ext_sst
        && ext_sst.start != family_end
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedContent,
            "edit cells",
            "ExtSST is not adjacent to its SST record family",
        ));
    }

    let effective = effective_sst_entries(source, resources)?;
    if effective.len() != shared_strings.len()
        || effective
            .iter()
            .zip(shared_strings)
            .any(|(entry, text)| entry.text != *text)
    {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "effective shared-string resource order is stale",
        });
    }

    let mut references: Vec<u32> = Vec::new();
    references
        .try_reserve_exact(effective.len())
        .map_err(|_error| Error::Allocation("retaining SST reference counts"))?;
    references.resize(effective.len(), 0);
    let records = raw_records(workbook)?;
    for record in records.iter().filter(|record| record.kind == LABEL_SST) {
        let index = binary::read_u32_le_at(workbook, record.start + 10)?;
        let index = usize::try_from(index)
            .map_err(|_error| Error::InvalidData("LabelSst index exceeds usize".into()))?;
        let count = references
            .get_mut(index)
            .ok_or_else(|| Error::InvalidRecord {
                record_type: LABEL_SST,
                message: format!("LabelSst index {index} is outside the effective SST"),
            })?;
        *count = (*count)
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("SST total reference count overflow".into()))?;
    }

    let mut remap: Vec<Option<u32>> = Vec::new();
    remap
        .try_reserve_exact(effective.len())
        .map_err(|_error| Error::Allocation("retaining SST index remapping"))?;
    remap.resize(effective.len(), None);
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(references.iter().filter(|count| **count != 0).count())
        .map_err(|_error| Error::Allocation("retaining referenced SST entries"))?;
    for (index, (entry, count)) in effective.iter().zip(references.iter()).enumerate() {
        if *count == 0 {
            if entry.source_index.is_some() && !semantic_sst_entry_is_reversible(entry)? {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "compacting an unreferenced rich or repaired SST entry would lose semantic inverse metadata",
                ));
            }
            continue;
        }
        let target = u32::try_from(retained.len())
            .map_err(|_error| Error::InvalidData("SST unique count exceeds u32".into()))?;
        remap[index] = Some(target);
        retained.push(entry.clone());
    }
    let total = references.iter().try_fold(0_u32, |total, count| {
        total
            .checked_add(*count)
            .ok_or_else(|| Error::InvalidData("SST total reference count overflow".into()))
    })?;
    let unique = u32::try_from(retained.len())
        .map_err(|_error| Error::InvalidData("SST unique count exceeds u32".into()))?;
    if total > i32::MAX as u32 || unique > i32::MAX as u32 {
        return Err(Error::InvalidData(
            "SST counts exceed the BIFF8 signed count range".into(),
        ));
    }

    for record in raw_records(workbook)?
        .into_iter()
        .filter(|record| record.kind == LABEL_SST)
    {
        let old = usize::try_from(binary::read_u32_le_at(workbook, record.start + 10)?)
            .map_err(|_error| Error::InvalidData("LabelSst index exceeds usize".into()))?;
        let target = remap.get(old).and_then(|value| *value).ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "final LabelSst points to a dead SST entry",
            )
        })?;
        super::write_field(workbook, record.start + 10, &target.to_le_bytes())?;
    }

    let encoded = encode_sst_records(&retained, total)?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(encoded.bytes_len())
        .map_err(|_error| Error::Allocation("retaining compacted SST records"))?;
    for record in &encoded.records {
        replacement.extend_from_slice(record);
    }
    if ext_sst.is_some() {
        let index = encoded.ext_sst_index(sst_start, unique)?;
        replacement.extend_from_slice(&index.to_record_bytes()?);
    }
    let old_end = ext_sst.as_ref().map_or(family_end, |record| record.end);
    replace_range_and_adjust_bounds(workbook, sst_start, old_end, &replacement)
}

#[derive(Clone)]
struct SstEntry {
    text: String,
    properties: Option<Box<crate::records::SharedStringProperties>>,
    source_index: Option<usize>,
    /// The source wire entry, when one was present. The public semantic model
    /// repairs malformed UTF-16 and drops invalid runs, so a semantic
    /// re-encode would silently change accepted source bytes. Retaining this
    /// opaque body preserves the source flags, code units, rich runs and
    /// ExtRst fields while still allowing the entry to move in the SST.
    raw: Option<RawSstEntry>,
}

fn semantic_sst_entry_is_reversible(entry: &SstEntry) -> Result<bool> {
    if entry.properties.as_deref().is_some_and(|properties| {
        !properties.formatting_runs.is_empty() || properties.phonetic.is_some()
    }) {
        return Ok(false);
    }
    let Some(raw) = entry.raw.as_ref() else {
        return Ok(true);
    };
    let flags = raw.prefix.get(2).copied().unwrap_or_default();
    if flags & 0x0c != 0 || !raw.formatting_runs.is_empty() || raw.extension.is_some() {
        return Ok(false);
    }
    let unit_count = raw.characters.iter().try_fold(0_usize, |count, segment| {
        let width = if segment.wide { 2 } else { 1 };
        if segment.bytes.len() % width != 0 {
            return Err(Error::InvalidData(
                "source SST character segment is not aligned to its encoding width".into(),
            ));
        }
        count
            .checked_add(segment.bytes.len() / width)
            .ok_or_else(|| Error::InvalidData("source SST character count overflows".into()))
    })?;
    if unit_count != entry.text.encode_utf16().count() {
        return Ok(false);
    }
    let mut units = Vec::new();
    units
        .try_reserve_exact(unit_count)
        .map_err(|_error| Error::Allocation("checking source SST character fidelity"))?;
    for segment in &raw.characters {
        if segment.wide {
            for pair in segment.bytes.chunks_exact(2) {
                units.push(u16::from_le_bytes([pair[0], pair[1]]));
            }
        } else {
            units.extend(segment.bytes.iter().copied().map(u16::from));
        }
    }
    Ok(String::from_utf16(&units).is_ok_and(|text| text == entry.text))
}

#[derive(Clone)]
struct RawSstEntry {
    prefix: Vec<u8>,
    characters: Vec<RawSstCharacters>,
    formatting_runs: Vec<u8>,
    extension: Option<Vec<u8>>,
}

#[derive(Clone)]
struct RawSstCharacters {
    wide: bool,
    bytes: Vec<u8>,
}

struct RawSstCursor<'a> {
    segments: Vec<&'a [u8]>,
    segment: usize,
    offset: usize,
}

impl<'a> RawSstCursor<'a> {
    fn new(segments: Vec<&'a [u8]>) -> Self {
        Self {
            segments,
            segment: 0,
            offset: 0,
        }
    }

    fn current(&self) -> &'a [u8] {
        self.segments.get(self.segment).copied().unwrap_or_default()
    }

    fn remaining(&self) -> usize {
        self.current().len().saturating_sub(self.offset)
    }

    fn advance(&mut self, context: &str) -> Result<()> {
        self.segment = self
            .segment
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData(format!("{context} segment overflow")))?;
        self.offset = 0;
        if self.segment >= self.segments.len() {
            return Err(Error::UnexpectedEndOfStream(context.into()));
        }
        Ok(())
    }

    fn read_u8(&mut self, context: &str) -> Result<u8> {
        if self.remaining() == 0 {
            self.advance(context)?;
        }
        let value = *self
            .current()
            .get(self.offset)
            .ok_or_else(|| Error::UnexpectedEndOfStream(context.into()))?;
        self.offset += 1;
        Ok(value)
    }

    fn read_exact(&mut self, output: &mut [u8], context: &str) -> Result<()> {
        let mut written = 0_usize;
        while written < output.len() {
            if self.remaining() == 0 {
                self.advance(context)?;
            }
            let count = self.remaining().min(output.len() - written);
            output[written..written + count]
                .copy_from_slice(&self.current()[self.offset..self.offset + count]);
            self.offset += count;
            written += count;
        }
        Ok(())
    }

    fn read_u32(&mut self, context: &str) -> Result<u32> {
        let mut bytes = [0_u8; 4];
        self.read_exact(&mut bytes, context)?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn read_bytes(&mut self, length: usize, context: &str) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|_error| Error::Allocation("retaining source ExtRst"))?;
        bytes.resize(length, 0);
        self.read_exact(&mut bytes, context)?;
        Ok(bytes)
    }

    fn read_character_segments(
        &mut self,
        count: u16,
        mut wide: bool,
    ) -> Result<Vec<RawSstCharacters>> {
        let count = usize::from(count);
        let mut consumed = 0_usize;
        let mut segments = Vec::new();
        segments
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining source SST character segments"))?;
        while consumed < count {
            let width = if wide { 2 } else { 1 };
            let available = self.remaining();
            let chunk = (available / width).min(count - consumed);
            if wide && !available.is_multiple_of(2) && chunk < count - consumed {
                return Err(Error::InvalidData(
                    "source SST UTF-16 data is split inside a code unit".into(),
                ));
            }
            let byte_count = chunk.saturating_mul(width);
            let end = self.offset.checked_add(byte_count).ok_or_else(|| {
                Error::InvalidData("source SST character offset overflows".into())
            })?;
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(byte_count)
                .map_err(|_error| Error::Allocation("retaining source SST character bytes"))?;
            bytes.extend_from_slice(&self.current()[self.offset..end]);
            segments.push(RawSstCharacters { wide, bytes });
            self.offset = end;
            consumed += chunk;
            if consumed == count {
                break;
            }
            if self.remaining() != 0 {
                return Err(Error::InvalidData(
                    "source SST character data does not end at a record boundary".into(),
                ));
            }
            self.advance("continued source SST character data")?;
            let flags = self.read_u8("source SST continuation flags")?;
            if flags > 1 {
                return Err(Error::InvalidData(
                    "source SST continuation flags are invalid".into(),
                ));
            }
            wide = flags == 1;
            segments
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("retaining source SST character segments"))?;
        }
        Ok(segments)
    }
}

fn source_sst_entries(source: &super::Snapshot) -> Result<Vec<Option<RawSstEntry>>> {
    let mut segments = Vec::new();
    segments
        .try_reserve(1)
        .map_err(|_error| Error::Allocation("retaining source SST records"))?;
    let mut found = false;
    for item in Records::new(
        &source.inner.workbook_stream,
        litchi_biff::Limits::default(),
    )? {
        let record = item?;
        if !found {
            if record.kind().get() == super::SST {
                found = true;
                segments.push(record.payload());
            }
        } else if record.kind().get() == CONTINUE {
            segments
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("retaining source SST records"))?;
            segments.push(record.payload());
        } else {
            break;
        }
    }
    if !found {
        if source.inner.shared_strings.is_empty() {
            return Ok(Vec::new());
        }
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source shared strings have no SST owner",
        ));
    }
    let mut cursor = RawSstCursor::new(segments);
    let total = cursor.read_u32("source SST total count")?;
    let unique = cursor.read_u32("source SST unique count")?;
    if total < unique || total > i32::MAX as u32 || unique > i32::MAX as u32 {
        return Err(Error::InvalidData(
            "source SST counts are outside the BIFF8 range".into(),
        ));
    }
    let unique = usize::try_from(unique)
        .map_err(|_error| Error::InvalidData("source SST unique count exceeds usize".into()))?;
    if unique > source.inner.workbook_stream.len().saturating_sub(8) / 3 {
        return Err(Error::InvalidData(
            "source SST declares more entries than its stream can contain".into(),
        ));
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(unique)
        .map_err(|_error| Error::Allocation("retaining source SST entries"))?;
    for _ in 0..unique {
        let character_count_bytes = cursor.read_bytes(2, "source SST character count")?;
        let character_count =
            u16::from_le_bytes(character_count_bytes[..2].try_into().map_err(|_error| {
                Error::InvalidData("source SST character count is truncated".into())
            })?);
        let flags_bytes = cursor.read_bytes(1, "source SST flags")?;
        let flags = flags_bytes[0];
        let prefix_length = 3_usize
            .checked_add(if flags & 0x08 != 0 { 2 } else { 0 })
            .and_then(|length| length.checked_add(if flags & 0x04 != 0 { 4 } else { 0 }))
            .ok_or_else(|| Error::InvalidData("source SST entry header length overflows".into()))?;
        let mut prefix = Vec::new();
        prefix
            .try_reserve_exact(prefix_length)
            .map_err(|_error| Error::Allocation("retaining source SST entry header"))?;
        prefix.extend_from_slice(&character_count_bytes);
        prefix.extend_from_slice(&flags_bytes);
        let run_count = if flags & 0x08 != 0 {
            let bytes = cursor.read_bytes(2, "source SST rich-run count")?;
            prefix.extend_from_slice(&bytes);
            u16::from_le_bytes(bytes[..2].try_into().map_err(|_error| {
                Error::InvalidData("source SST rich-run count is truncated".into())
            })?)
        } else {
            0
        };
        let extension_length = if flags & 0x04 != 0 {
            let bytes = cursor.read_bytes(4, "source SST extension length")?;
            prefix.extend_from_slice(&bytes);
            usize::try_from(u32::from_le_bytes(bytes[..4].try_into().map_err(
                |_error| Error::InvalidData("source SST extension length is truncated".into()),
            )?))
            .map_err(|_error| Error::InvalidData("source SST extension length overflows".into()))?
        } else {
            0
        };
        let characters = cursor.read_character_segments(character_count, flags & 0x01 != 0)?;
        let run_bytes = usize::from(run_count)
            .checked_mul(4)
            .ok_or_else(|| Error::InvalidData("source SST rich-run length overflows".into()))?;
        let formatting_runs = cursor.read_bytes(run_bytes, "source SST rich runs")?;
        let extension = if flags & 0x04 != 0 {
            Some(cursor.read_bytes(extension_length, "source SST ExtRst")?)
        } else {
            None
        };
        entries.push(Some(RawSstEntry {
            prefix,
            characters,
            formatting_runs,
            extension,
        }));
    }
    Ok(entries)
}

#[cfg(test)]
fn source_sst_extensions(source: &super::Snapshot) -> Result<Vec<Option<Vec<u8>>>> {
    Ok(source_sst_entries(source)?
        .into_iter()
        .map(|entry| entry.and_then(|entry| entry.extension))
        .collect())
}

fn effective_sst_entries(
    source: &super::Snapshot,
    resources: &[super::ResourceChange],
) -> Result<Vec<SstEntry>> {
    let insert_count = resources
        .iter()
        .filter(|resource| {
            matches!(
                resource,
                super::ResourceChange::SharedString { insert: true, .. }
                    | super::ResourceChange::RichSharedString { insert: true, .. }
            )
        })
        .count();
    let capacity = source
        .inner
        .shared_strings
        .len()
        .checked_add(insert_count)
        .ok_or_else(|| Error::InvalidData("effective SST entry count overflows".into()))?;
    let mut raw_entries = source_sst_entries(source)?;
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(capacity)
        .map_err(|_error| Error::Allocation("retaining source SST entries"))?;
    for (index, text) in source.inner.shared_strings.iter().enumerate() {
        entries.push(SstEntry {
            text: text.clone(),
            properties: source
                .inner
                .shared_string_properties
                .get(index)
                .cloned()
                .flatten(),
            source_index: Some(index),
            raw: raw_entries.get_mut(index).and_then(Option::take),
        });
    }
    let removal_count = resources
        .iter()
        .filter(|resource| {
            matches!(
                resource,
                super::ResourceChange::SharedString { insert: false, .. }
                    | super::ResourceChange::RichSharedString { insert: false, .. }
            )
        })
        .count();
    let mut removed_source_indices = Vec::new();
    removed_source_indices
        .try_reserve_exact(removal_count)
        .map_err(|_error| Error::Allocation("retaining removed SST resource indices"))?;
    for resource in resources.iter().filter(|resource| {
        matches!(
            resource,
            super::ResourceChange::SharedString { insert: false, .. }
                | super::ResourceChange::RichSharedString { insert: false, .. }
        )
    }) {
        let text = resource_text(resource)
            .ok_or_else(|| Error::InvalidData("SST resource has no text".into()))?;
        let index =
            entries
                .iter()
                .position(|entry| entry.text == text)
                .ok_or(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "SST resource removal identity is stale",
                })?;
        if !sst_resource_removal_matches(&entries[index], resource)? {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "SST resource removal identity is stale",
            });
        }
        if let Some(source_index) = entries[index].source_index {
            removed_source_indices.push(source_index);
        }
        entries.remove(index);
    }
    removed_source_indices.sort_unstable();
    if let Some(&first) = removed_source_indices.first() {
        let expected = source.inner.shared_strings.len().saturating_sub(first);
        if removed_source_indices.len() != expected
            || removed_source_indices
                .iter()
                .enumerate()
                .any(|(offset, index)| *index != first.saturating_add(offset))
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "SST resource removals must form a source-table suffix",
            ));
        }
    }
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(insert_count)
        .map_err(|_error| Error::Allocation("retaining inserted SST resources"))?;
    inserted.extend(resources.iter().filter(|resource| {
        matches!(
            resource,
            super::ResourceChange::SharedString { insert: true, .. }
                | super::ResourceChange::RichSharedString { insert: true, .. }
        )
    }));
    inserted.sort_by_cached_key(|resource| super::resource_target(resource));
    for resource in inserted {
        let (text, properties) = match resource {
            super::ResourceChange::SharedString { text, .. } => (text, None),
            super::ResourceChange::RichSharedString {
                text,
                formatting_runs,
                ..
            } => (
                text,
                Some(Box::new(crate::records::SharedStringProperties {
                    formatting_runs: formatting_runs.clone(),
                    phonetic: None,
                })),
            ),
            _ => unreachable!("filtered SST resource"),
        };
        entries.push(SstEntry {
            text: text.clone(),
            properties,
            source_index: None,
            raw: None,
        });
    }
    Ok(entries)
}

fn resource_text(resource: &super::ResourceChange) -> Option<&str> {
    match resource {
        super::ResourceChange::SharedString { text, .. }
        | super::ResourceChange::RichSharedString { text, .. } => Some(text),
        super::ResourceChange::ExtendedFormat { .. }
        | super::ResourceChange::FormulaCell { .. } => None,
    }
}

fn sst_resource_removal_matches(
    entry: &SstEntry,
    resource: &super::ResourceChange,
) -> Result<bool> {
    match resource {
        super::ResourceChange::SharedString { .. } => {
            Ok(entry.properties.is_none() && semantic_sst_entry_is_reversible(entry)?)
        },
        super::ResourceChange::RichSharedString {
            formatting_runs, ..
        } => canonical_rich_sst_entry_matches(entry, formatting_runs),
        super::ResourceChange::ExtendedFormat { .. }
        | super::ResourceChange::FormulaCell { .. } => Ok(false),
    }
}

fn canonical_rich_sst_entry_matches(
    entry: &SstEntry,
    formatting_runs: &[crate::records::SharedStringFormatRun],
) -> Result<bool> {
    let Some(properties) = entry.properties.as_deref() else {
        return Ok(entry.source_index.is_none() && entry.raw.is_none());
    };
    if properties.phonetic.is_some() || properties.formatting_runs != formatting_runs {
        return Ok(false);
    }
    let unit_count = entry.text.encode_utf16().count();
    let character_count = match u16::try_from(unit_count) {
        Ok(count) => count,
        Err(_) => return Ok(false),
    };
    let compressed = entry.text.encode_utf16().all(|unit| unit <= 0xff);
    let flags = u8::from(!compressed) | (1 << 3);
    let run_count = match u16::try_from(formatting_runs.len()) {
        Ok(count) => count,
        Err(_) => return Ok(false),
    };
    let mut prefix = Vec::new();
    prefix
        .try_reserve_exact(5)
        .map_err(|_error| Error::Allocation("checking rich SST resource identity"))?;
    prefix.extend_from_slice(&character_count.to_le_bytes());
    prefix.push(flags);
    prefix.extend_from_slice(&run_count.to_le_bytes());
    let Some(raw) = entry.raw.as_ref() else {
        return Ok(entry.source_index.is_none());
    };
    if raw.prefix != prefix || raw.extension.is_some() {
        return Ok(false);
    }

    let mut characters = Vec::new();
    characters
        .try_reserve_exact(if compressed {
            unit_count
        } else {
            unit_count.saturating_mul(2)
        })
        .map_err(|_error| Error::Allocation("checking rich SST characters"))?;
    for unit in entry.text.encode_utf16() {
        if compressed {
            characters.push(u8::try_from(unit).map_err(|_error| {
                Error::InvalidData("compressed rich SST character exceeds u8".into())
            })?);
        } else {
            characters.extend_from_slice(&unit.to_le_bytes());
        }
    }
    let mut offset = 0_usize;
    for segment in &raw.characters {
        if segment.wide != !compressed
            || characters.get(offset..offset.saturating_add(segment.bytes.len()))
                != Some(segment.bytes.as_slice())
        {
            return Ok(false);
        }
        offset = offset
            .checked_add(segment.bytes.len())
            .ok_or_else(|| Error::InvalidData("rich SST character offset overflows".into()))?;
    }
    if offset != characters.len() {
        return Ok(false);
    }

    let mut runs = Vec::new();
    runs.try_reserve_exact(formatting_runs.len().saturating_mul(4))
        .map_err(|_error| Error::Allocation("checking rich SST formatting runs"))?;
    for run in formatting_runs {
        runs.extend_from_slice(&run.character_index.to_le_bytes());
        runs.extend_from_slice(&run.font_index.to_le_bytes());
    }
    Ok(raw.formatting_runs == runs)
}

/// The worksheet-name rule of [MS-XLS] 2.4.28, which the reader enforces
/// ([`crate::records::sheet_name_fault`]).
pub(super) fn validate_sheet_name(name: &str) -> Result<()> {
    use crate::records::SheetNameFault;
    match crate::records::sheet_name_fault(name) {
        None => Ok(()),
        Some(SheetNameFault::Empty | SheetNameFault::TooLong { .. }) => Err(Error::InvalidData(
            "BIFF8 worksheet names must contain 1 through 31 UTF-16 units".into(),
        )),
        Some(SheetNameFault::ForbiddenCharacter('\0')) | Some(SheetNameFault::EdgeApostrophe) => {
            Err(Error::InvalidData(
                "BIFF8 worksheet name has a forbidden quote or NUL".into(),
            ))
        },
        Some(SheetNameFault::ForbiddenCharacter(_)) => Err(Error::InvalidData(
            "BIFF8 worksheet name contains a forbidden character".into(),
        )),
    }
}

pub(super) fn certify_sst_insertion(candidate: &super::ResourceChange) -> Result<()> {
    // Commit publishes every staged SST resource through
    // `compact_shared_strings`, which re-encodes the whole table and its
    // ExtSST and re-checks the source SST family.  The only refusal that
    // depends on the candidate alone is its own encoding; checking it here
    // keeps the refusal at the staging call, in O(text) rather than by
    // re-applying every staged insertion to a copy of the workbook.
    let (text, formatting_runs) = match candidate {
        super::ResourceChange::SharedString { text, .. } => (text.as_str(), &[][..]),
        super::ResourceChange::RichSharedString {
            text,
            formatting_runs,
            ..
        } => (text.as_str(), formatting_runs.as_slice()),
        super::ResourceChange::ExtendedFormat { .. }
        | super::ResourceChange::FormulaCell { .. } => {
            return Ok(());
        },
    };
    encode_sst_tail_records(text, formatting_runs).map(drop)
}

fn operation_sheet(change: &StructuralChange) -> usize {
    match change {
        StructuralChange::Cell { sheet, .. }
        | StructuralChange::Rows { sheet, .. }
        | StructuralChange::Columns { sheet, .. }
        | StructuralChange::RenameSheet { sheet, .. } => *sheet,
    }
}

pub(super) fn certify_shift(
    source: &super::Snapshot,
    sheet: usize,
    shift: AxisShift,
) -> Result<()> {
    let workbook = &source.inner.workbook_stream;
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet dependency index is stale",
        })?
        .workbook_index;
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(workbook.len())
        .map_err(|_error| Error::Allocation("certifying workbook dependencies"))?;
    candidate.extend_from_slice(workbook);
    let limits = source.raw_record_limits()?;
    patch_workbook_dependencies(&mut candidate, workbook_index, &limits, shift)?;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or(Error::Internal("worksheet dependency owner disappeared"))?;
    let start = usize::try_from(bound.position)
        .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
    let end = bounds
        .iter()
        .filter_map(|candidate| {
            let position = usize::try_from(candidate.position).ok()?;
            (position > start).then_some(position)
        })
        .min()
        .unwrap_or(workbook.len());
    let worksheet = workbook.get(start..end).ok_or_else(|| {
        Error::InvalidData("worksheet dependency range is outside Workbook".into())
    })?;
    certify_coordinate_shift(worksheet, &raw_records(worksheet)?, shift)
}

fn resource_insert(resource: &super::ResourceChange) -> bool {
    match resource {
        super::ResourceChange::SharedString { insert, .. }
        | super::ResourceChange::RichSharedString { insert, .. }
        | super::ResourceChange::ExtendedFormat { insert, .. }
        | super::ResourceChange::FormulaCell { insert, .. } => *insert,
    }
}

fn apply_resource(workbook: &mut Vec<u8>, resource: &super::ResourceChange) -> Result<()> {
    match resource {
        super::ResourceChange::SharedString { text, insert } => {
            apply_shared_string_resource(workbook, text, &[], *insert)
        },
        super::ResourceChange::RichSharedString {
            text,
            formatting_runs,
            insert,
        } => apply_shared_string_resource(workbook, text, formatting_runs, *insert),
        super::ResourceChange::ExtendedFormat {
            index,
            payload,
            insert,
        } => apply_xf_resource(workbook, *index, payload, *insert),
        super::ResourceChange::FormulaCell { .. } => Ok(()),
    }
}

fn workbook_globals(workbook: &[u8]) -> Result<Vec<RawRecord>> {
    let mut globals = Vec::new();
    for record in raw_records(workbook)? {
        let eof = record.kind == 0x000a;
        globals.push(record);
        if eof {
            break;
        }
    }
    Ok(globals)
}

fn apply_shared_string_resource(
    workbook: &mut Vec<u8>,
    text: &str,
    formatting_runs: &[crate::records::SharedStringFormatRun],
    insert: bool,
) -> Result<()> {
    let globals = workbook_globals(workbook)?;
    let sst_index = unique_kind(&globals, super::SST, "SST")?;
    let sst = &globals[sst_index];
    if sst.end - sst.start < 12 {
        return Err(Error::InvalidData("SST header is truncated".into()));
    }
    let unique = binary::read_u32_le_at(workbook, sst.start + 8)?;
    let mut family_end = sst.end;
    let mut last = sst;
    for record in &globals[sst_index + 1..] {
        if record.kind != CONTINUE {
            break;
        }
        family_end = record.end;
        last = record;
    }
    let ext_sst = globals
        .iter()
        .filter(|record| record.kind == EXT_SST)
        .try_fold(None, |found, record| {
            if found.is_some() {
                Err(Error::InvalidData("duplicate ExtSST record".into()))
            } else {
                Ok(Some(record.clone()))
            }
        })?;
    if let Some(ext_sst) = &ext_sst
        && ext_sst.start != family_end
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedContent,
            "edit cells",
            "ExtSST is not adjacent to its SST record family",
        ));
    }
    if insert {
        let records = encode_sst_tail_records(text, formatting_runs)?;
        let total = records
            .iter()
            .try_fold(0_usize, |total, record| total.checked_add(record.len()));
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(
                total.ok_or_else(|| Error::InvalidData("SST tail length overflow".into()))?,
            )
            .map_err(|_error| Error::Allocation("retaining continued SST tail"))?;
        for record in records {
            bytes.extend_from_slice(&record);
        }
        let updated = unique
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("SST unique count overflow".into()))?;
        let updated_ext_sst = ext_sst
            .as_ref()
            .map(|record| {
                let string_position = family_end
                    .checked_add(4)
                    .and_then(|value| u32::try_from(value).ok())
                    .ok_or_else(|| {
                        Error::InvalidData("ExtSST string position exceeds u32".into())
                    })?;
                update_ext_sst_record(workbook, record, unique, updated, Some(string_position))
            })
            .transpose()?;
        replace_range_and_adjust_bounds(workbook, family_end, family_end, &bytes)?;
        workbook[sst.start + 8..sst.start + 12].copy_from_slice(&updated.to_le_bytes());
        if let (Some(record), Some(updated_record)) = (&ext_sst, updated_ext_sst) {
            let start = record
                .start
                .checked_add(bytes.len())
                .ok_or_else(|| Error::InvalidData("shifted ExtSST offset overflow".into()))?;
            let end = record
                .end
                .checked_add(bytes.len())
                .ok_or_else(|| Error::InvalidData("shifted ExtSST range overflow".into()))?;
            replace_range_and_adjust_bounds(workbook, start, end, &updated_record)?;
        }
        return Ok(());
    }
    let expected = encode_sst_tail_records(text, formatting_runs)?;
    let authored_count = expected.len();
    let family: Vec<_> = globals[sst_index + 1..]
        .iter()
        .take_while(|record| record.kind == CONTINUE)
        .collect();
    if authored_count == 0 || family.len() < authored_count {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "SST inverse cannot locate its complete tail resource",
        ));
    }
    let tail = &family[family.len() - authored_count..];
    for (record, expected) in tail.iter().zip(&expected) {
        if &workbook[record.start..record.end] != expected.as_slice() {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "SST inverse tail resource precondition is stale",
            });
        }
    }
    let removed_index = unique
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidData("SST unique count underflow".into()))?;
    if label_sst_reference_exists(workbook, removed_index)? {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "SST inverse resource is still referenced by a LabelSst cell",
        ));
    }
    let updated_ext_sst = ext_sst
        .as_ref()
        .map(|record| update_ext_sst_record(workbook, record, unique, removed_index, None))
        .transpose()?;
    let tail_start = tail
        .first()
        .ok_or_else(|| Error::InvalidData("SST inverse tail is empty".into()))?
        .start;
    let removed_len = last
        .end
        .checked_sub(tail_start)
        .ok_or_else(|| Error::InvalidData("SST inverse tail range is reversed".into()))?;
    workbook[sst.start + 8..sst.start + 12].copy_from_slice(&removed_index.to_le_bytes());
    replace_range_and_adjust_bounds(workbook, tail_start, last.end, &[])?;
    if let (Some(record), Some(updated_record)) = (&ext_sst, updated_ext_sst) {
        let start = record
            .start
            .checked_sub(removed_len)
            .ok_or_else(|| Error::InvalidData("restored ExtSST offset underflow".into()))?;
        let end = record
            .end
            .checked_sub(removed_len)
            .ok_or_else(|| Error::InvalidData("restored ExtSST range underflow".into()))?;
        replace_range_and_adjust_bounds(workbook, start, end, &updated_record)?;
    }
    Ok(())
}

fn update_ext_sst_record(
    workbook: &[u8],
    record: &RawRecord,
    current_unique: u32,
    target_unique: u32,
    appended_string_position: Option<u32>,
) -> Result<Vec<u8>> {
    let payload_start = record
        .start
        .checked_add(4)
        .ok_or_else(|| Error::InvalidData("ExtSST payload offset overflow".into()))?;
    let payload = workbook
        .get(payload_start..record.end)
        .ok_or_else(|| Error::InvalidData("ExtSST payload is truncated".into()))?;
    let current =
        crate::shared_string_index::SharedStringIndex::parse_payload(current_unique, payload)?;
    let target_bucket_size = ext_sst_bucket_size(target_unique)?;
    if target_bucket_size != current.strings_per_bucket() {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "SST authoring crosses an ExtSST bucket-size boundary",
        ));
    }
    let required = ext_sst_bucket_count(target_unique, target_bucket_size)?;
    let mut buckets = current.buckets().to_vec();
    match required.cmp(&buckets.len()) {
        std::cmp::Ordering::Equal => {},
        std::cmp::Ordering::Greater if required == buckets.len().saturating_add(1) => {
            let position = appended_string_position.ok_or_else(|| {
                Error::InvalidData("ExtSST insertion has no appended string position".into())
            })?;
            buckets.push(crate::shared_string_index::SharedStringBucket::try_new(
                position, 4,
            )?);
        },
        std::cmp::Ordering::Less if buckets.len() == required.saturating_add(1) => {
            buckets.pop();
        },
        _ => {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "SST authoring changes unsupported ExtSST bucket geometry",
            ));
        },
    }
    crate::shared_string_index::SharedStringIndex::try_new(target_unique, buckets)?
        .to_record_bytes()
}

fn ext_sst_bucket_size(unique: u32) -> Result<u16> {
    let size = (unique / 128 + 1).max(8);
    u16::try_from(size)
        .map_err(|_error| Error::InvalidData("ExtSST bucket size exceeds u16".into()))
}

fn ext_sst_bucket_count(unique: u32, size: u16) -> Result<usize> {
    if size == 0 {
        return Err(Error::InvalidData("ExtSST bucket size is zero".into()));
    }
    let count = if unique == 0 {
        0
    } else {
        (unique - 1) / u32::from(size) + 1
    };
    usize::try_from(count)
        .map_err(|_error| Error::InvalidData("ExtSST bucket count exceeds usize".into()))
}

fn encode_sst_tail_records(
    text: &str,
    formatting_runs: &[crate::records::SharedStringFormatRun],
) -> Result<Vec<Vec<u8>>> {
    let unit_count = text.encode_utf16().count();
    let count = u16::try_from(unit_count).map_err(|_error| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "shared string exceeds u16 characters",
        )
    })?;
    let mut units = Vec::new();
    units
        .try_reserve_exact(unit_count)
        .map_err(|_error| Error::Allocation("retaining shared-string UTF-16 units"))?;
    units.extend(text.encode_utf16());
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let width = if compressed { 1 } else { 2 };
    let mut offset = 0_usize;
    let mut first = true;
    let mut records = Vec::new();
    records
        .try_reserve(1 + unit_count / 4_111)
        .map_err(|_error| Error::Allocation("retaining continued SST records"))?;
    loop {
        let prefix = if first {
            3 + if formatting_runs.is_empty() { 0 } else { 2 }
        } else {
            1
        };
        let capacity = (8_224 - prefix) / width;
        let mut end = units.len().min(offset.saturating_add(capacity));
        if end < units.len()
            && end > offset
            && (0xd800..=0xdbff).contains(&units[end - 1])
            && (0xdc00..=0xdfff).contains(&units[end])
        {
            end -= 1;
        }
        let payload_len = prefix
            .checked_add((end - offset).saturating_mul(width))
            .ok_or_else(|| Error::InvalidData("continued SST payload size overflows".into()))?;
        let mut payload = Vec::new();
        payload
            .try_reserve_exact(payload_len)
            .map_err(|_error| Error::Allocation("retaining continued SST payload"))?;
        if first {
            payload.extend_from_slice(&count.to_le_bytes());
        }
        payload.push(
            u8::from(!compressed)
                | if first && !formatting_runs.is_empty() {
                    1 << 3
                } else {
                    0
                },
        );
        if first && !formatting_runs.is_empty() {
            payload.extend_from_slice(
                &u16::try_from(formatting_runs.len())
                    .map_err(|_error| Error::InvalidData("SST rich-run count exceeds u16".into()))?
                    .to_le_bytes(),
            );
        }
        for unit in &units[offset..end] {
            if compressed {
                payload.push(u8::try_from(*unit).map_err(|_error| {
                    Error::InvalidData("compressed shared-string unit exceeds u8".into())
                })?);
            } else {
                payload.extend_from_slice(&unit.to_le_bytes());
            }
        }
        let record_len = 4usize
            .checked_add(payload.len())
            .ok_or_else(|| Error::InvalidData("continued SST record size overflows".into()))?;
        let mut record = Vec::new();
        record
            .try_reserve_exact(record_len)
            .map_err(|_error| Error::Allocation("retaining continued SST record"))?;
        record.extend_from_slice(&CONTINUE.to_le_bytes());
        record.extend_from_slice(
            &u16::try_from(payload.len())
                .map_err(|_error| Error::InvalidData("SST string payload exceeds u16".into()))?
                .to_le_bytes(),
        );
        record.extend_from_slice(&payload);
        records.push(record);
        if end == units.len() {
            break;
        }
        offset = end;
        first = false;
    }
    if !formatting_runs.is_empty() {
        records
            .try_reserve(1 + formatting_runs.len() / 2_056)
            .map_err(|_error| Error::Allocation("retaining rich SST Continue records"))?;
        let mut run_bytes = Vec::new();
        run_bytes
            .try_reserve_exact(formatting_runs.len().saturating_mul(4))
            .map_err(|_error| Error::Allocation("retaining SST formatting runs"))?;
        for run in formatting_runs {
            run_bytes.extend_from_slice(&run.character_index.to_le_bytes());
            run_bytes.extend_from_slice(&run.font_index.to_le_bytes());
        }
        let mut offset = 0_usize;
        while offset < run_bytes.len() {
            let last = records.last_mut().ok_or_else(|| {
                Error::InvalidData("SST rich string has no character record".into())
            })?;
            let available = 8_228_usize.saturating_sub(last.len());
            let aligned = available - available % 4;
            if aligned == 0 {
                let record_len = 8_228.min(
                    4usize
                        .checked_add(run_bytes.len() - offset)
                        .ok_or_else(|| {
                            Error::InvalidData("SST formatting record size overflows".into())
                        })?,
                );
                let mut record = Vec::new();
                record
                    .try_reserve_exact(record_len)
                    .map_err(|_error| Error::Allocation("retaining SST formatting record"))?;
                record.extend_from_slice(&CONTINUE.to_le_bytes());
                record.extend_from_slice(&0_u16.to_le_bytes());
                records.push(record);
                continue;
            }
            let count = aligned.min(run_bytes.len() - offset);
            last.extend_from_slice(&run_bytes[offset..offset + count]);
            let payload_len = u16::try_from(last.len() - 4)
                .map_err(|_error| Error::InvalidData("SST rich payload exceeds u16".into()))?;
            last[2..4].copy_from_slice(&payload_len.to_le_bytes());
            offset += count;
        }
    }
    Ok(records)
}

struct EncodedSst {
    records: Vec<Vec<u8>>,
    locations: Vec<(usize, u16)>,
}

impl EncodedSst {
    fn bytes_len(&self) -> usize {
        self.records.iter().map(Vec::len).sum()
    }

    fn ext_sst_index(
        &self,
        sst_start: usize,
        unique: u32,
    ) -> Result<crate::shared_string_index::SharedStringIndex> {
        let bucket_size = ext_sst_bucket_size(unique)?;
        let mut buckets = Vec::new();
        buckets
            .try_reserve((self.locations.len() / usize::from(bucket_size)).saturating_add(1))
            .map_err(|_error| Error::Allocation("retaining ExtSST buckets"))?;
        let mut index = 0_usize;
        while index < self.locations.len() {
            let (relative, record_offset) = self.locations[index];
            let absolute = sst_start
                .checked_add(relative)
                .and_then(|value| u32::try_from(value).ok())
                .ok_or_else(|| Error::InvalidData("ExtSST string position exceeds u32".into()))?;
            buckets.push(crate::shared_string_index::SharedStringBucket::try_new(
                absolute,
                record_offset,
            )?);
            index = index.saturating_add(usize::from(bucket_size));
        }
        crate::shared_string_index::SharedStringIndex::try_new(unique, buckets)
    }
}

fn encode_sst_records(entries: &[SstEntry], total: u32) -> Result<EncodedSst> {
    let unique = u32::try_from(entries.len())
        .map_err(|_error| Error::InvalidData("SST unique count exceeds u32".into()))?;
    let mut records = Vec::new();
    records
        .try_reserve(entries.len().saturating_div(2).saturating_add(1))
        .map_err(|_error| Error::Allocation("retaining compacted SST frames"))?;
    let mut locations = Vec::new();
    locations
        .try_reserve_exact(entries.len())
        .map_err(|_error| Error::Allocation("retaining compacted SST locations"))?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(MAX_SST_PAYLOAD)
        .map_err(|_error| Error::Allocation("retaining compacted SST payload"))?;
    payload.extend_from_slice(&total.to_le_bytes());
    payload.extend_from_slice(&unique.to_le_bytes());
    let mut first_record = true;
    let mut record_start = 0_usize;

    for entry in entries {
        if let Some(raw) = entry.raw.as_ref() {
            append_raw_sst_entry(
                raw,
                &mut records,
                &mut locations,
                &mut payload,
                &mut first_record,
                &mut record_start,
            )?;
            continue;
        }
        let unit_count = entry.text.encode_utf16().count();
        let mut units = Vec::new();
        units
            .try_reserve_exact(unit_count)
            .map_err(|_error| Error::Allocation("retaining compacted SST UTF-16 units"))?;
        units.extend(entry.text.encode_utf16());
        let character_count = u16::try_from(units.len()).map_err(|_error| {
            Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "shared string exceeds u16 characters",
            )
        })?;
        let compressed = units.iter().all(|unit| *unit <= 0xff);
        let runs = entry
            .properties
            .as_deref()
            .map_or(&[][..], |properties| properties.formatting_runs.as_slice());
        let ext = entry
            .properties
            .as_deref()
            .and_then(|properties| properties.phonetic.as_ref())
            .map(encode_phonetic_extension)
            .transpose()?;
        let run_count = u16::try_from(runs.len()).map_err(|_error| {
            Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "SST rich-run count exceeds u16",
            )
        })?;
        let ext_len = ext.as_ref().map_or(0, Vec::len);
        let ext_len_u32 = u32::try_from(ext_len)
            .map_err(|_error| Error::InvalidData("SST ExtRst length exceeds u32".into()))?;
        let header_len = 3_usize
            .checked_add(if runs.is_empty() { 0 } else { 2 })
            .and_then(|length| length.checked_add(if ext.is_some() { 4 } else { 0 }))
            .ok_or_else(|| Error::InvalidData("SST string header length overflow".into()))?;
        if payload.len().saturating_add(header_len) > MAX_SST_PAYLOAD {
            flush_sst_record(
                &mut records,
                &mut payload,
                &mut first_record,
                &mut record_start,
            )?;
        }
        if payload.len().saturating_add(header_len) > MAX_SST_PAYLOAD {
            return Err(Error::InvalidData(
                "SST string header exceeds record payload".into(),
            ));
        }
        let string_start = record_start
            .checked_add(4)
            .and_then(|value| value.checked_add(payload.len()))
            .ok_or_else(|| Error::InvalidData("SST string position overflows".into()))?;
        let record_offset = u16::try_from(4usize.saturating_add(payload.len()))
            .map_err(|_error| Error::InvalidData("SST record offset exceeds u16".into()))?;
        locations.push((string_start, record_offset));
        payload.extend_from_slice(&character_count.to_le_bytes());
        let flags = u8::from(!compressed)
            | if !runs.is_empty() { 1 << 3 } else { 0 }
            | if ext.is_some() { 1 << 2 } else { 0 };
        payload.push(flags);
        if !runs.is_empty() {
            payload.extend_from_slice(&run_count.to_le_bytes());
        }
        if ext.is_some() {
            payload.extend_from_slice(&ext_len_u32.to_le_bytes());
        }

        let width = if compressed { 1 } else { 2 };
        let mut offset = 0_usize;
        while offset < units.len() {
            let available = MAX_SST_PAYLOAD.saturating_sub(payload.len());
            let mut count = available / width;
            count = count.min(units.len().saturating_sub(offset));
            if count == 0 {
                flush_sst_record(
                    &mut records,
                    &mut payload,
                    &mut first_record,
                    &mut record_start,
                )?;
                payload.push(u8::from(!compressed));
                continue;
            }
            if !compressed && count < units.len() - offset && count > 0 {
                let end = offset + count;
                if end < units.len()
                    && (0xd800..=0xdbff).contains(&units[end - 1])
                    && (0xdc00..=0xdfff).contains(&units[end])
                {
                    count -= 1;
                }
                if count == 0 {
                    flush_sst_record(
                        &mut records,
                        &mut payload,
                        &mut first_record,
                        &mut record_start,
                    )?;
                    payload.push(1);
                    continue;
                }
            }
            if compressed {
                for unit in &units[offset..offset + count] {
                    payload.push(u8::try_from(*unit).map_err(|_error| {
                        Error::InvalidData("compressed SST unit exceeds u8".into())
                    })?);
                }
            } else {
                for unit in &units[offset..offset + count] {
                    payload.extend_from_slice(&unit.to_le_bytes());
                }
            }
            offset += count;
            if offset < units.len() {
                flush_sst_record(
                    &mut records,
                    &mut payload,
                    &mut first_record,
                    &mut record_start,
                )?;
                payload.push(u8::from(!compressed));
            }
        }
        if !runs.is_empty() {
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(runs.len().saturating_mul(4))
                .map_err(|_error| Error::Allocation("retaining SST formatting runs"))?;
            for run in runs {
                bytes.extend_from_slice(&run.character_index.to_le_bytes());
                bytes.extend_from_slice(&run.font_index.to_le_bytes());
            }
            append_sst_bytes(
                &mut records,
                &mut payload,
                &mut first_record,
                &mut record_start,
                &bytes,
            )?;
        }
        if let Some(ext) = ext {
            append_sst_bytes(
                &mut records,
                &mut payload,
                &mut first_record,
                &mut record_start,
                &ext,
            )?;
        }
    }
    flush_sst_record(
        &mut records,
        &mut payload,
        &mut first_record,
        &mut record_start,
    )?;
    Ok(EncodedSst { records, locations })
}

fn append_raw_sst_entry(
    entry: &RawSstEntry,
    records: &mut Vec<Vec<u8>>,
    locations: &mut Vec<(usize, u16)>,
    payload: &mut Vec<u8>,
    first_record: &mut bool,
    record_start: &mut usize,
) -> Result<()> {
    if entry.prefix.len() > MAX_SST_PAYLOAD {
        return Err(Error::InvalidData(
            "source SST entry header exceeds record payload".into(),
        ));
    }
    if payload.len().saturating_add(entry.prefix.len()) > MAX_SST_PAYLOAD {
        flush_sst_record(records, payload, first_record, record_start)?;
    }
    let string_start = record_start
        .checked_add(4)
        .and_then(|value| value.checked_add(payload.len()))
        .ok_or_else(|| Error::InvalidData("SST string position overflows".into()))?;
    let record_offset = u16::try_from(4usize.saturating_add(payload.len()))
        .map_err(|_error| Error::InvalidData("SST record offset exceeds u16".into()))?;
    locations.push((string_start, record_offset));
    payload.extend_from_slice(&entry.prefix);
    append_raw_sst_characters(
        records,
        payload,
        first_record,
        record_start,
        &entry.characters,
    )?;
    append_sst_bytes(
        records,
        payload,
        first_record,
        record_start,
        &entry.formatting_runs,
    )?;
    if let Some(extension) = &entry.extension {
        append_sst_bytes(records, payload, first_record, record_start, extension)?;
    }
    Ok(())
}

fn append_raw_sst_characters(
    records: &mut Vec<Vec<u8>>,
    payload: &mut Vec<u8>,
    first_record: &mut bool,
    record_start: &mut usize,
    segments: &[RawSstCharacters],
) -> Result<()> {
    for (segment_index, segment) in segments.iter().enumerate() {
        if segment_index != 0 {
            flush_sst_record(records, payload, first_record, record_start)?;
            payload.push(u8::from(segment.wide));
        }
        let mut offset = 0_usize;
        while offset < segment.bytes.len() {
            if payload.len() == MAX_SST_PAYLOAD {
                flush_sst_record(records, payload, first_record, record_start)?;
                payload.push(u8::from(segment.wide));
            }
            let mut count = (MAX_SST_PAYLOAD - payload.len()).min(segment.bytes.len() - offset);
            if segment.wide {
                count -= count % 2;
            }
            if count == 0 {
                flush_sst_record(records, payload, first_record, record_start)?;
                payload.push(u8::from(segment.wide));
                continue;
            }
            payload.extend_from_slice(&segment.bytes[offset..offset + count]);
            offset += count;
        }
    }
    Ok(())
}

fn flush_sst_record(
    records: &mut Vec<Vec<u8>>,
    payload: &mut Vec<u8>,
    first_record: &mut bool,
    record_start: &mut usize,
) -> Result<()> {
    if payload.is_empty() {
        return Ok(());
    }
    let length = u16::try_from(payload.len())
        .map_err(|_error| Error::InvalidData("SST record payload exceeds u16".into()))?;
    let kind = if *first_record { super::SST } else { CONTINUE };
    let record_len = 4usize
        .checked_add(payload.len())
        .ok_or_else(|| Error::InvalidData("SST record size overflows".into()))?;
    let mut record = Vec::new();
    record
        .try_reserve_exact(record_len)
        .map_err(|_error| Error::Allocation("retaining SST record"))?;
    record.extend_from_slice(&kind.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(payload);
    *record_start = record_start
        .checked_add(record.len())
        .ok_or_else(|| Error::InvalidData("SST record position overflows".into()))?;
    records
        .try_reserve(1)
        .map_err(|_error| Error::Allocation("retaining SST records"))?;
    records.push(record);
    payload.clear();
    *first_record = false;
    Ok(())
}

fn append_sst_bytes(
    records: &mut Vec<Vec<u8>>,
    payload: &mut Vec<u8>,
    first_record: &mut bool,
    record_start: &mut usize,
    bytes: &[u8],
) -> Result<()> {
    let mut offset = 0_usize;
    while offset < bytes.len() {
        if payload.len() == MAX_SST_PAYLOAD {
            flush_sst_record(records, payload, first_record, record_start)?;
        }
        let count = (MAX_SST_PAYLOAD - payload.len()).min(bytes.len() - offset);
        payload.extend_from_slice(&bytes[offset..offset + count]);
        offset += count;
    }
    Ok(())
}

fn encode_phonetic_extension(phonetic: &crate::records::PhoneticString) -> Result<Vec<u8>> {
    let text_units: Vec<u16> = phonetic.text.encode_utf16().collect();
    let text_count = u16::try_from(text_units.len())
        .map_err(|_error| Error::InvalidData("ExtRst text exceeds u16 characters".into()))?;
    if text_count > 32_767 {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "ExtRst text exceeds signed character limit",
        ));
    }
    let run_count = u16::try_from(phonetic.runs.len())
        .map_err(|_error| Error::InvalidData("ExtRst run count exceeds u16".into()))?;
    if run_count > 32_767 {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "ExtRst run count exceeds signed limit",
        ));
    }
    let mut data = Vec::new();
    data.try_reserve(
        14usize
            .saturating_add(text_units.len().saturating_mul(2))
            .saturating_add(phonetic.runs.len().saturating_mul(6))
            .saturating_add(phonetic.extra_data.len()),
    )
    .map_err(|_error| Error::Allocation("retaining SST ExtRst"))?;
    data.extend_from_slice(&0_u16.to_le_bytes());
    data.extend_from_slice(&0_u16.to_le_bytes());
    data.extend_from_slice(&phonetic.font_index.to_le_bytes());
    let kind: u16 = match phonetic.phonetic_type {
        crate::records::PhoneticType::NarrowKatakana => 0,
        crate::records::PhoneticType::WideKatakana => 1,
        crate::records::PhoneticType::Hiragana => 2,
        crate::records::PhoneticType::Any => 3,
    };
    let alignment: u16 = match phonetic.alignment {
        crate::records::PhoneticAlignment::General => 0,
        crate::records::PhoneticAlignment::Left => 1,
        crate::records::PhoneticAlignment::Center => 2,
        crate::records::PhoneticAlignment::Distributed => 3,
    };
    data.extend_from_slice(&(kind | (alignment << 2)).to_le_bytes());
    data.extend_from_slice(&run_count.to_le_bytes());
    data.extend_from_slice(&text_count.to_le_bytes());
    data.extend_from_slice(&text_count.to_le_bytes());
    for unit in text_units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    for run in &phonetic.runs {
        data.extend_from_slice(&run.phonetic_text_index.to_le_bytes());
        data.extend_from_slice(&run.base_text_index.to_le_bytes());
        data.extend_from_slice(&run.base_text_length.to_le_bytes());
    }
    data.extend_from_slice(&phonetic.extra_data);
    let payload_length = u16::try_from(data.len())
        .map_err(|_error| Error::InvalidData("ExtRst payload exceeds u16".into()))?;
    data[2..4].copy_from_slice(&payload_length.to_le_bytes());
    Ok(data)
}

fn label_sst_reference_exists(workbook: &[u8], index: u32) -> Result<bool> {
    for record in raw_records(workbook)? {
        if record.kind == LABEL_SST && binary::read_u32_le_at(workbook, record.start + 10)? == index
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn apply_xf_resource(
    workbook: &mut Vec<u8>,
    index: StyleIndex,
    payload: &[u8],
    insert: bool,
) -> Result<()> {
    super::validate_xf_payload(payload)?;
    let globals = workbook_globals(workbook)?;
    let xfs: Vec<_> = globals.iter().filter(|record| record.kind == XF).collect();
    let crc_record = globals
        .iter()
        .find(|record| record.kind == crate::number_format::XFCRC_RECORD);
    if let Some(record) = crc_record {
        // A stale CRC can mean the retained XFExt records no longer describe
        // these XFs. Do not silently validate those extensions by repairing it.
        let mut crc = litchi_core::mso_crc32::MsoCrc32::new();
        for xf in &xfs {
            crc.update(&workbook[xf.start + 4..xf.end]);
        }
        if record.end - record.start != 24
            || binary::read_u32_le_at(workbook, record.start + 20)? != crc.finalize()
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "XFCRC does not validate the source XF table",
            ));
        }
        if insert && xfs.len() >= crate::writer::formatting::MAX_XF_RECORDS_WITH_XFCRC {
            return Err(Error::TooMany {
                collection: "XF records of a workbook with XFCRC",
                limit: crate::writer::formatting::MAX_XF_RECORDS_WITH_XFCRC,
            });
        }
    }
    if insert {
        if usize::from(index.get()) != xfs.len() {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "new XF index is not the next resource",
            ));
        }
        let last = xfs.last().ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "workbook has no XF insertion owner",
            )
        })?;
        let mut record = Vec::with_capacity(4 + payload.len());
        record.extend_from_slice(&XF.to_le_bytes());
        record.extend_from_slice(
            &u16::try_from(payload.len())
                .map_err(|_error| Error::InvalidData("XF payload exceeds u16".into()))?
                .to_le_bytes(),
        );
        record.extend_from_slice(payload);
        replace_range_and_adjust_bounds(workbook, last.end, last.end, &record)?;
        return refresh_xfcrc(workbook);
    }
    if usize::from(index.get()) + 1 != xfs.len() {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "XF inverse can remove only the final resource",
        ));
    }
    let last = xfs.last().ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "workbook has no XF removal owner",
        )
    })?;
    if &workbook[last.start + 4..last.end] != payload {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "XF inverse payload is stale",
        });
    }
    if xf_reference_exists(workbook, index.get())? {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "XF inverse resource remains referenced by a cell",
        ));
    }
    replace_range_and_adjust_bounds(workbook, last.start, last.end, &[])?;
    refresh_xfcrc(workbook)
}

/// MS-XLS 2.4.354: checksum only the XF payloads using MS-OSHARED 2.4.3.
/// Keep the existing future-record header and every XFExt byte unchanged.
fn refresh_xfcrc(workbook: &mut [u8]) -> Result<()> {
    let globals = workbook_globals(workbook)?;
    let Some(record) = globals
        .iter()
        .find(|record| record.kind == crate::number_format::XFCRC_RECORD)
    else {
        return Ok(());
    };
    let mut crc = litchi_core::mso_crc32::MsoCrc32::new();
    let mut count = 0_u16;
    for xf in globals.iter().filter(|record| record.kind == XF) {
        count += 1;
        crc.update(&workbook[xf.start + 4..xf.end]);
    }
    workbook[record.start + 18..record.start + 20].copy_from_slice(&count.to_le_bytes());
    workbook[record.start + 20..record.end].copy_from_slice(&crc.finalize().to_le_bytes());
    Ok(())
}

fn xf_reference_exists(workbook: &[u8], index: u16) -> Result<bool> {
    for record in raw_records(workbook)? {
        if record.kind == MUL_RK {
            let payload = &workbook[record.start + 4..record.end];
            let count = (payload.len().saturating_sub(6)) / 6;
            for item in 0..count {
                if binary::read_u16_le_at(payload, 4 + item * 6)? == index {
                    return Ok(true);
                }
            }
        } else if record.kind == MUL_BLANK {
            let payload = &workbook[record.start + 4..record.end];
            let count = payload.len().saturating_sub(6) / 2;
            for item in 0..count {
                if binary::read_u16_le_at(payload, 4 + item * 2)? == index {
                    return Ok(true);
                }
            }
        } else if is_cell_record(record.kind)
            && binary::read_u16_le_at(workbook, record.start + 8)? == index
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn rename_sheet(
    workbook: &mut Vec<u8>,
    source: &super::Snapshot,
    sheet: usize,
    before: &str,
    after: &str,
) -> Result<()> {
    validate_sheet_name(after)?;
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet rename index is stale",
        })?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or(Error::Internal("worksheet BoundSheet disappeared"))?;
    let payload = &workbook[bound.record.start + 4..bound.record.end];
    let decoded = decode_bound_name(payload)?;
    if decoded != before {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet rename precondition is stale",
        });
    }
    let encoded = encode_bound_name(after)?;
    let payload_len = 7usize
        .checked_add(encoded.len())
        .ok_or_else(|| Error::InvalidData("BoundSheet length overflow".into()))?;
    let payload_len = u16::try_from(payload_len)
        .map_err(|_error| Error::InvalidData("BoundSheet payload exceeds u16".into()))?;
    let mut replacement = Vec::with_capacity(4 + usize::from(payload_len));
    replacement.extend_from_slice(&BOUND_SHEET.to_le_bytes());
    replacement.extend_from_slice(&payload_len.to_le_bytes());
    replacement.extend_from_slice(&payload[..6]);
    replacement.push(
        u8::try_from(after.encode_utf16().count())
            .map_err(|_error| Error::InvalidData("worksheet name length exceeds u8".into()))?,
    );
    replacement.extend_from_slice(&encoded);
    replace_range_and_adjust_bounds(workbook, bound.record.start, bound.record.end, &replacement)
}

fn rewrite_worksheet(
    workbook: &mut Vec<u8>,
    source: &super::Snapshot,
    sheet: usize,
    operations: &[&StructuralChange],
    resources: &[super::ResourceChange],
    layout: &[&super::layout::LayoutChange],
    shared_strings: &[String],
) -> Result<()> {
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet operation index is stale",
        })?
        .workbook_index;
    let limits = source.raw_record_limits()?;
    for operation in operations {
        if let Some(shift) = axis_shift(operation) {
            patch_workbook_dependencies(workbook, workbook_index, &limits, shift)?;
        }
    }
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or(Error::Internal("worksheet BoundSheet disappeared"))?;
    let start = usize::try_from(bound.position)
        .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
    let end = bounds
        .iter()
        .filter_map(|candidate| {
            let position = usize::try_from(candidate.position).ok()?;
            (position > start).then_some(position)
        })
        .min()
        .unwrap_or(workbook.len());
    let worksheet = workbook
        .get(start..end)
        .ok_or_else(|| Error::InvalidData("worksheet range is outside Workbook".into()))?;
    let replacement = rebuild_sheet(
        worksheet,
        start,
        sheet,
        source,
        operations,
        resources,
        layout,
        shared_strings,
    )?;
    replace_range_and_adjust_bounds(workbook, start, end, &replacement)
}

/// The workbook's sheet-reference table: the self-referencing `SupBook`
/// (MS-XLS 2.4.271) and every `XTI` of `ExternSheet` (2.4.105) as
/// `(SupBook index, first tab, last tab)`.
struct ExternSheets {
    self_book: Option<u16>,
    entries: Vec<(u16, i16, i16)>,
}

const SUP_BOOK: u16 = 0x01ae;
const EXTERN_SHEET: u16 = 0x0017;
const EXTERN_NAME: u16 = 0x0023;
const LBL: u16 = 0x0018;
const BRAI: u16 = 0x1051;
const EOF: u16 = 0x000a;

fn extern_sheets(workbook: &[u8], records: &[RawRecord]) -> Result<ExternSheets> {
    let mut sheets = ExternSheets {
        self_book: None,
        entries: Vec::new(),
    };
    let mut books = 0_u16;
    for (position, record) in records.iter().enumerate() {
        match record.kind {
            EOF => break,
            SUP_BOOK => {
                let payload = record_payload(workbook, record, "SupBook")?;
                if payload.len() == 4 && payload[2..4] == [0x01, 0x04] {
                    sheets.self_book.get_or_insert(books);
                }
                books = books
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("SupBook count overflows".into()))?;
            },
            EXTERN_SHEET => {
                if records
                    .get(position + 1)
                    .is_some_and(|next| next.kind == CONTINUE)
                {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "row/column movement cannot close a continued ExternSheet record",
                    ));
                }
                let payload = record_payload(workbook, record, "ExternSheet")?;
                let count = usize::from(binary::read_u16_le_at(payload, 0)?);
                if payload.len() != 2 + count * 6 {
                    return Err(Error::InvalidData(
                        "ExternSheet count does not match its payload".into(),
                    ));
                }
                for entry in payload[2..].as_chunks::<6>().0 {
                    sheets.entries.push((
                        u16::from_le_bytes([entry[0], entry[1]]),
                        i16::from_le_bytes([entry[2], entry[3]]),
                        i16::from_le_bytes([entry[4], entry[5]]),
                    ));
                }
            },
            _ => {},
        }
    }
    Ok(sheets)
}

/// Where a 3-D reference points relative to the shifted worksheet.
enum Target3d {
    /// Another sheet or another workbook: the shift does not touch it.
    Elsewhere,
    /// Exactly the shifted worksheet.
    Shifted,
}

fn target_3d(sheets: &ExternSheets, extern_sheet: u16, tab: usize) -> Result<Target3d> {
    let &(book, first, last) = sheets
        .entries
        .get(usize::from(extern_sheet))
        .ok_or_else(|| Error::InvalidData("3-D reference names a missing XTI".into()))?;
    if Some(book) != sheets.self_book || first < 0 || last < 0 {
        return Ok(Target3d::Elsewhere);
    }
    let (first, last) = (first as usize, last as usize);
    if tab < first.min(last) || tab > first.max(last) {
        return Ok(Target3d::Elsewhere);
    }
    if first != last {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot close a 3-D reference spanning several sheets",
        ));
    }
    Ok(Target3d::Shifted)
}

/// Shifts the 3-D references of `tokens` that name the shifted worksheet.
/// `name_formula` marks a `NameParsedFormula`, whose relative references are
/// offsets (`RgceLocRel`), which a shift cannot move: they are refused.
fn patch_3d_references(
    tokens: &mut [u8],
    sheets: &ExternSheets,
    tab: usize,
    shift: AxisShift,
    name_formula: bool,
) -> Result<()> {
    let references = crate::formula::formula_reference_spans(tokens).ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot read a workbook formula's references",
        )
    })?;
    if name_formula && !references.same_sheet.is_empty() {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot close a defined name with a sheet-less reference",
        ));
    }
    for reference in references.three_d {
        if matches!(
            target_3d(sheets, reference.extern_sheet, tab)?,
            Target3d::Elsewhere
        ) {
            continue;
        }
        let Some(span) = reference.reference else {
            continue;
        };
        if name_formula {
            let flags = match span {
                crate::formula::FormulaReferenceSpan::Cell { column, .. } => {
                    binary::read_u16_le_at(tokens, column)?
                },
                crate::formula::FormulaReferenceSpan::Area {
                    first_column,
                    last_column,
                    ..
                } => {
                    binary::read_u16_le_at(tokens, first_column)?
                        | binary::read_u16_le_at(tokens, last_column)?
                },
            };
            if flags & 0xc000 != 0 {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement cannot close a relative reference in a defined name",
                ));
            }
        }
        patch_reference_span(tokens, span, shift)?;
    }
    Ok(())
}

/// Shifts the workbook-wide dependencies of one worksheet's row or column
/// shift in place, and refuses those it cannot close.
///
/// Defined names (`Lbl`), chart series links (`BRAI`) and cell formulas
/// anywhere in the stream keep 3-D references into the shifted worksheet;
/// their coordinate fields are rewritten at fixed width, so no record moves.
/// External names belong to another workbook's `SupBook` and are untouched.
/// Records whose coordinates this owner does not model stay refusals.
fn patch_workbook_dependencies(
    workbook: &mut [u8],
    tab: usize,
    limits: &SourceLimits,
    shift: AxisShift,
) -> Result<()> {
    // AutoFilter keeps column positions relative to a filter range this
    // owner does not shift; it stays refused in the shifted worksheet and
    // wherever a 3-D reference can reach it.
    const GLOBAL_OR_CROSS_SHEET_DEPENDENCIES: &[u16] = &[0x009e];
    let records = raw_records(workbook)?;
    let sheets = extern_sheets(workbook, &records)?;
    // A record of another worksheet reaches the shifted one only through a
    // 3-D reference, which needs an XTI of this workbook covering its tab;
    // the shifted worksheet's own records are certified with it
    // (`certify_coordinate_shift`). Record 0308.
    let bounds = bound_sheets_in(workbook, &records)?;
    let target = worksheet_range_in(&bounds, workbook.len(), tab)?;
    let target_name = bounds
        .get(tab)
        .map(|bound| {
            record_payload(workbook, &bound.record, "BoundSheet").and_then(decode_bound_name)
        })
        .transpose()?
        .unwrap_or_default();
    let reachable = sheets.entries.iter().any(|&(book, first, last)| {
        Some(book) == sheets.self_book
            && first >= 0
            && last >= 0
            && (first.min(last)..=first.max(last)).contains(&i16::try_from(tab).unwrap_or(i16::MAX))
    });
    if records.iter().any(|record| {
        GLOBAL_OR_CROSS_SHEET_DEPENDENCIES.contains(&record.kind)
            && (reachable || target.contains(&record.start))
    }) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot close workbook dependency record",
        ));
    }
    let mut current_book_is_self = false;
    for (position, record) in records.iter().enumerate() {
        let continued = records
            .get(position + 1)
            .is_some_and(|next| next.kind == CONTINUE);
        let payload_start = record
            .start
            .checked_add(4)
            .ok_or_else(|| Error::InvalidData("dependency payload offset overflow".into()))?;
        match record.kind {
            SUP_BOOK => {
                let payload = record_payload(workbook, record, "SupBook")?;
                current_book_is_self = payload.len() == 4 && payload[2..4] == [0x01, 0x04];
            },
            EXTERN_NAME if current_book_is_self => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement cannot close an ExternName of this workbook",
                ));
            },
            LBL | BRAI if continued => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement cannot close a continued record",
                ));
            },
            LBL => {
                let payload = workbook
                    .get_mut(payload_start..record.end)
                    .ok_or_else(|| Error::InvalidData("Lbl payload is truncated".into()))?;
                if payload.len() < 14 {
                    return Err(Error::InvalidData("Lbl payload is truncated".into()));
                }
                let name_length = usize::from(payload[3]);
                let formula_length = usize::from(binary::read_u16_le_at(payload, 4)?);
                let wide = payload[14..].first().is_some_and(|flags| flags & 1 != 0);
                let start = 15 + name_length * if wide { 2 } else { 1 };
                if name_length == 0 || formula_length == 0 {
                    continue;
                }
                let tokens = payload
                    .get_mut(start..start + formula_length)
                    .ok_or_else(|| Error::InvalidData("Lbl formula is truncated".into()))?;
                patch_3d_references(tokens, &sheets, tab, shift, true)?;
            },
            BRAI => {
                let payload = workbook
                    .get_mut(payload_start..record.end)
                    .ok_or_else(|| Error::InvalidData("BRAI payload is truncated".into()))?;
                let formula_length = usize::from(binary::read_u16_le_at(payload, 6)?);
                if formula_length == 0 {
                    continue;
                }
                if payload.len() != 8 + formula_length {
                    return Err(Error::InvalidData(
                        "BRAI formula length does not match its payload".into(),
                    ));
                }
                let tokens = &mut payload[8..];
                let references =
                    crate::formula::formula_reference_spans(tokens).ok_or_else(|| {
                        Error::refused(
                            RefusalKind::UnsupportedOperation,
                            "edit cells",
                            "row/column movement cannot read a chart series formula",
                        )
                    })?;
                if !references.same_sheet.is_empty() {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "row/column movement cannot close a sheet-less chart reference",
                    ));
                }
                patch_3d_references(tokens, &sheets, tab, shift, false)?;
            },
            FORMULA => {
                let payload = workbook
                    .get_mut(payload_start..record.end)
                    .ok_or_else(|| Error::InvalidData("Formula payload is truncated".into()))?;
                let length = formula_tokens(payload)?.len();
                let tokens = &mut payload[22..22 + length];
                // A PtgExp or PtgTbl cell holds no 3-D reference.
                if formula_anchor(tokens).is_none() {
                    patch_3d_references(tokens, &sheets, tab, shift, false)?;
                }
            },
            HLINK => {
                let payload = record_payload(workbook, record, "HLINK")?;
                certify_hyperlink_target(
                    payload,
                    limits,
                    target.contains(&record.start),
                    &target_name,
                )?;
            },
            // Shared, array, conditional-format and validation formulas of
            // any worksheet keep their 3-D references into the shifted one;
            // their same-sheet references move with their own worksheet
            // (record 0316).
            SHARED_FORMULA | ARRAY | CF | rules::CF12 | DV if continued => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement cannot close a continued record",
                ));
            },
            SHARED_FORMULA | ARRAY if !reachable => {},
            SHARED_FORMULA | ARRAY => {
                let payload = workbook
                    .get_mut(payload_start..record.end)
                    .ok_or_else(|| Error::InvalidData("formula payload is truncated".into()))?;
                let (length_at, tokens_at) = if record.kind == ARRAY {
                    (12, 14)
                } else {
                    (8, 10)
                };
                let length = usize::from(binary::read_u16_le_at(payload, length_at)?);
                let tokens = payload
                    .get_mut(tokens_at..tokens_at + length)
                    .ok_or_else(|| Error::InvalidData("formula tokens are truncated".into()))?;
                rules::patch_rule_3d_references(tokens, &sheets, tab, shift)?;
            },
            CF | rules::CF12 | DV if reachable => {
                let payload = workbook
                    .get_mut(payload_start..record.end)
                    .ok_or_else(|| Error::InvalidData("rule payload is truncated".into()))?;
                rules::patch_rule_record_3d(record.kind, payload, &sheets, tab, shift)?;
            },
            _ => {},
        }
    }
    Ok(())
}

const CONDFMT: u16 = 0x01b0;
const CF: u16 = 0x01b1;
const DV: u16 = 0x01be;

/// Shifts a `CondFmt` record (`[MS-XLS]` 2.4.42): its bounding range and
/// every range of its `sqref`. Its `CF` rules are certified separately
/// (`cf_formulas`); record 0308.
fn patch_condfmt(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() < 14 {
        return Err(Error::InvalidData("CondFmt record is truncated".into()));
    }
    patch_ref8(&mut payload[4..12], shift)?;
    patch_sqref(&mut payload[12..], shift)
}

/// Shifts every `Ref8U` of a `SqRefU` (count, then the ranges), which must
/// fill `payload`.
fn patch_sqref(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let count = usize::from(binary::read_u16_le_at(payload, 0)?);
    if payload.len() != 2 + count * 8 {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "a range list does not fill its record and is not shifted",
        ));
    }
    for range in payload[2..].as_chunks_mut::<8>().0.iter_mut() {
        patch_ref8(range, shift)?;
    }
    Ok(())
}

/// The byte range of the worksheet substream at workbook tab `tab`.
fn worksheet_range_in(
    bounds: &[BoundSheet],
    length: usize,
    tab: usize,
) -> Result<std::ops::Range<usize>> {
    let Some(bound) = bounds.get(tab) else {
        // A stream without the tab's BoundSheet holds none of its records.
        return Ok(0..0);
    };
    let start = usize::try_from(bound.position)
        .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
    let end = bounds
        .iter()
        .filter_map(|candidate| {
            let position = usize::try_from(candidate.position).ok()?;
            (position > start).then_some(position)
        })
        .min()
        .unwrap_or(length);
    Ok(start..end)
}

fn rebuild_sheet(
    worksheet: &[u8],
    absolute_start: usize,
    sheet: usize,
    source: &super::Snapshot,
    operations: &[&StructuralChange],
    resources: &[super::ResourceChange],
    layout: &[&super::layout::LayoutChange],
    shared_strings: &[String],
) -> Result<Vec<u8>> {
    let records = raw_records(worksheet)?;
    // An embedded chart substream (BOF ... EOF inside the worksheet) has
    // its own DIMENSIONS; the worksheet's records are the top-level ones.
    let top = top_level(&records);
    let dimensions = unique_kind(&top, DIMENSIONS, "DIMENSIONS")?;
    let def_col = unique_kind(&top, DEF_COL_WIDTH, "DEFCOLWIDTH")?;
    let first_row = records.iter().position(|record| record.kind == ROW);
    let last_dbcell = records.iter().rposition(|record| record.kind == DBCELL);
    // A producer may write no INDEX and no DBCELL (the grammar requires
    // INDEX; readers do not). The row table is then regenerated in the same
    // row blocks and written back without either record (record 0308).
    let unindexed = !top.iter().any(|(_, record)| record.kind == INDEX) && last_dbcell.is_none();
    let index = if unindexed {
        def_col
    } else {
        unique_kind(&top, INDEX, "INDEX")?
    };
    let last_row_record = if unindexed {
        first_row.map(|first| {
            first
                + records[first..]
                    .iter()
                    .take_while(|record| {
                        is_cell_record(record.kind)
                            || matches!(
                                record.kind,
                                ROW | STRING | CONTINUE | TABLE | ARRAY | SHARED_FORMULA
                            )
                    })
                    .count()
                - 1
        })
    } else {
        last_dbcell
    };
    // An empty sheet whose producer still wrote a DBCELL (and an INDEX
    // pointing at it) keeps that row table while it stays empty (record
    // 0316).
    // Only a sheet without ROW records asks.
    let top_holds_cells = || {
        top.iter().any(|(_, record)| {
            is_cell_record(record.kind)
                || matches!(
                    record.kind,
                    STRING | TABLE | ARRAY | SHARED_FORMULA | 0x0091
                )
        })
    };
    let lone_dbcells = match (first_row, last_dbcell) {
        (None, Some(last)) if !unindexed && !top_holds_cells() => records[..=last]
            .iter()
            .position(|record| record.kind == DBCELL)
            .filter(|first| {
                *first > index
                    && records[*first..=last]
                        .iter()
                        .all(|record| record.kind == DBCELL)
            }),
        _ => None,
    };
    let (row_records, row_table_start, row_table_end) = match (first_row, last_row_record) {
        (Some(first), Some(last)) if first > index && last >= first => (
            &records[first..=last],
            records[first].start,
            records[last].end,
        ),
        (None, Some(last)) if lone_dbcells.is_some() => {
            let first = lone_dbcells.unwrap_or(last);
            (&records[0..0], records[first].start, records[last].end)
        },
        (None, None) if unindexed => {
            if top_holds_cells() {
                return Err(Error::Internal(
                    "empty worksheet row-table metadata is inconsistent",
                ));
            }
            (
                &records[0..0],
                records[dimensions].end,
                records[dimensions].end,
            )
        },
        (None, None) if !unindexed => {
            // An empty indexed sheet has no row blocks. Require empty extents,
            // no INDEX block pointers and no orphaned cells before creating its
            // first row block directly after DIMENSIONS.
            let index_payload = &worksheet[records[index].start + 4..records[index].end];
            let dimensions_payload =
                &worksheet[records[dimensions].start + 4..records[dimensions].end];
            if index_payload.len() != 16
                || index_payload[..12].iter().any(|byte| *byte != 0)
                || dimensions_payload.len() != 14
                || dimensions_payload.iter().any(|byte| *byte != 0)
                || dimensions <= index
                || def_col <= index
                || def_col >= dimensions
                || top_holds_cells()
            {
                return Err(Error::Internal(
                    "empty worksheet row-table metadata is inconsistent",
                ));
            }
            (
                &records[0..0],
                records[dimensions].end,
                records[dimensions].end,
            )
        },
        _ => {
            return Err(Error::InvalidData(
                "worksheet row-block order is invalid".into(),
            ));
        },
    };
    let shifting = operations.iter().any(|operation| {
        matches!(
            operation,
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. }
        )
    });
    if shifting {
        for operation in operations {
            if let Some(shift) = axis_shift(operation) {
                certify_coordinate_shift(worksheet, &records, shift)?;
            }
        }
    }

    let mut rows = collect_rows(worksheet, row_records)?;
    let mut formula_resources: Vec<_> = resources
        .iter()
        .filter(|resource| {
            matches!(resource, super::ResourceChange::FormulaCell { sheet: owner, .. } if *owner == sheet)
        })
        .collect();
    formula_resources.sort_by_cached_key(|resource| super::resource_target(resource));
    // Authored Formula removals vacate their cells before cell operations
    // run, and insertions follow them, so a scalar and an authored Formula
    // can replace each other at one coordinate in either direction.
    for insert_phase in [false, true] {
        if insert_phase {
            for operation in operations {
                apply_rebuild_operation(&mut rows, operation, source, resources, shared_strings)?;
            }
        }
        for resource in &formula_resources {
            if let super::ResourceChange::FormulaCell {
                sheet: _,
                reference,
                style,
                tokens,
                insert,
            } = resource
                && *insert == insert_phase
            {
                apply_formula_cell_resource(&mut rows, *reference, *style, tokens, *insert)?;
            }
        }
    }
    apply_row_layout(&mut rows, layout)?;
    patch_row_extents(&mut rows)?;

    // Without INDEX, the records from DEFCOLWIDTH on stand where the
    // records after INDEX stand.
    let (index_start, index_end) = if unindexed {
        (records[def_col].start, records[def_col].start)
    } else {
        (records[index].start, records[index].end)
    };
    let source_between = worksheet
        .get(index_end..row_table_start)
        .ok_or_else(|| Error::InvalidData("INDEX-to-row range is invalid".into()))?;
    // Column layout rewrites the ColInfo run after DefColWidth, which can
    // change the length of this region; DIMENSIONS is located again below.
    let between = super::layout::apply_columns(source_between, layout)?;
    let between = between.as_slice();
    let def_offset = records[def_col]
        .start
        .checked_sub(index_end)
        .ok_or_else(|| Error::InvalidData("DEFCOLWIDTH precedes INDEX".into()))?;
    let layout_rows = rows
        .into_iter()
        .map(|(row, data)| {
            let mut cells = Vec::new();
            for record in data.cell_records {
                cells.extend_from_slice(&record);
            }
            RowBlockLayoutRow::new(row, data.row_record, cells)
        })
        .collect();
    let absolute_index = absolute_start
        .checked_add(index_start)
        .ok_or_else(|| Error::InvalidData("absolute INDEX offset overflow".into()))?;
    let plan = RowBlockLayoutPlan::generate(
        u64::try_from(absolute_index)
            .map_err(|_error| Error::InvalidData("INDEX offset exceeds u64".into()))?,
        u64::try_from(between.len())
            .map_err(|_error| Error::InvalidData("INDEX gap exceeds u64".into()))?,
        u64::try_from(def_offset)
            .map_err(|_error| Error::InvalidData("DEFCOLWIDTH offset exceeds u64".into()))?,
        layout_rows,
    )
    .map_err(|_error| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "the BIFF8 row block cannot be regenerated",
        )
    })?;
    let (mut new_index, new_rows) = plan.into_records();
    // A producer INDEX whose `ibXF` does not locate DEFCOLWIDTH keeps its
    // value; readers locate DEFCOLWIDTH by record order (record 0316).
    if !unindexed {
        let source_ibxf = worksheet
            .get(records[index].start + 16..records[index].start + 20)
            .ok_or_else(|| Error::InvalidData("INDEX record is truncated".into()))?;
        let actual = u32::try_from(absolute_start + records[def_col].start)
            .map_err(|_error| Error::InvalidData("DEFCOLWIDTH offset exceeds u32".into()))?;
        if source_ibxf != actual.to_le_bytes() {
            new_index
                .get_mut(16..20)
                .ok_or_else(|| Error::InvalidData("regenerated INDEX is truncated".into()))?
                .copy_from_slice(source_ibxf);
        }
    }
    let (new_index, new_rows) = if lone_dbcells.is_some()
        && between.len() == source_between.len()
        && !new_rows_hold_rows(&new_rows)?
    {
        (
            worksheet[records[index].start..records[index].end].to_vec(),
            worksheet[row_table_start..row_table_end].to_vec(),
        )
    } else {
        (new_index, new_rows)
    };

    let computed = dimensions_from_row_table(&new_rows)?;
    // Producers write DIMENSIONS wider than the cells span (rows with only a
    // ROW record, say). When the source's extents are wider than its cells,
    // they are kept, moved by the shifts, while they still cover the cells,
    // so an edit and its inverse leave the record byte-identical (record
    // 0308). Exact source extents stay exact.
    let source_dimensions = worksheet
        .get(records[dimensions].start + 4..records[dimensions].end)
        .unwrap_or_default();
    // The source's exactness matters only when the shifted extents differ
    // from the regenerated ones, so the source row table is read only then.
    let stored = match stored_dimensions(source_dimensions, operations) {
        Some(stored) if stored != computed => {
            let source_exact = stored_dimensions(source_dimensions, &[])
                == Some(dimensions_from_row_table(
                    worksheet
                        .get(row_table_start..row_table_end)
                        .unwrap_or_default(),
                )?);
            (!source_exact).then_some(stored)
        },
        _ => None,
    };
    // Some producers write the last row and column, not one past them; the
    // shifted extents keep that convention while it still describes the
    // cells exactly (record 0316).
    let inclusive = |extents: (u32, u32, u16, u16)| {
        extents.1 > 0 && extents.3 > 0 && extents.1 > extents.0 && extents.3 > extents.2
    };
    let (first_used_row, last_used_row, first_used_col, last_used_col) = match stored {
        Some(stored)
            if computed == (0, 0, 0, 0)
                || (stored.0 <= computed.0
                    && stored.1 >= computed.1
                    && stored.2 <= computed.2
                    && stored.3 >= computed.3) =>
        {
            stored
        },
        Some(stored)
            if inclusive(computed)
                && stored == (computed.0, computed.1 - 1, computed.2, computed.3 - 1) =>
        {
            stored
        },
        _ => computed,
    };
    let mut new_dimensions = Vec::new();
    crate::writer::biff::write_dimensions(
        &mut new_dimensions,
        first_used_row,
        last_used_row,
        first_used_col,
        last_used_col,
    )?;
    if new_dimensions.len() != records[dimensions].end - records[dimensions].start {
        return Err(Error::refused(
            RefusalKind::UnsupportedContent,
            "edit cells",
            "DIMENSIONS encoding changed fixed record width",
        ));
    }
    let mut between = between.to_vec();
    let dimensions_relative = records[dimensions]
        .start
        .checked_sub(index_end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS precedes INDEX".into()))?
        .checked_add(between.len())
        .and_then(|value| value.checked_sub(source_between.len()))
        .ok_or_else(|| Error::InvalidData("DIMENSIONS offset adjustment overflows".into()))?;
    let dimensions_end = dimensions_relative
        .checked_add(new_dimensions.len())
        .ok_or_else(|| Error::InvalidData("DIMENSIONS range overflow".into()))?;
    between
        .get_mut(dimensions_relative..dimensions_end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS is outside INDEX gap".into()))?
        .copy_from_slice(&new_dimensions);

    let new_rows = if unindexed {
        without_dbcells(&new_rows)?
    } else {
        new_rows
    };
    let mut result = Vec::new();
    result.extend_from_slice(&worksheet[..index_start]);
    if !unindexed {
        result.extend_from_slice(&new_index);
    }
    result.extend_from_slice(&between);
    result.extend_from_slice(&new_rows);
    result.extend_from_slice(&worksheet[row_table_end..]);
    // Anchors that keep their place are checked against the sizes one
    // shift predicts, so a batch with several shifts or with row and column
    // layout changes keeps refusing them (record 0316).
    let predictable = layout.is_empty()
        && operations
            .iter()
            .filter(|operation| axis_shift(operation).is_some())
            .count()
            == 1;
    patch_coordinate_dependencies(
        &mut result,
        operations,
        predictable.then_some((worksheet, records.as_slice())),
    )?;
    Ok(result)
}

/// Whether a regenerated row table holds a `ROW` record.
fn new_rows_hold_rows(rows: &[u8]) -> Result<bool> {
    Ok(raw_records(rows)?.iter().any(|record| record.kind == ROW))
}

/// `rows`, a regenerated row table, without its DBCELL records.
fn without_dbcells(rows: &[u8]) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    out.try_reserve_exact(rows.len())
        .map_err(|_error| Error::Allocation("unindexed row table"))?;
    for record in raw_records(rows)? {
        if record.kind != DBCELL {
            out.extend_from_slice(
                rows.get(record.start..record.end)
                    .ok_or_else(|| Error::InvalidData("row table record is truncated".into()))?,
            );
        }
    }
    Ok(out)
}

/// Patch, create or reset `Row` records for staged row layout changes.
fn apply_row_layout(
    rows: &mut BTreeMap<u16, RowData>,
    layout: &[&super::layout::LayoutChange],
) -> Result<()> {
    for change in layout {
        let super::layout::LayoutChange::Row {
            row, before, after, ..
        } = change
        else {
            continue;
        };
        match rows.get_mut(row) {
            Some(data) => {
                let observed = super::layout::row_record_state(&data.row_record)?;
                // A cell insertion may have created this record in the same
                // edit; it then carries the default layout.
                if before.present && observed != *before {
                    return Err(Error::PatchConflict {
                        operation: "edit cells",
                        reason: "row layout precondition is stale",
                    });
                }
                if !after.present && data.cell_records.is_empty() {
                    rows.remove(row);
                } else {
                    let mut state = *after;
                    if !after.present {
                        state.custom = false;
                        state.hidden = false;
                    }
                    super::layout::patch_row_record(&mut data.row_record, state)?;
                }
            },
            None if before.present => {
                return Err(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "row layout precondition names a missing Row record",
                });
            },
            None if after.present => {
                rows.insert(
                    *row,
                    RowData {
                        row_record: super::layout::new_row_record(*row, *after)?,
                        cell_records: Vec::new(),
                        stored_extents: None,
                    },
                );
            },
            None => {},
        }
    }
    Ok(())
}

fn apply_rebuild_operation(
    rows: &mut BTreeMap<u16, RowData>,
    operation: &StructuralChange,
    source: &super::Snapshot,
    resources: &[super::ResourceChange],
    shared_strings: &[String],
) -> Result<()> {
    match operation {
        StructuralChange::Cell {
            reference,
            before,
            after,
            ..
        } => {
            apply_cell_change(
                rows,
                *reference,
                before.as_ref(),
                after.as_ref(),
                source,
                resources,
                shared_strings,
            )?;
        },
        StructuralChange::Rows {
            start,
            count,
            insert,
            ..
        } => {
            shift_rows(rows, *start, *count, *insert)?;
        },
        StructuralChange::Columns {
            start,
            count,
            insert,
            ..
        } => {
            shift_columns(rows, *start, *count, *insert)?;
        },
        StructuralChange::RenameSheet { .. } => {},
    }
    Ok(())
}

fn axis_shift(operation: &StructuralChange) -> Option<AxisShift> {
    match operation {
        StructuralChange::Rows {
            start,
            count,
            insert,
            ..
        } => Some(AxisShift::Rows {
            start: *start,
            count: *count,
            insert: *insert,
        }),
        StructuralChange::Columns {
            start,
            count,
            insert,
            ..
        } => Some(AxisShift::Columns {
            start: *start,
            count: *count,
            insert: *insert,
        }),
        StructuralChange::Cell { .. } | StructuralChange::RenameSheet { .. } => None,
    }
}

fn patch_coordinate_dependencies(
    worksheet: &mut [u8],
    operations: &[&StructuralChange],
    source: Option<(&[u8], &[RawRecord])>,
) -> Result<()> {
    patch_selection_records(worksheet, operations)?;
    let records = raw_records(worksheet)?;
    for operation in operations {
        let Some(shift) = axis_shift(operation) else {
            continue;
        };
        let mut rules = false;
        for record in &records {
            let payload_start = record
                .start
                .checked_add(4)
                .ok_or_else(|| Error::InvalidData("coordinate payload offset overflow".into()))?;
            let payload = worksheet
                .get_mut(payload_start..record.end)
                .ok_or_else(|| {
                    Error::InvalidData("coordinate record payload is truncated".into())
                })?;
            match record.kind {
                FORMULA => patch_formula_references(payload, shift)?,
                SHARED_FORMULA => patch_shared_formula(payload, shift)?,
                // Their ranges moved with their rows (`shift_rows`,
                // `shift_columns`).
                ARRAY => rules::patch_array_formula(payload, shift)?,
                TABLE => rules::patch_table_inputs(payload, shift)?,
                kind if rules::RULE_RECORDS.contains(&kind) => rules = true,
                MERGED_CELLS => patch_merged_cells(payload, shift)?,
                HLINK => patch_hyperlink(payload, shift)?,
                HLINK_TOOLTIP => patch_hyperlink_tooltip(payload, shift)?,
                COL_INFO => patch_col_info(payload, shift)?,
                NOTE => patch_note(payload, shift)?,
                VERTICAL_PAGE_BREAKS => patch_vertical_breaks(payload, shift)?,
                HORIZONTAL_PAGE_BREAKS => patch_horizontal_breaks(payload, shift)?,
                _ => {},
            }
        }
        if rules {
            rules::patch_rule_records(worksheet, &records, shift)?;
        }
        let sizes = match source {
            Some((source, source_records)) if !drawing_ranges(&records).is_empty() => {
                Some(sizes::SizeCheck::predict(source, source_records, shift)?)
            },
            _ => None,
        };
        patch_drawing_stream(worksheet, &records, shift, sizes.as_ref())?;
    }
    Ok(())
}

/// The payload ranges of the worksheet's OfficeArt stream: every top-level
/// `MsoDrawing` record and the `Continue` records that directly follow one.
/// A drawing's records are split across them (record 0308).
fn drawing_ranges(records: &[RawRecord]) -> Vec<std::ops::Range<usize>> {
    let mut ranges = Vec::new();
    let mut chained = false;
    for (_, record) in top_level(records) {
        chained = match record.kind {
            MSO_DRAWING => true,
            CONTINUE if chained => true,
            _ => false,
        };
        if chained {
            ranges.push(record.start + 4..record.end);
        }
    }
    ranges
}

/// The worksheet's OfficeArt stream, read as the concatenation of its
/// `MsoDrawing` payloads, with its ClientAnchors shifted (anchors keep their
/// width, so every record keeps its length). `None` without a drawing.
fn shifted_drawing_stream(
    worksheet: &[u8],
    ranges: &[std::ops::Range<usize>],
    shift: AxisShift,
    sizes: Option<&sizes::SizeCheck>,
) -> Result<Option<Vec<u8>>> {
    if ranges.is_empty() {
        return Ok(None);
    }
    let mut stream = Vec::new();
    for range in ranges {
        stream.extend_from_slice(
            worksheet
                .get(range.clone())
                .ok_or_else(|| Error::InvalidData("MsoDrawing payload is truncated".into()))?,
        );
    }
    patch_officeart(&mut stream, shift, sizes)?;
    Ok(Some(stream))
}

/// Shifts the ClientAnchors of the worksheet's OfficeArt stream in place.
/// `sizes` are the shifted axis's sizes, for anchors that keep their place.
fn patch_drawing_stream(
    worksheet: &mut [u8],
    records: &[RawRecord],
    shift: AxisShift,
    sizes: Option<&sizes::SizeCheck>,
) -> Result<()> {
    let ranges = drawing_ranges(records);
    let Some(stream) = shifted_drawing_stream(worksheet, &ranges, shift, sizes)? else {
        return Ok(());
    };
    let mut offset = 0;
    for range in ranges {
        let length = range.len();
        let source = stream
            .get(offset..offset + length)
            .ok_or_else(|| Error::InvalidData("OfficeArt stream is truncated".into()))?;
        worksheet
            .get_mut(range)
            .ok_or_else(|| Error::InvalidData("MsoDrawing payload is truncated".into()))?
            .copy_from_slice(source);
        offset += length;
    }
    Ok(())
}

fn patch_selection_records(worksheet: &mut [u8], operations: &[&StructuralChange]) -> Result<()> {
    let records = raw_records(worksheet)?;
    for operation in operations {
        let Some(shift) = axis_shift(operation) else {
            continue;
        };
        for record in records.iter().filter(|record| record.kind == SELECTION) {
            let payload = worksheet
                .get_mut(record.start + 4..record.end)
                .ok_or_else(|| Error::InvalidData("SELECTION payload is truncated".into()))?;
            patch_selection_shift(payload, shift)?;
        }
    }
    Ok(())
}

fn patch_selection_shift(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() < 9 {
        return Err(Error::InvalidData(
            "SELECTION payload is shorter than nine bytes".into(),
        ));
    }
    let range_count = usize::from(binary::read_u16_le_at(payload, 7)?);
    let expected = 9usize
        .checked_add(
            range_count
                .checked_mul(6)
                .ok_or_else(|| Error::InvalidData("SELECTION range size overflow".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("SELECTION payload size overflow".into()))?;
    if payload.len() != expected {
        return Err(Error::InvalidData(
            "SELECTION range count does not match its payload".into(),
        ));
    }
    match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => {
            let active = binary::read_u16_le_at(payload, 1)?;
            let active = shift_axis_point(active, start, count, insert, 65_536)?;
            payload[1..3].copy_from_slice(&active.to_le_bytes());
            for range in payload[9..].as_chunks_mut::<6>().0.iter_mut() {
                let first = binary::read_u16_le_at(range, 0)?;
                let last = binary::read_u16_le_at(range, 2)?;
                let (first, last) = shift_axis_range(first, last, start, count, insert, 65_536)?;
                range[0..2].copy_from_slice(&first.to_le_bytes());
                range[2..4].copy_from_slice(&last.to_le_bytes());
            }
        },
        AxisShift::Columns {
            start,
            count,
            insert,
        } => {
            let active = binary::read_u16_le_at(payload, 3)?;
            let active = shift_axis_point(active, u16::from(start), u16::from(count), insert, 256)?;
            payload[3..5].copy_from_slice(&active.to_le_bytes());
            for range in payload[9..].as_chunks_mut::<6>().0.iter_mut() {
                let (first, last) = shift_axis_range(
                    u16::from(range[4]),
                    u16::from(range[5]),
                    u16::from(start),
                    u16::from(count),
                    insert,
                    256,
                )?;
                range[4] = u8::try_from(first).map_err(|_error| {
                    Error::InvalidData("shifted SELECTION column exceeds u8".into())
                })?;
                range[5] = u8::try_from(last).map_err(|_error| {
                    Error::InvalidData("shifted SELECTION column exceeds u8".into())
                })?;
            }
        },
    }
    Ok(())
}

fn shift_axis_point(value: u16, start: u16, count: u16, insert: bool, limit: u32) -> Result<u16> {
    let value = u32::from(value);
    let start = u32::from(start);
    let count = u32::from(count);
    let shifted = if insert {
        if value < start {
            value
        } else {
            value
                .checked_add(count)
                .ok_or_else(|| Error::InvalidData("dependency coordinate shift overflows".into()))?
        }
    } else {
        let end = start.saturating_add(count).min(limit);
        let removed = end - start;
        if value < start {
            value
        } else if value >= end {
            value - removed
        } else {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "row/column deletion would erase a dependency coordinate",
            ));
        }
    };
    if shifted >= limit {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "row/column insertion moves a dependency coordinate outside BIFF8",
        ));
    }
    u16::try_from(shifted)
        .map_err(|_error| Error::InvalidData("dependency coordinate exceeds u16".into()))
}

fn shift_axis_range(
    first: u16,
    last: u16,
    start: u16,
    count: u16,
    insert: bool,
    limit: u32,
) -> Result<(u16, u16)> {
    if first > last {
        return Err(Error::InvalidData("dependency range is inverted".into()));
    }
    if first < last && u32::from(last) + 1 == limit {
        // A range through the last row or column of the grid (a whole column
        // `A:A`, a format to the sheet edge) keeps that end: an insertion
        // cannot push it outside the grid and a deletion does not pull it
        // in, so insert followed by delete stays an exact inverse.
        let first = if insert && first < start {
            first
        } else {
            shift_axis_point(first, start, count, insert, limit)?
        };
        return Ok((first, last));
    }
    if insert && first < start && last >= start {
        let shifted_last = shift_axis_point(last, start, count, true, limit)?;
        return Ok((first, shifted_last));
    }
    let first = shift_axis_point(first, start, count, insert, limit)?;
    let last = shift_axis_point(last, start, count, insert, limit)?;
    Ok((first.min(last), first.max(last)))
}

fn patch_merged_cells(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let count = usize::from(binary::read_u16_le_at(payload, 0)?);
    let expected = count
        .checked_mul(8)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| Error::InvalidData("MergedCells range size overflow".into()))?;
    if payload.len() != expected {
        return Err(Error::InvalidData(
            "MergedCells count does not match its payload".into(),
        ));
    }
    for range in payload[2..].as_chunks_mut::<8>().0.iter_mut() {
        patch_ref8(range, shift)?;
    }
    Ok(())
}

fn patch_hyperlink(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let range = payload
        .get_mut(..8)
        .ok_or_else(|| Error::InvalidData("HLINK range is truncated".into()))?;
    patch_ref8(range, shift)
}

fn patch_hyperlink_tooltip(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if binary::read_u16_le_at(payload, 0)? != HLINK_TOOLTIP {
        return Err(Error::InvalidData(
            "HLinkTooltip internal record type is invalid".into(),
        ));
    }
    let range = payload
        .get_mut(2..10)
        .ok_or_else(|| Error::InvalidData("HLinkTooltip range is truncated".into()))?;
    patch_ref8(range, shift)
}

/// `ColInfo` (MS-XLS 2.4.53): a column span's width and format. Rows do not
/// touch it. A column shift moves the span like a range, except that a span
/// ending at the last BIFF8 column (`colLast` 255, or 256 as some producers
/// write it) keeps that end: it means "through the end of the grid", which an
/// insertion cannot push outside the grid and a deletion cannot pull in. That
/// keeps insert followed by delete of the same columns an exact inverse.
fn patch_col_info(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let AxisShift::Columns {
        start,
        count,
        insert,
    } = shift
    else {
        return Ok(());
    };
    if payload.len() < 4 {
        return Err(Error::InvalidData("ColInfo payload is truncated".into()));
    }
    let first = binary::read_u16_le_at(payload, 0)?;
    let last = binary::read_u16_le_at(payload, 2)?;
    if first > last || first > 255 || last > 256 {
        return Err(Error::InvalidData("ColInfo column span is invalid".into()));
    }
    let (start, count) = (u16::from(start), u16::from(count));
    let (first, last) = if last >= 255 {
        let first = if insert && first >= start {
            first.checked_add(count).filter(|first| *first <= 255)
        } else if !insert && first >= start {
            if first < start + count {
                // The deleted span holds the open-ended span's first column:
                // its new start would be ambiguous.
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column deletion would erase a dependency coordinate",
                ));
            }
            Some(first - count)
        } else {
            Some(first)
        }
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "row/column insertion moves a dependency coordinate outside BIFF8",
            )
        })?;
        (first, last)
    } else {
        shift_axis_range(first, last, start, count, insert, 256)?
    };
    payload[0..2].copy_from_slice(&first.to_le_bytes());
    payload[2..4].copy_from_slice(&last.to_le_bytes());
    Ok(())
}

/// `Note` (MS-XLS 2.4.179): the commented cell.
fn patch_note(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() < 4 {
        return Err(Error::InvalidData("Note payload is truncated".into()));
    }
    match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => {
            let row = binary::read_u16_le_at(payload, 0)?;
            let row = shift_axis_point(row, start, count, insert, 65_536)?;
            payload[0..2].copy_from_slice(&row.to_le_bytes());
        },
        AxisShift::Columns {
            start,
            count,
            insert,
        } => {
            let column = binary::read_u16_le_at(payload, 2)?;
            let column = shift_axis_point(column, u16::from(start), u16::from(count), insert, 256)?;
            payload[2..4].copy_from_slice(&column.to_le_bytes());
        },
    }
    Ok(())
}

/// The breaks of a `VerticalPageBreaks` or `HorizontalPageBreaks` record
/// (MS-XLS 2.4.350, 2.4.122): a count, then six-byte `(position, span
/// first, span last)` entries.
fn page_breaks(payload: &mut [u8]) -> Result<&mut [[u8; 6]]> {
    let count = usize::from(binary::read_u16_le_at(payload, 0)?);
    if payload.len() != 2 + count * 6 {
        return Err(Error::InvalidData(
            "page-break count does not match its payload".into(),
        ));
    }
    Ok(payload[2..].as_chunks_mut::<6>().0)
}

/// Shifts the breaks of one record. `position_axis_rows` names the axis of
/// the break position (rows for horizontal breaks); the span runs along the
/// other axis, and a span covering that whole axis is left as it is.
fn patch_page_breaks(payload: &mut [u8], shift: AxisShift, position_axis_rows: bool) -> Result<()> {
    let (start, count, insert, rows) = match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => (start, count, insert, true),
        AxisShift::Columns {
            start,
            count,
            insert,
        } => (u16::from(start), u16::from(count), insert, false),
    };
    let limit = if rows { 65_536 } else { 256 };
    for entry in page_breaks(payload)? {
        if rows == position_axis_rows {
            // A break sits before its position; it moves with the row or
            // column it precedes, and deleting that row or column erases it.
            let position = u16::from_le_bytes([entry[0], entry[1]]);
            let position = shift_axis_point(position, start, count, insert, limit)?;
            entry[0..2].copy_from_slice(&position.to_le_bytes());
        } else {
            let first = u16::from_le_bytes([entry[2], entry[3]]);
            let last = u16::from_le_bytes([entry[4], entry[5]]);
            let whole = u32::from(last) + 1 >= limit && first == 0;
            if !whole {
                let (first, last) = shift_axis_range(first, last, start, count, insert, limit)?;
                entry[2..4].copy_from_slice(&first.to_le_bytes());
                entry[4..6].copy_from_slice(&last.to_le_bytes());
            }
        }
    }
    Ok(())
}

fn patch_vertical_breaks(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    patch_page_breaks(payload, shift, false)
}

fn patch_horizontal_breaks(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    patch_page_breaks(payload, shift, true)
}

fn patch_ref8(range: &mut [u8], shift: AxisShift) -> Result<()> {
    if range.len() != 8 {
        return Err(Error::InvalidData("Ref8 payload is not eight bytes".into()));
    }
    match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => {
            let first = binary::read_u16_le_at(range, 0)?;
            let last = binary::read_u16_le_at(range, 2)?;
            let (first, last) = shift_axis_range(first, last, start, count, insert, 65_536)?;
            range[0..2].copy_from_slice(&first.to_le_bytes());
            range[2..4].copy_from_slice(&last.to_le_bytes());
        },
        AxisShift::Columns {
            start,
            count,
            insert,
        } => {
            let first = binary::read_u16_le_at(range, 4)?;
            let last = binary::read_u16_le_at(range, 6)?;
            let (first, last) =
                shift_axis_range(first, last, u16::from(start), u16::from(count), insert, 256)?;
            range[4..6].copy_from_slice(&first.to_le_bytes());
            range[6..8].copy_from_slice(&last.to_le_bytes());
        },
    }
    Ok(())
}

fn patch_officeart(
    data: &mut [u8],
    shift: AxisShift,
    sizes: Option<&sizes::SizeCheck>,
) -> Result<()> {
    patch_officeart_records(data, shift, sizes)
}

fn patch_officeart_records(
    data: &mut [u8],
    shift: AxisShift,
    sizes: Option<&sizes::SizeCheck>,
) -> Result<()> {
    let mut offset = 0_usize;
    while offset < data.len() {
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| Error::InvalidData("OfficeArt header range overflow".into()))?;
        let header = data.get(offset..header_end).ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "split OfficeArt records are not shiftable",
            )
        })?;
        let options = u16::from_le_bytes([header[0], header[1]]);
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let length = usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ]))
        .map_err(|_error| Error::InvalidData("OfficeArt length exceeds usize".into()))?;
        let end = header_end
            .checked_add(length)
            .ok_or_else(|| Error::InvalidData("OfficeArt record range overflow".into()))?;
        let payload = data.get_mut(header_end..end).ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "split OfficeArt records are not shiftable",
            )
        })?;
        let version = options & 0x000f;
        let instance = options >> 4;
        if kind == 0xf010 {
            if version != 0 || instance != 0 || payload.len() != 18 {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "worksheet drawing has a noncanonical ClientAnchor",
                ));
            }
            patch_client_anchor(payload, shift, sizes)?;
        } else if version == 0x000f {
            patch_officeart_records(payload, shift, sizes)?;
        }
        offset = end;
    }
    Ok(())
}

/// Shifts one `OfficeArtClientAnchorSheet` (`[MS-XLS]` 2.5.193).
///
/// Its first word holds `fMove` (bit 0) and `fSize` (bit 1), each set when
/// the shape is kept intact as cells move or resize: 0 is
/// `AnchorBehavior::MoveAndResize`, 2 `MoveWithoutResize`, 3 `Absolute`
/// (Apache POI's `AnchorType` values; LibreOffice's position- and
/// size-locked flags; record 0308 corrected the reading). A moving anchor's
/// endpoints shift like cells; a size-locked one shifts whole when the
/// shift point is before it. An endpoint that keeps its place (both of a
/// position-locked anchor, the far end of a size-locked one the shift point
/// falls inside) is unchanged by a shift after it, and otherwise keeps its
/// cell coordinates when every row or column from the shift point to it
/// keeps its size (`sizes`, record 0316); anything else is refused, since
/// keeping its place would need the sizes converted. The spec-invalid 1
/// (`fMove` alone) is read as Apache POI and LibreOffice read it: the
/// top-left keeps its place and the bottom-right moves with its cell.
fn patch_client_anchor(
    payload: &mut [u8],
    shift: AxisShift,
    sizes: Option<&sizes::SizeCheck>,
) -> Result<()> {
    let behavior = binary::read_u16_le_at(payload, 0)?;
    let (first_offset, last_offset, start, count, insert, limit) = match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => (6, 14, start, count, insert, 65_536),
        AxisShift::Columns {
            start,
            count,
            insert,
        } => (2, 10, u16::from(start), u16::from(count), insert, 256),
    };
    let first = binary::read_u16_le_at(payload, first_offset)?;
    let last = binary::read_u16_le_at(payload, last_offset)?;
    // Whether an endpoint that keeps its place keeps its coordinates.
    let in_place = |to: u16| sizes.is_some_and(|sizes| sizes.unchanged(start, to));
    let (first, last) = match behavior {
        0 => shift_axis_range(first, last, start, count, insert, limit)?,
        _ if start > last => return Ok(()),
        0b10 if (insert && start <= first) || (!insert && start + count <= first) => (
            shift_axis_point(first, start, count, insert, limit)?,
            shift_axis_point(last, start, count, insert, limit)?,
        ),
        // The top-left cell stays and the size is kept.
        0b10 if start > first && in_place(last) => return Ok(()),
        0b10 => {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "a size-locked drawing anchor that spans the shift point is not reversibly shiftable",
            ));
        },
        0b11 if in_place(last) => return Ok(()),
        // `fMove` without `fSize` breaks the spec's "fSize MUST be 1 when
        // fMove is 1", but Excel writes it on AutoFilter and PivotTable
        // drop-downs. Apache POI (`DONT_MOVE_DO_RESIZE`) and LibreOffice
        // read it as position-locked and resizing: the top-left keeps its
        // place and the bottom-right moves with its cell (record 0316).
        0b01 if start > first || in_place(first) => {
            let moved = shift_axis_point(last, start, count, insert, limit)?;
            if moved < first {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "a row/column deletion would invert a resizing drawing anchor",
                ));
            }
            (first, moved)
        },
        _ => {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "a position-locked drawing anchor is not shiftable without row and column sizes",
            ));
        },
    };
    payload[first_offset..first_offset + 2].copy_from_slice(&first.to_le_bytes());
    payload[last_offset..last_offset + 2].copy_from_slice(&last.to_le_bytes());
    Ok(())
}

fn collect_rows(worksheet: &[u8], records: &[RawRecord]) -> Result<BTreeMap<u16, RowData>> {
    let mut rows = BTreeMap::new();
    let mut last_cell_row = None;
    for record in records {
        let bytes = worksheet[record.start..record.end].to_vec();
        match record.kind {
            ROW => {
                let row = binary::read_u16_le_at(&bytes, 4)?;
                if rows
                    .insert(
                        row,
                        RowData {
                            row_record: bytes,
                            cell_records: Vec::new(),
                            stored_extents: None,
                        },
                    )
                    .is_some()
                {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedContent,
                        "edit cells",
                        "duplicate ROW record in row table",
                    ));
                }
            },
            DBCELL => {},
            TABLE | SHARED_FORMULA | ARRAY => {
                let row = last_cell_row.ok_or_else(|| {
                    Error::InvalidData("formula continuation has no preceding cell".into())
                })?;
                rows.get_mut(&row)
                    .ok_or_else(|| Error::InvalidData("formula continuation row is absent".into()))?
                    .cell_records
                    .push(bytes);
            },
            STRING | CONTINUE => {
                let row = last_cell_row.ok_or_else(|| {
                    Error::InvalidData("formula String companion has no preceding cell".into())
                })?;
                rows.get_mut(&row)
                    .ok_or_else(|| Error::InvalidData("formula String row is absent".into()))?
                    .cell_records
                    .push(bytes);
            },
            kind if is_cell_record(kind) => {
                let row = binary::read_u16_le_at(&bytes, 4)?;
                rows.get_mut(&row)
                    .ok_or_else(|| Error::InvalidData("cell record has no ROW owner".into()))?
                    .cell_records
                    .push(bytes);
                last_cell_row = Some(row);
            },
            _ => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "record inside the BIFF8 row table is not losslessly classified",
                ));
            },
        }
    }
    for data in rows.values_mut() {
        let stored = row_extents_field(&data.row_record)?;
        if stored != cell_extents(&data.cell_records)? {
            data.stored_extents = Some(stored);
        }
    }
    Ok(rows)
}

/// The `(colMic, colMac)` fields of a `Row` record.
fn row_extents_field(row_record: &[u8]) -> Result<(u16, u16)> {
    Ok((
        binary::read_u16_le_at(row_record, 6)?,
        binary::read_u16_le_at(row_record, 8)?,
    ))
}

/// The `(first, last + 1)` column extents of a row's cell records, `(0, 0)`
/// for a row without cells.
fn cell_extents(cell_records: &[Vec<u8>]) -> Result<(u16, u16)> {
    let mut first = u16::MAX;
    let mut last = 0_u16;
    for record in cell_records {
        let kind = record_kind(record)?;
        if matches!(kind, TABLE | SHARED_FORMULA | ARRAY | STRING | CONTINUE) {
            continue;
        }
        let start = u16::from(cell_column(record)?);
        let end = if matches!(kind, MUL_RK | MUL_BLANK) {
            let payload_len = usize::from(u16::from_le_bytes([record[2], record[3]]));
            binary::read_u16_le_at(record, 4 + payload_len - 2)?
        } else {
            start
        };
        first = first.min(start);
        last = last.max(end);
    }
    Ok(if first == u16::MAX {
        (0, 0)
    } else {
        (first, last + 1)
    })
}

/// Moves stored `(colMic, colMac)` extents with a column shift; `None` when
/// they no longer describe a span.
fn shift_stored_extents(
    (first, end): (u16, u16),
    start: u16,
    count: u16,
    insert: bool,
) -> Option<(u16, u16)> {
    let (first, end) = if insert {
        (
            if first >= start { first + count } else { first },
            if end > start { end + count } else { end },
        )
    } else {
        let deleted_end = start + count;
        let shift = |value: u16| {
            if value >= deleted_end {
                value - count
            } else {
                value.min(start)
            }
        };
        (shift(first), shift(end))
    };
    (first < end && end <= 256).then_some((first, end))
}

fn apply_cell_change(
    rows: &mut BTreeMap<u16, RowData>,
    reference: Reference,
    before: Option<&(Storage, Value, StyleIndex)>,
    after: Option<&(Storage, Value, StyleIndex)>,
    source: &super::Snapshot,
    resources: &[super::ResourceChange],
    shared_strings: &[String],
) -> Result<()> {
    let mut found = None;
    if let Some(row) = rows.get_mut(&reference.row()) {
        for (index, record) in row.cell_records.iter().enumerate() {
            if record_contains_reference(record, reference)? && found.replace(index).is_some() {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "structural cell target is ambiguous",
                ));
            }
        }
    }
    match (before, after, found) {
        (None, Some(after), None) => {
            if after.0 == Storage::MulRk {
                return insert_packed_rk(rows, reference, after);
            }
            if let std::collections::btree_map::Entry::Vacant(row) = rows.entry(reference.row()) {
                let mut row_record = Vec::new();
                crate::writer::biff::write_row(
                    &mut row_record,
                    u32::from(reference.row()),
                    u16::from(reference.column()),
                    u16::from(reference.column()) + 1,
                    255,
                    false,
                )?;
                row.insert(RowData {
                    row_record,
                    cell_records: Vec::new(),
                    stored_extents: None,
                });
            }
            let record = encode_cell(reference, after, source, resources, shared_strings)?;
            rows.get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("inserted ROW disappeared".into()))?
                .cell_records
                .push(record);
        },
        (Some(expected), None, Some(index)) => {
            verify_record_state(
                &rows[&reference.row()].cell_records[index],
                reference,
                expected,
                source,
            )?;
            let row = rows
                .get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("removed ROW disappeared".into()))?;
            remove_cell_record(&mut row.cell_records, index, reference)?;
        },
        (Some(expected), Some(after), Some(index)) => {
            let row = rows
                .get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("styled ROW disappeared".into()))?;
            verify_record_state(&row.cell_records[index], reference, expected, source)?;
            if expected.0 == Storage::Formula
                && after.0 == Storage::Formula
                && !super::values_equal(&expected.1, &after.1)
            {
                replace_formula_cache(&mut row.cell_records, index, &after.1)?;
            } else if expected.0 != after.0 || !super::values_equal(&expected.1, &after.1) {
                // A scalar family change cannot reuse the fixed-width field.
                // Re-encode one standalone BIFF8 cell with the requested XF,
                // replacing the source record in the existing row block. A
                // packed MulRk member is split around the target first so
                // its neighbouring numeric cells remain intact.
                let replacement = encode_cell(reference, after, source, resources, shared_strings)?;
                let kind = record_kind(&row.cell_records[index])?;
                if kind == MUL_BLANK {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "scalar replacement of a packed blank record is unsupported",
                    ));
                }
                if kind == MUL_RK {
                    remove_cell_record(&mut row.cell_records, index, reference)?;
                    let column = u16::from(reference.column());
                    let mut insertion = row.cell_records.len();
                    for (position, record) in row.cell_records.iter().enumerate() {
                        if u16::from(cell_column(record)?) > column {
                            insertion = position;
                            break;
                        }
                    }
                    row.cell_records
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("retaining scalar replacement"))?;
                    row.cell_records.insert(insertion, replacement);
                } else {
                    row.cell_records[index] = replacement;
                }
            } else {
                patch_record_style(&mut row.cell_records[index], reference, after.2)?;
            }
        },
        (None, Some(_), Some(_)) => {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "cell insertion precondition is stale",
            });
        },
        (Some(_), _, None) => {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "structural cell precondition is stale",
            });
        },
        (None, None, _) => {
            return Err(Error::InvalidData(
                "structural cell operation has no outcome".into(),
            ));
        },
    }
    Ok(())
}

fn apply_formula_cell_resource(
    rows: &mut BTreeMap<u16, RowData>,
    reference: Reference,
    style: StyleIndex,
    tokens: &[u8],
    insert: bool,
) -> Result<()> {
    let mut found = None;
    if let Some(row) = rows.get(&reference.row()) {
        for (index, record) in row.cell_records.iter().enumerate() {
            if record_contains_reference(record, reference)? && found.replace(index).is_some() {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "authored Formula target is ambiguous",
                ));
            }
        }
    }
    if insert {
        if found.is_some() {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "authored Formula insertion target is occupied",
            ));
        }
        if let std::collections::btree_map::Entry::Vacant(row) = rows.entry(reference.row()) {
            let mut row_record = Vec::new();
            crate::writer::biff::write_row(
                &mut row_record,
                u32::from(reference.row()),
                u16::from(reference.column()),
                u16::from(reference.column()) + 1,
                255,
                false,
            )?;
            row.insert(RowData {
                row_record,
                cell_records: Vec::new(),
                stored_extents: None,
            });
        }
        let mut record = Vec::new();
        crate::writer::biff::write_formula(
            &mut record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            style.get(),
            tokens,
        )?;
        rows.get_mut(&reference.row())
            .ok_or_else(|| Error::InvalidData("authored Formula ROW disappeared".into()))?
            .cell_records
            .push(record);
        return Ok(());
    }
    let index = found.ok_or_else(|| {
        Error::refused(
            RefusalKind::NotFound,
            "edit cells",
            "authored Formula target is absent",
        )
    })?;
    let row = rows
        .get_mut(&reference.row())
        .ok_or_else(|| Error::InvalidData("authored Formula ROW disappeared".into()))?;
    let record = &row.cell_records[index];
    let mut expected = Vec::new();
    crate::writer::biff::write_formula(
        &mut expected,
        u32::from(reference.row()),
        u16::from(reference.column()),
        style.get(),
        tokens,
    )?;
    if record.as_slice() != expected.as_slice() {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "authored Formula inverse state is stale",
        });
    }
    if row
        .cell_records
        .get(index.saturating_add(1))
        .is_some_and(|next| record_kind(next).is_ok_and(|kind| matches!(kind, STRING | CONTINUE)))
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "authored Formula inverse refuses a string-cache owner group",
        ));
    }
    row.cell_records.remove(index);
    Ok(())
}

fn remove_cell_record(
    records: &mut Vec<Vec<u8>>,
    index: usize,
    reference: Reference,
) -> Result<()> {
    match record_kind(&records[index])? {
        MUL_RK => remove_packed_rk(records, index, reference),
        MUL_BLANK => remove_packed_blank(records, index, reference),
        _ => {
            records.remove(index);
            Ok(())
        },
    }
}

fn remove_packed_rk(records: &mut Vec<Vec<u8>>, index: usize, reference: Reference) -> Result<()> {
    let record = &mut records[index];
    let payload_len = usize::from(binary::read_u16_le_at(record, 2)?);
    let item_bytes = payload_len
        .checked_sub(6)
        .ok_or_else(|| Error::InvalidData("MulRk payload is truncated".into()))?;
    if item_bytes % 6 != 0 {
        return Err(Error::InvalidData("MulRk item framing is invalid".into()));
    }
    let count = item_bytes / 6;
    let first = binary::read_u16_le_at(record, 6)?;
    let target = u16::from(reference.column());
    let item = usize::from(target.checked_sub(first).ok_or_else(|| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "packed deletion target precedes its MulRk range",
        )
    })?);
    if item >= count {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "packed deletion target exceeds its MulRk range",
        ));
    }
    if item != 0 && item + 1 != count {
        let (_, _, items) = decode_rk_items(record)?;
        let right_first = target
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("MulRk split column overflow".into()))?;
        let left = encode_rk_segment(reference.row(), first, &items[..item])?;
        let right = encode_rk_segment(reference.row(), right_first, &items[item + 1..])?;
        records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining MulRk split"))?;
        records[index] = left;
        records.insert(index + 1, right);
        return Ok(());
    }
    if count == 2 {
        let remaining = usize::from(item == 0);
        let item_offset = 8 + remaining * 6;
        let column = first
            .checked_add(u16::try_from(remaining).map_err(|_error| {
                Error::InvalidData("MulRk remaining column exceeds u16".into())
            })?)
            .ok_or_else(|| Error::InvalidData("MulRk remaining column overflow".into()))?;
        let mut rk = Vec::with_capacity(14);
        rk.extend_from_slice(&RK.to_le_bytes());
        rk.extend_from_slice(&10_u16.to_le_bytes());
        rk.extend_from_slice(&reference.row().to_le_bytes());
        rk.extend_from_slice(&column.to_le_bytes());
        rk.extend_from_slice(
            record
                .get(item_offset..item_offset + 6)
                .ok_or_else(|| Error::InvalidData("MulRk remaining item is truncated".into()))?,
        );
        records[index] = rk;
        return Ok(());
    }
    let item_offset = 8usize
        .checked_add(
            item.checked_mul(6)
                .ok_or_else(|| Error::InvalidData("MulRk deletion item offset overflow".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("MulRk deletion offset overflow".into()))?;
    record.drain(item_offset..item_offset + 6);
    let new_payload = u16::try_from(payload_len - 6)
        .map_err(|_error| Error::InvalidData("MulRk payload exceeds u16".into()))?;
    record[2..4].copy_from_slice(&new_payload.to_le_bytes());
    if item == 0 {
        record[6..8].copy_from_slice(&(first + 1).to_le_bytes());
    } else {
        let last = target - 1;
        let end = record.len();
        record[end - 2..end].copy_from_slice(&last.to_le_bytes());
    }
    Ok(())
}

fn remove_packed_blank(
    records: &mut Vec<Vec<u8>>,
    index: usize,
    reference: Reference,
) -> Result<()> {
    let record = &mut records[index];
    let payload_len = usize::from(binary::read_u16_le_at(record, 2)?);
    let item_bytes = payload_len
        .checked_sub(6)
        .ok_or_else(|| Error::InvalidData("MulBlank payload is truncated".into()))?;
    if item_bytes == 0 || item_bytes % 2 != 0 {
        return Err(Error::InvalidData(
            "MulBlank item framing is invalid".into(),
        ));
    }
    let count = item_bytes / 2;
    let first = binary::read_u16_le_at(record, 6)?;
    let target = u16::from(reference.column());
    let item = usize::from(target.checked_sub(first).ok_or_else(|| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "packed blank deletion target precedes its MulBlank range",
        )
    })?);
    if item >= count {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "packed blank deletion target exceeds its MulBlank range",
        ));
    }

    // An interior removal is represented by two records. Keeping standalone
    // BLANK records for one-item sides avoids emitting an invalid one-item
    // MulBlank while retaining each neighboring XF exactly.
    if item != 0 && item + 1 != count {
        let styles = decode_blank_styles(record)?;
        let left = encode_blank_segment(reference.row(), first, &styles[..item])?;
        let right_first = target
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("MulBlank split column overflow".into()))?;
        let right = encode_blank_segment(reference.row(), right_first, &styles[item + 1..])?;
        records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining MulBlank split"))?;
        records[index] = left;
        records.insert(index + 1, right);
        return Ok(());
    }

    if count == 2 {
        let remaining = usize::from(item == 0);
        let item_offset = 8 + remaining * 2;
        let column = first
            .checked_add(u16::try_from(remaining).map_err(|_error| {
                Error::InvalidData("MulBlank remaining column exceeds u16".into())
            })?)
            .ok_or_else(|| Error::InvalidData("MulBlank remaining column overflow".into()))?;
        let style = binary::read_u16_le_at(record, item_offset)?;
        let remaining_reference = Reference::new(u32::from(reference.row()), u32::from(column))?;
        records[index] = scalar_record(BLANK, remaining_reference, &StyleIndex(style), &[]);
        return Ok(());
    }

    let item_offset =
        8usize
            .checked_add(item.checked_mul(2).ok_or_else(|| {
                Error::InvalidData("MulBlank deletion item offset overflow".into())
            })?)
            .ok_or_else(|| Error::InvalidData("MulBlank deletion offset overflow".into()))?;
    record.drain(item_offset..item_offset + 2);
    let new_payload = u16::try_from(payload_len - 2)
        .map_err(|_error| Error::InvalidData("MulBlank payload exceeds u16".into()))?;
    record[2..4].copy_from_slice(&new_payload.to_le_bytes());
    if item == 0 {
        record[6..8].copy_from_slice(&(first + 1).to_le_bytes());
    } else {
        let last = target - 1;
        let end = record.len();
        record[end - 2..end].copy_from_slice(&last.to_le_bytes());
    }
    Ok(())
}

fn decode_blank_styles(record: &[u8]) -> Result<Vec<u16>> {
    let payload_len = usize::from(binary::read_u16_le_at(record, 2)?);
    let item_bytes = payload_len
        .checked_sub(6)
        .ok_or_else(|| Error::InvalidData("MulBlank payload is truncated".into()))?;
    if item_bytes == 0 || item_bytes % 2 != 0 {
        return Err(Error::InvalidData(
            "MulBlank item framing is invalid".into(),
        ));
    }
    let count = item_bytes / 2;
    let mut styles = Vec::new();
    styles
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation("retaining MulBlank styles"))?;
    for item in 0..count {
        styles.push(binary::read_u16_le_at(record, 8 + item * 2)?);
    }
    Ok(styles)
}

fn encode_blank_segment(row: u16, first_column: u16, styles: &[u16]) -> Result<Vec<u8>> {
    if styles.is_empty() {
        return Err(Error::InvalidData("MulBlank segment is empty".into()));
    }
    let last_column = first_column
        .checked_add(
            u16::try_from(styles.len() - 1)
                .map_err(|_error| Error::InvalidData("MulBlank segment is too wide".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("MulBlank segment column overflow".into()))?;
    if last_column > 255 {
        return Err(Error::InvalidData(
            "MulBlank segment exceeds the BIFF8 worksheet grid".into(),
        ));
    }
    if styles.len() == 1 {
        let reference = Reference::new(u32::from(row), u32::from(first_column))?;
        return Ok(scalar_record(BLANK, reference, &StyleIndex(styles[0]), &[]));
    }
    let item_bytes = styles
        .len()
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidData("MulBlank item length overflow".into()))?;
    let payload_len = 6usize
        .checked_add(item_bytes)
        .ok_or_else(|| Error::InvalidData("MulBlank payload length overflow".into()))?;
    let mut record = Vec::with_capacity(4 + payload_len);
    record.extend_from_slice(&MUL_BLANK.to_le_bytes());
    record.extend_from_slice(
        &u16::try_from(payload_len)
            .map_err(|_error| Error::InvalidData("MulBlank payload exceeds u16".into()))?
            .to_le_bytes(),
    );
    record.extend_from_slice(&row.to_le_bytes());
    record.extend_from_slice(&first_column.to_le_bytes());
    for style in styles {
        record.extend_from_slice(&style.to_le_bytes());
    }
    record.extend_from_slice(&last_column.to_le_bytes());
    Ok(record)
}

fn insert_packed_rk(
    rows: &mut BTreeMap<u16, RowData>,
    reference: Reference,
    state: &(Storage, Value, StyleIndex),
) -> Result<()> {
    let Value::Number(number) = &state.1 else {
        return Err(Error::InvalidData(
            "MulRk insertion value is not numeric".into(),
        ));
    };
    let encoded = super::encode_rk(*number).ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "MulRk insertion value is not exactly RK-representable",
        )
    })?;
    if let std::collections::btree_map::Entry::Vacant(row) = rows.entry(reference.row()) {
        let mut row_record = Vec::new();
        crate::writer::biff::write_row(
            &mut row_record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            u16::from(reference.column()) + 1,
            255,
            false,
        )?;
        row.insert(RowData {
            row_record,
            cell_records: Vec::new(),
            stored_extents: None,
        });
    }
    let row = rows
        .get_mut(&reference.row())
        .ok_or_else(|| Error::InvalidData("MulRk inverse ROW disappeared".into()))?;
    let target = u16::from(reference.column());
    let mut candidates = Vec::new();
    for (index, record) in row.cell_records.iter().enumerate() {
        let kind = record_kind(record)?;
        let adjacent = if kind == MUL_RK {
            let first = binary::read_u16_le_at(record, 6)?;
            let last = binary::read_u16_le_at(record, record.len() - 2)?;
            target.checked_add(1) == Some(first) || last.checked_add(1) == Some(target)
        } else if kind == RK {
            let column = binary::read_u16_le_at(record, 6)?;
            target.checked_add(1) == Some(column) || column.checked_add(1) == Some(target)
        } else {
            false
        };
        if adjacent {
            candidates
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("retaining adjacent packed records"))?;
            candidates.push(index);
        }
    }
    if candidates.len() == 2 {
        let left_index = candidates[0];
        let right_index = candidates[1];
        let (left_row, left_first, mut left_items) =
            decode_rk_items(&row.cell_records[left_index])?;
        let (right_row, right_first, right_items) =
            decode_rk_items(&row.cell_records[right_index])?;
        if left_row != reference.row() || right_row != reference.row() {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "MulRk inverse adjacent records belong to another row",
            ));
        }
        let left_last =
            left_first
                .checked_add(u16::try_from(left_items.len().saturating_sub(1)).map_err(
                    |_error| Error::InvalidData("packed inverse span exceeds u16".into()),
                )?)
                .ok_or_else(|| Error::InvalidData("packed inverse span overflows".into()))?;
        if left_last.checked_add(1) != Some(target) || target.checked_add(1) != Some(right_first) {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "MulRk inverse adjacent records do not bracket the restored cell",
            ));
        }
        left_items
            .try_reserve(
                right_items
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("MulRk merge size overflow".into()))?,
            )
            .map_err(|_error| Error::Allocation("merging split MulRk records"))?;
        left_items.push((state.2.get(), encoded));
        left_items.extend(right_items);
        row.cell_records[left_index] = encode_mul_rk(reference.row(), left_first, &left_items)?;
        row.cell_records.remove(right_index);
        return Ok(());
    }
    if candidates.len() > 2 {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "MulRk inverse has ambiguous adjacent packed records",
        ));
    }
    let Some(index) = candidates.first().copied() else {
        // If a transaction removed every member of a packed source record,
        // the first semantic inverse has no surviving neighbor to merge with.
        // Restore one standalone RK in sorted row order; subsequent inverse
        // operations can merge adjacent RK records back into a packed family.
        let replacement = encode_rk_segment(reference.row(), target, &[(state.2.get(), encoded)])?;
        let mut insertion = row.cell_records.len();
        for (position, record) in row.cell_records.iter().enumerate() {
            let kind = record_kind(record)?;
            if is_cell_record(kind) && u16::from(cell_column(record)?) > target {
                insertion = position;
                break;
            }
        }
        row.cell_records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining standalone RK inverse"))?;
        row.cell_records.insert(insertion, replacement);
        return Ok(());
    };
    if record_kind(&row.cell_records[index])? == RK {
        let neighbor = &row.cell_records[index];
        let neighbor_column = binary::read_u16_le_at(neighbor, 6)?;
        let neighbor_style = binary::read_u16_le_at(neighbor, 8)?;
        let neighbor_rk = binary::read_u32_le_at(neighbor, 10)?;
        let (first_column, items) = if target < neighbor_column {
            (
                target,
                [(state.2.get(), encoded), (neighbor_style, neighbor_rk)],
            )
        } else {
            (
                neighbor_column,
                [(neighbor_style, neighbor_rk), (state.2.get(), encoded)],
            )
        };
        row.cell_records[index] = encode_mul_rk(reference.row(), first_column, &items)?;
        return Ok(());
    }
    let record = &mut row.cell_records[index];
    let payload_len = binary::read_u16_le_at(record, 2)?;
    let first = binary::read_u16_le_at(record, 6)?;
    let last = binary::read_u16_le_at(record, record.len() - 2)?;
    let mut item = Vec::with_capacity(6);
    item.extend_from_slice(&state.2.get().to_le_bytes());
    item.extend_from_slice(&encoded.to_le_bytes());
    if target < first {
        drop(record.splice(8..8, item));
        record[6..8].copy_from_slice(&target.to_le_bytes());
    } else if target > last {
        let offset = record.len() - 2;
        drop(record.splice(offset..offset, item));
        let end = record.len();
        record[end - 2..end].copy_from_slice(&target.to_le_bytes());
    } else {
        return Err(Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "MulRk inverse target is not an edge extension",
        ));
    }
    record[2..4].copy_from_slice(
        &payload_len
            .checked_add(6)
            .ok_or_else(|| Error::InvalidData("MulRk payload length overflow".into()))?
            .to_le_bytes(),
    );
    Ok(())
}

fn encode_mul_rk(row: u16, first_column: u16, items: &[(u16, u32)]) -> Result<Vec<u8>> {
    if items.is_empty() {
        return Err(Error::InvalidData(
            "MulRk record requires at least one item".into(),
        ));
    }
    let item_bytes = items
        .len()
        .checked_mul(6)
        .ok_or_else(|| Error::InvalidData("MulRk item length overflow".into()))?;
    let payload_len = 6usize
        .checked_add(item_bytes)
        .ok_or_else(|| Error::InvalidData("MulRk payload length overflow".into()))?;
    let last_column = first_column
        .checked_add(
            u16::try_from(items.len() - 1)
                .map_err(|_error| Error::InvalidData("MulRk column count exceeds u16".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("MulRk last column overflow".into()))?;
    let mut record = Vec::with_capacity(4 + payload_len);
    record.extend_from_slice(&MUL_RK.to_le_bytes());
    record.extend_from_slice(
        &u16::try_from(payload_len)
            .map_err(|_error| Error::InvalidData("MulRk payload exceeds u16".into()))?
            .to_le_bytes(),
    );
    record.extend_from_slice(&row.to_le_bytes());
    record.extend_from_slice(&first_column.to_le_bytes());
    for (style, rk) in items {
        record.extend_from_slice(&style.to_le_bytes());
        record.extend_from_slice(&rk.to_le_bytes());
    }
    record.extend_from_slice(&last_column.to_le_bytes());
    Ok(record)
}

fn decode_rk_items(record: &[u8]) -> Result<(u16, u16, Vec<(u16, u32)>)> {
    let kind = record_kind(record)?;
    let row = binary::read_u16_le_at(record, 4)?;
    let first = binary::read_u16_le_at(record, 6)?;
    if kind == RK {
        return Ok((
            row,
            first,
            vec![(
                binary::read_u16_le_at(record, 8)?,
                binary::read_u32_le_at(record, 10)?,
            )],
        ));
    }
    if kind != MUL_RK {
        return Err(Error::InvalidData(
            "packed RK segment has an unsupported record kind".into(),
        ));
    }
    let payload_len = usize::from(binary::read_u16_le_at(record, 2)?);
    let item_bytes = payload_len
        .checked_sub(6)
        .ok_or_else(|| Error::InvalidData("MulRk segment payload is truncated".into()))?;
    if item_bytes == 0 || !item_bytes.is_multiple_of(6) {
        return Err(Error::InvalidData(
            "MulRk segment item framing is invalid".into(),
        ));
    }
    let count = item_bytes / 6;
    let encoded_last = binary::read_u16_le_at(record, record.len() - 2)?;
    let expected_last = first
        .checked_add(
            u16::try_from(count - 1)
                .map_err(|_error| Error::InvalidData("MulRk segment count exceeds u16".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("MulRk segment last column overflows".into()))?;
    if encoded_last != expected_last {
        return Err(Error::InvalidData(
            "MulRk segment last column disagrees with its item count".into(),
        ));
    }
    let mut items = Vec::new();
    items
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation("retaining MulRk segment items"))?;
    for item in 0..count {
        let offset = 8 + item * 6;
        items.push((
            binary::read_u16_le_at(record, offset)?,
            binary::read_u32_le_at(record, offset + 2)?,
        ));
    }
    Ok((row, first, items))
}

fn encode_rk_segment(row: u16, first_column: u16, items: &[(u16, u32)]) -> Result<Vec<u8>> {
    match items {
        [(style, value)] => {
            let mut record = Vec::with_capacity(14);
            record.extend_from_slice(&RK.to_le_bytes());
            record.extend_from_slice(&10_u16.to_le_bytes());
            record.extend_from_slice(&row.to_le_bytes());
            record.extend_from_slice(&first_column.to_le_bytes());
            record.extend_from_slice(&style.to_le_bytes());
            record.extend_from_slice(&value.to_le_bytes());
            Ok(record)
        },
        [] => Err(Error::InvalidData("packed RK segment is empty".into())),
        _ => encode_mul_rk(row, first_column, items),
    }
}

fn replace_formula_cache(records: &mut Vec<Vec<u8>>, formula: usize, value: &Value) -> Result<()> {
    let Value::FormulaCache(cache) = value else {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "Formula record replacement requires a formula cache",
        ));
    };
    records[formula]
        .get_mut(10..18)
        .ok_or_else(|| Error::InvalidData("Formula cached-value field is truncated".into()))?
        .copy_from_slice(&super::encode_formula_cache(cache));

    let mut string_start = formula + 1;
    while string_start < records.len()
        && matches!(
            record_kind(&records[string_start])?,
            TABLE | SHARED_FORMULA | ARRAY
        )
    {
        string_start += 1;
    }
    let mut string_end = string_start;
    while string_end < records.len()
        && matches!(record_kind(&records[string_end])?, STRING | CONTINUE)
    {
        string_end += 1;
    }
    records.drain(string_start..string_end);
    if let super::FormulaCache::String(text) = cache {
        records.insert(string_start, encode_formula_string(text)?);
    }
    Ok(())
}

fn encode_formula_string(text: &str) -> Result<Vec<u8>> {
    let units: Vec<u16> = text.encode_utf16().collect();
    let count = u16::try_from(units.len()).map_err(|_error| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "formula string cache exceeds u16",
        )
    })?;
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let byte_len = units
        .len()
        .checked_mul(if compressed { 1 } else { 2 })
        .and_then(|length| length.checked_add(3))
        .ok_or_else(|| Error::InvalidData("formula String payload length overflow".into()))?;
    if byte_len > 8_224 {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "formula string cache requires unsupported Continue splitting",
        ));
    }
    let payload_len = u16::try_from(byte_len)
        .map_err(|_error| Error::InvalidData("formula String payload exceeds u16".into()))?;
    let mut record = Vec::with_capacity(4 + byte_len);
    record.extend_from_slice(&STRING.to_le_bytes());
    record.extend_from_slice(&payload_len.to_le_bytes());
    record.extend_from_slice(&count.to_le_bytes());
    record.push(u8::from(!compressed));
    if compressed {
        for unit in units {
            record.push(u8::try_from(unit).map_err(|_error| {
                Error::InvalidData("compressed formula string unit exceeds u8".into())
            })?);
        }
    } else {
        for unit in units {
            record.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(record)
}

fn shift_rows(
    rows: &mut BTreeMap<u16, RowData>,
    start: u16,
    count: u16,
    insert: bool,
) -> Result<()> {
    let mut shifted = BTreeMap::new();
    let delete_end = start.saturating_add(count);
    for (row, mut data) in std::mem::take(rows) {
        let target = if insert {
            if row < start {
                row
            } else {
                row.checked_add(count).ok_or_else(|| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "row insertion moves content outside BIFF8",
                    )
                })?
            }
        } else if row < start {
            row
        } else if row < delete_end {
            continue;
        } else {
            row - count
        };
        patch_row_number(&mut data, row, target, start, count, insert)?;
        shifted.insert(target, data);
    }
    *rows = shifted;
    Ok(())
}

fn shift_columns(
    rows: &mut BTreeMap<u16, RowData>,
    start: u8,
    count: u8,
    insert: bool,
) -> Result<()> {
    let delete_end = start.saturating_add(count);
    for data in rows.values_mut() {
        data.stored_extents = data.stored_extents.and_then(|extents| {
            shift_stored_extents(extents, u16::from(start), u16::from(count), insert)
        });
        let mut retained = Vec::new();
        for mut record in std::mem::take(&mut data.cell_records) {
            let kind = record_kind(&record)?;
            if matches!(kind, TABLE | ARRAY) {
                // The range moves whole (record 0316).
                let shift = AxisShift::Columns {
                    start,
                    count,
                    insert,
                };
                rules::move_owner_range(&mut record, shift, 0, 0)?;
                retained.push(record);
                continue;
            }
            if !is_cell_record(kind) {
                retained.push(record);
                continue;
            }
            if kind == MUL_RK || kind == MUL_BLANK {
                // A packed run moves whole when the shift point is outside
                // it; a run the point falls inside is not split (record
                // 0308).
                let length = record.len();
                if length < 10 {
                    return Err(Error::InvalidData("packed cell record is truncated".into()));
                }
                let first = binary::read_u16_le_at(&record, 6)?;
                let last = binary::read_u16_le_at(&record, length - 2)?;
                let (start16, count16) = (u16::from(start), u16::from(count));
                let delta: Option<i32> = if last < start16 {
                    Some(0)
                } else if insert && first >= start16 {
                    Some(i32::from(count16))
                } else if !insert && first >= start16 + count16 {
                    Some(-i32::from(count16))
                } else {
                    None
                };
                let Some(delta) = delta else {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "column movement across packed cell records is refused",
                    ));
                };
                let moved = |value: u16| {
                    u16::try_from(i32::from(value) + delta)
                        .ok()
                        .filter(|value| *value < 256)
                        .ok_or_else(|| {
                            Error::refused(
                                RefusalKind::InvalidValue,
                                "edit cells",
                                "column insertion moves content outside BIFF8",
                            )
                        })
                };
                let (first, last) = (moved(first)?, moved(last)?);
                record[6..8].copy_from_slice(&first.to_le_bytes());
                record[length - 2..].copy_from_slice(&last.to_le_bytes());
                retained.push(record);
                continue;
            }
            let column = cell_column(&record)?;
            let target = if insert {
                if column < start {
                    column
                } else {
                    column.checked_add(count).ok_or_else(|| {
                        Error::refused(
                            RefusalKind::InvalidValue,
                            "edit cells",
                            "column insertion moves content outside BIFF8",
                        )
                    })?
                }
            } else if column < start {
                column
            } else if column < delete_end {
                continue;
            } else {
                column - count
            };
            record[6..8].copy_from_slice(&u16::from(target).to_le_bytes());
            if kind == FORMULA {
                shift_formula_anchor(
                    &mut record,
                    AxisShift::Columns {
                        start,
                        count,
                        insert,
                    },
                )?;
            }
            retained.push(record);
        }
        data.cell_records = retained;
    }
    Ok(())
}

fn patch_row_number(
    data: &mut RowData,
    from: u16,
    row: u16,
    start: u16,
    count: u16,
    insert: bool,
) -> Result<()> {
    data.row_record
        .get_mut(4..6)
        .ok_or_else(|| Error::InvalidData("ROW record is truncated".into()))?
        .copy_from_slice(&row.to_le_bytes());
    for record in &mut data.cell_records {
        let kind = record_kind(record)?;
        if matches!(kind, TABLE | ARRAY) {
            // The range moves whole with its first row (record 0316).
            if from != row {
                let shift = AxisShift::Rows {
                    start,
                    count,
                    insert,
                };
                rules::move_owner_range(record, shift, from, row)?;
            }
            continue;
        }
        if !is_cell_record(kind) {
            continue;
        }
        record
            .get_mut(4..6)
            .ok_or_else(|| Error::InvalidData("cell record is truncated".into()))?
            .copy_from_slice(&row.to_le_bytes());
        if kind == FORMULA && from != row {
            shift_formula_anchor(
                record,
                AxisShift::Rows {
                    start,
                    count,
                    insert,
                },
            )?;
        }
    }
    Ok(())
}

fn patch_row_extents(rows: &mut BTreeMap<u16, RowData>) -> Result<()> {
    for data in rows.values_mut() {
        let (first, last_plus_one) = cell_extents(&data.cell_records)?;
        let (first, last_plus_one) = match data.stored_extents {
            Some((stored_first, stored_end))
                if last_plus_one == 0 || (stored_first <= first && last_plus_one <= stored_end) =>
            {
                (stored_first, stored_end)
            },
            _ => (first, last_plus_one),
        };
        data.row_record
            .get_mut(6..8)
            .ok_or_else(|| Error::InvalidData("ROW colMic is truncated".into()))?
            .copy_from_slice(&first.to_le_bytes());
        data.row_record
            .get_mut(8..10)
            .ok_or_else(|| Error::InvalidData("ROW colMac is truncated".into()))?
            .copy_from_slice(&last_plus_one.to_le_bytes());
    }
    Ok(())
}

/// The source DIMENSIONS `(rwMic, rwMac, colMic, colMac)` moved by the row
/// and column shifts of `operations`, or `None` when the record is not the
/// BIFF8 14-byte form or a shift leaves its range.
fn stored_dimensions(
    payload: &[u8],
    operations: &[&StructuralChange],
) -> Option<(u32, u32, u16, u16)> {
    if payload.len() != 14 {
        return None;
    }
    let mut rows = (
        binary::read_u32_le_at(payload, 0).ok()?,
        binary::read_u32_le_at(payload, 4).ok()?,
    );
    let mut columns = (
        u32::from(binary::read_u16_le_at(payload, 8).ok()?),
        u32::from(binary::read_u16_le_at(payload, 10).ok()?),
    );
    // Moves a `[first, end)` extent by one insertion or deletion.
    let shift = |(first, end): (u32, u32), start: u32, count: u32, insert: bool| {
        if insert {
            (
                if first >= start { first + count } else { first },
                if end > start { end + count } else { end },
            )
        } else {
            let moved = |value: u32| {
                if value >= start + count {
                    value - count
                } else {
                    value.min(start)
                }
            };
            (moved(first), moved(end))
        }
    };
    for operation in operations {
        match axis_shift(operation) {
            Some(AxisShift::Rows {
                start,
                count,
                insert,
            }) => rows = shift(rows, u32::from(start), u32::from(count), insert),
            Some(AxisShift::Columns {
                start,
                count,
                insert,
            }) => columns = shift(columns, u32::from(start), u32::from(count), insert),
            None => {},
        }
    }
    if rows.1 > 65_536 || columns.1 > 256 || rows.0 > rows.1 || columns.0 > columns.1 {
        return None;
    }
    Some((
        rows.0,
        rows.1,
        u16::try_from(columns.0).ok()?,
        u16::try_from(columns.1).ok()?,
    ))
}

fn dimensions_from_row_table(bytes: &[u8]) -> Result<(u32, u32, u16, u16)> {
    let records = raw_records(bytes)?;
    let mut first_row = u32::MAX;
    let mut last_row = 0_u32;
    let mut first_col = u16::MAX;
    let mut last_col = 0_u16;
    for record in records {
        if !is_cell_record(record.kind) {
            continue;
        }
        let raw = &bytes[record.start..record.end];
        let row = u32::from(binary::read_u16_le_at(raw, 4)?);
        let col = u16::from(cell_column(raw)?);
        let final_col = if matches!(record.kind, MUL_RK | MUL_BLANK) {
            binary::read_u16_le_at(raw, raw.len() - 2)?
        } else {
            col
        };
        first_row = first_row.min(row);
        last_row = last_row.max(row + 1);
        first_col = first_col.min(col);
        last_col = last_col.max(final_col + 1);
    }
    if first_row == u32::MAX {
        Ok((0, 0, 0, 0))
    } else {
        Ok((first_row, last_row, first_col, last_col))
    }
}

fn encode_cell(
    reference: Reference,
    state: &(Storage, Value, StyleIndex),
    source: &super::Snapshot,
    resources: &[super::ResourceChange],
    shared_strings: &[String],
) -> Result<Vec<u8>> {
    let (storage, value, style) = state;
    if usize::from(style.get()) >= source.inner.xf_records.len()
        && !resources.iter().any(|resource| {
            matches!(
                resource,
                super::ResourceChange::ExtendedFormat {
                    index,
                    insert: true,
                    ..
                } if index == style
            )
        })
    {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "inserted cell XF resource is stale",
        });
    }
    let mut record = Vec::new();
    match (storage, value) {
        (Storage::Number, Value::Number(number)) => crate::writer::biff::write_number(
            &mut record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            style.get(),
            *number,
        )?,
        (Storage::LabelSst, Value::Text(text)) => {
            let index = shared_strings
                .iter()
                .position(|candidate| candidate == text)
                .ok_or(Error::Internal("inserted text is absent from SST"))?;
            crate::writer::biff::write_labelsst(
                &mut record,
                u32::from(reference.row()),
                u16::from(reference.column()),
                style.get(),
                u32::try_from(index)
                    .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?,
            )?;
        },
        (Storage::BoolErr, Value::Boolean(value)) => crate::writer::biff::write_boolerr(
            &mut record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            style.get(),
            *value,
        )?,
        (Storage::BoolErr, Value::Error(error)) => {
            record = scalar_record(BOOL_ERR, reference, style, &[error.code(), 1]);
        },
        (Storage::Blank, Value::Blank) => {
            record = scalar_record(BLANK, reference, style, &[]);
        },
        _ => {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "inserted cell value has no standalone BIFF8 encoding",
            ));
        },
    }
    Ok(record)
}

fn scalar_record(kind: u16, reference: Reference, style: &StyleIndex, tail: &[u8]) -> Vec<u8> {
    let length = 6_u16 + u16::try_from(tail.len()).unwrap_or(0);
    let mut record = Vec::with_capacity(4 + usize::from(length));
    record.extend_from_slice(&kind.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(&reference.row().to_le_bytes());
    record.extend_from_slice(&u16::from(reference.column()).to_le_bytes());
    record.extend_from_slice(&style.get().to_le_bytes());
    record.extend_from_slice(tail);
    record
}

fn verify_record_state(
    record: &[u8],
    reference: Reference,
    expected: &(Storage, Value, StyleIndex),
    source: &super::Snapshot,
) -> Result<()> {
    let kind = record_kind(record)?;
    if !record_storage_matches(kind, expected.0) || record_style(record, reference)? != expected.2 {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "structural cell state is stale",
        });
    }
    // Snapshot parsing already validated the semantic value. Reconfirm text
    // resources because the structural path can change Workbook offsets.
    if let Value::Text(text) = &expected.1
        && !source
            .inner
            .shared_strings
            .iter()
            .any(|candidate| candidate == text)
    {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "structural cell SST dependency is stale",
        });
    }
    Ok(())
}

fn record_storage_matches(kind: u16, expected: Storage) -> bool {
    match expected {
        // A packed numeric record becomes a standalone RK when its final
        // neighboring member is removed. Likewise, a packed blank record can
        // reduce to a standalone BLANK. Both forms retain the same cell
        // state and are valid intermediate owners for a multi-cell edit.
        Storage::MulRk => matches!(kind, RK | MUL_RK),
        Storage::Blank => matches!(kind, BLANK | MUL_BLANK),
        _ => kind == super::storage_record_kind(expected),
    }
}

fn patch_record_style(record: &mut [u8], reference: Reference, style: StyleIndex) -> Result<()> {
    if matches!(record_kind(record)?, MUL_RK | MUL_BLANK) {
        let first = binary::read_u16_le_at(record, 6)?;
        let index = usize::from(
            u16::from(reference.column())
                .checked_sub(first)
                .ok_or_else(|| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "MulRk style target precedes packed range",
                    )
                })?,
        );
        let item_size = if record_kind(record)? == MUL_RK { 6 } else { 2 };
        let offset = 8usize
            .checked_add(
                index
                    .checked_mul(item_size)
                    .ok_or_else(|| Error::InvalidData("packed style offset overflow".into()))?,
            )
            .ok_or_else(|| Error::InvalidData("MulRk style offset overflow".into()))?;
        record
            .get_mut(offset..offset + 2)
            .ok_or_else(|| Error::InvalidData("MulRk style field is truncated".into()))?
            .copy_from_slice(&style.get().to_le_bytes());
    } else {
        record
            .get_mut(8..10)
            .ok_or_else(|| Error::InvalidData("cell XF field is truncated".into()))?
            .copy_from_slice(&style.get().to_le_bytes());
    }
    Ok(())
}

fn record_style(record: &[u8], reference: Reference) -> Result<StyleIndex> {
    let kind = record_kind(record)?;
    let index = if matches!(kind, MUL_RK | MUL_BLANK) {
        let first = binary::read_u16_le_at(record, 6)?;
        let index = usize::from(
            u16::from(reference.column())
                .checked_sub(first)
                .ok_or_else(|| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "MulRk style target precedes packed range",
                    )
                })?,
        );
        let item_size = if kind == MUL_RK { 6 } else { 2 };
        binary::read_u16_le_at(record, 8 + index * item_size)?
    } else {
        binary::read_u16_le_at(record, 8)?
    };
    Ok(StyleIndex(index))
}

fn record_contains_reference(record: &[u8], reference: Reference) -> Result<bool> {
    let kind = record_kind(record)?;
    if !is_cell_record(kind) {
        return Ok(false);
    }
    if binary::read_u16_le_at(record, 4)? != reference.row() {
        return Ok(false);
    }
    let first = binary::read_u16_le_at(record, 6)?;
    let column = u16::from(reference.column());
    if matches!(kind, MUL_RK | MUL_BLANK) {
        let last = binary::read_u16_le_at(record, record.len() - 2)?;
        Ok((first..=last).contains(&column))
    } else {
        Ok(first == column)
    }
}

fn cell_column(record: &[u8]) -> Result<u8> {
    u8::try_from(binary::read_u16_le_at(record, 6)?)
        .map_err(|_error| Error::InvalidData("cell column exceeds BIFF8".into()))
}

fn certify_coordinate_shift(
    worksheet: &[u8],
    records: &[RawRecord],
    shift: AxisShift,
) -> Result<()> {
    const DEPENDENT: &[u16] = &[0x009e, 0x0041, 0x0090];
    if records
        .iter()
        .any(|record| DEPENDENT.contains(&record.kind))
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot close dependency record",
        ));
    }
    let drawings = drawing_ranges(records);
    if !drawings.is_empty() {
        let sizes = sizes::SizeCheck::predict(worksheet, records, shift)?;
        shifted_drawing_stream(worksheet, &drawings, shift, Some(&sizes))?;
    }
    let mut has_rules = false;
    for record in records {
        let payload_start = record
            .start
            .checked_add(4)
            .ok_or_else(|| Error::InvalidData("dependency payload offset overflow".into()))?;
        let payload = worksheet
            .get(payload_start..record.end)
            .ok_or_else(|| Error::InvalidData("dependency record payload is truncated".into()))?;
        match record.kind {
            FORMULA => certify_mutation(payload, shift, patch_formula_record)?,
            SHARED_FORMULA => certify_mutation(payload, shift, patch_shared_formula)?,
            ARRAY => certify_mutation(payload, shift, rules::patch_array)?,
            TABLE => certify_mutation(payload, shift, rules::patch_table)?,
            kind if rules::RULE_RECORDS.contains(&kind) => has_rules = true,
            SELECTION => certify_mutation(payload, shift, patch_selection_shift)?,
            MERGED_CELLS => certify_mutation(payload, shift, patch_merged_cells)?,
            HLINK => certify_mutation(payload, shift, patch_hyperlink)?,
            HLINK_TOOLTIP => certify_mutation(payload, shift, patch_hyperlink_tooltip)?,
            OBJ => certify_simple_drawing_obj(payload)?,
            COL_INFO => certify_mutation(payload, shift, patch_col_info)?,
            NOTE => certify_mutation(payload, shift, patch_note)?,
            VERTICAL_PAGE_BREAKS => certify_mutation(payload, shift, patch_vertical_breaks)?,
            HORIZONTAL_PAGE_BREAKS => {
                certify_mutation(payload, shift, patch_horizontal_breaks)?;
            },
            _ => {},
        }
    }
    if has_rules {
        // Conditional-format rules are read against the ranges of the
        // records before them, so the rules are certified together.
        let mut candidate = Vec::new();
        candidate
            .try_reserve_exact(worksheet.len())
            .map_err(|_error| Error::Allocation("certifying conditional formats"))?;
        candidate.extend_from_slice(worksheet);
        rules::patch_rule_records(&mut candidate, records, shift)?;
    }
    Ok(())
}

/// Refuses a hyperlink whose location can name a cell of the shifted
/// worksheet (record 0308): a location on another sheet of this workbook,
/// a defined name (which moves with its own record) and a location inside
/// another document do not.
fn certify_hyperlink_target(
    payload: &[u8],
    limits: &SourceLimits,
    in_target: bool,
    target_name: &str,
) -> Result<()> {
    let hyperlink = crate::hyperlinks::parse_hlink_record(payload, limits)?;
    let Some(location) = hyperlink.location() else {
        return Ok(());
    };
    if hyperlink.moniker().is_some() || !location_reaches(location, in_target, target_name) {
        return Ok(());
    }
    Err(Error::refused(
        RefusalKind::UnsupportedOperation,
        "edit cells",
        "row/column movement refuses a hyperlink with an internal location dependency",
    ))
}

/// Whether a hyperlink location of this workbook can name a cell of the
/// worksheet `target_name` (`in_target`: the hyperlink is on that sheet).
fn location_reaches(location: &str, in_target: bool, target_name: &str) -> bool {
    let location = location.trim().trim_start_matches('#');
    let Some((sheet, _cell)) = location.rsplit_once('!') else {
        return in_target && looks_like_reference(location);
    };
    let sheet = sheet.trim();
    let sheet = match sheet
        .strip_prefix('\'')
        .and_then(|inner| inner.strip_suffix('\''))
    {
        Some(inner) => inner.replace("''", "'"),
        None => sheet.to_owned(),
    };
    if sheet.contains(':') || sheet.contains('[') {
        return true;
    }
    sheet.eq_ignore_ascii_case(target_name)
}

/// Whether `text` reads as an A1 cell, range, row or column reference
/// rather than a defined name.
fn looks_like_reference(text: &str) -> bool {
    // `(letters, digits)` of one `$`-optional part such as `$B$12`.
    let split = |part: &str| {
        let part = part.trim_start_matches('$');
        let letters = part.bytes().take_while(u8::is_ascii_alphabetic).count();
        let rest = part
            .get(letters..)
            .unwrap_or_default()
            .trim_start_matches('$');
        let digits = rest.bytes().all(|byte| byte.is_ascii_digit());
        (letters, rest.len(), digits)
    };
    let cell = |part: &str| {
        let (letters, rest, digits) = split(part);
        (1..=3).contains(&letters) && rest > 0 && digits
    };
    let line = |part: &str| {
        let (letters, rest, digits) = split(part);
        ((1..=3).contains(&letters) && rest == 0) || (letters == 0 && rest > 0 && digits)
    };
    let reference = match text.split_once(':') {
        None => cell(text),
        Some((first, last)) => (cell(first) && cell(last)) || (line(first) && line(last)),
    };
    // R1C1-style and anything unusual are treated as references.
    reference
        || text
            .bytes()
            .any(|byte| !(byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.'))
}

fn certify_mutation(
    payload: &[u8],
    shift: AxisShift,
    patch: fn(&mut [u8], AxisShift) -> Result<()>,
) -> Result<()> {
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(payload.len())
        .map_err(|_error| Error::Allocation("certifying coordinate dependency"))?;
    candidate.extend_from_slice(payload);
    patch(&mut candidate, shift)
}

fn record_payload<'a>(bytes: &'a [u8], record: &RawRecord, label: &str) -> Result<&'a [u8]> {
    let start = record
        .start
        .checked_add(4)
        .ok_or_else(|| Error::InvalidData(format!("{label} payload offset overflow")))?;
    bytes
        .get(start..record.end)
        .ok_or_else(|| Error::InvalidData(format!("{label} payload is truncated")))
}

fn formula_tokens(payload: &[u8]) -> Result<&[u8]> {
    if payload.len() < 22 {
        return Err(Error::InvalidData("Formula record is truncated".into()));
    }
    let (tokens, extra) = crate::formula_metadata::formula_payload_parts(payload)
        .map_err(|error| Error::InvalidData(error.to_string()))?;
    if !extra.is_empty() {
        crate::formula_metadata::validate_formula_extra(tokens, extra).map_err(|error| {
            Error::InvalidData(format!("Formula RgbExtra is malformed: {error}"))
        })?;
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement refuses Formula records with RgbExtra ancillary references",
        ));
    }
    Ok(tokens)
}

/// The cell a Formula's `PtgExp` or `PtgTbl` (`[MS-XLS]` 2.5.198.58,
/// 2.5.198.91) names: the first cell of its shared formula, array formula or
/// data table. Such a formula holds no other token.
fn formula_anchor(tokens: &[u8]) -> Option<(u16, u16)> {
    crate::formula::ptg_exp_anchor(tokens).or_else(|| {
        (tokens.len() == 5 && tokens[0] == 0x02 && tokens[4] == 0).then(|| {
            (
                u16::from_le_bytes([tokens[1], tokens[2]]),
                u16::from_le_bytes([tokens[3], tokens[4]]),
            )
        })
    })
}

/// Shifts the references of a Formula record of the rebuilt row table. The
/// `PtgExp`/`PtgTbl` anchor of a shared, array or table cell already moved
/// with its row (`shift_formula_anchor`), so the row-block planner sees an
/// `Array` or `Table` record next to its anchor (record 0316).
fn patch_formula_references(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let token_length = formula_tokens(payload)?.len();
    let tokens = payload
        .get_mut(22..22 + token_length)
        .ok_or_else(|| Error::InvalidData("Formula tokens are truncated".into()))?;
    if formula_anchor(tokens).is_some() {
        return Ok(());
    }
    patch_formula_tokens(tokens, shift)
}

/// Moves the `PtgExp`/`PtgTbl` anchor of a Formula cell record (`record`,
/// header included) under `shift`; other formulas are unchanged.
fn shift_formula_anchor(record: &mut [u8], shift: AxisShift) -> Result<()> {
    let payload = record
        .get_mut(4..)
        .ok_or_else(|| Error::InvalidData("Formula record is truncated".into()))?;
    let Some(tokens) = payload.get(22..) else {
        return Err(Error::InvalidData("Formula record is truncated".into()));
    };
    let length = usize::from(binary::read_u16_le_at(payload, 20)?);
    if tokens.len() < length || formula_anchor(&tokens[..length]).is_none() {
        return Ok(());
    }
    patch_formula_record(payload, shift)
}

fn patch_formula_record(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let token_length = formula_tokens(payload)?.len();
    let tokens = payload
        .get_mut(22..22 + token_length)
        .ok_or_else(|| Error::InvalidData("Formula tokens are truncated".into()))?;
    if formula_anchor(tokens).is_some() {
        // A cell of a shared or array formula: its PtgExp names the range's
        // first cell, which moves with the range (record 0308); a data-table
        // cell's PtgTbl names the table's first cell likewise (record 0316).
        return match shift {
            AxisShift::Rows {
                start,
                count,
                insert,
            } => patch_formula_row(tokens, 1, start, count, insert),
            AxisShift::Columns {
                start,
                count,
                insert,
            } => {
                let column = binary::read_u16_le_at(tokens, 3)?;
                let column =
                    shift_axis_point(column, u16::from(start), u16::from(count), insert, 256)?;
                write_formula_u16(tokens, 3, column)
            },
        };
    }
    patch_formula_tokens(tokens, shift)
}

/// Shifts a `ShrFmla` record (`[MS-XLS]` 2.4.260): its range moves whole
/// (a shift point inside it would split it and is refused), absolute
/// references shift like ordinary ones, and a relative `PtgRefN`/`PtgAreaN`
/// component stays when every cell and the cell it names lie on the same
/// side of the shift point. 3-D references are closed by the workbook pass
/// (record 0316). Record 0308.
fn patch_shared_formula(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() < 10 {
        return Err(Error::InvalidData("ShrFmla record is truncated".into()));
    }
    let token_length = usize::from(binary::read_u16_le_at(payload, 8)?);
    let (range_first, range_last, start, count, insert, limit) = match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => (
            binary::read_u16_le_at(payload, 0)?,
            binary::read_u16_le_at(payload, 2)?,
            start,
            count,
            insert,
            65_536_u32,
        ),
        AxisShift::Columns {
            start,
            count,
            insert,
        } => (
            u16::from(payload[4]),
            u16::from(payload[5]),
            u16::from(start),
            u16::from(count),
            insert,
            256,
        ),
    };
    let end = start.saturating_add(count);
    let moves = if range_last < start {
        false
    } else if (insert && range_first >= start) || (!insert && range_first >= end) {
        true
    } else {
        return Err(Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "a row/column shift through a shared formula's range would split it",
        ));
    };
    let tokens = payload
        .get_mut(10..10 + token_length)
        .ok_or_else(|| Error::InvalidData("ShrFmla tokens are truncated".into()))?;
    // 3-D references, including those naming this worksheet, are closed by
    // `patch_workbook_dependencies` over the whole stream (record 0316).
    let (same_sheet, relative, _three_d) = crate::formula::shared_formula_reference_spans(tokens)
        .ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "row/column movement cannot read a shared formula's references",
        )
    })?;
    for reference in same_sheet {
        patch_reference_span(tokens, reference, shift)?;
    }
    // A relative offset must keep naming the same cell for every cell of
    // the range.
    for reference in relative {
        match reference {
            crate::formula::FormulaReferenceSpan::Cell { row, column } => {
                rules::patch_relative_component(
                    tokens,
                    row,
                    column,
                    shift,
                    range_first,
                    range_last,
                )?;
            },
            crate::formula::FormulaReferenceSpan::Area {
                first_row,
                last_row,
                first_column,
                last_column,
            } => {
                for (row, column) in [(first_row, first_column), (last_row, last_column)] {
                    rules::patch_relative_component(
                        tokens,
                        row,
                        column,
                        shift,
                        range_first,
                        range_last,
                    )?;
                }
            },
        }
    }
    if moves {
        let shifted = |value: u16| shift_axis_point(value, start, count, insert, limit);
        match shift {
            AxisShift::Rows { .. } => {
                let (first, last) = (shifted(range_first)?, shifted(range_last)?);
                payload[0..2].copy_from_slice(&first.to_le_bytes());
                payload[2..4].copy_from_slice(&last.to_le_bytes());
            },
            AxisShift::Columns { .. } => {
                let (first, last) = (shifted(range_first)?, shifted(range_last)?);
                payload[4] = u8::try_from(first).map_err(|_error| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "column outside BIFF8",
                    )
                })?;
                payload[5] = u8::try_from(last).map_err(|_error| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "column outside BIFF8",
                    )
                })?;
            },
        }
    }
    Ok(())
}

fn patch_formula_tokens(tokens: &mut [u8], shift: AxisShift) -> Result<()> {
    // 3-D references, including those naming this worksheet, are closed by
    // `patch_workbook_dependencies` over the whole stream.
    let references = crate::formula::formula_reference_spans(tokens).ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement supports only ordinary Ref/Area and 3-D Formula tokens",
        )
    })?;
    for reference in references.same_sheet {
        patch_reference_span(tokens, reference, shift)?;
    }
    Ok(())
}

fn patch_reference_span(
    tokens: &mut [u8],
    reference: crate::formula::FormulaReferenceSpan,
    shift: AxisShift,
) -> Result<()> {
    match (reference, shift) {
        (
            crate::formula::FormulaReferenceSpan::Cell { row, .. },
            AxisShift::Rows {
                start,
                count,
                insert,
            },
        ) => patch_formula_row(tokens, row, start, count, insert),
        (
            crate::formula::FormulaReferenceSpan::Cell { column, .. },
            AxisShift::Columns {
                start,
                count,
                insert,
            },
        ) => patch_formula_column(tokens, column, start, count, insert),
        (
            crate::formula::FormulaReferenceSpan::Area {
                first_row,
                last_row,
                ..
            },
            AxisShift::Rows {
                start,
                count,
                insert,
            },
        ) => patch_formula_row_range(tokens, first_row, last_row, start, count, insert),
        (
            crate::formula::FormulaReferenceSpan::Area {
                first_column,
                last_column,
                ..
            },
            AxisShift::Columns {
                start,
                count,
                insert,
            },
        ) => patch_formula_column_range(tokens, first_column, last_column, start, count, insert),
    }
}

fn patch_formula_row(
    tokens: &mut [u8],
    offset: usize,
    start: u16,
    count: u16,
    insert: bool,
) -> Result<()> {
    let row = binary::read_u16_le_at(tokens, offset)?;
    let row = shift_axis_point(row, start, count, insert, 65_536)?;
    write_formula_u16(tokens, offset, row)
}

fn patch_formula_row_range(
    tokens: &mut [u8],
    first_offset: usize,
    last_offset: usize,
    start: u16,
    count: u16,
    insert: bool,
) -> Result<()> {
    let first = binary::read_u16_le_at(tokens, first_offset)?;
    let last = binary::read_u16_le_at(tokens, last_offset)?;
    let (first, last) = shift_axis_range(first, last, start, count, insert, 65_536)?;
    write_formula_u16(tokens, first_offset, first)?;
    write_formula_u16(tokens, last_offset, last)
}

fn patch_formula_column(
    tokens: &mut [u8],
    offset: usize,
    start: u8,
    count: u8,
    insert: bool,
) -> Result<()> {
    let flags = binary::read_u16_le_at(tokens, offset)?;
    let column = flags & 0x00ff;
    let column = shift_axis_point(column, u16::from(start), u16::from(count), insert, 256)?;
    write_formula_u16(tokens, offset, (flags & 0xff00) | column)
}

fn patch_formula_column_range(
    tokens: &mut [u8],
    first_offset: usize,
    last_offset: usize,
    start: u8,
    count: u8,
    insert: bool,
) -> Result<()> {
    let first_flags = binary::read_u16_le_at(tokens, first_offset)?;
    let last_flags = binary::read_u16_le_at(tokens, last_offset)?;
    let (first, last) = shift_axis_range(
        first_flags & 0x00ff,
        last_flags & 0x00ff,
        u16::from(start),
        u16::from(count),
        insert,
        256,
    )?;
    write_formula_u16(tokens, first_offset, (first_flags & 0xff00) | first)?;
    write_formula_u16(tokens, last_offset, (last_flags & 0xff00) | last)
}

fn write_formula_u16(bytes: &mut [u8], offset: usize, value: u16) -> Result<()> {
    bytes
        .get_mut(offset..offset + 2)
        .ok_or_else(|| Error::InvalidData("Formula reference field is truncated".into()))?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn certify_simple_drawing_obj(payload: &[u8]) -> Result<()> {
    let mut offset = 0_usize;
    let mut subrecord = 0_usize;
    loop {
        let kind = binary::read_u16_le_at(payload, offset)?;
        let length_offset = offset
            .checked_add(2)
            .ok_or_else(|| Error::InvalidData("Obj length offset overflow".into()))?;
        let length = usize::from(binary::read_u16_le_at(payload, length_offset)?);
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
            .ok_or_else(|| Error::InvalidData("Obj subrecord range overflow".into()))?;
        if end > payload.len() {
            return Err(Error::InvalidData("Obj subrecord is truncated".into()));
        }
        match (subrecord, kind, length) {
            (0, 0x0015, 18) => {},
            (_, 0x0006, 2) if subrecord == 1 => {},
            // FtNts (a comment's note data), FtCf (picture clipboard
            // format) and FtPioGrbit (picture flags) hold no cell
            // coordinates ([MS-XLS] 2.5.200, 2.5.140, 2.5.200; record 0308).
            (1.., 0x000d, 22) | (1.., 0x0007 | 0x0008, 2) => {},
            (_, 0, 0) if end == payload.len() => return Ok(()),
            _ => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement only supports coordinate-free drawing Obj records",
                ));
            },
        }
        offset = end;
        subrecord = subrecord
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("Obj subrecord count overflow".into()))?;
    }
}

pub(super) fn replace_range_and_adjust_bounds(
    workbook: &mut Vec<u8>,
    start: usize,
    end: usize,
    replacement: &[u8],
) -> Result<()> {
    let old_len = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidData("replacement range is reversed".into()))?;
    let delta = i64::try_from(replacement.len())
        .ok()
        .and_then(|new| {
            i64::try_from(old_len)
                .ok()
                .and_then(|old| new.checked_sub(old))
        })
        .ok_or_else(|| Error::InvalidData("Workbook replacement delta overflow".into()))?;
    let old_bounds = bound_sheets(workbook)?;
    workbook.splice(start..end, replacement.iter().copied());
    if delta == 0 {
        return Ok(());
    }
    let threshold = u32::try_from(end)
        .map_err(|_error| Error::InvalidData("Workbook replacement offset exceeds u32".into()))?;
    let shifted_bounds: Vec<_> = old_bounds
        .iter()
        .filter(|bound| bound.position >= threshold)
        .map(|bound| bound.position)
        .collect();
    for bound in old_bounds {
        if bound.position < threshold {
            continue;
        }
        let record_start = adjusted_offset(bound.record.start, start, end, delta)?;
        let position = i64::from(bound.position)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "BoundSheet position adjustment overflows",
                )
            })?;
        workbook
            .get_mut(record_start + 4..record_start + 8)
            .ok_or_else(|| Error::InvalidData("BoundSheet position field is truncated".into()))?
            .copy_from_slice(&position.to_le_bytes());
    }
    for old_position in shifted_bounds {
        let position = i64::from(old_position)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "worksheet position adjustment overflows",
                )
            })?;
        adjust_worksheet_index_offsets(workbook, position, delta)?;
    }
    Ok(())
}

/// Applies worksheet-local replacements in descending source order, then
/// repairs every affected `BoundSheet8` and downstream `INDEX` exactly once.
///
/// This narrower batch primitive deliberately rejects workbook-global ranges:
/// its callers must already have proved that every replacement is inside a
/// worksheet and after that worksheet's final `DBCELL`.
pub(super) fn replace_ranges_and_adjust_bounds(
    workbook: &mut Vec<u8>,
    replacements: &[(usize, usize, Vec<u8>)],
) -> Result<()> {
    if replacements.is_empty() {
        return Ok(());
    }
    let old_bounds = bound_sheets(workbook)?;
    let first_sheet = old_bounds
        .iter()
        .try_fold(None, |first: Option<usize>, bound| {
            let position = usize::try_from(bound.position)
                .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
            Ok::<_, Error>(Some(first.map_or(position, |value| value.min(position))))
        })?
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "worksheet replacement requires a BoundSheet",
            )
        })?;
    let mut prior_start = workbook.len();
    let mut deltas = Vec::new();
    deltas
        .try_reserve_exact(replacements.len())
        .map_err(|_error| Error::Allocation("retaining worksheet replacement deltas"))?;
    for (start, end, replacement) in replacements {
        if *start < first_sheet || *start > *end || *end > workbook.len() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "batched replacement is outside a worksheet source range",
            ));
        }
        if *end > prior_start {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "batched worksheet replacement ranges overlap or are not descending",
            ));
        }
        let old_len = end
            .checked_sub(*start)
            .ok_or_else(|| Error::InvalidData("replacement range is reversed".into()))?;
        let delta = i64::try_from(replacement.len())
            .ok()
            .and_then(|new| {
                i64::try_from(old_len)
                    .ok()
                    .and_then(|old| new.checked_sub(old))
            })
            .ok_or_else(|| Error::InvalidData("Workbook replacement delta overflow".into()))?;
        deltas.push((*end, delta));
        prior_start = *start;
    }

    let total_delta = deltas.iter().try_fold(0_i64, |total, (_, delta)| {
        total
            .checked_add(*delta)
            .ok_or_else(|| Error::InvalidData("Workbook replacement total overflow".into()))
    })?;
    let target_len = i64::try_from(workbook.len())
        .ok()
        .and_then(|length| length.checked_add(total_delta))
        .and_then(|length| usize::try_from(length).ok())
        .ok_or_else(|| Error::InvalidData("Workbook replacement length overflows".into()))?;
    let mut rebuilt = Vec::new();
    rebuilt
        .try_reserve_exact(target_len)
        .map_err(|_error| Error::Allocation("publishing batched worksheet replacements"))?;
    let mut cursor = 0_usize;
    for (start, end, replacement) in replacements.iter().rev() {
        rebuilt.extend_from_slice(workbook.get(cursor..*start).ok_or_else(|| {
            Error::InvalidData("batched replacement prefix is outside Workbook".into())
        })?);
        rebuilt.extend_from_slice(replacement);
        cursor = *end;
    }
    rebuilt.extend_from_slice(workbook.get(cursor..).ok_or_else(|| {
        Error::InvalidData("batched replacement suffix is outside Workbook".into())
    })?);
    if rebuilt.len() != target_len {
        return Err(Error::InvalidData(
            "batched Workbook replacement length disagrees with its plan".into(),
        ));
    }
    *workbook = rebuilt;

    for bound in old_bounds {
        let before = usize::try_from(bound.position)
            .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
        let shift = deltas.iter().try_fold(0_i64, |total, (end, delta)| {
            if *end <= before {
                total
                    .checked_add(*delta)
                    .ok_or_else(|| Error::InvalidData("worksheet shift overflow".into()))
            } else {
                Ok(total)
            }
        })?;
        if shift == 0 {
            continue;
        }
        let position = i64::from(bound.position)
            .checked_add(shift)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "BoundSheet position adjustment overflows",
                )
            })?;
        workbook
            .get_mut(bound.record.start + 4..bound.record.start + 8)
            .ok_or_else(|| Error::InvalidData("BoundSheet position field is truncated".into()))?
            .copy_from_slice(&position.to_le_bytes());
        adjust_worksheet_index_offsets(workbook, position, shift)?;
    }
    Ok(())
}

fn adjust_worksheet_index_offsets(
    workbook: &mut [u8],
    worksheet_position: u32,
    delta: i64,
) -> Result<()> {
    let worksheet_start = usize::try_from(worksheet_position)
        .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
    let records = raw_records(
        workbook
            .get(worksheet_start..)
            .ok_or_else(|| Error::InvalidData("worksheet position is outside Workbook".into()))?,
    )?;
    let index = records
        .iter()
        .take_while(|record| record.kind != 0x000a)
        .find(|record| record.kind == INDEX);
    let Some(index) = index else {
        return Ok(());
    };
    let payload_len = index
        .end
        .checked_sub(index.start + 4)
        .ok_or_else(|| Error::InvalidData("INDEX payload range is invalid".into()))?;
    if payload_len < 16 || (payload_len - 16) % 4 != 0 {
        return Err(Error::InvalidData("INDEX payload is malformed".into()));
    }
    for relative in (12..payload_len).step_by(4) {
        let offset = worksheet_start
            .checked_add(index.start + 4 + relative)
            .ok_or_else(|| Error::InvalidData("INDEX field offset overflow".into()))?;
        let before = binary::read_u32_le_at(workbook, offset)?;
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
                    "INDEX absolute offset adjustment overflows",
                )
            })?;
        workbook
            .get_mut(offset..offset + 4)
            .ok_or_else(|| Error::InvalidData("INDEX absolute offset is truncated".into()))?
            .copy_from_slice(&after.to_le_bytes());
    }
    Ok(())
}

fn adjusted_offset(offset: usize, start: usize, end: usize, delta: i64) -> Result<usize> {
    if offset < start {
        return Ok(offset);
    }
    if offset >= end {
        return usize::try_from(
            i64::try_from(offset)
                .map_err(|_error| Error::InvalidData("record offset exceeds i64".into()))?
                .checked_add(delta)
                .ok_or_else(|| Error::InvalidData("record offset adjustment overflows".into()))?,
        )
        .map_err(|_error| Error::InvalidData("adjusted record offset exceeds usize".into()));
    }
    // The only supported overlap is replacement of that exact BoundSheet
    // record during a sheet rename; its rebuilt record starts at `start`.
    Ok(start)
}

fn bound_sheets(workbook: &[u8]) -> Result<Vec<BoundSheet>> {
    bound_sheets_in(workbook, &raw_records(workbook)?)
}

/// The BoundSheet records among `records`, the records of `workbook`.
fn bound_sheets_in(workbook: &[u8], records: &[RawRecord]) -> Result<Vec<BoundSheet>> {
    let mut bounds = Vec::new();
    for record in records.iter().cloned() {
        if record.kind == 0x000a {
            break;
        }
        if record.kind != BOUND_SHEET {
            continue;
        }
        let payload = workbook
            .get(record.start + 4..record.end)
            .ok_or_else(|| Error::InvalidData("BoundSheet payload is truncated".into()))?;
        if payload.len() < 8 {
            return Err(Error::InvalidData("BoundSheet payload is truncated".into()));
        }
        bounds.push(BoundSheet {
            position: binary::read_u32_le_at(payload, 0)?,
            record,
        });
    }
    Ok(bounds)
}

fn decode_bound_name(payload: &[u8]) -> Result<String> {
    if payload.len() < 8 {
        return Err(Error::InvalidData("BoundSheet name is truncated".into()));
    }
    let units = usize::from(payload[6]);
    let wide = payload[7] & 1 != 0;
    if wide {
        let byte_len = units
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidData("BoundSheet name length overflow".into()))?;
        let bytes = payload
            .get(8..8 + byte_len)
            .ok_or_else(|| Error::InvalidData("wide BoundSheet name is truncated".into()))?;
        let utf16: Vec<u16> = bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16(&utf16)
            .map_err(|error| Error::InvalidData(format!("invalid BoundSheet UTF-16: {error}")))
    } else {
        let bytes = payload
            .get(8..8 + units)
            .ok_or_else(|| Error::InvalidData("compressed BoundSheet name is truncated".into()))?;
        Ok(bytes.iter().map(|byte| char::from(*byte)).collect())
    }
}

fn encode_bound_name(name: &str) -> Result<Vec<u8>> {
    let units: Vec<u16> = name.encode_utf16().collect();
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let mut encoded = Vec::new();
    encoded.push(u8::from(!compressed));
    if compressed {
        for unit in units {
            encoded.push(u8::try_from(unit).map_err(|_error| {
                Error::InvalidData("compressed BoundSheet character exceeds u8".into())
            })?);
        }
    } else {
        for unit in units {
            encoded.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(encoded)
}

/// Every BIFF record frame of `bytes`, stopping at stream padding.
///
/// Producers pad the Workbook stream after its final EOF with zero bytes
/// (Excel and LibreOffice round small streams up to the 4,096-byte
/// mini-stream cutoff). That tail is no record: a zero run reads as empty
/// records of type 0 and then a truncated header. The scan stops where only
/// zero bytes remain, and the callers keep that tail as untouched bytes.
fn raw_records(bytes: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    let mut end_of_last = 0_usize;
    let padding_from = |offset: usize| {
        bytes
            .get(offset..)
            .is_some_and(|tail| tail.iter().all(|byte| *byte == 0))
    };
    for item in Records::new(bytes, litchi_biff::Limits::default())? {
        let record = match item {
            Ok(record) => record,
            Err(_) if end_of_last > 0 && padding_from(end_of_last) => break,
            Err(error) => return Err(error.into()),
        };
        let start = record.offset();
        if record.kind().get() == 0 && start > 0 && padding_from(start) {
            break;
        }
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        records.push(RawRecord {
            kind: record.kind().get(),
            start,
            end,
        });
        end_of_last = end;
    }
    Ok(records)
}

/// A record of a record list, with its position in the worksheet's list.
trait TopRecord {
    fn kind(&self) -> u16;
    fn position(&self, index: usize) -> usize;
}

impl TopRecord for RawRecord {
    fn kind(&self) -> u16 {
        self.kind
    }

    fn position(&self, index: usize) -> usize {
        index
    }
}

impl TopRecord for (usize, RawRecord) {
    fn kind(&self) -> u16 {
        self.1.kind
    }

    fn position(&self, _index: usize) -> usize {
        self.0
    }
}

/// The records of `records` outside embedded substreams (a nested
/// `BOF` through its `EOF`), with their positions in `records`.
fn top_level(records: &[RawRecord]) -> Vec<(usize, RawRecord)> {
    let mut depth = 0_usize;
    let mut top = Vec::with_capacity(records.len());
    for (position, record) in records.iter().enumerate() {
        match record.kind {
            0x0809 => {
                depth += 1;
                if depth == 1 {
                    top.push((position, record.clone()));
                }
            },
            0x000a => {
                if depth == 1 {
                    top.push((position, record.clone()));
                }
                depth = depth.saturating_sub(1);
            },
            _ if depth <= 1 => top.push((position, record.clone())),
            _ => {},
        }
    }
    top
}

fn unique_kind<R: TopRecord>(records: &[R], kind: u16, _label: &str) -> Result<usize> {
    let mut matches = records
        .iter()
        .enumerate()
        .filter(|(_, record)| record.kind() == kind);
    let index = matches
        .next()
        .map(|(index, record)| record.position(index))
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "the worksheet lacks a required record",
            )
        })?;
    if matches.next().is_some() {
        return Err(Error::refused(
            RefusalKind::UnsupportedContent,
            "edit cells",
            "the worksheet has a duplicated singleton record",
        ));
    }
    Ok(index)
}

fn record_kind(record: &[u8]) -> Result<u16> {
    Ok(binary::read_u16_le_at(record, 0)?)
}

fn is_cell_record(kind: u16) -> bool {
    matches!(
        kind,
        FORMULA | BLANK | NUMBER | LABEL | BOOL_ERR | RK | MUL_RK | MUL_BLANK | RSTRING | LABEL_SST
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn raw_record(kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut record = Vec::new();
        record.extend_from_slice(&kind.to_le_bytes());
        record.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
        record.extend_from_slice(payload);
        record
    }

    #[test]
    fn xf_resource_refuses_stale_crc_and_count_overflow_before_mutation() {
        fn globals(count: u16) -> Vec<u8> {
            let mut bytes = Vec::new();
            for _ in 0..count {
                bytes.extend(raw_record(XF, &[0; 20]));
            }
            let mut crc = [0; 20];
            crc[..2].copy_from_slice(&crate::number_format::XFCRC_RECORD.to_le_bytes());
            crc[14..16].copy_from_slice(&count.to_le_bytes());
            bytes.extend(raw_record(crate::number_format::XFCRC_RECORD, &crc));
            bytes.extend(raw_record(0x000a, &[]));
            bytes
        }
        let mut at_limit = globals(4050);
        let original = at_limit.clone();
        assert!(matches!(
            apply_xf_resource(&mut at_limit, StyleIndex(4050), &[0; 20], true),
            Err(Error::TooMany { limit: 4050, .. })
        ));
        assert_eq!(at_limit, original);

        let mut stale = globals(16);
        stale[16 * 24 + 20] = 1;
        let original = stale.clone();
        assert!(matches!(
            apply_xf_resource(&mut stale, StyleIndex(16), &[0; 20], true),
            Err(Error::Refused(r)) if r.kind() == RefusalKind::UnsupportedContent
        ));
        assert_eq!(stale, original);
    }

    #[test]
    fn ranges_and_move_size_anchors_shift_and_inverse_exactly() {
        let insert_rows = AxisShift::Rows {
            start: 3,
            count: 2,
            insert: true,
        };
        let delete_rows = AxisShift::Rows {
            start: 3,
            count: 2,
            insert: false,
        };
        let original_range = [1, 0, 5, 0, 2, 0, 4, 0];
        let mut range = original_range;
        patch_ref8(&mut range, insert_rows).unwrap();
        patch_ref8(&mut range, delete_rows).unwrap();
        assert_eq!(range, original_range);

        // "Move and size with cells" (flags 0) shifts like cells.
        let original_anchor = [0, 0, 2, 0, 0, 0, 1, 0, 0, 0, 4, 0, 0, 0, 5, 0, 0, 0];
        let mut anchor = original_anchor;
        patch_client_anchor(&mut anchor, insert_rows, None).unwrap();
        assert_eq!(&anchor[14..16], &[7, 0]);
        patch_client_anchor(&mut anchor, delete_rows, None).unwrap();
        assert_eq!(anchor, original_anchor);
        // "Don't move or size" (3) is refused when the shift reaches it and
        // unchanged after it; "move but don't size" (2) moves whole only
        // when the shift point is outside it (record 0308).
        let mut locked = original_anchor;
        locked[0] = 3;
        assert!(patch_client_anchor(&mut locked, insert_rows, None).is_err());
        let late = AxisShift::Rows {
            start: 9,
            count: 2,
            insert: true,
        };
        patch_client_anchor(&mut locked, late, None).unwrap();
        assert_eq!(&locked[6..8], &[1, 0]);
        let mut moving = original_anchor;
        moving[0] = 2;
        assert!(patch_client_anchor(&mut moving, insert_rows, None).is_err());
        let early = AxisShift::Rows {
            start: 0,
            count: 2,
            insert: true,
        };
        patch_client_anchor(&mut moving, early, None).unwrap();
        assert_eq!((&moving[6..8], &moving[14..16]), (&[3, 0][..], &[7, 0][..]));

        let insert_columns = AxisShift::Columns {
            start: 3,
            count: 2,
            insert: true,
        };
        let delete_columns = AxisShift::Columns {
            start: 3,
            count: 2,
            insert: false,
        };
        let mut range = original_range;
        patch_ref8(&mut range, insert_columns).unwrap();
        patch_ref8(&mut range, delete_columns).unwrap();
        assert_eq!(range, original_range);
        let mut anchor = original_anchor;
        patch_client_anchor(&mut anchor, insert_columns, None).unwrap();
        patch_client_anchor(&mut anchor, delete_columns, None).unwrap();
        assert_eq!(anchor, original_anchor);

        let destructive = AxisShift::Rows {
            start: 4,
            count: 2,
            insert: false,
        };
        let mut destructive_range = original_range;
        assert!(patch_ref8(&mut destructive_range, destructive).is_err());
    }

    const INSERT_ROWS: AxisShift = AxisShift::Rows {
        start: 3,
        count: 2,
        insert: true,
    };
    const DELETE_ROWS: AxisShift = AxisShift::Rows {
        start: 3,
        count: 2,
        insert: false,
    };
    const INSERT_COLUMNS: AxisShift = AxisShift::Columns {
        start: 3,
        count: 2,
        insert: true,
    };
    const DELETE_COLUMNS: AxisShift = AxisShift::Columns {
        start: 3,
        count: 2,
        insert: false,
    };

    fn round_trip(
        original: &[u8],
        patch: fn(&mut [u8], AxisShift) -> Result<()>,
        forward: AxisShift,
        back: AxisShift,
    ) -> Vec<u8> {
        let mut bytes = original.to_vec();
        patch(&mut bytes, forward).unwrap();
        let shifted = bytes.clone();
        patch(&mut bytes, back).unwrap();
        assert_eq!(bytes, original);
        shifted
    }

    #[test]
    fn zero_padding_after_the_final_eof_is_not_a_record() {
        for padding in [1, 3, 4, 9] {
            let mut stream = raw_record(0x0809, &[0; 16]);
            stream.extend(raw_record(EOF, &[]));
            stream.extend(std::iter::repeat_n(0, padding));
            let records = raw_records(&stream).unwrap();
            assert_eq!(records.len(), 2, "padding {padding}");
            assert_eq!(records[1].kind, EOF);
        }
        let mut garbage = raw_record(0x0809, &[0; 16]);
        garbage.extend([0, 0, 7]);
        assert!(raw_records(&garbage).is_err());
    }

    #[test]
    fn col_info_shifts_with_columns_and_keeps_the_grid_end() {
        // Columns D:F (3..=5): an insertion at D moves the span.
        let span = [3, 0, 5, 0, 0x6d, 0x08, 0x0f, 0, 0, 0, 0, 0];
        let shifted = round_trip(&span, patch_col_info, INSERT_COLUMNS, DELETE_COLUMNS);
        assert_eq!(&shifted[..4], &[5, 0, 7, 0]);
        // Rows never touch a ColInfo.
        round_trip(&span, patch_col_info, INSERT_ROWS, DELETE_ROWS);
        // A span through the grid end (B:IV) keeps its end.
        let open = [1, 0, 0xff, 0, 0x6d, 0x08, 0x0f, 0, 0, 0, 0, 0];
        let shifted = round_trip(&open, patch_col_info, INSERT_COLUMNS, DELETE_COLUMNS);
        assert_eq!(&shifted[..4], &[1, 0, 0xff, 0]);
        // Deleting the open span's first column is refused.
        let mut first = [3, 0, 0, 1, 0x6d, 0x08, 0x0f, 0, 0, 0, 0, 0];
        assert!(patch_col_info(&mut first, DELETE_COLUMNS).is_err());
    }

    #[test]
    fn notes_and_page_breaks_shift_and_inverse_exactly() {
        let note = [7, 0, 4, 0, 0, 0, 1, 0];
        let shifted = round_trip(&note, patch_note, INSERT_ROWS, DELETE_ROWS);
        assert_eq!(&shifted[..4], &[9, 0, 4, 0]);
        round_trip(&note, patch_note, INSERT_COLUMNS, DELETE_COLUMNS);
        let mut erased = note;
        assert!(
            patch_note(
                &mut erased,
                AxisShift::Rows {
                    start: 7,
                    count: 1,
                    insert: false,
                }
            )
            .is_err()
        );

        // A horizontal break before row 10 spanning every column.
        let horizontal = [1, 0, 10, 0, 0, 0, 0xff, 0];
        let shifted = round_trip(
            &horizontal,
            patch_horizontal_breaks,
            INSERT_ROWS,
            DELETE_ROWS,
        );
        assert_eq!(&shifted[2..4], &[12, 0]);
        let shifted = round_trip(
            &horizontal,
            patch_horizontal_breaks,
            INSERT_COLUMNS,
            DELETE_COLUMNS,
        );
        assert_eq!(shifted, horizontal);
        // A vertical break before column E spanning every row.
        let vertical = [1, 0, 4, 0, 0, 0, 0xff, 0xff];
        let shifted = round_trip(
            &vertical,
            patch_vertical_breaks,
            INSERT_COLUMNS,
            DELETE_COLUMNS,
        );
        assert_eq!(&shifted[2..4], &[6, 0]);
        round_trip(&vertical, patch_vertical_breaks, INSERT_ROWS, DELETE_ROWS);
    }

    #[test]
    fn whole_axis_formula_ranges_keep_the_grid_end() {
        // PtgArea A:A ($A$1:$A$65536), a PtgRef to row 6, and PtgAdd.
        let original = [
            0x25, 0, 0, 0xff, 0xff, 0, 0xc0, 0, 0xc0, 0x24, 5, 0, 1, 0xc0, 0x03,
        ];
        let mut tokens = original;
        patch_formula_tokens(&mut tokens, INSERT_ROWS).unwrap();
        assert_eq!(&tokens[1..5], &[0, 0, 0xff, 0xff]);
        assert_eq!(&tokens[10..12], &[7, 0]);
        patch_formula_tokens(&mut tokens, DELETE_ROWS).unwrap();
        assert_eq!(tokens, original);
    }

    /// A Workbook stream with a self `SupBook`, two XTIs (sheet 0 and sheet
    /// 1), a print-area `Lbl` on sheet 0 and a formula pointing at sheet 1.
    fn three_d_workbook() -> Vec<u8> {
        let mut stream = raw_record(
            0x0809,
            &[0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        );
        stream.extend(raw_record(SUP_BOOK, &[2, 0, 0x01, 0x04]));
        stream.extend(raw_record(
            EXTERN_SHEET,
            &[2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0],
        ));
        // Lbl: builtin Print_Area (0x06), sheet-scoped (itab 1), formula
        // PtgArea3d ixti 0 $A$4:$C$9.
        let mut name = vec![0x20, 0, 0, 1, 11, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0x06];
        name.extend([0x3b, 0, 0, 3, 0, 8, 0, 0, 0, 2, 0]);
        stream.extend(raw_record(LBL, &name));
        stream.extend(raw_record(EOF, &[]));
        // A sheet formula =Sheet2!$B$7 (PtgRef3d ixti 1).
        let mut formula = vec![0; 20];
        formula.extend([7, 0, 0x3a, 1, 0, 6, 0, 1, 0]);
        stream.extend(raw_record(FORMULA, &formula));
        stream
    }

    #[test]
    fn defined_names_and_3d_formulas_follow_their_sheet_and_inverse_exactly() {
        let original = three_d_workbook();
        let mut workbook = original.clone();
        let limits = SourceLimits::default();
        patch_workbook_dependencies(&mut workbook, 0, &limits, INSERT_ROWS).unwrap();
        let name_start = original.len() - 4 - 29 - 4 - 4 - 27;
        // $A$4:$C$9 on sheet 0 became $A$6:$C$11; the formula on sheet 1 kept $B$7.
        assert_eq!(
            &workbook[name_start + 4 + 16..name_start + 4 + 27],
            &[0x3b, 0, 0, 5, 0, 10, 0, 0, 0, 2, 0]
        );
        assert_eq!(
            workbook[workbook.len() - 6..],
            original[original.len() - 6..]
        );
        patch_workbook_dependencies(&mut workbook, 0, &limits, DELETE_ROWS).unwrap();
        assert_eq!(workbook, original);

        // The same shift on sheet 1 moves only the formula's reference.
        patch_workbook_dependencies(&mut workbook, 1, &limits, INSERT_ROWS).unwrap();
        assert_eq!(&workbook[workbook.len() - 4..workbook.len() - 2], &[8, 0]);
        patch_workbook_dependencies(&mut workbook, 1, &limits, DELETE_ROWS).unwrap();
        assert_eq!(workbook, original);

        // Deleting a row the name names is refused.
        let mut erasing = original.clone();
        assert!(
            patch_workbook_dependencies(
                &mut erasing,
                0,
                &limits,
                AxisShift::Rows {
                    start: 8,
                    count: 1,
                    insert: false,
                }
            )
            .is_err()
        );
    }

    #[test]
    fn producer_row_extents_survive_a_shift_and_its_inverse() {
        // A Row record for row 0 declaring columns 0..5 with one cell in
        // column 1: a producer's sheet-wide extents.
        let row = raw_record(ROW, &[0, 0, 0, 0, 5, 0, 0xff, 0, 0, 0, 0, 0, 0, 1, 0x0f, 0]);
        let cell = raw_record(0x0201, &[0, 0, 1, 0, 0x0f, 0]);
        let mut worksheet = row.clone();
        worksheet.extend(&cell);
        let mut rows = collect_rows(&worksheet, &raw_records(&worksheet).unwrap()).unwrap();
        assert_eq!(rows[&0].stored_extents, Some((0, 5)));
        shift_columns(&mut rows, 0, 1, true).unwrap();
        patch_row_extents(&mut rows).unwrap();
        assert_eq!(row_extents_field(&rows[&0].row_record).unwrap(), (1, 6));
        shift_columns(&mut rows, 0, 1, false).unwrap();
        patch_row_extents(&mut rows).unwrap();
        assert_eq!(rows[&0].row_record, row);
    }

    #[test]
    fn ext_sst_tail_resource_updates_bucket_and_inverts_exactly() {
        let mut workbook = raw_record(0x0809, &[0; 12]);
        let sst_start = workbook.len();
        let mut sst_payload = Vec::new();
        sst_payload.extend_from_slice(&8_u32.to_le_bytes());
        sst_payload.extend_from_slice(&8_u32.to_le_bytes());
        for _ in 0..8 {
            sst_payload.extend_from_slice(&[0, 0, 0]);
        }
        workbook.extend_from_slice(&raw_record(super::super::SST, &sst_payload));
        let bucket = crate::shared_string_index::SharedStringBucket::try_new(
            u32::try_from(sst_start + 12).unwrap(),
            12,
        )
        .unwrap();
        let ext = crate::shared_string_index::SharedStringIndex::try_new(8, vec![bucket])
            .unwrap()
            .to_record_bytes()
            .unwrap();
        workbook.extend_from_slice(&ext);
        workbook.extend_from_slice(&raw_record(0x000a, &[]));
        let original = workbook.clone();

        apply_shared_string_resource(&mut workbook, "ninth", &[], true).unwrap();
        let globals = workbook_globals(&workbook).unwrap();
        let ext = globals
            .iter()
            .find(|record| record.kind == EXT_SST)
            .unwrap();
        let index = crate::shared_string_index::SharedStringIndex::parse_payload(
            9,
            &workbook[ext.start + 4..ext.end],
        )
        .unwrap();
        assert_eq!(index.buckets().len(), 2);
        apply_shared_string_resource(&mut workbook, "ninth", &[], false).unwrap();
        assert_eq!(workbook, original);
    }

    #[test]
    fn interior_mulrk_deletion_splits_and_inverse_merges_exactly() {
        let original = encode_mul_rk(
            7,
            2,
            &[
                (0, super::super::encode_rk(1.0).unwrap()),
                (1, super::super::encode_rk(2.0).unwrap()),
                (2, super::super::encode_rk(3.0).unwrap()),
                (3, super::super::encode_rk(4.0).unwrap()),
                (4, super::super::encode_rk(5.0).unwrap()),
            ],
        )
        .unwrap();
        let reference = Reference::new(7, 4).unwrap();
        let mut records = vec![original.clone()];
        remove_cell_record(&mut records, 0, reference).unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(binary::read_u16_le_at(&records[0], 6).unwrap(), 2);
        assert_eq!(
            binary::read_u16_le_at(&records[0], records[0].len() - 2).unwrap(),
            3
        );
        assert_eq!(binary::read_u16_le_at(&records[1], 6).unwrap(), 5);
        assert_eq!(
            binary::read_u16_le_at(&records[1], records[1].len() - 2).unwrap(),
            6
        );

        let mut rows = BTreeMap::from([(
            7,
            RowData {
                row_record: Vec::new(),
                cell_records: records,
                stored_extents: None,
            },
        )]);
        insert_packed_rk(
            &mut rows,
            reference,
            &(Storage::MulRk, Value::Number(3.0), StyleIndex(2)),
        )
        .unwrap();
        assert_eq!(rows[&7].cell_records, [original]);
    }

    #[test]
    fn mulblank_deletions_preserve_interior_and_edge_styles() {
        let original = encode_blank_segment(7, 2, &[10, 11, 12, 13, 14]).unwrap();
        let target = Reference::new(7, 4).unwrap();
        let mut records = vec![original.clone()];
        remove_cell_record(&mut records, 0, target).unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(record_kind(&records[0]).unwrap(), MUL_BLANK);
        assert_eq!(record_kind(&records[1]).unwrap(), MUL_BLANK);
        assert_eq!(decode_blank_styles(&records[0]).unwrap(), vec![10, 11]);
        assert_eq!(decode_blank_styles(&records[1]).unwrap(), vec![13, 14]);
        assert!(record_contains_reference(&records[0], Reference::new(7, 2).unwrap()).unwrap());
        assert!(record_contains_reference(&records[1], Reference::new(7, 6).unwrap()).unwrap());
        assert_eq!(
            record_style(&records[0], Reference::new(7, 3).unwrap()).unwrap(),
            StyleIndex(11)
        );
        assert_eq!(
            record_style(&records[1], Reference::new(7, 5).unwrap()).unwrap(),
            StyleIndex(13)
        );

        // Each split side can then be reduced at an edge without dropping its
        // remaining neighbor or changing that neighbor's XF.
        remove_cell_record(&mut records, 0, Reference::new(7, 2).unwrap()).unwrap();
        remove_cell_record(&mut records, 1, Reference::new(7, 6).unwrap()).unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(record_kind(&records[0]).unwrap(), BLANK);
        assert_eq!(record_kind(&records[1]).unwrap(), BLANK);
        assert_eq!(
            record_style(&records[0], Reference::new(7, 3).unwrap()).unwrap(),
            StyleIndex(11)
        );
        assert_eq!(
            record_style(&records[1], Reference::new(7, 5).unwrap()).unwrap(),
            StyleIndex(13)
        );
    }

    #[test]
    fn several_mulrk_deletions_accept_reduced_standalone_owners() {
        let original = encode_mul_rk(
            7,
            2,
            &[
                (0, super::super::encode_rk(1.0).unwrap()),
                (1, super::super::encode_rk(2.0).unwrap()),
                (2, super::super::encode_rk(3.0).unwrap()),
                (3, super::super::encode_rk(4.0).unwrap()),
                (4, super::super::encode_rk(5.0).unwrap()),
            ],
        )
        .unwrap();
        let mut records = vec![original];
        for column in [4, 2, 6, 3, 5] {
            let reference = Reference::new(7, column).unwrap();
            let index = records
                .iter()
                .position(|record| record_contains_reference(record, reference).unwrap())
                .expect("remaining packed member");
            remove_cell_record(&mut records, index, reference).unwrap();
        }
        assert!(records.is_empty());
    }

    #[test]
    fn packed_inverse_can_build_a_family_from_standalone_records() {
        let mut rows = BTreeMap::from([(
            7,
            RowData {
                row_record: Vec::new(),
                cell_records: Vec::new(),
                stored_extents: None,
            },
        )]);
        for (column, value, style) in [(2, 1.0, 0), (3, 2.0, 1), (4, 3.0, 2)] {
            insert_packed_rk(
                &mut rows,
                Reference::new(7, column).unwrap(),
                &(Storage::MulRk, Value::Number(value), StyleIndex(style)),
            )
            .unwrap();
        }
        let items = decode_rk_items(&rows[&7].cell_records[0]).unwrap();
        assert_eq!(items.1, 2);
        assert_eq!(items.2.len(), 3);
    }

    #[test]
    fn selection_coordinates_follow_row_insertion_and_inverse_deletion() {
        let mut payload = vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5];
        let insert = StructuralChange::Rows {
            sheet: 0,
            start: 2,
            count: 2,
            insert: true,
        };
        patch_selection_shift(&mut payload, axis_shift(&insert).unwrap()).unwrap();
        assert_eq!(binary::read_u16_le_at(&payload, 1).unwrap(), 5);
        assert_eq!(binary::read_u16_le_at(&payload, 9).unwrap(), 5);
        assert_eq!(binary::read_u16_le_at(&payload, 11).unwrap(), 6);

        let delete = StructuralChange::Rows {
            sheet: 0,
            start: 2,
            count: 2,
            insert: false,
        };
        patch_selection_shift(&mut payload, axis_shift(&delete).unwrap()).unwrap();
        assert_eq!(payload, vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5]);
    }

    #[test]
    fn selection_records_skip_scalar_changes_but_keep_frame_validation() {
        let payload = vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5];
        let mut worksheet = raw_record(SELECTION, &payload);
        let original = worksheet.clone();
        let scalar = StructuralChange::Cell {
            sheet: 0,
            reference: Reference { row: 3, column: 2 },
            before: None,
            after: None,
        };

        patch_selection_records(&mut worksheet, &[&scalar]).unwrap();
        assert_eq!(worksheet, original);

        let mut malformed = vec![(SELECTION & 0xff) as u8, (SELECTION >> 8) as u8, 9, 0, 0];
        assert!(patch_selection_records(&mut malformed, &[&scalar]).is_err());
    }

    #[test]
    fn selection_records_apply_axis_changes_after_scalar_changes() {
        let payload = vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5];
        let mut worksheet = raw_record(SELECTION, &payload);
        let scalar = StructuralChange::Cell {
            sheet: 0,
            reference: Reference { row: 3, column: 2 },
            before: None,
            after: None,
        };
        let rows = StructuralChange::Rows {
            sheet: 0,
            start: 2,
            count: 2,
            insert: true,
        };

        patch_selection_records(&mut worksheet, &[&scalar, &rows]).unwrap();
        assert_eq!(
            &worksheet[4..],
            [3, 5, 0, 2, 0, 0, 0, 1, 0, 5, 0, 6, 0, 2, 5]
        );
    }

    /// A `ShrFmla` over rows 5..=7 whose template is `=A{r-1}` (a
    /// row-relative `PtgRefN` of offset -1) and `+$B$2` (absolute).
    fn shared_formula(offset: i16) -> Vec<u8> {
        let mut payload = vec![5, 0, 7, 0, 2, 2, 0, 3];
        let tokens = [
            [0x2c].as_slice(),
            &offset.to_le_bytes(),
            &[0x00, 0xc0],
            &[0x24, 1, 0, 1, 0],
            &[0x03],
        ]
        .concat();
        payload.extend_from_slice(&u16::try_from(tokens.len()).unwrap().to_le_bytes());
        payload.extend_from_slice(&tokens);
        payload
    }

    #[test]
    fn shared_formulas_move_with_their_range_and_keep_relative_offsets() {
        let original = shared_formula(-1);
        let mut payload = original.clone();
        // Rows 3 and 4 inserted: the range (5..=7) and its targets (4..=6)
        // are below the shift point.
        patch_shared_formula(&mut payload, INSERT_ROWS).unwrap();
        assert_eq!(&payload[0..4], &[7, 0, 9, 0]);
        // The absolute $B$2 is above it and stays; the offset stays.
        assert_eq!(&payload[10..20], &original[10..20]);
        patch_shared_formula(&mut payload, DELETE_ROWS).unwrap();
        assert_eq!(payload, original);
        // An offset of -3 names rows 2..=4 for cells 5..=7: across the
        // point, so the shift is refused.
        assert!(patch_shared_formula(&mut shared_formula(-3), INSERT_ROWS).is_err());
        // A shift point inside the range would split it.
        let inside = AxisShift::Rows {
            start: 6,
            count: 1,
            insert: true,
        };
        assert!(patch_shared_formula(&mut shared_formula(-1), inside).is_err());
    }

    #[test]
    fn packed_runs_move_whole_outside_the_shift_point() {
        let mut row = RowData {
            row_record: raw_record(ROW, &[0; 16]),
            cell_records: vec![raw_record(MUL_BLANK, &[0, 0, 4, 0, 15, 0, 15, 0, 5, 0])],
            stored_extents: None,
        };
        let mut rows = BTreeMap::from([(0, row)]);
        shift_columns(&mut rows, 2, 2, true).unwrap();
        row = rows.remove(&0).unwrap();
        assert_eq!(&row.cell_records[0][4..], &[0, 0, 6, 0, 15, 0, 15, 0, 7, 0]);
        let mut rows = BTreeMap::from([(0, row)]);
        assert!(shift_columns(&mut rows, 7, 1, true).is_err());
    }

    #[test]
    fn hyperlink_locations_reach_only_the_shifted_sheet() {
        assert!(location_reaches("Data!B3", false, "data"));
        assert!(location_reaches("'Q1 ''24'!A1:B2", false, "Q1 '24"));
        assert!(!location_reaches("Other!B3", true, "Data"));
        assert!(location_reaches("B3", true, "Data"));
        assert!(!location_reaches("B3", false, "Data"));
        assert!(!location_reaches("SalesTotal", true, "Data"));
        assert!(location_reaches("A:A", true, "Data"));
        assert!(location_reaches("Sheet1:Sheet3!A1", false, "Data"));
    }
}

mod rules;
mod sizes;

#[cfg(test)]
#[path = "structural_sst_tests.rs"]
mod structural_sst_tests;
