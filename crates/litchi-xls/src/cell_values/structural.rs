use super::{
    BLANK, BOOL_ERR, Error, FORMULA, LABEL_SST, MUL_RK, NUMBER, RK, Reference, Result, Storage,
    StructuralChange, StyleIndex, Value,
};
use crate::writer::row_blocks::{RowBlockLayoutPlan, RowBlockLayoutRow};
use litchi_biff::Records;
use litchi_core::binary;
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
}

pub(super) fn apply(
    mut workbook: Vec<u8>,
    source: &super::Snapshot,
    changes: &[StructuralChange],
    fixed_changes: &[super::Change],
    resources: &[super::ResourceChange],
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
    inserted_resources.sort_by_key(|resource| super::resource_target(resource));
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
    // Rewriting later streams first keeps earlier absolute offsets stable and
    // minimizes the number of BoundSheet position adjustments.
    for (sheet, operations) in by_sheet.into_iter().rev() {
        rewrite_worksheet(
            &mut workbook,
            source,
            sheet,
            &operations,
            resources,
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
    removed_resources.sort_by_key(|resource| std::cmp::Reverse(super::resource_target(resource)));
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
            .ok_or_else(|| Error::UnsafeEdit("SST record disappeared from globals".into()))?;
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
            .ok_or_else(|| Error::UnsafeEdit("workbook globals have no EOF owner".into()))?;
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
        return Err(Error::UnsafeEdit(
            "ExtSST is not adjacent to its SST record family".into(),
        ));
    }

    let effective = effective_sst_entries(source, resources)?;
    if effective.len() != shared_strings.len()
        || effective
            .iter()
            .zip(shared_strings)
            .any(|(entry, text)| entry.text != *text)
    {
        return Err(Error::UnsafeEdit(
            "effective shared-string resource order is stale".into(),
        ));
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
                return Err(Error::UnsafeEdit(
                    "compacting an unreferenced rich or repaired SST entry would lose semantic inverse metadata".into(),
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
        let target = remap
            .get(old)
            .and_then(|value| *value)
            .ok_or_else(|| Error::UnsafeEdit("final LabelSst points to a dead SST entry".into()))?;
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
    for item in Records::new(&source.inner.workbook_stream) {
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
        return Err(Error::UnsafeEdit(
            "source shared strings have no SST owner".into(),
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
        let index = entries
            .iter()
            .position(|entry| entry.text == text)
            .ok_or_else(|| Error::UnsafeEdit("SST resource removal identity is stale".into()))?;
        if !sst_resource_removal_matches(&entries[index], resource)? {
            return Err(Error::UnsafeEdit(
                "SST resource removal identity is stale".into(),
            ));
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
            return Err(Error::UnsafeEdit(
                "SST resource removals must form a source-table suffix".into(),
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
    inserted.sort_by_key(|resource| super::resource_target(resource));
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

pub(super) fn certify_sst_insertion(
    source: &super::Snapshot,
    staged: &[super::ResourceChange],
    candidate: &super::ResourceChange,
) -> Result<()> {
    let mut workbook = Vec::new();
    workbook
        .try_reserve_exact(source.inner.workbook_stream.len())
        .map_err(|_error| Error::Allocation("certifying SST resource insertion"))?;
    workbook.extend_from_slice(&source.inner.workbook_stream);
    let insertion_count = staged
        .iter()
        .chain(std::iter::once(candidate))
        .filter(|resource| {
            matches!(
                resource,
                super::ResourceChange::SharedString { insert: true, .. }
                    | super::ResourceChange::RichSharedString { insert: true, .. }
            )
        })
        .count();
    let mut resources = Vec::new();
    resources
        .try_reserve_exact(insertion_count)
        .map_err(|_error| Error::Allocation("certifying SST resource insertion"))?;
    resources.extend(
        staged
            .iter()
            .chain(std::iter::once(candidate))
            .filter(|resource| {
                matches!(
                    resource,
                    super::ResourceChange::SharedString { insert: true, .. }
                        | super::ResourceChange::RichSharedString { insert: true, .. }
                )
            }),
    );
    resources.sort_by_key(|resource| super::resource_target(resource));
    if !workbook_globals(&workbook)?
        .iter()
        .any(|record| record.kind == super::SST)
    {
        for resource in resources {
            let (text, formatting_runs) = match resource {
                super::ResourceChange::SharedString { text, .. } => (text.as_str(), &[][..]),
                super::ResourceChange::RichSharedString {
                    text,
                    formatting_runs,
                    ..
                } => (text.as_str(), formatting_runs.as_slice()),
                _ => continue,
            };
            encode_sst_tail_records(text, formatting_runs)?;
        }
        return Ok(());
    }
    for resource in resources {
        apply_resource(&mut workbook, resource)?;
    }
    Ok(())
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
    certify_workbook_shift(workbook)?;
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or_else(|| Error::UnsafeEdit("worksheet dependency index is stale".into()))?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or_else(|| Error::UnsafeEdit("worksheet dependency owner disappeared".into()))?;
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
        return Err(Error::UnsafeEdit(
            "ExtSST is not adjacent to its SST record family".into(),
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
        return Err(Error::UnsafeEdit(
            "SST inverse cannot locate its complete tail resource".into(),
        ));
    }
    let tail = &family[family.len() - authored_count..];
    for (record, expected) in tail.iter().zip(&expected) {
        if &workbook[record.start..record.end] != expected.as_slice() {
            return Err(Error::UnsafeEdit(
                "SST inverse tail resource precondition is stale".into(),
            ));
        }
    }
    let removed_index = unique
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidData("SST unique count underflow".into()))?;
    if label_sst_reference_exists(workbook, removed_index)? {
        return Err(Error::UnsafeEdit(
            "SST inverse resource is still referenced by a LabelSst cell".into(),
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
        return Err(Error::UnsafeEdit(
            "SST authoring crosses an ExtSST bucket-size boundary".into(),
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
            return Err(Error::UnsafeEdit(
                "SST authoring changes unsupported ExtSST bucket geometry".into(),
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
    let count = u16::try_from(unit_count)
        .map_err(|_error| Error::UnsafeEdit("shared string exceeds u16 characters".into()))?;
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
        let character_count = u16::try_from(units.len())
            .map_err(|_error| Error::UnsafeEdit("shared string exceeds u16 characters".into()))?;
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
        let run_count = u16::try_from(runs.len())
            .map_err(|_error| Error::UnsafeEdit("SST rich-run count exceeds u16".into()))?;
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
        return Err(Error::UnsafeEdit(
            "ExtRst text exceeds signed character limit".into(),
        ));
    }
    let run_count = u16::try_from(phonetic.runs.len())
        .map_err(|_error| Error::InvalidData("ExtRst run count exceeds u16".into()))?;
    if run_count > 32_767 {
        return Err(Error::UnsafeEdit(
            "ExtRst run count exceeds signed limit".into(),
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
    if insert {
        if usize::from(index.get()) != xfs.len() {
            return Err(Error::UnsafeEdit(
                "new XF index is not the next resource".into(),
            ));
        }
        let last = xfs
            .last()
            .ok_or_else(|| Error::UnsafeEdit("workbook has no XF insertion owner".into()))?;
        let mut record = Vec::with_capacity(4 + payload.len());
        record.extend_from_slice(&XF.to_le_bytes());
        record.extend_from_slice(
            &u16::try_from(payload.len())
                .map_err(|_error| Error::InvalidData("XF payload exceeds u16".into()))?
                .to_le_bytes(),
        );
        record.extend_from_slice(payload);
        return replace_range_and_adjust_bounds(workbook, last.end, last.end, &record);
    }
    if usize::from(index.get()) + 1 != xfs.len() {
        return Err(Error::UnsafeEdit(
            "XF inverse can remove only the final resource".into(),
        ));
    }
    let last = xfs
        .last()
        .ok_or_else(|| Error::UnsafeEdit("workbook has no XF removal owner".into()))?;
    if &workbook[last.start + 4..last.end] != payload {
        return Err(Error::UnsafeEdit("XF inverse payload is stale".into()));
    }
    if xf_reference_exists(workbook, index.get())? {
        return Err(Error::UnsafeEdit(
            "XF inverse resource remains referenced by a cell".into(),
        ));
    }
    replace_range_and_adjust_bounds(workbook, last.start, last.end, &[])
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
        .ok_or_else(|| Error::UnsafeEdit("worksheet rename index is stale".into()))?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or_else(|| Error::UnsafeEdit("worksheet BoundSheet disappeared".into()))?;
    let payload = &workbook[bound.record.start + 4..bound.record.end];
    let decoded = decode_bound_name(payload)?;
    if decoded != before {
        return Err(Error::UnsafeEdit(
            "worksheet rename precondition is stale".into(),
        ));
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
    shared_strings: &[String],
) -> Result<()> {
    if operations.iter().any(|operation| {
        matches!(
            operation,
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. }
        )
    }) {
        certify_workbook_shift(workbook)?;
    }
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or_else(|| Error::UnsafeEdit("worksheet operation index is stale".into()))?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or_else(|| Error::UnsafeEdit("worksheet BoundSheet disappeared".into()))?;
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
        shared_strings,
    )?;
    replace_range_and_adjust_bounds(workbook, start, end, &replacement)
}

fn certify_workbook_shift(workbook: &[u8]) -> Result<()> {
    // These records can retain formulas or coordinates into any worksheet.
    // Without a complete token/range rewriter, moving coordinates would make
    // them stale even when the selected worksheet itself has no Formula cell.
    const GLOBAL_OR_CROSS_SHEET_DEPENDENCIES: &[u16] = &[
        TABLE,
        SHARED_FORMULA,
        ARRAY,
        0x0018,
        0x0023,
        0x1051,
        0x01b0,
        0x0879,
        0x087a,
        0x01be,
        0x009e,
    ];
    let records = raw_records(workbook)?;
    if let Some(record) = records
        .iter()
        .find(|record| GLOBAL_OR_CROSS_SHEET_DEPENDENCIES.contains(&record.kind))
    {
        return Err(Error::UnsafeEdit(format!(
            "row/column movement cannot close workbook dependency record 0x{:04X}",
            record.kind
        )));
    }
    for record in records.iter().filter(|record| record.kind == FORMULA) {
        let payload = record_payload(workbook, record, "Formula")?;
        certify_formula_owner(payload)?;
    }
    for record in records.iter().filter(|record| record.kind == HLINK) {
        let payload_start = record
            .start
            .checked_add(4)
            .ok_or_else(|| Error::InvalidData("HLINK payload offset overflow".into()))?;
        let payload = workbook
            .get(payload_start..record.end)
            .ok_or_else(|| Error::InvalidData("HLINK payload is truncated".into()))?;
        certify_hyperlink_target(payload)?;
    }
    Ok(())
}

fn rebuild_sheet(
    worksheet: &[u8],
    absolute_start: usize,
    sheet: usize,
    source: &super::Snapshot,
    operations: &[&StructuralChange],
    resources: &[super::ResourceChange],
    shared_strings: &[String],
) -> Result<Vec<u8>> {
    let records = raw_records(worksheet)?;
    let index = unique_kind(&records, INDEX, "INDEX")?;
    let dimensions = unique_kind(&records, DIMENSIONS, "DIMENSIONS")?;
    let def_col = unique_kind(&records, DEF_COL_WIDTH, "DEFCOLWIDTH")?;
    let first_row = records.iter().position(|record| record.kind == ROW);
    let last_dbcell = records.iter().rposition(|record| record.kind == DBCELL);
    let (row_records, row_table_start, row_table_end) = match (first_row, last_dbcell) {
        (Some(first), Some(last)) if first > index && last >= first => (
            &records[first..=last],
            records[first].start,
            records[last].end,
        ),
        (None, None) => {
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
                || records.iter().any(|record| {
                    is_cell_record(record.kind)
                        || matches!(
                            record.kind,
                            STRING | TABLE | ARRAY | SHARED_FORMULA | CONTINUE | 0x0091
                        )
                })
            {
                return Err(Error::UnsafeEdit(
                    "empty worksheet row-table metadata is inconsistent".into(),
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
    formula_resources.sort_by_key(|resource| super::resource_target(resource));
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
    patch_row_extents(&mut rows)?;

    let index_record = &records[index];
    let between = worksheet
        .get(index_record.end..row_table_start)
        .ok_or_else(|| Error::InvalidData("INDEX-to-row range is invalid".into()))?;
    let def_offset = records[def_col]
        .start
        .checked_sub(index_record.end)
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
        .checked_add(index_record.start)
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
    .map_err(|error| Error::UnsafeEdit(format!("BIFF8 row-block regeneration: {error}")))?;
    let (new_index, new_rows) = plan.into_records();

    let (first_used_row, last_used_row, first_used_col, last_used_col) =
        dimensions_from_row_table(&new_rows)?;
    let mut new_dimensions = Vec::new();
    crate::writer::biff::write_dimensions(
        &mut new_dimensions,
        first_used_row,
        last_used_row,
        first_used_col,
        last_used_col,
    )?;
    if new_dimensions.len() != records[dimensions].end - records[dimensions].start {
        return Err(Error::UnsafeEdit(
            "DIMENSIONS encoding changed fixed record width".into(),
        ));
    }
    let mut between = between.to_vec();
    let dimensions_relative = records[dimensions]
        .start
        .checked_sub(index_record.end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS precedes INDEX".into()))?;
    let dimensions_end = dimensions_relative
        .checked_add(new_dimensions.len())
        .ok_or_else(|| Error::InvalidData("DIMENSIONS range overflow".into()))?;
    between
        .get_mut(dimensions_relative..dimensions_end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS is outside INDEX gap".into()))?
        .copy_from_slice(&new_dimensions);

    let mut result = Vec::new();
    result.extend_from_slice(&worksheet[..index_record.start]);
    result.extend_from_slice(&new_index);
    result.extend_from_slice(&between);
    result.extend_from_slice(&new_rows);
    result.extend_from_slice(&worksheet[row_table_end..]);
    patch_coordinate_dependencies(&mut result, operations)?;
    Ok(result)
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
) -> Result<()> {
    patch_selection_records(worksheet, operations)?;
    let records = raw_records(worksheet)?;
    for operation in operations {
        let Some(shift) = axis_shift(operation) else {
            continue;
        };
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
                FORMULA => patch_formula_record(payload, shift)?,
                MERGED_CELLS => patch_merged_cells(payload, shift)?,
                HLINK => patch_hyperlink(payload, shift)?,
                HLINK_TOOLTIP => patch_hyperlink_tooltip(payload, shift)?,
                MSO_DRAWING => patch_officeart(payload, shift)?,
                _ => {},
            }
        }
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
            return Err(Error::UnsafeEdit(
                "row/column deletion would erase a dependency coordinate".into(),
            ));
        }
    };
    if shifted >= limit {
        return Err(Error::UnsafeEdit(
            "row/column insertion moves a dependency coordinate outside BIFF8".into(),
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

fn patch_officeart(data: &mut [u8], shift: AxisShift) -> Result<()> {
    patch_officeart_records(data, shift)
}

fn patch_officeart_records(data: &mut [u8], shift: AxisShift) -> Result<()> {
    let mut offset = 0_usize;
    while offset < data.len() {
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| Error::InvalidData("OfficeArt header range overflow".into()))?;
        let header = data
            .get(offset..header_end)
            .ok_or_else(|| Error::UnsafeEdit("split OfficeArt records are not shiftable".into()))?;
        let options = u16::from_le_bytes([header[0], header[1]]);
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let length = usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ]))
        .map_err(|_error| Error::InvalidData("OfficeArt length exceeds usize".into()))?;
        let end = header_end
            .checked_add(length)
            .ok_or_else(|| Error::InvalidData("OfficeArt record range overflow".into()))?;
        let payload = data
            .get_mut(header_end..end)
            .ok_or_else(|| Error::UnsafeEdit("split OfficeArt records are not shiftable".into()))?;
        let version = options & 0x000f;
        let instance = options >> 4;
        if kind == 0xf010 {
            if version != 0 || instance != 0 || payload.len() != 18 {
                return Err(Error::UnsafeEdit(
                    "worksheet drawing has a noncanonical ClientAnchor".into(),
                ));
            }
            patch_client_anchor(payload, shift)?;
        } else if version == 0x000f {
            patch_officeart_records(payload, shift)?;
        }
        offset = end;
    }
    Ok(())
}

fn patch_client_anchor(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let behavior = binary::read_u16_le_at(payload, 0)?;
    match behavior {
        0 => return Ok(()),
        0b10 => {
            return Err(Error::UnsafeEdit(
                "size-only drawing anchors are not reversibly shiftable".into(),
            ));
        },
        0b11 => {},
        _ => {
            return Err(Error::InvalidData(
                "worksheet ClientAnchor has reserved behavior flags".into(),
            ));
        },
    }
    match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => {
            let first = binary::read_u16_le_at(payload, 6)?;
            let last = binary::read_u16_le_at(payload, 14)?;
            let (first, last) = shift_axis_range(first, last, start, count, insert, 65_536)?;
            payload[6..8].copy_from_slice(&first.to_le_bytes());
            payload[14..16].copy_from_slice(&last.to_le_bytes());
        },
        AxisShift::Columns {
            start,
            count,
            insert,
        } => {
            let first = binary::read_u16_le_at(payload, 2)?;
            let last = binary::read_u16_le_at(payload, 10)?;
            let (first, last) =
                shift_axis_range(first, last, u16::from(start), u16::from(count), insert, 256)?;
            payload[2..4].copy_from_slice(&first.to_le_bytes());
            payload[10..12].copy_from_slice(&last.to_le_bytes());
        },
    }
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
                        },
                    )
                    .is_some()
                {
                    return Err(Error::UnsafeEdit(
                        "duplicate ROW record in row table".into(),
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
                return Err(Error::UnsafeEdit(format!(
                    "record 0x{:04X} inside the BIFF8 row table is not losslessly classified",
                    record.kind
                )));
            },
        }
    }
    Ok(rows)
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
                return Err(Error::UnsafeEdit(
                    "structural cell target is ambiguous".into(),
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
                    return Err(Error::UnsafeEdit(
                        "scalar replacement of a packed blank record is unsupported".into(),
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
            return Err(Error::UnsafeEdit(
                "cell insertion precondition is stale".into(),
            ));
        },
        (Some(_), _, None) => {
            return Err(Error::UnsafeEdit(
                "structural cell precondition is stale".into(),
            ));
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
                return Err(Error::UnsafeEdit(
                    "authored Formula target is ambiguous".into(),
                ));
            }
        }
    }
    if insert {
        if found.is_some() {
            return Err(Error::UnsafeEdit(
                "authored Formula insertion target is occupied".into(),
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
    let index =
        found.ok_or_else(|| Error::UnsafeEdit("authored Formula target is absent".into()))?;
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
        return Err(Error::UnsafeEdit(
            "authored Formula inverse state is stale".into(),
        ));
    }
    if row
        .cell_records
        .get(index.saturating_add(1))
        .is_some_and(|next| record_kind(next).is_ok_and(|kind| matches!(kind, STRING | CONTINUE)))
    {
        return Err(Error::UnsafeEdit(
            "authored Formula inverse refuses a string-cache owner group".into(),
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
        Error::UnsafeEdit("packed deletion target precedes its MulRk range".into())
    })?);
    if item >= count {
        return Err(Error::UnsafeEdit(
            "packed deletion target exceeds its MulRk range".into(),
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
        Error::UnsafeEdit("packed blank deletion target precedes its MulBlank range".into())
    })?);
    if item >= count {
        return Err(Error::UnsafeEdit(
            "packed blank deletion target exceeds its MulBlank range".into(),
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
        Error::UnsafeEdit("MulRk insertion value is not exactly RK-representable".into())
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
            return Err(Error::UnsafeEdit(
                "MulRk inverse adjacent records belong to another row".into(),
            ));
        }
        let left_last =
            left_first
                .checked_add(u16::try_from(left_items.len().saturating_sub(1)).map_err(
                    |_error| Error::InvalidData("packed inverse span exceeds u16".into()),
                )?)
                .ok_or_else(|| Error::InvalidData("packed inverse span overflows".into()))?;
        if left_last.checked_add(1) != Some(target) || target.checked_add(1) != Some(right_first) {
            return Err(Error::UnsafeEdit(
                "MulRk inverse adjacent records do not bracket the restored cell".into(),
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
        return Err(Error::UnsafeEdit(
            "MulRk inverse has ambiguous adjacent packed records".into(),
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
        return Err(Error::UnsafeEdit(
            "MulRk inverse target is not an edge extension".into(),
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
        return Err(Error::UnsafeEdit(
            "Formula record replacement requires a formula cache".into(),
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
    let count = u16::try_from(units.len())
        .map_err(|_error| Error::UnsafeEdit("formula string cache exceeds u16".into()))?;
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let byte_len = units
        .len()
        .checked_mul(if compressed { 1 } else { 2 })
        .and_then(|length| length.checked_add(3))
        .ok_or_else(|| Error::InvalidData("formula String payload length overflow".into()))?;
    if byte_len > 8_224 {
        return Err(Error::UnsafeEdit(
            "formula string cache requires unsupported Continue splitting".into(),
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
                    Error::UnsafeEdit("row insertion moves content outside BIFF8".into())
                })?
            }
        } else if row < start {
            row
        } else if row < delete_end {
            continue;
        } else {
            row - count
        };
        patch_row_number(&mut data, target)?;
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
        let mut retained = Vec::new();
        for mut record in std::mem::take(&mut data.cell_records) {
            let kind = record_kind(&record)?;
            if matches!(kind, TABLE | SHARED_FORMULA | ARRAY) {
                return Err(Error::UnsafeEdit(
                    "column movement of formula ownership records is refused".into(),
                ));
            }
            if !is_cell_record(kind) {
                retained.push(record);
                continue;
            }
            if kind == MUL_RK || kind == MUL_BLANK {
                return Err(Error::UnsafeEdit(
                    "column movement across packed cell records is refused".into(),
                ));
            }
            let column = cell_column(&record)?;
            let target = if insert {
                if column < start {
                    column
                } else {
                    column.checked_add(count).ok_or_else(|| {
                        Error::UnsafeEdit("column insertion moves content outside BIFF8".into())
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
            retained.push(record);
        }
        data.cell_records = retained;
    }
    Ok(())
}

fn patch_row_number(data: &mut RowData, row: u16) -> Result<()> {
    data.row_record
        .get_mut(4..6)
        .ok_or_else(|| Error::InvalidData("ROW record is truncated".into()))?
        .copy_from_slice(&row.to_le_bytes());
    for record in &mut data.cell_records {
        let kind = record_kind(record)?;
        if matches!(kind, TABLE | SHARED_FORMULA | ARRAY) {
            return Err(Error::UnsafeEdit(
                "row movement of formula ownership records is refused".into(),
            ));
        }
        if !is_cell_record(kind) {
            continue;
        }
        record
            .get_mut(4..6)
            .ok_or_else(|| Error::InvalidData("cell record is truncated".into()))?
            .copy_from_slice(&row.to_le_bytes());
    }
    Ok(())
}

fn patch_row_extents(rows: &mut BTreeMap<u16, RowData>) -> Result<()> {
    for data in rows.values_mut() {
        let mut first = u16::MAX;
        let mut last = 0_u16;
        for record in &data.cell_records {
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
        let (first, last_plus_one) = if first == u16::MAX {
            (0, 0)
        } else {
            (first, last + 1)
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
        return Err(Error::UnsafeEdit(
            "inserted cell XF resource is stale".into(),
        ));
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
                .ok_or_else(|| Error::UnsafeEdit("inserted text is absent from SST".into()))?;
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
            return Err(Error::UnsafeEdit(
                "inserted cell value has no standalone BIFF8 encoding".into(),
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
        return Err(Error::UnsafeEdit("structural cell state is stale".into()));
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
        return Err(Error::UnsafeEdit(
            "structural cell SST dependency is stale".into(),
        ));
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
                    Error::UnsafeEdit("MulRk style target precedes packed range".into())
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
                    Error::UnsafeEdit("MulRk style target precedes packed range".into())
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
    const DEPENDENT: &[u16] = &[
        TABLE,
        SHARED_FORMULA,
        ARRAY,
        0x01be,
        0x01b2,
        0x01b0,
        0x0879,
        0x087a,
        0x001c,
        0x007d,
        0x009e,
        0x0041,
        0x001a,
        0x001b,
        0x0090,
    ];
    if let Some(record) = records
        .iter()
        .find(|record| DEPENDENT.contains(&record.kind))
    {
        return Err(Error::UnsafeEdit(format!(
            "row/column movement cannot close dependency record 0x{:04X}",
            record.kind
        )));
    }
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
            SELECTION => certify_mutation(payload, shift, patch_selection_shift)?,
            MERGED_CELLS => certify_mutation(payload, shift, patch_merged_cells)?,
            HLINK => certify_mutation(payload, shift, patch_hyperlink)?,
            HLINK_TOOLTIP => certify_mutation(payload, shift, patch_hyperlink_tooltip)?,
            MSO_DRAWING => certify_mutation(payload, shift, patch_officeart)?,
            OBJ => certify_simple_drawing_obj(payload)?,
            _ => {},
        }
    }
    Ok(())
}

fn certify_hyperlink_target(payload: &[u8]) -> Result<()> {
    let hyperlink = crate::hyperlinks::parse_hlink_record(payload)?;
    if hyperlink.location().is_some() {
        return Err(Error::UnsafeEdit(
            "row/column movement refuses a hyperlink with an internal location dependency".into(),
        ));
    }
    Ok(())
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
        return Err(Error::UnsafeEdit(
            "row/column movement refuses Formula records with RgbExtra ancillary references".into(),
        ));
    }
    Ok(tokens)
}

fn certify_formula_owner(payload: &[u8]) -> Result<()> {
    if crate::formula::formula_2d_reference_spans(formula_tokens(payload)?).is_none() {
        return Err(Error::UnsafeEdit(
            "row/column movement supports only ordinary same-sheet Ref/Area Formula tokens".into(),
        ));
    }
    Ok(())
}

fn patch_formula_record(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let token_length = formula_tokens(payload)?.len();
    let tokens = payload
        .get_mut(22..22 + token_length)
        .ok_or_else(|| Error::InvalidData("Formula tokens are truncated".into()))?;
    patch_formula_tokens(tokens, shift)
}

fn patch_formula_tokens(tokens: &mut [u8], shift: AxisShift) -> Result<()> {
    let references = crate::formula::formula_2d_reference_spans(tokens).ok_or_else(|| {
        Error::UnsafeEdit(
            "row/column movement supports only ordinary same-sheet Ref/Area Formula tokens".into(),
        )
    })?;
    for reference in references {
        match (reference, shift) {
            (
                crate::formula::FormulaReferenceSpan::Cell { row, .. },
                AxisShift::Rows {
                    start,
                    count,
                    insert,
                },
            ) => patch_formula_row(tokens, row, start, count, insert)?,
            (
                crate::formula::FormulaReferenceSpan::Cell { column, .. },
                AxisShift::Columns {
                    start,
                    count,
                    insert,
                },
            ) => patch_formula_column(tokens, column, start, count, insert)?,
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
            ) => patch_formula_row_range(tokens, first_row, last_row, start, count, insert)?,
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
            ) => {
                patch_formula_column_range(
                    tokens,
                    first_column,
                    last_column,
                    start,
                    count,
                    insert,
                )?;
            },
        }
    }
    Ok(())
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
            (_, 0, 0) if end == payload.len() => return Ok(()),
            _ => {
                return Err(Error::UnsafeEdit(
                    "row/column movement only supports coordinate-free drawing Obj records".into(),
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
            .ok_or_else(|| Error::UnsafeEdit("BoundSheet position adjustment overflows".into()))?;
        workbook
            .get_mut(record_start + 4..record_start + 8)
            .ok_or_else(|| Error::InvalidData("BoundSheet position field is truncated".into()))?
            .copy_from_slice(&position.to_le_bytes());
    }
    for old_position in shifted_bounds {
        let position = i64::from(old_position)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| Error::UnsafeEdit("worksheet position adjustment overflows".into()))?;
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
        .ok_or_else(|| Error::UnsafeEdit("worksheet replacement requires a BoundSheet".into()))?;
    let mut prior_start = workbook.len();
    let mut deltas = Vec::new();
    deltas
        .try_reserve_exact(replacements.len())
        .map_err(|_error| Error::Allocation("retaining worksheet replacement deltas"))?;
    for (start, end, replacement) in replacements {
        if *start < first_sheet || *start > *end || *end > workbook.len() {
            return Err(Error::UnsafeEdit(
                "batched replacement is outside a worksheet source range".into(),
            ));
        }
        if *end > prior_start {
            return Err(Error::UnsafeEdit(
                "batched worksheet replacement ranges overlap or are not descending".into(),
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
            .ok_or_else(|| Error::UnsafeEdit("BoundSheet position adjustment overflows".into()))?;
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
                Error::UnsafeEdit("INDEX absolute offset adjustment overflows".into())
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
    let records = raw_records(workbook)?;
    let mut bounds = Vec::new();
    for record in records {
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

fn raw_records(bytes: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    for item in Records::new(bytes) {
        let record = item?;
        let start = record.offset();
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        records.push(RawRecord {
            kind: record.kind().get(),
            start,
            end,
        });
    }
    Ok(records)
}

fn unique_kind(records: &[RawRecord], kind: u16, label: &str) -> Result<usize> {
    let mut matches = records
        .iter()
        .enumerate()
        .filter(|(_, record)| record.kind == kind);
    let index = matches
        .next()
        .map(|(index, _)| index)
        .ok_or_else(|| Error::UnsafeEdit(format!("worksheet has no {label} record")))?;
    if matches.next().is_some() {
        return Err(Error::UnsafeEdit(format!(
            "worksheet has duplicate {label} records"
        )));
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

        let original_anchor = [3, 0, 2, 0, 0, 0, 1, 0, 0, 0, 4, 0, 0, 0, 5, 0, 0, 0];
        let mut anchor = original_anchor;
        patch_client_anchor(&mut anchor, insert_rows).unwrap();
        patch_client_anchor(&mut anchor, delete_rows).unwrap();
        assert_eq!(anchor, original_anchor);

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
        patch_client_anchor(&mut anchor, insert_columns).unwrap();
        patch_client_anchor(&mut anchor, delete_columns).unwrap();
        assert_eq!(anchor, original_anchor);

        let destructive = AxisShift::Rows {
            start: 4,
            count: 2,
            insert: false,
        };
        let mut destructive_range = original_range;
        assert!(patch_ref8(&mut destructive_range, destructive).is_err());
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
}

#[cfg(test)]
#[path = "structural_sst_tests.rs"]
mod structural_sst_tests;
