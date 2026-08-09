//! Workbook and CFB transaction layer for XLS OLE objects.

use super::super::codec::{ranges, u32_at};
use super::super::semantic::{FormControl, ObjectMetadataEdit, OleObjectRecord};
use super::super::{BOUNDSHEET, CFB_STREAM, CONTINUE, EOF, Limits, OBJ, TXO, invalid};
use crate::error::{Error, Result};
use litchi_cfb::OleFile;
use litchi_ole_common::object::{Editor as ObjectEditor, Target, Targets};
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

#[derive(Clone)]
pub struct Editor {
    package: ObjectEditor,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    sheets: Vec<Vec<OleObjectRecord>>,
    form_controls: Vec<Vec<FormControl>>,
    /// Number of form-control Obj records already present in each source
    /// worksheet. Newly authored controls are appended at the worksheet EOF;
    /// existing controls remain in their original byte representation.
    preserved_control_counts: Vec<usize>,
}

impl Editor {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        // Workbook metadata is XLS-owned. Read and parse it before handing
        // the original CFB bytes to the neutral object editor so the target
        // catalog can be derived solely from Obj/FtPictFmla records.
        let (workbook_path, workbook) = read_workbook(&bytes, limits)?;
        let (sheets, form_controls) = parse_workbook(&workbook)?;
        let targets = targets_for_sheets(&sheets)?;
        let package = ObjectEditor::open(bytes, targets, limits)?;
        let preserved_control_counts = form_controls.iter().map(Vec::len).collect();
        Ok(Self {
            package,
            workbook_path,
            workbook,
            sheets,
            form_controls,
            preserved_control_counts,
        })
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn objects(&self, worksheet: usize) -> Result<&[OleObjectRecord]> {
        self.sheets
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    /// Form controls (checkboxes, list boxes, scroll bars, ...) anchored in a
    /// worksheet, in Obj record order.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn form_controls(&self, worksheet: usize) -> Result<&[FormControl]> {
        self.form_controls
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub(crate) fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Adds a typed worksheet form-control Obj record transactionally.
    ///
    /// The operation authors only the BIFF `Obj` metadata. It does not load,
    /// instantiate, or execute any external control/runtime content. Unknown
    /// subrecords supplied by the caller are serialized unchanged, and all
    /// controls already present in the source worksheet remain byte-identical.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_form_control(&mut self, worksheet: usize, control: FormControl) -> Result<()> {
        control.validate()?;
        let object_id = control.object_id();
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.object_id() == object_id)
            || self
                .form_controls
                .iter()
                .flatten()
                .any(|value| value.object_id() == object_id)
        {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        let mut candidate = self.clone();
        candidate
            .form_controls
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .push(control);
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add(
        &mut self,
        worksheet: usize,
        object: OleObjectRecord,
        compound_file: Vec<u8>,
    ) -> Result<()> {
        object.validate()?;
        let storage = object
            .storage_name()
            .ok_or_else(|| invalid(OBJ, "new Obj has no MBD/LNK reference"))?;
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.object_id() == object.object_id())
            || self
                .form_controls
                .iter()
                .flatten()
                .any(|value| value.object_id() == object.object_id())
        {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        let mut candidate = self.clone();
        candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .push(object);
        let target = target_for_storage(storage)?;
        candidate.package.add_storage(target, compound_file)?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove(&mut self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let index = sheet
            .iter()
            .position(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        let removed = sheet.remove(index);
        if let Some(storage) = removed.storage_name()
            && !candidate
                .sheets
                .iter()
                .flatten()
                .any(|value| value.storage_name().as_deref() == Some(&storage))
        {
            let target = target_for_storage(storage)?;
            candidate.package.remove_storage(target.key())?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(removed)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn reorder(&mut self, worksheet: usize, ids: &[u16]) -> Result<()> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        if ids.len() != sheet.len() {
            return Err(invalid(
                OBJ,
                "reorder must contain every worksheet OLE object",
            ));
        }
        let mut remaining = sheet.clone();
        let mut reordered = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining
                .iter()
                .position(|value| value.object_id() == *id)
                .ok_or_else(|| invalid(OBJ, "unknown or repeated OLE object ID"))?;
            reordered.push(remaining.remove(index));
        }
        *sheet = reordered;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn update_object_metadata(
        &mut self,
        worksheet: usize,
        object_id: u16,
        edit: ObjectMetadataEdit,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let object = sheet
            .iter_mut()
            .find(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        edit.apply(object)?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> Result<()> {
        let storage = self
            .sheets
            .iter()
            .flatten()
            .find_map(|value| {
                value
                    .storage_name()
                    .filter(|value| value.as_str() == storage_name)
            })
            .ok_or_else(|| invalid(OBJ, "storage has no Obj reference"))?;
        let target = target_for_storage(storage)?;
        self.package
            .replace(target.key(), compound_file)
            .map_err(Into::into)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    fn commit(&mut self) -> Result<()> {
        validate_entries(&self.sheets, &self.form_controls)?;
        let workbook = rewrite_workbook(
            &self.workbook,
            &self.sheets,
            &self.form_controls,
            &self.preserved_control_counts,
        )?;
        self.package
            .put_stream(&self.workbook_path, workbook.clone())?;
        self.workbook = workbook;
        self.preserved_control_counts = self.form_controls.iter().map(Vec::len).collect();
        Ok(())
    }
}

pub(crate) fn read_workbook(bytes: &[u8], limits: Limits) -> Result<(Vec<String>, Vec<u8>)> {
    let max_size = limits.max_stream_size.min(limits.max_total_size);
    if max_size == 0 {
        return Err(Error::InvalidData(
            "Workbook stream limits must be non-zero".into(),
        ));
    }
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    let entries = ole.list_directory_entries(&[])?;
    for name in ["Workbook", "Book"] {
        let Some((actual_name, declared_size)) = entries
            .iter()
            .find(|entry| entry.entry_type == CFB_STREAM && entry.name.eq_ignore_ascii_case(name))
            .map(|entry| (entry.name.clone(), entry.size))
        else {
            continue;
        };
        if declared_size > max_size {
            return Err(Error::InvalidData(format!(
                "{actual_name} stream exceeds configured read limit"
            )));
        }
        let workbook = ole.open_stream(&[actual_name.as_str()])?;
        if workbook.len() as u64 > max_size {
            return Err(Error::InvalidData(format!(
                "{actual_name} stream exceeds configured read limit"
            )));
        }
        return Ok((vec![actual_name], workbook));
    }
    Err(Error::InvalidData("Workbook stream not found".into()))
}

fn target_for_storage(storage: String) -> Result<Target> {
    Ok(Target::new(storage.clone(), [storage])?)
}

pub(crate) fn targets_for_sheets(sheets: &[Vec<OleObjectRecord>]) -> Result<Targets> {
    let mut seen = HashSet::new();
    let mut targets = Vec::new();
    for object in sheets.iter().flatten() {
        let Some(storage) = object.storage_name() else {
            continue;
        };
        if seen.insert(storage.clone()) {
            targets.push(target_for_storage(storage)?);
        }
    }
    Ok(Targets::new(targets)?)
}

#[allow(
    clippy::type_complexity,
    reason = "type mirrors the decoded BIFF record structure"
)]
fn parse_workbook(input: &[u8]) -> Result<(Vec<Vec<OleObjectRecord>>, Vec<Vec<FormControl>>)> {
    let (_, starts) = bindings(input)?;
    let mut sheets = Vec::new();
    let mut form_controls = Vec::new();
    for (index, (start, worksheet)) in starts.iter().enumerate() {
        if !worksheet {
            continue;
        }
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        let (objects, controls) = parse_sheet(&input[*start..end])?;
        sheets.push(objects);
        form_controls.push(controls);
    }
    validate_entries(&sheets, &form_controls)?;
    Ok((sheets, form_controls))
}

fn parse_sheet(input: &[u8]) -> Result<(Vec<OleObjectRecord>, Vec<FormControl>)> {
    let records = ranges(input)?;
    let mut objects = Vec::new();
    let mut controls = Vec::new();
    for (index, value) in records.iter().enumerate() {
        if value.2 != OBJ {
            continue;
        }
        let txo = if records.get(index + 1).is_some_and(|next| next.2 == TXO) {
            if records
                .get(index + 2)
                .is_some_and(|next| next.2 == CONTINUE)
            {
                return Err(invalid(
                    TXO,
                    "Continue-based TxO beside OLE Obj is unsupported",
                ));
            }
            let next = records[index + 1];
            Some(input[next.0..next.1].to_vec())
        } else {
            None
        };
        let body = &input[value.3..value.4];
        if let Ok(object) = OleObjectRecord::parse(body, txo.clone()) {
            objects.push(object);
        } else if let Some(control) = FormControl::parse(body, txo) {
            controls.push(control);
        }
    }
    Ok((objects, controls))
}

fn validate_entries(
    sheets: &[Vec<OleObjectRecord>],
    form_controls: &[Vec<FormControl>],
) -> Result<()> {
    let mut ids = HashSet::new();
    let mut count = 0usize;
    for object in sheets.iter().flatten() {
        count += 1;
        if count > 4_096 {
            return Err(invalid(OBJ, "workbook object count exceeds limit"));
        }
        object.validate()?;
        if !ids.insert(object.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
    }
    for control in form_controls.iter().flatten() {
        count += 1;
        if count > 4_096 {
            return Err(invalid(OBJ, "workbook object count exceeds limit"));
        }
        if !ids.insert(control.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
    }
    Ok(())
}

fn rewrite_workbook(
    input: &[u8],
    sheets: &[Vec<OleObjectRecord>],
    form_controls: &[Vec<FormControl>],
    preserved_control_counts: &[usize],
) -> Result<Vec<u8>> {
    let (refs, starts) = bindings(input)?;
    let first = starts.first().map_or(input.len(), |value| value.0);
    let mut output = input[..first].to_vec();
    let mut new_offsets = HashMap::new();
    let mut worksheet = 0usize;
    for (index, (start, is_worksheet)) in starts.iter().enumerate() {
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        new_offsets.insert(*start, output.len());
        if *is_worksheet {
            output.extend_from_slice(&rewrite_sheet(
                &input[*start..end],
                sheets
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "worksheet list missing"))?,
                form_controls
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "form-control list missing"))?,
                *preserved_control_counts
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "form-control baseline missing"))?,
            )?);
            worksheet += 1;
        } else {
            output.extend_from_slice(&input[*start..end]);
        }
    }
    if worksheet != sheets.len() {
        return Err(invalid(BOUNDSHEET, "worksheet count mismatch"));
    }
    for (payload, old) in refs {
        let new = *new_offsets
            .get(&old)
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet target missing"))?;
        output[payload..payload + 4].copy_from_slice(
            &u32::try_from(new)
                .map_err(|_error| invalid(BOUNDSHEET, "sheet offset exceeds u32"))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_sheet(
    input: &[u8],
    objects: &[OleObjectRecord],
    form_controls: &[FormControl],
    preserved_control_count: usize,
) -> Result<Vec<u8>> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    let mut next = 0usize;
    let mut skip_txo = false;
    for (index, value) in records.iter().enumerate() {
        if skip_txo && value.2 == TXO {
            skip_txo = false;
            continue;
        }
        if value.2 == OBJ && OleObjectRecord::parse(&input[value.3..value.4], None).is_ok() {
            if let Some(object) = objects.get(next) {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
                next += 1;
            }
            skip_txo = records
                .get(index + 1)
                .is_some_and(|following| following.2 == TXO);
            continue;
        }
        if value.2 == EOF {
            for object in &objects[next..] {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
            }
            next = objects.len();
            for control in form_controls
                .get(preserved_control_count..)
                .ok_or_else(|| {
                    invalid(OBJ, "form-control baseline exceeds current control count")
                })?
            {
                output.extend_from_slice(&control.to_record_bytes()?);
                if let Some(txo) = &control.text_object {
                    output.extend_from_slice(txo);
                }
            }
        }
        output.extend_from_slice(&input[value.0..value.1]);
    }
    if next != objects.len() {
        return Err(invalid(EOF, "worksheet has no EOF"));
    }
    Ok(output)
}

#[allow(
    clippy::type_complexity,
    reason = "type mirrors the decoded BIFF record structure"
)]
fn bindings(input: &[u8]) -> Result<(Vec<(usize, usize)>, Vec<(usize, bool)>)> {
    let mut refs = Vec::new();
    for (start, _, kind, body_start, body_end) in ranges(input)? {
        if kind != BOUNDSHEET {
            continue;
        }
        let body = &input[body_start..body_end];
        if body.len() < 6 {
            return Err(invalid(BOUNDSHEET, "BoundSheet is truncated"));
        }
        refs.push((
            start
                .checked_add(4)
                .ok_or_else(|| invalid(BOUNDSHEET, "record offset overflow"))?,
            u32_at(body, 0).ok_or_else(|| invalid(BOUNDSHEET, "sheet offset is truncated"))?
                as usize,
            body[5] == 0,
        ));
    }
    let mut starts = refs
        .iter()
        .map(|(_, offset, sheet)| (*offset, *sheet))
        .collect::<Vec<_>>();
    starts.sort_by_key(|value| value.0);
    if starts.windows(2).any(|value| value[0].0 >= value[1].0)
        || starts.iter().any(|value| value.0 >= input.len())
    {
        return Err(invalid(BOUNDSHEET, "invalid or duplicate sheet offsets"));
    }
    Ok((
        refs.into_iter()
            .map(|(payload, offset, _)| (payload, offset))
            .collect(),
        starts,
    ))
}
