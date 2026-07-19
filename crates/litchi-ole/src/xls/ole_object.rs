//! Strict BIFF8 Obj/FtPictFmla parsing and transactional OLE-object editing.

use super::{XlsError, XlsResult};
use crate::{LegacyOfficeObjectEditor, LegacyOfficeObjectFormat, LegacyOfficeObjectLimits};
use std::collections::{HashMap, HashSet};

const OBJ: u16 = 0x005D;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;
const BOUNDSHEET: u16 = 0x0085;
const EOF: u16 = 0x000A;
const FT_CMO: u16 = 0x0015;
const FT_CF: u16 = 0x0007;
const FT_PIO: u16 = 0x0008;
const FT_PICT_FMLA: u16 = 0x0009;
const FT_END: u16 = 0;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtCmo {
    pub object_type: u16,
    pub object_id: u16,
    pub flags: u16,
    pub reserved: [u8; 12],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XlsFtPioGrbit {
    pub raw: u16,
}

impl XlsFtPioGrbit {
    pub fn is_dde(self) -> bool { self.raw & 2 != 0 }
    pub fn display_as_icon(self) -> bool { self.raw & 8 != 0 }
    pub fn is_control(self) -> bool { self.raw & 0x10 != 0 }
    pub fn uses_control_stream(self) -> bool { self.raw & 0x20 != 0 }
    pub fn camera_picture(self) -> bool { self.raw & 0x80 != 0 }
    pub fn default_size(self) -> bool { self.raw & 0x100 != 0 }
    pub fn auto_load(self) -> bool { self.raw & 0x200 != 0 }
    fn validate(self) -> XlsResult<()> {
        if self.is_dde() && self.is_control() {
            return Err(invalid(OBJ, "FtPioGrbit DDE and control flags are mutually exclusive"));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFtPictFmla {
    pub formula: Vec<u8>,
    pub storage_position: Option<u32>,
    pub control_buffer_size: Option<u32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsObjSubrecord {
    Common(XlsFtCmo),
    ClipboardFormat(Vec<u8>),
    PictureFlags(XlsFtPioGrbit),
    PictureFormula(XlsFtPictFmla),
    Unknown { kind: u16, data: Vec<u8> },
    End,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsOleObjectRecord {
    pub subrecords: Vec<XlsObjSubrecord>,
    /// Complete adjacent TxO record, retained byte-for-byte.
    pub text_object: Option<Vec<u8>>,
}

impl XlsOleObjectRecord {
    pub fn parse(data: &[u8], text_object: Option<Vec<u8>>) -> XlsResult<Self> {
        let mut offset = 0usize;
        let mut subrecords = Vec::new();
        while offset < data.len() {
            let header = data.get(offset..offset + 4)
                .ok_or_else(|| invalid(OBJ, "truncated Obj subrecord header"))?;
            let kind = u16::from_le_bytes([header[0], header[1]]);
            let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
            offset += 4;
            let end = offset.checked_add(len).ok_or_else(|| invalid(OBJ, "Obj subrecord overflow"))?;
            let body = data.get(offset..end).ok_or_else(|| invalid(OBJ, "truncated Obj subrecord"))?;
            let value = match (kind, len) {
                (FT_CMO, 18) => XlsObjSubrecord::Common(XlsFtCmo {
                    object_type: u16::from_le_bytes([body[0], body[1]]),
                    object_id: u16::from_le_bytes([body[2], body[3]]),
                    flags: u16::from_le_bytes([body[4], body[5]]),
                    reserved: body[6..18].try_into().expect("length checked"),
                }),
                (FT_CMO, _) => return Err(invalid(OBJ, "FtCmo must contain 18 bytes")),
                (FT_CF, _) => XlsObjSubrecord::ClipboardFormat(body.to_vec()),
                (FT_PIO, 2) => XlsObjSubrecord::PictureFlags(XlsFtPioGrbit { raw: u16::from_le_bytes([body[0], body[1]]) }),
                (FT_PIO, _) => return Err(invalid(OBJ, "FtPioGrbit must contain 2 bytes")),
                (FT_PICT_FMLA, _) => XlsObjSubrecord::PictureFormula(parse_formula(body)?),
                (FT_END, 0) => XlsObjSubrecord::End,
                (FT_END, _) => return Err(invalid(OBJ, "FtEnd must be empty")),
                _ => XlsObjSubrecord::Unknown { kind, data: body.to_vec() },
            };
            subrecords.push(value);
            offset = end;
        }
        let value = Self { subrecords, text_object };
        value.validate()?;
        Ok(value)
    }

    pub fn object_id(&self) -> u16 {
        self.subrecords.iter().find_map(|value| match value {
            XlsObjSubrecord::Common(value) => Some(value.object_id), _ => None,
        }).unwrap_or(0)
    }

    pub fn storage_position(&self) -> Option<u32> {
        self.subrecords.iter().find_map(|value| match value {
            XlsObjSubrecord::PictureFormula(value) => value.storage_position, _ => None,
        })
    }

    pub fn storage_name(&self) -> Option<String> {
        let position = self.storage_position()?;
        let dde = self.subrecords.iter().find_map(|value| match value {
            XlsObjSubrecord::PictureFlags(value) => Some(value.is_dde()), _ => None,
        }).unwrap_or(false);
        Some(format!("{}{:08X}", if dde { "LNK" } else { "MBD" }, position))
    }

    pub fn validate(&self) -> XlsResult<()> {
        if self.subrecords.len() > 1_024 { return Err(invalid(OBJ, "too many Obj subrecords")); }
        let common = self.subrecords.iter().filter_map(|value| match value {
            XlsObjSubrecord::Common(value) => Some(value), _ => None,
        }).collect::<Vec<_>>();
        if common.len() != 1 || common[0].object_type != 8 || common[0].object_id == 0
            || !matches!(self.subrecords.first(), Some(XlsObjSubrecord::Common(_))) {
            return Err(invalid(OBJ, "OLE Obj requires a leading FtCmo type 8 with nonzero ID"));
        }
        let pio = self.subrecords.iter().filter_map(|value| match value {
            XlsObjSubrecord::PictureFlags(value) => Some(*value), _ => None,
        }).collect::<Vec<_>>();
        if pio.len() != 1 { return Err(invalid(OBJ, "OLE Obj requires one FtPioGrbit")); }
        pio[0].validate()?;
        if self.subrecords.iter().filter(|value| matches!(value, XlsObjSubrecord::PictureFormula(_))).count() > 1 {
            return Err(invalid(OBJ, "duplicate FtPictFmla"));
        }
        if !matches!(self.subrecords.last(), Some(XlsObjSubrecord::End)) {
            return Err(invalid(OBJ, "OLE Obj must end with FtEnd"));
        }
        Ok(())
    }

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        self.validate()?;
        let mut body = Vec::new();
        for value in &self.subrecords {
            let (kind, data) = serialize_subrecord(value)?;
            let len = u16::try_from(data.len()).map_err(|_| invalid(OBJ, "Obj subrecord exceeds u16"))?;
            body.extend_from_slice(&kind.to_le_bytes());
            body.extend_from_slice(&len.to_le_bytes());
            body.extend_from_slice(&data);
        }
        record(OBJ, &body)
    }
}

#[derive(Clone)]
pub struct XlsOleObjectEditor {
    package: LegacyOfficeObjectEditor,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    sheets: Vec<Vec<XlsOleObjectRecord>>,
}

impl XlsOleObjectEditor {
    pub fn new(bytes: Vec<u8>, limits: LegacyOfficeObjectLimits) -> XlsResult<Self> {
        let package = LegacyOfficeObjectEditor::open(&bytes, LegacyOfficeObjectFormat::Xls, limits)?;
        let workbook_path = [vec!["Workbook".into()], vec!["Book".into()]].into_iter()
            .find(|path| package.package_stream(path).is_some())
            .ok_or_else(|| XlsError::InvalidData("Workbook stream not found".into()))?;
        let workbook = package.package_stream(&workbook_path).expect("selected stream").to_vec();
        let sheets = parse_workbook(&workbook)?;
        Ok(Self { package, workbook_path, workbook, sheets })
    }

    pub fn objects(&self, worksheet: usize) -> XlsResult<&[XlsOleObjectRecord]> {
        self.sheets.get(worksheet).map(Vec::as_slice)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub fn add(&mut self, worksheet: usize, object: XlsOleObjectRecord, compound_file: Vec<u8>) -> XlsResult<()> {
        object.validate()?;
        let storage = object.storage_name().ok_or_else(|| invalid(OBJ, "new Obj has no MBD/LNK reference"))?;
        if self.sheets.iter().flatten().any(|value| value.object_id() == object.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        let mut candidate = self.clone();
        candidate.sheets.get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?.push(object);
        candidate.package.add_referenced_storage(&storage, compound_file)?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    pub fn remove(&mut self, worksheet: usize, object_id: u16) -> XlsResult<XlsOleObjectRecord> {
        let mut candidate = self.clone();
        let sheet = candidate.sheets.get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let index = sheet.iter().position(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        let removed = sheet.remove(index);
        if let Some(storage) = removed.storage_name() {
            if !candidate.sheets.iter().flatten().any(|value| value.storage_name().as_deref() == Some(&storage)) {
                candidate.package.remove_referenced_storage(&storage)?;
            }
        }
        candidate.commit()?;
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder(&mut self, worksheet: usize, ids: &[u16]) -> XlsResult<()> {
        let mut candidate = self.clone();
        let sheet = candidate.sheets.get_mut(worksheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        if ids.len() != sheet.len() { return Err(invalid(OBJ, "reorder must contain every worksheet OLE object")); }
        let mut remaining = sheet.clone();
        let mut reordered = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining.iter().position(|value| value.object_id() == *id)
                .ok_or_else(|| invalid(OBJ, "unknown or repeated OLE object ID"))?;
            reordered.push(remaining.remove(index));
        }
        *sheet = reordered;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> XlsResult<()> {
        if !self.sheets.iter().flatten().any(|value| value.storage_name().as_deref() == Some(storage_name)) {
            return Err(invalid(OBJ, "storage has no Obj reference"));
        }
        self.package.replace(storage_name, compound_file).map_err(Into::into)
    }

    pub fn finish(self) -> XlsResult<Vec<u8>> { self.package.finish().map_err(Into::into) }

    fn commit(&mut self) -> XlsResult<()> {
        validate_objects(&self.sheets)?;
        let workbook = rewrite_workbook(&self.workbook, &self.sheets)?;
        self.package.replace_package_stream(&self.workbook_path, workbook.clone())?;
        self.workbook = workbook;
        Ok(())
    }
}

fn parse_formula(body: &[u8]) -> XlsResult<XlsFtPictFmla> {
    if body.len() < 2 { return Err(invalid(OBJ, "FtPictFmla is truncated")); }
    let len = usize::from(u16::from_le_bytes([body[0], body[1]]));
    let end = 2usize.checked_add(len).ok_or_else(|| invalid(OBJ, "formula overflow"))?;
    let formula = body.get(2..end).ok_or_else(|| invalid(OBJ, "formula is truncated"))?.to_vec();
    let tail = &body[end..];
    let (storage_position, control_buffer_size) = match tail.len() {
        0 => (None, None),
        8 => (Some(u32::from_le_bytes(tail[..4].try_into().unwrap())), Some(u32::from_le_bytes(tail[4..].try_into().unwrap()))),
        _ => return Err(invalid(OBJ, "unsupported FtPictFmla trailing layout")),
    };
    Ok(XlsFtPictFmla { formula, storage_position, control_buffer_size })
}

fn serialize_subrecord(value: &XlsObjSubrecord) -> XlsResult<(u16, Vec<u8>)> {
    Ok(match value {
        XlsObjSubrecord::Common(value) => {
            let mut body = Vec::with_capacity(18);
            body.extend_from_slice(&value.object_type.to_le_bytes()); body.extend_from_slice(&value.object_id.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes()); body.extend_from_slice(&value.reserved);
            (FT_CMO, body)
        }
        XlsObjSubrecord::ClipboardFormat(data) => (FT_CF, data.clone()),
        XlsObjSubrecord::PictureFlags(value) => (FT_PIO, value.raw.to_le_bytes().to_vec()),
        XlsObjSubrecord::PictureFormula(value) => {
            let len = u16::try_from(value.formula.len()).map_err(|_| invalid(OBJ, "formula exceeds u16"))?;
            let mut body = len.to_le_bytes().to_vec(); body.extend_from_slice(&value.formula);
            match (value.storage_position, value.control_buffer_size) {
                (Some(position), Some(size)) => { body.extend_from_slice(&position.to_le_bytes()); body.extend_from_slice(&size.to_le_bytes()); }
                (None, None) => {}
                _ => return Err(invalid(OBJ, "FtPictFmla optional fields must occur together")),
            }
            (FT_PICT_FMLA, body)
        }
        XlsObjSubrecord::Unknown { kind, data } => (*kind, data.clone()),
        XlsObjSubrecord::End => (FT_END, Vec::new()),
    })
}

fn parse_workbook(input: &[u8]) -> XlsResult<Vec<Vec<XlsOleObjectRecord>>> {
    let (_, starts) = bindings(input)?;
    let mut sheets = Vec::new();
    for (index, (start, worksheet)) in starts.iter().enumerate() {
        if !worksheet { continue; }
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        sheets.push(parse_sheet(&input[*start..end])?);
    }
    validate_objects(&sheets)?;
    Ok(sheets)
}

fn parse_sheet(input: &[u8]) -> XlsResult<Vec<XlsOleObjectRecord>> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    for (index, value) in records.iter().enumerate() {
        if value.2 != OBJ { continue; }
        let txo = if records.get(index + 1).is_some_and(|next| next.2 == TXO) {
            if records.get(index + 2).is_some_and(|next| next.2 == CONTINUE) {
                return Err(invalid(TXO, "Continue-based TxO beside OLE Obj is unsupported"));
            }
            let next = records[index + 1]; Some(input[next.0..next.1].to_vec())
        } else { None };
        if let Ok(object) = XlsOleObjectRecord::parse(&input[value.3..value.4], txo) { output.push(object); }
    }
    Ok(output)
}

fn validate_objects(sheets: &[Vec<XlsOleObjectRecord>]) -> XlsResult<()> {
    let mut ids = HashSet::new();
    for (index, object) in sheets.iter().flatten().enumerate() {
        if index >= 4_096 { return Err(invalid(OBJ, "workbook object count exceeds limit")); }
        object.validate()?;
        if !ids.insert(object.object_id()) { return Err(invalid(OBJ, "duplicate workbook object ID")); }
    }
    Ok(())
}

fn rewrite_workbook(input: &[u8], sheets: &[Vec<XlsOleObjectRecord>]) -> XlsResult<Vec<u8>> {
    let (refs, starts) = bindings(input)?;
    let first = starts.first().map_or(input.len(), |value| value.0);
    let mut output = input[..first].to_vec();
    let mut new_offsets = HashMap::new();
    let mut worksheet = 0usize;
    for (index, (start, is_worksheet)) in starts.iter().enumerate() {
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        new_offsets.insert(*start, output.len());
        if *is_worksheet {
            output.extend_from_slice(&rewrite_sheet(&input[*start..end], sheets.get(worksheet).ok_or_else(|| invalid(BOUNDSHEET, "worksheet list missing"))?)?);
            worksheet += 1;
        } else { output.extend_from_slice(&input[*start..end]); }
    }
    if worksheet != sheets.len() { return Err(invalid(BOUNDSHEET, "worksheet count mismatch")); }
    for (payload, old) in refs {
        let new = *new_offsets.get(&old).ok_or_else(|| invalid(BOUNDSHEET, "sheet target missing"))?;
        output[payload..payload + 4].copy_from_slice(&u32::try_from(new).map_err(|_| invalid(BOUNDSHEET, "sheet offset exceeds u32"))?.to_le_bytes());
    }
    Ok(output)
}

fn rewrite_sheet(input: &[u8], objects: &[XlsOleObjectRecord]) -> XlsResult<Vec<u8>> {
    let records = ranges(input)?;
    let mut output = Vec::new(); let mut next = 0usize; let mut skip_txo = false;
    for (index, value) in records.iter().enumerate() {
        if skip_txo && value.2 == TXO { skip_txo = false; continue; }
        if value.2 == OBJ && XlsOleObjectRecord::parse(&input[value.3..value.4], None).is_ok() {
            if let Some(object) = objects.get(next) {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object { output.extend_from_slice(txo); }
                next += 1;
            }
            skip_txo = records.get(index + 1).is_some_and(|following| following.2 == TXO);
            continue;
        }
        if value.2 == EOF {
            for object in &objects[next..] {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object { output.extend_from_slice(txo); }
            }
            next = objects.len();
        }
        output.extend_from_slice(&input[value.0..value.1]);
    }
    if next != objects.len() { return Err(invalid(EOF, "worksheet has no EOF")); }
    Ok(output)
}

fn bindings(input: &[u8]) -> XlsResult<(Vec<(usize, usize)>, Vec<(usize, bool)>)> {
    let mut refs = Vec::new();
    for (start, _, kind, body_start, body_end) in ranges(input)? {
        if kind != BOUNDSHEET { continue; }
        let body = &input[body_start..body_end];
        if body.len() < 6 { return Err(invalid(BOUNDSHEET, "BoundSheet is truncated")); }
        refs.push((start + 4, u32::from_le_bytes(body[..4].try_into().unwrap()) as usize, body[5] == 0));
    }
    let mut starts = refs.iter().map(|(_, offset, sheet)| (*offset, *sheet)).collect::<Vec<_>>();
    starts.sort_by_key(|value| value.0);
    if starts.windows(2).any(|value| value[0].0 >= value[1].0) || starts.iter().any(|value| value.0 >= input.len()) {
        return Err(invalid(BOUNDSHEET, "invalid or duplicate sheet offsets"));
    }
    Ok((refs.into_iter().map(|(payload, offset, _)| (payload, offset)).collect(), starts))
}

fn ranges(input: &[u8]) -> XlsResult<Vec<(usize, usize, u16, usize, usize)>> {
    let mut output = Vec::new(); let mut offset = 0usize;
    while offset < input.len() {
        let header = input.get(offset..offset + 4).ok_or(XlsError::InvalidLength { expected: offset + 4, found: input.len() })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset.checked_add(4 + len).ok_or_else(|| invalid(kind, "record size overflow"))?;
        if end > input.len() { return Err(XlsError::InvalidLength { expected: end, found: input.len() }); }
        output.push((offset, end, kind, offset + 4, end)); offset = end;
    }
    Ok(output)
}

fn record(kind: u16, body: &[u8]) -> XlsResult<Vec<u8>> {
    if body.len() > 8_224 { return Err(invalid(kind, "record exceeds BIFF8 limit")); }
    let mut output = Vec::with_capacity(body.len() + 4);
    output.extend_from_slice(&kind.to_le_bytes()); output.extend_from_slice(&(body.len() as u16).to_le_bytes()); output.extend_from_slice(body); Ok(output)
}

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord { record_type, message: message.into() }
}
