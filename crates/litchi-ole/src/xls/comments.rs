//! Strict BIFF8 cell comment/note retention.
//!
//! Comments are described by a later `NOTE` record and by an earlier drawing
//! sequence containing `OBJ`, `MSODRAWING`, `TXO`, and `CONTINUE` records. This
//! module links those records without rendering shapes or evaluating object
//! formulas.

use crate::xls::error::{XlsError, XlsResult};
use std::collections::{HashMap, HashSet};

pub const RECORD_TYPE: u16 = 0x001C;
pub const OBJ_TYPE: u16 = 0x005D;
pub const TXO_TYPE: u16 = 0x01B6;
pub const CONTINUE_TYPE: u16 = 0x003C;
const MSODRAWING_TYPE: u16 = 0x00EC;
const COMMENT_OBJECT_TYPE: u16 = 0x0019;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CommentVisibility {
    Hidden,
    Visible,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCommentHorizontalAlignment {
    Left,
    Centered,
    Right,
    Justified,
    Distributed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCommentVerticalAlignment {
    Top,
    Centered,
    Bottom,
    Justified,
    Distributed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCommentTextOrientation {
    None,
    Stacked,
    CounterClockwise,
    Clockwise,
}

/// Stable identity supplied by the comment's `OBJ`/`FtNts` structures.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsCommentObjectIdentity {
    object_id: u16,
    guid: [u8; 16],
    shared: bool,
}

impl XlsCommentObjectIdentity {
    pub fn object_id(&self) -> u16 {
        self.object_id
    }
    pub fn guid(&self) -> &[u8; 16] {
        &self.guid
    }
    pub fn shared(&self) -> bool {
        self.shared
    }
}

/// One formatting run from the comment's `TxORuns` structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCommentTextRun {
    character_index: u16,
    font_index: u16,
}

impl XlsCommentTextRun {
    pub fn character_index(&self) -> u16 {
        self.character_index
    }
    pub fn font_index(&self) -> u16 {
        self.font_index
    }
}

/// Inert text properties retained from the comment's `TXO` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsCommentTextProperties {
    horizontal_alignment: XlsCommentHorizontalAlignment,
    vertical_alignment: XlsCommentVerticalAlignment,
    orientation: XlsCommentTextOrientation,
    locked: bool,
    justify_last_line: bool,
    secret_edit: bool,
    font_when_empty: u16,
    formula_bytes: Vec<u8>,
}

impl XlsCommentTextProperties {
    pub fn horizontal_alignment(&self) -> XlsCommentHorizontalAlignment {
        self.horizontal_alignment
    }
    pub fn vertical_alignment(&self) -> XlsCommentVerticalAlignment {
        self.vertical_alignment
    }
    pub fn orientation(&self) -> XlsCommentTextOrientation {
        self.orientation
    }
    pub fn locked(&self) -> bool {
        self.locked
    }
    pub fn justify_last_line(&self) -> bool {
        self.justify_last_line
    }
    pub fn secret_edit(&self) -> bool {
        self.secret_edit
    }
    pub fn font_when_empty(&self) -> u16 {
        self.font_when_empty
    }
    /// Raw, unevaluated ObjFmla payload bytes, excluding its length field.
    pub fn formula_bytes(&self) -> &[u8] {
        &self.formula_bytes
    }
}

/// A fully linked, immutable BIFF8 cell comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsComment {
    row: u16,
    column: u8,
    visibility: CommentVisibility,
    row_hidden: bool,
    column_hidden: bool,
    identity: XlsCommentObjectIdentity,
    author: String,
    text: String,
    text_properties: XlsCommentTextProperties,
    text_runs: Vec<XlsCommentTextRun>,
}

impl XlsComment {
    pub fn row(&self) -> u16 {
        self.row
    }
    pub fn column(&self) -> u8 {
        self.column
    }
    pub fn visibility(&self) -> CommentVisibility {
        self.visibility
    }
    pub fn row_hidden(&self) -> bool {
        self.row_hidden
    }
    pub fn column_hidden(&self) -> bool {
        self.column_hidden
    }
    pub fn identity(&self) -> &XlsCommentObjectIdentity {
        &self.identity
    }
    pub fn author(&self) -> &str {
        &self.author
    }
    pub fn text(&self) -> &str {
        &self.text
    }
    pub fn text_properties(&self) -> &XlsCommentTextProperties {
        &self.text_properties
    }
    pub fn text_runs(&self) -> &[XlsCommentTextRun] {
        &self.text_runs
    }
}

#[derive(Debug)]
struct NoteRecord {
    row: u16,
    column: u8,
    visibility: CommentVisibility,
    row_hidden: bool,
    column_hidden: bool,
    object_id: u16,
    author: String,
}

#[derive(Debug)]
struct CommentObject {
    identity: XlsCommentObjectIdentity,
    text: Option<String>,
    text_properties: Option<XlsCommentTextProperties>,
    text_runs: Vec<XlsCommentTextRun>,
}

#[derive(Debug)]
struct PendingTxo {
    object_id: u16,
    character_count: usize,
    run_byte_count: usize,
    code_units: Vec<u16>,
    run_bytes: Vec<u8>,
    properties: XlsCommentTextProperties,
}

/// Worksheet-scoped comment linkage state.
#[derive(Debug, Default)]
pub(crate) struct CommentCollector {
    object_ids: HashSet<u16>,
    comment_guids: HashSet<[u8; 16]>,
    objects: HashMap<u16, CommentObject>,
    notes: Vec<NoteRecord>,
    note_cells: HashSet<(u16, u8)>,
    note_object_ids: HashSet<u16>,
    awaiting_drawing: Option<u16>,
    awaiting_txo: Option<u16>,
    pending_txo: Option<PendingTxo>,
}

impl CommentCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if let Some(mut pending) = self.pending_txo.take() {
            if record_type != CONTINUE_TYPE {
                return invalid(format!(
                    "incomplete TXO for comment object {} must be followed by CONTINUE",
                    pending.object_id
                ));
            }
            let complete = feed_txo_continue(&mut pending, data)?;
            if complete {
                self.complete_txo(pending)?;
            } else {
                self.pending_txo = Some(pending);
            }
            return Ok(());
        }

        if let Some(object_id) = self.awaiting_drawing.take() {
            if record_type != MSODRAWING_TYPE {
                return invalid(format!(
                    "comment OBJ {object_id} must be followed by its MSODRAWING textbox"
                ));
            }
            if data.len() != 8
                || u16_at(data, 0)? != 0
                || u16_at(data, 2)? != 0xF00D
                || u32_at(data, 4)? != 0
            {
                return invalid(format!(
                    "comment object {object_id} must be followed by an exact OfficeArtClientTextbox boundary"
                ));
            }
            self.awaiting_txo = Some(object_id);
            return Ok(());
        }

        if let Some(object_id) = self.awaiting_txo {
            if record_type != TXO_TYPE {
                return invalid(format!(
                    "comment object {object_id} textbox must be followed by TXO"
                ));
            }
            self.awaiting_txo = None;
            let pending = parse_txo(data, object_id)?;
            if pending.character_count == 0 {
                self.complete_txo(pending)?;
            } else {
                self.pending_txo = Some(pending);
            }
            return Ok(());
        }

        match record_type {
            OBJ_TYPE => {
                if let Some(object) = parse_obj(data)? {
                    let object_id = object.identity.object_id;
                    if !self.comment_guids.insert(object.identity.guid) {
                        return invalid("duplicate comment FtNts GUID".to_string());
                    }
                    self.objects.insert(object_id, object);
                    self.awaiting_drawing = Some(object_id);
                }
                let object_id = parse_cmo_id(data)?;
                if !self.object_ids.insert(object_id) {
                    return invalid(format!("duplicate OBJ object id: {object_id}"));
                }
            },
            RECORD_TYPE => {
                let note = parse_note_record(data)?;
                if !self.note_cells.insert((note.row, note.column)) {
                    return invalid(format!(
                        "duplicate NOTE for cell ({}, {})",
                        note.row, note.column
                    ));
                }
                if !self.note_object_ids.insert(note.object_id) {
                    return invalid(format!(
                        "duplicate NOTE object reference: {}",
                        note.object_id
                    ));
                }
                self.notes.push(note);
            },
            _ => {},
        }
        Ok(())
    }

    fn complete_txo(&mut self, pending: PendingTxo) -> XlsResult<()> {
        let text = String::from_utf16(&pending.code_units)
            .map_err(|_| XlsError::InvalidData("TXO text contains invalid UTF-16".to_string()))?;
        let runs = parse_txo_runs(&pending.run_bytes, pending.character_count as u16)?;
        let object = self.objects.get_mut(&pending.object_id).ok_or_else(|| {
            XlsError::InvalidData(format!(
                "TXO references unknown comment object {}",
                pending.object_id
            ))
        })?;
        if object.text.is_some() {
            return invalid(format!(
                "comment object {} has more than one TXO",
                pending.object_id
            ));
        }
        object.text = Some(text);
        object.text_properties = Some(pending.properties);
        object.text_runs = runs;
        Ok(())
    }

    pub(crate) fn finish(self) -> XlsResult<Vec<XlsComment>> {
        if self.awaiting_drawing.is_some()
            || self.awaiting_txo.is_some()
            || self.pending_txo.is_some()
        {
            return invalid(
                "worksheet ended with an incomplete comment object sequence".to_string(),
            );
        }
        for object_id in self.objects.keys() {
            if !self.note_object_ids.contains(object_id) {
                return invalid(format!("comment OBJ {object_id} has no matching NOTE"));
            }
        }
        let mut comments = Vec::with_capacity(self.notes.len());
        for note in self.notes {
            let object = self.objects.get(&note.object_id).ok_or_else(|| {
                XlsError::InvalidData(format!(
                    "NOTE references missing comment OBJ {}",
                    note.object_id
                ))
            })?;
            comments.push(XlsComment {
                row: note.row,
                column: note.column,
                visibility: note.visibility,
                row_hidden: note.row_hidden,
                column_hidden: note.column_hidden,
                identity: object.identity.clone(),
                author: note.author,
                text: object.text.clone().ok_or_else(|| {
                    XlsError::InvalidData(format!("comment OBJ {} has no TXO", note.object_id))
                })?,
                text_properties: object.text_properties.clone().ok_or_else(|| {
                    XlsError::InvalidData(format!(
                        "comment OBJ {} has no TXO properties",
                        note.object_id
                    ))
                })?,
                text_runs: object.text_runs.clone(),
            });
        }
        Ok(comments)
    }
}

fn parse_note_record(data: &[u8]) -> XlsResult<NoteRecord> {
    if data.len() < 13 {
        return invalid(format!("NOTE payload is too short: {}", data.len()));
    }
    let row = u16_at(data, 0)?;
    let column = u16_at(data, 2)?;
    if column > 255 {
        return invalid(format!("NOTE column exceeds BIFF8 limit: {column}"));
    }
    let flags = u16_at(data, 4)?;
    if flags & !0x018A != 0 {
        return invalid(format!("NOTE contains reserved flag bits: {flags:#06x}"));
    }
    let object_id = u16_at(data, 6)?;
    if object_id == 0 {
        return invalid("NOTE object id must not be zero".to_string());
    }
    let character_count = u16_at(data, 8)? as usize;
    if !(1..=54).contains(&character_count) {
        return invalid(format!(
            "NOTE author length must be 1..=54, got {character_count}"
        ));
    }
    let string_flags = data[10];
    if string_flags & !1 != 0 {
        return invalid(format!(
            "NOTE author contains reserved string flags: {string_flags:#04x}"
        ));
    }
    let width = if string_flags & 1 == 0 { 1 } else { 2 };
    let byte_count = character_count
        .checked_mul(width)
        .ok_or_else(|| XlsError::InvalidData("NOTE author size overflow".to_string()))?;
    let expected = 12usize
        .checked_add(byte_count)
        .ok_or_else(|| XlsError::InvalidData("NOTE size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(format!(
            "NOTE payload length must be {expected}, got {}",
            data.len()
        ));
    }
    let author = decode_unicode(&data[11..11 + byte_count], width == 2)?;
    Ok(NoteRecord {
        row,
        column: column as u8,
        visibility: if flags & 0x0002 != 0 {
            CommentVisibility::Visible
        } else {
            CommentVisibility::Hidden
        },
        row_hidden: flags & 0x0080 != 0,
        column_hidden: flags & 0x0100 != 0,
        object_id,
        author,
    })
}

fn parse_cmo_id(data: &[u8]) -> XlsResult<u16> {
    if data.len() < 22 {
        return invalid(format!(
            "OBJ payload is too short for FtCmo: {}",
            data.len()
        ));
    }
    if u16_at(data, 0)? != 0x0015 || u16_at(data, 2)? != 0x0012 {
        return invalid("OBJ must begin with a 22-byte FtCmo".to_string());
    }
    if u16_at(data, 8)? & 0x0002 != 0 {
        return invalid("FtCmo contains its reserved option bit".to_string());
    }
    Ok(u16_at(data, 6)?)
}

fn parse_obj(data: &[u8]) -> XlsResult<Option<CommentObject>> {
    let object_id = parse_cmo_id(data)?;
    if u16_at(data, 4)? != COMMENT_OBJECT_TYPE {
        return Ok(None);
    }
    if object_id == 0 {
        return invalid("comment OBJ id must not be zero".to_string());
    }
    let mut position = 22usize;
    let mut identity = None;
    let mut found_end = false;
    while position < data.len() {
        if data.len() - position < 4 {
            if data[position..].iter().all(|&byte| byte == 0) {
                break;
            }
            return invalid("truncated OBJ subrecord header".to_string());
        }
        let sub_type = u16_at(data, position)?;
        let size = u16_at(data, position + 2)? as usize;
        position += 4;
        if sub_type == 0 {
            if size != 0 {
                return invalid("FtEnd size must be zero".to_string());
            }
            found_end = true;
            if data[position..].iter().any(|&byte| byte != 0) {
                return invalid("non-zero OBJ padding after FtEnd".to_string());
            }
            break;
        }
        let end = position
            .checked_add(size)
            .ok_or_else(|| XlsError::InvalidData("OBJ subrecord size overflow".to_string()))?;
        let body = data
            .get(position..end)
            .ok_or_else(|| XlsError::InvalidData("truncated OBJ subrecord".to_string()))?;
        if sub_type == 0x000D {
            if size != 0x0016 {
                return invalid(format!("FtNts size must be 22, got {size}"));
            }
            if identity.is_some() {
                return invalid("comment OBJ contains more than one FtNts".to_string());
            }
            let guid: [u8; 16] = body[0..16].try_into().unwrap();
            let shared = match u16::from_le_bytes([body[16], body[17]]) {
                0 => false,
                1 => true,
                value => return invalid(format!("invalid FtNts shared-note value: {value}")),
            };
            identity = Some(XlsCommentObjectIdentity {
                object_id,
                guid,
                shared,
            });
        }
        position = end;
    }
    if !found_end {
        return invalid("comment OBJ is missing FtEnd".to_string());
    }
    Ok(Some(CommentObject {
        identity: identity
            .ok_or_else(|| XlsError::InvalidData("comment OBJ is missing FtNts".to_string()))?,
        text: None,
        text_properties: None,
        text_runs: Vec::new(),
    }))
}

fn parse_txo(data: &[u8], object_id: u16) -> XlsResult<PendingTxo> {
    if data.len() < 18 {
        return invalid(format!("TXO payload is too short: {}", data.len()));
    }
    let options = u16_at(data, 0)?;
    if options & !0xC27E != 0 {
        return invalid(format!("TXO contains reserved option bits: {options:#06x}"));
    }
    let horizontal_alignment = match (options >> 1) & 7 {
        1 => XlsCommentHorizontalAlignment::Left,
        2 => XlsCommentHorizontalAlignment::Centered,
        3 => XlsCommentHorizontalAlignment::Right,
        4 => XlsCommentHorizontalAlignment::Justified,
        7 => XlsCommentHorizontalAlignment::Distributed,
        value => return invalid(format!("invalid TXO horizontal alignment: {value}")),
    };
    let vertical_alignment = match (options >> 4) & 7 {
        1 => XlsCommentVerticalAlignment::Top,
        2 => XlsCommentVerticalAlignment::Centered,
        3 => XlsCommentVerticalAlignment::Bottom,
        4 => XlsCommentVerticalAlignment::Justified,
        7 => XlsCommentVerticalAlignment::Distributed,
        value => return invalid(format!("invalid TXO vertical alignment: {value}")),
    };
    let orientation = match u16_at(data, 2)? {
        0 => XlsCommentTextOrientation::None,
        1 => XlsCommentTextOrientation::Stacked,
        2 => XlsCommentTextOrientation::CounterClockwise,
        3 => XlsCommentTextOrientation::Clockwise,
        value => return invalid(format!("invalid TXO text orientation: {value}")),
    };
    if u16_at(data, 4)? != 0 || u32_at(data, 6)? != 0 {
        return invalid("TXO reserved fields must be zero".to_string());
    }
    let character_count = u16_at(data, 10)? as usize;
    let run_byte_count = u16_at(data, 12)? as usize;
    if character_count == 0 {
        if run_byte_count != 0 {
            return invalid("empty TXO text must have zero run bytes".to_string());
        }
    } else if run_byte_count < 16 || run_byte_count % 8 != 0 {
        return invalid(format!(
            "non-empty TXO run size must be a multiple of 8 and at least 16, got {run_byte_count}"
        ));
    }
    let formula_size = u16_at(data, 16)? as usize;
    if formula_size % 2 != 0 {
        return invalid(format!("TXO ObjFmla size must be even, got {formula_size}"));
    }
    let expected = 18usize
        .checked_add(formula_size)
        .ok_or_else(|| XlsError::InvalidData("TXO formula size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(format!(
            "TXO payload length must be {expected}, got {}",
            data.len()
        ));
    }
    Ok(PendingTxo {
        object_id,
        character_count,
        run_byte_count,
        code_units: Vec::with_capacity(character_count),
        run_bytes: Vec::with_capacity(run_byte_count),
        properties: XlsCommentTextProperties {
            horizontal_alignment,
            vertical_alignment,
            orientation,
            locked: options & 0x0200 != 0,
            justify_last_line: options & 0x4000 != 0,
            secret_edit: options & 0x8000 != 0,
            font_when_empty: u16_at(data, 14)?,
            formula_bytes: data[18..].to_vec(),
        },
    })
}

fn feed_txo_continue(pending: &mut PendingTxo, data: &[u8]) -> XlsResult<bool> {
    if pending.code_units.len() < pending.character_count {
        let (&flags, characters) = data
            .split_first()
            .ok_or_else(|| XlsError::InvalidData("empty TXO text CONTINUE".to_string()))?;
        if flags & !1 != 0 {
            return invalid(format!(
                "TXO text CONTINUE contains reserved flags: {flags:#04x}"
            ));
        }
        let wide = flags & 1 != 0;
        if wide && characters.len() % 2 != 0 {
            return invalid("TXO UTF-16 segment has an odd byte length".to_string());
        }
        let segment_count = characters.len() / if wide { 2 } else { 1 };
        let remaining = pending.character_count - pending.code_units.len();
        if segment_count == 0 || segment_count > remaining {
            return invalid("TXO text CONTINUE character count does not match cchText".to_string());
        }
        if wide {
            pending.code_units.extend(
                characters
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
            );
        } else {
            pending
                .code_units
                .extend(characters.iter().map(|&byte| u16::from(byte)));
        }
        return Ok(false);
    }

    let remaining = pending.run_byte_count - pending.run_bytes.len();
    if data.is_empty() || data.len() > remaining {
        return invalid("TXO run CONTINUE length exceeds cbRuns".to_string());
    }
    pending.run_bytes.extend_from_slice(data);
    Ok(pending.run_bytes.len() == pending.run_byte_count)
}

fn parse_txo_runs(data: &[u8], character_count: u16) -> XlsResult<Vec<XlsCommentTextRun>> {
    if data.is_empty() {
        return Ok(Vec::new());
    }
    if data.len() < 16 || data.len() % 8 != 0 {
        return invalid("invalid TxORuns byte count".to_string());
    }
    let run_count = data.len() / 8 - 1;
    let mut runs = Vec::with_capacity(run_count);
    for index in 0..run_count {
        let offset = index * 8;
        let character_index = u16_at(data, offset)?;
        if character_index > character_count {
            return invalid(format!(
                "TXO formatting run index {character_index} exceeds cchText {character_count}"
            ));
        }
        runs.push(XlsCommentTextRun {
            character_index,
            font_index: u16_at(data, offset + 2)?,
        });
    }
    let last_offset = run_count * 8;
    if u16_at(data, last_offset)? != character_count {
        return invalid("TxOLastRun character count does not match TXO cchText".to_string());
    }
    Ok(runs)
}

fn decode_unicode(data: &[u8], wide: bool) -> XlsResult<String> {
    if wide {
        let words: Vec<u16> = data
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect();
        String::from_utf16(&words)
            .map_err(|_| XlsError::InvalidData("NOTE author contains invalid UTF-16".to_string()))
    } else {
        Ok(data.iter().map(|&byte| char::from(byte)).collect())
    }
}

fn u16_at(data: &[u8], offset: usize) -> XlsResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| XlsError::InvalidData("truncated comment record".to_string()))?;
    Ok(u16::from_le_bytes(bytes.try_into().unwrap()))
}

fn u32_at(data: &[u8], offset: usize) -> XlsResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| XlsError::InvalidData("truncated comment record".to_string()))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn invalid<T>(message: String) -> XlsResult<T> {
    Err(XlsError::InvalidData(message))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn obj(object_id: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0015u16.to_le_bytes());
        data.extend_from_slice(&0x0012u16.to_le_bytes());
        data.extend_from_slice(&COMMENT_OBJECT_TYPE.to_le_bytes());
        data.extend_from_slice(&object_id.to_le_bytes());
        data.extend_from_slice(&0x4011u16.to_le_bytes());
        data.extend_from_slice(&[0; 12]);
        data.extend_from_slice(&0x000Du16.to_le_bytes());
        data.extend_from_slice(&0x0016u16.to_le_bytes());
        let mut guid = [0u8; 16];
        guid[0..2].copy_from_slice(&object_id.to_le_bytes());
        data.extend_from_slice(&guid);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(&[0; 4]);
        data
    }

    fn txo(character_count: u16, run_bytes: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0212u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[0; 6]);
        data.extend_from_slice(&character_count.to_le_bytes());
        data.extend_from_slice(&run_bytes.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data
    }

    fn client_textbox() -> [u8; 8] {
        [0, 0, 0x0D, 0xF0, 0, 0, 0, 0]
    }

    fn note(object_id: u16, flags: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&5u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&object_id.to_le_bytes());
        data.extend_from_slice(&4u16.to_le_bytes());
        data.push(0);
        data.extend_from_slice(b"User");
        data.push(0xD0);
        data
    }

    fn runs(character_count: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&6u16.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(&character_count.to_le_bytes());
        data.extend_from_slice(&[0; 6]);
        data
    }

    #[test]
    fn links_obj_txo_continues_and_note() {
        let mut collector = CommentCollector::new();
        collector.feed_record(OBJ_TYPE, &obj(7)).unwrap();
        collector
            .feed_record(MSODRAWING_TYPE, &client_textbox())
            .unwrap();
        collector.feed_record(TXO_TYPE, &txo(5, 16)).unwrap();
        collector.feed_record(CONTINUE_TYPE, b"\0Hello").unwrap();
        collector.feed_record(CONTINUE_TYPE, &runs(5)).unwrap();
        collector
            .feed_record(RECORD_TYPE, &note(7, 0x0182))
            .unwrap();
        let comments = collector.finish().unwrap();
        let comment = &comments[0];
        assert_eq!(comment.row(), 5);
        assert_eq!(comment.column(), 3);
        assert_eq!(comment.visibility(), CommentVisibility::Visible);
        assert!(comment.row_hidden() && comment.column_hidden());
        assert_eq!(comment.identity().object_id(), 7);
        assert_eq!(comment.author(), "User");
        assert_eq!(comment.text(), "Hello");
        assert_eq!(comment.text_runs()[0].font_index(), 6);
    }

    #[test]
    fn assembles_mixed_segmented_unicode_without_splitting_surrogates() {
        let mut collector = CommentCollector::new();
        collector.feed_record(OBJ_TYPE, &obj(8)).unwrap();
        collector
            .feed_record(MSODRAWING_TYPE, &client_textbox())
            .unwrap();
        collector.feed_record(TXO_TYPE, &txo(3, 16)).unwrap();
        collector.feed_record(CONTINUE_TYPE, b"\0A").unwrap();
        let mut wide = vec![1];
        wide.extend_from_slice(&0xD83Du16.to_le_bytes());
        wide.extend_from_slice(&0xDE00u16.to_le_bytes());
        collector.feed_record(CONTINUE_TYPE, &wide).unwrap();
        collector.feed_record(CONTINUE_TYPE, &runs(3)).unwrap();
        collector.feed_record(RECORD_TYPE, &note(8, 0)).unwrap();
        assert_eq!(collector.finish().unwrap()[0].text(), "A😀");
    }

    #[test]
    fn rejects_reserved_bits_broken_order_and_bad_last_run() {
        let mut malformed_note = note(1, 1);
        assert!(parse_note_record(&malformed_note).is_err());
        malformed_note[4..6].copy_from_slice(&0u16.to_le_bytes());
        malformed_note.pop();
        assert!(parse_note_record(&malformed_note).is_err());

        let mut collector = CommentCollector::new();
        collector.feed_record(OBJ_TYPE, &obj(1)).unwrap();
        assert!(collector.feed_record(TXO_TYPE, &txo(1, 16)).is_err());

        let mut collector = CommentCollector::new();
        collector.feed_record(OBJ_TYPE, &obj(2)).unwrap();
        assert!(collector.feed_record(MSODRAWING_TYPE, &[0; 8]).is_err());

        let mut bad_runs = runs(2);
        bad_runs[8..10].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_txo_runs(&bad_runs, 2).is_err());
    }

    #[test]
    fn reads_poi_comment_fixtures() {
        use crate::xls::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../3rdparty/poi/test-data/spreadsheet")
                .join(name)
        };

        let simple =
            XlsWorkbook::new(File::open(fixture("SimpleWithComments.xls")).unwrap()).unwrap();
        let comments = simple.xls_worksheet(0).unwrap().comments();
        assert_eq!(comments.len(), 3);
        assert_eq!(comments[0].author(), "Yegor Kozlov");
        assert_eq!(comments[0].text(), "Yegor Kozlov:\nfirst cell");
        assert_eq!(comments[1].text(), "Yegor Kozlov:\nsecond cell");
        assert_eq!(comments[2].visibility(), CommentVisibility::Visible);
        assert_eq!(comments[0].identity().object_id(), 1);
        assert_ne!(comments[0].identity().guid(), &[0; 16]);
        assert_eq!(comments[0].text_runs().len(), 2);

        let drawing =
            XlsWorkbook::new(File::open(fixture("DrawingAndComments.xls")).unwrap()).unwrap();
        let comments = drawing.xls_worksheet(0).unwrap().comments();
        assert_eq!(comments.len(), 3);
        assert!(comments.iter().all(|comment| !comment.text().is_empty()));

        let libreoffice = XlsWorkbook::new(File::open(fixture("comments.xls")).unwrap()).unwrap();
        let comments = libreoffice.xls_worksheet(0).unwrap().comments();
        assert_eq!(comments.len(), 3);
        assert!(
            comments
                .iter()
                .all(|comment| comment.author() == "Sven Nissel")
        );
        assert_eq!(comments[0].text(), "comment top row1 (index0)\n");
    }
}
