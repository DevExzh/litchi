//! BIFF8 comment record linkage and bounded codecs.

use super::model::{
    COMMENT_OBJECT_TYPE, CONTINUE_TYPE, Comment, CommentRecord, HorizontalAlignment, MAX_COMMENTS,
    MAX_RECORD_BYTES, MAX_RETAINED_BYTES, MAX_RETAINED_RECORDS, MSODRAWING_TYPE, NoteMetadata,
    OBJ_TYPE, ObjectIdentity, ObjectPadding, ObjectProperties, ObjectSubrecord, RECORD_TYPE,
    RecordKind, TXO_TYPE, TextOrientation, TextProperties, TextRun, VerticalAlignment, Visibility,
    boxed_bytes,
};
use crate::error::{Error, Result};
use std::collections::{HashMap, HashSet};

#[derive(Debug)]
pub(crate) struct NoteRecord {
    pub(crate) row: u16,
    pub(crate) column: u8,
    pub(crate) visibility: Visibility,
    pub(crate) row_hidden: bool,
    pub(crate) column_hidden: bool,
    pub(crate) object_id: u16,
    pub(crate) author: String,
    pub(crate) metadata: NoteMetadata,
    pub(crate) record: CommentRecord,
    pub(crate) order: usize,
}

#[derive(Debug)]
struct OrderedRecord {
    order: usize,
    record: CommentRecord,
}

#[derive(Debug)]
struct CommentObject {
    identity: ObjectIdentity,
    properties: ObjectProperties,
    subrecords: Vec<ObjectSubrecord>,
    padding: ObjectPadding,
    text: Option<String>,
    text_properties: Option<TextProperties>,
    text_runs: Vec<TextRun>,
    records: Vec<OrderedRecord>,
}

#[derive(Debug)]
struct PendingTxo {
    object_id: u16,
    character_count: usize,
    run_byte_count: usize,
    code_units: Vec<u16>,
    run_bytes: Vec<u8>,
    properties: TextProperties,
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
    retained_records: usize,
    retained_bytes: usize,
    next_order: usize,
}

impl CommentCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        if matches!(
            record_type,
            OBJ_TYPE | MSODRAWING_TYPE | TXO_TYPE | CONTINUE_TYPE | RECORD_TYPE
        ) && data.len() > MAX_RECORD_BYTES
        {
            return invalid("comment record payload exceeds the BIFF8 record bound".to_string());
        }
        if let Some(mut pending) = self.pending_txo.take() {
            if record_type != CONTINUE_TYPE {
                return invalid(format!(
                    "incomplete TXO for comment object {} must be followed by CONTINUE",
                    pending.object_id
                ));
            }
            let complete = feed_txo_continue(&mut pending, data)?;
            self.retain_object_record(pending.object_id, RecordKind::Continue, data)?;
            if complete {
                self.complete_txo(pending)?;
            } else {
                self.pending_txo = Some(pending);
            }
            return Ok(());
        }

        if let Some(object_id) = self.awaiting_drawing {
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
                    "comment object {object_id} must be followed by an OfficeArtClientTextbox boundary"
                ));
            }
            self.retain_object_record(object_id, RecordKind::Drawing, data)?;
            self.awaiting_drawing = None;
            self.awaiting_txo = Some(object_id);
            return Ok(());
        }

        if let Some(object_id) = self.awaiting_txo {
            if record_type != TXO_TYPE {
                return invalid(format!(
                    "comment object {object_id} textbox must be followed by TXO"
                ));
            }
            let pending = parse_txo(data, object_id)?;
            self.retain_object_record(object_id, RecordKind::TextObject, data)?;
            self.awaiting_txo = None;
            if pending.character_count == 0 {
                self.complete_txo(pending)?;
            } else {
                self.pending_txo = Some(pending);
            }
            return Ok(());
        }

        match record_type {
            OBJ_TYPE => {
                let object_id = parse_cmo(data)?.object_id;
                if object_id == 0 {
                    return invalid("OBJ object id must not be zero".to_string());
                }
                self.object_ids
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("retaining OBJ identifiers"))?;
                if !self.object_ids.insert(object_id) {
                    return invalid(format!("duplicate OBJ object id: {object_id}"));
                }
                if let Some(object) = parse_obj(data)? {
                    if self.objects.len() >= MAX_COMMENTS {
                        return invalid(
                            "worksheet comment count exceeds the BIFF8 object bound".to_string(),
                        );
                    }
                    self.comment_guids
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("retaining comment identities"))?;
                    if !self.comment_guids.insert(object.identity.guid) {
                        return invalid("duplicate comment FtNts GUID".to_string());
                    }
                    self.objects
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("retaining worksheet comments"))?;
                    self.objects.insert(object_id, object);
                    self.retain_object_record(object_id, RecordKind::Object, data)?;
                    self.awaiting_drawing = Some(object_id);
                }
            },
            RECORD_TYPE => {
                let mut note = parse_note_record(data)?;
                self.note_cells
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("retaining NOTE cell identities"))?;
                if !self.note_cells.insert((note.row, note.column)) {
                    return invalid(format!(
                        "duplicate NOTE for cell ({}, {})",
                        note.row, note.column
                    ));
                }
                self.note_object_ids
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("retaining NOTE object identities"))?;
                if !self.note_object_ids.insert(note.object_id) {
                    return invalid(format!(
                        "duplicate NOTE object reference: {}",
                        note.object_id
                    ));
                }
                if self.notes.len() >= MAX_COMMENTS {
                    return invalid(
                        "worksheet NOTE count exceeds the BIFF8 object bound".to_string(),
                    );
                }
                self.reserve_retained(data.len())?;
                note.order = self.take_order()?;
                self.notes
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("retaining worksheet NOTE records"))?;
                self.notes.push(note);
            },
            _ => {},
        }
        Ok(())
    }

    fn reserve_retained(&mut self, byte_count: usize) -> Result<()> {
        if byte_count > MAX_RECORD_BYTES {
            return invalid("comment record payload exceeds the BIFF8 record bound".to_string());
        }
        let records = self
            .retained_records
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("comment record count overflow".to_string()))?;
        if records > MAX_RETAINED_RECORDS {
            return invalid("retained comment record count exceeds the safety bound".to_string());
        }
        let bytes = self.retained_bytes.checked_add(byte_count).ok_or_else(|| {
            Error::InvalidData("retained comment byte count overflow".to_string())
        })?;
        if bytes > MAX_RETAINED_BYTES {
            return invalid("retained comment bytes exceed the safety bound".to_string());
        }
        self.retained_records = records;
        self.retained_bytes = bytes;
        Ok(())
    }

    fn take_order(&mut self) -> Result<usize> {
        let order = self.next_order;
        self.next_order = self
            .next_order
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("comment record order overflow".to_string()))?;
        Ok(order)
    }

    fn retain_object_record(
        &mut self,
        object_id: u16,
        kind: RecordKind,
        data: &[u8],
    ) -> Result<()> {
        self.reserve_retained(data.len())?;
        let record = CommentRecord::new(kind, data)?;
        let order = self.take_order()?;
        let object = self.objects.get_mut(&object_id).ok_or_else(|| {
            Error::InvalidData(format!(
                "comment record references unknown object {object_id}"
            ))
        })?;
        object
            .records
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining comment record order"))?;
        object.records.push(OrderedRecord { order, record });
        Ok(())
    }

    fn complete_txo(&mut self, pending: PendingTxo) -> Result<()> {
        let text = String::from_utf16(&pending.code_units)
            .map_err(|_error| Error::InvalidData("TXO text contains invalid UTF-16".to_string()))?;
        let runs = parse_txo_runs(
            &pending.run_bytes,
            crate::utils::truncate_usize_to_u16(pending.character_count),
        )?;
        let object = self.objects.get_mut(&pending.object_id).ok_or_else(|| {
            Error::InvalidData(format!(
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

    pub(crate) fn finish(self) -> Result<Vec<Comment>> {
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
        let mut comments = Vec::new();
        comments
            .try_reserve_exact(self.notes.len())
            .map_err(|_error| Error::Allocation("finishing worksheet comments"))?;
        for note in self.notes {
            let object = self.objects.get(&note.object_id).ok_or_else(|| {
                Error::InvalidData(format!(
                    "NOTE references missing comment OBJ {}",
                    note.object_id
                ))
            })?;
            let text = object.text.as_deref().ok_or_else(|| {
                Error::InvalidData(format!("comment OBJ {} has no TXO", note.object_id))
            })?;
            let text_properties = object.text_properties.as_ref().ok_or_else(|| {
                Error::InvalidData(format!(
                    "comment OBJ {} has no TXO properties",
                    note.object_id
                ))
            })?;

            let mut ordered = Vec::new();
            let capacity = object
                .records
                .len()
                .checked_add(1)
                .ok_or(Error::Allocation("ordering comment records"))?;
            ordered
                .try_reserve_exact(capacity)
                .map_err(|_error| Error::Allocation("ordering comment records"))?;
            ordered.extend(object.records.iter().map(|value| OrderedRecord {
                order: value.order,
                record: value.record.clone(),
            }));
            ordered.push(OrderedRecord {
                order: note.order,
                record: note.record,
            });
            ordered.sort_by_key(|value| value.order);
            let mut records = Vec::new();
            records
                .try_reserve_exact(ordered.len())
                .map_err(|_error| Error::Allocation("finishing comment record order"))?;
            records.extend(ordered.into_iter().map(|value| value.record));

            let mut text_runs = Vec::new();
            text_runs
                .try_reserve_exact(object.text_runs.len())
                .map_err(|_error| Error::Allocation("finishing comment text runs"))?;
            text_runs.extend_from_slice(&object.text_runs);

            let mut subrecords = Vec::new();
            subrecords
                .try_reserve_exact(object.subrecords.len())
                .map_err(|_error| Error::Allocation("finishing OBJ subrecords"))?;
            subrecords.extend(object.subrecords.iter().cloned());

            comments.push(Comment {
                row: note.row,
                column: note.column,
                visibility: note.visibility,
                row_hidden: note.row_hidden,
                column_hidden: note.column_hidden,
                identity: object.identity.clone(),
                object_properties: object.properties.clone(),
                object_subrecords: subrecords,
                object_padding: object.padding.clone(),
                note_metadata: note.metadata,
                author: clone_string(&note.author, "finishing comment author")?,
                text: clone_string(text, "finishing comment text")?,
                text_properties: text_properties.clone(),
                text_runs,
                records,
            });
        }
        Ok(comments)
    }
}

pub(crate) fn parse_note_record(data: &[u8]) -> Result<NoteRecord> {
    if data.len() < 13 {
        return invalid(format!("NOTE payload is too short: {}", data.len()));
    }
    let row = u16_at(data, 0)?;
    let column = u16_at(data, 2)?;
    if column > 255 {
        return invalid(format!("NOTE column exceeds BIFF8 limit: {column}"));
    }
    let flags = u16_at(data, 4)?;
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
    let width = if string_flags & 1 == 0 { 1 } else { 2 };
    let byte_count = character_count
        .checked_mul(width)
        .ok_or_else(|| Error::InvalidData("NOTE author size overflow".to_string()))?;
    let expected = 12usize
        .checked_add(byte_count)
        .ok_or_else(|| Error::InvalidData("NOTE size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(format!(
            "NOTE payload length must be {expected}, got {}",
            data.len()
        ));
    }
    let author = decode_unicode(&data[11..11 + byte_count], width == 2)?;
    Ok(NoteRecord {
        row,
        column: crate::utils::truncate_u16_to_u8(column),
        visibility: if flags & 0x0002 != 0 {
            Visibility::Visible
        } else {
            Visibility::Hidden
        },
        row_hidden: flags & 0x0080 != 0,
        column_hidden: flags & 0x0100 != 0,
        object_id,
        author,
        metadata: NoteMetadata {
            reserved_flags: flags & !0x0182,
            reserved_string_flags: string_flags & !1,
            unused: data[11 + byte_count],
        },
        record: CommentRecord::new(RecordKind::Note, data)?,
        order: 0,
    })
}

fn parse_cmo(data: &[u8]) -> Result<ObjectProperties> {
    if data.len() < 22 {
        return invalid(format!(
            "OBJ payload is too short for FtCmo: {}",
            data.len()
        ));
    }
    let flags = u16_at(data, 8)?;
    Ok(ObjectProperties {
        object_type: u16_at(data, 4)?,
        object_id: u16_at(data, 6)?,
        flags,
        reserved_flags: flags & 0x0002,
        reserved_header: data[0..4].try_into().unwrap(),
        unused: data[10..22].try_into().unwrap(),
    })
}

fn parse_obj(data: &[u8]) -> Result<Option<CommentObject>> {
    let properties = parse_cmo(data)?;
    if properties.object_type != COMMENT_OBJECT_TYPE {
        return Ok(None);
    }
    if properties.object_id == 0 {
        return invalid("comment OBJ id must not be zero".to_string());
    }
    let mut position = 22usize;
    let mut identity = None;
    let mut subrecords = Vec::new();
    subrecords
        .try_reserve(4)
        .map_err(|_error| Error::Allocation("retaining OBJ subrecords"))?;
    let mut padding = ObjectPadding::new(&[])?;
    let mut found_end = false;
    while position < data.len() {
        if data.len() - position < 4 {
            padding = ObjectPadding::new(&data[position..])?;
            break;
        }
        let sub_type = u16_at(data, position)?;
        let size = u16_at(data, position + 2)? as usize;
        let body_start = position + 4;
        let end = body_start
            .checked_add(size)
            .ok_or_else(|| Error::InvalidData("OBJ subrecord size overflow".to_string()))?;
        let body = data
            .get(body_start..end)
            .ok_or_else(|| Error::InvalidData("truncated OBJ subrecord".to_string()))?;
        let known = sub_type == 0 || sub_type == 0x000D;
        subrecords
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("retaining OBJ subrecords"))?;
        subrecords.push(ObjectSubrecord::new(sub_type, body, known)?);
        position = end;
        if sub_type == 0 && size == 0 {
            padding = ObjectPadding::new(&data[position..])?;
            found_end = true;
            break;
        }
        if sub_type == 0x000D {
            if size != 0x0016 {
                return invalid(format!("FtNts size must be 22, got {size}"));
            }
            if identity.is_some() {
                return invalid("comment OBJ contains more than one FtNts".to_string());
            }
            let guid: [u8; 16] = body[0..16].try_into().unwrap();
            let shared_value = u16::from_le_bytes([body[16], body[17]]);
            identity = Some(ObjectIdentity {
                object_id: properties.object_id,
                guid,
                shared: shared_value != 0,
                shared_value,
                unused: body[18..22].try_into().unwrap(),
            });
        }
    }
    if !found_end {
        return invalid("comment OBJ is missing FtEnd".to_string());
    }
    Ok(Some(CommentObject {
        identity: identity
            .ok_or_else(|| Error::InvalidData("comment OBJ is missing FtNts".to_string()))?,
        properties,
        subrecords,
        padding,
        text: None,
        text_properties: None,
        text_runs: Vec::new(),
        records: Vec::new(),
    }))
}

fn parse_txo(data: &[u8], object_id: u16) -> Result<PendingTxo> {
    if data.len() < 18 {
        return invalid(format!("TXO payload is too short: {}", data.len()));
    }
    let options = u16_at(data, 0)?;
    let horizontal_alignment = horizontal_alignment(((options >> 1) & 7) as u8);
    let vertical_alignment = vertical_alignment(((options >> 4) & 7) as u8);
    let orientation = match u16_at(data, 2)? {
        0 => TextOrientation::None,
        1 => TextOrientation::Stacked,
        2 => TextOrientation::CounterClockwise,
        3 => TextOrientation::Clockwise,
        value => TextOrientation::Unknown(value),
    };
    let character_count = u16_at(data, 10)? as usize;
    let run_byte_count = u16_at(data, 12)? as usize;
    if character_count == 0 {
        if run_byte_count != 0 {
            return invalid("empty TXO text must have zero run bytes".to_string());
        }
    } else if run_byte_count < 16 || !run_byte_count.is_multiple_of(8) {
        return invalid(format!(
            "non-empty TXO run size must be a multiple of 8 and at least 16, got {run_byte_count}"
        ));
    }
    let formula_size = u16_at(data, 16)? as usize;
    if !formula_size.is_multiple_of(2) {
        return invalid(format!("TXO ObjFmla size must be even, got {formula_size}"));
    }
    let expected = 18usize
        .checked_add(formula_size)
        .ok_or_else(|| Error::InvalidData("TXO formula size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(format!(
            "TXO payload length must be {expected}, got {}",
            data.len()
        ));
    }
    let mut code_units = Vec::new();
    code_units
        .try_reserve_exact(character_count)
        .map_err(|_error| Error::Allocation("retaining TXO text"))?;
    let mut run_bytes = Vec::new();
    run_bytes
        .try_reserve_exact(run_byte_count)
        .map_err(|_error| Error::Allocation("retaining TXO formatting runs"))?;
    Ok(PendingTxo {
        object_id,
        character_count,
        run_byte_count,
        code_units,
        run_bytes,
        properties: TextProperties {
            horizontal_alignment,
            vertical_alignment,
            orientation,
            locked: options & 0x0200 != 0,
            justify_last_line: options & 0x4000 != 0,
            secret_edit: options & 0x8000 != 0,
            font_when_empty: u16_at(data, 14)?,
            reserved_options: options & !0xC27E,
            reserved_fields: data[4..10].try_into().unwrap(),
            formula_bytes: boxed_bytes(&data[18..], "retaining TXO formula bytes")?,
        },
    })
}

fn horizontal_alignment(value: u8) -> HorizontalAlignment {
    match value {
        1 => HorizontalAlignment::Left,
        2 => HorizontalAlignment::Centered,
        3 => HorizontalAlignment::Right,
        4 => HorizontalAlignment::Justified,
        7 => HorizontalAlignment::Distributed,
        value => HorizontalAlignment::Unknown(value),
    }
}

fn vertical_alignment(value: u8) -> VerticalAlignment {
    match value {
        1 => VerticalAlignment::Top,
        2 => VerticalAlignment::Centered,
        3 => VerticalAlignment::Bottom,
        4 => VerticalAlignment::Justified,
        7 => VerticalAlignment::Distributed,
        value => VerticalAlignment::Unknown(value),
    }
}

fn feed_txo_continue(pending: &mut PendingTxo, data: &[u8]) -> Result<bool> {
    if pending.code_units.len() < pending.character_count {
        let (&flags, characters) = data
            .split_first()
            .ok_or_else(|| Error::InvalidData("empty TXO text CONTINUE".to_string()))?;
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
            pending
                .code_units
                .try_reserve(segment_count)
                .map_err(|_error| Error::Allocation("retaining TXO text"))?;
            pending.code_units.extend(
                characters
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
            );
        } else {
            pending
                .code_units
                .try_reserve(segment_count)
                .map_err(|_error| Error::Allocation("retaining TXO text"))?;
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
    pending
        .run_bytes
        .try_reserve(data.len())
        .map_err(|_error| Error::Allocation("retaining TXO formatting runs"))?;
    pending.run_bytes.extend_from_slice(data);
    Ok(pending.run_bytes.len() == pending.run_byte_count)
}

pub(crate) fn parse_txo_runs(data: &[u8], character_count: u16) -> Result<Vec<TextRun>> {
    if data.is_empty() {
        return Ok(Vec::new());
    }
    if data.len() < 16 || !data.len().is_multiple_of(8) {
        return invalid("invalid TxORuns byte count".to_string());
    }
    let run_count = data.len() / 8 - 1;
    let mut runs = Vec::new();
    runs.try_reserve_exact(run_count)
        .map_err(|_error| Error::Allocation("retaining TXO formatting runs"))?;
    for index in 0..run_count {
        let offset = index * 8;
        let character_index = u16_at(data, offset)?;
        if character_index > character_count {
            return invalid(format!(
                "TXO formatting run index {character_index} exceeds cchText {character_count}"
            ));
        }
        runs.push(TextRun {
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

fn decode_unicode(data: &[u8], wide: bool) -> Result<String> {
    if wide {
        let mut words = Vec::new();
        words
            .try_reserve_exact(data.len() / 2)
            .map_err(|_error| Error::Allocation("decoding comment UTF-16 text"))?;
        words.extend(
            data.chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
        );
        String::from_utf16(&words)
            .map_err(|_error| Error::InvalidData("NOTE author contains invalid UTF-16".to_string()))
    } else {
        let mut value = String::new();
        value
            .try_reserve(data.len())
            .map_err(|_error| Error::Allocation("decoding comment compressed text"))?;
        value.extend(data.iter().map(|&byte| char::from(byte)));
        Ok(value)
    }
}

fn clone_string(value: &str, context: &'static str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_error| Error::Allocation(context))?;
    copy.push_str(value);
    Ok(copy)
}

fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| Error::InvalidData("truncated comment record".to_string()))?;
    Ok(u16::from_le_bytes(bytes.try_into().unwrap()))
}

fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| Error::InvalidData("truncated comment record".to_string()))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn invalid<T>(message: String) -> Result<T> {
    Err(Error::InvalidData(message))
}
