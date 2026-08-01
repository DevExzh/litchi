//! Transactional writer infrastructure for supported DOC embedded-object fields.
//!
//! The editor accepts only Word 97+ single-generation CLX layouts. It appends
//! new physical structures and retargets FIB entries; existing physical text,
//! FKPs, table structures, and unknown bytes are not relocated.

use super::package::{DocError, Result};
use super::writer::ChpxFkpBuilder;
use crate::{LegacyOfficeObjectEditor, LegacyOfficeObjectFormat, LegacyOfficeObjectLimits};

const FIB_CCP_TEXT: usize = 76;
const FIB_FC_LCB: usize = 154;
const PLCFBTE_CHPX: usize = 12;
const PLCFFLD_MOM: usize = 16;
const CLX: usize = 33;
const SPRM_C_PIC_LOCATION: u16 = 0x6A03;
const SPRM_C_F_OLE2: u16 = 0x080A;
const SPRM_C_F_SPEC: u16 = 0x0855;
const SPRM_C_F_OBJ: u16 = 0x0856;
const MAX_PIECES: usize = 65_536;
const MAX_FIELDS: usize = 65_536;
const MAX_PICF: usize = 128 * 1024 * 1024;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocEmbeddedObjectWriteOptions {
    pub storage_id: u32,
    pub instruction: String,
    /// Complete PICFAndOfficeArtData block for the Data stream.
    pub picture_data: Vec<u8>,
    /// Standalone CFB to install as `ObjectPool/_<storage_id>`.
    pub compound_file: Vec<u8>,
}

impl DocEmbeddedObjectWriteOptions {
    pub fn new(storage_id: u32, compound_file: Vec<u8>, picture_data: Vec<u8>) -> Self {
        Self {
            storage_id,
            instruction: format!(" EMBED LITCHI_OBJECT _{storage_id} "),
            picture_data,
            compound_file,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocEmbeddedObjectReference {
    pub storage_id: u32,
    pub storage_name: String,
    pub start_cp: u32,
    pub separator_cp: u32,
    pub end_cp: u32,
    pub data_offset: u32,
}

#[derive(Clone, Debug)]
struct RawPiece {
    start: u32,
    end: u32,
    fc: u32,
    unicode: bool,
    pcd_prefix: [u8; 2],
    prm: [u8; 2],
}

#[derive(Clone, Debug)]
struct FieldMarker {
    cp: u32,
    descriptor: [u8; 2],
}

#[derive(Clone)]
pub struct DocEmbeddedObjectEditor {
    package: LegacyOfficeObjectEditor,
    word_path: Vec<String>,
    table_path: Vec<String>,
    data_path: Vec<String>,
    word: Vec<u8>,
    table: Vec<u8>,
    data: Vec<u8>,
    pieces: Vec<RawPiece>,
    fields: Vec<FieldMarker>,
    main_ccp: u32,
    changed: bool,
}

impl DocEmbeddedObjectEditor {
    pub fn open(bytes: Vec<u8>, limits: LegacyOfficeObjectLimits) -> Result<Self> {
        let package = LegacyOfficeObjectEditor::open(&bytes, LegacyOfficeObjectFormat::Doc, limits)
            .map_err(DocError::from)?;
        let word_path = vec!["WordDocument".to_string()];
        let word = package
            .package_stream(&word_path)
            .ok_or_else(|| corrupted("WordDocument stream is missing"))?
            .to_vec();
        if word.len() < FIB_FC_LCB + (CLX + 1) * 8 || u16_at(&word, 0)? != 0xA5EC {
            return Err(corrupted("unsupported pre-Word-97 or truncated FIB"));
        }
        let flags = u16_at(&word, 10)?;
        if flags & 0x0100 != 0 || u32_at(&word, 14)? != 0 {
            return Err(corrupted("encrypted DOC cannot be edited"));
        }
        let table_path = vec![
            if flags & 0x0200 != 0 {
                "1Table"
            } else {
                "0Table"
            }
            .to_string(),
        ];
        let table = package
            .package_stream(&table_path)
            .ok_or_else(|| corrupted("selected Table stream is missing"))?
            .to_vec();
        let data_path = vec!["Data".to_string()];
        let data = package.package_stream(&data_path).unwrap_or(&[]).to_vec();
        let main_ccp = u32_at(&word, FIB_CCP_TEXT)?;
        let pieces = parse_clx(&word, &table)?;
        if pieces.last().is_none_or(|piece| piece.end < main_ccp) {
            return Err(corrupted("piece table does not cover the main story"));
        }
        let fields = parse_fields(&word, &table, main_ccp)?;
        validate_existing_fields(&fields, main_ccp)?;
        Ok(Self {
            package,
            word_path,
            table_path,
            data_path,
            word,
            table,
            data,
            pieces,
            fields,
            main_ccp,
            changed: false,
        })
    }

    pub fn is_changed(&self) -> bool {
        self.changed
    }

    pub fn objects(&self) -> Result<Vec<DocEmbeddedObjectReference>> {
        managed_objects(&self.word, &self.pieces, &self.fields)
    }

    /// Adds an object at the main-story boundary. No existing logical range is shifted
    /// outside the main story; subsequent story piece CPs are shifted consistently.
    pub fn add(
        &mut self,
        options: DocEmbeddedObjectWriteOptions,
    ) -> Result<DocEmbeddedObjectReference> {
        validate_options(&options, self)?;
        let mut candidate = self.clone();
        let storage_name = format!("_{}", options.storage_id);
        if candidate.package.objects().find(&storage_name).is_some() {
            return Err(corrupted("ObjectPool storage identifier already exists"));
        }

        let data_offset = align4(candidate.data.len())?;
        candidate.data.resize(data_offset, 0);
        candidate.data.extend_from_slice(&options.picture_data);
        let data_offset =
            u32::try_from(data_offset).map_err(|_| corrupted("Data offset exceeds u32"))?;

        let start_cp = candidate.main_ccp;
        let mut text = vec![0x13u16];
        text.extend(options.instruction.encode_utf16());
        let separator_cp =
            start_cp + u32::try_from(text.len()).map_err(|_| corrupted("field length overflow"))?;
        text.push(0x14);
        let result_cp = separator_cp + 1;
        text.push(0x01);
        let end_cp = result_cp + 1;
        text.push(0x15);
        let added_cps =
            u32::try_from(text.len()).map_err(|_| corrupted("field length overflow"))?;

        let text_fc = align2(candidate.word.len())?;
        candidate.word.resize(text_fc, 0);
        for unit in &text {
            candidate.word.extend_from_slice(&unit.to_le_bytes());
        }
        let text_fc_u32 =
            u32::try_from(text_fc).map_err(|_| corrupted("WordDocument offset exceeds u32"))?;
        let text_end_u32 = u32::try_from(candidate.word.len())
            .map_err(|_| corrupted("WordDocument offset exceeds u32"))?;
        insert_piece_at_main_end(&mut candidate.pieces, start_cp, added_cps, text_fc_u32)?;
        shift_markers(&mut candidate.fields, start_cp, added_cps)?;
        candidate.fields.extend([
            FieldMarker {
                cp: start_cp,
                descriptor: [0x13, 0x3A],
            },
            FieldMarker {
                cp: separator_cp,
                descriptor: [0x14, 0],
            },
            FieldMarker {
                cp: end_cp,
                descriptor: [0x15, 0x80],
            },
        ]);
        candidate.fields.sort_by_key(|marker| marker.cp);
        candidate.main_ccp = candidate
            .main_ccp
            .checked_add(added_cps)
            .ok_or_else(|| corrupted("main story CP overflow"))?;

        candidate.append_object_chpx(
            text_fc_u32,
            &text,
            separator_cp - start_cp,
            result_cp - start_cp,
            options.storage_id,
            data_offset,
        )?;
        candidate.append_table_replacements()?;
        put_u32(&mut candidate.word, FIB_CCP_TEXT, candidate.main_ccp)?;
        put_u32(&mut candidate.word, 28, text_end_u32)?;
        let cb_mac = u32::try_from(candidate.word.len())
            .map_err(|_| corrupted("WordDocument size exceeds u32"))?;
        put_u32(&mut candidate.word, 64, cb_mac)?;
        candidate
            .package
            .replace_package_stream(&candidate.word_path, candidate.word.clone())
            .map_err(DocError::from)?;
        candidate
            .package
            .replace_package_stream(&candidate.table_path, candidate.table.clone())
            .map_err(DocError::from)?;
        if candidate
            .package
            .package_stream(&candidate.data_path)
            .is_some()
        {
            candidate
                .package
                .replace_package_stream(&candidate.data_path, candidate.data.clone())
                .map_err(DocError::from)?;
        } else {
            candidate
                .package
                .add_package_stream(candidate.data_path.clone(), candidate.data.clone())
                .map_err(DocError::from)?;
        }
        candidate
            .package
            .add_referenced_storage(&storage_name, options.compound_file)
            .map_err(DocError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(DocEmbeddedObjectReference {
            storage_id: options.storage_id,
            storage_name,
            start_cp,
            separator_cp,
            end_cp,
            data_offset,
        })
    }

    pub fn remove(&mut self, storage_id: u32) -> Result<DocEmbeddedObjectReference> {
        let object = self
            .objects()?
            .into_iter()
            .find(|value| value.storage_id == storage_id)
            .ok_or_else(|| corrupted("managed embedded-object field was not found"))?;
        let mut candidate = self.clone();
        let remove_end = object
            .end_cp
            .checked_add(1)
            .ok_or_else(|| corrupted("field CP overflow"))?;
        let amount = remove_end - object.start_cp;
        let index = candidate
            .pieces
            .iter()
            .position(|piece| piece.start == object.start_cp && piece.end == remove_end)
            .ok_or_else(|| corrupted("object field does not occupy one dedicated piece"))?;
        candidate.pieces.remove(index);
        for piece in candidate.pieces.iter_mut().skip(index) {
            piece.start -= amount;
            piece.end -= amount;
        }
        candidate
            .fields
            .retain(|marker| marker.cp < object.start_cp || marker.cp >= remove_end);
        for marker in candidate
            .fields
            .iter_mut()
            .filter(|marker| marker.cp >= remove_end)
        {
            marker.cp -= amount;
        }
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(amount)
            .ok_or_else(|| corrupted("main CP underflow"))?;
        candidate.append_table_replacements()?;
        put_u32(&mut candidate.word, FIB_CCP_TEXT, candidate.main_ccp)?;
        candidate
            .package
            .replace_package_stream(&candidate.word_path, candidate.word.clone())
            .map_err(DocError::from)?;
        candidate
            .package
            .replace_package_stream(&candidate.table_path, candidate.table.clone())
            .map_err(DocError::from)?;
        candidate
            .package
            .remove_referenced_storage(&object.storage_name)
            .map_err(DocError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(object)
    }

    /// Reorders a contiguous suffix of editor-managed object fields.
    pub fn reorder(&mut self, storage_ids: &[u32]) -> Result<()> {
        let objects = self.objects()?;
        if storage_ids.len() != objects.len() || objects.is_empty() {
            return Err(corrupted(
                "reorder must contain every managed object exactly once",
            ));
        }
        let first_cp = objects.first().unwrap().start_cp;
        if objects.last().unwrap().end_cp + 1 != self.main_ccp
            || objects
                .windows(2)
                .any(|pair| pair[0].end_cp + 1 != pair[1].start_cp)
        {
            return Err(corrupted(
                "managed object fields are not a contiguous main-story suffix",
            ));
        }
        let mut candidate = self.clone();
        let prefix_piece_count = candidate
            .pieces
            .iter()
            .position(|piece| piece.start == first_cp)
            .ok_or_else(|| corrupted("managed object piece suffix is missing"))?;
        let suffix = candidate.pieces[prefix_piece_count..].to_vec();
        if suffix.len() != objects.len() {
            return Err(corrupted("managed object fields are not dedicated pieces"));
        }
        let mut reordered = Vec::with_capacity(suffix.len());
        let mut new_fields = candidate
            .fields
            .iter()
            .filter(|marker| marker.cp < first_cp)
            .cloned()
            .collect::<Vec<_>>();
        let mut cp = first_cp;
        let mut used = HashSet::new();
        for id in storage_ids {
            if !used.insert(*id) {
                return Err(corrupted("reorder contains a repeated storage ID"));
            }
            let ordinal = objects
                .iter()
                .position(|object| object.storage_id == *id)
                .ok_or_else(|| corrupted("reorder contains an unknown storage ID"))?;
            let object = &objects[ordinal];
            let mut piece = suffix[ordinal].clone();
            let len = piece.end - piece.start;
            piece.start = cp;
            piece.end = cp + len;
            new_fields.extend([
                FieldMarker {
                    cp,
                    descriptor: [0x13, 0x3A],
                },
                FieldMarker {
                    cp: cp + (object.separator_cp - object.start_cp),
                    descriptor: [0x14, 0],
                },
                FieldMarker {
                    cp: cp + (object.end_cp - object.start_cp),
                    descriptor: [0x15, 0x80],
                },
            ]);
            cp += len;
            reordered.push(piece);
        }
        candidate.pieces.truncate(prefix_piece_count);
        candidate.pieces.extend(reordered);
        candidate.fields = new_fields;
        candidate.append_table_replacements()?;
        candidate
            .package
            .replace_package_stream(&candidate.word_path, candidate.word.clone())
            .map_err(DocError::from)?;
        candidate
            .package
            .replace_package_stream(&candidate.table_path, candidate.table.clone())
            .map_err(DocError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(DocError::from)
    }

    fn append_object_chpx(
        &mut self,
        text_fc: u32,
        text: &[u16],
        separator: u32,
        result: u32,
        storage_id: u32,
        data_offset: u32,
    ) -> Result<()> {
        let byte_end = text_fc
            .checked_add(
                u32::try_from(text.len() * 2).map_err(|_| corrupted("text bytes overflow"))?,
            )
            .ok_or_else(|| corrupted("text FC overflow"))?;
        let sep_fc = text_fc + separator * 2;
        let result_fc = text_fc + result * 2;
        let (old_fc, old_pages) = parse_bte(&self.word, &self.table, PLCFBTE_CHPX)?;
        let fkp_start = old_fc.last().copied().unwrap_or(text_fc);
        if fkp_start > text_fc {
            return Err(corrupted(
                "new object FC overlaps the existing CHPX bin table",
            ));
        }
        let mut builder = ChpxFkpBuilder::new();
        if fkp_start < sep_fc {
            builder.add_entry(fkp_start, sep_fc, Vec::new());
        }
        builder.add_entry(sep_fc, sep_fc + 2, object_separator_sprms(storage_id));
        builder.add_entry(result_fc, result_fc + 2, object_preview_sprms(data_offset));
        if result_fc + 2 < byte_end {
            builder.add_entry(result_fc + 2, byte_end, Vec::new());
        }
        let pages = builder.generate_pages().map_err(DocError::from)?;
        if pages.pages.len() != 1 {
            return Err(corrupted("object CHPX unexpectedly spans multiple FKPs"));
        }
        let page_offset = align512(self.word.len())?;
        self.word.resize(page_offset, 0);
        self.word.extend_from_slice(&pages.pages[0]);
        let page_number =
            u32::try_from(page_offset / 512).map_err(|_| corrupted("FKP page exceeds u32"))?;
        let mut fc = old_fc;
        let mut page_numbers = old_pages;
        if fc.is_empty() {
            fc.push(fkp_start);
        }
        fc.push(byte_end);
        page_numbers.push(page_number);
        let mut plc = Vec::new();
        for value in fc {
            plc.extend_from_slice(&value.to_le_bytes());
        }
        for value in page_numbers {
            plc.extend_from_slice(&value.to_le_bytes());
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_CHPX, &plc)?;
        Ok(())
    }

    fn append_table_replacements(&mut self) -> Result<()> {
        let clx = serialize_clx(&self.pieces)?;
        append_table_block(&mut self.word, &mut self.table, CLX, &clx)?;
        let fields = serialize_fields(&self.fields, self.main_ccp)?;
        append_table_block(&mut self.word, &mut self.table, PLCFFLD_MOM, &fields)?;
        Ok(())
    }
}

use std::collections::HashSet;

fn managed_objects(
    word: &[u8],
    pieces: &[RawPiece],
    fields: &[FieldMarker],
) -> Result<Vec<DocEmbeddedObjectReference>> {
    let mut stack: Vec<(u32, u8, Option<u32>)> = Vec::new();
    let mut output = Vec::new();
    for marker in fields {
        match marker.descriptor[0] & 0x1F {
            0x13 => stack.push((marker.cp, marker.descriptor[1], None)),
            0x14 => {
                if let Some(value) = stack.last_mut() {
                    value.2 = Some(marker.cp);
                }
            },
            0x15 => {
                let Some((start, kind, Some(separator))) = stack.pop() else {
                    continue;
                };
                if kind != 0x3A || !stack.is_empty() {
                    continue;
                }
                let code = text_range(word, pieces, start + 1, separator)?;
                let Some(id_text) = [" EMBED LITCHI_OBJECT _", " EMBED Equation.3 _"]
                    .into_iter()
                    .find_map(|prefix| code.strip_prefix(prefix).and_then(|v| v.strip_suffix(' ')))
                else {
                    continue;
                };
                let Ok(storage_id) = id_text.parse::<u32>() else {
                    continue;
                };
                let end = marker.cp;
                if pieces
                    .iter()
                    .any(|piece| piece.start == start && piece.end == end + 1)
                {
                    output.push(DocEmbeddedObjectReference {
                        storage_id,
                        storage_name: format!("_{storage_id}"),
                        start_cp: start,
                        separator_cp: separator,
                        end_cp: end,
                        data_offset: 0,
                    });
                }
            },
            _ => {},
        }
    }
    Ok(output)
}

fn text_range(word: &[u8], pieces: &[RawPiece], start: u32, end: u32) -> Result<String> {
    let mut units = Vec::new();
    let mut cp = start;
    while cp < end {
        let piece = pieces
            .iter()
            .find(|piece| cp >= piece.start && cp < piece.end)
            .ok_or_else(|| corrupted("field code CP has no text piece"))?;
        let take_end = end.min(piece.end);
        let count = (take_end - cp) as usize;
        let relative = (cp - piece.start) as usize;
        if piece.unicode {
            let offset = piece.fc as usize + relative * 2;
            let bytes = word
                .get(offset..offset + count * 2)
                .ok_or_else(|| corrupted("field code exceeds WordDocument"))?;
            for pair in bytes.chunks_exact(2) {
                units.push(u16::from_le_bytes([pair[0], pair[1]]));
            }
        } else {
            let offset = piece.fc as usize + relative;
            let bytes = word
                .get(offset..offset + count)
                .ok_or_else(|| corrupted("field code exceeds WordDocument"))?;
            units.extend(bytes.iter().map(|byte| u16::from(*byte)));
        }
        cp = take_end;
    }
    String::from_utf16(&units).map_err(|_| corrupted("field instruction contains invalid UTF-16"))
}

fn serialize_clx(pieces: &[RawPiece]) -> Result<Vec<u8>> {
    if pieces.is_empty() || pieces.len() > MAX_PIECES {
        return Err(corrupted("piece table cardinality is invalid"));
    }
    let plc_size = pieces
        .len()
        .checked_mul(12)
        .and_then(|value| value.checked_add(4))
        .ok_or_else(|| corrupted("PlcPcd size overflow"))?;
    let mut output = vec![2];
    output.extend_from_slice(
        &u32::try_from(plc_size)
            .map_err(|_| corrupted("PlcPcd exceeds u32"))?
            .to_le_bytes(),
    );
    for piece in pieces {
        output.extend_from_slice(&piece.start.to_le_bytes());
    }
    output.extend_from_slice(&pieces.last().expect("nonempty").end.to_le_bytes());
    for piece in pieces {
        output.extend_from_slice(&piece.pcd_prefix);
        let raw_fc = if piece.unicode {
            piece.fc
        } else {
            piece
                .fc
                .checked_mul(2)
                .ok_or_else(|| corrupted("compressed FC overflow"))?
                | 0x4000_0000
        };
        if raw_fc & 0x8000_0000 != 0 {
            return Err(corrupted("FC uses reserved high bit"));
        }
        output.extend_from_slice(&raw_fc.to_le_bytes());
        output.extend_from_slice(&piece.prm);
    }
    Ok(output)
}

fn parse_clx(word: &[u8], table: &[u8]) -> Result<Vec<RawPiece>> {
    let (offset, length) = fib_pair(word, CLX)?;
    let data = slice(table, offset, length, "CLX")?;
    if data.first() != Some(&2) {
        return Err(corrupted("fast-save RgPrc CLX is unsupported"));
    }
    let size = u32_at(data, 1)? as usize;
    if size + 5 != data.len() || size < 4 || !(size - 4).is_multiple_of(12) {
        return Err(corrupted("CLX PlcPcd size is invalid"));
    }
    let count = (size - 4) / 12;
    if count == 0 || count > MAX_PIECES {
        return Err(corrupted("piece count is unsupported"));
    }
    let cps = &data[5..5 + (count + 1) * 4];
    let pcds = &data[5 + (count + 1) * 4..];
    let mut pieces = Vec::with_capacity(count);
    for index in 0..count {
        let start = u32_at(cps, index * 4)?;
        let end = u32_at(cps, (index + 1) * 4)?;
        if start >= end
            || pieces
                .last()
                .is_some_and(|last: &RawPiece| last.end != start)
        {
            return Err(corrupted("piece CPs overlap or contain gaps"));
        }
        let pcd = &pcds[index * 8..index * 8 + 8];
        let raw_fc = u32_at(pcd, 2)?;
        let unicode = raw_fc & 0x4000_0000 == 0;
        let fc = if unicode {
            raw_fc & 0x3FFF_FFFF
        } else {
            (raw_fc & 0x3FFF_FFFF) / 2
        };
        let byte_len = (end - start)
            .checked_mul(if unicode { 2 } else { 1 })
            .ok_or_else(|| corrupted("piece length overflow"))?;
        if fc
            .checked_add(byte_len)
            .is_none_or(|end| end as usize > word.len())
        {
            return Err(corrupted("piece text exceeds WordDocument"));
        }
        pieces.push(RawPiece {
            start,
            end,
            fc,
            unicode,
            pcd_prefix: pcd[..2].try_into().unwrap(),
            prm: pcd[6..8].try_into().unwrap(),
        });
    }
    if pieces.iter().any(|piece| piece.prm != [0, 0]) {
        return Err(corrupted("piece-level fast-save SPRMs are unsupported"));
    }
    Ok(pieces)
}

fn parse_fields(word: &[u8], table: &[u8], main_ccp: u32) -> Result<Vec<FieldMarker>> {
    let (offset, length) = fib_pair(word, PLCFFLD_MOM)?;
    if length == 0 {
        return Ok(Vec::new());
    }
    let data = slice(table, offset, length, "PlcfFldMom")?;
    if data.len() < 4 || (data.len() - 4) % 6 != 0 {
        return Err(corrupted("PlcfFldMom length is invalid"));
    }
    let count = (data.len() - 4) / 6;
    if count > MAX_FIELDS {
        return Err(corrupted("field count exceeds limit"));
    }
    let cp_bytes = (count + 1) * 4;
    let mut output = Vec::with_capacity(count);
    for index in 0..count {
        let cp = u32_at(data, index * 4)?;
        if cp >= main_ccp
            || output
                .last()
                .is_some_and(|last: &FieldMarker| last.cp >= cp)
        {
            return Err(corrupted("field marker CPs are invalid"));
        }
        output.push(FieldMarker {
            cp,
            descriptor: data[cp_bytes + index * 2..cp_bytes + index * 2 + 2]
                .try_into()
                .unwrap(),
        });
    }
    Ok(output)
}

fn validate_existing_fields(fields: &[FieldMarker], main_ccp: u32) -> Result<()> {
    let mut stack = Vec::new();
    for marker in fields {
        match marker.descriptor[0] & 0x1F {
            0x13 => stack.push((marker.cp, marker.descriptor[1], false)),
            0x14 => {
                let Some(value) = stack.last_mut() else {
                    return Err(corrupted("orphan field separator"));
                };
                if value.2 {
                    return Err(corrupted("duplicate field separator"));
                }
                value.2 = true;
            },
            0x15 => {
                if stack.pop().is_none() {
                    return Err(corrupted("orphan field end"));
                }
            },
            _ => return Err(corrupted("invalid field marker descriptor")),
        }
    }
    if !stack.is_empty() || fields.last().is_some_and(|marker| marker.cp >= main_ccp) {
        return Err(corrupted("unclosed field structure"));
    }
    Ok(())
}

fn validate_options(
    value: &DocEmbeddedObjectWriteOptions,
    editor: &DocEmbeddedObjectEditor,
) -> Result<()> {
    if value.storage_id == 0 || value.storage_id > i32::MAX as u32 {
        return Err(corrupted("storage ID must be a positive signed integer"));
    }
    if value.instruction.is_empty()
        || value.instruction.encode_utf16().count() > 4_096
        || value
            .instruction
            .chars()
            .any(|c| matches!(c, '\u{13}' | '\u{14}' | '\u{15}'))
    {
        return Err(corrupted("object instruction is invalid"));
    }
    if value.picture_data.len() < 4
        || value.picture_data.len() > MAX_PICF
        || u32_at(&value.picture_data, 0)? as usize != value.picture_data.len()
    {
        return Err(corrupted("PICF block length prefix is invalid"));
    }
    if editor.pieces.len() + 2 > MAX_PIECES || editor.fields.len() + 3 > MAX_FIELDS {
        return Err(corrupted("object insertion exceeds resource limits"));
    }
    Ok(())
}

fn insert_piece_at_main_end(pieces: &mut Vec<RawPiece>, cp: u32, len: u32, fc: u32) -> Result<()> {
    let index = pieces
        .iter()
        .position(|piece| cp >= piece.start && cp <= piece.end)
        .ok_or_else(|| corrupted("main story boundary is outside piece table"))?;
    let old = pieces[index].clone();
    let width = if old.unicode { 2 } else { 1 };
    let mut replacement = Vec::new();
    if old.start < cp {
        let mut before = old.clone();
        before.end = cp;
        replacement.push(before);
    }
    replacement.push(RawPiece {
        start: cp,
        end: cp + len,
        fc,
        unicode: true,
        pcd_prefix: [0; 2],
        prm: [0; 2],
    });
    if cp < old.end {
        let mut after = old.clone();
        after.start = cp + len;
        after.end = old.end + len;
        after.fc = old.fc + (cp - old.start) * width;
        replacement.push(after);
    }
    let replacement_len = replacement.len();
    pieces.splice(index..=index, replacement);
    for piece in pieces.iter_mut().skip(index + replacement_len) {
        piece.start = piece
            .start
            .checked_add(len)
            .ok_or_else(|| corrupted("piece CP overflow"))?;
        piece.end = piece
            .end
            .checked_add(len)
            .ok_or_else(|| corrupted("piece CP overflow"))?;
    }
    Ok(())
}

fn shift_markers(fields: &mut [FieldMarker], cp: u32, amount: u32) -> Result<()> {
    for marker in fields.iter_mut().filter(|marker| marker.cp >= cp) {
        marker.cp = marker
            .cp
            .checked_add(amount)
            .ok_or_else(|| corrupted("field CP overflow"))?;
    }
    Ok(())
}

fn serialize_fields(fields: &[FieldMarker], terminal: u32) -> Result<Vec<u8>> {
    validate_existing_fields(fields, terminal)?;
    let mut output = Vec::new();
    for marker in fields {
        output.extend_from_slice(&marker.cp.to_le_bytes());
    }
    output.extend_from_slice(&terminal.to_le_bytes());
    for marker in fields {
        output.extend_from_slice(&marker.descriptor);
    }
    Ok(output)
}

fn object_separator_sprms(storage_id: u32) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
    output.extend_from_slice(&storage_id.to_le_bytes());
    for (opcode, value) in [
        (SPRM_C_F_OLE2, true),
        (SPRM_C_F_SPEC, true),
        (SPRM_C_F_OBJ, true),
    ] {
        output.extend_from_slice(&opcode.to_le_bytes());
        output.push(u8::from(value));
    }
    output
}

fn object_preview_sprms(data_offset: u32) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
    output.extend_from_slice(&data_offset.to_le_bytes());
    output.extend_from_slice(&SPRM_C_F_SPEC.to_le_bytes());
    output.push(1);
    output
}

fn parse_bte(word: &[u8], table: &[u8], index: usize) -> Result<(Vec<u32>, Vec<u32>)> {
    let (offset, length) = fib_pair(word, index)?;
    let data = slice(table, offset, length, "PlcBteChpx")?;
    if data.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    if data.len() < 4 || (data.len() - 4) % 8 != 0 {
        return Err(corrupted(format!(
            "PlcBteChpx length {} is invalid",
            data.len()
        )));
    }
    let count = (data.len() - 4) / 8;
    let mut fc = Vec::with_capacity(count + 1);
    let mut pages = Vec::with_capacity(count);
    for i in 0..=count {
        fc.push(u32_at(data, i * 4)?);
    }
    for i in 0..count {
        pages.push(u32_at(data, (count + 1) * 4 + i * 4)?);
    }
    if fc.windows(2).any(|v| v[0] >= v[1])
        || pages
            .iter()
            .any(|pn| (*pn as usize) * 512 + 512 > word.len())
    {
        return Err(corrupted("PlcBteChpx references invalid FKPs"));
    }
    Ok((fc, pages))
}

fn append_table_block(
    word: &mut [u8],
    table: &mut Vec<u8>,
    index: usize,
    data: &[u8],
) -> Result<()> {
    let offset = u32::try_from(table.len()).map_err(|_| corrupted("Table stream exceeds u32"))?;
    table.extend_from_slice(data);
    put_fib_pair(
        word,
        index,
        offset,
        u32::try_from(data.len()).map_err(|_| corrupted("table block exceeds u32"))?,
    )
}
fn fib_pair(word: &[u8], index: usize) -> Result<(u32, u32)> {
    Ok((
        u32_at(word, FIB_FC_LCB + index * 8)?,
        u32_at(word, FIB_FC_LCB + index * 8 + 4)?,
    ))
}
fn put_fib_pair(word: &mut [u8], index: usize, fc: u32, lcb: u32) -> Result<()> {
    put_u32(word, FIB_FC_LCB + index * 8, fc)?;
    put_u32(word, FIB_FC_LCB + index * 8 + 4, lcb)
}
fn slice<'a>(data: &'a [u8], offset: u32, length: u32, name: &str) -> Result<&'a [u8]> {
    let start = offset as usize;
    let end = start
        .checked_add(length as usize)
        .ok_or_else(|| corrupted(format!("{name} range overflow")))?;
    data.get(start..end)
        .ok_or_else(|| corrupted(format!("{name} exceeds stream")))
}
fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    data.get(offset..offset + 2)
        .map(|v| u16::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| corrupted("truncated u16"))
}
fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| corrupted("truncated u32"))
}
fn put_u32(data: &mut [u8], offset: usize, value: u32) -> Result<()> {
    let slot = data
        .get_mut(offset..offset + 4)
        .ok_or_else(|| corrupted("truncated FIB field"))?;
    slot.copy_from_slice(&value.to_le_bytes());
    Ok(())
}
fn align2(value: usize) -> Result<usize> {
    value
        .checked_add(1)
        .map(|v| v & !1)
        .ok_or_else(|| corrupted("alignment overflow"))
}
fn align4(value: usize) -> Result<usize> {
    value
        .checked_add(3)
        .map(|v| v & !3)
        .ok_or_else(|| corrupted("alignment overflow"))
}
fn align512(value: usize) -> Result<usize> {
    value
        .checked_add(511)
        .map(|v| v & !511)
        .ok_or_else(|| corrupted("alignment overflow"))
}
fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}
