//! Embedded OLE object fields of the revision editor: the main-story
//! `EMBED` field, its `ObjectPool` storage and the preview picture block in
//! the Data stream ([MS-DOC] 2.1.4 and 2.6, "Embedded Object" and
//! "Pictures").
//!
//! An object field is five runs: the field begin (`0x13`), the instruction,
//! the separator (`0x14`) whose `sprmCPicLocation` names the storage
//! `ObjectPool/_<id>`, the picture character (`0x01`) whose
//! `sprmCPicLocation` is the offset of its PICF block in the Data stream,
//! and the field end (`0x15`).

use super::{FcRun, PLCFFLD_MOM, Result, RevisionEditor, corrupted, put_fib_pair, read_units};
use crate::sprm_operations::{
    SPRM_C_F_OBJ, SPRM_C_F_OLE2, SPRM_C_F_SPEC, SPRM_C_HPS_POS, SPRM_C_PIC_LOCATION,
};
use crate::tracked_revision::codec::{retain_sprms, strict_sprms};

/// The storage that holds the embedded objects of a Word document.
const OBJECT_POOL: &str = "ObjectPool";
/// `flt` of an `EMBED` field ([MS-DOC] 2.9.90).
const FLT_EMBED: u8 = 0x3A;
/// The `FLD` of an object field's separator as Word writes it: `ch` 0x14
/// with its reserved high bit set, and a reserved `0xFF`.
const FIELD_SEPARATOR: [u8; 2] = [0x94, 0xFF];
/// The first byte of an object field's end `FLD`: `ch` 0x15 with its
/// reserved high bit set, as Word writes it.
const FIELD_END: u8 = 0x95;
/// The length of a PICF header before its `OfficeArtSpContainer`.
const PICF_HEADER: usize = 0x44;

/// A new object field and the object it shows.
pub(crate) struct ObjectFieldInsert<'a> {
    /// The main-story CP the field begins at.
    pub(crate) at: u32,
    /// The field instruction, ` EMBED Equation.DSMT4 ` and its switches.
    pub(crate) instruction: &'a str,
    /// The character formatting of the text the field is inserted into,
    /// without picture or special-character SPRMs.
    pub(crate) base: &'a [u8],
    /// The `ObjectPool` storage number.
    pub(crate) storage_id: u32,
    /// The `grffld` of the field end: `fHasSep`, with `fResultDirty` for
    /// a MathType object as Word writes it.
    pub(crate) end_flags: u8,
    /// The complete PICF block of the preview picture.
    pub(crate) picture_block: &'a [u8],
    /// How far the preview is lowered below the baseline, in half-points.
    pub(crate) lowering: u16,
    /// The object storage, as a standalone compound file.
    pub(crate) compound_file: Vec<u8>,
}

fn object_pool_path() -> Vec<String> {
    vec![OBJECT_POOL.to_owned()]
}

fn storage_path(storage_id: u32) -> Vec<String> {
    vec![OBJECT_POOL.to_owned(), format!("_{storage_id}")]
}

/// `sprmCPicLocation` at `offset` and, for a lowered picture,
/// `sprmCHpsPos` moving it down by `lowering` half-points.
fn picture_sprms(offset: u32, lowering: u16) -> Result<Vec<u8>> {
    let mut grpprl = with_sprm(&[], SPRM_C_PIC_LOCATION, &offset.to_le_bytes());
    if lowering > 0 {
        let position = i16::try_from(lowering)
            .map(|lowering| -lowering)
            .map_err(|_error| corrupted("picture lowering exceeds sprmCHpsPos"))?;
        grpprl.extend_from_slice(&with_sprm(&[], SPRM_C_HPS_POS, &position.to_le_bytes()));
    }
    Ok(grpprl)
}

fn with_sprm(base: &[u8], opcode: u16, operand: &[u8]) -> Vec<u8> {
    let mut grpprl = base.to_vec();
    grpprl.extend_from_slice(&opcode.to_le_bytes());
    grpprl.extend_from_slice(operand);
    grpprl
}

impl RevisionEditor {
    /// The numbers of the `ObjectPool` storages named `_<decimal>`.
    pub(crate) fn object_storage_ids(&self) -> Vec<u32> {
        let pool = object_pool_path();
        self.package
            .child_storage_paths(&pool)
            .filter_map(|path| path.last()?.strip_prefix('_')?.parse().ok())
            .collect()
    }

    /// One stream of the object storage `ObjectPool/_<storage_id>`.
    pub(crate) fn object_stream(&self, storage_id: u32, name: &str) -> Option<&[u8]> {
        let mut path = storage_path(storage_id);
        path.push(name.to_owned());
        self.package.stream(&path)
    }

    /// The main-story text `start..end` as UTF-16 units.
    pub(crate) fn main_units(&self, start: u32, end: u32) -> Result<Vec<u16>> {
        read_units(&self.word, &self.pieces, start, end)
    }

    /// Insert an object field at `insert.at`, its preview at the end of the
    /// Data stream and its storage in `ObjectPool`, shifting every modeled
    /// CP table as a text insertion does.
    pub(crate) fn insert_object_field(&mut self, insert: ObjectFieldInsert<'_>) -> Result<()> {
        let mut candidate = self.clone();
        let data_offset = u32::try_from(candidate.data.len())
            .map_err(|_error| corrupted("Data stream exceeds u32"))?;
        candidate.data.extend_from_slice(insert.picture_block);
        candidate.data_changed = true;

        let special = with_sprm(insert.base, SPRM_C_F_SPEC, &[1]);
        let mut separator = with_sprm(&[], SPRM_C_PIC_LOCATION, &insert.storage_id.to_le_bytes());
        separator.extend_from_slice(&with_sprm(&[], SPRM_C_F_OLE2, &[1]));
        separator.extend_from_slice(&special);
        separator.extend_from_slice(&with_sprm(&[], SPRM_C_F_OBJ, &[1]));
        let mut picture = picture_sprms(data_offset, insert.lowering)?;
        picture.extend_from_slice(&with_sprm(
            &retain_sprms(insert.base, &[SPRM_C_HPS_POS])?,
            SPRM_C_F_SPEC,
            &[1],
        ));
        let instruction = insert.instruction.encode_utf16().collect::<Vec<_>>();
        let instruction_len = u32::try_from(instruction.len())
            .map_err(|_error| corrupted("field instruction exceeds u32"))?;
        let separator_cp = insert
            .at
            .checked_add(1)
            .and_then(|cp| cp.checked_add(instruction_len))
            .ok_or_else(|| corrupted("field CP overflow"))?;
        let end_cp = separator_cp
            .checked_add(2)
            .ok_or_else(|| corrupted("field CP overflow"))?;
        let runs = vec![
            (vec![0x13], special.clone()),
            (instruction, insert.base.to_vec()),
            (vec![0x14], separator),
            (vec![0x01], picture),
            (vec![0x15], special),
        ];
        let fields = [
            (insert.at, [0x13, FLT_EMBED]),
            (separator_cp, FIELD_SEPARATOR),
            (end_cp, [FIELD_END, insert.end_flags]),
        ];
        candidate.insert_runs_with_fields(insert.at, &runs, &fields)?;
        candidate.add_object_storage(insert.storage_id, insert.compound_file)?;
        // The package changed after the last stream publication; a render
        // retained by that publication no longer shows it.
        candidate.release_retained_render();
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    fn add_object_storage(&mut self, storage_id: u32, compound_file: Vec<u8>) -> Result<()> {
        let name = format!("_{storage_id}");
        let pool = object_pool_path();
        let has_pool = self
            .package
            .child_storage_paths(&[])
            .any(|path| path == pool.as_slice());
        // The storages are not opened targets: the fields own them, and a
        // later removal or insertion addresses them by path.
        if has_pool {
            self.package
                .add_storage_at(&storage_path(storage_id), compound_file)?;
        } else {
            let limits = self.package.limits();
            let wrapped =
                crate::embedded_object::wrap_object_storage(&name, compound_file, limits)?;
            self.package.add_storage_at(&pool, wrapped)?;
        }
        Ok(())
    }

    /// Replace the stream `name` of the object storage `_<storage_id>` and
    /// show `picture_block` as the preview of the picture character at
    /// `picture_cp`. The old preview block is released (see
    /// [`Self::release_picture_block`]).
    pub(crate) fn replace_object_content(
        &mut self,
        storage_id: u32,
        name: &str,
        bytes: Vec<u8>,
        picture_cp: u32,
        picture_block: &[u8],
        lowering: u16,
    ) -> Result<()> {
        let mut path = storage_path(storage_id);
        path.push(name.to_owned());
        if self.package.stream(&path).is_none() {
            return Err(corrupted("object storage stream is missing"));
        }
        let old_offset = self.picture_location_at_cp(picture_cp)?;
        let mut candidate = self.clone();
        candidate.release_picture_block(old_offset)?;
        let offset = u32::try_from(candidate.data.len())
            .map_err(|_error| corrupted("Data stream exceeds u32"))?;
        candidate.data.extend_from_slice(picture_block);
        candidate.data_changed = true;
        let end = picture_cp
            .checked_add(1)
            .ok_or_else(|| corrupted("picture CP overflow"))?;
        // `set_character_sprms` publishes the Data stream with the CHPX.
        candidate.set_character_sprms(
            picture_cp,
            end,
            &[SPRM_C_PIC_LOCATION, SPRM_C_HPS_POS],
            &picture_sprms(offset, lowering)?,
        )?;
        candidate.package.put_stream(&path, bytes)?;
        // The package changed after the last stream publication; a render
        // retained by that publication no longer shows it.
        candidate.release_retained_render();
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Remove the object field `start..end` (its end mark included), its
    /// storage `_<storage_id>` and, when `picture_cp` names its picture
    /// character, its preview block.
    pub(crate) fn remove_object_field(
        &mut self,
        start: u32,
        end: u32,
        storage_id: u32,
        picture_cp: Option<u32>,
    ) -> Result<()> {
        let offset = picture_cp
            .map(|cp| self.picture_location_at_cp(cp))
            .transpose()?;
        // Another object field that shows the same storage keeps it.
        let shared = self.object_references(storage_id)? > 1;
        let mut candidate = self.clone();
        if let Some(offset) = offset {
            candidate.release_picture_block(offset)?;
        }
        candidate.delete_range_with_fields(start, end)?;
        if candidate
            .cp_tables
            .iter()
            .any(|table| table.index == PLCFFLD_MOM && table.records.is_empty())
        {
            // The last field is gone: drop the table instead of keeping an
            // empty one.
            candidate
                .cp_tables
                .retain(|table| table.index != PLCFFLD_MOM);
            put_fib_pair(&mut candidate.word, PLCFFLD_MOM, 0, 0)?;
            candidate.commit()?;
        }
        if !shared {
            candidate
                .package
                .remove_storage_at(&storage_path(storage_id))?;
        }
        // The package changed after the last stream publication; a render
        // retained by that publication no longer shows it.
        candidate.release_retained_render();
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Release the PICF block at Data offset `offset` that only one picture
    /// character shows: drop it when it ends the Data stream, otherwise
    /// overwrite it with zeros so that no other Data offset moves. A block
    /// another character also shows, or one that does not fit the stream,
    /// is kept.
    fn release_picture_block(&mut self, offset: u32) -> Result<()> {
        let start = usize::try_from(offset).map_err(|_error| corrupted("Data offset overflow"))?;
        let Some(length) = self
            .data
            .get(start..start.saturating_add(4))
            .and_then(|bytes| bytes.try_into().ok())
            .map(u32::from_le_bytes)
            .and_then(|length| usize::try_from(length).ok())
        else {
            return Ok(());
        };
        let Some(end) = start.checked_add(length) else {
            return Ok(());
        };
        if length < 4 || end > self.data.len() || self.picture_references(offset)? > 1 {
            return Ok(());
        }
        if end == self.data.len() {
            self.data.truncate(start);
        } else {
            self.data[start..end].fill(0);
        }
        self.data_changed = true;
        Ok(())
    }

    /// The PICF shape id of the Data block at `offset` (its
    /// `OfficeArtFSP` `spid`), when the block holds one.
    pub(crate) fn picture_shape_id(&self, offset: u32) -> Option<u32> {
        let start = usize::try_from(offset).ok()?;
        let header = self.data.get(start + 4..start + 6)?;
        if usize::from(u16::from_le_bytes([header[0], header[1]])) != PICF_HEADER {
            return None;
        }
        // OfficeArtSpContainer header (8 bytes), OfficeArtFSP header (8
        // bytes), then the spid.
        let at = start + PICF_HEADER;
        let records = self.data.get(at..at + 20)?;
        if records[2..4] != [0x04, 0xF0] || records[10..12] != [0x0A, 0xF0] {
            return None;
        }
        Some(u32::from_le_bytes([
            records[16],
            records[17],
            records[18],
            records[19],
        ]))
    }

    /// How many characters show the Data block at `offset`: characters
    /// whose `sprmCPicLocation` is `offset` and that are not OLE object
    /// separators, whose location is a storage number instead.
    fn picture_references(&self, offset: u32) -> Result<usize> {
        self.characters_located_at(offset, false)
    }

    /// How many object field separators name the storage `_<storage_id>`.
    fn object_references(&self, storage_id: u32) -> Result<usize> {
        self.characters_located_at(storage_id, true)
    }

    /// How many characters of the text have `sprmCPicLocation` equal to
    /// `location`, among OLE object separators (`ole`) or the other
    /// characters. Only characters a text piece still reaches count: a
    /// formatting run also describes text that an earlier edit deleted.
    fn characters_located_at(&self, location: u32, ole: bool) -> Result<usize> {
        let mut count = 0usize;
        for FcRun { start, end, grpprl } in &self.chpx {
            let sprms = strict_sprms(grpprl)?;
            let is_ole = sprms
                .iter()
                .any(|sprm| sprm.opcode == SPRM_C_F_OLE2 && sprm.operand_bytes() != [0]);
            let located = sprms
                .iter()
                .rev()
                .find(|sprm| sprm.opcode == SPRM_C_PIC_LOCATION)
                .is_some_and(|sprm| sprm.operand_bytes() == location.to_le_bytes());
            if !located || is_ole != ole {
                continue;
            }
            for piece in &self.pieces {
                let width = if piece.unicode { 2 } else { 1 };
                let Some(piece_end) = piece
                    .end
                    .checked_sub(piece.start)
                    .and_then(|length| length.checked_mul(width))
                    .and_then(|length| piece.fc.checked_add(length))
                else {
                    continue;
                };
                let overlap = piece_end.min(*end).saturating_sub(piece.fc.max(*start));
                let characters = usize::try_from(overlap / width).unwrap_or(usize::MAX);
                count = count.saturating_add(characters);
            }
        }
        Ok(count)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::equation::{EquationKind, MtefEquationObject, equation_object_compound_file};
    use crate::tracked_revision::Limits;
    use crate::writer::images::{FIRST_PICTURE_SHAPE_ID, Picture, write_picture_block};
    use std::io::Cursor;

    /// A 1x1 transparent PNG.
    const PNG: &[u8] = &[
        0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, b'I', b'H', b'D',
        b'R', 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, b'I', b'D', b'A', b'T', 0x08, 0xD7, 0x63, 0x60,
        0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x05, 0x00, 0x01, 0xE2, 0x26, 0x05, 0x9B, 0x00, 0x00,
        0x00, 0x00, b'I', b'E', b'N', b'D', 0xAE, 0x42, 0x60, 0x82,
    ];

    fn editor() -> RevisionEditor {
        let mut builder = crate::builder::Builder::new();
        builder.add_paragraph("one two three").unwrap();
        let mut output = Cursor::new(Vec::new());
        builder.build().unwrap().write_to(&mut output).unwrap();
        RevisionEditor::open(output.into_inner(), Limits::default()).unwrap()
    }

    fn block() -> Vec<u8> {
        let mut block = Vec::new();
        write_picture_block(
            &Picture::new(PNG.to_vec()).unwrap(),
            FIRST_PICTURE_SHAPE_ID,
            &mut block,
            litchi_odraw::image::Limits::default(),
        )
        .unwrap();
        block
    }

    fn insert(editor: &mut RevisionEditor, at: u32, storage_id: u32) {
        let equation =
            MtefEquationObject::from_mtef_payload(vec![3, 1, 1, 3, 0, 2, 0x86, 0x3D, 0, 0])
                .unwrap();
        let picture_block = block();
        editor
            .insert_object_field(ObjectFieldInsert {
                at,
                instruction: " EMBED Equation.3  ",
                base: &[],
                storage_id,
                end_flags: 0x80,
                picture_block: &picture_block,
                lowering: 4,
                compound_file: equation_object_compound_file(
                    EquationKind::EquationEditor3,
                    equation.equation_native(),
                )
                .unwrap(),
            })
            .unwrap();
    }

    /// The object field that begins at `start`: the CP after its end mark,
    /// and its picture character.
    fn field(editor: &RevisionEditor, start: u32) -> (u32, u32) {
        let units = editor.main_units(start, editor.main_ccp).unwrap();
        let picture = units.iter().position(|unit| *unit == 0x01).unwrap();
        let end = units.iter().position(|unit| *unit == 0x15).unwrap();
        let cp = |index: usize| start + u32::try_from(index).unwrap();
        (cp(end) + 1, cp(picture))
    }

    #[test]
    fn a_removed_field_no_longer_holds_its_storage_or_preview() {
        let mut editor = editor();
        let data = editor.data.len();
        for _ in 0..2 {
            // The same storage number and Data offset are reused each time.
            insert(&mut editor, 4, 1);
            let (end, picture) = field(&editor, 4);
            editor
                .remove_object_field(4, end, 1, Some(picture))
                .unwrap();
            assert!(editor.object_storage_ids().is_empty());
            assert_eq!(editor.data.len(), data);
        }
    }

    #[test]
    fn a_storage_two_fields_name_is_removed_with_the_last() {
        let mut editor = editor();
        insert(&mut editor, 4, 1);
        let (end, _) = field(&editor, 4);
        // A second field naming the same storage, after the first.
        let units = editor.main_units(4, end).unwrap();
        let separator = units.iter().position(|unit| *unit == 0x14).unwrap();
        let mut runs = Vec::new();
        for (index, unit) in units.iter().enumerate() {
            let at = 4 + u32::try_from(index).unwrap();
            runs.push((vec![*unit], editor.character_grpprl_at(at).unwrap()));
        }
        let length = u32::try_from(units.len()).unwrap();
        let second = end;
        let separator_cp = second + u32::try_from(separator).unwrap();
        editor
            .insert_runs_with_fields(
                second,
                &runs,
                &[
                    (second, [0x13, FLT_EMBED]),
                    (separator_cp, FIELD_SEPARATOR),
                    (second + length - 1, [FIELD_END, 0x80]),
                ],
            )
            .unwrap();
        assert_eq!(editor.object_references(1).unwrap(), 2);

        let (second_end, second_picture) = field(&editor, second);
        editor
            .remove_object_field(second, second_end, 1, Some(second_picture))
            .unwrap();
        assert_eq!(
            editor.object_storage_ids(),
            vec![1],
            "the first field keeps the storage"
        );
        let (end, picture) = field(&editor, 4);
        editor
            .remove_object_field(4, end, 1, Some(picture))
            .unwrap();
        assert!(editor.object_storage_ids().is_empty());
    }
}
