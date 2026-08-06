//! Candidate-snapshot mutations for managed embedded-object fields.

use super::super::codec::{align2, align4, corrupted, managed_objects, put_u32, validate_options};
use super::super::model::{Editor, FieldMarker, Reference, WriteOptions};
use super::super::storage::object_target;
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

use super::inventory::object_for_reference;

impl Editor {
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// # Errors
    ///
    /// Returns an error when a managed field does not resolve to its owning
    /// `ObjectPool` storage.
    pub fn objects(&self) -> Result<Vec<Reference>> {
        managed_objects(&self.word, &self.pieces, &self.fields)?
            .into_iter()
            .map(|mut reference| {
                let object =
                    object_for_reference(self.package.objects(), &reference).ok_or_else(|| {
                        corrupted(format!(
                            "managed embedded-object storage {:?} is missing",
                            reference.storage_name
                        ))
                    })?;
                // Field instructions carry the numeric storage identity. Use
                // the resolved directory spelling for subsequent CFB edits,
                // so producer forms such as `_00042` remain addressable.
                reference.storage_name = object.key().to_owned();
                Ok(reference)
            })
            .collect()
    }

    /// Adds an object at the main-story boundary. No existing logical range
    /// is shifted outside the main story; subsequent story piece CPs shift
    /// consistently.
    pub fn add(&mut self, options: WriteOptions) -> Result<Reference> {
        validate_options(&options, self)?;
        let mut candidate = self.clone();
        let target = object_target(options.storage_id)?;
        let storage_name = target.key().to_owned();
        if candidate.package.targets().get(target.key()).is_some() {
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
        put_u32(&mut candidate.word, 76, candidate.main_ccp)?;
        put_u32(&mut candidate.word, 28, text_end_u32)?;
        let cb_mac = u32::try_from(candidate.word.len())
            .map_err(|_| corrupted("WordDocument size exceeds u32"))?;
        put_u32(&mut candidate.word, 64, cb_mac)?;
        candidate
            .package
            .put_stream(&candidate.word_path, candidate.word.clone())
            .map_err(PackageError::from)?;
        candidate
            .package
            .put_stream(&candidate.table_path, candidate.table.clone())
            .map_err(PackageError::from)?;
        if candidate.package.stream(&candidate.data_path).is_some() {
            candidate
                .package
                .put_stream(&candidate.data_path, candidate.data.clone())
                .map_err(PackageError::from)?;
        } else {
            candidate
                .package
                .add_stream(candidate.data_path.clone(), candidate.data.clone())
                .map_err(PackageError::from)?;
        }
        let limits = candidate.limits;
        candidate.add_object_storage(target, options.compound_file, limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(Reference {
            storage_id: options.storage_id,
            storage_name,
            start_cp,
            separator_cp,
            end_cp,
            data_offset,
        })
    }

    pub fn remove(&mut self, storage_id: u32) -> Result<Reference> {
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
        put_u32(&mut candidate.word, 76, candidate.main_ccp)?;
        candidate
            .package
            .put_stream(&candidate.word_path, candidate.word.clone())
            .map_err(PackageError::from)?;
        candidate
            .package
            .put_stream(&candidate.table_path, candidate.table.clone())
            .map_err(PackageError::from)?;
        let target = candidate
            .package
            .targets()
            .get(&object.storage_name)
            .cloned()
            .ok_or_else(|| corrupted("ObjectPool storage target is missing"))?;
        candidate
            .package
            .remove_storage(target.key())
            .map_err(PackageError::from)?;
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
        let first_cp = objects
            .first()
            .ok_or_else(|| corrupted("managed object list is empty"))?
            .start_cp;
        let suffix_end = objects
            .last()
            .and_then(|object| object.end_cp.checked_add(1))
            .ok_or_else(|| corrupted("managed object CP overflow"))?;
        if suffix_end != self.main_ccp
            || objects
                .windows(2)
                .any(|pair| pair[0].end_cp.checked_add(1) != Some(pair[1].start_cp))
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
            .put_stream(&candidate.word_path, candidate.word.clone())
            .map_err(PackageError::from)?;
        candidate
            .package
            .put_stream(&candidate.table_path, candidate.table.clone())
            .map_err(PackageError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }
}

fn insert_piece_at_main_end(
    pieces: &mut Vec<super::super::model::RawPiece>,
    cp: u32,
    len: u32,
    fc: u32,
) -> Result<()> {
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
    replacement.push(super::super::model::RawPiece {
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
