//! MathType and Equation Editor objects on the primary body transaction:
//! insert one inline in a main-story paragraph, replace one's content, or
//! remove one with its storage and preview.
//!
//! An equation is the `EMBED` field of an OLE object ([MS-DOC] 2.1.4 and
//! 2.6) whose `ObjectPool/_<id>` storage has an `Equation Native` stream.
//! A new one is written as Word writes it: the field begin, the
//! ` EMBED Equation.DSMT4 ` instruction, the separator naming the storage
//! through `sprmCPicLocation` with `fOle2`, `fObj` and `fSpec`, the picture
//! character naming its PICF block in the Data stream, and the field end.

use litchi_core::Position;

use super::properties::{EquationChange, MarkChange};
use super::{
    Change, Edit, Error, Refusal, Result, cp_overflow, keyword_of, source_paragraphs,
    text_revision_intersects,
};
use crate::equation::{EquationContent, EquationKind, equation_object_compound_file};
use crate::package::Error as PackageError;
use crate::sprm_operations::{
    SPRM_C_F_OBJ, SPRM_C_F_OLE2, SPRM_C_F_RMARK, SPRM_C_F_RMARK_DEL, SPRM_C_F_SPEC,
    SPRM_C_PIC_LOCATION,
};
use crate::tracked_revision::ObjectFieldInsert;
use crate::writer::images::{FIRST_PICTURE_SHAPE_ID, write_picture_block};

/// `grffld` of an object field end: `fHasSep`.
const FIELD_END_HAS_SEPARATOR: u8 = 0x80;
/// `grffld` `fResultDirty`, which Word sets on a MathType field end.
const FIELD_END_RESULT_DIRTY: u8 = 0x04;

/// The stream of an equation object that holds its MTEF.
const EQUATION_NATIVE: &str = "Equation Native";
/// The length of the `EQNOLEFILEHDR` before the MTEF.
const EQUATION_NATIVE_HEADER: usize = 28;

/// One equation object field of the main story.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct EquationField {
    /// The CP of the field begin.
    pub(crate) start_cp: u32,
    /// The CP of the field end.
    pub(crate) end_cp: u32,
    /// The `ObjectPool` storage number.
    pub(crate) storage_id: u32,
    /// The CP of the picture character that shows the preview.
    pub(crate) picture_cp: Option<u32>,
    /// The MTEF version of the object's `Equation Native` stream.
    pub(crate) mtef_version: u8,
}

fn invalid(error: PackageError) -> Error {
    match error {
        PackageError::Corrupted(message)
            if message.contains("anchored range") || message.contains("paragraph") =>
        {
            Error::Refused(Refusal::AnchoredContent)
        },
        other => Error::Invalid(other),
    }
}

impl Edit {
    /// The equation object field that begins at `start_cp`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] when no field begins there or
    /// the field is not an `EMBED` field of an object with an
    /// `Equation Native` stream.
    pub(crate) fn equation_field(&self, start_cp: u32) -> Result<EquationField> {
        let not_found = || Error::Refused(Refusal::EquationNotFound);
        let (begin, separator, end) = self
            .main_fields()?
            .into_iter()
            .find(|(begin, _, _)| *begin == start_cp)
            .ok_or_else(not_found)?;
        let separator = separator.ok_or_else(not_found)?;
        let after_begin = begin.checked_add(1).ok_or_else(cp_overflow)?;
        let instruction = self
            .editor
            .main_units(after_begin, separator)
            .map_err(Error::Invalid)?;
        if keyword_of(&String::from_utf16_lossy(&instruction)).as_deref() != Some("EMBED") {
            return Err(not_found());
        }
        let storage_id = self
            .editor
            .picture_location_at_cp(separator)
            .map_err(|_error| not_found())?;
        let mtef_version = self
            .editor
            .object_stream(storage_id, EQUATION_NATIVE)
            .and_then(|stream| stream.get(EQUATION_NATIVE_HEADER).copied())
            .ok_or_else(not_found)?;
        let result_start = separator.checked_add(1).ok_or_else(cp_overflow)?;
        let result = self
            .editor
            .main_units(result_start, end)
            .map_err(Error::Invalid)?;
        let picture_cp = result
            .iter()
            .position(|unit| *unit == 0x01)
            .and_then(|index| u32::try_from(index).ok())
            .map(|index| result_start + index);
        Ok(EquationField {
            start_cp: begin,
            end_cp: end,
            storage_id,
            picture_cp,
            mtef_version,
        })
    }

    /// Insert an equation object of `kind` holding `content` into the
    /// ordinary main-story paragraph at `position`, before UTF-16 offset
    /// `offset` of its text, and return its storage number. The field takes
    /// the formatting of the text it is inserted before.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidProperty`] when `content` holds MTEF of
    /// another kind than `kind`, and the refusals of a hyperlink insertion
    /// for the position.
    pub(crate) fn insert_equation(
        &mut self,
        position: Position,
        offset: u32,
        kind: EquationKind,
        content: &EquationContent,
    ) -> Result<u32> {
        if EquationKind::for_mtef_version(content.mtef_version()) != kind {
            return Err(Error::Refused(Refusal::InvalidProperty));
        }
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        let length =
            u32::try_from(paragraph.text.encode_utf16().count()).map_err(|_error| cp_overflow())?;
        if offset > length {
            return Err(Error::Refused(Refusal::TargetNotFound));
        }
        let at = paragraph
            .start_cp
            .checked_add(offset)
            .ok_or_else(cp_overflow)?;
        if self
            .main_fields()?
            .iter()
            .any(|(begin, _, end)| *begin < at && at <= *end)
        {
            return Err(Error::Refused(Refusal::ComplexField));
        }
        let next = at.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, at, next)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.length_change_admitted(at, at, 1)?;
        let base = self
            .editor
            .character_grpprl_at(at)
            .or_else(|_error| match at.checked_sub(1) {
                Some(previous) => self.editor.character_grpprl_at(previous),
                None => Ok(Vec::new()),
            })
            .unwrap_or_default();
        let base = super::properties::grpprl_without(
            &base,
            &[
                SPRM_C_F_SPEC,
                SPRM_C_F_OBJ,
                SPRM_C_F_OLE2,
                SPRM_C_PIC_LOCATION,
                SPRM_C_F_RMARK,
                SPRM_C_F_RMARK_DEL,
            ],
        )?;
        let storage_id = self
            .editor
            .object_storage_ids()
            .into_iter()
            .max()
            .map_or(Some(1), |last| last.checked_add(1))
            .filter(|id| i32::try_from(*id).is_ok())
            .ok_or(Error::Refused(Refusal::ResourceCollision {
                storage_id: u32::MAX,
            }))?;
        // Word writes the ProgID followed by two spaces.
        let instruction = format!(" EMBED {}  ", kind.prog_id());
        let units = instruction.encode_utf16().count() + 4;
        self.ensure_replacement_capacity(units)?;
        self.ensure_operation_capacity()?;
        // Word numbers the preview shapes of a document's objects 0x401,
        // 0x402 and on.
        let existing = u32::try_from(self.editor.object_storage_ids().len()).unwrap_or(u32::MAX);
        let shape_id = FIRST_PICTURE_SHAPE_ID.saturating_add(existing);
        let picture_block = picture_block(content, shape_id)?;
        let end_flags = match kind {
            EquationKind::MathType => FIELD_END_HAS_SEPARATOR | FIELD_END_RESULT_DIRTY,
            EquationKind::EquationEditor3 => FIELD_END_HAS_SEPARATOR,
        };
        let compound_file =
            equation_object_compound_file(kind, content.equation().equation_native())
                .map_err(Error::Invalid)?;
        self.editor
            .insert_object_field(ObjectFieldInsert {
                at,
                instruction: &instruction,
                base: &base,
                storage_id,
                end_flags,
                picture_block: &picture_block,
                lowering: content.lowering(),
                compound_file,
            })
            .map_err(Error::Invalid)?;
        self.replacement_units = self.replacement_units.saturating_add(units);
        self.changes.push(Change::Mark(MarkChange::Equation {
            start_cp: at,
            change: EquationChange::Inserted {
                position,
                offset,
                kind,
                content: Box::new(content.clone()),
            },
        }));
        Ok(storage_id)
    }

    /// Replace the content of the equation object whose field begins at
    /// `start_cp`: its `Equation Native` stream and its preview picture.
    /// The field, the object's other streams and its class stay.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] when no equation object field
    /// begins there, [`Refusal::InvalidProperty`] when `content` holds MTEF
    /// of another kind than the object, and [`Refusal::StructuralContent`]
    /// when the field shows no picture character.
    pub(crate) fn set_equation(&mut self, start_cp: u32, content: &EquationContent) -> Result<()> {
        let field = self.equation_field(start_cp)?;
        if EquationKind::for_mtef_version(field.mtef_version)
            != EquationKind::for_mtef_version(content.mtef_version())
        {
            return Err(Error::Refused(Refusal::InvalidProperty));
        }
        let picture_cp = field
            .picture_cp
            .ok_or(Error::Refused(Refusal::StructuralContent))?;
        let after = field.end_cp.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, start_cp, after)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.ensure_operation_capacity()?;
        // The new preview keeps the old one's shape id.
        let shape_id = self
            .editor
            .picture_location_at_cp(picture_cp)
            .ok()
            .and_then(|offset| self.editor.picture_shape_id(offset))
            .unwrap_or(FIRST_PICTURE_SHAPE_ID);
        let picture_block = picture_block(content, shape_id)?;
        self.editor
            .replace_object_content(
                field.storage_id,
                EQUATION_NATIVE,
                content.equation().equation_native().to_vec(),
                picture_cp,
                &picture_block,
                content.lowering(),
            )
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::Equation {
            start_cp,
            change: EquationChange::Updated {
                content: Box::new(content.clone()),
            },
        }));
        Ok(())
    }

    /// Remove the equation object whose field begins at `start_cp`: the
    /// whole field, its `ObjectPool` storage (kept when another field shows
    /// it) and its preview picture block (dropped when it ends the Data
    /// stream, otherwise overwritten with zeros so that no other picture
    /// moves, and kept when another character shows it). A field table left
    /// without fields is dropped.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] when no equation object field
    /// begins there, and a typed refusal when the field holds tracked text
    /// or another anchored range.
    pub(crate) fn remove_equation(&mut self, start_cp: u32) -> Result<()> {
        let field = self.equation_field(start_cp)?;
        let after = field.end_cp.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, start_cp, after)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.length_change_admitted(start_cp, after, 0)?;
        self.ensure_operation_capacity()?;
        self.editor
            .remove_object_field(start_cp, after, field.storage_id, field.picture_cp)
            .map_err(invalid)?;
        self.changes.push(Change::Mark(MarkChange::Equation {
            start_cp,
            change: EquationChange::Removed,
        }));
        Ok(())
    }
}

/// The PICF block of `content`'s preview, as Word stores an inline picture,
/// its shape numbered `shape_id`.
fn picture_block(content: &EquationContent, shape_id: u32) -> Result<Vec<u8>> {
    let mut block = Vec::new();
    write_picture_block(
        content.preview(),
        shape_id,
        &mut block,
        litchi_odraw::image::Limits::default(),
    )
    .map_err(|error| Error::Invalid(PackageError::InvalidFormat(error.to_string())))?;
    Ok(block)
}
