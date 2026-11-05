//! Immutable legacy-note edits of a [`Package`] (unified-ops record 0320).
//!
//! [`Package::edit_comments`] starts an [`Edit`] over an immutable snapshot.
//! [`Edit::set`] adds or replaces the note of one cell, [`Edit::remove`]
//! deletes it, [`Edit::replace`] writes a worksheet's whole typed note graph
//! (with [`Edit::set_comment`], [`Edit::set_author`],
//! [`Edit::rename_author`] and [`Edit::clear`] over it), and
//! [`Edit::commit`] returns a [`Commit`] with the new
//! package, the reversible [`Patch`] of the parts it changed and the
//! [`Diagnostics`]; [`Package::apply_comments`] replays a patch onto an
//! equal package.
//!
//! Each change is a splice: the `<comment>` elements, authors and VML
//! shapes of the other notes keep their bytes, rich text included. A new
//! note gets the VML note shape Excel needs to show it; the worksheet's
//! comments part, VML drawing and `<legacyDrawing>` element are created with
//! the first note and removed with the last one, when the VML drawing holds
//! nothing else. A VML drawing holding other shapes (form controls) keeps
//! them. An exact no-op keeps every byte, and a changed edit of a signed
//! workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//!
//! # fn main() -> Result<(), litchi_xlsx::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_comments()?;
//! edit.set(0, "B2", "Ann", "Check this")?;
//! let after = edit.commit()?.into_package();
//! let comments = after.edit_comments()?.comments(0)?.unwrap();
//! assert_eq!(comments.get("B2").unwrap().text, "Check this");
//! # Ok(())
//! # }
//! ```

use std::fmt::Write as _;
use std::ops::Range;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI};
use litchi_spreadsheet_drawing::vml::{
    Limits as VmlLimits, SHAPE_TYPE, contains, new_drawing, next_shape_id, note_shape, rfind,
    shapes,
};
use quick_xml::events::Event;

use super::model::{Comment, Comments, RunColor, TextRun};
use super::package::{
    COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, load_from_worksheet, remove_from_worksheet,
    replace_on_worksheet,
};
use super::validation;
use crate::Package;
use crate::error::{Result, invalid};
use crate::part_edit::{Staged, WorksheetLayout, family};

family!("comment");

/// The most notes one worksheet may carry after an edit.
pub const MAX_COMMENTS: usize = super::MAX_COMMENTS;
/// The longest note text or author, in bytes.
pub const MAX_TEXT_BYTES: usize = super::MAX_TEXT_BYTES;

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Notes of one worksheet ([`MAX_COMMENTS`]).
    Comments,
    /// Bytes of one note text or author ([`MAX_TEXT_BYTES`]).
    TextBytes,
}

/// A staged legacy-note edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a legacy-note edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_comments(&self) -> Result<Edit> {
        self.edit_comments_with_options(crate::EditOptions::default())
    }

    /// Start a legacy-note edit with an explicit retained-patch budget.
    pub fn edit_comments_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_comments", options)?,
        })
    }

    /// Apply a legacy-note patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_comments(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_comments")?,
            patch.0.clone(),
        )))
    }
}

/// The checked A1 reference and zero-based (row, column) of a cell.
fn cell(value: &str) -> Result<(String, u32, u32)> {
    let cell = litchi_sheet::Cell::from_a1(value)?;
    Ok((cell.a1(), cell.row().get(), cell.column().get()))
}

impl Edit {
    /// The staged notes of the worksheet at `sheet`; `None` without a
    /// comments part.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, or a malformed
    /// comments part.
    pub fn comments(&self, sheet: usize) -> Result<Option<Comments>> {
        let worksheet = self.staged.worksheet(sheet)?;
        Ok(load_from_worksheet(self.staged.opc(), &worksheet)?.map(|part| part.comments))
    }

    /// Add or replace the note of `cell` (A1). Returns whether anything
    /// changed; a replaced note keeps its other attributes and its VML shape.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid cell,
    /// author or text, too many notes, a VML drawing this editor cannot read,
    /// or a changed edit of a signed workbook.
    pub fn set(&mut self, sheet: usize, cell_ref: &str, author: &str, text: &str) -> Result<bool> {
        self.write(sheet, cell_ref, author, Body::Text(text), false)
    }

    /// Add or replace the note of `cell` (A1) with rich-text `runs`: each
    /// run is written as an `<r>` with the run properties its
    /// [`RunFont`](super::RunFont) sets. Returns whether anything changed;
    /// a note whose author and runs are equal is left as it is.
    ///
    /// # Errors
    ///
    /// As [`Edit::set`], and for no runs or an invalid run property.
    pub fn set_runs(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        runs: &[TextRun],
    ) -> Result<bool> {
        self.write(sheet, cell_ref, author, Body::Runs(runs), false)
    }

    /// [`Edit::add`] with rich-text `runs`, as [`Edit::set_runs`] writes
    /// them.
    ///
    /// # Errors
    ///
    /// As [`Edit::set_runs`].
    pub fn add_runs(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        runs: &[TextRun],
    ) -> Result<Option<usize>> {
        self.insert(sheet, cell_ref, author, Body::Runs(runs))
    }

    /// Add a note to `cell` (A1) when the cell has none. Returns the new
    /// note's position among the worksheet's notes in row-major order of
    /// their cells; `None` leaves a cell that already has a note unchanged.
    ///
    /// # Errors
    ///
    /// As [`Edit::set`].
    pub fn add(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        text: &str,
    ) -> Result<Option<usize>> {
        self.insert(sheet, cell_ref, author, Body::Text(text))
    }

    fn insert(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        body: Body<'_>,
    ) -> Result<Option<usize>> {
        let (_, row, column) = cell(cell_ref)?;
        let worksheet = self.staged.worksheet(sheet)?;
        let before = load_from_worksheet(self.staged.opc(), &worksheet)?;
        let position = before.as_ref().map_or(0, |part| {
            part.comments
                .iter()
                .filter(|(other, _)| {
                    litchi_sheet::Cell::from_a1(other).is_ok_and(|other| {
                        (other.row().get(), other.column().get()) < (row, column)
                    })
                })
                .count()
        });
        Ok(self
            .write_loaded(sheet, cell_ref, author, body, true, before)?
            .then_some(position))
    }

    fn write(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        body: Body<'_>,
        insert_only: bool,
    ) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let existing = load_from_worksheet(self.staged.opc(), &worksheet)?;
        self.write_loaded(sheet, cell_ref, author, body, insert_only, existing)
    }

    fn write_loaded(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        body: Body<'_>,
        insert_only: bool,
        existing: Option<super::model::Part>,
    ) -> Result<bool> {
        let (reference, row, column) = cell(cell_ref)?;
        let joined;
        let text = match body {
            Body::Text(text) => text,
            Body::Runs(runs) => {
                check_runs(runs)?;
                joined = runs.iter().map(|run| run.text.as_str()).collect::<String>();
                joined.as_str()
            },
        };
        for (value, what) in [(author, "author"), (text, "text")] {
            if value.len() > MAX_TEXT_BYTES {
                return Err(Error::limit(
                    LimitKind::TextBytes,
                    value.len(),
                    MAX_TEXT_BYTES,
                ));
            }
            if value.contains('\0') {
                return Err(invalid(format!("a note {what} must hold no NUL")));
            }
        }
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(part) = &existing {
            if let Some(comment) = part.comments.get(&reference) {
                let same = match body {
                    Body::Text(_) => comment.text == text,
                    Body::Runs(runs) => comment.runs == runs,
                };
                if insert_only || (comment.author == author && same) {
                    return Ok(false);
                }
            } else if part.comments.len() >= MAX_COMMENTS {
                return Err(Error::limit(
                    LimitKind::Comments,
                    part.comments.len() + 1,
                    MAX_COMMENTS,
                ));
            }
        }
        let added = existing
            .as_ref()
            .is_none_or(|part| part.comments.get(&reference).is_none());
        match existing {
            Some(part) => {
                let name = PackURI::new(part.part_name.as_str())
                    .map_err(|error| invalid(error.to_string()))?;
                let xml = splice_comment(self.staged.blob(&name)?, &reference, author, body)?;
                self.staged.replace_blob(&name, xml)?;
            },
            None => {
                let name = self.staged.free_name("/xl/comments", ".xml")?;
                let xml = new_comments_part(&reference, author, body);
                let opc = self.staged.opc_mut()?;
                opc.add_part(Box::new(BlobPart::new(
                    name.clone(),
                    COMMENTS_CONTENT_TYPE.to_owned(),
                    xml.into_bytes(),
                )));
                opc.get_part_mut(&worksheet)?.relate_to(
                    &name.relative_ref(worksheet.base_uri()),
                    COMMENTS_RELATIONSHIP_TYPE,
                );
                self.staged.touch(&name);
                self.staged.touch(&worksheet);
            },
        }
        if added {
            self.add_note_shape(sheet, &worksheet, row, column)?;
        }
        if load_from_worksheet(self.staged.opc(), &worksheet)?
            .is_none_or(|part| part.comments.get(&reference).is_none_or(|c| c.text != text))
        {
            return Err(invalid("the note did not read back"));
        }
        Ok(true)
    }

    /// Remove the note of `cell` (A1). Returns whether a note was removed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// cell, a VML drawing this editor cannot read, or a signed workbook.
    pub fn remove(&mut self, sheet: usize, cell_ref: &str) -> Result<bool> {
        let (reference, row, column) = cell(cell_ref)?;
        let worksheet = self.staged.worksheet(sheet)?;
        let Some(part) = load_from_worksheet(self.staged.opc(), &worksheet)? else {
            return Ok(false);
        };
        if part.comments.get(&reference).is_none() {
            return Ok(false);
        }
        let name =
            PackURI::new(part.part_name.as_str()).map_err(|error| invalid(error.to_string()))?;
        if part.comments.len() == 1 {
            let opc = self.staged.opc_mut()?;
            opc.get_part_mut(&worksheet)?
                .rels_mut()
                .remove(&part.relationship_id);
            opc.remove_part(&name);
            self.staged.touch(&name);
            self.staged.touch(&worksheet);
        } else {
            let xml = remove_comment(self.staged.blob(&name)?, &reference)?;
            self.staged.replace_blob(&name, xml)?;
        }
        self.remove_note_shape(&worksheet, row, column)?;
        Ok(true)
    }

    /// Replace the complete legacy-note graph of the worksheet at `sheet`;
    /// `None` removes the comments part and its worksheet relationship.
    /// Returns whether anything changed.
    ///
    /// The graph is written whole through the comments codec: inert VML
    /// shape IDs are kept as note metadata, and VML drawing parts are left as
    /// they are (no note shape is added or removed).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// graph, or a changed edit of a signed workbook.
    pub fn replace(&mut self, sheet: usize, value: Option<Comments>) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(value) = &value {
            validation::comments(value)?;
        }
        let current = load_from_worksheet(self.staged.opc(), &worksheet)?.map(|part| part.comments);
        if current == value {
            return Ok(false);
        }
        self.staged.update(|candidate| {
            match &value {
                Some(value) => {
                    replace_on_worksheet(candidate, &worksheet, value)?;
                },
                None => {
                    remove_from_worksheet(candidate, &worksheet)?;
                },
            }
            super::package::validate_graph(candidate)?;
            let published = load_from_worksheet(candidate, &worksheet)?.map(|part| part.comments);
            if published != value {
                return Err(invalid(
                    "classic comments package publication changed the staged graph",
                ));
            }
            Ok(())
        })?;
        Ok(true)
    }

    /// Add or replace a fully typed note entry of the worksheet at `sheet`,
    /// written through [`Self::replace`]. The cell reference and author
    /// relationship are checked by the comments validator; shape IDs remain
    /// opaque and inert.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`], and for an invalid cell, author or text.
    pub fn set_comment(&mut self, sheet: usize, value: Comment) -> Result<bool> {
        litchi_sheet::Cell::from_a1(&value.cell_ref)?;
        validation::text(&value.author, "author")?;
        validation::text(&value.text, "text")?;
        let mut comments = self.comments(sheet)?.unwrap_or_default();
        if comments.comments.get(&value.cell_ref) == Some(&value) {
            return Ok(false);
        }
        comments.comments.insert(value.cell_ref.clone(), value);
        self.replace(sheet, Some(comments))
    }

    /// Rename one checked author-table entry of the worksheet at `sheet` and
    /// update every note referencing it, written through [`Self::replace`].
    ///
    /// # Errors
    ///
    /// As [`Self::replace`], and for an absent comments part or author ID.
    pub fn set_author(
        &mut self,
        sheet: usize,
        author_id: u32,
        author: impl Into<String>,
    ) -> Result<bool> {
        let author = author.into();
        validation::text(&author, "author")?;
        let mut comments = self
            .comments(sheet)?
            .ok_or_else(|| invalid("cannot edit an author on an absent comments part"))?;
        let index = usize::try_from(author_id)
            .ok()
            .filter(|index| *index < comments.authors.len())
            .ok_or_else(|| invalid(format!("classic comments author ID {author_id} is absent")))?;
        if comments.authors[index] == author {
            return Ok(false);
        }
        comments.authors[index].clone_from(&author);
        for comment in comments.comments.values_mut() {
            if comment.author_id == author_id {
                comment.author.clone_from(&author);
            }
        }
        self.replace(sheet, Some(comments))
    }

    /// Rename a unique author of the worksheet at `sheet` without exposing
    /// its native author-table index. Ambiguous names are refused rather
    /// than guessed.
    ///
    /// # Errors
    ///
    /// As [`Self::set_author`], and for an absent or ambiguous author.
    pub fn rename_author(
        &mut self,
        sheet: usize,
        current: &str,
        replacement: impl Into<String>,
    ) -> Result<bool> {
        validation::text(current, "author")?;
        let comments = self
            .comments(sheet)?
            .ok_or_else(|| invalid("cannot edit an author on an absent comments part"))?;
        let mut found = None;
        for (index, author) in comments.authors.iter().enumerate() {
            if author == current {
                if found.is_some() {
                    return Err(invalid(format!(
                        "classic comments author '{current}' is ambiguous"
                    )));
                }
                found = Some(index);
            }
        }
        let index = found
            .ok_or_else(|| invalid(format!("classic comments author '{current}' is absent")))?;
        let index = u32::try_from(index)
            .map_err(|_source| invalid("classic-comments author index exceeds u32"))?;
        self.set_author(sheet, index, replacement)
    }

    /// Remove every legacy note of the worksheet at `sheet`, written through
    /// [`Self::replace`]: the comments part and its relationship go, VML
    /// drawing parts stay.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`].
    pub fn clear(&mut self, sheet: usize) -> Result<bool> {
        self.replace(sheet, None)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The worksheet's VML drawing: its relationship id and part name.
    fn vml(&self, worksheet: &PackURI) -> Result<Option<(String, PackURI)>> {
        let layout = WorksheetLayout::scan(self.staged.blob(worksheet)?)?;
        let Some(child) = layout.child("legacyDrawing") else {
            return Ok(None);
        };
        let id = child
            .relationship_id
            .clone()
            .ok_or_else(|| invalid("legacyDrawing has no relationship id"))?;
        let part = self.staged.opc().get_part(worksheet)?;
        let relationship = part
            .rels()
            .get(&id)
            .ok_or_else(|| invalid("legacyDrawing relationship is missing"))?;
        if relationship.is_external() {
            return Err(invalid("legacyDrawing relationship cannot be external"));
        }
        Ok(Some((id, relationship.target_partname()?)))
    }

    fn add_note_shape(
        &mut self,
        sheet: usize,
        worksheet: &PackURI,
        row: u32,
        column: u32,
    ) -> Result<()> {
        match self.vml(worksheet)? {
            Some((_, name)) => {
                let xml = self.staged.blob(&name)?;
                let limits = VmlLimits::from_source_limits(self.staged.opc().read_limits());
                let shapes = shapes(xml, &limits)?;
                if shapes.iter().any(|shape| shape.note == Some((row, column))) {
                    return Ok(());
                }
                let id = next_shape_id(xml, &shapes, sheet);
                let close = rfind(xml, b"</xml>")
                    .ok_or_else(|| invalid("the VML drawing has no closing xml element"))?;
                let mut out = xml.to_vec();
                let mut fragment = String::new();
                if !contains(xml, b"_x0000_t202") {
                    fragment.push_str(SHAPE_TYPE);
                }
                fragment.push_str(&note_shape(id, row, column));
                out.splice(close..close, fragment.into_bytes());
                self.staged.replace_blob(&name, out)?;
            },
            None => {
                let name = self.staged.free_name("/xl/drawings/vmlDrawing", ".vml")?;
                let xml = new_drawing(sheet, row, column);
                let layout = WorksheetLayout::scan(self.staged.blob(worksheet)?)?;
                let opc = self.staged.opc_mut()?;
                opc.add_part(Box::new(BlobPart::new(
                    name.clone(),
                    ct::OFC_VML_DRAWING.to_owned(),
                    xml.into_bytes(),
                )));
                let id = opc
                    .get_part_mut(worksheet)?
                    .relate_to(&name.relative_ref(worksheet.base_uri()), rt::VML_DRAWING);
                let ws = layout.insert_relationship_child(
                    opc.get_part(worksheet)?.blob(),
                    "legacyDrawing",
                    &id,
                )?;
                opc.get_part_mut(worksheet)?.set_blob(ws);
                self.staged.touch(&name);
                self.staged.touch(worksheet);
            },
        }
        Ok(())
    }

    fn remove_note_shape(&mut self, worksheet: &PackURI, row: u32, column: u32) -> Result<()> {
        let Some((id, name)) = self.vml(worksheet)? else {
            return Ok(());
        };
        let xml = self.staged.blob(&name)?;
        let limits = VmlLimits::from_source_limits(self.staged.opc().read_limits());
        let shapes = shapes(xml, &limits)?;
        let Some(shape) = shapes
            .iter()
            .find(|shape| shape.note == Some((row, column)))
        else {
            return Ok(());
        };
        if shapes.len() > 1 {
            let mut out = xml.to_vec();
            out.drain(shape.span.clone());
            self.staged.replace_blob(&name, out)?;
            return Ok(());
        }
        // The last shape: drop the drawing, its relationship and the
        // worksheet's `legacyDrawing` element.
        let layout = WorksheetLayout::scan(self.staged.blob(worksheet)?)?;
        let ws = layout
            .remove_child(self.staged.blob(worksheet)?, "legacyDrawing")
            .ok_or_else(|| invalid("legacyDrawing vanished"))?;
        let opc = self.staged.opc_mut()?;
        let part = opc.get_part_mut(worksheet)?;
        part.set_blob(ws);
        part.rels_mut().remove(&id);
        let referenced = opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship.target_partname().ok() == Some(name.clone())
            })
        });
        if !referenced {
            opc.remove_part(&name);
        }
        self.staged.touch(&name);
        self.staged.touch(worksheet);
        Ok(())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        super::package::validate_graph(self.staged.opc())?;
        Ok(Commit::new(self.staged.commit()?))
    }
}

fn escape(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn text_element(prefix: &str, text: &str) -> String {
    format!(
        "<{prefix}text><{prefix}t xml:space=\"preserve\">{}</{prefix}t></{prefix}text>",
        escape(text)
    )
}

/// The text of a note to write: plain, or rich-text runs.
#[derive(Clone, Copy)]
enum Body<'a> {
    Text(&'a str),
    Runs(&'a [TextRun]),
}

impl Body<'_> {
    /// The note's `<text>` element in the comments part's `prefix`.
    fn element(self, prefix: &str) -> String {
        match self {
            Self::Text(text) => text_element(prefix, text),
            Self::Runs(runs) => runs_element(prefix, runs),
        }
    }
}

/// Refuse runs this editor cannot write: none at all, or a property
/// outside its value space.
fn check_runs(runs: &[TextRun]) -> Result<()> {
    if runs.is_empty() {
        return Err(invalid("a rich note needs at least one run"));
    }
    for run in runs {
        let font = &run.font;
        let token = |value: &Option<String>, what: &str| {
            if value
                .as_deref()
                .is_some_and(|v| v.is_empty() || !v.bytes().all(|b| b.is_ascii_alphabetic()))
            {
                return Err(invalid(format!("a note run {what} is a token")));
            }
            Ok(())
        };
        token(&font.underline, "underline")?;
        token(&font.vert_align, "vertical alignment")?;
        if font
            .size
            .is_some_and(|size| !size.is_finite() || !(1.0..=409.0).contains(&size))
        {
            return Err(invalid("a note run size must be 1 to 409 points"));
        }
        if let Some(RunColor::Rgb(rgb)) = &font.color
            && (rgb.len() != 8 || !rgb.bytes().all(|b| b.is_ascii_hexdigit()))
        {
            return Err(invalid("a note run color is eight hexadecimal digits"));
        }
        if font
            .name
            .as_deref()
            .is_some_and(|name| name.is_empty() || name.len() > 31 || name.contains('\0'))
        {
            return Err(invalid("a note run font name must hold 1 to 31 characters"));
        }
    }
    Ok(())
}

/// A note's `<text>` of rich runs. Run properties follow the
/// `CT_RPrElt` order Excel writes: `b`, `i`, `strike`, `u`, `vertAlign`,
/// `sz`, `color`, `rFont`.
fn runs_element(prefix: &str, runs: &[TextRun]) -> String {
    let mut out = format!("<{prefix}text>");
    for run in runs {
        let font = &run.font;
        let mut properties = String::new();
        for (on, tag) in [
            (font.bold, "b"),
            (font.italic, "i"),
            (font.strike, "strike"),
        ] {
            match on {
                Some(true) => {
                    let _ = write!(properties, "<{prefix}{tag}/>");
                },
                Some(false) => {
                    let _ = write!(properties, "<{prefix}{tag} val=\"0\"/>");
                },
                None => {},
            }
        }
        if let Some(underline) = &font.underline {
            let _ = write!(properties, "<{prefix}u val=\"{underline}\"/>");
        }
        if let Some(align) = &font.vert_align {
            let _ = write!(properties, "<{prefix}vertAlign val=\"{align}\"/>");
        }
        if let Some(size) = font.size {
            let _ = write!(properties, "<{prefix}sz val=\"{size}\"/>");
        }
        match &font.color {
            Some(RunColor::Rgb(rgb)) => {
                let _ = write!(
                    properties,
                    "<{prefix}color rgb=\"{}\"/>",
                    rgb.to_ascii_uppercase()
                );
            },
            Some(RunColor::Indexed(index)) => {
                let _ = write!(properties, "<{prefix}color indexed=\"{index}\"/>");
            },
            Some(RunColor::Theme(theme)) => {
                let _ = write!(properties, "<{prefix}color theme=\"{theme}\"/>");
            },
            None => {},
        }
        if let Some(name) = &font.name {
            let _ = write!(properties, "<{prefix}rFont val=\"{}\"/>", escape(name));
        }
        out.push_str(&format!("<{prefix}r>"));
        if !properties.is_empty() {
            let _ = write!(out, "<{prefix}rPr>{properties}</{prefix}rPr>");
        }
        let _ = write!(
            out,
            "<{prefix}t xml:space=\"preserve\">{}</{prefix}t></{prefix}r>",
            escape(&run.text)
        );
    }
    let _ = write!(out, "</{prefix}text>");
    out
}

fn new_comments_part(reference: &str, author: &str, body: Body<'_>) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>{}</author></authors><commentList><comment ref=\"{reference}\" authorId=\"0\">{}</comment></commentList></comments>",
        escape(author),
        body.element("")
    )
}

/// The layout of a comments part.
struct Layout {
    prefix: String,
    authors: Vec<String>,
    /// The offset of `</authors>`, or of an empty `<authors/>` span.
    authors_close: Option<usize>,
    authors_empty: Option<Range<usize>>,
    list_close: Option<usize>,
    list_empty: Option<Range<usize>>,
    /// Each `comment`: its ref, its span and its start tag's span.
    comments: Vec<(String, Range<usize>, Range<usize>)>,
}

fn scan(xml: &[u8]) -> Result<Layout> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut layout = Layout {
        prefix: String::new(),
        authors: Vec::new(),
        authors_close: None,
        authors_empty: None,
        list_close: None,
        list_empty: None,
        comments: Vec::new(),
    };
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut open: Option<(String, usize, Range<usize>)> = None;
    let mut author = None::<String>;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("comments part markup: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    let name = element.name();
                    if let Some(at) = name.as_ref().iter().position(|byte| *byte == b':') {
                        layout.prefix = String::from_utf8_lossy(&name.as_ref()[..=at]).into_owned();
                    }
                }
                if stack.len() == 2 && local == b"comment" {
                    let reference = litchi_ooxml_common::xml::attributes::first_wins(&element)
                        .flatten()
                        .find(|attribute| attribute.key.as_ref() == b"ref")
                        .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                        .unwrap_or_default();
                    open = Some((reference, start, start..end));
                }
                if stack.len() == 2 && local == b"author" {
                    author = Some(String::new());
                }
                stack.push(local);
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if stack.len() == 1 && local.as_ref() == b"authors" {
                    layout.authors_empty = Some(start..end);
                }
                if stack.len() == 1 && local.as_ref() == b"commentList" {
                    layout.list_empty = Some(start..end);
                }
                if stack.len() == 2 && local.as_ref() == b"author" {
                    layout.authors.push(String::new());
                }
            },
            Event::Text(text) => {
                if let Some(author) = author.as_mut() {
                    let raw = String::from_utf8_lossy(text.as_ref()).into_owned();
                    author.push_str(
                        &quick_xml::escape::unescape(&raw).map_or(raw.clone(), |v| v.into_owned()),
                    );
                }
            },
            Event::GeneralRef(reference) => {
                if let (Some(author), Some(value)) = (
                    author.as_mut(),
                    crate::part_edit::entity(reference.as_ref()),
                ) {
                    author.push_str(&value);
                }
            },
            Event::End(_) => {
                let local = stack.pop().unwrap_or_default();
                if stack.len() == 1 && local == b"authors" {
                    layout.authors_close = Some(start);
                }
                if stack.len() == 1 && local == b"commentList" {
                    layout.list_close = Some(start);
                }
                if stack.len() == 2 && local == b"author" {
                    layout.authors.push(author.take().unwrap_or_default());
                }
                if stack.len() == 2
                    && local == b"comment"
                    && let Some((reference, begin, tag)) = open.take()
                {
                    layout.comments.push((reference, begin..end, tag));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(layout)
}

/// Replace or append the note of `reference` in a comments part.
fn splice_comment(xml: &[u8], reference: &str, author: &str, body: Body<'_>) -> Result<Vec<u8>> {
    let layout = scan(xml)?;
    let prefix = layout.prefix.as_str();
    let mut out = xml.to_vec();
    // Edits run back to front so earlier offsets stay valid.
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    let author_id = match layout.authors.iter().position(|known| known == author) {
        Some(index) => index,
        None => {
            let element = format!("<{prefix}author>{}</{prefix}author>", escape(author));
            match (layout.authors_close, &layout.authors_empty) {
                (Some(close), _) => edits.push((close..close, element)),
                (None, Some(span)) => edits.push((
                    span.clone(),
                    format!("<{prefix}authors>{element}</{prefix}authors>"),
                )),
                (None, None) => return Err(invalid("the comments part has no authors element")),
            }
            layout.authors.len()
        },
    };
    let body = body.element(prefix);
    match layout
        .comments
        .iter()
        .find(|(known, _, _)| known.eq_ignore_ascii_case(reference))
    {
        Some((_, span, tag)) => {
            let start_tag = String::from_utf8_lossy(&xml[tag.clone()]).into_owned();
            let start_tag = set_attribute(&start_tag, "authorId", &author_id.to_string())?;
            let close = format!("</{prefix}comment>");
            edits.push((span.clone(), format!("{start_tag}{body}{close}")));
        },
        None => {
            let element = format!(
                "<{prefix}comment ref=\"{reference}\" authorId=\"{author_id}\">{body}</{prefix}comment>"
            );
            match (layout.list_close, &layout.list_empty) {
                (Some(close), _) => edits.push((close..close, element)),
                (None, Some(span)) => edits.push((
                    span.clone(),
                    format!("<{prefix}commentList>{element}</{prefix}commentList>"),
                )),
                (None, None) => return Err(invalid("the comments part has no commentList")),
            }
        },
    }
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        out.splice(range, replacement.into_bytes());
    }
    Ok(out)
}

fn remove_comment(xml: &[u8], reference: &str) -> Result<Vec<u8>> {
    let layout = scan(xml)?;
    let (_, span, _) = layout
        .comments
        .iter()
        .find(|(known, _, _)| known.eq_ignore_ascii_case(reference))
        .ok_or_else(|| invalid("the note is missing"))?;
    let mut out = xml.to_vec();
    out.drain(span.clone());
    Ok(out)
}

/// Set (or add) one attribute of a start tag.
fn set_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let needle = format!(" {name}=\"");
    if let Some(at) = tag.find(&needle) {
        let value_start = at + needle.len();
        let value_end = tag[value_start..]
            .find('"')
            .map(|offset| value_start + offset)
            .ok_or_else(|| invalid("unterminated attribute"))?;
        let mut out = String::with_capacity(tag.len());
        out.push_str(&tag[..value_start]);
        out.push_str(value);
        out.push_str(&tag[value_end..]);
        return Ok(out);
    }
    let close = tag
        .rfind('>')
        .ok_or_else(|| invalid("unterminated start tag"))?;
    let mut out = String::new();
    let _ = write!(out, "{} {name}=\"{value}\"{}", &tag[..close], &tag[close..]);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use crate::Package;

    #[test]
    fn rich_runs_round_trip_and_replace_plain_text() {
        use crate::workbook::comments::{RunColor, RunFont, TextRun};
        let package = Package::new().unwrap();
        let mut edit = package.edit_comments().unwrap();
        assert!(edit.set(0, "B2", "Ann", "plain").unwrap());
        let runs = vec![
            TextRun {
                text: "Alice:".into(),
                font: RunFont {
                    bold: Some(true),
                    size: Some(9.0),
                    color: Some(RunColor::Rgb("FFFF0000".into())),
                    name: Some("Tahoma".into()),
                    ..RunFont::default()
                },
            },
            TextRun {
                text: " check <this>".into(),
                font: RunFont {
                    italic: Some(true),
                    underline: Some("double".into()),
                    vert_align: Some("superscript".into()),
                    strike: Some(false),
                    color: Some(RunColor::Indexed(81)),
                    ..RunFont::default()
                },
            },
        ];
        assert!(edit.set_runs(0, "B2", "Ann", &runs).unwrap());
        assert!(!edit.set_runs(0, "B2", "Ann", &runs).unwrap());
        assert_eq!(edit.add_runs(0, "A1", "Bob", &runs[..1]).unwrap(), Some(0));
        let after = edit.commit().unwrap().into_package();
        let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let comments = reopened
            .edit_comments()
            .unwrap()
            .comments(0)
            .unwrap()
            .unwrap();
        let note = comments.get("B2").unwrap();
        assert_eq!(note.text, "Alice: check <this>");
        assert_eq!(note.runs, runs);
        assert_eq!(comments.get("A1").unwrap().runs, runs[..1]);
        let mut edit = reopened.edit_comments().unwrap();
        assert!(edit.set(0, "B2", "Ann", "plain again").unwrap());
        let comments = edit.comments(0).unwrap().unwrap();
        assert!(comments.get("B2").unwrap().runs.is_empty());
        assert!(edit.set_runs(0, "C3", "Ann", &[]).is_err());
        assert!(
            edit.set_runs(
                0,
                "C3",
                "Ann",
                &[TextRun {
                    text: "x".into(),
                    font: RunFont {
                        size: Some(0.0),
                        ..RunFont::default()
                    }
                }]
            )
            .is_err()
        );
    }

    #[test]
    fn notes_add_replace_remove_with_vml() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_comments().unwrap();
        assert!(edit.set(0, "B2", "Ann", "first").unwrap());
        assert!(edit.set(0, "A10", "Bob & Co", "second & <more>").unwrap());
        assert!(!edit.set(0, "B2", "Ann", "first").unwrap());
        let (after, patch) = edit.commit().unwrap().into_parts();
        let comments = after.edit_comments().unwrap().comments(0).unwrap().unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments.get("A10").unwrap().text, "second & <more>");
        assert_eq!(comments.get("A10").unwrap().author, "Bob & Co");
        assert_eq!(comments.authors.len(), 2);
        let bytes = after.to_bytes().unwrap();
        let reopened = Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            reopened
                .edit_comments()
                .unwrap()
                .comments(0)
                .unwrap()
                .unwrap()
                .len(),
            2
        );
        assert!(
            package
                .apply_comments(&patch)
                .unwrap()
                .package()
                .same_content(&after)
                .unwrap()
        );

        let mut edit = after.edit_comments().unwrap();
        assert!(edit.set(0, "B2", "Bob & Co", "changed").unwrap());
        assert!(edit.remove(0, "A10").unwrap());
        let after = edit.commit().unwrap().into_package();
        let comments = after.edit_comments().unwrap().comments(0).unwrap().unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments.get("B2").unwrap().author, "Bob & Co");
        assert_eq!(comments.authors.len(), 2);

        let mut edit = after.edit_comments().unwrap();
        assert!(edit.remove(0, "B2").unwrap());
        assert!(!edit.remove(0, "B2").unwrap());
        let empty = edit.commit().unwrap().into_package();
        assert!(
            empty
                .edit_comments()
                .unwrap()
                .comments(0)
                .unwrap()
                .is_none()
        );
        assert!(empty.same_content(&package).unwrap());
    }
}
