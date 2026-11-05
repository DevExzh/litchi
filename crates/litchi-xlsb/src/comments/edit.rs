//! Immutable note (legacy comment) edits of an XLSB [`Package`]
//! (unified-ops record 0322).
//!
//! A worksheet's notes live in its comments part ([MS-XLSB] section
//! 2.1.7.8): an author list and one `BrtBeginComment` collection per note.
//! Excel draws each note from a hidden VML shape in the worksheet's VML
//! drawing, which the worksheet names with `BrtLegacyDrawing`. [`Edit`]
//! adds, replaces and removes notes one at a time:
//!
//! - every other note keeps its bytes (rich text runs and alternate-content
//!   identifiers included); a new author is appended to the author list;
//! - a new note gets a hidden VML note shape, and the comments part, the
//!   VML part and `BrtLegacyDrawing` are created with the first note and
//!   removed with the last one when the VML drawing holds nothing else.
//!
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_comments`]
//! replays a patch onto an equal package.
//!
//! ```
//! use litchi_xlsb::Package;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_comments()?;
//! assert_eq!(edit.add(0, 1, 1, "Ann", "Check this")?, Some(0));
//! let after = edit.commit()?.into_package();
//! let notes = after.edit_comments()?.comments(0)?;
//! assert_eq!((notes[0].row, notes[0].col, notes[0].text.as_str()), (1, 1, "Check this"));
//! # Ok(())
//! # }
//! ```

use std::ops::Range;

use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{BlobPart, PackURI};

use super::{MAX_COLUMNS, MAX_ROWS, MAX_TEXT_UNITS, Record};
use crate::package::Package;
use crate::package::error::{Error as PackageError, Result};
use crate::part_edit::{
    Rank, SheetLayout, Staged, family, invalid, record_with_limits, snapshot_options, splice,
};
use crate::raw::{Encoder, Records, kind};
use litchi_spreadsheet_drawing::vml;

family!("note");
snapshot_options!("note");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Notes of one worksheet ([`MAX_COMMENTS`]).
    Comments,
}

/// The most notes one worksheet holds through this edit.
pub const MAX_COMMENTS: usize = 65_530;

/// The content type of an XLSB comments part.
const COMMENTS_CONTENT_TYPE: &str = "application/vnd.ms-excel.comments";

/// A staged note edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start a note edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_comments(&self) -> Result<Edit> {
        self.edit_comments_with_options(EditOptions::default())
    }

    /// Start a note edit with an explicit retained-payload ceiling.
    pub fn edit_comments_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a note patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_comments(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// The comments part a worksheet names, with its relationship id.
pub(crate) fn comments_part(
    relationships: &litchi_opc::Relationships,
) -> Result<Option<(String, PackURI)>> {
    let mut found = relationships
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::COMMENTS);
    let Some(first) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid("a worksheet names more than one comments part"));
    }
    if first.is_external() {
        return Err(invalid("a comments relationship cannot be external"));
    }
    Ok(Some((first.r_id().to_owned(), first.target_partname()?)))
}

/// The record layout of a comments stream.
struct Layout {
    authors: Vec<String>,
    /// The offset of `BrtEndCommentAuthors`.
    authors_end: usize,
    /// Each note's span (its alternate-content identifier included) and
    /// zero-based cell.
    notes: Vec<(Range<usize>, (u32, u32))>,
    /// The offset of `BrtEndCommentList`.
    list_end: usize,
}

fn layout(bytes: &[u8], limits: crate::raw::Limits) -> Result<Layout> {
    let mut authors = Vec::new();
    let mut authors_end = None;
    let mut notes = Vec::new();
    let mut list_end = None;
    let mut note_start: Option<usize> = None;
    let mut cell = None;
    let mut records = Records::new(bytes, limits);
    while let Some(item) = records.next() {
        let item = item?;
        let start = item.offset();
        let end = records.offset();
        match item.kind() {
            kind::COMMENT_AUTHOR => {
                let payload = item.payload();
                let units = payload
                    .get(..4)
                    .map(|count| u32::from_le_bytes([count[0], count[1], count[2], count[3]]))
                    .ok_or_else(|| invalid("a truncated comment author"))?;
                let units =
                    usize::try_from(units).map_err(|_overflow| invalid("comment author"))?;
                let text: Vec<u16> = payload
                    .get(4..4 + units.saturating_mul(2))
                    .ok_or_else(|| invalid("a truncated comment author"))?
                    .chunks_exact(2)
                    .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                    .collect();
                authors.push(String::from_utf16_lossy(&text));
            },
            kind::END_COMMENT_AUTHORS => authors_end = Some(start),
            kind::AC_BEGIN if authors_end.is_some() && note_start.is_none() => {
                note_start = Some(start);
            },
            kind::BEGIN_COMMENT => {
                note_start.get_or_insert(start);
                let payload = item.payload();
                if payload.len() != 36 {
                    return Err(invalid("a comment header is not 36 bytes"));
                }
                let row = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
                let col = u32::from_le_bytes([payload[12], payload[13], payload[14], payload[15]]);
                cell = Some((row, col));
            },
            kind::END_COMMENT => {
                let begin = note_start
                    .take()
                    .ok_or_else(|| invalid("a comment end without a comment"))?;
                let cell = cell
                    .take()
                    .ok_or_else(|| invalid("a comment end without a comment"))?;
                notes.push((begin..end, cell));
            },
            kind::END_COMMENT_LIST => list_end = Some(start),
            _ => {},
        }
    }
    Ok(Layout {
        authors,
        authors_end: authors_end.ok_or_else(|| invalid("a comments part has no author list"))?,
        notes,
        list_end: list_end.ok_or_else(|| invalid("a comments part has no comment list"))?,
    })
}

fn check(row: u32, col: u32, author: &str, text: &str) -> Result<()> {
    if row >= MAX_ROWS || col >= MAX_COLUMNS {
        return Err(invalid("the note cell is outside the worksheet grid"));
    }
    if author.encode_utf16().count() > super::MAX_AUTHOR_UNITS {
        return Err(invalid("a note author holds at most 54 characters"));
    }
    if text.encode_utf16().count() > MAX_TEXT_UNITS || text.contains('\0') || author.contains('\0')
    {
        return Err(invalid(
            "a note text holds at most 32767 characters and no NUL",
        ));
    }
    Ok(())
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    fn part(&self, worksheet: &PackURI) -> Result<Option<(String, PackURI)>> {
        comments_part(self.staged.opc().get_part(worksheet)?.rels())
    }

    /// The staged notes of the worksheet at complete-catalog position
    /// `sheet`, in stream order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// comments part.
    pub fn comments(&self, sheet: usize) -> Result<Vec<Record>> {
        let worksheet = self.staged.worksheet(sheet)?;
        match self.part(&worksheet)? {
            Some((_, name)) => read(self.staged.blob(&name)?, self.staged.raw_limits()),
            None => Ok(Vec::new()),
        }
    }

    /// Add a note to cell `(row, col)` (zero-based) when it has none.
    /// Returns the new note's position among the worksheet's notes in
    /// row-major order of their cells; `None` leaves a cell that already has
    /// a note unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a cell outside
    /// the grid, an author or text the record cannot hold, more than
    /// [`MAX_COMMENTS`] notes, a VML drawing this editor cannot read, or a
    /// changed edit of a signed workbook.
    pub fn add(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        author: &str,
        text: &str,
    ) -> Result<Option<usize>> {
        let notes = self.comments(sheet)?;
        if notes.iter().any(|note| (note.row, note.col) == (row, col)) {
            return Ok(None);
        }
        let position = notes
            .iter()
            .filter(|note| (note.row, note.col) < (row, col))
            .count();
        self.write(sheet, row, col, author, text, true)?;
        Ok(Some(position))
    }

    /// Add or replace the note of cell `(row, col)`. A replaced note keeps
    /// its identifiers and its VML shape; its rich-text runs are dropped
    /// when the text changes. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// As [`Edit::add`].
    pub fn set(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        author: &str,
        text: &str,
    ) -> Result<bool> {
        self.write(sheet, row, col, author, text, false)
    }

    fn write(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        author: &str,
        text: &str,
        insert_only: bool,
    ) -> Result<bool> {
        check(row, col, author, text)?;
        let worksheet = self.staged.worksheet(sheet)?;
        let existing = self.part(&worksheet)?;
        let Some((_, name)) = existing else {
            // The first note: a new comments part.
            let note = Record::new(row, col, author.to_owned(), text.to_owned());
            let mut bytes = Vec::new();
            super::write(
                &mut Encoder::new(&mut bytes, self.staged.raw_limits()),
                std::slice::from_ref(&note),
            )?;
            let name = self.staged.free_name("/xl/comments", ".bin")?;
            self.staged.add_part(BlobPart::new(
                name.clone(),
                COMMENTS_CONTENT_TYPE.to_owned(),
                bytes,
            ))?;
            let opc = self.staged.opc_mut()?;
            opc.get_part_mut(&worksheet)?.relate_to(
                &name.relative_ref(worksheet.base_uri()),
                relationship_type::COMMENTS,
            );
            self.staged.touch(&worksheet);
            self.add_note_shape(sheet, &worksheet, row, col)?;
            return Ok(true);
        };
        let bytes = self.staged.blob(&name)?.to_vec();
        let limits = self.staged.raw_limits();
        let notes = read(&bytes, limits)?;
        let layout = layout(&bytes, limits)?;
        let current = notes
            .iter()
            .position(|note| (note.row, note.col) == (row, col));
        if let Some(index) = current {
            let note = &notes[index];
            if insert_only || (note.author == author && note.text == text) {
                return Ok(false);
            }
        } else if notes.len() >= MAX_COMMENTS {
            return Err(invalid("a worksheet holds at most notes"));
        }
        // The author index, appending a new author.
        let (author_index, author_record) =
            match layout.authors.iter().position(|name| name == author) {
                Some(index) => (index, None),
                None => (
                    layout.authors.len(),
                    Some(record_with_limits(
                        kind::COMMENT_AUTHOR,
                        &super::codec::encode_author(author)?,
                        limits,
                    )?),
                ),
            };
        let author_index =
            u32::try_from(author_index).map_err(|_overflow| invalid("too many authors"))?;
        let mut note = match current {
            Some(index) => notes[index].clone(),
            None => Record::new(row, col, String::new(), String::new()),
        };
        if note.text != text {
            note.runs.clear();
        }
        note.author = author.to_owned();
        note.text = text.to_owned();
        let mut encoded = Vec::new();
        super::codec::write_comment(&mut Encoder::new(&mut encoded, limits), &note, author_index)
            .map_err(|error| match error {
            super::codec::Error::UnsupportedFeature(_) => PackageError::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_comments",
                "the note contains unsupported rich-text metadata",
            ),
            error @ (super::codec::Error::Wire(_)
            | super::codec::Error::InvalidRecordType(_)
            | super::codec::Error::InvalidLength { .. }
            | super::codec::Error::Unrecognized { .. }
            | super::codec::Error::Encoding(_)
            | super::codec::Error::Allocation { .. }) => error.into(),
        })?;
        let span = match current {
            Some(index) => {
                let (span, _) = &layout.notes[index];
                // Keep an alternate-content identifier that precedes the
                // note's header.
                let header = header_offset(&bytes, span.clone(), limits)?;
                header..span.end
            },
            None => {
                let at = layout
                    .notes
                    .iter()
                    .find(|(_, cell)| *cell > (row, col))
                    .map_or(layout.list_end, |(span, _)| span.start);
                at..at
            },
        };
        let mut updated = splice(&bytes, span.clone(), &encoded);
        if let Some(author_record) = author_record {
            // The author list precedes every note, so its offset is
            // unchanged by the note splice above.
            updated = splice(
                &updated,
                layout.authors_end..layout.authors_end,
                &author_record,
            );
        }
        self.staged.replace_blob(&name, updated)?;
        if current.is_none() {
            self.add_note_shape(sheet, &worksheet, row, col)?;
        }
        Ok(true)
    }

    /// Remove the note of cell `(row, col)`. Returns whether a note was
    /// removed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a VML drawing
    /// this editor cannot read, or a changed edit of a signed workbook.
    pub fn remove(&mut self, sheet: usize, row: u32, col: u32) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let Some((id, name)) = self.part(&worksheet)? else {
            return Ok(false);
        };
        let bytes = self.staged.blob(&name)?.to_vec();
        let layout = layout(&bytes, self.staged.raw_limits())?;
        let Some(index) = layout
            .notes
            .iter()
            .position(|(_, cell)| *cell == (row, col))
        else {
            return Ok(false);
        };
        if layout.notes.len() == 1 {
            let opc = self.staged.opc_mut()?;
            opc.get_part_mut(&worksheet)?.rels_mut().remove(&id);
            self.staged.remove_part(&name)?;
            self.staged.touch(&worksheet);
        } else {
            let (span, _) = &layout.notes[index];
            let updated = splice(&bytes, span.clone(), &[]);
            self.staged.replace_blob(&name, updated)?;
        }
        self.remove_note_shape(&worksheet, row, col)?;
        Ok(true)
    }

    /// The worksheet's VML drawing: its relationship id and part name.
    fn vml(&self, worksheet: &PackURI) -> Result<Option<(String, PackURI)>> {
        let bytes = self.staged.blob(worksheet)?;
        let layout = SheetLayout::scan_with_limits(bytes, self.staged.raw_limits())?;
        let Some(unit) = layout.units_of(kind::LEGACY_DRAWING).next() else {
            return Ok(None);
        };
        let limits = self.staged.raw_limits();
        let item = Records::new(&bytes[unit.span.clone()], limits)
            .next()
            .transpose()?
            .ok_or_else(|| invalid("an empty legacy-drawing record"))?;
        let mut cursor = crate::raw::Cursor::new(item.payload(), "BrtLegacyDrawing", limits);
        let id = cursor.read_wide_string()?;
        let part = self.staged.opc().get_part(worksheet)?;
        let relationship = part
            .rels()
            .get(&id)
            .ok_or_else(|| invalid("the legacy-drawing relationship is missing"))?;
        if relationship.is_external() {
            return Err(invalid("a legacy-drawing relationship cannot be external"));
        }
        Ok(Some((id, relationship.target_partname()?)))
    }

    fn add_note_shape(
        &mut self,
        sheet: usize,
        worksheet: &PackURI,
        row: u32,
        col: u32,
    ) -> Result<()> {
        match self.vml(worksheet)? {
            Some((_, name)) => {
                let xml = self.staged.blob(&name)?;
                let limits = vml::Limits::from_source_limits(self.staged.opc().read_limits());
                let shapes = vml::shapes(xml, &limits)?;
                if shapes.iter().any(|shape| shape.note == Some((row, col))) {
                    return Ok(());
                }
                let id = vml::next_shape_id(xml, &shapes, sheet);
                let close = vml::rfind(xml, b"</xml>")
                    .ok_or_else(|| invalid("the VML drawing has no closing xml element"))?;
                let mut fragment = String::new();
                if !vml::contains(xml, b"_x0000_t202") {
                    fragment.push_str(vml::SHAPE_TYPE);
                }
                fragment.push_str(&vml::note_shape(id, row, col));
                let updated = splice(xml, close..close, fragment.as_bytes());
                self.staged.replace_blob(&name, updated)?;
            },
            None => {
                let name = self.staged.free_name("/xl/drawings/vmlDrawing", ".vml")?;
                let xml = vml::new_drawing(sheet, row, col);
                self.staged.add_part(BlobPart::new(
                    name.clone(),
                    content_type::OFC_VML_DRAWING.to_owned(),
                    xml.into_bytes(),
                ))?;
                let id = self.staged.opc_mut()?.get_part_mut(worksheet)?.relate_to(
                    &name.relative_ref(worksheet.base_uri()),
                    relationship_type::VML_DRAWING,
                );
                let bytes = self.staged.blob(worksheet)?;
                let limits = self.staged.raw_limits();
                let layout = SheetLayout::scan_with_limits(bytes, limits)?;
                let at = layout.insertion(Rank::LegacyDrawing);
                let mut payload = Vec::new();
                Encoder::new(&mut payload, limits).write_wide_string(&id)?;
                let updated = splice(
                    bytes,
                    at..at,
                    &record_with_limits(kind::LEGACY_DRAWING, &payload, limits)?,
                );
                self.staged.replace_worksheet(worksheet, updated)?;
                self.staged.touch(worksheet);
            },
        }
        Ok(())
    }

    fn remove_note_shape(&mut self, worksheet: &PackURI, row: u32, col: u32) -> Result<()> {
        let Some((id, name)) = self.vml(worksheet)? else {
            return Ok(());
        };
        let xml = self.staged.blob(&name)?;
        let limits = vml::Limits::from_source_limits(self.staged.opc().read_limits());
        let shapes = vml::shapes(xml, &limits)?;
        let Some(shape) = shapes.iter().find(|shape| shape.note == Some((row, col))) else {
            return Ok(());
        };
        if shapes.len() > 1 {
            let updated = splice(xml, shape.span.clone(), &[]);
            self.staged.replace_blob(&name, updated)?;
            return Ok(());
        }
        // The last shape: drop the drawing, its relationship and the
        // worksheet's `BrtLegacyDrawing`.
        let bytes = self.staged.blob(worksheet)?;
        let layout = SheetLayout::scan_with_limits(bytes, self.staged.raw_limits())?;
        let span = layout
            .units_of(kind::LEGACY_DRAWING)
            .next()
            .map(|unit| unit.span.clone())
            .ok_or_else(|| invalid("the legacy-drawing record vanished"))?;
        let updated = splice(bytes, span, &[]);
        self.staged.replace_worksheet(worksheet, updated)?;
        let opc = self.staged.opc_mut()?;
        opc.get_part_mut(worksheet)?.rels_mut().remove(&id);
        let referenced = opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship.target_partname().ok().as_ref() == Some(&name)
            })
        });
        if !referenced {
            self.staged.remove_part(&name)?;
        }
        self.staged.touch(worksheet);
        Ok(())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

/// Decode a comments stream, recording (not refusing) repairs as an opened
/// worksheet does.
pub(crate) fn read(bytes: &[u8], limits: crate::raw::Limits) -> Result<Vec<Record>> {
    let mut report = crate::leniency::ToleranceReport::default();
    Ok(super::read_recording(bytes, limits, &mut report)?)
}

/// The offset of the `BrtBeginComment` inside a note's span.
fn header_offset(bytes: &[u8], span: Range<usize>, limits: crate::raw::Limits) -> Result<usize> {
    let mut records = Records::new(&bytes[span.clone()], limits);
    while let Some(item) = records.next() {
        let item = item?;
        if item.kind() == kind::BEGIN_COMMENT {
            return Ok(span.start + item.offset());
        }
    }
    Err(invalid("a note has no header"))
}

#[cfg(test)]
mod tests {
    use crate::Package;

    fn part(package: &Package, name: &str) -> Option<Vec<u8>> {
        package
            .opc_package()
            .get_part(&litchi_opc::PackURI::new(name).unwrap())
            .ok()
            .map(|part| part.blob().to_vec())
    }

    #[test]
    fn notes_add_replace_remove_with_vml() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_comments().unwrap();
        assert_eq!(edit.add(0, 4, 1, "Ann", "second").unwrap(), Some(0));
        assert_eq!(edit.add(0, 0, 0, "Bob", "first").unwrap(), Some(0));
        assert_eq!(edit.add(0, 0, 0, "Bob", "again").unwrap(), None);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let notes = after.edit_comments().unwrap().comments(0).unwrap();
        let cells: Vec<_> = notes
            .iter()
            .map(|note| (note.row, note.col, note.author.as_str()))
            .collect();
        assert_eq!(cells, [(0, 0, "Bob"), (4, 1, "Ann")]);
        let vml = String::from_utf8(part(&after, "/xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert_eq!(vml.matches("ObjectType=\"Note\"").count(), 2);
        assert_eq!(
            package
                .apply_comments(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        // The opened workbook reads the notes.
        let workbook = after.workbook().unwrap();
        let grid = workbook.grid(0usize).unwrap().unwrap();
        assert_eq!(grid.worksheet().comments().len(), 2);

        let mut edit = after.edit_comments().unwrap();
        assert!(edit.set(0, 4, 1, "Ann", "changed").unwrap());
        assert!(!edit.set(0, 4, 1, "Ann", "changed").unwrap());
        let set = edit.commit().unwrap().into_package();
        assert_eq!(
            set.edit_comments().unwrap().comments(0).unwrap()[1].text,
            "changed"
        );

        let mut edit = set.edit_comments().unwrap();
        assert!(edit.remove(0, 0, 0).unwrap());
        assert!(edit.remove(0, 4, 1).unwrap());
        assert!(!edit.remove(0, 4, 1).unwrap());
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(removed.to_bytes().unwrap(), package.to_bytes().unwrap());
    }

    #[test]
    fn fixture_notes_keep_their_bytes() {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/comments.xlsb"
        ))
        .unwrap();
        let package = Package::from_bytes(&bytes).unwrap();
        let before = package.edit_comments().unwrap().comments(0).unwrap();
        assert!(!before.is_empty());
        let mut edit = package.edit_comments().unwrap();
        assert!(edit.add(0, 20, 5, "New author", "added").unwrap().is_some());
        let added = edit.commit().unwrap().into_package();
        let notes = added.edit_comments().unwrap().comments(0).unwrap();
        assert_eq!(notes.len(), before.len() + 1);
        assert_eq!(&notes[..before.len()], &before[..]);
        let mut edit = added.edit_comments().unwrap();
        assert!(edit.remove(0, 20, 5).unwrap());
        let back = edit.commit().unwrap().into_package();
        // The added author stays; every original note keeps its bytes.
        assert_eq!(back.edit_comments().unwrap().comments(0).unwrap(), before);
        assert_eq!(
            part(&back, "/xl/drawings/vmlDrawing1.vml"),
            part(&package, "/xl/drawings/vmlDrawing1.vml")
        );
        assert_eq!(
            part(&back, "/xl/worksheets/sheet1.bin"),
            part(&package, "/xl/worksheets/sheet1.bin")
        );
    }
}
