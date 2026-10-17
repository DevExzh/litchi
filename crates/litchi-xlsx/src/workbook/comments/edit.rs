//! Immutable legacy-note edits of a [`Package`] (unified-ops record 0320).
//!
//! [`Package::edit_comments`] starts an [`Edit`] over an immutable snapshot.
//! [`Edit::set`] adds or replaces the note of one cell, [`Edit::remove`]
//! deletes it, and [`Edit::commit`] returns a [`Commit`] with the new
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
use quick_xml::events::Event;

use super::model::Comments;
use super::package::{COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, load_from_worksheet};
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
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a legacy-note edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_comments(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_comments")?,
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
    pub fn apply_comments(&self, patch: &Patch) -> Result<Self> {
        patch.0.apply(self, "apply_comments")
    }
}

/// The checked A1 reference and zero-based (row, column) of a cell.
fn cell(value: &str) -> Result<(String, u32, u32)> {
    let cell = litchi_sheet::Cell::from_a1(value)?;
    Ok((cell.a1(), cell.row().get(), cell.column().get()))
}

impl Edit<'_> {
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
        self.write(sheet, cell_ref, author, text, false)
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
            .write_loaded(sheet, cell_ref, author, text, true, before)?
            .then_some(position))
    }

    fn write(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        text: &str,
        insert_only: bool,
    ) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let existing = load_from_worksheet(self.staged.opc(), &worksheet)?;
        self.write_loaded(sheet, cell_ref, author, text, insert_only, existing)
    }

    fn write_loaded(
        &mut self,
        sheet: usize,
        cell_ref: &str,
        author: &str,
        text: &str,
        insert_only: bool,
        existing: Option<super::model::Part>,
    ) -> Result<bool> {
        let (reference, row, column) = cell(cell_ref)?;
        for (value, what) in [(author, "author"), (text, "text")] {
            if value.len() > MAX_TEXT_BYTES || value.contains('\0') {
                return Err(invalid(format!(
                    "a note {what} must hold at most {MAX_TEXT_BYTES} bytes and no NUL"
                )));
            }
        }
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(part) = &existing {
            if let Some(comment) = part.comments.get(&reference) {
                if insert_only || (comment.author == author && comment.text == text) {
                    return Ok(false);
                }
            } else if part.comments.len() >= MAX_COMMENTS {
                return Err(invalid(format!(
                    "a worksheet holds at most {MAX_COMMENTS} notes"
                )));
            }
        }
        let added = existing
            .as_ref()
            .is_none_or(|part| part.comments.get(&reference).is_none());
        match existing {
            Some(part) => {
                let name = PackURI::new(part.part_name.as_str())
                    .map_err(|error| invalid(error.to_string()))?;
                let xml = splice_comment(self.staged.blob(&name)?, &reference, author, text)?;
                self.staged.replace_blob(&name, xml)?;
            },
            None => {
                let name = self.staged.free_name("/xl/comments", ".xml")?;
                let xml = new_comments_part(&reference, author, text);
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
                let shapes = shapes(xml)?;
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
                let xml = format!(
                    "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{}\"/></o:shapelayout>{SHAPE_TYPE}{}</xml>",
                    sheet + 1,
                    note_shape(first_shape_id(sheet), row, column)
                );
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
        let shapes = shapes(xml)?;
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

const SHAPE_TYPE: &str = "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>";

fn first_shape_id(sheet: usize) -> u64 {
    (sheet as u64 + 1) * 1024 + 1
}

/// One VML shape: its byte span and, for a note, its zero-based cell.
struct Shape {
    span: Range<usize>,
    note: Option<(u32, u32)>,
    id: Option<u64>,
}

fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

fn find(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack
        .get(from..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| at + from)
}

fn rfind(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .rposition(|window| window == needle)
}

fn element_text(xml: &[u8], tag: &[u8]) -> Option<u32> {
    let open = [b"<x:", tag, b">"].concat();
    let close = [b"</x:", tag, b">"].concat();
    let start = find(xml, &open, 0)? + open.len();
    let end = find(xml, &close, start)?;
    std::str::from_utf8(&xml[start..end])
        .ok()?
        .trim()
        .parse()
        .ok()
}

/// The `v:shape` elements of a VML drawing, scanned bytewise: VML parts are
/// not always well-formed XML.
fn shapes(xml: &[u8]) -> Result<Vec<Shape>> {
    let mut shapes = Vec::new();
    let mut cursor = 0;
    while let Some(start) = find(xml, b"<v:shape", cursor) {
        let after = xml.get(start + 8).copied();
        if !matches!(after, Some(b' ' | b'>' | b'\t' | b'\r' | b'\n')) {
            cursor = start + 8;
            continue;
        }
        let tag_end = find(xml, b">", start).ok_or_else(|| invalid("unterminated VML shape"))?;
        let end = if xml[tag_end - 1] == b'/' {
            tag_end + 1
        } else {
            find(xml, b"</v:shape>", tag_end).ok_or_else(|| invalid("unterminated VML shape"))?
                + b"</v:shape>".len()
        };
        let body = &xml[start..end];
        let note = (contains(body, b"ObjectType=\"Note\"") || contains(body, b"ObjectType='Note'"))
            .then(|| element_text(body, b"Row").zip(element_text(body, b"Column")))
            .flatten();
        let id = find(body, b"id=\"_x0000_s", 0).and_then(|at| {
            let digits = &body[at + 11..];
            let len = digits
                .iter()
                .take_while(|byte| byte.is_ascii_digit())
                .count();
            std::str::from_utf8(&digits[..len]).ok()?.parse().ok()
        });
        shapes.push(Shape {
            span: start..end,
            note,
            id,
        });
        cursor = end;
    }
    Ok(shapes)
}

fn next_shape_id(xml: &[u8], shapes: &[Shape], sheet: usize) -> u64 {
    let base = find(xml, b"data=\"", 0)
        .and_then(|at| {
            let digits = &xml[at + 6..];
            let len = digits
                .iter()
                .take_while(|byte| byte.is_ascii_digit())
                .count();
            std::str::from_utf8(&digits[..len])
                .ok()?
                .parse::<u64>()
                .ok()
        })
        .map_or_else(|| first_shape_id(sheet), |block| block * 1024 + 1);
    shapes
        .iter()
        .filter_map(|shape| shape.id)
        .max()
        .map_or(base, |id| id.max(base - 1) + 1)
}

/// A hidden note box anchored beside the cell, as Excel draws one.
fn note_shape(id: u64, row: u32, column: u32) -> String {
    let anchor_col = column + 1;
    let anchor_row = row.saturating_sub(1);
    format!(
        "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor_col}, 15, {anchor_row}, 2, {}, 15, {}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{column}</x:Column></x:ClientData></v:shape>",
        anchor_col + 2,
        anchor_row + 4
    )
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

fn new_comments_part(reference: &str, author: &str, text: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>{}</author></authors><commentList><comment ref=\"{reference}\" authorId=\"0\">{}</comment></commentList></comments>",
        escape(author),
        text_element("", text)
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
fn splice_comment(xml: &[u8], reference: &str, author: &str, text: &str) -> Result<Vec<u8>> {
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
    let body = text_element(prefix, text);
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
