//! Element edits of the root transaction (record 0323): slide comments,
//! shape and text hyperlinks, and table-cell text.
//!
//! Each verb computes complete record replacements from the working
//! artifact's exact records and publishes them through
//! [`Transaction::stage_records`], which keeps every other record byte for
//! byte and carries the change through exact, inverse and durable patches.

use std::io::Cursor;

use super::records::Replacement;
use super::{Error, PackageError, Position, Refusal, Result, Transaction};
use crate::comments::{CommentChange, NewComment};

impl Transaction {
    /// The working artifact's presentation projection: the opened source
    /// while nothing is staged, else a parse of the working bytes.
    pub(super) fn working_presentation(&self) -> Result<crate::Presentation> {
        if let Some(opened) = self.opened_source() {
            return Ok(opened.clone());
        }
        let mut reader = crate::raw::Reader::from_reader_with_limits(
            Cursor::new(self.working.bytes()),
            self.working.limits,
        )?;
        Ok(reader.presentation_with_options(self.working.open_options())?)
    }

    /// The next per-author comment index: one more than the author's highest
    /// index on any slide and than the author's document-level seed.
    fn next_comment_index(&self, author: &str) -> Result<i32> {
        let presentation = self.working_presentation()?;
        let mut highest = 0_i32;
        for slide in presentation.comments()? {
            for comment in slide.comments {
                if comment.author == author {
                    highest = highest.max(comment.index);
                }
            }
        }
        let authors = crate::comments::Authors::parse(&presentation.live_document_record()?)?;
        if let Some(seed) = authors
            .find(author)
            .and_then(|entry| entry.comment_index_seed)
        {
            highest = highest.max(seed);
        }
        highest
            .checked_add(1)
            .ok_or(Error::Refused(Refusal::InvalidElementValue {
                element: "comment",
                reason: "the author's comment index is exhausted",
            }))
    }

    /// Append a comment to the slide at a projected position; returns its
    /// zero-based position among the slide's comments.
    ///
    /// The comment goes after the slide's existing `Comment10Container`
    /// records in its `___PPT10` programmable tag, which is created when
    /// absent. Its index is the author's next one; a document-level author
    /// entry has its seed raised to it. The author list is not extended.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent or inserted slide, a value the
    /// records cannot hold, a slide whose tags are ambiguous, or a changed
    /// signed source.
    pub fn add_comment(&mut self, position: Position, comment: &NewComment) -> Result<usize> {
        let index = self.next_comment_index(&comment.author)?;
        let mut added = 0;
        self.stage_records(Some(position), true, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            let (record, at) = crate::comments::edit::add(slide.bytes, comment, index)?;
            added = at;
            let document = match view.document {
                Some(document) => {
                    crate::comments::edit::raise_author_seed(document, &comment.author, index)?
                },
                None => None,
            };
            Ok(Replacement {
                slide: Some(record),
                document,
            })
        })?;
        Ok(added)
    }

    /// Change the text or anchor of the comment at `comment` on the slide at
    /// a projected position. An unchanged value is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide or comment, a value the
    /// records cannot hold, or a changed signed source.
    pub fn set_comment(
        &mut self,
        position: Position,
        comment: Position,
        change: &CommentChange,
    ) -> Result<()> {
        self.stage_records(Some(position), false, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            Ok(Replacement {
                slide: crate::comments::edit::set(slide.bytes, comment.get(), change)?,
                document: None,
            })
        })
        .map(drop)
    }

    /// Remove the comment at `comment` from the slide at a projected
    /// position. The slide's programmable tag stays, possibly empty.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide or comment, or a changed
    /// signed source.
    pub fn remove_comment(&mut self, position: Position, comment: Position) -> Result<()> {
        self.stage_records(Some(position), false, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            Ok(Replacement {
                slide: Some(crate::comments::edit::remove(slide.bytes, comment.get())?),
                document: None,
            })
        })
        .map(drop)
    }
}

fn missing_record() -> Error {
    PackageError::Corrupted("PPT element edit lost its selected record".into()).into()
}

impl Transaction {
    /// The `OfficeArt` shape ID of the cell at `row`, `column` of the table
    /// at `target` in the working artifact.
    fn table_cell_shape_id(
        &self,
        target: crate::text_edit::Target,
        row: Position,
        column: Position,
    ) -> Result<u32> {
        let refusal = |element| Error::Refused(Refusal::ElementNotFound { element });
        let presentation = self.working_presentation()?;
        let projected = self.document.slides()?;
        let persist_id = projected
            .get(target.slide().get())
            .ok_or(Error::Refused(Refusal::SlideNotFound {
                position: target.slide(),
            }))?
            .persist_id();
        let slide = presentation
            .slides()?
            .into_iter()
            .find(|slide| slide.persist_id() == persist_id)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let shapes = slide.shapes()?;
        let table = shapes
            .get(target.shape().get())
            .and_then(crate::shapes::ShapeEnum::as_table)
            .ok_or_else(|| refusal("table"))?;
        if row.get() >= table.rows() {
            return Err(refusal("table row"));
        }
        if column.get() >= table.columns() {
            return Err(refusal("table column"));
        }
        table
            .cell_shape_id(row.get(), column.get())
            .ok_or(Error::Refused(Refusal::UnsupportedElement {
                element: "table cell",
                reason: "the cell is covered by a merged cell",
            }))
    }

    /// Replace the text of one table cell (record 0323).
    ///
    /// A legacy table is a group of cell shapes; the cell's shape text is
    /// rewritten under the ordinary shape-text closure rules, so a changed
    /// length refuses where the cell's text has several formatting runs or
    /// other position-dependent records. Equal text is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide, table, row or column, a
    /// merged-away cell, a cell without editable text, an unrepresentable
    /// text, or a changed signed source.
    pub fn set_table_cell_text(
        &mut self,
        target: crate::text_edit::Target,
        row: Position,
        column: Position,
        text: &str,
    ) -> Result<()> {
        self.require_position(target.slide())?;
        let shape_id = self.table_cell_shape_id(target, row, column)?;
        self.stage_records(Some(target.slide()), false, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            Ok(Replacement {
                slide: crate::text_edit::rewrite_shape_text_by_id(slide.bytes, shape_id, text)?,
                document: None,
            })
        })
        .map(drop)
    }
}
