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
