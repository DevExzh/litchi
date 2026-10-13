//! Positional legacy-comment verbs for the unified API (record 0252).

use litchi_core::Position;

use super::Edit;
use crate::comments::{Author, Comment, Conformance};
use crate::{Error, Result};

impl Edit {
    fn comment_conformance(&self, slide: &litchi_opc::PackURI) -> Result<Conformance> {
        let owner = self.working.get_part(slide)?;
        Ok(
            match crate::media_parts::document_conformance(owner.blob())? {
                crate::media_parts::Conformance::Transitional => Conformance::Transitional,
                crate::media_parts::Conformance::Strict => Conformance::Strict,
            },
        )
    }

    /// Append a legacy comment to a slide, written by the comment author
    /// named `author` (matched exactly), which is created with `initials`
    /// when the presentation has none of that name. The comment takes the
    /// author's next index; `date_time` is an optional `xsd:dateTime`
    /// stamp, supplied by the caller (the editor consults no clock).
    /// Returns the comment's position in the slide's comment list.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown slide, an author name shared by
    /// several authors, exhausted author identifiers or indexes, a signed
    /// package, or malformed comment parts. A failure leaves the edit
    /// unchanged.
    pub fn add_comment_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author: &str,
        initials: &str,
        text: &str,
        location: (i64, i64),
        date_time: Option<&str>,
    ) -> Result<Position> {
        let selected = self.resolve_slide(slide.into())?;
        let conformance = self.comment_conformance(&selected.part_name)?;
        let graph = crate::comments::load_presentation_comments_strict(&self.working)?;
        let (authors, existing) = graph.as_ref().map_or((&[][..], 0), |graph| {
            let existing = graph
                .slides
                .iter()
                .find(|list| list.slide_part_name == selected.part_name.as_str())
                .map_or(0, |list| list.comments.len());
            (graph.authors.as_slice(), existing)
        });
        let mut named = authors.iter().filter(|candidate| candidate.name == author);
        let found = named.next().cloned();
        if named.next().is_some() {
            return Err(Error::UnsafeEdit {
                operation: "add_comment_text",
                reason: "several comment authors share the requested name",
            });
        }
        let before = self.working.clone();
        let result = (|| {
            let author = match found {
                Some(found) => found,
                None => {
                    let id = authors
                        .iter()
                        .map(|candidate| candidate.id)
                        .max()
                        .map_or(Some(0), |maximum| maximum.checked_add(1))
                        .ok_or_else(|| {
                            Error::Invalid("comment author identifiers are exhausted".into())
                        })?;
                    let created = Author::new(id, author, initials);
                    crate::comments::add_presentation_comment_author(
                        &mut self.working,
                        created.clone(),
                        conformance,
                    )?;
                    created
                },
            };
            let index = author
                .last_index
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("comment indexes are exhausted".into()))?;
            let mut comment = Comment::new(author.id, text, location.0, location.1);
            comment.index = index;
            if let Some(date_time) = date_time {
                comment = comment.with_date_time(date_time);
            }
            crate::comments::add_presentation_comment(
                &mut self.working,
                selected.part_name.as_str(),
                comment,
                conformance,
            )
        })();
        if let Err(error) = result {
            self.working = before;
            return Err(error);
        }
        Ok(Position::new(existing))
    }

    /// Replace the text and/or location of the comment at `position` in the
    /// slide's comment list. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment and the
    /// errors of the comment codec. A failure leaves the edit unchanged.
    pub fn set_comment_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        position: Position,
        text: Option<&str>,
        location: Option<(i64, i64)>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let conformance = self.comment_conformance(&selected.part_name)?;
        let graph = crate::comments::load_presentation_comments_strict(&self.working)?;
        let comments = graph
            .as_ref()
            .and_then(|graph| {
                graph
                    .slides
                    .iter()
                    .find(|list| list.slide_part_name == selected.part_name.as_str())
            })
            .map_or(&[][..], |list| list.comments.as_slice());
        let current = comments
            .get(position.get())
            .ok_or(Error::IndexOutOfBounds {
                index: position.get(),
                len: comments.len(),
            })?;
        let mut replacement = current.clone();
        if let Some(text) = text {
            replacement.text = text.to_owned();
        }
        if let Some((x, y)) = location {
            replacement.x = x;
            replacement.y = y;
        }
        if &replacement == current {
            return Ok(false);
        }
        let before = self.working.clone();
        if let Err(error) = crate::comments::update_presentation_comment(
            &mut self.working,
            selected.part_name.as_str(),
            current.author_id,
            current.index,
            replacement,
            conformance,
        ) {
            self.working = before;
            return Err(error);
        }
        Ok(true)
    }
}
