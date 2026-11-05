//! Positional legacy- and modern-comment verbs for the unified API (records
//! 0252 and 0333).

use litchi_core::Position;

use super::Edit;
use crate::comments::{Author, Comment, Conformance};
use crate::{Error, Result};

impl Edit {
    fn comment_conformance(&self, slide: &litchi_opc::PackURI) -> Result<Conformance> {
        let owner = self.working.get_part(slide)?;
        Ok(
            match crate::media_parts::document_conformance_with_limits(
                owner.blob(),
                self.semantic_limits()
                    .media_parts()
                    .constrained_by_source(self.working.read_limits()),
            )? {
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
        let graph = crate::comments::load_presentation_comments_strict_with_limits(
            &self.working,
            self.comments_limits(),
        )?;
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
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "add_comment_text",
                "several comment authors share the requested name",
            ));
        }
        let limits = self.comments_limits();
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
                    crate::comments::add_presentation_comment_author_with_limits(
                        &mut self.working,
                        created.clone(),
                        conformance,
                        limits,
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
            crate::comments::add_presentation_comment_with_limits(
                &mut self.working,
                selected.part_name.as_str(),
                comment,
                conformance,
                limits,
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
        let graph = crate::comments::load_presentation_comments_strict_with_limits(
            &self.working,
            self.comments_limits(),
        )?;
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
        let limits = self.comments_limits();
        if let Err(error) = crate::comments::update_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            current.author_id,
            current.index,
            replacement,
            conformance,
            limits,
        ) {
            self.working = before;
            return Err(error);
        }
        Ok(true)
    }

    /// Replace the `xsd:dateTime` stamp of the comment at `position` in the
    /// slide's comment list (record 0346). Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment and the
    /// errors of the comment codec. A failure leaves the edit unchanged.
    pub fn set_comment_date_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        position: Position,
        date_time: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let conformance = self.comment_conformance(&selected.part_name)?;
        let current = self.legacy_comment(&selected.part_name, position)?;
        if current.date_time.as_deref() == Some(date_time) {
            return Ok(false);
        }
        let mut replacement = current.clone();
        replacement.date_time = Some(date_time.to_owned());
        let before = self.working.clone();
        let limits = self.comments_limits();
        if let Err(error) = crate::comments::update_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            current.author_id,
            current.index,
            replacement,
            conformance,
            limits,
        ) {
            self.working = before;
            return Err(error);
        }
        Ok(true)
    }

    /// Re-attribute the comment at `position` in the slide's comment list to
    /// the author `name` with `initials` (record 0346); `None` keeps the
    /// current author's value. The author is renamed in place when no other
    /// comment of the presentation is theirs; otherwise the comment moves to
    /// the author of that name and initials (created when absent) and takes
    /// that author's next index, so the other comments keep their
    /// attribution. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment, an error
    /// for exhausted author identifiers or indexes or a signed package, and
    /// the errors of the comment codec. A failure leaves the edit unchanged.
    pub fn set_comment_author_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        position: Position,
        name: Option<&str>,
        initials: Option<&str>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let conformance = self.comment_conformance(&selected.part_name)?;
        let current = self.legacy_comment(&selected.part_name, position)?;
        let before = self.working.clone();
        let limits = self.comments_limits();
        match crate::comments::reattribute_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            current.author_id,
            current.index,
            name,
            initials,
            conformance,
            limits,
        ) {
            Ok(changed) => Ok(changed),
            Err(error) => {
                self.working = before;
                Err(error)
            },
        }
    }

    /// The legacy comment at `position` of a slide part.
    fn legacy_comment(&self, slide: &litchi_opc::PackURI, position: Position) -> Result<Comment> {
        let graph = crate::comments::load_presentation_comments_strict_with_limits(
            &self.working,
            self.comments_limits(),
        )?;
        let comments = graph
            .as_ref()
            .and_then(|graph| {
                graph
                    .slides
                    .iter()
                    .find(|list| list.slide_part_name == slide.as_str())
            })
            .map_or(&[][..], |list| list.comments.as_slice());
        comments
            .get(position.get())
            .cloned()
            .ok_or(Error::IndexOutOfBounds {
                index: position.get(),
                len: comments.len(),
            })
    }

    /// Remove the comment at `position` in the slide's legacy comment list
    /// (record 0333). The slide's comments part and its relationship go
    /// when the last comment does; the comment authors are kept, since
    /// other slides may name them, and so is each author's last index, so
    /// a later comment takes a fresh index. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment, an error
    /// for a signed package, and the errors of the comment codec. A failure
    /// leaves the edit unchanged.
    pub fn remove_comment_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        position: Position,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let conformance = self.comment_conformance(&selected.part_name)?;
        let graph = crate::comments::load_presentation_comments_strict_with_limits(
            &self.working,
            self.comments_limits(),
        )?;
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
        let before = self.working.clone();
        let limits = self.comments_limits();
        match crate::comments::remove_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            current.author_id,
            current.index,
            conformance,
            limits,
        ) {
            Ok(changed) => Ok(changed),
            Err(error) => {
                self.working = before;
                Err(error)
            },
        }
    }

    /// Remove the modern (threaded) comment at `position` in the slide's
    /// modern comment list, with its replies (record 0333). The slide's
    /// comment part goes with its last comment; authors are kept.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment, an error
    /// for a signed package, and the errors of the modern-comment owner.
    pub fn remove_modern_comment_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        position: Position,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let id = self.modern_comment_id(&selected.part_name, position)?;
        let mut candidate = self.working.clone();
        let changed = crate::modern_comments::remove_modern_comment(
            &mut candidate,
            &selected.part_name,
            &id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(changed)
    }

    /// Remove the reply at `reply` of the modern comment at `comment` in the
    /// slide's modern comment list (record 0333).
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing comment or reply,
    /// an error for a signed package, and the errors of the modern-comment
    /// owner.
    pub fn remove_modern_comment_reply_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: Position,
        reply: Position,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let graph = crate::modern_comments::load_modern_comment_graph(&self.working)?;
        let comments = modern_comments_of(&graph, &selected.part_name);
        let record = comments.get(comment.get()).ok_or(Error::IndexOutOfBounds {
            index: comment.get(),
            len: comments.len(),
        })?;
        let target = record
            .replies
            .get(reply.get())
            .ok_or(Error::IndexOutOfBounds {
                index: reply.get(),
                len: record.replies.len(),
            })?;
        let (comment_id, reply_id) = (record.id.clone(), target.id.clone());
        let mut candidate = self.working.clone();
        let changed = crate::modern_comments::remove_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            &comment_id,
            &reply_id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(changed)
    }

    /// The stable identifier of the modern comment at `position` of a slide.
    fn modern_comment_id(&self, slide: &litchi_opc::PackURI, position: Position) -> Result<String> {
        let graph = crate::modern_comments::load_modern_comment_graph(&self.working)?;
        let comments = modern_comments_of(&graph, slide);
        comments
            .get(position.get())
            .map(|comment| comment.id.clone())
            .ok_or(Error::IndexOutOfBounds {
                index: position.get(),
                len: comments.len(),
            })
    }
}

/// The modern comments of one slide part, in document order.
fn modern_comments_of<'g>(
    graph: &'g crate::modern_comments::Graph,
    slide: &litchi_opc::PackURI,
) -> &'g [crate::modern_comments::Comment] {
    graph
        .comments
        .iter()
        .find(|part| part.slide_part_name == slide.as_str())
        .map_or(&[][..], |part| part.comments.comments.as_slice())
}
