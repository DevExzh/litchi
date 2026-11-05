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
        let mut reader = crate::raw::Reader::from_reader_with_options(
            Cursor::new(self.working.bytes()),
            self.working.open_options(),
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
        let authors = crate::comments::Authors::parse(
            &presentation.live_document_record()?,
            &presentation.source_limits(),
        )?;
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
    pub(crate) fn add_comment(
        &mut self,
        position: Position,
        comment: &NewComment,
    ) -> Result<usize> {
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
    pub(crate) fn set_comment(
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
    pub(crate) fn remove_comment(&mut self, position: Position, comment: Position) -> Result<()> {
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
    pub(crate) fn set_table_cell_text(
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

impl Transaction {
    /// The working artifact's slide at a projected position, its shape at a
    /// source-order index, the shape's `OfficeArt` ID, and the shape's
    /// current click hyperlink.
    fn shape_click(
        &self,
        target: crate::text_edit::Target,
    ) -> Result<(u32, Option<crate::hyperlink::Click>)> {
        let presentation = self.working_presentation()?;
        let projected = self.document.slides()?;
        let persist_id = projected
            .get(target.slide().get())
            .ok_or(Error::Refused(Refusal::SlideNotFound {
                position: target.slide(),
            }))?
            .persist_id();
        let slide_ids = projected
            .iter()
            .map(|slide| slide.slide_id())
            .collect::<Vec<_>>();
        let slide = presentation
            .slides()?
            .into_iter()
            .find(|slide| slide.persist_id() == persist_id)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let shapes = slide.shapes()?;
        let shape =
            shapes
                .get(target.shape().get())
                .ok_or(Error::Refused(Refusal::ElementNotFound {
                    element: "shape",
                }))?;
        let shape_id = crate::text_edit::native_shape_id(shape);
        let hyperlinks = presentation.hyperlinks()?;
        let current = slide
            .shape_interactions()?
            .into_iter()
            .find(|entry| entry.shape_id == shape_id)
            .and_then(|entry| {
                entry.interactions.iter().find_map(|interaction| {
                    crate::hyperlink::Click::resolve(interaction, &hyperlinks, &slide_ids)
                })
            });
        Ok((shape_id, current))
    }

    /// Set (`Some`) or remove (`None`) the click hyperlink of the top-level
    /// shape at `target` (record 0323).
    ///
    /// A URL or slide link appends an `ExHyperlinkContainer` to the live
    /// document's external-object list (raising its identifier seed) and,
    /// with a tooltip, an `ExHyperlink9Container` screen tip to its
    /// `___PPT9` tag; a jump needs no definition. Definitions a replaced or
    /// removed link leaves unused are kept, as `PowerPoint` keeps them. An
    /// equal link is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide or shape, a slide target
    /// outside the presentation, an empty URL, ambiguous client data, or a
    /// changed signed source.
    pub(crate) fn set_shape_hyperlink(
        &mut self,
        target: crate::text_edit::Target,
        click: Option<&crate::hyperlink::Click>,
    ) -> Result<()> {
        use crate::hyperlink::Target as LinkTarget;
        use crate::hyperlink::edit::{
            Definition, add_definition, click_atom, click_container, set_shape_click,
        };

        self.require_position(target.slide())?;
        let (shape_id, current) = self.shape_click(target)?;
        if current.as_ref() == click {
            return Ok(());
        }
        let invalid = |reason| {
            Error::Refused(Refusal::InvalidElementValue {
                element: "hyperlink",
                reason,
            })
        };
        let slide_location = match click.map(|click| &click.target) {
            Some(LinkTarget::Url(url)) if url.is_empty() || url.contains('\0') => {
                return Err(invalid("the URL is empty or holds a NUL character"));
            },
            Some(LinkTarget::Slide(position)) => {
                let slide_id = self
                    .document
                    .slides()?
                    .get(position.get())
                    .ok_or(invalid("the target slide is outside the presentation"))?
                    .slide_id();
                let number = position.get() + 1;
                Some((
                    format!("Slide {number}"),
                    format!("{slide_id},{number},Slide {number}"),
                ))
            },
            Some(LinkTarget::Jump(
                crate::hyperlink::InteractionJump::None
                | crate::hyperlink::InteractionJump::LastSlideViewed,
            )) => return Err(invalid("the jump is not a hyperlink target")),
            _ => None,
        };
        let needs_definition = matches!(
            click.map(|click| &click.target),
            Some(LinkTarget::Url(_) | LinkTarget::Slide(_))
        );
        if click.is_some_and(|click| click.tooltip.is_some()) && !needs_definition {
            return Err(invalid("a slide-show jump has no screen tip"));
        }
        let source_limits = self.source.limits();
        self.stage_records(Some(target.slide()), needs_definition, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            let Some(click) = click else {
                return Ok(Replacement {
                    slide: set_shape_click(slide.bytes, shape_id, None, &source_limits)?,
                    document: None,
                });
            };
            let (document, id) = match (&click.target, view.document) {
                (LinkTarget::Url(url), Some(document)) => {
                    let (record, id) = add_definition(
                        document,
                        &Definition {
                            friendly_name: url,
                            target: Some(url),
                            location: None,
                            tooltip: click.tooltip.as_deref(),
                        },
                        &source_limits,
                    )?;
                    (Some(record), id)
                },
                (LinkTarget::Slide(_), Some(document)) => {
                    let (friendly, location) =
                        slide_location.as_ref().ok_or_else(missing_record)?;
                    let (record, id) = add_definition(
                        document,
                        &Definition {
                            friendly_name: friendly,
                            target: None,
                            location: Some(location),
                            tooltip: click.tooltip.as_deref(),
                        },
                        &source_limits,
                    )?;
                    (Some(record), id)
                },
                _ => (None, 0),
            };
            let container = click_container(&click_atom(&click.target, id), &source_limits)?;
            Ok(Replacement {
                slide: set_shape_click(slide.bytes, shape_id, Some(&container), &source_limits)?,
                document,
            })
        })?;
        if self.shape_click(target)?.1.as_ref() != click {
            return Err(PackageError::Corrupted(
                "published PPT shape hyperlink did not read back".into(),
            )
            .into());
        }
        Ok(())
    }
}

impl Transaction {
    /// The persist ID and the `OfficeArt` shape ID of the top-level shape at
    /// `target`, and the shape's text runs in the working artifact.
    fn shape_runs(
        &self,
        target: crate::text_edit::Target,
    ) -> Result<(
        u32,
        u32,
        Option<crate::text_edit::TextRuns>,
        crate::Presentation,
    )> {
        let presentation = self.working_presentation()?;
        let persist_id = self
            .document
            .slides()?
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
        let shape_id = shapes
            .get(target.shape().get())
            .map(crate::text_edit::native_shape_id)
            .ok_or(Error::Refused(Refusal::ElementNotFound {
                element: "shape",
            }))?;
        let runs = slide.shape_text_runs(shape_id)?;
        drop(slide);
        Ok((persist_id, shape_id, runs, presentation))
    }

    /// Apply character formatting to the UTF-16 units `start..start +
    /// length` of the text of the top-level shape at `target` (record 0323).
    ///
    /// The body's `StyleTextPropAtom` runs are split at the range ends and
    /// the format applied to the runs inside; a body without a style atom
    /// gains one. An outline-backed placeholder's slide-list body is edited
    /// in the live document. A range that already shows the format is an
    /// exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide or shape, a shape without
    /// editable text, a range outside its text, an unrepresentable size, a
    /// style atom the run model does not reproduce byte for byte, or a
    /// changed signed source.
    pub(crate) fn set_run_format(
        &mut self,
        target: crate::text_edit::Target,
        start: usize,
        length: usize,
        format: &crate::text_edit::RunFormat,
    ) -> Result<()> {
        self.require_position(target.slide())?;
        if format.size.is_some_and(|size| !(1..=4000).contains(&size)) {
            return Err(Error::Refused(Refusal::InvalidElementValue {
                element: "run",
                reason: "a font size is 1 to 4000 points",
            }));
        }
        let (persist_id, shape_id, runs, _) = self.shape_runs(target)?;
        let runs = runs.ok_or(Error::Refused(Refusal::ElementNotFound { element: "run" }))?;
        let end = start
            .checked_add(length)
            .filter(|end| length > 0 && *end <= runs.text().encode_utf16().count() + 1)
            .ok_or(Error::Refused(Refusal::InvalidElementValue {
                element: "run",
                reason: "the range is outside the shape's text",
            }))?;
        let interaction_limits =
            crate::text_edit::runs::interaction_limits_for_source(self.source.limits());
        self.stage_records(Some(target.slide()), true, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            let document = view.document.ok_or_else(missing_record)?;
            let (slide, document) = crate::text_edit::runs::rewrite(
                slide.bytes,
                document,
                persist_id,
                shape_id,
                &|records| {
                    crate::text_edit::runs::format_body(
                        records,
                        start,
                        end,
                        format,
                        interaction_limits,
                    )
                },
            )?;
            Ok(Replacement { slide, document })
        })
        .map(drop)
    }

    /// Set (`Some`) or remove (`None`) the click hyperlink over exactly the
    /// UTF-16 units `start..start + length` of the text of the top-level
    /// shape at `target` (record 0323), as an `InteractiveInfo` and
    /// `TextInteractiveInfoAtom` pair; definitions follow
    /// [`Self::set_shape_hyperlink`]. An equal link is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent slide or shape, a range outside
    /// the text or overlapping another link, an invalid target, or a changed
    /// signed source.
    pub(crate) fn set_text_hyperlink(
        &mut self,
        target: crate::text_edit::Target,
        start: usize,
        length: usize,
        click: Option<&crate::hyperlink::Click>,
    ) -> Result<()> {
        use crate::hyperlink::Target as LinkTarget;
        use crate::hyperlink::edit::{Definition, add_definition, click_atom, click_container};

        self.require_position(target.slide())?;
        let (persist_id, shape_id, runs, presentation) = self.shape_runs(target)?;
        let runs = runs.ok_or(Error::Refused(Refusal::ElementNotFound { element: "run" }))?;
        let end = start
            .checked_add(length)
            .filter(|end| length > 0 && *end <= runs.text().encode_utf16().count() + 1)
            .ok_or(Error::Refused(Refusal::InvalidElementValue {
                element: "hyperlink",
                reason: "the range is outside the shape's text",
            }))?;
        let current = runs
            .links()
            .iter()
            .filter(|link| link.start == start && link.end == end)
            .find_map(|link| presentation.resolve_click(&link.interaction).transpose())
            .transpose()?;
        if current.as_ref() == click {
            return Ok(());
        }
        let invalid = |reason| {
            Error::Refused(Refusal::InvalidElementValue {
                element: "hyperlink",
                reason,
            })
        };
        let slide_location = match click.map(|click| &click.target) {
            Some(LinkTarget::Url(url)) if url.is_empty() || url.contains('\0') => {
                return Err(invalid("the URL is empty or holds a NUL character"));
            },
            Some(LinkTarget::Slide(position)) => {
                let slide_id = self
                    .document
                    .slides()?
                    .get(position.get())
                    .ok_or(invalid("the target slide is outside the presentation"))?
                    .slide_id();
                let number = position.get() + 1;
                Some((
                    format!("Slide {number}"),
                    format!("{slide_id},{number},Slide {number}"),
                ))
            },
            Some(LinkTarget::Jump(
                crate::hyperlink::InteractionJump::None
                | crate::hyperlink::InteractionJump::LastSlideViewed,
            )) => return Err(invalid("the jump is not a hyperlink target")),
            _ => None,
        };
        if click.is_some_and(|click| {
            click.tooltip.is_some()
                && !matches!(click.target, LinkTarget::Url(_) | LinkTarget::Slide(_))
        }) {
            return Err(invalid("a slide-show jump has no screen tip"));
        }
        let source_limits = self.source.limits();
        self.stage_records(Some(target.slide()), true, |view| {
            let slide = view.slide.ok_or_else(missing_record)?;
            let mut document = view.document.ok_or_else(missing_record)?.to_vec();
            let mut document_changed = false;
            let container = match click {
                None => None,
                Some(click) => {
                    let definition = match &click.target {
                        LinkTarget::Url(url) => Some(Definition {
                            friendly_name: url,
                            target: Some(url),
                            location: None,
                            tooltip: click.tooltip.as_deref(),
                        }),
                        LinkTarget::Slide(_) => {
                            let (friendly, location) =
                                slide_location.as_ref().ok_or_else(missing_record)?;
                            Some(Definition {
                                friendly_name: friendly,
                                target: None,
                                location: Some(location),
                                tooltip: click.tooltip.as_deref(),
                            })
                        },
                        LinkTarget::Jump(_) => None,
                    };
                    let id = match definition {
                        Some(definition) => {
                            let (record, id) =
                                add_definition(&document, &definition, &source_limits)?;
                            document = record;
                            document_changed = true;
                            id
                        },
                        None => 0,
                    };
                    Some(click_container(
                        &click_atom(&click.target, id),
                        &source_limits,
                    )?)
                },
            };
            let (slide, outline) = crate::text_edit::runs::rewrite(
                slide.bytes,
                &document,
                persist_id,
                shape_id,
                &|records| {
                    crate::text_edit::runs::link_body(records, start, end, container.as_deref())
                },
            )?;
            let document = match outline {
                Some(outline) => Some(outline),
                None => document_changed.then_some(document),
            };
            Ok(Replacement { slide, document })
        })?;
        let (_, _, runs, presentation) = self.shape_runs(target)?;
        let published = runs
            .ok_or_else(missing_record)?
            .links()
            .iter()
            .filter(|link| link.start == start && link.end == end)
            .find_map(|link| presentation.resolve_click(&link.interaction).transpose())
            .transpose()?;
        if published.as_ref() != click {
            return Err(PackageError::Corrupted(
                "published PPT text hyperlink did not read back".into(),
            )
            .into());
        }
        Ok(())
    }
}
