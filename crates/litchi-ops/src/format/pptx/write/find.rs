//! PPTX find/replace (records 0273, 0285): `set` with `find=`.
//!
//! The scope is the target's subtree, as OfficeCLI's find reads it: the
//! presentation searches every slide and its speaker notes; a slide its
//! shapes (inside groups too) and its tables' cells; a shape or placeholder
//! its text body; a table, row or cell its cells; the notes page its shapes;
//! a paragraph itself; and a run or line break only its own text. A
//! paragraph is one search group; line breaks and fields separate text.
//! Matches become run text edits (`TextTarget::Run`, or `CellRun` in a
//! table cell), and format properties become run-span edits (`RunSpan`,
//! `CellRunSpan`) on exactly the replaced or matched text, all through one
//! native edit. The run rewrite is the one template merge uses
//! (`find::rewrite_ranges`, record 0285).
//!
//! Masters and layouts are not searched yet.

use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_drawingml::text::body::{Body, RunKind};
use litchi_pptx::opened::{TextAssignment, TextTarget};

use super::text::{TextValues, admit};
use super::{Commit, Pptx, allocation, error};
use crate::find::{Piece, Scratch, Spec, plan_with};
use crate::format::pptx::Nav;
use crate::format::pptx::node::{Item, TableItem, TreeItem};
use crate::format::pptx::spec::{Kind, RUN_WRITE_TABLE, TextSetter};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::{Code, Error, Result};

/// Where a searched text body lives.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum Place {
    /// A slide's shape tree.
    Slide,
    /// A slide's notes page.
    Notes,
}

/// Where one text body is: its slide or notes page, its shape (positional
/// scene index) and, for a table, its cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Locus {
    pub(super) place: Place,
    pub(super) slide: usize,
    pub(super) shape: usize,
    pub(super) cell: Option<(usize, usize)>,
}

/// One text body to search, optionally narrowed to a paragraph or one run.
pub(super) struct Scope {
    pub(super) locus: Locus,
    pub(super) body: Arc<Body>,
    paragraph: Option<usize>,
    run: Option<usize>,
}

/// A planned change to one run: its new text and the spans to format.
pub(super) struct RunChange {
    pub(super) locus: Locus,
    pub(super) paragraph: usize,
    pub(super) run: usize,
    pub(super) text: Option<String>,
    pub(super) marks: Vec<std::ops::Range<usize>>,
}

impl RunChange {
    pub(super) fn target(&self, span: Option<&std::ops::Range<usize>>) -> TextTarget {
        let shape = Position::new(self.locus.shape);
        let paragraph = Position::new(self.paragraph);
        let run = Position::new(self.run);
        match (self.locus.cell, span) {
            (None, None) => TextTarget::Run {
                shape,
                paragraph,
                run,
            },
            (None, Some(span)) => TextTarget::RunSpan {
                shape,
                paragraph,
                run,
                start: span.start,
                end: span.end,
            },
            (Some((row, column)), None) => TextTarget::CellRun {
                shape,
                row: Position::new(row),
                column: Position::new(column),
                paragraph,
                run,
            },
            (Some((row, column)), Some(span)) => TextTarget::CellRunSpan {
                shape,
                row: Position::new(row),
                column: Position::new(column),
                paragraph,
                run,
                start: span.start,
                end: span.end,
            },
        }
    }
}

impl Pptx {
    /// Find (and replace or format) text in the subtree of `target`,
    /// returning the commit and the match count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        self.check_source()?;
        let result = self
            .find_inner(target, spec, limits)
            .map_err(|error| error.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }

    fn find_inner(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        let format = if spec.formats() {
            let values = admit(&RUN_WRITE_TABLE, spec.format(), limits)?;
            if values.has(TextSetter::Text) {
                return Err(error(
                    Code::InvalidArgument,
                    "`text` replaces a whole run; with `find`, use `replace`",
                )
                .with_property("text"));
            }
            Some(values)
        } else {
            None
        };
        let matcher = spec.matcher();
        let scopes = self.find_scopes(target, &|text: &str| matcher.may_match(text))?;
        let mut matches = 0usize;
        let mut staged = 0usize;
        let mut changes: Vec<RunChange> = Vec::new();
        let mut scratch = Scratch::default();
        for scope in &scopes {
            let mut texts: Vec<(usize, usize)> = Vec::new();
            let mut pieces: Vec<Piece<'_>> = Vec::new();
            let mut group = 0;
            for (index, paragraph) in scope.body.paragraphs.iter().enumerate() {
                if scope.paragraph.is_some_and(|wanted| wanted != index) {
                    continue;
                }
                group += 1;
                for (position, run) in paragraph.runs.iter().enumerate() {
                    if run.kind != RunKind::Text {
                        group += 1;
                        continue;
                    }
                    pieces.try_reserve(1).map_err(allocation)?;
                    texts.try_reserve(1).map_err(allocation)?;
                    pieces.push(Piece {
                        text: run.text.as_str(),
                        group,
                        scoped: scope.run.is_none_or(|wanted| wanted == position),
                    });
                    texts.push((index, position));
                }
            }
            let planned = plan_with(
                &mut scratch,
                &pieces,
                &matcher,
                spec.replacement(),
                format.is_some(),
                limits.max_staged_bytes().saturating_sub(staged),
            )?;
            matches = matches.saturating_add(planned.matches);
            for rewrite in planned.rewrites {
                let (paragraph, run) = texts[rewrite.piece];
                staged = staged.saturating_add(rewrite.text.as_ref().map_or(0, String::len));
                changes.try_reserve(1).map_err(allocation)?;
                changes.push(RunChange {
                    locus: scope.locus,
                    paragraph,
                    run,
                    text: rewrite.text,
                    marks: rewrite.marks,
                });
            }
        }
        // Splits shift the runs after them: edit each paragraph from its
        // last run back, a run's text before its spans, spans from the last.
        changes.sort_by(|left, right| {
            (left.locus, left.paragraph, left.run)
                .cmp(&(right.locus, right.paragraph, right.run))
                .reverse()
        });
        let properties = format.as_ref().map(TextValues::properties);
        let mut slides: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
        let mut notes: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
        for change in &changes {
            let groups = match change.locus.place {
                Place::Slide => &mut slides,
                Place::Notes => &mut notes,
            };
            if groups
                .last()
                .is_none_or(|(slide, _)| *slide != change.locus.slide)
            {
                groups.try_reserve(1).map_err(allocation)?;
                groups.push((change.locus.slide, Vec::new()));
            }
            let Some((_, edits)) = groups.last_mut() else {
                unreachable!("a group was just pushed")
            };
            if let Some(text) = &change.text {
                edits.try_reserve(1).map_err(allocation)?;
                edits.push(TextAssignment::new(
                    change.target(None),
                    litchi_pptx::opened::TextProperties::default().with_text(text),
                ));
            }
            if let Some(properties) = properties {
                for mark in change.marks.iter().rev() {
                    edits.try_reserve(1).map_err(allocation)?;
                    edits.push(TextAssignment::new(change.target(Some(mark)), properties));
                }
            }
        }
        slides.retain(|(_, edits)| !edits.is_empty());
        notes.retain(|(_, edits)| !edits.is_empty());
        Ok((
            self.commit_story_text_edits(&slides, &notes, limits)?,
            matches,
        ))
    }

    /// The text bodies the target nodes search, each once; `may_match`
    /// skips a body whose plain text cannot hold a match.
    pub(super) fn find_scopes(
        &self,
        target: &TargetArg<'_>,
        may_match: &impl Fn(&str) -> bool,
    ) -> Result<Vec<Scope>> {
        let nav = Nav::new(self);
        let mut scopes = Vec::new();
        if let Some(path) = target.path() {
            self.collect(
                &nav,
                crate::path::resolve::get(&nav, path)?,
                may_match,
                &mut scopes,
            )?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(&nav, query)?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(error(
                    Code::InvalidSelector,
                    "mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                self.collect(&nav, node?, may_match, &mut scopes)?;
            }
        } else {
            return Err(error(Code::InvalidArgument, "mutation target is empty"));
        }
        // A body reached twice (a query naming a shape and its paragraph)
        // searches once, with the widest scope.
        scopes.sort_by_key(|scope| scope.locus);
        let mut merged: Vec<Scope> = Vec::with_capacity(scopes.len());
        for scope in scopes {
            match merged.last_mut() {
                Some(last) if last.locus == scope.locus => {
                    if last.paragraph != scope.paragraph || last.run != scope.run {
                        if last.paragraph == scope.paragraph {
                            last.run = None;
                        } else {
                            last.paragraph = None;
                            last.run = None;
                        }
                    }
                },
                _ => merged.push(scope),
            }
        }
        Ok(merged)
    }

    /// Every text body of the presentation: each slide's shapes, groups and
    /// table cells, then its notes page (the bodies template merge fills,
    /// record 0285).
    pub(super) fn presentation_scopes(
        &self,
        may_match: &impl Fn(&str) -> bool,
    ) -> Result<Vec<Scope>> {
        let nav = Nav::new(self);
        let mut scopes = Vec::new();
        self.collect(&nav, nav.root()?, may_match, &mut scopes)?;
        Ok(scopes)
    }

    fn collect(
        &self,
        nav: &Nav<'_>,
        node: crate::format::pptx::Node<'_>,
        may_match: &impl Fn(&str) -> bool,
        scopes: &mut Vec<Scope>,
    ) -> Result<()> {
        let push = |tree: &TreeItem,
                    paragraph: Option<usize>,
                    run: Option<usize>,
                    scopes: &mut Vec<Scope>|
         -> Result<()> {
            let slide = slide_of(tree)?;
            let Some(body) = tree.body()? else {
                return Ok(());
            };
            scopes.try_reserve(1).map_err(allocation)?;
            scopes.push(Scope {
                locus: Locus {
                    place: Place::Slide,
                    slide,
                    shape: tree.handle.position().get(),
                    cell: None,
                },
                body: Arc::clone(body),
                paragraph,
                run,
            });
            Ok(())
        };
        match node.item() {
            Item::Presentation => {
                for child in nav.children(&node, KindSet::of(Kind::Slide))? {
                    let child = child?;
                    if let Item::Slide(slide) = child.item() {
                        self.notes_scopes(slide.position as usize, may_match, scopes)?;
                    }
                    self.collect(nav, child, may_match, scopes)?;
                }
                Ok(())
            },
            Item::Slide(_) => {
                let kinds = KindSet::of(Kind::Shape)
                    .with(Kind::Placeholder)
                    .with(Kind::Group)
                    .with(Kind::Table);
                for child in nav.children(&node, kinds)? {
                    self.collect(nav, child?, may_match, scopes)?;
                }
                Ok(())
            },
            Item::Notes(notes) => self.notes_scopes(notes.slide as usize, may_match, scopes),
            Item::Tree(tree) => match tree.kind {
                Kind::Shape | Kind::Placeholder => {
                    // The scene's plain text (runs, fields and breaks joined)
                    // proves the absence of a match without decoding the body.
                    if tree
                        .handle
                        .view()
                        .text()
                        .is_some_and(|text| !may_match(text))
                    {
                        return Ok(());
                    }
                    push(tree, None, None, scopes)
                },
                Kind::Group => {
                    let kinds = KindSet::of(Kind::Shape)
                        .with(Kind::Placeholder)
                        .with(Kind::Group)
                        .with(Kind::Table);
                    for child in nav.children(&node, kinds)? {
                        self.collect(nav, child?, may_match, scopes)?;
                    }
                    Ok(())
                },
                Kind::Table => cell_scopes(tree, None, None, scopes),
                // Pictures, charts, media and the like carry no searchable text.
                _ => Ok(()),
            },
            Item::Row(TableItem { table, row, .. }) => {
                cell_scopes(table, Some(*row as usize), None, scopes)
            },
            Item::Cell(TableItem { table, row, column }) => cell_scopes(
                table,
                Some(*row as usize),
                Some(*column as usize),
                scopes,
            ),
            Item::Paragraph(text) => text_scope(text, None, scopes),
            Item::Run(text) | Item::LineBreak(text) => {
                text_scope(text, Some(text.run as usize), scopes)
            },
            _ => Err(error(
                Code::UnsupportedCapability,
                "find searches the presentation, slides, notes, shapes, groups, tables, rows, cells, paragraphs and runs",
            )
            .with_property(crate::find::FIND)
            .with_suggestion("comments, masters and layouts are not searched yet (record 0285)")),
        }
    }

    /// The shapes of one slide's notes page.
    fn notes_scopes(
        &self,
        slide: usize,
        may_match: &impl Fn(&str) -> bool,
        scopes: &mut Vec<Scope>,
    ) -> Result<()> {
        let Some(scene) = self.slide_notes_scene(slide)? else {
            return Ok(());
        };
        for handle in scene.iter() {
            let view = handle.view();
            if view.text().is_none_or(|text| !may_match(text)) {
                continue;
            }
            let Some(body) = view.text_body()? else {
                continue;
            };
            scopes.try_reserve(1).map_err(allocation)?;
            scopes.push(Scope {
                locus: Locus {
                    place: Place::Notes,
                    slide,
                    shape: handle.position().get(),
                    cell: None,
                },
                body: Arc::new(body),
                paragraph: None,
                run: None,
            });
        }
        Ok(())
    }
}

/// The body of a paragraph, run or line break, narrowed to it: a slide
/// shape's, a table cell's or the notes page's (record 0333).
fn text_scope(
    text: &crate::format::pptx::node::TextItem,
    run: Option<usize>,
    scopes: &mut Vec<Scope>,
) -> Result<()> {
    let (place, slide) = match text.shape.container.notes() {
        Some(slide) => (Place::Notes, slide as usize),
        None => (Place::Slide, slide_of(&text.shape)?),
    };
    scopes.try_reserve(1).map_err(allocation)?;
    scopes.push(Scope {
        locus: Locus {
            place,
            slide,
            shape: text.shape.handle.position().get(),
            cell: text
                .cell
                .get()
                .map(|(row, column)| (row as usize, column as usize)),
        },
        body: Arc::clone(&text.body),
        paragraph: Some(text.paragraph as usize),
        run,
    });
    Ok(())
}

/// The slide of a shape; masters and layouts are not searched.
fn slide_of(tree: &TreeItem) -> Result<usize> {
    tree.container
        .slide()
        .map(|slide| slide as usize)
        .ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "find does not search slide masters and layouts yet",
            )
            .with_property(crate::find::FIND)
        })
}

/// The cells of a table frame (one row, or one cell, when given), each a
/// text body run edits address (record 0285).
fn cell_scopes(
    tree: &TreeItem,
    only_row: Option<usize>,
    only_column: Option<usize>,
    scopes: &mut Vec<Scope>,
) -> Result<()> {
    let slide = slide_of(tree)?;
    let table = tree.handle.table().map_err(Error::from)?;
    for (row_index, row) in table.rows().map_err(Error::from)?.iter().enumerate() {
        if only_row.is_some_and(|wanted| wanted != row_index) {
            continue;
        }
        for (column, cell) in row.cells().map_err(Error::from)?.iter().enumerate() {
            if only_column.is_some_and(|wanted| wanted != column) {
                continue;
            }
            let Some(body) = cell.text_body().map_err(Error::from)? else {
                continue;
            };
            scopes.try_reserve(1).map_err(allocation)?;
            scopes.push(Scope {
                locus: Locus {
                    place: Place::Slide,
                    slide,
                    shape: tree.handle.position().get(),
                    cell: Some((row_index, column)),
                },
                body: Arc::new(body),
                paragraph: None,
                run: None,
            });
        }
    }
    Ok(())
}
