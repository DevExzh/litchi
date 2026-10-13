//! PPTX find/replace (record 0273): `set` with `find=`.
//!
//! The scope is the target's subtree: the presentation searches every
//! slide, a slide its shapes (inside groups too), a shape or placeholder its
//! text body, a paragraph itself, and a run or line break only its own text.
//! A paragraph is one search group; line breaks and fields separate text.
//! Matches become run text edits (`TextTarget::Run`), and format properties
//! become run-span edits (`TextTarget::RunSpan`) on exactly the replaced or
//! matched text, all through one native edit.
//!
//! Table cells, notes, masters and layouts are not searched yet: a table
//! cell in scope whose text matches is refused rather than skipped.

use litchi_core::{FileFormat, Position};
use litchi_drawingml::text::body::{Body, RunKind};
use litchi_pptx::opened::{TextEdit, TextTarget};

use super::text::{TextValues, admit};
use super::{Commit, Pptx, allocation, error};
use crate::find::{Piece, Spec, plan};
use crate::format::pptx::Nav;
use crate::format::pptx::node::{Item, TreeItem};
use crate::format::pptx::spec::{Kind, RUN_WRITE_TABLE, TextSetter};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::{Code, Error, Result};

/// One shape text body to search, optionally narrowed to a paragraph or
/// one run.
struct Scope {
    slide: usize,
    shape: Position,
    body: std::sync::Arc<Body>,
    paragraph: Option<usize>,
    run: Option<usize>,
}

/// A planned change to one run: its new text and the spans to format.
struct RunChange {
    slide: usize,
    shape: Position,
    paragraph: usize,
    run: usize,
    text: Option<String>,
    marks: Vec<std::ops::Range<usize>>,
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
        let scopes = self.find_scopes(target, &matcher)?;
        let mut matches = 0usize;
        let mut staged = 0usize;
        let mut changes: Vec<RunChange> = Vec::new();
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
            let planned = plan(
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
                    slide: scope.slide,
                    shape: scope.shape,
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
            (left.slide, left.shape, left.paragraph, left.run)
                .cmp(&(right.slide, right.shape, right.paragraph, right.run))
                .reverse()
        });
        let properties = format.as_ref().map(TextValues::properties);
        let mut groups: Vec<(usize, Vec<TextEdit<'_>>)> = Vec::new();
        for change in &changes {
            if groups
                .last()
                .is_none_or(|(slide, _)| *slide != change.slide)
            {
                groups.try_reserve(1).map_err(allocation)?;
                groups.push((change.slide, Vec::new()));
            }
            let Some((_, edits)) = groups.last_mut() else {
                unreachable!("a group was just pushed")
            };
            if let Some(text) = &change.text {
                edits.try_reserve(1).map_err(allocation)?;
                edits.push(TextEdit::new(
                    TextTarget::Run {
                        shape: change.shape,
                        paragraph: Position::new(change.paragraph),
                        run: Position::new(change.run),
                    },
                    litchi_pptx::opened::TextProperties::default().with_text(text),
                ));
            }
            if let Some(properties) = properties {
                for mark in change.marks.iter().rev() {
                    edits.try_reserve(1).map_err(allocation)?;
                    edits.push(TextEdit::new(
                        TextTarget::RunSpan {
                            shape: change.shape,
                            paragraph: Position::new(change.paragraph),
                            run: Position::new(change.run),
                            start: mark.start,
                            end: mark.end,
                        },
                        properties,
                    ));
                }
            }
        }
        groups.retain(|(_, edits)| !edits.is_empty());
        Ok((self.commit_text_edits(&groups)?, matches))
    }

    /// The text bodies the target nodes search, each once.
    fn find_scopes(
        &self,
        target: &TargetArg<'_>,
        matcher: &crate::find::Matcher<'_>,
    ) -> Result<Vec<Scope>> {
        let nav = Nav::new(self);
        let mut scopes = Vec::new();
        let visit = |node: crate::format::pptx::Node<'_>, scopes: &mut Vec<Scope>| {
            collect(&nav, node, matcher, scopes)
        };
        if let Some(path) = target.path() {
            visit(crate::path::resolve::get(&nav, path)?, &mut scopes)?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(&nav, query.clone())?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(error(
                    Code::InvalidSelector,
                    "mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                visit(node?, &mut scopes)?;
            }
        } else {
            return Err(error(Code::InvalidArgument, "mutation target is empty"));
        }
        // A shape reached twice (a query naming it and its paragraph)
        // searches once, with the widest scope.
        scopes.sort_by_key(|scope| (scope.slide, scope.shape.get()));
        let mut merged: Vec<Scope> = Vec::with_capacity(scopes.len());
        for scope in scopes {
            match merged.last_mut() {
                Some(last) if last.slide == scope.slide && last.shape == scope.shape => {
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
}

fn collect(
    nav: &Nav<'_>,
    node: crate::format::pptx::Node<'_>,
    matcher: &crate::find::Matcher<'_>,
    scopes: &mut Vec<Scope>,
) -> Result<()> {
    let mut push = |tree: &TreeItem, paragraph: Option<usize>, run: Option<usize>| -> Result<()> {
        let Some(slide) = tree.container.slide() else {
            return Err(error(
                Code::UnsupportedCapability,
                "find does not search slide masters and layouts yet",
            )
            .with_property(crate::find::FIND));
        };
        let Some(body) = tree.body()? else {
            return Ok(());
        };
        scopes.try_reserve(1).map_err(allocation)?;
        scopes.push(Scope {
            slide: slide as usize,
            shape: tree.handle.position(),
            body: std::sync::Arc::clone(body),
            paragraph,
            run,
        });
        Ok(())
    };
    match node.item() {
        Item::Presentation | Item::Slide(_) => {
            let kinds = KindSet::of(Kind::Slide)
                .with(Kind::Shape)
                .with(Kind::Placeholder)
                .with(Kind::Group)
                .with(Kind::Table);
            for child in nav.children(&node, kinds)? {
                collect(nav, child?, matcher, scopes)?;
            }
            Ok(())
        },
        Item::Tree(tree) => match tree.kind {
            Kind::Shape | Kind::Placeholder => push(tree, None, None),
            Kind::Group => {
                let kinds = KindSet::of(Kind::Shape)
                    .with(Kind::Placeholder)
                    .with(Kind::Group)
                    .with(Kind::Table);
                for child in nav.children(&node, kinds)? {
                    collect(nav, child?, matcher, scopes)?;
                }
                Ok(())
            },
            Kind::Table => refuse_matching_table(nav, &node, matcher),
            // Pictures, charts, media and the like carry no searchable text.
            _ => Ok(()),
        },
        Item::Paragraph(text) => push(&text.shape, Some(text.paragraph as usize), None),
        Item::Run(text) | Item::LineBreak(text) => push(
            &text.shape,
            Some(text.paragraph as usize),
            Some(text.run as usize),
        ),
        Item::Row(_) | Item::Cell(_) => refuse_matching_table(nav, &node, matcher),
        _ => Err(error(
            Code::UnsupportedCapability,
            "find searches the presentation, slides, shapes, groups, paragraphs and runs",
        )
        .with_property(crate::find::FIND)
        .with_suggestion(
            "notes, comments, masters and layouts are not searched yet (record 0273)",
        )),
    }
}

/// Table cells cannot be rewritten run by run yet: a match inside one is
/// refused, so a broad find never skips text it would have changed.
fn refuse_matching_table(
    nav: &Nav<'_>,
    node: &crate::format::pptx::Node<'_>,
    matcher: &crate::find::Matcher<'_>,
) -> Result<()> {
    let cells: Vec<crate::format::pptx::Node<'_>> = match node.item() {
        Item::Cell(_) => vec![node.clone()],
        _ => {
            let mut cells = Vec::new();
            let kinds = KindSet::of(Kind::Row).with(Kind::Cell);
            let mut stack = vec![node.clone()];
            while let Some(parent) = stack.pop() {
                for child in nav.children(&parent, kinds)? {
                    let child = child?;
                    if matches!(child.item(), Item::Cell(_)) {
                        cells.push(child);
                    } else {
                        stack.push(child);
                    }
                }
            }
            cells
        },
    };
    for cell in &cells {
        let Item::Cell(item) = cell.item() else {
            continue;
        };
        let text = crate::format::pptx::text::cell_text(item)?;
        if text
            .split('\n')
            .any(|line| matcher.find_iter(line).next().is_some())
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "find/replace inside PPTX table cells is not integrated",
            )
            .with_format(FileFormat::Pptx)
            .with_property(crate::find::FIND)
            .with_suggestion("narrow the find to the slide's shapes (record 0273 follow-up)"));
        }
    }
    Ok(())
}
