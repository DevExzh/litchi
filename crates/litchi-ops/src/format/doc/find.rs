//! DOC find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! The document root and the body search every ordinary main-story
//! paragraph; a paragraph searches itself, and a run only its own text
//! (OfficeCLI's run scope). A paragraph's character runs are one group; a
//! match spanning runs collapses into the run where it starts. Every
//! changed run is one native `replace_text` of its UTF-16 span
//! (`TextTarget::BodySpan`), staged last to first so no change moves a
//! span staged after it, in one native edit and one commit: each run keeps
//! its character formatting.
//!
//! Runs that hold fields, drawings or other structural characters cannot
//! be rewritten as text: a find that touches one is refused, and merge
//! treats it as fixed content. Text in tables is searched too and refused
//! when it would change (the native table-cell edit replaces a cell
//! whole). Format properties are refused: the native editor formats whole
//! runs and cannot split one at a match.

use litchi_core::{FileFormat, Position};
use litchi_doc::body_text::TextTarget;

use super::spec::Kind;
use super::{Doc, Nav, Node, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// What a run is in the native editor.
#[derive(Clone, Copy)]
enum Key {
    /// A main-story run: its paragraph and UTF-16 span.
    Span {
        paragraph: Position,
        start: u32,
        end: u32,
    },
    /// Text the surface cannot rewrite (a table cell).
    Fixed,
}

/// The paragraphs a find searches: every one, or each named with an
/// optional run scope (its span start).
enum Scope {
    Every,
    Some(Vec<(Position, Option<u32>)>),
}

impl Doc {
    /// Find and replace text below `target`: the commit and the match
    /// count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, usize)> {
        self.check_source()?;
        let result = (|| {
            let replacement = surface::text_only(
                spec,
                FileFormat::Doc,
                "the native DOC editor formats whole runs and cannot split one at a match",
            )?;
            let mut every = false;
            let mut wanted = Vec::new();
            surface::each_target(Nav::new(self), target, |node: Node<'_>| {
                match node.element_kind() {
                    Kind::Document | Kind::Body => every = true,
                    Kind::Paragraph if node.paragraph_position().is_some() => {
                        wanted.extend(node.paragraph_position().map(|at| (at, None)));
                    },
                    Kind::Run if node.body_run().is_some() => {
                        wanted.extend(
                            node.body_run()
                                .map(|(paragraph, start, _)| (paragraph, Some(start))),
                        );
                    },
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedType,
                            "a DOC find searches the main story's paragraphs and runs",
                        )
                        .with_suggestion(
                            "target the document, the body, or a body paragraph or run",
                        ));
                    },
                }
                Ok(())
            })?;
            let scope = if every {
                Scope::Every
            } else {
                Scope::Some(wanted)
            };
            let matcher = spec.matcher();
            let (runs, keys) = self.surface(&scope, |text| matcher.may_match(text))?;
            let (changes, matches) =
                surface::find(&runs, spec, replacement, FileFormat::Doc, limits)?;
            Ok((self.commit_changes(&keys, changes, limits)?, matches))
        })()
        .map_err(|error: Error| error.with_format(FileFormat::Doc));
        self.check_source()?;
        result
    }

    /// Fill `record` into the main story's placeholders. `None` means
    /// nothing matched.
    pub(crate) fn merge(
        &self,
        record: &Record,
        options: &Options,
        limits: &SemanticLimits,
    ) -> Result<(Option<WriteCommit>, Summary)> {
        self.check_source()?;
        let result = (|| {
            let (runs, keys) = self.surface(&Scope::Every, |text| text.contains("{{"))?;
            let mut planner = Planner::new(record, options, limits);
            let changes = surface::merge(&runs, &mut planner, false, FileFormat::Doc, limits)?;
            let summary = planner.finish();
            if changes.is_empty() {
                return Ok((None, summary));
            }
            Ok((Some(self.commit_changes(&keys, changes, limits)?), summary))
        })()
        .map_err(|error: Error| error.with_format(FileFormat::Doc));
        self.check_source()?;
        result
    }

    /// The main story's runs in `scope`, and their native keys. A
    /// paragraph whose visible text `wanted` rejects contributes nothing:
    /// its runs are not read (every match or placeholder lies inside one
    /// paragraph, and the visible text holds every run's text).
    fn surface(
        &self,
        scope: &Scope,
        wanted: impl Fn(&str) -> bool,
    ) -> Result<(Vec<Run<'static>>, Vec<Key>)> {
        let nav = Nav::new(self);
        let root = nav.root()?;
        let mut runs = Vec::new();
        let mut keys = Vec::new();
        let mut group = 0usize;
        for body in nav.children(&root, KindSet::of(Kind::Body))? {
            let body = body?;
            for block in nav.children(&body, KindSet::of(Kind::Paragraph).with(Kind::Table))? {
                let block = block?;
                group += 1;
                if block.element_kind() == Kind::Table {
                    // Cells are searched as fixed text when the whole body is.
                    if matches!(scope, Scope::Every) {
                        for text in table_texts(&nav, &block)?
                            .into_iter()
                            .filter(|text| wanted(text))
                        {
                            runs.push(Run::new(text, group).editable(false));
                            keys.push(Key::Fixed);
                            group += 1;
                        }
                    }
                    continue;
                }
                let Some(position) = block.paragraph_position() else {
                    continue;
                };
                let scopes: Vec<Option<u32>> = match scope {
                    Scope::Every => vec![None],
                    Scope::Some(wanted) => wanted
                        .iter()
                        .filter(|(at, _)| *at == position)
                        .map(|(_, run)| *run)
                        .collect(),
                };
                if scopes.is_empty() || !block.paragraph_text()?.is_some_and(|text| wanted(&text)) {
                    continue;
                }
                for run in nav.children(&block, KindSet::of(Kind::Run))? {
                    let run = run?;
                    let (Some((paragraph, start, end)), Some(native)) = (run.body_run(), run.run())
                    else {
                        continue;
                    };
                    let raw = native.text().map_err(Error::from)?;
                    let editable = !raw
                        .chars()
                        .any(|character| character.is_control() && character != '\t');
                    let text = if editable {
                        raw.to_owned()
                    } else {
                        run.text()?.unwrap_or_default().into_owned()
                    };
                    let scoped = scopes
                        .iter()
                        .any(|wanted| wanted.is_none_or(|wanted| wanted == start));
                    runs.push(Run::new(text, group).scoped(scoped).editable(editable));
                    keys.push(Key::Span {
                        paragraph,
                        start,
                        end,
                    });
                }
            }
        }
        Ok((runs, keys))
    }

    /// Stage one native span replacement per change, last first, and
    /// commit once.
    fn commit_changes(
        &self,
        keys: &[Key],
        changes: Vec<Change>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let mut spans = Vec::with_capacity(changes.len());
        for change in changes {
            match keys[change.run] {
                Key::Span {
                    paragraph,
                    start,
                    end,
                } => spans.push((paragraph, start, end, change.text)),
                Key::Fixed => {
                    return Err(Error::new(
                        Code::InternalError,
                        "a fixed DOC run was planned for a change",
                    ));
                },
            }
        }
        spans.sort_by_key(|(paragraph, start, ..)| std::cmp::Reverse((paragraph.get(), *start)));
        let mut edit = self.native_edit(limits)?;
        let operations = spans.len();
        for (paragraph, start, end, text) in spans {
            edit.replace_text(
                TextTarget::BodySpan {
                    paragraph,
                    start,
                    end,
                },
                &text,
            )
            .map_err(Error::from)?;
        }
        self.publish(edit, operations, limits)
    }
}

/// The text of every cell of a table, in order.
fn table_texts<'f>(nav: &Nav<'f>, table: &Node<'f>) -> Result<Vec<String>> {
    let mut texts = Vec::new();
    for row in nav.children(table, KindSet::of(Kind::Row))? {
        for cell in nav.children(&row?, KindSet::of(Kind::Cell))? {
            texts.push(cell?.text()?.unwrap_or_default().into_owned());
        }
    }
    Ok(texts)
}
