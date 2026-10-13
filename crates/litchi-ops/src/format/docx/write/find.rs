//! DOCX find/replace (record 0273): `set` with `find=`.
//!
//! The scope is the target's subtree: the document root and the body search
//! every paragraph of the main story, including table cells and block
//! content controls; a table, row, cell or content control searches its
//! paragraphs; a paragraph searches itself; a hyperlink or a run searches
//! only its own text, as OfficeCLI's run scope does. Each paragraph's text
//! nodes come from the leaf ([`litchi_docx::document::Edit::text_nodes`]),
//! the shared engine plans the rewrite, and one leaf call
//! ([`litchi_docx::document::Edit::rewrite_text_nodes`]) writes every
//! changed paragraph. Format properties then go through the ordinary run
//! setters on the runs the rewrite isolated.

use std::collections::HashMap;

use litchi_core::FileFormat;
use litchi_docx::document::{ElementPath, ElementStep, ParagraphRewrite, TextNodeRewrite};

use super::super::Docx;
use super::super::Nav;
use super::super::element;
use super::super::node::{Node, WriteSelection};
use super::super::spec::Kind;
use super::{Commit, Selection, Target, TargetKind, WorkEdit, admit_assignments};
use crate::error::{Code, Error, Result};
use crate::find::{Piece, Scratch, Spec, plan_with};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// One paragraph to search: everything, or only the text of some runs.
struct Scope {
    path: ElementPath,
    /// `None` searches the whole paragraph; otherwise the paths of the runs
    /// (or hyperlinks) whose text is in scope.
    within: Option<Vec<ElementPath>>,
}

impl Docx {
    /// Find (and replace or format) text in the subtree of `target`,
    /// returning the commit and the match count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        let result = self
            .find_inner(target, spec, limits)
            .map_err(|error| error.with_format(FileFormat::Docx));
        self.owner().finish(result)
    }

    fn find_inner(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        let editor = self.edit()?;
        if spec.format().get("text").is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "`text` replaces a whole run; with `find`, use `replace`",
            )
            .with_property("text"));
        }
        // Admit the format properties as run properties before reading.
        let assignments = if spec.formats() {
            admit_assignments(TargetKind::Run, spec.format())?
        } else {
            Vec::new()
        };
        let nodes = self.resolve_targets(target)?;
        let scopes = scopes(&nodes)?;
        let mut work = WorkEdit::from(editor)?;
        let matcher = spec.matcher();
        let mut matches = 0usize;
        let mut staged = 0usize;
        let mut rewrites = Vec::new();
        let mut scratch = Scratch::default();
        for scope in &scopes {
            let nodes = both!(&work, |edit| edit
                .text_nodes(&scope.path)
                .map_err(Error::from))?;
            if nodes.is_empty() {
                continue;
            }
            let pieces: Vec<Piece<'_>> = nodes
                .iter()
                .map(|node| Piece {
                    text: node.text.as_str(),
                    group: node.group,
                    scoped: scope.within.as_ref().is_none_or(|within| {
                        node.run
                            .as_ref()
                            .is_some_and(|run| within.iter().any(|path| starts_with(run, path)))
                    }),
                })
                .collect();
            let planned = plan_with(
                &mut scratch,
                &pieces,
                &matcher,
                spec.replacement(),
                spec.formats(),
                limits.max_staged_bytes().saturating_sub(staged),
            )?;
            if planned.matches == 0 {
                continue;
            }
            matches = matches.saturating_add(planned.matches);
            let mut slots = Vec::with_capacity(planned.rewrites.len());
            for rewrite in planned.rewrites {
                let text = rewrite
                    .text
                    .unwrap_or_else(|| pieces[rewrite.piece].text.to_owned());
                staged = staged.saturating_add(text.len());
                slots.push(TextNodeRewrite::new(rewrite.piece, text).isolating(rewrite.marks));
            }
            if !slots.is_empty() {
                rewrites.push(ParagraphRewrite::new(scope.path.clone(), slots));
            }
        }
        if rewrites.is_empty() {
            return Ok((work.commit(self)?, matches));
        }
        let isolated = both!(&mut work, |edit| edit
            .rewrite_text_nodes(&rewrites)
            .map_err(Error::from))?;
        for path in isolated {
            let target = Target {
                selection: Selection::Element(path),
                kind: TargetKind::Run,
            };
            for assignment in &assignments {
                work.apply(&target, assignment)?;
            }
        }
        Ok((work.commit(self)?, matches))
    }
}

/// Whether `path` is `prefix` or below it.
fn starts_with(path: &ElementPath, prefix: &ElementPath) -> bool {
    path.steps().starts_with(prefix.steps())
}

/// The paragraphs the target nodes search, each once.
fn scopes(nodes: &[Node<'_>]) -> Result<Vec<Scope>> {
    let mut scopes: Vec<Scope> = Vec::new();
    let mut index: HashMap<ElementPath, usize> = HashMap::new();
    let mut body_index = None;
    let mut add =
        |path: ElementPath, within: Option<ElementPath>, scopes: &mut Vec<Scope>| match index
            .get(&path)
        {
            Some(&at) => match (&mut scopes[at].within, within) {
                (Some(runs), Some(run)) => runs.push(run),
                (slot, None) => *slot = None,
                (None, Some(_)) => {},
            },
            None => {
                index.insert(path.clone(), scopes.len());
                scopes.push(Scope {
                    path,
                    within: within.map(|run| vec![run]),
                });
            },
        };
    for node in nodes {
        match node.element_kind() {
            Kind::Document | Kind::Body => {
                // Direct body paragraphs come from the body's block index;
                // only tables and content controls are walked as nodes.
                let blocks = super::super::nav::load_body(node.docx)?;
                let mut nested = false;
                for block in blocks.iter() {
                    match block.kind() {
                        litchi_docx::BlockKind::Paragraph => add(
                            ElementPath::new([ElementStep::Paragraph(
                                block.transaction_position().map_err(Error::from)?,
                            )]),
                            None,
                            &mut scopes,
                        ),
                        litchi_docx::BlockKind::Table | litchi_docx::BlockKind::ContentControl => {
                            nested = true;
                        },
                        _ => {},
                    }
                }
                if nested {
                    let nav = Nav::new(node.docx);
                    let body = if node.element_kind() == Kind::Body {
                        node.clone()
                    } else {
                        nav.children(node, KindSet::of(Kind::Body))?
                            .next()
                            .transpose()?
                            .ok_or_else(|| {
                                Error::new(Code::InternalError, "the DOCX document has no body")
                            })?
                    };
                    let mut paragraphs = Vec::new();
                    for child in nav.children(&body, KindSet::of(Kind::Table).with(Kind::Sdt))? {
                        descend(&child?, &mut paragraphs)?;
                    }
                    for paragraph in &paragraphs {
                        add(
                            paragraph_path(paragraph, &mut body_index)?,
                            None,
                            &mut scopes,
                        );
                    }
                }
            },
            Kind::Table | Kind::Row | Kind::Cell | Kind::Sdt => {
                let mut paragraphs = Vec::new();
                descend(node, &mut paragraphs)?;
                for paragraph in &paragraphs {
                    add(
                        paragraph_path(paragraph, &mut body_index)?,
                        None,
                        &mut scopes,
                    );
                }
            },
            Kind::Paragraph => add(paragraph_path(node, &mut body_index)?, None, &mut scopes),
            Kind::Run | Kind::Hyperlink => {
                let path = element::element_path(node)?;
                let paragraph = path
                    .steps()
                    .iter()
                    .rposition(|step| matches!(step, ElementStep::Paragraph(_)))
                    .ok_or_else(|| {
                        Error::new(Code::InternalError, "a DOCX run path names no paragraph")
                    })?;
                let paragraph = ElementPath::new(&path.steps()[..=paragraph]);
                add(paragraph, Some(path), &mut scopes);
            },
            _ => {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "find searches the document, the body, tables, rows, cells, content controls, paragraphs, hyperlinks and runs",
                )
                .with_property(crate::find::FIND)
                .with_suggestion("headers, footers, notes and comments are not searched yet (record 0273)"));
            },
        }
    }
    Ok(scopes)
}

/// The paragraphs below `node`, in document order.
fn descend<'f>(node: &Node<'f>, out: &mut Vec<Node<'f>>) -> Result<()> {
    let nav = Nav::new(node.docx);
    let kinds = match node.element_kind() {
        Kind::Document => KindSet::of(Kind::Body),
        Kind::Table => KindSet::of(Kind::Row),
        Kind::Row => KindSet::of(Kind::Cell),
        _ => KindSet::of(Kind::Paragraph)
            .with(Kind::Table)
            .with(Kind::Sdt),
    };
    for child in nav.children(node, kinds)? {
        let child = child?;
        if child.element_kind() == Kind::Paragraph {
            out.push(child);
        } else {
            descend(&child, out)?;
        }
    }
    Ok(())
}

/// The leaf path of a paragraph node: its body position when it is a direct
/// body paragraph, else its element path.
fn paragraph_path(node: &Node<'_>, body: &mut Option<litchi_docx::Blocks>) -> Result<ElementPath> {
    match node.write_selection_cached(body) {
        Ok(WriteSelection::Paragraph(position)) => {
            Ok(ElementPath::new([ElementStep::Paragraph(position)]))
        },
        Ok(WriteSelection::Run { .. }) => Err(Error::new(
            Code::InternalError,
            "a DOCX paragraph resolved to a run",
        )),
        Err(error) if error.code() == Code::UnsupportedCapability => element::element_path(node),
        Err(error) => Err(error),
    }
}
