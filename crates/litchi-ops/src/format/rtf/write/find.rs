//! RTF find/replace (record 0273): `set` with `find=` and `replace=`.
//!
//! The document root and the body search every main-story paragraph, a
//! paragraph itself, and a run only its own text (OfficeCLI's run scope).
//! A paragraph's runs are one search group; a match spanning runs
//! collapses into the run where it starts. Every changed run is one native
//! run-text replacement (`Edit::replace_body_run_texts`) in one native
//! commit, so each run keeps its character formatting. Format properties
//! would need a run split the native editor does not offer yet: they are
//! refused.

use litchi_core::{FileFormat, Position};
use litchi_rtf::edit::{EditOptions, RunTextReplacement, SemanticLimits as NativeLimits};

use super::super::{Kind, Nav, Node, Rtf};
use super::{Commit, patch_snapshot_bytes, snapshot_limit};
use crate::error::{Code, Error, Result};
use crate::find::{Piece, Scratch, Spec, plan_with};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, SemanticLimits};

impl Rtf {
    /// Find and replace text below `target`, returning the commit and the
    /// match count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        self.check_source()?;
        let result = self
            .find_inner(target, spec, limits)
            .map_err(|error| error.with_format(FileFormat::Rtf));
        self.check_source()?;
        result
    }

    fn find_inner(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        self.admit_edit()?;
        if spec.formats() {
            let key = spec
                .format()
                .iter()
                .next()
                .map_or(crate::find::FIND, |(key, _)| key.as_str());
            return Err(Error::new(
                Code::UnsupportedProperty,
                "an RTF find replaces text; formatting the replacement needs a native run split (record 0273 follow-up)",
            )
            .with_property(key.to_owned()));
        }
        let Some(replacement) = spec.replacement() else {
            return Err(
                Error::new(Code::MissingProperty, "an RTF find needs `replace`")
                    .with_property(crate::find::REPLACE),
            );
        };
        let nav = Nav::new(self);
        // `None`: every paragraph; else the paragraphs to search, each with
        // an optional run scope (`None` searches the whole paragraph).
        let mut every = false;
        let mut wanted: Vec<(Position, Option<Position>)> = Vec::new();
        let mut collect = |node: Node<'_>| -> Result<()> {
            match node.element_kind() {
                Kind::Document | Kind::Body => every = true,
                Kind::Paragraph | Kind::Run => {
                    let paragraph = node.paragraph_position().ok_or_else(|| {
                        Error::new(Code::UnsupportedType, "an RTF find needs a body paragraph")
                    })?;
                    wanted.push((paragraph, node.run_position()));
                },
                // Record 0275 added tables, stories, links, pictures, styles
                // and lists; RTF find still searches body paragraphs only.
                _ => {
                    return Err(Error::new(
                        Code::UnsupportedType,
                        "an RTF find searches the body's paragraphs and runs",
                    )
                    .with_suggestion("target the document, the body, or a body paragraph or run"));
                },
            }
            Ok(())
        };
        if let Some(path) = target.path() {
            collect(crate::path::resolve::get(&nav, path)?)?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(&nav, query.clone())?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(Error::new(
                    Code::InvalidSelector,
                    "the mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                collect(node?)?;
            }
        } else {
            return Err(Error::new(
                Code::InvalidArgument,
                "the mutation target is empty",
            ));
        }
        let mut paragraphs: Vec<(Node<'_>, Vec<Option<Position>>)> = Vec::new();
        let root = nav.root()?;
        for body in nav.children(&root, KindSet::of(Kind::Body))? {
            for paragraph in nav.children(&body?, KindSet::of(Kind::Paragraph))? {
                let paragraph = paragraph?;
                let Some(position) = paragraph.paragraph_position() else {
                    continue;
                };
                let scopes: Vec<Option<Position>> = if every {
                    vec![None]
                } else {
                    wanted
                        .iter()
                        .filter(|(at, _)| *at == position)
                        .map(|(_, run)| *run)
                        .collect()
                };
                if !scopes.is_empty() {
                    paragraphs.push((paragraph, scopes));
                }
            }
        }

        let matcher = spec.matcher();
        let mut matches = 0usize;
        let mut scratch = Scratch::default();
        let mut staged = 0usize;
        let mut updates: Vec<RunTextReplacement> = Vec::new();
        for (paragraph, scopes) in &paragraphs {
            let Some(position) = paragraph.paragraph_position() else {
                continue;
            };
            let runs: Vec<Node<'_>> = nav
                .children(paragraph, KindSet::of(Kind::Run))?
                .collect::<Result<_>>()?;
            let texts: Vec<std::borrow::Cow<'_, str>> = runs
                .iter()
                .map(|run| Ok(run.text()?.unwrap_or_default()))
                .collect::<Result<_>>()?;
            let pieces: Vec<Piece<'_>> = runs
                .iter()
                .zip(&texts)
                .map(|(run, text)| Piece {
                    text,
                    group: 0,
                    scoped: scopes
                        .iter()
                        .any(|scope| scope.is_none_or(|wanted| run.run_position() == Some(wanted))),
                })
                .collect();
            let planned = plan_with(
                &mut scratch,
                &pieces,
                &matcher,
                Some(replacement),
                false,
                limits.max_staged_bytes().saturating_sub(staged),
            )?;
            matches = matches.saturating_add(planned.matches);
            for rewrite in planned.rewrites {
                let (Some(text), Some(run)) = (rewrite.text, runs[rewrite.piece].run_position())
                else {
                    continue;
                };
                staged = staged.saturating_add(text.len());
                updates
                    .try_reserve(1)
                    .map_err(|source| TransactionError::Allocation { source })?;
                updates.push(RunTextReplacement::new(position, run, text));
            }
        }
        let native_limits = NativeLimits::builder()
            .max_operations(limits.max_operations())
            .map_err(Error::from)?
            .max_replacement_bytes(limits.max_staged_bytes())
            .map_err(Error::from)?
            .build()
            .map_err(Error::from)?;
        let mut edit = self
            .0
            .package
            .edit_with_options(EditOptions::new().semantic_limits(native_limits));
        if !updates.is_empty() {
            edit.replace_body_run_texts(&updates).map_err(Error::from)?;
        }
        let native = edit.commit().map_err(Error::from)?;
        let changed = native.diagnostics().changed();
        let operations = native.diagnostics().operation_count();
        let (package, patch) = native.into_parts();
        let snapshot_bytes = patch_snapshot_bytes(&patch);
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
        }
        let after = if changed {
            self.with_package(package)
        } else {
            self.clone()
        };
        Ok((
            Commit {
                before: self.clone(),
                after,
                patch,
                operations,
                snapshot_bytes,
            },
            matches,
        ))
    }
}
