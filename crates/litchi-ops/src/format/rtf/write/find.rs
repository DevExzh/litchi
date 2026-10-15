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
use crate::find::Spec;
use crate::find::surface::{self, Run};
use crate::merge::{Options, Planner, Record, Summary};
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
        let replacement = surface::text_only(
            spec,
            FileFormat::Rtf,
            "formatting the replacement needs a native RTF run split",
        )?;
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
        surface::each_target(&nav, target, &mut collect)?;
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

        let mut runs: Vec<Run<'_>> = Vec::new();
        let mut keys: Vec<(Position, Position)> = Vec::new();
        for (group, (paragraph, scopes)) in paragraphs.iter().enumerate() {
            let Some(position) = paragraph.paragraph_position() else {
                continue;
            };
            for run in nav.children(paragraph, KindSet::of(Kind::Run))? {
                let run = run?;
                let Some(at) = run.run_position() else {
                    continue;
                };
                let scoped = scopes
                    .iter()
                    .any(|scope| scope.is_none_or(|wanted| at == wanted));
                runs.try_reserve(1)
                    .map_err(|source| TransactionError::Allocation { source })?;
                runs.push(Run::new(run.text()?.unwrap_or_default(), group).scoped(scoped));
                keys.push((position, at));
            }
        }
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Rtf, limits)?;
        let updates: Vec<RunTextReplacement> = changes
            .into_iter()
            .map(|change| {
                let (paragraph, run) = keys[change.run];
                RunTextReplacement::new(paragraph, run, change.text)
            })
            .collect();
        Ok((self.commit_run_texts(&updates, limits)?, matches))
    }

    /// Commit run-text replacements in one native edit (none: an exact
    /// no-op commit).
    fn commit_run_texts(
        &self,
        updates: &[RunTextReplacement],
        limits: &SemanticLimits,
    ) -> Result<Commit> {
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
            edit.replace_body_run_texts(updates).map_err(Error::from)?;
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
        Ok(Commit {
            before: self.clone(),
            after,
            patch,
            operations,
            snapshot_bytes,
        })
    }

    /// Template merge (record 0295): placeholders in every main-story
    /// paragraph, rewritten run by run so each run keeps its formatting;
    /// one native commit. `None` means nothing matched.
    pub(crate) fn merge(
        &self,
        record: &Record,
        options: &Options,
        limits: &SemanticLimits,
    ) -> Result<(Option<Commit>, Summary)> {
        self.check_source()?;
        let result = (|| {
            self.admit_edit()?;
            let nav = Nav::new(self);
            let root = nav.root()?;
            let mut runs: Vec<Run<'_>> = Vec::new();
            let mut keys: Vec<(Position, Position)> = Vec::new();
            let mut group = 0usize;
            for body in nav.children(&root, KindSet::of(Kind::Body))? {
                for paragraph in nav.children(&body?, KindSet::of(Kind::Paragraph))? {
                    let paragraph = paragraph?;
                    group += 1;
                    let Some(position) = paragraph.paragraph_position() else {
                        continue;
                    };
                    for run in nav.children(&paragraph, KindSet::of(Kind::Run))? {
                        let run = run?;
                        let Some(at) = run.run_position() else {
                            continue;
                        };
                        runs.push(Run::new(
                            run.text()?.unwrap_or_default().into_owned(),
                            group,
                        ));
                        keys.push((position, at));
                    }
                }
            }
            let mut planner = Planner::new(record, options, limits);
            let changes = surface::merge(&runs, &mut planner, false, FileFormat::Rtf, limits)?;
            let summary = planner.finish();
            if changes.is_empty() {
                return Ok((None, summary));
            }
            let updates: Vec<RunTextReplacement> = changes
                .into_iter()
                .map(|change| {
                    let (paragraph, run) = keys[change.run];
                    RunTextReplacement::new(paragraph, run, change.text)
                })
                .collect();
            Ok((Some(self.commit_run_texts(&updates, limits)?), summary))
        })()
        .map_err(|error: Error| error.with_format(FileFormat::Rtf));
        self.check_source()?;
        result
    }
}
