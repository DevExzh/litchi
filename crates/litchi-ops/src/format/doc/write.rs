//! Coalesced ordinary body-paragraph text writes for the DOC adapter.
//!
//! A run owns one native `Edit`.  Direct paragraph paths are stable
//! while their text is replaced, so consecutive `set` calls can resolve once,
//! stage in order, and publish one variable-width commit.  Queries and any
//! non-text property remain barriers and use the ordinary isolated setter.

use litchi_core::Position;
use litchi_doc::Edit;

use super::{Doc, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::transaction::SemanticLimits;
use crate::{Props, TargetArg};

/// A direct paragraph text assignment admitted to a native run.
pub(crate) struct Plan {
    pub(crate) position: Position,
    pub(crate) text: String,
}

struct Item {
    index: usize,
    position: Position,
    text: String,
}

/// One coalesced body-text transaction.
pub(crate) struct Run {
    base: Doc,
    edit: Edit,
    items: Vec<Item>,
    last_target: Option<(Box<str>, Position)>,
}

impl Doc {
    /// Admit only the direct `/body/p[N]` text shape. Returning `None` keeps
    /// query targets and other property families on the normal setter path,
    /// where their typed capability error can be reported precisely.
    pub(crate) fn paragraph_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        if target.path().is_none() || props.len() != 1 {
            return Ok(None);
        }
        let Some((key, value)) = props.iter().next() else {
            return Ok(None);
        };
        if !key.eq_ignore_ascii_case("text") {
            return Ok(None);
        }
        let Some(text) = value.as_text() else {
            return Ok(None);
        };
        let path = target.path().expect("path checked above");
        let node = crate::path::resolve::get(&super::Nav::new(self), path)?;
        let Some(position) = node.paragraph_position() else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "DOC text writes target paragraph nodes",
            )
            .with_format(litchi_core::FileFormat::Doc));
        };
        Ok(Some(Plan {
            position,
            text: text.to_owned(),
        }))
    }
}

impl Run {
    pub(crate) fn new(base: &Doc, limits: &SemanticLimits) -> Result<Self> {
        base.check_source()?;
        let package = base.editable_package()?.ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the DOC document was opened without retaining its source for editing",
            )
            .with_format(litchi_core::FileFormat::Doc)
            .with_suggestion("reopen the document with source retention enabled")
        })?;
        let edit = package
            .edit_with_options(super::edit_options(super::body_transaction_limits(*limits)))?;
        Ok(Self {
            base: base.clone(),
            edit,
            items: Vec::new(),
            last_target: None,
        })
    }

    /// Resolve another direct path against the immutable base. Paragraph
    /// positions do not move when text width changes, while a query does not
    /// meet the run's direct-path proof and is handled as a barrier.
    pub(crate) fn plan(
        &self,
        spelling: Option<&str>,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        self.base.check_source()?;
        if let (Some(spelling), Some((last, position))) = (spelling, &self.last_target)
            && last.as_ref() == spelling
        {
            let Some((key, value)) = props.iter().next() else {
                return Ok(None);
            };
            if props.len() == 1
                && key.eq_ignore_ascii_case("text")
                && let Some(text) = value.as_text()
            {
                return Ok(Some(Plan {
                    position: *position,
                    text: text.to_owned(),
                }));
            }
            return Ok(None);
        }
        self.base.paragraph_plan(target, props)
    }

    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }

    /// Conservative retained-snapshot charge used before commit. The final
    /// rendered snapshot is charged again by `flush_pending`.
    pub(crate) fn snapshot_bytes(&self) -> usize {
        let source = self.edit.source_bytes().len().saturating_mul(2);
        let text = self.items.iter().fold(0usize, |total, item| {
            total.saturating_add(item.text.len().saturating_mul(2))
        });
        source
            .saturating_add(text)
            .saturating_add(size_of::<litchi_doc::Patch>())
    }

    pub(crate) fn stage(
        &mut self,
        plan: Plan,
        index: usize,
        spelling: Option<&str>,
    ) -> Result<bool> {
        self.base.check_source()?;
        let prior = self
            .items
            .iter()
            .rev()
            .find(|item| item.position == plan.position)
            .map_or_else(
                || self.base.body_paragraph_at(plan.position),
                |item| Ok(Some(item.text.as_str())),
            )?
            .ok_or_else(|| {
                Error::new(
                    Code::PathNotFound,
                    "the DOC paragraph is outside the main story",
                )
                .with_format(litchi_core::FileFormat::Doc)
            })?;
        let changed = prior != plan.text;
        self.edit.replace_paragraph(plan.position, &plan.text)?;
        self.base.check_source()?;
        self.items
            .try_reserve(1)
            .map_err(|source| crate::transaction::Error::Allocation { source })?;
        if let Some(spelling) = spelling {
            self.last_target = Some((spelling.into(), plan.position));
        } else {
            self.last_target = None;
        }
        self.items.push(Item {
            index,
            position: plan.position,
            text: plan.text,
        });
        Ok(changed)
    }

    pub(crate) fn commit(self) -> Result<WriteCommit> {
        let Self {
            base, edit, items, ..
        } = self;
        base.check_source()?;
        let operations = items.len();
        let before = base.clone();
        let commit = edit.commit()?;
        base.check_source()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        let snapshot_bytes = super::patch_snapshot_bytes(&patch);
        let after = if changed {
            base.with_package(package)
        } else {
            base.clone()
        };
        base.check_source()?;
        Ok(WriteCommit {
            before,
            after,
            patch: super::StepPatch::Primary(patch),
            operations,
            snapshot_bytes,
        })
    }
}
