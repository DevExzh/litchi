//! Coalesced ordinary body-paragraph writes for the DOC adapter.
//!
//! A run owns one native `Edit`.  Direct paragraph paths are stable
//! while their text is replaced or their characters toggled, so consecutive
//! `set` calls can resolve once, stage in order, and publish one commit:
//! `text` (variable width) and, since record 0359a, `bold` and `italic` on
//! the whole paragraph.  A paragraph takes one kind of change per run (a
//! toggle after a text replacement of the same paragraph, or the reverse,
//! closes the run first).  Queries and any other property remain barriers
//! and use the ordinary isolated setter.

use litchi_core::Position;
use litchi_doc::Edit;
use litchi_doc::body_text::{CharacterProperty, TextTarget};

use super::{Doc, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::transaction::SemanticLimits;
use crate::{Props, TargetArg};

/// A direct paragraph assignment admitted to a native run.
pub(crate) struct Plan {
    pub(crate) position: Position,
    pub(crate) change: Change,
}

/// The one property a run stages for a paragraph.
pub(crate) enum Change {
    /// Replace the paragraph's text.
    Text(String),
    /// Set a character toggle on the whole paragraph; `holds` is whether
    /// every run of the base paragraph already reads as `enabled`.
    Toggle {
        property: CharacterProperty,
        enabled: bool,
        holds: bool,
    },
}

impl Change {
    const fn is_text(&self) -> bool {
        matches!(self, Self::Text(_))
    }
}

struct Item {
    index: usize,
    position: Position,
    change: Change,
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
        let toggle = key.eq_ignore_ascii_case("bold") || key.eq_ignore_ascii_case("italic");
        if !key.eq_ignore_ascii_case("text") && !toggle {
            return Ok(None);
        }
        let text = value.as_text();
        if !toggle && text.is_none() {
            return Ok(None);
        }
        let path = target.path().expect("path checked above");
        let node = crate::path::resolve::get(&super::Nav::new(self), path)?;
        // Run text and other nodes take the ordinary setter path (0289).
        let Some(position) = node.paragraph_position() else {
            return Ok(None);
        };
        let change = if toggle {
            let Some((property, enabled, holds)) =
                super::edit::paragraph_toggle(&node, key.as_str(), value)?
            else {
                return Ok(None);
            };
            Change::Toggle {
                property,
                enabled,
                holds,
            }
        } else {
            Change::Text(text.expect("text checked above").to_owned())
        };
        Ok(Some(Plan { position, change }))
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
        let plan = if let (Some(spelling), Some((last, position))) = (spelling, &self.last_target)
            && last.as_ref() == spelling
        {
            let Some((key, value)) = props.iter().next() else {
                return Ok(None);
            };
            if props.len() == 1
                && key.eq_ignore_ascii_case("text")
                && let Some(text) = value.as_text()
            {
                Some(Plan {
                    position: *position,
                    change: Change::Text(text.to_owned()),
                })
            } else {
                self.base.paragraph_plan(target, props)?
            }
        } else {
            self.base.paragraph_plan(target, props)?
        };
        // One kind of change per paragraph per run: a different kind closes
        // the run so the isolated setter sees the published candidate.
        Ok(plan.filter(|plan| {
            !self.items.iter().any(|item| {
                item.position == plan.position && item.change.is_text() != plan.change.is_text()
            })
        }))
    }

    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }

    /// Conservative retained-snapshot charge used before commit. The final
    /// rendered snapshot is charged again by `flush_pending`.
    pub(crate) fn snapshot_bytes(&self) -> usize {
        let source = self.edit.source_bytes().len().saturating_mul(2);
        let text = self
            .items
            .iter()
            .fold(0usize, |total, item| match &item.change {
                Change::Text(text) => total.saturating_add(text.len().saturating_mul(2)),
                Change::Toggle { .. } => total,
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
        let changed = match &plan.change {
            Change::Text(text) => {
                let prior = self
                    .items
                    .iter()
                    .rev()
                    .find_map(|item| match &item.change {
                        Change::Text(text) if item.position == plan.position => Some(text.as_str()),
                        _ => None,
                    })
                    .map_or_else(
                        || self.base.body_paragraph_at(plan.position),
                        |text| Ok(Some(text)),
                    )?
                    .ok_or_else(|| {
                        Error::new(
                            Code::PathNotFound,
                            "the DOC paragraph is outside the main story",
                        )
                        .with_format(litchi_core::FileFormat::Doc)
                    })?;
                let changed = prior != text;
                self.edit.replace_paragraph(plan.position, text)?;
                changed
            },
            Change::Toggle {
                property,
                enabled,
                holds,
            } => {
                // The latest staged value of this toggle, else the base's.
                let holds = self
                    .items
                    .iter()
                    .rev()
                    .find_map(|item| match &item.change {
                        Change::Toggle {
                            property: staged,
                            enabled: value,
                            ..
                        } if item.position == plan.position && staged == property => {
                            Some(value == enabled)
                        },
                        _ => None,
                    })
                    .unwrap_or(*holds);
                if !holds {
                    self.edit.set_character_property(
                        TextTarget::body_paragraph(plan.position),
                        *property,
                        *enabled,
                    )?;
                }
                !holds
            },
        };
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
            change: plan.change,
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
