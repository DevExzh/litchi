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

use litchi_core::{FileFormat, Position};
use litchi_doc::Edit;
use litchi_doc::body_text::{CharacterProperty, TextTarget};
use std::mem::size_of;

use super::{Doc, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
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
    edit: Edit<'static>,
    items: Vec<Item>,
    last_target: Option<(Box<str>, Position)>,
    signed: bool,
    max_snapshot_bytes: usize,
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
    pub(crate) fn new(
        base: &Doc,
        limits: &SemanticLimits,
        remaining_snapshot: usize,
    ) -> Result<Self> {
        if remaining_snapshot == 0 {
            return Err(Error::from(TransactionError::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: 1,
                maximum: 0,
            }));
        }
        base.check_source_for_write()?;
        // Check the live native edit handoff before the signed setter
        // materializes its unsigned private package; the final commit still
        // charges the exact retained before/after proof.
        base.preflight_snapshot_bytes(
            remaining_snapshot,
            size_of::<litchi_doc::body_text::ChangeRef<'static>>(),
        )?;
        let package = (if base.signature() == crate::format::signature::Signature::Signed {
            base.editable_package_for_write()?
        } else {
            base.editable_package()?
        })
        .ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the DOC document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Doc)
            .with_suggestion("reopen the document with source retention enabled")
        })?;
        let source_bytes = base
            .retained_source_len()
            .map_or(0, |bytes| usize::try_from(bytes).unwrap_or(usize::MAX));
        let edit = package.edit_with_options(super::edit_options(
            super::body_transaction_limits_with_source(*limits, source_bytes, remaining_snapshot),
        ))?;
        Ok(Self {
            base: base.clone(),
            edit,
            items: Vec::new(),
            last_target: None,
            signed: base.signature() == crate::format::signature::Signature::Signed,
            max_snapshot_bytes: remaining_snapshot,
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
        self.base.check_source_for_write()?;
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
    pub(crate) fn snapshot_bytes(&mut self) -> Result<usize> {
        let source = self.edit.source_bytes()?.len().saturating_mul(2);
        let text = self
            .items
            .iter()
            .fold(0usize, |total, item| match &item.change {
                Change::Text(text) => total.saturating_add(text.len().saturating_mul(2)),
                Change::Toggle { .. } => total,
            });
        Ok(source
            .saturating_add(text)
            .saturating_add(size_of::<litchi_doc::Patch>()))
    }

    pub(crate) fn stage(
        &mut self,
        plan: Plan,
        index: usize,
        spelling: Option<&str>,
    ) -> Result<bool> {
        self.base.check_source_for_write()?;
        let prospective = self.snapshot_bytes_for(&plan)?.ok_or_else(|| {
            Error::from(TransactionError::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: usize::MAX,
                maximum: self.max_snapshot_bytes,
            })
        })?;
        limit(
            LimitKind::SnapshotBytes,
            prospective,
            self.max_snapshot_bytes,
        )?;
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
                        .with_format(FileFormat::Doc)
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
        self.base.check_source_for_write()?;
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

    fn snapshot_bytes_for(&mut self, plan: &Plan) -> Result<Option<usize>> {
        let Some(source) = self.edit.source_bytes()?.len().checked_mul(2) else {
            return Ok(None);
        };
        let text = self
            .items
            .iter()
            .fold(0usize, |total, item| match &item.change {
                Change::Text(text) => total.saturating_add(text.len().saturating_mul(2)),
                Change::Toggle { .. } => total,
            });
        let incoming = match &plan.change {
            Change::Text(text) => text.len().saturating_mul(2),
            Change::Toggle { .. } => 0,
        };
        let Some(total) = source.checked_add(text.saturating_add(incoming)) else {
            return Ok(None);
        };
        let Some(total) = total.checked_add(size_of::<litchi_doc::Patch>()) else {
            return Ok(None);
        };
        let total = total.checked_add(
            self.items
                .len()
                .saturating_add(1)
                .saturating_mul(size_of::<litchi_doc::body_text::ChangeRef<'static>>()),
        );
        Ok(total)
    }

    pub(crate) fn commit(self) -> Result<WriteCommit> {
        let Self {
            base,
            edit,
            items,
            signed,
            max_snapshot_bytes,
            ..
        } = self;
        base.check_source_for_write()?;
        let operations = items.len();
        let before = base.clone();
        let commit = edit.commit()?;
        base.check_source_for_write()?;
        let changed = commit.diagnostics().changed();
        let staged_snapshot_bytes = super::patch_snapshot_bytes(commit.patch());
        // A native DOC edit may rewrite the compound stream even when a
        // sequence of paragraph assignments returns every touched value to
        // its original state.  Signed transactions must publish the exact
        // retained source in that case; compare the final authored values
        // against the immutable base rather than relying on stream bytes.
        let logical_noop = signed && logical_noop(&base, &items)?;
        let (package, patch) = commit.into_parts();
        let (after, patch, snapshot_bytes) = if logical_noop {
            // The staged edit was still committed above, so native refusal
            // and invalid intermediate operations retain their normal error
            // semantics.  Its changed patch is discarded after the authored
            // plan proves cancellation; the retained signed package gets a
            // fresh canonical identity patch of its own.
            let original = base.original_package_for_noop()?;
            let (_, patch) = empty_commit(&base, &original, max_snapshot_bytes)?;
            let snapshot_bytes = staged_snapshot_bytes.max(super::patch_snapshot_bytes(&patch));
            (base.clone(), patch, snapshot_bytes)
        } else {
            let snapshot_bytes = staged_snapshot_bytes.max(super::patch_snapshot_bytes(&patch));
            let after = if changed {
                base.with_package(package)
            } else {
                base.clone()
            };
            (after, patch, snapshot_bytes)
        };
        limit(LimitKind::SnapshotBytes, snapshot_bytes, max_snapshot_bytes)?;
        base.check_source_for_write()?;
        Ok(WriteCommit {
            before,
            after,
            patch: super::StepPatch::Primary(patch),
            operations,
            snapshot_bytes,
        })
    }
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TransactionError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

fn empty_commit(
    base: &Doc,
    source: &litchi_doc::Package,
    max_snapshot_bytes: usize,
) -> Result<(litchi_doc::Package, litchi_doc::Patch)> {
    base.check_source_for_write()?;
    base.preflight_snapshot_bytes(max_snapshot_bytes, 0)?;
    let source_bytes = base
        .retained_source_len()
        .map_or(0, |bytes| usize::try_from(bytes).unwrap_or(usize::MAX));
    let transaction_limits =
        litchi_doc::body_text::TransactionLimits::new(1, usize::MAX, usize::MAX)
            .with_max_retained_render_bytes(max_snapshot_bytes.saturating_sub(source_bytes));
    let commit = source
        .edit_with_options(super::edit_options(transaction_limits))?
        .commit()?;
    base.check_source_for_write()?;
    Ok(commit.into_parts())
}

/// Whether every coalesced paragraph assignment leaves the base's authored
/// value unchanged.  This proof is deliberately limited to the properties
/// admitted by [`Run::plan`]; unsupported or mixed families never enter this
/// run and therefore cannot be mistaken for a no-op.
fn logical_noop(base: &Doc, items: &[Item]) -> Result<bool> {
    use std::collections::HashMap;

    let mut final_text = HashMap::<Position, String>::new();
    let mut final_toggles = HashMap::<(Position, CharacterProperty), (bool, bool)>::new();
    for item in items {
        match &item.change {
            Change::Text(text) => {
                final_text.insert(item.position, text.clone());
            },
            Change::Toggle {
                property,
                enabled,
                holds,
            } => {
                final_toggles
                    .entry((item.position, *property))
                    .and_modify(|entry| entry.0 = *enabled)
                    .or_insert((*enabled, *holds));
            },
        }
    }
    for (position, text) in final_text {
        let Some(original) = base.body_paragraph_at(position)? else {
            return Ok(false);
        };
        if original != text {
            return Ok(false);
        }
    }
    Ok(final_toggles
        .into_values()
        .all(|(enabled, originally_holds)| enabled == originally_holds))
}
