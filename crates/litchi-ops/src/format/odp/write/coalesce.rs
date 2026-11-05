//! Coalesced ODP property sets for explicit transactions and batches.
//!
//! Consecutive direct-path sets on one snapshot become one list of scene
//! changes, staged by one leaf `splice_scene` call on one `Package::edit()`
//! and published by one commit. A set joins the pending run only when its
//! targets cannot observe an earlier pending set: a later set of the same
//! text target replaces the earlier one (the earlier value is never
//! visible), attribute sets merge into their start tags, and a set whose
//! text target overlaps a pending one at another granularity (a story and
//! one of its paragraphs, a paragraph and one of its runs) closes the run
//! first, so it resolves against the published result.

use std::collections::{HashMap, HashSet};

use litchi_odp::scene::{Attribute, Change, ObjectRef, StoryRef};

use super::{Plan, WriteCommit, commit_changes, plan_set, resolve_nodes};
use crate::error::Result;
use crate::format::odp::Odp;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// What one pending change writes, for overlap checks.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum Footprint {
    Story(StoryRef),
    Paragraph(StoryRef, usize),
    Run(StoryRef, usize, usize),
    Attribute(ObjectRef, Attribute),
    PageName(usize),
    /// Character formatting of a paragraph or run of the story (record
    /// 0308): it reads the story's runs, so it never shares a run with a
    /// text write of the same story.
    Format(StoryRef),
}

impl Footprint {
    fn of(change: &Change) -> Option<Self> {
        Some(match change {
            Change::SetStoryText { story, .. } => Self::Story(story.clone()),
            Change::SetParagraphText {
                story, paragraph, ..
            } => Self::Paragraph(story.clone(), *paragraph),
            Change::SetRunText {
                story,
                paragraph,
                run,
                ..
            } => Self::Run(story.clone(), *paragraph, *run),
            Change::SetAttribute {
                object, attribute, ..
            } => Self::Attribute(object.clone(), *attribute),
            Change::SetPageName { page, .. } => Self::PageName(*page),
            Change::SetTextProperties { story, .. } => Self::Format(story.clone()),
            _ => return None,
        })
    }
}

/// What the pending run writes inside one story.
#[derive(Default)]
struct StoryWrites {
    /// The whole story is replaced.
    whole: bool,
    /// Paragraph sets, and the run sets per paragraph.
    paragraphs: HashSet<usize>,
    runs: HashMap<usize, usize>,
}

/// The pending writes, indexed so a new set's overlap check and same-target
/// replacement cost O(1) in the run's length.
#[derive(Default)]
struct Index {
    stories: HashMap<StoryRef, StoryWrites>,
    pages: HashSet<usize>,
    /// Stories with pending character formatting.
    formats: HashSet<StoryRef>,
    /// The position in `Run::changes` of each text footprint.
    text: HashMap<Footprint, usize>,
}

impl Index {
    /// Whether `footprint` must see the result of a pending write: it
    /// overlaps one at another granularity, or renames a page twice.
    fn overlaps(&self, footprint: &Footprint) -> bool {
        if let Footprint::Story(story) | Footprint::Paragraph(story, _) | Footprint::Run(story, ..) =
            footprint
            && self.formats.contains(story)
        {
            return true;
        }
        match footprint {
            Footprint::Story(story) => self
                .stories
                .get(story)
                .is_some_and(|writes| !writes.paragraphs.is_empty() || !writes.runs.is_empty()),
            Footprint::Paragraph(story, paragraph) => self
                .stories
                .get(story)
                .is_some_and(|writes| writes.whole || writes.runs.contains_key(paragraph)),
            Footprint::Run(story, paragraph, _) => self
                .stories
                .get(story)
                .is_some_and(|writes| writes.whole || writes.paragraphs.contains(paragraph)),
            Footprint::Attribute(..) => false,
            Footprint::PageName(page) => self.pages.contains(page),
            Footprint::Format(story) => self.stories.get(story).is_some_and(|writes| {
                writes.whole || !writes.paragraphs.is_empty() || !writes.runs.is_empty()
            }),
        }
    }

    fn record(&mut self, footprint: &Footprint) {
        match footprint {
            Footprint::Story(story) => self.stories.entry(story.clone()).or_default().whole = true,
            Footprint::Paragraph(story, paragraph) => {
                self.stories
                    .entry(story.clone())
                    .or_default()
                    .paragraphs
                    .insert(*paragraph);
            },
            Footprint::Run(story, paragraph, _) => {
                *self
                    .stories
                    .entry(story.clone())
                    .or_default()
                    .runs
                    .entry(*paragraph)
                    .or_default() += 1;
            },
            Footprint::Attribute(..) => {},
            Footprint::PageName(page) => {
                self.pages.insert(*page);
            },
            Footprint::Format(story) => {
                self.formats.insert(story.clone());
            },
        }
    }
}

/// A pending run of coalesced sets on one base snapshot.
pub(crate) struct Run {
    base: Odp,
    changes: Vec<Change>,
    index: Index,
    first_index: Option<usize>,
    staged: usize,
}

impl Odp {
    /// Plan a coalescable set: a direct positional path (never a query,
    /// `/metadata` or an identity predicate, which an earlier pending
    /// rename could change), resolved against this snapshot. `None` when the
    /// set must take the ordinary route.
    pub(crate) fn set_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        if crate::format::odf_metadata::is_target(target) {
            return Ok(None);
        }
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let positional = path.path().steps().iter().all(|step| {
            step.predicates()
                .iter()
                .all(|predicate| matches!(predicate, litchi_core::selector::Predicate::Index(_)))
        });
        if !positional {
            return Ok(None);
        }
        let nodes = resolve_nodes(self, target, limits)?;
        // Hyperlink sets (record 0328) take the ordinary route.
        if nodes
            .iter()
            .any(|node| node.element() == crate::format::odp::Kind::Hyperlink)
        {
            return Ok(None);
        }
        plan_set(&nodes, props, limits).map(Some)
    }
}

impl Run {
    pub(crate) fn new(base: &Odp) -> Self {
        Self {
            base: base.clone(),
            changes: Vec::new(),
            index: Index::default(),
            first_index: None,
            staged: 0,
        }
    }

    /// The snapshot the run's sets resolve against.
    pub(crate) fn base(&self) -> &Odp {
        &self.base
    }

    /// The wire index of the first staged set.
    pub(crate) fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Whether `plan` must see the published result of the pending run, or
    /// would exceed its staging bound.
    pub(crate) fn requires_barrier(&self, plan: &Plan, limits: &SemanticLimits) -> bool {
        if self.changes.len() + plan.changes.len() > limits.max_expanded_targets()
            || self.staged.saturating_add(plan.staged) > limits.max_staged_bytes()
        {
            return !self.changes.is_empty();
        }
        plan.changes.iter().any(|change| {
            Footprint::of(change).is_none_or(|footprint| self.index.overlaps(&footprint))
        })
    }

    /// Stage a plan admitted by [`Self::requires_barrier`]: a set of a text
    /// target already pending replaces the pending value.
    pub(crate) fn stage(&mut self, plan: Plan, index: usize) {
        self.first_index.get_or_insert(index);
        self.staged = self.staged.saturating_add(plan.staged);
        for change in plan.changes {
            let Some(footprint) = Footprint::of(&change) else {
                continue;
            };
            let text = matches!(
                footprint,
                Footprint::Story(_) | Footprint::Paragraph(..) | Footprint::Run(..)
            );
            if text {
                if let Some(&position) = self.index.text.get(&footprint) {
                    self.changes[position] = change;
                    continue;
                }
                self.index
                    .text
                    .insert(footprint.clone(), self.changes.len());
            }
            self.index.record(&footprint);
            self.changes.push(change);
        }
    }

    /// Publish the run as one leaf commit.
    pub(crate) fn commit(self, limits: &SemanticLimits) -> Result<(Odp, WriteCommit)> {
        let commit = commit_changes(&self.base, &self.changes, limits)?;
        Ok((self.base, commit))
    }
}
