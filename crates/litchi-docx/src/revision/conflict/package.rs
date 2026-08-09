#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Story discovery and failure-atomic package publication.

use super::model::{Binding, Source};
use super::transaction::{Commit, Patch};
use super::{Limits, Snapshot};
use crate::package::story::{
    StoryInventory as Capture, StoryLimits, StoryPart as Located, capture_with_policy,
};
use crate::{Error, Package, Result};
use litchi_opc::{OpcPackage, PackURI};
use std::sync::Arc;

/// One reachable `WordprocessingML` story and its conflict snapshot.
#[derive(Clone, Debug)]
pub struct Story {
    part: PackURI,
    content_type: String,
    snapshot: Snapshot,
}

impl Story {
    /// Canonical OPC part name that owns this story.
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    /// Declared content type of the story part.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Bounded semantic conflict snapshot for this story.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Consume the wrapper and retain the story-bound snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

impl Package {
    /// Read the conflict revisions in the resolved main-document story.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn conflicts(&self) -> Result<Snapshot> {
        self.conflicts_with_limits(Limits::default())
    }

    /// Read the main-story conflicts with explicit semantic resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn conflicts_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        self.ensure_conflict_opc_current("conflicts_with_limits")?;
        let captured = capture(self.opc_package(), limits, false)?;
        let story = captured
            .stories
            .iter()
            .find(|story| story.part == captured.main)
            .ok_or_else(|| invalid("resolved main document is missing from the story inventory"))?;
        snapshot_single(story, Arc::clone(&captured.topology), limits)
    }

    /// Read every reachable `WordprocessingML` story independently.
    ///
    /// Stories are returned with the main document first and remaining parts
    /// ordered by canonical part name. Conflict ranges never pair across two
    /// entries in this inventory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn conflict_stories(&self) -> Result<Vec<Story>> {
        self.conflict_stories_with_limits(Limits::default())
    }

    /// Read every reachable story with explicit semantic resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn conflict_stories_with_limits(&self, limits: Limits) -> Result<Vec<Story>> {
        let limits = limits.validate()?;
        self.ensure_conflict_opc_current("conflict_stories_with_limits")?;
        let captured = capture(self.opc_package(), limits, true)?;
        let mut stories = Vec::new();
        stories
            .try_reserve_exact(captured.stories.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX conflict story inventory",
                source,
            })?;
        let mut aggregate = Aggregate::default();
        for located in &captured.stories {
            let parsed = snapshot_accounted(
                located,
                Arc::clone(&captured.topology),
                limits,
                &mut aggregate,
                None,
            )?;
            stories.push(Story {
                part: located.part.clone(),
                content_type: located.content_type.clone(),
                snapshot: parsed,
            });
        }
        Ok(stories)
    }

    /// Publish a prepared story-bound conflict transaction.
    ///
    /// The commit remains reusable when publication fails.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply_conflicts(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_conflict_patch(commit.patch())
    }

    /// Publish a reversible story-bound conflict patch.
    ///
    /// Exact no-ops bypass raw OPC editing, retaining the original blob `Arc`,
    /// signatures, and managed-writer state. Real changes clone-stage the OPC
    /// graph through an internal semantic publication path, replace only the
    /// bound story blob,
    /// and reparse it before the candidate is published.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply_conflict_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        let claim = patch.claim_publication()?;
        self.ensure_conflict_opc_current("apply_conflict_patch")?;
        let binding = patch
            .story_binding()
            .ok_or_else(|| invalid("conflict patch is not bound to a package story"))?
            .clone();
        let source_limits = patch.source_limits();
        let target_limits = patch.target_limits();
        let before = patch.before_source();
        let after = patch.after_source();
        let captured = capture(self.opc_package(), source_limits, true)?;
        let current = bound_story(&captured, &binding)?;
        if current.source.as_slice() != before.as_ref() {
            return Err(invalid("conflict patch source story is stale"));
        }
        let current_selection =
            parse_selected_with_totals(&captured, current.part.as_str(), source_limits, None)?;
        if patch.is_noop() {
            claim.finalize();
            return Ok(current_selection.snapshot);
        }

        // Validate the complete target before the package or its signatures
        // can change. This is deliberately independent of XML serialization.
        preflight_source_len(after.len(), target_limits)?;
        preflight_replacement_total(&captured, current.source.len(), after.len(), target_limits)?;
        let mut replacement_aggregate = current_selection
            .totals
            .without(current_selection.selected)?;
        let expected = snapshot_accounted(
            current,
            Arc::clone(&captured.topology),
            target_limits,
            &mut replacement_aggregate,
            Some(&after),
        )?;
        if expected.source() != after.as_ref() {
            return Err(invalid("conflict patch target did not round-trip exactly"));
        }

        let target = current.part.clone();
        let replacement = source_blob(after)?;
        let expected_topology = Arc::clone(&captured.topology);
        let expected_published = expected.clone();
        let published = self.edit_semantic_opc("apply_conflict_patch", move |candidate| {
            let staged = capture(candidate, source_limits, true)?;
            if staged.topology.as_ref() != expected_topology.as_ref() {
                return Err(invalid("conflict story topology is stale"));
            }
            let source = bound_story(&staged, &binding)?;
            if source.part != target || source.source.as_slice() != before.as_ref() {
                return Err(invalid("conflict patch source story is stale"));
            }
            preflight_replacement_total(
                &staged,
                source.source.len(),
                replacement.len(),
                target_limits,
            )?;
            candidate
                .get_part_mut(&target)?
                .set_blob_shared(Arc::clone(&replacement));
            Ok(expected_published)
        })?;
        claim.finalize();
        Ok(published)
    }
}

fn capture(package: &OpcPackage, limits: Limits, enforce_story_bytes: bool) -> Result<Capture> {
    let policy = StoryLimits {
        max_stories: limits.max_stories,
        max_story_bytes: limits.max_source_bytes,
        max_total_story_bytes: limits.max_total_story_bytes,
        max_relationships_per_owner: limits.max_relationships_per_story,
        max_total_relationships: limits.max_total_relationships,
        max_topology_bytes: limits.max_topology_bytes,
        ..StoryLimits::default()
    };
    capture_with_policy(package, policy, enforce_story_bytes)
}

#[derive(Clone, Copy, Default)]
struct Aggregate {
    story_bytes: usize,
    conflicts: usize,
    ranges: usize,
    metadata_bytes: usize,
    text_segments: usize,
}

struct Selection {
    snapshot: Snapshot,
    totals: Aggregate,
    selected: Aggregate,
}

fn snapshot_single(story: &Located, topology: Arc<[u8]>, limits: Limits) -> Result<Snapshot> {
    preflight_source_len(story.source.len(), limits)?;
    Snapshot::from_source_with_parse_and_retained_limits(
        Source::Blob(Arc::clone(&story.source)),
        limits,
        limits,
    )
    .map(|snapshot| {
        snapshot.with_binding(Binding::new(
            story.part.as_str().to_owned(),
            story.content_type.clone(),
            topology,
        ))
    })
}

fn parse_selected_with_totals(
    captured: &Capture,
    selected: &str,
    limits: Limits,
    replacement: Option<&Source>,
) -> Result<Selection> {
    let mut aggregate = Aggregate::default();
    let mut selected_snapshot = None;
    let mut selected_contribution = None;
    for story in &captured.stories {
        let source = if story.part.as_str() == selected {
            replacement
        } else {
            None
        };
        let parsed = snapshot_accounted(
            story,
            Arc::clone(&captured.topology),
            limits,
            &mut aggregate,
            source,
        )?;
        if story.part.as_str() == selected {
            selected_contribution = Some(contribution(&parsed, parsed.source().len())?);
            selected_snapshot = Some(parsed);
        }
    }
    Ok(Selection {
        snapshot: selected_snapshot.ok_or_else(|| invalid("selected conflict story is missing"))?,
        totals: aggregate,
        selected: selected_contribution
            .ok_or_else(|| invalid("selected conflict story contribution is missing"))?,
    })
}

fn snapshot_accounted(
    story: &Located,
    topology: Arc<[u8]>,
    limits: Limits,
    aggregate: &mut Aggregate,
    replacement: Option<&Source>,
) -> Result<Snapshot> {
    let length = replacement.map_or_else(|| story.source.len(), Source::len);
    let mut effective = limits;
    effective.max_source_bytes = effective.max_source_bytes.min(remaining(
        limits.max_total_story_bytes,
        aggregate.story_bytes,
        "story bytes",
    )?);
    effective.max_conflicts = effective.max_conflicts.min(remaining(
        limits.max_total_conflicts,
        aggregate.conflicts,
        "conflict elements",
    )?);
    effective.max_ranges = effective.max_ranges.min(remaining(
        limits.max_total_ranges,
        aggregate.ranges,
        "conflict ranges",
    )?);
    effective.max_metadata_bytes = effective.max_metadata_bytes.min(remaining(
        limits.max_total_metadata_bytes,
        aggregate.metadata_bytes,
        "metadata bytes",
    )?);
    effective.max_text_segments = effective.max_text_segments.min(remaining(
        limits.max_total_text_segments,
        aggregate.text_segments,
        "text segments",
    )?);
    effective = effective.validate()?;
    preflight_source_len(length, effective)?;
    let source = replacement
        .cloned()
        .unwrap_or_else(|| Source::Blob(Arc::clone(&story.source)));
    let parsed = Snapshot::from_source_with_parse_and_retained_limits(source, effective, limits)?
        .with_binding(Binding::new(
            story.part.as_str().to_owned(),
            story.content_type.clone(),
            topology,
        ));

    aggregate.add(contribution(&parsed, length)?, limits)?;
    Ok(parsed)
}

impl Aggregate {
    fn add(&mut self, value: Self, limits: Limits) -> Result<()> {
        self.story_bytes = checked_total(
            self.story_bytes,
            value.story_bytes,
            limits.max_total_story_bytes,
            "story byte",
        )?;
        self.conflicts = checked_total(
            self.conflicts,
            value.conflicts,
            limits.max_total_conflicts,
            "conflict element",
        )?;
        self.ranges = checked_total(
            self.ranges,
            value.ranges,
            limits.max_total_ranges,
            "conflict range",
        )?;
        self.metadata_bytes = checked_total(
            self.metadata_bytes,
            value.metadata_bytes,
            limits.max_total_metadata_bytes,
            "metadata byte",
        )?;
        self.text_segments = checked_total(
            self.text_segments,
            value.text_segments,
            limits.max_total_text_segments,
            "text segment",
        )?;
        Ok(())
    }

    fn without(self, value: Self) -> Result<Self> {
        Ok(Self {
            story_bytes: subtract(self.story_bytes, value.story_bytes, "story bytes")?,
            conflicts: subtract(self.conflicts, value.conflicts, "conflict elements")?,
            ranges: subtract(self.ranges, value.ranges, "conflict ranges")?,
            metadata_bytes: subtract(self.metadata_bytes, value.metadata_bytes, "metadata bytes")?,
            text_segments: subtract(self.text_segments, value.text_segments, "text segments")?,
        })
    }
}

fn contribution(snapshot: &Snapshot, story_bytes: usize) -> Result<Aggregate> {
    let inventory = snapshot.inventory();
    let mut text_segments = 0usize;
    for conflict in &inventory.conflicts {
        text_segments = text_segments
            .checked_add(conflict.text_spans().len())
            .ok_or_else(|| invalid("DOCX conflict aggregate text segment count overflow"))?;
    }
    Ok(Aggregate {
        story_bytes,
        conflicts: inventory.conflicts.len(),
        ranges: inventory.ranges.len(),
        metadata_bytes: metadata_bytes(inventory)?,
        text_segments,
    })
}

fn subtract(total: usize, removed: usize, resource: &'static str) -> Result<usize> {
    total
        .checked_sub(removed)
        .ok_or_else(|| invalid(format!("DOCX conflict aggregate {resource} underflow")))
}

fn remaining(limit: usize, used: usize, resource: &'static str) -> Result<usize> {
    limit.checked_sub(used).ok_or_else(|| {
        invalid(format!(
            "DOCX conflict aggregate {resource} limit is exhausted"
        ))
    })
}

fn metadata_bytes(inventory: &super::Inventory) -> Result<usize> {
    let mut bytes = 0usize;
    for metadata in inventory
        .conflicts
        .iter()
        .map(|conflict| &conflict.metadata)
        .chain(inventory.ranges.iter().map(|range| &range.metadata))
    {
        bytes = bytes
            .checked_add(metadata.author.len())
            .and_then(|value| value.checked_add(metadata.date.as_deref().map_or(0, str::len)))
            .ok_or_else(|| invalid("DOCX conflict aggregate metadata size overflow"))?;
    }
    Ok(bytes)
}

fn preflight_source_len(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_source_bytes {
        return Err(invalid(format!(
            "DOCX conflict story has {length} bytes, exceeding {}",
            limits.max_source_bytes
        )));
    }
    Ok(())
}

fn preflight_replacement_total(
    captured: &Capture,
    before: usize,
    after: usize,
    limits: Limits,
) -> Result<()> {
    let total = captured
        .total_story_bytes
        .checked_sub(before)
        .and_then(|value| value.checked_add(after))
        .ok_or_else(|| invalid("DOCX conflict aggregate story size overflow"))?;
    if total > limits.max_total_story_bytes {
        return Err(invalid(format!(
            "DOCX conflict aggregate story bytes exceed {}",
            limits.max_total_story_bytes
        )));
    }
    Ok(())
}

fn source_blob(source: Source) -> Result<Arc<Vec<u8>>> {
    match source {
        Source::Blob(source) => Ok(source),
        Source::Detached(_) => Err(invalid(
            "package-bound conflict patch target is not OPC-native",
        )),
    }
}

fn checked_total(
    total: usize,
    added: usize,
    limit: usize,
    resource: &'static str,
) -> Result<usize> {
    let total = total
        .checked_add(added)
        .ok_or_else(|| invalid(format!("DOCX conflict aggregate {resource} count overflow")))?;
    if total > limit {
        return Err(invalid(format!(
            "DOCX conflict aggregate {resource} count exceeds {limit}"
        )));
    }
    Ok(total)
}

fn bound_story<'a>(captured: &'a Capture, binding: &Binding) -> Result<&'a Located> {
    if captured.topology.as_ref() != binding.topology() {
        return Err(invalid("conflict story topology is stale"));
    }
    let story = captured
        .stories
        .iter()
        .find(|story| story.part.as_str() == binding.part())
        .ok_or_else(|| invalid("bound conflict story part is missing"))?;
    if story.content_type != binding.content_type() {
        return Err(invalid("bound conflict story content type is stale"));
    }
    Ok(story)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
