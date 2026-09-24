//! Source-checked annotation authoring and reversible package publication.

use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::{OpcPackage, OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI};

use super::authoring::{FallbackImage, Style};
use litchi_drawingml::ink::Prepared;

use super::codec::Form;
use super::graph::{AddDelta, Binding, Delta};
use super::{CONTENT_TYPE, Limits, Location, Snapshot, host};
use crate::{Error, Package, Result};

mod authoring;
mod drawing_ids;
mod durable;

/// A paragraph in the edit's immutable story snapshot.
///
/// New annotations append a run to this paragraph. Paragraph positions count
/// active Word paragraphs in document order within the selected story.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Destination {
    story: Location,
    paragraph: Position,
}

impl Destination {
    /// Select a paragraph in a semantic story location.
    #[must_use]
    pub const fn new(story: Location, paragraph: Position) -> Self {
        Self { story, paragraph }
    }

    /// Select a paragraph in the main document story.
    #[must_use]
    pub const fn main(paragraph: Position) -> Self {
        Self::new(
            Location::new(crate::package::story::StoryKind::Main, Position::new(0)),
            paragraph,
        )
    }

    /// Selected story role and position.
    #[must_use]
    pub const fn story(self) -> Location {
        self.story
    }

    /// Selected zero-based active paragraph position within the story.
    #[must_use]
    pub const fn paragraph(self) -> Position {
        self.paragraph
    }
}

/// Finite limits for a source-bound annotation edit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EditLimits {
    /// Annotation, story, payload and graph metadata limits.
    pub inventory: Limits,
    /// Maximum aggregate uncompressed Part bytes retained by the edit source.
    pub max_package_bytes: usize,
    /// Conservative aggregate byte budget for retained story and graph changes.
    ///
    /// Shared token/payload storage can be counted more than once; this is a
    /// staging bound, not a measurement of process memory.
    pub max_staged_bytes: usize,
    /// Maximum annotation insertions, replacements and removals in one edit.
    pub max_operations: usize,
}

impl Default for EditLimits {
    fn default() -> Self {
        Self {
            inventory: Limits::default(),
            max_package_bytes: 128 * 1024 * 1024,
            max_staged_bytes: 128 * 1024 * 1024,
            max_operations: 16_384,
        }
    }
}

impl EditLimits {
    fn validate(self) -> Result<Self> {
        self.inventory.validate()?;
        for (resource, actual, maximum) in [
            (
                "edit package bytes",
                self.max_package_bytes,
                512 * 1024 * 1024,
            ),
            (
                "edit staged bytes",
                self.max_staged_bytes,
                512 * 1024 * 1024,
            ),
            ("edit operations", self.max_operations, 65_536),
        ] {
            if actual == 0 || actual > maximum {
                return Err(Error::Invalid(format!(
                    "DOCX Ink {resource} must be in 1..={maximum}"
                )));
            }
        }
        Ok(self)
    }
}

struct State {
    package: OpcPackage,
    binding: Binding,
    snapshot: Snapshot,
}

/// Isolated annotation edit using positions from its immutable base snapshot.
///
/// Removing an annotation removes its complete modeled host and exclusively
/// owned resources. Unsupported multi-object/fallback dependencies are refused.
/// Insertions and replacements consume detached validated InkML values. Native
/// resource identities are allocated only inside the source-bound commit.
pub struct Edit {
    base: Arc<State>,
    removals: Vec<usize>,
    additions: Vec<Addition>,
    replacements: Vec<Replacement>,
    input_bytes: usize,
    limits: EditLimits,
}

/// The source-bound semantic request retained beside a committed patch.
///
/// Durable replay uses this typed request to drive the same [`Edit`] API that
/// authored the original package transition.  Native relationship names and
/// story XML are intentionally absent from this value.
#[derive(Clone)]
pub(crate) enum SemanticIntent {
    Insert {
        destination: Destination,
        payload: Prepared,
        style: Style,
    },
    Replace {
        position: usize,
        payload: Prepared,
        fallback: Option<FallbackImage>,
    },
    Remove {
        position: usize,
    },
}

impl Edit {
    fn take_semantic_intent(&mut self) -> Result<Arc<Vec<SemanticIntent>>> {
        let count = self.operation_count();
        let mut intents = Vec::new();
        intents
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink durable semantic intents",
                source,
            })?;
        intents.extend(
            std::mem::take(&mut self.removals)
                .into_iter()
                .map(|position| SemanticIntent::Remove { position }),
        );
        intents.extend(
            std::mem::take(&mut self.replacements)
                .into_iter()
                .map(|replacement| SemanticIntent::Replace {
                    position: replacement.position,
                    payload: replacement.payload,
                    fallback: replacement.fallback,
                }),
        );
        intents.extend(
            std::mem::take(&mut self.additions)
                .into_iter()
                .map(|addition| SemanticIntent::Insert {
                    destination: addition.destination,
                    payload: addition.payload,
                    style: addition.style,
                }),
        );
        Ok(Arc::new(intents))
    }
}

struct Addition {
    destination: Destination,
    payload: Prepared,
    style: Style,
}

struct Replacement {
    position: usize,
    payload: Prepared,
    fallback: Option<FallbackImage>,
}

impl Edit {
    /// Borrow the immutable annotation inventory used by every selector.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.base.snapshot
    }

    /// Append a new annotation run to a paragraph in the base story snapshot.
    ///
    /// Repeated insertions at one destination retain call order. The story and
    /// paragraph selector, host grammar and resource closure are resolved at
    /// commit, before publishing any package state.
    ///
    /// # Errors
    /// Returns an error when payload, operation or retained-input limits exceed
    /// the edit's policy. A refused insertion does not change the draft.
    pub fn insert(
        &mut self,
        destination: Destination,
        payload: Prepared,
        style: Style,
    ) -> Result<()> {
        self.check_operation_count(self.operation_count().saturating_add(1))?;
        bound(
            "payload bytes",
            payload.as_bytes().len(),
            self.limits.inventory.max_payload_bytes,
        )?;
        let charge = payload
            .as_bytes()
            .len()
            .saturating_add(style.fallback().map_or(0, |image| image.as_bytes().len()));
        let input_bytes = self.input_bytes.saturating_add(charge);
        bound(
            "edit staged bytes",
            input_bytes,
            self.limits.max_staged_bytes,
        )?;
        self.additions
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink insertion operations",
                source,
            })?;
        self.additions.push(Addition {
            destination,
            payload,
            style,
        });
        self.input_bytes = input_bytes;
        Ok(())
    }

    /// Replace only the selected annotation's payload, retaining its host form.
    ///
    /// Changed drawing payloads require a complete replacement fallback image;
    /// the existing fallback's presentation geometry and XML are retained while
    /// its image relationship is retargeted. Base content parts take no fallback.
    /// Shared targets are cloned and only
    /// this anchor is retargeted. Equal payload bytes with no fallback change
    /// preserve the exact source. Selectors always refer to the base snapshot.
    ///
    /// # Errors
    /// Returns an error for a missing or removed selection, or an exhausted
    /// payload/operation/input budget. Host and fallback closure is checked at
    /// commit; any refusal leaves the original package unchanged.
    pub fn replace(
        &mut self,
        position: Position,
        payload: Prepared,
        fallback: Option<FallbackImage>,
    ) -> Result<bool> {
        let annotation =
            self.base.snapshot.get(position).ok_or_else(|| {
                Error::Invalid("DOCX Ink annotation selector is out of range".into())
            })?;
        if self.removals.binary_search(&position.get()).is_ok() {
            return Err(unsupported(
                "an annotation cannot be both removed and replaced",
            ));
        }
        bound(
            "payload bytes",
            payload.as_bytes().len(),
            self.limits.inventory.max_payload_bytes,
        )?;
        let existing = self
            .replacements
            .binary_search_by_key(&position.get(), |entry| entry.position);
        let previous_charge = existing
            .as_ref()
            .ok()
            .map_or(0, |&index| replacement_charge(&self.replacements[index]));
        if annotation.document.source() == payload.as_bytes() && fallback.is_none() {
            if let Ok(index) = existing {
                self.replacements.remove(index);
                self.input_bytes -= previous_charge;
                return Ok(true);
            }
            return Ok(false);
        }
        if let Ok(index) = existing {
            let previous = &self.replacements[index];
            if previous.payload.as_bytes() == payload.as_bytes() && previous.fallback == fallback {
                return Ok(false);
            }
        }
        let replacement = Replacement {
            position: position.get(),
            payload,
            fallback,
        };
        let input_bytes = self
            .input_bytes
            .saturating_sub(previous_charge)
            .saturating_add(replacement_charge(&replacement));
        bound(
            "edit staged bytes",
            input_bytes,
            self.limits.max_staged_bytes,
        )?;
        match existing {
            Ok(index) => self.replacements[index] = replacement,
            Err(index) => {
                self.check_operation_count(self.operation_count().saturating_add(1))?;
                self.replacements
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX Ink replacement operations",
                        source,
                    })?;
                self.replacements.insert(index, replacement);
            },
        }
        self.input_bytes = input_bytes;
        Ok(true)
    }

    fn operation_count(&self) -> usize {
        self.removals
            .len()
            .saturating_add(self.additions.len())
            .saturating_add(self.replacements.len())
    }

    fn check_operation_count(&self, count: usize) -> Result<()> {
        bound("edit operations", count, self.limits.max_operations)
    }

    /// Queue removal of one annotation at its base-snapshot position.
    ///
    /// A repeated selection returns false. Position and operation bounds are
    /// checked before changing the draft. Dependency checks occur at commit.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing annotation or an exhausted operation limit.
    pub fn remove(&mut self, position: Position) -> Result<bool> {
        if self.base.snapshot.get(position).is_none() {
            return Err(Error::Invalid(
                "DOCX Ink annotation selector is out of range".into(),
            ));
        }
        let position = position.get();
        if self
            .replacements
            .binary_search_by_key(&position, |entry| entry.position)
            .is_ok()
        {
            return Err(unsupported(
                "an annotation cannot be both replaced and removed",
            ));
        }
        match self.removals.binary_search(&position) {
            Ok(_) => Ok(false),
            Err(index) => {
                bound(
                    "edit operations",
                    self.operation_count().saturating_add(1),
                    self.limits.max_operations,
                )?;
                self.removals
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX Ink removal operations",
                        source,
                    })?;
                self.removals.insert(index, position);
                Ok(true)
            },
        }
    }

    /// Stage all host and graph changes, returning an immutable reversible commit.
    ///
    /// The original package remains unchanged. Exact no-ops share their source.
    /// Signed packages must be explicitly unsigned before starting a changed edit.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported host/resource dependencies, invalid XML,
    /// resource limits, source inconsistency, or candidate readback failure.
    pub fn commit(mut self) -> Result<Commit> {
        if !self.additions.is_empty() || !self.replacements.is_empty() {
            return authoring::commit(self);
        }
        if self.removals.is_empty() {
            let intent = self.take_semantic_intent()?;
            return Ok(Commit {
                patch: Patch {
                    before: Arc::clone(&self.base),
                    after: self.base,
                    changes: Arc::new(Vec::new()),
                    reversed: false,
                    limits: self.limits,
                    intent,
                },
            });
        }
        let stories =
            crate::package::story::capture(&self.base.package, self.limits.inventory.stories)?;
        let mut changes = Vec::new();
        let mut candidate = self.base.package.clone();
        let mut ordinal = 0usize;
        let mut staged_bytes = 0usize;
        for story in stories.stories() {
            let hosts = host::capture(story.source(), stories.dialect(), self.limits.inventory)?;
            let mut edits = Vec::new();
            for entry in hosts {
                if !is_annotation(&self.base.package, story.part(), &entry.anchor)? {
                    continue;
                }
                let selected = self.removals.binary_search(&ordinal).is_ok();
                ordinal += 1;
                if !selected {
                    continue;
                }
                if !entry.removable {
                    return Err(unsupported(
                        "annotation host contains unmodeled sibling or fallback dependencies",
                    ));
                }
                let start_tag = entry.removal_start_tag.ok_or_else(|| {
                    unsupported(
                        "annotation removal requires a modeled complete host and fallback closure",
                    )
                })?;
                edits.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX Ink story removals",
                    source,
                })?;
                edits.push(OwnedElementUpdate {
                    start_tag,
                    edit: OwnedElementEdit::Remove,
                });
            }
            if edits.is_empty() {
                continue;
            }
            let before = self.base.package.source_xml_part(story.part())?;
            let initial_charge = before
                .bytes()
                .len()
                .saturating_mul(2)
                .saturating_add(
                    candidate
                        .source_content_types()?
                        .bytes()
                        .len()
                        .saturating_mul(2),
                )
                .saturating_add(
                    candidate
                        .source_relationships(story.part())?
                        .bytes()
                        .len()
                        .saturating_mul(2),
                );
            bound(
                "edit staged bytes",
                staged_bytes.saturating_add(initial_charge),
                self.limits.max_staged_bytes,
            )?;
            let after =
                before.update_elements(&edits, self.limits.inventory.stories.max_story_bytes)?;
            let mut previous_refs = host::reference_values(before.bytes(), self.limits.inventory)?;
            previous_refs.sort_unstable();
            let mut remaining_refs = host::reference_values(after.bytes(), self.limits.inventory)?;
            remaining_refs.sort_unstable();
            let owner = self.base.package.get_part(story.part())?;
            let mut removed_ids = Vec::new();
            for relationship in owner.rels().iter() {
                if previous_refs
                    .binary_search_by(|value| value.as_str().cmp(relationship.r_id()))
                    .is_ok()
                    && remaining_refs
                        .binary_search_by(|value| value.as_str().cmp(relationship.r_id()))
                        .is_err()
                {
                    removed_ids
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "DOCX Ink removed relationships",
                            source,
                        })?;
                    removed_ids.push(relationship.r_id().to_owned());
                }
            }
            removed_ids.sort_unstable();
            candidate.try_replace_owned_xml_part(before.bytes(), after.clone())?;
            let graph = Delta::remove(
                &mut candidate,
                story.part(),
                &removed_ids,
                self.limits.inventory,
            )?;
            staged_bytes = staged_bytes
                .saturating_add(before.bytes().len())
                .saturating_add(after.bytes().len())
                .saturating_add(graph.retained_bytes()?);
            bound(
                "edit staged bytes",
                staged_bytes,
                self.limits.max_staged_bytes,
            )?;
            changes.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "DOCX Ink story changes",
                source,
            })?;
            changes.push(StoryChange {
                owner: story.part().clone(),
                before,
                after,
                graph: GraphChange::Removal(graph),
            });
        }
        if ordinal != self.base.snapshot.annotations().len() {
            return Err(Error::Invalid(
                "DOCX Ink host inventory diverged from its source snapshot".into(),
            ));
        }
        let snapshot = super::package::load(&candidate, self.limits.inventory)?;
        if snapshot.annotations().len() != ordinal - self.removals.len() {
            return Err(Error::Invalid(
                "DOCX Ink removal readback changed an unselected annotation".into(),
            ));
        }
        let binding = Binding::capture(&candidate, self.limits.inventory)?;
        let intent = self.take_semantic_intent()?;
        let after = Arc::new(State {
            package: candidate,
            binding,
            snapshot,
        });
        Ok(Commit {
            patch: Patch {
                before: self.base,
                after,
                changes: Arc::new(changes),
                reversed: false,
                limits: self.limits,
                intent,
            },
        })
    }
}

struct StoryChange {
    owner: PackURI,
    before: OwnedXmlPart,
    after: OwnedXmlPart,
    graph: GraphChange,
}

enum GraphChange {
    Removal(Delta),
    Authored { added: AddDelta, removed: Delta },
}

impl GraphChange {
    fn apply(&self, package: &mut OpcPackage, reversed: bool) -> Result<()> {
        match (self, reversed) {
            (Self::Removal(delta), false) => delta.apply(package),
            (Self::Removal(delta), true) => delta.inverse().apply(package),
            (Self::Authored { added, removed }, false) => {
                added.apply(package)?;
                removed.apply(package)
            },
            (Self::Authored { added, removed }, true) => {
                removed.inverse().apply(package)?;
                added.inverse().apply(package)
            },
        }
    }
}

/// A prepared annotation result and its reversible source-checked patch.
#[derive(Clone)]
pub struct Commit {
    patch: Patch,
}

impl Commit {
    /// Whether the prepared commit changes any annotation host.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Read the prepared annotation inventory without publishing it.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.patch.after.snapshot
    }

    /// Borrow the reusable patch for explicit package publication.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// In-memory reversible annotation patch with private package/source guards.
///
/// Clones share their source and target state. Native package identities remain
/// private. The durable codec retains only the typed semantic intent and a
/// bounded inverse proof closure.
#[derive(Clone)]
pub struct Patch {
    before: Arc<State>,
    after: Arc<State>,
    changes: Arc<Vec<StoryChange>>,
    reversed: bool,
    limits: EditLimits,
    intent: Arc<Vec<SemanticIntent>>,
}

impl Patch {
    /// Whether the patch preserves its exact source.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Return the inverse while sharing retained source and target bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            changes: Arc::clone(&self.changes),
            reversed: !self.reversed,
            limits: self.limits,
            intent: Arc::clone(&self.intent),
        }
    }

    fn matches(&self, package: &OpcPackage) -> Result<bool> {
        if !self
            .before
            .binding
            .matches(package, self.limits.inventory)?
        {
            return Ok(false);
        }
        // Payloads are compared, so both sides decode the part here
        // (ADR 0030); iteration itself yields only metadata.
        for part in self.before.package.iter_parts() {
            let current = package.get_part(part.partname())?;
            let expected = self.before.package.get_part(part.partname())?.blob();
            let actual = current.blob();
            if !std::ptr::eq(expected, actual) && expected != actual {
                return Ok(false);
            }
        }
        Ok(true)
    }

    fn apply_candidate(&self, candidate: &mut OpcPackage) -> Result<()> {
        // Story changes and their graph deltas are replayed in their original
        // order, or reversed together so content-type tokens compose exactly.
        for index in 0..self.changes.len() {
            let index = if self.reversed {
                self.changes.len() - 1 - index
            } else {
                index
            };
            let change = &self.changes[index];
            let (before, after) = if self.reversed {
                (&change.after, &change.before)
            } else {
                (&change.before, &change.after)
            };
            candidate.try_replace_owned_xml_part(before.bytes(), after.clone())?;
            change.graph.apply(candidate, self.reversed)?;
            if candidate.get_part(&change.owner)?.blob() != after.bytes() {
                return Err(Error::Invalid(
                    "DOCX Ink host publication did not retain its exact target".into(),
                ));
            }
        }
        if !self
            .after
            .binding
            .matches(candidate, self.limits.inventory)?
        {
            return Err(Error::Invalid(
                "DOCX Ink graph publication did not match its target".into(),
            ));
        }
        Ok(())
    }
}

impl Package {
    /// Start an isolated annotation edit from this package's active inventory.
    ///
    /// # Errors
    /// Returns an error for dirty state, invalid ownership/XML, or resource limits.
    pub fn edit_ink(&self) -> Result<Edit> {
        self.edit_ink_with_limits(EditLimits::default())
    }

    /// Start an annotation edit with explicit source and operation bounds.
    ///
    /// # Errors
    /// Returns an error for invalid limits, dirty state, malformed data or limits.
    pub fn edit_ink_with_limits(&self, limits: EditLimits) -> Result<Edit> {
        let limits = limits.validate()?;
        self.ensure_story_opc_current("edit_ink")?;
        let package = self.opc_package();
        let mut total = 0usize;
        // The bound is over actual payload bytes, so every part is decoded
        // here (ADR 0030).
        for part in package.try_iter_parts() {
            let part = part?;
            total = total
                .checked_add(part.blob().len())
                .ok_or_else(|| Error::Invalid("DOCX Ink edit source size overflow".into()))?;
            bound("edit package bytes", total, limits.max_package_bytes)?;
        }
        let snapshot = super::package::load(package, limits.inventory)?;
        let binding = Binding::capture(package, limits.inventory)?;
        Ok(Edit {
            base: Arc::new(State {
                package: package.clone(),
                binding,
                snapshot,
            }),
            removals: Vec::new(),
            additions: Vec::new(),
            replacements: Vec::new(),
            input_bytes: 0,
            limits,
        })
    }

    /// Atomically apply a prepared annotation patch to its guarded source.
    ///
    /// Exact no-ops preserve signatures; changed signed publication requires
    /// explicit `Package::unsign` before starting the edit. A failed patch leaves
    /// the package unchanged.
    ///
    /// # Errors
    /// Returns an error for stale state, signature policy, dependency or readback failure.
    pub fn apply_ink_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        self.ensure_story_opc_current("apply_ink_patch")?;
        if !patch.matches(self.opc_package())? {
            return Err(Error::Invalid("DOCX Ink patch source is stale".into()));
        }
        if patch.is_empty() {
            return Ok(patch.after.snapshot.clone());
        }
        self.edit_semantic_opc("apply_ink_patch", |candidate| {
            patch.apply_candidate(candidate)?;
            let snapshot = super::package::load(candidate, patch.limits.inventory)?;
            if snapshot.annotations().len() != patch.after.snapshot.annotations().len() {
                return Err(Error::Invalid(
                    "DOCX Ink publication annotation readback mismatch".into(),
                ));
            }
            Ok(snapshot)
        })
    }

    /// Commit and atomically publish an isolated annotation edit.
    ///
    /// # Errors
    /// Returns an error for unsafe dependencies, stale state, or publication failure.
    pub fn publish_ink_edit(&mut self, edit: Edit) -> Result<Commit> {
        let commit = edit.commit()?;
        self.apply_ink_patch(commit.patch())?;
        Ok(commit)
    }
}

fn replacement_charge(replacement: &Replacement) -> usize {
    replacement.payload.as_bytes().len().saturating_add(
        replacement
            .fallback
            .as_ref()
            .map_or(0, |image| image.as_bytes().len()),
    )
}

fn is_annotation(
    package: &OpcPackage,
    owner: &PackURI,
    anchor: &super::codec::Anchor,
) -> Result<bool> {
    let relationship = package
        .get_part(owner)?
        .rels()
        .get(&anchor.relationship_id)
        .ok_or_else(|| {
            Error::InvalidRelationship("DOCX Ink host relationship disappeared".into())
        })?;
    let part = package.get_part(&relationship.target_partname()?)?;
    Ok(part.content_type() == CONTENT_TYPE
        || (anchor.form == Form::Base
            && part.content_type() == "text/xml"
            && super::package::has_ink_root(part.blob())?))
}

fn bound(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        Err(Error::InkLimit {
            resource,
            actual,
            maximum,
        })
    } else {
        Ok(())
    }
}

fn unsupported(reason: &'static str) -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "edit_ink",
        reason,
    }
}
