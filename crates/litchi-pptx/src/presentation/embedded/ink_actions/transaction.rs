//! Source-checked existing-target transactions for slide InkAction parts.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_drawingml::ink::actions::{self, Edit as ProfileEdit, Profile};
use litchi_opc::{OpcPackage, PackURI};

use super::model::{AnchorFingerprint, AnchorSelector, Snapshot};
use crate::{Error, Result};

/// Compact fingerprint of a complete slide/action source closure.
pub type Revision = u64;

/// A source-backed editor over already classified action targets.
#[derive(Clone, Debug)]
pub struct Edit {
    source: Snapshot,
    replacements: HashMap<usize, Profile>,
}

impl Edit {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            source,
            replacements: HashMap::new(),
        }
    }

    /// Immutable snapshot against which selectors and source contexts resolve.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Whether one or more target profiles differ from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.replacements.iter().any(|(index, profile)| {
            self.source
                .anchors()
                .get(*index)
                .is_some_and(|anchor| anchor.profile() != profile)
        })
    }

    /// Obtain a detached shared profile edit for one selected target.
    ///
    /// The returned edit remains independent until it is passed back through
    /// [`Self::replace_profile`] or [`Self::edit_profile`].
    pub fn profile_edit(&self, selector: AnchorSelector) -> Result<ProfileEdit> {
        let index = self.source.find(selector)?;
        actions::Edit::with_limits(
            self.source.anchors()[index].profile().clone(),
            profile_limits(self.source.limits()),
        )
        .map_err(Into::into)
    }

    /// Alias emphasizing that the returned edit is the shared DrawingML edit.
    pub fn shared_edit(&self, selector: AnchorSelector) -> Result<ProfileEdit> {
        self.profile_edit(selector)
    }

    /// Run shared action-profile mutations and stage the resulting profile.
    pub fn edit_profile<F>(&mut self, selector: AnchorSelector, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut ProfileEdit) -> Result<()>,
    {
        let index = self.source.find(selector)?;
        let mut profile_edit = actions::Edit::with_limits(
            self.source.anchors()[index].profile().clone(),
            profile_limits(self.source.limits()),
        )?;
        edit(&mut profile_edit)?;
        let committed = profile_edit.finish()?;
        self.replace_profile(selector, committed.profile().clone())
    }

    /// Replace one existing target with an already validated shared profile.
    pub fn replace_profile(&mut self, selector: AnchorSelector, profile: Profile) -> Result<bool> {
        let index = self.source.find(selector)?;
        if profile.source().len() > self.source.limits().target_bytes {
            return Err(Error::Limit {
                resource: "ink-action target bytes",
                limit: self.source.limits().target_bytes,
            });
        }
        if self.source.anchors()[index].profile() == &profile {
            self.replacements.remove(&index);
            return Ok(false);
        }
        self.replacements
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ink-action profile replacements",
                source,
            })?;
        self.replacements.insert(index, profile);
        Ok(true)
    }

    /// Parse and replace one existing target from bounded XML bytes.
    pub fn replace_profile_bytes(
        &mut self,
        selector: AnchorSelector,
        bytes: impl AsRef<[u8]>,
    ) -> Result<bool> {
        if bytes.as_ref().len() > self.source.limits().target_bytes {
            return Err(Error::Limit {
                resource: "ink-action target bytes",
                limit: self.source.limits().target_bytes,
            });
        }
        let profile = actions::read_profile(bytes.as_ref())?;
        self.replace_profile(selector, profile)
    }

    /// Commit all staged target replacements into an immutable snapshot and
    /// reversible source-checked patch.
    pub fn commit(self) -> Result<Commit> {
        if self.replacements.is_empty() {
            return Ok(Commit {
                patch: Patch {
                    before: self.source.clone(),
                    after: self.source.clone(),
                },
                snapshot: self.source,
            });
        }
        let replacement_targets = self.validate_replacements()?;
        let mut anchors = clone_slice(self.source.anchors(), "ink-action commit anchors")?;
        let mut raw_anchors = clone_slice(
            self.source.source_anchors(),
            "ink-action commit raw anchors",
        )?;
        let mut changed = false;
        for (target_name, (target_bytes, profile)) in &replacement_targets {
            for anchor in anchors
                .iter_mut()
                .filter(|anchor| anchor.target_part_name() == target_name)
            {
                if anchor.target_bytes() == target_bytes.as_ref()
                    && anchor.profile() == profile.as_ref()
                {
                    continue;
                }
                anchor.target_bytes = Arc::clone(target_bytes);
                anchor.profile = Arc::clone(profile);
                anchor.fingerprint = AnchorFingerprint::from_closure(
                    anchor.owner_xml(),
                    anchor.choice_xml(),
                    anchor.relationship_id.as_bytes(),
                    anchor.relationship_type.as_bytes(),
                    anchor.target_ref.as_bytes(),
                    anchor.target_part_name.as_str().as_bytes(),
                    anchor.target_bytes(),
                )
                .with_graph(
                    anchor.content_type.as_bytes(),
                    anchor.target_mode,
                    &anchor.inbound,
                    &anchor.outbound,
                );
                if let Some(raw) = raw_anchors
                    .iter_mut()
                    .find(|raw| raw.source_ordinal == anchor.source_ordinal)
                {
                    raw.fingerprint = anchor.fingerprint;
                }
                changed = true;
            }
        }
        if !changed {
            return Ok(Commit {
                patch: Patch {
                    before: self.source.clone(),
                    after: self.source.clone(),
                },
                snapshot: self.source,
            });
        }
        let owner_relationships = clone_slice(
            self.source.owner_relationships.as_ref(),
            "ink-action commit owner relationships",
        )?;
        let package_relationships = clone_slice(
            self.source.package_relationships.as_ref(),
            "ink-action commit package relationships",
        )?;
        let after = Snapshot::from_parts(
            self.source.slide_index,
            self.source.slide_part_name.clone(),
            Arc::clone(&self.source.source_xml),
            anchors,
            raw_anchors,
            owner_relationships,
            self.source.content_types_source.clone(),
            self.source.owner_relationship_source.clone(),
            self.source.package_relationship_source.clone(),
            package_relationships,
            self.source.read_limits,
            self.source.limits(),
        );
        Ok(Commit {
            patch: Patch {
                before: self.source,
                after: after.clone(),
            },
            snapshot: after,
        })
    }

    /// Alias for [`Self::commit`].
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    fn validate_replacements(&self) -> Result<HashMap<PackURI, (Arc<[u8]>, Arc<Profile>)>> {
        let limits = self.source.limits();
        let mut final_sizes = HashMap::<PackURI, usize>::new();
        final_sizes
            .try_reserve(self.source.anchors().len())
            .map_err(|source| Error::Allocation {
                resource: "ink-action final target size index",
                source,
            })?;
        for anchor in self.source.anchors() {
            final_sizes
                .entry(anchor.target_part_name().clone())
                .or_insert(anchor.target_bytes().len());
        }
        let mut replacement_indices = HashMap::<PackURI, usize>::new();
        replacement_indices
            .try_reserve(self.replacements.len())
            .map_err(|source| Error::Allocation {
                resource: "ink-action replacement target index",
                source,
            })?;
        for (index, profile) in &self.replacements {
            let anchor = self
                .source
                .anchors()
                .get(*index)
                .ok_or(Error::IndexOutOfBounds {
                    index: *index,
                    len: self.source.anchors().len(),
                })?;
            let size = profile.source().len();
            if size > limits.target_bytes {
                return Err(Error::Limit {
                    resource: "ink-action target bytes",
                    limit: limits.target_bytes,
                });
            }
            if let Some(existing_index) =
                replacement_indices.get(anchor.target_part_name()).copied()
            {
                let existing = self.replacements.get(&existing_index).ok_or_else(|| {
                    Error::Invalid("ink-action replacement index is missing".into())
                })?;
                if existing.source() != profile.source() {
                    return Err(Error::Invalid(
                        "ink-action shared target has conflicting replacements".into(),
                    ));
                }
            } else {
                replacement_indices.insert(anchor.target_part_name().clone(), *index);
            }
            final_sizes.insert(anchor.target_part_name().clone(), size);
        }
        let mut aggregate = 0usize;
        for size in final_sizes.values() {
            aggregate = aggregate.checked_add(*size).ok_or(Error::Limit {
                resource: "ink-action aggregate target bytes",
                limit: limits.total_target_bytes,
            })?;
            if aggregate > limits.total_target_bytes {
                return Err(Error::Limit {
                    resource: "ink-action aggregate target bytes",
                    limit: limits.total_target_bytes,
                });
            }
        }

        let mut replacement_targets = HashMap::<PackURI, (Arc<[u8]>, Arc<Profile>)>::new();
        replacement_targets
            .try_reserve(replacement_indices.len())
            .map_err(|source| Error::Allocation {
                resource: "ink-action staged target index",
                source,
            })?;
        for (target_name, index) in replacement_indices {
            let profile = self
                .replacements
                .get(&index)
                .ok_or_else(|| Error::Invalid("ink-action replacement index is missing".into()))?;
            let profile = Arc::new(profile.clone());
            replacement_targets.insert(target_name, (profile.shared_source(), profile));
        }
        Ok(replacement_targets)
    }
}

fn clone_slice<T: Clone>(values: &[T], resource: &'static str) -> Result<Vec<T>> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(values.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.extend(values.iter().cloned());
    Ok(cloned)
}

/// A committed immutable action graph and its source-checked patch.
#[derive(Clone, Debug)]
pub struct Commit {
    patch: Patch,
    snapshot: Snapshot,
}

impl Commit {
    /// Snapshot expected after publication.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible package patch represented by this commit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether publication changes any owner or target bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Consume the commit and return its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked patch for an existing action target graph.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Source snapshot required by this patch.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Resulting snapshot produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the patch is an exact semantic/source no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias used by package owners for no-op checks.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Reverse this patch while retaining exact source checks.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch to an OPC package atomically.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }
}

/// Shared action profile limits selected by the owner transaction.
#[must_use]
pub fn default_profile_limits() -> actions::Limits {
    actions::Limits::default()
}

fn profile_limits(limits: super::model::Limits) -> actions::Limits {
    let mut profile_limits = actions::Limits::default();
    profile_limits.max_output_bytes = profile_limits.max_output_bytes.min(limits.target_bytes);
    profile_limits.max_payload_bytes = profile_limits.max_payload_bytes.min(limits.target_bytes);
    profile_limits
}
