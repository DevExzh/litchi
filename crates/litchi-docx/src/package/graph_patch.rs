//! The exact, reversible package-graph patch of a package-level edit family
//! (record 0355).
//!
//! A family stages its changes on a working copy of the package; the patch
//! is the difference between the source graph and the working graph: for
//! every part whose presence, content type, payload or relationships changed,
//! its exact state before and after, and the package relationships when they
//! changed. Untouched parts share their payload allocations with the source,
//! so the difference is found without decoding them
//! ([`OpcPackage::changed_parts`]).
//!
//! Applying the patch requires the target to hold exactly the before state
//! of every recorded part (a patch made for another base is a
//! [`crate::Error::PatchConflict`]); the inverse swaps the two states.

use std::sync::Arc;

use litchi_opc::OpcPackage;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::rel::Relationships;

use crate::error::{Error, Result};

/// One part's complete observable state.
#[derive(Clone)]
struct PartState {
    content_type: Arc<str>,
    blob: Arc<Vec<u8>>,
    rels: Arc<Relationships>,
}

impl PartState {
    fn of(part: &dyn Part) -> Result<Self> {
        part.ensure_payload()?;
        Ok(Self {
            content_type: Arc::from(part.content_type()),
            blob: part.blob_arc(),
            rels: Arc::new(part.rels().clone()),
        })
    }

    fn matches(&self, part: &dyn Part) -> Result<bool> {
        if part.content_type() != &*self.content_type || !part.rels().same_entries(&self.rels) {
            return Ok(false);
        }
        part.ensure_payload()?;
        let blob = part.blob_arc();
        Ok(Arc::ptr_eq(&blob, &self.blob) || blob.as_slice() == self.blob.as_slice())
    }

    fn same(&self, other: &Self) -> bool {
        self.content_type == other.content_type
            && self.rels.same_entries(&other.rels)
            && (Arc::ptr_eq(&self.blob, &other.blob) || self.blob == other.blob)
    }
}

#[derive(Clone)]
struct PartChange {
    name: PackURI,
    before: Option<PartState>,
    after: Option<PartState>,
}

impl PartChange {
    fn inverse(&self) -> Self {
        Self {
            name: self.name.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    fn same(&self, other: &Self) -> bool {
        fn same_state(left: Option<&PartState>, right: Option<&PartState>) -> bool {
            match (left, right) {
                (Some(left), Some(right)) => left.same(right),
                (None, None) => true,
                _ => false,
            }
        }
        self.name == other.name
            && same_state(self.before.as_ref(), other.before.as_ref())
            && same_state(self.after.as_ref(), other.after.as_ref())
    }
}

/// The exact before and after state of everything an edit changed.
#[derive(Clone, Default)]
pub(crate) struct GraphPatch {
    parts: Arc<[PartChange]>,
    package_rels: Option<Arc<(Relationships, Relationships)>>,
    /// Unchanged owners of a changed part (their relationships target it):
    /// the topology the change was made in. `None` is the package itself.
    /// Each entry holds the owner's relationships in the before and the
    /// after state; replay requires the before ones.
    owners: Arc<[(Option<PackURI>, Arc<(Relationships, Relationships)>)]>,
}

impl GraphPatch {
    /// The difference between `base` and `next`, which `next` was derived
    /// from by an edit of the working copy.
    pub(crate) fn between(base: &OpcPackage, next: &OpcPackage) -> Result<Self> {
        let mut parts = Vec::new();
        for name in next.changed_parts(base) {
            let before = match base.get_part(&name) {
                Ok(part) => Some(PartState::of(part)?),
                Err(litchi_opc::error::OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            let after = match next.get_part(&name) {
                Ok(part) => Some(PartState::of(part)?),
                Err(litchi_opc::error::OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            let change = PartChange {
                name,
                before,
                after,
            };
            // A part rewritten with equal bytes is no change.
            if change
                .before
                .as_ref()
                .zip(change.after.as_ref())
                .is_some_and(|(left, right)| left.same(right))
            {
                continue;
            }
            parts.push(change);
        }
        let package_rels = (!base.rels().same_entries(next.rels()))
            .then(|| Arc::new((base.rels().clone(), next.rels().clone())));
        let owners = owners_of(base, next, &parts, package_rels.is_some());
        Ok(Self {
            parts: parts.into(),
            package_rels,
            owners: owners.into(),
        })
    }

    /// Whether the patch changes nothing.
    pub(crate) fn is_noop(&self) -> bool {
        self.parts.is_empty() && self.package_rels.is_none()
    }

    /// The patch that restores the before state.
    pub(crate) fn inverse(&self) -> Self {
        Self {
            parts: self.parts.iter().map(PartChange::inverse).collect(),
            package_rels: self
                .package_rels
                .as_ref()
                .map(|rels| Arc::new((rels.1.clone(), rels.0.clone()))),
            owners: self
                .owners
                .iter()
                .map(|(owner, rels)| (owner.clone(), Arc::new((rels.1.clone(), rels.0.clone()))))
                .collect(),
        }
    }

    /// Check that `package` holds exactly the before state.
    pub(crate) fn check(&self, package: &OpcPackage, operation: &'static str) -> Result<()> {
        const REASON: &str = "the package does not hold the patch's before state";
        if let Some(rels) = &self.package_rels
            && !package.rels().same_entries(&rels.0)
        {
            return Err(Error::conflict(operation, REASON));
        }
        for (owner, rels) in self.owners.iter() {
            let current = match owner {
                None => package.rels(),
                Some(name) => match package.part_metadata(name) {
                    Some(part) => part.rels(),
                    None => return Err(Error::conflict(operation, REASON)),
                },
            };
            if !current.same_entries(&rels.0) {
                return Err(Error::conflict(operation, REASON));
            }
        }
        for change in self.parts.iter() {
            let current = match package.get_part(&change.name) {
                Ok(part) => Some(part),
                Err(litchi_opc::error::OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            let matches = match (&change.before, current) {
                (Some(before), Some(part)) => before.matches(part)?,
                (None, None) => true,
                _ => false,
            };
            if !matches {
                return Err(Error::conflict(operation, REASON));
            }
        }
        Ok(())
    }

    /// Visit every changed part's before and after state without applying the
    /// graph mutation. Family-specific callers use this to validate a final
    /// closure ledger before `apply_to` allocates new package parts.
    pub(crate) fn for_each_part_delta(
        &self,
        mut visit: impl FnMut(
            &PackURI,
            Option<(&str, &[u8], &Relationships)>,
            Option<(&str, &[u8], &Relationships)>,
        ) -> Result<()>,
    ) -> Result<()> {
        for change in self.parts.iter() {
            let before = change.before.as_ref().map(|state| {
                (
                    state.content_type.as_ref(),
                    state.blob.as_slice(),
                    state.rels.as_ref(),
                )
            });
            let after = change.after.as_ref().map(|state| {
                (
                    state.content_type.as_ref(),
                    state.blob.as_slice(),
                    state.rels.as_ref(),
                )
            });
            visit(&change.name, before, after)?;
        }
        Ok(())
    }

    /// Validate the after-state relationship sets without applying the patch.
    pub(crate) fn validate_after_relationships(
        &self,
        mut validate: impl FnMut(Option<&PackURI>, &Relationships) -> Result<()>,
    ) -> Result<()> {
        if let Some((_, after)) = self.package_rels.as_deref() {
            validate(None, after)?;
        }
        for (owner, rels) in self.owners.iter() {
            validate(owner.as_ref(), &rels.1)?;
        }
        Ok(())
    }

    /// Write the after state into `package`, which holds the before state.
    pub(crate) fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        for change in self.parts.iter() {
            match (&change.before, &change.after) {
                (Some(_), None) => {
                    package.remove_part(&change.name);
                },
                (None, Some(after)) => {
                    let mut part = BlobPart::new_shared(
                        change.name.clone(),
                        after.content_type.to_string(),
                        Arc::clone(&after.blob),
                    );
                    *part.rels_mut() = Relationships::clone(&after.rels);
                    package.try_add_part(Box::new(part))?;
                },
                (Some(before), Some(after)) => {
                    let part = package.get_part_mut(&change.name)?;
                    if before.content_type != after.content_type {
                        part.set_content_type(after.content_type.to_string())?;
                    }
                    if !Arc::ptr_eq(&before.blob, &after.blob) {
                        part.set_blob_shared(Arc::clone(&after.blob));
                    }
                    if !before.rels.same_entries(&after.rels) {
                        *part.rels_mut() = Relationships::clone(&after.rels);
                    }
                },
                (None, None) => {},
            }
        }
        if let Some(rels) = &self.package_rels {
            *package.rels_mut() = rels.1.clone();
        }
        Ok(())
    }
}

/// The unchanged owners whose relationships target a changed part.
fn owners_of(
    base: &OpcPackage,
    next: &OpcPackage,
    parts: &[PartChange],
    package_rels_changed: bool,
) -> Vec<(Option<PackURI>, Arc<(Relationships, Relationships)>)> {
    let mut owners = Vec::new();
    if parts.is_empty() {
        return owners;
    }
    let changed = |name: &PackURI| parts.iter().any(|change| &change.name == name);
    let targets = |rels: &Relationships| {
        rels.iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| changed(&target))
        })
    };
    if !package_rels_changed && targets(base.rels()) {
        owners.push((None, Arc::new((base.rels().clone(), next.rels().clone()))));
    }
    for part in base.iter_parts() {
        let name = part.partname();
        if changed(name) || !targets(part.rels()) {
            continue;
        }
        let after = next
            .part_metadata(name)
            .map_or_else(|| part.rels().clone(), |part| part.rels().clone());
        owners.push((Some(name.clone()), Arc::new((part.rels().clone(), after))));
    }
    owners
}

impl PartialEq for GraphPatch {
    fn eq(&self, other: &Self) -> bool {
        let rels_equal = match (&self.package_rels, &other.package_rels) {
            (Some(left), Some(right)) => {
                left.0.same_entries(&right.0) && left.1.same_entries(&right.1)
            },
            (None, None) => true,
            _ => false,
        };
        rels_equal
            && self.parts.len() == other.parts.len()
            && self
                .parts
                .iter()
                .zip(other.parts.iter())
                .all(|(left, right)| left.same(right))
    }
}

impl std::fmt::Debug for GraphPatch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Content-free: part payloads may carry document text.
        formatter
            .debug_struct("GraphPatch")
            .field("parts", &self.parts.len())
            .field("package_relationships", &self.package_rels.is_some())
            .finish()
    }
}
