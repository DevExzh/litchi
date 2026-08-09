//! Source-bound semantic edits and reversible exact-source patches.

use super::codec::{self, Owner};
use super::model::{Id, Snapshot, State};
use crate::shape::Key;
use crate::{Error, Result};

/// Diagnostics emitted by a successful change-tracking edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed_identifiers: usize,
}

impl Diagnostics {
    /// Number of identifier elements added, replaced, or removed.
    #[inline]
    #[must_use]
    pub const fn changed_identifiers(self) -> usize {
        self.changed_identifiers
    }
}

/// Isolated edit over one immutable slide snapshot.
#[derive(Debug, Clone)]
pub struct Edit {
    original: Snapshot,
    working: State,
}

impl Edit {
    pub(super) fn new(original: Snapshot) -> Self {
        Self {
            working: original.state.clone(),
            original,
        }
    }

    /// Set or replace the slide creation identifier.
    pub fn set_creation_id(&mut self, id: impl Into<Id>) {
        *self.working.creation_id_mut() = Some(id.into());
    }

    /// Remove the slide creation identifier.
    pub fn clear_creation_id(&mut self) {
        *self.working.creation_id_mut() = None;
    }

    /// Set one selected shape's modification identifier.
    ///
    /// # Errors
    ///
    /// Returns a typed error for a missing, ambiguous, out-of-range selector,
    /// or a value already used by another shape in this slide.
    pub fn set_shape_modification_id<'a>(
        &mut self,
        selector: impl Into<Key<'a>>,
        requested_id: impl Into<Id>,
    ) -> Result<()> {
        let position = resolve(&self.working, selector.into())?;
        if !self.working.shapes()[position].supports_modification_id() {
            return Err(Error::Invalid(
                "the selected extension shape has no standard p:nvPr owner".into(),
            ));
        }
        let id = requested_id.into();
        if self
            .working
            .shapes()
            .iter()
            .enumerate()
            .any(|(index, candidate)| index != position && candidate.modification_id() == Some(id))
        {
            return Err(Error::Invalid(
                "shape modification identifiers must be unique within one slide".into(),
            ));
        }
        self.working.shapes_mut()[position].modification_id = Some(id);
        Ok(())
    }

    /// Remove one selected shape's modification identifier.
    ///
    /// # Errors
    ///
    /// Returns a typed error for a missing, ambiguous, or out-of-range selector.
    pub fn clear_shape_modification_id<'a>(&mut self, shape: impl Into<Key<'a>>) -> Result<()> {
        let position = resolve(&self.working, shape.into())?;
        if !self.working.shapes()[position].supports_modification_id() {
            return Err(Error::Invalid(
                "the selected extension shape has no standard p:nvPr owner".into(),
            ));
        }
        self.working.shapes_mut()[position].modification_id = None;
        Ok(())
    }

    /// Borrow the projected semantic state.
    #[inline]
    #[must_use]
    pub const fn state(&self) -> &State {
        &self.working
    }

    /// Validate and consume the edit into an immutable commit.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged XML cannot be losslessly patched or
    /// does not round-trip to the requested semantic state.
    pub fn commit(self) -> Result<Commit> {
        self.working.validate()?;
        let before = self.original.state.clone();
        let mut target = self.original.source.clone();
        let mut changed_identifiers = 0usize;

        if before.creation_id() != self.working.creation_id() {
            target = match self.working.creation_id() {
                Some(id) => codec::set(&target, Owner::Slide, id)?,
                None => codec::remove(&target, Owner::Slide)?,
            };
            changed_identifiers += 1;
        }
        for (position, (old, new)) in before
            .shapes()
            .iter()
            .zip(self.working.shapes())
            .enumerate()
        {
            if old.modification_id() == new.modification_id() {
                continue;
            }
            target = match new.modification_id() {
                Some(id) => codec::set_shape(&target, position, id)?,
                None => codec::remove_shape(&target, position)?,
            };
            changed_identifiers += 1;
        }

        let snapshot = Snapshot::from_source(self.original.owner.clone(), target.clone())?;
        if snapshot.state != self.working {
            return Err(Error::Invalid(
                "change-tracking candidate did not round-trip to the requested state".into(),
            ));
        }
        let patch = Patch {
            owner: self.original.owner,
            source: self.original.source,
            target,
            before,
            after: self.working,
        };
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                changed_identifiers,
            },
        })
    }
}

/// Validated result of one change-tracking edit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the candidate snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source reversible patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return edit diagnostics.
    #[inline]
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Whether publication would change the slide XML.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    pub(super) fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Exact-source-checked reversible patch for one slide's identifiers.
#[derive(Debug, Clone)]
pub struct Patch {
    pub(super) owner: litchi_opc::PackURI,
    pub(super) source: Vec<u8>,
    pub(super) target: Vec<u8>,
    before: State,
    after: State,
}

impl Patch {
    /// Semantic state required at the patch source.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &State {
        &self.before
    }

    /// Semantic state produced by the patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &State {
        &self.after
    }

    /// Whether source and target slide XML differ.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.source != self.target
    }

    /// Construct the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            owner: self.owner.clone(),
            source: self.target.clone(),
            target: self.source.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn resolve(state: &State, key: Key<'_>) -> Result<usize> {
    match key {
        Key::Index(index) => {
            if index >= state.shapes().len() {
                return Err(Error::Invalid(format!(
                    "shape index {index} is outside a scene of length {}",
                    state.shapes().len()
                )));
            }
            Ok(index)
        },
        Key::Name(name) => {
            let mut found = None;
            let mut matches = 0usize;
            for (index, shape) in state.shapes().iter().enumerate() {
                if shape.name() == Some(name) {
                    matches += 1;
                    found = Some(index);
                }
            }
            match (matches, found) {
                (0, _) => Err(Error::Invalid(format!("shape name '{name}' was not found"))),
                (1, Some(index)) => Ok(index),
                _ => Err(Error::Invalid(format!(
                    "shape name '{name}' is ambiguous ({matches} exact matches)"
                ))),
            }
        },
    }
}
