//! Source-bound semantic edits and reversible exact-source patches.

use super::codec::{self, Owner};
use super::model::{Id, Snapshot, State};
use crate::shape::Key;
use crate::{Error, Result};

/// Tally emitted by a successful change-tracking edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) struct Tally {
    changed_identifiers: usize,
}

impl Tally {
    /// Number of identifier elements added, replaced, or removed.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn changed_identifiers(self) -> usize {
        self.changed_identifiers
    }
}

/// Isolated edit over one immutable slide snapshot.
#[derive(Debug, Clone)]
pub(crate) struct Draft {
    original: Snapshot,
    working: State,
}

impl Draft {
    pub(super) fn new(original: Snapshot) -> Self {
        Self {
            working: original.state.clone(),
            original,
        }
    }

    /// Set or replace the slide creation identifier.
    pub(crate) fn set_creation_id(&mut self, id: impl Into<Id>) {
        *self.working.creation_id_mut() = Some(id.into());
    }

    /// Remove the slide creation identifier.
    pub(crate) fn clear_creation_id(&mut self) {
        *self.working.creation_id_mut() = None;
    }

    /// Set one selected shape's modification identifier.
    ///
    /// # Errors
    ///
    /// Returns a typed error for a missing, ambiguous, out-of-range selector,
    /// or a value already used by another shape in this slide.
    pub(crate) fn set_shape_modification_id<'a>(
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
    pub(crate) fn clear_shape_modification_id<'a>(
        &mut self,
        shape: impl Into<Key<'a>>,
    ) -> Result<()> {
        let position = resolve(&self.working, shape.into())?;
        if !self.working.shapes()[position].supports_modification_id() {
            return Err(Error::Invalid(
                "the selected extension shape has no standard p:nvPr owner".into(),
            ));
        }
        self.working.shapes_mut()[position].modification_id = None;
        Ok(())
    }

    /// Validate and consume the edit into an immutable commit.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged XML cannot be losslessly patched or
    /// does not round-trip to the requested semantic state.
    pub(crate) fn commit(self) -> Result<Prepared> {
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
                Some(id) => {
                    codec::set_shape_with_limits(&target, position, id, self.original.shape_limits)?
                },
                None => {
                    codec::remove_shape_with_limits(&target, position, self.original.shape_limits)?
                },
            };
            changed_identifiers += 1;
        }

        let snapshot = Snapshot::from_source_with_limits(
            self.original.owner.clone(),
            target.clone(),
            self.original.shape_limits,
        )?;
        if snapshot.state != self.working {
            return Err(Error::Invalid(
                "change-tracking candidate did not round-trip to the requested state".into(),
            ));
        }
        let patch = Change {
            owner: self.original.owner,
            source: self.original.source,
            target,
            before,
            after: self.working,
        };
        Ok(Prepared {
            snapshot,
            patch,
            diagnostics: Tally {
                changed_identifiers,
            },
        })
    }
}

/// Validated result of one change-tracking edit.
#[derive(Debug, Clone)]
pub(crate) struct Prepared {
    snapshot: Snapshot,
    patch: Change,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    diagnostics: Tally,
}

impl Prepared {
    /// Borrow the candidate snapshot.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source reversible patch.
    #[inline]
    #[must_use]
    pub(crate) const fn patch(&self) -> &Change {
        &self.patch
    }

    /// Return edit diagnostics.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn diagnostics(&self) -> Tally {
        self.diagnostics
    }

    pub(super) fn into_parts(self) -> (Snapshot, Change) {
        (self.snapshot, self.patch)
    }
}

/// Exact-source-checked reversible patch for one slide's identifiers.
#[derive(Debug, Clone)]
pub(crate) struct Change {
    pub(super) owner: litchi_opc::PackURI,
    pub(super) source: Vec<u8>,
    pub(super) target: Vec<u8>,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    before: State,
    after: State,
}

impl Change {
    /// Semantic state produced by the patch.
    #[inline]
    #[must_use]
    pub(crate) const fn after(&self) -> &State {
        &self.after
    }

    fn is_changed(&self) -> bool {
        self.source != self.target
    }

    /// Whether this patch leaves its source unchanged.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Construct the exact inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
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
