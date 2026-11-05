//! Shared staging of the package-level edit families (record 0355).
//!
//! Every package-level family stages its changes on a working copy of the
//! source [`Package`]: cloning a package is O(1) (its graph, properties and
//! caches are shared), and the first change detaches only the OPC graph,
//! whose built-in parts keep sharing their payload allocations. A commit
//! diffs the working graph against the source ([`GraphPatch::between`]) and
//! publishes the working package, or the source itself when nothing changed,
//! so an exact no-op keeps every byte (a signature included).

use litchi_core::edit::Diagnostics;

use super::graph_patch::GraphPatch;
use crate::Package;
use crate::error::Result;

/// The source and working copy of one family edit.
#[derive(Clone)]
pub(crate) struct Staged {
    source: Package,
    working: Package,
}

impl Staged {
    pub(crate) fn new(source: &Package) -> Self {
        Self {
            source: source.clone(),
            working: source.clone(),
        }
    }

    /// The package with the staged changes.
    pub(crate) const fn working(&self) -> &Package {
        &self.working
    }

    /// The working copy, for a staging verb.
    pub(crate) fn working_mut(&mut self) -> &mut Package {
        &mut self.working
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.working = self.source.clone();
    }

    /// The published package, its exact patch and its diagnostics. A
    /// changed commit of a signed package is refused (ADR 0006): remove the
    /// signature explicitly with [`Package::unsign`] first.
    pub(crate) fn commit(self) -> Result<(Package, GraphPatch, Diagnostics)> {
        let patch = GraphPatch::between(&self.source.opc, &self.working.opc)?;
        if patch.is_noop() {
            return Ok((self.source, patch, Diagnostics::UNCHANGED));
        }
        crate::signature_policy::ensure_unsigned(&self.source.opc)?;
        Ok((self.working, patch, Diagnostics::CHANGED))
    }
}

impl Package {
    /// Replay a family's graph patch onto this package: the package must hold
    /// the patch's before state. A changed patch refuses a signed package
    /// (ADR 0006); an exact no-op returns this package unchanged.
    pub(crate) fn apply_graph_patch(
        &self,
        patch: &GraphPatch,
        operation: &'static str,
    ) -> Result<(Self, Diagnostics)> {
        if patch.is_noop() {
            return Ok((self.clone(), Diagnostics::UNCHANGED));
        }
        let mut next = self.clone();
        next.edit_semantic_opc(operation, |opc| {
            patch.check(opc, operation)?;
            patch.apply_to(opc)
        })?;
        Ok((next, Diagnostics::CHANGED))
    }
}

/// Declare a package-level family's `Patch`, `Commit`, `Diagnostics`,
/// `Error`, `Result` and `LimitKind` over the shared graph patch.
macro_rules! graph_family {
    // A family whose budgets are the package's open limits.
    ($what:literal) => {
        crate::package::family::graph_family!($what, own_limits);

        #[doc = concat!(
            "The configured budgets of the ", $what, " family. It has none of its ",
            "own yet: the package's open limits bound every part it reads."
        )]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        impl std::fmt::Display for LimitKind {
            fn fmt(&self, _formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                match *self {}
            }
        }

    };
    // A family that declares its own `LimitKind`.
    ($what:literal, own_limits) => {
        #[doc = concat!(
            "A replayable ", $what, " change: the exact before and after state of ",
            "every part and relationship the edit changed. Apply it with the ",
            "family's `Package::apply_*` to a package holding the before state; ",
            "[`Patch::inverse`] restores it."
        )]
        #[derive(Clone, Default, PartialEq)]
        pub struct Patch(pub(crate) crate::package::graph_patch::GraphPatch);

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The patch that restores the before state.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }
        }

        impl std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.debug_tuple("Patch").field(&self.0).finish()
            }
        }

        #[doc = concat!("A published ", $what, " change: the new package, its patch and diagnostics.")]
        pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!("The error of the ", $what, " family: the crate's error.")]
        pub type Error = crate::Error;

        #[doc = concat!("Result of the ", $what, " family.")]
        pub(crate) type Result<T> = std::result::Result<T, Error>;

        #[allow(dead_code, reason = "a family with its own publication path commits without staging")]
        fn commit_staged(staged: crate::package::family::Staged) -> Result<Commit> {
            let (package, patch, diagnostics) = staged.commit()?;
            Ok(Commit::new(package, Patch(patch), diagnostics))
        }

        fn apply_patch(
            package: &crate::Package,
            patch: &Patch,
            operation: &'static str,
        ) -> Result<Commit> {
            let (package, diagnostics) = package.apply_graph_patch(&patch.0, operation)?;
            Ok(Commit::new(package, patch.clone(), diagnostics))
        }
    };
}

/// Declare a package-level family's `Edit` over [`Staged`] and the
/// package entries `edit_x` and `apply_x`. Requires [`graph_family!`] in the
/// same module.
macro_rules! graph_edit {
    ($what:literal, $edit_fn:ident, $apply_fn:ident) => {
        #[doc = concat!(
                    "A staged ", $what, " change over one package snapshot. The source ",
                    "package is never changed: [`Edit::commit`] publishes a new package, ",
                    "and dropping the edit discards it."
                )]
        #[derive(Clone)]
        pub struct Edit {
            staged: crate::package::family::Staged,
        }

        impl Edit {
            /// The package with the changes staged so far.
            #[must_use]
            pub const fn package(&self) -> &crate::Package {
                self.staged.working()
            }

            /// Discard every staged change, keeping the edit.
            pub fn reset(&mut self) {
                self.staged.reset();
            }

            /// Publish the staged changes as a new package with their exact
            /// patch. Nothing staged, or changes that cancel out, publish the
            /// source package unchanged.
            ///
            /// # Errors
            ///
            /// Returns a package-graph failure while capturing the patch.
            pub fn commit(self) -> Result<Commit> {
                commit_staged(self.staged)
            }

            /// Run one staging verb on a copy of the working package and keep
            /// the copy only when the verb succeeds.
            #[allow(
                dead_code,
                reason = "a family may stage through the working package directly"
            )]
            fn stage<T>(
                &mut self,
                verb: impl FnOnce(&mut crate::Package) -> crate::Result<T>,
            ) -> Result<T> {
                let mut next = self.staged.working().clone();
                let value = verb(&mut next)?;
                *self.staged.working_mut() = next;
                Ok(value)
            }
        }

        impl std::fmt::Debug for Edit {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.debug_struct("Edit").finish_non_exhaustive()
            }
        }

        impl crate::Package {
            #[doc = concat!(
                        "Stage a ", $what, " change over this package. O(1): the edit ",
                        "shares the package until a verb changes it."
                    )]
            #[must_use]
            pub fn $edit_fn(&self) -> Edit {
                Edit {
                    staged: crate::package::family::Staged::new(self),
                }
            }

            #[doc = concat!(
                        "Replay a ", $what, " patch onto this package, which must hold the ",
                        "patch's before state."
                    )]
            ///
            /// # Errors
            ///
            /// Returns [`crate::Error::PatchConflict`] for a package that does
            /// not hold the before state, and the signed-package refusal for a
            /// changed patch on a signed package.
            pub fn $apply_fn(&self, patch: &Patch) -> Result<Commit> {
                apply_patch(self, patch, stringify!($apply_fn))
            }
        }
    };
}

pub(crate) use {graph_edit, graph_family};
