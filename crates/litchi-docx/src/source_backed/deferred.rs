//! Committed source-backed packages (record 0355a).
//!
//! A source-backed family commit returns a new [`Package`] without writing
//! anything. Its output is the family's byte-preserving source splice, run
//! when the package is written ([`Package::write_to`], [`Package::save`]),
//! so a commit followed by a save streams exactly as the former
//! `publish_*_to_stream` routes did. A read of the committed package
//! materializes that output once into memory and serves every later read
//! from it; edits of a committed package chain the same way.

use std::io::Write;
use std::sync::{Arc, OnceLock};

use litchi_opc::{SourceBackedPackage, SourceLimits};

use super::Package;
use crate::{Error, Result};

/// Writes the committed artifact to a sink.
pub(crate) type Publish = dyn Fn(&mut dyn Write) -> Result<()> + Send + Sync;

/// The pending output of a committed package.
pub(crate) struct Deferred {
    /// The package the commit was made on.
    pub(crate) base: Package,
    publish: Box<Publish>,
    limits: SourceLimits,
    materialized: OnceLock<SourceBackedPackage>,
}

impl Deferred {
    /// The committed artifact as a physical source, written into memory on
    /// first use.
    pub(crate) fn materialized(&self) -> litchi_opc::Result<&SourceBackedPackage> {
        if let Some(package) = self.materialized.get() {
            return Ok(package);
        }
        let mut bytes = Vec::new();
        (self.publish)(&mut bytes).map_err(|error| match error {
            Error::Opc(error) => error,
            other => litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: format!("materializing a committed source-backed package: {other}"),
            },
        })?;
        let package = SourceBackedPackage::from_vec_with_limits(bytes, self.limits)?;
        Ok(self.materialized.get_or_init(|| package))
    }

    /// Write the committed artifact.
    pub(crate) fn write(&self, writer: &mut dyn Write) -> Result<()> {
        match self.materialized.get() {
            // A materialized artifact is the published bytes.
            Some(package) => Ok(package
                .clone()
                .write_topology_to_stream(writer, litchi_opc::SourceTopologyPlan::new())?),
            None => (self.publish)(writer),
        }
    }
}

impl Package {
    /// A committed package over `self` whose artifact `publish` writes.
    pub(crate) fn deferred(
        &self,
        publish: impl Fn(&mut dyn Write) -> Result<()> + Send + Sync + 'static,
    ) -> Result<Self> {
        // A changed write of a signed or encrypted source is refused at
        // commit, before any output, as the writer would refuse it.
        let physical = self.physical_source()?;
        if physical.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if physical.has_encrypted_entries() {
            return Err(litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "source XML publication refuses encrypted ZIP members".into(),
            }
            .into());
        }
        let base = self.clone();
        let limits = physical.read_limits();
        let mut next =
            Self::with_execution(self.physical_source()?.clone(), self.execution.clone())?
                .with_semantic_limits(self.semantic_limits().clone());
        next.deferred = Some(Arc::new(Deferred {
            base,
            publish: Box::new(publish),
            limits,
            materialized: OnceLock::new(),
        }));
        Ok(next)
    }
}

impl Package {
    /// Whether `self` and `other` are shares of one package (the same
    /// source, caches and pending output).
    pub(crate) fn same_owner(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.paragraph_index_cache, &other.paragraph_index_cache)
    }
}

/// One family's replayable change: its detached record patch, which knows
/// its exact before state and how the package's output carries it.
pub(crate) trait Change: Clone + Send + Sync + 'static {
    /// The family's error.
    type Error: From<Error>;
    /// Whether the change leaves the package unchanged.
    fn is_noop(&self) -> bool;
    /// The change that restores the before state.
    #[must_use]
    fn inverse(&self) -> Self;
    /// Check the change against `package`'s current state and return the
    /// committed package (a share of `package` for a no-op).
    fn publish(&self, package: &Package) -> std::result::Result<Package, Self::Error>;
}

/// Declare a source-backed family's `Patch`, `Commit`, `Diagnostics`,
/// `Error` and the commit and replay functions over its [`Change`].
macro_rules! source_family {
    ($what:literal, $change:ty) => {
        crate::source_backed::deferred::source_family!($what, $change, crate::Error);
    };
    ($what:literal, $change:ty, $error:ty) => {
        #[doc = concat!("The error of the ", $what, " family.")]
        pub type Error = $error;

        crate::source_backed::deferred::source_family!(@core $what, $change, $error);
    };
    // A family module that defines its own `Error`.
    ($what:literal, $change:ty, own_error $error:ty) => {
        crate::source_backed::deferred::source_family!(@core $what, $change, $error);
    };
    (@core $what:literal, $change:ty, $error:ty) => {
        #[doc = concat!(
                    "A replayable ", $what, " change of a source-backed package: the ",
                    "exact before and after state of what it changed. Replay it with the ",
                    "family's `Package::apply_*` onto a package holding the before state; ",
                    "[`Patch::inverse`] restores it. The inverse of a commit applied to that ",
                    "commit's own package restores the exact source artifact."
                )]
        #[derive(Clone)]
        pub struct Patch {
            change: $change,
            restore: Option<(crate::source_backed::Package, crate::source_backed::Package)>,
        }

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                crate::source_backed::deferred::Change::is_noop(&self.change)
            }

            /// The patch that restores the before state.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    change: crate::source_backed::deferred::Change::inverse(&self.change),
                    restore: self
                        .restore
                        .as_ref()
                        .map(|(before, after)| (after.clone(), before.clone())),
                }
            }
        }

        impl std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter
                    .debug_struct("Patch")
                    .field("noop", &self.is_noop())
                    .finish_non_exhaustive()
            }
        }

        #[doc = concat!(
                    "A committed ", $what, " change: the new source-backed package, its ",
                    "patch and diagnostics."
                )]
        pub type Commit = litchi_core::edit::Commit<crate::source_backed::Package, Patch>;

        pub use litchi_core::edit::Diagnostics;

        fn commit_change(
            package: &crate::source_backed::Package,
            change: $change,
        ) -> std::result::Result<Commit, $error> {
            let changed = !crate::source_backed::deferred::Change::is_noop(&change);
            let next = crate::source_backed::deferred::Change::publish(&change, package)?;
            let patch = Patch {
                change,
                restore: Some((package.clone(), next.clone())),
            };
            Ok(Commit::new(
                next,
                patch,
                litchi_core::edit::Diagnostics::new(changed),
            ))
        }

        fn apply_change(
            package: &crate::source_backed::Package,
            patch: &Patch,
        ) -> std::result::Result<Commit, $error> {
            if let Some((before, after)) = &patch.restore
                && package.same_owner(before)
            {
                let changed = !patch.is_noop();
                return Ok(Commit::new(
                    after.clone(),
                    patch.clone(),
                    litchi_core::edit::Diagnostics::new(changed),
                ));
            }
            commit_change(package, patch.change.clone())
        }
    };
}

pub(crate) use source_family;
