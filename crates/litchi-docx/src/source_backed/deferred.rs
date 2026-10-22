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

use super::{ArtifactFingerprintMemo, Package};
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
    /// The length and SHA-256 of the written artifact.
    fingerprint: OnceLock<(u64, [u8; 32])>,
}

impl Deferred {
    /// The committed artifact as a physical source, written into memory on
    /// first use.
    pub(crate) fn materialized(&self) -> litchi_opc::Result<&SourceBackedPackage> {
        if let Some(package) = self.materialized.get() {
            return Ok(package);
        }
        // A publication releases the base's clean paragraph indexes first,
        // as the stream publishers did, so their memory is free for it.
        self.base.trim_clean_paragraph_index_for_operation();
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
            None => {
                self.base.trim_clean_paragraph_index_for_operation();
                (self.publish)(writer)
            },
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
            fingerprint: OnceLock::new(),
        }));
        Ok(next)
    }
}

/// A sink that hashes and counts what it is given.
struct Hashing(sha2::Sha256, u64);

impl Write for Hashing {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        sha2::Digest::update(&mut self.0, bytes);
        self.1 += bytes.len() as u64;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Package {
    /// The SHA-256 of the complete artifact this package writes (record
    /// 0355a): the source archive for an opened package, the output for a
    /// committed or edited one.
    pub(crate) fn output_fingerprint(&self) -> Result<[u8; 32]> {
        Ok(self.output_artifact()?.1)
    }

    /// The length and SHA-256 of the complete artifact this package writes.
    /// A committed package hashes its output once (streaming, without
    /// keeping it) and remembers it.
    pub(crate) fn output_artifact(&self) -> Result<(u64, [u8; 32])> {
        if self.effective.is_none() {
            if let Some(deferred) = &self.deferred {
                if let Some(artifact) = deferred.fingerprint.get() {
                    return Ok(*artifact);
                }
            } else {
                let physical = self.physical_source()?;
                let version = physical.source_version()?;
                let artifact = physical.source_artifact();
                if let Some(ArtifactFingerprintMemo::Source(memo)) =
                    self.artifact_fingerprint.as_ref()
                    && let Some((_cached_version, fingerprint)) = memo
                        .lock()
                        .unwrap_or_else(std::sync::PoisonError::into_inner)
                        .as_ref()
                        .copied()
                        .filter(|(cached_version, _)| *cached_version == version)
                {
                    return Ok((artifact.len(), fingerprint));
                }
                let fingerprint = artifact.fingerprint()?.into_sha256();
                if let Some(ArtifactFingerprintMemo::Source(memo)) =
                    self.artifact_fingerprint.as_ref()
                {
                    *memo
                        .lock()
                        .unwrap_or_else(std::sync::PoisonError::into_inner) =
                        Some((version, fingerprint));
                }
                return Ok((artifact.len(), fingerprint));
            }
        }
        self.check_source()?;
        if let Some(ArtifactFingerprintMemo::Effective(memo)) = self.artifact_fingerprint.as_ref()
            && let Some(artifact) = memo.get()
        {
            return Ok(*artifact);
        }
        let mut sink = Hashing(<sha2::Sha256 as sha2::Digest>::new(), 0);
        self.write_to(&mut sink)?;
        let artifact = (sink.1, sha2::Digest::finalize(sink.0).into());
        if let Some(ArtifactFingerprintMemo::Effective(memo)) = self.artifact_fingerprint.as_ref() {
            let _ = memo.set(artifact);
        }
        if let (None, Some(deferred)) = (&self.effective, &self.deferred) {
            let _ = deferred.fingerprint.set(artifact);
        }
        Ok(artifact)
    }

    /// Whether `self` writes exactly the artifact `other` writes: a share of
    /// it, or a reopened copy of its bytes.
    pub(crate) fn same_artifact(&self, other: &Self) -> Result<bool> {
        if self.same_owner(other) {
            return Ok(true);
        }
        Ok(self.output_artifact()? == other.output_artifact()?)
    }

    /// The package that restores `after` over `self`, a package writing
    /// exactly `before`'s artifact (record 0355a).
    ///
    /// A share of `before` gets `after` itself. A reopened copy of an
    /// opened archive restoring an opened archive streams the original
    /// artifact through `self`'s execution context: output is charged to
    /// it, and both artifacts are authenticated again as they are written
    /// (`SourceArtifactMismatch` before any byte otherwise).
    pub(crate) fn restored(&self, before: &Self, after: &Self) -> Result<Self> {
        if self.same_owner(before) || !self.is_opened() || !after.is_opened() {
            return Ok(after.clone());
        }
        let current = self.physical_source()?.clone();
        let original = after.physical_source()?.source_artifact();
        let current_artifact = current.source_artifact();
        let proof = litchi_opc::SourceArtifactRestoreProof {
            current_len: current_artifact.len(),
            current_sha256: current_artifact.fingerprint()?,
            original_len: original.len(),
            original_sha256: original.fingerprint()?,
        };
        let max_output_bytes = proof.original_len.max(1);
        let mut next = Self::with_execution(current.clone(), self.execution.clone())?
            .with_semantic_limits(after.semantic_limits().clone());
        next.deferred = Some(Arc::new(Deferred {
            base: self.clone(),
            publish: Box::new(move |writer| {
                Ok(current.restore_source_artifact_to_stream(
                    &original,
                    proof,
                    max_output_bytes,
                    writer,
                )?)
            }),
            limits: after.physical_source()?.read_limits(),
            materialized: OnceLock::new(),
            fingerprint: OnceLock::from((proof.original_len, proof.original_sha256.into_sha256())),
        }));
        Ok(next)
    }

    /// Whether the package writes its opened source archive unchanged.
    fn is_opened(&self) -> bool {
        self.effective.is_none() && self.deferred.is_none()
    }

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

        pub use litchi_core::edit::Diagnostics;

        crate::source_backed::deferred::source_family!(
            @core $what, $change, $error, litchi_core::edit::Diagnostics,
            |change: &$change| litchi_core::edit::Diagnostics::new(
                !crate::source_backed::deferred::Change::is_noop(change)
            )
        );
    };
    // A family module that defines its own `Error`.
    ($what:literal, $change:ty, own_error $error:ty) => {
        pub use litchi_core::edit::Diagnostics;

        crate::source_backed::deferred::source_family!(
            @core $what, $change, $error, litchi_core::edit::Diagnostics,
            |change: &$change| litchi_core::edit::Diagnostics::new(
                !crate::source_backed::deferred::Change::is_noop(change)
            )
        );
    };
    // A family with its own diagnostics type, reported by the change's
    // `diagnostics()`; the module declares `Diagnostics`.
    ($what:literal, $change:ty, $error:ty, diagnostics $diag:ty) => {
        #[doc = concat!("The error of the ", $what, " family.")]
        pub type Error = $error;

        crate::source_backed::deferred::source_family!(
            @core $what, $change, $error, $diag, |change: &$change| change.diagnostics()
        );
    };
    (@core $what:literal, $change:ty, $error:ty, $diag:ty, $report:expr) => {
        #[doc = concat!(
                    "A replayable ", $what, " change of a source-backed package: the ",
                    "exact before and after state of what it changed. Replay it with the ",
                    "family's `Package::apply_*` onto a package holding the before state; ",
                    "[`Patch::inverse`] restores it. The inverse of a commit applied to that ",
                    "commit's own package restores the exact source artifact; a family that "
                    "is forward-only may refuse an inverse on another package."
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
        pub type Commit = litchi_core::edit::Commit<crate::source_backed::Package, Patch, $diag>;

        fn commit_change(
            package: &crate::source_backed::Package,
            change: $change,
        ) -> std::result::Result<Commit, $error> {
            let diagnostics = ($report)(&change);
            let next = crate::source_backed::deferred::Change::publish(&change, package)?;
            let patch = Patch {
                change,
                restore: Some((package.clone(), next.clone())),
            };
            Ok(Commit::new(next, patch, diagnostics))
        }

        fn apply_change(
            package: &crate::source_backed::Package,
            patch: &Patch,
        ) -> std::result::Result<Commit, $error> {
            // The package that writes exactly the patch's before artifact (a
            // share of it, or a reopened copy of its bytes) gets the exact
            // after package: an inverse restores the original artifact.
            if let Some((before, after)) = &patch.restore
                && package.same_artifact(before)?
            {
                return Ok(Commit::new(
                    package.restored(before, after)?,
                    patch.clone(),
                    ($report)(&patch.change),
                ));
            }
            commit_change(package, patch.change.clone())
        }
    };
}

pub(crate) use source_family;
