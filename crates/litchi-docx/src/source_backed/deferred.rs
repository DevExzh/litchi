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

use litchi_core::OwnedSource;
use litchi_opc::{SourceBackedPackage, SourceLimitKind, SourceLimits, SourceReadPolicy};

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
            _ => litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                // The OPC boundary cannot carry a document-family error
                // without exposing its content-bearing display text. Keep
                // this refusal fixed and content-free; typed OPC failures
                // already pass through the branch above.
                reason: "materializing a committed source-backed package failed".into(),
            },
        })?;
        // Preserve a managed recipient's execution context when the deferred
        // artifact is first admitted. The owned-vector constructor is the
        // zero-copy unmanaged path; reopening the same moved bytes through the
        // explicit context constructor keeps cancellation, work and output
        // accounting attached to the recipient for later exact restoration.
        let package = match self.base.execution_context() {
            Some(context) =>
                SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
                    Arc::new(OwnedSource::new(bytes)),
                    self.limits,
                    Default::default(),
                    SourceReadPolicy::exact(),
                    context,
                )?,
            None => SourceBackedPackage::from_vec_with_limits(bytes, self.limits)?,
        };
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
                let artifact = self.physical_source()?.source_artifact();
                return Ok((artifact.len(), artifact.fingerprint()?.into_sha256()));
            }
        }
        let mut sink = Hashing(<sha2::Sha256 as sha2::Digest>::new(), 0);
        self.write_to(&mut sink)?;
        let artifact = (sink.1, sha2::Digest::finalize(sink.0).into());
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

    /// The package that restores `after` over `self`, writing the exact
    /// artifact admitted from `after` (record 0355a).
    ///
    /// A matching recipient streams the original artifact through `self`'s
    /// execution context: output is charged to it, and both artifacts are
    /// authenticated again as they are written (`SourceArtifactMismatch`
    /// before any byte otherwise).
    pub(crate) fn restored(&self, after: &Self) -> Result<Self> {
        // Rebuild the deferred owner from the recipient even when either
        // endpoint is already deferred or effective. Cloning `after` here
        // would retain its producer execution context, source limits and
        // semantic policy, bypassing the recipient boundary.
        let current = self.physical_source()?.clone();
        let recipient_read_limits = current.read_limits();
        let original = after.physical_source()?.source_artifact();
        let original_len = original.len();
        if original_len > recipient_read_limits.max_input_bytes() {
            return Err(litchi_opc::OpcError::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: original_len,
                maximum: recipient_read_limits.max_input_bytes(),
            }
            .into());
        }
        let current_artifact = current.source_artifact();
        let proof = litchi_opc::SourceArtifactRestoreProof {
            current_len: current_artifact.len(),
            current_sha256: current_artifact.fingerprint()?,
            original_len,
            original_sha256: original.fingerprint()?,
        };
        let max_output_bytes = proof.original_len.max(1);
        let mut next = Self::with_execution(current.clone(), self.execution.clone())?
            .with_semantic_limits(self.semantic_limits().clone());
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
            limits: recipient_read_limits,
            materialized: OnceLock::new(),
            fingerprint: OnceLock::from((proof.original_len, proof.original_sha256.into_sha256())),
        }));
        Ok(next)
    }

    /// Whether `self` and `other` are shares of one package (the same source,
    /// caches, pending output, and execution context).
    ///
    /// The cache identity is minted by [`Package::with_execution`]. Package
    /// clones and [`Package::with_semantic_limits`] retain it, while every
    /// physical-source, execution-context, effective-topology, and deferred
    /// publication boundary calls `with_execution` and therefore gets a new
    /// identity. This makes the identity a proof for the same-owner replay
    /// fast path; equal bytes alone are deliberately insufficient for that
    /// path because an independent recipient must be rebuilt under its own
    /// policy and context.
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
            // The same-owner branch below bypasses artifact authentication;
            // retain the source freshness fence that the publishing path
            // would otherwise perform before it can share an owner.
            package.check_source()?;
            // The package that writes exactly the patch's before artifact (a
            // share of it, or a reopened copy of its bytes) gets the exact
            // after package: an inverse restores the original artifact.
            if let Some((before, after)) = &patch.restore {
                let same_semantic_limits = package.semantic_limits() == before.semantic_limits();
                // same_owner is the package-local provenance proof. Package
                // clones and semantic-limit projections retain the cache;
                // every physical, execution, effective-topology, and
                // deferred boundary goes through with_execution and mints a
                // fresh cache. Reuse the committed owner only under that
                // proof, so an ordinary same-owner inverse does not hash or
                // materialize its artifact again.
                if same_semantic_limits
                    && package.same_owner(before)
                    && after.semantic_limits() == package.semantic_limits()
                {
                    return Ok(Commit::new(
                        after.clone(),
                        patch.clone(),
                        ($report)(&patch.change),
                    ));
                }
                if same_semantic_limits && package.same_artifact(before)? {
                    return Ok(Commit::new(
                        package.restored(after)?,
                        patch.clone(),
                        ($report)(&patch.change),
                    ));
                }
            }
            commit_change(package, patch.change.clone())
        }
    };
}

pub(crate) use source_family;

#[cfg(test)]
mod tests {
    use super::Package;
    use crate::source_backed::story_text;
    use litchi_core::Position;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::constants::relationship_type as rt;
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

    fn package() -> Package {
        let mut archive = OpcPackage::new();
        archive
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").unwrap(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>"#.to_vec(),
            )))
            .unwrap();
        archive.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        Package::from_bytes(&PackageWriter::to_bytes(&archive).unwrap()).unwrap()
    }

    #[test]
    fn same_owner_restore_does_not_materialize_or_hash_deferred_output() {
        let package = package();
        let mut edit = package.edit_story_text(story_text::Selector::Main).unwrap();
        edit.replace_paragraph_text(Position::new(0), "after")
            .unwrap();
        let commit = edit.commit().unwrap();
        let inverse = commit.patch().inverse();

        let deferred = commit
            .package()
            .deferred
            .as_ref()
            .expect("a source-backed edit publishes a deferred owner");
        assert!(deferred.materialized.get().is_none());

        let restored = commit.package().apply_story_text(&inverse).unwrap();
        assert!(deferred.materialized.get().is_none());
        assert!(restored.package().deferred.is_none());
    }
}
