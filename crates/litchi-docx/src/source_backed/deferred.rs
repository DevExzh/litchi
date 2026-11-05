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

use litchi_core::{ExecutionContext, OwnedSource};
use litchi_opc::{
    ArtifactStream, ArtifactStreamPolicy, ArtifactStreamReport, ArtifactStreamTarget,
    SourceBackedPackage, SourceCacheLimits, SourceLimitKind, SourceLimits, SourceReadPolicy,
};

use super::{ArtifactFingerprintMemo, Package};
use crate::{Error, Result};

/// Writes the committed artifact to a sink.
///
/// `None` is the ordinary package publication path.  A `Some` policy is used
/// when an exact cross-owner restore is nested inside another bounded OPC
/// publication; it must reach the underlying publisher so the producer's
/// output budget is not charged a second time.
pub(crate) type Publish =
    dyn for<'a> Fn(&mut dyn Write, Option<&ArtifactStreamPolicy<'a>>) -> Result<()> + Send + Sync;

/// Publishes a deferred artifact through the sealed OPC artifact target.
///
/// This is deliberately a separate route from [`Publish`]. Ordinary DOCX
/// output still owns its caller-provided `Write` sink; cross-owner artifact
/// streams must instead select one native OPC publication method and return
/// its target-bound proof.
pub(crate) type StreamPublish =
    dyn for<'a> Fn(&mut ArtifactStreamTarget<'a>) -> Result<ArtifactStreamReport> + Send + Sync;

/// The pending output of a committed package.
pub(crate) struct Deferred {
    /// The package the commit was made on.
    pub(crate) base: Package,
    publish: Box<Publish>,
    stream_publish: Box<StreamPublish>,
    limits: SourceLimits,
    cache_limits: SourceCacheLimits,
    source_read_policy: SourceReadPolicy,
    execution_context: Option<ExecutionContext>,
    materialized: OnceLock<SourceBackedPackage>,
    /// The length and SHA-256 of the written artifact.
    fingerprint: OnceLock<(u64, [u8; 32])>,
}

impl Deferred {
    /// The committed artifact as a physical source, written into memory on
    /// first use.
    pub(crate) fn materialized(&self) -> Result<&SourceBackedPackage> {
        if let Some(package) = self.materialized.get() {
            return Ok(package);
        }
        // A publication releases the base's clean paragraph indexes first,
        // as the stream publishers did, so their memory is free for it.
        self.base.trim_clean_paragraph_index_for_operation();
        let mut bytes = Vec::new();
        // Keep the producer's typed error while materializing. This boundary
        // is used by DOCX readers as well as the artifact-stream adapter, so
        // converting non-OPC failures into a generic OPC refusal would lose
        // the operation's limit, refusal, or conflict classification.
        (self.publish)(&mut bytes, None)?;
        // Retain the native package context exactly, including None. The
        // enclosing Package separately retains its facade context; substituting
        // it here would change native read and resource accounting after the
        // first materialization. The read-ahead policy also remains unchanged.
        let package = SourceBackedPackage::from_read_at(
            Arc::new(OwnedSource::new(bytes)),
            self.limits,
            self.cache_limits,
            self.source_read_policy,
            self.execution_context.clone(),
        )?;
        Ok(self.materialized.get_or_init(|| package))
    }

    /// Write the committed artifact.
    pub(crate) fn write(&self, writer: &mut dyn Write) -> Result<()> {
        self.write_with_policy(writer, None)
    }

    /// Write the committed artifact under an enclosing OPC stream policy.
    pub(crate) fn write_with_policy(
        &self,
        writer: &mut dyn Write,
        policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> Result<()> {
        match self.materialized.get() {
            // A materialized artifact is the published bytes.
            Some(package) => match policy {
                Some(policy) => Ok(package
                    .clone()
                    .write_exact_source_with_policy(writer, policy)?),
                None => Ok(package
                    .clone()
                    .write_topology_to_stream(writer, litchi_opc::SourceTopologyPlan::new())?),
            },
            None => {
                self.base.trim_clean_paragraph_index_for_operation();
                (self.publish)(writer, policy)
            },
        }
    }

    /// Publish the committed artifact through the sealed OPC target.
    pub(crate) fn stream_to(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
    ) -> Result<ArtifactStreamReport> {
        let publication = (self.stream_publish)(target)?;
        let _ = self.fingerprint.set((
            publication.written(),
            publication.fingerprint().into_sha256(),
        ));
        Ok(publication)
    }
}

impl Package {
    /// A committed package over `self` whose artifact `publish` writes.
    pub(crate) fn deferred(
        &self,
        publish: impl for<'a> Fn(&mut dyn Write, Option<&ArtifactStreamPolicy<'a>>) -> Result<()>
        + Send
        + Sync
        + 'static,
        stream_publish: impl for<'a> Fn(&mut ArtifactStreamTarget<'a>) -> Result<ArtifactStreamReport>
        + Send
        + Sync
        + 'static,
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
        let cache_limits = physical.cache_limits();
        let source_read_policy = physical.source_read_policy();
        let execution_context = physical.execution_context();
        let mut next =
            Self::with_execution(self.physical_source()?.clone(), self.execution.clone())?
                .with_semantic_limits(self.semantic_limits().clone());
        next.deferred = Some(Arc::new(Deferred {
            base,
            publish: Box::new(publish),
            stream_publish: Box::new(stream_publish),
            limits,
            cache_limits,
            source_read_policy,
            execution_context,
            materialized: OnceLock::new(),
            fingerprint: OnceLock::new(),
        }));
        Ok(next)
    }
}

impl ArtifactStream for Package {
    type Error = Error;

    fn check_fresh(&self) -> std::result::Result<(), Self::Error> {
        self.check_source()
    }

    fn stream_to(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
    ) -> std::result::Result<ArtifactStreamReport, Self::Error> {
        match (&self.effective, &self.deferred) {
            (Some(topology), _) => target.publish_topology(topology).map_err(Error::Opc),
            (None, Some(deferred)) => deferred.stream_to(target),
            (None, None) => target
                .publish_source_package(self.physical.clone())
                .map_err(Error::Opc),
        }
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
        let (length, fingerprint) = self.physical.probe_artifact_stream(self, u64::MAX, None)?;
        let artifact = (length, fingerprint.into_sha256());
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

    /// The package that restores `after` over `self`, writing the exact
    /// artifact admitted from `after` (record 0355a).
    ///
    /// A matching recipient streams the original artifact through `self`'s
    /// execution context: output is charged to it, and both artifacts are
    /// authenticated again as they are written (`SourceArtifactMismatch`
    /// before any byte otherwise).
    pub(crate) fn restored(&self, after: &Self) -> Result<Self> {
        if self.deferred.is_none()
            && after.deferred.is_none()
            && self.effective.is_none()
            && after.effective.is_none()
        {
            // Keep the ordinary opened-source path on the original bounded
            // SourceArtifact restore. It avoids hashing the same source twice
            // and preserves the established no-allocation fast path. A
            // nested policy still uses the stream seam below, so an enclosing
            // recipient publication never charges this producer's output.
            let current = self.physical.clone();
            let recipient_read_limits = current.read_limits();
            let original = after.physical.source_artifact();
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
            let current_package = self.clone();
            let original_package = after.clone();
            let cache_limits = current.cache_limits();
            let source_read_policy = current.source_read_policy();
            let execution_context = current.execution_context();
            let stream_current = current_package.clone();
            let stream_original = original_package.clone();
            let stream_physical = current.clone();
            let mut next = Self::with_execution(current.clone(), self.execution.clone())?
                .with_semantic_limits(self.semantic_limits().clone());
            next.deferred = Some(Arc::new(Deferred {
                base: self.clone(),
                publish: Box::new(move |writer, policy| match policy {
                    Some(policy) => Ok(current.restore_streamed_artifact_to_stream(
                        &current_package,
                        &original_package,
                        proof,
                        max_output_bytes,
                        writer,
                        Some(policy),
                    )?),
                    None => Ok(current.restore_source_artifact_to_stream(
                        &original,
                        proof,
                        max_output_bytes,
                        writer,
                    )?),
                }),
                stream_publish: Box::new(move |target| {
                    target
                        .publish_restore(
                            &stream_physical,
                            &stream_current,
                            &stream_original,
                            proof,
                            max_output_bytes,
                        )
                        .map_err(Error::from)
                }),
                limits: recipient_read_limits,
                cache_limits,
                source_read_policy,
                execution_context,
                materialized: OnceLock::new(),
                fingerprint: OnceLock::from((
                    proof.original_len,
                    proof.original_sha256.into_sha256(),
                )),
            }));
            return Ok(next);
        }
        // Rebuild the deferred owner from the recipient even when either
        // endpoint is already deferred or effective. Cloning `after` here
        // would retain its producer execution context, source limits and
        // semantic policy, bypassing the recipient boundary. The stream
        // adapter keeps both endpoints lazy; a deferred/effective package is
        // never materialized merely to obtain the restore proof.
        let recipient_read_limits = self.retained_read_limits();
        let current_package = self.clone();
        let original_package = after.clone();
        let (current_len, current_fingerprint) = self.physical.probe_artifact_stream(
            &current_package,
            recipient_read_limits.max_input_bytes(),
            None,
        )?;
        let (original_len, original_fingerprint) = self.physical.probe_artifact_stream(
            &original_package,
            recipient_read_limits.max_input_bytes(),
            None,
        )?;
        if original_len > recipient_read_limits.max_input_bytes() {
            return Err(litchi_opc::OpcError::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: original_len,
                maximum: recipient_read_limits.max_input_bytes(),
            }
            .into());
        }
        let proof = litchi_opc::SourceArtifactRestoreProof {
            current_len,
            current_sha256: current_fingerprint,
            original_len,
            original_sha256: original_fingerprint,
        };
        let max_output_bytes = proof.original_len.max(1);
        let recipient = self.physical.clone();
        let cache_limits = recipient.cache_limits();
        let source_read_policy = recipient.source_read_policy();
        let execution_context = recipient.execution_context();
        let stream_current = current_package.clone();
        let stream_original = original_package.clone();
        let stream_recipient = recipient.clone();
        let mut next = Self::with_execution(recipient, self.execution.clone())?
            .with_semantic_limits(self.semantic_limits().clone());
        next.deferred = Some(Arc::new(Deferred {
            base: self.clone(),
            publish: Box::new(move |writer, policy| {
                Ok(current_package
                    .physical
                    .clone()
                    .restore_streamed_artifact_to_stream(
                        &current_package,
                        &original_package,
                        proof,
                        max_output_bytes,
                        writer,
                        policy,
                    )?)
            }),
            stream_publish: Box::new(move |target| {
                target
                    .publish_restore(
                        &stream_recipient,
                        &stream_current,
                        &stream_original,
                        proof,
                        max_output_bytes,
                    )
                    .map_err(Error::from)
            }),
            limits: recipient_read_limits,
            cache_limits,
            source_read_policy,
            execution_context,
            materialized: OnceLock::new(),
            fingerprint: OnceLock::from((proof.original_len, proof.original_sha256.into_sha256())),
        }));
        Ok(next)
    }

    fn retained_read_limits(&self) -> SourceLimits {
        self.deferred
            .as_ref()
            .map(|deferred| deferred.limits)
            .unwrap_or_else(|| self.physical.read_limits())
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
    /// Whether a detached record may restore an independently opened package
    /// whose complete content matches the recorded before artifact.
    ///
    /// The exact retained owner pair is handled by the `same_owner` proof
    /// before this capability is consulted. Families whose record is
    /// forward-only can therefore keep their own inverse on its retained
    /// owner while refusing the detached-record fallback.
    fn allows_foreign_restore(&self) -> bool {
        true
    }
    /// Validate a foreign recipient before an authenticated mismatch is
    /// reported. Families may use this to preserve a source parse failure's
    /// precedence over the complete-artifact proof; the default keeps replay
    /// on its existing fast path.
    fn preflight_foreign(&self, _package: &Package) -> std::result::Result<(), Self::Error> {
        Ok(())
    }
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
                    "commit's own package restores the exact source artifact. An exact retained ",
                    "owner pair is checked before the detached-record fallback; a family that ",
                    "is forward-only may refuse that fallback on another package."
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
            // An exact retained owner pair gets the exact after package: an
            // inverse restores the original artifact without the detached
            // record path. An independently opened package may use the
            // content-matching fallback only when the change explicitly
            // permits foreign restoration.
            if let Some((before, after)) = &patch.restore {
                let same_semantic_limits = package.semantic_limits() == before.semantic_limits();
                // Cache identity proves the recorded base without hashing. A
                // changed recipient policy still runs the family publisher.
                if package.same_owner(before) {
                    if same_semantic_limits && after.semantic_limits() == package.semantic_limits() {
                        return Ok(Commit::new(
                            after.clone(),
                            patch.clone(),
                            ($report)(&patch.change),
                        ));
                    }
                } else {
                    // Forward-only redaction refuses every foreign inverse,
                    // including a different artifact. Preserve that typed
                    // refusal before attempting content authentication.
                    if !crate::source_backed::deferred::Change::allows_foreign_restore(
                        &patch.change,
                    ) {
                        return commit_change(package, patch.change.clone());
                    }
                    // A retained owner patch records a complete artifact base.
                    // A mismatch must not fall back to a scoped record patch.
                    if !package.same_artifact(before)? {
                        crate::source_backed::deferred::Change::preflight_foreign(
                            &patch.change,
                            package,
                        )?;
                        return Err(crate::Error::conflict(
                            $what,
                            "the patch base does not match the supplied source",
                        ).into());
                    }
                    if same_semantic_limits {
                        return Ok(Commit::new(
                            package.restored(after)?,
                            patch.clone(),
                            ($report)(&patch.change),
                        ));
                    }
                }
            }
            commit_change(package, patch.change.clone())
        }
    };
}

pub(crate) use source_family;

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::{
        Arc,
        atomic::{AtomicU64, AtomicUsize, Ordering},
    };

    use super::Package;
    use crate::source_backed::story_text;
    use litchi_core::{Position, ReadAt, SourceVersion};
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::constants::relationship_type as rt;
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

    fn package_bytes() -> Vec<u8> {
        let mut archive = OpcPackage::new();
        archive
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").unwrap(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>"#.to_vec(),
            )))
            .unwrap();
        archive.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        PackageWriter::to_bytes(&archive).unwrap()
    }

    fn package() -> Package {
        Package::from_bytes(&package_bytes()).unwrap()
    }

    struct CountingVersionedSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
    }

    impl CountingVersionedSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
            }
        }

        fn bump(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingVersionedSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0xD0C5,
                self.revision.load(Ordering::SeqCst),
            ))
        }
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

    #[test]
    fn stale_same_owner_replay_checks_without_materializing_deferred_output() {
        let source = Arc::new(CountingVersionedSource::new(package_bytes()));
        let package = Package::from_read_at(source.clone()).unwrap();
        let mut edit = package.edit_story_text(story_text::Selector::Main).unwrap();
        edit.replace_paragraph_text(Position::new(0), "after")
            .unwrap();
        let commit = edit.commit().unwrap();
        let deferred = commit
            .package()
            .deferred
            .as_ref()
            .expect("a source-backed edit publishes a deferred owner");
        assert!(deferred.materialized.get().is_none());
        let reads_before_replay = source.reads.load(Ordering::SeqCst);

        source.bump();
        let result = commit.package().apply_story_text(&commit.patch().inverse());
        assert!(matches!(
            result,
            Err(story_text::Error::Document(crate::Error::Opc(
                litchi_opc::OpcError::SourceChanged { .. }
            )))
        ));
        assert!(deferred.materialized.get().is_none());
        assert_eq!(
            source.reads.load(Ordering::SeqCst),
            reads_before_replay,
            "freshness proof must not materialize or read the deferred archive"
        );
    }

    fn execution() -> (
        litchi_core::CancellationSource,
        litchi_core::ExecutionContext,
    ) {
        let (cancel, token) = litchi_core::CancellationSource::pair();
        let budget = litchi_core::Budget::root(
            "deferred-context-test",
            litchi_core::Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let limits = litchi_core::ExecutionLimits::new(
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroU64::new(1 << 20).unwrap(),
            0,
        )
        .unwrap();
        (
            cancel,
            litchi_core::ExecutionContext::new(budget, token, limits),
        )
    }

    fn deferred_route(owner: &Package, route: usize) -> Package {
        if route == 1 {
            return owner.restored(&package()).unwrap();
        }
        // Exercise the shared deferred boundary directly. The compatibility
        // story editor deliberately refuses a managed native source, while
        // deferred publication itself supports independent native and facade
        // contexts.
        let physical = owner.physical_source().unwrap().clone();
        let stream_physical = physical.clone();
        let deferred = owner
            .deferred(
                move |writer, policy| match policy {
                    Some(policy) => Ok(physical
                        .clone()
                        .write_exact_source_with_policy(writer, policy)?),
                    None => Ok(physical
                        .clone()
                        .write_topology_to_stream(writer, litchi_opc::SourceTopologyPlan::new())?),
                },
                move |target| {
                    target
                        .publish_source_package(stream_physical.clone())
                        .map_err(crate::Error::from)
                },
            )
            .unwrap();
        if route == 2 {
            owner.restored(&deferred).unwrap()
        } else {
            deferred
        }
    }

    #[test]
    fn deferred_native_and_facade_contexts_remain_independent_after_materialization() {
        for route in 0..3 {
            let (native_cancel, native) = execution();
            let (facade_cancel, facade) = execution();
            let source = litchi_opc::SourceBackedPackage::from_vec(
                package_bytes(),
                litchi_opc::SourceLimits::default(),
                litchi_opc::SourceCacheLimits::default(),
                Some(native),
            )
            .unwrap();
            let package = Package::with_execution(source, Some(facade)).unwrap();
            let next = deferred_route(&package, route);
            let materialized = next.deferred.as_ref().unwrap().materialized().unwrap();
            let native_context = materialized.execution_context().unwrap();
            facade_cancel.cancel();
            assert!(
                next.check_source().is_err(),
                "facade cancellation must survive route {route}"
            );
            assert!(
                native_context.check().is_ok(),
                "native reads must retain their own context"
            );
            native_cancel.cancel();
            assert!(
                native_context.check().is_err(),
                "native cancellation must survive route {route}"
            );
        }
    }

    #[test]
    fn deferred_unmanaged_native_source_stays_unmanaged_with_a_facade_context() {
        for route in 0..3 {
            let (facade_cancel, facade) = execution();
            let source = litchi_opc::SourceBackedPackage::from_vec(
                package_bytes(),
                litchi_opc::SourceLimits::default(),
                litchi_opc::SourceCacheLimits::default(),
                None,
            )
            .unwrap();
            let package = Package::with_execution(source, Some(facade)).unwrap();
            let next = deferred_route(&package, route);
            let materialized = next.deferred.as_ref().unwrap().materialized().unwrap();
            assert!(materialized.execution_context().is_none());
            facade_cancel.cancel();
            assert!(next.check_source().is_err());
        }
    }

    #[test]
    fn retained_story_patch_rejects_a_changed_unrelated_archive_part() {
        fn with_opaque(marker: u8) -> Package {
            let mut archive =
                OpcPackage::from_vec(package_bytes(), litchi_opc::SourceLimits::default()).unwrap();
            archive
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/opaque.bin").unwrap(),
                    "application/octet-stream".to_owned(),
                    vec![marker],
                )))
                .unwrap();
            Package::from_bytes(&PackageWriter::to_bytes(&archive).unwrap()).unwrap()
        }
        let producer = with_opaque(1);
        let mut edit = producer
            .edit_story_text(story_text::Selector::Main)
            .unwrap();
        edit.replace_paragraph_text(Position::new(0), "after")
            .unwrap();
        let commit = edit.commit().unwrap();
        let foreign = with_opaque(2);
        assert!(matches!(
            foreign.apply_story_text(commit.patch()),
            Err(story_text::Error::Document(
                crate::Error::PatchConflict { .. }
            ))
        ));
    }
}
