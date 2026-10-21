//! Committed source-backed workbooks (record 0360, the shape of DOCX
//! record 0355a §1).
//!
//! A source-backed family commit returns a new [`SourceBackedWorkbook`]
//! without writing anything. Its output is the family's byte-preserving
//! source splice, run when the workbook is written
//! ([`SourceBackedWorkbook::write_to`], [`SourceBackedWorkbook::save`]), so a
//! commit followed by a save streams exactly as the former
//! `publish_*_to_stream` routes did: untouched members are raw-copied from
//! the source. A read of the committed workbook (or another edit of it)
//! writes that output into memory once and serves every later read from the
//! reopened artifact.

use std::io::Write;
use std::sync::{Arc, OnceLock};

use litchi_opc::{SourceArtifactRestoreProof, SourceBackedPackage};

use super::SourceBackedWorkbook;
use crate::error::Result;

/// Writes a committed artifact from the base package to a sink.
pub(crate) type Publish = dyn Fn(&SourceBackedPackage, &mut dyn Write) -> Result<()> + Send + Sync;

/// The pending output of a committed workbook.
pub(crate) struct Deferred {
    publish: Box<Publish>,
    materialized: OnceLock<SourceBackedWorkbook>,
    /// The complete deferred artifact identity, computed at most once when a
    /// replay needs to compare owners.
    fingerprint: OnceLock<(u64, [u8; 32])>,
    /// The committed workbook catalog of a workbook-level change, which the
    /// infallible catalog accessors read without materializing the output.
    catalog: Option<Catalog>,
}

/// The catalog fields a workbook-level family can change.
#[derive(Clone)]
pub(crate) struct Catalog {
    pub(crate) defined_names: Box<[crate::raw::DefinedName]>,
    pub(crate) active_sheet: Option<usize>,
}

impl Deferred {
    /// The reopened output, once a read has materialized it.
    pub(super) fn materialized(&self) -> Option<&SourceBackedWorkbook> {
        self.materialized.get()
    }

    /// The committed catalog of a workbook-level change.
    pub(super) const fn catalog(&self) -> Option<&Catalog> {
        self.catalog.as_ref()
    }
}

impl SourceBackedWorkbook {
    /// A committed workbook over `self` whose artifact `publish` writes from
    /// `self`'s package.
    ///
    /// A changed write of a signed or encrypted source is refused here,
    /// before any output exists, as the stream writer would refuse it.
    pub(crate) fn deferred(
        &self,
        catalog_xml: Option<&[u8]>,
        publish: impl Fn(&SourceBackedPackage, &mut dyn Write) -> Result<()> + Send + Sync + 'static,
    ) -> Result<Self> {
        let base = self.view()?;
        let package = &base.inner.package;
        if package.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if package.has_encrypted_entries() {
            return Err(litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "source XML publication refuses encrypted ZIP members".into(),
            }
            .into());
        }
        let catalog = match catalog_xml {
            None => None,
            Some(xml) => {
                let parsed = crate::raw::parse_catalog(xml, &package.read_limits())?;
                Some(Catalog {
                    active_sheet: (!parsed.sheets.is_empty()).then_some(parsed.active_sheet_index),
                    defined_names: parsed.defined_names.into_boxed_slice(),
                })
            },
        };
        Ok(Self {
            inner: Arc::clone(&base.inner),
            pending: Some(Arc::new(Deferred {
                publish: Box::new(publish),
                materialized: OnceLock::new(),
                fingerprint: OnceLock::new(),
                catalog,
            })),
        })
    }

    /// The workbook every read and edit observes: `self` when opened, the
    /// reopened output (written into memory once) when committed.
    pub(crate) fn view(&self) -> Result<&Self> {
        let Some(pending) = &self.pending else {
            return Ok(self);
        };
        if let Some(view) = pending.materialized.get() {
            return Ok(view);
        }
        let mut bytes = Vec::new();
        (pending.publish)(&self.inner.package, &mut bytes)?;
        let view = Self::from_read_at_with_options(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            self.inner.options.clone(),
        )?;
        Ok(pending.materialized.get_or_init(|| view))
    }

    /// Stream this workbook's complete archive.
    ///
    /// An opened workbook writes its source archive; a committed one writes
    /// its family's source splice (or the bytes it already materialized for
    /// a read). Repeated writes of one workbook are byte-identical. A source
    /// that changed after it was opened is refused before any byte is
    /// written.
    ///
    /// # Errors
    ///
    /// Returns the OPC writer's error, a source-version error, or the
    /// family's publication error (for example a patch conflict).
    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        match &self.pending {
            None => Ok(self
                .inner
                .package
                .clone()
                .write_part_overlays_shared_to_stream(&mut writer, Vec::new())?),
            Some(pending) => match pending.materialized.get() {
                Some(view) => view.write_to(writer),
                None => (pending.publish)(&self.inner.package, &mut writer),
            },
        }
    }

    /// Serialize this workbook's complete archive into owned bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`].
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }

    /// Atomically save this workbook's archive at a path.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`], and the file-system errors of the atomic
    /// replacement.
    pub fn save(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<std::path::Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        litchi_opc::atomic::replace_with_durability::<crate::Error>(
            path.as_ref(),
            durability,
            |sink| self.write_to(sink),
        )
    }
}

impl SourceBackedWorkbook {
    /// The semantic budgets this workbook was opened with
    /// ([`crate::OpenOptions::semantic_limits`]).
    #[must_use]
    pub fn semantic_limits(&self) -> &crate::SemanticLimits {
        &self.inner.options.semantic_limits
    }

    /// The indexed package every source-backed edit reads: the reopened
    /// output of a committed workbook.
    pub(crate) fn edit_package(&self) -> Result<&SourceBackedPackage> {
        let view = self.view()?;
        view.inner.package.check_execution()?;
        Ok(&view.inner.package)
    }

    /// Return the exact artifact this owner writes, hashing a deferred
    /// publication only when an operation needs to compare it with another
    /// owner. Ordinary commits and saves continue to stream their splice.
    pub(crate) fn output_artifact(&self) -> Result<(u64, [u8; 32])> {
        if self.pending.is_none() {
            let artifact = self.inner.package.source_artifact();
            return Ok((artifact.len(), artifact.fingerprint()?.into_sha256()));
        }
        let pending = self.pending.as_ref().expect("pending checked above");
        if let Some(artifact) = pending.fingerprint.get() {
            return Ok(*artifact);
        }
        if let Some(view) = pending.materialized() {
            let artifact = view.output_artifact()?;
            let _ = pending.fingerprint.set(artifact);
            return Ok(artifact);
        }
        let mut sink = Hashing(<sha2::Sha256 as sha2::Digest>::new(), 0);
        self.write_to(&mut sink)?;
        let artifact = (sink.1, sha2::Digest::finalize(sink.0).into());
        let _ = pending.fingerprint.set(artifact);
        Ok(artifact)
    }

    /// Whether two owners represent the same source or pending publication.
    /// The pending publication is part of ownership: sharing its base source
    /// alone does not make a changed owner equal to the opened one.
    pub(crate) fn same_owner(&self, other: &Self) -> bool {
        if !Arc::ptr_eq(&self.inner, &other.inner) {
            return false;
        }
        match (&self.pending, &other.pending) {
            (None, None) => true,
            (Some(left), Some(right)) => Arc::ptr_eq(left, right),
            (None, Some(_)) | (Some(_), None) => false,
        }
    }

    /// Whether two owners write the same complete artifact.
    pub(crate) fn same_artifact(&self, other: &Self) -> Result<bool> {
        if self.same_owner(other) {
            return Ok(true);
        }
        Ok(self.output_artifact()? == other.output_artifact()?)
    }

    /// Build a deferred owner that restores `after`'s exact artifact over this
    /// owner. The source and output are authenticated by the OPC restore
    /// primitive or by the deferred publisher, and the recipient owner's
    /// execution context and source limits perform the replay.
    pub(crate) fn restored(&self, before: &Self, after: &Self) -> Result<Self> {
        if self.same_owner(before) {
            return Ok(after.clone());
        }

        // A pending owner has an indexed base package, not the bytes it will
        // publish. Materialize that owner once so the recipient package below
        // is always the exact artifact authenticated by the replay closure.
        // This is only reached by patch replay, after same_artifact has already
        // established that the recipient and the patch base are equal.
        let recipient = self.view()?.clone();
        let current = recipient.inner.package.clone();
        let expected_before = before.output_artifact()?;
        let options = recipient.inner.options.clone();
        let catalog = after_catalog(after);

        let publish: Box<Publish> = if after.is_opened() {
            let original = after.inner.package.source_artifact();
            let current_artifact = current.source_artifact();
            let proof = SourceArtifactRestoreProof {
                current_len: current_artifact.len(),
                current_sha256: current_artifact.fingerprint()?,
                original_len: original.len(),
                original_sha256: original.fingerprint()?,
            };
            let max_output_bytes = proof.original_len.max(1);
            Box::new(move |package, writer| {
                Ok(package.restore_source_artifact_to_stream(
                    &original,
                    proof,
                    max_output_bytes,
                    writer,
                )?)
            })
        } else {
            let expected_after = after.output_artifact()?;
            let pending = Arc::clone(
                after
                    .pending
                    .as_ref()
                    .expect("a non-opened owner has a deferred publication"),
            );
            Box::new(move |package, writer| {
                authenticate_artifact(package, expected_before, "current")?;
                let mut authenticated = AuthenticatedWriter::new(writer);
                (pending.publish)(package, &mut authenticated)?;
                authenticated.finish(expected_after)
            })
        };

        let mut next = SourceBackedWorkbook::open_source_backed_package(current.clone(), options)?;
        next.pending = Some(Arc::new(Deferred {
            publish,
            materialized: OnceLock::new(),
            fingerprint: OnceLock::new(),
            catalog,
        }));
        Ok(next)
    }

    /// Whether this owner is backed directly by an opened source artifact.
    fn is_opened(&self) -> bool {
        self.pending.is_none()
    }
}

fn after_catalog(workbook: &SourceBackedWorkbook) -> Option<Catalog> {
    if let Some(pending) = workbook.pending.as_ref()
        && pending.materialized().is_none()
        && let Some(catalog) = pending.catalog()
    {
        return Some(catalog.clone());
    }
    Some(Catalog {
        defined_names: workbook.current().defined_names.clone(),
        active_sheet: workbook.current().active_sheet,
    })
}

fn authenticate_artifact(
    package: &SourceBackedPackage,
    expected: (u64, [u8; 32]),
    role: &'static str,
) -> Result<()> {
    let artifact = package.source_artifact();
    if artifact.len() != expected.0 {
        return Err(litchi_opc::OpcError::SourceArtifactMismatch {
            artifact: role,
            field: "length",
        }
        .into());
    }
    if artifact.fingerprint()?.into_sha256() != expected.1 {
        return Err(litchi_opc::OpcError::SourceArtifactMismatch {
            artifact: role,
            field: "fingerprint",
        }
        .into());
    }
    Ok(())
}

/// A sink that hashes and counts a streamed deferred artifact.
struct Hashing(sha2::Sha256, u64);

impl Write for Hashing {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        sha2::Digest::update(&mut self.0, bytes);
        self.1 = self.1.saturating_add(bytes.len() as u64);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// Authenticates a deferred publisher's output while keeping it streamed.
struct AuthenticatedWriter<'a> {
    writer: &'a mut dyn Write,
    hasher: sha2::Sha256,
    length: u64,
}

impl<'a> AuthenticatedWriter<'a> {
    fn new(writer: &'a mut dyn Write) -> Self {
        Self {
            writer,
            hasher: <sha2::Sha256 as sha2::Digest>::new(),
            length: 0,
        }
    }

    fn finish(self, expected: (u64, [u8; 32])) -> Result<()> {
        if self.length != expected.0 {
            return Err(litchi_opc::OpcError::SourceArtifactMismatch {
                artifact: "after",
                field: "length",
            }
            .into());
        }
        let actual: [u8; 32] = sha2::Digest::finalize(self.hasher).into();
        if actual != expected.1 {
            return Err(litchi_opc::OpcError::SourceArtifactMismatch {
                artifact: "after",
                field: "fingerprint",
            }
            .into());
        }
        Ok(())
    }
}

impl Write for AuthenticatedWriter<'_> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.writer.write(bytes)?;
        sha2::Digest::update(&mut self.hasher, &bytes[..written]);
        self.length = self.length.saturating_add(written as u64);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.writer.flush()
    }
}

/// One source-backed family's replayable change: its detached record patch,
/// which knows its exact before state and how the workbook's output carries
/// its after state.
pub(crate) trait Change: Clone + Send + Sync + 'static {
    /// Whether the change leaves the workbook unchanged.
    fn is_noop(&self) -> bool;
    /// The change that restores the before state.
    #[must_use]
    fn inverse(&self) -> Self;
    /// Refuse a package that does not hold the change's before state.
    fn check(&self, package: &SourceBackedPackage) -> Result<()>;
    /// Check `package` again and write the after artifact: the selected
    /// parts replaced, every other member raw-copied.
    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()>;
    /// The after workbook XML of a workbook-level change, whose catalog the
    /// committed workbook's infallible accessors read.
    fn catalog_xml(&self) -> Option<&[u8]> {
        None
    }
}

/// What a committed change reports.
pub(crate) trait Report {
    /// The family's diagnostics.
    type Diagnostics;
    /// The diagnostics of committing (or replaying) the change.
    fn diagnostics(&self) -> Self::Diagnostics;
}

/// The committed workbook of `change` over `workbook`: `workbook` itself
/// for a no-op, otherwise a deferred workbook writing the change's splice.
pub(crate) fn publish<C: Change>(
    workbook: &SourceBackedWorkbook,
    change: &Arc<C>,
) -> Result<SourceBackedWorkbook> {
    if change.is_noop() {
        return Ok(workbook.clone());
    }
    let catalog = change.catalog_xml().map(<[u8]>::to_vec);
    // The committed workbook shares the patch's record; nothing is copied.
    let change = Arc::clone(change);
    workbook.deferred(catalog.as_deref(), move |package, writer| {
        change.write(package, writer)
    })
}

/// Replay `change` onto `workbook`: refuse a workbook without the before
/// state (`PatchConflict`), then publish.
pub(crate) fn replay<C: Change>(
    workbook: &SourceBackedWorkbook,
    change: &Arc<C>,
) -> Result<SourceBackedWorkbook> {
    change.check(workbook.edit_package()?)?;
    publish(workbook, change)
}

/// Declare a source-backed family's `Patch`, `Commit`, `Diagnostics` and
/// `Error` over its record [`Change`], and its commit and replay helpers.
macro_rules! source_family {
    // A family bounded only by the workbook's source limits.
    ($what:literal, $change:ty) => {
        crate::workbook::source::deferred::source_family!($what, $change, own_limits);

        #[doc = concat!(
                    "The configured budgets of the source-backed ", $what, " family. It has ",
                    "none of its own: the workbook's source limits bound every part it reads."
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
    // A family module that declares (or re-exports) its own `LimitKind`.
    ($what:literal, $change:ty, own_limits) => {
        pub use litchi_core::edit::Diagnostics;

        impl crate::workbook::source::deferred::Report for $change {
            type Diagnostics = Diagnostics;

            fn diagnostics(&self) -> Diagnostics {
                Diagnostics::new(!crate::workbook::source::deferred::Change::is_noop(self))
            }
        }

        crate::workbook::source::deferred::source_family!(@core $what, $change, Diagnostics);
    };
    // A family that declares its own `LimitKind` and a `Diagnostics` type
    // with a `changed()` accessor, reported by its change's `Report` impl.
    ($what:literal, $change:ty, own_limits, own_diagnostics) => {
        crate::workbook::source::deferred::source_family!(@core $what, $change, Diagnostics);
    };
    (@core $what:literal, $change:ty, $diag:ty) => {
        pub use crate::Error;

        #[doc = concat!(
                    "A replayable ", $what, " change of a source-backed workbook: the exact ",
                    "before and after state of the part it changed. Replay it with the ",
                    "family's `SourceBackedWorkbook::apply_*` onto a workbook holding the ",
                    "before state; [`Patch::inverse`] restores it."
                )]
        #[derive(Clone, Debug)]
        pub struct Patch(
            pub(crate) std::sync::Arc<$change>,
            pub(crate) Option<(
                crate::SourceBackedWorkbook,
                crate::SourceBackedWorkbook,
            )>,
        );

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                crate::workbook::source::deferred::Change::is_noop(&*self.0)
            }

            /// The patch that restores the before state.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(
                    std::sync::Arc::new(crate::workbook::source::deferred::Change::inverse(
                        &*self.0,
                    )),
                    self.1
                        .as_ref()
                        .map(|(before, after)| (after.clone(), before.clone())),
                )
            }
        }

        #[doc = concat!(
                    "A committed ", $what, " change: the new source-backed workbook, its ",
                    "patch and diagnostics."
                )]
        pub type Commit = litchi_core::edit::Commit<crate::SourceBackedWorkbook, Patch, $diag>;

        #[allow(dead_code, reason = "not every family commits from a staged edit")]
        fn commit_change(
            workbook: &crate::SourceBackedWorkbook,
            change: $change,
        ) -> crate::Result<Commit> {
            let change = std::sync::Arc::new(change);
            let next = crate::workbook::source::deferred::publish(workbook, &change)?;
            let diagnostics = crate::workbook::source::deferred::Report::diagnostics(&*change);
            Ok(Commit::new(
                next.clone(),
                Patch(change, Some((workbook.clone(), next))),
                diagnostics,
            ))
        }

        fn apply_change(
            workbook: &crate::SourceBackedWorkbook,
            patch: &Patch,
        ) -> crate::Result<Commit> {
            let next = if let Some((before, after)) = &patch.1
                && workbook.same_artifact(before)?
            {
                workbook.restored(before, after)?
            } else {
                crate::workbook::source::deferred::replay(workbook, &patch.0)?
            };
            let diagnostics =
                crate::workbook::source::deferred::Report::diagnostics(&*patch.0);
            Ok(Commit::new(next, patch.clone(), diagnostics))
        }
    };
}

pub(crate) use source_family;
