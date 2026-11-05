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
use std::sync::{Arc, OnceLock, Weak};

use litchi_opc::{SourceArtifact, SourceArtifactRestoreProof, SourceBackedPackage};

use super::{SourceBackedWorkbook, SourceCatalogProof};
use crate::error::Result;

/// Writes a committed artifact from the base package to a sink.
///
/// A publisher is held by both the deferred owner and its patch. The blanket
/// implementation below lets those two views share the change's existing
/// `Arc` allocation, instead of allocating one closure for every retained
/// commit. Restore publishers use the same small, typed endpoint object.
pub(crate) trait DeferredPublisher: Send + Sync {
    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()>;

    /// Return the target workbook catalog parsed under the recipient policy,
    /// when this publisher owns one. Exact replays use the recorded catalog
    /// directly only when the recipient retains its source policy.
    fn catalog(&self, _package: &SourceBackedPackage) -> Result<Option<Catalog>> {
        Ok(None)
    }

    /// Write after the recipient's complete before artifact was authenticated.
    /// Families may omit their redundant semantic selector/dependency check
    /// while retaining the same streaming writer and source fences.
    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        self.write(package, writer)
    }
}

/// The pending output of a committed workbook.
pub(crate) struct Deferred {
    publish: Arc<dyn DeferredPublisher>,
    materialized: OnceLock<SourceBackedWorkbook>,
    /// The complete deferred artifact identity, computed at most once when a
    /// replay needs to compare owners.
    fingerprint: OnceLock<(u64, [u8; 32])>,
    /// The committed workbook catalog of a workbook-level change, which the
    /// infallible catalog accessors read without materializing the output.
    catalog: Option<Catalog>,
    /// Weak owner identity used to reconstruct a deferred workbook while its
    /// pending output is still alive. The endpoint keeps only the pending
    /// token, avoiding one repeated owner weak pointer per retained patch.
    owner: Weak<super::SourceInner>,
    /// Cached exact endpoint shared by this pending owner. The endpoint is
    /// referenced by the commit that produced this owner and by the next
    /// commit's before side, so keeping one handle avoids duplicating the
    /// endpoint state in adjacent patch records.
    endpoint: OnceLock<Arc<RestoreEndpoint>>,
}

/// The compact exact-artifact endpoint retained by a source-backed patch.
///
/// A patch must retain the exact before/after publication pair, but retaining
/// the public workbook owners would also retain each owner's parsed catalog,
/// sheet handles and semantic caches. The endpoint keeps only the shared OPC
/// package handle, the source artifact or publisher, and weak owner identity
/// tokens for the same-owner fast path. `SourceBackedPackage::clone` shares
/// its index, cache and source without retaining the XLSX facade graph.
pub(crate) enum RestoreEndpoint {
    Opened {
        artifact: SourceArtifact,
        inner: Weak<super::SourceInner>,
    },
    Deferred {
        package: Arc<SourceBackedPackage>,
        publisher: Arc<dyn DeferredPublisher>,
        fingerprint: OnceLock<(u64, [u8; 32])>,
        pending: Weak<Deferred>,
    },
}

impl Clone for RestoreEndpoint {
    fn clone(&self) -> Self {
        match self {
            Self::Opened { artifact, inner } => Self::Opened {
                artifact: artifact.clone(),
                inner: inner.clone(),
            },
            Self::Deferred {
                package,
                publisher,
                fingerprint,
                pending,
            } => Self::Deferred {
                package: Arc::clone(package),
                publisher: Arc::clone(publisher),
                fingerprint: clone_fingerprint(fingerprint),
                pending: pending.clone(),
            },
        }
    }
}

fn clone_fingerprint(fingerprint: &OnceLock<(u64, [u8; 32])>) -> OnceLock<(u64, [u8; 32])> {
    let cloned = OnceLock::new();
    if let Some(value) = fingerprint.get() {
        let _ = cloned.set(*value);
    }
    cloned
}

impl std::fmt::Debug for RestoreEndpoint {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("RestoreEndpoint")
            .field(
                "kind",
                &match self {
                    Self::Opened { .. } => "opened",
                    Self::Deferred { .. } => "deferred",
                },
            )
            .finish_non_exhaustive()
    }
}

impl RestoreEndpoint {
    pub(crate) fn from_workbook(workbook: &SourceBackedWorkbook) -> Arc<Self> {
        if let Some(pending) = &workbook.pending {
            return Arc::clone(pending.endpoint.get_or_init(|| {
                Arc::new(Self::Deferred {
                    package: workbook.inner.package_handle(),
                    publisher: Arc::clone(&pending.publish),
                    fingerprint: OnceLock::new(),
                    pending: Arc::downgrade(pending),
                })
            }));
        }
        Arc::clone(workbook.inner.restore_endpoint.get_or_init(|| {
            Arc::new(Self::Opened {
                artifact: workbook.inner.package.source_artifact(),
                inner: Arc::downgrade(&workbook.inner),
            })
        }))
    }

    pub(crate) fn output_artifact(&self) -> Result<(u64, [u8; 32])> {
        match self {
            Self::Opened { artifact, .. } => {
                Ok((artifact.len(), artifact.fingerprint()?.into_sha256()))
            },
            Self::Deferred {
                package,
                publisher,
                fingerprint,
                ..
            } => {
                if let Some(artifact) = fingerprint.get() {
                    return Ok(*artifact);
                }
                let mut sink = Hashing(<sha2::Sha256 as sha2::Digest>::new(), 0);
                publisher.write(package, &mut sink)?;
                let artifact = (sink.1, sha2::Digest::finalize(sink.0).into());
                let _ = fingerprint.set(artifact);
                Ok(artifact)
            },
        }
    }

    pub(crate) fn same_owner(&self, workbook: &SourceBackedWorkbook) -> bool {
        match self {
            Self::Opened { inner, .. } => {
                workbook.pending.is_none() && inner.as_ptr() == Arc::as_ptr(&workbook.inner)
            },
            Self::Deferred { pending, .. } => {
                let Some(current) = &workbook.pending else {
                    return false;
                };
                pending.as_ptr() == Arc::as_ptr(current)
            },
        }
    }

    fn owner(&self) -> Option<SourceBackedWorkbook> {
        match self {
            Self::Opened { inner, .. } => inner.upgrade().map(|inner| SourceBackedWorkbook {
                inner,
                pending: None,
            }),
            Self::Deferred { pending, .. } => {
                let pending = pending.upgrade()?;
                let inner = pending.owner.upgrade()?;
                Some(SourceBackedWorkbook {
                    inner,
                    pending: Some(pending),
                })
            },
        }
    }

    fn artifact(&self) -> Option<&SourceArtifact> {
        match self {
            Self::Opened { artifact, .. } => Some(artifact),
            Self::Deferred { .. } => None,
        }
    }

    fn publisher(&self) -> Option<&Arc<dyn DeferredPublisher>> {
        match self {
            Self::Opened { .. } => None,
            Self::Deferred { publisher, .. } => Some(publisher),
        }
    }

    fn catalog_for(&self, package: &SourceBackedPackage) -> Result<Option<Catalog>> {
        let Self::Deferred {
            package: recorded_package,
            publisher,
            pending,
            ..
        } = self
        else {
            return Ok(None);
        };
        // The endpoint's package is the package under which its catalog was
        // parsed. Keep the policy comparison, but derive the recorded policy
        // from that already-retained package instead of copying all of the
        // SourceLimits fields into every deferred endpoint.
        if package.read_limits() == recorded_package.read_limits() {
            if let Some(catalog) = pending
                .upgrade()
                .and_then(|pending| pending.catalog().cloned())
            {
                return Ok(Some(catalog));
            }
        }
        publisher.catalog(package)
    }
}

/// The catalog fields a workbook-level family can change.
///
/// Defined names are immutable for tab-state changes and are shared through
/// an `Arc` across the snapshot, deferred owner, and exact restore endpoint.
#[derive(Clone)]
pub(crate) struct Catalog {
    pub(crate) defined_names: Arc<[crate::raw::DefinedName]>,
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
    fn deferred_from_base(
        base: &Self,
        catalog: Option<Catalog>,
        publish: Arc<dyn DeferredPublisher>,
    ) -> Self {
        Self {
            inner: Arc::clone(&base.inner),
            pending: Some(Arc::new(Deferred {
                publish,
                materialized: OnceLock::new(),
                fingerprint: OnceLock::new(),
                catalog,
                owner: Arc::downgrade(&base.inner),
                endpoint: OnceLock::new(),
            })),
        }
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
        // Preserve the established source-first error precedence, and fence
        // the deferred publication before it allocates an output buffer. A
        // facade context is an additional policy boundary around an external
        // package; checking it only after reopening would publish bytes after
        // cancellation.
        self.inner.package.source_version()?;
        self.inner.execution_check()?;
        // A committed workbook either carries the catalog captured while its
        // change was authored, or changes only non-workbook parts. Reuse one
        // immutable proof whenever the output keeps the same catalog. The
        // proof carries the lineage's policy Arc as well as its names: a
        // OnceLock local to each reopened SourceInner would otherwise create
        // one fresh policy allocation for every materialized history step.
        let lineage_limits = self.inner.source_limits();
        let shared_catalog = match pending.catalog.as_ref() {
            None => self.inner.catalog_proof(),
            Some(catalog) if catalog.defined_names.as_ref() == self.inner.defined_names_slice() => {
                self.inner.catalog_proof()
            },
            Some(catalog) => Arc::new(SourceCatalogProof {
                defined_names: Arc::clone(&catalog.defined_names),
                limits: Arc::clone(&lineage_limits),
            }),
        };
        let mut bytes = Vec::new();
        pending.publish.write(&self.inner.package, &mut bytes)?;
        // A facade built around an externally supplied package may have
        // options whose source profile differs from that package's admitted
        // profile. Materialization follows the package lineage, so preserve
        // the retained policy while carrying the other facade options over.
        let options = self
            .inner
            .owner_options_ref()
            .to_open_options(*lineage_limits);
        let view = Self::from_read_at_with_options_and_shared_catalog(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            options,
            Some(shared_catalog),
            Some(self.inner.owner_options()),
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
        // Check the original source before any output, then apply both the
        // package-owned and optional facade-owned execution fences. In
        // particular, an uncached deferred owner must not publish into the
        // caller's sink after its facade context has been cancelled.
        self.inner.package.source_version()?;
        self.inner.execution_check()?;
        match &self.pending {
            None => Ok(self
                .inner
                .package
                .clone()
                .write_part_overlays_shared_to_stream(&mut writer, Vec::new())?),
            Some(pending) => match pending.materialized.get() {
                Some(view) => view.write_to(writer),
                None => pending.publish.write(&self.inner.package, &mut writer),
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
        self.inner.owner_options_ref().semantic_limits()
    }

    /// The indexed package every source-backed edit reads: the reopened
    /// output of a committed workbook.
    pub(crate) fn edit_package(&self) -> Result<&SourceBackedPackage> {
        let view = self.view()?;
        view.inner.package.check_execution()?;
        Ok(&view.inner.package)
    }

    /// Check source and execution fences without materializing a pending
    /// publication. Lazy source-backed edit entries use this for bounded
    /// caller-input admission; payload resolution remains in `LazySource`.
    pub(crate) fn check_execution(&self) -> Result<()> {
        self.inner.execution_check()
    }

    pub(crate) fn edit_source(&self) -> Result<(&SourceBackedPackage, &super::SourceInner)> {
        let view = self.view()?;
        view.inner.package.check_execution()?;
        Ok((&view.inner.package, &view.inner))
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

    /// Build a deferred owner that restores `after`'s exact artifact over this
    /// owner. The source and output are authenticated by the OPC restore
    /// primitive or by the deferred publisher, and the recipient owner's
    /// execution context and source limits perform the replay.
    pub(crate) fn restored(
        &self,
        before: &RestoreEndpoint,
        after: &RestoreEndpoint,
    ) -> Result<Self> {
        if before.same_owner(self) {
            if let Some(owner) = after.owner() {
                return Ok(owner);
            }
            // The patch may outlive the committed owner. Rebuild its deferred
            // endpoint over this recipient's source without retaining the
            // producer's facade graph. A pending owner is materialized only
            // when its own current output is needed as the recipient base.
            if let Some(publisher) = after.publisher() {
                let expected_before = before.output_artifact()?;
                check_deferred_policy(&self.inner.package)?;
                let catalog = after.catalog_for(&self.inner.package)?;
                return self.deferred_endpoint(
                    Arc::new(RestorePublisher::Deferred {
                        publisher: Arc::clone(publisher),
                        expected_before,
                    }),
                    catalog,
                );
            }
        }

        // A pending owner has an indexed base package, not the bytes it will
        // publish. Materialize that owner once so the recipient package below
        // is always the exact artifact authenticated by the replay closure.
        // This is only reached by patch replay, after same_artifact has already
        // established that the recipient and the patch base are equal.
        let recipient = self.view()?.clone();
        let expected_before = before.output_artifact()?;
        let catalog = after.catalog_for(&recipient.inner.package)?;

        let publish: Arc<dyn DeferredPublisher> = if let Some(original) = after.artifact() {
            let original = original.clone();
            let current_artifact = recipient.inner.package.source_artifact();
            let proof = SourceArtifactRestoreProof {
                current_len: current_artifact.len(),
                current_sha256: current_artifact.fingerprint()?,
                original_len: original.len(),
                original_sha256: original.fingerprint()?,
            };
            let max_output_bytes = proof.original_len.max(1);
            Arc::new(RestorePublisher::SourceArtifact {
                original,
                proof,
                max_output_bytes,
            })
        } else {
            let publisher = Arc::clone(
                after
                    .publisher()
                    .expect("a non-opened owner has a deferred publication"),
            );
            Arc::new(RestorePublisher::Deferred {
                publisher,
                expected_before,
            })
        };

        // `recipient` is already the exact, indexed artifact authenticated
        // above. Reusing its owner keeps the recipient's execution context,
        // source limits, caches and admission proof intact while avoiding a
        // second catalog/relationship parse for every exact replay. The
        // deferred publisher still streams through that same recipient
        // package and checks its source freshness at write time.
        let mut next = recipient;
        next.pending = Some(Arc::new(Deferred {
            publish,
            materialized: OnceLock::new(),
            fingerprint: OnceLock::new(),
            catalog,
            owner: Arc::downgrade(&next.inner),
            endpoint: OnceLock::new(),
        }));
        Ok(next)
    }

    fn deferred_endpoint(
        &self,
        publish: Arc<dyn DeferredPublisher>,
        catalog: Option<Catalog>,
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
        Ok(Self {
            inner: Arc::clone(&base.inner),
            pending: Some(Arc::new(Deferred {
                publish,
                materialized: OnceLock::new(),
                fingerprint: OnceLock::new(),
                catalog,
                owner: Arc::downgrade(&base.inner),
                endpoint: OnceLock::new(),
            })),
        })
    }
}

fn check_deferred_policy(package: &SourceBackedPackage) -> Result<()> {
    if package.has_signature_infrastructure() {
        return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
    }
    if package.has_encrypted_entries() {
        return Err(litchi_opc::OpcError::SourceBackedOverlayUnavailable {
            reason: "source XML publication refuses encrypted ZIP members".into(),
        }
        .into());
    }
    Ok(())
}

/// A deferred output endpoint used by exact replay. Keeping this as an
/// `Arc`-backed object lets the restored owner retain one typed endpoint
/// without allocating another closure environment.
enum RestorePublisher {
    SourceArtifact {
        original: SourceArtifact,
        proof: SourceArtifactRestoreProof,
        max_output_bytes: u64,
    },
    Deferred {
        publisher: Arc<dyn DeferredPublisher>,
        expected_before: (u64, [u8; 32]),
    },
}

impl DeferredPublisher for RestorePublisher {
    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        match self {
            Self::SourceArtifact {
                original,
                proof,
                max_output_bytes,
            } => Ok(package.restore_source_artifact_to_stream(
                original,
                *proof,
                *max_output_bytes,
                writer,
            )?),
            Self::Deferred {
                publisher,
                expected_before,
            } => {
                authenticate_artifact(package, *expected_before, "current")?;
                // The exact before endpoint has already matched the
                // recipient. A deferred Change is immutable and its splice
                // is deterministic over those exact source bytes, so one
                // recipient-policy publication is the authenticated after
                // artifact. Hashing the same complete archive once before
                // this write would duplicate the bounded splice and defeat
                // source-backed streaming.
                publisher.write_authenticated(package, writer)
            },
        }
    }

    fn catalog(&self, package: &SourceBackedPackage) -> Result<Option<Catalog>> {
        match self {
            Self::SourceArtifact { .. } => Ok(None),
            Self::Deferred { publisher, .. } => publisher.catalog(package),
        }
    }
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
    /// Write after exact artifact authentication has established this
    /// change's complete before state. Families may omit their redundant
    /// semantic selector/dependency check.
    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        self.write(package, writer)
    }
    /// Return the already-parsed workbook catalog when the change owns one and
    /// the recipient retains the same physical read policy.
    ///
    /// Workbook-level families parse their target catalog while authoring the
    /// change. Reusing that result avoids a second XML parse during commit
    /// without changing the infallible read-only projections.
    fn catalog(&self, _package: &SourceBackedPackage) -> Option<Catalog> {
        None
    }
    /// The after workbook XML of a workbook-level change, whose catalog the
    /// committed workbook's infallible accessors read.
    fn catalog_xml(&self) -> Option<&[u8]> {
        None
    }
}

impl<C: Change> DeferredPublisher for C {
    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        Change::write(self, package, writer)
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        Change::write_authenticated(self, package, writer)
    }

    fn catalog(&self, package: &SourceBackedPackage) -> Result<Option<Catalog>> {
        if let Some(catalog) = Change::catalog(self, package) {
            return Ok(Some(catalog));
        }
        let Some(xml) = Change::catalog_xml(self) else {
            return Ok(None);
        };
        let parsed = crate::raw::parse_catalog(xml, &package.read_limits())?;
        Ok(Some(Catalog {
            active_sheet: (!parsed.sheets.is_empty()).then_some(parsed.active_sheet_index),
            defined_names: Arc::from(parsed.defined_names.into_boxed_slice()),
        }))
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
    let publisher: Arc<dyn DeferredPublisher> = change.clone();
    let base = workbook.view()?;
    let package = &base.inner.package;
    check_deferred_policy(package)?;
    if let Some(catalog) = change.catalog(package) {
        return Ok(SourceBackedWorkbook::deferred_from_base(
            base,
            Some(catalog),
            publisher,
        ));
    }
    let catalog = change.catalog_xml().map(<[u8]>::to_vec);
    // The committed workbook shares the patch's record; nothing is copied.
    // Coercing the existing Arc to the publisher trait object keeps the
    // deferred owner and patch on one allocation.
    let catalog = match catalog {
        None => None,
        Some(xml) => {
            let parsed = crate::raw::parse_catalog(&xml, &package.read_limits())?;
            Some(Catalog {
                active_sheet: (!parsed.sheets.is_empty()).then_some(parsed.active_sheet_index),
                defined_names: Arc::from(parsed.defined_names.into_boxed_slice()),
            })
        },
    };
    Ok(SourceBackedWorkbook::deferred_from_base(
        base, catalog, publisher,
    ))
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
                std::sync::Arc<crate::workbook::source::deferred::RestoreEndpoint>,
                std::sync::Arc<crate::workbook::source::deferred::RestoreEndpoint>,
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
                        .map(|(before, after)| (
                            std::sync::Arc::clone(after),
                            std::sync::Arc::clone(before),
                        )),
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
                Patch(
                    change,
                    Some((
                        crate::workbook::source::deferred::RestoreEndpoint::from_workbook(
                            workbook,
                        ),
                        crate::workbook::source::deferred::RestoreEndpoint::from_workbook(&next),
                    )),
                ),
                diagnostics,
            ))
        }

        fn apply_change(
            workbook: &crate::SourceBackedWorkbook,
            patch: &Patch,
        ) -> crate::Result<Commit> {
            let exact_base = patch.1.as_ref().and_then(|(before, after)| {
                if before.same_owner(workbook) {
                    return Some(Ok((before, after)));
                }
                // Fingerprint the producer endpoint first. If that provider
                // advanced after commit, its exact endpoint is unavailable;
                // the detached record below still binds replay to immutable
                // before bytes and its selector/dependency closure.
                let expected = match before.output_artifact() {
                    Ok(expected) => expected,
                    Err(crate::Error::Package(
                        litchi_opc::OpcError::SourceChanged { .. },
                    )) => return None,
                    Err(error) => return Some(Err(error)),
                };
                let current = match workbook.output_artifact() {
                    Ok(current) => current,
                    // A recipient source failure is its own error and must
                    // never be hidden by the producer fallback.
                    Err(error) => return Some(Err(error)),
                };
                if current == expected {
                    Some(Ok((before, after)))
                } else {
                    None
                }
            });
            let next = match exact_base {
                Some(Ok((before, after))) => workbook.restored(before, after)?,
                Some(Err(error)) => return Err(error),
                None => crate::workbook::source::deferred::replay(workbook, &patch.0)?,
            };
            let diagnostics =
                crate::workbook::source::deferred::Report::diagnostics(&*patch.0);
            Ok(Commit::new(next, patch.clone(), diagnostics))
        }
    };
}

pub(crate) use source_family;
