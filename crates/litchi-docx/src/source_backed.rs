//! Lazy DOCX access and guarded main-document publication over an immutable
//! positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes. Main-document
//! transactions retain the raw XML and may be published to a sequential sink
//! while raw-copying every unselected ZIP member.

mod deferred;
pub(crate) mod document_policy;
mod ingress;
mod ink;
mod owner;
#[path = "source_backed/paragraph_copy_record.rs"]
#[allow(unreachable_pub)]
mod paragraph_copy_record;
mod svg_lifecycle;

#[cfg(test)]
mod document_memo_tests;
#[cfg(test)]
mod paragraph_index_tests;

pub mod alt_chunk_payload;
pub mod document_variables;
pub mod glossary_text;
pub mod hyperlink_detachment;
pub mod hyperlink_redaction;
pub mod paragraph_copy;
pub mod paragraph_stream;
pub mod plain_paragraph_copy;
pub mod plain_paragraph_removal;
pub mod section_layout;
pub mod story_hyperlink_redaction;
pub mod story_text;
pub mod story_text_edit;
pub mod svg;
pub mod tail_append;
pub mod tail_append_stream;

pub use litchi_opc::{SourceReadDiagnostics, SourceReadPolicy, SourceReadPolicyError};

use crate::document::{Result as TransactionResult, Snapshot};
use crate::error::{Error, Result};
use crate::namespace::scan_word_element_ranges;
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{
    DocumentIndexAdmission, ParagraphIndex, ParagraphIndexBuild, ParagraphIndexCache,
    ParagraphIndexMemo, ParagraphIndexRoute, body_block_ranges, document_blocks, document_elements,
    document_paragraph, document_paragraph_count, document_paragraph_from_index,
    document_paragraphs, document_paragraphs_from_index, document_tables, is_xml_outer_whitespace,
    visible_document_capabilities,
};
use crate::parts::memo::{RawPositions, VisiblePayload};
use crate::redact;
use crate::sanitize::{self, RelationshipState};
use crate::section::layout;
use crate::settings::DocumentSettings;
use crate::variables;
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{ExecutionContext, ExecutionError, ReadAt, Reservation, Resource, SourceVersion};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    ArtifactStreamPolicy, ArtifactStreamReport, ArtifactStreamTarget, BlobPart, PackURI, Part,
    PartData, PartView, SourceArtifactFingerprint, SourceBackedPackage, SourceCacheDiagnostics,
    SourceTopologyPlan,
};
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest as _, Sha256};
use smallvec::SmallVec;
use std::borrow::Cow;
use std::io::{Read, Write};
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};

/// A DOCX package that leaves ordinary part bodies cold at open.
///
/// Clones share the source lineage, index, caches and execution reservations.
/// Publishing borrows this snapshot: callers can publish again or keep reading
/// the original package while any share remains alive.
#[derive(Clone)]
pub struct Package {
    physical: SourceBackedPackage,
    // Installed only through with_prepared_topology, which creates fresh caches.
    // Never replace this while retaining the prior owner's paragraph/admission caches.
    effective: Option<litchi_opc::PreparedTopology>,
    execution: Option<ExecutionContext>,
    /// The semantic read budgets the package was opened with.
    semantic_limits: Arc<crate::SemanticLimits>,
    paragraph_index_cache: Arc<ParagraphIndexCache>,
    /// Everything repaired while the package was read, aggregated from the
    /// codecs that record. See `crate::leniency`. Interior mutability keeps
    /// the lazy read facades on `&self`.
    tolerance: Arc<Mutex<crate::leniency::ToleranceReport>>,
    /// Whether [`Package::metadata`] has folded the core-properties repairs
    /// into `tolerance`. The source is pinned to one revision, so every read
    /// repairs the same defects; they are recorded once, not per call.
    properties_recorded: Arc<std::sync::atomic::AtomicBool>,
    /// The source revision whose main document the markup-compatibility
    /// admission last proved for an unmanaged transaction. One revision has
    /// one main-document payload, so publishing an edit started from it does
    /// not prove the same bytes again.
    main_admission: Arc<Mutex<Option<SourceVersion>>>,
    /// The unmanaged main-document transaction snapshot last built from the
    /// physical source (record 0305). A snapshot is cheap to clone and
    /// carries its source identity (lineage, revision, part name); it is
    /// reused only when that identity matches and its bytes equal the
    /// current payload, so one revision is scanned once, not once per edit
    /// and again per commit.
    main_snapshot: Arc<Mutex<Option<Snapshot>>>,
    /// The pending output of a committed package (record 0355a); its
    /// physical source is the materialized output.
    deferred: Option<Arc<deferred::Deferred>>,
    /// A complete-artifact fingerprint memo for a generic source or an
    /// immutable effective topology. Owned physical sources use the memo
    /// already retained by `SourceArtifact` and keep this field empty.
    artifact_fingerprint: Option<ArtifactFingerprintMemo>,
}

#[derive(Clone)]
enum ArtifactFingerprintMemo {
    Source(Arc<Mutex<Option<(SourceVersion, [u8; 32])>>>),
    Effective(Arc<OnceLock<(u64, [u8; 32])>>),
}

impl ArtifactFingerprintMemo {
    fn source() -> Self {
        Self::Source(Arc::new(Mutex::new(None)))
    }

    fn effective() -> Self {
        Self::Effective(Arc::new(OnceLock::new()))
    }
}

/// A checked selector for one active `w:altChunk` anchor in the main story.
///
/// The selector is deliberately positional: an altChunk has no semantic name
/// of its own, and relationship IDs are package vocabulary rather than an
/// ordinary document selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AltChunkSelector {
    /// Zero-based active main-story altChunk position.
    Index(usize),
}

impl AltChunkSelector {
    /// Construct a checked-by-use positional selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }
}

impl From<usize> for AltChunkSelector {
    fn from(index: usize) -> Self {
        Self::Index(index)
    }
}

struct DocumentVariablesSource {
    partname: PackURI,
    snapshot: variables::Snapshot,
    protected: bool,
    unique_inbound_owner: bool,
}

struct FingerprintingWriter<W> {
    inner: W,
    hasher: Sha256,
}

struct SourceCheckedTextSink<'a, W: ?Sized> {
    output: &'a mut W,
    package: &'a Package,
    failure: Arc<Mutex<Option<Error>>>,
}

impl<'a, W: ?Sized> SourceCheckedTextSink<'a, W> {
    fn record_failure(&self, error: Error) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check(&self) -> std::io::Result<()> {
        let result = self.package.check_source();
        match result {
            Ok(()) => Ok(()),
            Err(error) => {
                let message = error.to_string();
                self.record_failure(error);
                Err(std::io::Error::other(message))
            },
        }
    }
}

impl<'a, W: Write + ?Sized> Write for SourceCheckedTextSink<'a, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.check()?;
        let result = self.output.write(bytes);
        let _ = self.check();
        result
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.check()?;
        let result = self.output.flush();
        let _ = self.check();
        result
    }
}

fn take_source_text_failure(failure: &Arc<Mutex<Option<Error>>>) -> Option<Error> {
    failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take()
}

impl<W: Write> Write for FingerprintingWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        if written > bytes.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "source-backed publication sink reported {written} bytes for a {}-byte write",
                    bytes.len()
                ),
            ));
        }
        self.hasher.update(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Package {
    fn materialize_document_edit(
        package: &Self,
        options: crate::document::EditOptions,
    ) -> TransactionResult<crate::document::TransactionEdit> {
        package
            .main_document_snapshot_with_before("edit", true, None, Some(&options))
            .map(|(_, snapshot)| snapshot.edit())
    }

    /// Central boundary for operations that require the original OPC graph.
    fn physical_package(&self) -> Result<&SourceBackedPackage> {
        self.physical_opc_package()
    }

    fn physical_opc_package(&self) -> Result<&SourceBackedPackage> {
        if let Some(effective) = &self.effective {
            effective.check_source()?;
            return Err(Error::Opc(
                litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                    reason: "this DOCX operation requires migration to the effective graph".into(),
                },
            ));
        }
        self.physical_source()
    }

    /// The physical source: the opened archive, or the materialized output
    /// of a committed package (record 0355a).
    fn physical_source(&self) -> Result<&SourceBackedPackage> {
        match &self.deferred {
            Some(deferred) => deferred.materialized(),
            None => Ok(&self.physical),
        }
    }

    /// Open an immutable filesystem source using the standard finite policies.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, crate::OpenOptions::default())
    }

    /// Open a filesystem source with selected source, cache, read-ahead and
    /// execution policies. Ordinary payloads stay cold until requested.
    /// Cancellation is checked before opening the file and during OPC admission.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: crate::OpenOptions) -> Result<Self> {
        if let Some(context) = &options.execution_context {
            context.check().map_err(|error| match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })?;
        }
        Self::from_read_at_with_options(file_source(path)?, options)
    }

    /// Open an immutable positional source using standard finite policies.
    /// The main relationship and content type are checked without loading its payload.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, crate::OpenOptions::default())
    }

    /// Open a positional source with caller-selected source, cache, read-ahead
    /// and execution policies. The context remains attached to deferred reads;
    /// managed payload handles retain their reservations through publication.
    /// No runtime or implicit scheduler is installed.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        let cache_limits = options.cache_limits.unwrap_or_default();
        let read_policy = options.source_read_policy.unwrap_or_default();
        let package = if let Some(context) = options.execution_context.clone() {
            SourceBackedPackage::from_read_at(
                source,
                options.source_limits,
                cache_limits,
                read_policy,
                Some(context),
            )?
        } else {
            SourceBackedPackage::from_read_at(
                source,
                options.source_limits,
                cache_limits,
                read_policy,
                None,
            )?
        };
        let limits = options.semantic_limits.clone();
        Ok(Self::with_execution(package, options.execution_context)?.with_semantic_limits(limits))
    }

    /// Open a sequential reader with the standard finite source policies.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, crate::OpenOptions::default())
    }

    /// Consume a reader once into a bounded immutable source.
    ///
    /// Managed ingress charges accepted input bytes and copy/retry work. The
    /// source retains its buffer-capacity and owner memory reservations; growth
    /// also admits the peak of old and replacement buffers plus an 8 KiB scratch
    /// buffer. Subsequent OPC reads charge input independently. An EOF probe
    /// needs one byte of available input budget, but commits zero at EOF.
    /// Cache, read-ahead and execution policies continue through deferred reads.
    pub fn from_reader_with_options<R: Read>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        if options.execution_context.is_none()
            && options.cache_limits.is_none()
            && options.source_read_policy.is_none()
        {
            // Preserve the existing move-owned OPC fast path and its memoized
            // immutable-source proofs when no lazy policy is selected.
            return Ok(Self::from_source_backed(SourceBackedPackage::from_reader(
                reader,
                options.source_limits,
                litchi_opc::SourceCacheLimits::default(),
                None,
            )?)?
            .with_semantic_limits(options.semantic_limits));
        }
        let source = ingress::from_reader(reader, &options)?;
        Self::from_read_at_with_options(source, options)
    }

    /// Open a byte slice with standard finite policies. The input is copied.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, crate::OpenOptions::default())
    }

    /// Open a byte slice with caller-selected policies, copying it into an
    /// immutable source under the same accounting as sequential ingress.
    pub fn from_bytes_with_options(bytes: &[u8], options: crate::OpenOptions) -> Result<Self> {
        Self::from_reader_with_options(bytes, options)
    }

    /// The semantic read budgets this package keeps (record 0355).
    #[must_use]
    pub fn semantic_limits(&self) -> &crate::SemanticLimits {
        &self.semantic_limits
    }

    /// The same package under other semantic read budgets. O(1): the source
    /// and every cache are shared.
    #[must_use]
    pub fn with_semantic_limits(&self, limits: crate::SemanticLimits) -> Self {
        let mut next = self.clone();
        next.semantic_limits = Arc::new(limits);
        next
    }

    fn from_source_backed(package: SourceBackedPackage) -> Result<Self> {
        Self::with_execution(package, None)
    }

    fn with_execution(
        package: SourceBackedPackage,
        execution: Option<ExecutionContext>,
    ) -> Result<Self> {
        package.check_execution()?;
        let main = package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        package.check_execution()?;
        let artifact_fingerprint = if package.source_artifact().has_fingerprint_memo() {
            None
        } else {
            Some(ArtifactFingerprintMemo::source())
        };
        // The report starts with the packaging repairs made at open; see
        // `crate::leniency`.
        let mut tolerance = crate::leniency::ToleranceReport::default();
        tolerance.append_package(package.tolerance_report());
        Ok(Self {
            physical: package,
            effective: None,
            execution,
            semantic_limits: Arc::default(),
            paragraph_index_cache: Arc::new(ParagraphIndexCache::new()),
            tolerance: Arc::new(Mutex::new(tolerance)),
            properties_recorded: Arc::new(std::sync::atomic::AtomicBool::new(false)),
            main_admission: Arc::new(Mutex::new(None)),
            main_snapshot: Arc::new(Mutex::new(None)),
            deferred: None,
            artifact_fingerprint,
        })
    }

    /// Adopt an already-indexed source-backed OPC package.
    ///
    /// The package is validated exactly once at this DOCX boundary: the
    /// caller-owned OPC catalog is retained without reparsing, while the
    /// unique main-document relationship and WordprocessingML content type
    /// receive the same checks as the regular source constructors. The
    /// package's execution context is carried into every deferred semantic
    /// query.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        let execution = package.execution_context();
        Self::with_execution(package, execution)
    }

    /// Borrow the original OPC graph when no effective candidate is installed.
    ///
    /// Returns a typed refusal for effective owners so callers cannot silently
    /// discard retained edits by operating on the physical archive.
    ///
    /// This is the package indexed when the document was opened: no ZIP
    /// directory, content-type catalog, or relationship graph is rebuilt.
    /// Borrowed package services, such as
    /// [`SourceBackedPackage::source_artifact`], read through the same
    /// positional source and observe the same source-version fence.
    pub fn source_backed_package(&self) -> Result<&SourceBackedPackage> {
        self.physical_package()
    }

    /// Return the OPC read limits retained by this owner.
    ///
    /// An effective owner still uses the physical source's admission policy;
    /// reading that policy must not ask the effective graph to masquerade as
    /// the physical package or perform another catalog read.
    #[must_use]
    pub const fn read_limits(&self) -> litchi_opc::SourceLimits {
        self.physical.read_limits()
    }

    /// Clone the execution context retained by this owner, including an
    /// owner with an effective edited topology.
    ///
    /// This accessor only returns the context captured at source admission; it
    /// never asks the effective graph to materialize or reopens the physical
    /// package. Callers performing source-backed I/O must still use
    /// [`Self::check_source`] around that operation.
    #[must_use]
    pub fn execution_context(&self) -> Option<ExecutionContext> {
        self.execution.clone()
    }

    /// Materialize this source-backed DOCX into the established owning
    /// package facade without reparsing the source catalog.
    ///
    /// This is the explicit compatibility seam for consumers whose semantic
    /// operation needs package-wide relationship access (for example, the
    /// Markdown adapter). Unmanaged sources copy the validated catalog and
    /// payloads into one owning package; managed sources are refused by the
    /// OPC layer before ordinary payload reads so their budget reservations
    /// cannot be detached accidentally.
    pub fn to_owned_package(&self) -> Result<crate::Package> {
        let package = match &self.effective {
            Some(effective) => effective.to_opc_package()?,
            None => self.physical_source()?.to_opc_package()?,
        };
        let result = crate::Package::from_opc_package(package);
        // The OPC conversion checks freshness through its final payload read,
        // but the owning DOCX constructor still performs graph/property
        // validation after that point. Re-check the original source before
        // exposing either success or a constructor error so a mutation in
        // this publication window cannot yield a stale owner (or mask the
        // stale-source failure with an eager parse error).
        self.check_source()?;
        result
    }

    /// Materialize this source-backed DOCX for an explicit Eager owner while
    /// retaining the exact source archive for later owning publication.
    ///
    /// In-memory and reader-owned sources share their existing bytes. A
    /// positional source is copied once under its OPC input limit; the copy
    /// retains opaque ZIP members and original framing so an owning edit can
    /// preserve untouched members. Source freshness is checked by the OPC
    /// conversion before and after the copy and materialization.
    pub fn to_owned_package_preserving_source(&self) -> Result<crate::Package> {
        let package = match &self.effective {
            Some(effective) => effective.to_opc_package_preserving_source()?,
            None => self.physical_source()?.to_opc_package_preserving_source()?,
        };
        let result = crate::Package::from_opc_package(package);
        self.check_source()?;
        result
    }

    /// Load and pin the main document for read-only semantic queries.
    ///
    /// The first call materializes only the main-document part. An explicit
    /// source-read policy may also prefetch adjacent compressed bytes. The
    /// returned document owns its validated visible XML bytes, so repeated text and selective
    /// queries do not revisit the positional source. Managed opens retain the
    /// budgeted [`PartData`] handle instead of detaching an `Arc`.
    ///
    /// On an unmanaged package the markup-compatibility projection of the
    /// main document is memoized for the package's source revision: the
    /// first call projects it, and later calls share that projection without
    /// reading the part again. Every call still checks the execution context
    /// and the source revision first, so a changed source is reported as
    /// before. [`Package::trim_document_memos`] releases the memo.
    pub fn document(&self) -> Result<Document> {
        if let Some(effective) = &self.effective {
            return self.effective_document(effective);
        }
        let result: Result<Document> = (|| {
            self.physical_package()?.check_execution()?;
            let main = self.physical_package()?.main_document_part()?;
            validate_document_main_content_type(main.content_type())?;
            let managed = self.physical_package()?.cache_diagnostics().budget_managed;
            if !managed {
                return self.visible_document(main);
            }
            let data = self.main_data(main)?;
            let source_version = self.physical_package()?.source_version()?;
            let document: Result<Document> = (|| {
                let (xml, paragraph_memo, index_admission, paragraph_route) = {
                    let context =
                        self.physical_package()?
                            .execution_context()
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "managed document context is unavailable".into(),
                                )
                            })?;

                    let build = self.paragraph_index_cache.get_or_build_source(
                        source_version,
                        ParagraphIndexRoute::Managed,
                        || {
                            // The cache lock has already removed any clean old
                            // entry. Admit and validate only on a miss, so cache
                            // hits do not repeat the workspace/work envelope or
                            // contend for budget before joining the memo.
                            ensure_source_document_xml(
                                data.as_bytes(),
                                self.physical_package()?.execution_context().as_ref(),
                            )?;
                            let index_admission =
                                DocumentIndexAdmission::new(&context, data.as_bytes().len())?;
                            let _parser =
                                admit_document_query_parser(Some(&context), data.as_bytes().len())?;
                            // Budget-managed opens keep the eager scan. Its parser
                            // admission is reserved and its work consumed here, so
                            // deferring only the scan would run it with fewer live
                            // reservations than the managed contract established.
                            let paragraph_index = ParagraphIndex::from_xml(data.as_bytes()).ok();
                            // Do not publish an index whose source or execution
                            // fence was crossed while the managed scan ran.
                            self.physical_package()?.source_version()?;
                            self.physical_package()?.check_execution()?;
                            Ok(match paragraph_index {
                                Some(index) => {
                                    ParagraphIndexBuild::indexed(index, Some(index_admission))
                                },
                                None => ParagraphIndexBuild::unavailable(Some(index_admission)),
                            })
                        },
                    )?;
                    (
                        DocumentPayload::Managed(data),
                        OnceLock::from(build.memo),
                        build.admission,
                        ParagraphIndexRoute::Managed,
                    )
                };
                Ok(Document {
                    xml,
                    _retained: None,
                    paragraph_memo,
                    _index_admission: index_admission,
                    source_version,
                    execution: self.execution.clone(),
                    limits: Arc::clone(&self.semantic_limits),
                    paragraph_index_cache: Arc::clone(&self.paragraph_index_cache),
                    paragraph_route,
                    // A managed view reads its source bytes, which only a
                    // lexical admission lets through.
                    positions: RawPositions::Preserved,
                    body_memo: OnceLock::new(),
                    allows_detached_views: false,
                    // A pinned view starts from the packaging repairs made
                    // when its package was opened.
                    tolerance: Arc::new(Mutex::new(self.opened_tolerance())),
                })
            })();
            // The semantic/MCE stage can fail after the payload read has
            // completed. Check freshness once more before exposing that error so
            // a source mutation during parsing wins over a stale parse result.
            let freshness = (|| {
                self.physical_package()?.source_version()?;
                self.physical_package()?.check_execution()?;
                Ok::<_, Error>(())
            })();
            match freshness {
                Ok(()) => {
                    if document.is_err() {
                        self.paragraph_index_cache.trim_clean();
                    }
                    document
                },
                Err(error) => {
                    drop(document);
                    self.paragraph_index_cache.trim_clean();
                    Err(error)
                },
            }
        })();
        if result.is_err() {
            self.paragraph_index_cache.trim_clean();
        }
        result
    }

    /// [`Self::document`] on an unmanaged package: the visible projection is
    /// memoized per source revision, and the part is read only to build it
    /// or when the projection is the part's own allocation.
    fn visible_document(&self, main: PartView<'_>) -> Result<Document> {
        let source_version = self.physical_package()?.source_version()?;
        let document: Result<Document> = (|| {
            let mut loaded = None;
            let payload =
                self.paragraph_index_cache
                    .visible_payload_source(source_version, || {
                        let raw = self.main_data(main)?.into_arc()?;
                        let payload = VisiblePayload::project(raw.as_slice())?;
                        loaded = Some(raw);
                        Ok(payload)
                    })?;
            // A hit reads no payload: re-deriving it here, even in debug
            // builds, would change the observable source-cache counters, so
            // the memo's value identity is proven by the suite instead.
            let xml = match (payload.projected(), loaded) {
                (Some(projected), _) => Arc::clone(projected),
                (None, Some(raw)) => raw,
                (None, None) => self.main_data(main)?.into_arc()?,
            };
            Ok(Document {
                xml: DocumentPayload::Owned(xml),
                _retained: None,
                // Unmanaged opens reserve nothing for the index and swallow
                // its failure, so the scan is deferred to the first paragraph
                // query that can use it.
                paragraph_memo: OnceLock::new(),
                _index_admission: None,
                source_version,
                execution: self.execution.clone(),
                limits: Arc::clone(&self.semantic_limits),
                paragraph_index_cache: Arc::clone(&self.paragraph_index_cache),
                paragraph_route: ParagraphIndexRoute::Visible,
                positions: payload.positions(),
                body_memo: OnceLock::new(),
                allows_detached_views: true,
                // A pinned view starts from the packaging repairs made when its
                // package was opened.
                tolerance: Arc::new(Mutex::new(self.opened_tolerance())),
            })
        })();
        // The semantic/MCE stage can fail after the payload read has
        // completed. Check freshness once more before exposing that error so a
        // source mutation during parsing wins over a stale parse result.
        self.physical_package()?.source_version()?;
        self.physical_package()?.check_execution()?;
        document
    }

    /// Release the package's clean document memos: the visible main-document
    /// projection, the paragraph offsets and the body block index that no
    /// live document view or value still holds.
    ///
    /// Views keep what they hold; the next [`Self::document`] rebuilds what
    /// was released. Results never depend on what the memos hold.
    pub fn trim_document_memos(&self) {
        self.paragraph_index_cache.trim_clean_all();
    }

    fn ensure_source_document_xml(
        &self,
        xml: &[u8],
        context: Option<&ExecutionContext>,
    ) -> Result<()> {
        match ensure_source_document_xml(xml, context) {
            Ok(()) => Ok(()),
            Err(error) if is_index_pressure_error(&error) => {
                // A clean memo may retain a managed index reservation after
                // its last document view was dropped. Release that
                // cache-only reference and retry only the failed admission.
                // Active documents keep their memo and reservation pinned.
                self.paragraph_index_cache.trim_clean();
                ensure_source_document_xml(xml, context)
            },
            Err(error) => Err(error),
        }
    }

    fn trim_clean_paragraph_index_for_operation(&self) {
        self.paragraph_index_cache.trim_clean();
    }

    fn main_data(&self, main: PartView<'_>) -> Result<PartData> {
        match main.data().map_err(Error::from) {
            Ok(data) => Ok(data),
            Err(error) if is_index_pressure_error(&error) => {
                self.paragraph_index_cache.trim_clean();
                main.data().map_err(Error::from)
            },
            Err(error) => Err(error),
        }
    }

    fn preflight_external_hyperlink_document_xml(&self, limits: sanitize::Limits) -> Result<()> {
        let main = self.physical_package()?.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let declared = main.declared_uncompressed_size()?;
        let maximum = u64::try_from(limits.max_document_bytes()).unwrap_or(u64::MAX);
        if declared > maximum {
            let actual = usize::try_from(declared).unwrap_or(usize::MAX);
            return Err(Error::ExternalHyperlinkDetachmentLimit {
                kind: sanitize::LimitKind::MainDocumentXmlBytes,
                maximum: limits.max_document_bytes(),
                actual,
            });
        }
        Ok(())
    }

    /// Stream visible main-document paragraphs to a caller-owned sink.
    ///
    /// Resolves the current main part from the effective graph after a commit,
    /// including a redirected main part. Retained donor payloads and execution
    /// contexts stay pinned through output and the final source check.
    ///
    /// The source catalog and execution policy are checked before deferred
    /// payload work. The parser checks them before and after every XML event
    /// and immediately before each paragraph emission; the source-checked
    /// sink checks both fences around every underlying write and after the
    /// final object. The main-document declared size is checked before
    /// payload materialization; the payload reader still verifies its actual
    /// decoded length and source boundary. Only one bounded paragraph text
    /// value is retained by the semantic parser.
    ///
    /// Foreign text subtrees (OMML `m:t`, DrawingML `a:t`) are skipped as
    /// opaque and recorded as
    /// [`crate::leniency::TextDefect::SkippedForeignMarkup`], merged into the
    /// package's report, enumerable through [`Package::tolerance_report`].
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: litchi_core::TextOutputOptions<'_>,
    ) -> std::result::Result<litchi_core::TextOutputReport, litchi_core::TextOutputError<Error>>
    {
        let source_failure = Arc::new(Mutex::new(None));
        let mut checked_output = SourceCheckedTextSink {
            output,
            package: self,
            failure: Arc::clone(&source_failure),
        };
        let mut writer = litchi_core::SequentialTextWriter::new(&mut checked_output, options);
        let mut repaired = crate::leniency::ToleranceReport::default();
        // Keep payload reservations and donor provenance through the final fence.
        let mut retained_data = None;

        let parsed = (|| -> std::result::Result<(), litchi_core::TextOutputError<Error>> {
            self.check_source()
                .map_err(|source| writer.document_error(source))?;
            let data = retained_data.insert(
                self.main_text_data()
                    .map_err(|source| writer.document_error(source))?,
            );
            let managed = self.execution.is_some();
            let mce_limit = crate::paragraph::semantic_text_raw_xml_limit();
            let visible: Cow<'_, [u8]> = if managed {
                self.ensure_source_document_xml(data.as_bytes(), self.execution.as_ref())
                    .map_err(|source| writer.document_error(source))?;
                Cow::Borrowed(data.as_bytes())
            } else {
                let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
                capabilities
                    .understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
                capabilities.understand_namespace(crate::revision::WORD_2023_DATE_UTC_NAMESPACE);
                litchi_ooxml_common::mce::process_markup_compatibility(
                    data.as_bytes(),
                    &capabilities,
                    &litchi_ooxml_common::mce::Limits {
                        max_input_bytes: mce_limit,
                        max_output_bytes: mce_limit,
                        ..Default::default()
                    },
                )
                .map_err(|source| writer.document_error(source.into()))?
                .xml
            };
            self.check_source()
                .map_err(|source| writer.document_error(source))?;
            crate::paragraph::write_text_to_with_operation_check(
                visible.as_ref(),
                &mut writer,
                || self.check_source(),
                &mut repaired,
            )
        })();
        let progress = writer.progress();
        let source = self
            .check_source()
            .err()
            .or_else(|| take_source_text_failure(&source_failure));
        // Every repair the pass applied is merged, output error or not: a
        // failed call still skipped the subtrees it recorded.
        if !repaired.is_clean() {
            self.tolerance
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .append(&repaired);
        }
        if let Some(source) = source {
            return Err(litchi_core::TextOutputError::Document { source, progress });
        }
        parsed.map(|()| writer.finish())
    }

    /// The packaging repairs made when this package was opened, folded into
    /// a document-level report.
    fn opened_tolerance(&self) -> crate::leniency::ToleranceReport {
        let mut report = crate::leniency::ToleranceReport::default();
        report.append_package(self.physical.tolerance_report());
        report
    }

    /// A snapshot of everything repaired while this package was read, in
    /// discovery order.
    ///
    /// The report is the exhaustive record of the non-structural
    /// substitutions described by [`crate::leniency`]: nothing is repaired
    /// silently. It starts with the OPC packaging repairs made at open
    /// ([`Defect::Package`](crate::Defect::Package)); [`Self::metadata`] adds
    /// the core-properties repairs ([`Defect::Properties`](crate::Defect::Properties))
    /// the first time it reads them successfully, and the semantic text sink (`write_text_to`)
    /// records foreign `*:t` skips.
    #[must_use]
    pub fn tolerance_report(&self) -> crate::leniency::ToleranceReport {
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .clone()
    }

    /// Load only the mandatory main-document payload and capture its immutable
    /// section inventory at the exact opened source version.
    ///
    /// Header/footer relationship IDs are retained as inert values. Their
    /// target parts are neither resolved nor read.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_bounded(self.semantic_limits.sections_value())
    }

    /// Capture a source-bound section inventory with explicit semantic limits.
    pub(crate) fn section_inventory_snapshot_bounded(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Snapshot> {
        self.trim_clean_paragraph_index_for_operation();
        self.physical_package()?.check_execution()?;
        let main = self.physical_package()?.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let source_version = self.physical_package()?.source_version()?;
        let raw = self.main_data(main)?;
        let snapshot = (|| {
            ensure_source_section_inventory_xml(raw.as_bytes(), limits)?;
            crate::section::Snapshot::from_source_xml(raw.as_bytes(), source_version, limits)
        })();
        // Detect hostile adapters that mutate immediately after the payload
        // read and semantic scan, before publishing the closure to the caller.
        let observed = self.physical_package()?.source_version()?;
        if observed != source_version {
            return Err(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            }
            .into());
        }
        self.physical_package()?.check_execution()?;
        if let Ok(snapshot) = &snapshot {
            let repaired = snapshot.inventory().tolerance_report();
            if !repaired.is_clean() {
                self.tolerance
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .append(repaired);
            }
        }
        snapshot
    }

    /// Capture an exact-source snapshot for editing locally authored page
    /// layout on one existing main-story section.
    pub fn section_layout_snapshot(&self) -> Result<layout::Snapshot> {
        self.section_layout_snapshot_bounded(self.semantic_limits.sections_value())
    }

    /// Capture a bounded exact-source page-layout snapshot.
    pub(crate) fn section_layout_snapshot_bounded(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<layout::Snapshot> {
        let (_, snapshot) =
            self.main_section_layout_snapshot("section_layout_snapshot", limits, None)?;
        Ok(snapshot)
    }

    /// Return content-free cache activity for this source-backed DOCX.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.physical.cache_diagnostics()
    }

    /// Inspect the opt-in source-read window; exact constructors return `None`.
    pub fn source_read_diagnostics(&self) -> Result<Option<SourceReadDiagnostics>> {
        Ok(self.physical_package()?.source_read_diagnostics()?)
    }

    /// Return the physical source version after checking all effective guards.
    /// Logical candidates can share this version; it is not a semantic cache key.
    pub fn source_version(&self) -> Result<SourceVersion> {
        match &self.effective {
            Some(topology) => Ok(topology.physical_source_version()?),
            None => Ok(self.physical_source()?.source_version()?),
        }
    }

    /// The comments of the main document's comments part, in part order;
    /// empty when the document has none.
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the comments part is
    /// unreadable or malformed.
    pub fn comments(&self) -> Result<Vec<crate::comment::Comment>> {
        let target = self.with_main_relationships(|relationships| {
            relationships
                .iter()
                .find(|relationship| is_comments_relationship(relationship))
                .map(litchi_opc::rel::Relationship::target_partname)
                .transpose()
                .map_err(Error::from)
        })?;
        let Some(target) = target else {
            return Ok(Vec::new());
        };
        let payload = self
            .effective_part_xml(&target)?
            .ok_or_else(|| Error::PartNotFound(target.as_str().to_owned()))?;
        crate::comment::Comment::extract_from_payload(&payload)
    }

    /// The entries of the `commentsExtended` part, as the eager
    /// `Document::comments_extended` reads them (record 0345).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the part is malformed.
    pub fn comments_extended(&self) -> Result<Vec<crate::modern_comments::Extended>> {
        match self.related_xml("/commentsExtended")? {
            Some(xml) => {
                let limits = crate::mail_merge::mce_limits_for_source(&self.read_limits());
                crate::modern_comments::parse_comments_extended(&xml, &limits)
            },
            None => Ok(Vec::new()),
        }
    }

    /// Whether the main document relates a comments part; retained edits
    /// are included (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the relationships cannot
    /// be read.
    pub fn has_comments(&self) -> Result<bool> {
        self.with_main_relationships(|relationships| {
            Ok(relationships.iter().any(is_comments_relationship))
        })
    }

    /// The styles part, located, or `None` when the document has none; only
    /// the main part's relationships and the styles part are read (record
    /// 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the part is malformed.
    pub fn style_catalog(&self) -> Result<Option<crate::styles::catalog::Catalog>> {
        match self.related_xml("/styles")? {
            Some(xml) => Ok(Some(crate::styles::catalog::Catalog::from_xml(xml)?)),
            None => Ok(None),
        }
    }

    /// The numbering part, located, or `None` when the document has none;
    /// only the main part's relationships and the numbering part are read
    /// (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the part is malformed.
    pub fn numbering_catalog(&self) -> Result<Option<crate::numbering::catalog::Catalog>> {
        match self.related_xml("/numbering")? {
            Some(xml) => Ok(Some(crate::numbering::catalog::Catalog::from_xml(xml)?)),
            None => Ok(None),
        }
    }

    /// The header and footer parts the main document relates, in
    /// relationship order; only the main part's relationships are read
    /// (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the relationships cannot
    /// be read.
    pub fn header_footer_entries(&self) -> Result<Vec<crate::header_footer::catalog::Entry>> {
        self.with_main_relationships(crate::header_footer::catalog::entries_of)
    }

    /// The header and footer references of every section, in document
    /// order, each with its role; only the main part is read (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the main part is not
    /// well-formed.
    pub fn header_footer_references(
        &self,
    ) -> Result<Vec<(crate::header_footer::Role, crate::section::Reference)>> {
        let payload = match &self.effective {
            Some(topology) => Arc::new(topology.main_document_part()?.data()?.as_bytes().to_vec()),
            None => {
                let package = self.source_backed_package()?;
                let name = package.main_document_part()?.partname().clone();
                package.part(&name)?.data()?.into_arc()?
            },
        };
        crate::header_footer::catalog::references_of(&payload)
    }

    /// The payload of part `name`, or `None` when the package has no such
    /// part (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the name is invalid.
    pub fn part_xml(&self, name: &str) -> Result<Option<Arc<Vec<u8>>>> {
        let uri = PackURI::new(name.to_owned()).map_err(Error::Invalid)?;
        self.effective_part_xml(&uri)
    }

    /// The footnotes of the main document, in part order, without
    /// separators; only the main part's relationships and the footnotes part
    /// are read (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the part is malformed.
    pub fn footnotes(&self) -> Result<Vec<crate::footnote::Note>> {
        match self.related_xml("/footnotes")? {
            Some(xml) => crate::footnote::Note::footnotes_from_xml(xml),
            None => Ok(Vec::new()),
        }
    }

    /// The endnotes of the main document, in part order, without
    /// separators (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the part is malformed.
    pub fn endnotes(&self) -> Result<Vec<crate::footnote::Note>> {
        match self.related_xml("/endnotes")? {
            Some(xml) => crate::footnote::Note::endnotes_from_xml(xml),
            None => Ok(Vec::new()),
        }
    }

    /// The payload of the first internal part the main part relates by a
    /// relationship type ending in `suffix`.
    fn related_xml(&self, suffix: &str) -> Result<Option<Arc<Vec<u8>>>> {
        let target = self.with_main_relationships(|relationships| {
            relationships
                .iter()
                .find(|relationship| {
                    !relationship.is_external() && relationship.reltype().ends_with(suffix)
                })
                .map(litchi_opc::rel::Relationship::target_partname)
                .transpose()
                .map_err(Error::from)
        })?;
        match target {
            Some(target) => self.effective_part_xml(&target),
            None => Ok(None),
        }
    }

    /// Run `read` over the main part's relationships as this owner's
    /// effective graph has them: retained edits included (record 0319).
    fn with_main_relationships<T>(
        &self,
        read: impl FnOnce(&litchi_opc::Relationships) -> Result<T>,
    ) -> Result<T> {
        match &self.effective {
            Some(topology) => read(topology.main_document_part()?.rels()),
            None => {
                let package = self.source_backed_package()?;
                read(package.main_document_part()?.rels())
            },
        }
    }

    /// The payload of part `uri` in this owner's effective graph, or `None`
    /// when it has no such part (record 0319).
    fn effective_part_xml(&self, uri: &PackURI) -> Result<Option<Arc<Vec<u8>>>> {
        match &self.effective {
            Some(topology) => match topology.part(uri) {
                Ok(part) => Ok(Some(Arc::new(part.data()?.as_bytes().to_vec()))),
                Err(_) => Ok(None),
            },
            None => match self.source_backed_package()?.part(uri) {
                Ok(part) => Ok(Some(part.data()?.into_arc()?)),
                Err(_) => Ok(None),
            },
        }
    }

    /// The header, footer, footnotes, endnotes and comments parts the main
    /// document relates, in relationship order, each once; only the main
    /// part's relationships are read (record 0285).
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed or the main part's
    /// relationships cannot be read.
    pub fn story_parts(&self) -> Result<Vec<crate::document::substitute::StoryPart>> {
        let package = self.source_backed_package()?;
        let main = package.main_document_part()?;
        crate::document::story_parts_of(main.rels())
    }

    /// The chart part (a `c:chartSpace` or `cx:chartSpace`) the main
    /// document relates through `relationship_id` — an
    /// [`InlineChart`](crate::chart::InlineChart)'s — or `None` when the main
    /// document has no chart relationship of that ID. Only the main part's
    /// relationships and the chart part are read.
    ///
    /// # Errors
    ///
    /// Returns an error when the source changed, for an external chart
    /// relationship, or for a target that is not a chart part.
    pub fn chart_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        let package = self.source_backed_package()?;
        let main = package.main_document_part()?;
        let Some(relationship) = main.rels().get(relationship_id) else {
            return Ok(None);
        };
        let Some(content_type) = crate::chart::inline::chart_content_type(relationship.reltype())
        else {
            return Ok(None);
        };
        if relationship.is_external() {
            return Err(Error::InvalidFormat(
                "a chart relationship cannot be external".into(),
            ));
        }
        let part = package.part(&relationship.target_partname()?)?;
        if part.content_type() != content_type {
            return Err(Error::InvalidFormat(
                "a chart relationship targets a part that is not a chart".into(),
            ));
        }
        Ok(Some(part.data()?.as_bytes().to_vec()))
    }

    /// Read relationship-selected OOXML core properties without loading the
    /// main document or unrelated package members. Retained owners resolve the
    /// effective graph, including edited relationships and property payloads.
    ///
    /// A missing core-properties relationship yields empty metadata. Present
    /// properties retain the shared graph validation and dialect repair rules
    /// used by eager OOXML readers. Source freshness is checked before and after
    /// the selected payload read by the common source-backed property reader;
    /// execution cancellation is checked at the same boundaries.
    pub fn metadata(&self) -> Result<litchi_core::Metadata> {
        self.trim_clean_paragraph_index_for_operation();
        self.check_source()?;
        let mut repaired = litchi_ooxml_common::leniency::ToleranceReport::default();
        let properties = match &self.effective {
            Some(topology) => {
                litchi_ooxml_common::properties::read_prepared_with_report(topology, &mut repaired)
            },
            None => litchi_ooxml_common::properties::read_source_backed_with_report(
                self.physical_source()?,
                &mut repaired,
            ),
        };
        self.check_source()?;
        let properties = properties?;
        let mut tolerance = self
            .tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if !self
            .properties_recorded
            .swap(true, std::sync::atomic::Ordering::AcqRel)
        {
            tolerance.append_properties(&repaired);
        }
        Ok(properties.unwrap_or_default())
    }

    /// Inventory embedded-object and embedded-package relationships without
    /// materializing the OPC package or any payload bytes.
    ///
    /// The returned catalog entries borrow this source-backed package. Their
    /// payload views remain deferred; reading a payload is an explicit,
    /// fallible operation on the returned entry rather than part of this
    /// inventory call.
    pub fn embedded(&self) -> Result<Vec<litchi_ooxml_common::embedded::SourceEntry<'_>>> {
        self.trim_clean_paragraph_index_for_operation();
        Ok(litchi_ooxml_common::embedded::scan_source(
            self.physical_package()?,
            &self.semantic_limits.embedded_value(),
        )?)
    }

    /// Load the exact raw main-document bytes as the source for the canonical
    /// owner edit. Standalone XML changes use [`crate::document::raw::Tree`]
    /// with a required [`crate::document::raw::Limits`] policy.
    ///
    /// The owner edits and publishes the retained source bytes while readers
    /// see the markup-compatibility projection, so a document is admitted only
    /// when the two agree on everything either reads (one bounded proof, no
    /// projection built): it never mentions the markup-compatibility
    /// namespace, binds it and uses nothing in it, or carries one valid
    /// `mc:Ignorable` on `w:document` whose dropped markup holds no element
    /// or attribute of an understood namespace (and, below it, only
    /// `Ignorable` directives listing understood namespaces). Names are
    /// decided on their decoded values. A document whose only selections are
    /// `mc:AlternateContent` wrappers inside runs (a drawing or text box with
    /// a fallback) is admitted once its projection succeeds; each wrapper is
    /// one opaque inline object to the snapshot's edits (record 0307). A
    /// document whose projection otherwise selects, unwraps or hides content
    /// (`AlternateContent` outside a run, `ProcessContent`,
    /// `MustUnderstand`, ignorable markup holding content, or an `Ignorable`
    /// below `w:document` listing a namespace that is not understood) is
    /// refused, and a projection error is returned as is. Under an
    /// execution context the projection is never built, so only documents
    /// whose projection differs from the source in spelling alone are
    /// admitted. Keeping the raw bytes prevents semantic selectors from being
    /// applied to a different XML representation than the one eventually
    /// published.
    ///
    /// # Errors
    ///
    /// Returns a typed package, document, resource-limit, markup-
    /// compatibility, or unsafe-edit error.
    pub fn document_snapshot(&self) -> TransactionResult<Snapshot> {
        let (_, snapshot) = self.main_document_snapshot("document_snapshot", true)?;
        Ok(snapshot)
    }

    /// Start an isolated semantic edit retaining this complete source owner.
    ///
    /// Commit returns a new immutable package without serializing a ZIP.
    /// For standalone XML editing, use [`crate::document::raw::Tree`] with a
    /// required [`crate::document::raw::Limits`] policy; package publication
    /// continues through this owner edit.
    ///
    /// # Errors
    ///
    /// The same failures as [`Self::document_snapshot`] are deferred until
    /// the first operation, projection or commit.
    pub fn edit(&self) -> TransactionResult<crate::document::Edit<Self>> {
        self.edit_with_options(crate::document::EditOptions::default())
    }

    /// Start an isolated semantic edit under one explicit document-edit
    /// profile while retaining this complete source owner.
    pub fn edit_with_options(
        &self,
        options: crate::document::EditOptions,
    ) -> TransactionResult<crate::document::Edit<Self>> {
        Ok(crate::document::Edit::deferred(
            self.clone(),
            options,
            Self::materialize_document_edit,
        ))
    }

    /// Publish one joined document composition through this source owner. The
    /// source identity and current source version are checked before staging,
    /// then all accepted operations share one validation and deferred package
    /// publication pass.
    pub fn apply_composition(
        &self,
        composition: crate::document::composition::Composition,
    ) -> TransactionResult<crate::document::Commit<Self>> {
        let (source, operations, artifact_fingerprint) = composition.into_parts();
        let (_, current) = self.main_document_snapshot("apply_composition", true)?;
        if current.xml_bytes() != source.xml_bytes() {
            return Err(crate::document::Error::PatchConflict);
        }
        if let Some(expected) = artifact_fingerprint
            && self.output_fingerprint()? != expected
        {
            return Err(crate::document::Error::PatchConflict);
        }
        let mut edit = crate::document::Edit::new(self.clone(), current.edit());
        edit.apply_composition_operations(&operations)?;
        edit.commit()
    }

    /// Capture document variables from the existing internal settings Part.
    ///
    /// This source-backed capability never creates a settings relationship or
    /// Part. It accepts exactly one Strict or Transitional internal settings
    /// relationship with the Word settings content type. The settings part is
    /// admitted by the main document's markup-compatibility rule, with the
    /// baseline profile settings readers use: a valid `mc:Ignorable` on
    /// `w:settings` may drop Word's foreign leaves (such as `w14:docId`),
    /// which hold no WordprocessingML. Branch selection and every other
    /// directive are refused, so the semantic snapshot and publish bytes
    /// cannot diverge.
    ///
    /// # Errors
    ///
    /// Returns a typed source, relationship, content-type, MCE, XML, or limit
    /// error.
    pub fn document_variables_snapshot(&self) -> Result<variables::Snapshot> {
        Ok(self
            .settings_document_variables_source("document_variables_snapshot")?
            .snapshot)
    }

    /// Check a document-variable patch against the current settings part
    /// and return the part and its rewritten XML (record 0355a). Exact
    /// no-ops skip the rewrite checks.
    pub(crate) fn document_variables_target(
        &self,
        patch: &variables::Patch,
        operation: &'static str,
    ) -> Result<(PackURI, Arc<Vec<u8>>)> {
        self.trim_clean_paragraph_index_for_operation();
        let current = self.settings_document_variables_source(operation)?;
        let target = patch.apply_content(&current.snapshot)?;
        if !patch.is_empty() {
            if !current.unique_inbound_owner {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    operation,
                    "document-variable markup cannot be preserved for this settings owner",
                ));
            }
            if current.protected {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::Protected,
                    operation,
                    "document or write protection is enforced",
                ));
            }
            variables::ensure_source_backed_rewrite_safe(
                current.snapshot.xml_bytes(),
                current.snapshot.variables(),
                &current.snapshot.limits(),
            )?;
        }
        Ok((current.partname, target.shared_xml()))
    }

    /// Write this package with one part's payload replaced (record 0355a):
    /// the output of a committed single-part family change.
    pub(crate) fn write_part_overlay(
        &self,
        writer: &mut dyn Write,
        partname: &PackURI,
        payload: &Arc<Vec<u8>>,
        policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> Result<()> {
        let package = self.physical_package()?.clone();
        match policy {
            Some(policy) => package.write_part_overlay_shared_to_stream_with_policy(
                writer,
                partname,
                Arc::clone(payload),
                policy,
            )?,
            None => package.write_part_overlay_shared_to_stream(
                writer,
                partname,
                Arc::clone(payload),
            )?,
        }
        Ok(())
    }

    /// Publish one changed Part through the sealed OPC artifact target.
    ///
    /// The target owns the recipient sink and the complete ZIP/OPC census;
    /// this helper only constructs the bounded native topology plan that the
    /// ordinary publisher already uses for the same replacement.
    pub(crate) fn stream_part_overlay(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
        partname: &PackURI,
        payload: &Arc<Vec<u8>>,
    ) -> Result<ArtifactStreamReport> {
        let package = self.physical_package()?.clone();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_part_shared(partname.clone(), Arc::clone(payload))
            .map_err(Error::Opc)?;
        target.publish_overlay(package, plan).map_err(Error::Opc)
    }

    /// Publish several changed Parts and their explicitly removed external
    /// relationships through the sealed OPC artifact target.
    pub(crate) fn stream_relationship_overlays(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
        overlays: impl IntoIterator<Item = (PackURI, Vec<u8>, Vec<String>)>,
    ) -> Result<ArtifactStreamReport> {
        let package = self.physical_package()?.clone();
        let mut plan = SourceTopologyPlan::new();
        for (partname, replacement, removed_relationship_ids) in overlays {
            plan.try_replace_part(partname.clone(), replacement)
                .map_err(Error::Opc)?;
            for relationship_id in removed_relationship_ids {
                plan.try_remove_external_relationship(partname.clone(), relationship_id)
                    .map_err(Error::Opc)?;
            }
        }
        target.publish_overlay(package, plan).map_err(Error::Opc)
    }

    /// Create a deferred one-Part publication with both ordinary and sealed
    /// OPC output routes. The two routes deliberately share only immutable
    /// package/payload ownership; neither route borrows a caller sink.
    pub(crate) fn deferred_part_overlay(
        &self,
        partname: PackURI,
        payload: Arc<Vec<u8>>,
    ) -> Result<Self> {
        let ordinary_base = self.clone();
        let stream_base = ordinary_base.clone();
        let ordinary_part = partname.clone();
        let stream_part = partname;
        let ordinary_payload = Arc::clone(&payload);
        self.deferred(
            move |writer, policy| {
                ordinary_base.write_part_overlay(writer, &ordinary_part, &ordinary_payload, policy)
            },
            move |target| stream_base.stream_part_overlay(target, &stream_part, &payload),
        )
    }

    /// Check a section-layout patch against the current main document and
    /// return the part and its rewritten XML (record 0355a). `commit`, the
    /// commit that made the patch, lets its retained base inventory spare a
    /// second scan of an unchanged source.
    pub(crate) fn section_layout_target(
        &self,
        patch: &layout::Patch,
        commit: Option<&layout::Commit>,
        operation: &'static str,
    ) -> Result<(PackURI, Arc<Vec<u8>>, SourceArtifactFingerprint)> {
        self.trim_clean_paragraph_index_for_operation();
        let recipient_limits = self.semantic_limits.sections_value();
        let (main, current) = self.main_section_layout_snapshot(
            operation,
            recipient_limits,
            commit.and_then(layout::Commit::base_snapshot_for_publication),
        )?;
        let target = match commit.map(|commit| commit.apply_to_current(&current)) {
            Some(Ok(target)) => target,
            _ => patch.apply_content(&current)?,
        };
        let fingerprint = match current.artifact_fingerprint() {
            Some(fingerprint) => fingerprint,
            None => self.source_artifact_fingerprint()?,
        };
        Ok((main, target.shared_xml(), fingerprint))
    }

    /// The fingerprint of the complete source artifact.
    pub(crate) fn source_artifact_fingerprint(&self) -> Result<SourceArtifactFingerprint> {
        Ok(self.physical_package()?.source_artifact().fingerprint()?)
    }

    /// Capture the exact source closure used to detach external hyperlinks in
    /// the main document.
    ///
    /// The closure binds the raw main-document XML and its complete outbound
    /// relationship set. Enforced document or write protection is refused.
    /// Encrypted OOXML is rejected earlier because it is not a plaintext OPC
    /// package. No external target is fetched or executed.
    ///
    /// # Errors
    ///
    /// Returns a typed package, protection, markup-compatibility, XML, or
    /// resource-limit error.
    pub fn external_hyperlink_sanitization_snapshot(&self) -> Result<sanitize::Snapshot> {
        self.external_hyperlink_sanitization_snapshot_bounded(
            self.semantic_limits.hyperlinks_value(),
        )
    }

    /// Capture the external-hyperlink sanitization closure with explicit
    /// semantic scanner limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot`].
    pub(crate) fn external_hyperlink_sanitization_snapshot_bounded(
        &self,
        limits: sanitize::Limits,
    ) -> Result<sanitize::Snapshot> {
        self.preflight_external_hyperlink_document_xml(limits)?;
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_sanitization_snapshot", false)
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        let main = self.physical_package()?.main_document_part()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(main.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink relationship closure",
                source,
            })?;
        for relationship in main.rels().iter() {
            relationships.push(RelationshipState::new(
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            ));
        }
        sanitize::Snapshot::from_source(
            document
                .shared_xml()
                .map_err(transaction_error_to_document)?,
            relationships,
            limits,
        )
    }

    /// Build a non-mutating plan that detaches all external main-document
    /// hyperlink wrappers while retaining their visible child markup.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot`].
    pub fn plan_external_hyperlink_detachment(&self) -> Result<sanitize::SanitizePlan> {
        Ok(self.external_hyperlink_sanitization_snapshot()?.plan())
    }

    /// Capture the exact main-document closure for explicit irreversible
    /// external-hyperlink redaction.
    ///
    /// Unlike reversible wrapper detachment, this inventory is intended for a
    /// later consuming publication that removes relationship records. It
    /// refuses external link owners outside the main document, external
    /// relationship forms outside `w:hyperlink`, protection, signatures, and
    /// unsupported owner syntax. No target is fetched or executed.
    pub fn external_hyperlink_redaction_snapshot(&self) -> Result<redact::Snapshot> {
        self.external_hyperlink_redaction_snapshot_bounded(self.semantic_limits.hyperlinks_value())
    }

    /// Capture an irreversible redaction inventory with explicit XML limits.
    pub(crate) fn external_hyperlink_redaction_snapshot_bounded(
        &self,
        limits: sanitize::Limits,
    ) -> Result<redact::Snapshot> {
        self.preflight_external_hyperlink_document_xml(limits)?;
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_redaction_snapshot", false)
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        self.refuse_external_hyperlink_redaction_topology()?;
        let main = self.physical_package()?.main_document_part()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(main.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink redaction relationship closure",
                source,
            })?;
        for relationship in main.rels().iter() {
            relationships.push(RelationshipState::new(
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            ));
        }
        let source_version = self.physical_package()?.source_version()?;
        let source_fingerprint = self.physical_package()?.source_artifact().fingerprint()?;
        redact::Snapshot::from_source(
            document
                .shared_xml()
                .map_err(transaction_error_to_document)?,
            relationships,
            source_version,
            source_fingerprint,
            limits,
        )
    }

    /// Build a non-mutating forward-only plan for exact target URL values.
    pub fn plan_external_hyperlink_redaction(&self, target_urls: &[&str]) -> Result<redact::Plan> {
        self.external_hyperlink_redaction_snapshot()?
            .plan_target_urls(target_urls)
    }

    /// The payload part behind one active main-story altChunk anchor, for
    /// a replacement of `media_type` and `extension` (record 0355a).
    pub(crate) fn alt_chunk_payload_target(
        &self,
        selector: AltChunkSelector,
        media_type: &str,
        extension: &str,
    ) -> Result<PackURI> {
        self.trim_clean_paragraph_index_for_operation();
        self.alt_chunk_target(selector, media_type, extension)
    }

    /// The exact payload bytes of `part`.
    pub(crate) fn part_payload(&self, part: &PackURI) -> Result<Arc<Vec<u8>>> {
        Ok(self.physical_package()?.part(part)?.data()?.into_arc()?)
    }

    /// The main-document part name.
    pub(crate) fn main_part_name(&self) -> Result<PackURI> {
        Ok(self
            .physical_package()?
            .main_document_part()?
            .partname()
            .clone())
    }

    /// Check an external-hyperlink detachment patch against the current main
    /// document and return the part and its rewritten XML (record 0355a).
    pub(crate) fn external_hyperlink_detachment_target(
        &self,
        patch: &sanitize::Patch,
    ) -> Result<(PackURI, Arc<Vec<u8>>)> {
        self.trim_clean_paragraph_index_for_operation();
        let main = self
            .physical_package()?
            .main_document_part()?
            .partname()
            .clone();
        let current = self.external_hyperlink_sanitization_snapshot_bounded(
            self.semantic_limits.hyperlinks_value(),
        )?;
        let target = patch.apply(&current)?;
        Ok((main, target.shared_xml()))
    }

    /// Check an external-hyperlink redaction commit against the current main
    /// document and return the main part (record 0355a).
    pub(crate) fn external_hyperlink_redaction_target(
        &self,
        patch: &redact::IrreversibleRedactionPatch,
    ) -> Result<PackURI> {
        self.trim_clean_paragraph_index_for_operation();
        let main = self
            .physical_package()?
            .main_document_part()?
            .partname()
            .clone();
        let current = self.external_hyperlink_redaction_snapshot_bounded(
            self.semantic_limits.hyperlinks_value(),
        )?;
        patch.validate_source(&current)?;
        Ok(main)
    }

    /// Write this package with a redaction's main-document replacement and
    /// relationship removals (record 0355a).
    pub(crate) fn write_external_hyperlink_redaction(
        &self,
        writer: &mut dyn Write,
        main: &PackURI,
        patch: &redact::IrreversibleRedactionPatch,
        policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> Result<()> {
        let (replacement, removed_ids) = redact::publication_parts(patch);
        let package = self.physical_package()?.clone();
        match policy {
            Some(policy) => package
                .write_part_overlay_with_external_relationship_removals_to_stream_with_policy(
                    writer,
                    main,
                    replacement,
                    removed_ids,
                    policy,
                )?,
            None => package.write_part_overlay_with_external_relationship_removals_to_stream(
                writer,
                main,
                replacement,
                removed_ids,
            )?,
        }
        Ok(())
    }

    /// Capture the package-wide, relationship-owned story hyperlink closure.
    ///
    /// The inventory covers the main, header, footer, footnote, endnote,
    /// comments, and glossary stories. It does not alter the established
    /// main-only sanitization or redaction APIs above.
    pub fn story_hyperlinks_only_snapshot(&self) -> Result<crate::story_hyperlinks::Snapshot> {
        self.story_hyperlinks_only_snapshot_bounded(self.semantic_limits.story_hyperlinks_value())
    }

    /// Capture the story hyperlink closure under an explicit bounded policy.
    pub(crate) fn story_hyperlinks_only_snapshot_bounded(
        &self,
        limits: crate::story_hyperlinks::Limits,
    ) -> Result<crate::story_hyperlinks::Snapshot> {
        self.trim_clean_paragraph_index_for_operation();
        self.physical_package()?.check_execution()?;
        crate::story_hyperlinks::capture_source(self.physical_package()?, limits)
    }

    /// Plan exact target URL removal across all relationship-owned stories.
    ///
    /// Strict mode rejects unsupported external-reference classes before any
    /// publication. Best-effort mode retains explicit diagnostics and still
    /// fails closed for this tranche's security-sensitive classes.
    pub fn plan_story_hyperlink_redaction(
        &self,
        target_urls: &[&str],
        mode: crate::story_hyperlinks::Mode,
    ) -> Result<crate::story_hyperlinks::Plan> {
        self.plan_story_hyperlink_redaction_bounded(
            target_urls,
            mode,
            self.semantic_limits.story_hyperlinks_value(),
        )
    }

    /// Plan story hyperlink removal with explicit semantic bounds.
    pub(crate) fn plan_story_hyperlink_redaction_bounded(
        &self,
        target_urls: &[&str],
        mode: crate::story_hyperlinks::Mode,
        limits: crate::story_hyperlinks::Limits,
    ) -> Result<crate::story_hyperlinks::Plan> {
        self.story_hyperlinks_only_snapshot_bounded(limits)?
            .plan_target_urls_with_mode(target_urls, mode)
    }

    /// Check a story-hyperlink redaction against this exact source (record
    /// 0355a).
    pub(crate) fn story_hyperlink_redaction_admit(
        &self,
        patch: &crate::story_hyperlinks::ForwardOnlyPatch,
    ) -> Result<()> {
        self.trim_clean_paragraph_index_for_operation();
        self.physical_package()?.check_execution()?;
        patch.validate_source_identity(self.physical_package()?)
    }

    /// Write this package with a story-hyperlink redaction's story
    /// replacements and relationship removals (record 0355a).
    pub(crate) fn write_story_hyperlink_redaction(
        &self,
        writer: &mut dyn Write,
        patch: &crate::story_hyperlinks::ForwardOnlyPatch,
        policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> Result<()> {
        let parts = crate::story_hyperlinks::publication_parts(patch)?;
        let package = self.physical_package()?.clone();
        match policy {
            Some(policy) => package
                .write_part_overlays_with_external_relationship_removals_to_stream_with_policy(
                    writer, parts, policy,
                )?,
            None => package
                .write_part_overlays_with_external_relationship_removals_to_stream(writer, parts)?,
        }
        Ok(())
    }

    fn alt_chunk_target(
        &self,
        selector: AltChunkSelector,
        media_type: &str,
        extension: &str,
    ) -> Result<PackURI> {
        self.physical_package()?.check_execution()?;
        let main = self.physical_package()?.main_document_part()?;
        let main_name = main.partname().clone();
        let data = self.main_data(main)?;
        let chunks = crate::alt::scan(data.as_bytes())?;
        // `body_block_ranges` uses this stable private kind ordering for the
        // third variant, the direct-body altChunk block.
        const ALT_CHUNK_BLOCK_KIND: usize = 2;
        let body_ranges = body_block_ranges(data.as_bytes())?;
        let body_len = body_ranges
            .iter()
            .filter(|(kind, start, _, _)| {
                *kind == ALT_CHUNK_BLOCK_KIND && chunks.contains_key(start)
            })
            .count();
        let AltChunkSelector::Index(index) = selector;
        let chunk = body_ranges
            .into_iter()
            .filter_map(|(kind, start, _, _)| {
                (kind == ALT_CHUNK_BLOCK_KIND)
                    .then_some(chunks.get(&start))
                    .flatten()
            })
            .nth(index)
            .ok_or(Error::OutOfBounds {
                object: "altChunk",
                index,
                len: body_len,
            })?;
        let relationship = main
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                Error::InvalidRelationship(format!(
                    "altChunk relationship '{}' is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if !crate::alt::is_relationship(relationship.reltype()) {
            return Err(Error::InvalidRelationship(format!(
                "relationship '{}' is not an altChunk relationship",
                chunk.relationship().as_str()
            )));
        }
        if relationship.is_external() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "external altChunk targets cannot be replaced by a source-backed payload overlay",
            ));
        }
        let target = relationship.target_partname()?;
        let target_part = self.physical_package()?.part(&target)?;
        if target == main_name {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "an altChunk target cannot be the main document Part",
            ));
        }
        let target_kind = crate::alt::Kind::from_media_type(target_part.content_type());
        let replacement_kind = crate::alt::Kind::from_media_type(media_type);
        if target_kind != replacement_kind || matches!(target_kind, crate::alt::Kind::Unknown) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::InvalidValue,
                "publish_alt_chunk_to_stream",
                "replacement media type family must match the existing altChunk target Part",
            ));
        }
        if !target_part.rels().is_empty() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "altChunk targets with dependent relationships are outside the safe payload-only closure",
            ));
        }
        let relative = target
            .as_str()
            .strip_prefix("/word/")
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "altChunk target must be a direct child of the Word package directory",
            ))?;
        let Some((stem, target_extension)) = relative.rsplit_once('.') else {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "altChunk target must have a bounded foreign-payload extension",
            ));
        };
        if stem.is_empty()
            || relative.contains('/')
            || target_extension.is_empty()
            || !target_extension.eq_ignore_ascii_case(extension)
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "altChunk target path or extension is outside the safe payload-only layout",
            ));
        }

        let mut inbound = 0usize;
        for candidate in self.physical_package()?.rels().iter() {
            if !candidate.is_external() && candidate.target_partname()? == target {
                inbound = inbound.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("altChunk inbound count overflow".into())
                })?;
            }
        }
        for part in self.physical_package()?.iter_parts() {
            for candidate in part.rels().iter() {
                if !candidate.is_external() && candidate.target_partname()? == target {
                    inbound = inbound.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("altChunk inbound count overflow".into())
                    })?;
                }
            }
        }
        if inbound != 1 {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "publish_alt_chunk_to_stream",
                "the selected altChunk target is shared by an ambiguous relationship closure",
            ));
        }
        self.physical_package()?.source_version()?;
        Ok(target)
    }

    fn main_document_snapshot(
        &self,
        operation: &'static str,
        source_authorized: bool,
    ) -> TransactionResult<(PackURI, Snapshot)> {
        self.main_document_snapshot_with_before(operation, source_authorized, None, None)
    }

    fn main_document_snapshot_with_before(
        &self,
        operation: &'static str,
        source_authorized: bool,
        before: Option<&Snapshot>,
        edit_options: Option<&crate::document::EditOptions>,
    ) -> TransactionResult<(PackURI, Snapshot)> {
        if let Some(topology) = &self.effective {
            return self.effective_document_snapshot(
                topology,
                operation,
                source_authorized,
                before,
                edit_options,
            );
        }

        self.physical_package()?
            .source_version()
            .map_err(Error::from)?;
        let result = (|| {
            self.physical_package()?
                .check_execution()
                .map_err(Error::from)?;
            let main = self
                .physical_package()?
                .main_document_part()
                .map_err(Error::from)?;
            validate_document_main_content_type(main.content_type())?;
            if let Some(options) = edit_options
                && options.max_snapshot_bytes_value().is_some()
            {
                options
                    .admit_source_bytes(main.declared_uncompressed_size().map_err(Error::from)?)?;
            }
            let partname = main.partname();
            let context = self.physical_package()?.execution_context();
            let managed = self.physical_package()?.cache_diagnostics().budget_managed;
            if managed {
                // Transaction snapshots do not use the paragraph memo. Drop
                // only an unpinned clean entry before their own admissions.
                self.paragraph_index_cache.trim_clean();
            }
            let source_version = self
                .physical_package()?
                .source_version()
                .map_err(Error::from)?;
            let source_lineage = self.physical_package()?.source_lineage();
            let snapshot = if managed {
                if source_authorized {
                    let hinted_before = before
                        .filter(|snapshot| {
                            snapshot.source_identity_matches(
                                &source_lineage,
                                source_version,
                                partname,
                            )
                        })
                        .and_then(|snapshot| snapshot.source_xml().map(|hint| (snapshot, hint)));
                    let source = match &hinted_before {
                        Some((_, hint)) => main.source_xml_with_hint(hint).map_err(Error::from),
                        None => main.source_xml().map_err(Error::from),
                    };
                    match source {
                        Ok(source) => {
                            // The retained layout was admitted when the edit
                            // captured its original source. Reuse it only after
                            // OPC has checked the current Part and the exact
                            // snapshot source precondition still matches.
                            let reused = hinted_before.and_then(|(before, _)| {
                                before.reuse_if_source_xml_matches(
                                    &source,
                                    &source_lineage,
                                    source_version,
                                    partname,
                                )
                            });
                            if let Some(snapshot) = reused {
                                snapshot
                            } else {
                                self.ensure_source_document_xml(source.bytes(), context.as_ref())?;
                                Snapshot::from_source_xml(
                                    source,
                                    source_lineage.clone(),
                                    source_version,
                                    partname,
                                    context.clone().ok_or_else(|| {
                                        Error::InvalidFormat(
                                            "managed document context is unavailable".into(),
                                        )
                                    })?,
                                )?
                            }
                        },
                        Err(Error::Opc(
                            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
                        )) => {
                            let data = self.main_data(main)?;
                            self.ensure_source_document_xml(data.as_bytes(), context.as_ref())?;
                            Snapshot::from_managed_part(
                                data,
                                source_lineage.clone(),
                                source_version,
                                partname,
                                context.clone().ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "managed document context is unavailable".into(),
                                    )
                                })?,
                            )?
                        },
                        Err(error) => return Err(error.into()),
                    }
                } else {
                    let data = self.main_data(main)?;
                    self.ensure_source_document_xml(
                        data.as_bytes(),
                        self.physical_package()?.execution_context().as_ref(),
                    )?;
                    Snapshot::from_managed_part(
                        data,
                        source_lineage.clone(),
                        source_version,
                        partname,
                        context.clone().ok_or_else(|| {
                            Error::InvalidFormat("managed document context is unavailable".into())
                        })?,
                    )?
                }
            } else {
                let data = self.main_data(main)?;
                let raw = data.into_arc().map_err(Error::from)?;
                let snapshot: TransactionResult<Snapshot> = (|| {
                    self.admit_main_document(raw.as_slice(), source_version, operation)?;
                    if let Some(snapshot) =
                        self.memoized_main_snapshot(&raw, &source_lineage, source_version, partname)
                    {
                        return Ok(self.with_retained_main_proof(snapshot, || {
                            self.physical_package()
                                .ok()
                                .and_then(|package| package.xml_audit_proof(partname))
                        }));
                    }
                    let snapshot = Snapshot::from_shared_xml_with_source_identity(
                        raw,
                        source_lineage.clone(),
                        source_version,
                        partname,
                    )?;
                    self.remember_main_snapshot(&snapshot);
                    Ok(self.with_retained_main_proof(snapshot, || {
                        self.physical_package()
                            .ok()
                            .and_then(|package| package.xml_audit_proof(partname))
                    }))
                })();
                snapshot?
            };
            Ok((partname.clone(), snapshot))
        })();
        self.physical_package()?
            .source_version()
            .map_err(Error::from)?;
        self.physical_package()?
            .check_execution()
            .map_err(Error::from)?;
        result
    }

    /// `snapshot` with the source XML audit proof the OPC package retained
    /// for its exact bytes, when it has none yet and `retained` finds one
    /// (record 0317); the memo then keeps the proven snapshot.
    ///
    /// The proof is the package's own audit of these bytes (its indexed
    /// proof of the original, or the proof an earlier commit's candidate
    /// earned), so the preserving commit proves its candidate from it
    /// without scanning the base again. A proof for any other allocation is
    /// dropped by [`Snapshot::with_publication_proof`].
    pub(super) fn with_retained_main_proof(
        &self,
        snapshot: Snapshot,
        retained: impl FnOnce() -> Option<xml_minifier::audit::VerifiedSource>,
    ) -> Snapshot {
        if snapshot.publication_proof().is_some() {
            return snapshot;
        }
        let Some(proof) = retained() else {
            return snapshot;
        };
        let proven = snapshot.with_publication_proof(proof);
        if proven.publication_proof().is_some() {
            self.remember_main_snapshot(&proven);
        }
        proven
    }

    /// Keep `snapshot` as the unmanaged transaction snapshot memo.
    fn remember_main_snapshot(&self, snapshot: &Snapshot) {
        *self
            .main_snapshot
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner) = Some(snapshot.clone());
    }

    /// The memoized unmanaged snapshot, when it was built from this source
    /// revision's main part, or from an effective revision over it, and its
    /// bytes equal `raw`.
    fn memoized_main_snapshot(
        &self,
        raw: &[u8],
        lineage: &litchi_opc::SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
    ) -> Option<Snapshot> {
        let memo = self
            .main_snapshot
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let snapshot = memo.as_ref()?;
        let held = snapshot.xml_bytes();
        (snapshot.source_identity_matches(lineage, version, partname)
            && held.len() == raw.len()
            && (held.as_ptr() == raw.as_ptr() || held == raw))
            .then(|| snapshot.clone())
    }

    /// Prove the unmanaged main document admissible to a transaction over its
    /// source bytes (see [`crate::parts::document_part::admit_document_source`]),
    /// once per source revision.
    fn admit_main_document(
        &self,
        raw: &[u8],
        version: SourceVersion,
        operation: &'static str,
    ) -> Result<()> {
        // Admission only examines these already-retained bytes; no source
        // callback or publication runs while this shared memo is locked.
        let mut admitted = self
            .main_admission
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if *admitted == Some(version) {
            return Ok(());
        }
        crate::parts::document_part::admit_document_source(
            raw,
            operation,
            "source-backed document transactions do not support markup-compatibility branch selection",
        )?;
        *admitted = Some(version);
        Ok(())
    }

    fn main_section_layout_snapshot(
        &self,
        operation: &'static str,
        limits: &crate::section::Limits,
        retained_base: Option<&layout::Snapshot>,
    ) -> Result<(PackURI, layout::Snapshot)> {
        self.physical_package()?.check_execution()?;
        let main = self.physical_package()?.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let partname = main.partname().clone();
        if self.physical_package()?.cache_diagnostics().budget_managed {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                operation,
                "source-backed section layout edits require an owned immutable main-document snapshot",
            ));
        }
        let source_version = self.physical_package()?.source_version()?;
        let lineage = self.physical_package()?.source_lineage();
        let data = self.main_data(main)?;
        let raw = data.into_arc()?;
        let artifact_fingerprint = self.physical_package()?.source_artifact().fingerprint()?;
        let snapshot = if let Some(base) =
            retained_base.filter(|base| base.xml_matches(&raw) && base.limits() == limits)
        {
            Ok(layout::Snapshot::from_reused_inventory(
                Arc::clone(&raw),
                source_version,
                lineage.clone(),
                artifact_fingerprint,
                limits,
                base,
            ))
        } else {
            layout::Snapshot::from_source_xml(
                Arc::clone(&raw),
                source_version,
                lineage.clone(),
                artifact_fingerprint,
                limits,
            )
        };
        let observed = self.physical_package()?.source_version()?;
        if observed != source_version {
            return Err(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            }
            .into());
        }
        self.physical_package()?.check_execution()?;
        Ok((partname, snapshot?))
    }

    fn settings_document_variables_source(
        &self,
        operation: &'static str,
    ) -> Result<DocumentVariablesSource> {
        self.physical_package()?.check_execution()?;
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        let main = self.physical_package()?.main_document_part()?;
        let package_strict = self
            .physical_package()?
            .rels()
            .iter()
            .find_map(|relationship| {
                matches!(
                    relationship.reltype(),
                    rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
                )
                .then_some(relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT)
            });
        let mut matching = main.rels().iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
        });
        let relationship = matching.next().ok_or_else(|| {
            Error::InvalidRelationship(
                "source-backed document-variable editing requires an existing settings relationship"
                    .into(),
            )
        })?;
        if matching.next().is_some() {
            return Err(Error::InvalidRelationship(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::InvalidRelationship(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let mut inbound_owners = 0usize;
        for candidate in self.physical_package()?.rels().iter() {
            if !candidate.is_external() && candidate.target_partname()? == target {
                inbound_owners = inbound_owners.saturating_add(1);
            }
        }
        for part in self.physical_package()?.iter_parts() {
            for candidate in part.rels().iter() {
                if !candidate.is_external() && candidate.target_partname()? == target {
                    inbound_owners = inbound_owners.saturating_add(1);
                }
            }
        }
        let settings_part = self.physical_package()?.part(&target)?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::InvalidContentType {
                expected: ct::WML_SETTINGS.into(),
                got: settings_part.content_type().into(),
            });
        }
        if self.physical_package()?.cache_diagnostics().budget_managed {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                operation,
                "managed source-backed settings snapshots require an owned edit snapshot; use an unmanaged compatibility constructor",
            ));
        }
        let data = settings_part.data()?;
        let raw = data.into_arc()?;
        let variable_limits = variables::raw::Limits::default();
        admit_settings_source(raw.as_slice(), operation)?;
        let policy = variables::inspect_source_policy(raw.as_slice(), &variable_limits)?;
        let settings_strict = relationship.reltype() == STRICT_SETTINGS;
        let root_strict = policy.dialect == variables::SettingsDialect::Strict;
        if package_strict != Some(settings_strict) || settings_strict != root_strict {
            return Err(Error::InvalidRelationship(
                "main document, settings relationship, and settings XML use mixed OOXML conformance families"
                    .into(),
            ));
        }
        Ok(DocumentVariablesSource {
            partname: target,
            snapshot: variables::Snapshot::from_source_xml(
                raw,
                self.physical_package()?.source_version()?,
                variable_limits,
            )?,
            protected: policy.protected,
            unique_inbound_owner: inbound_owners == 1,
        })
    }

    fn refuse_protected_external_hyperlink_detachment(&self) -> Result<()> {
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        if self.physical_package()?.cache_diagnostics().budget_managed {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "external_hyperlink_wrapper_detachment",
                "managed source-backed hyperlink sanitization requires an owned settings snapshot; use an unmanaged compatibility constructor",
            ));
        }

        let main = self.physical_package()?.main_document_part()?;
        let mut settings_relationships = main.rels().iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
        });
        let Some(relationship) = settings_relationships.next() else {
            return Ok(());
        };
        if settings_relationships.next().is_some() {
            return Err(Error::InvalidRelationship(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::InvalidRelationship(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let settings_part = self.physical_package()?.part(&target)?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::InvalidContentType {
                expected: ct::WML_SETTINGS.into(),
                got: settings_part.content_type().into(),
            });
        }
        let mut staged = BlobPart::new(
            target,
            ct::WML_SETTINGS.to_owned(),
            settings_part.data()?.as_bytes().to_vec(),
        );
        for relationship in settings_part.rels().iter() {
            staged.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )?;
        }
        let limits =
            crate::mail_merge::mce_limits_for_source(&self.physical_package()?.read_limits());
        let settings = DocumentSettings::extract_from_part_with_mce_limits(&staged, &limits)?;
        if settings.is_protected() || settings.is_write_protected() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "external_hyperlink_wrapper_detachment",
                "document or write protection is enforced",
            ));
        }
        Ok(())
    }

    fn refuse_external_hyperlink_redaction_topology(&self) -> Result<()> {
        const SIGNATURE_TYPES: &[&str] = &[
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
        ];
        let main = self.physical_package()?.main_document_part()?;
        if self
            .physical_package()?
            .rels()
            .iter()
            .any(|relationship| SIGNATURE_TYPES.contains(&relationship.reltype()))
        {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if self
            .physical_package()?
            .rels()
            .iter()
            .any(|relationship| relationship.is_external())
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "external_hyperlink_redaction",
                "a package-level relationship has an external target outside the redaction closure",
            ));
        }
        for part in self.physical_package()?.iter_parts() {
            if part.partname().as_str().starts_with("/_xmlsignatures/")
                || matches!(
                    part.content_type(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN
                        | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                        | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
                )
                || part
                    .rels()
                    .iter()
                    .any(|relationship| SIGNATURE_TYPES.contains(&relationship.reltype()))
            {
                return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
            }
            if part.partname() != main.partname()
                && part
                    .rels()
                    .iter()
                    .any(|relationship| relationship.is_external())
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "external_hyperlink_redaction",
                    "a non-main-document part owns an external hyperlink relationship",
                ));
            }
        }
        Ok(())
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn transaction_error_to_document(error: crate::document::Error) -> Error {
    match error {
        crate::document::Error::Document(error) => error,
        other => Error::InvalidFormat(format!(
            "external hyperlink-wrapper detachment snapshot could not be captured: {other}"
        )),
    }
}

/// Managed source-backed document reads must not run the MCE processor: its
/// owned `Cow` output would be an unbudgeted duplicate of the retained
/// `PartData` payload. The existing source-backed edit boundary already
/// refuses any MCE projection, so reject actual MCE elements, attributes, or
/// namespace bindings before a materialization can occur.
// These ceilings mirror the default OPC per-part/event/depth policy. The
// payload read has already enforced the caller's selected part-byte limit;
// the explicit byte ceiling also keeps a caller-selected looser policy from
// turning this namespace-only pass into an unbounded operation.
const SOURCE_DOCUMENT_SCAN_MAX_BYTES: usize = 512 * 1024 * 1024;
const SOURCE_DOCUMENT_SCAN_MAX_EVENTS: usize = 1_000_000;
const SOURCE_DOCUMENT_SCAN_MAX_DEPTH: usize = 256;
const SOURCE_DOCUMENT_SCAN_MAX_NAME_BYTES: usize = 64 * 1024;

fn source_document_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::Cancelled => Error::Opc(litchi_opc::OpcError::Cancelled),
        error => Error::Opc(litchi_opc::OpcError::Execution(error)),
    }
}

fn is_index_pressure_error(error: &Error) -> bool {
    matches!(
        error,
        Error::Opc(litchi_opc::OpcError::Execution(
            ExecutionError::ResourceLimit(limit)
        )) if matches!(limit.resource, Resource::Memory | Resource::Objects)
    )
}

fn ensure_source_document_xml(xml: &[u8], context: Option<&ExecutionContext>) -> Result<()> {
    if xml.len() > SOURCE_DOCUMENT_SCAN_MAX_BYTES {
        return Err(Error::InvalidFormat(
            "source-backed document XML exceeds the bounded MCE scan byte limit".into(),
        ));
    }

    // Admit a conservative envelope before quick-xml or the topology stack
    // allocates. Namespace buffers and their per-event clone are bounded by
    // the input, with headroom for collection capacity and descriptor storage.
    // These are scoped reservations, not a claim about measured peak usage.
    let _workspace = context
        .map(|context| {
            let bytes = (xml.len() as u64)
                .saturating_mul(32)
                .saturating_add(131_072);
            let memory = context.reserve(Resource::Memory, bytes)?;
            let objects = context.reserve(Resource::Objects, xml.len() as u64 + 1024)?;
            let depth = context.reserve(Resource::Depth, SOURCE_DOCUMENT_SCAN_MAX_DEPTH as u64)?;
            Ok::<_, ExecutionError>((memory, objects, depth))
        })
        .transpose()
        .map_err(source_document_execution_error)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    // Namespace resolution is all this pass needs. Disable quick-xml's own
    // end-name comparison path; the fixed-capacity name stack below performs
    // exact topology validation while the scalar depth counter bounds
    // resolver scope growth.
    reader.config_mut().check_end_names = false;
    let mut open_names = Vec::new();
    open_names
        .try_reserve_exact(SOURCE_DOCUMENT_SCAN_MAX_DEPTH)
        .map_err(|source| Error::Allocation {
            resource: "source-backed document XML topology stack",
            source,
        })?;
    let mut open_name_bytes = 0usize;
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;
    let mut saw_declaration = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("document XML event counter overflow".into()))?;
        if events > SOURCE_DOCUMENT_SCAN_MAX_EVENTS {
            return Err(Error::InvalidFormat(
                "source-backed document XML exceeds the bounded MCE scan event limit".into(),
            ));
        }
        if let Some(context) = context {
            // Resolver cloning can revisit the complete namespace buffer on
            // each event. Charge its input-sized upper bound before that work.
            context
                .consume(Resource::Work, xml.len() as u64 + 1)
                .map_err(source_document_execution_error)?;
        }
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "source-backed document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("document XML depth overflow".into()))?;
                if depth > SOURCE_DOCUMENT_SCAN_MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML exceeds the bounded MCE scan depth limit"
                            .into(),
                    ));
                }
                retain_source_document_name(
                    &mut open_names,
                    &mut open_name_bytes,
                    element.name().as_ref(),
                )?;
                ensure_source_mce_element(&resolver, decoder, &namespace, &element, "document")?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "source-backed document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                    root_closed = true;
                }
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("document XML depth overflow".into()))?;
                if element_depth > SOURCE_DOCUMENT_SCAN_MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML exceeds the bounded MCE scan depth limit"
                            .into(),
                    ));
                }
                ensure_source_document_name_length(element.name().as_ref())?;
                ensure_source_mce_element(&resolver, decoder, &namespace, &element, "document")?;
            },
            Event::End(element) => {
                let expected = open_names
                    .pop()
                    .ok_or_else(|| Error::InvalidFormat("document XML depth underflow".into()))?;
                open_name_bytes = open_name_bytes.checked_sub(expected.len()).ok_or_else(|| {
                    Error::InvalidFormat("document XML name-byte counter underflow".into())
                })?;
                if expected.as_slice() != element.name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "mismatched source-backed document XML end tag".into(),
                    ));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("document XML depth underflow".into()))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof => {
                if !saw_root {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML does not contain exactly one root".into(),
                    ));
                }
                if !open_names.is_empty() || depth != 0 {
                    return Err(Error::InvalidFormat(
                        "unclosed source-backed document XML element".into(),
                    ));
                }
                break;
            },
            Event::Text(text) => {
                if depth == 0 && !is_xml_outer_whitespace(text.as_ref()) {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML has CDATA outside its root".into(),
                    ));
                }
            },
            Event::Comment(_) => {},
            Event::Decl(_) => {
                if saw_declaration || saw_root || root_closed {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML declaration is not in the prolog".into(),
                    ));
                }
                saw_declaration = true;
            },
            Event::PI(_) => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "source-backed document read",
                    "processing instructions are not accepted in a managed source-backed document",
                ));
            },
            Event::DocType(_) => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "source-backed document read",
                    "DTD declarations are not accepted in a managed source-backed document",
                ));
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "source-backed document XML has a reference outside its root".into(),
                    ));
                }
                if reference.is_char_ref() {
                    let value = reference
                        .resolve_char_ref()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "numeric source-backed document reference did not resolve".into(),
                            )
                        })?;
                    if !is_legal_xml_character(value) {
                        return Err(Error::InvalidFormat(
                            "numeric source-backed document reference is not a legal XML character"
                                .into(),
                        ));
                    }
                } else {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if !matches!(name.as_ref(), "amp" | "apos" | "gt" | "lt" | "quot") {
                        return Err(Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedContent,
                            "source-backed document read",
                            "non-predefined entity references are not accepted in a managed source-backed document",
                        ));
                    }
                }
            },
        }
    }
    admit_managed_document_markup_compatibility(xml, context)
}

/// The markup-compatibility admission of a managed source-backed document.
///
/// A managed document is read and edited over its source bytes: there is no
/// projection, because building one would need an unbudgeted owned payload.
/// The shared proof admits it only when the projection would differ from
/// the source in spelling alone (no markup-compatibility use at all, or
/// valid `Ignorable` directives that drop nothing), so a managed
/// reader sees what an unmanaged reader sees in the projection. The proof
/// runs inside the caller's workspace reservation and charges its work per
/// XML event; cancellation and budget failures keep their typed identity.
///
/// Everything else, including ignorable markup the projection would drop,
/// keeps the historical refusal.
fn admit_managed_document_markup_compatibility(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> Result<()> {
    use litchi_ooxml_common::mce;

    if !mce::mentions_namespace(xml) {
        return Ok(());
    }
    let capabilities = visible_document_capabilities();
    let verdict =
        mce::admission::admit_with(xml, &capabilities, &mce::Limits::default(), |bytes| {
            if let Some(context) = context {
                let work = u64::try_from(bytes).unwrap_or(u64::MAX).saturating_add(1);
                context
                    .consume(Resource::Work, work)
                    .map_err(source_document_execution_error)?;
            }
            Ok::<(), Error>(())
        })?;
    match verdict {
        Ok(admission) if admission.is_lexical() => Ok(()),
        Ok(_) | Err(_) => Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "source-backed document read",
            "markup-compatibility preprocessing would require an unbudgeted owned payload",
        )),
    }
}

/// Admit a settings part to a document-variable transaction over its source
/// bytes: the main-document rule with the OOXML baseline profile that the
/// settings readers use. Word's `w14`/`w15` settings leaves are ignorable
/// and not understood by that profile, so the projection drops them; they
/// hold no WordprocessingML, and the transaction reads only `w:docVars`.
///
/// A part the proof cannot admit keeps the historical answer: the
/// projection's own error, or the refusal of a projection that differs from
/// the source.
fn admit_settings_source(raw: &[u8], operation: &'static str) -> Result<()> {
    use litchi_ooxml_common::mce;

    if !mce::mentions_namespace(raw) {
        return Ok(());
    }
    let capabilities = mce::Capabilities::default();
    let limits = mce::Limits::default();
    if mce::admission::admit(raw, &capabilities, &limits).is_ok() {
        return Ok(());
    }
    match mce::process_markup_compatibility(raw, &capabilities, &limits)?.xml {
        Cow::Borrowed(_) => Ok(()),
        Cow::Owned(_) => Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            operation,
            "source-backed document-variable transactions do not support markup-compatibility branch selection",
        )),
    }
}

fn retain_source_document_name(
    open_names: &mut Vec<Vec<u8>>,
    open_name_bytes: &mut usize,
    name: &[u8],
) -> Result<()> {
    ensure_source_document_name_length(name)?;
    let next_name_bytes = open_name_bytes
        .checked_add(name.len())
        .ok_or_else(|| Error::InvalidFormat("document XML name-byte counter overflow".into()))?;
    if next_name_bytes > SOURCE_DOCUMENT_SCAN_MAX_NAME_BYTES {
        return Err(Error::InvalidFormat(
            "source-backed document XML exceeds the bounded aggregate name-byte limit".into(),
        ));
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed document XML element name",
            source,
        })?;
    owned.extend_from_slice(name);
    open_names
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "source-backed document XML topology stack",
            source,
        })?;
    open_names.push(owned);
    *open_name_bytes = next_name_bytes;
    Ok(())
}

fn ensure_source_document_name_length(name: &[u8]) -> Result<()> {
    if name.len() > SOURCE_DOCUMENT_SCAN_MAX_NAME_BYTES {
        return Err(Error::InvalidFormat(
            "source-backed document XML exceeds the bounded name-byte limit".into(),
        ));
    }
    Ok(())
}

/// Keep the source-backed section facade on the exact retained PartData
/// representation. The general in-memory section parser intentionally
/// supports MCE branch selection and ignores DTD/entity events because it is
/// used by compatibility callers. A source-bound inventory cannot publish a
/// descriptor whose semantics came from an unbudgeted MCE projection or from
/// entity expansion, so reject those constructs before that parser runs.
pub(crate) fn ensure_source_section_inventory_xml(
    xml: &[u8],
    limits: &crate::section::Limits,
) -> Result<()> {
    if xml.len() > limits.max_input_bytes {
        return Err(Error::SectionInventoryLimit {
            kind: crate::section::LimitKind::InputBytes,
            maximum: limits.max_input_bytes,
            actual: xml.len(),
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;
    let mut saw_declaration = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("section XML event counter overflow".into()))?;
        if events > limits.max_events {
            return Err(Error::SectionInventoryLimit {
                kind: crate::section::LimitKind::XmlEvents,
                maximum: limits.max_events,
                actual: events,
            });
        }
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "source-backed section XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?;
                if depth > limits.max_depth {
                    return Err(Error::SectionInventoryLimit {
                        kind: crate::section::LimitKind::XmlDepth,
                        maximum: limits.max_depth,
                        actual: depth,
                    });
                }
                validate_source_section_element(resolver, decoder, &namespace, &element)?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "source-backed section XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                    root_closed = true;
                }
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?;
                if element_depth > limits.max_depth {
                    return Err(Error::SectionInventoryLimit {
                        kind: crate::section::LimitKind::XmlDepth,
                        maximum: limits.max_depth,
                        actual: element_depth,
                    });
                }
                validate_source_section_element(resolver, decoder, &namespace, &element)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth underflow".into()))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "source-backed section inventory",
                    "DTD declarations are not accepted in a source-bound inventory",
                ));
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "source-backed section XML has a reference outside its root".into(),
                    ));
                }
                if reference.is_char_ref() {
                    let value = reference
                        .resolve_char_ref()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "numeric XML character reference did not resolve".into(),
                            )
                        })?;
                    if !is_legal_xml_character(value) {
                        return Err(Error::InvalidFormat(
                            "numeric XML character reference is not a legal XML character".into(),
                        ));
                    }
                } else {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if !matches!(name.as_ref(), "amp" | "apos" | "gt" | "lt" | "quot") {
                        return Err(Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedContent,
                            "source-backed section inventory",
                            "non-predefined entity references are not accepted in a source-bound inventory",
                        ));
                    }
                }
            },
            Event::Eof => {
                if !saw_root {
                    return Err(Error::InvalidFormat(
                        "source-backed section XML does not contain exactly one root".into(),
                    ));
                }
                if depth != 0 {
                    return Err(Error::InvalidFormat(
                        "unclosed source-backed section XML element".into(),
                    ));
                }
                break;
            },
            Event::Text(text) => {
                if depth == 0 && !is_xml_outer_whitespace(text.as_ref()) {
                    return Err(Error::InvalidFormat(
                        "source-backed section XML has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "source-backed section XML has CDATA outside its root".into(),
                ));
            },
            Event::CData(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Decl(_) => {
                if saw_declaration || saw_root || root_closed {
                    return Err(Error::InvalidFormat(
                        "source-backed section XML declaration is not in the prolog".into(),
                    ));
                }
                saw_declaration = true;
            },
        }
    }
    Ok(())
}

fn validate_source_section_element(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<()> {
    ensure_source_mce_element(resolver, decoder, namespace, element, "section inventory")?;

    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        validate_source_section_attribute_value(attribute.value.as_ref())?;
    }
    Ok(())
}

fn ensure_source_mce_element(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    operation: &'static str,
) -> Result<()> {
    let reason = if operation == "document" {
        "markup-compatibility preprocessing would require an unbudgeted owned payload"
    } else {
        "markup-compatibility elements are not accepted in a source-bound inventory"
    };
    if is_mce_namespace(namespace) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            if operation == "document" {
                "source-backed document read"
            } else {
                "source-backed section inventory"
            },
            reason,
        ));
    }

    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        validate_source_attribute_value(
            attribute.value.as_ref(),
            if operation == "document" {
                "source-backed document read"
            } else {
                "source-backed section inventory"
            },
        )?;
        if operation == "document" {
            // A document's declarations and directives are decided on
            // decoded names by the shared admission proof once this scan
            // has finished (`admit_managed_document_markup_compatibility`).
            continue;
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        if is_mce_namespace(&attribute_namespace)
            || is_namespace_declaration(attribute.key)
                && source_namespace_binding_is_mce(
                    attribute.value.as_ref(),
                    decoder,
                    if operation == "document" {
                        "source-backed document read"
                    } else {
                        "source-backed section inventory"
                    },
                )?
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                if operation == "document" {
                    "source-backed document read"
                } else {
                    "source-backed section inventory"
                },
                if operation == "document" {
                    "markup-compatibility preprocessing would require an unbudgeted owned payload"
                } else {
                    "markup-compatibility namespace bindings are not accepted in a source-bound inventory"
                },
            ));
        }
    }
    Ok(())
}

fn is_namespace_declaration(key: quick_xml::name::QName<'_>) -> bool {
    (key.prefix().is_none() && key.local_name().as_ref() == b"xmlns")
        || key
            .prefix()
            .as_ref()
            .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
}

fn source_namespace_binding_is_mce(
    value: &[u8],
    decoder: quick_xml::encoding::Decoder,
    operation: &'static str,
) -> Result<bool> {
    // The MCE namespace is ASCII. Rejecting non-ASCII source bytes here
    // avoids asking the decoder to allocate a converted copy merely to prove
    // that the value cannot equal the bounded target namespace.
    if value.iter().any(|byte| !byte.is_ascii()) {
        return Ok(false);
    }
    let decoded = decoder
        .decode(value)
        .map_err(|error| Error::Xml(error.to_string()))?;
    let source = decoded.as_ref();
    let target = litchi_ooxml_common::mce::NAMESPACE.as_bytes();
    let mut target_position = 0usize;
    let mut matches = true;
    let mut compare = |character: char| {
        if !matches {
            return;
        }
        if character.is_ascii()
            && target_position < target.len()
            && target[target_position] == character as u8
        {
            target_position += 1;
        } else {
            matches = false;
        }
    };
    let mut start = 0usize;
    while let Some(relative) = source[start..].find('&') {
        let ampersand = start + relative;
        for character in source[start..ampersand].chars() {
            compare(character);
        }
        let Some(relative_end) = source[ampersand + 1..].find(';') else {
            for character in source[ampersand..].chars() {
                compare(character);
            }
            return Ok(matches && target_position == target.len());
        };
        let end = ampersand + 1 + relative_end;
        let name = &source[ampersand + 1..end];
        if let Some(entity) = predefined_xml_entity(name) {
            for character in entity.chars() {
                compare(character);
            }
        } else if name.starts_with('#') {
            let reference = BytesRef::new(name);
            let character = reference
                .resolve_char_ref()
                .map_err(|error| Error::Xml(error.to_string()))?
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "numeric source namespace reference did not resolve".into(),
                    )
                })?;
            if !is_legal_xml_character(character) {
                return Err(Error::InvalidFormat(
                    "numeric source namespace reference is not a legal XML character".into(),
                ));
            }
            compare(character);
        } else {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                operation,
                "non-predefined entity references are not accepted in a source-backed XML namespace binding",
            ));
        }
        start = end + 1;
    }
    for character in source[start..].chars() {
        compare(character);
    }
    Ok(matches && target_position == target.len())
}

fn predefined_xml_entity(name: &str) -> Option<&'static str> {
    match name {
        "amp" => Some("&"),
        "apos" => Some("'"),
        "gt" => Some(">"),
        "lt" => Some("<"),
        "quot" => Some("\""),
        _ => None,
    }
}

fn validate_source_section_attribute_value(value: &[u8]) -> Result<()> {
    validate_source_attribute_value(value, "source-backed section inventory")
}

fn validate_source_attribute_value(value: &[u8], operation: &'static str) -> Result<()> {
    let mut start = 0usize;
    while let Some(relative) = value[start..].iter().position(|byte| *byte == b'&') {
        let ampersand = start + relative;
        let end = value
            .get(ampersand + 1..)
            .and_then(|tail| tail.iter().position(|byte| *byte == b';'))
            .map(|relative| ampersand + 1 + relative)
            .ok_or_else(|| Error::InvalidFormat("unterminated XML attribute reference".into()))?;
        let name = std::str::from_utf8(&value[ampersand + 1..end])
            .map_err(|error| Error::Xml(error.to_string()))?;
        let reference = BytesRef::new(name);
        if reference.is_char_ref() {
            let value = reference
                .resolve_char_ref()
                .map_err(|error| Error::Xml(error.to_string()))?
                .ok_or_else(|| {
                    Error::InvalidFormat("numeric XML attribute reference did not resolve".into())
                })?;
            if !is_legal_xml_character(value) {
                return Err(Error::InvalidFormat(
                    "numeric XML attribute reference is not a legal XML character".into(),
                ));
            }
        } else if !matches!(name, "amp" | "apos" | "gt" | "lt" | "quot") {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                operation,
                "non-predefined entity references are not accepted in a source-bound attribute",
            ));
        }
        start = end + 1;
    }
    Ok(())
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == litchi_ooxml_common::mce::NAMESPACE.as_bytes()
    )
}

fn is_legal_xml_character(value: char) -> bool {
    matches!(
        value as u32,
        0x0009 | 0x000a | 0x000d | 0x0020..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff
    )
}

#[derive(Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "inline retained handles avoid an extra unadmitted allocation on managed reads"
)]
enum DocumentPayload {
    /// Compatibility payload whose Arc ownership is not attached to a
    /// hierarchical execution budget.
    Owned(Arc<Vec<u8>>),
    /// Managed payload retaining its OPC reservation for the document view.
    Managed(PartData),
    Effective(litchi_opc::PreparedPartData),
}

impl DocumentPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(xml) => xml.as_slice(),
            Self::Managed(data) => data.as_bytes(),
            Self::Effective(data) => data.as_bytes(),
        }
    }

    const fn is_managed(&self) -> bool {
        matches!(self, Self::Managed(_) | Self::Effective(_))
    }
}

struct DocumentQueryAdmission {
    _memory: Reservation,
    _objects: Reservation,
    _depth: Reservation,
}

fn admit_document_query_parser(
    context: Option<&ExecutionContext>,
    xml_len: usize,
) -> Result<Option<DocumentQueryAdmission>> {
    context
        .map(|context| {
            let bytes = xml_len as u64;
            let memory = context.reserve(
                Resource::Memory,
                bytes.saturating_mul(32).saturating_add(131_072),
            )?;
            let objects = context.reserve(Resource::Objects, bytes.saturating_add(1024))?;
            let depth = context.reserve(Resource::Depth, SOURCE_DOCUMENT_SCAN_MAX_DEPTH as u64)?;
            // Existing namespace scanners can revisit all active binding bytes per
            // event. Prepay the bounded input-by-event envelope before entering
            // those parsers; this is a work ceiling, not a measured operation count.
            context.consume(
                Resource::Work,
                bytes
                    .saturating_add(1)
                    .saturating_mul(bytes.saturating_add(1)),
            )?;
            Ok(DocumentQueryAdmission {
                _memory: memory,
                _objects: objects,
                _depth: depth,
            })
        })
        .transpose()
        .map_err(source_document_execution_error)
}

/// A pinned read-only view of a DOCX main document.
///
/// This view owns the main-document bytes loaded by [`Package::document`].
/// It intentionally exposes semantic text and paragraph queries only; use
/// the established [`crate::Package`] APIs for mutable package access.
#[derive(Clone)]
pub struct Document {
    xml: DocumentPayload,
    _retained: Option<litchi_opc::PreparedPartData>,
    /// Bounded offsets into the pinned XML, admitted separately from its
    /// payload. Managed opens fill this during [`Package::document`]; unmanaged
    /// opens fill it on the first paragraph query.
    paragraph_memo: OnceLock<Option<Arc<ParagraphIndexMemo>>>,
    _index_admission: Option<Arc<DocumentIndexAdmission>>,
    source_version: SourceVersion,
    execution: Option<ExecutionContext>,
    paragraph_index_cache: Arc<ParagraphIndexCache>,
    paragraph_route: ParagraphIndexRoute,
    /// Whether the visible XML keeps every element where the raw part has it.
    positions: RawPositions,
    /// The body block index, pinned by this view after its first body query.
    body_memo: OnceLock<Arc<crate::document::BlockIndex>>,
    /// Only unmanaged base and donor contexts can detach Arc-backed values.
    allows_detached_views: bool,
    /// Everything repaired while this pinned view was read, shared by its
    /// clones. See `crate::leniency`.
    tolerance: Arc<Mutex<crate::leniency::ToleranceReport>>,
    /// The semantic read budgets of the package the view came from.
    limits: Arc<crate::SemanticLimits>,
}

impl Document {
    /// A snapshot of everything repaired while this pinned view (or one of
    /// its clones) was read, in discovery order.
    ///
    /// The report starts with the OPC packaging repairs made when the
    /// package was opened ([`crate::leniency::Defect::Package`]).
    /// Tracked-revision repairs ([`crate::leniency::RevisionDefect`]) from
    /// [`Self::revisions`] and section column-model repairs
    /// ([`crate::leniency::SectionDefect`]) from the section inventory are
    /// recorded here; an empty report means every read conformed.
    #[must_use]
    pub fn tolerance_report(&self) -> crate::leniency::ToleranceReport {
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .clone()
    }

    /// Merge one read's repairs into this view's report.
    fn merge_tolerance(&self, repaired: &crate::leniency::ToleranceReport) {
        if repaired.is_clean() {
            return;
        }
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .append(repaired);
    }

    /// Borrow the bounded paragraph offsets for this pinned view.
    ///
    /// Managed opens fill the cell inside [`Package::document`], under the
    /// admission reserved there, so this never scans for them. Unmanaged opens
    /// reserve nothing for the index and swallow its failure with `.ok()`, so
    /// the scan is deferred to the first paragraph query: the same operations
    /// can observe the same swallowed failure, and text, table, block and
    /// element reads - which never consult the index - no longer pay for it.
    fn paragraph_ranges(&self) -> Option<&ParagraphIndex> {
        self.paragraph_memo
            .get_or_init(|| {
                self.paragraph_index_cache
                    .get_or_build_source(self.source_version, self.paragraph_route, || {
                        let build = ParagraphIndex::from_xml(self.xml.as_bytes())
                            .ok()
                            .map(|index| ParagraphIndexBuild::indexed(index, None))
                            .unwrap_or_else(|| ParagraphIndexBuild::unavailable(None));
                        Ok(build)
                    })
                    .ok()
                    .and_then(|build| build.memo)
            })
            .as_ref()
            .map(|memo| &memo.index)
    }

    fn check_execution(&self) -> Result<()> {
        if let Some(data) = &self._retained {
            data.check_execution()?;
        }
        let Some(context) = self.execution.as_ref() else {
            return Ok(());
        };
        context.check().map_err(|error| {
            Error::Opc(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                other => litchi_opc::OpcError::Execution(other),
            })
        })
    }

    fn admit_query_parser(&self, xml_len: usize) -> Result<Option<DocumentQueryAdmission>> {
        self.paragraph_index_cache.trim_clean();
        admit_document_query_parser(self.execution.as_ref(), xml_len)
    }

    fn check_selective_operation(&self, operation: &'static str) -> Result<()> {
        self.check_execution()?;
        if self.xml.is_managed() || !self.allows_detached_views {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                operation,
                "this query would return Arc-backed semantic views that cannot retain the managed PartData reservation",
            ));
        }
        Ok(())
    }

    /// Return the physical source version captured for this view.
    /// Logical candidates can share this version; it is not a semantic cache key.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Extract all visible paragraph text from the pinned document.
    pub fn extract_text(&self) -> Result<String> {
        self.check_execution()?;
        let _parser = self.admit_query_parser(self.xml.as_bytes().len())?;
        let result = crate::paragraph::extract_word_text(self.xml.as_bytes());
        self.check_execution()?;
        result
    }

    /// Extract all visible paragraph text from the pinned document, with
    /// `separator` between paragraphs.
    ///
    /// The text is [`Self::extract_text`]'s, with `separator` before every
    /// paragraph after the first. The paragraphs are the ones
    /// [`Self::paragraphs`] returns: a paragraph nested in another one, such
    /// as a text box's, stays part of the paragraph that anchors it, and an
    /// empty paragraph adds an empty segment. For well-formed XML the result
    /// is the paragraph texts joined by `separator`; malformed XML fails as
    /// [`Self::extract_text`] fails, where [`Self::paragraphs`] may return
    /// the paragraphs it found.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::extract_text`].
    pub fn extract_text_with_paragraph_separator(&self, separator: &str) -> Result<String> {
        self.check_execution()?;
        let _parser = self.admit_query_parser(self.xml.as_bytes().len())?;
        let result = crate::paragraph::extract_word_paragraph_text(self.xml.as_bytes(), separator);
        self.check_execution()?;
        result
    }

    /// Read recognized tracked revisions throughout the visible main document.
    /// Includes body-final section properties and table-grid/row metadata that
    /// paragraph-only traversal cannot visit. Unknown markup remains in source.
    ///
    /// # Errors
    /// Returns malformed metadata, XML, or revision resource-limit errors.
    pub fn revisions(&self) -> Result<SmallVec<[crate::revision::Revision; 4]>> {
        self.revisions_bounded(self.limits.revisions_value())
    }

    /// Read tracked revision metadata with explicit parser and retention limits.
    ///
    /// # Errors
    /// Returns malformed metadata, XML, or revision resource-limit errors.
    pub(crate) fn revisions_bounded(
        &self,
        limits: crate::revision::Limits,
    ) -> Result<SmallVec<[crate::revision::Revision; 4]>> {
        self.check_execution()?;
        // The revision limits bound the parse itself; a managed view still
        // admits and charges the same parser workspace as every other
        // whole-document query before it parses.
        let _parser = self.admit_query_parser(self.xml.as_bytes().len())?;
        let mut repaired = crate::leniency::ToleranceReport::default();
        let result = crate::revision::parse_revisions_tolerant(
            self.xml.as_bytes(),
            &[],
            limits,
            &mut repaired,
        );
        self.merge_tolerance(&repaired);
        self.check_execution()?;
        result
    }

    /// Count visible paragraphs in the pinned document.
    pub fn paragraph_count(&self) -> Result<usize> {
        self.check_execution()?;
        if let Some(index) = self.paragraph_ranges() {
            return Ok(index.len());
        }
        let _parser = self.admit_query_parser(self.xml.as_bytes().len())?;
        let result = document_paragraph_count(self.xml.as_bytes());
        self.check_execution()?;
        result
    }

    /// Extract one visible paragraph's text without constructing an
    /// Arc-backed [`Paragraph`] view. This is available for managed
    /// source-backed documents because the returned `String` is caller-owned
    /// semantic output, not a hidden payload alias.
    ///
    /// `index` counts every outermost paragraph of the main story in document
    /// order — body paragraphs, paragraphs in table cells and paragraphs in
    /// body-level content controls — as [`Self::paragraph`] and
    /// [`Self::paragraph_count`] do. For the direct body paragraph a path such
    /// as `/body/p[N]` names, use [`Self::body_paragraph`].
    pub fn paragraph_text(&self, index: usize) -> Result<Option<String>> {
        self.check_execution()?;
        let selected = self.paragraph_ranges().and_then(|paragraph_index| {
            paragraph_index
                .get(index)
                .map(|range| (range.start, range.length))
        });
        let (start, length) = if let Some(selected) = selected {
            selected
        } else {
            let _parser = self.admit_query_parser(self.xml.as_bytes().len())?;
            let mut position = 0usize;
            let mut selected = None;
            scan_word_element_ranges(self.xml.as_bytes(), &[b"p"], |_, start, length| {
                if position == index {
                    selected = Some((start, length));
                }
                position = position.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("document paragraph counter overflow".into())
                })?;
                Ok(())
            })?;
            let Some(selected) = selected else {
                return Ok(None);
            };
            selected
        };
        let start = usize::try_from(start)
            .map_err(|_| Error::InvalidFormat("document paragraph offset overflow".into()))?;
        let length = usize::try_from(length)
            .map_err(|_| Error::InvalidFormat("document paragraph length overflow".into()))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| Error::InvalidFormat("document paragraph range overflow".into()))?;
        let xml = self.xml.as_bytes().get(start..end).ok_or_else(|| {
            Error::InvalidFormat("document paragraph range is outside XML".into())
        })?;
        let _parser = self.admit_query_parser(xml.len())?;
        let text = crate::paragraph::extract_word_text(xml)?;
        self.check_execution()?;
        Ok(Some(text))
    }

    /// Return the fields of the main story in document order, as
    /// [`crate::Document::fields`] does.
    ///
    /// A `w:fldSimple` without `w:instr` is kept with an empty instruction
    /// and recorded as [`crate::leniency::FieldDefect`] in
    /// [`Self::tolerance_report`].
    ///
    /// # Errors
    ///
    /// Returns an error if the field markup is malformed.
    pub fn fields(&self) -> Result<Vec<crate::Field>> {
        self.check_selective_operation("document fields")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        let (fields, repaired) = crate::Field::extract_from_document_tolerant(xml)?;
        self.merge_tolerance(&repaired);
        Ok(fields)
    }

    /// Return the bookmarks of the main story in document order, as
    /// [`crate::Document::bookmarks`] does (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error if the bookmark markup is malformed.
    pub fn bookmarks(&self) -> Result<Vec<crate::bookmark::Bookmark>> {
        self.check_selective_operation("document bookmarks")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        crate::bookmark::Bookmark::extract_from_document(xml)
    }

    /// Return the text boxes of the main story in document order, as
    /// [`crate::Document::text_boxes`] does (record 0319).
    ///
    /// # Errors
    ///
    /// Returns an error if the drawing markup is malformed.
    pub fn text_boxes(&self) -> Result<Vec<crate::textbox::TextBox>> {
        self.check_selective_operation("document text boxes")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        crate::textbox::load_text_boxes(xml)
    }

    /// Return the drawing objects of the main-story paragraphs in document
    /// order.
    ///
    /// A shape whose preset geometry is unrecognized or missing is kept with
    /// its raw token and recorded as [`crate::leniency::DrawingDefect`] in
    /// [`Self::tolerance_report`].
    ///
    /// # Errors
    ///
    /// Returns an error if a paragraph's drawing markup is malformed.
    pub fn drawing_objects(&self) -> Result<Vec<crate::drawing::Object>> {
        let paragraphs = self.paragraphs()?;
        let mut repaired = crate::leniency::ToleranceReport::default();
        let objects = crate::drawing::collect_document_objects(&paragraphs, &mut repaired);
        self.merge_tolerance(&repaired);
        objects
    }

    /// Return visible paragraphs sharing the pinned main-document allocation.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        self.check_selective_operation("document paragraphs")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        if let Some(index) = self.paragraph_ranges() {
            return document_paragraphs_from_index(Arc::clone(xml), index);
        }
        document_paragraphs(Arc::clone(xml))
    }

    /// Return visible tables sharing the pinned main-document allocation.
    pub fn tables(&self) -> Result<SmallVec<[crate::Table; 8]>> {
        self.check_selective_operation("document tables")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        document_tables(Arc::clone(xml))
    }

    /// Return one visible paragraph without allocating all paragraph views.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        self.check_selective_operation("document paragraph")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        if let Some(paragraph_index) = self.paragraph_ranges() {
            return Ok(document_paragraph_from_index(
                Arc::clone(xml),
                paragraph_index,
                index,
            ));
        }
        document_paragraph(Arc::clone(xml), index)
    }

    /// Return all visible direct body blocks in source order, including inert
    /// alternative-format anchors and unknown direct children.
    pub fn blocks(&self) -> Result<Vec<crate::Block>> {
        self.check_selective_operation("document blocks")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        document_blocks(Arc::clone(xml), self.positions)
    }

    /// The direct block children of the main document's `w:body`, as
    /// [`crate::Document::body_blocks`] indexes them.
    ///
    /// The index is built once per source revision and shared by every view
    /// of the package; this view pins it after its first body query.
    ///
    /// # Errors
    ///
    /// Refuses managed views, whose Arc-backed values could not retain the
    /// managed reservation, and returns an error when the visible XML is
    /// malformed or a bound on the index is reached.
    pub fn body_blocks(&self) -> Result<crate::document::Blocks> {
        self.check_selective_operation("document body blocks")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        if let Some(index) = self.body_memo.get() {
            return Ok(crate::document::Blocks::new(
                Arc::clone(xml),
                Arc::clone(index),
            ));
        }
        let index = self.paragraph_index_cache.body_index_source(
            self.source_version,
            xml.as_slice(),
            self.positions,
        )?;
        self.check_execution()?;
        let index = self.body_memo.get_or_init(|| index);
        Ok(crate::document::Blocks::new(
            Arc::clone(xml),
            Arc::clone(index),
        ))
    }

    /// The number of direct body blocks of every kind.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::body_blocks`].
    pub fn body_block_count(&self) -> Result<usize> {
        Ok(self.body_blocks()?.len())
    }

    /// The direct body block at zero-based `index`, counting every kind.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::body_blocks`].
    pub fn body_block(
        &self,
        index: litchi_core::Position,
    ) -> Result<Option<crate::document::IndexedBlock>> {
        Ok(self.body_blocks()?.get(index))
    }

    /// The direct body paragraph at zero-based `index`: `/body/p[index + 1]`,
    /// with `w:customXml` wrappers flattened and display equations excluded.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::body_blocks`].
    pub fn body_paragraph(&self, index: litchi_core::Position) -> Result<Option<Paragraph>> {
        Ok(self.body_blocks()?.paragraph(index))
    }

    /// The direct body table at zero-based `index`.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::body_blocks`].
    pub fn body_table(&self, index: litchi_core::Position) -> Result<Option<crate::Table>> {
        Ok(self.body_blocks()?.table(index))
    }

    /// The direct body content control (`w:sdt`) at zero-based `index`.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::body_blocks`].
    pub fn body_sdt(
        &self,
        index: litchi_core::Position,
    ) -> Result<Option<crate::document::BlockContentControl>> {
        Ok(self.body_blocks()?.content_control(index))
    }

    /// Return visible paragraph, table, and unknown elements in direct body
    /// order. Alternative-format anchors are omitted as in the eager view.
    pub fn elements(&self) -> Result<Vec<crate::Element>> {
        self.check_selective_operation("document elements")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        document_elements(Arc::clone(xml))
    }

    /// Capture the immutable section inventory from the pinned main document.
    pub fn section_inventory(&self) -> Result<crate::section::Inventory> {
        self.section_inventory_bounded(self.limits.sections_value())
    }

    /// Capture the section inventory with caller-provided semantic limits.
    pub(crate) fn section_inventory_bounded(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Inventory> {
        self.check_execution()?;
        ensure_source_section_inventory_xml(self.xml.as_bytes(), limits)?;
        let inventory = crate::section::Inventory::parse(self.xml.as_bytes(), limits)?;
        self.merge_tolerance(inventory.tolerance_report());
        self.check_execution()?;
        Ok(inventory)
    }

    /// Capture a cheaply cloneable inventory bound to the exact opened source.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_bounded(self.limits.sections_value())
    }

    /// Capture a source-bound inventory with caller-provided semantic limits.
    pub(crate) fn section_inventory_snapshot_bounded(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Snapshot> {
        self.check_execution()?;
        ensure_source_section_inventory_xml(self.xml.as_bytes(), limits)?;
        let snapshot = crate::section::Snapshot::from_source_xml(
            self.xml.as_bytes(),
            self.source_version,
            limits,
        )?;
        self.merge_tolerance(snapshot.inventory().tolerance_report());
        self.check_execution()?;
        Ok(snapshot)
    }
}

/// Whether `relationship` relates a comments part.
fn is_comments_relationship(relationship: &litchi_opc::rel::Relationship) -> bool {
    !relationship.is_external()
        && matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::COMMENTS
                | litchi_opc::constants::relationship_type::STRICT_COMMENTS
        )
}
