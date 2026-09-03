//! Lazy DOCX access and guarded main-document publication over an immutable
//! positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes. Main-document
//! transactions retain the raw XML and may be published to a sequential sink
//! while raw-copying every unselected ZIP member.

pub mod paragraph_copy;
pub mod paragraph_remove;
pub mod story_text;

pub use story_text::{
    Commit as StoryTextCommit, Edit as StoryTextEdit, Error as StoryTextError,
    Limits as StoryTextLimits, Patch as StoryTextPatch, Publication as StoryTextPublication,
    Selector as StorySelector, Snapshot as StoryTextSnapshot,
};

use crate::alt::Data;
use crate::document::{Commit, Edit, Snapshot, TransactionResult};
use crate::error::{Error, Result};
use crate::namespace::scan_word_element_ranges;
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{
    ParagraphIndex, body_block_ranges, document_blocks, document_elements, document_paragraph,
    document_paragraph_count, document_paragraph_from_index, document_paragraphs,
    document_paragraphs_from_index, document_tables, is_xml_outer_whitespace, visible_document_xml,
};
use crate::redact;
use crate::sanitize::{self, RelationshipState};
use crate::section::layout;
use crate::settings::DocumentSettings;
use crate::variables;
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{ExecutionContext, ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, PackURI, Part, PartData, SourceArtifact, SourceArtifactFingerprint,
    SourceBackedPackage, SourceCacheDiagnostics, SourceCacheLimits,
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
use std::sync::{Arc, Mutex};

/// A DOCX package that leaves ordinary part bodies cold at open.
pub struct Package {
    package: SourceBackedPackage,
    execution: Option<ExecutionContext>,
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

/// Products of one exact source-backed document-variable publication.
pub struct DocumentVariablesPublication {
    snapshot: variables::Snapshot,
    original_snapshot: variables::Snapshot,
    original_artifact: SourceArtifact,
    published_artifact: SourceArtifactFingerprint,
}

impl DocumentVariablesPublication {
    /// Borrow the published settings snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &variables::Snapshot {
        &self.snapshot
    }
}

struct FingerprintingWriter<W> {
    inner: W,
    hasher: Sha256,
}

struct SourceCheckedTextSink<'a, W: ?Sized> {
    output: &'a mut W,
    package: &'a SourceBackedPackage,
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
        let result = self
            .package
            .check_execution()
            .map_err(Error::from)
            .and_then(|_| {
                self.package
                    .source_version()
                    .map(|_| ())
                    .map_err(Error::from)
            });
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
    /// Open a DOCX package from a regular filesystem file without slurping it
    /// into memory.
    ///
    /// The path is opened once as an immutable positional [`FileSource`].
    /// Ordinary package payloads remain cold until a query asks for them.
    ///
    /// This constructor is available on platforms where [`FileSource`] is
    /// implemented.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Open a DOCX package from a regular filesystem file without slurping it
    /// into memory.
    ///
    /// This constructor is available on platforms where [`FileSource`] is
    /// implemented.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open a filesystem-backed DOCX package with explicit OPC read limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(
        path: impl AsRef<Path>,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a filesystem-backed DOCX package with an explicit finite payload
    /// cache policy.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits(file_source(path)?, cache_limits)
    }

    /// Open a filesystem-backed DOCX package with explicit OPC read and
    /// payload-cache policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: litchi_opc::ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(file_source(path)?, limits, cache_limits)
    }

    /// Open a filesystem-backed DOCX package with the standard read limits
    /// and a caller-owned execution context.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits_and_execution_context(
            file_source(path)?,
            cache_limits,
            context,
        )
    }

    /// Open a filesystem-backed DOCX package with explicit OPC read limits
    /// and a caller-owned execution context.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_execution_context(
        path: impl AsRef<Path>,
        limits: litchi_opc::ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed DOCX package with explicit OPC read and
    /// execution policies while retaining the default finite payload cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: litchi_opc::ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed DOCX package with explicit OPC read,
    /// payload-cache, and caller-owned execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: litchi_opc::ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            file_source(path)?,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open a DOCX source using the standard bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at(source)?)
    }

    /// Open a DOCX source from a sequential reader using the standard bounded
    /// OPC read policy.
    ///
    /// The reader is consumed once into the source-backed positional owner;
    /// ordinary package payloads remain deferred until a query asks for them.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_reader(reader)?)
    }

    /// Open a DOCX source from a sequential reader with explicit OPC limits.
    ///
    /// Reader ingestion is bounded by `limits.max_input_bytes()`, and
    /// ordinary package payloads remain deferred after the source is indexed.
    pub fn from_reader_with_limits<R: Read>(
        reader: R,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_reader_with_limits(
            reader, limits,
        )?)
    }

    /// Open a DOCX source with an explicit bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Open a DOCX source with an explicit finite deferred-part cache policy.
    ///
    /// This compatibility constructor remains unmanaged: cache retention is
    /// bounded, but it is not charged to a hierarchical execution budget.
    /// Use one of the `*_with_execution_context` constructors when the caller
    /// owns the budget for lazy payloads.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open a DOCX source with explicit read and finite cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
        )
    }

    /// Open a DOCX source with an explicit cache policy and caller execution
    /// context while retaining the standard OPC read limits.
    pub fn from_read_at_with_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            litchi_opc::ReadLimits::default(),
            cache_limits,
            context,
        )
    }

    /// Open a DOCX source with an explicit caller-owned execution context.
    ///
    /// The context is checked before mandatory open work and before each
    /// deferred semantic read. Lazy part payloads retain their managed
    /// [`PartData`] handles for the lifetime of the returned read facade; no
    /// executor or global scheduler is installed.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            limits,
            SourceCacheLimits::default(),
            context,
        )
    }

    /// Open a DOCX source with explicit read and execution policies while
    /// retaining the default finite deferred-part cache.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, limits, context)
    }

    /// Open a DOCX source with explicit read, cache, and execution policies.
    ///
    /// This is the fully explicit managed constructor. The selective
    /// read-only document facade retains the bounded OPC [`PartData`] handle
    /// instead of detaching it into an unbudgeted `Arc` allocation. Existing
    /// edit operations whose snapshots require an owned `Arc` remain typed
    /// refusals on this path.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        let package_context = context.clone();
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context,
            )?;
        Self::with_execution(package, Some(package_context))
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
        Ok(Self { package, execution })
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
        let package = self.package.to_opc_package()?;
        let result = crate::Package::from_opc_package(package);
        // The OPC conversion checks freshness through its final payload read,
        // but the owning DOCX constructor still performs graph/property
        // validation after that point. Re-check the original source before
        // exposing either success or a constructor error so a mutation in
        // this publication window cannot yield a stale owner (or mask the
        // stale-source failure with an eager parse error).
        self.package.source_version()?;
        self.package.check_execution()?;
        result
    }

    /// Load and pin the main document for read-only semantic queries.
    ///
    /// The first call reads only the main-document part. The returned document
    /// owns its validated visible XML bytes, so repeated text and selective
    /// queries do not revisit the positional source. Managed opens retain the
    /// budgeted [`PartData`] handle instead of detaching an `Arc`.
    pub fn document(&self) -> Result<Document> {
        self.package.check_execution()?;
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let data = main.data()?;
        let managed = self.package.cache_diagnostics().budget_managed;
        let document: Result<Document> = (|| {
            let (xml, paragraph_index) = if managed {
                ensure_source_document_xml(data.as_bytes())?;
                let paragraph_index = ParagraphIndex::from_xml(data.as_bytes()).ok().map(Arc::new);
                (DocumentPayload::Managed(data), paragraph_index)
            } else {
                let xml = visible_document_xml(data.into_arc()?)?;
                let paragraph_index = ParagraphIndex::from_xml(xml.as_slice()).ok().map(Arc::new);
                (DocumentPayload::Owned(xml), paragraph_index)
            };
            let source_version = self.package.source_version()?;
            Ok(Document {
                xml,
                paragraph_index,
                source_version,
                execution: self.execution.clone(),
            })
        })();
        // The semantic/MCE stage can fail after the payload read has
        // completed. Check freshness once more before exposing that error so
        // a source mutation during parsing wins over a stale parse result.
        self.package.source_version()?;
        self.package.check_execution()?;
        document
    }

    /// Stream visible main-document paragraphs to a caller-owned sink.
    ///
    /// The source catalog and execution policy are checked before deferred
    /// payload work. The parser checks them before and after every XML event
    /// and immediately before each paragraph emission; the source-checked
    /// sink checks both fences around every underlying write and after the
    /// final object. The main-document ZIP declaration is checked before
    /// payload materialization; the payload reader still verifies its actual
    /// decoded length and source boundary. Only one bounded paragraph text
    /// value is retained by the semantic parser.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: litchi_core::TextOutputOptions<'_>,
    ) -> std::result::Result<litchi_core::TextOutputReport, litchi_core::TextOutputError<Error>>
    {
        let source_failure = Arc::new(Mutex::new(None));
        let mut checked_output = SourceCheckedTextSink {
            output,
            package: &self.package,
            failure: Arc::clone(&source_failure),
        };
        let mut writer = litchi_core::SequentialTextWriter::new(&mut checked_output, options);

        let parsed = (|| -> std::result::Result<(), litchi_core::TextOutputError<Error>> {
            self.package
                .check_execution()
                .map_err(|source| writer.document_error(source.into()))?;
            self.package
                .source_version()
                .map_err(|source| writer.document_error(source.into()))?;
            let main = self
                .package
                .main_document_part()
                .map_err(|source| writer.document_error(source.into()))?;
            validate_document_main_content_type(main.content_type())
                .map_err(|source| writer.document_error(source))?;
            let declared = main
                .declared_uncompressed_size()
                .map_err(|source| writer.document_error(source.into()))?;
            let limit =
                u64::try_from(crate::paragraph::semantic_text_raw_xml_limit()).map_err(|_| {
                    writer.document_error(Error::InvalidFormat(
                        "semantic DOCX XML limit overflow".into(),
                    ))
                })?;
            let mce_limit = crate::paragraph::semantic_text_raw_xml_limit();
            if declared > limit {
                return Err(writer.document_error(Error::InvalidFormat(format!(
                    "semantic DOCX declared XML exceeds {limit} bytes"
                ))));
            }
            let data = main
                .data()
                .map_err(|source| writer.document_error(source.into()))?;
            let managed = self.package.cache_diagnostics().budget_managed;
            let visible: Cow<'_, [u8]> = if managed {
                ensure_source_document_xml(data.as_bytes())
                    .map_err(|source| writer.document_error(source))?;
                Cow::Borrowed(data.as_bytes())
            } else {
                let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
                capabilities
                    .understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
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
            self.package
                .source_version()
                .map_err(|source| writer.document_error(source.into()))?;
            self.package
                .check_execution()
                .map_err(|source| writer.document_error(source.into()))?;
            crate::paragraph::write_text_to_with_operation_check(
                visible.as_ref(),
                &mut writer,
                || {
                    self.package.check_execution().map_err(Error::from)?;
                    self.package.source_version().map_err(Error::from)?;
                    Ok(())
                },
            )
        })();
        let progress = writer.progress();
        let source = self
            .package
            .check_execution()
            .err()
            .map(Error::from)
            .or_else(|| self.package.source_version().err().map(Error::from))
            .or_else(|| take_source_text_failure(&source_failure));
        if let Some(source) = source {
            return Err(litchi_core::TextOutputError::Document { source, progress });
        }
        parsed.map(|()| writer.finish())
    }

    /// Load only the mandatory main-document payload and capture its immutable
    /// section inventory at the exact opened source version.
    ///
    /// Header/footer relationship IDs are retained as inert values. Their
    /// target parts are neither resolved nor read.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_with_limits(&crate::section::Limits::default())
    }

    /// Capture a source-bound section inventory with explicit semantic limits.
    pub fn section_inventory_snapshot_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Snapshot> {
        self.package.check_execution()?;
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let source_version = self.package.source_version()?;
        let raw = main.data()?;
        let snapshot = (|| {
            ensure_source_section_inventory_xml(raw.as_bytes(), limits)?;
            crate::section::Snapshot::from_source_xml(raw.as_bytes(), source_version, limits)
        })();
        // Detect hostile adapters that mutate immediately after the payload
        // read and semantic scan, before publishing the closure to the caller.
        let observed = self.package.source_version()?;
        if observed != source_version {
            return Err(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            }
            .into());
        }
        self.package.check_execution()?;
        snapshot
    }

    /// Capture an exact-source snapshot for editing locally authored page
    /// layout on one existing main-story section.
    pub fn section_layout_snapshot(&self) -> Result<layout::Snapshot> {
        self.section_layout_snapshot_with_limits(&crate::section::Limits::default())
    }

    /// Capture a bounded exact-source page-layout snapshot.
    pub fn section_layout_snapshot_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<layout::Snapshot> {
        let (_, snapshot) = self.main_section_layout_snapshot("section_layout_snapshot", limits)?;
        Ok(snapshot)
    }

    /// Start an isolated edit of one existing locally authored section.
    pub fn edit_section_layout(
        &self,
        selector: impl Into<crate::section::Selector>,
    ) -> Result<layout::Edit> {
        self.section_layout_snapshot()?.edit(selector)
    }

    /// Start an isolated bounded edit of one existing locally authored
    /// section.
    pub fn edit_section_layout_with_limits(
        &self,
        selector: impl Into<crate::section::Selector>,
        limits: &crate::section::Limits,
    ) -> Result<layout::Edit> {
        self.section_layout_snapshot_with_limits(limits)?
            .edit(selector)
    }

    /// Return content-free cache activity for this source-backed DOCX.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Return the exact source identity and revision captured at open.
    pub fn source_version(&self) -> Result<SourceVersion> {
        Ok(self.package.source_version()?)
    }

    /// Read relationship-selected OOXML core properties without loading the
    /// main document or unrelated package members.
    ///
    /// A missing core-properties relationship yields empty metadata. Present
    /// properties retain the shared strict graph and dialect validation used
    /// by eager OOXML readers. Source freshness is checked before and after
    /// the selected payload read by the common source-backed property reader;
    /// execution cancellation is checked at the same boundaries.
    pub fn metadata(&self) -> Result<litchi_core::Metadata> {
        self.package.check_execution()?;
        let properties = litchi_ooxml_common::properties::read_source_backed(&self.package);
        self.package.source_version()?;
        self.package.check_execution()?;
        Ok(properties?.map(Into::into).unwrap_or_default())
    }

    /// Inventory embedded-object and embedded-package relationships without
    /// materializing the OPC package or any payload bytes.
    ///
    /// The returned catalog entries borrow this source-backed package. Their
    /// payload views remain deferred; reading a payload is an explicit,
    /// fallible operation on the returned entry rather than part of this
    /// inventory call.
    pub fn embedded(&self) -> Result<Vec<litchi_ooxml_common::embedded::SourceEntry<'_>>> {
        Ok(litchi_ooxml_common::embedded::scan_source(&self.package)?)
    }

    /// Inventory embedded relationships with explicit resource limits.
    ///
    /// This is a catalog-only operation over the retained source-backed OPC
    /// package. Payload bytes are not read or materialized; callers must use
    /// the deferred fallible payload operation on each returned entry.
    pub fn embedded_with_limits(
        &self,
        limits: &litchi_ooxml_common::embedded::Limits,
    ) -> Result<Vec<litchi_ooxml_common::embedded::SourceEntry<'_>>> {
        Ok(litchi_ooxml_common::embedded::scan_source_with(
            &self.package,
            limits,
        )?)
    }

    /// Load the exact raw main-document bytes as a semantic transaction
    /// snapshot.
    ///
    /// Source-backed edits currently refuse documents whose markup-
    /// compatibility preprocessing selects or rewrites branches. Keeping the
    /// raw bytes prevents semantic selectors from being applied to a different
    /// XML representation than the one eventually published.
    ///
    /// # Errors
    ///
    /// Returns a typed package, document, resource-limit, or unsafe-edit error.
    pub fn document_snapshot(&self) -> TransactionResult<Snapshot> {
        let (_, snapshot) = self.main_document_snapshot("document_snapshot")?;
        Ok(snapshot)
    }

    /// Start an isolated semantic edit over the exact raw main document.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::document_snapshot`].
    pub fn edit_document(&self) -> TransactionResult<Edit> {
        Ok(self.document_snapshot()?.edit())
    }

    /// Capture document variables from the existing internal settings Part.
    ///
    /// This source-backed capability never creates a settings relationship or
    /// Part. It accepts exactly one Strict or Transitional internal settings
    /// relationship with the Word settings content type. Markup-compatibility
    /// branch selection is refused so the semantic snapshot and publish bytes
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

    /// Start an isolated edit of variables in the existing settings Part.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::document_variables_snapshot`].
    pub fn edit_document_variables(&self) -> Result<variables::Transaction> {
        self.document_variables_snapshot()?.try_edit()
    }

    /// Publish a source-checked document-variable commit to a sequential sink.
    ///
    /// Only the existing settings payload may change. Every other physical ZIP
    /// member is raw-copied. Exact no-ops reproduce the complete source artifact
    /// byte for byte, including signatures and protected settings. A real change
    /// refuses signed, protected, MCE-selected, or unmodeled selected-variable
    /// sources before output begins.
    ///
    /// # Errors
    ///
    /// Returns a typed source-conflict, relationship, content-type, protection,
    /// preservation, signature, limit, XML-publication, or incomplete-output
    /// error.
    pub fn publish_document_variables_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &variables::Commit,
    ) -> Result<DocumentVariablesPublication> {
        let current =
            self.settings_document_variables_source("publish_document_variables_commit_to_stream")?;
        let target = commit.patch().apply(&current.snapshot)?;
        if !commit.patch().is_empty() {
            if !current.unique_inbound_owner {
                return Err(Error::DocumentVariablesPreservation(
                    "the settings Part has an additional internal inbound relationship",
                ));
            }
            if current.protected {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "publish_document_variables_commit_to_stream",
                    reason: "document or write protection is enforced",
                });
            }
            variables::ensure_source_backed_rewrite_safe(
                current.snapshot.xml_bytes(),
                current.snapshot.variables(),
            )?;
        }
        let original_artifact = self.package.source_artifact();
        let mut output = FingerprintingWriter {
            inner: writer,
            hasher: Sha256::new(),
        };
        if commit.patch().is_empty() {
            self.package
                .write_part_overlays_shared_to_stream(&mut output, Vec::new())?;
        } else {
            self.package.write_part_overlay_shared_to_stream(
                &mut output,
                &current.partname,
                target.shared_xml(),
            )?;
        }
        let published_artifact =
            SourceArtifactFingerprint::from_sha256(output.hasher.finalize().into());
        Ok(DocumentVariablesPublication {
            snapshot: target,
            original_snapshot: commit.patch().before_snapshot().clone(),
            original_artifact,
            published_artifact,
        })
    }

    /// Apply the exact inverse of a completed source-backed publication.
    ///
    /// The supplied package must be the byte-exact artifact emitted by
    /// `publication`; a reopened foreign or subsequently modified artifact is
    /// rejected before output begins. Successful publication copies the exact
    /// retained original artifact, including all untouched physical ZIP bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed conflict, source-change, allocation, or incomplete-
    /// output error.
    pub fn publish_document_variables_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &DocumentVariablesPublication,
    ) -> Result<variables::Snapshot> {
        let current = self
            .settings_document_variables_source("publish_document_variables_inverse_to_stream")?;
        if !current.snapshot.same_content(&publication.snapshot)
            || self.package.source_artifact().fingerprint()? != publication.published_artifact
        {
            return Err(Error::DocumentVariablesConflict);
        }
        publication.original_artifact.write_to_stream(writer)?;
        Ok(publication.original_snapshot.clone())
    }

    /// Publish one exact-source-checked existing-section layout commit to a
    /// sequential stream. Only the existing main-document Part is overlaid;
    /// all other ZIP members are copied through the source-backed preservation
    /// plan.
    pub fn publish_section_layout_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &layout::Commit,
    ) -> Result<layout::Publication> {
        self.publish_section_layout_patch_to_stream(writer, commit.patch())
    }

    /// Publish an exact-source-checked section-layout patch to a sequential
    /// stream. The returned publication retains an inverse authorization for
    /// the exact emitted artifact.
    pub fn publish_section_layout_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &layout::Patch,
    ) -> Result<layout::Publication> {
        let (main, current) = self.main_section_layout_snapshot(
            "publish_section_layout_patch_to_stream",
            patch.limits(),
        )?;
        let target = patch.apply(&current)?;
        let original_artifact = self.package.source_artifact();
        let mut output = FingerprintingWriter {
            inner: writer,
            hasher: Sha256::new(),
        };
        if patch.is_noop() {
            original_artifact.write_to_stream(&mut output)?;
        } else {
            self.package.write_part_overlay_shared_to_stream(
                &mut output,
                &main,
                target.shared_xml(),
            )?;
        }
        let published_fingerprint =
            SourceArtifactFingerprint::from_sha256(output.hasher.finalize().into());
        let mut inverse_patch = patch.inverse();
        inverse_patch.reauthorize_for_artifact(published_fingerprint);
        Ok(layout::Publication::new(
            target.with_artifact_fingerprint(published_fingerprint),
            current,
            original_artifact,
            inverse_patch,
        ))
    }

    /// Restore the exact original package from a section-layout publication.
    ///
    /// This is the explicit reauthorization boundary for a reopened emitted
    /// artifact. The complete emitted ZIP fingerprint and target main-story
    /// bytes are checked before the sink receives any output; a foreign or
    /// stale package is rejected without partial output.
    pub fn publish_section_layout_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &layout::Publication,
    ) -> Result<layout::Snapshot> {
        let (_, current) = self.main_section_layout_snapshot(
            "publish_section_layout_inverse_to_stream",
            publication.snapshot().limits(),
        )?;
        publication.inverse_patch().apply(&current)?;
        publication
            .original_artifact()
            .write_to_stream(writer)
            .map_err(Error::from)?;
        Ok(publication.original_snapshot().clone())
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
        self.external_hyperlink_sanitization_snapshot_with_limits(sanitize::Limits::default())
    }

    /// Capture the external-hyperlink sanitization closure with explicit
    /// semantic scanner limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot`].
    pub fn external_hyperlink_sanitization_snapshot_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<sanitize::Snapshot> {
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_sanitization_snapshot")
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        let main = self.package.main_document_part()?;
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
        sanitize::Snapshot::from_source(document.shared_xml(), relationships, limits)
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

    /// Build the external-hyperlink detachment plan with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot_with_limits`].
    pub fn plan_external_hyperlink_detachment_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<sanitize::SanitizePlan> {
        Ok(self
            .external_hyperlink_sanitization_snapshot_with_limits(limits)?
            .plan())
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
        self.external_hyperlink_redaction_snapshot_with_limits(sanitize::Limits::default())
    }

    /// Capture an irreversible redaction inventory with explicit XML limits.
    pub fn external_hyperlink_redaction_snapshot_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<redact::Snapshot> {
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_redaction_snapshot")
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        self.refuse_external_hyperlink_redaction_topology()?;
        let main = self.package.main_document_part()?;
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
        let source_version = self.package.source_version()?;
        let source_fingerprint = self.package.source_artifact().fingerprint()?;
        redact::Snapshot::from_source(
            document.shared_xml(),
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

    /// Publish one exact-source-checked main-document commit to a sequential
    /// stream while preserving every other physical ZIP member.
    ///
    /// Only operations confined to the main-document payload are accepted.
    /// Cross-package paragraph transfers are refused because their dependency
    /// graph requires package-level publication. A no-op commit reproduces the
    /// complete source artifact byte for byte. A changed signed package is
    /// refused by the underlying OPC publisher.
    ///
    /// All semantic, source-version, topology, signature, and replacement-XML
    /// checks happen before output. A sink failure after output begins is
    /// reported through the underlying typed incomplete-output error.
    ///
    /// # Errors
    ///
    /// Returns a typed transaction, package, unsafe-edit, signature, source,
    /// XML-publication, or sink error.
    pub fn publish_document_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> TransactionResult<Snapshot> {
        let (main, current) = self.main_document_snapshot("publish_document_commit_to_stream")?;
        let target = commit.patch().apply(&current)?;
        if commit
            .patch()
            .operations()
            .iter()
            .any(|operation| !operation.supports_source_backed_main_document_overlay())
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_document_commit_to_stream",
                reason: "paragraph transfers require package-level dependency publication",
            }
            .into());
        }
        // Semantic operations are bookkeeping, not the publication
        // authorization. An edit may temporarily change a document and then
        // restore the exact source bytes before commit. Only the patch's
        // byte-identity check is allowed to select the exact-source path;
        // this also avoids asking a rewritten target to stand in for an
        // unsupported source payload.
        if !commit.patch().changed() {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())
                .map_err(Error::from)?;
        } else {
            self.package
                .write_part_overlay_shared_to_stream(writer, &main, target.shared_xml())
                .map_err(Error::from)?;
        }
        Ok(target)
    }

    /// Replace the opaque payload behind one existing main-story altChunk.
    ///
    /// This is an opt-in one-edit save path for foreign HTML, RTF, text, XML,
    /// or nested-Office payloads. The selected relationship, Part URI, media
    /// type, and relationship closure are immutable; only the existing target
    /// payload is replaced. Every other ZIP member is copied through the
    /// source-backed OPC preservation plan, including unsupported members and
    /// the source's local ZIP framing.
    ///
    /// A positional selector resolves only active main-story anchors. An
    /// external relationship, an ambiguous/shared target, a target with its
    /// own relationships, a non-`/word/` direct-child target, a media-type or
    /// extension mismatch, and other unsafe layouts are refused before output.
    /// Exact payload no-ops preserve the complete source artifact, including
    /// signatures. A real change to signed input is refused until the caller
    /// chooses an explicit signature-stripping or resigning workflow.
    ///
    /// The payload is inert: it is never imported, parsed as a nested Office
    /// document, rendered, fetched, or executed.
    ///
    /// # Errors
    ///
    /// Returns a typed source, selector, relationship, target-layout,
    /// signature, XML-publication, or sink error. A sink failure after output
    /// begins is reported by the underlying incomplete-output error.
    pub fn publish_alt_chunk_to_stream<W: Write>(
        self,
        writer: W,
        selector: impl Into<AltChunkSelector>,
        replacement: Data,
    ) -> Result<()> {
        replacement.validate()?;
        let target = self.alt_chunk_target(
            selector.into(),
            replacement.media_type(),
            replacement.extension(),
        )?;
        self.package
            .write_part_overlay_to_stream(writer, &target, replacement.into_bytes())
            .map_err(Error::from)
    }

    /// Publish an explicit external-hyperlink sanitization commit to a
    /// sequential stream.
    ///
    /// Only the main-document payload is regenerated. Every other physical
    /// member and the relationship topology are preserved. Consequently the
    /// detached hyperlinks' external relationship records remain present and
    /// are counted in [`sanitize::EffectReport`]. An exact no-op reproduces
    /// every source byte, including signatures; a real change to a signed
    /// package is refused because this API has no resigning policy.
    ///
    /// All source, protection, patch, XML, signature, and preservation checks
    /// complete before output. A later sink failure is reported as the
    /// underlying typed incomplete-output error.
    ///
    /// # Errors
    ///
    /// Returns a typed package, source-conflict, protection, signature, XML,
    /// resource-limit, or sink error.
    pub fn publish_external_hyperlink_sanitization_to_stream<W: Write>(
        self,
        writer: W,
        commit: &sanitize::Commit,
    ) -> Result<sanitize::Snapshot> {
        let main = self.package.main_document_part()?.partname().clone();
        let current =
            self.external_hyperlink_sanitization_snapshot_with_limits(commit.patch().limits())?;
        let target = commit.patch().apply(&current)?;
        if commit.patch().is_noop() {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        } else {
            self.package
                .write_part_overlay_shared_to_stream(writer, &main, target.shared_xml())?;
        }
        Ok(target)
    }

    /// Irreversibly publish a source-checked external-hyperlink redaction.
    ///
    /// Selected wrappers are unwrapped, selected external relationship records
    /// are removed, visible children are retained, and every other ZIP member
    /// is raw-copied. This operation consumes the package and exposes no inverse
    /// API. It is never called by ordinary save or detachment operations.
    pub fn publish_external_hyperlink_redaction_to_stream<W: Write>(
        self,
        writer: W,
        commit: &redact::Commit,
    ) -> Result<redact::EffectReport> {
        let main = self.package.main_document_part()?.partname().clone();
        let current =
            self.external_hyperlink_redaction_snapshot_with_limits(commit.patch().limits())?;
        commit.patch().validate_source(&current)?;
        let report = commit.effect_report();
        if report.is_noop() {
            self.package.source_artifact().write_to_stream(writer)?;
            return Ok(report);
        }
        let (replacement, removed_ids) = redact::publication_parts(commit);
        self.package
            .write_part_overlay_with_external_relationship_removals_to_stream(
                writer,
                &main,
                replacement,
                removed_ids,
            )?;
        Ok(report)
    }

    /// Capture the package-wide, relationship-owned story hyperlink closure.
    ///
    /// The inventory covers the main, header, footer, footnote, endnote,
    /// comments, and glossary stories. It does not alter the established
    /// main-only sanitization or redaction APIs above.
    pub fn story_hyperlinks_only_snapshot(&self) -> Result<crate::story_hyperlinks::Snapshot> {
        self.story_hyperlinks_only_snapshot_with_limits(crate::story_hyperlinks::Limits::default())
    }

    /// Capture the story hyperlink closure under an explicit bounded policy.
    pub fn story_hyperlinks_only_snapshot_with_limits(
        &self,
        limits: crate::story_hyperlinks::Limits,
    ) -> Result<crate::story_hyperlinks::Snapshot> {
        self.package.check_execution()?;
        crate::story_hyperlinks::capture_source(&self.package, limits)
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
        self.plan_story_hyperlink_redaction_with_limits(
            target_urls,
            mode,
            crate::story_hyperlinks::Limits::default(),
        )
    }

    /// Plan story hyperlink removal with explicit semantic bounds.
    pub fn plan_story_hyperlink_redaction_with_limits(
        &self,
        target_urls: &[&str],
        mode: crate::story_hyperlinks::Mode,
        limits: crate::story_hyperlinks::Limits,
    ) -> Result<crate::story_hyperlinks::Plan> {
        self.story_hyperlinks_only_snapshot_with_limits(limits)?
            .plan_target_urls_with_mode(target_urls, mode)
    }

    /// Publish a sealed forward-only story hyperlink redaction.
    ///
    /// Selected story XML and `.rels` members are regenerated as one bounded
    /// preservation plan. Every untouched ZIP member is raw-copied. Exact
    /// no-ops copy the source byte-for-byte; changed signed/protected or stale
    /// sources are refused before output begins.
    pub fn publish_story_hyperlink_redaction_to_stream<W: Write>(
        self,
        writer: W,
        commit: &crate::story_hyperlinks::Commit,
    ) -> Result<crate::story_hyperlinks::Report> {
        self.package.check_execution()?;
        commit.patch().validate_source_identity(&self.package)?;
        let report = commit.report().clone();
        if report.effect().is_noop() {
            self.package.source_artifact().write_to_stream(writer)?;
            return Ok(report);
        }
        let parts = crate::story_hyperlinks::publication_parts(commit)?;
        self.package
            .write_part_overlays_with_external_relationship_removals_to_stream(writer, parts)?;
        Ok(report)
    }

    fn alt_chunk_target(
        &self,
        selector: AltChunkSelector,
        media_type: &str,
        extension: &str,
    ) -> Result<PackURI> {
        self.package.check_execution()?;
        let main = self.package.main_document_part()?;
        let main_name = main.partname().clone();
        let data = main.data()?;
        let chunks = crate::alt::scan(data.as_bytes())?;
        // `body_block_ranges` uses this stable private kind ordering for the
        // third variant, the direct-body altChunk block.
        const ALT_CHUNK_BLOCK_KIND: usize = 2;
        let body_ranges = body_block_ranges(data.as_bytes())?;
        let body_len = body_ranges
            .iter()
            .filter(|(kind, start, _)| *kind == ALT_CHUNK_BLOCK_KIND && chunks.contains_key(start))
            .count();
        let AltChunkSelector::Index(index) = selector;
        let chunk = body_ranges
            .into_iter()
            .filter_map(|(kind, start, _)| {
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
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "external altChunk targets cannot be replaced by a source-backed payload overlay",
            });
        }
        let target = relationship.target_partname()?;
        let target_part = self.package.part(&target)?;
        if target == main_name {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "an altChunk target cannot be the main document Part",
            });
        }
        let target_kind = crate::alt::Kind::from_media_type(target_part.content_type());
        let replacement_kind = crate::alt::Kind::from_media_type(media_type);
        if target_kind != replacement_kind || matches!(target_kind, crate::alt::Kind::Unknown) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "replacement media type family must match the existing altChunk target Part",
            });
        }
        if !target_part.rels().is_empty() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "altChunk targets with dependent relationships are outside the safe payload-only closure",
            });
        }
        let relative = target
            .as_str()
            .strip_prefix("/word/")
            .ok_or(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "altChunk target must be a direct child of the Word package directory",
            })?;
        let Some((stem, target_extension)) = relative.rsplit_once('.') else {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "altChunk target must have a bounded foreign-payload extension",
            });
        };
        if stem.is_empty()
            || relative.contains('/')
            || target_extension.is_empty()
            || !target_extension.eq_ignore_ascii_case(extension)
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "altChunk target path or extension is outside the safe payload-only layout",
            });
        }

        let mut inbound = 0usize;
        for candidate in self.package.rels().iter() {
            if !candidate.is_external() && candidate.target_partname()? == target {
                inbound = inbound.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("altChunk inbound count overflow".into())
                })?;
            }
        }
        for part in self.package.iter_parts() {
            for candidate in part.rels().iter() {
                if !candidate.is_external() && candidate.target_partname()? == target {
                    inbound = inbound.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("altChunk inbound count overflow".into())
                    })?;
                }
            }
        }
        if inbound != 1 {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_alt_chunk_to_stream",
                reason: "the selected altChunk target is shared by an ambiguous relationship closure",
            });
        }
        self.package.source_version()?;
        Ok(target)
    }

    fn main_document_snapshot(
        &self,
        operation: &'static str,
    ) -> TransactionResult<(PackURI, Snapshot)> {
        self.package.check_execution().map_err(Error::from)?;
        let main = self.package.main_document_part().map_err(Error::from)?;
        validate_document_main_content_type(main.content_type())?;
        let partname = main.partname().clone();
        if self.package.cache_diagnostics().budget_managed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "managed source-backed document transactions require an owned edit snapshot; use the selective read facade or an unmanaged compatibility constructor",
            }
            .into());
        }
        let data = main.data().map_err(Error::from)?;
        let raw = data.into_arc().map_err(Error::from)?;
        let snapshot: TransactionResult<Snapshot> = (|| {
            let visible = visible_document_xml(Arc::clone(&raw))?;
            if !Arc::ptr_eq(&raw, &visible) {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation,
                    reason: "source-backed document transactions do not support markup-compatibility branch selection",
                }
                .into());
            }
            Snapshot::from_shared_xml(raw)
        })();
        self.package.source_version().map_err(Error::from)?;
        self.package.check_execution().map_err(Error::from)?;
        Ok((partname, snapshot?))
    }

    fn main_section_layout_snapshot(
        &self,
        operation: &'static str,
        limits: &crate::section::Limits,
    ) -> Result<(PackURI, layout::Snapshot)> {
        self.package.check_execution()?;
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let partname = main.partname().clone();
        if self.package.cache_diagnostics().budget_managed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "source-backed section layout edits require an owned immutable main-document snapshot",
            });
        }
        let source_version = self.package.source_version()?;
        let lineage = self.package.source_lineage();
        let data = main.data()?;
        let raw = data.into_arc()?;
        let artifact_fingerprint = self.package.source_artifact().fingerprint()?;
        let snapshot = layout::Snapshot::from_source_xml(
            Arc::clone(&raw),
            source_version,
            lineage.clone(),
            artifact_fingerprint,
            limits,
        );
        let observed = self.package.source_version()?;
        if observed != source_version {
            return Err(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            }
            .into());
        }
        self.package.check_execution()?;
        Ok((partname, snapshot?))
    }

    fn settings_document_variables_source(
        &self,
        operation: &'static str,
    ) -> Result<DocumentVariablesSource> {
        self.package.check_execution()?;
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        let main = self.package.main_document_part()?;
        let package_strict = self.package.rels().iter().find_map(|relationship| {
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
        for candidate in self.package.rels().iter() {
            if !candidate.is_external() && candidate.target_partname()? == target {
                inbound_owners = inbound_owners.saturating_add(1);
            }
        }
        for part in self.package.iter_parts() {
            for candidate in part.rels().iter() {
                if !candidate.is_external() && candidate.target_partname()? == target {
                    inbound_owners = inbound_owners.saturating_add(1);
                }
            }
        }
        let settings_part = self.package.part(&target)?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::InvalidContentType {
                expected: ct::WML_SETTINGS.into(),
                got: settings_part.content_type().into(),
            });
        }
        if self.package.cache_diagnostics().budget_managed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "managed source-backed settings snapshots require an owned edit snapshot; use an unmanaged compatibility constructor",
            });
        }
        let data = settings_part.data()?;
        let raw = data.into_arc()?;
        if raw.len() > variables::MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(Error::InvalidFormat(format!(
                "settings XML exceeds the {} byte document-variable limit",
                variables::MAX_DOCUMENT_VARIABLE_XML_BYTES
            )));
        }
        let staged = BlobPart::new_shared(
            target.clone(),
            ct::WML_SETTINGS.to_owned(),
            Arc::clone(&raw),
        );
        let visible = litchi_ooxml_common::mce::process_part_arc(&staged)?;
        if !Arc::ptr_eq(&raw, &visible) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "source-backed document-variable transactions do not support markup-compatibility branch selection",
            });
        }
        let policy = variables::inspect_source_policy(raw.as_slice())?;
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
            snapshot: variables::Snapshot::from_source_xml(raw, self.package.source_version()?)?,
            protected: policy.protected,
            unique_inbound_owner: inbound_owners == 1,
        })
    }

    fn refuse_protected_external_hyperlink_detachment(&self) -> Result<()> {
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        if self.package.cache_diagnostics().budget_managed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "external_hyperlink_wrapper_detachment",
                reason: "managed source-backed hyperlink sanitization requires an owned settings snapshot; use an unmanaged compatibility constructor",
            });
        }

        let main = self.package.main_document_part()?;
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
        let settings_part = self.package.part(&target)?;
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
        let settings = DocumentSettings::extract_from_part(&staged)?;
        if settings.is_protected() || settings.is_write_protected() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "external_hyperlink_wrapper_detachment",
                reason: "document or write protection is enforced",
            });
        }
        Ok(())
    }

    fn refuse_external_hyperlink_redaction_topology(&self) -> Result<()> {
        const SIGNATURE_TYPES: &[&str] = &[
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
        ];
        let main = self.package.main_document_part()?;
        if self
            .package
            .rels()
            .iter()
            .any(|relationship| SIGNATURE_TYPES.contains(&relationship.reltype()))
        {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if self
            .package
            .rels()
            .iter()
            .any(|relationship| relationship.is_external())
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "external_hyperlink_redaction",
                reason: "a package-level relationship has an external target outside the redaction closure",
            });
        }
        for part in self.package.iter_parts() {
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
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "external_hyperlink_redaction",
                    reason: "a non-main-document part owns an external hyperlink relationship",
                });
            }
        }
        Ok(())
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn transaction_error_to_document(error: crate::document::TransactionError) -> Error {
    match error {
        crate::document::TransactionError::Document(error) => error,
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

fn ensure_source_document_xml(xml: &[u8]) -> Result<()> {
    if xml.len() > SOURCE_DOCUMENT_SCAN_MAX_BYTES {
        return Err(Error::InvalidFormat(
            "source-backed document XML exceeds the bounded MCE scan byte limit".into(),
        ));
    }

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
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "source-backed document read",
                    reason: "processing instructions are not accepted in a managed source-backed document",
                });
            },
            Event::DocType(_) => {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "source-backed document read",
                    reason: "DTD declarations are not accepted in a managed source-backed document",
                });
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
                        return Err(Error::UnsafeEdit {
                            format: "DOCX",
                            operation: "source-backed document read",
                            reason: "non-predefined entity references are not accepted in a managed source-backed document",
                        });
                    }
                }
            },
        }
    }
    Ok(())
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
            resource: "input bytes",
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
                resource: "XML events",
                maximum: limits.max_events,
                actual: events,
            });
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
                        resource: "XML depth",
                        maximum: limits.max_depth,
                        actual: depth,
                    });
                }
                validate_source_section_element(&resolver, decoder, &namespace, &element)?;
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
                        resource: "XML depth",
                        maximum: limits.max_depth,
                        actual: element_depth,
                    });
                }
                validate_source_section_element(&resolver, decoder, &namespace, &element)?;
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
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "source-backed section inventory",
                    reason: "DTD declarations are not accepted in a source-bound inventory",
                });
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
                        return Err(Error::UnsafeEdit {
                            format: "DOCX",
                            operation: "source-backed section inventory",
                            reason: "non-predefined entity references are not accepted in a source-bound inventory",
                        });
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

    for attribute in element.attributes() {
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
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: if operation == "document" {
                "source-backed document read"
            } else {
                "source-backed section inventory"
            },
            reason,
        });
    }

    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        validate_source_attribute_value(
            attribute.value.as_ref(),
            if operation == "document" {
                "source-backed document read"
            } else {
                "source-backed section inventory"
            },
        )?;
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
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: if operation == "document" {
                    "source-backed document read"
                } else {
                    "source-backed section inventory"
                },
                reason: if operation == "document" {
                    "markup-compatibility preprocessing would require an unbudgeted owned payload"
                } else {
                    "markup-compatibility namespace bindings are not accepted in a source-bound inventory"
                },
            });
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
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "non-predefined entity references are not accepted in a source-backed XML namespace binding",
            });
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
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "non-predefined entity references are not accepted in a source-bound attribute",
            });
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
enum DocumentPayload {
    /// Compatibility payload whose Arc ownership is not attached to a
    /// hierarchical execution budget.
    Owned(Arc<Vec<u8>>),
    /// Managed payload retaining its OPC reservation for the document view.
    Managed(PartData),
}

impl DocumentPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(xml) => xml.as_slice(),
            Self::Managed(data) => data.as_bytes(),
        }
    }

    const fn is_managed(&self) -> bool {
        matches!(self, Self::Managed(_))
    }
}

/// A pinned read-only view of a DOCX main document.
///
/// This view owns the main-document bytes loaded by [`Package::document`].
/// It intentionally exposes semantic text and paragraph queries only; use
/// the established [`crate::Package`] APIs for mutable package access.
#[derive(Clone)]
pub struct Document {
    xml: DocumentPayload,
    /// Bounded offsets into the pinned XML. The index owns no payload bytes;
    /// managed documents therefore retain the same `PartData` ownership and
    /// budget reservations as the uncached selective facade.
    paragraph_index: Option<Arc<ParagraphIndex>>,
    source_version: SourceVersion,
    execution: Option<ExecutionContext>,
}

impl Document {
    fn check_execution(&self) -> Result<()> {
        let Some(context) = self.execution.as_ref() else {
            return Ok(());
        };
        context.check().map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                other => litchi_opc::OpcError::Execution(other),
            })
        })
    }

    fn check_selective_operation(&self, operation: &'static str) -> Result<()> {
        self.check_execution()?;
        if self.xml.is_managed() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "this query would return Arc-backed semantic views that cannot retain the managed PartData reservation",
            });
        }
        Ok(())
    }

    /// Return the exact source identity and revision captured for this view.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Extract all visible paragraph text from the pinned document.
    pub fn extract_text(&self) -> Result<String> {
        self.check_execution()?;
        crate::paragraph::extract_word_text(self.xml.as_bytes())
    }

    /// Count visible paragraphs in the pinned document.
    pub fn paragraph_count(&self) -> Result<usize> {
        self.check_execution()?;
        if let Some(index) = self.paragraph_index.as_deref() {
            return Ok(index.len());
        }
        document_paragraph_count(self.xml.as_bytes())
    }

    /// Extract one direct body paragraph's visible text without constructing
    /// an Arc-backed [`Paragraph`] view. This is available for managed
    /// source-backed documents because the returned `String` is caller-owned
    /// semantic output, not a hidden payload alias.
    pub fn paragraph_text(&self, index: usize) -> Result<Option<String>> {
        self.check_execution()?;
        let selected = self.paragraph_index.as_deref().and_then(|paragraph_index| {
            paragraph_index
                .get(index)
                .map(|range| (range.start, range.length))
        });
        let (start, length) = if let Some(selected) = selected {
            selected
        } else {
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
        let text = crate::paragraph::extract_word_text(xml)?;
        self.check_execution()?;
        Ok(Some(text))
    }

    /// Return visible paragraphs sharing the pinned main-document allocation.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        self.check_selective_operation("document paragraphs")?;
        let DocumentPayload::Owned(xml) = &self.xml else {
            unreachable!("managed document payload rejected above")
        };
        if let Some(index) = self.paragraph_index.as_deref() {
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
        if let Some(paragraph_index) = self.paragraph_index.as_deref() {
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
        document_blocks(Arc::clone(xml))
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
        self.section_inventory_with_limits(&crate::section::Limits::default())
    }

    /// Capture the section inventory with caller-provided semantic limits.
    pub fn section_inventory_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Inventory> {
        self.check_execution()?;
        ensure_source_section_inventory_xml(self.xml.as_bytes(), limits)?;
        let inventory = crate::section::Inventory::parse_with_limits(self.xml.as_bytes(), limits)?;
        self.check_execution()?;
        Ok(inventory)
    }

    /// Capture a cheaply cloneable inventory bound to the exact opened source.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_with_limits(&crate::section::Limits::default())
    }

    /// Capture a source-bound inventory with caller-provided semantic limits.
    pub fn section_inventory_snapshot_with_limits(
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
        self.check_execution()?;
        Ok(snapshot)
    }
}
