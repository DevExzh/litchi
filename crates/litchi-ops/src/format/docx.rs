//! Handle for Word OOXML (`.docx`) inputs.
//!
//! [`Docx`] is one `Arc` over its state: the two-state owner of record
//! 0003 §2.3 (`Source` for the deferred, source-backed package; `Eager` for
//! the owning package), the metadata and tolerance sidecars of §1.6, and the
//! retention policy from [`OpenOptions`]. Cloning bumps one reference count.
//!
//! # Routes
//!
//! The constructors take what detection already produced, so an open does
//! one detection pass and one container index:
//!
//! | Open route | Detection result | Constructor |
//! |---|---|---|
//! | `open` (path) on `unix`/`windows`, `open_bytes`, `open_prepared` | source-backed OPC package | `Docx::from_source` |
//! | [`SourceMode::Eager`], and paths elsewhere | eager `OpcPackage` | `Docx::from_opc_package` |
//! | `Locked::unlock` | decrypted package | `Docx::from_owned_package` |
//!
//! `SourceMode::Eager` reaches this handle through the eager detection
//! route, so the owning package keeps the input's ZIP layout. Only
//! `open_prepared` under `SourceMode::Eager` still converts a source-backed
//! package with `to_owned_package()` (record 0008 A5), because its package
//! is already indexed.
//!
//! # Source freshness
//!
//! Every read on a source-backed owner ends with `DocxOwner::finish`: the
//! source revision is checked after the operation, and a changed source wins
//! over whatever the operation produced. This closes the window between the
//! leaf's payload read and the handle's return, as the facade's
//! `finish_source_docx_result` did.
//!
//! # Markdown
//!
//! With the `markdown` feature the handle also keeps the owning package the
//! Markdown adapter reads relationships, notes, images and numbering from.
//! An eager owner is that package already; a source-backed owner
//! materializes it once, on the first conversion, and every later conversion
//! of the handle or its clones reuses it. The cache never outlives the
//! source revision it was built from: each use checks the source first and
//! last, and the adapter checks it again after the whole conversion, so a
//! changed file is reported as [`Code::IoError`] instead of being answered
//! from the stale package. This is the facade's former `DocxSourceCache`,
//! moved onto the handle.
//!
//! # Unified reads
//!
//! The handle's typed nodes ([`Node`], [`Kind`]) are what `File::get`,
//! `File::find` and `File::query` return for a DOCX document (record 0010
//! §4.5). Paths address the main story on the markup-compatibility-visible
//! view (decision D4): `/body/p[N]` is the Nth direct body paragraph, with
//! `w:customXml` wrappers flattened and display equations counted as
//! `oMathPara[N]` instead; `/body/tbl[N]/tr[M]/tc[K]/p[J]`, `/body/sdt[N]`,
//! `/body/p[N]/r[M]` (OfficeCLI's run set) and `/body/p[N]/hyperlink[M]`
//! follow the same rules, and `p[@paraId=…]`/`sdt[@sdtId=…]` select by
//! identity. A node keeps the leaf values it was read from and nothing
//! else, so the leaf's memos stay the only caches (record 0010 §3.3).

use std::fmt;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata};
use litchi_docx::section::{Inventory, Limits as SectionLimits};
use litchi_docx::{ToleranceReport, source_backed};
use litchi_opc::{OpcPackage, SourceLimits};

use super::signature::{self, Opened, Signature};
use crate::error::{Code, Error, Result, Source};
use crate::options::{OpenOptions, SourceMode};
use crate::transaction::{Error as TransactionError, LimitKind};

// The unified read adapter (M1 step R4a): element kinds and spec table,
// typed nodes, path resolution and query access, the node view's members,
// and the view modes.
#[cfg(feature = "test-support")]
#[doc(hidden)]
#[path = "docx/support.rs"]
pub mod __unified;
mod chart;
mod children;
mod column;
pub(crate) mod create;
mod element;
mod few;
mod lists;
mod marks;
pub(crate) mod merge;
mod metadata;
mod nav;
mod nested;
mod node;
mod parts;
mod picture;
mod props;
mod publish;
pub(crate) mod refresh;
pub(crate) mod setters;
mod smartart;
mod spec;
mod story;
mod structure;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use props::Props;
pub use spec::Kind;
pub(crate) use spec::SPEC;

/// An opened Word OOXML document.
///
/// Cheap to clone and `Send + Sync`. Read methods take `&self`; on a
/// source-backed document each read checks that the file has not changed
/// since it was opened and reports [`Code::IoError`] when it has.
#[derive(Clone)]
pub struct Docx(Arc<Opened<DocxState>>);

/// Shared state behind a [`Docx`] handle.
pub(crate) struct DocxState {
    owner: DocxOwner,
    /// Core properties, read on first request. The eager owner parsed them
    /// at open; the source owner reads the selected part once and is then
    /// bound to the opened revision by [`DocxOwner::finish`].
    metadata: OnceLock<Box<Metadata>>,
    /// Successful explicit signature removal, shared by repeated publications.
    /// Failures are not cached, and source fences guard every use.
    unsigned: OnceLock<litchi_docx::Package>,
    /// The owning package a source-backed owner materializes for the writes
    /// that publish package-graph changes, built on the first such write
    /// (record 0274). Source fences guard every use.
    promoted: OnceLock<litchi_docx::Package>,
    /// The eager owner's open-time report. A source-backed report keeps
    /// growing as deferred parts are read, so it is never cached.
    tolerance: OnceLock<ToleranceReport>,
    /// Whether write verbs may build an editor from the owner
    /// ([`OpenOptions::retain_source`]). Both owner states already hold
    /// everything an edit needs, so this releases nothing; it only refuses.
    retain_source: bool,
    /// The owning package a source-backed owner materializes for the
    /// Markdown adapter, built on the first conversion.
    #[cfg(feature = "markdown")]
    markdown: OnceLock<Box<litchi_docx::Package>>,
    /// The OPC read limits of the open, which an unsigned copy is opened
    /// under (record 0013).
    ooxml_limits: SourceLimits,
    /// The eager owner's body block index, read on the first path step that
    /// needs it (record 0286). The owning package never changes under a
    /// handle, so the index stays valid; a source-backed owner is not cached
    /// here, because its reads are fenced against the source.
    body: OnceLock<litchi_docx::Blocks>,
    /// The styles, numbering, story and notes parts the unified adapter has
    /// read (record 0319), each on first use. A source-backed owner's reads
    /// through them are still fenced against the source.
    parts: parts::Cache,
}

/// What a DOCX write starts from (M2).
pub(crate) enum DocxEditor {
    /// An isolated main-document edit over the handle's own owner, whose
    /// commit publishes through that owner.
    Edit(Box<litchi_docx::document::Edit>),
    /// An isolated edit retaining the complete source-backed owner.
    Source(Box<litchi_docx::document::Edit<source_backed::Package>>),
    /// After [`Signed::unsign`](crate::Signed::unsign): an owning copy of the
    /// document whose signature the leaf removed explicitly
    /// ([`litchi_docx::Package::unsign`]). Edits start from it
    /// (`edit`) and publish into it, so the output carries no
    /// signature part.
    Unsigned(Box<litchi_docx::Package>),
}

impl fmt::Debug for DocxEditor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Edit(_) => "DocxEditor::Edit",
            Self::Source(_) => "DocxEditor::Source",
            Self::Unsigned(_) => "DocxEditor::Unsigned",
        })
    }
}

/// The two owner states of a DOCX handle (record 0003 §2.3).
#[allow(
    clippy::large_enum_variant,
    reason = "one owner per handle, inside its `Arc`; boxing the source owner would add an indirection to every read"
)]
pub(crate) enum DocxOwner {
    /// Deferred part payloads over the retained positional source.
    Source(Box<source_backed::Package>),
    /// The owning package, validated by one text pass at open. Unboxed: the
    /// package is a few handles, and a commit adopts it without an
    /// allocation (record 0305).
    Eager(litchi_docx::Package),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Docx>();
    assert!(size_of::<Docx>() == size_of::<usize>());
};

impl DocxOwner {
    /// Prefer a source revision failure over the result of an operation that
    /// read a source-backed owner. Eager owners pass the result through.
    ///
    /// Inside a resolver or query step for this owner (a
    /// [`crate::path::fence::Scope`] with [`Self::fence_key`]), whose bracket
    /// checks the source once after the whole step, the result passes
    /// through unchecked: the step's closing check observes a change made
    /// before or during the read, and its error wins (records 0262, 0315).
    pub(crate) fn finish<T>(&self, result: Result<T>) -> Result<T> {
        if crate::path::fence::is_active(self.fence_key()) {
            return normalize_transaction_limit(result);
        }
        self.ensure_current()?;
        normalize_transaction_limit(result)
    }

    /// Validate every retained source before or after a public read; inside
    /// a fenced step of this owner, the step's own check covers it.
    pub(crate) fn check_source(&self) -> Result<()> {
        self.finish(Ok(()))
    }

    /// Validate every retained source, whether or not a fenced step is
    /// active: the check a resolver or query bracket itself runs.
    pub(crate) fn ensure_current(&self) -> Result<()> {
        match self {
            Self::Source(package) => {
                package.check_source()?;
                Ok(())
            },
            Self::Eager(_) => Ok(()),
        }
    }

    /// This owner's key for [`crate::path::fence::Scope`]: its address in
    /// the handle's shared state, which every clone of the handle reads.
    pub(crate) fn fence_key(&self) -> usize {
        std::ptr::from_ref(self).addr()
    }
}

/// The DOCX leaf calls its retained before/after pair
/// `ProjectedXmlBytes`.  The unified transaction exposes that same charge as
/// `SnapshotBytes`, so the facade keeps one error vocabulary even when lazy
/// leaf admission rejects the source before a native step is published.
fn normalize_transaction_limit<T>(result: Result<T>) -> Result<T> {
    result.map_err(|error| match error.leaf() {
        Source::DocxTransaction(litchi_docx::document::Error::Limit {
            kind: litchi_docx::document::LimitKind::ProjectedXmlBytes,
            observed,
            maximum,
        }) => Error::from(TransactionError::Limit {
            kind: LimitKind::SnapshotBytes,
            observed: *observed,
            maximum: *maximum,
        }),
        _ => error,
    })
}

impl Docx {
    /// Identity includes signature disposition, origin and other sidecars.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        self.0.retain_source
    }

    /// Whether the eager owner came from an encrypted package whose
    /// publication still needs an explicit re-encryption profile.
    #[cfg(feature = "encryption")]
    pub(crate) fn requires_reencryption(&self) -> bool {
        matches!(&self.0.owner, DocxOwner::Eager(package) if package.encryption().is_some())
    }

    /// The OPC ingress policy retained by this handle for raw re-admission.
    pub(crate) fn ooxml_read_limits(&self) -> SourceLimits {
        self.0.ooxml_limits
    }

    /// Adopt the source-backed package detection built for a path or a
    /// byte buffer. No part payload is read.
    ///
    /// Under [`SourceMode::Eager`] the package is materialized into the
    /// owning form instead, with the eager route's validation pass. The
    /// openers reach this only from `open_prepared`; their other eager
    /// routes call [`Self::from_opc_package`].
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "compatibility constructor used by tests and byte callers"
        )
    )]
    pub(crate) fn from_source(
        package: source_backed::Package,
        options: &OpenOptions,
    ) -> Result<Self> {
        Self::from_source_with_origin(package, options, None)
    }

    /// Build a source-backed owner and retain the path which supplied it.
    /// `origin` is already resolved by the path opener; this constructor does
    /// not inspect a filename or change format detection.
    pub(crate) fn from_source_with_origin(
        package: source_backed::Package,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        if options.source_mode == SourceMode::Eager {
            let owned = package.to_owned_package_preserving_source()?;
            return Self::from_owned_package_with_origin(owned, options, origin);
        }
        // A committed source owner carries an effective topology.  Its
        // physical graph is intentionally unavailable to ordinary callers,
        // but the retained ingress policy remains valid for the new facade.
        let ooxml_limits = package.read_limits();
        Ok(Self::with_owner(
            DocxOwner::Source(Box::new(package)),
            None,
            options,
            ooxml_limits,
        )
        .with_origin(origin))
    }

    /// Build the eager owner from the OPC package detection already parsed
    /// (`DetectedFormat::Docx`).
    ///
    /// As on the facade's eager route, the main story is validated by one
    /// text pass now, and that pass's repairs become the tolerance report.
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "compatibility constructor used by tests and byte callers"
        )
    )]
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        Self::from_opc_package_with_origin(package, options, None)
    }

    /// Build an eager owner and retain a path origin supplied by the path
    /// opener. The package itself remains the sole source of format truth.
    pub(crate) fn from_opc_package_with_origin(
        package: OpcPackage,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        Self::from_owned_package_with_origin(
            litchi_docx::Package::from_opc_package(package)?,
            options,
            origin,
        )
    }

    /// Adopt an owning package, validating its main story with one text pass
    /// as the eager route does. [`crate::Locked::unlock`] uses this for a
    /// package the leaf decrypted, which keeps its encryption profile.
    pub(crate) fn from_owned_package(
        package: litchi_docx::Package,
        options: &OpenOptions,
    ) -> Result<Self> {
        Self::from_owned_package_with_origin(package, options, None)
    }

    /// Adopt an owning package with an optional path origin. This is used by
    /// the encrypted OOXML unlock path once its security wrapper forwards the
    /// origin captured at path open.
    pub(crate) fn from_owned_package_with_origin(
        package: litchi_docx::Package,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let ooxml_limits = package.opc_package().read_limits();
        // `document()` pins the main payload but defers XML traversal; the
        // eager route keeps its historical validation boundary with one
        // linear text pass, whose report then holds the packaging, metadata
        // and main-story repairs of this open.
        let document = package.document()?;
        document.text()?;
        let tolerance = document.tolerance_report();
        Ok(Self::with_owner(
            DocxOwner::Eager(package),
            Some(tolerance),
            options,
            ooxml_limits,
        )
        .with_origin(origin))
    }

    fn with_owner(
        owner: DocxOwner,
        tolerance: Option<ToleranceReport>,
        options: &OpenOptions,
        ooxml_limits: SourceLimits,
    ) -> Self {
        Self(Opened::new(Self::state(
            owner,
            tolerance,
            options,
            ooxml_limits,
        )))
    }

    fn state(
        owner: DocxOwner,
        tolerance: Option<ToleranceReport>,
        options: &OpenOptions,
        ooxml_limits: SourceLimits,
    ) -> DocxState {
        DocxState {
            owner,
            metadata: OnceLock::new(),
            unsigned: OnceLock::new(),
            promoted: OnceLock::new(),
            tolerance: tolerance.map_or_else(OnceLock::new, OnceLock::from),
            retain_source: options.retain_source,
            #[cfg(feature = "markdown")]
            markdown: OnceLock::new(),
            ooxml_limits,
            body: OnceLock::new(),
            parts: parts::Cache::default(),
        }
    }

    /// Adopt a committed owning package without serializing and reopening it.
    ///
    /// The edit transaction has already validated the package's changed
    /// parts.  Keeping the package graph here preserves the source/eager
    /// publication boundary and carries the original limits and retention
    /// policy onto the new immutable handle.
    pub(crate) fn adopt_eager_package(&self, package: litchi_docx::Package) -> Self {
        self.adopt_owner(DocxOwner::Eager(package))
    }

    /// Adopt a committed source-backed package without serializing and
    /// reopening it.
    pub(crate) fn adopt_source_package(&self, package: source_backed::Package) -> Self {
        self.adopt_owner(DocxOwner::Source(Box::new(package)))
    }

    fn adopt_owner(&self, owner: DocxOwner) -> Self {
        let options = OpenOptions::new()
            .retain_source(self.0.retain_source)
            .ooxml_limits(self.0.ooxml_limits);
        let tolerance = self.0.tolerance.get().cloned();
        let signature = match self.signature() {
            // `Docx::edit` only creates an unsigned editor for this state
            // after the leaf has removed the signature parts.
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Self(Opened::adopted(
            Self::state(owner, tolerance, &options, self.0.ooxml_limits),
            signature,
            self.0.origin(),
        ))
    }

    /// The body block index of this handle's package, if a path step has read it.
    pub(crate) fn cached_body(&self) -> Option<&litchi_docx::Blocks> {
        self.0.body.get()
    }

    /// Remember the body block index; a concurrent first read
    /// keeps the index stored first (both describe the same part).
    pub(crate) fn cache_body(&self, blocks: litchi_docx::Blocks) -> &litchi_docx::Blocks {
        self.0.body.get_or_init(|| blocks)
    }

    /// The part elements this handle has read (record 0319).
    pub(crate) fn parts(&self) -> &parts::Cache {
        &self.0.parts
    }

    /// The stable path origin captured by a path open, if any.
    pub(crate) fn origin(&self) -> Option<Arc<Path>> {
        self.0.origin()
    }

    /// Attach path metadata to a handle after a path route has finished
    /// classification. Byte and create routes leave this unset.
    pub(crate) fn with_origin(&self, origin: Option<Arc<Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    /// The same document under another signature disposition (record 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// The owning package of a source-backed owner, materialized once.
    pub(crate) fn promoted(
        &self,
        package: &source_backed::Package,
    ) -> Result<litchi_docx::Package> {
        if let Some(promoted) = self.0.promoted.get() {
            return self.owner().finish(Ok(promoted.clone()));
        }
        let owned = self.owner().finish(
            package
                .to_owned_package_preserving_source()
                .map_err(Error::from),
        )?;
        let _ = self.0.promoted.set(owned.clone());
        Ok(owned)
    }

    /// The owner, for the family views and write handlers built on this
    /// handle. Reads through it must end with [`DocxOwner::finish`].
    pub(crate) fn owner(&self) -> &DocxOwner {
        &self.0.owner
    }

    /// Whether the owner defers part payloads to the retained source.
    #[must_use]
    pub fn is_source_backed(&self) -> bool {
        matches!(self.0.owner, DocxOwner::Source(_))
    }

    /// The visible text of the main document: the text of each paragraph,
    /// table-cell paragraphs included, in document order, separated by
    /// `\n`. An empty paragraph is an empty line, so the paragraph breaks of
    /// the document stay visible, as in the other word-processing formats.
    /// The paragraphs are the ones
    /// [`DocumentView::paragraphs`](crate::DocumentView::paragraphs)
    /// returns, and each one's text is that paragraph's own text (which a
    /// line break inside the paragraph can split across lines).
    ///
    /// # Errors
    ///
    /// Returns [`Code::CorruptFile`] for malformed main-document XML (on a
    /// source-backed document this is first noticed here, not at open), and
    /// [`Code::IoError`] when the source changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// for line in docx.text()?.lines() {
    ///     println!("{line}");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn text(&self) -> Result<String> {
        // One validating pass over the main story, which writes the
        // separator before every paragraph after the first; the leaf's
        // `extract_text` joins the runs with no separator at all (oracle drift
        // B3 in `docs/unified-ops/s8a-python-oracle.md`).
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(package
                .document()?
                .text_with_paragraph_separator(PARAGRAPH_SEPARATOR)?),
            DocxOwner::Source(package) => {
                let result = package
                    .document()
                    .and_then(|document| {
                        document.extract_text_with_paragraph_separator(PARAGRAPH_SEPARATOR)
                    })
                    .map_err(Error::from);
                self.0.owner.finish(result)
            },
        }
    }

    /// The document's core properties. Always `Some` when it succeeds;
    /// a package without core properties reads as empty metadata, as the
    /// facade reports it.
    ///
    /// # Errors
    ///
    /// Returns [`Code::CorruptFile`] for a malformed core-properties part on
    /// a source-backed document, and [`Code::IoError`] when the source
    /// changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// if let Some(title) = docx.metadata()?.and_then(|metadata| metadata.title) {
    ///     println!("{title}");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        self.read_metadata().map(Some)
    }

    /// The application properties (`docProps/app.xml`) the root node reads
    /// as `extended.*` (record 0342). `None` for an absent or invalid part
    /// (a metadata edit reads the part strictly and reports it), and for a
    /// source owner with edited topology, whose physical archive no longer
    /// answers for it.
    pub(super) fn app_props(&self) -> Result<Option<litchi_ooxml_common::properties::app::Props>> {
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(package.app_props().ok().flatten()),
            DocxOwner::Source(package) => {
                let Ok(physical) = package.source_backed_package() else {
                    return Ok(None);
                };
                let props = litchi_ooxml_common::properties::app::read_source_backed(physical)
                    .ok()
                    .flatten();
                // A changed source wins over the read.
                self.0.owner.finish(Ok(props))
            },
        }
    }

    /// The application- and custom-properties part bytes, which the
    /// transaction change test compares (record 0342). `None` when they
    /// cannot be read: a source owner with edited topology, or an invalid
    /// property graph.
    pub(crate) fn property_part_bytes(&self) -> Result<Option<[Option<Vec<u8>>; 2]>> {
        use litchi_ooxml_common::custom::Host;
        use litchi_ooxml_common::properties::{Family, Snapshot};
        let snapshot = match &self.0.owner {
            DocxOwner::Eager(package) => Snapshot::read(package.opc_package(), Host::Word),
            DocxOwner::Source(package) => {
                let Ok(physical) = package.source_backed_package() else {
                    return Ok(None);
                };
                let snapshot = Snapshot::read_source_backed(physical, Host::Word);
                self.0.owner.finish(Ok(()))?;
                snapshot
            },
        };
        let Ok(snapshot) = snapshot else {
            return Ok(None);
        };
        Ok(Some([Family::App, Family::Custom].map(|family| {
            snapshot.part_bytes(family).map(<[u8]>::to_vec)
        })))
    }

    fn read_metadata(&self) -> Result<Metadata> {
        if let Some(metadata) = self.0.metadata.get() {
            return self.0.owner.finish(Ok(Metadata::clone(metadata)));
        }
        let metadata = match &self.0.owner {
            DocxOwner::Eager(package) => package.metadata().cloned().unwrap_or_default(),
            DocxOwner::Source(package) => {
                let result = package.metadata().map_err(Error::from);
                self.0.owner.finish(result)?
            },
        };
        // Concurrent first callers read the same pinned revision; the loser
        // of the race discards an equal value.
        let _ = self.0.metadata.set(Box::new(metadata.clone()));
        Ok(metadata)
    }

    /// Everything the reader repaired while this document was read.
    ///
    /// An eager document reports its open: packaging, core and custom
    /// properties, and the main story's text pass. A source-backed document
    /// reports packaging and core properties (read now if [`Self::metadata`]
    /// has not read them yet), plus the repairs of every deferred read so
    /// far.
    ///
    /// # Errors
    ///
    /// Returns an error when a source-backed document's source changed or
    /// its core-properties part cannot be read.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// assert!(docx.tolerance_report()?.is_clean());
    /// # Ok(())
    /// # }
    /// ```
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        match &self.0.owner {
            DocxOwner::Eager(_) => Ok(self.0.tolerance.get().cloned().unwrap_or_default()),
            DocxOwner::Source(package) => {
                // Core-properties repairs join the source report when the
                // metadata is first read; read it so they are part of this
                // report whether or not the caller asked for it.
                let result = self.read_metadata().map(|_| package.tolerance_report());
                self.0.owner.finish(result)
            },
        }
    }

    /// The main document's local section descriptors: ownership, logical
    /// paragraph ranges, locally authored page properties, and inert
    /// header/footer relationship IDs.
    ///
    /// The inventory owns its values. It does not resolve relationship
    /// targets or compute effective inherited layout. `limits` bound the
    /// scan (`litchi_docx::section::Limits::default()` for the defaults; record 0362).
    ///
    /// # Errors
    ///
    /// Returns [`Code::LimitExceeded`] when `limits` are exceeded,
    /// [`Code::CorruptFile`] for malformed main-document XML, and
    /// [`Code::IoError`] when the source changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// let inventory = docx.section_inventory(&litchi_docx::section::Limits::default())?;
    /// println!("{} sections", inventory.sections().len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn section_inventory(&self, limits: &SectionLimits) -> Result<Inventory> {
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(package
                .with_semantic_limits(package.semantic_limits().clone().sections(limits.clone()))
                .document()?
                .section_inventory()?),
            DocxOwner::Source(package) => {
                let result = package
                    .with_semantic_limits(
                        package.semantic_limits().clone().sections(limits.clone()),
                    )
                    .section_inventory_snapshot()
                    .map(|snapshot| snapshot.inventory().clone())
                    .map_err(Error::from);
                self.0.owner.finish(result)
            },
        }
    }

    /// Build what a write starts from.
    ///
    /// On an unsigned document, both owner states reach the leaf's
    /// `edit` directly, so no container is indexed again. After
    /// [`Signed::unsign`](crate::Signed::unsign) the editor is an owning copy
    /// whose signature the leaf removed explicitly (record 0013): it is opened
    /// over the exact input bytes, or cloned from the eager owner's package,
    /// and then `unsign`ed. Write verbs (M2) build on this.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`] while the document is still
    /// signed, with a suggestion naming `Signed::unsign`;
    /// [`Code::UnsupportedCapability`] when it was opened with
    /// `retain_source(false)`, or when an unlocked encrypted document would
    /// lose its encryption profile in the unsigned copy; [`Code::IoError`]
    /// when the source changed; and the leaf's typed failures otherwise.
    /// Leaf document admission is deferred until its first projection,
    /// mutation or commit.
    pub(crate) fn edit(&self) -> Result<DocxEditor> {
        self.edit_with_options(litchi_docx::document::EditOptions::default())
    }

    /// Build a native editor under one explicit main-document edit profile.
    /// The profile reaches the package owner before it captures the document
    /// snapshot, so a refused source-size floor does not inflate the payload.
    pub(crate) fn edit_with_options(
        &self,
        options: litchi_docx::document::EditOptions,
    ) -> Result<DocxEditor> {
        self.edit_policy()?;
        self.edit_inner(options)
    }

    /// Build an editor with the unified transaction's remaining snapshot
    /// budget admitted before native document capture.
    pub(crate) fn edit_with_snapshot_budget(&self, maximum: usize) -> Result<DocxEditor> {
        self.edit_with_options(
            litchi_docx::document::EditOptions::default().max_snapshot_bytes(maximum),
        )
    }

    /// Build a native writer under one explicit main-document edit profile.
    /// Signed documents retain their exact no-op exception while the selected
    /// profile is admitted before any owning or source-backed snapshot.
    pub(crate) fn edit_for_write_with_options(
        &self,
        options: litchi_docx::document::EditOptions,
    ) -> Result<DocxEditor> {
        if self.0.signature() != Signature::Signed {
            self.edit_policy()?;
        } else if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the DOCX document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Docx)
            .with_suggestion("reopen the document without `OpenOptions::retain_source(false)`"));
        }
        self.edit_inner(options)
    }

    /// Build a native writer with the unified transaction's remaining
    /// snapshot budget admitted before native document capture.
    pub(crate) fn edit_for_write_with_snapshot_budget(&self, maximum: usize) -> Result<DocxEditor> {
        self.edit_for_write_with_options(
            litchi_docx::document::EditOptions::default().max_snapshot_bytes(maximum),
        )
    }

    fn edit_inner(&self, options: litchi_docx::document::EditOptions) -> Result<DocxEditor> {
        if self.0.signature() == Signature::RemovalAuthorized {
            let package = self.unsigned_package()?;
            return if options.max_snapshot_bytes_value().is_some() {
                Ok(DocxEditor::Edit(Box::new(
                    package.edit_with_options(options).map_err(Error::from)?,
                )))
            } else {
                Ok(DocxEditor::Unsigned(Box::new(package)))
            };
        }
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(DocxEditor::Edit(Box::new(
                package.edit_with_options(options).map_err(Error::from)?,
            ))),
            DocxOwner::Source(package) => {
                let result = package.edit_with_options(options).map_err(Error::from);
                self.0
                    .owner
                    .finish(result)
                    .map(|edit| DocxEditor::Source(Box::new(edit)))
            },
        }
    }

    /// The refusals of [`Self::edit`] that come before any editor is built:
    /// a signed document, and a document opened without its source.
    pub(crate) fn edit_policy(&self) -> Result<()> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Docx));
        }
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the DOCX document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Docx)
            .with_suggestion("reopen the document without `OpenOptions::retain_source(false)`"));
        }
        #[cfg(feature = "encryption")]
        if self.requires_reencryption() {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "an unlocked encrypted DOCX document cannot be edited, signed or not",
            )
            .with_format(FileFormat::Docx));
        }
        Ok(())
    }

    /// An owning copy of the document whose signature the leaf removed
    /// explicitly ([`litchi_docx::Package::unsign`]).
    fn unsigned_package(&self) -> Result<litchi_docx::Package> {
        self.owner().check_source()?;
        let result = self.unsigned_package_inner();
        self.owner().finish(result)
    }

    fn unsigned_package_inner(&self) -> Result<litchi_docx::Package> {
        if let Some(package) = self.0.unsigned.get() {
            return Ok(package.clone());
        }
        let package = match &self.0.owner {
            DocxOwner::Eager(package) => {
                // A copy of the graph would drop the encryption profile the
                // eager owner keeps, and the leaf refuses to edit a package
                // it decrypted; refuse rather than publish it in the clear.
                #[cfg(feature = "encryption")]
                if package.encryption().is_some() {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "an unlocked encrypted DOCX document cannot be edited, signed or not",
                    )
                    .with_format(FileFormat::Docx));
                }
                litchi_docx::Package::from_opc_package(package.opc_package().clone())?
            },
            DocxOwner::Source(package) => package.to_owned_package_preserving_source()?,
        };
        let package = package.unsign();
        self.owner().check_source()?;
        // Concurrent first users may each build a candidate; only a complete
        // successfully fenced result is retained. Each writer gets a clone.
        let _ = self.0.unsigned.set(package.clone());
        Ok(package)
    }
}

#[cfg(feature = "markdown")]
impl Docx {
    /// The owning package the Markdown adapter reads relationships, notes,
    /// images and numbering from: the eager owner itself, or the package a
    /// source-backed owner materializes once and caches on the handle.
    ///
    /// The source is checked before the cache is consulted, so a package
    /// materialized by an earlier conversion never answers for a file that
    /// changed since, and again before the package is returned. Failures are
    /// not cached: two first callers may both materialize, and the loser
    /// drops its equal package.
    pub(crate) fn markdown_package(&self) -> Result<&litchi_docx::Package> {
        let source = match &self.0.owner {
            DocxOwner::Eager(package) => return Ok(package),
            DocxOwner::Source(source) => source,
        };
        self.0.owner.check_source()?;
        if let Some(package) = self.0.markdown.get() {
            return self.0.owner.finish(Ok(package.as_ref()));
        }
        // `to_owned_package` reads the pinned source revision and fails when
        // the file changed underneath it.
        let package = self
            .0
            .owner
            .finish(source.to_owned_package().map_err(Error::from))?;
        let _ = self.0.markdown.set(Box::new(package));
        let package = self.0.markdown.get().map(Box::as_ref).ok_or_else(|| {
            Error::new(
                Code::InternalError,
                "the DOCX Markdown package cache is empty after it was filled",
            )
            .with_format(FileFormat::Docx)
        });
        self.0.owner.finish(package)
    }
}

impl fmt::Debug for Docx {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Docx")
            .field("source_backed", &self.is_source_backed())
            .field("retain_source", &self.0.retain_source)
            .finish_non_exhaustive()
    }
}

/// The paragraph separator of [`Docx::text`].
const PARAGRAPH_SEPARATOR: &str = "\n";

#[cfg(test)]
mod tests;
#[cfg(test)]
mod unified_tests;
