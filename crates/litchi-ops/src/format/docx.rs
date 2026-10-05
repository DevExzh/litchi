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

use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata};
use litchi_docx::section::{Inventory, Limits as SectionLimits};
use litchi_docx::{ToleranceReport, source_backed};
use litchi_opc::OpcPackage;

use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

/// An opened Word OOXML document.
///
/// Cheap to clone and `Send + Sync`. Read methods take `&self`; on a
/// source-backed document each read checks that the file has not changed
/// since it was opened and reports [`Code::IoError`] when it has.
#[derive(Clone)]
pub struct Docx(Arc<DocxState>);

/// Shared state behind a [`Docx`] handle.
pub(crate) struct DocxState {
    owner: DocxOwner,
    /// Core properties, read on first request. The eager owner parsed them
    /// at open; the source owner reads the selected part once and is then
    /// bound to the opened revision by [`DocxOwner::finish`].
    metadata: OnceLock<Metadata>,
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
}

/// The two owner states of a DOCX handle (record 0003 §2.3).
#[allow(
    clippy::large_enum_variant,
    reason = "one owner per handle, inside its `Arc`; boxing the source owner would add an indirection to every read"
)]
pub(crate) enum DocxOwner {
    /// Deferred part payloads over the retained positional source.
    Source(source_backed::Package),
    /// The owning package, validated by one text pass at open.
    Eager(Box<litchi_docx::Package>),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Docx>();
    assert!(size_of::<Docx>() == size_of::<usize>());
};

impl DocxOwner {
    /// Prefer a source revision failure over the result of an operation that
    /// read a source-backed owner. Eager owners pass the result through.
    pub(crate) fn finish<T>(&self, result: Result<T>) -> Result<T> {
        match self {
            Self::Source(package) => {
                package.source_version()?;
                result
            },
            Self::Eager(_) => result,
        }
    }

    /// Check that a source-backed owner's file is unchanged: the facade's
    /// `check_source_freshness_after_markdown`, which the Markdown adapter
    /// runs after a whole conversion.
    #[cfg_attr(
        not(any(test, feature = "markdown")),
        expect(dead_code, reason = "used by the Markdown adapter")
    )]
    pub(crate) fn check_source(&self) -> Result<()> {
        self.finish(Ok(()))
    }
}

impl Docx {
    /// Adopt the source-backed package detection built for a path or a
    /// byte buffer. No part payload is read.
    ///
    /// Under [`SourceMode::Eager`] the package is materialized into the
    /// owning form instead, with the eager route's validation pass. The
    /// openers reach this only from `open_prepared`; their other eager
    /// routes call [`Self::from_opc_package`].
    pub(crate) fn from_source(
        package: source_backed::Package,
        options: &OpenOptions,
    ) -> Result<Self> {
        if options.source_mode == SourceMode::Eager {
            // The conversion re-checks the source revision before it returns.
            return Self::from_owned_package(package.to_owned_package()?, options);
        }
        Ok(Self::with_owner(DocxOwner::Source(package), None, options))
    }

    /// Build the eager owner from the OPC package detection already parsed
    /// (`DetectedFormat::Docx`).
    ///
    /// As on the facade's eager route, the main story is validated by one
    /// text pass now, and that pass's repairs become the tolerance report.
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        Self::from_owned_package(litchi_docx::Package::from_opc_package(package)?, options)
    }

    /// Adopt an owning package, validating its main story with one text pass
    /// as the eager route does. [`crate::Locked::unlock`] uses this for a
    /// package the leaf decrypted, which keeps its encryption profile.
    pub(crate) fn from_owned_package(
        package: litchi_docx::Package,
        options: &OpenOptions,
    ) -> Result<Self> {
        let package = Box::new(package);
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
        ))
    }

    fn with_owner(
        owner: DocxOwner,
        tolerance: Option<ToleranceReport>,
        options: &OpenOptions,
    ) -> Self {
        let state = DocxState {
            owner,
            metadata: OnceLock::new(),
            tolerance: tolerance.map_or_else(OnceLock::new, OnceLock::from),
            retain_source: options.retain_source,
            #[cfg(feature = "markdown")]
            markdown: OnceLock::new(),
        };
        Self(Arc::new(state))
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

    /// The visible text of the main document.
    ///
    /// # Errors
    ///
    /// Returns [`Code::CorruptFile`] for malformed main-document XML (on a
    /// source-backed document this is first noticed here, not at open), and
    /// [`Code::IoError`] when the source changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// let text = docx.text()?;
    /// println!("{text}");
    /// # Ok(())
    /// # }
    /// ```
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(package.document()?.text()?),
            DocxOwner::Source(package) => {
                let result = package
                    .document()
                    .and_then(|document| document.extract_text())
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

    fn read_metadata(&self) -> Result<Metadata> {
        if let Some(metadata) = self.0.metadata.get() {
            return self.0.owner.finish(Ok(metadata.clone()));
        }
        let metadata = match &self.0.owner {
            DocxOwner::Eager(package) => package
                .props()
                .cloned()
                .map(Metadata::from)
                .unwrap_or_default(),
            DocxOwner::Source(package) => {
                let result = package.metadata().map_err(Error::from);
                self.0.owner.finish(result)?
            },
        };
        // Concurrent first callers read the same pinned revision; the loser
        // of the race discards an equal value.
        let _ = self.0.metadata.set(metadata.clone());
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

    /// The main document's local section descriptors under the default
    /// semantic limits; see [`Self::section_inventory_with_limits`].
    ///
    /// # Errors
    ///
    /// The same as [`Self::section_inventory_with_limits`].
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// let inventory = docx.section_inventory()?;
    /// println!("{} sections", inventory.sections().len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn section_inventory(&self) -> Result<Inventory> {
        self.section_inventory_with_limits(&SectionLimits::default())
    }

    /// The main document's local section descriptors: ownership, logical
    /// paragraph ranges, locally authored page properties, and inert
    /// header/footer relationship IDs.
    ///
    /// The inventory owns its values. It does not resolve relationship
    /// targets or compute effective inherited layout.
    ///
    /// # Errors
    ///
    /// Returns [`Code::LimitExceeded`] when `limits` are exceeded,
    /// [`Code::CorruptFile`] for malformed main-document XML, and
    /// [`Code::IoError`] when the source changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// let limits = litchi_docx::section::Limits::default();
    /// let inventory = docx.section_inventory_with_limits(&limits)?;
    /// # let _ = inventory;
    /// # Ok(())
    /// # }
    /// ```
    pub fn section_inventory_with_limits(&self, limits: &SectionLimits) -> Result<Inventory> {
        match &self.0.owner {
            DocxOwner::Eager(package) => {
                Ok(package.document()?.section_inventory_with_limits(limits)?)
            },
            DocxOwner::Source(package) => {
                let result = package
                    .section_inventory_snapshot_with_limits(limits)
                    .map(|snapshot| snapshot.inventory().clone())
                    .map_err(Error::from);
                self.0.owner.finish(result)
            },
        }
    }

    /// Start an isolated main-document edit from the retained owner.
    ///
    /// Both owner states reach the leaf's `edit_document` directly, so no
    /// container is indexed again. Write verbs (M2) build on this.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedCapability`] when the document was opened
    /// with `retain_source(false)`, and the leaf's typed snapshot failures
    /// otherwise.
    #[cfg_attr(not(test), expect(dead_code, reason = "write verbs arrive in M2"))]
    pub(crate) fn edit_document(&self) -> Result<litchi_docx::document::Edit> {
        if !self.0.retain_source {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the DOCX document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Docx)
            .with_suggestion("reopen the document without `OpenOptions::retain_source(false)`"));
        }
        match &self.0.owner {
            DocxOwner::Eager(package) => Ok(package.edit_document()?),
            DocxOwner::Source(package) => {
                let result = package.edit_document().map_err(Error::from);
                self.0.owner.finish(result)
            },
        }
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

#[cfg(test)]
mod tests;
