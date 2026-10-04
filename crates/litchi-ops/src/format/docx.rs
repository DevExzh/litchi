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
//! | Facade route today | Detection result | Constructor |
//! |---|---|---|
//! | `Document::open` on `unix`/`windows` | `DocxSourcePathDetection::Docx` | `Docx::from_source` |
//! | `Document::from_bytes` | `DocxSourceBytesDetection::Docx` | `Docx::from_source` |
//! | `Document::open` elsewhere, and `from_detected` | `DetectedFormat::Docx` | `Docx::from_opc_package` |
//! | `Document::from_prepared` | `PreparedOoxmlSource` | `Docx::from_prepared` |
//!
//! The facade has no reader route for documents; a caller with a `Read`
//! detects it into bytes first.
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
//! The facade also keeps a lazily materialized owning package for its
//! Markdown adapter (`DocxSourceCache`). `litchi-ops` has no `markdown`
//! feature, so that cache is not ported here; S6/S7 add it next to
//! `impl ToMarkdown for File`, as a further `OnceLock` in `DocxState`,
//! bound to the same freshness guard.

use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata};
use litchi_docx::section::{Inventory, Limits as SectionLimits};
use litchi_docx::{ToleranceReport, source_backed};
use litchi_opc::{OpcPackage, PreparedOoxmlSource};

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
}

/// The two owner states of a DOCX handle (record 0003 §2.3).
#[allow(
    clippy::large_enum_variant,
    reason = "one owner per handle, inside its `Arc`; boxing the source owner would add an indirection to every read"
)]
#[cfg_attr(
    not(test),
    allow(dead_code, reason = "constructed by the `File` openers from S6")
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
    /// `check_source_freshness_after_markdown`, for the Markdown adapter S7
    /// moves onto `File`.
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "used by the Markdown adapter from S7")
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
    /// owning form instead, with the eager route's validation pass.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the `File` openers from S6")
    )]
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
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        Self::from_owned_package(litchi_docx::Package::from_opc_package(package)?, options)
    }

    /// Adopt a package prepared by `detect_and_prepare` without indexing the
    /// container again.
    ///
    /// Returns [`Code::UnsupportedFormat`] naming the detected format when
    /// the prepared package is another OOXML family, and [`Code::IoError`]
    /// when the file changed after it was prepared.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the `File` openers from S6")
    )]
    pub(crate) fn from_prepared(
        prepared: PreparedOoxmlSource,
        options: &OpenOptions,
    ) -> Result<Self> {
        let detected =
            crate::detect::ooxml::format_for_main_content_type(prepared.main_part_content_type())
                .ok_or(litchi_core::Error::NotOfficeFile)?;
        if detected != FileFormat::Docx {
            return Err(litchi_core::Error::UnexpectedFormat { detected }.into());
        }
        let package = prepared.into_package()?;
        let package = source_backed::Package::from_source_backed_package(package)?;
        Self::from_source(package, options)
    }

    fn from_owned_package(package: litchi_docx::Package, options: &OpenOptions) -> Result<Self> {
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
        };
        Self(Arc::new(state))
    }

    /// The owner, for the family views and write handlers built on this
    /// handle. Reads through it must end with [`DocxOwner::finish`].
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "used by the family views from S6")
    )]
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

    /// The document's core properties; empty when the package has none.
    ///
    /// # Errors
    ///
    /// Returns [`Code::CorruptFile`] for a malformed core-properties part on
    /// a source-backed document, and [`Code::IoError`] when the source
    /// changed after open.
    ///
    /// ```no_run
    /// # fn demo(docx: &litchi_ops::format::docx::Docx) -> litchi_ops::Result<()> {
    /// if let Some(title) = docx.metadata()?.title {
    ///     println!("{title}");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn metadata(&self) -> Result<Metadata> {
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
                let result = self.metadata().map(|_| package.tolerance_report());
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
