//! Handle for OpenDocument Text (`.odt`) inputs.
//!
//! ODT is tier 3 and read-only in M0a (record 0003 §2.3). The handle carries
//! the facade's two payloads over unchanged: the eager `litchi_odt::Document`
//! built from detection's prepared package (byte and reader routes), and the
//! filesystem-backed `litchi_odt::SourceBackedDocument` (path route on
//! `unix`/`windows`). Both keep the package they were opened from, so M6 can
//! add write without reshaping this state.

use std::fmt;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;

use crate::{Error, OpenOptions, Result};

/// An opened ODT document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Odt(Arc<OdtState>);

struct OdtState {
    owner: OdtOwner,
}

/// The ODT read owner, in the state today's route builds.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdtOwner {
    /// Parsed from the package detection prepared (byte and reader routes,
    /// and the path fallback on other targets).
    Eager(litchi_odt::Document),
    /// Filesystem-backed, with deferred package payloads (path route).
    #[cfg(any(unix, windows))]
    Source(litchi_odt::SourceBackedDocument),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Odt>();
    assert!(size_of::<Odt>() == size_of::<usize>());
};

fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odt)
}

impl Odt {
    /// Adopts the package detection prepared (`DetectedFormat::Odt`) without
    /// a second ZIP index. The ODF limits were applied when the package was
    /// prepared.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    pub(crate) fn from_prepared(prepared: PreparedPackage, _options: &OpenOptions) -> Result<Self> {
        let document = litchi_odt::Document::from_prepared_package(prepared).map_err(core_error)?;
        Ok(Self::from_owner(OdtOwner::Eager(document)))
    }

    /// Opens the source-pinned candidate the document path probe returned
    /// (`DocumentSourcePathDetection::Odt`). The candidate's own guards
    /// refuse a source that changed since detection.
    ///
    /// ODT-versus-OOXML polyglot arbitration stays with detection: the caller
    /// passes a candidate only once ODT owns the input.
    #[cfg(any(unix, windows))]
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    pub(crate) fn from_source_candidate(
        candidate: crate::detect::detected::OdtSourcePathCandidate,
        _options: &OpenOptions,
    ) -> Result<Self> {
        let document = candidate.into_document().map_err(core_error)?;
        Ok(Self::from_owner(OdtOwner::Source(document)))
    }

    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    fn from_owner(owner: OdtOwner) -> Self {
        Self(Arc::new(OdtState { owner }))
    }

    /// The document's plain text.
    ///
    /// # Errors
    ///
    /// Returns `corrupt_file` when the body cannot be parsed, and `io_error`
    /// when a filesystem-backed source changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.text().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.text()),
        }
    }

    /// The number of body paragraphs.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn paragraph_count(&self) -> Result<usize> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.paragraph_count().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.paragraph_count()),
        }
    }

    /// Document properties from `meta.xml`. Always `Some` for ODT.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        let metadata = match &self.0.owner {
            OdtOwner::Eager(document) => document.metadata().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.metadata()),
        }?;
        Ok(Some(metadata))
    }

    /// Everything the reader repaired while opening the package and parsing
    /// the body, which happens at open.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match &self.0.owner {
            OdtOwner::Eager(document) => document.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => document.tolerance_report().clone(),
        })
    }

    /// The retained leaf owner, for the crate's family views.
    #[allow(dead_code, reason = "consumed by the `File` family views (M0a S6)")]
    pub(crate) fn owner(&self) -> &OdtOwner {
        &self.0.owner
    }
}

impl OdtOwner {
    /// Binds a read's result to the source revision captured at open: on a
    /// filesystem-backed owner a final source-revision failure wins over the
    /// result the leaf produced after its own check. This closes the TOCTOU
    /// window between the leaf's payload read and the handle's return (the
    /// facade's `finish_source_odt_result`).
    pub(crate) fn finish<T>(&self, result: litchi_core::Result<T>) -> Result<T> {
        match self {
            Self::Eager(_) => result.map_err(core_error),
            #[cfg(any(unix, windows))]
            Self::Source(document) => match document.source_version() {
                Err(error) => Err(core_error(error)),
                Ok(_) => result.map_err(core_error),
            },
        }
    }

    /// Fails with `io_error` when a filesystem-backed source changed since
    /// open; always succeeds for the eager owner. This is the ODT half of
    /// the facade's `check_source_freshness_after_markdown`.
    #[allow(dead_code, reason = "consumed by the `File` markdown adapter (M0a S7)")]
    pub(crate) fn check_source(&self) -> Result<()> {
        match self {
            Self::Eager(_) => Ok(()),
            #[cfg(any(unix, windows))]
            Self::Source(document) => document.source_version().map(drop).map_err(core_error),
        }
    }
}

impl fmt::Debug for Odt {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match &self.0.owner {
            OdtOwner::Eager(_) => "eager",
            #[cfg(any(unix, windows))]
            OdtOwner::Source(_) => "source",
        };
        formatter
            .debug_struct("Odt")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
