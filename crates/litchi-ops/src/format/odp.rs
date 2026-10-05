//! Handle for OpenDocument Presentation (`.odp`) inputs.
//!
//! ODP is tier 3 and read-only in M0a (record 0003 §2.3). The handle carries
//! the facade's two payloads over unchanged: the eager
//! `litchi_odp::Presentation` (byte and reader routes) and the
//! filesystem-backed `litchi_odp::SourceBackedPresentation` (path route on
//! `unix`/`windows`). Both keep the package they were opened from, so M6 can
//! add write without reshaping this state.

use std::fmt;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;

use crate::{Error, OpenOptions, Result};

/// An opened ODP presentation. Cheap to clone (one `Arc` bump);
/// `Send + Sync`.
#[derive(Clone)]
pub struct Odp(Arc<OdpState>);

struct OdpState {
    owner: OdpOwner,
}

/// The ODP read owner, in the state today's route builds.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdpOwner {
    /// Parsed at open (byte and reader routes, and the path fallback on
    /// other targets).
    Eager(litchi_odp::Presentation),
    /// Filesystem-backed, with deferred package data (path route).
    #[cfg(any(unix, windows))]
    Source(litchi_odp::SourceBackedPresentation),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Odp>();
    assert!(size_of::<Odp>() == size_of::<usize>());
};

fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odp)
}

impl Odp {
    /// Adopts the package detection prepared without a second ZIP index.
    /// Every in-memory route uses this, including the one the facade served
    /// by turning a prepared package back into bytes and parsing them again
    /// (`DetectedFormat::Odp`).
    pub(crate) fn from_prepared(prepared: PreparedPackage, _options: &OpenOptions) -> Result<Self> {
        let presentation =
            litchi_odp::Presentation::from_prepared_package(prepared).map_err(core_error)?;
        Ok(Self::from_owner(OdpOwner::Eager(presentation)))
    }

    /// Adopts the source-backed owner the path route built over the file.
    #[cfg(any(unix, windows))]
    pub(crate) fn from_source(
        presentation: litchi_odp::SourceBackedPresentation,
        _options: &OpenOptions,
    ) -> Result<Self> {
        Ok(Self::from_owner(OdpOwner::Source(presentation)))
    }

    fn from_owner(owner: OdpOwner) -> Self {
        Self(Arc::new(OdpState { owner }))
    }

    /// The text of every slide.
    ///
    /// # Errors
    ///
    /// Returns `corrupt_file` when a slide cannot be parsed, and `io_error`
    /// when a filesystem-backed source changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            OdpOwner::Eager(presentation) => presentation.text(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.text(),
        }
        .map_err(core_error)
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn slide_count(&self) -> Result<usize> {
        match &self.0.owner {
            OdpOwner::Eager(presentation) => presentation.slide_count(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.slide_count(),
        }
        .map_err(core_error)
    }

    /// Document properties, as the facade reports them today: always
    /// `Some(Metadata::default())`. The leaf's `meta.xml` projection is not
    /// surfaced yet; changing that is a behaviour change left to M1.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when a filesystem-backed source changed.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        #[cfg(any(unix, windows))]
        if let OdpOwner::Source(presentation) = &self.0.owner {
            presentation.check_source().map_err(core_error)?;
        }
        Ok(Some(Metadata::default()))
    }

    /// Everything the reader repaired while opening the package and scanning
    /// the slides.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match &self.0.owner {
            OdpOwner::Eager(presentation) => presentation.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.tolerance_report().clone(),
        })
    }

    /// The retained leaf owner, for the crate's family views.
    pub(crate) fn owner(&self) -> &OdpOwner {
        &self.0.owner
    }
}

impl fmt::Debug for Odp {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match &self.0.owner {
            OdpOwner::Eager(_) => "eager",
            #[cfg(any(unix, windows))]
            OdpOwner::Source(_) => "source",
        };
        formatter
            .debug_struct("Odp")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
