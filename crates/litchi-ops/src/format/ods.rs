//! Handle for OpenDocument Spreadsheet (`.ods`) inputs.
//!
//! ODS is tier 3 and read-only in M0a (record 0003 §2.3). The handle carries
//! the facade's two payloads over: the eager `litchi_ods::Spreadsheet` (byte
//! and reader routes) and the filesystem-backed
//! `litchi_ods::SourceBackedSpreadsheet` (path route on `unix`/`windows`).
//!
//! The facade held the eager spreadsheet in a `RefCell`, which made its
//! `Workbook` `!Sync` (`litchi::sheet::workbook_types`). Every read API the
//! facade calls takes `&self`, so the owner is held directly here: no
//! interior mutability, and the handle is `Sync`.
//!
//! The facade's metadata sidecar is not copied at open: both owners keep the
//! projection they parsed, and [`Ods::metadata`] reads it on demand.

use std::fmt;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;

use crate::{Error, OpenOptions, Result};

/// An opened ODS spreadsheet. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Ods(Arc<OdsState>);

struct OdsState {
    owner: OdsOwner,
}

/// The ODS read owner, in the state today's route builds.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdsOwner {
    /// Parsed in full at open (byte and reader routes, and the path fallback
    /// on other targets).
    Eager(litchi_ods::Spreadsheet),
    /// Filesystem-backed, with deferred package-member reads (path route).
    #[cfg(any(unix, windows))]
    Source(litchi_ods::SourceBackedSpreadsheet),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Ods>();
    assert!(size_of::<Ods>() == size_of::<usize>());
};

fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Ods)
}

impl Ods {
    /// Adopts the package detection prepared (`detect_prepared_ods*`) without
    /// a second ZIP index. This is the facade's first byte-route arm.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    pub(crate) fn from_prepared(prepared: PreparedPackage, _options: &OpenOptions) -> Result<Self> {
        let spreadsheet =
            litchi_ods::Spreadsheet::from_prepared_package(prepared).map_err(core_error)?;
        Ok(Self::from_owner(OdsOwner::Eager(spreadsheet)))
    }

    /// Parses the bytes smart detection retained (`DetectedFormat::Ods`).
    /// The leaf adopts its own prepared index when the bytes carry one.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    pub(crate) fn from_bytes(bytes: Vec<u8>, _options: &OpenOptions) -> Result<Self> {
        let spreadsheet = litchi_ods::Spreadsheet::from_bytes(bytes).map_err(core_error)?;
        Ok(Self::from_owner(OdsOwner::Eager(spreadsheet)))
    }

    /// Adopts the source-backed owner the workbook path probe built
    /// (`WorkbookSourcePathDetection::Ods`).
    ///
    /// # Errors
    ///
    /// Returns `io_error` when the source changed since detection, as the
    /// facade's open did through its metadata read.
    #[cfg(any(unix, windows))]
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    pub(crate) fn from_source(
        spreadsheet: Box<litchi_ods::SourceBackedSpreadsheet>,
        _options: &OpenOptions,
    ) -> Result<Self> {
        spreadsheet.check_source().map_err(core_error)?;
        Ok(Self::from_owner(OdsOwner::Source(*spreadsheet)))
    }

    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "called by the single-pass `open` (M0a S6)")
    )]
    fn from_owner(owner: OdsOwner) -> Self {
        Self(Arc::new(OdsState { owner }))
    }

    /// Worksheet names in document order.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when a filesystem-backed source changed.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => Ok(spreadsheet.sheet_names()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.sheet_names().map_err(core_error),
        }
    }

    /// The number of worksheets.
    ///
    /// # Errors
    ///
    /// As for [`Self::worksheet_names`].
    pub fn worksheet_count(&self) -> Result<usize> {
        match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => Ok(spreadsheet.sheet_count()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.sheet_count().map_err(core_error),
        }
    }

    /// The typed worksheets in document order, borrowed from the owner.
    ///
    /// The facade has no ODS sheet adapter to port (its `sheet::adapters`
    /// covers XLSX and XLSB only), so this is the static base a workbook
    /// family view can build one on. Both owners hold their parsed
    /// worksheets, so nothing is decoded here.
    #[allow(dead_code, reason = "consumed by the `File` family views (M0a S6)")]
    pub(crate) fn worksheets(&self) -> Result<&[litchi_ods::Sheet]> {
        match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => Ok(spreadsheet.sheets()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.sheets().map_err(core_error),
        }
    }

    /// Displayed worksheet text: tab-separated cells, newline-separated rows,
    /// in sheet order.
    ///
    /// # Errors
    ///
    /// Returns an error when the bounded text projection exceeds its ceiling
    /// or cannot reserve its output, and `io_error` when a filesystem-backed
    /// source changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => spreadsheet.text(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.text(),
        }
        .map_err(core_error)
    }

    /// Document properties from `meta.xml`. Always `Some` for ODS.
    ///
    /// # Errors
    ///
    /// As for [`Self::worksheet_names`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        let metadata = match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => spreadsheet.metadata(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.metadata().map_err(core_error)?,
        };
        Ok(Some(metadata.clone()))
    }

    /// Everything the reader repaired while opening the package and parsing
    /// the worksheets, which happens at open.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match &self.0.owner {
            OdsOwner::Eager(spreadsheet) => spreadsheet.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source(spreadsheet) => spreadsheet.tolerance_report().clone(),
        })
    }

    /// The retained leaf owner, for the crate's family views.
    #[allow(dead_code, reason = "consumed by the `File` family views (M0a S6)")]
    pub(crate) fn owner(&self) -> &OdsOwner {
        &self.0.owner
    }
}

impl fmt::Debug for Ods {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match &self.0.owner {
            OdsOwner::Eager(_) => "eager",
            #[cfg(any(unix, windows))]
            OdsOwner::Source(_) => "source",
        };
        formatter
            .debug_struct("Ods")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
