//! Handle for OpenDocument Spreadsheet (`.ods`, flat `.fods`) inputs.
//!
//! The handle keeps the leaf owner each open route builds (record 0277):
//!
//! - the immutable [`litchi_ods::Package`] with its parsed
//!   [`litchi_ods::Workbook`] projection (byte and reader routes, fresh
//!   creation, and every committed edit);
//! - the filesystem-backed [`litchi_ods::SourceBackedSpreadsheet`] on the
//!   path route, which materializes its editable `Package` once, on the
//!   first edit;
//! - the flat single-XML [`litchi_ods::flat::Workbook`], read-only through
//!   the unified layer.
//!
//! Reads borrow the owner's typed worksheets; nothing is copied at open.
//! Writes stage one leaf [`litchi_ods::document::Edit`] per unified call or
//! coalesced batch run and publish its commit as a new handle, leaving this
//! one unchanged.

use std::fmt;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;
use litchi_ods::model::names::Definition;
use litchi_ods::{Package, Sheet};

use crate::{Code, Error, OpenOptions, Result};

mod children;
mod feature_write;
mod features;
mod index;
pub(crate) mod metadata;
mod nav;
mod node;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::{
    ADDABLE, CELL_WRITE_TABLE, REMOVABLE, SETTABLE, SPEC, WRITE_GAPS, capability_props,
};
pub(crate) use views::{Json as ViewJson, view_json, view_to};
pub(crate) use write::Commit as WriteCommit;

/// An opened ODS spreadsheet. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Ods(Arc<OdsState>);

struct OdsState {
    /// The owner and its read caches, shared by handles that differ only in
    /// their origin.
    inner: Arc<Inner>,
    /// The path a path-backed open read, for `File::save`.
    origin: Option<Arc<Path>>,
}

struct Inner {
    owner: OdsOwner,
    /// The caller-selected ODF ingress policy retained across publication
    /// and raw package reads.
    odf_options: litchi_odf_common::OpenOptions,
    /// Worksheet names for the generic sheet traits, which borrow them;
    /// built on their first use.
    names: OnceLock<Box<[String]>>,
    /// Lazily built read indexes (content cells, annotations).
    cache: index::Cache,
}

/// The ODS owner, in the state the open route or the last commit built.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
enum OdsOwner {
    /// The immutable package and its parsed projection (byte and reader
    /// routes, creation, and every committed edit).
    Eager {
        package: Package,
        workbook: litchi_ods::Workbook,
    },
    /// Filesystem-backed, with deferred package-member reads (path route).
    /// The editable package is materialized once, on the first edit.
    #[cfg(any(unix, windows))]
    Source {
        spreadsheet: Box<litchi_ods::SourceBackedSpreadsheet>,
        package: OnceLock<Package>,
    },
    /// A flat single-XML spreadsheet (`.fods`).
    Flat(litchi_ods::flat::Workbook),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Ods>();
    assert!(size_of::<Ods>() == size_of::<usize>());
};

/// Typed ODF limit failures stay `Source::Odf`; every other failure keeps
/// its common-core payload as `Source::Core`, exactly as before W3.
pub(crate) fn odf_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => core_error(error),
        other => Error::from(other).with_format(FileFormat::Ods),
    }
}

pub(crate) fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Ods)
}

impl Ods {
    /// Adopts the package detection prepared without a second ZIP index.
    /// Every in-memory route uses this, including the one the facade served
    /// by turning a prepared package back into bytes and parsing them again
    /// (`DetectedFormat::Ods`).
    pub(crate) fn from_prepared(prepared: PreparedPackage, options: &OpenOptions) -> Result<Self> {
        let odf_options = options.odf_open_options();
        let package = Package::from_prepared_package_with_options(prepared, odf_options)
            .map_err(odf_error)?;
        Self::from_package_with_options(package, None, odf_options)
    }

    /// Adopts the source-backed owner the path route built over the file.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when the source changed since detection, as the
    /// facade's open did through its metadata read.
    #[cfg(any(unix, windows))]
    pub(crate) fn from_source(
        spreadsheet: Box<litchi_ods::SourceBackedSpreadsheet>,
        options: &OpenOptions,
    ) -> Result<Self> {
        spreadsheet.check_source().map_err(odf_error)?;
        Ok(Self::from_owner(
            OdsOwner::Source {
                spreadsheet,
                package: OnceLock::new(),
            },
            None,
            options.odf_open_options(),
        ))
    }

    /// Parses a flat single-XML spreadsheet detection recognized.
    pub(crate) fn from_flat(bytes: Vec<u8>, options: &OpenOptions) -> Result<Self> {
        let flat =
            litchi_ods::flat::Workbook::from_bytes_with_options(&bytes, options.odf_open_options())
                .map_err(odf_error)?;
        Ok(Self::from_owner(
            OdsOwner::Flat(flat),
            None,
            options.odf_open_options(),
        ))
    }

    /// Adopts an immutable package (creation and committed edits), keeping
    /// `origin` for `File::save`.
    pub(crate) fn from_package(package: Package, origin: Option<Arc<Path>>) -> Result<Self> {
        Self::from_package_with_options(package, origin, litchi_odf_common::OpenOptions::new())
    }

    fn from_package_with_options(
        package: Package,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Result<Self> {
        let workbook = package.workbook().map_err(odf_error)?;
        Ok(Self::from_owner(
            OdsOwner::Eager { package, workbook },
            origin,
            odf_options,
        ))
    }

    fn from_owner(
        owner: OdsOwner,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Self {
        Self(Arc::new(OdsState {
            inner: Arc::new(Inner {
                owner,
                odf_options,
                names: OnceLock::new(),
                cache: index::Cache::default(),
            }),
            origin,
        }))
    }

    /// The same owner with `origin` as its save destination.
    pub(crate) fn with_origin(&self, origin: Option<Arc<Path>>) -> Self {
        Self(Arc::new(OdsState {
            inner: Arc::clone(&self.0.inner),
            origin,
        }))
    }

    fn owner(&self) -> &OdsOwner {
        &self.0.inner.owner
    }

    /// The ODF ingress policy retained by this handle.
    pub(crate) fn odf_options(&self) -> litchi_odf_common::OpenOptions {
        self.0.inner.odf_options
    }

    /// The path a path-backed open read.
    pub(crate) fn origin(&self) -> Option<Arc<Path>> {
        self.0.origin.clone()
    }

    /// Whether two handles share one owner (they may differ in origin).
    pub(crate) fn same_inner(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0.inner, &other.0.inner)
    }

    /// The exact bytes of a flat single-XML spreadsheet.
    pub(crate) fn flat_bytes(&self) -> Option<&[u8]> {
        match self.owner() {
            OdsOwner::Flat(flat) => Some(flat.as_bytes()),
            _ => None,
        }
    }

    /// The policy retained by a flat single-XML spreadsheet.
    pub(crate) fn flat_options(&self) -> Option<litchi_odf_common::OpenOptions> {
        match self.owner() {
            OdsOwner::Flat(flat) => Some(flat.options()),
            _ => None,
        }
    }

    /// Whether the owner defers package-member reads to the file, for the
    /// tests that pin each route's owner.
    #[cfg(test)]
    pub(crate) fn is_source_backed(&self) -> bool {
        match self.owner() {
            OdsOwner::Eager { .. } | OdsOwner::Flat(_) => false,
            #[cfg(any(unix, windows))]
            OdsOwner::Source { .. } => true,
        }
    }

    /// Fence a filesystem-backed owner against source changes.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when the file changed after open.
    pub(crate) fn check_source(&self) -> Result<()> {
        match self.owner() {
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.check_source().map_err(odf_error),
            _ => Ok(()),
        }
    }

    /// Run one read and fence the source after it; a freshness failure takes
    /// precedence over the read's own result.
    pub(crate) fn finish<T>(&self, operation: impl FnOnce() -> Result<T>) -> Result<T> {
        let value = operation();
        match self.check_source() {
            Err(error) => Err(error),
            Ok(()) => value,
        }
    }

    /// The typed worksheets in document order, borrowed from the owner.
    /// Every owner holds its parsed worksheets, so nothing is decoded here.
    pub(crate) fn worksheets(&self) -> Result<&[Sheet]> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.sheets()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.sheets().map_err(odf_error),
            OdsOwner::Flat(flat) => Ok(flat.sheets()),
        }
    }

    /// Named ranges and named expressions in document order. A flat
    /// spreadsheet's definitions are not projected by its leaf owner.
    pub(crate) fn definitions(&self) -> Result<&[Definition]> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.definitions()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.definitions().map_err(odf_error),
            OdsOwner::Flat(_) => Ok(&[]),
        }
    }

    /// The `content.xml` annotations are parsed from. A flat spreadsheet's
    /// annotations are not projected (record 0277 gap F2).
    pub(crate) fn content_xml(&self) -> Result<&str> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.content_xml()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.content_xml().map_err(odf_error),
            OdsOwner::Flat(_) => Err(Error::new(
                Code::UnsupportedCapability,
                "flat ODS annotations are not projected into unified nodes",
            )
            .with_format(FileFormat::Ods)
            .with_suggestion("follow-up: ODS adapter (record 0277 gap F2)")),
        }
    }

    /// The read caches of this snapshot.
    pub(crate) fn cache(&self) -> &index::Cache {
        &self.0.inner.cache
    }

    /// Worksheet names in document order.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when a filesystem-backed source changed.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.sheet_names()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.sheet_names().map_err(odf_error),
            OdsOwner::Flat(flat) => Ok(flat.sheets().iter().map(|s| s.name.clone()).collect()),
        }
    }

    /// The number of worksheets.
    ///
    /// # Errors
    ///
    /// As for [`Self::worksheet_names`].
    pub fn worksheet_count(&self) -> Result<usize> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.sheet_count()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.sheet_count().map_err(odf_error),
            OdsOwner::Flat(flat) => Ok(flat.sheets().len()),
        }
    }

    /// Worksheet names the generic sheet traits borrow: read once through
    /// [`Self::worksheets`], and empty when that read fails.
    pub(crate) fn trait_names(&self) -> &[String] {
        self.0.inner.names.get_or_init(|| {
            self.worksheets()
                .map(|sheets| sheets.iter().map(|sheet| sheet.name.clone()).collect())
                .unwrap_or_default()
        })
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
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => workbook.text(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.text(),
            OdsOwner::Flat(flat) => flat.text(),
        }
        .map_err(odf_error)
    }

    /// Document properties from `meta.xml`, borrowed from the owner.
    pub(crate) fn metadata_ref(&self) -> Result<&Metadata> {
        match self.owner() {
            OdsOwner::Eager { workbook, .. } => Ok(workbook.metadata()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.metadata().map_err(odf_error),
            OdsOwner::Flat(flat) => Ok(flat.metadata()),
        }
    }

    /// Document properties from `meta.xml`. Always `Some` for ODS.
    ///
    /// # Errors
    ///
    /// As for [`Self::worksheet_names`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(Some(self.metadata_ref()?.clone()))
    }

    /// Everything the reader repaired while opening the package and parsing
    /// the worksheets, which happens at open.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match self.owner() {
            OdsOwner::Eager { workbook, .. } => workbook.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { spreadsheet, .. } => spreadsheet.tolerance_report().clone(),
            OdsOwner::Flat(flat) => flat.tolerance_report().clone(),
        })
    }

    /// The immutable package of a packaged owner: the retained one, or the
    /// path route's file materialized once and fenced against change.
    ///
    /// # Errors
    ///
    /// `unsupported_capability` for a flat spreadsheet; `io_error` when the
    /// file changed after open; the leaf's typed failure otherwise.
    pub(crate) fn materialized(&self) -> Result<Package> {
        match self.owner() {
            OdsOwner::Eager { package, .. } => Ok(package.clone()),
            #[cfg(any(unix, windows))]
            OdsOwner::Source {
                spreadsheet,
                package,
            } => {
                if let Some(package) = package.get() {
                    spreadsheet.check_source().map_err(odf_error)?;
                    return Ok(package.clone());
                }
                let materialized = spreadsheet.materialize().map_err(odf_error)?;
                Ok(package.get_or_init(|| materialized).clone())
            },
            // A flat spreadsheet's edits start from its packaged projection
            // and republish into the flat bytes (record 0292).
            OdsOwner::Flat(flat) => {
                let flat_limits = litchi_odf_common::flat::Limits::from_options(flat.options());
                let parts = litchi_odf_common::flat::Parts::split(flat.as_bytes(), &flat_limits)
                    .map_err(odf_error)?;
                let bytes = parts.to_bytes(&flat_limits).map_err(odf_error)?;
                // The flat reader admitted the document; a projection the
                // packaged reader refuses is a gap of the edit.
                Package::from_bytes_with_options(&bytes, flat_limits.projection_options()).map_err(|error| {
                    odf_error(error)
                        .caused_by_argument(Code::UnsupportedCapability)
                        .with_suggestion(
                            "follow-up: 0292 flat ODF projection admission; the packaged reader refuses this flat document's parts",
                        )
                })
            },
        }
    }

    /// Create the minimal valid spreadsheet: one empty worksheet.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::create::{Conformance, Flavor};
        let refusal = if options.flavor_value() != Flavor::Document {
            Some((
                "ODS template creation is not integrated",
                "use Flavor::Document",
            ))
        } else if options.conformance_value() == Conformance::Strict {
            Some((
                "ODS does not support OOXML Strict conformance",
                "use Conformance::Transitional",
            ))
        } else {
            None
        };
        if let Some((message, suggestion)) = refusal {
            return Err(Error::new(Code::UnsupportedCapability, message)
                .with_format(FileFormat::Ods)
                .with_suggestion(suggestion));
        }
        let package = match options.metadata_value() {
            Some(metadata) => metadata::create(metadata)?,
            None => Package::new().map_err(odf_error)?,
        };
        Self::from_package(package, None)
    }

    /// Publish this snapshot's exact bytes to `sink`.
    pub(crate) fn write_to<W: std::io::Write>(&self, sink: &mut W) -> Result<()> {
        self.check_source()?;
        match self.owner() {
            OdsOwner::Eager { package, .. } => {
                sink.write_all(package.as_bytes())?;
            },
            #[cfg(any(unix, windows))]
            OdsOwner::Source { .. } => {
                let package = self.materialized()?;
                sink.write_all(package.as_bytes())?;
            },
            OdsOwner::Flat(flat) => sink.write_all(flat.as_bytes())?,
        }
        self.check_source()
    }

    /// Publish this snapshot to `path` through the shared sibling-temporary
    /// atomic replacement the leaf `Package::save` uses.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        self.check_source()?;
        let overwrite = match options.overwrite {
            crate::Overwrite::Replace => litchi_core::atomic::Overwrite::Replace,
            crate::Overwrite::Refuse => litchi_core::atomic::Overwrite::Refuse,
        };
        let package;
        let bytes: &[u8] = match self.owner() {
            OdsOwner::Eager { package, .. } => package.as_bytes(),
            #[cfg(any(unix, windows))]
            OdsOwner::Source { .. } => {
                package = self.materialized()?;
                package.as_bytes()
            },
            OdsOwner::Flat(flat) => flat.as_bytes(),
        };
        litchi_core::atomic::save(
            path,
            options.durability,
            overwrite,
            |writer| std::io::Write::write_all(writer, bytes).map_err(Error::from),
            || self.check_source(),
        )?;
        Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX))
    }
}

impl fmt::Debug for Ods {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match self.owner() {
            OdsOwner::Eager { .. } => "eager",
            #[cfg(any(unix, windows))]
            OdsOwner::Source { .. } => "source",
            OdsOwner::Flat(_) => "flat",
        };
        formatter
            .debug_struct("Ods")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
