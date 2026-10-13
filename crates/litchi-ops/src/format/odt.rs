//! Handle for OpenDocument Text (`.odt`) inputs.
//!
//! The handle retains the leaf read owner detection built — the eager
//! `litchi_odt::Document` of the byte and reader routes, with its immutable
//! `litchi_odt::Package`, or the filesystem-backed
//! `litchi_odt::SourceBackedDocument` of the path route — and reads through
//! the leaf body index (`litchi_odt::body`). Edits stage typed body edits on
//! the leaf `Package::edit()` transaction; a filesystem-backed handle
//! materializes its package once, on its first edit or publication.

pub(crate) mod children;
mod nav;
mod node;
mod props;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use props::Props;
pub use spec::Kind;
pub(crate) use spec::SPEC;

use std::fmt;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;
use litchi_odt::body::{Body, Tree};

use crate::{Error, OpenOptions, Result};

/// An opened ODT document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Odt(Arc<OdtState>);

#[allow(dead_code, reason = "TEMP-0276: read by the write adapter")]
struct OdtState {
    owner: OdtOwner,
    /// The immutable leaf package: set for eager owners, materialized on
    /// first use for a filesystem-backed owner.
    package: OnceLock<litchi_odt::Package>,
    /// The body index of a filesystem-backed owner (an eager document keeps
    /// its own).
    tree: OnceLock<Tree>,
    origin: Option<Arc<Path>>,
}

/// The ODT read owner, in the state today's route builds.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdtOwner {
    /// Parsed from the package detection prepared (byte and reader routes,
    /// the path fallback on other targets, and every edit result).
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

/// Typed ODF limit failures stay `Source::Odf`; every other failure keeps
/// its common-core payload as `Source::Core`, exactly as before W3.
pub(crate) fn odf_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => core_error(error),
        other => Error::from(other).with_format(FileFormat::Odt),
    }
}

pub(crate) fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odt)
}

#[allow(dead_code, reason = "TEMP-0276: used by the write adapter")]
impl Odt {
    /// Adopts the package detection prepared (`DetectedFormat::Odt`) without
    /// a second ZIP index. The ODF limits were applied when the package was
    /// prepared.
    pub(crate) fn from_prepared(prepared: PreparedPackage, _options: &OpenOptions) -> Result<Self> {
        let package = litchi_odt::Package::from_prepared_package(prepared).map_err(odf_error)?;
        Self::from_package(package, None)
    }

    /// Adopts an immutable leaf package (a commit result or a fresh one).
    pub(crate) fn from_package(
        package: litchi_odt::Package,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let document = package.document().map_err(core_error)?;
        Ok(Self(Arc::new(OdtState {
            owner: OdtOwner::Eager(document),
            package: OnceLock::from(package),
            tree: OnceLock::new(),
            origin,
        })))
    }

    /// Opens the source-pinned candidate the document path probe returned
    /// (`DocumentSourcePathDetection::Odt`). The candidate's own guards
    /// refuse a source that changed since detection.
    ///
    /// ODT-versus-OOXML polyglot arbitration stays with detection: the caller
    /// passes a candidate only once ODT owns the input.
    #[cfg(any(unix, windows))]
    pub(crate) fn from_source_candidate(
        candidate: crate::detect::detected::OdtSourcePathCandidate,
        options: &OpenOptions,
    ) -> Result<Self> {
        let document = candidate
            .into_document(options.odf_open_options())
            .map_err(core_error)?;
        Ok(Self::from_owner(OdtOwner::Source(document)))
    }

    fn from_owner(owner: OdtOwner) -> Self {
        Self(Arc::new(OdtState {
            owner,
            package: OnceLock::new(),
            tree: OnceLock::new(),
            origin: None,
        }))
    }

    /// The same document, saving to `origin`.
    pub(crate) fn with_origin(self, origin: Option<Arc<Path>>) -> Self {
        let state = match Arc::try_unwrap(self.0) {
            Ok(mut state) => {
                state.origin = origin;
                state
            },
            Err(shared) => {
                let owner = match &shared.owner {
                    OdtOwner::Eager(document) => OdtOwner::Eager(document.clone()),
                    #[cfg(any(unix, windows))]
                    OdtOwner::Source(_) => {
                        // A shared filesystem-backed owner cannot be cloned;
                        // keep sharing it and only record the origin on a
                        // fresh eager state when one exists.
                        return Self(shared);
                    },
                };
                OdtState {
                    owner,
                    package: shared
                        .package
                        .get()
                        .cloned()
                        .map_or_else(OnceLock::new, OnceLock::from),
                    tree: OnceLock::new(),
                    origin,
                }
            },
        };
        Self(Arc::new(state))
    }

    /// The path the document was opened from, if any.
    pub(crate) fn origin(&self) -> Option<&Arc<Path>> {
        self.0.origin.as_ref()
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
    pub(crate) fn owner(&self) -> &OdtOwner {
        &self.0.owner
    }

    /// The structural body index, built once per handle.
    pub(crate) fn body(&self) -> Result<Body<'_>> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.body().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => {
                let content = document.content_xml().map_err(core_error)?;
                let styles = document.styles_xml().map_err(core_error)?;
                let tree = match self.0.tree.get() {
                    Some(tree) => tree,
                    None => {
                        let tree = Tree::build(content, styles).map_err(core_error)?;
                        self.0.tree.get_or_init(|| tree)
                    },
                };
                Ok(tree.view(content, styles))
            },
        }
    }

    /// The immutable leaf package edits start from.
    pub(crate) fn package(&self) -> Result<&litchi_odt::Package> {
        if let Some(package) = self.0.package.get() {
            return Ok(package);
        }
        match &self.0.owner {
            OdtOwner::Eager(_) => Err(Error::new(
                crate::Code::InternalError,
                "an eager ODT handle lost its package",
            )
            .with_format(FileFormat::Odt)),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => {
                let package = self.0.owner.finish(document.to_package())?;
                Ok(self.0.package.get_or_init(|| package))
            },
        }
    }

    /// Fails with `io_error` when a filesystem-backed source changed.
    pub(crate) fn check_source(&self) -> Result<()> {
        self.0.owner.check_source()
    }

    /// Whether two handles retain the same immutable package bytes.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
            || match (self.0.package.get(), other.0.package.get()) {
                (Some(left), Some(right)) => left.as_bytes() == right.as_bytes(),
                _ => false,
            }
    }

    /// Publish the document's package bytes into `sink`.
    pub(crate) fn write_to(&self, sink: impl std::io::Write) -> Result<()> {
        self.package()?.write_to(sink).map_err(odf_error)
    }

    /// Atomically save the document's package to `path`.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let package = self.package()?;
        if options.overwrite == crate::Overwrite::Refuse && path.exists() {
            return Err(Error::new(
                crate::Code::FileExists,
                "the destination exists and the save options refuse to overwrite it",
            )
            .with_format(FileFormat::Odt));
        }
        package
            .save_with_durability(path, options.durability)
            .map_err(odf_error)?;
        Ok(u64::try_from(package.as_bytes().len()).unwrap_or(u64::MAX))
    }

    /// A fresh empty text document.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        write::create(options)
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
