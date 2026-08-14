//! Positional, source-backed access to legacy PowerPoint packages.
//!
//! The ordinary [`super::Package`] remains the compatibility path for
//! `Read + Seek` callers.  This owner adapts an immutable positional source to
//! the native CFB reader and lets the presentation defer the optional
//! `Pictures` stream until an image query is made.

use super::{Error, OpenOptions, RecordLimits, Result};
use crate::presentation::Presentation;
use litchi_cfb::{OleError, SharedOleFile, SharedOleFileLimits};
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::property_set::{Metadata, SharedPropertySetReader};
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::Arc;

/// An immutable positional legacy PowerPoint package.
///
/// Opening validates the complete CFB structure through [`SharedOleFile`].
/// The mandatory `PowerPoint Document` and `Current User` streams are read
/// when [`Self::presentation`] is requested; the optional `Pictures` stream
/// remains a descriptor until `images`, `image`, or `image_at` is called on
/// the returned presentation.
#[derive(Clone)]
pub struct SourceBackedPackage {
    shared: Arc<SharedOleFile>,
    record_limits: RecordLimits,
}

impl std::fmt::Debug for SourceBackedPackage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedPackage")
            .field("file_size", &self.shared.file_size())
            .field("record_limits", &self.record_limits)
            .finish_non_exhaustive()
    }
}

impl SourceBackedPackage {
    /// Open a positional source with the default finite PPT limits.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, RecordLimits::default())
    }

    /// Open a positional source with explicit finite PPT limits.
    ///
    /// The CFB index is validated before this method returns.  Stream payloads
    /// are not read by the package constructor.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        record_limits: RecordLimits,
    ) -> Result<Self> {
        let cfb_limit = record_limits
            .max_package_bytes
            .min(usize::try_from(SharedOleFileLimits::MAX_INPUT_BYTES).unwrap_or(usize::MAX));
        if cfb_limit == 0 {
            return Err(Error::ResourceLimit(
                "PPT package limit must be non-zero".to_string(),
            ));
        }
        let expected_version = source.version().map_err(OleError::Io)?;
        let source_length = source.len().map_err(OleError::Io)?;
        let observed_version = source.version().map_err(OleError::Io)?;
        if observed_version != expected_version {
            return Err(Error::Ole(OleError::SourceChanged {
                expected: expected_version,
                observed: observed_version,
            }));
        }
        let source_bytes = usize::try_from(source_length).map_err(|_error| {
            Error::ResourceLimit("PPT package size exceeds this platform".to_string())
        })?;
        if source_bytes > record_limits.max_package_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT package size {source_bytes} exceeds limit {}",
                record_limits.max_package_bytes
            )));
        }
        let cfb_limit = u64::try_from(cfb_limit).map_err(|_error| {
            Error::ResourceLimit("PPT package limit does not fit u64".to_string())
        })?;
        if source_length > cfb_limit {
            return Err(Error::ResourceLimit(format!(
                "PPT package size {source_length} exceeds shared CFB limit {cfb_limit}"
            )));
        }
        let cfb_limits = SharedOleFileLimits::new(cfb_limit)?;
        let shared = SharedOleFile::open_with_limits(source, cfb_limits)?;
        Ok(Self {
            shared: Arc::new(shared),
            record_limits,
        })
    }

    /// Adopts an already validated positional CFB owner when it contains
    /// native PowerPoint ownership streams.
    ///
    /// Classification happens before the PPT package-size policy is applied,
    /// so a generic facade can preserve its fallback behavior for non-PPT OLE
    /// inputs. A recognized PPT owner still receives the normal typed resource
    /// error when it exceeds the default package ceiling.
    pub fn from_shared_if_powerpoint(shared: SharedOleFile) -> Result<Option<Self>> {
        if !shared_is_powerpoint(&shared) {
            return Ok(None);
        }
        let source_length = shared.file_size();
        let record_limits = RecordLimits::default();
        let source_bytes = usize::try_from(source_length).map_err(|_error| {
            Error::ResourceLimit("PPT package size exceeds this platform".to_string())
        })?;
        if source_bytes > record_limits.max_package_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT package size {source_bytes} exceeds limit {}",
                record_limits.max_package_bytes
            )));
        }
        Ok(Some(Self {
            shared: Arc::new(shared),
            record_limits,
        }))
    }

    /// Open a filesystem-backed PPT without slurping the complete artifact.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path_with_limits(path, RecordLimits::default())
    }

    /// Open a filesystem-backed PPT with explicit finite limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(
        path: impl AsRef<Path>,
        record_limits: RecordLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits(Arc::new(FileSource::open(path)?), record_limits)
    }

    /// Source version captured by the validated CFB index.
    ///
    /// The source is checked before the version is returned.
    pub fn source_version(&self) -> Result<SourceVersion> {
        Ok(self.shared.source_version()?)
    }

    /// Physical source length captured by the validated CFB index.
    #[must_use]
    pub fn len(&self) -> u64 {
        self.shared.file_size()
    }

    /// Whether the validated positional source is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Read standard OLE metadata through the validated positional source.
    ///
    /// The projection uses the same property-set grammar as the ordinary
    /// cursor-backed package. Stream reads retain the shared source's version
    /// checks and do not require materializing an eager `OleFile`.
    pub fn metadata(&self) -> std::result::Result<Metadata, OleError> {
        <SharedOleFile as SharedPropertySetReader>::get_metadata(self.shared.as_ref())
    }

    /// Limits inherited by [`Self::presentation`].
    #[must_use]
    pub const fn record_limits(&self) -> RecordLimits {
        self.record_limits
    }

    /// Build a presentation using the default password options.
    ///
    /// Required native streams are materialized and parsed using the same
    /// semantic pipeline as [`super::Package::presentation`].  `Pictures` is
    /// only represented by its validated descriptor.
    pub fn presentation(&self) -> Result<Presentation> {
        self.presentation_with_options(OpenOptions::default())
    }

    /// Build a presentation with password-to-open options.
    pub fn presentation_with_options(&self, options: OpenOptions<'_>) -> Result<Presentation> {
        Presentation::from_shared_with_options(
            Arc::clone(&self.shared),
            options,
            self.record_limits,
        )
    }

    /// Build a presentation with stricter record limits.
    pub fn presentation_with_limits(&self, limits: RecordLimits) -> Result<Presentation> {
        self.presentation_with_options_and_limits(OpenOptions::default(), limits)
    }

    /// Build a presentation with password options and stricter limits.
    pub fn presentation_with_options_and_limits(
        &self,
        options: OpenOptions<'_>,
        limits: RecordLimits,
    ) -> Result<Presentation> {
        Presentation::from_shared_with_options(
            Arc::clone(&self.shared),
            options,
            limits.constrained_by(self.record_limits),
        )
    }
}

fn shared_is_powerpoint(shared: &SharedOleFile) -> bool {
    shared.exists(&["PowerPoint Document"]) || shared.exists(&["Current User"])
}
