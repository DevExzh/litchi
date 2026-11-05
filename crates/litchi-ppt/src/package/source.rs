//! Positional, source-backed access to legacy PowerPoint packages.
//!
//! The eager [`crate::raw::Reader`] remains the low-level path for `Read + Seek`
//! callers. This owner adapts an immutable positional source to
//! the native CFB reader and lets the presentation defer the optional
//! `Pictures` stream until an image query is made.

use super::{Error, OpenOptions, Result, SourceLimits};
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
    source_limits: SourceLimits,
    leniency: crate::leniency::Leniency,
    #[cfg(feature = "vba-inspection")]
    vba_project_limits: crate::VbaProjectLimits,
}

impl std::fmt::Debug for SourceBackedPackage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedPackage")
            .field("file_size", &self.shared.file_size())
            .field("source_limits", &self.source_limits)
            .finish_non_exhaustive()
    }
}

impl SourceBackedPackage {
    /// Adopt a compound-file index that a publisher already opened and
    /// validated over a composed target (the source-backed text family,
    /// record 0357a), under the source package's policies.
    pub(crate) fn from_validated_shared(
        shared: Arc<SharedOleFile>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let source_limits = options.source_limits_value();
        super::preflight_shared_ole_file(&shared, source_limits)?;
        Ok(Self {
            shared,
            source_limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
        })
    }

    /// Open a positional source with the default finite PPT limits.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, OpenOptions::default())
    }

    /// Open a positional source under explicit open options.
    ///
    /// The package keeps the options' source limits, leniency, and VBA
    /// inspection profile; every presentation it builds inherits them. The
    /// CFB index is validated before this method returns. Stream payloads are
    /// not read by the package constructor.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let cfb_limit = source_limits
            .max_package_bytes
            .min(usize::try_from(SharedOleFileLimits::MAX_INPUT_BYTES).unwrap_or(usize::MAX));
        if cfb_limit == 0 {
            return Err(Error::InvalidLimit {
                kind: super::SourceLimitKind::InputBytes,
                value: 0,
                maximum: SourceLimits::MAX_INPUT_BYTES,
            });
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
        let source_bytes = usize::try_from(source_length).map_err(|_error| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if source_bytes > source_limits.max_package_bytes {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: source_bytes,
                maximum: source_limits.max_package_bytes,
            });
        }
        let cfb_limit = u64::try_from(cfb_limit).map_err(|_error| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if source_length > cfb_limit {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: usize::try_from(source_length).unwrap_or(usize::MAX),
                maximum: usize::try_from(cfb_limit).unwrap_or(source_limits.max_package_bytes),
            });
        }
        let cfb_limits = SharedOleFileLimits::new(cfb_limit)?;
        let shared = SharedOleFile::open(source, cfb_limits)?;
        super::preflight_shared_ole_file(&shared, source_limits)?;
        Ok(Self {
            shared: Arc::new(shared),
            source_limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
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
        Self::from_shared_if_powerpoint_with_options(shared, OpenOptions::default())
    }

    /// Adopt an already validated positional CFB owner under explicit open
    /// options when it contains native PowerPoint streams.
    pub fn from_shared_if_powerpoint_with_options(
        shared: SharedOleFile,
        options: OpenOptions<'_>,
    ) -> Result<Option<Self>> {
        let source_limits = options.source_limits_value();
        shared.source_version()?;
        if !shared_is_powerpoint(&shared) {
            return Ok(None);
        }
        let source_length = shared.file_size();
        let source_bytes = usize::try_from(source_length).map_err(|_error| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if source_bytes > source_limits.max_package_bytes {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: source_bytes,
                maximum: source_limits.max_package_bytes,
            });
        }
        super::preflight_shared_ole_file(&shared, source_limits)?;
        Ok(Some(Self {
            shared: Arc::new(shared),
            source_limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
        }))
    }

    /// Open a filesystem-backed PPT without slurping the complete artifact.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a filesystem-backed PPT under explicit open options.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions<'_>) -> Result<Self> {
        Self::from_read_at_with_options(Arc::new(FileSource::open(path)?), options)
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
    pub fn metadata(&self) -> std::result::Result<Metadata, litchi_ole_common::Error> {
        <SharedOleFile as SharedPropertySetReader>::get_metadata(self.shared.as_ref())
    }

    /// Limits inherited by [`Self::presentation`].
    #[must_use]
    pub const fn source_limits(&self) -> SourceLimits {
        self.source_limits
    }

    /// Leniency inherited by [`Self::presentation`].
    #[must_use]
    pub const fn leniency(&self) -> crate::leniency::Leniency {
        self.leniency
    }

    /// Build a presentation under the package's own options and no password.
    ///
    /// Required native streams are materialized and parsed using the same
    /// semantic pipeline as [`crate::raw::Reader::presentation`]. `Pictures` is
    /// only represented by its validated descriptor.
    pub fn presentation(&self) -> Result<Presentation> {
        let options = OpenOptions::default().leniency(self.leniency);
        #[cfg(feature = "vba-inspection")]
        let options = options.vba_project_limits(self.vba_project_limits);
        self.presentation_with_options(options)
    }

    /// Build a presentation with per-call open options, such as the password
    /// of an encrypted presentation.
    ///
    /// Explicit source limits in `options` can only tighten the package's
    /// own; the options' leniency replaces the package's for this call.
    pub fn presentation_with_options(&self, options: OpenOptions<'_>) -> Result<Presentation> {
        #[cfg(feature = "vba-inspection")]
        let options = options.inherit_vba_project_limits(self.vba_project_limits);
        let source_limits = if options.has_explicit_source_limits() {
            self.source_limits
                .constrained_by(options.source_limits_value())
        } else {
            self.source_limits
        };
        Presentation::from_shared_with_options(Arc::clone(&self.shared), options, source_limits)
    }
}

impl SourceBackedPackage {
    /// Builds the source-backed text editor over this package's validated CFB
    /// owner, without indexing the container again.
    ///
    /// The editor inherits this package's record limits. It runs the checks
    /// [`crate::text_edit::SourceSnapshot::open_with_options`] runs after its
    /// own index (the protected-container refusal and the empty splice plan
    /// that fingerprints the complete artifact), so building it costs that
    /// open minus the index, paid here rather than when the package was
    /// opened.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is signed, encrypted or DRM
    /// protected, when its source changed, or when it exceeds the limits.
    pub fn text_edit_snapshot(&self) -> crate::text_edit::Result<crate::text_edit::SourceSnapshot> {
        crate::text_edit::SourceSnapshot::from_shared(
            Arc::clone(&self.shared),
            crate::text_edit::SourceBackedOptions {
                source_limits: self.source_limits,
                ..crate::text_edit::SourceBackedOptions::default()
            },
        )
    }

    /// Borrow the validated compound-file view this package reads from.
    ///
    /// This is the view opened over the positional source: its directory and
    /// allocation tables are not parsed again, and its reads observe the
    /// same source-version fence.
    #[must_use]
    pub fn shared_ole_file(&self) -> &SharedOleFile {
        &self.shared
    }

    /// Clone the validated compound-file view retained by this source owner.
    ///
    /// This is a read-only handoff for a canonical owner that already has a
    /// source-backed package. The caller must perform its own source-version
    /// check before adopting the returned view and before each operation that
    /// depends on source freshness.
    #[must_use]
    pub fn shared_ole_file_arc(&self) -> Arc<SharedOleFile> {
        Arc::clone(&self.shared)
    }
}

fn shared_is_powerpoint(shared: &SharedOleFile) -> bool {
    shared.exists(&["PowerPoint Document"]) || shared.exists(&["Current User"])
}
