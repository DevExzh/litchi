//! Cheaply shareable archive-free Keynote document snapshots.

use std::path::Path;
use std::sync::Arc;

use crate::package::{ReadError, SemanticLimits, Stats};
use crate::show::Show;

#[derive(Debug)]
struct State {
    show: Show,
    metadata: Option<litchi_core::Metadata>,
    stats: Option<Stats>,
}

/// Physical and semantic resource profiles for archive-free document ingress.
///
/// The source profile belongs to format detection rather than exact package
/// preservation, so it applies equally to complete ZIP artifacts and frozen
/// app-authored package directories. The canonical properties diagnostic has
/// the independent hard ceiling [`crate::MAX_DOCUMENT_PROPERTIES_BYTES`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentReadOptions {
    source: crate::DocumentSourceLimits,
    semantic: SemanticLimits,
}

impl DocumentReadOptions {
    /// Combine checked source-capture and semantic resource profiles.
    #[must_use]
    pub const fn new(source: crate::DocumentSourceLimits, semantic: SemanticLimits) -> Self {
        Self { source, semantic }
    }

    /// Return the bounded source-capture profile.
    #[must_use]
    pub const fn source(self) -> crate::DocumentSourceLimits {
        self.source
    }

    /// Return the semantic graph-decoding profile.
    #[must_use]
    pub const fn semantic(self) -> SemanticLimits {
        self.semantic
    }
}

/// An immutable, cheaply clonable semantic Keynote document snapshot.
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    /// Open a complete Keynote package or an app-authored package directory.
    ///
    /// This constructor eagerly freezes the source and completes the bounded
    /// semantic graph projection before publishing the archive-free snapshot.
    /// A directory contributes only its IWA index and canonical
    /// `Metadata/Properties.plist`; media, previews, exact package bytes,
    /// writing, and editing are intentionally not represented by `Document`.
    /// The properties diagnostic never exceeds
    /// [`crate::MAX_DOCUMENT_PROPERTIES_BYTES`], even when the broader source
    /// entry ceiling is larger.
    /// Use [`crate::Package`] for exact complete-package preservation.
    ///
    /// # Errors
    ///
    /// Returns an error when the source is missing, unsafe, ambiguous,
    /// encrypted, belongs to another iWork application, is malformed, changes
    /// during capture, or exceeds a checked resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, ReadError> {
        Self::open_with_options(path, DocumentReadOptions::default())
    }

    /// Open an archive-free Keynote document under explicit source and
    /// semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: DocumentReadOptions,
    ) -> Result<Self, ReadError> {
        let source = litchi_iwa_detect::PreparedSource::__from_path_with_keynote_properties(
            path,
            options.source(),
        )?
        .ok_or_else(|| {
            ReadError::InvalidFormat("source is not a recognized iWork document".to_owned())
        })?;
        if source.format() != litchi_iwa_detect::Format::Keynote {
            return Err(ReadError::NotKeynote);
        }
        crate::package::semantic_document_from_prepared_source(source, options.semantic())
    }

    /// Decode complete Keynote package bytes into an archive-free snapshot.
    ///
    /// The input is borrowed only for bounded capture. The returned document
    /// retains semantic show values and source diagnostics, but no package
    /// bytes, archive objects, protobuf messages, or native identifiers.
    /// Use [`crate::Package::from_bytes`] when exact package preservation or
    /// editing is required.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are missing a recognized Keynote root,
    /// violate a physical or semantic ceiling, or contain malformed semantic
    /// content.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, ReadError> {
        Self::from_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode complete Keynote package bytes under explicit resource limits.
    ///
    /// The source profile is translated once at this package boundary; the
    /// semantic profile is passed through unchanged. The package handle is
    /// dropped before this method returns, so the result remains archive-free.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_bytes`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: DocumentReadOptions,
    ) -> Result<Self, ReadError> {
        let source = options.source();
        let package_limits = crate::Limits::new(
            source.max_input_bytes(),
            source.max_files(),
            source.max_entry_size(),
            source.max_total_size(),
            source.max_iwa_stream_size(),
        )?;
        let package = crate::Package::from_bytes_with_options(
            bytes,
            crate::ReadOptions::new(package_limits, options.semantic()),
        )?;
        let show = package.show()?.clone();
        let metadata = package.metadata()?.ok_or_else(|| {
            ReadError::InvalidFormat("Keynote package has no semantic metadata".to_owned())
        })?;
        let stats = package.stats()?;
        Ok(Self::from_source(show, metadata, stats))
    }

    /// Create a snapshot from an already decoded semantic show.
    #[must_use]
    pub fn from_show(show: Show) -> Self {
        Self {
            state: Arc::new(State {
                show,
                metadata: None,
                stats: None,
            }),
        }
    }

    pub(crate) fn from_source(show: Show, metadata: litchi_core::Metadata, stats: Stats) -> Self {
        Self {
            state: Arc::new(State {
                show,
                metadata: Some(metadata),
                stats: Some(stats),
            }),
        }
    }

    /// Capture another cheap handle to the same snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the immutable semantic show.
    #[must_use]
    pub fn show(&self) -> &Show {
        &self.state.show
    }

    /// Borrow the slides without copying the snapshot.
    #[must_use]
    pub fn slides(&self) -> &[crate::Slide] {
        self.state.show.slides()
    }

    /// Extract rooted text in semantic presentation order.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the exact output buffer cannot be
    /// reserved.
    pub fn text(&self) -> Result<String, ReadError> {
        crate::package::semantic_text(&self.state.show)
    }

    /// Borrow source-derived metadata.
    ///
    /// The value combines semantic Show fields with the canonical properties
    /// diagnostic when that sidecar exists. `Some` denotes a validated source
    /// origin; it does not prove the optional sidecar was present.
    ///
    /// Values built with [`Self::from_show`] have no source diagnostics and
    /// return `None`.
    #[must_use]
    pub fn metadata(&self) -> Option<&litchi_core::Metadata> {
        self.state.metadata.as_ref()
    }

    /// Return source diagnostics retained during bounded ingress.
    ///
    /// Values built with [`Self::from_show`] have no physical source and
    /// return `None`.
    #[must_use]
    pub fn stats(&self) -> Option<Stats> {
        self.state.stats
    }

    /// Validate semantic invariants of the detached snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic error if the retained show settings are no
    /// longer canonical.
    pub fn validate(&self) -> crate::Result<()> {
        self.state.show.settings().validate()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn snapshots_are_send_sync_and_shareable() {
        assert_send_sync::<Document>();
        let document = Document::from_show(Show::builder().build());
        let snapshot = document.snapshot();
        assert_eq!(document.show(), snapshot.show());
    }
}
