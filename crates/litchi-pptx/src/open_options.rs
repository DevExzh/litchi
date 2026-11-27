//! Configuration for PPTX ingress.

use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceLimits};
use std::sync::Arc;

/// The semantic codec policies selected for an owning or source-backed
/// presentation.
///
/// The bundle is retained only when a caller overrides one of the codec
/// defaults.  This keeps the ordinary owner footprint unchanged while making
/// a selected policy available to every eager, source-backed, and edit path.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SemanticLimits {
    shape: crate::shape::Limits,
    transition: crate::transition::Limits,
    laser: crate::laser::Limits,
    media_parts: crate::media_parts::Limits,
    tag: crate::tag::Limits,
    notes: crate::notes::Limits,
    animations: crate::animations::Limits,
    comments: crate::comments::Limits,
    changes: crate::presentation_properties::metadata::changes::Limits,
    revision: crate::presentation_properties::metadata::revision::Limits,
    slide_sync: crate::presentation_properties::metadata::slide_sync::Limits,
    tracks: crate::presentation_properties::metadata::tracks::Limits,
    events: crate::presentation_properties::metadata::events::Limits,
}

pub(crate) type SemanticLimitsRef = Arc<SemanticLimits>;

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::defaults()
    }
}

impl SemanticLimits {
    const fn defaults() -> Self {
        Self {
            shape: crate::shape::Limits::DEFAULT,
            transition: crate::transition::Limits::DEFAULT,
            laser: crate::laser::Limits::DEFAULT,
            media_parts: crate::media_parts::Limits::DEFAULT,
            tag: crate::tag::Limits::DEFAULT,
            notes: crate::notes::Limits::DEFAULT,
            animations: crate::animations::Limits::DEFAULT,
            comments: crate::comments::Limits::DEFAULT,
            changes: crate::presentation_properties::metadata::changes::Limits::DEFAULT,
            revision: crate::presentation_properties::metadata::revision::Limits::DEFAULT,
            slide_sync: crate::presentation_properties::metadata::slide_sync::Limits::DEFAULT,
            tracks: crate::presentation_properties::metadata::tracks::Limits::DEFAULT,
            events: crate::presentation_properties::metadata::events::Limits::DEFAULT,
        }
    }

    pub(crate) const fn shape(self) -> crate::shape::Limits {
        self.shape
    }

    pub(crate) const fn transition(self) -> crate::transition::Limits {
        self.transition
    }

    pub(crate) const fn laser(self) -> crate::laser::Limits {
        self.laser
    }

    pub(crate) const fn media_parts(self) -> crate::media_parts::Limits {
        self.media_parts
    }

    pub(crate) const fn tag(self) -> crate::tag::Limits {
        self.tag
    }

    pub(crate) const fn notes(self) -> crate::notes::Limits {
        self.notes
    }

    pub(crate) const fn animations(self) -> crate::animations::Limits {
        self.animations
    }

    pub(crate) const fn comments(self) -> crate::comments::Limits {
        self.comments
    }

    pub(crate) const fn changes(self) -> crate::presentation_properties::metadata::changes::Limits {
        self.changes
    }

    pub(crate) const fn revision(
        self,
    ) -> crate::presentation_properties::metadata::revision::Limits {
        self.revision
    }

    pub(crate) const fn slide_sync(
        self,
    ) -> crate::presentation_properties::metadata::slide_sync::Limits {
        self.slide_sync
    }

    pub(crate) const fn tracks(self) -> crate::presentation_properties::metadata::tracks::Limits {
        self.tracks
    }

    pub(crate) const fn events(self) -> crate::presentation_properties::metadata::events::Limits {
        self.events
    }

    const fn with_shape(mut self, limits: crate::shape::Limits) -> Self {
        self.shape = limits;
        self
    }

    const fn with_transition(mut self, limits: crate::transition::Limits) -> Self {
        self.transition = limits;
        self
    }

    const fn with_laser(mut self, limits: crate::laser::Limits) -> Self {
        self.laser = limits;
        self
    }

    const fn with_media_parts(mut self, limits: crate::media_parts::Limits) -> Self {
        self.media_parts = limits;
        self
    }

    const fn with_tag(mut self, limits: crate::tag::Limits) -> Self {
        self.tag = limits;
        self
    }

    const fn with_notes(mut self, limits: crate::notes::Limits) -> Self {
        self.notes = limits;
        self
    }

    const fn with_animations(mut self, limits: crate::animations::Limits) -> Self {
        self.animations = limits;
        self
    }

    const fn with_comments(mut self, limits: crate::comments::Limits) -> Self {
        self.comments = limits;
        self
    }

    const fn with_changes(
        mut self,
        limits: crate::presentation_properties::metadata::changes::Limits,
    ) -> Self {
        self.changes = limits;
        self
    }

    const fn with_revision(
        mut self,
        limits: crate::presentation_properties::metadata::revision::Limits,
    ) -> Self {
        self.revision = limits;
        self
    }

    const fn with_slide_sync(
        mut self,
        limits: crate::presentation_properties::metadata::slide_sync::Limits,
    ) -> Self {
        self.slide_sync = limits;
        self
    }

    const fn with_tracks(
        mut self,
        limits: crate::presentation_properties::metadata::tracks::Limits,
    ) -> Self {
        self.tracks = limits;
        self
    }

    const fn with_events(
        mut self,
        limits: crate::presentation_properties::metadata::events::Limits,
    ) -> Self {
        self.events = limits;
        self
    }

    fn into_ref(self) -> Option<SemanticLimitsRef> {
        (self != Self::defaults()).then(|| Arc::new(self))
    }
}

#[inline]
pub(crate) fn selected(value: Option<&SemanticLimitsRef>) -> SemanticLimits {
    value.map(|value| **value).unwrap_or_default()
}

/// Policies for owning and positional-source PPTX ingress.
///
/// One envelope serves every configured open route:
///
/// - the owning [`crate::Package`] (`open_with_options`,
///   `from_reader_with_options`, `from_bytes_with_options` and, with the
///   `encryption` feature, the `*_with_password_and_options` routes);
/// - the positional-source owners
///   [`crate::SourceBackedPresentation::open_with_options`] and
///   [`crate::source::Package::open_with_options`].
///
/// Source limits bound the clear OPC archive. Cache and execution settings
/// apply only to positional-source owners; owning constructors reject them
/// with [`crate::Error::UnsupportedOpenPolicy`] before any input is read,
/// because an owning package retains no deferred payload cache or execution
/// session to charge them to. With the `encryption` feature, encryption
/// limits independently bound the outer encrypted container and decryption
/// work on password-bearing routes. Passwords are arguments of those routes
/// and are never retained here.
///
/// # Examples
///
/// ```
/// use litchi_pptx::{OpenOptions, Package, SourceLimits};
/// # fn example(bytes: &[u8]) -> litchi_pptx::Result<Package> {
/// let limits = SourceLimits::builder().max_input_bytes(32 * 1024 * 1024)?.build()?;
/// Package::from_bytes_with_options(bytes, OpenOptions::default().source_limits(limits))
/// # }
/// ```
///
/// Defaults retain the standard finite source and encryption policies.
/// Setters consume already validated policy values and perform no I/O.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    pub(crate) semantic_limits: Option<SemanticLimitsRef>,
    pub(crate) cache_limits: Option<SourceCacheLimits>,
    pub(crate) execution_context: Option<ExecutionContext>,
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: litchi_crypto::ooxml::Limits,
}

impl OpenOptions {
    /// Select the validated physical limits for the clear OPC archive.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select the bounded shape-scene policy retained by the opened owner.
    #[must_use]
    pub fn shape_limits(mut self, limits: crate::shape::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_shape(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded transition policy retained by the opened owner.
    #[must_use]
    pub fn transition_limits(mut self, limits: crate::transition::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_transition(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded laser-trace policy retained by the opened owner.
    #[must_use]
    pub fn laser_limits(mut self, limits: crate::laser::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_laser(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded media-part policy retained by the opened owner.
    #[must_use]
    pub fn media_parts_limits(mut self, limits: crate::media_parts::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_media_parts(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded Changes Information policy retained by the opened owner.
    #[must_use]
    pub fn changes_limits(
        mut self,
        limits: crate::presentation_properties::metadata::changes::Limits,
    ) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_changes(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded programmable-tag policy retained by the opened
    /// owner.
    #[must_use]
    pub fn tag_limits(mut self, limits: crate::tag::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_tag(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded revision policy retained by the opened owner.
    #[must_use]
    pub fn revision_limits(
        mut self,
        limits: crate::presentation_properties::metadata::revision::Limits,
    ) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_revision(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded speaker-notes policy retained by the opened owner.
    #[must_use]
    pub fn notes_limits(mut self, limits: crate::notes::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_notes(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded slide-sync policy retained by the opened owner.
    #[must_use]
    pub fn slide_sync_limits(
        mut self,
        limits: crate::presentation_properties::metadata::slide_sync::Limits,
    ) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_slide_sync(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded animation-timing policy retained by the opened
    /// owner.
    #[must_use]
    pub fn animations_limits(mut self, limits: crate::animations::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_animations(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded media-track policy retained by the opened owner.
    #[must_use]
    pub fn tracks_limits(
        mut self,
        limits: crate::presentation_properties::metadata::tracks::Limits,
    ) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_tracks(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded classic-comment policy retained by the opened
    /// owner.
    #[must_use]
    pub fn comments_limits(mut self, limits: crate::comments::Limits) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_comments(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select the bounded slide-show-events policy retained by the opened owner.
    #[must_use]
    pub fn events_limits(
        mut self,
        limits: crate::presentation_properties::metadata::events::Limits,
    ) -> Self {
        let profiles = selected(self.semantic_limits.as_ref()).with_events(limits);
        self.semantic_limits = profiles.into_ref();
        self
    }

    /// Select a validated finite payload-cache policy for positional-source
    /// ingress.
    #[must_use]
    pub const fn cache_limits(mut self, limits: SourceCacheLimits) -> Self {
        self.cache_limits = Some(limits);
        self
    }

    /// Use the caller's execution context for positional-source opening and
    /// deferred reads. Cloned options share its budget and cancellation token.
    #[must_use]
    pub fn execution_context(mut self, context: ExecutionContext) -> Self {
        self.execution_context = Some(context);
        self
    }

    /// Select the outer encryption policy for password-bearing ingress.
    ///
    /// The encryption owner validates this policy during decryption. It does
    /// not replace the clear archive's source limits and has no effect on
    /// plaintext ingress.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption_limits(mut self, limits: litchi_crypto::ooxml::Limits) -> Self {
        self.encryption_limits = limits;
        self
    }

    /// Refuse lazy-ingress policies on an owning route.
    pub(crate) fn check_owning(&self) -> crate::Result<()> {
        let policy = if self.execution_context.is_some() {
            Some(OpenPolicy::ExecutionContext)
        } else if self.cache_limits.is_some() {
            Some(OpenPolicy::CacheLimits)
        } else {
            None
        };
        match policy {
            Some(policy) => Err(crate::Error::UnsupportedOpenPolicy { policy }),
            None => Ok(()),
        }
    }
}

/// A resource policy that requires positional-source PPTX ownership.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenPolicy {
    /// Retention in the finite deferred payload cache.
    CacheLimits,
    /// Hierarchical budgets and cooperative cancellation.
    ExecutionContext,
}

impl std::fmt::Display for OpenPolicy {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::CacheLimits => "payload cache limits",
            Self::ExecutionContext => "execution context",
        })
    }
}
