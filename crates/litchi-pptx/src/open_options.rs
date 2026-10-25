//! Configuration for PPTX ingress.

use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceLimits};

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
}

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
}

#[inline]
pub(crate) fn selected(value: Option<&SemanticLimits>) -> SemanticLimits {
    value.copied().unwrap_or_default()
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
    pub(crate) semantic_limits: Option<SemanticLimits>,
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

    /// Select the bounded media-part policy retained by the opened owner.
    #[must_use]
    pub const fn media_parts_limits(mut self, limits: crate::media_parts::Limits) -> Self {
        let profiles = match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        };
        self.semantic_limits = Some(profiles.with_media_parts(limits));
        self
    }

    /// Select the bounded programmable-tag policy retained by the opened
    /// owner.
    #[must_use]
    pub const fn tag_limits(mut self, limits: crate::tag::Limits) -> Self {
        let profiles = match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        };
        self.semantic_limits = Some(profiles.with_tag(limits));
        self
    }

    /// Select the bounded speaker-notes policy retained by the opened owner.
    #[must_use]
    pub const fn notes_limits(mut self, limits: crate::notes::Limits) -> Self {
        let profiles = match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        };
        self.semantic_limits = Some(profiles.with_notes(limits));
        self
    }

    /// Select the bounded animation-timing policy retained by the opened
    /// owner.
    #[must_use]
    pub const fn animations_limits(mut self, limits: crate::animations::Limits) -> Self {
        let profiles = match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        };
        self.semantic_limits = Some(profiles.with_animations(limits));
        self
    }

    /// Select the bounded classic-comment policy retained by the opened
    /// owner.
    #[must_use]
    pub const fn comments_limits(mut self, limits: crate::comments::Limits) -> Self {
        let profiles = match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        };
        self.semantic_limits = Some(profiles.with_comments(limits));
        self
    }

    pub(crate) const fn semantic_limits_value(&self) -> SemanticLimits {
        match self.semantic_limits {
            Some(profiles) => profiles,
            None => SemanticLimits::defaults(),
        }
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
