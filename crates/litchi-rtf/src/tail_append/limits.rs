//! Checked policies for logical-tail transactions.
use super::{Error, LimitKind};

/// Finite append and wire bounds. Zero explicitly denies the selected resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SemanticLimits {
    pub(super) max_paragraphs: usize,
    pub(super) max_runs: usize,
    pub(super) max_input_bytes: usize,
    pub(super) max_inserted_bytes: usize,
    pub(super) max_output_bytes: usize,
    pub(super) max_patch_bytes: usize,
}

impl SemanticLimits {
    /// Largest addressable bound.
    pub const MAXIMUM: usize = isize::MAX as usize;
    /// Start with production defaults.
    #[must_use]
    pub fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder::default()
    }
    /// Start a builder retaining this profile.
    #[must_use]
    pub fn into_builder(self) -> SemanticLimitsBuilder {
        SemanticLimitsBuilder { limits: self }
    }
    /// Maximum paragraphs.
    #[must_use]
    pub const fn max_paragraphs(self) -> usize {
        self.max_paragraphs
    }
    /// Maximum runs.
    #[must_use]
    pub const fn max_runs(self) -> usize {
        self.max_runs
    }
    /// Maximum input bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }
    /// Maximum inserted bytes.
    #[must_use]
    pub const fn max_inserted_bytes(self) -> usize {
        self.max_inserted_bytes
    }
    /// Maximum output bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    /// Maximum patch bytes.
    #[must_use]
    pub const fn max_patch_bytes(self) -> usize {
        self.max_patch_bytes
    }
}
impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            max_paragraphs: 256,
            max_runs: 4096,
            max_input_bytes: 256 * 1024,
            max_inserted_bytes: 512 * 1024,
            max_output_bytes: 64 * 1024 * 1024,
            max_patch_bytes: 2 * 1024 * 1024,
        }
    }
}

/// Consuming checked profile builder.
#[derive(Debug, Clone, Copy, Default)]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}
impl SemanticLimitsBuilder {
    /// Select the paragraphs bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_paragraphs(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Paragraphs,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_paragraphs = value;
        Ok(self)
    }
    /// Select the runs bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_runs(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Runs,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_runs = value;
        Ok(self)
    }
    /// Select the input bytes bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::InputBytes,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_input_bytes = value;
        Ok(self)
    }
    /// Select the inserted bytes bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_inserted_bytes(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::InsertedBytes,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_inserted_bytes = value;
        Ok(self)
    }
    /// Select the output bytes bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_output_bytes(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::OutputBytes,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_output_bytes = value;
        Ok(self)
    }
    /// Select the patch bytes bound; zero denies this resource.
    ///
    /// # Errors
    /// Returns a typed error above the addressable ceiling.
    pub fn max_patch_bytes(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAXIMUM {
            return Err(Error::InvalidLimit {
                kind: LimitKind::PatchBytes,
                value,
                maximum: SemanticLimits::MAXIMUM,
            });
        }
        self.limits.max_patch_bytes = value;
        Ok(self)
    }
    /// Finish a profile whose setters have checked every bound.
    ///
    /// # Errors
    /// All values are admitted during setter calls.
    pub fn build(self) -> Result<SemanticLimits, Error> {
        Ok(self.limits)
    }
}

/// Per-call policy for logical-tail edits.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    semantic_limits: SemanticLimits,
}
impl EditOptions {
    /// Production-safe defaults.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
    /// Select a checked append profile.
    #[must_use]
    pub const fn semantic_limits(mut self, limits: SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }
    /// Return the admitted profile.
    #[must_use]
    pub const fn semantic_limits_value(self) -> SemanticLimits {
        self.semantic_limits
    }
}
