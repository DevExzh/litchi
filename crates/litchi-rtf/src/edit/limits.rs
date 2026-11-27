//! Checked semantic policies for primary RTF edits.

use super::Error;

/// Semantic admission dimension for a primary edit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Number of staged semantic operations.
    Operations,
    /// Aggregate authored replacement bytes retained by staged operations.
    ReplacementBytes,
}

/// Checked, finite primary-edit limits.
///
/// ```
/// use litchi_rtf::{EditOptions, Package, SemanticLimits};
/// let limits = SemanticLimits::builder().max_operations(1)?.build()?;
/// let package = Package::parse(r"{\rtf1 Before}")?;
/// let mut edit = package.edit_with_options(EditOptions::new().semantic_limits(limits));
/// edit.replace_paragraph_text(litchi_core::Position::new(0), "After")?;
/// assert_eq!(edit.commit()?.package().text(), "After");
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SemanticLimits {
    max_operations: usize,
    max_replacement_bytes: usize,
}

impl SemanticLimits {
    // A materialized slice of non-zero-sized operations cannot exceed the
    // addressable ceiling. Composition has already admitted these operations.
    pub(super) fn for_operations(
        operations: &[super::Operation],
        source_limits: crate::SourceLimits,
    ) -> Self {
        Self {
            max_operations: operations.len(),
            max_replacement_bytes: source_limits.max_input_bytes(),
        }
    }

    /// Ordinary primary-edit operation budget.
    pub const DEFAULT_MAX_OPERATIONS: usize = 256;
    /// Default aggregate replacement-payload budget (512 MiB).
    pub const DEFAULT_MAX_REPLACEMENT_BYTES: usize = 512 * 1024 * 1024;
    /// Largest addressable replacement-payload budget.
    pub const MAX_REPLACEMENT_BYTES: usize = isize::MAX as usize;

    /// Largest addressable operation budget.
    pub const MAX_OPERATIONS: usize = isize::MAX as usize;

    /// Start a checked profile builder with production defaults.
    #[must_use]
    pub fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder::default()
    }

    /// Maximum aggregate authored replacement bytes. Zero denies nonempty payloads.
    ///
    /// Text and metadata count UTF-8 bytes; paragraph insertion includes its
    /// separator, and pictures count hexadecimal wire transport including whitespace.
    /// This is independent of the retained source/output ceiling and does not
    /// measure inverse snapshots or total process memory.
    #[must_use]
    pub const fn max_replacement_bytes(self) -> usize {
        self.max_replacement_bytes
    }

    /// Maximum number of semantic operations. Zero permits an empty edit only.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.max_operations
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            max_operations: Self::DEFAULT_MAX_OPERATIONS,
            max_replacement_bytes: Self::DEFAULT_MAX_REPLACEMENT_BYTES,
        }
    }
}

/// Consuming builder for checked semantic admission limits.
#[derive(Clone, Copy, Debug, Default)]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}

impl SemanticLimitsBuilder {
    /// Select the operation ceiling; zero explicitly denies staged operations.
    ///
    /// # Errors
    /// Returns [`Error::InvalidLimit`] above the addressable maximum.
    pub fn max_operations(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAX_OPERATIONS {
            return Err(Error::InvalidLimit {
                kind: SemanticLimitKind::Operations,
                value,
                maximum: SemanticLimits::MAX_OPERATIONS,
            });
        }
        self.limits.max_operations = value;
        Ok(self)
    }

    /// Select the aggregate authored replacement-byte ceiling.
    ///
    /// # Errors
    /// Returns [`Error::InvalidLimit`] above the addressable maximum.
    pub fn max_replacement_bytes(mut self, value: usize) -> Result<Self, Error> {
        if value > SemanticLimits::MAX_REPLACEMENT_BYTES {
            return Err(Error::InvalidLimit {
                kind: SemanticLimitKind::ReplacementBytes,
                value,
                maximum: SemanticLimits::MAX_REPLACEMENT_BYTES,
            });
        }
        self.limits.max_replacement_bytes = value;
        Ok(self)
    }

    /// Finish the admitted profile.
    ///
    /// # Errors
    /// Every dimension is validated by its setter; the checked build result
    /// keeps construction uniform as dimensions are added.
    pub fn build(self) -> Result<SemanticLimits, Error> {
        Ok(self.limits)
    }
}

/// Per-call options for a primary immutable edit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EditOptions {
    semantic_limits: SemanticLimits,
}

impl EditOptions {
    /// Start with the finite production profile.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the checked semantic profile for this edit.
    #[must_use]
    pub const fn semantic_limits(mut self, value: SemanticLimits) -> Self {
        self.semantic_limits = value;
        self
    }

    /// The selected semantic profile.
    #[must_use]
    pub const fn semantic_limits_value(self) -> SemanticLimits {
        self.semantic_limits
    }
}
