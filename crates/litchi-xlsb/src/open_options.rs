//! Configuration for XLSB ingress.

use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceLimits};

use crate::external_link::ExternalLinkLimits;
use crate::leniency::Leniency;
use crate::workbook::DrawingLoadPolicy;

/// Policies for opening an XLSB workbook.
///
/// One envelope serves every configured open route:
///
/// - the owning [`crate::Package`] (`open_with_options`,
///   `from_bytes_with_options`, `from_reader_with_options`,
///   `from_opc_with_options`) and the eager [`crate::Workbook`] projection
///   constructors (`new_with_options`, `from_opc_package_with_options`);
/// - the positional-source owner [`crate::SourceBackedWorkbook`]
///   (`from_reader_with_options`, `from_read_at_with_options`).
///
/// Source limits bound the clear OPC archive and external-link limits bound
/// the semantic external-link catalog. The leniency selects the opt-in
/// repairs described in [`crate::leniency`]. Cache and execution settings
/// apply only to the positional-source owner, and skipping drawing decoding
/// applies only to the eager routes; a route refuses a policy it cannot
/// honour with [`crate::package::PackageError::UnsupportedOpenPolicy`]
/// before any input is read.
///
/// ```
/// use litchi_xlsb::{OpenOptions, Package, SourceLimits};
/// # fn example(bytes: &[u8]) -> litchi_xlsb::package::PackageResult<Package> {
/// let limits = SourceLimits::builder().max_input_bytes(32 * 1024 * 1024)?.build()?;
/// Package::from_bytes_with_options(bytes, OpenOptions::new().source_limits(limits))
/// # }
/// ```
///
/// Defaults retain the standard finite source and external-link policies,
/// the strict leniency and eager drawing decoding. Setters consume already
/// validated policy values and perform no I/O.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    /// Limits for detached BIFF12 record and string codecs reached by this
    /// owner.  The profile is retained so lazy worksheet reads and edits do
    /// not silently fall back to the raw kernel default.
    pub(crate) raw_limits: crate::raw::Limits,
    pub(crate) external_link_limits: ExternalLinkLimits,
    pub(crate) leniency: Leniency,
    pub(crate) skip_drawings: bool,
    pub(crate) cache_limits: Option<SourceCacheLimits>,
    pub(crate) execution_context: Option<ExecutionContext>,
    pub(crate) semantic_limits: crate::SemanticLimits,
}

impl OpenOptions {
    /// The default ingress policies.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the validated physical limits for the clear OPC archive.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select the finite BIFF12 payload and wide-string limits retained by
    /// this owner for every raw record codec it reaches.
    #[must_use]
    pub const fn raw_limits(mut self, limits: crate::raw::Limits) -> Self {
        self.raw_limits = limits;
        self
    }

    /// Select the validated semantic limits for the external-link catalog.
    #[must_use]
    pub const fn external_link_limits(mut self, limits: ExternalLinkLimits) -> Self {
        self.external_link_limits = limits;
        self
    }

    /// Select the opt-in repairs applied while the workbook is read.
    #[must_use]
    pub const fn leniency(mut self, leniency: Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Whether eager routes decode standard SpreadsheetDrawing parts at open
    /// (the default).
    ///
    /// `false` keeps every drawing, chart, image and relationship part in
    /// the package for lossless publication but leaves the typed
    /// [`crate::Workbook::sheet_drawings`] inventory empty. It suits
    /// cell and catalog edits that never read drawings.
    #[must_use]
    pub const fn parse_drawings(mut self, parse: bool) -> Self {
        self.skip_drawings = !parse;
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

    /// Select the semantic budgets the opened package keeps for its family
    /// reads and edits.
    #[must_use]
    pub const fn semantic_limits(mut self, limits: crate::SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The physical limits of the clear OPC archive.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// The retained BIFF12 raw record and string policy.
    #[must_use]
    pub const fn raw_limits_value(&self) -> crate::raw::Limits {
        self.raw_limits
    }

    /// The semantic limits of the external-link catalog.
    #[must_use]
    pub const fn external_link_limits_value(&self) -> ExternalLinkLimits {
        self.external_link_limits
    }

    /// The opt-in repairs applied while the workbook is read.
    #[must_use]
    pub const fn leniency_value(&self) -> Leniency {
        self.leniency
    }

    /// Whether eager routes decode drawings at open.
    #[must_use]
    pub const fn parse_drawings_value(&self) -> bool {
        !self.skip_drawings
    }

    /// The payload-cache policy of positional-source ingress, if any.
    #[must_use]
    pub const fn cache_limits_value(&self) -> Option<SourceCacheLimits> {
        self.cache_limits
    }

    /// The caller's execution context of positional-source ingress, if any.
    #[must_use]
    pub const fn execution_context_value(&self) -> Option<&ExecutionContext> {
        self.execution_context.as_ref()
    }

    /// The semantic budgets the opened package keeps.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> crate::SemanticLimits {
        self.semantic_limits
    }

    pub(crate) const fn drawing_load_policy(&self) -> DrawingLoadPolicy {
        if self.skip_drawings {
            DrawingLoadPolicy::Skipped
        } else {
            DrawingLoadPolicy::Eager
        }
    }

    /// Refuse positional-source policies on an owning route.
    pub(crate) fn check_owning(&self) -> crate::package::PackageResult<()> {
        let policy = if self.execution_context.is_some() {
            Some(OpenPolicy::ExecutionContext)
        } else if self.cache_limits.is_some() {
            Some(OpenPolicy::CacheLimits)
        } else {
            None
        };
        match policy {
            Some(policy) => Err(crate::package::PackageError::UnsupportedOpenPolicy { policy }),
            None => Ok(()),
        }
    }

    /// Refuse eager-only policies on a positional-source route.
    pub(crate) fn check_source(&self) -> crate::package::PackageResult<()> {
        if self.skip_drawings {
            return Err(crate::package::PackageError::UnsupportedOpenPolicy {
                policy: OpenPolicy::SkippedDrawings,
            });
        }
        Ok(())
    }
}

/// An open policy that only one owner kind can honour.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenPolicy {
    /// Retention in the finite deferred payload cache (positional source).
    CacheLimits,
    /// Hierarchical budgets and cooperative cancellation (positional source).
    ExecutionContext,
    /// Skipping drawing decoding at open (eager owners; the positional-source
    /// owner defers every drawing on its own).
    SkippedDrawings,
}

impl std::fmt::Display for OpenPolicy {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::CacheLimits => "payload cache limits",
            Self::ExecutionContext => "execution context",
            Self::SkippedDrawings => "skipped drawing decoding",
        })
    }
}
