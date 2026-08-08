//! Explicit, versioned compatibility profiles for nonconforming XLS inputs.
//!
//! Profiles are selected by the caller before parsing. They never enable
//! formula evaluation, external access, macro execution, or noncanonical
//! writing.

/// A narrow workbook-open compatibility contract.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum CompatibilityProfile {
    /// Accept only canonical MS-XLS structures.
    #[default]
    Strict,
    /// Preserve a Formula `fShrFmla` flag without the leading `PtgExp` token.
    ///
    /// `[MS-XLS]` 2.4.127 requires the token. This V1 profile exists for the
    /// real `ConditionalFormattingSamples.xls` fixture mirrored in the Apache
    /// POI spreadsheet test-data corpus. Local provenance does not identify
    /// the original producing application, so no producer is inferred.
    SharedFormulaFlagWithoutPtgExpV1,
}

impl CompatibilityProfile {
    /// Stable profile name suitable for diagnostics.
    pub const fn name(self) -> &'static str {
        match self {
            Self::Strict => "strict",
            Self::SharedFormulaFlagWithoutPtgExpV1 => "shared-formula-flag-without-ptg-exp-v1",
        }
    }

    /// Evidence provenance for this profile.
    pub const fn provenance(self) -> &'static str {
        match self {
            Self::Strict => "MS-XLS canonical structures",
            Self::SharedFormulaFlagWithoutPtgExpV1 => {
                "Apache POI spreadsheet test-data corpus mirror; original producer unknown"
            },
        }
    }

    pub(crate) const fn preserves_shared_formula_flag_without_ptg_exp(self) -> bool {
        matches!(self, Self::SharedFormulaFlagWithoutPtgExpV1)
    }
}
