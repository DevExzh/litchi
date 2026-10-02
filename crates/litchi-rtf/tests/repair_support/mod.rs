//! Shared assertion for malformed-input tests under the tolerance contract.
//!
//! A parameter-shape deviation (a parameter on a flag, a nonbinary toggle,
//! or a missing value) is repaired the way Word repairs it and recorded,
//! rather than rejected. Malformed-input tests therefore accept either
//! outcome for such sources: a hard error, or a successful parse whose only
//! repairs are recorded parameter defects.

use litchi_rtf::leniency::{Defect, ToleranceReport};

pub(crate) trait Outcome {
    fn tolerance_report(&self) -> Option<&ToleranceReport>;
}

impl Outcome for litchi_rtf::Result<litchi_rtf::RtfDocument<'_>> {
    fn tolerance_report(&self) -> Option<&ToleranceReport> {
        self.as_ref()
            .ok()
            .map(litchi_rtf::RtfDocument::tolerance_report)
    }
}

impl Outcome for litchi_rtf::Result<litchi_rtf::Document> {
    fn tolerance_report(&self) -> Option<&ToleranceReport> {
        self.as_ref()
            .ok()
            .map(litchi_rtf::Document::tolerance_report)
    }
}

/// Whether the source was rejected, or accepted only through recorded
/// Word-compatible parameter repairs.
pub(crate) fn rejected_or_parameter_repaired(outcome: &impl Outcome) -> bool {
    outcome.tolerance_report().is_none_or(|report| {
        !report.is_clean()
            && report
                .iter()
                .all(|tolerated| matches!(tolerated.defect, Defect::Parameter(_)))
    })
}
