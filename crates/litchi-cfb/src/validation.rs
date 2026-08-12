//! Non-mutating validation of CFB/OLE2 positional sources.

use crate::{OleError, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceDigest, EvidenceValue,
    IssueEvidence, IssueLocation, IssueSeverity, ReadAt, RepairAvailability, ValidateReport,
    ValidationCheck, ValidationIssue, ValidationLimits, ValidationReportError,
};
use std::{error::Error, fmt, sync::Arc};

const INGRESS_CHECK: &str = "cfb.container.ingress";

/// Failure to perform or retain a CFB validation report.
///
/// Definite format rejections are represented by issues in a successful
/// report. This error is reserved for source I/O or instability, bounded parser
/// allocation failure, and failure to construct the bounded report itself.
#[derive(Debug)]
#[non_exhaustive]
pub enum CfbValidationError {
    /// The canonical CFB ingress could not make an honest determination.
    Ingress(OleError),
    /// The requested validation-report bounds could not retain the result.
    Report(ValidationReportError),
}

impl fmt::Display for CfbValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ingress(error) => write!(formatter, "CFB validation ingress failed: {error}"),
            Self::Report(error) => write!(formatter, "CFB validation report failed: {error}"),
        }
    }
}

impl Error for CfbValidationError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Ingress(error) => Some(error),
            Self::Report(error) => Some(error),
        }
    }
}

impl From<ValidationReportError> for CfbValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

/// Validates a positional CFB source under the default finite source and
/// report limits without changing it.
///
/// The single declared `cfb.container.ingress` capability runs the crate's
/// canonical ingress pipeline. On accepted input that pipeline checks the CFB
/// header, DIFAT/FAT, directory graph, MiniFAT, exact stream chains, allocation
/// overlap, and physical-sector reconciliation. A definite structural
/// rejection completes the capability with an error issue; the parser may
/// stop at that first conclusive violation. Encryption, signatures, and
/// application-level stream semantics are outside this capability and are not
/// claimed by the report.
///
/// # Errors
///
/// Returns an error if the source cannot be read consistently, a bounded
/// parser allocation fails, or the bounded report cannot be retained.
pub fn validate_source(source: Arc<dyn ReadAt>) -> Result<ValidateReport, CfbValidationError> {
    validate_source_with_limits(
        source,
        SharedOleFileLimits::default(),
        ValidationLimits::default(),
    )
}

/// Validates a positional CFB source under explicit finite source and report
/// limits without changing it.
///
/// An input exceeding `source_limits` produces a complete report value whose
/// ingress capability is `Blocked`; it is not treated as malformed CFB. A
/// structural parser rejection instead produces a `Complete` capability plus
/// one deterministic error issue. Source I/O, source-version changes, and
/// allocation failures remain errors because no honest report can be made.
///
/// # Errors
///
/// Returns an error if the source cannot be read consistently, a bounded
/// parser allocation fails, or `report_limits` cannot retain the result.
pub fn validate_source_with_limits(
    source: Arc<dyn ReadAt>,
    source_limits: SharedOleFileLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport, CfbValidationError> {
    let (expected_version, source_size) =
        stable_source_state(source.as_ref()).map_err(CfbValidationError::Ingress)?;
    if source_size > source_limits.max_input_bytes() {
        let check = CheckCapabilityId::try_new(INGRESS_CHECK, report_limits)?;
        let status = CheckStatus::blocked(
            "source exceeds the configured CFB validation input ceiling",
            report_limits,
        )?;
        let report =
            ValidateReport::try_new([ValidationCheck::new(check, status)], [], report_limits)?;
        require_source_version(source.as_ref(), expected_version)
            .map_err(CfbValidationError::Ingress)?;
        return Ok(report);
    }

    let report = match SharedOleFile::open_with_limits(source.clone(), source_limits) {
        Ok(_validated) => complete_report(report_limits),
        Err(error) if is_structural_rejection(&error) => {
            rejected_report(&error, source_size, report_limits)
        },
        Err(error) => Err(CfbValidationError::Ingress(error)),
    }?;
    require_source_version(source.as_ref(), expected_version)
        .map_err(CfbValidationError::Ingress)?;
    Ok(report)
}

impl SharedOleFile {
    /// Reruns the full canonical CFB ingress pipeline against this view's
    /// positional source and returns a non-mutating bounded report.
    ///
    /// This deliberately reparses structural sectors rather than inferring
    /// results from the retained parsed index. Its I/O and parsing cost is
    /// therefore comparable to opening the same CFB source again. The source
    /// version captured by this view is checked both before and after that
    /// work; a changed or unstable source is returned as an error.
    ///
    /// # Errors
    ///
    /// Returns an error if the source cannot be read consistently, differs
    /// from the version used to open this view, a bounded parser allocation
    /// fails, or `report_limits` cannot retain the result.
    pub fn validate(
        &self,
        report_limits: ValidationLimits,
    ) -> Result<ValidateReport, CfbValidationError> {
        self.check_source_version()
            .map_err(CfbValidationError::Ingress)?;
        let source_limits =
            SharedOleFileLimits::new(self.file_size()).map_err(CfbValidationError::Ingress)?;
        let report =
            validate_source_with_limits(self.source.clone(), source_limits, report_limits)?;
        self.check_source_version()
            .map_err(CfbValidationError::Ingress)?;
        Ok(report)
    }
}

fn stable_source_state(source: &dyn ReadAt) -> Result<(litchi_core::SourceVersion, u64), OleError> {
    let expected = source.version()?;
    let length = source.len()?;
    require_source_version(source, expected)?;
    Ok((expected, length))
}

fn require_source_version(
    source: &dyn ReadAt,
    expected: litchi_core::SourceVersion,
) -> Result<(), OleError> {
    let observed = source.version()?;
    if observed != expected {
        return Err(OleError::SourceChanged { expected, observed });
    }
    Ok(())
}

fn complete_report(limits: ValidationLimits) -> Result<ValidateReport, CfbValidationError> {
    let check = CheckCapabilityId::try_new(INGRESS_CHECK, limits)?;
    ValidateReport::try_new(
        [ValidationCheck::new(check, CheckStatus::Complete)],
        [],
        limits,
    )
    .map_err(Into::into)
}

fn rejected_report(
    error: &OleError,
    source_size: u64,
    limits: ValidationLimits,
) -> Result<ValidateReport, CfbValidationError> {
    let check = CheckCapabilityId::try_new(INGRESS_CHECK, limits)?;
    let (code, message, diagnostic) = rejection_fields(error);
    let location = IssueLocation::try_new(Some("compound-file"), None, None, None, None, limits)?;
    let source_size =
        IssueEvidence::try_new("source.size", EvidenceValue::Size(source_size), limits)?;
    let diagnostic = IssueEvidence::try_new(
        "diagnostic.sha256",
        EvidenceValue::Sha256(EvidenceDigest::of(diagnostic.as_bytes())),
        limits,
    )?;
    let issue = ValidationIssue::try_new(
        check.clone(),
        code,
        IssueSeverity::Error,
        message,
        [location],
        [source_size, diagnostic],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?;
    ValidateReport::try_new(
        [ValidationCheck::new(check, CheckStatus::Complete)],
        [issue],
        limits,
    )
    .map_err(Into::into)
}

fn is_structural_rejection(error: &OleError) -> bool {
    matches!(
        error,
        OleError::InvalidFormat(_)
            | OleError::InvalidData(_)
            | OleError::NotOleFile
            | OleError::CorruptedFile(_)
            | OleError::StreamNotFound
    )
}

fn rejection_fields(error: &OleError) -> (&'static str, &'static str, &str) {
    match error {
        OleError::NotOleFile => (
            "cfb.container.not_ole",
            "The source is not a recognizable CFB compound file",
            "not-ole-file",
        ),
        OleError::InvalidFormat(diagnostic) => (
            "cfb.container.invalid_format",
            "The CFB ingress rejected invalid format metadata",
            diagnostic,
        ),
        OleError::InvalidData(diagnostic) => (
            "cfb.container.invalid_data",
            "The CFB ingress rejected invalid structural data",
            diagnostic,
        ),
        OleError::CorruptedFile(diagnostic) => (
            "cfb.container.corrupted",
            "The CFB ingress detected corrupted container topology",
            diagnostic,
        ),
        OleError::StreamNotFound => (
            "cfb.container.missing_stream",
            "The CFB ingress could not resolve a required structural stream",
            "required-structural-stream-not-found",
        ),
        OleError::Io(_)
        | OleError::Allocation { .. }
        | OleError::Committed { .. }
        | OleError::SourceChanged { .. } => unreachable!("non-structural error was classified"),
    }
}
