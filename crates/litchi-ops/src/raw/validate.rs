//! `File::validate`: the leaf owners' bounded validators over the published
//! package, plus OPC part-graph checks for OOXML.

#![cfg_attr(
    not(any(
        feature = "docx",
        feature = "pptx",
        feature = "xlsx",
        feature = "xlsb",
        feature = "xls",
        feature = "ppt",
        feature = "doc",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        odf_objects
    )),
    allow(
        unreachable_code,
        unused_variables,
        reason = "without a validating format every arm refuses"
    )
)]

use std::fmt;
#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "xls",
    feature = "ppt",
    feature = "doc"
))]
use std::sync::Arc;

#[allow(
    unused_imports,
    reason = "each import serves some format features only"
)]
use litchi_core::{FileFormat, ValidateReport, ValidationLimits, ValidationReportError};
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
use litchi_opc::SourceLimits;

use super::{Validation, with_format};
use crate::{Code, Error as OpsError, File, Result};

/// Why a validation could not produce a report (record 0336). Findings are
/// never errors; they are in the [`Validation`] report.
///
/// `Display` is content-free: the leaf validators report only structural
/// metadata.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// OPC ingress of the published package, or the part-graph checks.
    #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
    Opc(litchi_opc::OpcError),
    /// The DOCX validator.
    #[cfg(feature = "docx")]
    Docx(litchi_docx::validation::Error),
    /// The PPTX validator.
    #[cfg(feature = "pptx")]
    Pptx(litchi_pptx::validation::Error),
    /// The XLS validator.
    #[cfg(feature = "xls")]
    Xls(litchi_xls::validation::Error),
    /// The PPT validator.
    #[cfg(feature = "ppt")]
    Ppt(litchi_ppt::validation::Error),
    /// The compound-file validator (DOC).
    #[cfg(feature = "doc")]
    Cfb(litchi_cfb::CfbValidationError),
    /// The ODF package validator.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    Odf(litchi_odf_common::OdfValidationError),
    /// The merged report exceeded the report bounds.
    Report(ValidationReportError),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Report(_) => formatter.write_str("the validation report exceeded its bounds"),
            #[allow(unreachable_patterns, reason = "every leaf variant is feature-gated")]
            _ => formatter.write_str("the package validator could not complete"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
            Self::Opc(error) => Some(error),
            #[cfg(feature = "docx")]
            Self::Docx(error) => Some(error),
            #[cfg(feature = "pptx")]
            Self::Pptx(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::Xls(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::Ppt(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::Cfb(error) => Some(error),
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::Odf(error) => Some(error),
            Self::Report(error) => Some(error),
        }
    }
}

impl Error {
    /// The unified code: report bounds are `limit_exceeded`, everything else
    /// `internal_error` (the validators turn structural rejections into
    /// findings, so what remains is allocation or I/O on published bytes).
    pub(crate) const fn code(&self) -> Code {
        match self {
            Self::Report(_) => Code::LimitExceeded,
            #[allow(unreachable_patterns, reason = "every leaf variant is feature-gated")]
            _ => Code::InternalError,
        }
    }
}

pub(super) fn validate(file: &File) -> Result<Validation> {
    if let File::Locked(locked) = file {
        return Err(locked.read_error());
    }
    let Some(format) = file.format() else {
        return Err(refused(file));
    };
    let report = match format {
        #[cfg(feature = "docx")]
        FileFormat::Docx => {
            let limits = file.ooxml_read_limits();
            let bytes = Arc::new(file.to_bytes()?);
            let leaf = litchi_docx::validation::validate_read_at(
                Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&bytes))),
                limits,
                litchi_docx::validation::Limits::default(),
                ValidationLimits::default(),
            )
            .map_err(Error::Docx)?;
            with_parts(leaf, bytes, limits)?
        },
        #[cfg(feature = "pptx")]
        FileFormat::Pptx => {
            let limits = file.ooxml_read_limits();
            let bytes = Arc::new(file.to_bytes()?);
            let leaf = litchi_pptx::validation::validate_source(
                Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&bytes))),
                litchi_pptx::validation::Limits::default().with_read_limits(limits),
                ValidationLimits::default(),
            )
            .map_err(Error::Pptx)?;
            with_parts(leaf, bytes, limits)?
        },
        #[cfg(any(feature = "xlsx", feature = "xlsb"))]
        FileFormat::Xlsx | FileFormat::Xlsb => {
            let limits = file.ooxml_read_limits();
            let bytes = Arc::new(file.to_bytes()?);
            let leaf = litchi_opc::validate_read_at(
                Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&bytes))),
                limits,
                ValidationLimits::default(),
            )
            .map_err(Error::Opc)?;
            with_parts(leaf, bytes, limits)?
        },
        #[cfg(feature = "xls")]
        FileFormat::Xls => {
            let bytes = file.to_bytes()?;
            litchi_xls::validation::validate_source(
                Arc::new(litchi_core::OwnedSource::new(bytes)),
                litchi_xls::validation::Limits::default(),
            )
            .map_err(Error::Xls)?
        },
        #[cfg(feature = "ppt")]
        FileFormat::Ppt => {
            let bytes = file.to_bytes()?;
            litchi_ppt::validate_source(
                Arc::new(litchi_core::OwnedSource::new(bytes)),
                litchi_ppt::validation::Limits::default(),
            )
            .map_err(Error::Ppt)?
        },
        #[cfg(feature = "doc")]
        FileFormat::Doc => {
            let bytes = file.to_bytes()?;
            litchi_cfb::validate_source(
                Arc::new(litchi_core::OwnedSource::new(bytes)),
                litchi_cfb::SharedOleFileLimits::default(),
                ValidationLimits::default(),
            )
            .map_err(Error::Cfb)?
        },
        #[cfg(feature = "rtf")]
        FileFormat::Rtf => {
            let File::Rtf(handle) = file else {
                return Err(refused(file));
            };
            rtf(handle)?
        },
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        FileFormat::Odt
        | FileFormat::Oth
        | FileFormat::Odm
        | FileFormat::Ods
        | FileFormat::Odp
        | FileFormat::Odg
        | FileFormat::Odc
        | FileFormat::Odi
        | FileFormat::Odb
        | FileFormat::Odf => {
            let bytes = file.to_bytes()?;
            if bytes.starts_with(b"PK") {
                litchi_odf_common::validate_package(
                    &bytes,
                    odf_validation_limits(file),
                    ValidationLimits::default(),
                )
                .map_err(Error::Odf)?
            } else {
                flat_odf()?
            }
        },
        _ => return Err(refused(file)),
    };
    Ok(Validation { format, report })
}

/// Keep generic validation's own explicit report policy while intersecting
/// its package-admission dimensions with the ODF owner that admitted `file`.
/// Validation is a separate API: its XML-event, depth and report ceilings
/// intentionally remain the validator defaults rather than being invented as
/// new `OpenOptions` fields.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn odf_validation_limits(file: &File) -> litchi_odf_common::OdfValidationLimits {
    let source = file.odf_read_limits().source_limits_value();
    litchi_odf_common::OdfValidationLimits::default()
        .with_max_input_bytes(source.max_input_bytes())
        .with_max_entries(source.max_members())
        .with_max_mimetype_bytes(source.max_mimetype_bytes())
        .with_max_manifest_bytes(source.max_manifest_bytes())
        .with_max_archive_member_name_bytes(source.max_archive_member_name_bytes())
        .with_max_archive_metadata_bytes(source.max_archive_metadata_bytes())
        .with_max_archive_compressed_bytes(source.max_archive_compressed_bytes())
        .with_max_archive_entry_bytes(source.max_entry_bytes())
        .with_max_archive_total_bytes(source.max_aggregate_bytes())
}

fn refused(file: &File) -> OpsError {
    with_format(
        file,
        OpsError::new(
            Code::UnsupportedCapability,
            "structural validation is not available for this format",
        )
        .with_suggestion(
            "validation covers OOXML, ODF, RTF and the legacy binary formats; iWork packages are validated by their owners on open",
        ),
    )
}

/// The leaf report plus the OPC part-graph checks, when the published
/// package indexes.
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
fn with_parts(
    leaf: ValidateReport,
    bytes: Arc<Vec<u8>>,
    limits: SourceLimits,
) -> Result<ValidateReport> {
    let Ok(package) = litchi_opc::OpcPackage::from_shared_vec(bytes, limits) else {
        // The leaf report already records the ingress rejection.
        return Ok(leaf);
    };
    let parts =
        litchi_opc::validate_parts(&package, ValidationLimits::default()).map_err(Error::Opc)?;
    merge(&leaf, &parts)
}

#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
fn merge(first: &ValidateReport, second: &ValidateReport) -> Result<ValidateReport> {
    let checks = first.checks().iter().chain(second.checks()).cloned();
    let issues = first.issues().iter().chain(second.issues()).cloned();
    Ok(
        ValidateReport::try_new(checks, issues, ValidationLimits::default())
            .map_err(Error::Report)?,
    )
}

/// RTF: the leaf's syntax and security inventory as a report.
#[cfg(feature = "rtf")]
fn rtf(file: &crate::format::rtf::Rtf) -> Result<ValidateReport> {
    use litchi_core::{
        CheckCapabilityId, CheckStatus, CompatibilityImpact, IssueLocation, IssueSeverity,
        RepairAvailability, ValidationCheck, ValidationIssue,
    };
    let limits = ValidationLimits::default();
    let report = |error: ValidationReportError| OpsError::from(Error::Report(error));
    let syntax = CheckCapabilityId::try_new("rtf.syntax", limits).map_err(report)?;
    let security =
        CheckCapabilityId::try_new("rtf.security.surface_presence", limits).map_err(report)?;
    let mut issues = Vec::new();
    let inventory = file.validation_report();
    if !inventory.is_conservatively_clean() {
        issues.push(
            ValidationIssue::try_new(
                security.clone(),
                "rtf.security.surface_present",
                IssueSeverity::Info,
                "The document carries objects, fields or other active surfaces; they were not executed.",
                [IssueLocation::try_new(None, Some("document"), None, None, None, limits).map_err(report)?],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                limits,
            )
            .map_err(report)?,
        );
    }
    let security_status = CheckStatus::Complete;
    ValidateReport::try_new(
        [
            ValidationCheck::new(syntax, CheckStatus::Complete),
            ValidationCheck::new(security, security_status),
        ],
        issues,
        limits,
    )
    .map_err(report)
}

/// A flat ODF document: one XML file, already parsed on open.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn flat_odf() -> Result<ValidateReport> {
    use litchi_core::{CheckCapabilityId, CheckStatus, ValidationCheck};
    let limits = ValidationLimits::default();
    let id = CheckCapabilityId::try_new("odf.flat.package", limits).map_err(Error::Report)?;
    Ok(ValidateReport::try_new(
        [ValidationCheck::new(id, CheckStatus::NotApplicable)],
        [],
        limits,
    )
    .map_err(Error::Report)?)
}
