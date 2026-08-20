//! Bounded, non-mutating validation of generic ODF package ingress.
//!
//! The validator indexes ZIP metadata and materializes only `mimetype`, the
//! package manifest, and `content.xml`. It does not decrypt content, verify
//! signatures, execute macros, fetch links, or validate a document family's
//! semantics. One otherwise-valid `mimetype` local-extra issue is identified
//! with a separate repair ID; execution belongs to the bounded `repair`
//! module.

use std::{collections::HashSet, error::Error, fmt, str};

use crate::package::MAX_MANIFEST_ENTRIES;
use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceDigest, EvidenceValue,
    IssueEvidence, IssueLocation, IssueSeverity, RepairAvailability, ValidateReport,
    ValidationCheck, ValidationIssue, ValidationLimits, ValidationReportError,
};
use quick_xml::{
    XmlVersion,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use soapberry_zip::{
    Error as ZipError, ErrorKind as ZipErrorKind, LimitResource, ZipArchive,
    office::{ArchiveLimits, ArchiveReader},
};

const INGRESS: &str = "odf.package.ingress";
const CATALOG: &str = "odf.package.catalog";
const DECLARATIONS: &str = "odf.package.mimetype_manifest";
const ROOT_XML: &str = "odf.package.root_xml";
const ENCRYPTION: &str = "odf.package.encryption_presence";
const SIGNATURES: &str = "odf.package.signature_presence";
const EXTERNAL: &str = "odf.content_xml.external_reference_presence";
const MACROS: &str = "odf.package.macro_storage_presence";

const MANIFEST_PATH: &str = "META-INF/manifest.xml";
const CONTENT_PATH: &str = "content.xml";
const MIMETYPE_PATH: &str = "mimetype";
const DOCUMENT_SIGNATURE_PATH: &str = "META-INF/documentsignatures.xml";
const MACRO_SIGNATURE_PATH: &str = "META-INF/macrosignatures.xml";
pub(crate) const MANIFEST_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MIMETYPE_REPAIR_ID: &str = "odf.repair.mimetype_local_extra";

/// Default finite bounds for generic ODF validation.
pub const DEFAULT_ODF_VALIDATION_LIMITS: OdfValidationLimits = OdfValidationLimits {
    max_input_bytes: 2 * 1024 * 1024 * 1024,
    archive: ArchiveLimits {
        max_files: 100_000,
        max_member_name_bytes: 4 * 1024,
        max_metadata_bytes: 64 * 1024 * 1024,
        max_compressed_size: 512 * 1024 * 1024,
        max_entry_size: 512 * 1024 * 1024,
        max_total_size: 2 * 1024 * 1024 * 1024,
    },
    max_mimetype_bytes: 512,
    max_manifest_bytes: 16 * 1024 * 1024,
    max_manifest_entries: MAX_MANIFEST_ENTRIES,
    max_root_xml_bytes: 64 * 1024 * 1024,
    max_xml_events: 4_000_000,
    max_xml_depth: 4_096,
};

/// Finite input and parser bounds for one generic ODF validation pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OdfValidationLimits {
    max_input_bytes: u64,
    archive: ArchiveLimits,
    max_mimetype_bytes: u64,
    max_manifest_bytes: u64,
    max_manifest_entries: usize,
    max_root_xml_bytes: u64,
    max_xml_events: usize,
    max_xml_depth: usize,
}

impl OdfValidationLimits {
    /// Returns a copy with a new whole-source byte ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: u64) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    /// Returns a copy with a new ZIP member-count ceiling.
    #[must_use]
    pub const fn with_max_entries(mut self, maximum: usize) -> Self {
        self.archive.max_files = maximum;
        self
    }

    /// Returns a copy with a new per-member declared-size ceiling.
    #[must_use]
    pub const fn with_max_archive_entry_bytes(mut self, maximum: u64) -> Self {
        self.archive.max_entry_size = maximum;
        self
    }

    /// Returns a copy with a new per-member compressed-size ceiling.
    #[must_use]
    pub const fn with_max_archive_compressed_bytes(mut self, maximum: u64) -> Self {
        self.archive.max_compressed_size = maximum;
        self
    }

    /// Returns a copy with a new aggregate uncompressed ZIP-size ceiling.
    #[must_use]
    pub const fn with_max_archive_total_bytes(mut self, maximum: u64) -> Self {
        self.archive.max_total_size = maximum;
        self
    }

    /// Returns a copy with a new raw member-name ceiling.
    #[must_use]
    pub const fn with_max_archive_member_name_bytes(mut self, maximum: u64) -> Self {
        self.archive.max_member_name_bytes = maximum;
        self
    }

    /// Returns a copy with a new aggregate ZIP metadata byte ceiling.
    #[must_use]
    pub const fn with_max_archive_metadata_bytes(mut self, maximum: u64) -> Self {
        self.archive.max_metadata_bytes = maximum;
        self
    }

    /// Returns a copy with a new manifest materialization ceiling.
    #[must_use]
    pub const fn with_max_manifest_bytes(mut self, maximum: u64) -> Self {
        self.max_manifest_bytes = maximum;
        self
    }

    /// Returns a copy with a new manifest file-entry ceiling.
    #[must_use]
    pub const fn with_max_manifest_entries(mut self, maximum: usize) -> Self {
        self.max_manifest_entries = maximum;
        self
    }

    /// Returns a copy with a new `content.xml` materialization ceiling.
    #[must_use]
    pub const fn with_max_root_xml_bytes(mut self, maximum: u64) -> Self {
        self.max_root_xml_bytes = maximum;
        self
    }

    /// Returns a copy with a new XML-event ceiling.
    #[must_use]
    pub const fn with_max_xml_events(mut self, maximum: usize) -> Self {
        self.max_xml_events = maximum;
        self
    }

    /// Returns a copy with a new XML nesting-depth ceiling.
    #[must_use]
    pub const fn with_max_xml_depth(mut self, maximum: usize) -> Self {
        self.max_xml_depth = maximum;
        self
    }

    /// Returns the whole-source byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Returns the ZIP member-count ceiling.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.archive.max_files
    }

    /// Returns the manifest materialization ceiling.
    #[must_use]
    pub const fn max_manifest_bytes(self) -> u64 {
        self.max_manifest_bytes
    }

    /// Returns the manifest file-entry ceiling.
    #[must_use]
    pub const fn max_manifest_entries(self) -> usize {
        self.max_manifest_entries
    }

    /// Returns the `content.xml` materialization ceiling.
    #[must_use]
    pub const fn max_root_xml_bytes(self) -> u64 {
        self.max_root_xml_bytes
    }

    /// Returns the XML-event ceiling.
    #[must_use]
    pub const fn max_xml_events(self) -> usize {
        self.max_xml_events
    }

    /// Returns the XML nesting-depth ceiling.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.max_xml_depth
    }
}

impl Default for OdfValidationLimits {
    fn default() -> Self {
        DEFAULT_ODF_VALIDATION_LIMITS
    }
}

/// Failure to perform or retain an ODF validation report.
///
/// Definite package and XML rejections are issues in a successful report.
/// Configured input ceilings become blocked checks. This error is reserved for
/// allocation failure and bounded report construction failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum OdfValidationError {
    /// A bounded validator-owned allocation failed.
    Allocation {
        /// The allocation's fixed, content-free resource name.
        resource: &'static str,
    },
    /// The shared bounded report value rejected the result.
    Report(ValidationReportError),
}

impl fmt::Display for OdfValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Allocation { resource } => {
                write!(
                    formatter,
                    "allocation failed while validating ODF {resource}"
                )
            },
            Self::Report(error) => write!(formatter, "ODF validation report failed: {error}"),
        }
    }
}

impl Error for OdfValidationError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Allocation { .. } => None,
            Self::Report(error) => Some(error),
        }
    }
}

impl From<ValidationReportError> for OdfValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

/// Validate borrowed ODF package bytes under default finite input and report
/// bounds without changing or retaining the package.
///
/// The immutable slice is the source-identity bracket for the entire call.
/// ZIP metadata is bounded-indexed once. A separate allocation-free raw-header
/// pass checks path spelling and the mimetype local header; ordinary package
/// payloads remain unread.
///
/// # Errors
///
/// Returns an error only if a bounded allocation or report construction fails.
pub fn validate_package(data: &[u8]) -> Result<ValidateReport, OdfValidationError> {
    validate_package_with_limits(
        data,
        OdfValidationLimits::default(),
        ValidationLimits::default(),
    )
}

/// Validate borrowed ODF package bytes under explicit finite input and report
/// bounds without changing or retaining the package.
///
/// A configured input, ZIP, manifest, or XML ceiling is represented by the
/// affected capability's `Blocked` status. Structural rejection is instead a
/// completed check with a deterministic issue. Signature cryptography, macro
/// behavior, semantic document validity, and external fetching are not
/// capabilities of this function. The one supported `mimetype` local-extra
/// diagnostic is only an authorization hint for the separate repair module;
/// it does not execute a repair.
///
/// # Errors
///
/// Returns an error only if a bounded allocation or report construction fails.
pub fn validate_package_with_limits(
    data: &[u8],
    input_limits: OdfValidationLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport, OdfValidationError> {
    let source_size = u64::try_from(data.len()).unwrap_or(u64::MAX);
    if source_size > input_limits.max_input_bytes {
        return ingress_blocked_report(
            "source exceeds the configured ODF validation input ceiling",
            report_limits,
        );
    }

    let raw_catalog = match inspect_raw_catalog(data, input_limits.archive) {
        Ok(observations) => observations,
        Err(error) => {
            if matches!(error.kind(), ZipErrorKind::LimitExceeded { .. }) {
                return ingress_blocked_report(zip_limit_reason(&error), report_limits);
            }
            return ingress_rejected_report(&error, source_size, report_limits);
        },
    };

    let archive = match ArchiveReader::new_with_limits(data, input_limits.archive) {
        Ok(archive) => archive,
        Err(error) => {
            if allocation_failure(&error) {
                return Err(OdfValidationError::Allocation {
                    resource: "ZIP catalog",
                });
            }
            if matches!(error.kind(), ZipErrorKind::LimitExceeded { .. }) {
                return ingress_blocked_report(zip_limit_reason(&error), report_limits);
            }
            return ingress_rejected_report(&error, source_size, report_limits);
        },
    };

    let mut state = ValidationState::new(report_limits)?;
    state.ingress = CheckStatus::Complete;

    let hostile_paths = raw_catalog.hostile_paths;
    if hostile_paths != 0 {
        state.push_issue(count_issue(
            CATALOG,
            "odf.catalog.hostile_path",
            IssueSeverity::Error,
            "The ZIP catalog contains non-canonical or traversal-capable member paths.",
            "package-catalog",
            "hostile_paths",
            hostile_paths,
            CompatibilityImpact::Security,
            report_limits,
        )?)?;
    }

    let signature_count = u64::from(archive.contains(DOCUMENT_SIGNATURE_PATH))
        + u64::from(archive.contains(MACRO_SIGNATURE_PATH));
    state.signatures = presence_status(signature_count);
    if signature_count != 0 {
        state.push_issue(count_issue(
            SIGNATURES,
            "odf.signature.infrastructure_present",
            IssueSeverity::Info,
            "ODF signature infrastructure is present; cryptographic validity was not checked.",
            "signature-infrastructure",
            "signature_parts",
            signature_count,
            CompatibilityImpact::None,
            report_limits,
        )?)?;
    }

    let macro_storage_count = archive
        .file_names()
        .filter(|name| macro_storage_path(name))
        .count() as u64;
    state.macros = presence_status(macro_storage_count);
    if macro_storage_count != 0 {
        state.push_issue(count_issue(
            MACROS,
            "odf.macro.storage_present",
            IssueSeverity::Warning,
            "ODF macro storage is present; macro behavior and safety were not assessed.",
            "macro-storage",
            "macro_members",
            macro_storage_count,
            CompatibilityImpact::Security,
            report_limits,
        )?)?;
    }

    let declarations = inspect_declarations(
        raw_catalog,
        &archive,
        input_limits,
        report_limits,
        source_size,
        data,
        &mut state,
    )?;
    if let Some(declarations) = declarations.as_ref() {
        inspect_catalog(&archive, &declarations.manifest, report_limits, &mut state)?;
        inspect_encryption(declarations.encryption, report_limits, &mut state)?;
    } else if state.declarations.is_complete() {
        state.catalog = CheckStatus::blocked(
            "manifest was unavailable after a conclusive declaration rejection",
            report_limits,
        )?;
        state.encryption = CheckStatus::blocked(
            "manifest was unavailable after a conclusive declaration rejection",
            report_limits,
        )?;
    } else {
        let declaration_id = id(DECLARATIONS, report_limits)?;
        state.catalog = CheckStatus::stopped_by(declaration_id.clone());
        state.encryption = CheckStatus::stopped_by(declaration_id);
    }

    inspect_content_xml(
        &archive,
        declarations
            .as_ref()
            .is_some_and(|declarations| declarations.encryption.content_xml),
        input_limits,
        report_limits,
        &mut state,
    )?;

    state.finish()
}

struct Declarations {
    manifest: ValidationManifest,
    encryption: EncryptionPresence,
}

struct ValidationManifest {
    mimetype: String,
    paths: HashSet<String>,
}

#[derive(Clone, Copy)]
struct EncryptionPresence {
    count: u64,
    content_xml: bool,
}

struct ValidationState {
    limits: ValidationLimits,
    ingress: CheckStatus,
    catalog: CheckStatus,
    declarations: CheckStatus,
    root_xml: CheckStatus,
    encryption: CheckStatus,
    signatures: CheckStatus,
    external: CheckStatus,
    macros: CheckStatus,
    issues: Vec<ValidationIssue>,
}

impl ValidationState {
    fn new(limits: ValidationLimits) -> Result<Self, OdfValidationError> {
        let mut issues = Vec::new();
        issues
            .try_reserve_exact(limits.max_issues().min(16))
            .map_err(|_| OdfValidationError::Allocation {
                resource: "issue staging",
            })?;
        let pending = CheckStatus::blocked("validation phase has not run", limits)?;
        Ok(Self {
            limits,
            ingress: pending.clone(),
            catalog: pending.clone(),
            declarations: pending.clone(),
            root_xml: pending.clone(),
            encryption: pending.clone(),
            signatures: pending.clone(),
            external: pending.clone(),
            macros: pending,
            issues,
        })
    }

    fn push_issue(&mut self, issue: ValidationIssue) -> Result<(), OdfValidationError> {
        if self.issues.len() >= self.limits.max_issues() {
            return Err(OdfValidationError::Report(ValidationReportError::Limit {
                kind: litchi_core::ValidationLimitKind::Issues,
                observed: self.issues.len().saturating_add(1),
                limit: self.limits.max_issues(),
            }));
        }
        if self.issues.len() == self.issues.capacity() {
            self.issues
                .try_reserve(1)
                .map_err(|_| OdfValidationError::Allocation {
                    resource: "issue staging",
                })?;
        }
        self.issues.push(issue);
        Ok(())
    }

    fn finish(self) -> Result<ValidateReport, OdfValidationError> {
        let checks = [
            check(INGRESS, self.ingress, self.limits)?,
            check(CATALOG, self.catalog, self.limits)?,
            check(DECLARATIONS, self.declarations, self.limits)?,
            check(ROOT_XML, self.root_xml, self.limits)?,
            check(ENCRYPTION, self.encryption, self.limits)?,
            check(SIGNATURES, self.signatures, self.limits)?,
            check(EXTERNAL, self.external, self.limits)?,
            check(MACROS, self.macros, self.limits)?,
        ];
        ValidateReport::try_new(checks, self.issues, self.limits).map_err(Into::into)
    }
}

fn inspect_declarations(
    raw_catalog: RawCatalogObservations,
    archive: &ArchiveReader<'_>,
    input_limits: OdfValidationLimits,
    report_limits: ValidationLimits,
    source_size: u64,
    source: &[u8],
    state: &mut ValidationState,
) -> Result<Option<Declarations>, OdfValidationError> {
    if !archive.contains(MIMETYPE_PATH) || !archive.contains(MANIFEST_PATH) {
        state.declarations = CheckStatus::Complete;
        let missing = u64::from(!archive.contains(MIMETYPE_PATH))
            + u64::from(!archive.contains(MANIFEST_PATH));
        state.push_issue(count_issue(
            DECLARATIONS,
            "odf.declarations.missing",
            IssueSeverity::Error,
            "The ODF package is missing a required mimetype or manifest member.",
            "package-declarations",
            "missing_members",
            missing,
            CompatibilityImpact::Interoperability,
            report_limits,
        )?)?;
        return Ok(None);
    }

    let mimetype_size = match archive.metadata(MIMETYPE_PATH) {
        Ok(metadata) => metadata.uncompressed_size(),
        Err(error) => {
            return declaration_read_rejection(error, report_limits, state).map(|()| None);
        },
    };
    let manifest_size = match archive.metadata(MANIFEST_PATH) {
        Ok(metadata) => metadata.uncompressed_size(),
        Err(error) => {
            return declaration_read_rejection(error, report_limits, state).map(|()| None);
        },
    };
    if mimetype_size > input_limits.max_mimetype_bytes {
        state.declarations = CheckStatus::blocked(
            "mimetype exceeds the configured ODF declaration byte ceiling",
            report_limits,
        )?;
        return Ok(None);
    }
    if manifest_size > input_limits.max_manifest_bytes {
        state.declarations = CheckStatus::blocked(
            "manifest exceeds the configured ODF manifest byte ceiling",
            report_limits,
        )?;
        return Ok(None);
    }

    let mimetype = match archive.read(MIMETYPE_PATH) {
        Ok(bytes) => bytes,
        Err(error) => {
            return declaration_read_rejection(error, report_limits, state).map(|()| None);
        },
    };
    let manifest_bytes = match archive.read(MANIFEST_PATH) {
        Ok(bytes) => bytes,
        Err(error) => {
            return declaration_read_rejection(error, report_limits, state).map(|()| None);
        },
    };
    let observations = match inspect_xml(&manifest_bytes, input_limits) {
        Ok(observations) if observations.root_is_manifest => {
            if observations.manifest_entries > input_limits.max_manifest_entries as u64 {
                state.declarations = CheckStatus::blocked(
                    "manifest file entries exceed the configured ODF ceiling",
                    report_limits,
                )?;
                return Ok(None);
            }
            observations
        },
        Ok(_) => {
            state.declarations = CheckStatus::Complete;
            state.push_issue(simple_issue(
                DECLARATIONS,
                "odf.manifest.root_invalid",
                IssueSeverity::Error,
                "The required ODF manifest does not have a manifest:manifest document element.",
                MANIFEST_PATH,
                CompatibilityImpact::Interoperability,
                report_limits,
            )?)?;
            return Ok(None);
        },
        Err(XmlInspectionError::Limit(reason)) => {
            state.declarations = CheckStatus::blocked(reason, report_limits)?;
            return Ok(None);
        },
        Err(XmlInspectionError::Malformed { offset, doctype }) => {
            state.declarations = CheckStatus::Complete;
            let (code, message) = if doctype {
                (
                    "odf.manifest.doctype_forbidden",
                    "The required ODF manifest contains a forbidden document type declaration.",
                )
            } else {
                (
                    "odf.manifest.invalid_xml",
                    "The required ODF manifest is not well-formed XML.",
                )
            };
            state.push_issue(offset_issue(
                DECLARATIONS,
                code,
                message,
                MANIFEST_PATH,
                offset,
                report_limits,
            )?)?;
            return Ok(None);
        },
        Err(XmlInspectionError::ManifestStructure { offset }) => {
            state.declarations = CheckStatus::Complete;
            state.push_issue(offset_issue(
                DECLARATIONS,
                "odf.manifest.encryption_placement_invalid",
                "ODF encryption metadata is misplaced or repeated within a manifest file entry.",
                MANIFEST_PATH,
                offset,
                report_limits,
            )?)?;
            return Ok(None);
        },
    };
    let manifest_text = match str::from_utf8(&manifest_bytes) {
        Ok(text) => text,
        Err(_) => {
            state.declarations = CheckStatus::Complete;
            state.push_issue(simple_issue(
                DECLARATIONS,
                "odf.manifest.invalid_utf8",
                IssueSeverity::Error,
                "The required ODF manifest is not UTF-8 XML.",
                MANIFEST_PATH,
                CompatibilityImpact::Interoperability,
                report_limits,
            )?)?;
            return Ok(None);
        },
    };
    let manifest = match parse_validation_manifest(
        manifest_text,
        usize::try_from(observations.manifest_entries).unwrap_or(usize::MAX),
    ) {
        Ok(manifest) => manifest,
        Err(ValidationManifestError::Allocation { resource }) => {
            return Err(OdfValidationError::Allocation { resource });
        },
        Err(ValidationManifestError::Invalid { diagnostic }) => {
            state.declarations = CheckStatus::Complete;
            state.push_issue(diagnostic_issue(
                DECLARATIONS,
                "odf.manifest.invalid",
                "The required ODF manifest is structurally invalid.",
                MANIFEST_PATH,
                diagnostic,
                report_limits,
            )?)?;
            return Ok(None);
        },
    };
    let encryption = match inspect_manifest_encryption(&manifest_bytes) {
        Ok(presence) => presence,
        Err(XmlInspectionError::Malformed { offset, doctype }) => {
            state.declarations = CheckStatus::Complete;
            let (code, message) = if doctype {
                (
                    "odf.manifest.doctype_forbidden",
                    "The required ODF manifest contains a forbidden document type declaration.",
                )
            } else {
                (
                    "odf.manifest.invalid_xml",
                    "The required ODF manifest is not well-formed XML.",
                )
            };
            state.push_issue(offset_issue(
                DECLARATIONS,
                code,
                message,
                MANIFEST_PATH,
                offset,
                report_limits,
            )?)?;
            return Ok(None);
        },
        Err(XmlInspectionError::Limit(_)) => {
            unreachable!("the manifest presence scan has no independent resource ceiling")
        },
        Err(XmlInspectionError::ManifestStructure { offset }) => {
            state.declarations = CheckStatus::Complete;
            state.push_issue(offset_issue(
                DECLARATIONS,
                "odf.manifest.encryption_placement_invalid",
                "ODF encryption metadata is misplaced or repeated within a manifest file entry.",
                MANIFEST_PATH,
                offset,
                report_limits,
            )?)?;
            return Ok(None);
        },
    };

    state.declarations = CheckStatus::Complete;
    let mimetype_text = str::from_utf8(&mimetype).ok();
    if mimetype_text != Some(manifest.mimetype.as_str()) || !odf_mimetype(&mimetype) {
        state.push_issue(simple_issue(
            DECLARATIONS,
            "odf.mimetype.inconsistent",
            IssueSeverity::Error,
            "The mimetype member and manifest root media type are not the same supported ODF type.",
            MIMETYPE_PATH,
            CompatibilityImpact::Interoperability,
            report_limits,
        )?)?;
    }
    let first_and_stored = raw_catalog.mimetype_central_first
        && archive.is_stored(MIMETYPE_PATH).ok() == Some(true)
        && raw_catalog.mimetype_local_first;
    if !first_and_stored {
        state.push_issue(simple_issue(
            DECLARATIONS,
            "odf.mimetype.layout",
            IssueSeverity::Error,
            "The ODF mimetype member must be first, stored, and have no local-header extra field.",
            MIMETYPE_PATH,
            CompatibilityImpact::Interoperability,
            report_limits,
        )?)?;
    } else if raw_catalog.mimetype_repairable {
        state.push_issue(mimetype_extra_issue(
            source_size,
            EvidenceDigest::of(source),
            report_limits,
        )?)?;
    } else if raw_catalog.mimetype_local_extra {
        state.push_issue(simple_issue(
            DECLARATIONS,
            "odf.mimetype.layout",
            IssueSeverity::Error,
            "The ODF mimetype local-header extra field is not a supported bounded repair target.",
            MIMETYPE_PATH,
            CompatibilityImpact::Interoperability,
            report_limits,
        )?)?;
    }
    Ok(Some(Declarations {
        manifest,
        encryption,
    }))
}

fn inspect_catalog(
    archive: &ArchiveReader<'_>,
    manifest: &ValidationManifest,
    report_limits: ValidationLimits,
    state: &mut ValidationState,
) -> Result<(), OdfValidationError> {
    state.catalog = CheckStatus::Complete;
    let undeclared = archive
        .file_names()
        .filter(|name| {
            *name != MIMETYPE_PATH && *name != MANIFEST_PATH && !manifest.paths.contains(*name)
        })
        .count() as u64;
    let missing = manifest
        .paths
        .iter()
        .filter(|path| {
            path.as_str() != "/" && !path.ends_with('/') && !archive.contains(path.as_str())
        })
        .count() as u64;
    if undeclared != 0 || missing != 0 {
        let evidence = [
            IssueEvidence::try_new(
                "archive_members_without_manifest_entry",
                EvidenceValue::Count(undeclared),
                report_limits,
            )?,
            IssueEvidence::try_new(
                "manifest_entries_without_archive_member",
                EvidenceValue::Count(missing),
                report_limits,
            )?,
        ];
        state.push_issue(ValidationIssue::try_new(
            id(CATALOG, report_limits)?,
            "odf.catalog.inconsistent",
            IssueSeverity::Error,
            "The ZIP member catalog and ODF manifest file entries are inconsistent.",
            [location("package-catalog", report_limits)?],
            evidence,
            None,
            CompatibilityImpact::Interoperability,
            RepairAvailability::Unavailable,
            report_limits,
        )?)?;
    }
    if !archive.contains(CONTENT_PATH) || !manifest.paths.contains(CONTENT_PATH) {
        state.push_issue(simple_issue(
            CATALOG,
            "odf.catalog.content_missing",
            IssueSeverity::Error,
            "The required ODF content.xml member or its manifest entry is missing.",
            CONTENT_PATH,
            CompatibilityImpact::DataLoss,
            report_limits,
        )?)?;
    }
    Ok(())
}

fn inspect_encryption(
    presence: EncryptionPresence,
    report_limits: ValidationLimits,
    state: &mut ValidationState,
) -> Result<(), OdfValidationError> {
    let count = presence.count;
    state.encryption = presence_status(count);
    if count != 0 {
        state.push_issue(count_issue(
            ENCRYPTION,
            "odf.encryption.infrastructure_present",
            IssueSeverity::Info,
            "ODF encryption metadata is present; plaintext and password validity were not checked.",
            "encryption-infrastructure",
            "encrypted_entries",
            count,
            CompatibilityImpact::None,
            report_limits,
        )?)?;
    }
    Ok(())
}

fn inspect_content_xml(
    archive: &ArchiveReader<'_>,
    content_encrypted: bool,
    input_limits: OdfValidationLimits,
    report_limits: ValidationLimits,
    state: &mut ValidationState,
) -> Result<(), OdfValidationError> {
    if !archive.contains(CONTENT_PATH) {
        state.root_xml = CheckStatus::Complete;
        state.external = CheckStatus::blocked(
            "content.xml is absent after a conclusive catalog rejection",
            report_limits,
        )?;
        return Ok(());
    }
    if content_encrypted {
        state.root_xml = CheckStatus::blocked(
            "encrypted content.xml cannot be inspected without plaintext",
            report_limits,
        )?;
        state.external = CheckStatus::stopped_by(id(ROOT_XML, report_limits)?);
        return Ok(());
    }
    let size = match archive.metadata(CONTENT_PATH) {
        Ok(metadata) => metadata.uncompressed_size(),
        Err(error) => {
            return content_read_rejection(error, report_limits, state);
        },
    };
    if size > input_limits.max_root_xml_bytes {
        state.root_xml = CheckStatus::blocked(
            "content.xml exceeds the configured ODF root XML byte ceiling",
            report_limits,
        )?;
        state.external = CheckStatus::stopped_by(id(ROOT_XML, report_limits)?);
        return Ok(());
    }
    let content = match archive.read(CONTENT_PATH) {
        Ok(content) => content,
        Err(error) => return content_read_rejection(error, report_limits, state),
    };
    match inspect_xml(&content, input_limits) {
        Ok(observations) => {
            state.root_xml = CheckStatus::Complete;
            state.external = presence_status(observations.external_references);
            if observations.external_references != 0 {
                state.push_issue(count_issue(
                    EXTERNAL,
                    "odf.external_reference.present",
                    IssueSeverity::Warning,
                    "External XLink references are present; no target was fetched.",
                    CONTENT_PATH,
                    "external_references",
                    observations.external_references,
                    CompatibilityImpact::Security,
                    report_limits,
                )?)?;
            }
        },
        Err(XmlInspectionError::Limit(reason)) => {
            state.root_xml = CheckStatus::blocked(reason, report_limits)?;
            state.external = CheckStatus::stopped_by(id(ROOT_XML, report_limits)?);
        },
        Err(XmlInspectionError::Malformed { offset, doctype }) => {
            state.root_xml = CheckStatus::Complete;
            state.external = CheckStatus::blocked(
                "external-reference scan could not finish after malformed content XML",
                report_limits,
            )?;
            let code = if doctype {
                "odf.xml.doctype_forbidden"
            } else {
                "odf.xml.malformed"
            };
            let message = if doctype {
                "The required ODF content XML contains a forbidden document type declaration."
            } else {
                "The required ODF content XML is not well formed."
            };
            state.push_issue(offset_issue(
                ROOT_XML,
                code,
                message,
                CONTENT_PATH,
                offset,
                report_limits,
            )?)?;
        },
        Err(XmlInspectionError::ManifestStructure { offset }) => {
            state.root_xml = CheckStatus::Complete;
            state.external = CheckStatus::blocked(
                "external-reference scan could not finish after malformed content XML",
                report_limits,
            )?;
            state.push_issue(offset_issue(
                ROOT_XML,
                "odf.xml.malformed",
                "The required ODF content XML is not well formed.",
                CONTENT_PATH,
                offset,
                report_limits,
            )?)?;
        },
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct XmlObservations {
    external_references: u64,
    root_is_manifest: bool,
    manifest_entries: u64,
}

enum XmlInspectionError {
    Limit(&'static str),
    Malformed { offset: u64, doctype: bool },
    ManifestStructure { offset: u64 },
}

enum ValidationManifestError {
    Allocation { resource: &'static str },
    Invalid { diagnostic: &'static str },
}

fn parse_validation_manifest(
    xml: &str,
    entry_count: usize,
) -> Result<ValidationManifest, ValidationManifestError> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut paths = HashSet::new();
    paths
        .try_reserve(entry_count)
        .map_err(|_| ValidationManifestError::Allocation {
            resource: "manifest entry index",
        })?;
    let mut mimetype = String::new();
    let mut entry_open = false;

    loop {
        let (namespace, event) =
            reader
                .read_resolved_event()
                .map_err(|_| ValidationManifestError::Invalid {
                    diagnostic: "manifest XML event",
                })?;
        match event {
            Event::Start(element)
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                    && element.local_name().as_ref() == b"file-entry" =>
            {
                if entry_open {
                    return Err(ValidationManifestError::Invalid {
                        diagnostic: "nested manifest file entry",
                    });
                }
                insert_validation_manifest_entry(&reader, &element, &mut paths, &mut mimetype)?;
                entry_open = true;
            },
            Event::Empty(element)
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                    && element.local_name().as_ref() == b"file-entry" =>
            {
                if entry_open {
                    return Err(ValidationManifestError::Invalid {
                        diagnostic: "nested manifest file entry",
                    });
                }
                insert_validation_manifest_entry(&reader, &element, &mut paths, &mut mimetype)?;
            },
            Event::End(element)
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                    && element.local_name().as_ref() == b"file-entry" =>
            {
                entry_open = false;
            },
            Event::Eof if !entry_open => break,
            Event::Eof => {
                return Err(ValidationManifestError::Invalid {
                    diagnostic: "incomplete manifest file entry",
                });
            },
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    Ok(ValidationManifest { mimetype, paths })
}

fn insert_validation_manifest_entry(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    paths: &mut HashSet<String>,
    mimetype: &mut String,
) -> Result<(), ValidationManifestError> {
    let mut seen_attributes: Vec<Vec<u8>> = Vec::new();
    let mut full_path = None;
    let mut media_type = None;

    for result in element.attributes() {
        let attribute = result.map_err(|_| ValidationManifestError::Invalid {
            diagnostic: "manifest attribute syntax",
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE) {
            continue;
        }
        let local = local.as_ref();
        if seen_attributes.iter().any(|seen| seen.as_slice() == local) {
            return Err(ValidationManifestError::Invalid {
                diagnostic: "duplicate manifest attribute",
            });
        }
        if seen_attributes.len() == seen_attributes.capacity() {
            seen_attributes
                .try_reserve(1)
                .map_err(|_| ValidationManifestError::Allocation {
                    resource: "manifest attribute index",
                })?;
        }
        let mut owned_local = Vec::new();
        owned_local.try_reserve_exact(local.len()).map_err(|_| {
            ValidationManifestError::Allocation {
                resource: "manifest attribute name",
            }
        })?;
        owned_local.extend_from_slice(local);
        seen_attributes.push(owned_local);
        let value = decode_manifest_attribute(attribute.value.as_ref())?;
        match local {
            b"full-path" => full_path = Some(value),
            b"media-type" => media_type = Some(value),
            b"size" => {
                value
                    .parse::<u64>()
                    .map_err(|_| ValidationManifestError::Invalid {
                        diagnostic: "manifest entry size",
                    })?;
            },
            _ => {},
        }
    }

    let path =
        full_path
            .filter(|path| !path.is_empty())
            .ok_or(ValidationManifestError::Invalid {
                diagnostic: "manifest entry path",
            })?;
    if path == "/" {
        *mimetype = media_type.unwrap_or_default();
    }
    if !paths.insert(path) {
        return Err(ValidationManifestError::Invalid {
            diagnostic: "duplicate manifest entry path",
        });
    }
    Ok(())
}

fn decode_manifest_attribute(raw: &[u8]) -> Result<String, ValidationManifestError> {
    let text = str::from_utf8(raw).map_err(|_| ValidationManifestError::Invalid {
        diagnostic: "manifest attribute UTF-8",
    })?;
    let mut decoded = String::new();
    decoded
        .try_reserve_exact(raw.len())
        .map_err(|_| ValidationManifestError::Allocation {
            resource: "manifest attribute value",
        })?;
    let mut cursor = 0;
    while cursor < text.len() {
        let remainder = &text[cursor..];
        let character = remainder
            .chars()
            .next()
            .ok_or(ValidationManifestError::Invalid {
                diagnostic: "manifest attribute value",
            })?;
        if character == '&' {
            let end = remainder
                .find(';')
                .ok_or(ValidationManifestError::Invalid {
                    diagnostic: "manifest attribute reference",
                })?;
            let reference = &remainder[1..end];
            let resolved = match reference {
                "amp" => '&',
                "lt" => '<',
                "gt" => '>',
                "apos" => '\'',
                "quot" => '"',
                _ => BytesRef::new(reference)
                    .resolve_char_ref()
                    .ok()
                    .flatten()
                    .filter(|value| xml10_character(*value))
                    .ok_or(ValidationManifestError::Invalid {
                        diagnostic: "manifest attribute reference",
                    })?,
            };
            decoded.push(resolved);
            cursor += end + 1;
        } else {
            let width = character.len_utf8();
            cursor += width;
            match character {
                '\r' => {
                    if text[cursor..].starts_with('\n') {
                        cursor += 1;
                    }
                    decoded.push(' ');
                },
                '\n' | '\t' => decoded.push(' '),
                value if xml10_character(value) => decoded.push(value),
                _ => {
                    return Err(ValidationManifestError::Invalid {
                        diagnostic: "manifest attribute character",
                    });
                },
            }
        }
    }
    Ok(decoded)
}

fn xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn inspect_manifest_encryption(xml: &[u8]) -> Result<EncryptionPresence, XmlInspectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut current_entry: Option<EntryEncryptionPresence> = None;
    let mut count = 0_u64;
    let mut content_xml = false;

    loop {
        let offset = reader.buffer_position();
        let (namespace, event) =
            reader
                .read_resolved_event()
                .map_err(|_| XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                })?;
        match event {
            Event::Start(element) => {
                let manifest_element = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE);
                if manifest_element && element.local_name().as_ref() == b"file-entry" {
                    current_entry = Some(EntryEncryptionPresence {
                        content_xml: manifest_entry_is_content(&reader, &element, offset)?,
                        encrypted: false,
                    });
                } else if manifest_element && element.local_name().as_ref() == b"encryption-data" {
                    let entry = current_entry
                        .as_mut()
                        .ok_or(XmlInspectionError::ManifestStructure { offset })?;
                    if entry.encrypted {
                        return Err(XmlInspectionError::ManifestStructure { offset });
                    }
                    entry.encrypted = true;
                }
            },
            Event::Empty(element) => {
                let manifest_element = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE);
                if manifest_element && element.local_name().as_ref() == b"encryption-data" {
                    let entry = current_entry
                        .as_mut()
                        .ok_or(XmlInspectionError::ManifestStructure { offset })?;
                    if entry.encrypted {
                        return Err(XmlInspectionError::ManifestStructure { offset });
                    }
                    entry.encrypted = true;
                }
            },
            Event::End(element) => {
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                    && element.local_name().as_ref() == b"file-entry"
                {
                    let entry = current_entry
                        .take()
                        .ok_or(XmlInspectionError::ManifestStructure { offset })?;
                    if entry.encrypted {
                        count = count.saturating_add(1);
                        content_xml |= entry.content_xml;
                    }
                }
            },
            Event::DocType(_) => {
                return Err(XmlInspectionError::Malformed {
                    offset,
                    doctype: true,
                });
            },
            Event::Eof => {
                return Ok(EncryptionPresence { count, content_xml });
            },
            Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

struct EntryEncryptionPresence {
    content_xml: bool,
    encrypted: bool,
}

fn manifest_entry_is_content(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    offset: u64,
) -> Result<bool, XmlInspectionError> {
    for result in element.attributes() {
        let attribute = result.map_err(|_| XmlInspectionError::Malformed {
            offset,
            doctype: false,
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
            && local.as_ref() == b"full-path"
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| value == CONTENT_PATH)
                .map_err(|_| XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                });
        }
    }
    Ok(false)
}

fn inspect_xml(
    xml: &[u8],
    limits: OdfValidationLimits,
) -> Result<XmlObservations, XmlInspectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut events = 0_usize;
    let mut depth = 0_usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_is_manifest = false;
    let mut manifest_entries = 0_u64;
    let mut external_references = 0_u64;

    loop {
        if events >= limits.max_xml_events {
            return Err(XmlInspectionError::Limit(
                "XML part exceeds the configured ODF event ceiling",
            ));
        }
        let offset = reader.buffer_position();
        let event = reader
            .read_event()
            .map_err(|_| XmlInspectionError::Malformed {
                offset,
                doctype: false,
            })?;
        events += 1;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(XmlInspectionError::Malformed {
                            offset,
                            doctype: false,
                        });
                    }
                    root_seen = true;
                    root_is_manifest = matches!(
                        reader.resolver().resolve_element(element.name()).0,
                        ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE
                    ) && element.local_name().as_ref() == b"manifest";
                }
                if depth >= limits.max_xml_depth {
                    return Err(XmlInspectionError::Limit(
                        "XML part exceeds the configured ODF depth ceiling",
                    ));
                }
                depth += 1;
                if matches!(
                    reader.resolver().resolve_element(element.name()).0,
                    ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE
                ) && element.local_name().as_ref() == b"file-entry"
                {
                    manifest_entries = manifest_entries.saturating_add(1);
                }
                inspect_attributes(&reader, &element, &mut external_references, offset)?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(XmlInspectionError::Malformed {
                            offset,
                            doctype: false,
                        });
                    }
                    root_seen = true;
                    root_closed = true;
                    root_is_manifest = matches!(
                        reader.resolver().resolve_element(element.name()).0,
                        ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE
                    ) && element.local_name().as_ref() == b"manifest";
                }
                if depth >= limits.max_xml_depth {
                    return Err(XmlInspectionError::Limit(
                        "XML part exceeds the configured ODF depth ceiling",
                    ));
                }
                if matches!(
                    reader.resolver().resolve_element(element.name()).0,
                    ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE
                ) && element.local_name().as_ref() == b"file-entry"
                {
                    manifest_entries = manifest_entries.saturating_add(1);
                }
                inspect_attributes(&reader, &element, &mut external_references, offset)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(XmlInspectionError::Malformed {
                    offset,
                    doctype: true,
                });
            },
            Event::Eof if root_seen && root_closed && depth == 0 => {
                return Ok(XmlObservations {
                    external_references,
                    root_is_manifest,
                    manifest_entries,
                });
            },
            Event::Eof => {
                return Err(XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                });
            },
            Event::Text(text) if depth == 0 => {
                let bytes: &[u8] = text.as_ref();
                if !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(XmlInspectionError::Malformed {
                        offset,
                        doctype: false,
                    });
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                });
            },
            Event::GeneralRef(reference) if !valid_xml_reference(&reference) => {
                return Err(XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                });
            },
            Event::Decl(_)
            | Event::PI(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

pub(crate) fn valid_xml_reference(reference: &BytesRef<'_>) -> bool {
    let bytes: &[u8] = reference;
    matches!(bytes, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
        || (reference.is_char_ref()
            && reference
                .resolve_char_ref()
                .ok()
                .flatten()
                .is_some_and(xml10_character))
}

fn inspect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    external_references: &mut u64,
    offset: u64,
) -> Result<(), XmlInspectionError> {
    for result in element.attributes() {
        let attribute = result.map_err(|_| XmlInspectionError::Malformed {
            offset,
            doctype: false,
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE)
            && local.as_ref() == b"href"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|_| XmlInspectionError::Malformed {
                    offset,
                    doctype: false,
                })?;
            if provably_external(value.as_bytes()) {
                *external_references = external_references.saturating_add(1);
            }
        }
    }
    Ok(())
}

fn provably_external(value: &[u8]) -> bool {
    let value = trim_xml_schema_whitespace(value);
    if value.is_empty() || value.starts_with(b"#") {
        return false;
    }
    !safe_package_href(value)
}

fn trim_xml_schema_whitespace(mut value: &[u8]) -> &[u8] {
    while value
        .first()
        .is_some_and(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        value = &value[1..];
    }
    while value
        .last()
        .is_some_and(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        value = &value[..value.len() - 1];
    }
    value
}

pub(crate) fn safe_package_href(value: &[u8]) -> bool {
    if !decoded_href_is_utf8(value) {
        return false;
    }
    let mut decoded = DecodedHref::new(value);
    let mut depth = 0_usize;
    let mut segment = HrefSegment::new();
    let mut scheme_candidate = true;
    let mut scheme_index = 0_usize;

    while let Some(Ok(byte)) = decoded.next() {
        if matches!(byte, b'\\' | b'?' | b'#') || byte == 0 || byte.is_ascii_control() {
            return false;
        }
        if byte == b':' && scheme_candidate && scheme_index != 0 {
            return false;
        }
        if scheme_candidate {
            scheme_candidate = if scheme_index == 0 {
                byte.is_ascii_alphabetic()
            } else {
                byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.')
            };
            scheme_index = scheme_index.saturating_add(1);
        }
        if byte == b'/' {
            if scheme_index == 1 && segment.is_empty() {
                return false;
            }
            if !finish_href_segment(segment, &mut depth) {
                return false;
            }
            segment = HrefSegment::new();
            scheme_candidate = false;
        } else {
            segment.push(byte);
        }
    }
    finish_href_segment(segment, &mut depth) && depth != 0
}

fn finish_href_segment(segment: HrefSegment, depth: &mut usize) -> bool {
    if segment.is_empty() || segment.is_dot() {
        return true;
    }
    if segment.is_dot_dot() {
        let Some(next) = depth.checked_sub(1) else {
            return false;
        };
        *depth = next;
        return true;
    }
    if *depth == 0 && (segment.is_mimetype() || segment.is_meta_inf()) {
        return false;
    }
    *depth = depth.saturating_add(1);
    true
}

#[derive(Clone, Copy)]
struct HrefSegment {
    length: usize,
    dot: bool,
    dot_dot: bool,
    mimetype: bool,
    meta_inf: bool,
}

impl HrefSegment {
    const fn new() -> Self {
        Self {
            length: 0,
            dot: true,
            dot_dot: true,
            mimetype: true,
            meta_inf: true,
        }
    }

    fn push(&mut self, byte: u8) {
        const MIMETYPE: &[u8] = b"mimetype";
        const META_INF: &[u8] = b"META-INF";
        self.dot &= self.length == 0 && byte == b'.';
        self.dot_dot &= self.length < 2 && byte == b'.';
        self.mimetype &= MIMETYPE.get(self.length).copied() == Some(byte);
        self.meta_inf &= META_INF.get(self.length).copied() == Some(byte);
        self.length = self.length.saturating_add(1);
    }

    const fn is_empty(self) -> bool {
        self.length == 0
    }

    const fn is_dot(self) -> bool {
        self.dot && self.length == 1
    }

    const fn is_dot_dot(self) -> bool {
        self.dot_dot && self.length == 2
    }

    const fn is_mimetype(self) -> bool {
        self.mimetype && self.length == 8
    }

    const fn is_meta_inf(self) -> bool {
        self.meta_inf && self.length == 8
    }
}

#[derive(Clone, Copy)]
struct DecodedHref<'a> {
    value: &'a [u8],
    index: usize,
}

impl<'a> DecodedHref<'a> {
    const fn new(value: &'a [u8]) -> Self {
        Self { value, index: 0 }
    }
}

impl Iterator for DecodedHref<'_> {
    type Item = Result<u8, ()>;

    fn next(&mut self) -> Option<Self::Item> {
        let byte = *self.value.get(self.index)?;
        self.index += 1;
        if byte != b'%' {
            return Some(Ok(byte));
        }
        let Some(high) = self.value.get(self.index).copied().and_then(hex_value) else {
            return Some(Err(()));
        };
        let Some(low) = self.value.get(self.index + 1).copied().and_then(hex_value) else {
            return Some(Err(()));
        };
        self.index += 2;
        Some(Ok((high << 4) | low))
    }
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn decoded_href_is_utf8(value: &[u8]) -> bool {
    let mut bytes = DecodedHref::new(value);
    while let Some(result) = bytes.next() {
        let Ok(first) = result else {
            return false;
        };
        let continuations = match first {
            0x00..=0x7f => continue,
            0xc2..=0xdf => 1,
            0xe0..=0xef => 2,
            0xf0..=0xf4 => 3,
            _ => return false,
        };
        let Some(Ok(second)) = bytes.next() else {
            return false;
        };
        if !(0x80..=0xbf).contains(&second)
            || (first == 0xe0 && second < 0xa0)
            || (first == 0xed && second > 0x9f)
            || (first == 0xf0 && second < 0x90)
            || (first == 0xf4 && second > 0x8f)
        {
            return false;
        }
        for _ in 1..continuations {
            if !matches!(bytes.next(), Some(Ok(0x80..=0xbf))) {
                return false;
            }
        }
    }
    true
}

#[derive(Clone, Copy)]
struct RawCatalogObservations {
    hostile_paths: u64,
    mimetype_local_extra: bool,
    mimetype_repairable: bool,
    mimetype_local_first: bool,
    mimetype_central_first: bool,
}

fn inspect_raw_catalog(
    data: &[u8],
    limits: ArchiveLimits,
) -> Result<RawCatalogObservations, ZipError> {
    let archive = ZipArchive::from_slice(data)?;
    let mut hostile_paths = 0_u64;
    let mut mimetype_local_extra = false;
    let mut mimetype_local_extra_repairable = false;
    let mut mimetype_central_extra = false;
    let mut first_local_offset = None;
    let mut mimetype_local_offset = None;
    let mut next_local_offset = None;
    let mut mimetype_count = 0_usize;
    let mut mimetype_local_header_valid = false;
    let mut mimetype_local_payload_end = None;
    let mut first_central_name_is_mimetype = false;
    let mut first_central_seen = false;
    let mut files = 0_usize;
    let mut metadata_bytes = 0_u64;
    let mut total_size = 0_u64;
    for entry in archive.entries() {
        let entry = entry?;
        let raw_path = entry.file_path();
        if !first_central_seen {
            first_central_seen = true;
            first_central_name_is_mimetype = raw_path.as_ref() == MIMETYPE_PATH.as_bytes();
        }
        let local_offset = entry.local_header_offset();
        first_local_offset =
            Some(first_local_offset.map_or(local_offset, |first: u64| first.min(local_offset)));
        if local_offset > 0 {
            next_local_offset =
                Some(next_local_offset.map_or(local_offset, |next: u64| next.min(local_offset)));
        }
        let name_bytes = raw_path.as_ref().len() as u64;
        enforce_zip_limit(
            LimitResource::MemberNameBytes,
            name_bytes,
            limits.max_member_name_bytes,
        )?;
        let comment_bytes = central_file_comment_len(data, entry.central_directory_offset());
        let entry_metadata = name_bytes
            .saturating_add(
                u64::try_from(entry.extra_fields().remaining_bytes().len()).unwrap_or(u64::MAX),
            )
            .saturating_add(comment_bytes);
        metadata_bytes = metadata_bytes.saturating_add(entry_metadata);
        enforce_zip_limit(
            LimitResource::MetadataBytes,
            metadata_bytes,
            limits.max_metadata_bytes,
        )?;
        if !canonical_odf_path(raw_path.as_ref(), entry.is_dir()) {
            hostile_paths = hostile_paths.saturating_add(1);
        }
        if raw_path.as_ref() == MIMETYPE_PATH.as_bytes() {
            mimetype_count = mimetype_count.saturating_add(1);
            mimetype_local_offset = Some(local_offset);
            let mut central_fields = entry.extra_fields();
            mimetype_central_extra =
                central_fields.next().is_some() || !central_fields.remaining_bytes().is_empty();
            let local = archive.get_entry(entry.wayfinder())?;
            let mut fields = local.extra_fields();
            if fields.next().is_some() || !fields.remaining_bytes().is_empty() {
                mimetype_local_extra = true;
                mimetype_local_extra_repairable =
                    is_repairable_mimetype_extra(local.extra_fields());
            }

            // Keep the repair authorization predicate in lockstep with the
            // raw repair layout. In particular, this rejects descriptors,
            // flag/compression/CRC/size mismatches, and any opaque bytes after
            // the first stored payload before advertising the sole repair
            // diagnostic.
            let local_start = usize::try_from(local_offset).ok();
            let local_fixed = local_start
                .and_then(|start| start.checked_add(30).and_then(|end| data.get(start..end)));
            let central_start = usize::try_from(entry.central_directory_offset()).ok();
            let central_fixed = central_start
                .and_then(|start| start.checked_add(46).and_then(|end| data.get(start..end)));
            if let (Some(local_fixed), Some(central_fixed)) = (local_fixed, central_fixed) {
                let local_flags = validation_le_u16(local_fixed, 6);
                let central_flags = validation_le_u16(central_fixed, 8);
                let local_compression = validation_le_u16(local_fixed, 8);
                let central_compression = validation_le_u16(central_fixed, 10);
                let local_name_len = validation_le_u16(local_fixed, 26).map(usize::from);
                let local_extra_len = validation_le_u16(local_fixed, 28).map(usize::from);
                let central_name_len = validation_le_u16(central_fixed, 28).map(usize::from);
                let local_payload_end =
                    local_name_len
                        .zip(local_extra_len)
                        .and_then(|(name_len, extra_len)| {
                            30usize
                                .checked_add(name_len)
                                .and_then(|header_end| header_end.checked_add(extra_len))
                                .and_then(|payload_start| {
                                    usize::try_from(entry.compressed_size_hint())
                                        .ok()
                                        .and_then(|size| payload_start.checked_add(size))
                                })
                        });
                let local_name = local_name_len.and_then(|name_len| {
                    local_start.and_then(|start| {
                        let name_start = start.checked_add(30)?;
                        data.get(name_start..name_start.checked_add(name_len)?)
                    })
                });
                let central_name = central_name_len.and_then(|name_len| {
                    data.get(
                        central_start?.checked_add(46)?
                            ..central_start?.checked_add(46usize.checked_add(name_len)?)?,
                    )
                });
                let local_crc = validation_le_u32(local_fixed, 14);
                let central_crc = validation_le_u32(central_fixed, 16);
                let local_compressed = validation_le_u32(local_fixed, 18);
                let local_uncompressed = validation_le_u32(local_fixed, 22);
                let central_compressed = validation_le_u32(central_fixed, 20);
                let central_uncompressed = validation_le_u32(central_fixed, 24);
                let entry_compressed = u32::try_from(entry.compressed_size_hint()).ok();
                let entry_uncompressed = u32::try_from(entry.uncompressed_size_hint()).ok();
                mimetype_local_payload_end = local_payload_end.map(|end| end as u64);
                mimetype_local_header_valid = le_u32_validation(local_fixed, 0)
                    == Some(0x0403_4b50)
                    && local_flags == central_flags
                    && local_flags.is_some_and(|flags| flags & !0x0800 == 0)
                    && local_compression == Some(0)
                    && central_compression == Some(0)
                    && local_name == Some(MIMETYPE_PATH.as_bytes())
                    && central_name == Some(MIMETYPE_PATH.as_bytes())
                    && local_name_len == central_name_len
                    && local_crc == central_crc
                    && local_compressed == central_compressed
                    && local_uncompressed == central_uncompressed
                    && local_compressed == entry_compressed
                    && local_uncompressed == entry_uncompressed;
            }
        }
        if entry.is_dir() {
            continue;
        }
        files = files.saturating_add(1);
        enforce_zip_limit(
            LimitResource::FileCount,
            files as u64,
            limits.max_files as u64,
        )?;
        enforce_zip_limit(
            LimitResource::CompressedSize,
            entry.compressed_size_hint(),
            limits.max_compressed_size,
        )?;
        enforce_zip_limit(
            LimitResource::EntrySize,
            entry.uncompressed_size_hint(),
            limits.max_entry_size,
        )?;
        total_size = total_size.saturating_add(entry.uncompressed_size_hint());
        enforce_zip_limit(LimitResource::TotalSize, total_size, limits.max_total_size)?;
    }
    let local_span_end = next_local_offset.unwrap_or_else(|| {
        // `directory_offset` is a parsed ZIP offset and is therefore already
        // bounded by the caller's successful archive parse.
        archive.directory_offset()
    });
    let mimetype_repairable = mimetype_count == 1
        && first_central_name_is_mimetype
        && mimetype_local_offset == Some(0)
        && mimetype_local_offset == first_local_offset
        && mimetype_local_header_valid
        && mimetype_local_extra_repairable
        && !mimetype_central_extra
        && mimetype_local_payload_end == Some(local_span_end);
    Ok(RawCatalogObservations {
        hostile_paths,
        mimetype_local_extra,
        mimetype_repairable,
        mimetype_local_first: mimetype_local_offset == Some(0)
            && mimetype_local_offset == first_local_offset,
        mimetype_central_first: first_central_name_is_mimetype,
    })
}

fn validation_le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    bytes
        .get(offset..offset.checked_add(2)?)
        .and_then(|slice| slice.try_into().ok())
        .map(u16::from_le_bytes)
}

fn validation_le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    bytes
        .get(offset..offset.checked_add(4)?)
        .and_then(|slice| slice.try_into().ok())
        .map(u32::from_le_bytes)
}

fn le_u32_validation(bytes: &[u8], offset: usize) -> Option<u32> {
    validation_le_u32(bytes, offset)
}

fn is_repairable_mimetype_extra(mut fields: soapberry_zip::extra_fields::ExtraFields<'_>) -> bool {
    // The only local metadata this narrow repair is willing to remove is one
    // well-formed Extended Timestamp field.  No unknown, repeated, or
    // malformed field is silently discarded.
    let mut seen = false;
    for (id, body) in fields.by_ref() {
        if seen || id != soapberry_zip::extra_fields::ExtraFieldId::EXTENDED_TIMESTAMP {
            return false;
        }
        seen = true;
        if !valid_extended_timestamp(body) {
            return false;
        }
    }
    seen && fields.remaining_bytes().is_empty()
}

fn valid_extended_timestamp(body: &[u8]) -> bool {
    let Some(&flags) = body.first() else {
        return false;
    };
    if flags & !0b111 != 0 {
        return false;
    }
    let timestamps = flags.count_ones() as usize;
    body.len() == 1 + timestamps * 4
}

fn central_file_comment_len(data: &[u8], central_offset: u64) -> u64 {
    let Some(start) = usize::try_from(central_offset)
        .ok()
        .and_then(|offset| offset.checked_add(32))
    else {
        return u64::MAX;
    };
    let Some(bytes) = data.get(start..start.saturating_add(2)) else {
        return u64::MAX;
    };
    u64::from(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn enforce_zip_limit(resource: LimitResource, actual: u64, maximum: u64) -> Result<(), ZipError> {
    if actual > maximum {
        return Err(ZipErrorKind::LimitExceeded {
            resource,
            actual,
            maximum,
        }
        .into());
    }
    Ok(())
}

fn canonical_odf_path(raw: &[u8], directory: bool) -> bool {
    let Ok(text) = str::from_utf8(raw) else {
        return false;
    };
    if text.is_empty()
        || text.starts_with('/')
        || text.contains('\\')
        || text.as_bytes().iter().any(|byte| byte.is_ascii_control())
        || text.split('/').any(|part| part == "." || part == "..")
        || text.contains("//")
    {
        return false;
    }
    directory == text.ends_with('/')
}

fn macro_storage_path(path: &str) -> bool {
    path == "Basic"
        || path.starts_with("Basic/")
        || path == "Scripts"
        || path.starts_with("Scripts/")
}

fn odf_mimetype(bytes: &[u8]) -> bool {
    str::from_utf8(bytes)
        .ok()
        .is_some_and(|value| crate::constants::ODF_MIMETYPES.contains_key(value))
}

fn allocation_failure(error: &ZipError) -> bool {
    matches!(error.kind(), ZipErrorKind::InvalidInput { msg } if msg.contains("could not allocate"))
}

fn zip_limit_reason(error: &ZipError) -> &'static str {
    match error.kind() {
        ZipErrorKind::LimitExceeded { resource, .. } => match resource {
            LimitResource::FileCount => "ZIP member count exceeds the configured ODF ceiling",
            LimitResource::MemberNameBytes => {
                "ZIP member name exceeds the configured ODF byte ceiling"
            },
            LimitResource::MetadataBytes => {
                "ZIP catalog metadata exceeds the configured ODF byte ceiling"
            },
            LimitResource::CompressedSize => {
                "ZIP member compressed size exceeds the configured ODF ceiling"
            },
            LimitResource::EntrySize => {
                "ZIP member declared size exceeds the configured ODF ceiling"
            },
            LimitResource::TotalSize => {
                "ZIP total declared size exceeds the configured ODF ceiling"
            },
        },
        _ => "ZIP ingress exceeds a configured ODF ceiling",
    }
}

fn declaration_read_rejection(
    error: ZipError,
    limits: ValidationLimits,
    state: &mut ValidationState,
) -> Result<(), OdfValidationError> {
    if allocation_failure(&error) {
        return Err(OdfValidationError::Allocation {
            resource: "ODF declaration payload",
        });
    }
    state.declarations = CheckStatus::Complete;
    state.push_issue(diagnostic_issue(
        DECLARATIONS,
        "odf.declarations.unreadable",
        "A required ODF declaration member could not be decoded and verified.",
        "package-declarations",
        &error.to_string(),
        limits,
    )?)
}

fn content_read_rejection(
    error: ZipError,
    limits: ValidationLimits,
    state: &mut ValidationState,
) -> Result<(), OdfValidationError> {
    if allocation_failure(&error) {
        return Err(OdfValidationError::Allocation {
            resource: "ODF content XML payload",
        });
    }
    state.root_xml = CheckStatus::Complete;
    state.external = CheckStatus::blocked(
        "external-reference scan could not finish because content.xml was unreadable",
        limits,
    )?;
    state.push_issue(diagnostic_issue(
        ROOT_XML,
        "odf.xml.unreadable",
        "The required ODF content XML member could not be decoded and verified.",
        CONTENT_PATH,
        &error.to_string(),
        limits,
    )?)
}

fn ingress_blocked_report(
    reason: &'static str,
    limits: ValidationLimits,
) -> Result<ValidateReport, OdfValidationError> {
    let ingress_id = id(INGRESS, limits)?;
    let checks = [
        ValidationCheck::new(ingress_id.clone(), CheckStatus::blocked(reason, limits)?),
        check(CATALOG, CheckStatus::stopped_by(ingress_id.clone()), limits)?,
        check(
            DECLARATIONS,
            CheckStatus::stopped_by(ingress_id.clone()),
            limits,
        )?,
        check(
            ROOT_XML,
            CheckStatus::stopped_by(ingress_id.clone()),
            limits,
        )?,
        check(
            ENCRYPTION,
            CheckStatus::stopped_by(ingress_id.clone()),
            limits,
        )?,
        check(
            SIGNATURES,
            CheckStatus::stopped_by(ingress_id.clone()),
            limits,
        )?,
        check(
            EXTERNAL,
            CheckStatus::stopped_by(ingress_id.clone()),
            limits,
        )?,
        check(MACROS, CheckStatus::stopped_by(ingress_id), limits)?,
    ];
    ValidateReport::try_new(checks, [], limits).map_err(Into::into)
}

fn ingress_rejected_report(
    error: &ZipError,
    source_size: u64,
    limits: ValidationLimits,
) -> Result<ValidateReport, OdfValidationError> {
    let blocked = CheckStatus::blocked(
        "ZIP ingress was conclusively rejected before this capability could run",
        limits,
    )?;
    let checks = [
        check(INGRESS, CheckStatus::Complete, limits)?,
        check(CATALOG, blocked.clone(), limits)?,
        check(DECLARATIONS, blocked.clone(), limits)?,
        check(ROOT_XML, blocked.clone(), limits)?,
        check(ENCRYPTION, blocked.clone(), limits)?,
        check(SIGNATURES, blocked.clone(), limits)?,
        check(EXTERNAL, blocked.clone(), limits)?,
        check(MACROS, blocked, limits)?,
    ];
    let issue = ValidationIssue::try_new(
        id(INGRESS, limits)?,
        "odf.zip.invalid",
        IssueSeverity::Fatal,
        "The input is not a structurally readable ZIP package.",
        [location("zip-package", limits)?],
        [
            IssueEvidence::try_new("source_size", EvidenceValue::Size(source_size), limits)?,
            IssueEvidence::try_new(
                "diagnostic_sha256",
                EvidenceValue::Sha256(EvidenceDigest::of(error.to_string().as_bytes())),
                limits,
            )?,
        ],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?;
    ValidateReport::try_new(checks, [issue], limits).map_err(Into::into)
}

fn simple_issue(
    check_name: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    part: &str,
    compatibility: CompatibilityImpact,
    limits: ValidationLimits,
) -> Result<ValidationIssue, OdfValidationError> {
    ValidationIssue::try_new(
        id(check_name, limits)?,
        code,
        severity,
        message,
        [location(part, limits)?],
        [],
        None,
        compatibility,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn mimetype_extra_issue(
    source_size: u64,
    source_digest: EvidenceDigest,
    limits: ValidationLimits,
) -> Result<ValidationIssue, OdfValidationError> {
    ValidationIssue::try_new(
        id(DECLARATIONS, limits)?,
        "odf.mimetype.local_header_extra",
        IssueSeverity::Error,
        "The ODF mimetype has one recognized removable local-header extra field.",
        [location(MIMETYPE_PATH, limits)?],
        [
            IssueEvidence::try_new("source_size", EvidenceValue::Size(source_size), limits)?,
            IssueEvidence::try_new(
                "source_sha256",
                EvidenceValue::Sha256(source_digest),
                limits,
            )?,
        ],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::available(MIMETYPE_REPAIR_ID, limits)?,
        limits,
    )
    .map_err(Into::into)
}

fn count_issue(
    check_name: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    part: &str,
    evidence_key: &str,
    count: u64,
    compatibility: CompatibilityImpact,
    limits: ValidationLimits,
) -> Result<ValidationIssue, OdfValidationError> {
    ValidationIssue::try_new(
        id(check_name, limits)?,
        code,
        severity,
        message,
        [location(part, limits)?],
        [IssueEvidence::try_new(
            evidence_key,
            EvidenceValue::Count(count),
            limits,
        )?],
        None,
        compatibility,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn diagnostic_issue(
    check_name: &str,
    code: &str,
    message: &str,
    part: &str,
    diagnostic: &str,
    limits: ValidationLimits,
) -> Result<ValidationIssue, OdfValidationError> {
    ValidationIssue::try_new(
        id(check_name, limits)?,
        code,
        IssueSeverity::Error,
        message,
        [location(part, limits)?],
        [IssueEvidence::try_new(
            "diagnostic_sha256",
            EvidenceValue::Sha256(EvidenceDigest::of(diagnostic.as_bytes())),
            limits,
        )?],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn offset_issue(
    check_name: &str,
    code: &str,
    message: &str,
    part: &str,
    offset: u64,
    limits: ValidationLimits,
) -> Result<ValidationIssue, OdfValidationError> {
    ValidationIssue::try_new(
        id(check_name, limits)?,
        code,
        IssueSeverity::Error,
        message,
        [IssueLocation::try_new(
            Some(part),
            None,
            Some(offset),
            None,
            None,
            limits,
        )?],
        [],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn location(part: &str, limits: ValidationLimits) -> Result<IssueLocation, OdfValidationError> {
    IssueLocation::try_new(Some(part), None, None, None, None, limits).map_err(Into::into)
}

fn presence_status(count: u64) -> CheckStatus {
    if count == 0 {
        CheckStatus::NotApplicable
    } else {
        CheckStatus::Complete
    }
}

fn id(value: &str, limits: ValidationLimits) -> Result<CheckCapabilityId, OdfValidationError> {
    CheckCapabilityId::try_new(value, limits).map_err(Into::into)
}

fn check(
    value: &str,
    status: CheckStatus,
    limits: ValidationLimits,
) -> Result<ValidationCheck, OdfValidationError> {
    Ok(ValidationCheck::new(id(value, limits)?, status))
}
