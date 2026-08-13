//! Bounded, non-mutating validation of a DOCX/WordprocessingML source.
//!
//! The validator builds on the source-backed OPC catalog. ZIP metadata,
//! content types, and relationship manifests are checked by `litchi-opc`;
//! only the main-document part is materialized for the format-semantic pass.
//! No document text, XML attributes, relationship targets, or macro payloads
//! are retained in the report. External targets are inventoried as inert
//! presence findings and are never fetched.

use std::{
    collections::HashSet,
    error::Error as StdError,
    fmt, io,
    num::{NonZeroU64, NonZeroUsize},
    sync::{Arc, Mutex},
};

use litchi_core::{
    Budget, CancellationSource, CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceValue,
    ExecutionContext, ExecutionError, ExecutionLimits, IssueEvidence, IssueLocation, IssueSeverity,
    Limits as CoreLimits, ReadAt, RepairAvailability, ValidateReport, ValidationCheck,
    ValidationIssue, ValidationLimits, ValidationReportError,
};
use litchi_opc::{
    OpcError, PartView, ReadLimits, SourceBackedPackage,
    constants::{content_type as ct, relationship_type as rt},
};
use quick_xml::{
    events::Event,
    name::ResolveResult,
    reader::{NsReader, Reader},
};

const INGRESS: &str = "docx.package.ingress";
const MAIN_DOCUMENT: &str = "docx.package.main_document";
const MAIN_RELATIONSHIPS: &str = "docx.main_document.relationship_closure";
const MCE: &str = "docx.main_document.markup_compatibility";
const SEMANTICS: &str = "docx.main_document.semantics";
const EXTERNAL: &str = "docx.relationships.external_target_presence";
const MACROS: &str = "docx.security.macro_presence";
const EMBEDDED: &str = "docx.security.embedded_content_presence";
const SIGNATURES: &str = "docx.security.signature_presence";
const SIGNATURE_DIRECTORY: &str = "/_xmlsignatures/";
// quick-xml's namespace reader tracks nesting in a u16.  Keep the public
// policy strictly below that representation's maximum so an attacker cannot
// make the reader overflow before the validator observes its own ceiling.
const MAX_SAFE_XML_DEPTH: usize = u16::MAX as usize - 1;
const STRICT_RELATIONSHIP_PREFIX: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const ACTIVE_X_BINARY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const ACTIVE_X_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControl";
const ACTIVE_X_DESCRIPTOR_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeX";

/// Default finite limits for one DOCX semantic validation pass.
pub const DEFAULT_DOCX_VALIDATION_LIMITS: DocxValidationLimits = DocxValidationLimits {
    max_main_document_bytes: 64 * 1024 * 1024,
    max_xml_events: 4_000_000,
    max_xml_depth: 4_096,
    max_mce_output_bytes: 128 * 1024 * 1024,
    max_mce_directive_tokens: 4_096,
    max_mce_choices_per_alternate: 1_024,
};

/// Finite bounds for the DOCX-specific semantic validation pass.
///
/// ZIP member, declared-size, content-type, and relationship ceilings remain
/// owned by [`ReadLimits`]. These bounds only govern the materialized main
/// document and the XML/MCE work performed after OPC ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocxValidationLimits {
    max_main_document_bytes: u64,
    max_xml_events: usize,
    max_xml_depth: usize,
    max_mce_output_bytes: usize,
    max_mce_directive_tokens: usize,
    max_mce_choices_per_alternate: usize,
}

impl DocxValidationLimits {
    /// Returns a copy with a new main-document byte ceiling.
    #[must_use]
    pub const fn with_max_main_document_bytes(mut self, maximum: u64) -> Self {
        self.max_main_document_bytes = maximum;
        self
    }

    /// Returns a copy with a new XML event ceiling.
    #[must_use]
    pub const fn with_max_xml_events(mut self, maximum: usize) -> Self {
        self.max_xml_events = maximum;
        self
    }

    /// Returns a copy with a new XML nesting-depth ceiling.
    #[must_use]
    pub const fn with_max_xml_depth(mut self, maximum: usize) -> Self {
        self.max_xml_depth = if maximum > MAX_SAFE_XML_DEPTH {
            MAX_SAFE_XML_DEPTH
        } else {
            maximum
        };
        self
    }

    /// Returns a copy with a new MCE output byte ceiling.
    #[must_use]
    pub const fn with_max_mce_output_bytes(mut self, maximum: usize) -> Self {
        self.max_mce_output_bytes = maximum;
        self
    }

    /// Returns a copy with a new MCE directive-token ceiling.
    #[must_use]
    pub const fn with_max_mce_directive_tokens(mut self, maximum: usize) -> Self {
        self.max_mce_directive_tokens = maximum;
        self
    }

    /// Returns a copy with a new MCE-choice ceiling.
    #[must_use]
    pub const fn with_max_mce_choices_per_alternate(mut self, maximum: usize) -> Self {
        self.max_mce_choices_per_alternate = maximum;
        self
    }

    /// Maximum materialized main-document bytes.
    #[must_use]
    pub const fn max_main_document_bytes(self) -> u64 {
        self.max_main_document_bytes
    }

    /// Maximum XML events inspected in the visible main document.
    #[must_use]
    pub const fn max_xml_events(self) -> usize {
        self.max_xml_events
    }

    /// Maximum XML nesting depth inspected in the visible main document.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.max_xml_depth
    }

    /// Maximum bytes retained by MCE preprocessing.
    #[must_use]
    pub const fn max_mce_output_bytes(self) -> usize {
        self.max_mce_output_bytes
    }

    /// Maximum MCE directive tokens accepted by preprocessing.
    #[must_use]
    pub const fn max_mce_directive_tokens(self) -> usize {
        self.max_mce_directive_tokens
    }

    /// Maximum choices accepted in one MCE alternate-content element.
    #[must_use]
    pub const fn max_mce_choices_per_alternate(self) -> usize {
        self.max_mce_choices_per_alternate
    }
}

impl Default for DocxValidationLimits {
    fn default() -> Self {
        DEFAULT_DOCX_VALIDATION_LIMITS
    }
}

const fn effective_xml_depth(limits: DocxValidationLimits) -> usize {
    if limits.max_xml_depth > MAX_SAFE_XML_DEPTH {
        MAX_SAFE_XML_DEPTH
    } else {
        limits.max_xml_depth
    }
}

/// Failure to perform or retain a DOCX validation report.
///
/// Definite package, XML, and unsupported-semantic findings are returned as
/// issues in a successful report. This error is reserved for source I/O or
/// instability, bounded allocation, and bounded report-construction failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum DocxValidationError {
    /// The immutable OPC source could not be read consistently.
    Ingress(OpcError),
    /// A bounded validator-owned allocation failed.
    Allocation {
        /// Content-free resource name.
        resource: &'static str,
    },
    /// The shared bounded report could not retain the result.
    Report(ValidationReportError),
}

impl fmt::Display for DocxValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ingress(error) => write!(formatter, "DOCX validation ingress failed: {error}"),
            Self::Allocation { resource } => {
                write!(
                    formatter,
                    "allocation failed while validating DOCX {resource}"
                )
            },
            Self::Report(error) => write!(formatter, "DOCX validation report failed: {error}"),
        }
    }
}

impl StdError for DocxValidationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Ingress(error) => Some(error),
            Self::Allocation { .. } => None,
            Self::Report(error) => Some(error),
        }
    }
}

impl From<ValidationReportError> for DocxValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

/// Result returned by the DOCX validation-report API.
pub type Result<T> = std::result::Result<T, DocxValidationError>;

/// Validate an immutable positional DOCX source under finite default limits.
///
/// The source is never changed, the ZIP package is not rewritten, external
/// targets are not fetched, macros are not executed, and signatures are not
/// cryptographically verified.
pub fn validate_read_at(source: Arc<dyn ReadAt>) -> Result<ValidateReport> {
    validate_read_at_with_limits(
        source,
        ReadLimits::default(),
        DocxValidationLimits::default(),
        ValidationLimits::default(),
    )
}

/// Validate an immutable positional DOCX source with explicit finite limits.
pub fn validate_read_at_with_limits(
    source: Arc<dyn ReadAt>,
    read_limits: ReadLimits,
    docx_limits: DocxValidationLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport> {
    let expected = source
        .version()
        .map_err(OpcError::from)
        .map_err(DocxValidationError::Ingress)?;
    let tracked = Arc::new(ValidationSource::new(Arc::clone(&source)));
    let ingress_source: Arc<dyn ReadAt> = tracked.clone();
    let report = match open_validation_package(ingress_source, read_limits, docx_limits) {
        Ok(package) => {
            validate_package_with_tracker(&package, docx_limits, report_limits, Some(&tracked))?
        },
        Err(error) => {
            if let Some(io_error) = tracked.take_error() {
                return Err(DocxValidationError::Ingress(OpcError::IoError(io_error)));
            }
            match error {
                OpcError::ReadLimit { .. } => blocked_ingress_report(report_limits)?,
                error if is_structural_rejection(&error) => {
                    rejected_ingress_report(&error, report_limits)?
                },
                error => return Err(DocxValidationError::Ingress(error)),
            }
        },
    };
    if let Some(io_error) = tracked.take_error() {
        return Err(DocxValidationError::Ingress(OpcError::IoError(io_error)));
    }
    let actual = source
        .version()
        .map_err(OpcError::from)
        .map_err(DocxValidationError::Ingress)?;
    if actual != expected {
        return Err(DocxValidationError::Ingress(OpcError::SourceChanged {
            expected,
            actual,
        }));
    }
    Ok(report)
}

fn open_validation_package(
    source: Arc<dyn ReadAt>,
    read_limits: ReadLimits,
    docx_limits: DocxValidationLimits,
) -> std::result::Result<SourceBackedPackage, OpcError> {
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let max_memory = docx_limits.max_main_document_bytes();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(max_memory.max(1)).unwrap_or(NonZeroU64::MAX),
        0,
    )
    .map_err(OpcError::Execution)?;
    let context = ExecutionContext::new(
        Budget::root(
            "docx-validation",
            CoreLimits::new(max_memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        ),
        cancellation,
        execution_limits,
    );
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source,
        read_limits,
        litchi_opc::SourceCacheLimits::default(),
        context,
    )
}

fn validate_package_with_tracker(
    package: &SourceBackedPackage,
    docx_limits: DocxValidationLimits,
    report_limits: ValidationLimits,
    tracker: Option<&ValidationSource>,
) -> Result<ValidateReport> {
    if let Some(tracker) = tracker {
        if let Some(error) = tracker.take_error() {
            return Err(DocxValidationError::Ingress(OpcError::IoError(error)));
        }
    }
    let ingress = check(INGRESS, CheckStatus::Complete, report_limits)?;
    let mut issues = Vec::new();
    issues
        .try_reserve_exact(12)
        .map_err(|_| DocxValidationError::Allocation {
            resource: "validation issue staging",
        })?;

    let main = match package.main_document_part() {
        Ok(part) => Some(part),
        Err(error) if is_structural_rejection(&error) => {
            issues.push(simple_issue(
                MAIN_DOCUMENT,
                "docx.main_document.relationship_invalid",
                "The DOCX package does not expose one valid internal main-document relationship.",
                "package-relationships",
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )?);
            None
        },
        Err(error) => return Err(DocxValidationError::Ingress(error)),
    };

    let mut main_status = CheckStatus::Complete;
    let main_path = main.as_ref().map(|part| part.partname().as_str());
    let main_content_type_valid = main.as_ref().is_some_and(|part| {
        crate::package::validate_document_main_content_type(part.content_type()).is_ok()
    });
    if main.is_none() {
        main_status = CheckStatus::blocked(
            "main-document relationship validation did not produce a semantic owner",
            report_limits,
        )?;
    }
    if let Some(main) = main.as_ref() {
        if let Err(error) = crate::package::validate_document_main_content_type(main.content_type())
        {
            issues.push(simple_issue(
                MAIN_DOCUMENT,
                "docx.main_document.content_type_invalid",
                "The DOCX main-document relationship targets an unsupported WordprocessingML content type.",
                main_path.unwrap_or("package-relationships"),
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )?);
            let _ = error;
            main_status = CheckStatus::blocked(
                "main-document content type is not supported by the DOCX semantic validator",
                report_limits,
            )?;
        }
    }

    let main_id = id(MAIN_DOCUMENT, report_limits)?;
    let main_relationship_status = if let Some(main) = main.as_ref() {
        if main_content_type_valid {
            inspect_main_relationships(package, main, report_limits, &mut issues)?
        } else {
            CheckStatus::stopped_by(main_id.clone())
        }
    } else {
        CheckStatus::stopped_by(main_id.clone())
    };
    let mut mce_status = if main_content_type_valid {
        CheckStatus::NotApplicable
    } else {
        CheckStatus::stopped_by(main_id.clone())
    };
    let mut semantic_status = CheckStatus::stopped_by(main_id);
    if let Some(main) = main.as_ref() {
        if main_content_type_valid {
            match inspect_main_document(
                main,
                main.partname().as_str(),
                docx_limits,
                report_limits,
                tracker,
            ) {
                Ok((mce, semantic_issues)) => {
                    mce_status = mce.status;
                    issues.extend(semantic_issues);
                    semantic_status = if mce.blocks_semantics {
                        CheckStatus::stopped_by(id(MCE, report_limits)?)
                    } else if mce_status.is_complete() {
                        CheckStatus::Complete
                    } else {
                        CheckStatus::stopped_by(id(MCE, report_limits)?)
                    };
                    if let Some(issue) = mce.issue {
                        issues.push(issue);
                    }
                },
                Err(InspectionFailure::Blocked(reason)) => {
                    mce_status = CheckStatus::blocked(reason, report_limits)?;
                    semantic_status = CheckStatus::stopped_by(id(MCE, report_limits)?);
                },
                Err(InspectionFailure::Issue(issue)) => {
                    issues.push(*issue);
                    mce_status = CheckStatus::NotApplicable;
                    semantic_status = CheckStatus::Complete;
                },
                Err(InspectionFailure::Ingress(error)) => {
                    return Err(DocxValidationError::Ingress(error));
                },
                Err(InspectionFailure::Allocation { resource }) => {
                    return Err(DocxValidationError::Allocation { resource });
                },
                Err(InspectionFailure::Report(error)) => {
                    return Err(DocxValidationError::Report(error));
                },
            }
        }
    }

    let external = presence_report(
        package,
        |part| {
            part.rels()
                .iter()
                .filter(|relationship| relationship.is_external())
                .count()
        },
        EXTERNAL,
        "docx.external_target.present",
        "External relationship targets are present; no target was fetched.",
        IssueSeverity::Warning,
        CompatibilityImpact::Security,
        report_limits,
        &mut issues,
    )?;
    let macros = security_report(
        package,
        |part| {
            u64::from(is_macro_content_type(part.content_type()))
                + part
                    .rels()
                    .iter()
                    .filter(|relationship| is_macro_relationship(relationship.reltype()))
                    .count() as u64
        },
        is_macro_relationship,
        MACROS,
        "docx.macro.storage_present",
        "Macro-enabled DOCX or VBA storage is present; macro behavior was not assessed.",
        IssueSeverity::Warning,
        report_limits,
        &mut issues,
    )?;
    let embedded = security_report(
        package,
        |part| {
            u64::from(is_embedded_content_type(part.content_type()))
                + part
                    .rels()
                    .iter()
                    .filter(|relationship| is_embedded_relationship(relationship.reltype()))
                    .count() as u64
        },
        is_embedded_relationship,
        EMBEDDED,
        "docx.embedded_content.present",
        "Embedded OLE or package content is present; payloads were retained inert and not activated.",
        IssueSeverity::Warning,
        report_limits,
        &mut issues,
    )?;
    let signatures = signature_report(package, report_limits, &mut issues)?;

    let checks = [
        ingress,
        check(MAIN_DOCUMENT, main_status, report_limits)?,
        check(MAIN_RELATIONSHIPS, main_relationship_status, report_limits)?,
        check(MCE, mce_status, report_limits)?,
        check(SEMANTICS, semantic_status, report_limits)?,
        check(EXTERNAL, external, report_limits)?,
        check(MACROS, macros, report_limits)?,
        check(EMBEDDED, embedded, report_limits)?,
        check(SIGNATURES, signatures, report_limits)?,
    ];
    ValidateReport::try_new(checks, issues, report_limits).map_err(Into::into)
}

struct MceInspection {
    status: CheckStatus,
    blocks_semantics: bool,
    issue: Option<ValidationIssue>,
}

enum InspectionFailure {
    Blocked(&'static str),
    Issue(Box<ValidationIssue>),
    Ingress(OpcError),
    Allocation { resource: &'static str },
    Report(ValidationReportError),
}

/// Retains the first positional-source I/O failure while the OPC reader maps
/// low-level archive errors into content-free ZIP diagnostics.
struct ValidationSource {
    source: Arc<dyn ReadAt>,
    first_io_error: Mutex<Option<io::Error>>,
}

fn inspect_main_relationships(
    package: &SourceBackedPackage,
    main: &PartView<'_>,
    limits: ValidationLimits,
    issues: &mut Vec<ValidationIssue>,
) -> Result<CheckStatus> {
    let mut invalid_targets = 0_u64;
    let mut missing_targets = 0_u64;
    let mut content_type_mismatches = 0_u64;
    let mut pending: Vec<PartView<'_>> = Vec::new();
    pending
        .try_reserve_exact(1)
        .map_err(|_| DocxValidationError::Allocation {
            resource: "main-document relationship traversal",
        })?;
    let main_part = package
        .part(main.partname())
        .map_err(DocxValidationError::Ingress)?;
    pending.push(main_part);
    let mut scheduled = HashSet::new();
    scheduled
        .try_reserve(package.iter_parts().count().saturating_add(1))
        .map_err(|_| DocxValidationError::Allocation {
            resource: "main-document relationship traversal",
        })?;
    scheduled.insert(std::ptr::from_ref(main.partname()) as usize);
    while let Some(source) = pending.pop() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = match relationship.target_partname() {
                Ok(target) => target,
                Err(_) => {
                    invalid_targets = invalid_targets.saturating_add(1);
                    continue;
                },
            };
            let target_part = match package.part(&target) {
                Ok(part) => part,
                Err(OpcError::PartNotFound(_)) => {
                    missing_targets = missing_targets.saturating_add(1);
                    continue;
                },
                Err(error) if is_structural_rejection(&error) => {
                    missing_targets = missing_targets.saturating_add(1);
                    continue;
                },
                Err(error) => return Err(DocxValidationError::Ingress(error)),
            };
            if target_part.content_type().is_empty()
                || !relationship_target_content_type_matches(
                    relationship.reltype(),
                    target_part.content_type(),
                )
            {
                content_type_mismatches = content_type_mismatches.saturating_add(1);
            }
            // Keep the validated package view itself in the queue.  The
            // catalog is immutable, so its part-name address is a stable,
            // allocation-free identity; repeated edges do not clone an
            // attacker-controlled PackURI into the pending queue.
            let target_identity = std::ptr::from_ref(target_part.partname()) as usize;
            if scheduled.insert(target_identity) {
                pending
                    .try_reserve(1)
                    .map_err(|_| DocxValidationError::Allocation {
                        resource: "main-document relationship traversal",
                    })?;
                pending.push(target_part);
            }
        }
    }
    let path = main.partname().as_str();
    if invalid_targets != 0 {
        issues.push(count_issue(
            MAIN_RELATIONSHIPS,
            "docx.main_document.relationship_target_invalid",
            IssueSeverity::Error,
            "The main document contains an internal relationship target that cannot be resolved.",
            path,
            "invalid_internal_targets",
            invalid_targets,
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }
    if missing_targets != 0 {
        issues.push(count_issue(
            MAIN_RELATIONSHIPS,
            "docx.main_document.relationship_target_missing",
            IssueSeverity::Error,
            "The main document contains an internal relationship target with no catalogued part.",
            path,
            "missing_internal_targets",
            missing_targets,
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }
    if content_type_mismatches != 0 {
        issues.push(count_issue(
            MAIN_RELATIONSHIPS,
            "docx.main_document.relationship_target_content_type",
            IssueSeverity::Error,
            "A main-document relationship target has an incompatible or missing content type.",
            path,
            "content_type_mismatches",
            content_type_mismatches,
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }
    Ok(CheckStatus::Complete)
}

fn relationship_target_content_type_matches(reltype: &str, content_type: &str) -> bool {
    if is_office_relationship(reltype, rt::COMMENTS, "comments") {
        return content_type == ct::WML_COMMENTS;
    }
    if is_office_relationship(reltype, rt::ENDNOTES, "endnotes") {
        return content_type == ct::WML_ENDNOTES;
    }
    if is_office_relationship(reltype, rt::FONT_TABLE, "fontTable") {
        return content_type == ct::WML_FONT_TABLE;
    }
    if is_office_relationship(reltype, rt::FOOTER, "footer") {
        return content_type == ct::WML_FOOTER;
    }
    if is_office_relationship(reltype, rt::FOOTNOTES, "footnotes") {
        return content_type == ct::WML_FOOTNOTES;
    }
    if is_office_relationship(reltype, rt::HEADER, "header") {
        return content_type == ct::WML_HEADER;
    }
    if is_office_relationship(reltype, rt::NUMBERING, "numbering") {
        return content_type == ct::WML_NUMBERING;
    }
    if is_office_relationship(reltype, rt::SETTINGS, "settings") {
        return content_type == ct::WML_SETTINGS;
    }
    if is_office_relationship(reltype, rt::STYLES, "styles") {
        return content_type == ct::WML_STYLES;
    }
    if is_office_relationship(reltype, rt::WEB_SETTINGS, "webSettings") {
        return content_type == ct::WML_WEB_SETTINGS;
    }
    if is_office_relationship(reltype, rt::IMAGE, "image") {
        return is_image_content_type(content_type);
    }
    if is_office_relationship(reltype, rt::AUDIO, "audio") {
        return is_audio_content_type(content_type);
    }
    if is_office_relationship(reltype, rt::VIDEO, "video") {
        return is_video_content_type(content_type);
    }
    if reltype == rt::MEDIA {
        return is_image_content_type(content_type)
            || is_audio_content_type(content_type)
            || is_video_content_type(content_type);
    }
    if is_office_relationship(reltype, rt::CHART, "chart") {
        return content_type == ct::DML_CHART;
    }
    if is_office_relationship(reltype, rt::DRAWING, "drawing") {
        return content_type == ct::OFC_DRAWING;
    }
    if is_office_relationship(reltype, rt::VML_DRAWING, "vmlDrawing") {
        return content_type == ct::OFC_VML_DRAWING;
    }
    if is_office_relationship(reltype, rt::THEME, "theme") {
        return content_type == ct::OFC_THEME;
    }
    if is_office_relationship(reltype, rt::THEME_OVERRIDE, "themeOverride") {
        return content_type == ct::OFC_THEME_OVERRIDE;
    }
    if is_office_relationship(reltype, rt::OLE_OBJECT, "oleObject") {
        return content_type == ct::OFC_OLE_OBJECT;
    }
    if is_office_relationship(reltype, rt::PACKAGE, "package") {
        return content_type == ct::OFC_PACKAGE;
    }
    if reltype == rt::VBA_PROJECT {
        return content_type == ct::OFC_VBA_PROJECT;
    }
    if reltype == rt::WORD_VBA_DATA {
        return content_type == ct::WML_VBA_DATA;
    }
    if is_office_relationship(reltype, rt::CUSTOM_XML, "customXml") {
        return content_type == ct::OFC_CUSTOM_XML_PROPERTIES
            || content_type == ct::XML
            || content_type.ends_with("+xml");
    }
    if is_known_active_x_relationship(reltype) {
        return content_type == "application/vnd.ms-office.activeX"
            || content_type == "application/vnd.ms-office.activeX+xml";
    }
    true
}

fn is_office_relationship(reltype: &str, transitional: &str, strict_local: &str) -> bool {
    reltype == transitional
        || reltype
            .strip_prefix(STRICT_RELATIONSHIP_PREFIX)
            .is_some_and(|local| local == strict_local)
}

fn is_known_active_x_relationship(reltype: &str) -> bool {
    matches!(
        reltype,
        ACTIVE_X_BINARY_RELATIONSHIP
            | ACTIVE_X_RELATIONSHIP
            | ACTIVE_X_DESCRIPTOR_RELATIONSHIP
            | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"
    ) || reltype
        .strip_prefix(STRICT_RELATIONSHIP_PREFIX)
        .is_some_and(|local| {
            matches!(
                local,
                "control" | "activeXControlBinary" | "activeXControl" | "activeX"
            )
        })
}

fn is_image_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::BMP | ct::GIF | ct::JPEG | ct::PNG | ct::TIFF | ct::MS_PHOTO | ct::X_EMF | ct::X_WMF
    )
}

fn is_audio_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::AUDIO_MPEG | ct::AUDIO_WAV | ct::AUDIO_WMA | ct::AUDIO_M4A
    )
}

fn is_video_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::VIDEO_MP4 | ct::VIDEO_WMV | ct::VIDEO_AVI | ct::VIDEO_MOV
    )
}

impl ValidationSource {
    fn new(source: Arc<dyn ReadAt>) -> Self {
        Self {
            source,
            first_io_error: Mutex::new(None),
        }
    }

    fn record(&self, error: io::Error) -> io::Error {
        let forwarded = io::Error::new(error.kind(), "source read failed");
        let mut first = self
            .first_io_error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if first.is_none() {
            *first = Some(error);
        }
        forwarded
    }

    fn take_error(&self) -> Option<io::Error> {
        self.first_io_error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
    }
}

impl ReadAt for ValidationSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len().map_err(|error| self.record(error))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.source
            .read_at(offset, output)
            .map_err(|error| self.record(error))
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        self.source.version().map_err(|error| self.record(error))
    }
}

fn inspect_main_document(
    main: &PartView<'_>,
    main_path: &str,
    limits: DocxValidationLimits,
    report_limits: ValidationLimits,
    tracker: Option<&ValidationSource>,
) -> std::result::Result<(MceInspection, Vec<ValidationIssue>), InspectionFailure> {
    if limits.max_main_document_bytes == 0 {
        return Err(InspectionFailure::Blocked(
            "main-document XML exceeds the configured DOCX semantic byte ceiling",
        ));
    }
    let data_result = main.data();
    if let Some(tracker) = tracker {
        if let Some(error) = tracker.take_error() {
            return Err(InspectionFailure::Ingress(OpcError::IoError(error)));
        }
    }
    let data = match data_result {
        Ok(data) => data,
        Err(OpcError::ReadLimit { .. }) => {
            return Err(InspectionFailure::Blocked(
                "main-document payload exceeded the configured OPC read ceiling",
            ));
        },
        Err(error) if is_structural_rejection(&error) => {
            return Err(payload_issue(
                "The main-document payload could not be read as a structurally valid ZIP member.",
                main_path,
                report_limits,
            ));
        },
        Err(error) if is_payload_limit(&error) => {
            return Err(InspectionFailure::Blocked(
                "main-document payload exceeded the configured DOCX semantic byte ceiling",
            ));
        },
        Err(error) => return Err(InspectionFailure::Ingress(error)),
    };
    let bytes = data.as_bytes();
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > limits.max_main_document_bytes {
        return Err(InspectionFailure::Blocked(
            "main-document XML exceeds the configured DOCX semantic byte ceiling",
        ));
    }
    let max_xml_depth = effective_xml_depth(limits);
    // The namespace reader uses a u16 nesting counter internally.  Preflight
    // with end-name checking disabled so hostile nesting is rejected before
    // that shared parser can overflow; this pass also bounds every event that
    // the MCE processor may inspect.
    preflight_xml_structure(bytes, limits.max_xml_events, max_xml_depth)?;

    let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
    capabilities.understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: bytes.len(),
        max_output_bytes: limits.max_mce_output_bytes,
        max_depth: max_xml_depth,
        max_directive_tokens: limits.max_mce_directive_tokens,
        max_choices_per_alternate: limits.max_mce_choices_per_alternate,
        ..Default::default()
    };
    let visible = match litchi_ooxml_common::mce::process_markup_compatibility(
        bytes,
        &capabilities,
        &mce_limits,
    ) {
        Ok(output) => output,
        Err(litchi_ooxml_common::mce::Error::LimitExceeded(_)) => {
            return Err(InspectionFailure::Blocked(
                "main-document markup compatibility exceeded a configured finite limit",
            ));
        },
        Err(litchi_ooxml_common::mce::Error::Allocation { resource, .. }) => {
            return Err(InspectionFailure::Allocation { resource });
        },
        Err(error) => {
            let code = if matches!(error, litchi_ooxml_common::mce::Error::MustUnderstand(_)) {
                "docx.mce.unsupported_namespace"
            } else {
                "docx.mce.invalid"
            };
            let issue = simple_issue(
                MCE,
                code,
                "The main-document markup-compatibility layer is malformed or unsupported.",
                main_path,
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )
            .map_err(InspectionFailure::Report)?;
            return Ok((
                MceInspection {
                    status: CheckStatus::blocked(
                        "main-document markup compatibility is malformed or unsupported",
                        report_limits,
                    )
                    .map_err(InspectionFailure::Report)?,
                    blocks_semantics: true,
                    issue: Some(issue),
                },
                Vec::new(),
            ));
        },
    };

    let mce_status = if visible.report == litchi_ooxml_common::mce::Report::default() {
        CheckStatus::NotApplicable
    } else {
        CheckStatus::Complete
    };
    let mce_issue = if visible.report == litchi_ooxml_common::mce::Report::default() {
        None
    } else {
        Some(
            count_issue(
                MCE,
                "docx.mce.branch_selected",
                IssueSeverity::Info,
                "Markup-compatibility content was normalized before semantic validation.",
                main_path,
                "selected_or_ignored_constructs",
                u64::try_from(
                    visible
                        .report
                        .alternate_content_count
                        .saturating_add(visible.report.ignored_elements)
                        .saturating_add(visible.report.ignored_attributes),
                )
                .unwrap_or(u64::MAX),
                CompatibilityImpact::None,
                report_limits,
            )
            .map_err(InspectionFailure::Report)?,
        )
    };

    let semantic = inspect_visible_xml(visible.xml.as_ref(), main_path, limits, report_limits)?;
    Ok((
        MceInspection {
            status: mce_status,
            blocks_semantics: false,
            issue: mce_issue,
        },
        semantic,
    ))
}

fn preflight_xml_structure(
    xml: &[u8],
    maximum_events: usize,
    maximum_depth: usize,
) -> std::result::Result<(), InspectionFailure> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(event) => {
                events = events.saturating_add(1);
                if events > maximum_events {
                    return Err(InspectionFailure::Blocked(
                        "main-document XML exceeded the configured event ceiling",
                    ));
                }
                match event {
                    Event::Start(_) => {
                        if depth >= maximum_depth {
                            return Err(InspectionFailure::Blocked(
                                "main-document XML exceeded the configured depth ceiling",
                            ));
                        }
                        depth = depth.saturating_add(1);
                    },
                    Event::Empty(_) if depth >= maximum_depth => {
                        return Err(InspectionFailure::Blocked(
                            "main-document XML exceeded the configured depth ceiling",
                        ));
                    },
                    Event::End(_) => {
                        // A malformed underflow is left for the semantic pass
                        // to classify as malformed XML.
                        depth = depth.saturating_sub(1);
                    },
                    Event::Eof => return Ok(()),
                    _ => {},
                }
            },
            // The semantic/MCE passes retain malformed-versus-unsupported
            // classification. A parser error is still bounded by the finite
            // input-byte ceiling and does not justify another issue here.
            Err(_) => return Ok(()),
        }
        buffer.clear();
    }
}

fn payload_issue(message: &str, path: &str, limits: ValidationLimits) -> InspectionFailure {
    simple_issue(
        MAIN_DOCUMENT,
        "docx.main_document.payload_invalid",
        message,
        path,
        IssueSeverity::Error,
        CompatibilityImpact::Interoperability,
        limits,
    )
    .map(|issue| InspectionFailure::Issue(Box::new(issue)))
    .unwrap_or_else(InspectionFailure::Report)
}

fn malformed_xml_failure(path: &str, limits: ValidationLimits) -> InspectionFailure {
    simple_issue(
        SEMANTICS,
        "docx.main_document.malformed_xml",
        "The visible main-document XML is not well formed or contains an unsupported XML construct.",
        path,
        IssueSeverity::Error,
        CompatibilityImpact::Interoperability,
        limits,
    )
    .map(|issue| InspectionFailure::Issue(Box::new(issue)))
    .unwrap_or_else(InspectionFailure::Report)
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn is_valid_general_reference(reference: &quick_xml::events::BytesRef<'_>) -> bool {
    if let Ok(Some(character)) = reference.resolve_char_ref() {
        return is_legal_xml_character(character);
    }
    matches!(
        reference.as_ref(),
        b"amp" | b"lt" | b"gt" | b"apos" | b"quot"
    )
}

fn is_valid_xml_reference_bytes(bytes: &[u8]) -> bool {
    let mut remaining = bytes;
    while let Some(start) = remaining.iter().position(|byte| *byte == b'&') {
        remaining = &remaining[start.saturating_add(1)..];
        let Some(end) = remaining.iter().position(|byte| *byte == b';') else {
            return false;
        };
        let Ok(reference) = std::str::from_utf8(&remaining[..end]) else {
            return false;
        };
        if reference.is_empty()
            || !is_valid_general_reference(&quick_xml::events::BytesRef::new(reference))
        {
            return false;
        }
        remaining = &remaining[end.saturating_add(1)..];
    }
    true
}

fn is_valid_xml_characters(bytes: &[u8]) -> bool {
    std::str::from_utf8(bytes).is_ok_and(|text| text.chars().all(is_legal_xml_character))
}

fn is_valid_element_attributes(element: &quick_xml::events::BytesStart<'_>) -> bool {
    element.attributes().with_checks(true).all(|attribute| {
        attribute.is_ok_and(|attribute| {
            is_valid_xml_characters(attribute.value.as_ref())
                && is_valid_xml_reference_bytes(attribute.value.as_ref())
        })
    })
}

fn has_unknown_attribute_namespace<R: io::BufRead>(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<R>,
) -> bool {
    element.attributes().with_checks(true).any(|attribute| {
        attribute.is_ok_and(|attribute| {
            let key = attribute.key.as_ref();
            // Namespace declarations are consumed by NsReader itself and are
            // not ordinary qualified attributes.
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                return false;
            }
            matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        })
    })
}

fn resolved_namespace_flags(namespace: ResolveResult<'_>) -> (bool, bool) {
    match namespace {
        ResolveResult::Unknown(_) => (true, false),
        ResolveResult::Bound(namespace) => (
            false,
            namespace.into_inner() == crate::namespace::WORDPROCESSINGML_NAMESPACE
                || namespace.into_inner() == crate::namespace::STRICT_WORDPROCESSINGML_NAMESPACE,
        ),
        ResolveResult::Unbound => (false, false),
    }
}

fn is_legal_xml_character(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn inspect_visible_xml(
    xml: &[u8],
    path: &str,
    limits: DocxValidationLimits,
    report_limits: ValidationLimits,
) -> std::result::Result<Vec<ValidationIssue>, InspectionFailure> {
    let max_xml_depth = effective_xml_depth(limits);
    // MCE output is bounded separately, but it still enters NsReader below;
    // preflight the actual visible bytes as well as the raw source so shared
    // parser nesting can never overflow before our semantic ceiling runs.
    preflight_xml_structure(xml, limits.max_xml_events, max_xml_depth)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut body_depth = None;
    let mut body_seen = false;
    let mut unsupported_body_children = 0_u64;
    let mut paragraphs = 0_u64;
    let mut tables = 0_u64;
    let mut buffer = Vec::new();

    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(|_| {
            simple_issue(
                SEMANTICS,
                "docx.main_document.malformed_xml",
                "The visible main-document XML is not well formed.",
                path,
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )
            .map(|issue| InspectionFailure::Issue(Box::new(issue)))
            .unwrap_or_else(InspectionFailure::Report)
        })?;
        events = events.saturating_add(1);
        if events > limits.max_xml_events {
            return Err(InspectionFailure::Blocked(
                "visible main-document XML exceeded the configured event ceiling",
            ));
        }
        let (unknown_namespace, is_word_namespace) = resolved_namespace_flags(namespace);
        if unknown_namespace {
            return Err(malformed_xml_failure(path, report_limits));
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                if !is_valid_element_attributes(&element)
                    || has_unknown_attribute_namespace(&element, &reader)
                {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                depth = depth.saturating_add(1);
                if depth > max_xml_depth {
                    return Err(InspectionFailure::Blocked(
                        "visible main-document XML exceeded the configured depth ceiling",
                    ));
                }
                let local = element.local_name();
                if depth == 1 {
                    if root_seen {
                        return Err(malformed_xml_failure(path, report_limits));
                    }
                    root_seen = true;
                    if !is_word_namespace || local.as_ref() != b"document" {
                        return Err(InspectionFailure::Issue(Box::new(simple_issue(
                            SEMANTICS,
                            "docx.main_document.unsupported_root",
                            "The main-document XML root is not a supported WordprocessingML document element.",
                            path,
                            IssueSeverity::Error,
                            CompatibilityImpact::Interoperability,
                            report_limits,
                        ).map_err(InspectionFailure::Report)?)));
                    }
                } else if depth == 2 && root_seen && is_word_namespace && local.as_ref() == b"body"
                {
                    if body_seen {
                        return Err(malformed_xml_failure(path, report_limits));
                    }
                    body_seen = true;
                    body_depth = Some(depth);
                } else if body_depth == Some(depth.saturating_sub(1)) {
                    classify_body_child(
                        is_word_namespace,
                        local.as_ref(),
                        &mut unsupported_body_children,
                        &mut paragraphs,
                        &mut tables,
                    );
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                if !is_valid_element_attributes(&element)
                    || has_unknown_attribute_namespace(&element, &reader)
                {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                let child_depth = depth.saturating_add(1);
                if child_depth > max_xml_depth {
                    return Err(InspectionFailure::Blocked(
                        "visible main-document XML exceeded the configured depth ceiling",
                    ));
                }
                let local = element.local_name();
                if child_depth == 1 {
                    if root_seen {
                        return Err(malformed_xml_failure(path, report_limits));
                    }
                    root_seen = true;
                    root_closed = true;
                    if !is_word_namespace || local.as_ref() != b"document" {
                        return Err(InspectionFailure::Issue(Box::new(simple_issue(
                            SEMANTICS,
                            "docx.main_document.unsupported_root",
                            "The main-document XML root is not a supported WordprocessingML document element.",
                            path,
                            IssueSeverity::Error,
                            CompatibilityImpact::Interoperability,
                            report_limits,
                        ).map_err(InspectionFailure::Report)?)));
                    }
                }
                if child_depth == 2 && root_seen && is_word_namespace && local.as_ref() == b"body" {
                    if body_seen {
                        return Err(malformed_xml_failure(path, report_limits));
                    }
                    body_seen = true;
                    // An empty body has no matching End event.  Keeping a
                    // stale body depth would make later direct children look
                    // nested even though the body has already closed.
                    body_depth = None;
                } else if body_depth == Some(child_depth.saturating_sub(1)) {
                    classify_body_child(
                        is_word_namespace,
                        local.as_ref(),
                        &mut unsupported_body_children,
                        &mut paragraphs,
                        &mut tables,
                    );
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                let local = element.local_name();
                if body_depth == Some(depth) && is_word_namespace && local.as_ref() == b"body" {
                    body_depth = None;
                }
                if depth == 1 {
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| malformed_xml_failure(path, report_limits))?;
            },
            Event::Eof => break,
            Event::Text(text) => {
                if !is_valid_xml_characters(text.as_ref())
                    || !is_valid_xml_reference_bytes(text.as_ref())
                    || (depth == 0 && !is_xml_whitespace(text.as_ref()))
                {
                    return Err(malformed_xml_failure(path, report_limits));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(malformed_xml_failure(path, report_limits));
            },
            Event::CData(data) => {
                if !is_valid_xml_characters(data.as_ref()) {
                    return Err(malformed_xml_failure(path, report_limits));
                }
            },
            Event::Comment(comment) => {
                if !is_valid_xml_characters(comment.as_ref())
                    || comment.as_ref().windows(2).any(|window| window == b"--")
                {
                    return Err(malformed_xml_failure(path, report_limits));
                }
            },
            Event::Decl(declaration) => {
                let standalone_invalid = declaration.standalone().is_some_and(|value| {
                    value.map_or(true, |value| !matches!(value.as_ref(), b"yes" | b"no"))
                });
                if declaration_seen
                    || root_seen
                    || events != 1
                    || declaration.xml_version().is_err()
                    || declaration.encoding().is_some_and(|value| value.is_err())
                    || standalone_invalid
                {
                    return Err(malformed_xml_failure(path, report_limits));
                }
                declaration_seen = true;
            },
            Event::PI(_) | Event::DocType(_) => {
                return Err(malformed_xml_failure(path, report_limits));
            },
            Event::GeneralRef(reference) => {
                if depth == 0 || !is_valid_general_reference(&reference) {
                    return Err(malformed_xml_failure(path, report_limits));
                }
            },
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(InspectionFailure::Issue(Box::new(
            simple_issue(
                SEMANTICS,
                "docx.main_document.malformed_xml",
                "The visible main-document XML is not well formed.",
                path,
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )
            .map_err(InspectionFailure::Report)?,
        )));
    }
    let mut issues = Vec::new();
    issues
        .try_reserve_exact(3)
        .map_err(|_| InspectionFailure::Allocation {
            resource: "semantic issue staging",
        })?;
    if !body_seen {
        issues.push(
            simple_issue(
                SEMANTICS,
                "docx.main_document.body_missing",
                "The supported WordprocessingML document root does not contain a body element.",
                path,
                IssueSeverity::Error,
                CompatibilityImpact::Interoperability,
                report_limits,
            )
            .map_err(InspectionFailure::Report)?,
        );
    }
    if unsupported_body_children != 0 {
        issues.push(count_issue(
            SEMANTICS,
            "docx.main_document.unsupported_body_children",
            IssueSeverity::Warning,
            "The main-document body contains direct children outside the supported semantic block set; they remain inert.",
            path,
            "unsupported_body_children",
            unsupported_body_children,
            CompatibilityImpact::Interoperability,
            report_limits,
        ).map_err(InspectionFailure::Report)?);
    }
    issues.push(count_issue(
        SEMANTICS,
        "docx.main_document.supported_blocks",
        IssueSeverity::Info,
        "Supported WordprocessingML main-document blocks were validated without retaining their content.",
        path,
        "supported_paragraphs_and_tables",
        paragraphs.saturating_add(tables),
        CompatibilityImpact::None,
        report_limits,
    ).map_err(InspectionFailure::Report)?);
    Ok(issues)
}

fn classify_body_child(
    is_word: bool,
    local: &[u8],
    unsupported: &mut u64,
    paragraphs: &mut u64,
    tables: &mut u64,
) {
    if !is_word {
        *unsupported = unsupported.saturating_add(1);
    } else {
        match local {
            b"p" => *paragraphs = paragraphs.saturating_add(1),
            b"tbl" => *tables = tables.saturating_add(1),
            b"sectPr" | b"altChunk" => {},
            _ => *unsupported = unsupported.saturating_add(1),
        }
    }
}

fn presence_report(
    package: &SourceBackedPackage,
    count_part: impl Fn(PartView<'_>) -> usize,
    capability: &str,
    code: &str,
    message: &str,
    severity: IssueSeverity,
    impact: CompatibilityImpact,
    limits: ValidationLimits,
    issues: &mut Vec<ValidationIssue>,
) -> Result<CheckStatus> {
    let mut count = 0_u64;
    count = count.saturating_add(
        package
            .rels()
            .iter()
            .filter(|relationship| relationship.is_external())
            .count() as u64,
    );
    for part in package.iter_parts() {
        count = count.saturating_add(count_part(part) as u64);
    }
    if count == 0 {
        return Ok(CheckStatus::NotApplicable);
    }
    issues.push(count_issue(
        capability,
        code,
        severity,
        message,
        "package-relationships",
        "observations",
        count,
        impact,
        limits,
    )?);
    Ok(CheckStatus::Complete)
}

fn security_report(
    package: &SourceBackedPackage,
    count_part: impl Fn(PartView<'_>) -> u64,
    package_relationship: fn(&str) -> bool,
    capability: &str,
    code: &str,
    message: &str,
    severity: IssueSeverity,
    limits: ValidationLimits,
    issues: &mut Vec<ValidationIssue>,
) -> Result<CheckStatus> {
    let mut count = 0_u64;
    for relationship in package.rels().iter() {
        count = count.saturating_add(u64::from(package_relationship(relationship.reltype())));
    }
    for part in package.iter_parts() {
        count = count.saturating_add(count_part(part));
    }
    if count == 0 {
        return Ok(CheckStatus::NotApplicable);
    }
    issues.push(count_issue(
        capability,
        code,
        severity,
        message,
        "package-security",
        "observations",
        count,
        CompatibilityImpact::Security,
        limits,
    )?);
    Ok(CheckStatus::Complete)
}

fn signature_report(
    package: &SourceBackedPackage,
    limits: ValidationLimits,
    issues: &mut Vec<ValidationIssue>,
) -> Result<CheckStatus> {
    let mut count = package
        .rels()
        .iter()
        .filter(|relationship| is_signature_relationship(relationship.reltype()))
        .count() as u64;
    for part in package.iter_parts() {
        count = count.saturating_add(u64::from(
            part.partname()
                .as_str()
                .as_bytes()
                .get(..SIGNATURE_DIRECTORY.len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case(SIGNATURE_DIRECTORY.as_bytes()))
                || is_signature_content_type(part.content_type()),
        ));
        count = count.saturating_add(
            part.rels()
                .iter()
                .filter(|relationship| is_signature_relationship(relationship.reltype()))
                .count() as u64,
        );
    }
    if count == 0 {
        return Ok(CheckStatus::NotApplicable);
    }
    issues.push(count_issue(
        SIGNATURES,
        "docx.signature.infrastructure_present",
        IssueSeverity::Info,
        "DOCX signature infrastructure is present; cryptographic signature validity was not checked.",
        "signature-infrastructure",
        "observations",
        count,
        CompatibilityImpact::None,
        limits,
    )?);
    Ok(CheckStatus::Complete)
}

fn is_macro_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
            | ct::OFC_VBA_PROJECT
            | ct::OFC_VBA_PROJECT_SIGNATURE
            | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
            | ct::WML_VBA_DATA
    )
}

fn is_embedded_content_type(value: &str) -> bool {
    matches!(value, ct::OFC_OLE_OBJECT | ct::OFC_PACKAGE)
        || value == "application/vnd.ms-office.activeX"
        || value == "application/vnd.ms-office.activeX+xml"
}

fn is_macro_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::VBA_PROJECT
            | rt::VBA_PROJECT_SIGNATURE
            | rt::VBA_PROJECT_SIGNATURE_AGILE
            | rt::WORD_VBA_DATA
    )
}

fn is_embedded_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT | rt::PACKAGE | rt::STRICT_PACKAGE
    ) || is_known_active_x_relationship(value)
}

fn is_signature_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_signature_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN
            | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
            | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
    )
}

fn blocked_ingress_report(limits: ValidationLimits) -> Result<ValidateReport> {
    let ingress = id(INGRESS, limits)?;
    let blocked = CheckStatus::blocked(
        "DOCX OPC ingress reached a configured finite resource ceiling",
        limits,
    )?;
    let stopped = CheckStatus::stopped_by(ingress.clone());
    ValidateReport::try_new(
        [
            ValidationCheck::new(ingress, blocked),
            ValidationCheck::new(id(MAIN_DOCUMENT, limits)?, stopped.clone()),
            ValidationCheck::new(id(MAIN_RELATIONSHIPS, limits)?, stopped.clone()),
            ValidationCheck::new(id(MCE, limits)?, stopped.clone()),
            ValidationCheck::new(id(SEMANTICS, limits)?, stopped.clone()),
            ValidationCheck::new(id(EXTERNAL, limits)?, stopped.clone()),
            ValidationCheck::new(id(MACROS, limits)?, stopped.clone()),
            ValidationCheck::new(id(EMBEDDED, limits)?, stopped.clone()),
            ValidationCheck::new(id(SIGNATURES, limits)?, stopped),
        ],
        [],
        limits,
    )
    .map_err(Into::into)
}

fn rejected_ingress_report(error: &OpcError, limits: ValidationLimits) -> Result<ValidateReport> {
    let ingress = id(INGRESS, limits)?;
    let issue = simple_issue(
        INGRESS,
        "docx.package.invalid",
        rejection_message(error),
        "package",
        IssueSeverity::Error,
        CompatibilityImpact::Interoperability,
        limits,
    )?;
    let blocked = CheckStatus::blocked(
        "DOCX OPC ingress structural validation prevented dependent checks",
        limits,
    )?;
    ValidateReport::try_new(
        [
            ValidationCheck::new(ingress, CheckStatus::Complete),
            ValidationCheck::new(id(MAIN_DOCUMENT, limits)?, blocked.clone()),
            ValidationCheck::new(id(MAIN_RELATIONSHIPS, limits)?, blocked.clone()),
            ValidationCheck::new(id(MCE, limits)?, blocked.clone()),
            ValidationCheck::new(id(SEMANTICS, limits)?, blocked.clone()),
            ValidationCheck::new(id(EXTERNAL, limits)?, blocked.clone()),
            ValidationCheck::new(id(MACROS, limits)?, blocked.clone()),
            ValidationCheck::new(id(EMBEDDED, limits)?, blocked.clone()),
            ValidationCheck::new(id(SIGNATURES, limits)?, blocked),
        ],
        [issue],
        limits,
    )
    .map_err(Into::into)
}

fn rejection_message(error: &OpcError) -> &'static str {
    match error {
        OpcError::ZipError(_) => "The DOCX source is not a structurally readable OPC ZIP package.",
        OpcError::InvalidContentTypesManifest(_)
        | OpcError::ContentTypeNotFound(_)
        | OpcError::InvalidContentType { .. }
        | OpcError::DuplicateContentTypeDefault(_)
        | OpcError::DuplicateContentTypeOverride { .. }
        | OpcError::InvalidContentTypeExtension(_) => {
            "The DOCX OPC content-type catalog is structurally invalid."
        },
        OpcError::InvalidRelationship(_)
        | OpcError::InvalidRelationshipsManifest(_)
        | OpcError::DuplicateRelationshipId(_)
        | OpcError::InvalidRelationshipTargetMode(_)
        | OpcError::MultipleCorePropertiesRelationships => {
            "A DOCX OPC relationship manifest is structurally invalid."
        },
        _ => "The DOCX OPC package is structurally invalid.",
    }
}

fn is_structural_rejection(error: &OpcError) -> bool {
    matches!(
        error,
        OpcError::InvalidPackUri(_)
            | OpcError::PartNotFound(_)
            | OpcError::DuplicatePartName(_)
            | OpcError::EquivalentPartNames { .. }
            | OpcError::DerivedPartNames { .. }
            | OpcError::ContentTypeNotFound(_)
            | OpcError::InvalidContentType { .. }
            | OpcError::InvalidContentTypesManifest(_)
            | OpcError::DuplicateContentTypeDefault(_)
            | OpcError::DuplicateContentTypeOverride { .. }
            | OpcError::InvalidContentTypeExtension(_)
            | OpcError::InvalidRelationship(_)
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::DuplicateRelationshipId(_)
            | OpcError::InvalidRelationshipTargetMode(_)
            | OpcError::RelationshipPartCannotBeSource(_)
            | OpcError::MultipleCorePropertiesRelationships
            | OpcError::XmlError(_)
            | OpcError::ZipError(_)
            | OpcError::QuickXmlError(_)
            | OpcError::Utf8Error(_)
            | OpcError::ParseIntError(_)
            | OpcError::AttrError(_)
    )
}

fn is_payload_limit(error: &OpcError) -> bool {
    matches!(error, OpcError::Execution(ExecutionError::ResourceLimit(_)))
}

fn id(value: &str, limits: ValidationLimits) -> Result<CheckCapabilityId> {
    CheckCapabilityId::try_new(value, limits).map_err(Into::into)
}

fn check(value: &str, status: CheckStatus, limits: ValidationLimits) -> Result<ValidationCheck> {
    Ok(ValidationCheck::new(id(value, limits)?, status))
}

fn simple_issue(
    capability: &str,
    code: &str,
    message: &str,
    path: &str,
    severity: IssueSeverity,
    impact: CompatibilityImpact,
    limits: ValidationLimits,
) -> std::result::Result<ValidationIssue, ValidationReportError> {
    ValidationIssue::try_new(
        CheckCapabilityId::try_new(capability, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            None,
            Some(path),
            None,
            None,
            None,
            limits,
        )?],
        [],
        None,
        impact,
        RepairAvailability::Unavailable,
        limits,
    )
}

fn count_issue(
    capability: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    path: &str,
    evidence_key: &str,
    count: u64,
    impact: CompatibilityImpact,
    limits: ValidationLimits,
) -> std::result::Result<ValidationIssue, ValidationReportError> {
    ValidationIssue::try_new(
        CheckCapabilityId::try_new(capability, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            None,
            Some(path),
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            evidence_key,
            EvidenceValue::Count(count),
            limits,
        )?],
        None,
        impact,
        RepairAvailability::Unavailable,
        limits,
    )
}
