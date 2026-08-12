//! Bounded, format-neutral validation report values.
//!
//! This module defines a vocabulary for reporting what a format-specific
//! validator checked. It deliberately does not define a validator trait, a
//! repair executor, or a generic notion of document validity. Callers must
//! inspect both the reported issues and [`ValidateReport::is_complete`].

use std::{fmt, sync::Arc};

use serde::{Serialize, Serializer};
use sha2::{Digest as _, Sha256};
use thiserror::Error;

/// Default finite bounds for one validation report.
pub const DEFAULT_VALIDATION_LIMITS: ValidationLimits = ValidationLimits::new(
    256,
    4_096,
    16,
    32,
    128,
    4_096,
    1_024,
    256,
    256,
    8 * 1024 * 1024,
);

/// Finite limits applied before validation-report data is retained.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ValidationLimits {
    checks: usize,
    issues: usize,
    locations_per_issue: usize,
    evidence_per_issue: usize,
    identifier_bytes: usize,
    message_bytes: usize,
    location_text_bytes: usize,
    citation_text_bytes: usize,
    blocker_text_bytes: usize,
    total_text_bytes: usize,
}

impl ValidationLimits {
    /// Creates explicit finite validation-report limits.
    #[must_use]
    #[allow(
        clippy::too_many_arguments,
        reason = "every independent wire bound is explicit"
    )]
    pub const fn new(
        max_checks: usize,
        max_issues: usize,
        max_locations_per_issue: usize,
        max_evidence_per_issue: usize,
        max_identifier_bytes: usize,
        max_message_bytes: usize,
        max_location_text_bytes: usize,
        max_citation_text_bytes: usize,
        max_blocker_text_bytes: usize,
        max_total_text_bytes: usize,
    ) -> Self {
        Self {
            checks: max_checks,
            issues: max_issues,
            locations_per_issue: max_locations_per_issue,
            evidence_per_issue: max_evidence_per_issue,
            identifier_bytes: max_identifier_bytes,
            message_bytes: max_message_bytes,
            location_text_bytes: max_location_text_bytes,
            citation_text_bytes: max_citation_text_bytes,
            blocker_text_bytes: max_blocker_text_bytes,
            total_text_bytes: max_total_text_bytes,
        }
    }

    /// Maximum number of declared check capabilities.
    #[must_use]
    pub const fn max_checks(self) -> usize {
        self.checks
    }

    /// Maximum number of issues.
    #[must_use]
    pub const fn max_issues(self) -> usize {
        self.issues
    }

    /// Maximum number of locations attached to one issue.
    #[must_use]
    pub const fn max_locations_per_issue(self) -> usize {
        self.locations_per_issue
    }

    /// Maximum number of evidence entries attached to one issue.
    #[must_use]
    pub const fn max_evidence_per_issue(self) -> usize {
        self.evidence_per_issue
    }

    /// Maximum UTF-8 byte length of one semantic identifier.
    #[must_use]
    pub const fn max_identifier_bytes(self) -> usize {
        self.identifier_bytes
    }

    /// Maximum UTF-8 byte length of one issue message.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.message_bytes
    }

    /// Maximum UTF-8 byte length of one location text field.
    #[must_use]
    pub const fn max_location_text_bytes(self) -> usize {
        self.location_text_bytes
    }

    /// Maximum UTF-8 byte length of one specification citation field.
    #[must_use]
    pub const fn max_citation_text_bytes(self) -> usize {
        self.citation_text_bytes
    }

    /// Maximum UTF-8 byte length of a check blocker description.
    #[must_use]
    pub const fn max_blocker_text_bytes(self) -> usize {
        self.blocker_text_bytes
    }

    /// Maximum aggregate retained UTF-8 bytes in one report.
    #[must_use]
    pub const fn max_total_text_bytes(self) -> usize {
        self.total_text_bytes
    }
}

impl Default for ValidationLimits {
    fn default() -> Self {
        DEFAULT_VALIDATION_LIMITS
    }
}

/// A stable identifier for a format-specific validation capability.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
#[serde(transparent)]
pub struct CheckCapabilityId(String);

impl CheckCapabilityId {
    /// Validates and retains an identifier such as `xls.formatting.records`.
    pub fn try_new(value: &str, limits: ValidationLimits) -> Result<Self, ValidationReportError> {
        Ok(Self(copy_identifier(value, limits)?))
    }

    /// Returns the identifier text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for CheckCapabilityId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Why one declared check did or did not run to completion.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(tag = "status", rename_all = "snake_case")]
#[non_exhaustive]
pub enum CheckStatus {
    /// The capability ran to completion.
    Complete,
    /// The capability does not apply to this input.
    NotApplicable,
    /// The capability could not start or finish for the stated bounded reason.
    Blocked {
        /// A diagnostic reason that must not contain source document content.
        reason: String,
    },
    /// The capability stopped because another declared check did not complete.
    StoppedBy {
        /// The prerequisite capability.
        check: CheckCapabilityId,
    },
}

impl CheckStatus {
    /// Constructs a bounded blocked status.
    pub fn blocked(reason: &str, limits: ValidationLimits) -> Result<Self, ValidationReportError> {
        Ok(Self::Blocked {
            reason: copy_text(
                reason,
                "blocked reason",
                ValidationLimitKind::BlockerTextBytes,
                limits.blocker_text_bytes,
                false,
            )?,
        })
    }

    /// Constructs a dependency-stopped status.
    #[must_use]
    pub const fn stopped_by(check: CheckCapabilityId) -> Self {
        Self::StoppedBy { check }
    }

    /// Returns whether this status represents a fully evaluated capability.
    #[must_use]
    pub const fn is_complete(&self) -> bool {
        matches!(self, Self::Complete | Self::NotApplicable)
    }
}

/// One declared validation capability and its outcome.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ValidationCheck {
    id: CheckCapabilityId,
    status: CheckStatus,
}

impl ValidationCheck {
    /// Creates one check outcome.
    #[must_use]
    pub const fn new(id: CheckCapabilityId, status: CheckStatus) -> Self {
        Self { id, status }
    }

    /// Returns the capability identifier.
    #[must_use]
    pub const fn id(&self) -> &CheckCapabilityId {
        &self.id
    }

    /// Returns the capability outcome.
    #[must_use]
    pub const fn status(&self) -> &CheckStatus {
        &self.status
    }
}

/// Severity of a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
#[serde(rename_all = "snake_case")]
#[non_exhaustive]
pub enum IssueSeverity {
    /// Informational diagnostic with no known compatibility failure.
    Info,
    /// A condition callers should review.
    Warning,
    /// A format or compatibility error.
    Error,
    /// A condition that prevents safe validation of affected content.
    Fatal,
}

/// A structured source location without retaining document content.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
pub struct IssueLocation {
    #[serde(skip_serializing_if = "Option::is_none")]
    part: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    byte_offset: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    byte_length: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    object_index: Option<u64>,
}

impl IssueLocation {
    /// Creates a location. At least one field must be present, byte length
    /// requires byte offset, and the byte range must not overflow.
    #[allow(
        clippy::too_many_arguments,
        reason = "the five orthogonal location coordinates are explicit"
    )]
    pub fn try_new(
        part: Option<&str>,
        path: Option<&str>,
        byte_offset: Option<u64>,
        byte_length: Option<u64>,
        object_index: Option<u64>,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        if part.is_none()
            && path.is_none()
            && byte_offset.is_none()
            && byte_length.is_none()
            && object_index.is_none()
        {
            return Err(ValidationReportError::EmptyLocation);
        }
        if byte_length.is_some() && byte_offset.is_none() {
            return Err(ValidationReportError::InvalidByteRange);
        }
        if let (Some(offset), Some(length)) = (byte_offset, byte_length) {
            offset
                .checked_add(length)
                .ok_or(ValidationReportError::InvalidByteRange)?;
        }
        Ok(Self {
            part: copy_optional_text(
                part,
                "location part",
                ValidationLimitKind::LocationTextBytes,
                limits.location_text_bytes,
            )?,
            path: copy_optional_text(
                path,
                "location path",
                ValidationLimitKind::LocationTextBytes,
                limits.location_text_bytes,
            )?,
            byte_offset,
            byte_length,
            object_index,
        })
    }

    /// Returns the package part, stream, or story name when available.
    #[must_use]
    pub fn part(&self) -> Option<&str> {
        self.part.as_deref()
    }

    /// Returns a format-owned structural path when available.
    #[must_use]
    pub fn path(&self) -> Option<&str> {
        self.path.as_deref()
    }

    /// Returns the byte offset when available.
    #[must_use]
    pub const fn byte_offset(&self) -> Option<u64> {
        self.byte_offset
    }

    /// Returns the byte length when available.
    #[must_use]
    pub const fn byte_length(&self) -> Option<u64> {
        self.byte_length
    }

    /// Returns a format-owned logical object index when available.
    #[must_use]
    pub const fn object_index(&self) -> Option<u64> {
        self.object_index
    }

    fn retained_text_bytes(&self) -> usize {
        option_text_len(self.part.as_deref()) + option_text_len(self.path.as_deref())
    }
}

/// A SHA-256 diagnostic digest used as content-free validation evidence.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct EvidenceDigest([u8; 32]);

impl EvidenceDigest {
    /// Computes a digest without retaining the source bytes.
    #[must_use]
    pub fn of(bytes: &[u8]) -> Self {
        Self(Sha256::digest(bytes).into())
    }

    /// Constructs a digest from exact bytes.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 32]) -> Self {
        Self(bytes)
    }

    /// Returns the exact digest bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 32] {
        &self.0
    }

    /// Writes the canonical lowercase hexadecimal representation.
    fn write_hex(&self, formatter: &mut impl fmt::Write) -> fmt::Result {
        for byte in self.0 {
            write!(formatter, "{byte:02x}")?;
        }
        Ok(())
    }
}

impl fmt::Debug for EvidenceDigest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("EvidenceDigest(")?;
        self.write_hex(formatter)?;
        formatter.write_str(")")
    }
}

impl fmt::Display for EvidenceDigest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.write_hex(formatter)
    }
}

impl Serialize for EvidenceDigest {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut text = String::new();
        text.try_reserve_exact(64)
            .map_err(serde::ser::Error::custom)?;
        self.write_hex(&mut text)
            .map_err(serde::ser::Error::custom)?;
        serializer.serialize_str(&text)
    }
}

/// A typed evidence value that cannot retain arbitrary document text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
#[serde(tag = "kind", content = "value", rename_all = "snake_case")]
#[non_exhaustive]
pub enum EvidenceValue {
    /// A counted number of structural objects.
    Count(u64),
    /// A zero-based byte offset.
    Offset(u64),
    /// A byte size.
    Size(u64),
    /// A boolean observation.
    Boolean(bool),
    /// A content-free SHA-256 digest.
    Sha256(EvidenceDigest),
}

/// One named, typed evidence observation.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
pub struct IssueEvidence {
    key: String,
    value: EvidenceValue,
}

impl IssueEvidence {
    /// Creates evidence with a bounded semantic key.
    pub fn try_new(
        key: &str,
        value: EvidenceValue,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        Ok(Self {
            key: copy_identifier(key, limits)?,
            value,
        })
    }

    /// Returns the evidence key.
    #[must_use]
    pub fn key(&self) -> &str {
        &self.key
    }

    /// Returns the content-free evidence value.
    #[must_use]
    pub const fn value(&self) -> EvidenceValue {
        self.value
    }
}

/// A bounded citation to a published format specification.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
pub struct SpecCitation {
    standard: String,
    clause: String,
}

impl SpecCitation {
    /// Creates a specification citation.
    pub fn try_new(
        standard: &str,
        clause: &str,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        Ok(Self {
            standard: copy_text(
                standard,
                "citation standard",
                ValidationLimitKind::CitationTextBytes,
                limits.citation_text_bytes,
                false,
            )?,
            clause: copy_text(
                clause,
                "citation clause",
                ValidationLimitKind::CitationTextBytes,
                limits.citation_text_bytes,
                false,
            )?,
        })
    }

    /// Returns the standard name or identifier.
    #[must_use]
    pub fn standard(&self) -> &str {
        &self.standard
    }

    /// Returns the cited clause.
    #[must_use]
    pub fn clause(&self) -> &str {
        &self.clause
    }

    fn retained_text_bytes(&self) -> usize {
        self.standard.len() + self.clause.len()
    }
}

/// Expected compatibility consequence of an issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
#[serde(rename_all = "snake_case")]
#[non_exhaustive]
pub enum CompatibilityImpact {
    /// No known consumer-visible consequence.
    None,
    /// Behavior can differ between consumers while preserving content.
    ApplicationSpecific,
    /// Other conforming consumers may reject or misinterpret the artifact.
    Interoperability,
    /// A consumer may discard or corrupt affected content.
    DataLoss,
    /// Processing the artifact may cross a security boundary.
    Security,
}

/// Whether a format owner exposes a separately identified repair operation.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize)]
#[serde(tag = "availability", rename_all = "snake_case")]
#[non_exhaustive]
pub enum RepairAvailability {
    /// No repair is offered for this issue.
    Unavailable,
    /// A format-owned repair exists. This value does not execute it.
    Available {
        /// Stable identifier of the format-owned repair operation.
        repair_id: String,
    },
}

impl RepairAvailability {
    /// Constructs availability for a bounded repair capability identifier.
    pub fn available(
        repair_id: &str,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        Ok(Self::Available {
            repair_id: copy_identifier(repair_id, limits)?,
        })
    }

    /// Returns the repair identifier when a repair is available.
    #[must_use]
    pub fn repair_id(&self) -> Option<&str> {
        match self {
            Self::Unavailable => None,
            Self::Available { repair_id } => Some(repair_id),
        }
    }
}

/// A stable SHA-256 identifier derived from all canonical issue fields.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct IssueId([u8; 32]);

impl IssueId {
    /// Returns the exact identifier bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 32] {
        &self.0
    }

    fn write_hex(&self, formatter: &mut impl fmt::Write) -> fmt::Result {
        for byte in self.0 {
            write!(formatter, "{byte:02x}")?;
        }
        Ok(())
    }
}

impl fmt::Debug for IssueId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("IssueId(")?;
        self.write_hex(formatter)?;
        formatter.write_str(")")
    }
}

impl fmt::Display for IssueId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.write_hex(formatter)
    }
}

impl Serialize for IssueId {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut text = String::new();
        text.try_reserve_exact(64)
            .map_err(serde::ser::Error::custom)?;
        self.write_hex(&mut text)
            .map_err(serde::ser::Error::custom)?;
        serializer.serialize_str(&text)
    }
}

/// One bounded validation issue.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ValidationIssue {
    id: IssueId,
    check: CheckCapabilityId,
    code: String,
    severity: IssueSeverity,
    message: String,
    locations: Vec<IssueLocation>,
    evidence: Vec<IssueEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    specification: Option<SpecCitation>,
    compatibility: CompatibilityImpact,
    repair: RepairAvailability,
}

impl ValidationIssue {
    /// Creates an issue, canonically sorting and de-duplicating its locations
    /// and evidence before computing its stable identifier.
    #[allow(
        clippy::too_many_arguments,
        reason = "the issue schema is explicit at construction"
    )]
    pub fn try_new(
        check: CheckCapabilityId,
        code: &str,
        severity: IssueSeverity,
        message: &str,
        locations: impl IntoIterator<Item = IssueLocation>,
        evidence: impl IntoIterator<Item = IssueEvidence>,
        specification: Option<SpecCitation>,
        compatibility: CompatibilityImpact,
        repair: RepairAvailability,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        validate_identifier(check.as_str(), limits)?;
        let code = copy_identifier(code, limits)?;
        let message = copy_text(
            message,
            "issue message",
            ValidationLimitKind::MessageBytes,
            limits.message_bytes,
            false,
        )?;
        let mut locations = collect_bounded(
            locations,
            limits.locations_per_issue,
            ValidationLimitKind::LocationsPerIssue,
        )?;
        locations.sort_unstable();
        if adjacent_duplicate(&locations) {
            return Err(ValidationReportError::DuplicateLocation);
        }
        let mut evidence = collect_bounded(
            evidence,
            limits.evidence_per_issue,
            ValidationLimitKind::EvidencePerIssue,
        )?;
        evidence.sort_unstable();
        if evidence.windows(2).any(|pair| pair[0].key == pair[1].key) {
            return Err(ValidationReportError::DuplicateEvidenceKey);
        }
        validate_issue_fields(
            &locations,
            &evidence,
            specification.as_ref(),
            &repair,
            limits,
        )?;
        let mut issue = Self {
            id: IssueId([0; 32]),
            check,
            code,
            severity,
            message,
            locations,
            evidence,
            specification,
            compatibility,
            repair,
        };
        issue.id = issue.compute_id();
        Ok(issue)
    }

    /// Returns the stable identifier.
    #[must_use]
    pub const fn id(&self) -> IssueId {
        self.id
    }

    /// Returns the capability that produced the issue.
    #[must_use]
    pub const fn check(&self) -> &CheckCapabilityId {
        &self.check
    }

    /// Returns the format-owned issue code.
    #[must_use]
    pub fn code(&self) -> &str {
        &self.code
    }

    /// Returns the severity.
    #[must_use]
    pub const fn severity(&self) -> IssueSeverity {
        self.severity
    }

    /// Returns the bounded human-facing message.
    #[must_use]
    pub fn message(&self) -> &str {
        &self.message
    }

    /// Returns canonically ordered locations.
    #[must_use]
    pub fn locations(&self) -> &[IssueLocation] {
        &self.locations
    }

    /// Returns canonically ordered evidence.
    #[must_use]
    pub fn evidence(&self) -> &[IssueEvidence] {
        &self.evidence
    }

    /// Returns the optional specification citation.
    #[must_use]
    pub const fn specification(&self) -> Option<&SpecCitation> {
        self.specification.as_ref()
    }

    /// Returns the compatibility impact.
    #[must_use]
    pub const fn compatibility(&self) -> CompatibilityImpact {
        self.compatibility
    }

    /// Returns repair availability without executing a repair.
    #[must_use]
    pub const fn repair(&self) -> &RepairAvailability {
        &self.repair
    }

    fn retained_text_bytes(&self) -> Result<usize, ValidationReportError> {
        let mut total = self.check.as_str().len();
        add_text_bytes(&mut total, self.code.len())?;
        add_text_bytes(&mut total, self.message.len())?;
        for location in &self.locations {
            add_text_bytes(&mut total, location.retained_text_bytes())?;
        }
        for evidence in &self.evidence {
            add_text_bytes(&mut total, evidence.key.len())?;
        }
        if let Some(citation) = &self.specification {
            add_text_bytes(&mut total, citation.retained_text_bytes())?;
        }
        if let Some(repair_id) = self.repair.repair_id() {
            add_text_bytes(&mut total, repair_id.len())?;
        }
        Ok(total)
    }

    fn compute_id(&self) -> IssueId {
        let mut digest = Sha256::new();
        digest.update(b"litchi.validation.issue.v1\0");
        hash_text(&mut digest, self.check.as_str());
        hash_text(&mut digest, &self.code);
        hash_u8(&mut digest, severity_tag(self.severity));
        hash_text(&mut digest, &self.message);
        hash_u64(&mut digest, self.locations.len() as u64);
        for location in &self.locations {
            hash_optional_text(&mut digest, location.part.as_deref());
            hash_optional_text(&mut digest, location.path.as_deref());
            hash_optional_u64(&mut digest, location.byte_offset);
            hash_optional_u64(&mut digest, location.byte_length);
            hash_optional_u64(&mut digest, location.object_index);
        }
        hash_u64(&mut digest, self.evidence.len() as u64);
        for evidence in &self.evidence {
            hash_text(&mut digest, &evidence.key);
            match evidence.value {
                EvidenceValue::Count(value) => {
                    hash_u8(&mut digest, 0);
                    hash_u64(&mut digest, value);
                },
                EvidenceValue::Offset(value) => {
                    hash_u8(&mut digest, 1);
                    hash_u64(&mut digest, value);
                },
                EvidenceValue::Size(value) => {
                    hash_u8(&mut digest, 2);
                    hash_u64(&mut digest, value);
                },
                EvidenceValue::Boolean(value) => {
                    hash_u8(&mut digest, 3);
                    hash_u8(&mut digest, u8::from(value));
                },
                EvidenceValue::Sha256(value) => {
                    hash_u8(&mut digest, 4);
                    digest.update(value.as_bytes());
                },
            }
        }
        match &self.specification {
            None => hash_u8(&mut digest, 0),
            Some(citation) => {
                hash_u8(&mut digest, 1);
                hash_text(&mut digest, &citation.standard);
                hash_text(&mut digest, &citation.clause);
            },
        }
        hash_u8(&mut digest, compatibility_tag(self.compatibility));
        match &self.repair {
            RepairAvailability::Unavailable => hash_u8(&mut digest, 0),
            RepairAvailability::Available { repair_id } => {
                hash_u8(&mut digest, 1);
                hash_text(&mut digest, repair_id);
            },
        }
        IssueId(digest.finalize().into())
    }
}

#[derive(Debug, PartialEq, Eq, Serialize)]
struct ValidateReportInner {
    checks: Vec<ValidationCheck>,
    issues: Vec<ValidationIssue>,
}

/// A bounded, canonically ordered validation report.
///
/// Cloning a report is O(1) and shares immutable storage. A report with no
/// errors is not necessarily complete; use [`Self::is_complete`] separately.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidateReport(Arc<ValidateReportInner>);

impl ValidateReport {
    /// Builds a report under explicit finite bounds.
    pub fn try_new(
        checks: impl IntoIterator<Item = ValidationCheck>,
        issues: impl IntoIterator<Item = ValidationIssue>,
        limits: ValidationLimits,
    ) -> Result<Self, ValidationReportError> {
        let mut checks = collect_bounded(checks, limits.checks, ValidationLimitKind::Checks)?;
        checks.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        if checks.windows(2).any(|pair| pair[0].id == pair[1].id) {
            return Err(ValidationReportError::DuplicateCheck);
        }
        validate_check_statuses(&checks, limits)?;

        let mut issues = collect_bounded(issues, limits.issues, ValidationLimitKind::Issues)?;
        issues.sort_unstable_by_key(ValidationIssue::id);
        if issues.windows(2).any(|pair| pair[0].id == pair[1].id) {
            return Err(ValidationReportError::DuplicateIssue);
        }

        let mut total_text_bytes = 0_usize;
        for check in &checks {
            validate_identifier(check.id.as_str(), limits)?;
            add_text_bytes(&mut total_text_bytes, check.id.as_str().len())?;
            match &check.status {
                CheckStatus::Blocked { reason } => {
                    validate_text(
                        reason,
                        "blocked reason",
                        ValidationLimitKind::BlockerTextBytes,
                        limits.blocker_text_bytes,
                        false,
                    )?;
                    add_text_bytes(&mut total_text_bytes, reason.len())?;
                },
                CheckStatus::StoppedBy { check } => {
                    validate_identifier(check.as_str(), limits)?;
                    add_text_bytes(&mut total_text_bytes, check.as_str().len())?;
                },
                CheckStatus::Complete | CheckStatus::NotApplicable => {},
            }
        }
        for issue in &issues {
            if checks
                .binary_search_by(|check| check.id.cmp(&issue.check))
                .is_err()
            {
                return Err(ValidationReportError::UnknownIssueCheck);
            }
            validate_issue(issue, limits)?;
            add_text_bytes(&mut total_text_bytes, issue.retained_text_bytes()?)?;
        }
        if total_text_bytes > limits.total_text_bytes {
            return Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::TotalTextBytes,
                observed: total_text_bytes,
                limit: limits.total_text_bytes,
            });
        }

        Ok(Self(Arc::new(ValidateReportInner { checks, issues })))
    }

    /// Returns canonically ordered check outcomes.
    #[must_use]
    pub fn checks(&self) -> &[ValidationCheck] {
        &self.0.checks
    }

    /// Returns issues ordered by stable issue identifier.
    #[must_use]
    pub fn issues(&self) -> &[ValidationIssue] {
        &self.0.issues
    }

    /// Returns whether any issue has error or fatal severity.
    #[must_use]
    pub fn has_errors(&self) -> bool {
        self.issues()
            .iter()
            .any(|issue| issue.severity >= IssueSeverity::Error)
    }

    /// Returns whether any issue has fatal severity.
    #[must_use]
    pub fn has_fatal(&self) -> bool {
        self.issues()
            .iter()
            .any(|issue| issue.severity == IssueSeverity::Fatal)
    }

    /// Returns whether every declared capability completed or was not
    /// applicable. This intentionally says nothing about issue severity.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        self.checks().iter().all(|check| check.status.is_complete())
    }
}

impl Serialize for ValidateReport {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        self.0.serialize(serializer)
    }
}

/// A finite report bound that rejected input.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ValidationLimitKind {
    /// Number of declared checks.
    Checks,
    /// Number of issues.
    Issues,
    /// Locations attached to one issue.
    LocationsPerIssue,
    /// Evidence entries attached to one issue.
    EvidencePerIssue,
    /// UTF-8 bytes in one semantic identifier.
    IdentifierBytes,
    /// UTF-8 bytes in one issue message.
    MessageBytes,
    /// UTF-8 bytes in one location text field.
    LocationTextBytes,
    /// UTF-8 bytes in one specification citation field.
    CitationTextBytes,
    /// UTF-8 bytes in one blocked-check reason.
    BlockerTextBytes,
    /// Aggregate retained UTF-8 bytes in the report.
    TotalTextBytes,
}

/// Failure to construct a bounded validation report value.
#[derive(Debug, Error, PartialEq, Eq)]
#[non_exhaustive]
pub enum ValidationReportError {
    /// One explicit finite bound was exceeded.
    #[error("{kind:?} validation-report limit exceeded: observed {observed}, limit {limit}")]
    Limit {
        /// Bound that failed.
        kind: ValidationLimitKind,
        /// Requested or observed amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// Retaining a bounded report-owned allocation failed.
    #[error("allocation failed while retaining validation report data")]
    Allocation,
    /// A required text field was empty or contained a control code.
    #[error("invalid {field}")]
    InvalidText {
        /// Field rejected before it was retained.
        field: &'static str,
    },
    /// A location had no coordinate.
    #[error("validation issue location must contain at least one coordinate")]
    EmptyLocation,
    /// A byte range omitted its offset or overflowed `u64`.
    #[error("invalid validation issue byte range")]
    InvalidByteRange,
    /// Two declared checks used the same capability identifier.
    #[error("duplicate validation check capability")]
    DuplicateCheck,
    /// Two issues had the same stable identifier.
    #[error("duplicate validation issue")]
    DuplicateIssue,
    /// An issue repeated an identical structured location.
    #[error("duplicate validation issue location")]
    DuplicateLocation,
    /// An issue repeated an evidence key.
    #[error("duplicate validation issue evidence key")]
    DuplicateEvidenceKey,
    /// An issue named a check that was not declared by the report.
    #[error("validation issue refers to an undeclared check capability")]
    UnknownIssueCheck,
    /// A stopped check named a prerequisite that was not declared.
    #[error("stopped validation check refers to an undeclared prerequisite")]
    UnknownStoppedByCheck,
    /// A stopped check named itself as its prerequisite.
    #[error("validation check cannot be stopped by itself")]
    SelfStoppedCheck,
    /// A stopped check named a prerequisite that completed or did not apply.
    #[error("stopped validation check requires an incomplete prerequisite")]
    CompletedStoppedByCheck,
    /// Stopped-check dependencies contained a cycle instead of ending at a
    /// blocked capability.
    #[error("stopped validation check dependency cycle")]
    StoppedByCycle,
    /// Aggregate retained text length overflowed `usize`.
    #[error("validation report text length overflow")]
    TextLengthOverflow,
}

fn copy_identifier(value: &str, limits: ValidationLimits) -> Result<String, ValidationReportError> {
    validate_identifier(value, limits)?;
    copy_validated(value)
}

fn validate_identifier(value: &str, limits: ValidationLimits) -> Result<(), ValidationReportError> {
    validate_text(
        value,
        "identifier",
        ValidationLimitKind::IdentifierBytes,
        limits.identifier_bytes,
        false,
    )?;
    if value
        .bytes()
        .any(|byte| !(byte.is_ascii_lowercase() || byte.is_ascii_digit() || b"._-".contains(&byte)))
    {
        return Err(ValidationReportError::InvalidText {
            field: "identifier",
        });
    }
    Ok(())
}

fn copy_optional_text(
    value: Option<&str>,
    field: &'static str,
    kind: ValidationLimitKind,
    limit: usize,
) -> Result<Option<String>, ValidationReportError> {
    value
        .map(|value| copy_text(value, field, kind, limit, false))
        .transpose()
}

fn copy_text(
    value: &str,
    field: &'static str,
    kind: ValidationLimitKind,
    limit: usize,
    permit_empty: bool,
) -> Result<String, ValidationReportError> {
    validate_text(value, field, kind, limit, permit_empty)?;
    copy_validated(value)
}

fn validate_text(
    value: &str,
    field: &'static str,
    kind: ValidationLimitKind,
    limit: usize,
    permit_empty: bool,
) -> Result<(), ValidationReportError> {
    if (!permit_empty && value.is_empty()) || value.chars().any(char::is_control) {
        return Err(ValidationReportError::InvalidText { field });
    }
    if value.len() > limit {
        return Err(ValidationReportError::Limit {
            kind,
            observed: value.len(),
            limit,
        });
    }
    Ok(())
}

fn copy_validated(value: &str) -> Result<String, ValidationReportError> {
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_| ValidationReportError::Allocation)?;
    retained.push_str(value);
    Ok(retained)
}

fn collect_bounded<T>(
    values: impl IntoIterator<Item = T>,
    limit: usize,
    kind: ValidationLimitKind,
) -> Result<Vec<T>, ValidationReportError> {
    let iterator = values.into_iter();
    let (lower, _) = iterator.size_hint();
    if lower > limit {
        return Err(ValidationReportError::Limit {
            kind,
            observed: lower,
            limit,
        });
    }
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(lower.min(limit))
        .map_err(|_| ValidationReportError::Allocation)?;
    for value in iterator {
        if retained.len() == limit {
            return Err(ValidationReportError::Limit {
                kind,
                observed: limit.saturating_add(1),
                limit,
            });
        }
        if retained.len() == retained.capacity() {
            retained
                .try_reserve(1)
                .map_err(|_| ValidationReportError::Allocation)?;
        }
        retained.push(value);
    }
    Ok(retained)
}

fn adjacent_duplicate<T: PartialEq>(values: &[T]) -> bool {
    values.windows(2).any(|pair| pair[0] == pair[1])
}

fn validate_check_statuses(
    checks: &[ValidationCheck],
    limits: ValidationLimits,
) -> Result<(), ValidationReportError> {
    for (index, check) in checks.iter().enumerate() {
        if let CheckStatus::StoppedBy {
            check: prerequisite,
        } = &check.status
        {
            validate_identifier(prerequisite.as_str(), limits)?;
            if check.id == *prerequisite {
                return Err(ValidationReportError::SelfStoppedCheck);
            }
            let prerequisite_index = checks
                .binary_search_by(|candidate| candidate.id.cmp(prerequisite))
                .map_err(|_| ValidationReportError::UnknownStoppedByCheck)?;
            if checks[prerequisite_index].status.is_complete() {
                return Err(ValidationReportError::CompletedStoppedByCheck);
            }
            validate_stopped_chain(checks, index)?;
        }
    }
    Ok(())
}

fn validate_stopped_chain(
    checks: &[ValidationCheck],
    origin: usize,
) -> Result<(), ValidationReportError> {
    let mut current = origin;
    // A chain longer than the number of distinct checks necessarily contains
    // a cycle. No allocation is needed to prove that finite bound.
    for _ in 0..checks.len() {
        match &checks[current].status {
            CheckStatus::Blocked { .. } => return Ok(()),
            CheckStatus::StoppedBy { check } => {
                current = checks
                    .binary_search_by(|candidate| candidate.id.cmp(check))
                    .map_err(|_| ValidationReportError::UnknownStoppedByCheck)?;
            },
            CheckStatus::Complete | CheckStatus::NotApplicable => {
                return Err(ValidationReportError::CompletedStoppedByCheck);
            },
        }
    }
    Err(ValidationReportError::StoppedByCycle)
}

fn validate_issue_fields(
    locations: &[IssueLocation],
    evidence: &[IssueEvidence],
    specification: Option<&SpecCitation>,
    repair: &RepairAvailability,
    limits: ValidationLimits,
) -> Result<(), ValidationReportError> {
    for location in locations {
        for text in [location.part.as_deref(), location.path.as_deref()]
            .into_iter()
            .flatten()
        {
            validate_text(
                text,
                "location text",
                ValidationLimitKind::LocationTextBytes,
                limits.location_text_bytes,
                false,
            )?;
        }
    }
    for observation in evidence {
        validate_identifier(&observation.key, limits)?;
    }
    if let Some(citation) = specification {
        for (field, text) in [
            ("citation standard", citation.standard.as_str()),
            ("citation clause", citation.clause.as_str()),
        ] {
            validate_text(
                text,
                field,
                ValidationLimitKind::CitationTextBytes,
                limits.citation_text_bytes,
                false,
            )?;
        }
    }
    if let Some(repair_id) = repair.repair_id() {
        validate_identifier(repair_id, limits)?;
    }
    Ok(())
}

fn validate_issue(
    issue: &ValidationIssue,
    limits: ValidationLimits,
) -> Result<(), ValidationReportError> {
    validate_identifier(issue.check.as_str(), limits)?;
    validate_identifier(&issue.code, limits)?;
    validate_text(
        &issue.message,
        "issue message",
        ValidationLimitKind::MessageBytes,
        limits.message_bytes,
        false,
    )?;
    if issue.locations.len() > limits.locations_per_issue {
        return Err(ValidationReportError::Limit {
            kind: ValidationLimitKind::LocationsPerIssue,
            observed: issue.locations.len(),
            limit: limits.locations_per_issue,
        });
    }
    if issue.evidence.len() > limits.evidence_per_issue {
        return Err(ValidationReportError::Limit {
            kind: ValidationLimitKind::EvidencePerIssue,
            observed: issue.evidence.len(),
            limit: limits.evidence_per_issue,
        });
    }
    validate_issue_fields(
        &issue.locations,
        &issue.evidence,
        issue.specification.as_ref(),
        &issue.repair,
        limits,
    )
}

fn option_text_len(value: Option<&str>) -> usize {
    value.map_or(0, str::len)
}

fn add_text_bytes(total: &mut usize, amount: usize) -> Result<(), ValidationReportError> {
    *total = total
        .checked_add(amount)
        .ok_or(ValidationReportError::TextLengthOverflow)?;
    Ok(())
}

fn hash_u8(digest: &mut Sha256, value: u8) {
    digest.update([value]);
}

fn hash_u64(digest: &mut Sha256, value: u64) {
    digest.update(value.to_be_bytes());
}

fn hash_text(digest: &mut Sha256, value: &str) {
    hash_u64(digest, value.len() as u64);
    digest.update(value.as_bytes());
}

fn hash_optional_text(digest: &mut Sha256, value: Option<&str>) {
    match value {
        None => hash_u8(digest, 0),
        Some(value) => {
            hash_u8(digest, 1);
            hash_text(digest, value);
        },
    }
}

fn hash_optional_u64(digest: &mut Sha256, value: Option<u64>) {
    match value {
        None => hash_u8(digest, 0),
        Some(value) => {
            hash_u8(digest, 1);
            hash_u64(digest, value);
        },
    }
}

const fn severity_tag(severity: IssueSeverity) -> u8 {
    match severity {
        IssueSeverity::Info => 0,
        IssueSeverity::Warning => 1,
        IssueSeverity::Error => 2,
        IssueSeverity::Fatal => 3,
    }
}

const fn compatibility_tag(impact: CompatibilityImpact) -> u8 {
    match impact {
        CompatibilityImpact::None => 0,
        CompatibilityImpact::ApplicationSpecific => 1,
        CompatibilityImpact::Interoperability => 2,
        CompatibilityImpact::DataLoss => 3,
        CompatibilityImpact::Security => 4,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const fn limits() -> ValidationLimits {
        ValidationLimits::new(2, 2, 2, 2, 8, 8, 8, 8, 8, 128)
    }

    fn check_id(value: &str) -> CheckCapabilityId {
        CheckCapabilityId::try_new(value, limits()).unwrap()
    }

    fn issue(code: &str, severity: IssueSeverity) -> ValidationIssue {
        ValidationIssue::try_new(
            check_id("check.a"),
            code,
            severity,
            "message",
            [],
            [],
            None,
            CompatibilityImpact::None,
            RepairAvailability::Unavailable,
            limits(),
        )
        .unwrap()
    }

    #[test]
    fn exact_text_boundaries_are_accepted_and_above_is_rejected() {
        let bounds = ValidationLimits::new(2, 2, 2, 2, 4, 4, 4, 4, 4, 64);
        assert!(CheckCapabilityId::try_new("abc", bounds).is_ok());
        assert!(CheckCapabilityId::try_new("abcd", bounds).is_ok());
        assert!(matches!(
            CheckCapabilityId::try_new("abcde", bounds),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::IdentifierBytes,
                observed: 5,
                limit: 4
            })
        ));

        let id = CheckCapabilityId::try_new("a", bounds).unwrap();
        assert!(
            ValidationIssue::try_new(
                id.clone(),
                "c",
                IssueSeverity::Info,
                "123",
                [],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(
            ValidationIssue::try_new(
                id,
                "c",
                IssueSeverity::Info,
                "1234",
                [],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(matches!(
            ValidationIssue::try_new(
                CheckCapabilityId::try_new("a", bounds).unwrap(),
                "c",
                IssueSeverity::Info,
                "12345",
                [],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            ),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::MessageBytes,
                observed: 5,
                limit: 4
            })
        ));

        for text in ["123", "1234"] {
            assert!(IssueLocation::try_new(Some(text), None, None, None, None, bounds).is_ok());
            assert!(SpecCitation::try_new(text, "1", bounds).is_ok());
            assert!(CheckStatus::blocked(text, bounds).is_ok());
        }
        assert!(matches!(
            IssueLocation::try_new(Some("12345"), None, None, None, None, bounds),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::LocationTextBytes,
                observed: 5,
                limit: 4
            })
        ));
        assert!(matches!(
            SpecCitation::try_new("12345", "1", bounds),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::CitationTextBytes,
                observed: 5,
                limit: 4
            })
        ));
        assert!(matches!(
            CheckStatus::blocked("12345", bounds),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::BlockerTextBytes,
                observed: 5,
                limit: 4
            })
        ));
        assert!(matches!(
            CheckStatus::blocked("", bounds),
            Err(ValidationReportError::InvalidText {
                field: "blocked reason"
            })
        ));
    }

    #[test]
    fn exact_count_boundaries_are_accepted_and_above_is_rejected() {
        let bounds = limits();
        let first = ValidationCheck::new(check_id("check.a"), CheckStatus::Complete);
        let second = ValidationCheck::new(check_id("check.b"), CheckStatus::Complete);
        assert!(ValidateReport::try_new([first.clone()], [], bounds).is_ok());
        assert!(ValidateReport::try_new([first.clone(), second.clone()], [], bounds).is_ok());
        let third = ValidationCheck::new(check_id("check.c"), CheckStatus::Complete);
        assert!(matches!(
            ValidateReport::try_new([first.clone(), second.clone(), third], [], bounds),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::Checks,
                observed: 3,
                limit: 2
            })
        ));

        let issue_a = issue("a", IssueSeverity::Info);
        let issue_b = issue("b", IssueSeverity::Warning);
        assert!(ValidateReport::try_new([first.clone()], [issue_a.clone()], bounds).is_ok());
        assert!(
            ValidateReport::try_new([first.clone()], [issue_a.clone(), issue_b.clone()], bounds)
                .is_ok()
        );
        assert!(matches!(
            ValidateReport::try_new(
                [first],
                [issue_a, issue_b, issue("c", IssueSeverity::Error)],
                bounds,
            ),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::Issues,
                observed: 3,
                limit: 2
            })
        ));

        let location_a = IssueLocation::try_new(Some("a"), None, None, None, None, bounds).unwrap();
        let location_b = IssueLocation::try_new(Some("b"), None, None, None, None, bounds).unwrap();
        let location_c = IssueLocation::try_new(Some("c"), None, None, None, None, bounds).unwrap();
        assert!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [location_a.clone()],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [location_a.clone(), location_b.clone()],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(matches!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [location_a, location_b, location_c],
                [],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            ),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::LocationsPerIssue,
                observed: 3,
                limit: 2
            })
        ));

        let evidence_a = IssueEvidence::try_new("a", EvidenceValue::Count(1), bounds).unwrap();
        let evidence_b = IssueEvidence::try_new("b", EvidenceValue::Size(2), bounds).unwrap();
        let evidence_c = IssueEvidence::try_new("c", EvidenceValue::Boolean(true), bounds).unwrap();
        assert!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [],
                [evidence_a.clone()],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [],
                [evidence_a.clone(), evidence_b.clone()],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .is_ok()
        );
        assert!(matches!(
            ValidationIssue::try_new(
                check_id("check.a"),
                "c",
                IssueSeverity::Info,
                "m",
                [],
                [evidence_a, evidence_b, evidence_c],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            ),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::EvidencePerIssue,
                observed: 3,
                limit: 2
            })
        ));
    }

    #[test]
    fn aggregate_text_boundary_is_exact() {
        let base = ValidationLimits::new(1, 1, 0, 0, 8, 8, 8, 8, 8, 6);
        let check = CheckCapabilityId::try_new("a", base).unwrap();
        let issue = ValidationIssue::try_new(
            check.clone(),
            "b",
            IssueSeverity::Info,
            "xyz",
            [],
            [],
            None,
            CompatibilityImpact::None,
            RepairAvailability::Unavailable,
            base,
        )
        .unwrap();
        // check declaration (1) + issue check/code/message (1 + 1 + 3) = 6.
        assert!(
            ValidateReport::try_new(
                [ValidationCheck::new(check.clone(), CheckStatus::Complete)],
                [issue.clone()],
                base,
            )
            .is_ok()
        );
        let below = ValidationLimits::new(1, 1, 0, 0, 8, 8, 8, 8, 8, 7);
        assert!(
            ValidateReport::try_new(
                [ValidationCheck::new(check.clone(), CheckStatus::Complete)],
                [issue.clone()],
                below,
            )
            .is_ok()
        );
        let above = ValidationLimits::new(1, 1, 0, 0, 8, 8, 8, 8, 8, 5);
        assert!(matches!(
            ValidateReport::try_new(
                [ValidationCheck::new(check, CheckStatus::Complete)],
                [issue],
                above,
            ),
            Err(ValidationReportError::Limit {
                kind: ValidationLimitKind::TotalTextBytes,
                observed: 6,
                limit: 5
            })
        ));
    }

    #[test]
    fn canonical_ordering_duplicate_rejection_and_stable_id() {
        let bounds = limits();
        let location_a = IssueLocation::try_new(Some("a"), None, None, None, None, bounds).unwrap();
        let location_b = IssueLocation::try_new(Some("b"), None, None, None, None, bounds).unwrap();
        let evidence_a = IssueEvidence::try_new("a", EvidenceValue::Count(1), bounds).unwrap();
        let evidence_b = IssueEvidence::try_new("b", EvidenceValue::Size(2), bounds).unwrap();
        let build = |locations, evidence| {
            ValidationIssue::try_new(
                check_id("check.a"),
                "code",
                IssueSeverity::Error,
                "message",
                locations,
                evidence,
                Some(SpecCitation::try_new("iso", "1.2", bounds).unwrap()),
                CompatibilityImpact::Interoperability,
                RepairAvailability::available("repair.a", bounds).unwrap(),
                bounds,
            )
            .unwrap()
        };
        let first = build(
            [location_b.clone(), location_a.clone()],
            [evidence_b.clone(), evidence_a.clone()],
        );
        let second = build(
            [location_a.clone(), location_b.clone()],
            [evidence_a.clone(), evidence_b.clone()],
        );
        assert_eq!(first.id(), second.id());
        assert_eq!(first.locations(), [location_a.clone(), location_b.clone()]);
        assert_eq!(first.evidence(), [evidence_a.clone(), evidence_b.clone()]);
        assert_eq!(
            first.id().to_string(),
            "82792f094727416f0b54b7b8362bf4faf5d9404335cececb09be317df76ec1c6"
        );

        let duplicate_location = ValidationIssue::try_new(
            check_id("check.a"),
            "code",
            IssueSeverity::Error,
            "message",
            [location_b.clone(), location_b],
            [evidence_a.clone()],
            None,
            CompatibilityImpact::None,
            RepairAvailability::Unavailable,
            bounds,
        );
        assert_eq!(
            duplicate_location,
            Err(ValidationReportError::DuplicateLocation)
        );
        let duplicate_evidence = ValidationIssue::try_new(
            check_id("check.a"),
            "code",
            IssueSeverity::Error,
            "message",
            [],
            [
                evidence_a,
                IssueEvidence::try_new("a", EvidenceValue::Count(2), bounds).unwrap(),
            ],
            None,
            CompatibilityImpact::None,
            RepairAvailability::Unavailable,
            bounds,
        );
        assert_eq!(
            duplicate_evidence,
            Err(ValidationReportError::DuplicateEvidenceKey)
        );

        let check = ValidationCheck::new(check_id("check.a"), CheckStatus::Complete);
        assert_eq!(
            ValidateReport::try_new([check.clone(), check], [], bounds),
            Err(ValidationReportError::DuplicateCheck)
        );
        let duplicate = issue("a", IssueSeverity::Info);
        assert_eq!(
            ValidateReport::try_new(
                [ValidationCheck::new(
                    check_id("check.a"),
                    CheckStatus::Complete
                )],
                [duplicate.clone(), duplicate],
                bounds,
            ),
            Err(ValidationReportError::DuplicateIssue)
        );
    }

    #[test]
    fn completeness_is_separate_from_severity_and_clone_shares_storage() {
        let bounds = limits();
        let blocked = ValidateReport::try_new(
            [ValidationCheck::new(
                check_id("check.a"),
                CheckStatus::blocked("budget", bounds).unwrap(),
            )],
            [],
            bounds,
        )
        .unwrap();
        assert!(!blocked.is_complete());
        assert!(!blocked.has_errors());
        assert!(!blocked.has_fatal());

        let fatal = ValidateReport::try_new(
            [ValidationCheck::new(
                check_id("check.a"),
                CheckStatus::Complete,
            )],
            [issue("fatal", IssueSeverity::Fatal)],
            bounds,
        )
        .unwrap();
        assert!(fatal.is_complete());
        assert!(fatal.has_errors());
        assert!(fatal.has_fatal());
        let clone = fatal.clone();
        assert!(Arc::ptr_eq(&fatal.0, &clone.0));
        assert_eq!(fatal, clone);
    }

    #[test]
    fn stopped_by_requires_a_distinct_declared_capability() {
        let bounds = limits();
        let prerequisite = check_id("check.a");
        let dependent = check_id("check.b");
        let report = ValidateReport::try_new(
            [
                ValidationCheck::new(
                    prerequisite.clone(),
                    CheckStatus::blocked("budget", bounds).unwrap(),
                ),
                ValidationCheck::new(
                    dependent.clone(),
                    CheckStatus::stopped_by(prerequisite.clone()),
                ),
            ],
            [],
            bounds,
        )
        .unwrap();
        assert!(!report.is_complete());
        assert!(matches!(
            ValidateReport::try_new(
                [ValidationCheck::new(
                    dependent.clone(),
                    CheckStatus::stopped_by(prerequisite.clone())
                )],
                [],
                bounds,
            ),
            Err(ValidationReportError::UnknownStoppedByCheck)
        ));
        assert!(matches!(
            ValidateReport::try_new(
                [ValidationCheck::new(
                    dependent.clone(),
                    CheckStatus::stopped_by(dependent.clone())
                )],
                [],
                bounds,
            ),
            Err(ValidationReportError::SelfStoppedCheck)
        ));
        assert!(matches!(
            ValidateReport::try_new(
                [
                    ValidationCheck::new(prerequisite.clone(), CheckStatus::Complete),
                    ValidationCheck::new(
                        dependent.clone(),
                        CheckStatus::stopped_by(prerequisite.clone())
                    ),
                ],
                [],
                bounds,
            ),
            Err(ValidationReportError::CompletedStoppedByCheck)
        ));
        assert!(matches!(
            ValidateReport::try_new(
                [
                    ValidationCheck::new(
                        prerequisite.clone(),
                        CheckStatus::stopped_by(dependent.clone())
                    ),
                    ValidationCheck::new(dependent, CheckStatus::stopped_by(prerequisite)),
                ],
                [],
                bounds,
            ),
            Err(ValidationReportError::StoppedByCycle)
        ));
    }

    #[test]
    fn serialized_json_is_deterministic_and_content_free_evidence_is_typed() {
        let bounds = limits();
        let report = ValidateReport::try_new(
            [ValidationCheck::new(
                check_id("check.a"),
                CheckStatus::Complete,
            )],
            [ValidationIssue::try_new(
                check_id("check.a"),
                "code",
                IssueSeverity::Warning,
                "message",
                [],
                [IssueEvidence::try_new(
                    "digest",
                    EvidenceValue::Sha256(EvidenceDigest::of(b"secret document content")),
                    bounds,
                )
                .unwrap()],
                None,
                CompatibilityImpact::None,
                RepairAvailability::Unavailable,
                bounds,
            )
            .unwrap()],
            bounds,
        )
        .unwrap();
        let first = serde_json::to_string(&report).unwrap();
        let second = serde_json::to_string(&report).unwrap();
        assert_eq!(first, second);
        assert!(!first.contains("secret document content"));
        assert!(first.contains("sha256"));
    }

    #[test]
    fn report_and_json_order_are_independent_of_input_order() {
        let bounds = limits();
        let check_a = ValidationCheck::new(check_id("check.a"), CheckStatus::Complete);
        let check_b = ValidationCheck::new(check_id("check.b"), CheckStatus::Complete);
        let issue_a = issue("a", IssueSeverity::Info);
        let issue_b = ValidationIssue::try_new(
            check_id("check.b"),
            "b",
            IssueSeverity::Warning,
            "message",
            [],
            [],
            None,
            CompatibilityImpact::None,
            RepairAvailability::Unavailable,
            bounds,
        )
        .unwrap();

        let forward = ValidateReport::try_new(
            [check_a.clone(), check_b.clone()],
            [issue_a.clone(), issue_b.clone()],
            bounds,
        )
        .unwrap();
        let reversed =
            ValidateReport::try_new([check_b, check_a], [issue_b, issue_a], bounds).unwrap();

        assert_eq!(forward, reversed);
        assert_eq!(forward.checks()[0].id().as_str(), "check.a");
        assert!(forward.issues()[0].id() < forward.issues()[1].id());
        assert_eq!(
            serde_json::to_string(&forward).unwrap(),
            serde_json::to_string(&reversed).unwrap()
        );
    }
}
