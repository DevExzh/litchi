//! Bounded, non-mutating semantic validation of legacy PowerPoint sources.
//!
//! This owner deliberately validates only facts that can be established from
//! the outer CFB directory, the `Current User` atom, and the bounded record
//! framing of `PowerPoint Document`. It never decrypts a stream, opens or
//! executes a VBA project, resolves an external target, verifies a signature,
//! or repairs a malformed source.

use crate::{Error, RecordLimits, records::Record};
use litchi_cfb::{OleError, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceDigest, EvidenceValue,
    IssueEvidence, IssueLocation, IssueSeverity, ReadAt, RepairAvailability, ValidateReport,
    ValidationCheck, ValidationIssue, ValidationLimits, ValidationReportError,
};
use litchi_ole_common::protection::is_protected_component;
use std::{collections::HashMap, error::Error as StdError, fmt, sync::Arc};

const CFB: &str = "ppt.cfb.ingress";
const HIERARCHY: &str = "ppt.storage.hierarchy";
const DOCUMENT: &str = "ppt.document.stream";
const CURRENT_USER: &str = "ppt.current_user.stream";
const RECORDS: &str = "ppt.record.parse";
const PICTURES: &str = "ppt.pictures.stream";
const ENCRYPTION: &str = "ppt.encryption.presence";
const SIGNATURE: &str = "ppt.signature.presence";
const MACRO: &str = "ppt.macro.presence";
const PROTECTION: &str = "ppt.protection.presence";
const EXTERNAL: &str = "ppt.external_reference.presence";
const STREAM_BUDGET: &str = "ppt.stream.budget";

const CHECK_IDS: [&str; 12] = [
    CFB,
    HIERARCHY,
    DOCUMENT,
    CURRENT_USER,
    RECORDS,
    PICTURES,
    ENCRYPTION,
    SIGNATURE,
    MACRO,
    PROTECTION,
    EXTERNAL,
    STREAM_BUDGET,
];

const POWERPOINT_DOCUMENT: &str = "PowerPoint Document";
const CURRENT_USER_NAME: &str = "Current User";
const PICTURES_NAME: &str = "Pictures";
const DUAL_STORAGE: &str = "PP97_DUALSTORAGE";
const USER_EDIT_ATOM: u16 = 4085;
const DOCUMENT_RECORD: u16 = 1000;
const DOCUMENT_VERSION: u16 = 0x0F;
const DOCUMENT_ATOM: u16 = 1001;
const VBA_INFO: u16 = 1023;
const VBA_INFO_ATOM: u16 = 1024;
const CSTRING: u16 = 4026;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CurrentUserBootstrapError {
    ResourceLimit,
    Invalid,
}

#[derive(Debug, Clone, Copy)]
struct CurrentUserBootstrap {
    current_edit_offset: u32,
    encrypted: bool,
}

impl CurrentUserBootstrap {
    fn parse(data: &[u8], limits: RecordLimits) -> Result<Self, CurrentUserBootstrapError> {
        const HEADER_SIZE: usize = 8;
        const FIXED_PAYLOAD_SIZE: usize = 20;
        const MIN_SIZE: usize = HEADER_SIZE + FIXED_PAYLOAD_SIZE;
        const RECORD_TYPE: u16 = 0x0FF6;
        const UNENCRYPTED_TOKEN: u32 = 0xE391_C05F;
        const ENCRYPTED_TOKEN: u32 = 0xF3D1_C4DF;

        if data.len() > limits.max_input_bytes {
            return Err(CurrentUserBootstrapError::ResourceLimit);
        }
        if data.len() < MIN_SIZE {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let version_instance = u16::from_le_bytes([data[0], data[1]]);
        let record_type = u16::from_le_bytes([data[2], data[3]]);
        if version_instance != 0 || record_type != RECORD_TYPE {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let declared_payload =
            usize::try_from(u32::from_le_bytes([data[4], data[5], data[6], data[7]]))
                .map_err(|_error| CurrentUserBootstrapError::Invalid)?;
        let actual_payload = data
            .len()
            .checked_sub(HEADER_SIZE)
            .ok_or(CurrentUserBootstrapError::Invalid)?;
        if declared_payload != actual_payload {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let fixed_size = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        if fixed_size != 20 {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let encrypted = match u32::from_le_bytes([data[12], data[13], data[14], data[15]]) {
            UNENCRYPTED_TOKEN => false,
            ENCRYPTED_TOKEN => true,
            _ => return Err(CurrentUserBootstrapError::Invalid),
        };
        let current_edit_offset = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
        let document_version = u16::from_le_bytes([data[22], data[23]]);
        let major_version = data[24];
        let minor_version = data[25];
        if document_version != 0x03F4 || major_version != 3 || minor_version != 0 {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let username_len = usize::from(u16::from_le_bytes([data[20], data[21]]));
        if username_len > 255 {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let ansi_start = MIN_SIZE;
        let release_start = ansi_start
            .checked_add(username_len)
            .ok_or(CurrentUserBootstrapError::Invalid)?;
        let release_end = release_start
            .checked_add(4)
            .ok_or(CurrentUserBootstrapError::Invalid)?;
        if release_end > data.len() {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        if !is_printable_ansi(&data[ansi_start..release_start]) {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let release_version = u32::from_le_bytes([
            data[release_start],
            data[release_start + 1],
            data[release_start + 2],
            data[release_start + 3],
        ]);
        if !matches!(release_version, 8 | 9) {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        let unicode_len = username_len
            .checked_mul(2)
            .ok_or(CurrentUserBootstrapError::Invalid)?;
        let trailing = data
            .len()
            .checked_sub(release_end)
            .ok_or(CurrentUserBootstrapError::Invalid)?;
        if trailing != 0 && trailing != unicode_len {
            return Err(CurrentUserBootstrapError::Invalid);
        }
        if trailing != 0 && !is_printable_utf16(&data[release_end..]) {
            return Err(CurrentUserBootstrapError::Invalid);
        }

        Ok(Self {
            current_edit_offset,
            encrypted,
        })
    }

    const fn is_encrypted(self) -> bool {
        self.encrypted
    }

    const fn current_edit_offset(self) -> u32 {
        self.current_edit_offset
    }
}

fn is_printable_ansi(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| !matches!(*byte, 0x00..=0x1F | 0x7F..=0x9F))
}

fn is_printable_utf16(bytes: &[u8]) -> bool {
    if !bytes.len().is_multiple_of(2) {
        return false;
    }
    let mut offset = 0_usize;
    while offset < bytes.len() {
        let unit = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
        offset += 2;
        let codepoint = if (0xD800..=0xDBFF).contains(&unit) {
            if offset >= bytes.len() {
                return false;
            }
            let low = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
            offset += 2;
            if !(0xDC00..=0xDFFF).contains(&low) {
                return false;
            }
            0x1_0000 + ((u32::from(unit) - 0xD800) << 10) + (u32::from(low) - 0xDC00)
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return false;
        } else {
            u32::from(unit)
        };
        if codepoint < 0x20 || (0x7F..=0x9F).contains(&codepoint) {
            return false;
        }
    }
    true
}

/// Finite source, CFB directory, record, and report bounds for one PPT
/// validation pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PptValidationLimits {
    max_input_bytes: u64,
    max_document_stream_bytes: u64,
    max_current_user_stream_bytes: u64,
    max_pictures_stream_bytes: u64,
    max_aggregate_stream_bytes: u64,
    max_directory_entries: usize,
    max_directory_depth: usize,
    record: RecordLimits,
    report: ValidationLimits,
}

impl PptValidationLimits {
    /// Creates explicit finite PPT validation bounds.
    #[must_use]
    #[allow(
        clippy::too_many_arguments,
        reason = "each independent untrusted-resource bound is explicit"
    )]
    pub const fn new(
        max_input_bytes: u64,
        max_document_stream_bytes: u64,
        max_current_user_stream_bytes: u64,
        max_pictures_stream_bytes: u64,
        max_aggregate_stream_bytes: u64,
        max_directory_entries: usize,
        max_directory_depth: usize,
        record: RecordLimits,
        report: ValidationLimits,
    ) -> Self {
        Self {
            max_input_bytes,
            max_document_stream_bytes,
            max_current_user_stream_bytes,
            max_pictures_stream_bytes,
            max_aggregate_stream_bytes,
            max_directory_entries,
            max_directory_depth,
            record,
            report,
        }
    }

    /// Maximum physical CFB source length inspected.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum `PowerPoint Document` stream length materialized.
    #[must_use]
    pub const fn max_document_stream_bytes(self) -> u64 {
        self.max_document_stream_bytes
    }

    /// Maximum `Current User` stream length materialized.
    #[must_use]
    pub const fn max_current_user_stream_bytes(self) -> u64 {
        self.max_current_user_stream_bytes
    }

    /// Maximum optional `Pictures` stream length accepted by the inventory.
    #[must_use]
    pub const fn max_pictures_stream_bytes(self) -> u64 {
        self.max_pictures_stream_bytes
    }

    /// Maximum combined size of the three native PPT streams.
    #[must_use]
    pub const fn max_aggregate_stream_bytes(self) -> u64 {
        self.max_aggregate_stream_bytes
    }

    /// Maximum CFB directory entries traversed by the semantic inventory.
    #[must_use]
    pub const fn max_directory_entries(self) -> usize {
        self.max_directory_entries
    }

    /// Maximum directory nesting traversed by the semantic inventory.
    #[must_use]
    pub const fn max_directory_depth(self) -> usize {
        self.max_directory_depth
    }

    /// Bounded record-parser policy used for `PowerPoint Document` and
    /// `Current User`.
    #[must_use]
    pub const fn record(self) -> RecordLimits {
        self.record
    }

    /// Bounds for the retained format-neutral report.
    #[must_use]
    pub const fn report(self) -> ValidationLimits {
        self.report
    }
}

impl Default for PptValidationLimits {
    fn default() -> Self {
        let record = RecordLimits::default();
        Self {
            max_input_bytes: record.max_package_bytes as u64,
            max_document_stream_bytes: record.max_input_bytes as u64,
            max_current_user_stream_bytes: record.max_input_bytes as u64,
            max_pictures_stream_bytes: record.max_input_bytes as u64,
            max_aggregate_stream_bytes: record.max_aggregate_input_bytes as u64,
            max_directory_entries: 100_000,
            max_directory_depth: 256,
            record,
            report: ValidationLimits::default(),
        }
    }
}

/// Failure to perform or retain a PPT validation report.
#[derive(Debug)]
#[non_exhaustive]
pub enum PptValidationError {
    /// CFB positional ingress or source stability failed.
    Ingress(OleError),
    /// A bounded validator-owned allocation failed.
    Allocation(&'static str),
    /// The shared bounded report rejected the retained result.
    Report(ValidationReportError),
}

impl fmt::Display for PptValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ingress(error) => write!(formatter, "PPT validation ingress failed: {error}"),
            Self::Allocation(resource) => {
                write!(
                    formatter,
                    "allocation failed while validating PPT {resource}"
                )
            },
            Self::Report(error) => write!(formatter, "PPT validation report failed: {error}"),
        }
    }
}

impl StdError for PptValidationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Ingress(error) => Some(error),
            Self::Allocation(_) => None,
            Self::Report(error) => Some(error),
        }
    }
}

impl From<ValidationReportError> for PptValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

impl From<OleError> for PptValidationError {
    fn from(error: OleError) -> Self {
        Self::Ingress(error)
    }
}

/// Validates a positional PPT source under default finite bounds.
///
/// The source remains untouched. The report is presence-only at security
/// boundaries: it records encryption, signatures, macros, protection, and
/// external references without decrypting, verifying, executing, fetching, or
/// repairing any of them.
///
/// # Errors
///
/// Returns an error for source I/O/version instability, bounded validator
/// allocation failure, or failure to retain the bounded report. Definite
/// format rejection and configured limits are represented in the report.
pub fn validate_source(source: Arc<dyn ReadAt>) -> Result<ValidateReport, PptValidationError> {
    validate_source_with_limits(source, PptValidationLimits::default())
}

/// Validates a positional PPT source under explicit finite bounds.
///
/// The validator checks CFB topology, the canonical root or
/// `PP97_DUALSTORAGE` stream layout, `Current User`, strict bounded record
/// framing, and inert security/presence markers. It never mutates the source.
///
/// # Errors
///
/// Returns an error for source I/O/version instability, bounded validator
/// allocation failure, or failure to retain the bounded report.
pub fn validate_source_with_limits(
    source: Arc<dyn ReadAt>,
    limits: PptValidationLimits,
) -> Result<ValidateReport, PptValidationError> {
    let expected_version = source
        .version()
        .map_err(|error| PptValidationError::Ingress(OleError::Io(error)))?;
    let source_size = source
        .len()
        .map_err(|error| PptValidationError::Ingress(OleError::Io(error)))?;
    require_version(source.as_ref(), expected_version)?;

    let input_limit = limits
        .max_input_bytes
        .min(SharedOleFileLimits::MAX_INPUT_BYTES);
    if input_limit == 0 || source_size > input_limit {
        let report = blocked_before_cfb(
            limits.report,
            "source exceeds the configured PPT input ceiling",
        )?;
        require_version(source.as_ref(), expected_version)?;
        return Ok(report);
    }

    let cfb_limits = SharedOleFileLimits::new(input_limit).map_err(PptValidationError::Ingress)?;
    let shared = match SharedOleFile::open_with_limits(source.clone(), cfb_limits) {
        Ok(shared) => shared,
        Err(error) if is_structural_cfb_error(&error) => {
            let report = cfb_rejection_report(&error, source_size, limits.report)?;
            require_version(source.as_ref(), expected_version)?;
            return Ok(report);
        },
        Err(error) => return Err(PptValidationError::Ingress(error)),
    };

    let mut issues = Vec::new();
    issues
        .try_reserve(16)
        .map_err(|_error| PptValidationError::Allocation("PPT validation issues"))?;
    let inventory = inspect_directory(&shared, limits)?;
    let mut statuses = initial_statuses(limits.report)?;

    if inventory.entry_limit || inventory.depth_limit {
        statuses[1] = CheckStatus::blocked(
            if inventory.entry_limit {
                "CFB directory traversal reached the configured PPT entry ceiling"
            } else {
                "CFB directory traversal reached the configured PPT depth ceiling"
            },
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                HIERARCHY,
                if inventory.entry_limit {
                    "ppt.storage.directory_limit"
                } else {
                    "ppt.storage.depth_limit"
                },
                IssueSeverity::Error,
                if inventory.entry_limit {
                    "The CFB directory exceeded the configured PPT semantic-inventory entry bound."
                } else {
                    "The CFB directory exceeded the configured PPT semantic-inventory depth bound."
                },
                Some("compound-file"),
                limits.report,
            )?,
        )?;
        statuses[2] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[3] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[4] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[5] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[6] = if inventory.encryption_markers != 0 {
            CheckStatus::Complete
        } else {
            CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?)
        };
        statuses[7] = if inventory.signature_markers != 0 {
            CheckStatus::Complete
        } else {
            CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?)
        };
        statuses[8] = if inventory.macro_storages != 0 {
            CheckStatus::Complete
        } else {
            CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?)
        };
        statuses[9] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[10] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        statuses[11] = CheckStatus::stopped_by(check_id(HIERARCHY, limits.report)?);
        let report = finish_report(statuses, issues, limits.report)?;
        shared
            .source_version()
            .map_err(PptValidationError::Ingress)?;
        return Ok(report);
    }

    if inventory.invalid {
        push_issue(
            &mut issues,
            simple_issue(
                HIERARCHY,
                "ppt.storage.hierarchy_invalid",
                IssueSeverity::Error,
                "The CFB directory does not contain one canonical PPT stream hierarchy.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
        statuses[1] = CheckStatus::blocked(
            "the CFB directory does not establish an exact canonical PPT hierarchy",
            limits.report,
        )?;
        let hierarchy_check = check_id(HIERARCHY, limits.report)?;
        if statuses[2].is_complete() {
            statuses[2] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[3].is_complete() {
            statuses[3] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[4].is_complete() {
            statuses[4] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[5].is_complete() {
            statuses[5] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[6].is_complete() && inventory.encryption_markers == 0 {
            statuses[6] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[7].is_complete() && inventory.signature_markers == 0 {
            statuses[7] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[8].is_complete() && inventory.macro_storages == 0 {
            statuses[8] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[9].is_complete() {
            statuses[9] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[10].is_complete() {
            statuses[10] = CheckStatus::stopped_by(hierarchy_check.clone());
        }
        if statuses[11].is_complete() {
            statuses[11] = CheckStatus::stopped_by(hierarchy_check);
        }
    }
    if inventory.document_count == 0 {
        statuses[2] = CheckStatus::blocked(
            "the canonical PowerPoint Document stream is absent",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                DOCUMENT,
                "ppt.document.stream_missing",
                IssueSeverity::Error,
                "The canonical PPT hierarchy has no PowerPoint Document stream.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    } else if inventory.document_count != 1 || inventory.document_not_stream {
        statuses[2] = CheckStatus::blocked(
            "PowerPoint Document stream resolution is ambiguous or not a stream",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                DOCUMENT,
                "ppt.document.stream_ambiguous",
                IssueSeverity::Error,
                "The canonical PPT hierarchy contains an ambiguous PowerPoint Document owner.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    }
    if inventory.current_user_count == 0 {
        statuses[3] =
            CheckStatus::blocked("the canonical Current User stream is absent", limits.report)?;
        push_issue(
            &mut issues,
            simple_issue(
                CURRENT_USER,
                "ppt.current_user.stream_missing",
                IssueSeverity::Error,
                "The canonical PPT hierarchy has no Current User stream.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    } else if inventory.current_user_count != 1 || inventory.current_user_not_stream {
        statuses[3] = CheckStatus::blocked(
            "Current User stream resolution is ambiguous or not a stream",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                CURRENT_USER,
                "ppt.current_user.stream_ambiguous",
                IssueSeverity::Error,
                "The canonical PPT hierarchy contains an ambiguous Current User owner.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    }
    if inventory.pictures_noncanonical {
        statuses[5] = CheckStatus::blocked(
            "Pictures must be a root-level stream in the canonical PPT hierarchy",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                PICTURES,
                "ppt.pictures.stream_noncanonical",
                IssueSeverity::Error,
                "The PPT Pictures stream is not at the canonical root-level location.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    } else if inventory.pictures_count > 1 || inventory.pictures_not_stream {
        statuses[5] = CheckStatus::blocked(
            "Pictures stream resolution is ambiguous or not a stream",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                PICTURES,
                "ppt.pictures.stream_ambiguous",
                IssueSeverity::Error,
                "The PPT hierarchy contains an ambiguous Pictures owner.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    } else if inventory.pictures_count == 0 {
        statuses[5] = CheckStatus::NotApplicable;
    }

    if inventory.encryption_markers != 0 {
        push_presence_issue(
            &mut issues,
            ENCRYPTION,
            "ppt.encryption.infrastructure_present",
            IssueSeverity::Warning,
            "Encryption metadata is present; this validator did not decrypt or authenticate the presentation.",
            inventory.encryption_markers,
            limits.report,
        )?;
    }
    if inventory.signature_markers != 0 {
        push_presence_issue(
            &mut issues,
            SIGNATURE,
            "ppt.signature.infrastructure_present",
            IssueSeverity::Info,
            "Signature infrastructure is present; certificate trust and cryptographic validity were not checked.",
            inventory.signature_markers,
            limits.report,
        )?;
    }
    if inventory.drm_markers != 0 {
        push_presence_issue(
            &mut issues,
            ENCRYPTION,
            "ppt.drm.infrastructure_present",
            IssueSeverity::Warning,
            "DRM metadata is present; protected content was not opened or evaluated.",
            inventory.drm_markers,
            limits.report,
        )?;
    }
    if inventory.macro_storages != 0 {
        push_presence_issue(
            &mut issues,
            MACRO,
            "ppt.macro.storage_present",
            IssueSeverity::Warning,
            "A VBA-related CFB storage is present; macro bytes were not opened or executed.",
            inventory.macro_storages,
            limits.report,
        )?;
    }

    let document_size = inventory.document_size;
    let current_user_size = inventory.current_user_size;
    let pictures_size = inventory.pictures_size;
    let aggregate_size = document_size
        .checked_add(current_user_size)
        .and_then(|value| value.checked_add(pictures_size));
    let stream_budget_ok =
        aggregate_size.is_some_and(|value| value <= limits.max_aggregate_stream_bytes);
    if !stream_budget_ok {
        statuses[11] = CheckStatus::blocked(
            "the native PPT stream aggregate exceeds the configured byte ceiling",
            limits.report,
        )?;
        push_issue(
            &mut issues,
            simple_issue(
                STREAM_BUDGET,
                "ppt.stream.aggregate_limit",
                IssueSeverity::Error,
                "The combined native PPT streams exceed the configured semantic-validation byte bound.",
                Some("compound-file"),
                limits.report,
            )?,
        )?;
    }
    if !stream_budget_ok {
        let budget_check = check_id(STREAM_BUDGET, limits.report)?;
        if statuses[2].is_complete() {
            statuses[2] = CheckStatus::stopped_by(budget_check.clone());
        }
        if statuses[3].is_complete() {
            statuses[3] = CheckStatus::stopped_by(budget_check.clone());
        }
        if statuses[4].is_complete() {
            statuses[4] = CheckStatus::stopped_by(budget_check);
        }
    }
    if document_size > limits.max_document_stream_bytes {
        statuses[2] = CheckStatus::blocked(
            "PowerPoint Document exceeds the configured PPT stream byte ceiling",
            limits.report,
        )?;
        statuses[4] = CheckStatus::stopped_by(check_id(DOCUMENT, limits.report)?);
    }
    if current_user_size > limits.max_current_user_stream_bytes {
        statuses[3] = CheckStatus::blocked(
            "Current User exceeds the configured PPT stream byte ceiling",
            limits.report,
        )?;
    }
    if inventory.pictures_count != 0 && pictures_size > limits.max_pictures_stream_bytes {
        statuses[5] = CheckStatus::blocked(
            "Pictures exceeds the configured PPT stream byte ceiling",
            limits.report,
        )?;
    }

    let mut current_data = None;
    if stream_budget_ok && statuses[3].is_complete() {
        if let Some(path) = inventory.current_user_path {
            current_data = Some(
                shared
                    .open_stream(path)
                    .map_err(PptValidationError::Ingress)?,
            );
        }
    }

    let current = match current_data.as_deref() {
        None => None,
        Some(data) => match CurrentUserBootstrap::parse(data, limits.record) {
            Ok(value) => Some(value),
            Err(error) => {
                let limited = matches!(error, CurrentUserBootstrapError::ResourceLimit);
                let (code, message) = if limited {
                    (
                        "ppt.current_user.limit",
                        "The Current User atom exceeded the configured bounded parser policy.",
                    )
                } else {
                    (
                        "ppt.current_user.invalid",
                        "The Current User atom is malformed or violates the configured bounded parser policy.",
                    )
                };
                push_current_user_error(&mut issues, code, message, limits.report)?;
                if limited {
                    statuses[3] = CheckStatus::blocked(
                        "Current User parsing reached a configured bounded parser limit",
                        limits.report,
                    )?;
                } else {
                    statuses[3] = CheckStatus::blocked(
                        "Current User parsing did not establish encryption state",
                        limits.report,
                    )?;
                }
                None
            },
        },
    };
    if current.as_ref().is_some_and(|value| value.is_encrypted()) {
        push_presence_issue(
            &mut issues,
            ENCRYPTION,
            "ppt.encryption.password_to_open_present",
            IssueSeverity::Warning,
            "Current User identifies password-to-open encryption; the PowerPoint Document stream was not decrypted.",
            1,
            limits.report,
        )?;
        statuses[4] = CheckStatus::blocked(
            "password-to-open encryption prevents clear PPT record inspection",
            limits.report,
        )?;
    }

    // Read and classify Current User before touching PowerPoint Document.
    // Password-to-open sources must stop at the encryption boundary rather
    // than materializing clear-looking ciphertext into the record parser.
    let mut document_data = None;
    let current_is_unencrypted = matches!(current, Some(ref value) if !value.is_encrypted());
    let current_edit_in_bounds = if current_is_unencrypted && statuses[2].is_complete() {
        let current_offset = current
            .as_ref()
            .map(|value| value.current_edit_offset())
            .expect("the unencrypted Current User state was established");
        if u64::from(current_offset) < document_size {
            true
        } else {
            push_issue(
                &mut issues,
                simple_issue(
                    CURRENT_USER,
                    "ppt.current_user.edit_offset_out_of_bounds",
                    IssueSeverity::Error,
                    "Current User points outside the declared PowerPoint Document stream.",
                    Some(CURRENT_USER_NAME),
                    limits.report,
                )?,
            )?;
            statuses[3] = CheckStatus::blocked(
                "Current User points outside the declared PowerPoint Document stream",
                limits.report,
            )?;
            if statuses[4].is_complete() {
                statuses[4] = CheckStatus::stopped_by(check_id(CURRENT_USER, limits.report)?);
            }
            false
        }
    } else {
        false
    };
    if current_edit_in_bounds
        && stream_budget_ok
        && statuses[2].is_complete()
        && let Some(path) = inventory.document_path
    {
        document_data = Some(
            shared
                .open_stream(path)
                .map_err(PptValidationError::Ingress)?,
        );
    }

    if current.is_none() {
        if statuses[6].is_complete() {
            statuses[6] = CheckStatus::blocked(
                "Current User parsing did not establish whether the presentation is encrypted",
                limits.report,
            )?;
        }
        if statuses[4].is_complete() {
            statuses[4] = CheckStatus::stopped_by(check_id(CURRENT_USER, limits.report)?);
        }
    }

    let records = if current_is_unencrypted
        && statuses[2].is_complete()
        && statuses[4].is_complete()
    {
        match document_data.as_deref() {
            Some(data) => match Record::parse_sequence_strict_with_limits(
                data,
                "PowerPoint Document",
                limits.record,
            ) {
                Ok(records) => Some(records),
                Err(error @ Error::ResourceLimit(_)) => {
                    statuses[4] = CheckStatus::blocked(
                        "PPT record traversal reached a configured byte, record, or depth ceiling",
                        limits.report,
                    )?;
                    push_record_error(
                        &mut issues,
                        RECORDS,
                        POWERPOINT_DOCUMENT,
                        "ppt.record.limit",
                        "The PowerPoint Document record tree exceeded the configured bounded parser policy.",
                        &error,
                        limits.report,
                    )?;
                    None
                },
                Err(Error::AllocationFailed(_)) => {
                    return Err(PptValidationError::Allocation("PPT record parser"));
                },
                Err(error @ Error::Io(_)) => {
                    return Err(PptValidationError::Ingress(match error {
                        Error::Io(error) => OleError::Io(error),
                        _ => unreachable!("matched PPT I/O error"),
                    }));
                },
                Err(error) => {
                    push_record_error(
                        &mut issues,
                        RECORDS,
                        POWERPOINT_DOCUMENT,
                        "ppt.record.invalid",
                        "The PowerPoint Document contains malformed record framing or nesting.",
                        &error,
                        limits.report,
                    )?;
                    statuses[4] = CheckStatus::blocked(
                        "PPT record grammar rejected the bounded PowerPoint Document stream",
                        limits.report,
                    )?;
                    None
                },
            },
            None => None,
        }
    } else {
        None
    };

    if let Some(records) = records.as_deref() {
        let observation = inspect_records(records)?;
        if observation.document_count == 0 {
            push_issue(
                &mut issues,
                simple_issue(
                    RECORDS,
                    "ppt.record.document_missing",
                    IssueSeverity::Error,
                    "The PowerPoint Document stream has no DocumentContainer record.",
                    Some(POWERPOINT_DOCUMENT),
                    limits.report,
                )?,
            )?;
        }
        if observation.document_atom_missing {
            push_issue(
                &mut issues,
                simple_issue(
                    RECORDS,
                    "ppt.record.document_atom_missing",
                    IssueSeverity::Error,
                    "A DocumentContainer is missing its required bounded DocumentAtom owner.",
                    Some(POWERPOINT_DOCUMENT),
                    limits.report,
                )?,
            )?;
        }
        if observation.document_header_invalid {
            push_issue(
                &mut issues,
                simple_issue(
                    RECORDS,
                    "ppt.record.document_header_invalid",
                    IssueSeverity::Error,
                    "A top-level DocumentContainer has an invalid bounded record header.",
                    Some(POWERPOINT_DOCUMENT),
                    limits.report,
                )?,
            )?;
        }
        if observation.document_atom_invalid {
            push_issue(
                &mut issues,
                simple_issue(
                    RECORDS,
                    "ppt.record.document_atom_invalid",
                    IssueSeverity::Error,
                    "A DocumentAtom failed its bounded MS-PPT semantic validation.",
                    Some(POWERPOINT_DOCUMENT),
                    limits.report,
                )?,
            )?;
        }
        if observation.macro_records != 0 {
            push_presence_issue(
                &mut issues,
                MACRO,
                "ppt.macro.record_present",
                IssueSeverity::Warning,
                "VBA metadata records are present; macro storage was not opened or executed.",
                observation.macro_records,
                limits.report,
            )?;
            statuses[8] = CheckStatus::Complete;
        } else if inventory.macro_storages == 0 {
            statuses[8] = CheckStatus::NotApplicable;
        }
        if observation.modify_password_records != 0 {
            push_presence_issue(
                &mut issues,
                PROTECTION,
                "ppt.protection.modify_password_present",
                IssueSeverity::Warning,
                "A PPT modify-password marker is present; this validator did not verify or bypass it.",
                observation.modify_password_records,
                limits.report,
            )?;
            statuses[9] = CheckStatus::Complete;
        } else {
            statuses[9] = CheckStatus::NotApplicable;
        }
        if observation.external_records != 0 {
            push_presence_issue(
                &mut issues,
                EXTERNAL,
                "ppt.external_reference.present",
                IssueSeverity::Info,
                "External object, hyperlink, or media metadata is present; targets were not fetched or activated.",
                observation.external_records,
                limits.report,
            )?;
            statuses[10] = CheckStatus::Complete;
        } else {
            statuses[10] = CheckStatus::NotApplicable;
        }
    } else {
        if statuses[4].is_complete() {
            statuses[4] = if statuses[2].is_complete() {
                CheckStatus::blocked(
                    "bounded PPT record inspection did not produce a record tree",
                    limits.report,
                )?
            } else {
                CheckStatus::stopped_by(check_id(DOCUMENT, limits.report)?)
            };
        }
        statuses[8] = if inventory.macro_storages != 0 {
            CheckStatus::Complete
        } else {
            CheckStatus::stopped_by(check_id(RECORDS, limits.report)?)
        };
        statuses[9] = CheckStatus::stopped_by(check_id(RECORDS, limits.report)?);
        statuses[10] = CheckStatus::stopped_by(check_id(RECORDS, limits.report)?);
    }

    if let (Some(current), Some(data)) = (current.as_ref(), document_data.as_deref()) {
        if !current.is_encrypted() && statuses[4].is_complete() {
            let current_offset = current.current_edit_offset();
            let offset = usize::try_from(current_offset)
                .map_err(|_error| PptValidationError::Allocation("Current User edit offset"))?;
            if offset >= data.len() {
                push_issue(
                    &mut issues,
                    simple_issue(
                        CURRENT_USER,
                        "ppt.current_user.edit_offset_out_of_bounds",
                        IssueSeverity::Error,
                        "Current User points outside the bounded PowerPoint Document stream.",
                        Some(CURRENT_USER_NAME),
                        limits.report,
                    )?,
                )?;
            } else {
                match Record::parse_strict_with_limits(data, offset, limits.record) {
                    Ok((record, _consumed)) if is_valid_user_edit_atom(&record, current_offset) => {
                    },
                    Ok((_record, _consumed)) => {
                        push_issue(
                            &mut issues,
                            simple_issue(
                                CURRENT_USER,
                                "ppt.current_user.edit_target_invalid",
                                IssueSeverity::Error,
                                "Current User does not point to a UserEditAtom record.",
                                Some(CURRENT_USER_NAME),
                                limits.report,
                            )?,
                        )?;
                    },
                    Err(error) => {
                        push_record_error(
                            &mut issues,
                            CURRENT_USER,
                            CURRENT_USER_NAME,
                            "ppt.current_user.edit_target_malformed",
                            "The Current User edit target is not a bounded complete record.",
                            &error,
                            limits.report,
                        )?;
                    },
                }
            }
        }
    }

    shared
        .source_version()
        .map_err(PptValidationError::Ingress)?;
    require_version(source.as_ref(), expected_version)?;
    finish_report(statuses, issues, limits.report)
}

fn is_valid_user_edit_atom(record: &Record, current_offset: u32) -> bool {
    if record.record_type_raw != USER_EDIT_ATOM
        || record.version != 0
        || record.instance != 0
        || !matches!(record.data_length, 28 | 32)
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return false;
    }
    let data = record.data.as_slice();
    let Some(offset_last_edit) = little_u32_at(data, 8) else {
        return false;
    };
    let Some(offset_persist_directory) = little_u32_at(data, 12) else {
        return false;
    };
    let Some(doc_persist_id_ref) = little_u32_at(data, 16) else {
        return false;
    };
    data.get(6) == Some(&0)
        && data.get(7) == Some(&3)
        && offset_last_edit < current_offset
        && offset_persist_directory > offset_last_edit
        && offset_persist_directory < current_offset
        && doc_persist_id_ref == 1
}

fn little_u32_at(data: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        data.get(offset..offset.checked_add(4)?)?.try_into().ok()?,
    ))
}

#[derive(Debug, Default)]
struct DirectoryObservation {
    entry_count: usize,
    entry_limit: bool,
    depth_limit: bool,
    invalid: bool,
    document_count: usize,
    document_not_stream: bool,
    document_path: Option<&'static [&'static str]>,
    current_user_count: usize,
    current_user_not_stream: bool,
    current_user_path: Option<&'static [&'static str]>,
    pictures_count: usize,
    pictures_noncanonical: bool,
    pictures_not_stream: bool,
    document_size: u64,
    current_user_size: u64,
    pictures_size: u64,
    encryption_markers: u64,
    signature_markers: u64,
    drm_markers: u64,
    macro_storages: u64,
}

const ROOT_DOCUMENT_PATH: &[&str] = &[POWERPOINT_DOCUMENT];
const DUAL_DOCUMENT_PATH: &[&str] = &[DUAL_STORAGE, POWERPOINT_DOCUMENT];
const ROOT_CURRENT_USER_PATH: &[&str] = &[CURRENT_USER_NAME];
const DUAL_CURRENT_USER_PATH: &[&str] = &[DUAL_STORAGE, CURRENT_USER_NAME];

fn inspect_directory(
    shared: &SharedOleFile,
    limits: PptValidationLimits,
) -> Result<DirectoryObservation, PptValidationError> {
    let mut observation = DirectoryObservation::default();
    for entry in shared.directory_entries() {
        if entry.entry_type == litchi_cfb::consts::STGTY_ROOT {
            continue;
        }
        observation.entry_count = observation
            .entry_count
            .checked_add(1)
            .ok_or(PptValidationError::Allocation("PPT directory entry count"))?;
        if observation.entry_count > limits.max_directory_entries {
            observation.entry_limit = true;
            break;
        }
        if is_protected_component(&entry.name) {
            if is_signature_component(&entry.name) {
                observation.signature_markers = observation.signature_markers.saturating_add(1);
            } else if is_drm_component(&entry.name) {
                observation.drm_markers = observation.drm_markers.saturating_add(1);
            } else {
                observation.encryption_markers = observation.encryption_markers.saturating_add(1);
            }
        }
        if is_macro_storage_name(&entry.name) {
            observation.macro_storages = observation.macro_storages.saturating_add(1);
        }
    }
    if observation.entry_limit {
        return Ok(observation);
    }

    inspect_directory_depth(shared, limits, &mut observation)?;
    if observation.depth_limit {
        return Ok(observation);
    }

    let root_document = probe_stream(shared, ROOT_DOCUMENT_PATH)?;
    let dual_document = probe_stream(shared, DUAL_DOCUMENT_PATH)?;
    let root_current = probe_stream(shared, ROOT_CURRENT_USER_PATH)?;
    let dual_current = probe_stream(shared, DUAL_CURRENT_USER_PATH)?;
    let root_pictures = probe_stream(shared, &[PICTURES_NAME])?;
    let dual_pictures = probe_stream(shared, &[DUAL_STORAGE, PICTURES_NAME])?;

    let document_entry_count = count_named_entries(shared, POWERPOINT_DOCUMENT);
    let current_entry_count = count_named_entries(shared, CURRENT_USER_NAME);
    let pictures_entry_count = count_named_entries(shared, PICTURES_NAME);

    observation.document_count =
        usize::from(root_document.exists).saturating_add(usize::from(dual_document.exists));
    observation.current_user_count =
        usize::from(root_current.exists).saturating_add(usize::from(dual_current.exists));
    observation.pictures_count = usize::from(root_pictures.exists);
    observation.pictures_noncanonical = dual_pictures.exists
        || dual_pictures.non_stream
        || pictures_entry_count != observation.pictures_count;
    if document_entry_count != observation.document_count
        || current_entry_count != observation.current_user_count
        || pictures_entry_count != observation.pictures_count
    {
        // A named leaf outside the two canonical layouts is not a harmless
        // unknown extension: it makes the required native stream topology
        // ambiguous, so the semantic owner fails closed.
        observation.invalid = true;
        observation.document_count = observation.document_count.max(document_entry_count);
        observation.current_user_count = observation.current_user_count.max(current_entry_count);
        observation.pictures_count = observation.pictures_count.max(pictures_entry_count);
    }
    if observation.pictures_noncanonical {
        observation.invalid = true;
    }

    observation.document_not_stream = root_document.non_stream || dual_document.non_stream;
    observation.current_user_not_stream = root_current.non_stream || dual_current.non_stream;
    observation.pictures_not_stream = root_pictures.non_stream;
    if root_document.exists && dual_document.exists {
        observation.invalid = true;
    }
    if root_current.exists && dual_current.exists {
        observation.invalid = true;
    }
    if root_document.exists && dual_current.exists || dual_document.exists && root_current.exists {
        observation.invalid = true;
    }
    observation.document_path = if root_document.exists {
        Some(ROOT_DOCUMENT_PATH)
    } else if dual_document.exists {
        Some(DUAL_DOCUMENT_PATH)
    } else {
        None
    };
    observation.current_user_path = if root_current.exists {
        Some(ROOT_CURRENT_USER_PATH)
    } else if dual_current.exists {
        Some(DUAL_CURRENT_USER_PATH)
    } else {
        None
    };
    observation.document_size = root_document.size.or(dual_document.size).unwrap_or(0);
    observation.current_user_size = root_current.size.or(dual_current.size).unwrap_or(0);
    observation.pictures_size = root_pictures.size.unwrap_or(0);
    Ok(observation)
}

fn inspect_directory_depth(
    shared: &SharedOleFile,
    limits: PptValidationLimits,
    observation: &mut DirectoryObservation,
) -> Result<(), PptValidationError> {
    let entry_count = shared.directory_entries().count();
    let mut by_sid = HashMap::new();
    by_sid
        .try_reserve(entry_count)
        .map_err(|_error| PptValidationError::Allocation("PPT directory depth index"))?;
    for entry in shared.directory_entries() {
        by_sid.insert(entry.sid, entry);
    }

    let Some(root) = by_sid.get(&0).copied() else {
        observation.invalid = true;
        return Ok(());
    };
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_error| PptValidationError::Allocation("PPT directory depth stack"))?;
    if root.sid_child != litchi_cfb::consts::NOSTREAM {
        pending.push((root.sid_child, 1_usize));
    }

    while let Some((sid, depth)) = pending.pop() {
        if sid == litchi_cfb::consts::NOSTREAM {
            continue;
        }
        if depth > limits.max_directory_depth {
            observation.depth_limit = true;
            break;
        }
        let Some(entry) = by_sid.get(&sid).copied() else {
            observation.invalid = true;
            break;
        };
        for sibling in [entry.sid_left, entry.sid_right] {
            if sibling != litchi_cfb::consts::NOSTREAM {
                pending.try_reserve(1).map_err(|_error| {
                    PptValidationError::Allocation("PPT directory depth stack")
                })?;
                pending.push((sibling, depth));
            }
        }
        if entry.sid_child != litchi_cfb::consts::NOSTREAM {
            let Some(child_depth) = depth.checked_add(1) else {
                observation.depth_limit = true;
                break;
            };
            pending
                .try_reserve(1)
                .map_err(|_error| PptValidationError::Allocation("PPT directory depth stack"))?;
            pending.push((entry.sid_child, child_depth));
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, Default)]
struct StreamProbe {
    exists: bool,
    non_stream: bool,
    size: Option<u64>,
}

fn probe_stream(shared: &SharedOleFile, path: &[&str]) -> Result<StreamProbe, PptValidationError> {
    if !shared.exists(path) {
        return Ok(StreamProbe::default());
    }
    match shared.stream_len(path) {
        Ok(size) => Ok(StreamProbe {
            exists: true,
            non_stream: false,
            size: Some(size),
        }),
        Err(OleError::InvalidFormat(_)) => Ok(StreamProbe {
            exists: true,
            non_stream: true,
            size: None,
        }),
        Err(error) => Err(PptValidationError::Ingress(error)),
    }
}

fn count_named_entries(shared: &SharedOleFile, name: &str) -> usize {
    shared
        .directory_entries()
        .filter(|entry| entry.entry_type != litchi_cfb::consts::STGTY_ROOT)
        .filter(|entry| equal_name(&entry.name, name))
        .count()
}

#[derive(Debug, Default)]
struct RecordObservation {
    document_count: usize,
    document_header_invalid: bool,
    document_atom_missing: bool,
    document_atom_invalid: bool,
    macro_records: u64,
    modify_password_records: u64,
    external_records: u64,
}

fn inspect_records(records: &[Record]) -> Result<RecordObservation, PptValidationError> {
    let mut observation = RecordObservation::default();
    let mut pending = Vec::new();
    pending
        .try_reserve(records.len())
        .map_err(|_error| PptValidationError::Allocation("PPT record traversal stack"))?;
    pending.extend(records.iter().map(|record| (record, true)));
    while let Some((record, top_level)) = pending.pop() {
        if top_level && record.record_type_raw == DOCUMENT_RECORD {
            if record.version != DOCUMENT_VERSION || record.instance != 0 {
                observation.document_header_invalid = true;
            } else {
                observation.document_count = observation.document_count.saturating_add(1);
                let mut atom = None;
                let mut atom_count = 0_usize;
                for child in &record.children {
                    if child.record_type_raw == DOCUMENT_ATOM {
                        atom_count = atom_count.saturating_add(1);
                        atom = Some(child);
                    }
                }
                if atom_count != 1 {
                    observation.document_atom_missing = true;
                } else if atom.is_some_and(|value| crate::DocumentAtom::parse(value).is_err()) {
                    observation.document_atom_invalid = true;
                }
            }
        }
        if record.record_type_raw == VBA_INFO || record.record_type_raw == VBA_INFO_ATOM {
            observation.macro_records = observation.macro_records.saturating_add(1);
        }
        if record.record_type_raw == CSTRING && record.instance == 3 {
            observation.modify_password_records =
                observation.modify_password_records.saturating_add(1);
        }
        if is_external_record(record.record_type_raw) {
            observation.external_records = observation.external_records.saturating_add(1);
        }
        pending
            .try_reserve(record.children.len())
            .map_err(|_error| PptValidationError::Allocation("PPT record traversal stack"))?;
        pending.extend(record.children.iter().map(|child| (child, false)));
    }
    Ok(observation)
}

fn is_external_record(record_type: u16) -> bool {
    matches!(
        record_type,
        3009 | 0x0FC3
            | 0x0FCC
            | 0x0FCD
            | 0x0FCE
            | 0x0FD1
            | 0x0FEE
            | 0x0FFB
            | 0x1004
            | 0x1005
            | 0x1006
            | 0x1007
            | 0x100D
            | 0x100E
            | 0x100F
            | 0x1010
            | 0x1011
            | 4051
            | 4055
            | 4068
            | 4082
            | 4083
    )
}

fn is_macro_storage_name(name: &str) -> bool {
    ["VBA", "VBAProject", "VbaProjectStg", "_VBA_PROJECT"]
        .iter()
        .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn is_signature_component(name: &str) -> bool {
    [
        "_xmlsignatures",
        "_signatures",
        "DigitalSignature",
        "\u{0005}DigitalSignature",
    ]
    .iter()
    .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn is_drm_component(name: &str) -> bool {
    ["\u{0009}DRMContent", "\u{0009}DRMViewerContent"]
        .iter()
        .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn equal_name(left: &str, right: &str) -> bool {
    left.eq_ignore_ascii_case(right)
}

fn initial_statuses(
    limits: ValidationLimits,
) -> Result<[CheckStatus; CHECK_IDS.len()], PptValidationError> {
    let _ = limits;
    Ok(std::array::from_fn(|_| CheckStatus::Complete))
}

fn blocked_before_cfb(
    limits: ValidationLimits,
    reason: &str,
) -> Result<ValidateReport, PptValidationError> {
    let cfb = check_id(CFB, limits)?;
    let statuses = [
        CheckStatus::blocked(reason, limits)?,
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb.clone()),
        CheckStatus::stopped_by(cfb),
    ];
    finish_report(statuses, Vec::new(), limits)
}

fn cfb_rejection_report(
    error: &OleError,
    source_size: u64,
    limits: ValidationLimits,
) -> Result<ValidateReport, PptValidationError> {
    let blocked = CheckStatus::blocked("CFB ingress was structurally rejected", limits)?;
    let statuses = [
        CheckStatus::Complete,
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked,
    ];
    let issue = ValidationIssue::try_new(
        check_id(CFB, limits)?,
        cfb_error_code(error),
        IssueSeverity::Error,
        "The CFB ingress rejected the source before PPT semantic validation could begin.",
        [IssueLocation::try_new(
            Some("compound-file"),
            None,
            None,
            None,
            None,
            limits,
        )?],
        [
            IssueEvidence::try_new("source.size", EvidenceValue::Size(source_size), limits)?,
            IssueEvidence::try_new(
                "diagnostic.sha256",
                EvidenceValue::Sha256(EvidenceDigest::of(cfb_error_code(error).as_bytes())),
                limits,
            )?,
        ],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?;
    let mut issues = Vec::new();
    issues
        .try_reserve_exact(1)
        .map_err(|_error| PptValidationError::Allocation("PPT CFB validation issues"))?;
    issues.push(issue);
    finish_report(statuses, issues, limits)
}

fn is_structural_cfb_error(error: &OleError) -> bool {
    matches!(
        error,
        OleError::InvalidFormat(_)
            | OleError::InvalidData(_)
            | OleError::NotOleFile
            | OleError::CorruptedFile(_)
            | OleError::StreamNotFound
    )
}

fn cfb_error_code(error: &OleError) -> &'static str {
    match error {
        OleError::NotOleFile => "ppt.cfb.not_ole",
        OleError::InvalidFormat(_) => "ppt.cfb.invalid_format",
        OleError::InvalidData(_) => "ppt.cfb.invalid_data",
        OleError::CorruptedFile(_) => "ppt.cfb.corrupted",
        OleError::StreamNotFound => "ppt.cfb.stream_missing",
        OleError::Io(_)
        | OleError::Allocation { .. }
        | OleError::Committed { .. }
        | OleError::SourceChanged { .. } => unreachable!("non-structural CFB error"),
    }
}

fn simple_issue(
    check: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    part: Option<&str>,
    limits: ValidationLimits,
) -> Result<ValidationIssue, PptValidationError> {
    Ok(ValidationIssue::try_new(
        check_id(check, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            part, None, None, None, None, limits,
        )?],
        [],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?)
}

fn push_presence_issue(
    issues: &mut Vec<ValidationIssue>,
    check: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    count: u64,
    limits: ValidationLimits,
) -> Result<(), PptValidationError> {
    let issue = ValidationIssue::try_new(
        check_id(check, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            Some("PowerPoint Document"),
            None,
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            "presence.count",
            EvidenceValue::Count(count),
            limits,
        )?],
        None,
        CompatibilityImpact::None,
        RepairAvailability::Unavailable,
        limits,
    )?;
    push_issue(issues, issue)
}

fn push_record_error(
    issues: &mut Vec<ValidationIssue>,
    check: &str,
    part: &str,
    code: &str,
    message: &str,
    _error: &Error,
    limits: ValidationLimits,
) -> Result<(), PptValidationError> {
    let issue = ValidationIssue::try_new(
        check_id(check, limits)?,
        code,
        IssueSeverity::Error,
        message,
        [IssueLocation::try_new(
            Some(part),
            None,
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            "diagnostic.sha256",
            EvidenceValue::Sha256(EvidenceDigest::of(code.as_bytes())),
            limits,
        )?],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?;
    push_issue(issues, issue)
}

fn push_current_user_error(
    issues: &mut Vec<ValidationIssue>,
    code: &str,
    message: &str,
    limits: ValidationLimits,
) -> Result<(), PptValidationError> {
    push_issue(
        issues,
        simple_issue(
            CURRENT_USER,
            code,
            IssueSeverity::Error,
            message,
            Some(CURRENT_USER_NAME),
            limits,
        )?,
    )
}

fn push_issue(
    issues: &mut Vec<ValidationIssue>,
    issue: ValidationIssue,
) -> Result<(), PptValidationError> {
    issues
        .try_reserve(1)
        .map_err(|_error| PptValidationError::Allocation("PPT validation issues"))?;
    issues.push(issue);
    Ok(())
}

fn finish_report(
    statuses: [CheckStatus; CHECK_IDS.len()],
    issues: Vec<ValidationIssue>,
    limits: ValidationLimits,
) -> Result<ValidateReport, PptValidationError> {
    let mut checks = Vec::new();
    checks
        .try_reserve_exact(CHECK_IDS.len())
        .map_err(|_error| PptValidationError::Allocation("PPT validation checks"))?;
    for (id, status) in CHECK_IDS.iter().zip(statuses) {
        checks.push(ValidationCheck::new(check_id(id, limits)?, status));
    }
    ValidateReport::try_new(checks, issues, limits).map_err(Into::into)
}

fn check_id(id: &str, limits: ValidationLimits) -> Result<CheckCapabilityId, PptValidationError> {
    CheckCapabilityId::try_new(id, limits).map_err(Into::into)
}

fn require_version(
    source: &dyn ReadAt,
    expected: litchi_core::SourceVersion,
) -> Result<(), PptValidationError> {
    let observed = source
        .version()
        .map_err(|error| PptValidationError::Ingress(OleError::Io(error)))?;
    if observed != expected {
        return Err(PptValidationError::Ingress(OleError::SourceChanged {
            expected,
            observed,
        }));
    }
    Ok(())
}
