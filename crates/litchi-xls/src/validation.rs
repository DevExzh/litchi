//! Bounded, non-mutating validation of legacy XLS ingress.
//!
//! This validator deliberately stops at the boundary that can be established
//! without credentials or a format-specific repair policy.  It validates the
//! CFB container, the Workbook/Book stream frame grammar, workbook-global
//! `BoundSheet8` ownership, and the inert presence of protection, password
//! encryption, signatures, DRM, and external-link metadata.  It never changes
//! the positional source, resolves external targets, verifies certificates, or
//! decrypts a Workbook stream.

use std::{
    collections::{HashMap, HashSet},
    error::Error as StdError,
    fmt,
    hash::Hash,
    sync::Arc,
};

use litchi_biff::{Limits as BiffLimits, Records};
use litchi_cfb::{OleError, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceValue, IssueEvidence,
    IssueLocation, IssueSeverity, ReadAt, RepairAvailability, ValidateReport, ValidationCheck,
    ValidationIssue, ValidationLimits, ValidationReportError,
};

use crate::{
    encryption::{FilePassKind, inspect_filepass},
    external_link::{
        CONTINUE_RECORD_TYPE, CRN_RECORD_TYPE, EXTERN_NAME_RECORD_TYPE, EXTERN_SHEET_RECORD_TYPE,
        SUP_BOOK_RECORD_TYPE, XCT_RECORD_TYPE,
    },
    protection::{
        FILESHARING_TYPE, OBJECTPROTECT_TYPE, PASSWORD_TYPE, PROT4REV_TYPE, PROT4REVPASS_TYPE,
        PROTECT_TYPE, SCENPROTECT_TYPE, WINPROTECT_TYPE, WRITEPROTECT_TYPE,
    },
    records::{BiffVersion, Encoding, SheetType},
};

const CFB: &str = "xls.cfb.ingress";
const WORKBOOK_STREAM: &str = "xls.workbook.stream";
const BIFF: &str = "xls.biff.parse";
const WORKSHEETS: &str = "xls.worksheet.inventory";
const PROTECTION: &str = "xls.protection.presence";
const ENCRYPTION: &str = "xls.encryption.presence";
const SIGNATURE: &str = "xls.signature.presence";
const DRM: &str = "xls.drm.presence";
const EXTERNAL: &str = "xls.external_reference.presence";

const CHECK_IDS: [&str; 9] = [
    CFB,
    WORKBOOK_STREAM,
    BIFF,
    WORKSHEETS,
    PROTECTION,
    ENCRYPTION,
    SIGNATURE,
    DRM,
    EXTERNAL,
];

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CODEPAGE: u16 = 0x0042;
const BOUNDSHEET8: u16 = 0x0085;
const MAX_RECORD_BYTES: usize = litchi_biff::MAX_RECORD_BYTES;
const WORKBOOK_BOF_TYPE: u16 = 0x0005;
const WORKSHEET_BOF_TYPE: u16 = 0x0010;
const CHART_BOF_TYPE: u16 = 0x0020;
const MACRO_SHEET_BOF_TYPE: u16 = 0x0040;
const VB_MODULE_BOF_TYPE: u16 = 0x0100;
const MAX_EXTERNAL_BOOKS: usize = 1024;
const MAX_EXTERNAL_SHEETS: usize = 256;
const MAX_CACHED_CELLS: usize = 65_536;
const MAX_DDE_OLE_VALUES: usize = 65_536;
const MAX_RETAINED_ISSUES: usize = 12;

/// Finite input, BIFF, directory, semantic, and report bounds for one XLS
/// validation pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsValidationLimits {
    max_input_bytes: u64,
    max_workbook_stream_bytes: u64,
    max_biff_records: usize,
    max_worksheets: usize,
    max_external_records: usize,
    max_directory_entries: usize,
    report: ValidationLimits,
}

impl XlsValidationLimits {
    /// Creates explicit finite XLS validation limits.
    #[must_use]
    #[allow(
        clippy::too_many_arguments,
        reason = "each independent untrusted-resource bound is explicit"
    )]
    pub const fn new(
        max_input_bytes: u64,
        max_workbook_stream_bytes: u64,
        max_biff_records: usize,
        max_worksheets: usize,
        max_external_records: usize,
        max_directory_entries: usize,
        report: ValidationLimits,
    ) -> Self {
        Self {
            max_input_bytes,
            max_workbook_stream_bytes,
            max_biff_records,
            max_worksheets,
            max_external_records,
            max_directory_entries,
            report,
        }
    }

    /// Maximum physical CFB source length inspected.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum materialized Workbook/Book stream length.
    #[must_use]
    pub const fn max_workbook_stream_bytes(self) -> u64 {
        self.max_workbook_stream_bytes
    }

    /// Maximum BIFF frames traversed in one Workbook stream.
    #[must_use]
    pub const fn max_biff_records(self) -> usize {
        self.max_biff_records
    }

    /// Maximum `BoundSheet8` entries retained for owner validation.
    #[must_use]
    pub const fn max_worksheets(self) -> usize {
        self.max_worksheets
    }

    /// Maximum external-link frames handed to the semantic collector.
    #[must_use]
    pub const fn max_external_records(self) -> usize {
        self.max_external_records
    }

    /// Maximum CFB directory entries inspected for security-presence markers.
    #[must_use]
    pub const fn max_directory_entries(self) -> usize {
        self.max_directory_entries
    }

    /// Bounds for the retained format-neutral report.
    #[must_use]
    pub const fn report(self) -> ValidationLimits {
        self.report
    }
}

impl Default for XlsValidationLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: SharedOleFileLimits::MAX_INPUT_BYTES,
            max_workbook_stream_bytes: 128 * 1024 * 1024,
            max_biff_records: 1_000_000,
            max_worksheets: 65_535,
            max_external_records: 65_536,
            max_directory_entries: 100_000,
            report: ValidationLimits::default(),
        }
    }
}

/// Failure to perform or retain an XLS validation report.
#[derive(Debug)]
#[non_exhaustive]
pub enum XlsValidationError {
    /// CFB positional ingress or source stability failed.
    Ingress(OleError),
    /// A bounded validator-owned allocation failed.
    Allocation(&'static str),
    /// The shared bounded report rejected the retained result.
    Report(ValidationReportError),
}

impl fmt::Display for XlsValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ingress(error) => write!(formatter, "XLS validation ingress failed: {error}"),
            Self::Allocation(resource) => {
                write!(
                    formatter,
                    "allocation failed while validating XLS {resource}"
                )
            },
            Self::Report(error) => write!(formatter, "XLS validation report failed: {error}"),
        }
    }
}

impl StdError for XlsValidationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Ingress(error) => Some(error),
            Self::Allocation(_) => None,
            Self::Report(error) => Some(error),
        }
    }
}

impl From<ValidationReportError> for XlsValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

impl From<OleError> for XlsValidationError {
    fn from(error: OleError) -> Self {
        Self::Ingress(error)
    }
}

/// Validate a positional XLS source under default finite bounds.
///
/// The source is opened through the canonical CFB ingress validator.  Only
/// bounded CFB metadata and the Workbook/Book stream are inspected; no bytes
/// are written back and no external target or signature trust decision is
/// attempted.
///
/// # Errors
///
/// Returns an error only for source I/O/version changes, bounded allocation
/// failure, or failure to retain the report. Definite format rejection and a
/// configured resource ceiling are represented in the returned report.
pub fn validate_source(source: Arc<dyn ReadAt>) -> Result<ValidateReport, XlsValidationError> {
    validate_source_with_limits(source, XlsValidationLimits::default())
}

/// Validate a positional XLS source under explicit finite bounds.
///
/// A source or stream ceiling produces a blocked capability. A definite CFB or
/// BIFF grammar rejection produces a completed check with a content-free error
/// issue; protection and external-link metadata remain blocked when their
/// presence cannot be established safely. Encrypted semantic regions stop at
/// the encryption dependency because credentials are intentionally not accepted
/// by this presence-only API.
///
/// # Errors
///
/// Returns an error for source I/O/version changes, bounded allocation
/// failure, or bounded report-construction failure.
pub fn validate_source_with_limits(
    source: Arc<dyn ReadAt>,
    limits: XlsValidationLimits,
) -> Result<ValidateReport, XlsValidationError> {
    let expected_version = source
        .version()
        .map_err(|error| XlsValidationError::Ingress(OleError::Io(error)))?;
    let source_size = source
        .len()
        .map_err(|error| XlsValidationError::Ingress(OleError::Io(error)))?;
    require_version(source.as_ref(), expected_version)?;

    let input_ceiling = limits
        .max_input_bytes
        .min(SharedOleFileLimits::MAX_INPUT_BYTES);
    if limits.max_input_bytes == 0 || source_size > input_ceiling {
        let statuses = blocked_ingress_statuses(
            limits.report,
            "source exceeds the configured XLS validation input ceiling",
        )?;
        let report = finish_report(statuses, Vec::new(), limits.report)?;
        require_version(source.as_ref(), expected_version)?;
        return Ok(report);
    }

    let cfb_limits =
        SharedOleFileLimits::new(input_ceiling).map_err(XlsValidationError::Ingress)?;
    let shared = match SharedOleFile::open_with_limits(source.clone(), cfb_limits) {
        Ok(shared) => shared,
        Err(error) if is_structural_cfb_error(&error) => {
            let statuses = cfb_rejection_statuses(limits.report)?;
            let mut issues = Vec::new();
            try_reserve(&mut issues, 1, "CFB rejection report")?;
            try_push(
                &mut issues,
                cfb_issue(&error, source_size, limits.report)?,
                "CFB rejection report",
            )?;
            let report = finish_report(statuses, issues, limits.report)?;
            require_version(source.as_ref(), expected_version)?;
            return Ok(report);
        },
        Err(error) => return Err(XlsValidationError::Ingress(error)),
    };

    let (markers, directory_limited) = inspect_directory(&shared, limits)?;
    let workbook = locate_workbook_stream(&shared)?;

    let mut issues = Vec::new();
    try_reserve(&mut issues, MAX_RETAINED_ISSUES, "XLS validation issues")?;

    let mut statuses = std::array::from_fn(|_| CheckStatus::Complete);
    statuses[5] = presence_status(markers.encryption_count, directory_limited, limits.report)?;
    statuses[6] = presence_status(markers.signature_count, directory_limited, limits.report)?;
    statuses[7] = presence_status(markers.drm_count, directory_limited, limits.report)?;

    if !directory_limited {
        push_presence_issue(
            &mut issues,
            SIGNATURE,
            "xls.signature.infrastructure_present",
            IssueSeverity::Info,
            "CFB signature infrastructure is present; cryptographic validity and certificate trust were not checked.",
            markers.signature_count,
            "signature-infrastructure",
            limits.report,
        )?;
        push_presence_issue(
            &mut issues,
            DRM,
            "xls.drm.storage_present",
            IssueSeverity::Warning,
            "CFB DRM storage is present; protected-content access was not attempted.",
            markers.drm_count,
            "drm-storage",
            limits.report,
        )?;
    }

    let WorkbookLocation::Stream {
        name: workbook_name,
        size: workbook_size,
    } = workbook
    else {
        if let WorkbookLocation::Invalid = workbook {
            statuses[1] = CheckStatus::Complete;
            try_push(
                &mut issues,
                simple_issue(
                    WORKBOOK_STREAM,
                    "xls.workbook.stream_ambiguous",
                    IssueSeverity::Error,
                    "The CFB directory contains ambiguous or non-stream Workbook/Book entries.",
                    Some("compound-file"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        } else {
            statuses[1] = CheckStatus::Complete;
            try_push(
                &mut issues,
                simple_issue(
                    WORKBOOK_STREAM,
                    "xls.workbook.stream_missing",
                    IssueSeverity::Error,
                    "The CFB container has neither a Workbook nor a Book stream.",
                    Some("compound-file"),
                    Some("Workbook"),
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
        statuses[2] = CheckStatus::blocked(
            "Workbook stream resolution did not produce one unambiguous BIFF stream",
            limits.report,
        )?;
        statuses[3] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        statuses[4] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        statuses[8] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        let report = finish_report(statuses, issues, limits.report)?;
        shared
            .source_version()
            .map_err(XlsValidationError::Ingress)?;
        return Ok(report);
    };

    if workbook_size > limits.max_workbook_stream_bytes {
        statuses[1] = CheckStatus::blocked(
            "Workbook stream exceeds the configured XLS validation byte ceiling",
            limits.report,
        )?;
        statuses[2] = CheckStatus::stopped_by(check_id(WORKBOOK_STREAM, limits.report)?);
        statuses[3] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        statuses[4] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        statuses[8] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        let report = finish_report(statuses, issues, limits.report)?;
        shared
            .source_version()
            .map_err(XlsValidationError::Ingress)?;
        return Ok(report);
    }

    let workbook_data = shared
        .open_stream(&[workbook_name])
        .map_err(XlsValidationError::Ingress)?;
    let analysis = analyze_workbook(&workbook_data, markers.encryption_count != 0, limits)?;

    let encryption_count = markers
        .encryption_count
        .checked_add(analysis.filepass_count)
        .ok_or(XlsValidationError::Allocation("encryption presence count"))?;
    statuses[5] = presence_status(encryption_count, directory_limited, limits.report)?;
    if !directory_limited {
        push_presence_issue(
            &mut issues,
            ENCRYPTION,
            "xls.encryption.password_to_open_present",
            IssueSeverity::Warning,
            "Password-to-open or encrypted-container metadata is present; this validator did not decrypt the Workbook stream.",
            encryption_count,
            "encryption-metadata",
            limits.report,
        )?;
    }
    if analysis.filepass_invalid {
        try_push(
            &mut issues,
            simple_issue(
                ENCRYPTION,
                "xls.encryption.filepass_invalid",
                IssueSeverity::Error,
                "A FILEPASS record is present but its bounded encryption header is malformed.",
                Some("Workbook"),
                None,
                None,
                limits.report,
            )?,
            "XLS validation issues",
        )?;
    }
    if analysis.filepass_unsupported {
        try_push(
            &mut issues,
            simple_issue(
                ENCRYPTION,
                "xls.encryption.unsupported",
                IssueSeverity::Error,
                "The Workbook declares a password-encryption family not supported by this presence-only validator.",
                Some("Workbook"),
                None,
                None,
                limits.report,
            )?,
            "XLS validation issues",
        )?;
    }
    if analysis.filepass_placement_invalid {
        try_push(
            &mut issues,
            simple_issue(
                ENCRYPTION,
                "xls.encryption.filepass_placement",
                IssueSeverity::Error,
                "A FILEPASS record appears outside the bounded workbook-global region after the global BOF.",
                Some("Workbook"),
                None,
                None,
                limits.report,
            )?,
            "XLS validation issues",
        )?;
    }

    if analysis.biff_limit {
        statuses[2] = CheckStatus::blocked(
            "BIFF record traversal reached the configured record ceiling",
            limits.report,
        )?;
    } else if analysis.worksheet_limit {
        statuses[2] = CheckStatus::blocked(
            "BIFF ownership traversal reached the configured worksheet ceiling",
            limits.report,
        )?;
    } else if analysis.encrypted {
        statuses[2] = CheckStatus::blocked(
            "password-to-open encryption prevents clear BIFF semantic inspection",
            limits.report,
        )?;
    } else {
        statuses[2] = CheckStatus::Complete;
        if analysis.biff_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    BIFF,
                    "xls.biff.invalid",
                    IssueSeverity::Error,
                    "The Workbook stream contains malformed or incomplete BIFF ownership grammar.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
    }

    let biff_blocked = matches!(statuses[2], CheckStatus::Blocked { .. });
    if biff_blocked {
        statuses[3] = CheckStatus::stopped_by(check_id(BIFF, limits.report)?);
        statuses[4] = if analysis.protection_invalid {
            CheckStatus::blocked(
                "malformed protection metadata was proven before BIFF inspection stopped",
                limits.report,
            )?
        } else {
            CheckStatus::stopped_by(check_id(BIFF, limits.report)?)
        };
        statuses[8] = if analysis.external_invalid {
            CheckStatus::blocked(
                "malformed external-link metadata was proven before BIFF inspection stopped",
                limits.report,
            )?
        } else {
            CheckStatus::stopped_by(check_id(BIFF, limits.report)?)
        };
        if analysis.protection_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    PROTECTION,
                    "xls.protection.invalid",
                    IssueSeverity::Error,
                    "Protection records are malformed, duplicated, or out of their required BIFF order.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
        if analysis.external_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    EXTERNAL,
                    "xls.external_reference.invalid",
                    IssueSeverity::Error,
                    "External-link records are malformed or have an invalid owner relationship.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
    } else if analysis.biff_invalid {
        statuses[3] = CheckStatus::blocked(
            "BIFF grammar rejection prevented worksheet inventory",
            limits.report,
        )?;
        statuses[4] = CheckStatus::blocked(
            "BIFF grammar rejection prevented protection inspection",
            limits.report,
        )?;
        statuses[8] = CheckStatus::blocked(
            "BIFF grammar rejection prevented external-link inspection",
            limits.report,
        )?;
    } else {
        statuses[3] = if analysis.worksheet_limit {
            CheckStatus::blocked(
                "worksheet inventory exceeded the configured sheet bound",
                limits.report,
            )?
        } else {
            CheckStatus::Complete
        };
        if analysis.worksheet_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    WORKSHEETS,
                    "xls.worksheet.invalid",
                    IssueSeverity::Error,
                    "A BoundSheet8 entry did not resolve to one bounded BOF-to-EOF worksheet owner.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
        statuses[4] = if analysis.protection_invalid {
            CheckStatus::blocked(
                "malformed protection metadata prevented a safe presence result",
                limits.report,
            )?
        } else if analysis.protection_seen {
            CheckStatus::Complete
        } else {
            CheckStatus::NotApplicable
        };
        if analysis.protection_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    PROTECTION,
                    "xls.protection.invalid",
                    IssueSeverity::Error,
                    "Protection records are malformed, duplicated, or out of their required BIFF order.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
        statuses[8] = if analysis.external_limit {
            CheckStatus::blocked(
                "external-link metadata exceeded the configured record bound",
                limits.report,
            )?
        } else if analysis.external_invalid {
            CheckStatus::blocked(
                "malformed external-link metadata prevented a safe presence result",
                limits.report,
            )?
        } else if analysis.external_present {
            CheckStatus::Complete
        } else {
            CheckStatus::NotApplicable
        };
        if analysis.external_invalid {
            try_push(
                &mut issues,
                simple_issue(
                    EXTERNAL,
                    "xls.external_reference.invalid",
                    IssueSeverity::Error,
                    "External-link records are malformed or have an invalid owner relationship.",
                    Some("Workbook"),
                    None,
                    None,
                    limits.report,
                )?,
                "XLS validation issues",
            )?;
        }
        if analysis.external_present {
            push_presence_issue(
                &mut issues,
                EXTERNAL,
                "xls.external_reference.present",
                IssueSeverity::Info,
                "External-reference metadata is present; targets were not resolved or fetched.",
                analysis.external_count,
                "external-reference-records",
                limits.report,
            )?;
        }
    }

    let report = finish_report(statuses, issues, limits.report)?;
    shared
        .source_version()
        .map_err(XlsValidationError::Ingress)?;
    Ok(report)
}

#[derive(Debug, Default, Clone, Copy)]
struct MarkerCounts {
    encryption_count: u64,
    signature_count: u64,
    drm_count: u64,
}

#[derive(Debug, Default)]
struct Analysis {
    filepass_count: u64,
    filepass_invalid: bool,
    filepass_unsupported: bool,
    filepass_placement_invalid: bool,
    stopped_at_filepass: bool,
    encrypted: bool,
    biff_invalid: bool,
    biff_limit: bool,
    worksheet_invalid: bool,
    worksheet_limit: bool,
    protection_seen: bool,
    protection_invalid: bool,
    external_present: bool,
    external_count: u64,
    external_invalid: bool,
    external_limit: bool,
}

#[derive(Debug, Clone, Copy)]
struct SheetObservation {
    sheet_type: SheetType,
    expected_dt: u16,
    started: bool,
    ended: bool,
}

struct ActiveSheet {
    index: usize,
    protection: SheetProtectionState,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BofInfo {
    version: BiffVersion,
    dt: u16,
}

fn parse_bof(data: &[u8]) -> Option<BofInfo> {
    let version = BiffVersion::from_bof_version(read_u16(data, 0)?)?;
    let expected_len = match version {
        BiffVersion::Biff2 => 4,
        BiffVersion::Biff3 | BiffVersion::Biff4 | BiffVersion::Biff5 => 8,
        BiffVersion::Biff8 => 16,
    };
    if data.len() != expected_len {
        return None;
    }
    let dt = read_u16(data, 2)?;
    if !matches!(
        dt,
        WORKBOOK_BOF_TYPE
            | WORKSHEET_BOF_TYPE
            | CHART_BOF_TYPE
            | MACRO_SHEET_BOF_TYPE
            | VB_MODULE_BOF_TYPE
    ) {
        return None;
    }
    Some(BofInfo { version, dt })
}

fn expected_sheet_bof_type(sheet_type: SheetType) -> u16 {
    match sheet_type {
        SheetType::WorkSheet => WORKSHEET_BOF_TYPE,
        SheetType::MacroSheet => MACRO_SHEET_BOF_TYPE,
        SheetType::ChartSheet => CHART_BOF_TYPE,
        SheetType::VBModule => VB_MODULE_BOF_TYPE,
    }
}

fn parse_bound_sheet(data: &[u8]) -> Option<(u32, SheetType)> {
    if data.len() < 8 {
        return None;
    }
    if data[4] & !0x03 != 0 || data[4] & 0x03 == 0x03 {
        return None;
    }
    let sheet_type = match data[5] {
        0x00 => SheetType::WorkSheet,
        0x01 => SheetType::MacroSheet,
        0x02 => SheetType::ChartSheet,
        0x06 => SheetType::VBModule,
        _ => return None,
    };
    let character_count = usize::from(data[6]);
    if !(1..=31).contains(&character_count) || data[7] & !1 != 0 {
        return None;
    }
    let wide = data[7] == 1;
    let name_bytes = character_count.checked_mul(if wide { 2 } else { 1 })?;
    if data.len() != 8usize.checked_add(name_bytes)? {
        return None;
    }
    let name = data.get(8..)?;
    if wide {
        let mut pairs = name.chunks_exact(2);
        let mut units = pairs
            .by_ref()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        let mut first = None;
        let mut last = None;
        for decoded in std::char::decode_utf16(units.by_ref()) {
            let character = decoded.ok()?;
            if first.is_none() {
                first = Some(character);
            }
            last = Some(character);
            if matches!(
                character,
                '\0' | '\u{0003}' | ':' | '\\' | '*' | '?' | '/' | '[' | ']'
            ) {
                return None;
            }
        }
        if !pairs.remainder().is_empty() || first == Some('\'') || last == Some('\'') {
            return None;
        }
    } else {
        let first = *name.first()?;
        let last = *name.last()?;
        if first == b'\''
            || last == b'\''
            || name.iter().any(|character| {
                matches!(
                    *character,
                    0 | 3 | b':' | b'\\' | b'*' | b'?' | b'/' | b'[' | b']'
                )
            })
        {
            return None;
        }
    }
    Some((
        u32::from_le_bytes([data[0], data[1], data[2], data[3]]),
        sheet_type,
    ))
}

fn read_u16(data: &[u8], offset: usize) -> Option<u16> {
    let bytes = data.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn inspect_filepass_record(
    output: &mut Analysis,
    payload: &[u8],
) -> Result<(), XlsValidationError> {
    output.filepass_count = output
        .filepass_count
        .checked_add(1)
        .ok_or(XlsValidationError::Allocation("FILEPASS count"))?;
    if output.filepass_count > 1 {
        output.filepass_invalid = true;
        output.biff_invalid = true;
    }
    match inspect_filepass(payload) {
        Ok(FilePassKind::Unsupported(_)) => output.filepass_unsupported = true,
        Ok(FilePassKind::Xor | FilePassKind::BinaryRc4 | FilePassKind::CryptoApi) => {},
        Err(_) => {
            output.filepass_invalid = true;
            output.biff_invalid = true;
        },
    }
    output.encrypted = true;
    Ok(())
}

#[derive(Debug, Default, Clone, Copy)]
struct WorkbookProtectionState {
    structure: bool,
    structure_value: bool,
    password: bool,
    password_set: bool,
    windows: bool,
    windows_value: bool,
    revisions: bool,
    revisions_value: bool,
    revision_password: bool,
    revision_password_set: bool,
    write_protect: bool,
    file_sharing: bool,
    previous: Option<u16>,
}

impl WorkbookProtectionState {
    fn feed(&mut self, record_type: u16, data: &[u8]) -> bool {
        let valid = match record_type {
            PROTECT_TYPE => {
                if self.structure {
                    false
                } else if let Some(value) = protection_bool(data) {
                    self.structure = true;
                    self.structure_value = value;
                    true
                } else {
                    false
                }
            },
            PASSWORD_TYPE => {
                if self.password {
                    false
                } else if let Some(value) = protection_word(data) {
                    self.password = true;
                    self.password_set = value != 0;
                    true
                } else {
                    false
                }
            },
            WINPROTECT_TYPE => {
                if self.windows {
                    false
                } else if let Some(value) = protection_bool(data) {
                    self.windows = true;
                    self.windows_value = value;
                    true
                } else {
                    false
                }
            },
            PROT4REV_TYPE => {
                if self.revisions {
                    false
                } else if let Some(value) = protection_bool(data) {
                    self.revisions = true;
                    self.revisions_value = value;
                    true
                } else {
                    false
                }
            },
            PROT4REVPASS_TYPE => {
                if self.revision_password || self.previous != Some(PROT4REV_TYPE) {
                    false
                } else if let Some(value) = protection_word(data) {
                    self.revision_password = true;
                    self.revision_password_set = value != 0;
                    self.revisions_value || value == 0
                } else {
                    false
                }
            },
            WRITEPROTECT_TYPE => {
                if self.write_protect || !data.is_empty() {
                    false
                } else {
                    self.write_protect = true;
                    true
                }
            },
            FILESHARING_TYPE => {
                if self.file_sharing || !valid_file_sharing(data) {
                    false
                } else {
                    self.file_sharing = true;
                    true
                }
            },
            _ => true,
        };
        self.previous = Some(record_type);
        valid
    }

    fn finish(self) -> bool {
        self.revisions == self.revision_password
    }

    fn is_present(self) -> bool {
        self.structure_value
            || self.password_set
            || self.windows_value
            || self.revisions_value
            || self.revision_password_set
            || self.write_protect
            || self.file_sharing
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct SheetProtectionState {
    protect: bool,
    objects: bool,
    scenarios: bool,
    password: bool,
    present: bool,
}

impl SheetProtectionState {
    fn feed(&mut self, record_type: u16, data: &[u8]) -> bool {
        match record_type {
            PROTECT_TYPE => {
                if self.protect || protection_bool(data) != Some(true) {
                    false
                } else {
                    self.protect = true;
                    self.present = true;
                    true
                }
            },
            OBJECTPROTECT_TYPE => {
                if self.objects || protection_bool(data) != Some(true) {
                    false
                } else {
                    self.objects = true;
                    self.present = true;
                    true
                }
            },
            SCENPROTECT_TYPE => {
                if self.scenarios {
                    false
                } else if let Some(value) = protection_bool(data) {
                    self.scenarios = true;
                    self.present |= value;
                    true
                } else {
                    false
                }
            },
            PASSWORD_TYPE => {
                if self.password {
                    false
                } else if let Some(value) = protection_word(data) {
                    self.password = true;
                    if value == 0 {
                        false
                    } else {
                        self.present = true;
                        true
                    }
                } else {
                    false
                }
            },
            _ => true,
        }
    }

    fn finish(self) -> bool {
        self.protect || (!self.objects && !self.scenarios && !self.password)
    }

    const fn is_present(self) -> bool {
        self.present
    }
}

fn protection_word(data: &[u8]) -> Option<u16> {
    read_u16(data, 0).filter(|_| data.len() == 2)
}

fn protection_bool(data: &[u8]) -> Option<bool> {
    match protection_word(data)? {
        0 => Some(false),
        1 => Some(true),
        _ => None,
    }
}

fn valid_file_sharing(data: &[u8]) -> bool {
    if protection_bool(data.get(..2).unwrap_or_default()).is_none() {
        return false;
    }
    let Some(write_password) = protection_word(data.get(2..4).unwrap_or_default()) else {
        return false;
    };
    let Some(cch_or_marker) = protection_word(data.get(4..6).unwrap_or_default()) else {
        return false;
    };
    if write_password == 0 {
        return cch_or_marker == 0 && data.len() == 6;
    }
    let cch = usize::from(cch_or_marker);
    if cch > 54 || data.len() < 7 {
        return false;
    }
    let flags = data[6];
    if flags & !1 != 0 {
        return false;
    }
    let expected = 7 + cch * if flags == 1 { 2 } else { 1 };
    if data.len() != expected {
        return false;
    }
    if flags == 1 {
        valid_utf16le(data.get(7..).unwrap_or_default())
    } else {
        true
    }
}

fn valid_utf16le(data: &[u8]) -> bool {
    let mut units = data.chunks_exact(2);
    let valid = std::char::decode_utf16(
        units
            .by_ref()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
    )
    .all(|value| value.is_ok());
    valid && units.remainder().is_empty()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum ExternalBookKind {
    #[default]
    SelfReference,
    AddIn,
    External {
        sheets: usize,
    },
    SameSheet,
    Unused {
        sheets: usize,
    },
    DdeOrOle,
}

#[derive(Debug)]
struct ExternalPresenceState {
    books: [ExternalBookKind; MAX_EXTERNAL_BOOKS],
    book_count: usize,
    current_book: Option<usize>,
    names_allowed: bool,
    extern_sheet_seen: bool,
    closed: bool,
    pending_crn: usize,
    pending_dde_values: usize,
    cached_cells: usize,
    external_names: usize,
    sheet_references: usize,
    external_present: bool,
    max_self_sheet: Option<usize>,
    pending_xct: Option<(usize, usize)>,
    seen_xct: HashSet<(usize, usize)>,
    cache_cells: HashSet<(usize, usize, u16, u8)>,
}

impl Default for ExternalPresenceState {
    fn default() -> Self {
        Self {
            books: [ExternalBookKind::SelfReference; MAX_EXTERNAL_BOOKS],
            book_count: 0,
            current_book: None,
            names_allowed: false,
            extern_sheet_seen: false,
            closed: false,
            pending_crn: 0,
            pending_dde_values: 0,
            cached_cells: 0,
            external_names: 0,
            sheet_references: 0,
            external_present: false,
            max_self_sheet: None,
            pending_xct: None,
            seen_xct: HashSet::new(),
            cache_cells: HashSet::new(),
        }
    }
}

impl ExternalPresenceState {
    fn feed(&mut self, record_type: u16, data: &[u8]) -> Result<bool, XlsValidationError> {
        if self.pending_crn != 0 && record_type != CRN_RECORD_TYPE {
            return Ok(false);
        }
        if self.pending_dde_values != 0 && record_type != CONTINUE_RECORD_TYPE {
            return Ok(false);
        }
        if record_type == CONTINUE_RECORD_TYPE {
            if self.pending_dde_values == 0 {
                return Ok(self.current_book.is_none() || self.closed);
            }
            let Some((count, consumed)) = count_ser_ar_values(data, self.pending_dde_values) else {
                return Ok(false);
            };
            if count == 0 || consumed != data.len() || count > self.pending_dde_values {
                return Ok(false);
            }
            self.pending_dde_values -= count;
            return Ok(true);
        }

        let is_target = is_external_record(record_type);
        if !is_target {
            if self.current_book.is_some() && !self.closed {
                self.closed = true;
            }
            return Ok(true);
        }
        if self.closed {
            return Ok(false);
        }

        match record_type {
            SUP_BOOK_RECORD_TYPE => Ok(self.feed_sup_book(data)),
            EXTERN_NAME_RECORD_TYPE => Ok(self.feed_extern_name(data)),
            XCT_RECORD_TYPE => self.feed_xct(data),
            CRN_RECORD_TYPE => self.feed_crn(data),
            EXTERN_SHEET_RECORD_TYPE => Ok(self.feed_extern_sheet(data)),
            CONTINUE_RECORD_TYPE => unreachable!(),
            _ => Ok(false),
        }
    }

    fn feed_sup_book(&mut self, data: &[u8]) -> bool {
        if self.extern_sheet_seen || self.book_count >= MAX_EXTERNAL_BOOKS {
            return false;
        }
        let Some(kind) = parse_external_book(data) else {
            return false;
        };
        let index = self.book_count;
        self.books[index] = kind;
        self.book_count += 1;
        self.current_book = Some(index);
        self.names_allowed = true;
        self.external_present |= matches!(
            kind,
            ExternalBookKind::External { .. }
                | ExternalBookKind::AddIn
                | ExternalBookKind::DdeOrOle
        );
        true
    }

    fn feed_extern_name(&mut self, data: &[u8]) -> bool {
        let Some(book_index) = self.current_book else {
            return false;
        };
        if !self.names_allowed || book_index >= self.book_count {
            return false;
        }
        let Some(next_pending) = parse_external_name_presence(data, self.books[book_index]) else {
            return false;
        };
        self.pending_dde_values = next_pending;
        let Some(next_count) = self.external_names.checked_add(1) else {
            return false;
        };
        self.external_names = next_count;
        true
    }

    fn feed_xct(&mut self, data: &[u8]) -> Result<bool, XlsValidationError> {
        let Some(book_index) = self.current_book else {
            return Ok(false);
        };
        let Some(declared) = read_i16(data, 0) else {
            return Ok(false);
        };
        let Some(sheet_index) = read_u16(data, 2).map(usize::from) else {
            return Ok(false);
        };
        if data.len() != 4
            || declared == i16::MIN
            || book_index >= self.book_count
            || !matches!(self.books[book_index], ExternalBookKind::External { .. })
            || !matches!(self.books[book_index], ExternalBookKind::External { sheets } if sheet_index < sheets)
        {
            return Ok(false);
        }
        let owner = (book_index, sheet_index);
        if self.seen_xct.len() >= MAX_CACHED_CELLS || self.seen_xct.contains(&owner) {
            return Ok(false);
        }
        try_hash_set_reserve(&mut self.seen_xct, 1, "external-link XCT ownership")?;
        self.seen_xct.insert(owner);
        self.pending_xct = Some(owner);
        self.pending_crn = usize::from(declared.unsigned_abs());
        self.names_allowed = false;
        Ok(true)
    }

    fn feed_crn(&mut self, data: &[u8]) -> Result<bool, XlsValidationError> {
        if self.pending_crn == 0 || data.len() < 4 {
            return Ok(false);
        }
        let Some(last_column) = data.first().copied() else {
            return Ok(false);
        };
        let first_column = data[1];
        if last_column < first_column {
            return Ok(false);
        }
        let values = usize::from(last_column - first_column) + 1;
        let Some((count, consumed)) = count_ser_ar_values(&data[4..], values) else {
            return Ok(false);
        };
        if count != values || consumed.checked_add(4) != Some(data.len()) {
            return Ok(false);
        }
        let Some(owner) = self.pending_xct else {
            return Ok(false);
        };
        let Some(next_cached_cells) = self.cached_cells.checked_add(count) else {
            return Ok(false);
        };
        if next_cached_cells > MAX_CACHED_CELLS {
            return Ok(false);
        }
        let row = u16::from_le_bytes([data[2], data[3]]);
        for column in first_column..=last_column {
            if self.cache_cells.contains(&(owner.0, owner.1, row, column)) {
                return Ok(false);
            }
        }
        try_hash_set_reserve(
            &mut self.cache_cells,
            count,
            "external-link cached-cell ownership",
        )?;
        for column in first_column..=last_column {
            self.cache_cells.insert((owner.0, owner.1, row, column));
        }
        self.cached_cells = next_cached_cells;
        self.pending_crn -= 1;
        if self.pending_crn == 0 {
            self.pending_xct = None;
        }
        Ok(true)
    }

    fn feed_extern_sheet(&mut self, data: &[u8]) -> bool {
        let Some(count) = read_u16(data, 0).map(usize::from) else {
            return false;
        };
        let Some(expected) = count.checked_mul(6).and_then(|value| value.checked_add(2)) else {
            return false;
        };
        if data.len() != expected {
            return false;
        }
        let mut offset = 2;
        for _ in 0..count {
            let Some(book_index) = read_u16(data, offset).map(usize::from) else {
                return false;
            };
            let Some(first) = read_i16(data, offset + 2) else {
                return false;
            };
            let Some(last) = read_i16(data, offset + 4) else {
                return false;
            };
            if book_index >= self.book_count
                || !self.valid_scope(self.books[book_index], first, last)
            {
                return false;
            }
            offset += 6;
        }
        self.extern_sheet_seen = true;
        self.names_allowed = false;
        self.sheet_references = self.sheet_references.saturating_add(count);
        self.sheet_references <= 1_370
    }

    fn valid_scope(&mut self, book: ExternalBookKind, first: i16, last: i16) -> bool {
        if first == -2 {
            return last == -2;
        }
        if matches!(
            book,
            ExternalBookKind::AddIn | ExternalBookKind::SameSheet | ExternalBookKind::DdeOrOle
        ) {
            return false;
        }
        if first == -1 || last == -1 {
            return first >= -1 && last >= -1;
        }
        if first < 0 || last < first {
            return false;
        }
        let first = usize::try_from(first).ok();
        let last = usize::try_from(last).ok();
        let (Some(first), Some(last)) = (first, last) else {
            return false;
        };
        match book {
            ExternalBookKind::SelfReference => {
                self.max_self_sheet = Some(self.max_self_sheet.unwrap_or(0).max(last));
                first <= last
            },
            ExternalBookKind::External { sheets } | ExternalBookKind::Unused { sheets } => {
                last < sheets
            },
            ExternalBookKind::AddIn | ExternalBookKind::SameSheet | ExternalBookKind::DdeOrOle => {
                false
            },
        }
    }

    fn finish(&self, internal_sheet_count: usize) -> bool {
        self.pending_crn == 0
            && self.pending_dde_values == 0
            && self
                .max_self_sheet
                .is_none_or(|maximum| maximum < internal_sheet_count)
    }

    fn observed_count(&self) -> u64 {
        u64::try_from(self.book_count)
            .unwrap_or(u64::MAX)
            .saturating_add(u64::try_from(self.external_names).unwrap_or(u64::MAX))
            .saturating_add(u64::try_from(self.sheet_references).unwrap_or(u64::MAX))
    }
}

fn read_i16(data: &[u8], offset: usize) -> Option<i16> {
    read_u16(data, offset).map(|value| i16::from_le_bytes(value.to_le_bytes()))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum UnicodeMarker {
    Other,
    Nul,
    Space,
}

fn unicode_span(data: &[u8], offset: usize, count: usize) -> Option<(usize, UnicodeMarker)> {
    let flags = *data.get(offset)?;
    if flags & !1 != 0 {
        return None;
    }
    let wide = flags == 1;
    let byte_count = count.checked_mul(if wide { 2 } else { 1 })?;
    let start = offset.checked_add(1)?;
    let end = start.checked_add(byte_count)?;
    let encoded = data.get(start..end)?;
    if wide && !valid_utf16le(encoded) {
        return None;
    }
    let mut all_nul = count == 1;
    let mut all_space = count == 1;
    if wide {
        for pair in encoded.chunks_exact(2) {
            let value = u16::from_le_bytes([pair[0], pair[1]]);
            all_nul &= value == 0;
            all_space &= value == u16::from(b' ');
        }
    } else {
        for value in encoded {
            all_nul &= *value == 0;
            all_space &= *value == b' ';
        }
    }
    let marker = if all_nul {
        UnicodeMarker::Nul
    } else if all_space {
        UnicodeMarker::Space
    } else {
        UnicodeMarker::Other
    };
    Some((end, marker))
}

fn unicode_string_span(
    data: &[u8],
    offset: usize,
    max_count: usize,
) -> Option<(usize, UnicodeMarker)> {
    let count = usize::from(read_u16(data, offset)?);
    if count > max_count {
        return None;
    }
    unicode_span(data, offset.checked_add(2)?, count)
}

fn short_unicode_span(data: &[u8], offset: usize) -> Option<(usize, UnicodeMarker)> {
    unicode_span(
        data,
        offset.checked_add(1)?,
        usize::from(*data.get(offset)?),
    )
}

fn unicode_equals_ascii(data: &[u8], offset: usize, text: &[u8]) -> bool {
    if data.get(offset).copied().map(usize::from) != Some(text.len()) {
        return false;
    }
    let Some(flags) = data.get(offset + 1).copied() else {
        return false;
    };
    if flags != 0 {
        return false;
    }
    let Some(end) = offset
        .checked_add(2)
        .and_then(|value| value.checked_add(text.len()))
    else {
        return false;
    };
    data.get(offset + 2..end) == Some(text)
}

fn parse_external_book(data: &[u8]) -> Option<ExternalBookKind> {
    if !(4..=MAX_RECORD_BYTES).contains(&data.len()) {
        return None;
    }
    let sheet_count = usize::from(read_u16(data, 0)?);
    let cch = read_u16(data, 2)?;
    if data.len() == 4 {
        return match cch {
            0x0401 => Some(ExternalBookKind::SelfReference),
            0x3a01 if sheet_count == 1 => Some(ExternalBookKind::AddIn),
            _ => None,
        };
    }
    if !(1..=255).contains(&cch) {
        return None;
    }
    let (mut offset, path_marker) = unicode_span(data, 4, usize::from(cch))?;
    if path_marker == UnicodeMarker::Nul {
        return (sheet_count == 0 && offset == data.len()).then_some(ExternalBookKind::SameSheet);
    }
    if sheet_count == 0 {
        return (offset == data.len()).then_some(ExternalBookKind::DdeOrOle);
    }
    if sheet_count > MAX_EXTERNAL_SHEETS {
        return None;
    }
    let mut all_sheet_placeholders = true;
    for _ in 0..sheet_count {
        let sheet_name_length = usize::from(read_u16(data, offset)?);
        if !(1..=31).contains(&sheet_name_length) {
            return None;
        }
        let (next, marker) = unicode_string_span(data, offset, 31)?;
        if marker == UnicodeMarker::Nul {
            return None;
        }
        all_sheet_placeholders &= marker == UnicodeMarker::Space;
        offset = next;
    }
    if offset != data.len() {
        return None;
    }
    if path_marker == UnicodeMarker::Space {
        if cch == 1 {
            all_sheet_placeholders.then_some(ExternalBookKind::Unused {
                sheets: sheet_count,
            })
        } else {
            Some(ExternalBookKind::External {
                sheets: sheet_count,
            })
        }
    } else {
        Some(ExternalBookKind::External {
            sheets: sheet_count,
        })
    }
}

fn parse_external_name_presence(data: &[u8], book: ExternalBookKind) -> Option<usize> {
    if !(8..=MAX_RECORD_BYTES).contains(&data.len()) {
        return None;
    }
    let flags = read_u16(data, 0)?;
    if !valid_clipboard_format((flags >> 5) & 0x03ff) {
        return None;
    }
    let standard_document_name = flags & 0x0008 != 0;
    let ole_link = flags & 0x0010 != 0;
    let displayed_as_icon = flags & 0x8000 != 0;
    if standard_document_name && ole_link {
        return None;
    }
    let name = short_unicode_span(data, 6)?;
    let offset = name.0;
    match book {
        ExternalBookKind::External { sheets } => {
            if flags & !0x0001 != 0
                || read_u16(data, 4)? != 0
                || usize::from(read_u16(data, 2)?) > sheets
            {
                return None;
            }
            let formula_len = usize::from(read_u16(data, offset)?);
            let formula_start = offset.checked_add(2)?;
            let formula_end = formula_start.checked_add(formula_len)?;
            if formula_end != data.len()
                || data
                    .get(formula_start)
                    .is_some_and(|token| !matches!(token, 0x1c | 0x3a | 0x3b | 0x3c | 0x3d))
            {
                return None;
            }
            Some(0)
        },
        ExternalBookKind::AddIn => {
            if flags != 0 || read_u16(data, 2)? != 0 || read_u16(data, 4)? != 0 {
                return None;
            }
            let unused_len = usize::from(read_u16(data, offset)?);
            (offset.checked_add(2)?.checked_add(unused_len)? == data.len()).then_some(0)
        },
        ExternalBookKind::DdeOrOle => {
            let storage_id = read_u32(data, 2)?;
            if flags & 0x0001 != 0 {
                return None;
            }
            if standard_document_name {
                return (!displayed_as_icon
                    && storage_id == 0
                    && unicode_equals_ascii(data, 6, b"StdDocumentName")
                    && offset == data.len())
                .then_some(0);
            }
            if (!ole_link && storage_id != 0) || (displayed_as_icon && !ole_link) {
                return None;
            }
            if offset == data.len() {
                return Some(0);
            }
            let last_column = *data.get(offset)?;
            let last_row = usize::from(read_u16(data, offset.checked_add(1)?)?);
            let expected = (usize::from(last_column) + 1).checked_mul(last_row + 1)?;
            if expected > MAX_DDE_OLE_VALUES {
                return None;
            }
            let values_start = offset.checked_add(3)?;
            let (count, consumed) = count_ser_ar_values(data.get(values_start..)?, expected)?;
            if consumed != data.len().checked_sub(values_start)? || count > expected {
                return None;
            }
            Some(expected - count)
        },
        ExternalBookKind::SelfReference
        | ExternalBookKind::SameSheet
        | ExternalBookKind::Unused { .. } => None,
    }
}

fn valid_clipboard_format(value: u16) -> bool {
    let value = if value & 0x0200 != 0 {
        i32::from(value) - 1024
    } else {
        i32::from(value)
    };
    matches!(
        value,
        -1 | 0 | 2 | 5 | 6 | 7 | 8 | 9 | 16 | 20 | 30 | 36 | 44 | 63
    )
}

fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    let bytes = data.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn count_ser_ar_values(data: &[u8], maximum: usize) -> Option<(usize, usize)> {
    let mut offset = 0;
    let mut count = 0;
    while offset < data.len() {
        if count >= maximum {
            return None;
        }
        let tag = *data.get(offset)?;
        let next = match tag {
            0x00 | 0x01 => offset.checked_add(9)?,
            0x02 => unicode_string_span(data, offset.checked_add(1)?, 255)?.0,
            0x04 => {
                if *data.get(offset.checked_add(2)?)? != 0
                    || !matches!(*data.get(offset.checked_add(1)?)?, 0 | 1)
                {
                    return None;
                }
                offset.checked_add(9)?
            },
            0x10 => {
                if *data.get(offset.checked_add(2)?)? != 0
                    || !matches!(
                        *data.get(offset.checked_add(1)?)?,
                        0x00 | 0x07 | 0x0f | 0x17 | 0x1d | 0x24 | 0x2a | 0x2b
                    )
                {
                    return None;
                }
                offset.checked_add(9)?
            },
            _ => return None,
        };
        if next > data.len() {
            return None;
        }
        count += 1;
        offset = next;
    }
    Some((count, offset))
}

fn inspect_directory(
    shared: &SharedOleFile,
    limits: XlsValidationLimits,
) -> Result<(MarkerCounts, bool), XlsValidationError> {
    let mut counts = MarkerCounts::default();
    let mut entries = 0_usize;
    for entry in shared.directory_entries() {
        if entries >= limits.max_directory_entries {
            return Ok((counts, true));
        }
        entries = entries
            .checked_add(1)
            .ok_or(XlsValidationError::Allocation("CFB directory count"))?;
        if is_signature_component(&entry.name) {
            counts.signature_count = counts
                .signature_count
                .checked_add(1)
                .ok_or(XlsValidationError::Allocation("signature presence count"))?;
        }
        if is_encryption_component(&entry.name) {
            counts.encryption_count = counts
                .encryption_count
                .checked_add(1)
                .ok_or(XlsValidationError::Allocation("encryption presence count"))?;
        }
        if is_drm_component(&entry.name) {
            counts.drm_count = counts
                .drm_count
                .checked_add(1)
                .ok_or(XlsValidationError::Allocation("DRM presence count"))?;
        }
    }
    Ok((counts, false))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WorkbookLocation {
    Missing,
    Invalid,
    Stream { name: &'static str, size: u64 },
}

fn locate_workbook_stream(shared: &SharedOleFile) -> Result<WorkbookLocation, XlsValidationError> {
    // `directory_entries` is a flattened SID index and does not carry the
    // parent storage path. Resolve only the two root-level logical paths so a
    // nested `Workbook`/`Book` member cannot shadow the workbook stream.
    let workbook = logical_stream_len(shared, "Workbook")?;
    let book = logical_stream_len(shared, "Book")?;

    if matches!(workbook, LogicalEntry::NonStream)
        || matches!(book, LogicalEntry::NonStream)
        || matches!(workbook, LogicalEntry::Stream(_)) && matches!(book, LogicalEntry::Stream(_))
    {
        return Ok(WorkbookLocation::Invalid);
    }
    if let LogicalEntry::Stream(size) = workbook {
        return Ok(WorkbookLocation::Stream {
            name: "Workbook",
            size,
        });
    }
    if let LogicalEntry::Stream(size) = book {
        return Ok(WorkbookLocation::Stream { name: "Book", size });
    }
    Ok(WorkbookLocation::Missing)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LogicalEntry {
    Missing,
    NonStream,
    Stream(u64),
}

fn logical_stream_len(
    shared: &SharedOleFile,
    name: &str,
) -> Result<LogicalEntry, XlsValidationError> {
    match shared.stream_len(&[name]) {
        Ok(size) => Ok(LogicalEntry::Stream(size)),
        Err(OleError::StreamNotFound) => Ok(LogicalEntry::Missing),
        Err(OleError::InvalidFormat(_)) => Ok(LogicalEntry::NonStream),
        Err(error) => Err(XlsValidationError::Ingress(error)),
    }
}

fn analyze_workbook(
    data: &[u8],
    container_encrypted: bool,
    limits: XlsValidationLimits,
) -> Result<Analysis, XlsValidationError> {
    let mut output = Analysis {
        encrypted: container_encrypted,
        ..Analysis::default()
    };
    if container_encrypted {
        return Ok(output);
    }

    let biff_limits = BiffLimits {
        max_records: limits.max_biff_records,
        max_record_bytes: MAX_RECORD_BYTES,
        max_input_bytes: data.len(),
        max_output_bytes: data.len(),
    };
    let records = match Records::with_limits(data, biff_limits) {
        Ok(records) => records,
        Err(litchi_biff::Error::LimitExceeded { .. }) => {
            return Ok(Analysis {
                biff_limit: true,
                encrypted: container_encrypted,
                ..Analysis::default()
            });
        },
        Err(litchi_biff::Error::Allocation { .. }) => {
            return Err(XlsValidationError::Allocation("BIFF frame traversal"));
        },
        Err(_) => {
            return Ok(Analysis {
                biff_invalid: true,
                encrypted: container_encrypted,
                ..Analysis::default()
            });
        },
    };

    let mut first_record = true;
    let mut globals_done = false;
    let mut saw_global_eof = false;
    let mut global_bof = None;
    let mut codepage_seen = false;
    let mut boundsheet_seen = false;
    let mut writeprotect_seen = false;
    let mut filepass_slot_open = true;
    let mut bound_sheets = Vec::<SheetObservation>::new();
    try_reserve(
        &mut bound_sheets,
        limits.max_worksheets.min(4_096),
        "BoundSheet8 inventory",
    )?;
    let mut targets = HashMap::<u64, usize>::new();
    let mut protection_collector = WorkbookProtectionState::default();
    let mut external_collector = ExternalPresenceState::default();
    let mut protection_seen = false;
    let mut protection_feed_failed = false;
    let mut external_feed_failed = false;
    let mut external_record_count = 0_usize;
    let mut active_sheet: Option<ActiveSheet> = None;

    for next in records {
        let record = match next {
            Ok(record) => record,
            Err(litchi_biff::Error::LimitExceeded { .. }) => {
                output.biff_limit = true;
                break;
            },
            Err(litchi_biff::Error::Allocation { .. }) => {
                return Err(XlsValidationError::Allocation("BIFF frame traversal"));
            },
            Err(_) => {
                output.biff_invalid = true;
                break;
            },
        };
        let kind = record.kind().get();
        let is_first_record = first_record;

        if first_record {
            first_record = false;
            if kind != BOF {
                output.biff_invalid = true;
            } else if let Some(bof) = parse_bof(record.payload()) {
                if bof.dt != WORKBOOK_BOF_TYPE {
                    output.biff_invalid = true;
                } else {
                    global_bof = Some(bof);
                }
            } else {
                output.biff_invalid = true;
            }
        }

        if !globals_done {
            if kind == FILEPASS_RECORD_TYPE {
                let valid_position =
                    global_bof.is_some() && !output.biff_invalid && filepass_slot_open;
                inspect_filepass_record(&mut output, record.payload())?;
                if !valid_position {
                    output.filepass_placement_invalid = true;
                    output.biff_invalid = true;
                }
                output.stopped_at_filepass = true;
                break;
            }

            if !is_first_record {
                if kind == WRITEPROTECT_TYPE {
                    if writeprotect_seen {
                        output.biff_invalid = true;
                        filepass_slot_open = false;
                    } else {
                        writeprotect_seen = true;
                    }
                } else {
                    filepass_slot_open = false;
                }
            }

            if kind == BOF {
                if !is_first_record {
                    output.biff_invalid = true;
                }
            } else if kind == EOF {
                if !record.payload().is_empty() {
                    output.biff_invalid = true;
                }
                globals_done = true;
                saw_global_eof = true;
            } else if kind == CODEPAGE {
                if codepage_seen || boundsheet_seen {
                    output.biff_invalid = true;
                }
                codepage_seen = true;
                if record.payload().len() != 2 {
                    output.biff_invalid = true;
                } else {
                    let codepage = u16::from_le_bytes([record.payload()[0], record.payload()[1]]);
                    if Encoding::from_codepage(codepage).is_err() {
                        output.biff_invalid = true;
                    }
                }
            } else if kind == BOUNDSHEET8 {
                boundsheet_seen = true;
                if bound_sheets.len() >= limits.max_worksheets {
                    output.worksheet_limit = true;
                } else {
                    match parse_bound_sheet(record.payload()) {
                        Some((position, sheet_type)) => {
                            let position = u64::from(position);
                            if targets.contains_key(&position) {
                                output.biff_invalid = true;
                            } else {
                                let index = bound_sheets.len();
                                try_push(
                                    &mut bound_sheets,
                                    SheetObservation {
                                        sheet_type,
                                        expected_dt: expected_sheet_bof_type(sheet_type),
                                        started: false,
                                        ended: false,
                                    },
                                    "BoundSheet8 inventory",
                                )?;
                                try_hash_insert(
                                    &mut targets,
                                    position,
                                    index,
                                    "BoundSheet8 owners",
                                )?;
                            }
                        },
                        None => output.biff_invalid = true,
                    }
                }
            }

            if !protection_feed_failed && !protection_collector.feed(kind, record.payload()) {
                protection_feed_failed = true;
            }
            if !external_feed_failed && !output.external_limit {
                if is_external_record(kind) {
                    external_record_count = external_record_count
                        .checked_add(1)
                        .ok_or(XlsValidationError::Allocation("external-link record count"))?;
                    if external_record_count > limits.max_external_records {
                        output.external_limit = true;
                    }
                }
                if !output.external_limit {
                    match external_collector.feed(kind, record.payload())? {
                        true => {},
                        false => external_feed_failed = true,
                    }
                }
            }
            continue;
        }

        if kind == FILEPASS_RECORD_TYPE {
            inspect_filepass_record(&mut output, record.payload())?;
            output.filepass_placement_invalid = true;
            output.biff_invalid = true;
            output.stopped_at_filepass = true;
            break;
        }

        let offset = u64::try_from(record.offset())
            .map_err(|_| XlsValidationError::Allocation("BIFF record offset"))?;
        if active_sheet.is_none() {
            if let Some(index) = targets.get(&offset).copied() {
                let Some(observation) = bound_sheets.get_mut(index) else {
                    output.worksheet_invalid = true;
                    output.biff_invalid = true;
                    continue;
                };
                if observation.started || kind != BOF {
                    output.worksheet_invalid = true;
                    output.biff_invalid = true;
                } else {
                    observation.started = true;
                    let expected_dt = observation.expected_dt;
                    let bof = parse_bof(record.payload());
                    let version_matches = global_bof
                        .zip(bof)
                        .is_some_and(|(global, sheet)| global.version == sheet.version);
                    if !version_matches || bof.is_none_or(|value| value.dt != expected_dt) {
                        output.worksheet_invalid = true;
                        output.biff_invalid = true;
                    }
                    active_sheet = Some(ActiveSheet {
                        index,
                        protection: SheetProtectionState::default(),
                    });
                }
            } else {
                output.biff_invalid = true;
            }
        } else if let Some(active) = active_sheet.as_mut() {
            if kind == EOF {
                if !record.payload().is_empty() {
                    output.biff_invalid = true;
                    output.worksheet_invalid = true;
                }
                if let Some(observation) = bound_sheets.get_mut(active.index) {
                    observation.ended = true;
                }
                if !active.protection.finish() {
                    protection_feed_failed = true;
                } else if active.protection.is_present() {
                    protection_seen = true;
                }
                active_sheet = None;
            } else {
                if kind == BOF {
                    output.biff_invalid = true;
                    output.worksheet_invalid = true;
                }
                if bound_sheets
                    .get(active.index)
                    .is_some_and(|sheet| sheet.sheet_type == SheetType::WorkSheet)
                    && !protection_feed_failed
                    && !active.protection.feed(kind, record.payload())
                {
                    protection_feed_failed = true;
                }
            }
        }
    }

    if !output.stopped_at_filepass && (first_record || !saw_global_eof) {
        output.biff_invalid = true;
    }
    if output.stopped_at_filepass {
        if !protection_collector.finish() {
            protection_feed_failed = true;
        }
        if let Some(active) = active_sheet.take() {
            if !active.protection.finish() {
                protection_feed_failed = true;
            }
            if active.protection.is_present() {
                protection_seen = true;
            }
        }
        let external_finished = external_collector.finish(bound_sheets.len());
        if !output.external_limit && !external_finished {
            external_feed_failed = true;
        }
        output.protection_seen = protection_seen || protection_collector.is_present();
        output.protection_invalid = protection_feed_failed;
        output.external_present = external_collector.external_present;
        output.external_count = external_collector.observed_count();
        output.external_invalid = external_feed_failed;
        return Ok(output);
    }
    if active_sheet.is_some() {
        output.worksheet_invalid = true;
        output.biff_invalid = true;
    }
    if bound_sheets.is_empty() && !output.worksheet_limit {
        output.worksheet_invalid = true;
    }
    if bound_sheets
        .iter()
        .any(|sheet| !sheet.started || !sheet.ended)
    {
        output.worksheet_invalid = true;
    }

    if saw_global_eof && !output.biff_invalid {
        if !protection_collector.finish() {
            protection_feed_failed = true;
        } else if protection_collector.is_present() {
            protection_seen = true;
        }
        if !output.external_limit && !external_feed_failed {
            if !external_collector.finish(bound_sheets.len()) {
                external_feed_failed = true;
            } else {
                output.external_present = external_collector.external_present;
                output.external_count = external_collector.observed_count();
            }
        }
    }
    output.protection_seen = protection_seen;
    output.protection_invalid = protection_feed_failed;
    output.external_invalid = external_feed_failed;
    Ok(output)
}

const FILEPASS_RECORD_TYPE: u16 = 0x002f;

fn is_external_record(kind: u16) -> bool {
    matches!(
        kind,
        SUP_BOOK_RECORD_TYPE
            | EXTERN_NAME_RECORD_TYPE
            | EXTERN_SHEET_RECORD_TYPE
            | XCT_RECORD_TYPE
            | CRN_RECORD_TYPE
            | CONTINUE_RECORD_TYPE
    )
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

fn is_encryption_component(name: &str) -> bool {
    [
        "\u{0006}DataSpaces",
        "\u{0006}DataSpaceInfo",
        "\u{0006}TransformInfo",
        "\u{0006}Primary",
        "EncryptedPackage",
        "EncryptionInfo",
    ]
    .iter()
    .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn is_drm_component(name: &str) -> bool {
    ["\u{0009}DRMContent", "\u{0009}DRMViewerContent"]
        .iter()
        .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn presence_status(
    count: u64,
    directory_limited: bool,
    limits: ValidationLimits,
) -> Result<CheckStatus, XlsValidationError> {
    if directory_limited {
        CheckStatus::blocked(
            "CFB directory inspection reached the configured entry ceiling",
            limits,
        )
        .map_err(Into::into)
    } else if count == 0 {
        Ok(CheckStatus::NotApplicable)
    } else {
        Ok(CheckStatus::Complete)
    }
}

fn push_presence_issue(
    issues: &mut Vec<ValidationIssue>,
    check: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    count: u64,
    part: &str,
    limits: ValidationLimits,
) -> Result<(), XlsValidationError> {
    if count == 0 {
        return Ok(());
    }
    try_push(
        issues,
        ValidationIssue::try_new(
            check_id(check, limits)?,
            code,
            severity,
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
                "observed.count",
                EvidenceValue::Count(count),
                limits,
            )?],
            None,
            if severity >= IssueSeverity::Error {
                CompatibilityImpact::Interoperability
            } else {
                CompatibilityImpact::None
            },
            RepairAvailability::Unavailable,
            limits,
        )?,
        "XLS validation issues",
    )?;
    Ok(())
}

fn simple_issue(
    check: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    part: Option<&str>,
    path: Option<&str>,
    offset: Option<u64>,
    limits: ValidationLimits,
) -> Result<ValidationIssue, XlsValidationError> {
    Ok(ValidationIssue::try_new(
        check_id(check, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            part, path, offset, None, None, limits,
        )?],
        [],
        None,
        if severity >= IssueSeverity::Error {
            CompatibilityImpact::Interoperability
        } else {
            CompatibilityImpact::None
        },
        RepairAvailability::Unavailable,
        limits,
    )?)
}

fn cfb_issue(
    error: &OleError,
    source_size: u64,
    limits: ValidationLimits,
) -> Result<ValidationIssue, XlsValidationError> {
    let (code, message) = match error {
        OleError::NotOleFile => (
            "xls.cfb.not_ole",
            "The source is not a recognizable CFB compound file.",
        ),
        OleError::InvalidFormat(_) => (
            "xls.cfb.invalid_format",
            "CFB ingress rejected invalid format metadata.",
        ),
        OleError::InvalidData(_) => (
            "xls.cfb.invalid_data",
            "CFB ingress rejected invalid structural data.",
        ),
        OleError::CorruptedFile(_) => (
            "xls.cfb.corrupted",
            "CFB ingress detected corrupted container topology.",
        ),
        OleError::StreamNotFound => (
            "xls.cfb.missing_stream",
            "CFB ingress could not resolve a required structural stream.",
        ),
        OleError::Io(_)
        | OleError::Allocation { .. }
        | OleError::Committed { .. }
        | OleError::SourceChanged { .. } => (
            "xls.cfb.invalid_data",
            "CFB ingress did not produce a stable structural result.",
        ),
    };
    Ok(ValidationIssue::try_new(
        check_id(CFB, limits)?,
        code,
        IssueSeverity::Error,
        message,
        [IssueLocation::try_new(
            Some("compound-file"),
            None,
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            "source.size",
            EvidenceValue::Size(source_size),
            limits,
        )?],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?)
}

fn finish_report(
    statuses: [CheckStatus; CHECK_IDS.len()],
    issues: Vec<ValidationIssue>,
    limits: ValidationLimits,
) -> Result<ValidateReport, XlsValidationError> {
    let mut checks = Vec::new();
    try_reserve(&mut checks, CHECK_IDS.len(), "XLS validation checks")?;
    for (id, status) in CHECK_IDS.into_iter().zip(statuses) {
        try_push(
            &mut checks,
            ValidationCheck::new(check_id(id, limits)?, status),
            "XLS validation checks",
        )?;
    }
    ValidateReport::try_new(checks, issues, limits).map_err(Into::into)
}

fn check_id(id: &str, limits: ValidationLimits) -> Result<CheckCapabilityId, XlsValidationError> {
    CheckCapabilityId::try_new(id, limits).map_err(Into::into)
}

fn blocked_ingress_statuses(
    limits: ValidationLimits,
    reason: &str,
) -> Result<[CheckStatus; CHECK_IDS.len()], XlsValidationError> {
    let ingress = check_id(CFB, limits)?;
    let blocked = CheckStatus::blocked(reason, limits)?;
    Ok([
        blocked,
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress.clone()),
        CheckStatus::stopped_by(ingress),
    ])
}

fn cfb_rejection_statuses(
    limits: ValidationLimits,
) -> Result<[CheckStatus; CHECK_IDS.len()], XlsValidationError> {
    let blocked = CheckStatus::blocked("CFB ingress was structurally rejected", limits)?;
    Ok([
        CheckStatus::Complete,
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked.clone(),
        blocked,
    ])
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

fn require_version(
    source: &dyn ReadAt,
    expected: litchi_core::SourceVersion,
) -> Result<(), XlsValidationError> {
    let observed = source
        .version()
        .map_err(|error| XlsValidationError::Ingress(OleError::Io(error)))?;
    if observed == expected {
        Ok(())
    } else {
        Err(XlsValidationError::Ingress(OleError::SourceChanged {
            expected,
            observed,
        }))
    }
}

fn try_reserve<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> Result<(), XlsValidationError> {
    values
        .try_reserve(additional)
        .map_err(|_| XlsValidationError::Allocation(resource))
}

fn try_push<T>(
    values: &mut Vec<T>,
    value: T,
    resource: &'static str,
) -> Result<(), XlsValidationError> {
    if values.len() == values.capacity() {
        try_reserve(values, 1, resource)?;
    }
    values.push(value);
    Ok(())
}

fn try_hash_insert(
    values: &mut HashMap<u64, usize>,
    key: u64,
    value: usize,
    resource: &'static str,
) -> Result<(), XlsValidationError> {
    if values.len() == values.capacity() {
        values
            .try_reserve(1)
            .map_err(|_| XlsValidationError::Allocation(resource))?;
    }
    values.insert(key, value);
    Ok(())
}

fn try_hash_set_reserve<T: Eq + Hash>(
    values: &mut HashSet<T>,
    additional: usize,
    resource: &'static str,
) -> Result<(), XlsValidationError> {
    if additional > values.capacity().saturating_sub(values.len()) {
        values
            .try_reserve(additional)
            .map_err(|_| XlsValidationError::Allocation(resource))?;
    }
    Ok(())
}
