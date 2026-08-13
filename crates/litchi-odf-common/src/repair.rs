//! One deliberately narrow, source-backed ODF repair.
//!
//! The only repair exposed here removes one well-formed Extended Timestamp
//! field from the local header of an otherwise valid first, stored
//! `mimetype` member.  The central record is retained byte-for-byte apart
//! from the local-header offsets that must move when the local field is
//! removed.  Every other local member, central record field, EOCD field, and
//! archive comment is copied from the source.
//!
//! Planning is fully preflighted: no caller sink is touched until the source,
//! report, candidate archive, semantic member digests, and preservation proof
//! all succeed.  The returned plan does not own generated output and offers no
//! inverse operation.  Callers own any atomic replacement policy.

use std::{
    borrow::Cow,
    collections::HashSet,
    fmt,
    io::{self, Write},
    num::TryFromIntError,
    ops::Range,
};

use litchi_core::{EvidenceDigest, EvidenceValue, IssueSeverity, ValidateReport};
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesStart, Event, attributes::Attribute},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use serde_json::json;
use sha2::{Digest as _, Sha256};
use soapberry_zip::{
    Error as ZipError, PreservationIndex, ZipArchive,
    extra_fields::{ExtraFieldId, ExtraFields},
    office::{ArchiveLimits, ArchiveReader},
};

use crate::{OdfValidationError, OdfValidationLimits, validate_package_with_limits};

/// Stable report issue code for the supported repair.
pub const MIMETYPE_LOCAL_EXTRA_ISSUE: &str = "odf.mimetype.local_header_extra";
/// Stable format-owned repair identifier.
pub const MIMETYPE_LOCAL_EXTRA_REPAIR: &str = "odf.repair.mimetype_local_extra";

const MIMETYPE: &[u8] = b"mimetype";
const DECLARATION_CHECK: &str = "odf.package.mimetype_manifest";
const LOCAL_OFFSET: Range<usize> = 42..46;
const LOCAL_FIXED: usize = 30;
const CENTRAL_FIXED: usize = 46;
const EOCD_FIXED: usize = 22;
const PUBLICATION_SCRATCH: usize = 64 * 1024;
const NAME_SET_ENTRY_ESTIMATE: u64 = 128;
const LAYOUT_ENTRY_ESTIMATE: u64 = 64;
const PRESERVATION_ENTRY_ESTIMATE: u64 = 128;
const LOCAL_ORDER_ENTRY_ESTIMATE: u64 = 16;
const MANIFEST_PATH: &str = "META-INF/manifest.xml";

/// Finite bounds for one ODF local-extra repair plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct OdfRepairLimits {
    max_input_bytes: u64,
    max_output_bytes: u64,
    max_members: usize,
    max_member_bytes: u64,
    max_total_member_bytes: u64,
    max_member_name_bytes: u64,
    max_metadata_bytes: u64,
    max_plan_json_bytes: usize,
    max_extra_bytes: usize,
    max_scratch_bytes: usize,
    max_preflight_candidate_bytes: u64,
}

impl OdfRepairLimits {
    /// Creates finite repair bounds. Every bound must be non-zero.
    pub const fn new(
        max_input_bytes: u64,
        max_output_bytes: u64,
        max_members: usize,
        max_member_bytes: u64,
        max_total_member_bytes: u64,
        max_plan_json_bytes: usize,
        max_extra_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Result<Self, RepairError> {
        let limits = Self {
            max_input_bytes,
            max_output_bytes,
            max_members,
            max_member_bytes,
            max_total_member_bytes,
            // Keep the original constructor arity stable. Callers that need
            // tighter metadata bounds use the explicit `with_max_*` builders.
            max_member_name_bytes: 4 * 1024,
            max_metadata_bytes: 64 * 1024 * 1024,
            max_plan_json_bytes,
            max_extra_bytes,
            max_scratch_bytes,
            // Planning proves preservation with one bounded candidate. The
            // default is deliberately far below the 2 GiB I/O ceilings;
            // publication revalidates by streaming and does not build a
            // second full candidate.
            max_preflight_candidate_bytes: 64 * 1024 * 1024,
        };
        if !limits.is_valid() {
            return Err(RepairError::InvalidLimits);
        }
        Ok(limits)
    }

    const fn is_valid(self) -> bool {
        self.max_input_bytes != 0
            && self.max_output_bytes != 0
            && self.max_members != 0
            && self.max_member_bytes != 0
            && self.max_total_member_bytes != 0
            && self.max_member_name_bytes != 0
            && self.max_metadata_bytes != 0
            && self.max_plan_json_bytes != 0
            && self.max_extra_bytes != 0
            && self.max_scratch_bytes >= PUBLICATION_SCRATCH
            && self.max_preflight_candidate_bytes != 0
    }

    /// Maximum source bytes accepted by planning and publication.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum target bytes accepted by planning and publication.
    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }

    /// Maximum ZIP member count.
    #[must_use]
    pub const fn max_members(self) -> usize {
        self.max_members
    }

    /// Maximum uncompressed bytes in one member.
    #[must_use]
    pub const fn max_member_bytes(self) -> u64 {
        self.max_member_bytes
    }

    /// Maximum aggregate uncompressed member bytes.
    #[must_use]
    pub const fn max_total_member_bytes(self) -> u64 {
        self.max_total_member_bytes
    }

    /// Maximum raw bytes in one member name.
    #[must_use]
    pub const fn max_member_name_bytes(self) -> u64 {
        self.max_member_name_bytes
    }

    /// Maximum aggregate raw catalog metadata bytes retained during planning.
    #[must_use]
    pub const fn max_metadata_bytes(self) -> u64 {
        self.max_metadata_bytes
    }

    /// Maximum serialized plan JSON bytes.
    #[must_use]
    pub const fn max_plan_json_bytes(self) -> usize {
        self.max_plan_json_bytes
    }

    /// Maximum local-extra bytes retained by a plan.
    #[must_use]
    pub const fn max_extra_bytes(self) -> usize {
        self.max_extra_bytes
    }

    /// Maximum fixed publication scratch bytes.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// Maximum transient candidate bytes used by plan-time reopen proof.
    #[must_use]
    pub const fn max_preflight_candidate_bytes(self) -> u64 {
        self.max_preflight_candidate_bytes
    }

    /// Returns a copy with a smaller or larger transient candidate ceiling.
    #[must_use]
    pub const fn with_max_preflight_candidate_bytes(mut self, maximum: u64) -> Self {
        self.max_preflight_candidate_bytes = maximum;
        self
    }

    /// Returns a copy with a raw member-name ceiling.
    #[must_use]
    pub const fn with_max_member_name_bytes(mut self, maximum: u64) -> Self {
        self.max_member_name_bytes = maximum;
        self
    }

    /// Returns a copy with an aggregate raw metadata ceiling.
    #[must_use]
    pub const fn with_max_metadata_bytes(mut self, maximum: u64) -> Self {
        self.max_metadata_bytes = maximum;
        self
    }
}

impl Default for OdfRepairLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 2 * 1024 * 1024 * 1024,
            max_output_bytes: 2 * 1024 * 1024 * 1024,
            max_members: 100_000,
            max_member_bytes: 512 * 1024 * 1024,
            max_total_member_bytes: 2 * 1024 * 1024 * 1024,
            max_member_name_bytes: 4 * 1024,
            max_metadata_bytes: 64 * 1024 * 1024,
            max_plan_json_bytes: 4 * 1024,
            max_extra_bytes: 1024,
            max_scratch_bytes: PUBLICATION_SCRATCH,
            max_preflight_candidate_bytes: 64 * 1024 * 1024,
        }
    }
}

/// Exact progress known for a non-atomic sequential sink after failure.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OutputProgress {
    /// The sink accepted no bytes.
    Untouched,
    /// The sink accepted an exact artifact prefix.
    Prefix { accepted: u64, expected: u64 },
    /// All artifact bytes were accepted but flush failed.
    CompleteUnflushed { bytes: u64 },
    /// All artifact bytes were accepted, but the post-write integrity proof failed.
    CompleteUnverified { bytes: u64 },
    /// The sink reported more bytes than offered, so exact progress is not known.
    Indeterminate { accepted_before: u64 },
}

/// SHA-256 identity used by the repair plan and publication report.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RepairFingerprint([u8; 32]);

impl RepairFingerprint {
    /// Computes an exact SHA-256 fingerprint.
    #[must_use]
    pub fn of(bytes: &[u8]) -> Self {
        Self(Sha256::digest(bytes).into())
    }

    /// Returns the exact digest bytes.
    #[must_use]
    pub const fn as_bytes(self) -> [u8; 32] {
        self.0
    }

    /// Returns lowercase hexadecimal digest text.
    #[must_use]
    pub fn as_hex(self) -> String {
        let mut text = String::with_capacity(64);
        for byte in self.0 {
            use fmt::Write as _;
            let _ = write!(text, "{byte:02x}");
        }
        text
    }
}

/// The sole supported ODF repair action.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RemoveMimetypeLocalExtra {
    field_id: u16,
    field_bytes: u16,
}

impl RemoveMimetypeLocalExtra {
    /// ZIP extra-field identifier.
    #[must_use]
    pub const fn field_id(self) -> u16 {
        self.field_id
    }

    /// Complete encoded field length, including its four-byte header.
    #[must_use]
    pub const fn field_bytes(self) -> u16 {
        self.field_bytes
    }
}

/// A fully preflighted, borrowed ODF repair plan.
pub struct MimetypeRepairPlan<'source> {
    source: &'source [u8],
    limits: OdfRepairLimits,
    source_len: u64,
    source_fingerprint: RepairFingerprint,
    target_len: u64,
    target_fingerprint: RepairFingerprint,
    members: usize,
    action: RemoveMimetypeLocalExtra,
}

impl fmt::Debug for MimetypeRepairPlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MimetypeRepairPlan")
            .field("source_len", &self.source_len)
            .field("target_len", &self.target_len)
            .field("members", &self.members)
            .field("action", &self.action)
            .finish_non_exhaustive()
    }
}

impl<'source> MimetypeRepairPlan<'source> {
    /// Source length bound to this plan.
    #[must_use]
    pub const fn source_len(&self) -> u64 {
        self.source_len
    }

    /// Target length after removing the local field.
    #[must_use]
    pub const fn output_len(&self) -> u64 {
        self.target_len
    }

    /// Source SHA-256 fingerprint bound to this plan.
    #[must_use]
    pub const fn source_fingerprint(&self) -> RepairFingerprint {
        self.source_fingerprint
    }

    /// Target SHA-256 fingerprint established during preflight.
    #[must_use]
    pub const fn output_fingerprint(&self) -> RepairFingerprint {
        self.target_fingerprint
    }

    /// Number of central-directory members.
    #[must_use]
    pub const fn member_count(&self) -> usize {
        self.members
    }

    /// The sole forward action.
    #[must_use]
    pub const fn action(&self) -> RemoveMimetypeLocalExtra {
        self.action
    }

    /// Serializes deterministic bounded plan metadata without source bytes.
    pub fn to_json(&self) -> Result<String, RepairError> {
        let value = json!({
            "schema": "odf-repair/mimetype-local-extra/v1",
            "repair_id": MIMETYPE_LOCAL_EXTRA_REPAIR,
            "action": "remove_mimetype_local_extra",
            "source_len": self.source_len,
            "source_sha256": self.source_fingerprint.as_hex(),
            "output_len": self.target_len,
            "output_sha256": self.target_fingerprint.as_hex(),
            "member_count": self.members,
            "extra_field_id": self.action.field_id,
            "extra_field_bytes": self.action.field_bytes,
        });
        let encoded = serde_json::to_vec(&value).map_err(|_| RepairError::PlanJsonEncoding)?;
        if encoded.len() > self.limits.max_plan_json_bytes {
            return Err(RepairError::Limit {
                resource: "plan JSON bytes",
                observed: encoded.len() as u64,
                limit: self.limits.max_plan_json_bytes as u64,
            });
        }
        String::from_utf8(encoded).map_err(|_| RepairError::PlanJsonEncoding)
    }

    /// Streams the complete repaired archive to a caller-owned sink.
    ///
    /// All source and candidate checks are repeated before the first sink
    /// byte. A generic sequential sink is not atomic; failures after progress
    /// return [`RepairError::IncompleteOutput`].
    pub fn write_to<W: Write>(&self, sink: &mut W) -> Result<RepairPublication, RepairError> {
        let layout = self.preflight()?;
        let mut accepted = 0_u64;
        let mut target_hasher = Sha256::new();
        let result = emit_output(
            self.source,
            &layout,
            self.action,
            sink,
            &mut accepted,
            &mut target_hasher,
        );
        if let Err(error) = result {
            let indeterminate = matches!(&error, RepairError::SinkOverreported);
            return Err(with_progress(
                error,
                accepted,
                self.target_len,
                indeterminate,
            ));
        }

        let observed_target = RepairFingerprint(target_hasher.finalize().into());
        if observed_target != self.target_fingerprint {
            return Err(with_progress(
                RepairError::TargetChanged {
                    expected: self.target_fingerprint,
                    observed: observed_target,
                },
                accepted,
                self.target_len,
                false,
            ));
        }
        let observed_source = RepairFingerprint::of(self.source);
        if observed_source != self.source_fingerprint {
            return Err(with_progress(
                RepairError::SourceChanged {
                    expected: self.source_fingerprint,
                    observed: observed_source,
                },
                accepted,
                self.target_len,
                false,
            ));
        }
        if let Err(error) = sink.flush() {
            return Err(RepairError::IncompleteOutput {
                progress: OutputProgress::CompleteUnflushed { bytes: accepted },
                source: Box::new(RepairError::Io(error)),
            });
        }
        Ok(RepairPublication {
            bytes: accepted,
            source_fingerprint: self.source_fingerprint,
            target_fingerprint: self.target_fingerprint,
            action: self.action,
        })
    }

    fn preflight(&self) -> Result<Layout, RepairError> {
        let observed_source = RepairFingerprint::of(self.source);
        let observed_len = u64::try_from(self.source.len()).map_err(RepairError::Integer)?;
        if observed_len != self.source_len || observed_source != self.source_fingerprint {
            return Err(RepairError::SourceChanged {
                expected: self.source_fingerprint,
                observed: observed_source,
            });
        }
        let layout = inspect_layout(self.source, self.limits)?;
        if layout.member_count != self.members
            || layout.action != self.action
            || layout.output_len != self.target_len
        {
            return Err(RepairError::PlanMismatch);
        }
        // Do not allocate a second full candidate at publication time. The
        // source fingerprint above binds this pass to the exact bytes proved
        // during planning; replay the deterministic transform into a sink
        // that retains no output and recheck its digest, length, source
        // member semantics, and raw-preservation arithmetic before writing.
        let mut accepted = 0_u64;
        let mut target_hasher = Sha256::new();
        let mut sink = io::sink();
        emit_output(
            self.source,
            &layout,
            self.action,
            &mut sink,
            &mut accepted,
            &mut target_hasher,
        )?;
        if accepted != self.target_len {
            return Err(RepairError::PlanMismatch);
        }
        let target = RepairFingerprint(target_hasher.finalize().into());
        if target != self.target_fingerprint {
            return Err(RepairError::TargetChanged {
                expected: self.target_fingerprint,
                observed: target,
            });
        }
        verify_streaming_proof(self.source, &layout, self.limits)?;
        Ok(layout)
    }
}

/// Successful publication evidence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RepairPublication {
    bytes: u64,
    source_fingerprint: RepairFingerprint,
    target_fingerprint: RepairFingerprint,
    action: RemoveMimetypeLocalExtra,
}

impl RepairPublication {
    /// Complete target bytes accepted by the sink.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.bytes
    }

    /// Source identity rechecked during publication.
    #[must_use]
    pub const fn source_fingerprint(self) -> RepairFingerprint {
        self.source_fingerprint
    }

    /// Target identity emitted during publication.
    #[must_use]
    pub const fn target_fingerprint(self) -> RepairFingerprint {
        self.target_fingerprint
    }

    /// Action that was emitted.
    #[must_use]
    pub const fn action(self) -> RemoveMimetypeLocalExtra {
        self.action
    }
}

/// Failure to construct or publish the narrow repair.
#[derive(Debug)]
#[non_exhaustive]
pub enum RepairError {
    /// A zero or too-small bound was supplied.
    InvalidLimits,
    /// The supplied report did not identify exactly this repair target.
    ReportMismatch,
    /// The source is outside the bounded repair contract.
    Unsupported { reason: &'static str },
    /// A finite resource bound was exceeded.
    Limit {
        /// Content-free resource name.
        resource: &'static str,
        /// Observed bounded value.
        observed: u64,
        /// Configured finite ceiling.
        limit: u64,
    },
    /// Source identity changed after planning.
    SourceChanged {
        /// Identity captured at planning time.
        expected: RepairFingerprint,
        /// Identity observed during the later operation.
        observed: RepairFingerprint,
    },
    /// Candidate identity differed from preflight.
    TargetChanged {
        /// Identity captured at planning time.
        expected: RepairFingerprint,
        /// Identity observed during the later operation.
        observed: RepairFingerprint,
    },
    /// The plan no longer matches the source layout.
    PlanMismatch,
    /// ZIP substrate rejected the source or candidate.
    Zip(ZipError),
    /// ODF validator could not produce a bounded report.
    Validation(OdfValidationError),
    /// A bounded temporary allocation failed.
    Allocation {
        /// Content-free allocation name.
        resource: &'static str,
    },
    /// Plan JSON serialization failed.
    PlanJsonEncoding,
    /// Sequential sink I/O failed before or during publication.
    Io(io::Error),
    /// The sink may contain an exact or indeterminate prefix.
    IncompleteOutput {
        /// Known sink progress.
        progress: OutputProgress,
        /// Underlying failure.
        source: Box<RepairError>,
    },
    /// A source or candidate member could not be read consistently.
    MemberMismatch { name: &'static str },
    /// A ZIP integer could not fit the bounded host representation.
    Integer(TryFromIntError),
    /// A sink reported more bytes than the offered slice.
    SinkOverreported,
}

impl fmt::Display for RepairError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidLimits => formatter.write_str("ODF repair limits are invalid"),
            Self::ReportMismatch => {
                formatter.write_str("validation report does not authorize this ODF repair")
            },
            Self::Unsupported { reason } => {
                write!(formatter, "ODF repair is unavailable: {reason}")
            },
            Self::Limit {
                resource,
                observed,
                limit,
            } => {
                write!(
                    formatter,
                    "ODF repair {resource} {observed} exceeds limit {limit}"
                )
            },
            Self::SourceChanged { .. } => {
                formatter.write_str("ODF repair source fingerprint changed")
            },
            Self::TargetChanged { .. } => {
                formatter.write_str("ODF repair target fingerprint changed")
            },
            Self::PlanMismatch => {
                formatter.write_str("ODF repair plan no longer matches its source layout")
            },
            Self::Zip(error) => write!(formatter, "ODF repair ZIP validation failed: {error}"),
            Self::Validation(error) => {
                write!(formatter, "ODF repair report validation failed: {error}")
            },
            Self::Allocation { resource } => {
                write!(formatter, "ODF repair could not allocate {resource}")
            },
            Self::PlanJsonEncoding => formatter.write_str("ODF repair plan JSON encoding failed"),
            Self::Io(error) => write!(formatter, "ODF repair output I/O failed: {error}"),
            Self::IncompleteOutput { progress, source } => {
                write!(
                    formatter,
                    "ODF repair output is incomplete ({progress:?}): {source}"
                )
            },
            Self::MemberMismatch { name } => {
                write!(formatter, "ODF repair member mismatch: {name}")
            },
            Self::Integer(_) => {
                formatter.write_str("ODF repair integer does not fit the bounded representation")
            },
            Self::SinkOverreported => {
                formatter.write_str("ODF repair sink reported more bytes than offered")
            },
        }
    }
}

impl std::error::Error for RepairError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Zip(error) => Some(error),
            Self::Validation(error) => Some(error),
            Self::Io(error) => Some(error),
            Self::IncompleteOutput { source, .. } => Some(source),
            _ => None,
        }
    }
}

impl From<ZipError> for RepairError {
    fn from(error: ZipError) -> Self {
        Self::Zip(error)
    }
}

impl From<io::Error> for RepairError {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<OdfValidationError> for RepairError {
    fn from(error: OdfValidationError) -> Self {
        Self::Validation(error)
    }
}

impl From<TryFromIntError> for RepairError {
    fn from(error: TryFromIntError) -> Self {
        Self::Integer(error)
    }
}

/// Build the narrow repair plan from a source and its completed validation report.
pub fn plan_mimetype_local_extra<'source>(
    source: &'source [u8],
    report: &ValidateReport,
    limits: OdfRepairLimits,
) -> Result<MimetypeRepairPlan<'source>, RepairError> {
    if !limits.is_valid() {
        return Err(RepairError::InvalidLimits);
    }
    let source_len = u64::try_from(source.len()).map_err(RepairError::Integer)?;
    if source_len > limits.max_input_bytes {
        return Err(RepairError::Limit {
            resource: "input bytes",
            observed: source_len,
            limit: limits.max_input_bytes,
        });
    }
    check_report(source, report, limits)?;
    let layout = inspect_layout(source, limits)?;
    let candidate = build_candidate(source, &layout, layout.action, limits)?;
    verify_candidate(source, &candidate, &layout, limits)?;
    let target_len = u64::try_from(candidate.len()).map_err(RepairError::Integer)?;
    let source_fingerprint = RepairFingerprint::of(source);
    let target_fingerprint = RepairFingerprint::of(&candidate);
    let plan = MimetypeRepairPlan {
        source,
        limits,
        source_len,
        source_fingerprint,
        target_len,
        target_fingerprint,
        members: layout.member_count,
        action: layout.action,
    };
    let _ = plan.to_json()?;
    Ok(plan)
}

/// Alias with a shorter operation-oriented name.
pub fn plan_mimetype_repair<'source>(
    source: &'source [u8],
    report: &ValidateReport,
    limits: OdfRepairLimits,
) -> Result<MimetypeRepairPlan<'source>, RepairError> {
    plan_mimetype_local_extra(source, report, limits)
}

fn check_report(
    source: &[u8],
    report: &ValidateReport,
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    if !report.is_complete() || report.issues().len() != 1 {
        return Err(RepairError::ReportMismatch);
    }
    let issue = &report.issues()[0];
    if issue.check().as_str() != DECLARATION_CHECK
        || issue.code() != MIMETYPE_LOCAL_EXTRA_ISSUE
        || issue.severity() != IssueSeverity::Error
        || issue.repair().repair_id() != Some(MIMETYPE_LOCAL_EXTRA_REPAIR)
        || issue
            .locations()
            .iter()
            .all(|location| location.part() != Some("mimetype"))
    {
        return Err(RepairError::ReportMismatch);
    }
    if issue.evidence().len() != 2 {
        return Err(RepairError::ReportMismatch);
    }
    let expected_len = u64::try_from(source.len()).map_err(RepairError::Integer)?;
    let expected_digest = EvidenceDigest::of(source);
    let mut source_size = None;
    let mut source_sha256 = None;
    for evidence in issue.evidence() {
        match (evidence.key(), evidence.value()) {
            ("source_size", EvidenceValue::Size(value)) => source_size = Some(value),
            ("source_sha256", EvidenceValue::Sha256(value)) => source_sha256 = Some(value),
            _ => {},
        }
    }
    if source_size != Some(expected_len) || source_sha256 != Some(expected_digest) {
        return Err(RepairError::ReportMismatch);
    }
    let fresh = validate_package_with_limits(
        source,
        validation_limits(limits, limits.max_input_bytes),
        litchi_core::ValidationLimits::default(),
    )?;
    if !fresh.is_complete() || fresh.issues().len() != 1 || fresh.issues()[0].id() != issue.id() {
        return Err(RepairError::ReportMismatch);
    }
    Ok(())
}

fn archive_limits(limits: OdfRepairLimits) -> ArchiveLimits {
    ArchiveLimits {
        max_files: limits.max_members,
        max_member_name_bytes: limits.max_member_name_bytes,
        max_metadata_bytes: limits.max_metadata_bytes,
        max_compressed_size: limits.max_member_bytes,
        max_entry_size: limits.max_member_bytes,
        max_total_size: limits.max_total_member_bytes,
    }
}

fn validation_limits(limits: OdfRepairLimits, max_input_bytes: u64) -> OdfValidationLimits {
    OdfValidationLimits::default()
        .with_max_input_bytes(max_input_bytes)
        .with_max_entries(limits.max_members)
        .with_max_archive_member_name_bytes(limits.max_member_name_bytes)
        .with_max_archive_metadata_bytes(limits.max_metadata_bytes)
        .with_max_archive_compressed_bytes(limits.max_member_bytes)
        .with_max_archive_entry_bytes(limits.max_member_bytes)
        .with_max_archive_total_bytes(limits.max_total_member_bytes)
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct EntryLayout {
    local: Range<usize>,
    central: Range<usize>,
    local_offset: usize,
    compressed_size: usize,
    uncompressed_size: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Layout {
    entries: Vec<EntryLayout>,
    local_order: Vec<usize>,
    central_start: usize,
    eocd: usize,
    output_len: u64,
    action: RemoveMimetypeLocalExtra,
    member_count: usize,
    non_directory_count: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct CentralPreflight {
    record_count: usize,
    non_directory_count: usize,
    raw_name_bytes: u64,
    central_record_bytes: u64,
    archive_comment_bytes: u64,
}

/// Parse only fixed central-directory/EOCD bytes before the preservation index
/// can copy central records or allocate its entry table. This pass retains no
/// names, comments, extras, or record bytes; all observed quantities are
/// charged to the one repair metadata budget.
fn preflight_central_directory(
    source: &[u8],
    central_start: usize,
    eocd: usize,
    limits: OdfRepairLimits,
) -> Result<CentralPreflight, RepairError> {
    let eocd_end = eocd
        .checked_add(EOCD_FIXED)
        .ok_or(RepairError::Unsupported {
            reason: "EOCD bounds overflow",
        })?;
    if central_start > eocd || eocd_end > source.len() {
        return Err(RepairError::Unsupported {
            reason: "invalid central-directory or EOCD bounds",
        });
    }
    let eocd_bytes = &source[eocd..eocd_end];
    if le_u32(eocd_bytes, 0) != Some(0x0605_4b50) {
        return Err(RepairError::Unsupported {
            reason: "malformed EOCD signature",
        });
    }
    let disk = le_u16(eocd_bytes, 4).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD disk number",
    })?;
    let central_disk = le_u16(eocd_bytes, 6).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD central disk number",
    })?;
    let disk_entries = le_u16(eocd_bytes, 8).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD disk entry count",
    })?;
    let total_entries = le_u16(eocd_bytes, 10).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD entry count",
    })?;
    let central_size = le_u32(eocd_bytes, 12).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD central size",
    })?;
    let declared_central_start = le_u32(eocd_bytes, 16).ok_or(RepairError::Unsupported {
        reason: "malformed EOCD central offset",
    })?;
    let archive_comment_bytes =
        u64::from(le_u16(eocd_bytes, 20).ok_or(RepairError::Unsupported {
            reason: "malformed EOCD comment length",
        })?);
    if disk != 0
        || central_disk != 0
        || disk_entries != total_entries
        || total_entries == u16::MAX
        || central_size == u32::MAX
        || declared_central_start == u32::MAX
        || u64::from(declared_central_start) != central_start as u64
        || u64::from(central_size) != (eocd - central_start) as u64
    {
        return Err(RepairError::Unsupported {
            reason: "ZIP64, multidisk, or inconsistent central-directory EOCD fields",
        });
    }
    let archive_end = eocd_end
        .checked_add(usize::try_from(archive_comment_bytes).map_err(RepairError::Integer)?)
        .ok_or(RepairError::Unsupported {
            reason: "EOCD comment bounds overflow",
        })?;
    if archive_end != source.len() {
        return Err(RepairError::Unsupported {
            reason: "EOCD comment has trailing or truncated bytes",
        });
    }
    if archive_end - eocd > limits.max_scratch_bytes {
        return Err(RepairError::Limit {
            resource: "EOCD scratch bytes",
            observed: (archive_end - eocd) as u64,
            limit: limits.max_scratch_bytes as u64,
        });
    }
    let record_count = usize::from(total_entries);
    if record_count > limits.max_members {
        return Err(RepairError::Limit {
            resource: "member count",
            observed: record_count as u64,
            limit: limits.max_members as u64,
        });
    }

    let mut cursor = central_start;
    let mut non_directory_count = 0_usize;
    let mut raw_name_bytes = 0_u64;
    let mut central_record_bytes = 0_u64;
    let mut total_uncompressed = 0_u64;
    for index in 0..record_count {
        let fixed_end = cursor
            .checked_add(CENTRAL_FIXED)
            .ok_or(RepairError::Unsupported {
                reason: "central-directory record bounds overflow",
            })?;
        if fixed_end > eocd {
            return Err(RepairError::Unsupported {
                reason: "central-directory record is truncated",
            });
        }
        let fixed = &source[cursor..fixed_end];
        if le_u32(fixed, 0) != Some(0x0201_4b50) {
            return Err(RepairError::Unsupported {
                reason: "malformed central-directory signature",
            });
        }
        let flags = le_u16(fixed, 8).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory flags",
        })?;
        if flags & 1 != 0 {
            return Err(RepairError::Unsupported {
                reason: "encrypted ZIP member",
            });
        }
        let compressed_size = u64::from(le_u32(fixed, 20).ok_or(RepairError::Unsupported {
            reason: "malformed central compressed size",
        })?);
        let uncompressed_size = u64::from(le_u32(fixed, 24).ok_or(RepairError::Unsupported {
            reason: "malformed central uncompressed size",
        })?);
        let name_len = usize::from(le_u16(fixed, 28).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory name length",
        })?);
        let extra_len = usize::from(le_u16(fixed, 30).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory extra length",
        })?);
        let comment_len = usize::from(le_u16(fixed, 32).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory comment length",
        })?);
        let record_len = CENTRAL_FIXED
            .checked_add(name_len)
            .and_then(|value| value.checked_add(extra_len))
            .and_then(|value| value.checked_add(comment_len))
            .ok_or(RepairError::Unsupported {
                reason: "central-directory record length overflow",
            })?;
        let record_end = cursor
            .checked_add(record_len)
            .ok_or(RepairError::Unsupported {
                reason: "central-directory record bounds overflow",
            })?;
        if record_end > eocd {
            return Err(RepairError::Unsupported {
                reason: "central-directory record is truncated or has invalid lengths",
            });
        }
        if record_len > limits.max_scratch_bytes {
            return Err(RepairError::Limit {
                resource: "central scratch bytes",
                observed: record_len as u64,
                limit: limits.max_scratch_bytes as u64,
            });
        }
        let record = &source[cursor..record_end];
        let name = &record[CENTRAL_FIXED..CENTRAL_FIXED + name_len];
        let name_bytes = name_len as u64;
        if name_bytes > limits.max_member_name_bytes {
            return Err(RepairError::Limit {
                resource: "member name bytes",
                observed: name_bytes,
                limit: limits.max_member_name_bytes,
            });
        }
        raw_name_bytes =
            raw_name_bytes
                .checked_add(name_bytes)
                .ok_or(RepairError::Unsupported {
                    reason: "aggregate member-name bytes overflow",
                })?;
        central_record_bytes = central_record_bytes.checked_add(record_len as u64).ok_or(
            RepairError::Unsupported {
                reason: "aggregate central metadata overflow",
            },
        )?;
        let observed_metadata = repair_metadata_bytes(
            index + 1,
            raw_name_bytes,
            central_record_bytes,
            archive_comment_bytes,
        )?;
        if observed_metadata > limits.max_metadata_bytes {
            return Err(RepairError::Limit {
                resource: "repair metadata bytes",
                observed: observed_metadata,
                limit: limits.max_metadata_bytes,
            });
        }
        if name.last() == Some(&b'/') {
            // Directory records remain part of the raw preservation index but
            // are not returned by ArchiveReader::file_names().
        } else {
            non_directory_count = non_directory_count.saturating_add(1);
        }
        if compressed_size > limits.max_member_bytes {
            return Err(RepairError::Limit {
                resource: "compressed member bytes",
                observed: compressed_size,
                limit: limits.max_member_bytes,
            });
        }
        if uncompressed_size > limits.max_member_bytes {
            return Err(RepairError::Limit {
                resource: "member bytes",
                observed: uncompressed_size,
                limit: limits.max_member_bytes,
            });
        }
        total_uncompressed =
            total_uncompressed
                .checked_add(uncompressed_size)
                .ok_or(RepairError::Unsupported {
                    reason: "aggregate member size overflow",
                })?;
        if total_uncompressed > limits.max_total_member_bytes {
            return Err(RepairError::Limit {
                resource: "aggregate member bytes",
                observed: total_uncompressed,
                limit: limits.max_total_member_bytes,
            });
        }
        let disk_start = le_u16(fixed, 34).ok_or(RepairError::Unsupported {
            reason: "malformed central disk number",
        })?;
        let local_offset = le_u32(fixed, 42).ok_or(RepairError::Unsupported {
            reason: "malformed central local-header offset",
        })?;
        if disk_start != 0
            || local_offset == u32::MAX
            || u64::from(local_offset) >= central_start as u64
        {
            return Err(RepairError::Unsupported {
                reason: "multidisk or invalid local-header offset",
            });
        }
        cursor = record_end;
    }
    if cursor != eocd {
        return Err(RepairError::Unsupported {
            reason: "central records are not exactly contiguous through the EOCD",
        });
    }
    Ok(CentralPreflight {
        record_count,
        non_directory_count,
        raw_name_bytes,
        central_record_bytes,
        archive_comment_bytes,
    })
}

fn inspect_layout(source: &[u8], limits: OdfRepairLimits) -> Result<Layout, RepairError> {
    if !limits.is_valid() {
        return Err(RepairError::InvalidLimits);
    }
    let source_len = u64::try_from(source.len()).map_err(RepairError::Integer)?;
    if source_len > limits.max_input_bytes {
        return Err(RepairError::Limit {
            resource: "input bytes",
            observed: source_len,
            limit: limits.max_input_bytes,
        });
    }
    if limits.max_scratch_bytes < PUBLICATION_SCRATCH {
        return Err(RepairError::InvalidLimits);
    }

    let archive = ZipArchive::from_slice(source)?.into_zip_archive();
    if archive.end_offset() != source_len || archive.is_zip64() {
        return Err(RepairError::Unsupported {
            reason: "ZIP64, prefixed, or trailing source archive",
        });
    }
    let central_start =
        usize::try_from(archive.directory_offset()).map_err(RepairError::Integer)?;
    let eocd = usize::try_from(archive.eocd_offset()).map_err(RepairError::Integer)?;
    let central_preflight = preflight_central_directory(source, central_start, eocd, limits)?;
    if central_preflight.record_count == 0 {
        return Err(RepairError::Unsupported {
            reason: "source archive has no members",
        });
    }
    let mut index_buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut index_buffer)?;
    let entries = index.entries();
    if entries.len() != central_preflight.record_count {
        return Err(RepairError::Unsupported {
            reason: "central-directory record count changed during indexing",
        });
    }
    if central_start > eocd || eocd.checked_add(EOCD_FIXED).is_none() {
        return Err(RepairError::Unsupported {
            reason: "invalid central-directory or EOCD bounds",
        });
    }

    check_repair_metadata(
        central_preflight.record_count,
        central_preflight.raw_name_bytes,
        central_preflight.central_record_bytes,
        central_preflight.archive_comment_bytes,
        limits,
    )?;
    let mut layouts = Vec::new();
    layouts
        .try_reserve_exact(entries.len())
        .map_err(|_| RepairError::Allocation {
            resource: "repair member index",
        })?;
    let mut names: HashSet<&[u8]> = HashSet::new();
    // The fixed table estimate is charged together with the layout vector and
    // the running raw-name total, so the combined retained catalog metadata
    // never exceeds one repair metadata ceiling.
    names
        .try_reserve(entries.len())
        .map_err(|_| RepairError::Allocation {
            resource: "repair member names",
        })?;
    let mut total_uncompressed = 0_u64;
    for entry in entries {
        let local_start =
            usize::try_from(entry.local_span().start).map_err(RepairError::Integer)?;
        let local_end = usize::try_from(entry.local_span().end).map_err(RepairError::Integer)?;
        let central = checked_range(entry.central_record(), source.len())?;
        if central.start < central_start || central.end > eocd || central.len() < CENTRAL_FIXED {
            return Err(RepairError::Unsupported {
                reason: "central-directory record range",
            });
        }
        if central.len() > limits.max_scratch_bytes {
            return Err(RepairError::Limit {
                resource: "central scratch bytes",
                observed: central.len() as u64,
                limit: limits.max_scratch_bytes as u64,
            });
        }
        let central_bytes = &source[central.clone()];
        if le_u32(central_bytes, 0) != Some(0x0201_4b50) {
            return Err(RepairError::Unsupported {
                reason: "malformed central-directory signature",
            });
        }
        if le_u16(central_bytes, 8).is_some_and(|flags| flags & 1 != 0) {
            return Err(RepairError::Unsupported {
                reason: "encrypted ZIP member",
            });
        }
        let name_len = usize::from(le_u16(central_bytes, 28).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory name length",
        })?);
        let extra_len = usize::from(le_u16(central_bytes, 30).ok_or(RepairError::Unsupported {
            reason: "malformed central-directory extra length",
        })?);
        let comment_len =
            usize::from(le_u16(central_bytes, 32).ok_or(RepairError::Unsupported {
                reason: "malformed central-directory comment length",
            })?);
        let central_variable = CENTRAL_FIXED
            .checked_add(name_len)
            .and_then(|value| value.checked_add(extra_len))
            .and_then(|value| value.checked_add(comment_len))
            .ok_or(RepairError::Unsupported {
                reason: "central-directory record length overflow",
            })?;
        if central_variable != central.len() {
            return Err(RepairError::Unsupported {
                reason: "central-directory record truncation or trailing bytes",
            });
        }
        let name = &central_bytes[CENTRAL_FIXED..CENTRAL_FIXED + name_len];
        if !names.insert(name) {
            return Err(RepairError::Unsupported {
                reason: "duplicate ZIP member name",
            });
        }
        let local_offset =
            usize::try_from(le_u32(central_bytes, 42).ok_or(RepairError::Unsupported {
                reason: "central local-header offset",
            })?)
            .map_err(RepairError::Integer)?;
        let compressed_size =
            usize::try_from(le_u32(central_bytes, 20).ok_or(RepairError::Unsupported {
                reason: "central compressed size",
            })?)
            .map_err(RepairError::Integer)?;
        let uncompressed_size =
            usize::try_from(le_u32(central_bytes, 24).ok_or(RepairError::Unsupported {
                reason: "central uncompressed size",
            })?)
            .map_err(RepairError::Integer)?;
        if (uncompressed_size as u64) > limits.max_member_bytes {
            return Err(RepairError::Limit {
                resource: "member bytes",
                observed: uncompressed_size as u64,
                limit: limits.max_member_bytes,
            });
        }
        total_uncompressed = total_uncompressed
            .checked_add(uncompressed_size as u64)
            .ok_or(RepairError::Unsupported {
                reason: "aggregate member size overflow",
            })?;
        if total_uncompressed > limits.max_total_member_bytes {
            return Err(RepairError::Limit {
                resource: "aggregate member bytes",
                observed: total_uncompressed,
                limit: limits.max_total_member_bytes,
            });
        }
        if local_start != local_offset || local_start >= local_end || local_end > central_start {
            return Err(RepairError::Unsupported {
                reason: "local-header offset or span mismatch",
            });
        }
        layouts.push(EntryLayout {
            local: local_start..local_end,
            central,
            local_offset,
            compressed_size,
            uncompressed_size,
        });
    }

    let mut local_order: Vec<usize> = (0..layouts.len()).collect();
    local_order.sort_unstable_by_key(|index| layouts[*index].local.start);
    if layouts[local_order[0]].local.start != 0
        || local_order
            .windows(2)
            .any(|pair| layouts[pair[0]].local.end != layouts[pair[1]].local.start)
        || layouts[local_order[local_order.len() - 1]].local.end != central_start
    {
        return Err(RepairError::Unsupported {
            reason: "local members are not exactly contiguous",
        });
    }
    let mut expected_central = central_start;
    for entry in &layouts {
        if entry.central.start != expected_central {
            return Err(RepairError::Unsupported {
                reason: "central records are not exactly contiguous",
            });
        }
        expected_central = entry.central.end;
    }
    if expected_central != eocd {
        return Err(RepairError::Unsupported {
            reason: "central records do not end at the EOCD",
        });
    }
    let eocd_len = source.len().saturating_sub(eocd);
    if eocd_len < EOCD_FIXED || eocd_len > limits.max_scratch_bytes {
        return Err(RepairError::Limit {
            resource: "EOCD scratch bytes",
            observed: eocd_len as u64,
            limit: limits.max_scratch_bytes as u64,
        });
    }

    // PreservationIndex deliberately permits data descriptors and opaque
    // bytes after untouched member payloads inside each raw local span. Those
    // bytes are copied as part of the span; only the first target member is
    // required to end exactly at its payload because its local header is the
    // sole transformed span.
    for &index in &local_order {
        let entry = &layouts[index];
        let local = &source[entry.local.clone()];
        if local.len() < LOCAL_FIXED || le_u32(local, 0) != Some(0x0403_4b50) {
            return Err(RepairError::Unsupported {
                reason: "malformed local member header",
            });
        }
        let flags = le_u16(local, 6).ok_or(RepairError::Unsupported {
            reason: "local member flags",
        })?;
        let compression = le_u16(local, 8).ok_or(RepairError::Unsupported {
            reason: "local member compression",
        })?;
        let central_bytes = &source[entry.central.clone()];
        let central_flags = le_u16(central_bytes, 8).ok_or(RepairError::Unsupported {
            reason: "central member flags",
        })?;
        let central_compression = le_u16(central_bytes, 10).ok_or(RepairError::Unsupported {
            reason: "central member compression",
        })?;
        if flags != central_flags || compression != central_compression || flags & !0x0808 != 0 {
            return Err(RepairError::Unsupported {
                reason: "encrypted, unknown, or mismatched local member flags/compression",
            });
        }
        let local_name_len = usize::from(le_u16(local, 26).ok_or(RepairError::Unsupported {
            reason: "local member name length",
        })?);
        let local_extra_len = usize::from(le_u16(local, 28).ok_or(RepairError::Unsupported {
            reason: "local member extra length",
        })?);
        let local_name_end =
            LOCAL_FIXED
                .checked_add(local_name_len)
                .ok_or(RepairError::Unsupported {
                    reason: "local member name range",
                })?;
        let local_header_end =
            local_name_end
                .checked_add(local_extra_len)
                .ok_or(RepairError::Unsupported {
                    reason: "local member header range",
                })?;
        let central_name_len =
            usize::from(le_u16(central_bytes, 28).ok_or(RepairError::Unsupported {
                reason: "central member name length",
            })?);
        if local_name_len != central_name_len
            || local_header_end > local.len()
            || local[LOCAL_FIXED..local_name_end]
                != central_bytes[CENTRAL_FIXED..CENTRAL_FIXED + central_name_len]
        {
            return Err(RepairError::Unsupported {
                reason: "local and central member names differ",
            });
        }
        let payload_end = local_header_end.checked_add(entry.compressed_size).ok_or(
            RepairError::Unsupported {
                reason: "local member payload range",
            },
        )?;
        if payload_end > local.len() {
            return Err(RepairError::Unsupported {
                reason: "local member payload is truncated",
            });
        }
        if flags & 0x08 == 0
            && (le_u32(local, 14) != le_u32(central_bytes, 16)
                || le_u32(local, 18) != le_u32(central_bytes, 20)
                || le_u32(local, 22) != le_u32(central_bytes, 24))
        {
            return Err(RepairError::Unsupported {
                reason: "local and central member sizes differ",
            });
        }
    }
    if !central_name_is(&source[layouts[0].central.clone()], MIMETYPE)
        || !central_name_is(&source[layouts[local_order[0]].central.clone()], MIMETYPE)
    {
        return Err(RepairError::Unsupported {
            reason: "mimetype is not first in both ZIP orders",
        });
    }

    let first = &layouts[local_order[0]];
    let local = &source[first.local.clone()];
    if local.len() < LOCAL_FIXED || le_u32(local, 0) != Some(0x0403_4b50) {
        return Err(RepairError::Unsupported {
            reason: "malformed mimetype local header",
        });
    }
    let flags = le_u16(local, 6).ok_or(RepairError::Unsupported {
        reason: "mimetype flags",
    })?;
    let compression = le_u16(local, 8).ok_or(RepairError::Unsupported {
        reason: "mimetype compression",
    })?;
    let local_name_len = usize::from(le_u16(local, 26).ok_or(RepairError::Unsupported {
        reason: "mimetype local name length",
    })?);
    let extra_len = usize::from(le_u16(local, 28).ok_or(RepairError::Unsupported {
        reason: "mimetype local extra length",
    })?);
    if flags & !0x0800 != 0 || compression != 0 || local_name_len != MIMETYPE.len() {
        return Err(RepairError::Unsupported {
            reason: "mimetype is encrypted, descriptor-backed, or not stored",
        });
    }
    let name_start = LOCAL_FIXED;
    let name_end = name_start
        .checked_add(local_name_len)
        .ok_or(RepairError::Unsupported {
            reason: "mimetype local name range",
        })?;
    let extra_start = name_end;
    let extra_end = extra_start
        .checked_add(extra_len)
        .ok_or(RepairError::Unsupported {
            reason: "mimetype local extra range",
        })?;
    let payload_end =
        extra_end
            .checked_add(first.compressed_size)
            .ok_or(RepairError::Unsupported {
                reason: "mimetype payload range",
            })?;
    if payload_end != local.len()
        || &local[name_start..name_end] != MIMETYPE
        || first.compressed_size != first.uncompressed_size
        || source[extra_start..extra_end].is_empty()
    {
        return Err(RepairError::Unsupported {
            reason: "mimetype local layout is malformed",
        });
    }
    let central_first = &source[first.central.clone()];
    if le_u16(central_first, 30) != Some(0) {
        return Err(RepairError::Unsupported {
            reason: "mimetype central extra field is not removable",
        });
    }
    if extra_len > limits.max_extra_bytes {
        return Err(RepairError::Limit {
            resource: "local extra bytes",
            observed: extra_len as u64,
            limit: limits.max_extra_bytes as u64,
        });
    }
    let mut fields = ExtraFields::new(&local[extra_start..extra_end]);
    let (field_id, body) = fields.next().ok_or(RepairError::Unsupported {
        reason: "mimetype local extra field is malformed",
    })?;
    if fields.next().is_some()
        || !fields.remaining_bytes().is_empty()
        || field_id != ExtraFieldId::EXTENDED_TIMESTAMP
        || !valid_extended_timestamp(body)
    {
        return Err(RepairError::Unsupported {
            reason: "mimetype local extra field is unknown, repeated, or malformed",
        });
    }
    let field_bytes = u16::try_from(extra_len).map_err(RepairError::Integer)?;
    let output_len = source_len
        .checked_sub(extra_len as u64)
        .ok_or(RepairError::Unsupported {
            reason: "repair output length underflow",
        })?;
    if output_len > limits.max_output_bytes {
        return Err(RepairError::Limit {
            resource: "output bytes",
            observed: output_len,
            limit: limits.max_output_bytes,
        });
    }
    scan_package_xml(source, limits)?;
    Ok(Layout {
        member_count: layouts.len(),
        non_directory_count: central_preflight.non_directory_count,
        entries: layouts,
        local_order,
        central_start,
        eocd,
        output_len,
        action: RemoveMimetypeLocalExtra {
            field_id: field_id.as_u16(),
            field_bytes,
        },
    })
}

/// Scan every admitted XML-family member before a candidate can be built or a
/// sink can be touched. The generic validator deliberately inspects only the
/// declaration and root XML; this repair emits the complete package, so its
/// authorization must also reject hostile XML anywhere else in the package.
fn scan_package_xml(source: &[u8], limits: OdfRepairLimits) -> Result<(), RepairError> {
    let archive = ArchiveReader::new_with_limits(source, archive_limits(limits))?;
    for name in archive.file_names() {
        if name != MANIFEST_PATH && is_xml_family_name(name) {
            scan_archive_xml_member(&archive, name, limits)?;
        }
    }

    let manifest = archive.read(MANIFEST_PATH)?;
    scan_xml_security(MANIFEST_PATH, &manifest, limits)?;
    scan_manifest_xml_types(&archive, &manifest, limits)
}

fn scan_archive_xml_member(
    archive: &ArchiveReader<'_>,
    name: &str,
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    let bytes = archive.read(name)?;
    if bytes.len() as u64 > limits.max_member_bytes {
        return Err(RepairError::Limit {
            resource: "XML member bytes",
            observed: bytes.len() as u64,
            limit: limits.max_member_bytes,
        });
    }
    scan_xml_security(name, &bytes, limits)
}

fn scan_manifest_xml_types(
    archive: &ArchiveReader<'_>,
    manifest: &[u8],
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    let mut reader = NsReader::from_reader(manifest);
    reader.config_mut().check_end_names = true;
    let validation_limits = OdfValidationLimits::default();
    let mut events = 0_usize;
    loop {
        if events >= validation_limits.max_xml_events() {
            return Err(RepairError::Limit {
                resource: "manifest XML events",
                observed: events as u64,
                limit: validation_limits.max_xml_events() as u64,
            });
        }
        let event = reader
            .read_resolved_event()
            .map_err(|_| RepairError::Unsupported {
                reason: "manifest XML is malformed during package-wide security scan",
            })?;
        events = events.saturating_add(1);
        let (_, event) = event;
        let element = match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"file-entry" =>
            {
                element
            },
            Event::Eof => return Ok(()),
            _ => continue,
        };
        let mut full_path = None;
        let mut media_type = None;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|_| RepairError::Unsupported {
                reason: "manifest XML attributes are malformed during package-wide security scan",
            })?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == crate::validation::MANIFEST_NAMESPACE)
            {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|_| RepairError::Unsupported {
                    reason: "manifest XML attribute value is malformed during package-wide security scan",
                })?;
            match local.as_ref() {
                b"full-path" => full_path = Some(value.into_owned()),
                b"media-type" => media_type = Some(value.into_owned()),
                _ => {},
            }
        }
        let Some(path) = full_path else {
            continue;
        };
        if path != "/" && !safe_manifest_member_path(&path) {
            return Err(RepairError::Unsupported {
                reason: "manifest XML type path is not a safe package path",
            });
        }
        let is_xml_type = media_type.as_deref().is_some_and(is_xml_media_type);
        if is_xml_type && archive.contains(&path) {
            scan_archive_xml_member(archive, &path, limits)?;
        }
    }
}

fn is_xml_family_name(name: &str) -> bool {
    name.rsplit_once('.').is_some_and(|(_, extension)| {
        extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("rdf")
    })
}

fn is_xml_media_type(media_type: &str) -> bool {
    let media_type = media_type.split(';').next().unwrap_or_default().trim();
    media_type.eq_ignore_ascii_case("text/xml")
        || media_type.eq_ignore_ascii_case("application/xml")
        || media_type
            .get(media_type.len().saturating_sub(4)..)
            .is_some_and(|suffix| suffix.eq_ignore_ascii_case("+xml"))
}

fn safe_manifest_member_path(path: &str) -> bool {
    !path.is_empty()
        && path != "/"
        && !path.starts_with('/')
        && !path.contains('\\')
        && path
            .split('/')
            .all(|segment| !segment.is_empty() && segment != "." && segment != "..")
}

fn scan_xml_security(name: &str, xml: &[u8], limits: OdfRepairLimits) -> Result<(), RepairError> {
    if xml.len() as u64 > limits.max_member_bytes {
        return Err(RepairError::Limit {
            resource: "XML member bytes",
            observed: xml.len() as u64,
            limit: limits.max_member_bytes,
        });
    }
    let validation_limits = OdfValidationLimits::default();
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut events = 0_usize;
    let mut depth = 0_usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    loop {
        if events >= validation_limits.max_xml_events() {
            return Err(RepairError::Limit {
                resource: "XML events",
                observed: events as u64,
                limit: validation_limits.max_xml_events() as u64,
            });
        }
        let event = reader.read_event().map_err(|_| RepairError::Unsupported {
            reason: "XML member is malformed during package-wide security scan",
        })?;
        events = events.saturating_add(1);
        match event {
            Event::Start(element) => {
                validate_element_qname(&reader, &element)?;
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(RepairError::Unsupported {
                            reason: "XML member has multiple roots",
                        });
                    }
                    root_seen = true;
                }
                depth = depth.saturating_add(1);
                if depth > validation_limits.max_xml_depth() {
                    return Err(RepairError::Limit {
                        resource: "XML depth",
                        observed: depth as u64,
                        limit: validation_limits.max_xml_depth() as u64,
                    });
                }
                scan_xml_link_attributes(&reader, &element, name)?;
            },
            Event::Empty(element) => {
                validate_element_qname(&reader, &element)?;
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(RepairError::Unsupported {
                            reason: "XML member has multiple roots",
                        });
                    }
                    root_seen = true;
                    root_closed = true;
                }
                scan_xml_link_attributes(&reader, &element, name)?;
            },
            Event::End(element) => {
                validate_xml_qname(reader.decoder(), element.name().as_ref(), "element")?;
                depth = depth.checked_sub(1).ok_or(RepairError::Unsupported {
                    reason: "XML member has an unmatched closing element",
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(RepairError::Unsupported {
                    reason: "XML member contains a DTD or entity declaration",
                });
            },
            Event::GeneralRef(reference)
                if depth > 0 && crate::validation::valid_xml_reference(&reference) => {},
            Event::GeneralRef(_) => {
                return Err(RepairError::Unsupported {
                    reason: "XML member contains an entity reference",
                });
            },
            Event::Eof if root_seen && root_closed && depth == 0 => return Ok(()),
            Event::Eof => {
                return Err(RepairError::Unsupported {
                    reason: "XML member is malformed during package-wide security scan",
                });
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(|_| RepairError::Unsupported {
                    reason: "XML member text encoding is unsupported or malformed",
                })?;
                validate_xml_string(&decoded)?;
                if depth == 0
                    && !decoded
                        .chars()
                        .all(|value| matches!(value, ' ' | '\t' | '\r' | '\n'))
                {
                    return Err(RepairError::Unsupported {
                        reason: "XML member has non-whitespace text outside its root",
                    });
                }
            },
            Event::CData(text) => {
                let decoded = text.decode().map_err(|_| RepairError::Unsupported {
                    reason: "XML member CDATA encoding is unsupported or malformed",
                })?;
                validate_xml_string(&decoded)?;
                if depth == 0 {
                    return Err(RepairError::Unsupported {
                        reason: "XML member has CDATA outside its root",
                    });
                }
            },
            Event::Comment(comment) => {
                let decoded = comment.decode().map_err(|_| RepairError::Unsupported {
                    reason: "XML member comment encoding is unsupported or malformed",
                })?;
                validate_xml_string(&decoded)?;
            },
            Event::PI(_) => {
                return Err(RepairError::Unsupported {
                    reason: "XML member contains an unsupported processing instruction",
                });
            },
            Event::Decl(declaration)
                if events == 1 && !declaration_seen && !root_seen && depth == 0 =>
            {
                validate_xml_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Decl(_) => {
                return Err(RepairError::Unsupported {
                    reason: "XML member has a misplaced or repeated declaration",
                });
            },
        }
    }
}

fn scan_xml_link_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    _name: &str,
) -> Result<(), RepairError> {
    let mut namespace_pool: Vec<Box<[u8]>> = Vec::new();
    let mut expanded_names: HashSet<(Option<usize>, Box<[u8]>)> = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|_| RepairError::Unsupported {
            reason: "XML member attribute syntax is malformed",
        })?;
        let raw_name = reader
            .decoder()
            .decode(attribute.key.as_ref())
            .map_err(|_| RepairError::Unsupported {
                reason: "XML member attribute name encoding is unsupported or malformed",
            })?;
        if raw_name == "xmlns" || raw_name.starts_with("xmlns:") {
            validate_namespace_declaration_name(&raw_name)?;
        } else {
            validate_xml_qname(reader.decoder(), attribute.key.as_ref(), "attribute")?;
            if matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            ) {
                return Err(RepairError::Unsupported {
                    reason: "XML member attribute uses an unbound namespace prefix",
                });
            }
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|_| RepairError::Unsupported {
                reason: "XML member attribute encoding is unsupported or malformed",
            })?;
        validate_xml_string(&value)?;
        if raw_name == "xmlns" || raw_name.starts_with("xmlns:") {
            validate_namespace_binding(&raw_name, &value)?;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let expanded_namespace = match &namespace {
            ResolveResult::Bound(Namespace(uri)) => {
                let namespace_attribute = Attribute {
                    key: QName(b"xmlns"),
                    value: Cow::Borrowed(uri),
                };
                let normalized = namespace_attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|_| RepairError::Unsupported {
                        reason: "XML member namespace encoding is unsupported or malformed",
                    })?;
                validate_xml_string(&normalized)?;
                if let Some(index) = namespace_pool
                    .iter()
                    .position(|known| known.as_ref() == normalized.as_bytes())
                {
                    Some(index)
                } else {
                    namespace_pool
                        .try_reserve(1)
                        .map_err(|_| RepairError::Allocation {
                            resource: "XML expanded attribute namespaces",
                        })?;
                    namespace_pool.push(try_boxed_copy(
                        normalized.as_bytes(),
                        "XML expanded attribute namespace",
                    )?);
                    Some(namespace_pool.len() - 1)
                }
            },
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(_) => {
                return Err(RepairError::Unsupported {
                    reason: "XML member attribute uses an unbound namespace prefix",
                });
            },
        };
        if raw_name != "xmlns" && !raw_name.starts_with("xmlns:") {
            expanded_names
                .try_reserve(1)
                .map_err(|_| RepairError::Allocation {
                    resource: "XML expanded attribute names",
                })?;
            let local = try_boxed_copy(local.as_ref(), "XML expanded attribute local name")?;
            if !expanded_names.insert((expanded_namespace, local)) {
                return Err(RepairError::Unsupported {
                    reason: "XML member has duplicate expanded attribute names",
                });
            }
        }
        if expanded_namespace.is_some_and(|index| {
            namespace_pool[index].as_ref() == b"http://www.w3.org/XML/1998/namespace"
        }) && local.as_ref() == b"base"
        {
            return Err(RepairError::Unsupported {
                reason: "XML member contains xml:base",
            });
        }
        if local.as_ref().starts_with(b"dde-") {
            return Err(RepairError::Unsupported {
                reason: "XML member contains an external DDE reference",
            });
        }
        if matches!(
            local.as_ref(),
            b"schemaLocation" | b"noNamespaceSchemaLocation" | b"codebase" | b"archive"
        ) {
            return Err(RepairError::Unsupported {
                reason: "XML member contains an external or unsafe link",
            });
        }
        if !matches!(
            local.as_ref(),
            b"href"
                | b"src"
                | b"url"
                | b"location"
                | b"target"
                | b"resource"
                | b"about"
                | b"source"
                | b"topic"
                | b"data-source"
                | b"connection-resource"
        ) {
            continue;
        }
        let value = trim_xml_whitespace(value.as_bytes());
        if value.is_empty() || value.starts_with(b"#") {
            continue;
        }
        if !crate::validation::safe_package_href(value) {
            return Err(RepairError::Unsupported {
                reason: "XML member contains an external or unsafe link",
            });
        }
    }
    Ok(())
}

fn try_boxed_copy(bytes: &[u8], resource: &'static str) -> Result<Box<[u8]>, RepairError> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|_| RepairError::Allocation { resource })?;
    copy.extend_from_slice(bytes);
    Ok(copy.into_boxed_slice())
}

fn validate_xml_declaration(declaration: &BytesDecl<'_>) -> Result<(), RepairError> {
    if declaration
        .xml_version()
        .map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration is malformed",
        })?
        != XmlVersion::Explicit1_0
    {
        return Err(RepairError::Unsupported {
            reason: "XML member declaration is not XML 1.0",
        });
    }
    let declaration_text =
        std::str::from_utf8(declaration.as_ref()).map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration encoding is malformed",
        })?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut state = 0_u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration is malformed",
        })?;
        if attribute.key.prefix().is_some() {
            return Err(RepairError::Unsupported {
                reason: "XML member declaration has a prefixed attribute",
            });
        }
        state = match (state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return Err(RepairError::Unsupported {
                    reason: "XML member declaration has duplicate, unknown, or out-of-order attributes",
                });
            },
        };
        std::str::from_utf8(attribute.value.as_ref()).map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration attribute encoding is malformed",
        })?;
    }
    if state == 0 {
        return Err(RepairError::Unsupported {
            reason: "XML member declaration lacks a version",
        });
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration encoding is malformed",
        })?;
        if !encoding.eq_ignore_ascii_case(b"UTF-8") {
            return Err(RepairError::Unsupported {
                reason: "XML member encoding is unsupported",
            });
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|_| RepairError::Unsupported {
            reason: "XML member declaration standalone value is malformed",
        })?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return Err(RepairError::Unsupported {
                reason: "XML member declaration standalone value is invalid",
            });
        }
    }
    Ok(())
}

fn validate_element_qname(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(), RepairError> {
    if element.name().local_name().as_ref().starts_with(b"dde-") {
        return Err(RepairError::Unsupported {
            reason: "XML member contains an active DDE element",
        });
    }
    if element
        .name()
        .prefix()
        .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
    {
        return Err(RepairError::Unsupported {
            reason: "XML member element uses the reserved xmlns prefix",
        });
    }
    validate_xml_qname(reader.decoder(), element.name().as_ref(), "element")?;
    if matches!(
        reader.resolver().resolve_element(element.name()),
        (ResolveResult::Unknown(_), _)
    ) {
        return Err(RepairError::Unsupported {
            reason: "XML member element uses an unbound namespace prefix",
        });
    }
    Ok(())
}

fn validate_xml_qname(
    decoder: quick_xml::encoding::Decoder,
    bytes: &[u8],
    kind: &'static str,
) -> Result<(), RepairError> {
    let decoded = decoder
        .decode(bytes)
        .map_err(|_| RepairError::Unsupported {
            reason: "XML member name encoding is unsupported or malformed",
        })?;
    if is_xml_qname(&decoded) {
        Ok(())
    } else {
        Err(RepairError::Unsupported {
            reason: match kind {
                "attribute" => "XML member attribute name is not a valid QName",
                _ => "XML member element name is not a valid QName",
            },
        })
    }
}

fn validate_namespace_declaration_name(value: &str) -> Result<(), RepairError> {
    if value == "xmlns" || value.strip_prefix("xmlns:").is_some_and(is_xml_ncname) {
        Ok(())
    } else {
        Err(RepairError::Unsupported {
            reason: "XML member namespace declaration name is invalid",
        })
    }
}

fn validate_namespace_binding(name: &str, value: &str) -> Result<(), RepairError> {
    const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
    const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
    let prefix = name.strip_prefix("xmlns:");
    if prefix == Some("xmlns")
        || value == XMLNS_NAMESPACE
        || (prefix == Some("xml")) != (value == XML_NAMESPACE)
        || (prefix.is_some() && value.is_empty())
    {
        return Err(RepairError::Unsupported {
            reason: "XML member has an invalid reserved namespace binding",
        });
    }
    Ok(())
}

fn is_xml_qname(value: &str) -> bool {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return false;
    };
    is_xml_ncname(first) && parts.next().is_none_or(is_xml_ncname) && parts.next().is_none()
}

fn is_xml_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(xml_ncname_start) && characters.all(xml_ncname_character)
}

fn xml_ncname_start(value: char) -> bool {
    matches!(
        value,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{c0}'..='\u{d6}'
            | '\u{d8}'..='\u{f6}'
            | '\u{f8}'..='\u{2ff}'
            | '\u{370}'..='\u{37d}'
            | '\u{37f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn xml_ncname_character(value: char) -> bool {
    xml_ncname_start(value)
        || matches!(
            value,
            '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}'
        )
}

fn validate_xml_string(value: &str) -> Result<(), RepairError> {
    if value.chars().all(xml10_character) {
        Ok(())
    } else {
        Err(RepairError::Unsupported {
            reason: "XML member contains a forbidden XML 1.0 character",
        })
    }
}

fn xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&value)
        || ('\u{e000}'..='\u{fffd}').contains(&value)
        || ('\u{10000}'..='\u{10ffff}').contains(&value)
}

fn trim_xml_whitespace(mut value: &[u8]) -> &[u8] {
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

fn build_candidate(
    source: &[u8],
    layout: &Layout,
    action: RemoveMimetypeLocalExtra,
    limits: OdfRepairLimits,
) -> Result<Vec<u8>, RepairError> {
    let output_len = usize::try_from(layout.output_len).map_err(RepairError::Integer)?;
    if layout.output_len > limits.max_preflight_candidate_bytes {
        return Err(RepairError::Limit {
            resource: "preflight candidate bytes",
            observed: layout.output_len,
            limit: limits.max_preflight_candidate_bytes,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| RepairError::Allocation {
            resource: "repair preflight candidate",
        })?;
    let first_index = layout.local_order[0];
    let first = &layout.entries[first_index];
    let local = &source[first.local.clone()];
    let local_name_len = usize::from(le_u16(local, 26).ok_or(RepairError::Unsupported {
        reason: "mimetype local name length",
    })?);
    let extra_len = usize::from(le_u16(local, 28).ok_or(RepairError::Unsupported {
        reason: "mimetype local extra length",
    })?);
    output.extend_from_slice(&local[..28]);
    output.extend_from_slice(&[0, 0]);
    let name_end = LOCAL_FIXED + local_name_len;
    let payload_start = name_end
        .checked_add(extra_len)
        .ok_or(RepairError::Unsupported {
            reason: "mimetype payload start",
        })?;
    output.extend_from_slice(&local[LOCAL_FIXED..name_end]);
    output.extend_from_slice(&local[payload_start..]);
    for &index in layout.local_order.iter().skip(1) {
        output.extend_from_slice(&source[layout.entries[index].local.clone()]);
    }
    for entry in &layout.entries {
        if entry.central.len() > limits.max_scratch_bytes {
            return Err(RepairError::Limit {
                resource: "central scratch bytes",
                observed: entry.central.len() as u64,
                limit: limits.max_scratch_bytes as u64,
            });
        }
        let mut central = source[entry.central.clone()].to_vec();
        let new_offset = if entry.local_offset == 0 {
            0
        } else {
            entry
                .local_offset
                .checked_sub(usize::from(action.field_bytes))
                .ok_or(RepairError::Unsupported {
                    reason: "central local offset underflow",
                })?
        };
        central[LOCAL_OFFSET].copy_from_slice(
            &u32::try_from(new_offset)
                .map_err(RepairError::Integer)?
                .to_le_bytes(),
        );
        output.extend_from_slice(&central);
    }
    let eocd_len = source.len().saturating_sub(layout.eocd);
    if eocd_len < EOCD_FIXED || eocd_len > limits.max_scratch_bytes {
        return Err(RepairError::Limit {
            resource: "EOCD scratch bytes",
            observed: eocd_len as u64,
            limit: limits.max_scratch_bytes as u64,
        });
    }
    let mut eocd = source[layout.eocd..].to_vec();
    let new_central = layout
        .central_start
        .checked_sub(usize::from(action.field_bytes))
        .ok_or(RepairError::Unsupported {
            reason: "central-directory offset underflow",
        })?;
    eocd[16..20].copy_from_slice(
        &u32::try_from(new_central)
            .map_err(RepairError::Integer)?
            .to_le_bytes(),
    );
    output.extend_from_slice(&eocd);
    if output.len() != output_len {
        return Err(RepairError::Unsupported {
            reason: "candidate output length mismatch",
        });
    }
    Ok(output)
}

fn verify_candidate(
    source: &[u8],
    candidate: &[u8],
    layout: &Layout,
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    if candidate.len() as u64 != layout.output_len {
        return Err(RepairError::TargetChanged {
            expected: RepairFingerprint::of(&candidate[..candidate.len().min(1)]),
            observed: RepairFingerprint::of(candidate),
        });
    }
    let candidate_report = validate_package_with_limits(
        candidate,
        validation_limits(limits, limits.max_output_bytes),
        litchi_core::ValidationLimits::default(),
    )?;
    if !candidate_report.is_complete() || !candidate_report.issues().is_empty() {
        return Err(RepairError::Unsupported {
            reason: "repaired candidate did not pass complete ODF validation",
        });
    }
    verify_raw_preservation(source, candidate, layout)?;
    verify_member_digests(source, candidate, limits)
}

fn verify_streaming_proof(
    source: &[u8],
    layout: &Layout,
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    // Re-run the bounded ODF checks against the source. Its exact fingerprint
    // was checked immediately before this function, so the already-proved
    // report authorization still applies without retaining report data in the
    // borrowed plan.
    let report = validate_package_with_limits(
        source,
        validation_limits(limits, limits.max_input_bytes),
        litchi_core::ValidationLimits::default(),
    )?;
    if !report.is_complete()
        || report.issues().len() != 1
        || report.issues()[0].check().as_str() != DECLARATION_CHECK
        || report.issues()[0].code() != MIMETYPE_LOCAL_EXTRA_ISSUE
        || report.issues()[0].repair().repair_id() != Some(MIMETYPE_LOCAL_EXTRA_REPAIR)
    {
        return Err(RepairError::Unsupported {
            reason: "source no longer passes the authorized ODF repair validation",
        });
    }

    // The plan-time candidate proof compared decompressed member digests.
    // Revalidation has the identical source fingerprint, and the exact local
    // span proof below establishes that every member payload emitted at
    // publication is the same source byte range. Reading the source members
    // once here catches any decompression/catalog failure without allocating a
    // second full target archive.
    verify_source_member_digests(source, layout.non_directory_count, limits)?;
    verify_raw_preservation_layout(source, layout)
}

fn verify_source_member_digests(
    source: &[u8],
    expected_members: usize,
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    let archive = ArchiveReader::new_with_limits(source, archive_limits(limits))?;
    let mut member_count = 0_usize;
    for name in archive.file_names() {
        member_count = member_count.saturating_add(1);
        if member_count > limits.max_members {
            return Err(RepairError::Limit {
                resource: "member count",
                observed: member_count as u64,
                limit: limits.max_members as u64,
            });
        }
        let bytes = archive.read(name)?;
        if bytes.len() as u64 > limits.max_member_bytes {
            return Err(RepairError::Limit {
                resource: "member bytes",
                observed: bytes.len() as u64,
                limit: limits.max_member_bytes,
            });
        }
        // Force a complete bounded read while retaining only the digest state.
        let _ = Sha256::digest(&bytes);
    }
    if member_count != expected_members {
        return Err(RepairError::MemberMismatch {
            name: "member catalog",
        });
    }
    Ok(())
}

fn verify_raw_preservation_layout(source: &[u8], layout: &Layout) -> Result<(), RepairError> {
    let delta = usize::from(layout.action.field_bytes);
    if layout.local_order.is_empty()
        || layout.entries[layout.local_order[0]].local.start != 0
        || layout.entries[layout.local_order[0]].local.end > layout.central_start
        || layout.output_len != u64::try_from(source.len().saturating_sub(delta))?
    {
        return Err(RepairError::Unsupported {
            reason: "local preservation layout changed",
        });
    }
    for pair in layout.local_order.windows(2) {
        if layout.entries[pair[0]].local.end != layout.entries[pair[1]].local.start {
            return Err(RepairError::Unsupported {
                reason: "local members are not physically contiguous",
            });
        }
    }
    if layout.entries[*layout.local_order.last().unwrap()]
        .local
        .end
        != layout.central_start
    {
        return Err(RepairError::Unsupported {
            reason: "local members do not end at the central directory",
        });
    }
    let mut central_cursor = layout.central_start;
    for entry in &layout.entries {
        if entry.central.start != central_cursor {
            return Err(RepairError::Unsupported {
                reason: "central records are not physically contiguous",
            });
        }
        central_cursor = entry.central.end;
    }
    if central_cursor != layout.eocd {
        return Err(RepairError::Unsupported {
            reason: "central records do not end at the EOCD",
        });
    }
    Ok(())
}

fn verify_raw_preservation(
    source: &[u8],
    candidate: &[u8],
    layout: &Layout,
) -> Result<(), RepairError> {
    let delta = usize::from(layout.action.field_bytes);
    let first_index = layout.local_order[0];
    let first = &layout.entries[first_index];
    let source_local = &source[first.local.clone()];
    let candidate_first_end = first.local.end - delta;
    let candidate_first = candidate
        .get(..candidate_first_end)
        .ok_or(RepairError::Unsupported {
            reason: "candidate mimetype local range",
        })?;
    if candidate_first.len() < LOCAL_FIXED
        || source_local[..28] != candidate_first[..28]
        || candidate_first[28..30] != [0, 0]
        || source_local[30..30 + MIMETYPE.len()] != candidate_first[30..30 + MIMETYPE.len()]
        || source_local[30 + MIMETYPE.len() + delta..] != candidate_first[30 + MIMETYPE.len()..]
    {
        return Err(RepairError::MemberMismatch {
            name: "mimetype local header",
        });
    }
    for &index in layout.local_order.iter().skip(1) {
        let entry = &layout.entries[index];
        let start = entry.local.start - delta;
        let end = entry.local.end - delta;
        if candidate.get(start..end) != Some(&source[entry.local.clone()]) {
            return Err(RepairError::MemberMismatch {
                name: "untouched local member",
            });
        }
    }
    let central_start = layout.central_start - delta;
    let mut central_offset = central_start;
    for entry in &layout.entries {
        let length = entry.central.len();
        let candidate_record = candidate
            .get(central_offset..central_offset + length)
            .ok_or(RepairError::MemberMismatch {
                name: "central-directory record",
            })?;
        let source_record = &source[entry.central.clone()];
        if candidate_record.len() != source_record.len()
            || candidate_record[..LOCAL_OFFSET.start] != source_record[..LOCAL_OFFSET.start]
            || candidate_record[LOCAL_OFFSET.end..] != source_record[LOCAL_OFFSET.end..]
        {
            return Err(RepairError::MemberMismatch {
                name: "central-directory record",
            });
        }
        central_offset += length;
    }
    let source_eocd = &source[layout.eocd..];
    let candidate_eocd_start = layout.eocd - delta;
    let candidate_eocd = candidate
        .get(candidate_eocd_start..)
        .ok_or(RepairError::MemberMismatch { name: "EOCD" })?;
    if source_eocd.len() != candidate_eocd.len()
        || source_eocd[..16] != candidate_eocd[..16]
        || source_eocd[20..] != candidate_eocd[20..]
    {
        return Err(RepairError::MemberMismatch { name: "EOCD" });
    }
    Ok(())
}

fn verify_member_digests(
    source: &[u8],
    candidate: &[u8],
    limits: OdfRepairLimits,
) -> Result<(), RepairError> {
    let archive_limits = archive_limits(limits);
    let source_archive = ArchiveReader::new_with_limits(source, archive_limits)?;
    let candidate_archive = ArchiveReader::new_with_limits(candidate, archive_limits)?;
    let source_names: Vec<&str> = source_archive.file_names().collect();
    let candidate_names: Vec<&str> = candidate_archive.file_names().collect();
    if source_names != candidate_names || source_names.len() > limits.max_members {
        return Err(RepairError::MemberMismatch {
            name: "member catalog",
        });
    }
    for name in source_names {
        let source_bytes = source_archive.read(name)?;
        let candidate_bytes = candidate_archive.read(name)?;
        if source_bytes.len() as u64 > limits.max_member_bytes
            || candidate_bytes.len() as u64 > limits.max_member_bytes
            || Sha256::digest(&source_bytes) != Sha256::digest(&candidate_bytes)
        {
            return Err(RepairError::MemberMismatch {
                name: "member payload",
            });
        }
    }
    Ok(())
}

fn emit_output<W: Write>(
    source: &[u8],
    layout: &Layout,
    action: RemoveMimetypeLocalExtra,
    sink: &mut W,
    accepted: &mut u64,
    target_hasher: &mut Sha256,
) -> Result<(), RepairError> {
    let delta = usize::from(action.field_bytes);
    let first_index = layout.local_order[0];
    let first = &layout.entries[first_index];
    let local = &source[first.local.clone()];
    let name_len = usize::from(le_u16(local, 26).ok_or(RepairError::Unsupported {
        reason: "mimetype local name length",
    })?);
    let extra_len = usize::from(le_u16(local, 28).ok_or(RepairError::Unsupported {
        reason: "mimetype local extra length",
    })?);
    let name_end = LOCAL_FIXED + name_len;
    let payload_start = name_end
        .checked_add(extra_len)
        .ok_or(RepairError::Unsupported {
            reason: "mimetype payload start",
        })?;
    let mut fixed = [0_u8; LOCAL_FIXED];
    fixed.copy_from_slice(&local[..LOCAL_FIXED]);
    fixed[28..30].copy_from_slice(&[0, 0]);
    write_piece(&fixed[..28], sink, accepted, target_hasher)?;
    write_piece(&fixed[28..], sink, accepted, target_hasher)?;
    write_piece(&local[LOCAL_FIXED..name_end], sink, accepted, target_hasher)?;
    write_piece(&local[payload_start..], sink, accepted, target_hasher)?;
    for &index in layout.local_order.iter().skip(1) {
        write_piece(
            &source[layout.entries[index].local.clone()],
            sink,
            accepted,
            target_hasher,
        )?;
    }
    for entry in &layout.entries {
        let mut central = source[entry.central.clone()].to_vec();
        let new_offset = if entry.local_offset == 0 {
            0
        } else {
            entry
                .local_offset
                .checked_sub(delta)
                .ok_or(RepairError::Unsupported {
                    reason: "central local offset underflow",
                })?
        };
        central[LOCAL_OFFSET].copy_from_slice(
            &u32::try_from(new_offset)
                .map_err(RepairError::Integer)?
                .to_le_bytes(),
        );
        write_piece(&central, sink, accepted, target_hasher)?;
    }
    let new_central = layout
        .central_start
        .checked_sub(delta)
        .ok_or(RepairError::Unsupported {
            reason: "central-directory offset underflow",
        })?;
    write_piece(
        &source[layout.eocd..layout.eocd + 16],
        sink,
        accepted,
        target_hasher,
    )?;
    write_piece(
        &u32::try_from(new_central)
            .map_err(RepairError::Integer)?
            .to_le_bytes(),
        sink,
        accepted,
        target_hasher,
    )?;
    write_piece(&source[layout.eocd + 20..], sink, accepted, target_hasher)?;
    Ok(())
}

fn write_piece<W: Write>(
    mut bytes: &[u8],
    sink: &mut W,
    accepted: &mut u64,
    target_hasher: &mut Sha256,
) -> Result<(), RepairError> {
    target_hasher.update(bytes);
    while !bytes.is_empty() {
        match sink.write(bytes) {
            Ok(0) => {
                return Err(RepairError::Io(io::Error::new(
                    io::ErrorKind::WriteZero,
                    "ODF repair sink accepted no bytes",
                )));
            },
            Ok(written) if written > bytes.len() => {
                return Err(RepairError::SinkOverreported);
            },
            Ok(written) => {
                *accepted = accepted.saturating_add(written as u64);
                bytes = &bytes[written..];
            },
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => return Err(RepairError::Io(error)),
        }
    }
    Ok(())
}

fn with_progress(
    error: RepairError,
    accepted: u64,
    expected: u64,
    indeterminate: bool,
) -> RepairError {
    if accepted == 0 && !indeterminate {
        return error;
    }
    let progress = if indeterminate {
        OutputProgress::Indeterminate {
            accepted_before: accepted,
        }
    } else if accepted == expected {
        OutputProgress::CompleteUnverified { bytes: accepted }
    } else {
        OutputProgress::Prefix { accepted, expected }
    };
    RepairError::IncompleteOutput {
        progress,
        source: Box::new(error),
    }
}

fn repair_metadata_bytes(
    member_count: usize,
    raw_name_bytes: u64,
    central_record_bytes: u64,
    archive_comment_bytes: u64,
) -> Result<u64, RepairError> {
    u64::try_from(member_count)
        .map_err(RepairError::Integer)?
        .checked_mul(
            NAME_SET_ENTRY_ESTIMATE
                + LAYOUT_ENTRY_ESTIMATE
                + PRESERVATION_ENTRY_ESTIMATE
                + LOCAL_ORDER_ENTRY_ESTIMATE,
        )
        .and_then(|fixed| fixed.checked_add(raw_name_bytes))
        .and_then(|observed| observed.checked_add(central_record_bytes))
        .and_then(|observed| observed.checked_add(archive_comment_bytes))
        .ok_or(RepairError::Unsupported {
            reason: "repair metadata estimate overflow",
        })
}

fn check_repair_metadata(
    member_count: usize,
    raw_name_bytes: u64,
    central_record_bytes: u64,
    archive_comment_bytes: u64,
    limits: OdfRepairLimits,
) -> Result<u64, RepairError> {
    let observed = repair_metadata_bytes(
        member_count,
        raw_name_bytes,
        central_record_bytes,
        archive_comment_bytes,
    )?;
    if observed > limits.max_metadata_bytes {
        return Err(RepairError::Limit {
            resource: "repair metadata bytes",
            observed,
            limit: limits.max_metadata_bytes,
        });
    }
    Ok(observed)
}

fn checked_range(range: Range<u64>, source_len: usize) -> Result<Range<usize>, RepairError> {
    let start = usize::try_from(range.start).map_err(RepairError::Integer)?;
    let end = usize::try_from(range.end).map_err(RepairError::Integer)?;
    if start > end || end > source_len {
        return Err(RepairError::Unsupported {
            reason: "source range outside input",
        });
    }
    Ok(start..end)
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    bytes
        .get(offset..offset.checked_add(2)?)
        .and_then(|slice| slice.try_into().ok())
        .map(u16::from_le_bytes)
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    bytes
        .get(offset..offset.checked_add(4)?)
        .and_then(|slice| slice.try_into().ok())
        .map(u32::from_le_bytes)
}

fn valid_extended_timestamp(body: &[u8]) -> bool {
    let Some(&flags) = body.first() else {
        return false;
    };
    if flags & !0b111 != 0 {
        return false;
    }
    body.len() == 1 + flags.count_ones() as usize * 4
}

fn central_name_is(record: &[u8], expected: &[u8]) -> bool {
    let Some(name_len) = le_u16(record, 28).map(usize::from) else {
        return false;
    };
    let Some(name_end) = CENTRAL_FIXED
        .checked_add(name_len)
        .filter(|end| *end <= record.len())
    else {
        return false;
    };
    &record[CENTRAL_FIXED..name_end] == expected
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn repair_metadata_budget_accepts_exact_and_rejects_one_over() {
        let exact = repair_metadata_bytes(3, 17, 128, 4).unwrap();
        let limits = OdfRepairLimits::default().with_max_metadata_bytes(exact);
        assert_eq!(check_repair_metadata(3, 17, 128, 4, limits).unwrap(), exact);
        let one_over = repair_metadata_bytes(3, 18, 128, 4).unwrap();
        assert!(matches!(
            check_repair_metadata(
                3,
                18,
                128,
                4,
                OdfRepairLimits::default().with_max_metadata_bytes(exact)
            ),
            Err(RepairError::Limit {
                resource: "repair metadata bytes",
                observed,
                limit
            }) if observed == one_over && limit == exact
        ));
    }
}
