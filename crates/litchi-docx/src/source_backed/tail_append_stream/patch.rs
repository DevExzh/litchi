//! Compact durable proof patches for replayable DOCX paragraph appends.
//!
//! The wire representation contains only bounded scalar proofs, selected
//! limits, and explicit caller/provider tokens. It never retains source,
//! candidate, or generated XML.

use super::tail_append;
use super::{AuthoredReplayHandle, AuthoredReplayReference, AuthoredStreamProof};
use litchi_opc::{
    SourceArtifact, SourceArtifactFingerprint, SourceArtifactRestoreProof, SourceBackedPackage,
};
use serde_json::{Map, Number, Value};
use std::fmt;
use std::io::{self, Write};
use std::sync::Arc;
use thiserror::Error;

/// Absolute input ceiling checked before parsing an untrusted patch.
pub const ABSOLUTE_MAX_PATCH_BYTES: usize = 4 * 1024 * 1024;

const OPERATION: &str = "docx.tail_append_plain_paragraphs";
const INVERSE_OPERATION: &str = "docx.tail_append_exact_inverse";
const VERSION: u64 = 1;
const MAX_ARTIFACT_PROOF_BYTES: u64 = u64::MAX - 1;
const MAX_ORIGINAL_REFERENCE_BYTES: u64 = ABSOLUTE_MAX_PATCH_BYTES as u64;
const MAX_MAIN_PART_BYTES: u64 = 4 * 1024;

/// Result type for this module.
pub type Result<T> = std::result::Result<T, PatchError>;

/// Length/hash identity of a complete physical archive.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ArtifactProof {
    /// Exact archive byte length.
    pub len: u64,
    /// SHA-256 over every archive byte.
    pub sha256: [u8; 32],
}

impl ArtifactProof {
    /// Construct an artifact identity after authentication by the caller.
    #[must_use]
    pub const fn new(len: u64, sha256: [u8; 32]) -> Self {
        Self { len, sha256 }
    }

    fn validate(self, resource: &'static str, maximum: u64) -> Result<()> {
        if self.len == 0 {
            return Err(PatchError::InvalidFacts("artifact length must be nonzero"));
        }
        if self.len > maximum {
            return Err(PatchError::Limit {
                resource,
                actual: self.len,
                maximum,
            });
        }
        Ok(())
    }
}

/// Source semantic facts retained by one stream patch.
///
/// The process-local source version is intentionally absent. Complete raw
/// archive and main-member identities are the durable authorization.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StreamSourceProof {
    /// Canonical main-part URI.
    pub main_part: String,
    /// Decoded source main-member length.
    pub source_len: u64,
    /// SHA-256 over decoded source main-member bytes.
    pub source_sha256: [u8; 32],
    /// Decoded insertion offset before section properties or body close.
    pub insertion_offset: u64,
    /// Direct source paragraph count.
    pub paragraph_count: u64,
    /// Source parser event count.
    pub event_count: u64,
    /// Maximum source parser depth.
    pub max_depth: u64,
    /// Whether Strict WordprocessingML is used.
    pub strict_namespace: bool,
    /// Opaque direct section-properties length.
    pub sect_pr_len: u64,
    /// SHA-256 over the exact section-properties span.
    pub sect_pr_sha256: [u8; 32],
    /// Complete source archive identity.
    pub archive: ArtifactProof,
}

/// Candidate semantic facts retained by one stream patch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct StreamCandidateProof {
    /// Decoded candidate main-member length.
    pub candidate_len: u64,
    /// SHA-256 over decoded candidate main-member bytes.
    pub candidate_sha256: [u8; 32],
    /// Direct candidate paragraph count.
    pub paragraph_count: u64,
    /// Candidate parser event count.
    pub event_count: u64,
    /// Maximum candidate parser depth.
    pub max_depth: u64,
    /// Offset at which the generated range starts.
    pub generated_offset: u64,
    /// Number of generated encoded bytes.
    pub generated_len: u64,
    /// Number of generated direct paragraphs.
    pub generated_paragraph_count: u64,
    /// Whether the generated range occurred exactly once.
    pub generated_once: bool,
    /// Candidate section-properties length.
    pub sect_pr_len: u64,
    /// SHA-256 over the candidate section-properties span.
    pub sect_pr_sha256: [u8; 32],
    /// Complete candidate archive identity.
    pub archive: ArtifactProof,
}

/// Wire-owned copy of all finite limits selected for a stream operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PatchLimits {
    /// Maximum decoded source XML bytes.
    pub max_source_xml_bytes: u64,
    /// Maximum authored UTF-8 bytes.
    pub max_text_bytes: u64,
    /// Maximum generated authored XML bytes.
    pub max_fragment_bytes: u64,
    /// Maximum decoded candidate XML bytes.
    pub max_candidate_xml_bytes: u64,
    /// Maximum parser events.
    pub max_events: u64,
    /// Maximum parser depth.
    pub max_depth: u64,
    /// Maximum source direct paragraphs.
    pub max_paragraphs: u64,
    /// Maximum settings XML bytes.
    pub max_settings_xml_bytes: u64,
    /// Maximum semantic workspace bytes.
    pub max_workspace_bytes: u64,
    /// Maximum physical output bytes.
    pub max_output_bytes: u64,
    /// Maximum XML token bytes.
    pub max_token_bytes: u64,
    /// Maximum authored paragraphs.
    pub max_authored_paragraphs: u64,
    /// Maximum authored events.
    pub max_authored_events: u64,
    /// Maximum one authored text chunk.
    pub max_authored_chunk_bytes: u64,
    /// Maximum authored input text bytes.
    pub max_authored_text_bytes: u64,
    /// Maximum authored generated XML bytes.
    pub max_authored_xml_bytes: u64,
    /// Maximum retained replay bytes.
    pub max_replay_bytes: u64,
    /// Maximum one-reader replay window bytes.
    pub max_replay_window_bytes: u64,
    /// Maximum canonical patch wire bytes.
    pub max_patch_bytes: u64,
}

impl PatchLimits {
    /// Copy the public source-backed stream limits into the durable wire
    /// shape. The copy contains no execution context or runtime handle.
    #[must_use]
    pub const fn from_stream_limits(limits: &super::ParagraphStreamLimits) -> Self {
        Self {
            max_source_xml_bytes: limits.source.max_source_xml_bytes,
            max_text_bytes: limits.source.max_text_bytes,
            max_fragment_bytes: limits.source.max_fragment_bytes,
            max_candidate_xml_bytes: limits.source.max_candidate_xml_bytes,
            max_events: limits.source.max_events,
            max_depth: limits.source.max_depth,
            max_paragraphs: limits.source.max_paragraphs,
            max_settings_xml_bytes: limits.source.max_settings_xml_bytes,
            max_workspace_bytes: limits.source.max_workspace_bytes,
            max_output_bytes: limits.source.max_output_bytes,
            max_token_bytes: limits.source.max_token_bytes,
            max_authored_paragraphs: limits.max_authored_paragraphs,
            max_authored_events: limits.max_authored_events,
            max_authored_chunk_bytes: limits.max_authored_chunk_bytes,
            max_authored_text_bytes: limits.max_authored_text_bytes,
            max_authored_xml_bytes: limits.max_authored_xml_bytes,
            max_replay_bytes: limits.max_replay_bytes,
            max_replay_window_bytes: limits.max_replay_window_bytes,
            max_patch_bytes: limits.max_patch_bytes,
        }
    }

    /// Validate all finite nonzero ceilings.
    pub fn validate(self) -> Result<()> {
        let values = [
            ("source XML bytes", self.max_source_xml_bytes),
            ("text bytes", self.max_text_bytes),
            ("fragment bytes", self.max_fragment_bytes),
            ("candidate XML bytes", self.max_candidate_xml_bytes),
            ("XML events", self.max_events),
            ("XML depth", self.max_depth),
            ("paragraphs", self.max_paragraphs),
            ("settings XML bytes", self.max_settings_xml_bytes),
            ("parser workspace", self.max_workspace_bytes),
            ("output bytes", self.max_output_bytes),
            ("XML token bytes", self.max_token_bytes),
            ("authored paragraphs", self.max_authored_paragraphs),
            ("authored events", self.max_authored_events),
            ("authored chunk bytes", self.max_authored_chunk_bytes),
            ("authored text bytes", self.max_authored_text_bytes),
            ("authored XML bytes", self.max_authored_xml_bytes),
            ("replay bytes", self.max_replay_bytes),
            ("replay window bytes", self.max_replay_window_bytes),
            ("patch bytes", self.max_patch_bytes),
        ];
        for (resource, value) in values {
            if value == 0 || value == u64::MAX {
                return Err(PatchError::Limit {
                    resource,
                    actual: value,
                    maximum: value.saturating_sub(1),
                });
            }
        }
        if self.max_patch_bytes > ABSOLUTE_MAX_PATCH_BYTES as u64 {
            return Err(PatchError::Limit {
                resource: "patch bytes",
                actual: self.max_patch_bytes,
                maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
            });
        }
        Ok(())
    }
}

/// Compact authenticated recipe for one replayable authored append.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphStreamPatch {
    limits: PatchLimits,
    source: StreamSourceProof,
    authored: AuthoredStreamProof,
    candidate: StreamCandidateProof,
    replay_reference: Option<AuthoredReplayReference>,
}

/// Short transaction-family name.
pub type Patch = ParagraphStreamPatch;

impl ParagraphStreamPatch {
    /// Build a patch from the semantic proofs emitted by the source-backed
    /// scanner. The process-local source version is copied only long enough
    /// to be deliberately discarded; it is never present in this patch.
    pub fn from_tail_append_proofs(
        limits: PatchLimits,
        main_part: impl Into<String>,
        source: tail_append::SourceProof,
        authored: AuthoredStreamProof,
        candidate: tail_append::CandidateProof,
        source_archive: ArtifactProof,
        candidate_archive: ArtifactProof,
        replay_reference: Option<AuthoredReplayReference>,
    ) -> Result<Self> {
        let source = StreamSourceProof {
            main_part: main_part.into(),
            source_len: source.source_len,
            source_sha256: source.source_sha256,
            insertion_offset: source.insertion_offset,
            paragraph_count: source.paragraph_count,
            event_count: source.event_count,
            max_depth: source.max_depth,
            strict_namespace: source.strict_namespace,
            sect_pr_len: source.sect_pr_len,
            sect_pr_sha256: source.sect_pr_sha256,
            archive: source_archive,
        };
        let candidate = StreamCandidateProof {
            candidate_len: candidate.candidate_len,
            candidate_sha256: candidate.candidate_sha256,
            paragraph_count: candidate.paragraph_count,
            event_count: candidate.event_count,
            max_depth: candidate.max_depth,
            generated_offset: candidate.generated_offset,
            generated_len: authored.encoded_xml_bytes,
            generated_paragraph_count: authored.paragraph_count,
            generated_once: candidate.generated_once,
            sect_pr_len: candidate.sect_pr_len,
            sect_pr_sha256: candidate.sect_pr_sha256,
            archive: candidate_archive,
        };
        Self::new(limits, source, authored, candidate, replay_reference)
    }

    /// Construct a live patch from compact semantic and physical proofs.
    ///
    /// A missing replay reference is allowed for process-local use but makes
    /// durable serialization refuse with NotDurable.
    pub fn new(
        limits: PatchLimits,
        source: StreamSourceProof,
        authored: AuthoredStreamProof,
        candidate: StreamCandidateProof,
        replay_reference: Option<AuthoredReplayReference>,
    ) -> Result<Self> {
        limits.validate()?;
        validate_facts(&limits, &source, authored, &candidate)?;
        if let Some(reference) = replay_reference.as_ref() {
            validate_reference(reference.as_bytes(), limits)?;
        }
        Ok(Self {
            limits,
            source,
            authored,
            candidate,
            replay_reference,
        })
    }

    /// Return selected operation limits.
    #[must_use]
    pub const fn limits(&self) -> PatchLimits {
        self.limits
    }

    /// Borrow source semantic facts.
    #[must_use]
    pub const fn source(&self) -> &StreamSourceProof {
        &self.source
    }

    /// Return the authored proof.
    #[must_use]
    pub const fn authored_proof(&self) -> AuthoredStreamProof {
        self.authored
    }

    /// Return candidate semantic facts.
    #[must_use]
    pub const fn candidate(&self) -> StreamCandidateProof {
        self.candidate
    }

    /// Borrow the optional durable replay token.
    #[must_use]
    pub fn replay_reference(&self) -> Option<&AuthoredReplayReference> {
        self.replay_reference.as_ref()
    }

    /// Compare a fresh source scan with the durable source facts. The
    /// process-local `SourceVersion` is deliberately excluded; complete raw
    /// archive identity and every persisted semantic source fact are checked.
    #[must_use]
    pub fn matches_source_proof(
        &self,
        main_part: &str,
        source: tail_append::SourceProof,
        archive: ArtifactProof,
    ) -> bool {
        self.source.main_part == main_part
            && self.source.source_len == source.source_len
            && self.source.source_sha256 == source.source_sha256
            && self.source.insertion_offset == source.insertion_offset
            && self.source.paragraph_count == source.paragraph_count
            && self.source.event_count == source.event_count
            && self.source.max_depth == source.max_depth
            && self.source.strict_namespace == source.strict_namespace
            && self.source.sect_pr_len == source.sect_pr_len
            && self.source.sect_pr_sha256 == source.sect_pr_sha256
            && self.source.archive == archive
    }

    /// Compare a fresh candidate scan with the persisted semantic candidate
    /// facts. Candidate XML parser events are compared exactly to the sealed
    /// candidate proof; only the authored input framing proof is compared to
    /// the authored handle separately.
    #[must_use]
    pub fn matches_candidate_proof(&self, candidate: tail_append::CandidateProof) -> bool {
        self.candidate.candidate_len == candidate.candidate_len
            && self.candidate.candidate_sha256 == candidate.candidate_sha256
            && self.candidate.paragraph_count == candidate.paragraph_count
            && self.candidate.event_count == candidate.event_count
            && self.candidate.max_depth == candidate.max_depth
            && self.candidate.generated_offset == candidate.generated_offset
            && self.candidate.generated_paragraph_count == self.authored.paragraph_count
            && self.candidate.generated_once == candidate.generated_once
            && self.candidate.sect_pr_len == candidate.sect_pr_len
            && self.candidate.sect_pr_sha256 == candidate.sect_pr_sha256
    }

    /// Compare a fresh candidate archive identity with the persisted archive
    /// proof. This is useful after an OPC publication has completed its
    /// authenticated candidate hashing pass.
    #[must_use]
    pub fn matches_candidate_archive(&self, archive: ArtifactProof) -> bool {
        self.candidate.archive == archive
    }

    /// Whether a caller-resolvable replay token is present.
    #[must_use]
    pub const fn is_durable(&self) -> bool {
        self.replay_reference.is_some()
    }

    /// Encode deterministic canonical JSON with no XML payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let reference = self
            .replay_reference
            .as_ref()
            .ok_or(PatchError::NotDurable)?;
        validate_reference(reference.as_bytes(), self.limits)?;
        let maximum =
            usize::try_from(self.limits.max_patch_bytes).map_err(|_| PatchError::Limit {
                resource: "patch bytes",
                actual: self.limits.max_patch_bytes,
                maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
            })?;
        let mut writer = JsonWriter::new(maximum);
        serde_json::to_writer(&mut writer, &self.to_value(reference))
            .map_err(|error| PatchError::Encoding(error.to_string()))?;
        Ok(writer.finish())
    }

    /// Parse and validate a canonical bounded JSON envelope.
    ///
    /// This convenience form uses the absolute parser ceiling. Callers that
    /// already have the selected finite policy should use
    /// [`Self::from_bytes_with_limit`] so that policy is checked before
    /// `serde_json` allocates or parses anything.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limit(bytes, ABSOLUTE_MAX_PATCH_BYTES as u64)
    }

    /// Parse a canonical JSON envelope under a caller-selected finite wire
    /// ceiling. The byte length is rejected before `serde_json` runs.
    /// Re-encoding the parsed value must exactly equal the input, which
    /// rejects duplicate keys, unknown spellings, noncanonical
    /// escapes/numbers, and trailing bytes.
    pub fn from_bytes_with_limit(bytes: &[u8], maximum: u64) -> Result<Self> {
        validate_wire_limit(maximum)?;
        let maximum = usize::try_from(maximum).map_err(|_| PatchError::Limit {
            resource: "patch bytes",
            actual: maximum,
            maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
        })?;
        if bytes.len() > maximum {
            return Err(PatchError::Limit {
                resource: "patch bytes",
                actual: bytes.len() as u64,
                maximum: maximum as u64,
            });
        }
        let value: Value = serde_json::from_slice(bytes)
            .map_err(|error| PatchError::InvalidWire(error.to_string()))?;
        if canonical_bytes(&value)?.as_slice() != bytes {
            return Err(PatchError::InvalidWire(
                "durable patch is not canonical JSON".to_owned(),
            ));
        }
        let patch = Self::from_value(&value)?;
        let selected =
            usize::try_from(patch.limits.max_patch_bytes).map_err(|_| PatchError::Limit {
                resource: "patch bytes",
                actual: patch.limits.max_patch_bytes,
                maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
            })?;
        if bytes.len() > selected {
            return Err(PatchError::Limit {
                resource: "patch bytes",
                actual: bytes.len() as u64,
                maximum: selected as u64,
            });
        }
        Ok(patch)
    }

    /// Resolve and authenticate the same sealed authored provider before the
    /// coordinator reruns semantic preparation and physical publication.
    pub fn resolve_replay<R: AuthoredReplayResolver>(
        &self,
        resolver: &R,
    ) -> Result<Arc<dyn AuthoredReplayHandle>> {
        let reference = self
            .replay_reference
            .as_ref()
            .ok_or(PatchError::MissingReplayProvider)?;
        let handle = resolver
            .resolve(reference)
            .map_err(PatchError::ReplayResolver)?;
        if handle.proof() != self.authored {
            return Err(PatchError::ReplayProofMismatch);
        }
        let resolved = handle
            .durable_reference()
            .ok_or(PatchError::MissingReplayProvider)?;
        if resolved.as_bytes() != reference.as_bytes() {
            return Err(PatchError::ReplayReferenceMismatch);
        }
        Ok(handle)
    }

    /// Create an explicit durable inverse authorization. The caller supplies
    /// the original provider token; it is never inferred from a path.
    pub fn inverse_authorization(
        &self,
        reference: OriginalArtifactReference,
    ) -> Result<ExactInverseAuthorization> {
        if reference.is_empty() {
            return Err(PatchError::MissingOriginalProvider);
        }
        Ok(ExactInverseAuthorization {
            current: self.candidate.archive,
            original: self.source.archive,
            reference,
            max_output_bytes: self.limits.max_output_bytes,
        })
    }

    fn to_value(&self, reference: &AuthoredReplayReference) -> Value {
        object([
            ("authored", authored_value(self.authored)),
            ("candidate", candidate_value(self.candidate)),
            ("limits", limits_value(self.limits)),
            ("operation", Value::String(OPERATION.to_owned())),
            (
                "replay",
                object([("reference", Value::String(hex(reference.as_bytes())))]),
            ),
            ("source", source_value(&self.source)),
            ("version", number(VERSION)),
        ])
    }

    fn from_value(value: &Value) -> Result<Self> {
        let root = object_ref(value, "root")?;
        require_keys(
            root,
            &[
                "authored",
                "candidate",
                "limits",
                "operation",
                "replay",
                "source",
                "version",
            ],
        )?;
        if string_field(root, "operation")? != OPERATION {
            return Err(PatchError::InvalidWire(
                "unsupported stream patch operation".to_owned(),
            ));
        }
        if u64_field(root, "version")? != VERSION {
            return Err(PatchError::InvalidWire(
                "unsupported stream patch version".to_owned(),
            ));
        }
        let limits = parse_limits(value_field(root, "limits")?)?;
        let source = parse_source(value_field(root, "source")?)?;
        let authored = parse_authored(value_field(root, "authored")?)?;
        let candidate = parse_candidate(value_field(root, "candidate")?)?;
        let replay = object_ref(value_field(root, "replay")?, "replay")?;
        require_keys(replay, &["reference"])?;
        let token = decode_hex(string_field(replay, "reference")?)?;
        validate_reference(&token, limits)?;
        let replay_reference = AuthoredReplayReference::try_from_bytes(
            &token,
            limits.max_replay_bytes.min(limits.max_patch_bytes),
        )
        .map_err(|error| PatchError::InvalidWire(error.to_string()))?;
        Self::new(limits, source, authored, candidate, Some(replay_reference))
    }
}

/// Explicit resolver for a caller-owned durable authored source.
pub trait AuthoredReplayResolver {
    /// Resolve one bounded opaque token into a sealed replay handle.
    fn resolve(
        &self,
        reference: &AuthoredReplayReference,
    ) -> std::result::Result<Arc<dyn AuthoredReplayHandle>, ReplayResolverError>;
}

/// Typed resolver refusal.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum ReplayResolverError {
    /// No provider exists for the supplied token.
    #[error("no authored replay provider is registered for the reference")]
    Missing,
    /// Caller-owned provider failed.
    #[error("authored replay provider failed: {0}")]
    Provider(#[source] Box<dyn std::error::Error + Send + Sync>),
    /// Provider returned an unusable token.
    #[error("authored replay provider returned an invalid reference")]
    Invalid,
}

/// Bounded caller-owned reference for reopening an exact original archive.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct OriginalArtifactReference {
    bytes: Arc<Vec<u8>>,
}

impl fmt::Debug for OriginalArtifactReference {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("OriginalArtifactReference")
            .field("length", &self.bytes.len())
            .finish()
    }
}

impl OriginalArtifactReference {
    /// Copy a nonempty bounded provider token.
    pub fn try_from_bytes(bytes: &[u8], maximum: u64) -> Result<Self> {
        if maximum == 0 || maximum == u64::MAX || maximum > MAX_ORIGINAL_REFERENCE_BYTES {
            return Err(PatchError::Limit {
                resource: "original reference bytes",
                actual: maximum,
                maximum: MAX_ORIGINAL_REFERENCE_BYTES,
            });
        }
        let length = u64::try_from(bytes.len()).map_err(|_| PatchError::Limit {
            resource: "original reference bytes",
            actual: u64::MAX,
            maximum,
        })?;
        if length == 0 || length > maximum {
            return Err(PatchError::Limit {
                resource: "original reference bytes",
                actual: length,
                maximum,
            });
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| PatchError::Allocation {
                resource: "original reference bytes",
                source,
            })?;
        owned.extend_from_slice(bytes);
        Ok(Self {
            bytes: Arc::new(owned),
        })
    }

    /// Borrow the opaque token.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Return token length.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    /// Return whether the token is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }
}

/// Explicit caller capability for reopening an original archive.
pub trait OriginalArtifactProvider {
    /// Reopen the artifact represented by the explicit token.
    fn open(
        &self,
        reference: &OriginalArtifactReference,
    ) -> std::result::Result<SourceArtifact, OriginalArtifactError>;
}

/// Typed original-provider failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum OriginalArtifactError {
    /// No provider data is available.
    #[error("original artifact provider has no data for the reference")]
    Missing,
    /// Provider data no longer names the sealed source.
    #[error("original artifact provider is stale")]
    Stale,
    /// Caller-owned transport/storage failure.
    #[error("original artifact provider failed: {0}")]
    Provider(#[source] Box<dyn std::error::Error + Send + Sync>),
}

/// Proof and explicit provider token authorizing one exact inverse.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExactInverseAuthorization {
    current: ArtifactProof,
    original: ArtifactProof,
    reference: OriginalArtifactReference,
    max_output_bytes: u64,
}

impl ExactInverseAuthorization {
    /// Construct an inverse authorization from complete artifact proofs.
    pub fn new(
        current: ArtifactProof,
        original: ArtifactProof,
        reference: OriginalArtifactReference,
        max_output_bytes: u64,
    ) -> Result<Self> {
        if max_output_bytes == 0 || max_output_bytes == u64::MAX {
            return Err(PatchError::Limit {
                resource: "output bytes",
                actual: max_output_bytes,
                maximum: max_output_bytes.saturating_sub(1),
            });
        }
        if reference.is_empty() {
            return Err(PatchError::MissingOriginalProvider);
        }
        current.validate("current artifact bytes", MAX_ARTIFACT_PROOF_BYTES)?;
        original.validate("original artifact bytes", max_output_bytes)?;
        Ok(Self {
            current,
            original,
            reference,
            max_output_bytes,
        })
    }

    /// Complete current-candidate artifact proof.
    #[must_use]
    pub const fn current(&self) -> ArtifactProof {
        self.current
    }

    /// Complete original-artifact proof.
    #[must_use]
    pub const fn original(&self) -> ArtifactProof {
        self.original
    }

    /// Borrow the explicit original-provider token.
    #[must_use]
    pub fn reference(&self) -> &OriginalArtifactReference {
        &self.reference
    }

    /// Return the physical output ceiling retained by this authorization.
    #[must_use]
    pub const fn max_output_bytes(&self) -> u64 {
        self.max_output_bytes
    }

    /// Encode this exact-inverse authorization as bounded canonical JSON.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limit(ABSOLUTE_MAX_PATCH_BYTES as u64)
    }

    /// Encode this exact-inverse authorization under a caller-selected wire
    /// ceiling. The token's hex expansion is checked before its JSON value is
    /// allocated; the streaming writer enforces the complete envelope size.
    pub fn to_bytes_with_limit(&self, maximum: u64) -> Result<Vec<u8>> {
        validate_wire_limit(maximum)?;
        self.current
            .validate("current artifact bytes", MAX_ARTIFACT_PROOF_BYTES)?;
        self.original
            .validate("original artifact bytes", self.max_output_bytes)?;
        let token_length = u64::try_from(self.reference.len()).map_err(|_| PatchError::Limit {
            resource: "original reference bytes",
            actual: u64::MAX,
            maximum,
        })?;
        let encoded_token_length = token_length.checked_mul(2).ok_or(PatchError::Limit {
            resource: "original reference bytes",
            actual: u64::MAX,
            maximum,
        })?;
        if encoded_token_length > maximum {
            return Err(PatchError::Limit {
                resource: "original reference bytes",
                actual: encoded_token_length,
                maximum,
            });
        }
        let maximum = usize::try_from(maximum).map_err(|_| PatchError::Limit {
            resource: "patch bytes",
            actual: maximum,
            maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
        })?;
        let mut writer = JsonWriter::new(maximum);
        serde_json::to_writer(
            &mut writer,
            &object([
                ("current", artifact_value(self.current)),
                ("max_output_bytes", number(self.max_output_bytes)),
                ("operation", Value::String(INVERSE_OPERATION.to_owned())),
                ("original", artifact_value(self.original)),
                ("reference", Value::String(hex(self.reference.as_bytes()))),
                ("version", number(VERSION)),
            ]),
        )
        .map_err(|error| PatchError::Encoding(error.to_string()))?;
        Ok(writer.finish())
    }

    /// Parse an exact-inverse authorization under the absolute wire ceiling.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limit(bytes, ABSOLUTE_MAX_PATCH_BYTES as u64)
    }

    /// Parse a canonical exact-inverse authorization after checking the
    /// caller-selected finite wire ceiling and before invoking `serde_json`.
    pub fn from_bytes_with_limit(bytes: &[u8], maximum: u64) -> Result<Self> {
        validate_wire_limit(maximum)?;
        let maximum = usize::try_from(maximum).map_err(|_| PatchError::Limit {
            resource: "patch bytes",
            actual: maximum,
            maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
        })?;
        if bytes.len() > maximum {
            return Err(PatchError::Limit {
                resource: "patch bytes",
                actual: bytes.len() as u64,
                maximum: maximum as u64,
            });
        }
        let value: Value = serde_json::from_slice(bytes)
            .map_err(|error| PatchError::InvalidWire(error.to_string()))?;
        if canonical_bytes(&value)?.as_slice() != bytes {
            return Err(PatchError::InvalidWire(
                "exact-inverse authorization is not canonical JSON".to_owned(),
            ));
        }
        let root = object_ref(&value, "exact-inverse authorization")?;
        require_keys(
            root,
            &[
                "current",
                "max_output_bytes",
                "operation",
                "original",
                "reference",
                "version",
            ],
        )?;
        if string_field(root, "operation")? != INVERSE_OPERATION {
            return Err(PatchError::InvalidWire(
                "unsupported exact-inverse operation".to_owned(),
            ));
        }
        if u64_field(root, "version")? != VERSION {
            return Err(PatchError::InvalidWire(
                "unsupported exact-inverse version".to_owned(),
            ));
        }
        let current = parse_artifact(value_field(root, "current")?)?;
        let original = parse_artifact(value_field(root, "original")?)?;
        let max_output_bytes = u64_field(root, "max_output_bytes")?;
        let token = decode_hex(string_field(root, "reference")?)?;
        let reference = OriginalArtifactReference::try_from_bytes(&token, maximum as u64)?;
        Self::new(current, original, reference, max_output_bytes)
    }
}

/// Apply a durable exact inverse through an explicit original provider.
pub fn apply_exact_inverse<W: Write, P: OriginalArtifactProvider>(
    current: &super::Package,
    inverse: &ExactInverseAuthorization,
    provider: &P,
    writer: W,
) -> Result<()> {
    let current_package: &SourceBackedPackage = &current.package;
    authenticate_current_artifact(current_package, inverse.current)?;
    let original = match provider.open(&inverse.reference) {
        Ok(original) => original,
        Err(error) => {
            // Retain source-freshness precedence even when the provider
            // fails before returning an original artifact.
            authenticate_current_artifact(current_package, inverse.current)?;
            return Err(PatchError::OriginalProvider(error));
        },
    };
    let proof = SourceArtifactRestoreProof {
        current_len: inverse.current.len,
        current_sha256: SourceArtifactFingerprint::from_sha256(inverse.current.sha256),
        original_len: inverse.original.len,
        original_sha256: SourceArtifactFingerprint::from_sha256(inverse.original.sha256),
    };
    current_package
        .restore_source_artifact_to_stream(&original, proof, inverse.max_output_bytes, writer)
        .map_err(PatchError::Opc)
}

fn authenticate_current_artifact(
    current: &SourceBackedPackage,
    expected: ArtifactProof,
) -> Result<()> {
    let artifact = current.source_artifact();
    if artifact.len() != expected.len {
        return Err(PatchError::Opc(
            litchi_opc::error::OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "length",
            },
        ));
    }
    let actual = artifact.fingerprint().map_err(PatchError::Opc)?;
    if actual != SourceArtifactFingerprint::from_sha256(expected.sha256) {
        return Err(PatchError::Opc(
            litchi_opc::error::OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "sha256",
            },
        ));
    }
    Ok(())
}

/// Authenticate a retained artifact and return only its compact identity.
pub fn artifact_proof(artifact: &SourceArtifact) -> Result<ArtifactProof> {
    let fingerprint = artifact.fingerprint().map_err(PatchError::Opc)?;
    Ok(ArtifactProof::new(
        artifact.len(),
        fingerprint.into_sha256(),
    ))
}

fn validate_facts(
    limits: &PatchLimits,
    source: &StreamSourceProof,
    authored: AuthoredStreamProof,
    candidate: &StreamCandidateProof,
) -> Result<()> {
    let main_part_len = u64::try_from(source.main_part.len()).map_err(|_| PatchError::Limit {
        resource: "main-part name bytes",
        actual: u64::MAX,
        maximum: MAX_MAIN_PART_BYTES,
    })?;
    if main_part_len == 0 || main_part_len > MAX_MAIN_PART_BYTES {
        return Err(PatchError::Limit {
            resource: "main-part name bytes",
            actual: main_part_len,
            maximum: MAX_MAIN_PART_BYTES,
        });
    }
    if source.main_part.as_bytes().contains(&0) {
        return Err(PatchError::InvalidFacts("main-part name contains NUL"));
    }
    if source.source_len == 0 || source.source_len > limits.max_source_xml_bytes {
        return Err(PatchError::Limit {
            resource: "source XML bytes",
            actual: source.source_len,
            maximum: limits.max_source_xml_bytes,
        });
    }
    if source.insertion_offset > source.source_len || source.sect_pr_len > source.source_len {
        return Err(PatchError::InvalidFacts("source span is out of range"));
    }
    if source.paragraph_count > limits.max_paragraphs
        || source.event_count > limits.max_events
        || source.max_depth > limits.max_depth
    {
        return Err(PatchError::InvalidFacts(
            "source facts exceed selected limits",
        ));
    }
    if authored.paragraph_count == 0
        || authored.event_count == 0
        || authored.encoded_xml_bytes == 0
        || authored.paragraph_count > limits.max_authored_paragraphs
        || authored.event_count > limits.max_authored_events
        || authored.text_bytes > limits.max_authored_text_bytes
        || authored.encoded_xml_bytes > limits.max_authored_xml_bytes
        || authored.strict_namespace != source.strict_namespace
    {
        return Err(PatchError::InvalidFacts(
            "authored proof exceeds selected limits",
        ));
    }
    let expected_candidate_len = source
        .source_len
        .checked_add(authored.encoded_xml_bytes)
        .ok_or(PatchError::InvalidFacts("candidate XML length overflow"))?;
    if candidate.candidate_len != expected_candidate_len
        || candidate.candidate_len > limits.max_candidate_xml_bytes
    {
        return Err(PatchError::InvalidFacts(
            "candidate length does not match source plus authored XML",
        ));
    }
    if candidate.generated_offset != source.insertion_offset
        || candidate.generated_len != authored.encoded_xml_bytes
        || candidate.generated_offset > candidate.candidate_len
        || candidate.generated_len > candidate.candidate_len
        || candidate.generated_paragraph_count != authored.paragraph_count
        || !candidate.generated_once
    {
        return Err(PatchError::InvalidFacts(
            "candidate generated range is not authenticated",
        ));
    }
    let expected_paragraphs = source
        .paragraph_count
        .checked_add(authored.paragraph_count)
        .ok_or(PatchError::InvalidFacts(
            "candidate paragraph count overflow",
        ))?;
    if candidate.paragraph_count != expected_paragraphs {
        return Err(PatchError::InvalidFacts(
            "candidate paragraph count does not match authored count",
        ));
    }
    // `authored.event_count` counts caller framing (start/text/end), while
    // `candidate.event_count` counts XML parser events after encoding. Their
    // exact values are intentionally independent; only the candidate's
    // bounded growth over the source parser pass is authenticated here.
    if candidate.paragraph_count > limits.max_paragraphs
        || candidate.event_count <= source.event_count
        || candidate.event_count > limits.max_events
        || candidate.max_depth > limits.max_depth
        || candidate.max_depth < source.max_depth
        || candidate.sect_pr_len > candidate.candidate_len
        || candidate.sect_pr_len != source.sect_pr_len
        || candidate.sect_pr_sha256 != source.sect_pr_sha256
    {
        return Err(PatchError::InvalidFacts(
            "candidate facts exceed selected limits",
        ));
    }
    source
        .archive
        .validate("source archive bytes", limits.max_output_bytes)?;
    candidate
        .archive
        .validate("candidate archive bytes", limits.max_output_bytes)?;
    Ok(())
}

fn validate_reference(bytes: &[u8], limits: PatchLimits) -> Result<()> {
    let length = u64::try_from(bytes.len()).map_err(|_| PatchError::Limit {
        resource: "replay reference bytes",
        actual: u64::MAX,
        maximum: limits.max_replay_bytes,
    })?;
    let maximum = limits.max_replay_bytes.min(limits.max_patch_bytes);
    if length == 0 || length > maximum {
        return Err(PatchError::Limit {
            resource: "replay reference bytes",
            actual: length,
            maximum,
        });
    }
    Ok(())
}

fn validate_wire_limit(maximum: u64) -> Result<()> {
    if maximum == 0 || maximum == u64::MAX {
        return Err(PatchError::Limit {
            resource: "patch bytes",
            actual: maximum,
            maximum: maximum.saturating_sub(1),
        });
    }
    if maximum > ABSOLUTE_MAX_PATCH_BYTES as u64 {
        return Err(PatchError::Limit {
            resource: "patch bytes",
            actual: maximum,
            maximum: ABSOLUTE_MAX_PATCH_BYTES as u64,
        });
    }
    Ok(())
}

fn object<I>(fields: I) -> Value
where
    I: IntoIterator<Item = (&'static str, Value)>,
{
    let mut map = Map::new();
    for (key, value) in fields {
        map.insert(key.to_owned(), value);
    }
    Value::Object(map)
}

fn number(value: u64) -> Value {
    Value::Number(Number::from(value))
}

fn bool_value(value: bool) -> Value {
    Value::Bool(value)
}

fn hash_value(value: &[u8; 32]) -> Value {
    Value::String(hex(value))
}

fn artifact_value(value: ArtifactProof) -> Value {
    object([
        ("len", number(value.len)),
        ("sha256", hash_value(&value.sha256)),
    ])
}

fn source_value(value: &StreamSourceProof) -> Value {
    object([
        ("archive", artifact_value(value.archive)),
        ("event_count", number(value.event_count)),
        ("insertion_offset", number(value.insertion_offset)),
        ("main_part", Value::String(value.main_part.clone())),
        ("max_depth", number(value.max_depth)),
        ("paragraph_count", number(value.paragraph_count)),
        ("sect_pr_len", number(value.sect_pr_len)),
        ("sect_pr_sha256", hash_value(&value.sect_pr_sha256)),
        ("source_len", number(value.source_len)),
        ("source_sha256", hash_value(&value.source_sha256)),
        ("strict_namespace", bool_value(value.strict_namespace)),
    ])
}

fn candidate_value(value: StreamCandidateProof) -> Value {
    object([
        ("archive", artifact_value(value.archive)),
        ("candidate_len", number(value.candidate_len)),
        ("candidate_sha256", hash_value(&value.candidate_sha256)),
        ("event_count", number(value.event_count)),
        ("generated_len", number(value.generated_len)),
        ("generated_offset", number(value.generated_offset)),
        (
            "generated_paragraph_count",
            number(value.generated_paragraph_count),
        ),
        ("generated_once", bool_value(value.generated_once)),
        ("max_depth", number(value.max_depth)),
        ("paragraph_count", number(value.paragraph_count)),
        ("sect_pr_len", number(value.sect_pr_len)),
        ("sect_pr_sha256", hash_value(&value.sect_pr_sha256)),
    ])
}

fn authored_value(value: AuthoredStreamProof) -> Value {
    object([
        ("encoded_sha256", hash_value(&value.encoded_sha256)),
        ("encoded_xml_bytes", number(value.encoded_xml_bytes)),
        ("event_count", number(value.event_count)),
        ("event_sha256", hash_value(&value.event_sha256)),
        ("paragraph_count", number(value.paragraph_count)),
        ("strict_namespace", bool_value(value.strict_namespace)),
        ("text_bytes", number(value.text_bytes)),
    ])
}

fn limits_value(value: PatchLimits) -> Value {
    object([
        (
            "max_authored_chunk_bytes",
            number(value.max_authored_chunk_bytes),
        ),
        ("max_authored_events", number(value.max_authored_events)),
        (
            "max_authored_paragraphs",
            number(value.max_authored_paragraphs),
        ),
        (
            "max_authored_text_bytes",
            number(value.max_authored_text_bytes),
        ),
        (
            "max_authored_xml_bytes",
            number(value.max_authored_xml_bytes),
        ),
        (
            "max_candidate_xml_bytes",
            number(value.max_candidate_xml_bytes),
        ),
        ("max_depth", number(value.max_depth)),
        ("max_events", number(value.max_events)),
        ("max_fragment_bytes", number(value.max_fragment_bytes)),
        ("max_output_bytes", number(value.max_output_bytes)),
        ("max_paragraphs", number(value.max_paragraphs)),
        ("max_patch_bytes", number(value.max_patch_bytes)),
        ("max_replay_bytes", number(value.max_replay_bytes)),
        (
            "max_replay_window_bytes",
            number(value.max_replay_window_bytes),
        ),
        (
            "max_settings_xml_bytes",
            number(value.max_settings_xml_bytes),
        ),
        ("max_source_xml_bytes", number(value.max_source_xml_bytes)),
        ("max_text_bytes", number(value.max_text_bytes)),
        ("max_token_bytes", number(value.max_token_bytes)),
        ("max_workspace_bytes", number(value.max_workspace_bytes)),
    ])
}

fn object_ref<'a>(value: &'a Value, name: &'static str) -> Result<&'a Map<String, Value>> {
    value
        .as_object()
        .ok_or_else(|| PatchError::InvalidWire(format!("{name} must be a JSON object")))
}

fn value_field<'a>(object: &'a Map<String, Value>, field: &'static str) -> Result<&'a Value> {
    object
        .get(field)
        .ok_or_else(|| PatchError::InvalidWire(format!("missing {field} field")))
}

fn string_field<'a>(object: &'a Map<String, Value>, field: &'static str) -> Result<&'a str> {
    value_field(object, field)?
        .as_str()
        .ok_or_else(|| PatchError::InvalidWire(format!("{field} must be a JSON string")))
}

fn u64_field(object: &Map<String, Value>, field: &'static str) -> Result<u64> {
    value_field(object, field)?
        .as_u64()
        .ok_or_else(|| PatchError::InvalidWire(format!("{field} must be a nonnegative integer")))
}

fn bool_field(object: &Map<String, Value>, field: &'static str) -> Result<bool> {
    value_field(object, field)?
        .as_bool()
        .ok_or_else(|| PatchError::InvalidWire(format!("{field} must be a JSON boolean")))
}

fn hash_field(object: &Map<String, Value>, field: &'static str) -> Result<[u8; 32]> {
    let text = string_field(object, field)?;
    let bytes = decode_hex(text)?;
    if bytes.len() != 32 {
        return Err(PatchError::InvalidWire(format!(
            "{field} must contain 32 bytes"
        )));
    }
    let mut hash = [0_u8; 32];
    hash.copy_from_slice(&bytes);
    Ok(hash)
}

fn require_keys(object: &Map<String, Value>, expected: &[&str]) -> Result<()> {
    if object.len() != expected.len() || expected.iter().any(|key| !object.contains_key(*key)) {
        return Err(PatchError::InvalidWire(
            "unknown or missing patch field".to_owned(),
        ));
    }
    Ok(())
}

fn parse_artifact(value: &Value) -> Result<ArtifactProof> {
    let object = object_ref(value, "artifact")?;
    require_keys(object, &["len", "sha256"])?;
    Ok(ArtifactProof::new(
        u64_field(object, "len")?,
        hash_field(object, "sha256")?,
    ))
}

fn parse_source(value: &Value) -> Result<StreamSourceProof> {
    let object = object_ref(value, "source")?;
    require_keys(
        object,
        &[
            "archive",
            "event_count",
            "insertion_offset",
            "main_part",
            "max_depth",
            "paragraph_count",
            "sect_pr_len",
            "sect_pr_sha256",
            "source_len",
            "source_sha256",
            "strict_namespace",
        ],
    )?;
    let main_part = string_field(object, "main_part")?;
    let main_part_len = u64::try_from(main_part.len()).map_err(|_| PatchError::Limit {
        resource: "main-part name bytes",
        actual: u64::MAX,
        maximum: MAX_MAIN_PART_BYTES,
    })?;
    if main_part_len == 0 || main_part_len > MAX_MAIN_PART_BYTES {
        return Err(PatchError::Limit {
            resource: "main-part name bytes",
            actual: main_part_len,
            maximum: MAX_MAIN_PART_BYTES,
        });
    }
    if main_part.as_bytes().contains(&0) {
        return Err(PatchError::InvalidWire(
            "main-part name contains NUL".to_owned(),
        ));
    }
    Ok(StreamSourceProof {
        archive: parse_artifact(value_field(object, "archive")?)?,
        event_count: u64_field(object, "event_count")?,
        insertion_offset: u64_field(object, "insertion_offset")?,
        main_part: main_part.to_owned(),
        max_depth: u64_field(object, "max_depth")?,
        paragraph_count: u64_field(object, "paragraph_count")?,
        sect_pr_len: u64_field(object, "sect_pr_len")?,
        sect_pr_sha256: hash_field(object, "sect_pr_sha256")?,
        source_len: u64_field(object, "source_len")?,
        source_sha256: hash_field(object, "source_sha256")?,
        strict_namespace: bool_field(object, "strict_namespace")?,
    })
}

fn parse_candidate(value: &Value) -> Result<StreamCandidateProof> {
    let object = object_ref(value, "candidate")?;
    require_keys(
        object,
        &[
            "archive",
            "candidate_len",
            "candidate_sha256",
            "event_count",
            "generated_len",
            "generated_offset",
            "generated_paragraph_count",
            "generated_once",
            "max_depth",
            "paragraph_count",
            "sect_pr_len",
            "sect_pr_sha256",
        ],
    )?;
    Ok(StreamCandidateProof {
        archive: parse_artifact(value_field(object, "archive")?)?,
        candidate_len: u64_field(object, "candidate_len")?,
        candidate_sha256: hash_field(object, "candidate_sha256")?,
        event_count: u64_field(object, "event_count")?,
        generated_len: u64_field(object, "generated_len")?,
        generated_offset: u64_field(object, "generated_offset")?,
        generated_paragraph_count: u64_field(object, "generated_paragraph_count")?,
        generated_once: bool_field(object, "generated_once")?,
        max_depth: u64_field(object, "max_depth")?,
        paragraph_count: u64_field(object, "paragraph_count")?,
        sect_pr_len: u64_field(object, "sect_pr_len")?,
        sect_pr_sha256: hash_field(object, "sect_pr_sha256")?,
    })
}

fn parse_authored(value: &Value) -> Result<AuthoredStreamProof> {
    let object = object_ref(value, "authored")?;
    require_keys(
        object,
        &[
            "encoded_sha256",
            "encoded_xml_bytes",
            "event_count",
            "event_sha256",
            "paragraph_count",
            "strict_namespace",
            "text_bytes",
        ],
    )?;
    Ok(AuthoredStreamProof {
        encoded_sha256: hash_field(object, "encoded_sha256")?,
        encoded_xml_bytes: u64_field(object, "encoded_xml_bytes")?,
        event_count: u64_field(object, "event_count")?,
        event_sha256: hash_field(object, "event_sha256")?,
        paragraph_count: u64_field(object, "paragraph_count")?,
        strict_namespace: bool_field(object, "strict_namespace")?,
        text_bytes: u64_field(object, "text_bytes")?,
    })
}

fn parse_limits(value: &Value) -> Result<PatchLimits> {
    let object = object_ref(value, "limits")?;
    require_keys(
        object,
        &[
            "max_authored_chunk_bytes",
            "max_authored_events",
            "max_authored_paragraphs",
            "max_authored_text_bytes",
            "max_authored_xml_bytes",
            "max_candidate_xml_bytes",
            "max_depth",
            "max_events",
            "max_fragment_bytes",
            "max_output_bytes",
            "max_paragraphs",
            "max_patch_bytes",
            "max_replay_bytes",
            "max_replay_window_bytes",
            "max_settings_xml_bytes",
            "max_source_xml_bytes",
            "max_text_bytes",
            "max_token_bytes",
            "max_workspace_bytes",
        ],
    )?;
    let limits = PatchLimits {
        max_authored_chunk_bytes: u64_field(object, "max_authored_chunk_bytes")?,
        max_authored_events: u64_field(object, "max_authored_events")?,
        max_authored_paragraphs: u64_field(object, "max_authored_paragraphs")?,
        max_authored_text_bytes: u64_field(object, "max_authored_text_bytes")?,
        max_authored_xml_bytes: u64_field(object, "max_authored_xml_bytes")?,
        max_candidate_xml_bytes: u64_field(object, "max_candidate_xml_bytes")?,
        max_depth: u64_field(object, "max_depth")?,
        max_events: u64_field(object, "max_events")?,
        max_fragment_bytes: u64_field(object, "max_fragment_bytes")?,
        max_output_bytes: u64_field(object, "max_output_bytes")?,
        max_paragraphs: u64_field(object, "max_paragraphs")?,
        max_patch_bytes: u64_field(object, "max_patch_bytes")?,
        max_replay_bytes: u64_field(object, "max_replay_bytes")?,
        max_replay_window_bytes: u64_field(object, "max_replay_window_bytes")?,
        max_settings_xml_bytes: u64_field(object, "max_settings_xml_bytes")?,
        max_source_xml_bytes: u64_field(object, "max_source_xml_bytes")?,
        max_text_bytes: u64_field(object, "max_text_bytes")?,
        max_token_bytes: u64_field(object, "max_token_bytes")?,
        max_workspace_bytes: u64_field(object, "max_workspace_bytes")?,
    };
    limits.validate()?;
    Ok(limits)
}

fn canonical_bytes(value: &Value) -> Result<Vec<u8>> {
    let mut writer = JsonWriter::new(ABSOLUTE_MAX_PATCH_BYTES);
    serde_json::to_writer(&mut writer, value)
        .map_err(|error| PatchError::Encoding(error.to_string()))?;
    Ok(writer.finish())
}

struct JsonWriter {
    bytes: Vec<u8>,
    maximum: usize,
}

impl JsonWriter {
    fn new(maximum: usize) -> Self {
        Self {
            bytes: Vec::new(),
            maximum,
        }
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for JsonWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let next = self.bytes.len().checked_add(bytes.len()).ok_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidInput, "durable patch length overflow")
        })?;
        if next > self.maximum {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "durable patch exceeds selected byte limit",
            ));
        }
        self.bytes
            .try_reserve_exact(bytes.len())
            .map_err(io::Error::other)?;
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut result = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        result.push(char::from(DIGITS[usize::from(byte >> 4)]));
        result.push(char::from(DIGITS[usize::from(byte & 0x0f)]));
    }
    result
}

fn decode_hex(value: &str) -> Result<Vec<u8>> {
    if value.len() % 2 != 0 {
        return Err(PatchError::InvalidWire(
            "hex field has odd length".to_owned(),
        ));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(value.len() / 2)
        .map_err(|source| PatchError::Allocation {
            resource: "durable patch hex field",
            source,
        })?;
    for pair in value.as_bytes().as_chunks::<2>().0 {
        let high = hex_digit(pair[0]).ok_or_else(|| {
            PatchError::InvalidWire("hex field contains a noncanonical digit".to_owned())
        })?;
        let low = hex_digit(pair[1]).ok_or_else(|| {
            PatchError::InvalidWire("hex field contains a noncanonical digit".to_owned())
        })?;
        bytes.push((high << 4) | low);
    }
    Ok(bytes)
}

fn hex_digit(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        _ => None,
    }
}

/// Durable patch construction and inverse errors.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum PatchError {
    /// The patch has no caller-resolvable replay token.
    #[error("stream patch is process-local and has no durable replay reference")]
    NotDurable,
    /// A provider was required but was not supplied or did not expose a
    /// durable reference.
    #[error("durable authored replay provider is missing")]
    MissingReplayProvider,
    /// A resolved provider token differed from the sealed token.
    #[error("durable authored replay reference does not match the patch")]
    ReplayReferenceMismatch,
    /// A resolved provider proof differed from the sealed proof.
    #[error("durable authored replay proof does not match the patch")]
    ReplayProofMismatch,
    /// Resolver failed before publication.
    #[error(transparent)]
    ReplayResolver(#[from] ReplayResolverError),
    /// No original provider token was supplied.
    #[error("durable exact inverse requires an explicit original-artifact provider")]
    MissingOriginalProvider,
    /// Original provider failed before output.
    #[error(transparent)]
    OriginalProvider(#[from] OriginalArtifactError),
    /// Compact proof facts are internally inconsistent.
    #[error("invalid stream patch facts: {0}")]
    InvalidFacts(&'static str),
    /// Canonical JSON was malformed or noncanonical.
    #[error("invalid durable stream patch wire: {0}")]
    InvalidWire(String),
    /// Canonical JSON encoding failed.
    #[error("durable stream patch encoding failed: {0}")]
    Encoding(String),
    /// A selected finite bound was exceeded.
    #[error("stream patch {resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// Resource name.
        resource: &'static str,
        /// Observed value.
        actual: u64,
        /// Selected maximum.
        maximum: u64,
    },
    /// A bounded allocation failed before output.
    #[error("stream patch allocation failed for {resource}: {source}")]
    Allocation {
        /// Allocation owner.
        resource: &'static str,
        /// Allocator failure.
        #[source]
        source: std::collections::TryReserveError,
    },
    /// OPC rejected or interrupted an exact restore.
    #[error(transparent)]
    Opc(#[from] litchi_opc::error::OpcError),
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits() -> PatchLimits {
        PatchLimits {
            max_source_xml_bytes: 1_000,
            max_text_bytes: 1_000,
            max_fragment_bytes: 1_000,
            max_candidate_xml_bytes: 2_000,
            max_events: 1_000,
            max_depth: 32,
            max_paragraphs: 100,
            max_settings_xml_bytes: 1_000,
            max_workspace_bytes: 4_000,
            max_output_bytes: 10_000,
            max_token_bytes: 256,
            max_authored_paragraphs: 20,
            max_authored_events: 100,
            max_authored_chunk_bytes: 256,
            max_authored_text_bytes: 1_000,
            max_authored_xml_bytes: 1_000,
            max_replay_bytes: 1_000,
            max_replay_window_bytes: 256,
            max_patch_bytes: 65_536,
        }
    }

    fn patch(reference: Option<AuthoredReplayReference>) -> ParagraphStreamPatch {
        let source = StreamSourceProof {
            main_part: "/word/document.xml".to_owned(),
            source_len: 100,
            source_sha256: [1; 32],
            insertion_offset: 80,
            paragraph_count: 3,
            event_count: 20,
            max_depth: 4,
            strict_namespace: false,
            sect_pr_len: 10,
            sect_pr_sha256: [2; 32],
            archive: ArtifactProof::new(500, [3; 32]),
        };
        let authored = AuthoredStreamProof {
            strict_namespace: false,
            paragraph_count: 2,
            event_count: 6,
            text_bytes: 5,
            encoded_xml_bytes: 20,
            event_sha256: [4; 32],
            encoded_sha256: [5; 32],
        };
        let candidate = StreamCandidateProof {
            candidate_len: 120,
            candidate_sha256: [6; 32],
            paragraph_count: 5,
            // XML parser events need not equal source events plus the
            // authored input framing count: text chunk boundaries are a
            // format-layer proof detail.
            event_count: 27,
            max_depth: 4,
            generated_offset: 80,
            generated_len: 20,
            generated_paragraph_count: 2,
            generated_once: true,
            sect_pr_len: 10,
            sect_pr_sha256: [2; 32],
            archive: ArtifactProof::new(600, [7; 32]),
        };
        ParagraphStreamPatch::new(limits(), source, authored, candidate, reference).unwrap()
    }

    #[test]
    fn canonical_wire_round_trips() {
        let reference = AuthoredReplayReference::try_from_bytes(b"provider-token", 1_000).unwrap();
        let original = patch(Some(reference));
        let bytes = original.to_bytes().unwrap();
        assert!(bytes.len() < ABSOLUTE_MAX_PATCH_BYTES);
        let decoded = ParagraphStreamPatch::from_bytes(&bytes).unwrap();
        assert_eq!(decoded, original);
    }

    #[test]
    fn candidate_parser_events_are_independent_of_authored_framing() {
        let reference = AuthoredReplayReference::try_from_bytes(b"provider-token", 1_000).unwrap();
        let original = patch(Some(reference));
        assert_eq!(original.authored_proof().event_count, 6);
        assert_eq!(original.candidate().event_count, 27);
        assert!(original.candidate().event_count > original.source().event_count);
    }

    #[test]
    fn process_local_patch_cannot_be_serialized() {
        assert!(matches!(
            patch(None).to_bytes(),
            Err(PatchError::NotDurable)
        ));
    }

    #[test]
    fn exact_inverse_wire_round_trips() {
        let replay = AuthoredReplayReference::try_from_bytes(b"provider-token", 1_000).unwrap();
        let original = patch(Some(replay));
        let reference =
            OriginalArtifactReference::try_from_bytes(b"original-provider", 1_000).unwrap();
        let inverse = original.inverse_authorization(reference).unwrap();
        let bytes = inverse.to_bytes_with_limit(65_536).unwrap();
        let decoded = ExactInverseAuthorization::from_bytes_with_limit(&bytes, 65_536).unwrap();
        assert_eq!(decoded, inverse);
    }

    #[test]
    fn inverse_allows_larger_current_candidate() {
        let reference = OriginalArtifactReference::try_from_bytes(b"original", 1_000).unwrap();
        let inverse = ExactInverseAuthorization::new(
            ArtifactProof::new(200, [1; 32]),
            ArtifactProof::new(100, [2; 32]),
            reference,
            100,
        )
        .unwrap();
        assert_eq!(inverse.current().len, 200);
        assert_eq!(inverse.original().len, inverse.max_output_bytes());
    }

    #[test]
    fn duplicate_and_unknown_fields_are_rejected() {
        let reference = AuthoredReplayReference::try_from_bytes(b"provider-token", 1_000).unwrap();
        let bytes = patch(Some(reference)).to_bytes().unwrap();
        let mut duplicate = bytes.clone();
        duplicate.pop();
        duplicate.extend_from_slice(br#","version":1}"#);
        assert!(matches!(
            ParagraphStreamPatch::from_bytes(&duplicate),
            Err(PatchError::InvalidWire(_))
        ));

        let mut unknown = bytes;
        unknown.pop();
        unknown.extend_from_slice(br#","unknown":0}"#);
        assert!(matches!(
            ParagraphStreamPatch::from_bytes(&unknown),
            Err(PatchError::InvalidWire(_))
        ));
    }
}
