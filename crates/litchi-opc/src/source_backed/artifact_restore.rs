//! Authenticated exact restoration using an explicitly supplied original artifact.

use super::PackURI;
use super::splice::{
    check_inverse_state, fingerprint_inverse_snapshot, finish_inverse_publication,
};
use super::{
    ContextCheckedSink, Counted, OpcError, OutputBudgetedSink, Result,
    SOURCE_PUBLICATION_CHUNK_BYTES, SourceArtifact, SourceArtifactFingerprint, SourceBackedPackage,
    SourceCheckedSink, SourceSnapshot, finish_source_publication, map_execution_error,
    map_io_error, overlay_unavailable, read_source_at_with_context,
};
use crate::SpliceResource;
use crate::limits::{SourceLimitKind, SourceLimits};
use crate::pkgreader::PackageReader;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use sha2::{Digest as _, Sha256};
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::io::{self, Write};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

/// The phase in which an artifact producer is being invoked.
///
/// A probe authenticates a producer into a bounded digest sink and therefore
/// must not consume the recipient's output allowance. A publication writes to
/// the caller's sink and charges that allowance exactly once.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ArtifactStreamPhase {
    Probe,
    Publish,
}

/// Recipient-owned policy passed through a source-backed artifact stream.
///
/// The policy contains the recipient's complete source/output profile. A
/// producer keeps using its own source, work and freshness context while it
/// writes, while the recipient profile owns every output admission and byte
/// charge. This split is what lets a deferred format producer publish into a
/// different package without inheriting the producer's output allowance or
/// charging the recipient twice. Format crates receive this value from a
/// callback; they do not construct or retain it.
pub struct ArtifactStreamPolicy<'recipient> {
    phase: ArtifactStreamPhase,
    recipient_limits: SourceLimits,
    recipient_context: Option<&'recipient ExecutionContext>,
    recipient_output_reservation_failures: Option<Arc<super::DiagnosticCounter>>,
    failure: Arc<Mutex<Option<ExecutionError>>>,
}

struct SourceCatalogView<'source> {
    source: &'source SourceSnapshot,
    archive: &'source soapberry_zip::office::IndexedArchive<super::SourceReader>,
    catalog: &'source super::AdmittedCatalog,
    /// An optional final decoded size for one ordinary Part being replaced.
    /// Source-backed splice admission uses this to avoid charging the
    /// recipient for an oversized original that the final plan shrinks.
    replacement: Option<(usize, u64)>,
}

/// One bounded folded-name index for structural ZIP members used during a
/// recipient census. Relationship owners are numerous on real OPC packages;
/// retaining this index once avoids rescanning every physical entry for each
/// `.rels` member while still refusing ASCII-folded ambiguities.
struct AdmissionPhysicalMemberLookup {
    by_folded_name: HashMap<String, (soapberry_zip::office::EntryId, usize)>,
    _memory: Option<Reservation>,
}

struct AdmissionPartLookup {
    by_folded_name: HashMap<String, usize>,
    _memory: Option<Reservation>,
}

struct AdmissionRelationshipLookup<'a> {
    by_folded_name: HashMap<String, &'a super::Relationships>,
    _memory: Option<Reservation>,
}

fn build_admission_part_lookup(
    source: &SourceSnapshot,
    catalog: &super::AdmittedCatalog,
) -> Result<AdmissionPartLookup> {
    let mut name_bytes = 0_u64;
    for (index, part) in catalog.parts.iter().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        name_bytes = name_bytes
            .checked_add(part.partname.as_str().len() as u64)
            .ok_or_else(|| overlay_unavailable("catalog Part name bytes overflow"))?;
    }
    let part_count = catalog.parts.len() as u64;
    let memory_bound = name_bytes
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(part_count.checked_mul(64)?))
        .ok_or_else(|| overlay_unavailable("catalog Part lookup memory overflows"))?;
    let memory = source
        .context
        .as_ref()
        .map(|context| {
            context
                .reserve(Resource::Memory, memory_bound)
                .map_err(map_execution_error)
        })
        .transpose()?;
    let mut by_folded_name: HashMap<String, usize> = HashMap::new();
    by_folded_name
        .try_reserve(catalog.parts.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC folded catalog Part lookup",
            source,
        })?;
    for (index, part) in catalog.parts.iter().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        let folded = super::folded_part_name(&part.partname)?;
        if by_folded_name.insert(folded, index).is_some() {
            return Err(overlay_unavailable(
                "catalog Part name is ambiguous under ASCII case folding",
            ));
        }
    }
    Ok(AdmissionPartLookup {
        by_folded_name,
        _memory: memory,
    })
}

fn build_admission_physical_lookup(
    source: &SourceSnapshot,
    archive: &soapberry_zip::office::IndexedArchive<super::SourceReader>,
) -> Result<AdmissionPhysicalMemberLookup> {
    let mut name_bytes = 0_u64;
    for (index, (_, name)) in archive.file_entries().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        name_bytes = name_bytes
            .checked_add(name.len() as u64)
            .ok_or_else(|| overlay_unavailable("structural member name bytes overflow"))?;
    }
    let member_count = archive.len() as u64;
    let memory_bound = name_bytes
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(member_count.checked_mul(128)?))
        .ok_or_else(|| overlay_unavailable("structural member lookup memory overflows"))?;
    let memory = source
        .context
        .as_ref()
        .map(|context| {
            context
                .reserve(Resource::Memory, memory_bound)
                .map_err(map_execution_error)
        })
        .transpose()?;
    let mut by_folded_name: HashMap<String, (soapberry_zip::office::EntryId, usize)> =
        HashMap::new();
    by_folded_name
        .try_reserve(archive.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC structural member lookup",
            source,
        })?;
    for (index, (entry_id, name)) in archive.file_entries().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        let folded =
            super::folded_ascii_name(name, "source-backed OPC folded structural member name")?;
        if let Some((_, count)) = by_folded_name.get_mut(&folded) {
            *count = count
                .checked_add(1)
                .ok_or_else(|| overlay_unavailable("structural member count overflows"))?;
        } else {
            by_folded_name.insert(folded, (entry_id, 1));
        }
    }
    Ok(AdmissionPhysicalMemberLookup {
        by_folded_name,
        _memory: memory,
    })
}

fn structural_member_entry_id(
    archive: &soapberry_zip::office::IndexedArchive<super::SourceReader>,
    physical_members: &AdmissionPhysicalMemberLookup,
    member_name: &str,
) -> Result<Option<soapberry_zip::office::EntryId>> {
    let folded = super::folded_ascii_name(
        member_name,
        "source-backed OPC folded structural member name",
    )?;
    match physical_members.by_folded_name.get(&folded) {
        Some((entry_id, 1)) => Ok(archive.entry_id(member_name).or(Some(*entry_id))),
        Some(_) => Err(overlay_unavailable(
            "structural OPC member name is ambiguous under ASCII case folding",
        )),
        None => Ok(archive.entry_id(member_name)),
    }
}

struct StructuralMember {
    bytes: Vec<u8>,
    _memory: Option<Reservation>,
}

/// Read one structural XML member for a recipient re-admission pass.
///
/// The source catalog retains parsed relationship/content-type facts, but it
/// intentionally does not retain XML bytes. A stricter recipient must still
/// re-run the bounded XML grammar (not just compare the producer's scalar
/// metrics), so this helper reads at most one already-size-checked structural
/// member at a time. Ordinary Part payloads never use this path.
fn read_structural_member(
    source: &SourceSnapshot,
    archive: &soapberry_zip::office::IndexedArchive<super::SourceReader>,
    physical_members: &AdmissionPhysicalMemberLookup,
    member_name: &str,
    kind: SourceLimitKind,
    maximum: u64,
    limits: SourceLimits,
) -> Result<Option<StructuralMember>> {
    let entry_id = structural_member_entry_id(archive, physical_members, member_name)?;
    let Some(entry_id) = entry_id else {
        return Ok(None);
    };
    let metadata = archive
        .metadata_for(entry_id)
        .map_err(super::map_preservation_error)?;
    let declared = metadata.uncompressed_size();
    limits.check(kind, declared, maximum)?;
    limits.check(
        SourceLimitKind::EntryBytes,
        declared,
        limits.max_entry_bytes(),
    )?;
    let memory = source
        .context
        .as_ref()
        .map(|context| {
            let bound = declared
                .checked_add(SOURCE_PUBLICATION_CHUNK_BYTES as u64)
                .ok_or_else(|| overlay_unavailable("structural XML memory bound overflows"))?;
            context
                .reserve(Resource::Memory, bound)
                .map_err(map_execution_error)
        })
        .transpose()?;
    source.ensure_current()?;
    source
        .context
        .as_ref()
        .map(|context| context.check().map_err(map_execution_error))
        .transpose()?;
    let bytes = archive
        .read_entry(entry_id)
        .map_err(super::map_preservation_error)?;
    source.ensure_current()?;
    source
        .context
        .as_ref()
        .map(|context| context.check().map_err(map_execution_error))
        .transpose()?;
    if let Some(context) = source.context.as_ref() {
        context
            .consume(Resource::Work, bytes.len() as u64)
            .map_err(map_execution_error)?;
        source.ensure_current()?;
    }
    if u64::try_from(bytes.len()).ok() != Some(declared) {
        return Err(OpcError::ZipError(format!(
            "structural OPC member '{member_name}' declared {declared} uncompressed bytes but read {}",
            bytes.len()
        )));
    }
    Ok(Some(StructuralMember {
        bytes,
        _memory: memory,
    }))
}

fn relationship_member_name(owner: &PackURI) -> Result<String> {
    if owner.as_str() == super::PACKAGE_URI {
        Ok("_rels/.rels".to_owned())
    } else {
        owner
            .rels_uri()
            .map_err(OpcError::InvalidPackUri)
            .map(|uri| uri.membername().to_owned())
    }
}

fn relationship_owner_from_member(member_name: &str) -> Result<Option<PackURI>> {
    if !super::is_relationship_member_name(member_name) {
        return Ok(None);
    }
    if member_name.eq_ignore_ascii_case("_rels/.rels") {
        return PackURI::new(super::PACKAGE_URI)
            .map(Some)
            .map_err(OpcError::InvalidPackUri);
    }
    let marker = b"/_rels/";
    let marker_start = member_name
        .as_bytes()
        .windows(marker.len())
        .rposition(|window| {
            window
                .iter()
                .zip(marker)
                .all(|(actual, expected)| actual.eq_ignore_ascii_case(expected))
        });
    let Some(marker_start) = marker_start else {
        return Err(overlay_unavailable(
            "relationship member owner cannot be derived",
        ));
    };
    let directory = &member_name[..marker_start];
    let filename = &member_name[marker_start + marker.len()..];
    let Some((filename, extension)) = filename.rsplit_once('.') else {
        return Err(overlay_unavailable(
            "relationship member owner has an invalid suffix",
        ));
    };
    if !extension.eq_ignore_ascii_case("rels") {
        return Err(overlay_unavailable(
            "relationship member owner has an invalid suffix",
        ));
    }
    let owner = format!("/{directory}/{filename}");
    PackURI::new(owner)
        .map(Some)
        .map_err(OpcError::InvalidPackUri)
}

fn catalog_part_index(
    catalog: &super::AdmittedCatalog,
    folded_parts: &AdmissionPartLookup,
    owner: &PackURI,
) -> Result<Option<usize>> {
    // The catalog's admission already rejects ambiguous OPC-equivalent Part
    // names. Reuse its exact/case-folded lookup index instead of scanning all
    // Parts for every relationship owner.
    if let Some(index) = catalog.parts_by_name.get(owner).copied() {
        return Ok(Some(index));
    }
    if let Some(order) = catalog.casefold_order.as_ref() {
        let wanted = owner.as_str();
        return Ok(order
            .binary_search_by(|index| {
                super::cmp_ascii_case_insensitive(catalog.parts[*index].partname.as_str(), wanted)
            })
            .ok()
            .map(|position| order[position]));
    }
    let folded = super::folded_part_name(owner)?;
    Ok(folded_parts.by_folded_name.get(&folded).copied())
}

fn canonical_relationship_owner(
    catalog: &super::AdmittedCatalog,
    folded_parts: &AdmissionPartLookup,
    owner: &PackURI,
) -> Result<PackURI> {
    let index = catalog_part_index(catalog, folded_parts, owner)?;
    Ok(index
        .and_then(|index| catalog.parts.get(index))
        .map_or_else(|| owner.clone(), |part| part.partname.clone()))
}

fn admit_content_types_xml(xml: &[u8], limits: SourceLimits) -> Result<()> {
    let _ = super::ContentTypeMap::from_xml(xml, limits)?;
    Ok(())
}

fn parse_relationship_xml(
    xml: &[u8],
    owner: &PackURI,
    limits: SourceLimits,
) -> Result<super::Relationships> {
    PackageReader::parse_owned_relationships(xml, owner, limits)
}

fn folded_relationship_lookup<'a>(
    source: &SourceSnapshot,
    relationships: &'a HashMap<PackURI, super::Relationships>,
) -> Result<AdmissionRelationshipLookup<'a>> {
    let mut name_bytes = 0_u64;
    for (index, (owner, _)) in relationships.iter().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        name_bytes = name_bytes
            .checked_add(owner.as_str().len() as u64)
            .ok_or_else(|| overlay_unavailable("relationship owner name bytes overflow"))?;
    }
    let relationship_count = relationships.len() as u64;
    let memory_bound = name_bytes
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(relationship_count.checked_mul(96)?))
        .ok_or_else(|| overlay_unavailable("relationship owner lookup memory overflows"))?;
    let memory = source
        .context
        .as_ref()
        .map(|context| {
            context
                .reserve(Resource::Memory, memory_bound)
                .map_err(map_execution_error)
        })
        .transpose()?;
    let mut lookup = HashMap::new();
    lookup
        .try_reserve(relationships.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC folded relationship owner lookup",
            source,
        })?;
    for (index, (owner, value)) in relationships.iter().enumerate() {
        if index & 0xff == 0 {
            source.ensure_current()?;
            source
                .context
                .as_ref()
                .map(|context| context.check().map_err(map_execution_error))
                .transpose()?;
        }
        let folded = super::folded_ascii_name(
            owner.as_str(),
            "source-backed OPC folded relationship owner",
        )?;
        if lookup.insert(folded, value).is_some() {
            return Err(overlay_unavailable(
                "relationship owner is ambiguous under ASCII case folding",
            ));
        }
    }
    Ok(AdmissionRelationshipLookup {
        by_folded_name: lookup,
        _memory: memory,
    })
}

fn admit_preservation_member_names(
    source: &SourceSnapshot,
    archive: &soapberry_zip::office::IndexedArchive<super::SourceReader>,
    limits: SourceLimits,
) -> Result<()> {
    source.ensure_current()?;
    for entry in archive.preservation_entries() {
        limits.check(
            SourceLimitKind::ArchiveMemberNameBytes,
            entry.raw_name_len(),
            limits.max_archive_member_name_bytes(),
        )?;
    }
    source.ensure_current()
}

/// The typed failure bridge for a source-backed producer.
///
/// `Opc` retains failures raised by the OPC publication seam. `Producer`
/// retains the format crate's original error value. When a sink or freshness
/// fence fails during publication, `Publication` keeps the
/// exact prefix and both observed causes when they are available. The enum's
/// display/debug text is intentionally fixed; callers inspect the typed
/// accessors for details.
pub enum ArtifactStreamError<E> {
    /// Failure raised by the OPC source or publication layer.
    Opc(OpcError),
    /// Failure returned by the producer format crate.
    Producer(E),
    /// A failure observed while publishing, with an exact accepted prefix.
    Publication {
        /// Exact bytes accepted by the caller's sink. This may be zero when a
        /// second freshness failure accompanies a producer error before output.
        written: u64,
        /// The primary failure.
        source: Box<Self>,
        /// A second failure observed while closing the operation, if any.
        secondary: Option<Box<Self>>,
    },
}

impl<E> ArtifactStreamError<E> {
    fn opc(error: OpcError) -> Self {
        Self::Opc(error)
    }

    fn producer(error: E) -> Self {
        Self::Producer(error)
    }

    fn incomplete(written: u64, source: Self, secondary: Option<Self>) -> Self {
        Self::Publication {
            written,
            source: Box::new(source),
            secondary: secondary.map(Box::new),
        }
    }

    /// Return the exact accepted prefix, when output was partially emitted.
    #[must_use]
    pub fn written(&self) -> Option<u64> {
        match self {
            Self::Publication { written, .. } => Some(*written),
            _ => None,
        }
    }

    /// Return the first OPC failure in this error tree.
    #[must_use]
    pub fn opc_error(&self) -> Option<&OpcError> {
        match self {
            Self::Opc(error) => Some(error),
            Self::Producer(_) => None,
            Self::Publication {
                source, secondary, ..
            } => source
                .opc_error()
                .or_else(|| secondary.as_deref().and_then(Self::opc_error)),
        }
    }

    /// Return the first producer failure in this error tree.
    #[must_use]
    pub fn producer_error(&self) -> Option<&E> {
        match self {
            Self::Opc(_) => None,
            Self::Producer(error) => Some(error),
            Self::Publication {
                source, secondary, ..
            } => source
                .producer_error()
                .or_else(|| secondary.as_deref().and_then(Self::producer_error)),
        }
    }

    /// Return a secondary failure observed after the primary failure.
    #[must_use]
    pub fn secondary(&self) -> Option<&Self> {
        match self {
            Self::Publication { secondary, .. } => secondary.as_deref(),
            _ => None,
        }
    }
}

fn with_secondary<E>(
    primary: ArtifactStreamError<E>,
    secondary: Option<ArtifactStreamError<E>>,
) -> ArtifactStreamError<E> {
    match secondary {
        Some(secondary) => ArtifactStreamError::incomplete(0, primary, Some(secondary)),
        None => primary,
    }
}

impl<E> fmt::Debug for ArtifactStreamError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("artifact stream operation failed")
    }
}

impl<E> fmt::Display for ArtifactStreamError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("artifact stream operation failed")
    }
}

impl<E> std::error::Error for ArtifactStreamError<E> {}

/// A format-owned execution/cancellation fence for one native publication.
///
/// The OPC target invokes this guard before publication and before every write
/// and flush performed by the native writer.  The guard error remains typed in
/// [`ArtifactStreamError::Producer`]; it is never flattened into an
/// `io::Error` visible to the format caller.
pub trait ArtifactStreamGuard {
    type Error;

    fn check(&self) -> std::result::Result<(), Self::Error>;
}

/// The exact report from one mediated artifact stream.
///
/// The producer cannot construct this value or write to the target directly.
/// It is minted only after one of the native OPC publication methods has
/// completed its recipient census and emitted one complete artifact. The
/// fingerprint is over the bytes accepted by the target sink.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ArtifactStreamReport {
    written: u64,
    fingerprint: SourceArtifactFingerprint,
    target_token: u64,
}

impl ArtifactStreamReport {
    /// Return the exact number of bytes accepted by the publication target.
    #[must_use]
    pub const fn written(self) -> u64 {
        self.written
    }

    /// Return the SHA-256 fingerprint of the accepted artifact bytes.
    #[must_use]
    pub const fn fingerprint(self) -> SourceArtifactFingerprint {
        self.fingerprint
    }
}

static NEXT_ARTIFACT_STREAM_TARGET: AtomicU64 = AtomicU64::new(1);

/// A sealed recipient target for one complete native OPC artifact.
///
/// `ArtifactStreamTarget` intentionally does not implement [`Write`]. A
/// format producer can publish only through the native source, topology,
/// overlay, restore, and splice methods below; arbitrary bytes cannot bypass
/// the OPC census or the ZIP preservation profile. The target is created by
/// the OPC facade and is valid for exactly one publication attempt.
pub struct ArtifactStreamTarget<'a> {
    writer: &'a mut dyn Write,
    policy: &'a ArtifactStreamPolicy<'a>,
    token: u64,
    claimed: bool,
    invalidated: bool,
    written: u64,
    hasher: Sha256,
    publication: Option<ArtifactStreamReport>,
}

impl<'a> ArtifactStreamTarget<'a> {
    fn new(writer: &'a mut dyn Write, policy: &'a ArtifactStreamPolicy<'a>) -> Self {
        Self {
            writer,
            policy,
            token: NEXT_ARTIFACT_STREAM_TARGET.fetch_add(1, Ordering::Relaxed),
            claimed: false,
            invalidated: false,
            written: 0,
            hasher: Sha256::new(),
            publication: None,
        }
    }

    fn publish<F>(&mut self, operation: F) -> Result<ArtifactStreamReport>
    where
        F: FnOnce(&mut dyn Write, &ArtifactStreamPolicy<'a>) -> Result<()>,
    {
        if self.claimed {
            self.invalidated = true;
            return Err(overlay_unavailable(
                "artifact stream target already has a publication attempt",
            ));
        }
        self.claimed = true;
        let policy = self.policy;
        let mut sink = ArtifactTargetSink { target: self };
        operation(&mut sink, policy)?;
        let publication = ArtifactStreamReport {
            written: sink.target.written,
            fingerprint: SourceArtifactFingerprint::from_sha256(
                sink.target.hasher.clone().finalize().into(),
            ),
            target_token: sink.target.token,
        };
        sink.target.publication = Some(publication);
        Ok(publication)
    }

    fn publish_with_guard<G, F>(
        &mut self,
        guard: &G,
        operation: F,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>>
    where
        G: ArtifactStreamGuard,
        F: FnOnce(&mut dyn Write, &ArtifactStreamPolicy<'a>) -> Result<()>,
    {
        if self.claimed {
            self.invalidated = true;
            return Err(ArtifactStreamError::opc(overlay_unavailable(
                "artifact stream target already has a publication attempt",
            )));
        }
        self.claimed = true;
        if let Err(error) = guard.check() {
            return Err(ArtifactStreamError::producer(error));
        }
        let policy = self.policy;
        let (operation_result, guard_error) = {
            let mut sink = ArtifactTargetSink { target: self };
            let mut guarded = GuardedArtifactSink {
                inner: &mut sink,
                guard,
                failure: None,
            };
            let operation_result = operation(&mut guarded, policy);
            let guard_error = guarded.failure.take();
            (operation_result, guard_error)
        };
        let policy_error = policy
            .take_failure()
            .map(|error| ArtifactStreamError::opc(map_execution_error(error)));
        let operation_result = operation_result.map_err(ArtifactStreamError::opc);

        match (operation_result, guard_error, policy_error) {
            (Ok(()), None, None) => Ok(self.mint_publication()),
            (Err(operation), None, None) => Err(operation),
            (Ok(()), Some(guard), None) => Err(ArtifactStreamError::producer(guard)),
            (Err(operation), Some(guard), None) => Err(ArtifactStreamError::incomplete(
                self.written,
                ArtifactStreamError::producer(guard),
                Some(operation),
            )),
            (Ok(()), None, Some(policy)) => Err(policy),
            (Err(operation), None, Some(policy)) => Err(ArtifactStreamError::incomplete(
                self.written,
                operation,
                Some(policy),
            )),
            (Ok(()), Some(guard), Some(policy)) => Err(ArtifactStreamError::incomplete(
                self.written,
                ArtifactStreamError::producer(guard),
                Some(policy),
            )),
            (Err(operation), Some(guard), Some(policy)) => {
                let secondary =
                    ArtifactStreamError::incomplete(self.written, operation, Some(policy));
                Err(ArtifactStreamError::incomplete(
                    self.written,
                    ArtifactStreamError::producer(guard),
                    Some(secondary),
                ))
            },
        }
    }

    fn mint_publication(&mut self) -> ArtifactStreamReport {
        let publication = ArtifactStreamReport {
            written: self.written,
            fingerprint: self.fingerprint(),
            target_token: self.token,
        };
        self.publication = Some(publication);
        publication
    }

    /// Publish an exact retained source artifact.
    pub fn publish_source_artifact(
        &mut self,
        artifact: &SourceArtifact,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| artifact.write_to_with_policy(writer, policy))
    }

    /// Publish an exact retained source artifact with a typed guard checked
    /// around every native write and flush.
    pub fn publish_source_artifact_with_guard<G: ArtifactStreamGuard>(
        &mut self,
        artifact: &SourceArtifact,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            artifact.write_to_with_policy(writer, policy)
        })
    }

    /// Publish an exact source-backed package through the recipient policy.
    pub fn publish_source_package(
        &mut self,
        package: SourceBackedPackage,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| package.write_exact_source_with_policy(writer, policy))
    }

    /// Publish an exact source-backed package with a typed native-write guard.
    pub fn publish_source_package_with_guard<G: ArtifactStreamGuard>(
        &mut self,
        package: SourceBackedPackage,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            package.write_exact_source_with_policy(writer, policy)
        })
    }

    /// Publish a prepared effective topology.
    pub fn publish_topology(
        &mut self,
        topology: &super::PreparedTopology,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| topology.publish_to_stream_with_policy(writer, policy))
    }

    /// Publish a prepared topology with a typed native-write guard.
    pub fn publish_topology_with_guard<G: ArtifactStreamGuard>(
        &mut self,
        topology: &super::PreparedTopology,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            topology.publish_to_stream_with_policy(writer, policy)
        })
    }

    /// Prepare and publish one bounded source-backed topology plan.
    ///
    /// Relationship removals must be represented in `plan`, not as a
    /// separately generated `.rels` payload. The topology publisher applies
    /// those removals to the source XML ranges and therefore retains the
    /// source declaration, whitespace, comments, and unrelated lexical bytes.
    pub fn publish_overlay(
        &mut self,
        package: SourceBackedPackage,
        plan: super::SourceTopologyPlan,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| {
            package.write_topology_to_stream_with_policy(writer, plan, policy)
        })
    }

    /// Publish a source topology plan with a typed native-write guard.
    ///
    /// Relationship removals must be represented by this topology plan.  The
    /// OPC publisher then retains the source relationship XML ranges and
    /// applies the removal in place, preserving declarations, whitespace,
    /// comments, and every unrelated lexical byte.
    pub fn publish_overlay_with_guard<G: ArtifactStreamGuard>(
        &mut self,
        package: SourceBackedPackage,
        plan: super::SourceTopologyPlan,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            package.write_topology_to_stream_with_policy(writer, plan, policy)
        })
    }

    /// Publish a bounded decoded-member splice plan.
    pub fn publish_splice(
        &mut self,
        plan: &super::SourcePartSplicePlan<'_>,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| {
            plan.duplicate()
                .write_to_with_policy(writer, policy)
                .map(|_| ())
        })
    }

    /// Authenticate the complete candidate artifact and publish the splice.
    ///
    /// The preview is bounded and performed by the existing splice planner;
    /// the caller's sink is touched only after both the length and SHA-256
    /// fingerprint match.
    pub fn publish_splice_with_expected_artifact(
        &mut self,
        plan: &super::SourcePartSplicePlan<'_>,
        expected_len: u64,
        expected_sha256: SourceArtifactFingerprint,
    ) -> Result<ArtifactStreamReport> {
        self.publish(|writer, policy| {
            plan.duplicate()
                .write_to_with_expected_artifact_with_policy(
                    writer,
                    expected_len,
                    expected_sha256,
                    policy,
                )
                .map(|_| ())
        })
    }

    /// Publish the splice with a format-owned execution/cancellation guard.
    pub fn publish_splice_with_guard<G: ArtifactStreamGuard>(
        &mut self,
        plan: &super::SourcePartSplicePlan<'_>,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            plan.duplicate()
                .write_to_with_policy(writer, policy)
                .map(|_| ())
        })
    }

    /// Authenticate and publish a splice while retaining typed guard errors.
    pub fn publish_splice_with_expected_artifact_and_guard<G: ArtifactStreamGuard>(
        &mut self,
        plan: &super::SourcePartSplicePlan<'_>,
        expected_len: u64,
        expected_sha256: SourceArtifactFingerprint,
        guard: &G,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<G::Error>> {
        self.publish_with_guard(guard, |writer, policy| {
            plan.duplicate()
                .write_to_with_expected_artifact_with_policy(
                    writer,
                    expected_len,
                    expected_sha256,
                    policy,
                )
                .map(|_| ())
        })
    }

    fn written(&self) -> u64 {
        self.written
    }

    fn fingerprint(&self) -> SourceArtifactFingerprint {
        SourceArtifactFingerprint::from_sha256(self.hasher.clone().finalize().into())
    }

    fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }

    fn validate_publication(&self, publication: ArtifactStreamReport) -> Result<()> {
        if self.invalidated
            || self.publication != Some(publication)
            || publication.target_token != self.token
            || publication.written != self.written
            || publication.fingerprint != self.fingerprint()
        {
            return Err(overlay_unavailable(
                "artifact stream publication proof does not belong to this target",
            ));
        }
        Ok(())
    }

    /// Mediate an exact inverse through this one-shot target.
    ///
    /// The current and original streams are probed and freshness-checked by
    /// the authenticated restore implementation.  The private target sink is
    /// the only writer passed to that implementation, so the format producer
    /// cannot bypass the recipient policy or publish a second artifact.
    pub fn publish_restore<C, O>(
        &mut self,
        package: &SourceBackedPackage,
        current: &C,
        original: &O,
        proof: SourceArtifactRestoreProof,
        max_output_bytes: u64,
    ) -> std::result::Result<ArtifactStreamReport, ArtifactStreamError<O::Error>>
    where
        C: ArtifactStream + ?Sized,
        O: ArtifactStream + ?Sized,
        C: ArtifactStream<Error = O::Error>,
    {
        if self.claimed {
            self.invalidated = true;
            return Err(ArtifactStreamError::opc(overlay_unavailable(
                "artifact stream target already has a publication attempt",
            )));
        }
        self.claimed = true;
        let policy = self.policy;
        let result = {
            let mut sink = ArtifactTargetSink { target: self };
            package.restore_streamed_artifact_to_stream(
                current,
                original,
                proof,
                max_output_bytes,
                &mut sink,
                Some(policy),
            )
        };
        match result {
            Ok(()) => Ok(self.mint_publication()),
            Err(error) => Err(error),
        }
    }
}

struct ArtifactTargetSink<'target, 'a> {
    target: &'target mut ArtifactStreamTarget<'a>,
}

impl Write for ArtifactTargetSink<'_, '_> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = self.target.writer.write(bytes)?;
        if accepted > bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "artifact stream target sink overreported accepted bytes",
            ));
        }
        if accepted != 0 {
            self.target.hasher.update(&bytes[..accepted]);
            self.target.written = self
                .target
                .written
                .checked_add(accepted as u64)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidData,
                        "artifact stream target byte count overflow",
                    )
                })?;
        }
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.target.writer.flush()
    }
}

struct GuardedArtifactSink<'guard, 'sink, G: ArtifactStreamGuard> {
    inner: &'sink mut dyn Write,
    guard: &'guard G,
    failure: Option<G::Error>,
}

impl<G: ArtifactStreamGuard> Write for GuardedArtifactSink<'_, '_, G> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.failure.is_none() {
            if let Err(error) = self.guard.check() {
                self.failure = Some(error);
            }
        }
        if self.failure.is_some() {
            // `Write::write_all` retries `Interrupted` indefinitely. A
            // format-owned guard refusal is terminal for this one-shot
            // publication, while `failure` retains the typed cause.
            return Err(io::Error::other(
                "artifact stream publication guard refused output",
            ));
        }
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.failure.is_none() {
            if let Err(error) = self.guard.check() {
                self.failure = Some(error);
            }
        }
        if self.failure.is_some() {
            return Err(io::Error::other(
                "artifact stream publication guard refused output",
            ));
        }
        self.inner.flush()
    }
}

/// A bounded source-backed producer that can be authenticated before output
/// is accepted by a caller-owned sink.
pub trait ArtifactStream {
    type Error;

    /// Check the producer's source/version and execution context.
    fn check_fresh(&self) -> std::result::Result<(), Self::Error>;

    /// Stream this artifact through the sealed native OPC target.
    ///
    /// Implementations must call exactly one target publication method. The
    /// returned proof is tied to that target invocation and cannot be created
    /// by a producer or reused with another target.
    fn stream_to(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
    ) -> std::result::Result<ArtifactStreamReport, Self::Error>;
}

impl<'recipient> ArtifactStreamPolicy<'recipient> {
    fn for_recipient(package: &'recipient SourceBackedPackage) -> Result<Self> {
        let recipient_context = package.cache.context();
        let recipient_output_reservation_failures =
            package.source.output_reservation_failures.clone();
        if recipient_context.is_some() && recipient_output_reservation_failures.is_none() {
            return Err(overlay_unavailable(
                "managed artifact stream output counter is unavailable",
            ));
        }
        Ok(Self {
            phase: ArtifactStreamPhase::Publish,
            recipient_limits: package.limits,
            recipient_context,
            recipient_output_reservation_failures,
            failure: Arc::new(Mutex::new(None)),
        })
    }

    pub(crate) fn probe(&self) -> Self {
        Self {
            phase: ArtifactStreamPhase::Probe,
            recipient_limits: self.recipient_limits,
            recipient_context: self.recipient_context,
            recipient_output_reservation_failures: self
                .recipient_output_reservation_failures
                .clone(),
            failure: Arc::new(Mutex::new(None)),
        }
    }

    fn is_publish(&self) -> bool {
        self.phase == ArtifactStreamPhase::Publish
    }

    fn recipient_context(&self) -> Option<&ExecutionContext> {
        self.recipient_context
    }

    fn failure(&self) -> Arc<Mutex<Option<ExecutionError>>> {
        Arc::clone(&self.failure)
    }

    pub(crate) fn take_failure(&self) -> Option<ExecutionError> {
        self.failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
    }

    pub(crate) fn check_recipient(&self) -> Result<()> {
        self.recipient_context.map_or(Ok(()), |context| {
            context.check().map_err(map_execution_error)
        })
    }

    /// Return the complete resource profile of the recipient package.
    ///
    /// The profile is used by a deferred/effective producer to admit its
    /// final OPC topology before publication.  It is deliberately separate
    /// from the producer's source profile: a producer keeps its own input and
    /// work checks while the recipient owns all output admission.
    #[must_use]
    pub const fn recipient_limits(&self) -> SourceLimits {
        self.recipient_limits
    }

    /// Check one recipient-owned resource without charging output bytes.
    ///
    /// Callers use this while census-ing a final topology.  Publication
    /// writers still charge accepted bytes through the policy sink after this
    /// check succeeds.
    pub fn check_recipient_limit(
        &self,
        resource: SourceLimitKind,
        actual: u64,
        maximum: u64,
    ) -> Result<()> {
        self.recipient_limits.check(resource, actual, maximum)
    }

    /// Admit the unchanged source catalog against the recipient profile.
    ///
    /// This is the bounded metadata census used by policy-aware OPC writers
    /// before they touch an external sink. It checks the retained ZIP member,
    /// Part, relationship, content-type, XML-event, attribute, target, and
    /// aggregate dimensions that are available without materializing ordinary
    /// payloads. A changed/effective producer must add its final-plan census
    /// through the same accessors before calling a policy writer.
    pub(crate) fn admit_source_catalog(&self, package: &SourceBackedPackage) -> Result<()> {
        self.admit_source_view(SourceCatalogView {
            source: &package.source,
            archive: &package.archive,
            catalog: &package.catalog,
            replacement: None,
        })
    }

    pub(crate) fn admit_source_artifact(&self, artifact: &SourceArtifact) -> Result<()> {
        self.admit_source_view(SourceCatalogView {
            source: &artifact.snapshot,
            archive: &artifact.archive,
            catalog: &artifact.catalog,
            replacement: None,
        })
    }

    /// Admit a decoded splice's final topology while allowing its target Part
    /// to shrink or grow relative to the source catalog. The ZIP profile
    /// callback still owns the exact final compressed/archive measurements;
    /// this pass covers the recipient's OPC Part and structural XML policy
    /// before replay can touch the publication sink.
    pub(crate) fn admit_splice_candidate(
        &self,
        package: &SourceBackedPackage,
        target: usize,
        candidate_bytes: u64,
    ) -> Result<()> {
        if target >= package.catalog.parts.len() {
            return Err(OpcError::PartNotFound(target.to_string()));
        }
        self.admit_source_view(SourceCatalogView {
            source: &package.source,
            archive: &package.archive,
            catalog: &package.catalog,
            replacement: Some((target, candidate_bytes)),
        })
    }

    fn admit_source_view(&self, view: SourceCatalogView<'_>) -> Result<()> {
        self.check_recipient()?;
        view.source.ensure_current()?;
        view.source
            .context
            .as_ref()
            .map(|context| context.check().map_err(map_execution_error))
            .transpose()?;
        let limits = self.recipient_limits;
        let physical_members = build_admission_physical_lookup(view.source, view.archive)?;
        let folded_parts = build_admission_part_lookup(view.source, view.catalog)?;
        if view.replacement.is_none() {
            admit_preservation_member_names(view.source, view.archive, limits)?;
        }
        if view.replacement.is_none() {
            limits.check(
                SourceLimitKind::InputBytes,
                view.source.length,
                limits.max_input_bytes(),
            )?;
        }
        let member_count = u64::try_from(view.archive.len())
            .map_err(|_| overlay_unavailable("recipient member count exceeds u64"))?;
        limits.check(
            SourceLimitKind::Members,
            member_count,
            limits.max_members() as u64,
        )?;
        let total_entries = u64::try_from(view.archive.preservation_entry_count())
            .map_err(|_| overlay_unavailable("recipient ZIP entry count exceeds u64"))?;
        limits.check(
            SourceLimitKind::ArchiveTotalEntries,
            total_entries,
            limits.max_archive_total_entries() as u64,
        )?;
        if view.replacement.is_none() {
            limits.check(
                SourceLimitKind::ArchiveMetadataBytes,
                view.archive.preservation_metadata_bytes(),
                limits.max_archive_metadata_bytes(),
            )?;
        }

        let mut aggregate_bytes = 0_u64;
        let mut relationship_parts = 0_u64;
        let replacement_entry = view
            .replacement
            .and_then(|(index, _)| view.catalog.parts.get(index))
            .map(|part| part.entry_id);
        for (index, (entry_id, name)) in view.archive.file_entries().enumerate() {
            if index & 0xff == 0 {
                view.source.ensure_current()?;
                view.source
                    .context
                    .as_ref()
                    .map(|context| context.check().map_err(map_execution_error))
                    .transpose()?;
            }
            let name_bytes = u64::try_from(name.len())
                .map_err(|_| overlay_unavailable("recipient member name exceeds u64"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                name_bytes,
                limits.max_archive_member_name_bytes(),
            )?;
            let metadata = view.archive.metadata_for(entry_id)?;
            let compressed = metadata.compressed_size();
            let replacement = replacement_entry == Some(entry_id);
            let entry = if replacement {
                view.replacement
                    .map(|(_, bytes)| bytes)
                    .ok_or_else(|| overlay_unavailable("replacement entry bytes are unavailable"))?
            } else {
                metadata.uncompressed_size()
            };
            if !replacement {
                limits.check(
                    SourceLimitKind::ArchiveCompressedBytes,
                    compressed,
                    limits.max_archive_compressed_bytes(),
                )?;
            }
            limits.check(SourceLimitKind::EntryBytes, entry, limits.max_entry_bytes())?;
            aggregate_bytes = aggregate_bytes
                .checked_add(entry)
                .ok_or_else(|| overlay_unavailable("recipient aggregate bytes overflow"))?;
            if super::is_relationship_member_name(name) {
                relationship_parts = relationship_parts
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("recipient relationship parts overflow"))?;
            }
        }
        limits.check(
            SourceLimitKind::AggregateBytes,
            aggregate_bytes,
            limits.max_aggregate_bytes(),
        )?;
        limits.check(
            SourceLimitKind::RelationshipParts,
            relationship_parts,
            limits.max_relationship_parts() as u64,
        )?;

        let part_count = u64::try_from(view.catalog.parts.len())
            .map_err(|_| overlay_unavailable("recipient Part count exceeds u64"))?;
        limits.check(
            SourceLimitKind::Parts,
            part_count,
            limits.max_parts() as u64,
        )?;
        let mut total_part_bytes = 0_u64;
        for (index, part) in view.catalog.parts.iter().enumerate() {
            if index & 0xff == 0 {
                view.source.ensure_current()?;
                view.source
                    .context
                    .as_ref()
                    .map(|context| context.check().map_err(map_execution_error))
                    .transpose()?;
            }
            let metadata = view.archive.metadata_for(part.entry_id)?;
            let bytes = if view.replacement.is_some_and(|(target, _)| target == index) {
                view.replacement
                    .map(|(_, bytes)| bytes)
                    .ok_or_else(|| overlay_unavailable("replacement Part bytes are unavailable"))?
            } else {
                metadata.uncompressed_size()
            };
            limits.check(SourceLimitKind::PartBytes, bytes, limits.max_part_bytes())?;
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            total_part_bytes = total_part_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient Part bytes overflow"))?;
        }
        limits.check(
            SourceLimitKind::TotalPartBytes,
            total_part_bytes,
            limits.max_total_part_bytes(),
        )?;

        let content_types_entry = view
            .archive
            .entry_id(&view.catalog.content_types_member)
            .ok_or_else(|| OpcError::PartNotFound(view.catalog.content_types_member.clone()))?;
        let content_types_bytes = view
            .archive
            .metadata_for(content_types_entry)?
            .uncompressed_size();
        limits.check(
            SourceLimitKind::ContentTypesBytes,
            content_types_bytes,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::PartBytes,
            content_types_bytes,
            limits.max_part_bytes(),
        )?;
        limits.check(
            SourceLimitKind::EntryBytes,
            content_types_bytes,
            limits.max_entry_bytes(),
        )?;
        limits.check(
            SourceLimitKind::ContentTypeMappings,
            view.catalog.content_types.mapping_count() as u64,
            limits.max_content_type_mappings() as u64,
        )?;
        if let Some(xml) = read_structural_member(
            view.source,
            view.archive,
            &physical_members,
            &view.catalog.content_types_member,
            SourceLimitKind::ContentTypesBytes,
            limits.max_content_types_bytes() as u64,
            limits,
        )? {
            admit_content_types_xml(&xml.bytes, limits)?;
        }

        let mut relationship_xml_bytes = 0_u64;
        let mut relationship_events = 0_u64;
        let mut total_relationships = 0_u64;
        let mut relationship_owners = HashSet::new();
        let mut admitted_relationships = HashMap::new();
        for owner in view.catalog.relationship_xml_metrics.keys() {
            relationship_owners.insert(owner.clone());
        }
        for (_, member_name) in view.archive.file_entries() {
            if let Some(owner) = relationship_owner_from_member(member_name)? {
                relationship_owners.insert(canonical_relationship_owner(
                    view.catalog,
                    &folded_parts,
                    &owner,
                )?);
            }
        }

        let mut check_relationships = |relationships: &super::Relationships| -> Result<()> {
            let count = u64::try_from(relationships.len())
                .map_err(|_| overlay_unavailable("recipient relationship count exceeds u64"))?;
            limits.check(
                SourceLimitKind::RelationshipsPerPart,
                count,
                limits.max_relationships_per_part() as u64,
            )?;
            total_relationships = total_relationships
                .checked_add(count)
                .ok_or_else(|| overlay_unavailable("recipient relationship count overflow"))?;
            limits.check(
                SourceLimitKind::TotalRelationships,
                total_relationships,
                limits.max_total_relationships() as u64,
            )?;
            for relationship in relationships.iter() {
                for value in [
                    relationship.r_id(),
                    relationship.reltype(),
                    relationship.target_ref(),
                ] {
                    limits.check(
                        SourceLimitKind::XmlAttributeBytes,
                        value.len() as u64,
                        limits.max_xml_attribute_bytes() as u64,
                    )?;
                    limits.check(
                        SourceLimitKind::XmlAttributeBytes,
                        super::escaped_xml_attribute_len(value)? as u64,
                        limits.max_xml_attribute_bytes() as u64,
                    )?;
                }
                limits.check(
                    SourceLimitKind::RelationshipTargetBytes,
                    relationship.target_ref().len() as u64,
                    limits.max_relationship_target_bytes() as u64,
                )?;
            }
            Ok(())
        };

        for owner in relationship_owners {
            let member_name = relationship_member_name(&owner)?;
            let Some(xml) = read_structural_member(
                view.source,
                view.archive,
                &physical_members,
                &member_name,
                SourceLimitKind::RelationshipXmlBytes,
                limits.max_relationship_xml_bytes() as u64,
                limits,
            )?
            else {
                return Err(overlay_unavailable(
                    "source relationship XML is unavailable for recipient admission",
                ));
            };
            let bytes = u64::try_from(xml.bytes.len())
                .map_err(|_| overlay_unavailable("relationship XML length exceeds u64"))?;
            let events = super::relationship_xml_event_count(&xml.bytes, limits)?;
            limits.check(
                SourceLimitKind::RelationshipXmlBytes,
                bytes,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            relationship_xml_bytes = relationship_xml_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient relationship bytes overflow"))?;
            relationship_events = relationship_events
                .checked_add(events)
                .ok_or_else(|| overlay_unavailable("recipient relationship events overflow"))?;
            let parsed = parse_relationship_xml(&xml.bytes, &owner, limits)?;
            if !view.catalog.relationship_xml_metrics.contains_key(&owner) {
                check_relationships(&parsed)?;
            }
            admitted_relationships.insert(owner, parsed);
        }
        limits.check(
            SourceLimitKind::TotalRelationshipXmlBytes,
            relationship_xml_bytes,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalRelationshipXmlEvents,
            relationship_events,
            limits.max_total_relationship_xml_events() as u64,
        )?;

        check_relationships(&view.catalog.package_relationships)?;
        for part in &view.catalog.parts {
            check_relationships(&part.relationships)?;
        }
        limits.check(
            SourceLimitKind::TotalRelationships,
            total_relationships,
            limits.max_total_relationships() as u64,
        )?;
        let admitted_relationships_by_owner =
            folded_relationship_lookup(view.source, &admitted_relationships)?;

        // Count the reachable relationship graph from the bounded, recipient-
        // admitted XML map. Keeping orphan owners in this map is important:
        // an orphan relationship part can still lead to another relationship
        // part, and the graph limit must cover that chain rather than falling
        // back to the source catalog only.
        let mut graph = HashSet::new();
        let mut work = Vec::new();
        enqueue_graph_targets(
            &view.catalog.package_relationships,
            &mut graph,
            &mut work,
            limits,
        )?;
        while let Some(owner) = work.pop() {
            if let Some(relationships) = admitted_relationships_by_owner.by_folded_name.get(&owner)
            {
                enqueue_graph_targets(relationships, &mut graph, &mut work, limits)?;
            }
        }
        Ok(())
    }

    /// Admit the exact final ZIP layout prepared by the preservation writer.
    ///
    /// The source/catalog checks above cannot know the compressed size or
    /// ZIP64 promotion of a changed member. The preservation profile is
    /// produced after those facts are final and before the first sink write,
    /// so this pass closes that gap without retaining archive bytes.
    pub(crate) fn admit_preservation_profile(
        &self,
        profile: &soapberry_zip::PreservationOutputProfile<'_>,
    ) -> Result<()> {
        self.check_recipient()?;
        let limits = self.recipient_limits;
        limits.check(
            SourceLimitKind::InputBytes,
            profile.archive_bytes(),
            limits.max_input_bytes(),
        )?;
        limits.check(
            SourceLimitKind::ArchiveTotalEntries,
            profile.entry_count(),
            limits.max_archive_total_entries() as u64,
        )?;
        let metadata_bytes = profile
            .prefix_bytes()
            .checked_add(profile.central_directory_bytes())
            .and_then(|bytes| bytes.checked_add(profile.tail_bytes()))
            .and_then(|bytes| bytes.checked_add(profile.archive_comment_bytes()))
            .ok_or_else(|| overlay_unavailable("recipient ZIP metadata bytes overflow"))?;
        let geometry = profile
            .prefix_bytes()
            .checked_add(profile.local_bytes())
            .and_then(|bytes| bytes.checked_add(profile.central_directory_bytes()))
            .and_then(|bytes| bytes.checked_add(profile.tail_bytes()))
            .and_then(|bytes| bytes.checked_add(profile.archive_comment_bytes()))
            .ok_or_else(|| overlay_unavailable("recipient ZIP archive geometry overflows"))?;
        if geometry != profile.archive_bytes() {
            return Err(overlay_unavailable(
                "recipient ZIP output profile geometry is inconsistent",
            ));
        }
        limits.check(
            SourceLimitKind::ArchiveMetadataBytes,
            metadata_bytes,
            limits.max_archive_metadata_bytes(),
        )?;

        let mut aggregate_bytes = 0_u64;
        let mut central_metadata_bytes = 0_u64;
        if u64::try_from(profile.entries().len()).ok() != Some(profile.entry_count()) {
            return Err(overlay_unavailable(
                "recipient ZIP entry count profile is inconsistent",
            ));
        }
        for entry in profile.entries() {
            let variable_metadata = entry
                .raw_name_len()
                .checked_add(entry.extra_len())
                .and_then(|bytes| bytes.checked_add(entry.comment_len()))
                .ok_or_else(|| overlay_unavailable("recipient ZIP entry metadata overflows"))?;
            if entry.metadata_len() != variable_metadata
                || entry.central_record_len() < variable_metadata
            {
                return Err(overlay_unavailable(
                    "recipient ZIP entry metadata profile is inconsistent",
                ));
            }
            central_metadata_bytes = central_metadata_bytes
                .checked_add(entry.metadata_len())
                .ok_or_else(|| overlay_unavailable("recipient ZIP metadata bytes overflow"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                entry.raw_name_len(),
                limits.max_archive_member_name_bytes(),
            )?;
            limits.check(
                SourceLimitKind::ArchiveCompressedBytes,
                entry.compressed_size(),
                limits.max_archive_compressed_bytes(),
            )?;
            limits.check(
                SourceLimitKind::EntryBytes,
                entry.uncompressed_size(),
                limits.max_entry_bytes(),
            )?;
            aggregate_bytes = aggregate_bytes
                .checked_add(entry.uncompressed_size())
                .ok_or_else(|| overlay_unavailable("recipient ZIP aggregate bytes overflow"))?;
        }
        if central_metadata_bytes != profile.central_metadata_bytes() {
            return Err(overlay_unavailable(
                "recipient ZIP central metadata profile is inconsistent",
            ));
        }
        limits.check(
            SourceLimitKind::AggregateBytes,
            aggregate_bytes,
            limits.max_aggregate_bytes(),
        )?;
        Ok(())
    }

    /// Admit a prepared effective topology against the recipient profile.
    ///
    /// A prepared plan may remove source members or replace their logical
    /// payloads, so checking the producer's original catalog here would both
    /// over-refuse valid shrinking plans and miss replacement-specific XML and
    /// graph facts.  This census uses the retained effective catalog and
    /// scalar plan metadata.  ZIP central-directory metadata for retained
    /// members remains exact; generated compressed sizes are deliberately
    /// checked by the writer's generated-entry admission hook before that
    /// hook exposes a publication plan.
    pub(super) fn admit_prepared_topology(
        &self,
        package: &SourceBackedPackage,
        state: &super::PreparedTopologyState,
    ) -> Result<()> {
        self.check_recipient()?;
        package.source.ensure_current()?;
        package.cache.check_context().map_err(map_execution_error)?;
        if state.no_op {
            return self.admit_source_catalog(package);
        }
        let limits = self.recipient_limits;
        let physical_members = build_admission_physical_lookup(&package.source, &package.archive)?;
        let folded_parts = build_admission_part_lookup(&package.source, &package.catalog)?;
        let effective = state.effective_data();
        let retained = state.retained();
        let totals = state
            .totals
            .as_ref()
            .ok_or_else(|| overlay_unavailable("prepared topology totals are unavailable"))?;

        let part_count = u64::try_from(effective.parts.len())
            .map_err(|_| overlay_unavailable("recipient effective Part count exceeds u64"))?;
        limits.check(
            SourceLimitKind::Parts,
            part_count,
            limits.max_parts() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalPartBytes,
            totals.parts,
            limits.max_total_part_bytes(),
        )?;

        let is_omitted = |name: &str| {
            retained
                .omitted_members
                .iter()
                .any(|omitted| omitted.eq_ignore_ascii_case(name))
        };
        let changed_payload = |name: &str| {
            state.changed.iter().find_map(|change| {
                (super::replacement_member_name(change, &package.catalog.parts)
                    .eq_ignore_ascii_case(name))
                .then(|| change.replacement.as_slice())
            })
        };
        let appended_payload = |name: &str| {
            retained
                .policy_appended
                .iter()
                .find(|member| member.name.eq_ignore_ascii_case(name))
                .map(|member| member.payload.as_slice())
        };

        let mut aggregate_bytes = 0_u64;
        let mut relationship_parts = 0_u64;
        let mut retained_members = 0_u64;
        for (index, (entry_id, name)) in package.archive.file_entries().enumerate() {
            if index & 0xff == 0 {
                package.source.ensure_current()?;
                package.cache.check_context().map_err(map_execution_error)?;
            }
            if is_omitted(name) {
                continue;
            }
            retained_members = retained_members
                .checked_add(1)
                .ok_or_else(|| overlay_unavailable("recipient member count overflows"))?;
            let name_bytes = u64::try_from(name.len())
                .map_err(|_| overlay_unavailable("recipient member name exceeds u64"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                name_bytes,
                limits.max_archive_member_name_bytes(),
            )?;
            let metadata = package.archive.metadata_for(entry_id)?;
            let bytes = changed_payload(name)
                .map_or(metadata.uncompressed_size(), |payload| payload.len() as u64);
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            // A changed member is regenerated with the source compression
            // method.  Store's compressed size is exact; Deflate size is
            // admitted by the generated-entry writer immediately before its
            // plan is exposed, rather than guessed from the source member.
            if changed_payload(name).is_none() {
                limits.check(
                    SourceLimitKind::ArchiveCompressedBytes,
                    metadata.compressed_size(),
                    limits.max_archive_compressed_bytes(),
                )?;
            }
            aggregate_bytes = aggregate_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient aggregate bytes overflow"))?;
            if super::is_relationship_member_name(name) {
                relationship_parts = relationship_parts
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("recipient relationship parts overflow"))?;
            }
        }
        let generated_names = effective.added_member_names.as_ref();
        if generated_names.len() != retained.policy_appended.len()
            || generated_names.iter().any(|name| {
                !retained
                    .policy_appended
                    .iter()
                    .any(|member| member.name.eq_ignore_ascii_case(name))
            })
        {
            return Err(overlay_unavailable(
                "prepared appended member metadata is not bound to its emitted entries",
            ));
        }
        for name in generated_names {
            let name_bytes = u64::try_from(name.len())
                .map_err(|_| overlay_unavailable("generated member name exceeds u64"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                name_bytes,
                limits.max_archive_member_name_bytes(),
            )?;
            let payload = retained
                .policy_appended
                .iter()
                .find(|member| member.name.eq_ignore_ascii_case(name))
                .map(|member| member.payload.as_slice())
                .ok_or_else(|| {
                    overlay_unavailable("generated member payload metadata is unavailable")
                })?;
            let bytes = payload.len() as u64;
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            aggregate_bytes = aggregate_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient aggregate bytes overflow"))?;
            if super::is_relationship_member_name(name) {
                relationship_parts = relationship_parts
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("recipient relationship parts overflow"))?;
            }
        }
        let member_count = retained_members
            .checked_add(
                u64::try_from(generated_names.len())
                    .map_err(|_| overlay_unavailable("generated member count exceeds u64"))?,
            )
            .ok_or_else(|| overlay_unavailable("recipient member count overflows"))?;
        limits.check(
            SourceLimitKind::Members,
            member_count,
            limits.max_members() as u64,
        )?;
        let source_entry_count = u64::try_from(package.archive.preservation_entry_count())
            .map_err(|_| overlay_unavailable("source ZIP entry count exceeds u64"))?;
        let omitted_entry_count = u64::try_from(retained.omitted_members.len())
            .map_err(|_| overlay_unavailable("omitted ZIP entry count exceeds u64"))?;
        let final_entry_count = source_entry_count
            .checked_sub(omitted_entry_count)
            .and_then(|count| count.checked_add(generated_names.len() as u64))
            .ok_or_else(|| overlay_unavailable("recipient ZIP entry count underflows"))?;
        limits.check(
            SourceLimitKind::ArchiveTotalEntries,
            final_entry_count,
            limits.max_archive_total_entries() as u64,
        )?;
        limits.check(
            SourceLimitKind::AggregateBytes,
            aggregate_bytes,
            limits.max_aggregate_bytes(),
        )?;
        limits.check(
            SourceLimitKind::RelationshipParts,
            relationship_parts,
            limits.max_relationship_parts() as u64,
        )?;

        let mut total_part_bytes = 0_u64;
        for (index, part) in effective.parts.iter().enumerate() {
            if index & 0xff == 0 {
                package.source.ensure_current()?;
                package.cache.check_context().map_err(map_execution_error)?;
            }
            let name_bytes = u64::try_from(part.partname(package).membername().len())
                .map_err(|_| overlay_unavailable("recipient Part name exceeds u64"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                name_bytes,
                limits.max_archive_member_name_bytes(),
            )?;
            let bytes = if let Some(payload) = part.replacement.as_deref() {
                payload.len() as u64
            } else {
                let source_index = part
                    .source_index
                    .ok_or_else(|| overlay_unavailable("effective Part payload is unavailable"))?;
                package
                    .archive
                    .metadata_for(package.catalog.parts[source_index].entry_id)?
                    .uncompressed_size()
            };
            limits.check(SourceLimitKind::PartBytes, bytes, limits.max_part_bytes())?;
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            total_part_bytes = total_part_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient Part bytes overflow"))?;
        }
        limits.check(
            SourceLimitKind::TotalPartBytes,
            total_part_bytes,
            limits.max_total_part_bytes(),
        )?;
        let package_relationships = effective
            .package_relationships
            .as_deref()
            .unwrap_or(&package.catalog.package_relationships);

        let content_types_name = package.catalog.content_types_member.as_str();
        let changed_content_types = changed_payload(content_types_name);
        let content_types_bytes = if let Some(payload) = changed_content_types {
            admit_content_types_xml(payload, limits)?;
            payload.len() as u64
        } else {
            let entry = package
                .archive
                .entry_id(content_types_name)
                .ok_or_else(|| OpcError::PartNotFound(content_types_name.to_owned()))?;
            let bytes = package.archive.metadata_for(entry)?.uncompressed_size();
            if let Some(xml) = read_structural_member(
                &package.source,
                &package.archive,
                &physical_members,
                content_types_name,
                SourceLimitKind::ContentTypesBytes,
                limits.max_content_types_bytes() as u64,
                limits,
            )? {
                admit_content_types_xml(&xml.bytes, limits)?;
            }
            bytes
        };
        limits.check(
            SourceLimitKind::ContentTypesBytes,
            content_types_bytes,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::PartBytes,
            content_types_bytes,
            limits.max_part_bytes(),
        )?;
        limits.check(
            SourceLimitKind::EntryBytes,
            content_types_bytes,
            limits.max_entry_bytes(),
        )?;
        let mapping_count = effective
            .content_types
            .as_ref()
            .map_or(package.catalog.content_types.mapping_count(), |map| {
                map.mapping_count()
            });
        limits.check(
            SourceLimitKind::ContentTypeMappings,
            mapping_count as u64,
            limits.max_content_type_mappings() as u64,
        )?;

        let mut relationship_xml_bytes = 0_u64;
        let mut relationship_events = 0_u64;
        let mut total_relationships = 0_u64;
        let mut relationship_owners = HashSet::new();
        for owner in package.catalog.relationship_xml_metrics.keys() {
            relationship_owners.insert(owner.clone());
        }
        for (_, member_name) in package.archive.file_entries() {
            if let Some(owner) = relationship_owner_from_member(member_name)? {
                relationship_owners.insert(canonical_relationship_owner(
                    &package.catalog,
                    &folded_parts,
                    &owner,
                )?);
            }
        }
        relationship_owners
            .insert(PackURI::new(super::PACKAGE_URI).map_err(OpcError::InvalidPackUri)?);
        for part in &effective.parts {
            relationship_owners.insert(part.partname(package).clone());
        }
        for name in generated_names.iter() {
            if let Some(owner) = relationship_owner_from_member(name)? {
                relationship_owners.insert(canonical_relationship_owner(
                    &package.catalog,
                    &folded_parts,
                    &owner,
                )?);
            }
        }
        let mut admitted_relationships = HashMap::new();
        for (index, owner) in relationship_owners.into_iter().enumerate() {
            if index & 0xff == 0 {
                package.source.ensure_current()?;
                package.cache.check_context().map_err(map_execution_error)?;
            }
            let member = relationship_member_name(&owner)?;
            // An omitted source member may be reintroduced as a generated
            // final member. Census the final payload before honoring the
            // omission; otherwise a replacement relationship part vanishes
            // from the XML and graph limits while publication still emits it.
            let replacement = changed_payload(&member).or_else(|| appended_payload(&member));
            if is_omitted(&member) && replacement.is_none() {
                continue;
            }
            let (bytes, events, relationships) = if let Some(xml) = replacement {
                let relationships = parse_relationship_xml(xml, &owner, limits)?;
                let bytes = xml.len() as u64;
                (
                    bytes,
                    super::relationship_xml_event_count(xml, limits)?,
                    relationships,
                )
            } else if package
                .catalog
                .relationship_xml_metrics
                .contains_key(&owner)
            {
                let Some(xml) = read_structural_member(
                    &package.source,
                    &package.archive,
                    &physical_members,
                    &member,
                    SourceLimitKind::RelationshipXmlBytes,
                    limits.max_relationship_xml_bytes() as u64,
                    limits,
                )?
                else {
                    return Err(overlay_unavailable(
                        "source relationship XML is unavailable for effective topology",
                    ));
                };
                let relationships = parse_relationship_xml(&xml.bytes, &owner, limits)?;
                let bytes = u64::try_from(xml.bytes.len())
                    .map_err(|_| overlay_unavailable("relationship XML length exceeds u64"))?;
                let events = super::relationship_xml_event_count(&xml.bytes, limits)?;
                (bytes, events, relationships)
            } else {
                let Some(xml) = read_structural_member(
                    &package.source,
                    &package.archive,
                    &physical_members,
                    &member,
                    SourceLimitKind::RelationshipXmlBytes,
                    limits.max_relationship_xml_bytes() as u64,
                    limits,
                )?
                else {
                    continue;
                };
                let relationships = parse_relationship_xml(&xml.bytes, &owner, limits)?;
                let bytes = u64::try_from(xml.bytes.len())
                    .map_err(|_| overlay_unavailable("relationship XML length exceeds u64"))?;
                let events = super::relationship_xml_event_count(&xml.bytes, limits)?;
                (bytes, events, relationships)
            };
            // Count and validate the parsed final XML before retaining the map.
            // This includes physical relationship parts that are not represented
            // in the source catalog (orphan .rels members).
            check_relationships(limits, &relationships, &mut total_relationships)?;
            limits.check(
                SourceLimitKind::RelationshipXmlBytes,
                bytes,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            limits.check(SourceLimitKind::Events, events, limits.max_events() as u64)?;
            relationship_xml_bytes = relationship_xml_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient relationship bytes overflow"))?;
            relationship_events = relationship_events
                .checked_add(events)
                .ok_or_else(|| overlay_unavailable("recipient relationship events overflow"))?;
            admitted_relationships.insert(owner, relationships);
        }
        let package_owner = PackURI::new(super::PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        if let std::collections::hash_map::Entry::Vacant(entry) =
            admitted_relationships.entry(package_owner)
        {
            let relationships = package_relationships.clone();
            check_relationships(limits, &relationships, &mut total_relationships)?;
            entry.insert(relationships);
        }
        for part in &effective.parts {
            let owner = part.partname(package).clone();
            if let std::collections::hash_map::Entry::Vacant(entry) =
                admitted_relationships.entry(owner)
            {
                let relationships = part.relationships(package).clone();
                check_relationships(limits, &relationships, &mut total_relationships)?;
                entry.insert(relationships);
            }
        }
        limits.check(
            SourceLimitKind::TotalRelationships,
            total_relationships,
            limits.max_total_relationships() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalRelationshipXmlBytes,
            relationship_xml_bytes,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalRelationshipXmlEvents,
            relationship_events,
            limits.max_total_relationship_xml_events() as u64,
        )?;
        let admitted_relationships_by_owner =
            folded_relationship_lookup(&package.source, &admitted_relationships)?;

        // Traverse the effective graph, checking before each new target is
        // retained so the admission itself cannot grow beyond the recipient
        // graph bound.
        let mut graph = HashSet::new();
        let mut work = Vec::new();
        enqueue_graph_targets(package_relationships, &mut graph, &mut work, limits)?;
        while let Some(owner) = work.pop() {
            if let Some(relationships) = admitted_relationships_by_owner.by_folded_name.get(&owner)
            {
                enqueue_graph_targets(relationships, &mut graph, &mut work, limits)?;
            }
        }
        Ok(())
    }

    /// Admit a direct replacement/deletion publication against the final
    /// physical member set.  This path has no prepared effective catalog, so
    /// it derives the final scalar gauges from source metadata plus the
    /// replacement bytes and reparses changed relationship XML before output.
    pub(super) fn admit_changed_overlays(
        &self,
        package: &SourceBackedPackage,
        replacements: &[super::ChangedOverlay],
        omitted_members: &[String],
        appended: &[super::PolicyAppendedEntry],
    ) -> Result<()> {
        self.check_recipient()?;
        package.source.ensure_current()?;
        package.cache.check_context().map_err(map_execution_error)?;
        let limits = self.recipient_limits;
        let physical_members = build_admission_physical_lookup(&package.source, &package.archive)?;
        let folded_parts = build_admission_part_lookup(&package.source, &package.catalog)?;
        let content_types_name = package.catalog.content_types_member.as_str();
        let is_omitted = |name: &str| {
            omitted_members
                .iter()
                .any(|omitted| omitted.eq_ignore_ascii_case(name))
        };
        let changed_payload = |name: &str| {
            replacements.iter().find_map(|change| {
                (super::replacement_member_name(change, &package.catalog.parts)
                    .eq_ignore_ascii_case(name))
                .then(|| change.replacement.as_slice())
            })
        };
        let appended_payload = |name: &str| {
            appended
                .iter()
                .find(|member| member.name.eq_ignore_ascii_case(name))
                .map(|member| member.payload.as_slice())
        };
        let final_payload = |name: &str| changed_payload(name).or_else(|| appended_payload(name));
        let mut appended_names = HashSet::new();
        appended_names
            .try_reserve(appended.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC appended member names",
                source,
            })?;
        for member in appended {
            let folded = super::folded_ascii_name(
                &member.name,
                "source-backed OPC folded appended member name",
            )?;
            if !appended_names.insert(folded) {
                return Err(overlay_unavailable(
                    "multiple appended members resolve to one ASCII-folded name",
                ));
            }
            if !is_omitted(&member.name)
                && (package
                    .archive
                    .file_names()
                    .any(|name| name.eq_ignore_ascii_case(&member.name))
                    || changed_payload(&member.name).is_some())
            {
                return Err(overlay_unavailable(
                    "appended member collides with a retained source or replacement member",
                ));
            }
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                member.name.len() as u64,
                limits.max_archive_member_name_bytes(),
            )?;
            let bytes = member.payload.len() as u64;
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
        }
        let mut members = 0_u64;
        let mut relationships_parts = 0_u64;
        let mut aggregate_bytes = 0_u64;
        for (index, (entry_id, name)) in package.archive.file_entries().enumerate() {
            if index & 0xff == 0 {
                package.source.ensure_current()?;
                package.cache.check_context().map_err(map_execution_error)?;
            }
            if is_omitted(name) {
                continue;
            }
            members = members
                .checked_add(1)
                .ok_or_else(|| overlay_unavailable("recipient member count overflows"))?;
            limits.check(
                SourceLimitKind::ArchiveMemberNameBytes,
                name.len() as u64,
                limits.max_archive_member_name_bytes(),
            )?;
            let metadata = package.archive.metadata_for(entry_id)?;
            let changed = changed_payload(name);
            let bytes =
                changed.map_or(metadata.uncompressed_size(), |payload| payload.len() as u64);
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            if changed.is_none() {
                limits.check(
                    SourceLimitKind::ArchiveCompressedBytes,
                    metadata.compressed_size(),
                    limits.max_archive_compressed_bytes(),
                )?;
            }
            aggregate_bytes = aggregate_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient aggregate bytes overflow"))?;
            if super::is_relationship_member_name(name) {
                relationships_parts = relationships_parts
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("recipient relationship parts overflow"))?;
            }
        }
        for member in appended {
            members = members
                .checked_add(1)
                .ok_or_else(|| overlay_unavailable("recipient member count overflows"))?;
            aggregate_bytes = aggregate_bytes
                .checked_add(member.payload.len() as u64)
                .ok_or_else(|| overlay_unavailable("recipient aggregate bytes overflow"))?;
            if super::is_relationship_member_name(&member.name) {
                relationships_parts = relationships_parts
                    .checked_add(1)
                    .ok_or_else(|| overlay_unavailable("recipient relationship parts overflow"))?;
            }
        }
        limits.check(
            SourceLimitKind::Members,
            members,
            limits.max_members() as u64,
        )?;
        let source_entry_count = u64::try_from(package.archive.preservation_entry_count())
            .map_err(|_| overlay_unavailable("source ZIP entry count exceeds u64"))?;
        let omitted_entry_count = u64::try_from(omitted_members.len())
            .map_err(|_| overlay_unavailable("omitted ZIP entry count exceeds u64"))?;
        let final_entry_count = source_entry_count
            .checked_sub(omitted_entry_count)
            .and_then(|count| count.checked_add(appended.len() as u64))
            .ok_or_else(|| overlay_unavailable("recipient ZIP entry count underflows"))?;
        limits.check(
            SourceLimitKind::ArchiveTotalEntries,
            final_entry_count,
            limits.max_archive_total_entries() as u64,
        )?;
        limits.check(
            SourceLimitKind::AggregateBytes,
            aggregate_bytes,
            limits.max_aggregate_bytes(),
        )?;
        limits.check(
            SourceLimitKind::RelationshipParts,
            relationships_parts,
            limits.max_relationship_parts() as u64,
        )?;

        let mut total_part_bytes = 0_u64;
        for (index, part) in package.catalog.parts.iter().enumerate() {
            if index & 0xff == 0 {
                package.source.ensure_current()?;
                package.cache.check_context().map_err(map_execution_error)?;
            }
            let member = part.partname.membername();
            if is_omitted(member) {
                continue;
            }
            let changed = changed_payload(member);
            let bytes = changed.map_or(
                package
                    .archive
                    .metadata_for(part.entry_id)?
                    .uncompressed_size(),
                |payload| payload.len() as u64,
            );
            limits.check(SourceLimitKind::PartBytes, bytes, limits.max_part_bytes())?;
            limits.check(SourceLimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
            total_part_bytes = total_part_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient Part bytes overflow"))?;
        }
        let mut appended_part_count = 0_u64;
        for member in appended {
            if member.name.eq_ignore_ascii_case(content_types_name)
                || super::is_relationship_member_name(&member.name)
            {
                continue;
            }
            appended_part_count = appended_part_count
                .checked_add(1)
                .ok_or_else(|| overlay_unavailable("recipient Part count overflows"))?;
            let bytes = member.payload.len() as u64;
            limits.check(SourceLimitKind::PartBytes, bytes, limits.max_part_bytes())?;
            total_part_bytes = total_part_bytes
                .checked_add(bytes)
                .ok_or_else(|| overlay_unavailable("recipient Part bytes overflow"))?;
        }
        let part_count = package
            .catalog
            .parts
            .iter()
            .filter(|part| !is_omitted(part.partname.membername()))
            .count();
        let part_count = u64::try_from(part_count)
            .map_err(|_| overlay_unavailable("recipient Part count exceeds u64"))?
            .checked_add(appended_part_count)
            .ok_or_else(|| overlay_unavailable("recipient Part count overflows"))?;
        limits.check(
            SourceLimitKind::Parts,
            part_count,
            limits.max_parts() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalPartBytes,
            total_part_bytes,
            limits.max_total_part_bytes(),
        )?;

        let (content_types_bytes, content_type_mapping_count) =
            if let Some(payload) = final_payload(content_types_name) {
                let parsed = super::ContentTypeMap::from_xml(payload, limits)?;
                (payload.len() as u64, parsed.mapping_count() as u64)
            } else {
                let entry = package
                    .archive
                    .entry_id(content_types_name)
                    .ok_or_else(|| OpcError::PartNotFound(content_types_name.to_owned()))?;
                let bytes = package.archive.metadata_for(entry)?.uncompressed_size();
                if let Some(xml) = read_structural_member(
                    &package.source,
                    &package.archive,
                    &physical_members,
                    content_types_name,
                    SourceLimitKind::ContentTypesBytes,
                    limits.max_content_types_bytes() as u64,
                    limits,
                )? {
                    admit_content_types_xml(&xml.bytes, limits)?;
                }
                (bytes, package.catalog.content_types.mapping_count() as u64)
            };
        limits.check(
            SourceLimitKind::ContentTypesBytes,
            content_types_bytes,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::PartBytes,
            content_types_bytes,
            limits.max_part_bytes(),
        )?;
        limits.check(
            SourceLimitKind::EntryBytes,
            content_types_bytes,
            limits.max_entry_bytes(),
        )?;
        limits.check(
            SourceLimitKind::ContentTypeMappings,
            content_type_mapping_count,
            limits.max_content_type_mappings() as u64,
        )?;

        let mut relationship_xml_bytes = 0_u64;
        let mut relationship_events = 0_u64;
        let mut relationship_total = 0_u64;
        let mut relationship_map = HashMap::new();
        let mut relationship_owners = HashSet::new();
        for owner in package.catalog.relationship_xml_metrics.keys() {
            relationship_owners.insert(owner.clone());
        }
        for (_, member_name) in package.archive.file_entries() {
            if let Some(owner) = relationship_owner_from_member(member_name)? {
                relationship_owners.insert(canonical_relationship_owner(
                    &package.catalog,
                    &folded_parts,
                    &owner,
                )?);
            }
        }
        for member in appended {
            if let Some(owner) = relationship_owner_from_member(&member.name)? {
                relationship_owners.insert(canonical_relationship_owner(
                    &package.catalog,
                    &folded_parts,
                    &owner,
                )?);
            }
        }
        for owner in relationship_owners {
            let member = relationship_member_name(&owner)?;
            let replacement = final_payload(&member);
            if is_omitted(&member) && replacement.is_none() {
                continue;
            }
            let relationships = if let Some(xml) = replacement {
                let parsed = PackageReader::parse_owned_relationships(xml, &owner, limits)?;
                let events = super::relationship_xml_event_count(xml, limits)?;
                relationship_xml_bytes = relationship_xml_bytes
                    .checked_add(xml.len() as u64)
                    .ok_or_else(|| overlay_unavailable("recipient relationship bytes overflow"))?;
                relationship_events = relationship_events
                    .checked_add(events)
                    .ok_or_else(|| overlay_unavailable("recipient relationship events overflow"))?;
                parsed
            } else {
                let Some(xml) = read_structural_member(
                    &package.source,
                    &package.archive,
                    &physical_members,
                    &member,
                    SourceLimitKind::RelationshipXmlBytes,
                    limits.max_relationship_xml_bytes() as u64,
                    limits,
                )?
                else {
                    return Err(overlay_unavailable(
                        "source relationship XML is unavailable for overlay admission",
                    ));
                };
                let parsed = parse_relationship_xml(&xml.bytes, &owner, limits)?;
                let events = super::relationship_xml_event_count(&xml.bytes, limits)?;
                relationship_xml_bytes = relationship_xml_bytes
                    .checked_add(xml.bytes.len() as u64)
                    .ok_or_else(|| overlay_unavailable("recipient relationship bytes overflow"))?;
                relationship_events = relationship_events
                    .checked_add(events)
                    .ok_or_else(|| overlay_unavailable("recipient relationship events overflow"))?;
                parsed
            };
            check_relationships(limits, &relationships, &mut relationship_total)?;
            relationship_map.insert(owner.clone(), relationships);
        }
        limits.check(
            SourceLimitKind::TotalRelationshipXmlBytes,
            relationship_xml_bytes,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalRelationshipXmlEvents,
            relationship_events,
            limits.max_total_relationship_xml_events() as u64,
        )?;
        limits.check(
            SourceLimitKind::TotalRelationships,
            relationship_total,
            limits.max_total_relationships() as u64,
        )?;
        let relationship_map_by_owner =
            folded_relationship_lookup(&package.source, &relationship_map)?;

        let root_owner = PackURI::new(super::PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let root_key = super::folded_ascii_name(
            root_owner.as_str(),
            "source-backed OPC folded relationship graph owner",
        )?;
        let Some(root) = relationship_map_by_owner.by_folded_name.get(&root_key) else {
            return Ok(());
        };
        let mut graph = HashSet::new();
        let mut work = Vec::new();
        enqueue_graph_targets(root, &mut graph, &mut work, limits)?;
        while let Some(owner) = work.pop() {
            if let Some(relationships) = relationship_map_by_owner.by_folded_name.get(&owner) {
                enqueue_graph_targets(relationships, &mut graph, &mut work, limits)?;
            }
        }
        Ok(())
    }
}

fn check_relationships(
    limits: SourceLimits,
    relationships: &super::Relationships,
    total: &mut u64,
) -> Result<()> {
    let count = u64::try_from(relationships.len())
        .map_err(|_| overlay_unavailable("recipient relationship count exceeds u64"))?;
    limits.check(
        SourceLimitKind::RelationshipsPerPart,
        count,
        limits.max_relationships_per_part() as u64,
    )?;
    *total = total
        .checked_add(count)
        .ok_or_else(|| overlay_unavailable("recipient relationship count overflows"))?;
    limits.check(
        SourceLimitKind::TotalRelationships,
        *total,
        limits.max_total_relationships() as u64,
    )?;
    for relationship in relationships.iter() {
        for value in [
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        ] {
            limits.check(
                SourceLimitKind::XmlAttributeBytes,
                value.len() as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            limits.check(
                SourceLimitKind::XmlAttributeBytes,
                super::escaped_xml_attribute_len(value)? as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
        }
        limits.check(
            SourceLimitKind::RelationshipTargetBytes,
            relationship.target_ref().len() as u64,
            limits.max_relationship_target_bytes() as u64,
        )?;
    }
    Ok(())
}

fn enqueue_graph_targets(
    relationships: &super::Relationships,
    graph: &mut HashSet<String>,
    work: &mut Vec<String>,
    limits: SourceLimits,
) -> Result<()> {
    for relationship in relationships.iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        let folded = super::folded_part_name(&target)?;
        if graph.insert(folded.clone()) {
            let count = u64::try_from(graph.len())
                .map_err(|_| overlay_unavailable("recipient graph node count exceeds u64"))?;
            limits.check(
                SourceLimitKind::RelationshipGraphNodes,
                count,
                limits.max_relationship_graph_nodes() as u64,
            )?;
            work.push(folded);
        }
    }
    Ok(())
}

pub(crate) enum ArtifactPolicySink<W> {
    Budgeted(OutputBudgetedSink<ContextCheckedSink<ContextCheckedSink<SourceCheckedSink<W>>>>),
    Cooperative(ContextCheckedSink<ContextCheckedSink<SourceCheckedSink<W>>>),
}

impl<W: Write> Write for ArtifactPolicySink<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        match self {
            Self::Budgeted(sink) => sink.write(bytes),
            Self::Cooperative(sink) => sink.write(bytes),
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        match self {
            Self::Budgeted(sink) => sink.flush(),
            Self::Cooperative(sink) => sink.flush(),
        }
    }
}

pub(crate) fn policy_source_sink<W: Write>(
    policy: &ArtifactStreamPolicy<'_>,
    writer: W,
    source: &SourceSnapshot,
    producer_context: Option<&ExecutionContext>,
) -> Result<ArtifactPolicySink<W>> {
    let failure = policy.failure();
    let checked = SourceCheckedSink {
        inner: writer,
        snapshot: source.clone(),
    };
    let producer = ContextCheckedSink {
        inner: checked,
        context: producer_context.cloned(),
        failure: Arc::clone(&failure),
    };
    let recipient = ContextCheckedSink {
        inner: producer,
        context: policy.recipient_context().cloned(),
        failure,
    };
    if policy.is_publish() && policy.recipient_context().is_some() {
        let output_reservation_failures = policy
            .recipient_output_reservation_failures
            .clone()
            .ok_or_else(|| {
                overlay_unavailable("managed artifact stream output counter is unavailable")
            })?;
        Ok(ArtifactPolicySink::Budgeted(OutputBudgetedSink {
            inner: recipient,
            context: policy
                .recipient_context()
                .expect("recipient context checked above")
                .clone(),
            failure: policy.failure(),
            output_reservation_failures,
        }))
    } else {
        Ok(ArtifactPolicySink::Cooperative(recipient))
    }
}

/// Copy one exact source snapshot with producer input/work checks and the
/// recipient output policy supplied separately.
pub(crate) fn write_exact_snapshot_with_policy<W: Write>(
    source: &SourceSnapshot,
    writer: W,
    policy: &ArtifactStreamPolicy<'_>,
) -> Result<()> {
    policy.check_recipient()?;
    if let Some(context) = source.context.as_ref() {
        context.check().map_err(map_execution_error)?;
    }
    let _monitored = source.monitor_publication();
    source.ensure_current()?;
    let _workspace = reserve_policy_workspace(source.context.as_ref())?;
    const STACK_BYTES: usize = 4 * 1024;
    let mut buffer = [0_u8; STACK_BYTES];
    let mut written = 0_u64;
    let mut sink = policy_source_sink(
        policy,
        Counted::new(writer, &mut written),
        source,
        source.context.as_ref(),
    )?;
    let mut offset = 0_u64;
    let result = (|| {
        while offset < source.length {
            if let Some(context) = source.context.as_ref() {
                context.check().map_err(map_execution_error)?;
            }
            let count = usize::try_from((source.length - offset).min(buffer.len() as u64))
                .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
            let read = read_source_at_with_context(
                source,
                source.context.as_ref(),
                offset,
                &mut buffer[..count],
                "policy publication",
            )?;
            if read == 0 {
                return Err(OpcError::IoError(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "source-backed OPC source ended during policy publication",
                )));
            }
            if let Some(context) = source.context.as_ref() {
                context
                    .consume(Resource::Work, read as u64)
                    .map_err(map_execution_error)?;
            }
            sink.write_all(&buffer[..read]).map_err(map_io_error)?;
            offset = offset
                .checked_add(read as u64)
                .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
        }
        sink.flush().map_err(map_io_error)
    })();
    let result = policy
        .take_failure()
        .map_or(result, |error| Err(map_execution_error(error)));
    finish_source_publication(result, source, written)
}

fn reserve_policy_workspace(context: Option<&ExecutionContext>) -> Result<Option<Reservation>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let requested = SOURCE_PUBLICATION_CHUNK_BYTES as u64;
    match context.reserve(Resource::Memory, requested) {
        Ok(reservation) => Ok(Some(reservation)),
        Err(ExecutionError::ResourceLimit(limit)) if limit.resource == Resource::Memory => {
            let used = limit.observed.saturating_sub(requested);
            let available = limit.limit.saturating_sub(used);
            if available == 0 {
                Ok(None)
            } else {
                Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            }
        },
        Err(error) => Err(map_execution_error(error)),
    }
}

impl ArtifactStream for SourceArtifact {
    type Error = OpcError;

    fn check_fresh(&self) -> Result<()> {
        self.snapshot.ensure_current()
    }

    fn stream_to(
        &self,
        target: &mut ArtifactStreamTarget<'_>,
    ) -> std::result::Result<ArtifactStreamReport, Self::Error> {
        target.publish_source_artifact(self)
    }
}

/// Complete artifact identities for a durable exact inverse.
///
/// These identities survive reopening through a different positional source.
/// Runtime source versions are checked separately throughout the operation.
/// A format-owned patch must authenticate this proof before invoking the
/// low-level restoration capability.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceArtifactRestoreProof {
    pub current_len: u64,
    pub current_sha256: SourceArtifactFingerprint,
    pub original_len: u64,
    pub original_sha256: SourceArtifactFingerprint,
}

struct BoundedProbeSink {
    accepted: u64,
    limit: u64,
}

impl BoundedProbeSink {
    fn new(limit: u64) -> Self {
        Self { accepted: 0, limit }
    }
}

impl Write for BoundedProbeSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let remaining = self.limit.saturating_sub(self.accepted);
        let accepted = bytes
            .len()
            .min(usize::try_from(remaining).unwrap_or(usize::MAX));
        if accepted == 0 {
            return Ok(0);
        }
        self.accepted = self.accepted.saturating_add(accepted as u64);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn preflight_stream<S: ArtifactStream + ?Sized>(
    artifact: &S,
    expected_len: u64,
    expected_fingerprint: SourceArtifactFingerprint,
    name: &'static str,
    policy: &ArtifactStreamPolicy<'_>,
) -> std::result::Result<(), ArtifactStreamError<S::Error>> {
    let policy_error = policy.check_recipient().err().map(ArtifactStreamError::opc);
    let initial_freshness = artifact
        .check_fresh()
        .err()
        .map(ArtifactStreamError::producer);
    let initial_error = match (policy_error, initial_freshness) {
        (None, None) => None,
        (Some(error), None) | (None, Some(error)) => Some(error),
        (Some(policy), Some(freshness)) => {
            Some(ArtifactStreamError::incomplete(0, policy, Some(freshness)))
        },
    };
    if let Some(initial_error) = initial_error {
        let final_freshness = artifact
            .check_fresh()
            .err()
            .map(ArtifactStreamError::producer);
        return Err(with_secondary(initial_error, final_freshness));
    }
    let mut probe_sink = BoundedProbeSink::new(expected_len.saturating_add(1));
    let mut target = ArtifactStreamTarget::new(&mut probe_sink, policy);
    let stream_result = artifact
        .stream_to(&mut target)
        .map_err(ArtifactStreamError::producer);
    let result = match stream_result {
        Ok(publication) => target
            .validate_publication(publication)
            .map_err(ArtifactStreamError::opc),
        Err(error) => Err(error),
    };
    let accepted = target.written();
    let fingerprint = target.fingerprint();
    drop(target);
    let policy_result = policy
        .take_failure()
        .map(|error| ArtifactStreamError::opc(map_execution_error(error)));
    let result = match (result, policy_result) {
        (Ok(()), None) => Ok(()),
        (Err(error), None) | (Ok(()), Some(error)) => Err(error),
        (Err(error), Some(policy_error)) => Err(ArtifactStreamError::incomplete(
            0,
            error,
            Some(policy_error),
        )),
    };
    let freshness = artifact
        .check_fresh()
        .map_err(ArtifactStreamError::producer);
    if let Err(freshness) = freshness {
        return Err(match result {
            Ok(()) => freshness,
            Err(error) => ArtifactStreamError::incomplete(0, freshness, Some(error)),
        });
    }
    if let Err(error) = result {
        if accepted > expected_len {
            let mismatch = ArtifactStreamError::opc(OpcError::SourceArtifactMismatch {
                artifact: name,
                field: "length",
            });
            return Err(ArtifactStreamError::incomplete(0, mismatch, Some(error)));
        }
        return Err(error);
    }
    if accepted != expected_len {
        return Err(ArtifactStreamError::opc(OpcError::SourceArtifactMismatch {
            artifact: name,
            field: "length",
        }));
    }
    if fingerprint != expected_fingerprint {
        return Err(ArtifactStreamError::opc(OpcError::SourceArtifactMismatch {
            artifact: name,
            field: "fingerprint",
        }));
    }
    Ok(())
}

fn finish_streamed_publication<SCurrent, SOriginal>(
    result: std::result::Result<(), ArtifactStreamError<SCurrent::Error>>,
    current: &SCurrent,
    original: &SOriginal,
    written: u64,
) -> std::result::Result<(), ArtifactStreamError<SCurrent::Error>>
where
    SCurrent: ArtifactStream + ?Sized,
    SOriginal: ArtifactStream<Error = SCurrent::Error> + ?Sized,
{
    let current_freshness = current.check_fresh().map_err(ArtifactStreamError::producer);
    let original_freshness = original
        .check_fresh()
        .map_err(ArtifactStreamError::producer);
    let freshness = match (current_freshness, original_freshness) {
        (Ok(()), Ok(())) => Ok(()),
        (Err(current), Ok(())) => Err(current),
        (Ok(()), Err(original)) => Err(original),
        (Err(current), Err(original)) => {
            Err(ArtifactStreamError::incomplete(0, current, Some(original)))
        },
    };
    if let Err(freshness) = freshness {
        return match (written != 0, result) {
            (true, Err(error)) | (false, Err(error)) => Err(ArtifactStreamError::incomplete(
                written,
                freshness,
                Some(error),
            )),
            (true, Ok(())) => Err(ArtifactStreamError::incomplete(written, freshness, None)),
            (false, Ok(())) => Err(freshness),
        };
    }
    match result {
        Err(error @ ArtifactStreamError::Publication { .. }) => Err(error),
        Err(error) if written != 0 => Err(ArtifactStreamError::incomplete(written, error, None)),
        other => other,
    }
}

impl SourceBackedPackage {
    /// Probe one producer into a bounded digest sink without accepting caller
    /// output or charging `OutputBytes`.
    pub fn probe_artifact_stream<S: ArtifactStream + ?Sized>(
        &self,
        stream: &S,
        max_output_bytes: u64,
        outer_policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> std::result::Result<(u64, SourceArtifactFingerprint), ArtifactStreamError<S::Error>> {
        self.disable_read_ahead_for_publication()
            .map_err(ArtifactStreamError::opc)?;
        let policy_owned;
        let policy = if let Some(policy) = outer_policy {
            policy
        } else {
            policy_owned =
                ArtifactStreamPolicy::for_recipient(self).map_err(ArtifactStreamError::opc)?;
            &policy_owned
        };
        let probe = policy.probe();
        let policy_error = policy.check_recipient().err().map(ArtifactStreamError::opc);
        let initial_freshness = stream
            .check_fresh()
            .err()
            .map(ArtifactStreamError::producer);
        let initial_error = match (policy_error, initial_freshness) {
            (None, None) => None,
            (Some(error), None) | (None, Some(error)) => Some(error),
            (Some(policy), Some(freshness)) => {
                Some(ArtifactStreamError::incomplete(0, policy, Some(freshness)))
            },
        };
        if let Some(initial_error) = initial_error {
            let final_freshness = stream
                .check_fresh()
                .err()
                .map(ArtifactStreamError::producer);
            return Err(with_secondary(initial_error, final_freshness));
        }
        let mut probe_sink = BoundedProbeSink::new(max_output_bytes.saturating_add(1));
        let mut target = ArtifactStreamTarget::new(&mut probe_sink, &probe);
        let stream_result = stream
            .stream_to(&mut target)
            .map_err(ArtifactStreamError::producer);
        let result = match stream_result {
            Ok(publication) => target
                .validate_publication(publication)
                .map_err(ArtifactStreamError::opc),
            Err(error) => Err(error),
        };
        let accepted = target.written();
        let fingerprint = target.fingerprint();
        drop(target);
        let policy_result = probe
            .take_failure()
            .map(|error| ArtifactStreamError::opc(map_execution_error(error)));
        let result = match (result, policy_result) {
            (Ok(()), None) => Ok(()),
            (Err(error), None) | (Ok(()), Some(error)) => Err(error),
            (Err(error), Some(policy_error)) => Err(ArtifactStreamError::incomplete(
                0,
                error,
                Some(policy_error),
            )),
        };
        let freshness = stream.check_fresh().map_err(ArtifactStreamError::producer);
        if let Err(freshness) = freshness {
            return match result {
                Ok(()) => Err(freshness),
                Err(error) => Err(ArtifactStreamError::incomplete(0, freshness, Some(error))),
            };
        }
        if let Err(error) = result {
            if accepted <= max_output_bytes {
                return Err(error);
            }
            let limit = ArtifactStreamError::opc(OpcError::SourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                actual: accepted,
                maximum: max_output_bytes,
            });
            return Err(ArtifactStreamError::incomplete(0, limit, Some(error)));
        }
        if accepted > max_output_bytes {
            return Err(ArtifactStreamError::opc(OpcError::SourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                actual: accepted,
                maximum: max_output_bytes,
            }));
        }
        Ok((accepted, fingerprint))
    }

    /// Authenticate two bounded artifact streams before publishing the
    /// retained original through the recipient sink.
    ///
    /// The current and original streams are each probed through a fixed
    /// max-plus-one digest sink before the caller's writer is touched. During
    /// publication, producer source/work checks remain attached to each
    /// producer while `ArtifactStreamPolicy` charges recipient output once.
    /// Passing an outer policy is required for nested deferred streams so the
    /// recipient allowance and probe/publication phase are shared.
    pub fn restore_streamed_artifact_to_stream<C, O, W>(
        &self,
        current: &C,
        original: &O,
        proof: SourceArtifactRestoreProof,
        max_output_bytes: u64,
        writer: W,
        outer_policy: Option<&ArtifactStreamPolicy<'_>>,
    ) -> std::result::Result<(), ArtifactStreamError<O::Error>>
    where
        C: ArtifactStream<Error = O::Error> + ?Sized,
        O: ArtifactStream + ?Sized,
        W: Write,
    {
        self.disable_read_ahead_for_publication()
            .map_err(ArtifactStreamError::opc)?;
        let policy_owned;
        let policy = if let Some(policy) = outer_policy {
            policy
        } else {
            policy_owned =
                ArtifactStreamPolicy::for_recipient(self).map_err(ArtifactStreamError::opc)?;
            &policy_owned
        };
        if max_output_bytes == 0 || max_output_bytes == u64::MAX {
            return Err(ArtifactStreamError::opc(
                OpcError::InvalidSourcePartSpliceLimit {
                    resource: SpliceResource::OutputBytes,
                    value: max_output_bytes,
                },
            ));
        }
        if proof.original_len > max_output_bytes {
            return Err(ArtifactStreamError::opc(OpcError::SourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                actual: proof.original_len,
                maximum: max_output_bytes,
            }));
        }
        let probe = policy.probe();
        let current_preflight = preflight_stream(
            current,
            proof.current_len,
            proof.current_sha256,
            "current",
            &probe,
        );
        let original_preflight = preflight_stream(
            original,
            proof.original_len,
            proof.original_sha256,
            "original",
            &probe,
        );
        match (current_preflight, original_preflight) {
            (Ok(()), Ok(())) => {},
            (Err(current), Ok(())) => return Err(current),
            (Ok(()), Err(original)) => return Err(original),
            (Err(current), Err(original)) => {
                return Err(ArtifactStreamError::incomplete(0, current, Some(original)));
            },
        }

        let mut written = 0_u64;
        let mut counted = Counted::new(writer, &mut written);
        let mut target = ArtifactStreamTarget::new(&mut counted, policy);
        let result = match policy.check_recipient().map_err(ArtifactStreamError::opc) {
            Err(error) => Err(error),
            Ok(()) => original
                .stream_to(&mut target)
                .map_err(ArtifactStreamError::producer)
                .and_then(|publication| {
                    target
                        .validate_publication(publication)
                        .map_err(ArtifactStreamError::opc)
                        .and_then(|()| {
                            if publication.written() != proof.original_len {
                                return Err(ArtifactStreamError::opc(
                                    OpcError::SourceArtifactMismatch {
                                        artifact: "original",
                                        field: "length",
                                    },
                                ));
                            }
                            if publication.fingerprint() != proof.original_sha256 {
                                return Err(ArtifactStreamError::opc(
                                    OpcError::SourceArtifactMismatch {
                                        artifact: "original",
                                        field: "fingerprint",
                                    },
                                ));
                            }
                            Ok(())
                        })
                })
                .and_then(|()| {
                    target
                        .flush()
                        .map_err(map_io_error)
                        .map_err(ArtifactStreamError::opc)
                }),
        };
        drop(target);
        let policy_result = policy
            .take_failure()
            .map(|error| ArtifactStreamError::opc(map_execution_error(error)));
        let result = match (result, policy_result) {
            (Ok(()), None) => Ok(()),
            (Err(error), None) | (Ok(()), Some(error)) => Err(error),
            (Err(error), Some(policy_error)) => Err(ArtifactStreamError::incomplete(
                written,
                error,
                Some(policy_error),
            )),
        };
        finish_streamed_publication(result, current, original, written)
    }

    /// Restore an explicit original artifact after authenticating both complete
    /// artifacts, without materializing either artifact.
    ///
    /// The current package's execution context controls the operation. The
    /// retained provider is still checked for freshness and cancellation, but
    /// its resource policy is never inherited by the recipient. Both
    /// providers must remain fresh and uncancelled. A single 64 KiB workspace
    /// is reserved when the current context has one; output is charged once by
    /// that context. The original bytes are authenticated again as they are
    /// emitted. Any failure after accepted output is reported with its exact
    /// accepted byte count.
    pub fn restore_source_artifact_to_stream<W: Write>(
        &self,
        original: &SourceArtifact,
        proof: SourceArtifactRestoreProof,
        max_output_bytes: u64,
        writer: W,
    ) -> Result<()> {
        self.disable_read_ahead_for_publication()?;
        let retained = &original.snapshot;
        let context = self.cache.context().cloned();
        check_inverse_state(self, retained, context.as_ref())?;
        if max_output_bytes == 0 || max_output_bytes == u64::MAX {
            return Err(OpcError::InvalidSourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                value: max_output_bytes,
            });
        }
        if retained.length > max_output_bytes {
            return Err(OpcError::SourcePartSpliceLimit {
                resource: SpliceResource::OutputBytes,
                actual: retained.length,
                maximum: max_output_bytes,
            });
        }
        if self.source.length != proof.current_len {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "length",
            });
        }
        if retained.length != proof.original_len {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "length",
            });
        }
        let output_failures = if self.cache.context().is_some() {
            self.source.output_reservation_failures.clone()
        } else {
            retained.output_reservation_failures.clone()
        };
        if context.is_some() && output_failures.is_none() {
            return Err(overlay_unavailable(
                "managed restore output counter is unavailable",
            ));
        }
        let _monitored_current = self.source.monitor_publication();
        let _monitored_retained = retained.monitor_publication();
        let _workspace = if let Some(context) = context.as_ref() {
            let requested = SOURCE_PUBLICATION_CHUNK_BYTES;
            match context.reserve(Resource::Memory, requested as u64) {
                Ok(reservation) => Some(reservation),
                Err(ExecutionError::ResourceLimit(limit)) if limit.resource == Resource::Memory => {
                    let used = limit.observed.saturating_sub(requested as u64);
                    let available = limit.limit.saturating_sub(used);
                    if available == 0 {
                        // The package may already retain its complete
                        // catalog under an exact memory allowance. A
                        // bounded stack window keeps authentication and
                        // streaming available without an uncharged heap
                        // allocation or a false policy refusal.
                        None
                    } else {
                        return Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)));
                    }
                },
                Err(error) => return Err(map_execution_error(error)),
            }
        } else {
            None
        };
        const RESTORE_STACK_BYTES: usize = 4 * 1024;
        let mut stack_buffer = [0_u8; RESTORE_STACK_BYTES];
        // Keep the managed 64 KiB reservation and its refusal boundary above,
        // but use the bounded stack window for the actual reads. This avoids
        // one heap buffer per exact inverse while retaining recipient policy.
        let buffer: &mut [u8] = &mut stack_buffer;
        let current_hash =
            fingerprint_inverse_snapshot(self, &self.source, retained, context.as_ref(), buffer)?;
        if current_hash != proof.current_sha256 {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "current",
                field: "fingerprint",
            });
        }
        let original_hash =
            fingerprint_inverse_snapshot(self, retained, retained, context.as_ref(), buffer)?;
        if original_hash != proof.original_sha256 {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "fingerprint",
            });
        }
        let failure = Arc::new(Mutex::new(None));
        let mut written = 0;
        let counted = Counted::new(writer, &mut written);
        let checked_current = SourceCheckedSink {
            inner: counted,
            snapshot: self.source.clone(),
        };
        let checked_original = SourceCheckedSink {
            inner: checked_current,
            snapshot: retained.clone(),
        };
        let original_cooperative = ContextCheckedSink {
            inner: checked_original,
            context: retained.context.clone(),
            failure: Arc::clone(&failure),
        };
        let cooperative = ContextCheckedSink {
            inner: original_cooperative,
            context: context.clone(),
            failure: Arc::clone(&failure),
        };
        let result = if let (Some(operation_context), Some(output_reservation_failures)) =
            (context.clone(), output_failures)
        {
            let mut output = OutputBudgetedSink {
                inner: cooperative,
                context: operation_context,
                failure: Arc::clone(&failure),
                output_reservation_failures,
            };
            copy_authenticated(
                self,
                retained,
                context.as_ref(),
                proof.original_sha256,
                buffer,
                &mut output,
            )
            .and_then(|()| output.flush().map_err(map_io_error))
        } else {
            let mut output = cooperative;
            copy_authenticated(
                self,
                retained,
                context.as_ref(),
                proof.original_sha256,
                buffer,
                &mut output,
            )
            .and_then(|()| output.flush().map_err(map_io_error))
        };
        let result = failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = match (
            result,
            check_inverse_state(self, retained, context.as_ref()),
        ) {
            (_, Err(error @ OpcError::SourceChanged { .. })) => Err(error),
            (Err(error), _) | (Ok(()), Err(error)) => Err(error),
            (Ok(()), Ok(())) => Ok(()),
        };
        finish_inverse_publication(result, &self.source, retained, written)
    }
}

fn copy_authenticated(
    current: &SourceBackedPackage,
    retained: &SourceSnapshot,
    context: Option<&ExecutionContext>,
    expected: SourceArtifactFingerprint,
    buffer: &mut [u8],
    output: &mut dyn Write,
) -> Result<()> {
    let mut offset = 0_u64;
    let mut hash = Sha256::new();
    while offset < retained.length {
        check_inverse_state(current, retained, context)?;
        let count = usize::try_from((retained.length - offset).min(buffer.len() as u64))
            .map_err(|_| overlay_unavailable("artifact restore range exceeds usize"))?;
        let read = read_source_at_with_context(
            retained,
            context,
            offset,
            &mut buffer[..count],
            "authenticated restore",
        )
        .map_err(super::map_source_backed_error)?;
        if read == 0 {
            return Err(OpcError::IoError(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "original artifact ended during restore",
            )));
        }
        check_inverse_state(current, retained, context)?;
        if let Some(context) = context {
            context
                .consume(Resource::Work, read as u64)
                .map_err(map_execution_error)?;
        }
        hash.update(&buffer[..read]);
        output.write_all(&buffer[..read]).map_err(map_io_error)?;
        check_inverse_state(current, retained, context)?;
        offset = offset
            .checked_add(read as u64)
            .ok_or_else(|| overlay_unavailable("artifact restore offset overflow"))?;
    }
    if SourceArtifactFingerprint::from_sha256(hash.finalize().into()) != expected {
        return Err(OpcError::SourceArtifactMismatch {
            artifact: "original",
            field: "replayed fingerprint",
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused policy tests deliberately panic on fixture errors"
    )]

    use super::super::{PolicyAppendedEntry, SourceCacheLimits};
    use super::*;
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::sync::atomic::{AtomicUsize, Ordering as AtomicOrdering};

    fn fingerprint(bytes: &[u8]) -> SourceArtifactFingerprint {
        SourceArtifactFingerprint::from_sha256(Sha256::digest(bytes).into())
    }

    fn append_fixture(with_orphan_relationship: bool) -> SourceBackedPackage {
        const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        const DOCUMENT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../custom/orphan.xml"/></Relationships>"#;
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("[Content_Types].xml", CONTENT_TYPES)
            .unwrap();
        writer
            .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<document/>")
            .unwrap();
        if with_orphan_relationship {
            writer
                .write_stored("word/_rels/document.xml.rels", DOCUMENT_RELATIONSHIPS)
                .unwrap();
            writer
                .write_stored("custom/orphan.xml", b"<orphan/>")
                .unwrap();
        }
        SourceBackedPackage::from_vec(
            writer.finish_to_bytes().unwrap(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .unwrap()
    }

    fn physical_orphan_relationship_fixture() -> SourceBackedPackage {
        const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        const DOCUMENT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../custom/orphan.xml"/></Relationships>"#;
        const ORPHAN_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../custom/orphan.xml"/></Relationships>"#;
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("[Content_Types].xml", CONTENT_TYPES)
            .unwrap();
        writer
            .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<document/>")
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", DOCUMENT_RELATIONSHIPS)
            .unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        writer
            .write_stored("custom/_rels/missing.xml.rels", ORPHAN_RELATIONSHIPS)
            .unwrap();
        SourceBackedPackage::from_vec(
            writer.finish_to_bytes().unwrap(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .unwrap()
    }

    fn test_policy(limits: SourceLimits) -> ArtifactStreamPolicy<'static> {
        ArtifactStreamPolicy {
            phase: ArtifactStreamPhase::Publish,
            recipient_limits: limits,
            recipient_context: None,
            recipient_output_reservation_failures: None,
            failure: Arc::new(Mutex::new(None)),
        }
    }

    fn appended_member(name: &str, payload: &[u8]) -> PolicyAppendedEntry {
        PolicyAppendedEntry::new_shared(
            name,
            Arc::new(payload.to_vec()),
            soapberry_zip::CompressionMethod::Store,
        )
    }

    fn swapping_artifact(document: &[u8]) -> SourceArtifact {
        const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("[Content_Types].xml", CONTENT_TYPES)
            .unwrap();
        writer
            .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        SourceBackedPackage::from_vec(
            writer.finish_to_bytes().unwrap(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .unwrap()
        .source_artifact()
    }

    struct SwappingArtifact {
        first: SourceArtifact,
        second: SourceArtifact,
        calls: AtomicUsize,
    }

    impl ArtifactStream for SwappingArtifact {
        type Error = OpcError;

        fn check_fresh(&self) -> Result<()> {
            Ok(())
        }

        fn stream_to(
            &self,
            target: &mut ArtifactStreamTarget<'_>,
        ) -> std::result::Result<ArtifactStreamReport, Self::Error> {
            if self.calls.fetch_add(1, AtomicOrdering::SeqCst) == 0 {
                target.publish_source_artifact(&self.first)
            } else {
                target.publish_source_artifact(&self.second)
            }
        }
    }

    struct RefusingGuard {
        checks: AtomicUsize,
    }

    impl ArtifactStreamGuard for RefusingGuard {
        type Error = ();

        fn check(&self) -> std::result::Result<(), Self::Error> {
            self.checks.fetch_add(1, AtomicOrdering::SeqCst);
            Err(())
        }
    }

    #[test]
    fn guard_refusal_is_terminal_for_write_all() {
        let guard = RefusingGuard {
            checks: AtomicUsize::new(0),
        };
        let mut output = Vec::new();
        let mut sink = GuardedArtifactSink {
            inner: &mut output,
            guard: &guard,
            failure: None,
        };
        let error = sink
            .write_all(b"guarded output")
            .expect_err("a refused guard must stop write_all");
        assert_eq!(error.kind(), io::ErrorKind::Other);
        assert_eq!(guard.checks.load(AtomicOrdering::SeqCst), 1);
        assert!(output.is_empty());
    }

    #[test]
    fn preflight_caps_an_overlong_stream_before_publication() {
        let stream = append_fixture(false).source_artifact();
        let expected_fingerprint = stream.fingerprint().unwrap();
        let policy = ArtifactStreamPolicy {
            phase: ArtifactStreamPhase::Probe,
            recipient_limits: SourceLimits::default(),
            recipient_context: None,
            recipient_output_reservation_failures: None,
            failure: Arc::new(Mutex::new(None)),
        };
        let error = preflight_stream(
            &stream,
            stream.len() - 1,
            expected_fingerprint,
            "original",
            &policy,
        )
        .expect_err("the bounded probe must reject a fourth byte");
        assert!(matches!(
            error,
            ArtifactStreamError::Opc(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "length"
            })
        ));
    }

    #[test]
    fn preflight_authenticates_exact_length_and_digest() {
        let stream = append_fixture(false).source_artifact();
        let expected_fingerprint = stream.fingerprint().unwrap();
        let policy = ArtifactStreamPolicy {
            phase: ArtifactStreamPhase::Probe,
            recipient_limits: SourceLimits::default(),
            recipient_context: None,
            recipient_output_reservation_failures: None,
            failure: Arc::new(Mutex::new(None)),
        };
        preflight_stream(
            &stream,
            stream.len(),
            expected_fingerprint,
            "original",
            &policy,
        )
        .expect("matching length and digest should pass");
        let error = preflight_stream(
            &stream,
            stream.len(),
            fingerprint(b"abd"),
            "original",
            &policy,
        )
        .expect_err("a changed digest must be rejected");
        assert!(matches!(
            error,
            ArtifactStreamError::Opc(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "fingerprint"
            })
        ));
    }

    #[test]
    fn restore_rejects_a_valid_but_swapped_final_original() {
        let first = swapping_artifact(b"<a/>");
        let second = swapping_artifact(b"<b/>");
        assert_eq!(first.len(), second.len());
        let fingerprint = first.fingerprint().unwrap();
        let proof = SourceArtifactRestoreProof {
            current_len: first.len(),
            current_sha256: fingerprint,
            original_len: first.len(),
            original_sha256: fingerprint,
        };
        let swapping = SwappingArtifact {
            first: first.clone(),
            second,
            calls: AtomicUsize::new(0),
        };
        let recipient = append_fixture(false);
        let mut output = Vec::new();
        let error = recipient
            .restore_streamed_artifact_to_stream(
                &first,
                &swapping,
                proof,
                first.len(),
                &mut output,
                None,
            )
            .expect_err("the final valid-but-different artifact must be rejected");
        assert_eq!(swapping.calls.load(AtomicOrdering::SeqCst), 2);
        assert!(matches!(
            error.opc_error(),
            Some(OpcError::SourceArtifactMismatch {
                artifact: "original",
                field: "fingerprint"
            })
        ));
        assert!(!output.is_empty());
    }

    #[test]
    fn appended_ordinary_part_checks_exact_and_one_under_part_limits() {
        let package = append_fixture(false);
        let appended = appended_member("custom/new.bin", b"new-part");
        let source_part_bytes = package
            .catalog
            .parts
            .iter()
            .map(|part| {
                package
                    .archive
                    .metadata_for(part.entry_id)
                    .unwrap()
                    .uncompressed_size()
            })
            .sum::<u64>();
        let exact = SourceLimits::builder()
            .max_parts(package.catalog.parts.len() + 1)
            .unwrap()
            .max_total_part_bytes(source_part_bytes + 8)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(exact)
                .admit_changed_overlays(&package, &[], &[], &[appended])
                .is_ok()
        );

        let one_under = SourceLimits::builder()
            .max_parts(package.catalog.parts.len())
            .unwrap()
            .max_total_part_bytes(source_part_bytes + 7)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(one_under)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &[],
                    &[appended_member("custom/new.bin", b"new-part")],
                )
                .is_err()
        );
    }

    #[test]
    fn appended_content_types_checks_exact_and_one_under_bytes() {
        const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let package = append_fixture(false);
        let exact = SourceLimits::builder()
            .max_content_types_bytes(CONTENT_TYPES.len())
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(exact)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["[Content_Types].xml".to_owned()],
                    &[appended_member("[Content_Types].xml", CONTENT_TYPES)],
                )
                .is_ok()
        );

        let one_under = SourceLimits::builder()
            .max_content_types_bytes(CONTENT_TYPES.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(one_under)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["[Content_Types].xml".to_owned()],
                    &[appended_member("[Content_Types].xml", CONTENT_TYPES)],
                )
                .is_err()
        );
    }

    #[test]
    fn appended_relationship_checks_exact_and_one_under_xml_bytes() {
        const RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let package = append_fixture(false);
        let exact = SourceLimits::builder()
            .max_total_relationship_xml_bytes(RELATIONSHIPS.len())
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(exact)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["_rels/.rels".to_owned()],
                    &[appended_member("_rels/.rels", RELATIONSHIPS)],
                )
                .is_ok()
        );

        let one_under = SourceLimits::builder()
            .max_total_relationship_xml_bytes(RELATIONSHIPS.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(one_under)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["_rels/.rels".to_owned()],
                    &[appended_member("_rels/.rels", RELATIONSHIPS)],
                )
                .is_err()
        );
    }

    #[test]
    fn appended_relationship_graph_checks_exact_and_one_under_nodes() {
        const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let package = append_fixture(true);
        let exact = SourceLimits::builder()
            .max_relationship_graph_nodes(2)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(exact)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["_rels/.rels".to_owned()],
                    &[appended_member("_rels/.rels", ROOT_RELATIONSHIPS)],
                )
                .is_ok()
        );

        let one_under = SourceLimits::builder()
            .max_relationship_graph_nodes(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            test_policy(one_under)
                .admit_changed_overlays(
                    &package,
                    &[],
                    &["_rels/.rels".to_owned()],
                    &[appended_member("_rels/.rels", ROOT_RELATIONSHIPS)],
                )
                .is_err()
        );
    }

    #[test]
    fn prepared_topology_counts_physical_orphan_relationships() {
        let package = physical_orphan_relationship_fixture();
        let mut plan = super::super::SourceTopologyPlan::new();
        plan.try_replace_part(
            PackURI::new("/word/document.xml").unwrap(),
            b"<changed/>".to_vec(),
        )
        .unwrap();
        let prepared = package.prepare_topology(plan).unwrap();
        let limits = SourceLimits::builder()
            .max_total_relationships(2)
            .unwrap()
            .build()
            .unwrap();
        let policy = test_policy(limits);
        let mut output = Vec::new();
        let error = prepared
            .publish_to_stream_with_policy(&mut output, &policy)
            .expect_err("the orphan relationship must count toward the final total");
        assert!(matches!(
            error,
            OpcError::Limit {
                kind: SourceLimitKind::TotalRelationships,
                ..
            }
        ));
        assert!(output.is_empty());
    }
}
