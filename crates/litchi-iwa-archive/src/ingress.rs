//! Allocation-free candidate ingress facts for one immutable iWork ZIP.
//!
//! This seam stops at physical storage. It validates the ZIP envelope and
//! reports the storage required to materialize one selected member, one
//! Snappy stream, or one decoded IWA archive. It deliberately does not select
//! Pages, Numbers, or Keynote messages and does not replace semantic
//! preflight.

use flate2::read::DeflateDecoder;
use litchi_iwa_core::{
    Archive, ArchiveCloneInventory, ArchiveCloneInventoryChunk, ArchiveCloneInventoryError,
    ArchiveCloneInventoryVisitor, ArchiveIngressRequirements, SnappyStream,
};
use soapberry_zip::{CompressionMethod, ZipArchive as RawZipArchive, ZipFileHeaderRecord};
use std::io::{self, Cursor, Read};

use crate::package::{
    EntryEdit, EntryInsertion, PreparedReassembly, ReassemblyPreparationRequirements,
};
use crate::{Error, LimitKind, Limits, Result, SourceCatalog};

const ZIP_CENTRAL_FIXED_METADATA_BYTES: usize = 46;

/// The physical phase whose storage is being admitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CandidateIngressPhase {
    /// Central-directory metadata and physical entry traversal.
    ZipDirectory,
    /// One decoded ZIP member retained by a package candidate.
    ZipEntryStorage,
    /// One checksum-free Apple IWA Snappy stream.
    SnappyStream,
    /// One decoded IWA archive and its header closure.
    IwaArchive,
    /// One deep clone of an already parsed IWA archive.
    IwaClone,
    /// ZIP reader, mutation indexes, and insertion indexes used to prepare a
    /// physical reassembly plan.
    ReassemblyPreparation,
}

/// A checked physical charge supplied to a caller-owned transaction ledger.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CandidateIngressCharge {
    phase: CandidateIngressPhase,
    retained_bytes: usize,
    traversal_bytes: usize,
    allocations: usize,
}

impl CandidateIngressCharge {
    /// Physical phase represented by this charge.
    #[must_use]
    pub const fn phase(self) -> CandidateIngressPhase {
        self.phase
    }

    /// Bytes retained or reserved by the phase.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Bytes traversed by the phase.
    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }

    /// Fallible logical allocation events represented by the phase.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    const fn new(
        phase: CandidateIngressPhase,
        retained_bytes: usize,
        traversal_bytes: usize,
        allocations: usize,
    ) -> Self {
        Self {
            phase,
            retained_bytes,
            traversal_bytes,
            allocations,
        }
    }
}

/// Visitor used to charge each physical phase before it allocates or decodes.
pub trait CandidateIngressVisitor {
    /// Error returned by the caller's aggregate transaction ledger.
    type Error;

    /// Admit one checked physical charge before the corresponding work.
    fn admit(&mut self, charge: CandidateIngressCharge) -> std::result::Result<(), Self::Error>;
}

impl<F, E> CandidateIngressVisitor for F
where
    F: FnMut(CandidateIngressCharge) -> std::result::Result<(), E>,
{
    type Error = E;

    fn admit(&mut self, charge: CandidateIngressCharge) -> std::result::Result<(), E> {
        self(charge)
    }
}

/// Typed failure from a candidate ingress visitor or physical preflight.
#[derive(Debug)]
#[non_exhaustive]
pub enum CandidateIngressError<E> {
    /// ZIP, Snappy, or IWA physical validation failed.
    Physical(Error),
    /// The caller's aggregate ledger rejected the next phase.
    Admission(E),
    /// The allocation-free clone inventory counter overflowed.
    CloneArithmetic { resource: &'static str },
}

impl<E> From<Error> for CandidateIngressError<E> {
    fn from(error: Error) -> Self {
        Self::Physical(error)
    }
}

/// Aggregate physical facts found in one ZIP central directory.
///
/// The source ZIP bytes are borrowed and therefore excluded. `materialized`
/// counts the capacity needed for each decoded file result: Store members use
/// their exact declared length and Deflate members include the one-byte
/// overrun sentinel used by the bounded reader. Catalog map/vector storage is
/// a separate shape-dependent charge and is exposed by the post-construction
/// [`crate::package::Catalog::storage_requirements`] census.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CandidateIngressRequirements {
    input_bytes: usize,
    physical_entries: usize,
    file_entries: usize,
    iwa_entries: usize,
    metadata_bytes: usize,
    compressed_bytes: usize,
    declared_entry_bytes: usize,
    materialized_entry_bytes: usize,
    retained_bytes: usize,
    traversal_bytes: usize,
    allocations: usize,
}

impl CandidateIngressRequirements {
    /// Complete borrowed source length checked by the ZIP locator.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Number of central-directory records, including directories.
    #[must_use]
    pub const fn physical_entries(self) -> usize {
        self.physical_entries
    }

    /// Number of non-directory records.
    #[must_use]
    pub const fn file_entries(self) -> usize {
        self.file_entries
    }

    /// Number of non-directory records whose physical name ends in `.iwa`.
    #[must_use]
    pub const fn iwa_entries(self) -> usize {
        self.iwa_entries
    }

    /// Central fixed and variable metadata bytes traversed by the plan.
    #[must_use]
    pub const fn metadata_bytes(self) -> usize {
        self.metadata_bytes
    }

    /// Declared compressed file bytes.
    #[must_use]
    pub const fn compressed_bytes(self) -> usize {
        self.compressed_bytes
    }

    /// Declared uncompressed file bytes.
    #[must_use]
    pub const fn declared_entry_bytes(self) -> usize {
        self.declared_entry_bytes
    }

    /// Capacity needed to materialize all file members through bounded ZIP
    /// decoding.
    #[must_use]
    pub const fn materialized_entry_bytes(self) -> usize {
        self.materialized_entry_bytes
    }

    /// Retained member storage plus central metadata facts.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Central metadata and compressed member bytes traversed by the plan.
    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }

    /// Logical output allocation events represented by the file members.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// ZIP compression class without exposing the dependency's raw method IDs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CandidateCompression {
    /// ZIP Store, whose member bytes can be copied directly.
    Stored,
    /// ZIP Deflate, whose member bytes require bounded inflation.
    Deflated,
    /// Any method outside the physical iWork reader's supported set.
    Unsupported,
}

/// Neutral classification of one physical ZIP record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CandidateEntryKind {
    /// A directory record with no package payload allocation.
    Directory,
    /// The legacy nested package index marker.
    Index,
    /// An IWA component member.
    Iwa,
    /// Any other physical file member.
    Other,
}

/// A borrowed, allocation-free descriptor for one ZIP record.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CandidateEntry {
    ordinal: usize,
    wayfinder: soapberry_zip::ZipArchiveEntryWayfinder,
    kind: CandidateEntryKind,
    compression: CandidateCompression,
    metadata_bytes: usize,
    compressed_bytes: usize,
    declared_bytes: usize,
    materialized_bytes: usize,
}

impl std::fmt::Debug for CandidateEntry {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CandidateEntry")
            .field("ordinal", &self.ordinal)
            .field("kind", &self.kind)
            .field("compression", &self.compression)
            .field("metadata_bytes", &self.metadata_bytes)
            .field("compressed_bytes", &self.compressed_bytes)
            .field("declared_bytes", &self.declared_bytes)
            .field("materialized_bytes", &self.materialized_bytes)
            .finish()
    }
}

impl CandidateEntry {
    /// Stable ordinal within this candidate's central-directory order.
    #[must_use]
    pub const fn ordinal(self) -> usize {
        self.ordinal
    }

    /// Neutral physical kind used by package owners to select a target.
    #[must_use]
    pub const fn kind(self) -> CandidateEntryKind {
        self.kind
    }

    /// Compression class without exposing raw ZIP method IDs.
    #[must_use]
    pub const fn compression(self) -> CandidateCompression {
        self.compression
    }

    /// Whether this record is a directory.
    #[must_use]
    pub const fn is_directory(self) -> bool {
        matches!(self.kind, CandidateEntryKind::Directory)
    }

    /// Whether this record is an IWA component target.
    #[must_use]
    pub const fn is_iwa(self) -> bool {
        matches!(self.kind, CandidateEntryKind::Iwa)
    }

    /// Central fixed and variable metadata bytes for this record.
    #[must_use]
    pub const fn metadata_bytes(self) -> usize {
        self.metadata_bytes
    }

    /// Declared compressed member bytes.
    #[must_use]
    pub const fn compressed_bytes(self) -> usize {
        self.compressed_bytes
    }

    /// Declared uncompressed member bytes.
    #[must_use]
    pub const fn declared_bytes(self) -> usize {
        self.declared_bytes
    }

    /// Capacity required by bounded ZIP materialization.
    #[must_use]
    pub const fn materialized_bytes(self) -> usize {
        self.materialized_bytes
    }
}

/// One ZIP target whose output capacity has been admitted.
pub struct AdmittedCandidateEntry<'plan, 'data> {
    plan: &'plan CandidateIngressPlan<'data>,
    entry: CandidateEntry,
}

impl std::fmt::Debug for AdmittedCandidateEntry<'_, '_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("AdmittedCandidateEntry")
            .field("entry", &self.entry)
            .finish_non_exhaustive()
    }
}

impl<'plan, 'data> AdmittedCandidateEntry<'plan, 'data> {
    /// Borrow the neutral descriptor after its output charge was admitted.
    #[must_use]
    pub const fn entry(&self) -> CandidateEntry {
        self.entry
    }

    /// Materialize the selected member after the caller's ledger admission.
    ///
    /// The returned bytes are ZIP-decoded payload bytes. Snappy and IWA
    /// admission are separate explicit phases and must be charged through
    /// [`CandidateIngressPlan::admit_snappy_stream`] and
    /// [`CandidateIngressPlan::admit_archive_stream`].
    pub fn read(self) -> Result<Vec<u8>> {
        self.plan.read_entry(self.entry)
    }
}

/// Exact decoded-length facts for one Snappy stream.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SnappyIngressRequirements {
    compressed_bytes: usize,
    decoded_bytes: usize,
}

impl SnappyIngressRequirements {
    /// Compressed Snappy bytes traversed.
    #[must_use]
    pub const fn compressed_bytes(self) -> usize {
        self.compressed_bytes
    }

    /// Exact decoded bytes that must be retained.
    #[must_use]
    pub const fn decoded_bytes(self) -> usize {
        self.decoded_bytes
    }
}

/// Allocation-free physical candidate plan for one ZIP source.
#[derive(Clone)]
pub struct CandidateIngressPlan<'data> {
    archive: soapberry_zip::ZipSliceArchive<&'data [u8]>,
    limits: Limits,
    requirements: CandidateIngressRequirements,
}

impl std::fmt::Debug for CandidateIngressPlan<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CandidateIngressPlan")
            .field("limits", &self.limits)
            .field("requirements", &self.requirements)
            .finish_non_exhaustive()
    }
}

impl<'data> CandidateIngressPlan<'data> {
    /// Validate the ZIP envelope and produce physical facts without retaining
    /// a central-directory map, member names, or decoded payloads.
    pub fn from_bytes_with_limits(data: &'data [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        limits.check_input_size(
            u64::try_from(data.len()).map_err(|_error| {
                Error::InvalidBundle("ZIP input length does not fit u64".to_owned())
            })?,
            "ZIP input",
        )?;
        let archive = RawZipArchive::from_slice(data)?;
        let mut requirements = CandidateIngressRequirements {
            input_bytes: data.len(),
            ..CandidateIngressRequirements::default()
        };
        for result in archive.entries() {
            let entry = result?;
            requirements.physical_entries = requirements
                .physical_entries
                .checked_add(1)
                .ok_or_else(|| invalid("ZIP physical entry count overflows usize"))?;
            if requirements.physical_entries > limits.max_entries() {
                return Err(limit(
                    LimitKind::Entries,
                    requirements.physical_entries,
                    limits.max_entries() as u64,
                ));
            }
            let metadata_bytes = ZIP_CENTRAL_FIXED_METADATA_BYTES
                .checked_add(usize_from_u64(
                    entry.metadata_size_hint(),
                    "ZIP metadata length",
                )?)
                .ok_or_else(|| invalid("ZIP metadata length overflows usize"))?;
            requirements.metadata_bytes =
                requirements
                    .metadata_bytes
                    .checked_add(metadata_bytes)
                    .ok_or_else(|| invalid("ZIP metadata total overflows usize"))?;
            if requirements.metadata_bytes
                > usize_from_u64(limits.max_metadata_bytes(), "ZIP metadata limit")?
            {
                return Err(limit(
                    LimitKind::MetadataBytes,
                    requirements.metadata_bytes,
                    limits.max_metadata_bytes(),
                ));
            }
            let raw_path = entry.file_path();
            let raw_name = raw_path.as_ref();
            if raw_name.len()
                > usize_from_u64(
                    limits.zip_limits().max_member_name_bytes,
                    "ZIP member name limit",
                )?
            {
                return Err(limit(
                    LimitKind::MemberNameBytes,
                    raw_name.len(),
                    limits.zip_limits().max_member_name_bytes,
                ));
            }
            if entry.is_encrypted() {
                return Err(Error::Encrypted);
            }
            // Validate local framing and the complete compressed range while
            // the source is still borrowed. No member output is allocated.
            let _ = archive.get_entry(entry.wayfinder())?;

            let kind = classify_entry(raw_name, entry.is_dir());
            if matches!(kind, CandidateEntryKind::Directory) {
                continue;
            }
            requirements.file_entries = requirements
                .file_entries
                .checked_add(1)
                .ok_or_else(|| invalid("ZIP file entry count overflows usize"))?;
            if matches!(kind, CandidateEntryKind::Iwa) {
                requirements.iwa_entries = requirements
                    .iwa_entries
                    .checked_add(1)
                    .ok_or_else(|| invalid("IWA entry count overflows usize"))?;
            }
            let compressed_bytes =
                usize_from_u64(entry.compressed_size_hint(), "ZIP compressed member length")?;
            if entry.compressed_size_hint() > limits.zip_limits().max_compressed_size {
                return Err(limit(
                    LimitKind::CompressedEntryBytes,
                    compressed_bytes,
                    limits.zip_limits().max_compressed_size,
                ));
            }
            requirements.compressed_bytes = requirements
                .compressed_bytes
                .checked_add(compressed_bytes)
                .ok_or_else(|| invalid("ZIP compressed total overflows usize"))?;
            let declared_bytes =
                usize_from_u64(entry.uncompressed_size_hint(), "ZIP entry length")?;
            if entry.uncompressed_size_hint() > limits.max_entry_bytes() {
                return Err(limit(
                    LimitKind::EntryBytes,
                    declared_bytes,
                    limits.max_entry_bytes(),
                ));
            }
            requirements.declared_entry_bytes = requirements
                .declared_entry_bytes
                .checked_add(declared_bytes)
                .ok_or_else(|| invalid("ZIP entry total overflows usize"))?;
            let total = u64::try_from(requirements.declared_entry_bytes)
                .map_err(|_error| invalid("ZIP entry total does not fit u64"))?;
            if total > limits.max_total_bytes() {
                return Err(limit(
                    LimitKind::TotalBytes,
                    requirements.declared_entry_bytes,
                    limits.max_total_bytes(),
                ));
            }
            let materialized_bytes = match entry.compression_method() {
                CompressionMethod::Store => declared_bytes,
                CompressionMethod::Deflate => declared_bytes
                    .checked_add(1)
                    .ok_or_else(|| invalid("ZIP Deflate output capacity overflows usize"))?,
                _ => 0,
            };
            requirements.materialized_entry_bytes = requirements
                .materialized_entry_bytes
                .checked_add(materialized_bytes)
                .ok_or_else(|| invalid("ZIP materialized total overflows usize"))?;
        }
        requirements.retained_bytes = requirements
            .metadata_bytes
            .checked_add(requirements.materialized_entry_bytes)
            .ok_or_else(|| invalid("candidate retained bytes overflow"))?;
        requirements.traversal_bytes = requirements
            .metadata_bytes
            .checked_add(requirements.compressed_bytes)
            .ok_or_else(|| invalid("candidate traversal bytes overflow"))?;
        requirements.allocations = requirements.file_entries;
        Ok(Self {
            archive,
            limits,
            requirements,
        })
    }

    /// Validate one source with the default physical limits.
    pub fn from_bytes(data: &'data [u8]) -> Result<Self> {
        Self::from_bytes_with_limits(data, Limits::default())
    }

    /// Borrow the aggregate ZIP facts.
    #[must_use]
    pub const fn requirements(&self) -> CandidateIngressRequirements {
        self.requirements
    }

    /// Return descriptors in central-directory order without allocating names.
    pub fn entries(&self) -> impl Iterator<Item = Result<CandidateEntry>> + '_ {
        self.archive
            .entries()
            .enumerate()
            .map(|(ordinal, result)| result.map(|entry| self.entry_from_record(ordinal, entry)))
            .map(|result| result.map_err(Error::from))
    }

    /// Admit the central-directory phase before constructing a package reader.
    pub fn admit_zip_directory<V>(
        &self,
        visitor: &mut V,
    ) -> std::result::Result<(), CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        visitor
            .admit(CandidateIngressCharge::new(
                CandidateIngressPhase::ZipDirectory,
                self.requirements.metadata_bytes,
                self.requirements.metadata_bytes,
                0,
            ))
            .map_err(CandidateIngressError::Admission)
    }

    /// Admit one member's ZIP-decoded output before materializing it.
    pub fn admit_entry<V>(
        &self,
        entry: CandidateEntry,
        visitor: &mut V,
    ) -> std::result::Result<AdmittedCandidateEntry<'_, 'data>, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        if entry.is_directory() {
            return Err(CandidateIngressError::Physical(invalid(
                "cannot admit a directory as a ZIP payload",
            )));
        }
        visitor
            .admit(CandidateIngressCharge::new(
                CandidateIngressPhase::ZipEntryStorage,
                entry.materialized_bytes,
                entry.compressed_bytes,
                1,
            ))
            .map_err(CandidateIngressError::Admission)?;
        Ok(AdmittedCandidateEntry { plan: self, entry })
    }

    /// Preflight a ZIP-decoded Snappy stream without allocating its output.
    pub fn preflight_snappy_stream(&self, compressed: &[u8]) -> Result<SnappyIngressRequirements> {
        let limits = self.limits.snappy_limits()?;
        let decoded_bytes =
            SnappyStream::decompressed_len_with_limits(compressed, limits).map_err(Error::Iwa)?;
        Ok(SnappyIngressRequirements {
            compressed_bytes: compressed.len(),
            decoded_bytes,
        })
    }

    /// Admit a Snappy stream before its decoded buffer is allocated.
    pub fn admit_snappy_stream<V>(
        &self,
        compressed: &[u8],
        visitor: &mut V,
    ) -> std::result::Result<SnappyIngressRequirements, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        let requirements = self
            .preflight_snappy_stream(compressed)
            .map_err(CandidateIngressError::Physical)?;
        visitor
            .admit(CandidateIngressCharge::new(
                CandidateIngressPhase::SnappyStream,
                requirements.decoded_bytes,
                requirements.compressed_bytes,
                usize::from(requirements.decoded_bytes != 0),
            ))
            .map_err(CandidateIngressError::Admission)?;
        Ok(requirements)
    }

    /// Preflight a decoded IWA stream without constructing generated values.
    pub fn preflight_archive_stream(&self, decoded: &[u8]) -> Result<ArchiveIngressRequirements> {
        Archive::preflight_with_limits(decoded, self.limits.effective_archive_limits()?)
            .map_err(Error::Iwa)
    }

    /// Admit archive/header storage before parsing the decoded IWA stream.
    pub fn admit_archive_stream<V>(
        &self,
        decoded: &[u8],
        visitor: &mut V,
    ) -> std::result::Result<ArchiveIngressRequirements, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        let requirements = self
            .preflight_archive_stream(decoded)
            .map_err(CandidateIngressError::Physical)?;
        visitor
            .admit(CandidateIngressCharge::new(
                CandidateIngressPhase::IwaArchive,
                requirements.retained_bytes(),
                requirements.traversal_bytes(),
                requirements.allocations(),
            ))
            .map_err(CandidateIngressError::Admission)?;
        Ok(requirements)
    }

    /// Admit a parsed archive clone through the shared neutral inventory.
    pub fn admit_archive_clone<V>(
        &self,
        archive: &Archive,
        visitor: &mut V,
    ) -> std::result::Result<ArchiveCloneInventory, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        let mut adapter = CloneVisitor { visitor };
        archive
            .clone_inventory_with_admission(&mut adapter)
            .map_err(|error| match error {
                ArchiveCloneInventoryError::Admission(error) => {
                    CandidateIngressError::Admission(error)
                },
                ArchiveCloneInventoryError::ArithmeticOverflow { resource } => {
                    CandidateIngressError::CloneArithmetic { resource }
                },
                _ => CandidateIngressError::CloneArithmetic {
                    resource: "non-exhaustive clone inventory error",
                },
            })
    }

    fn entry_from_record(&self, ordinal: usize, entry: ZipFileHeaderRecord<'_>) -> CandidateEntry {
        let raw_path = entry.file_path();
        let raw_name = raw_path.as_ref();
        let kind = classify_entry(raw_name, entry.is_dir());
        let declared_bytes = usize_from_u64(entry.uncompressed_size_hint(), "ZIP entry length")
            .unwrap_or(usize::MAX);
        let compressed_bytes =
            usize_from_u64(entry.compressed_size_hint(), "ZIP compressed member length")
                .unwrap_or(usize::MAX);
        let materialized_bytes = match entry.compression_method() {
            CompressionMethod::Store => declared_bytes,
            CompressionMethod::Deflate => declared_bytes.saturating_add(1),
            _ => 0,
        };
        CandidateEntry {
            ordinal,
            wayfinder: entry.wayfinder(),
            kind,
            compression: compression(entry.compression_method()),
            metadata_bytes: ZIP_CENTRAL_FIXED_METADATA_BYTES.saturating_add(
                usize_from_u64(entry.metadata_size_hint(), "ZIP metadata length")
                    .unwrap_or(usize::MAX),
            ),
            compressed_bytes,
            declared_bytes,
            materialized_bytes,
        }
    }

    fn read_entry(&self, entry: CandidateEntry) -> Result<Vec<u8>> {
        let source = self.archive.get_entry_borrowed(entry.wayfinder)?;
        match entry.compression {
            CandidateCompression::Stored => {
                let data = source.data();
                let mut verifier = source.verifying_reader(Cursor::new(data));
                io::copy(&mut verifier, &mut io::sink())?;
                let mut output = Vec::new();
                output
                    .try_reserve_exact(entry.materialized_bytes)
                    .map_err(|_error| Error::Allocation {
                        resource: "candidate ZIP entry output",
                        amount: entry.materialized_bytes,
                    })?;
                output.extend_from_slice(data);
                Ok(output)
            },
            CandidateCompression::Deflated => {
                let mut decoder = DeflateDecoder::new(source.data());
                let verifier = source.verifying_reader(&mut decoder);
                let mut output = Vec::new();
                output
                    .try_reserve_exact(entry.materialized_bytes)
                    .map_err(|_error| Error::Allocation {
                        resource: "candidate ZIP entry output",
                        amount: entry.materialized_bytes,
                    })?;
                verifier
                    .take(
                        u64::try_from(entry.declared_bytes)
                            .map_err(|_error| {
                                Error::InvalidBundle(
                                    "candidate ZIP entry length does not fit u64".to_owned(),
                                )
                            })?
                            .saturating_add(1),
                    )
                    .read_to_end(&mut output)?;
                if output.len() != entry.declared_bytes {
                    return Err(Error::InvalidBundle(
                        "candidate ZIP entry decoded length differs from its declaration"
                            .to_owned(),
                    ));
                }
                Ok(output)
            },
            CandidateCompression::Unsupported => Err(Error::InvalidBundle(
                "candidate ZIP entry uses unsupported compression".to_owned(),
            )),
        }
    }
}

struct CloneVisitor<'a, V> {
    visitor: &'a mut V,
}

impl<V> ArchiveCloneInventoryVisitor for CloneVisitor<'_, V>
where
    V: CandidateIngressVisitor,
{
    type Error = V::Error;

    fn admit(&mut self, chunk: ArchiveCloneInventoryChunk) -> std::result::Result<(), Self::Error> {
        self.visitor.admit(CandidateIngressCharge::new(
            CandidateIngressPhase::IwaClone,
            chunk.retained_bytes(),
            chunk.traversal_bytes(),
            chunk.allocations(),
        ))
    }
}

fn classify_entry(raw_name: &[u8], directory: bool) -> CandidateEntryKind {
    if directory {
        CandidateEntryKind::Directory
    } else if raw_name == b"Index.zip" || raw_name.ends_with(b"/Index.zip") {
        CandidateEntryKind::Index
    } else if raw_name.ends_with(b".iwa") {
        CandidateEntryKind::Iwa
    } else {
        CandidateEntryKind::Other
    }
}

const fn compression(method: CompressionMethod) -> CandidateCompression {
    match method {
        CompressionMethod::Store => CandidateCompression::Stored,
        CompressionMethod::Deflate => CandidateCompression::Deflated,
        _ => CandidateCompression::Unsupported,
    }
}

fn usize_from_u64(value: u64, label: &'static str) -> Result<usize> {
    usize::try_from(value).map_err(|_error| invalid(label))
}

fn invalid(message: &'static str) -> Error {
    Error::InvalidBundle(message.to_owned())
}

fn limit(kind: LimitKind, observed: usize, maximum: u64) -> Error {
    Error::Limit {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum,
    }
}

impl SourceCatalog {
    /// Admit the allocation-bearing preparation phase before calling one of
    /// the `prepare_reassembly*` methods.
    ///
    /// This method only charges physical ZIP/reassembly storage. Semantic
    /// target validation remains the format owner's responsibility, and final
    /// output/offset storage remains the separate execution phase reported by
    /// [`PreparedReassembly::execution_requirements`].
    #[doc(hidden)]
    pub fn __admit_reassembly_preparation<V>(
        &self,
        insertions: &[EntryInsertion<'_>],
        edits: &[EntryEdit<'_>],
        deleted_names: &[&str],
        limits: Limits,
        visitor: &mut V,
    ) -> std::result::Result<ReassemblyPreparationRequirements, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        let requirements = self
            .package()
            .__reassembly_preparation_requirements(insertions, edits, deleted_names, limits)
            .map_err(CandidateIngressError::Physical)?;
        visitor
            .admit(CandidateIngressCharge::new(
                CandidateIngressPhase::ReassemblyPreparation,
                requirements.retained_bytes(),
                requirements.traversal_bytes(),
                requirements.allocations(),
            ))
            .map_err(CandidateIngressError::Admission)?;
        Ok(requirements)
    }

    /// Admit and prepare a physical reassembly in one fallible operation.
    ///
    /// The visitor is called after the allocation-free request/ZIP census and
    /// before the eager ZIP reader or mutation collections are constructed.
    /// Call [`PreparedReassembly::execution_requirements`] and charge that
    /// report before consuming the returned plan with `execute`.
    #[doc(hidden)]
    pub fn __prepare_reassembly_with_changes_admitted<'source, V>(
        &'source self,
        insertions: &'source [EntryInsertion<'source>],
        edits: &'source [EntryEdit<'source>],
        deleted_names: &'source [&'source str],
        limits: Limits,
        visitor: &mut V,
    ) -> std::result::Result<PreparedReassembly<'source>, CandidateIngressError<V::Error>>
    where
        V: CandidateIngressVisitor,
    {
        self.__admit_reassembly_preparation(insertions, edits, deleted_names, limits, visitor)?;
        self.package()
            .prepare_reassembly_with_changes(insertions, edits, deleted_names, limits)
            .map_err(CandidateIngressError::Physical)
    }

    /// Build the allocation-free physical candidate plan used by a caller's
    /// aggregate transaction ledger.
    #[doc(hidden)]
    pub fn __candidate_ingress_with_limits(
        bytes: &[u8],
        limits: Limits,
    ) -> Result<CandidateIngressPlan<'_>> {
        CandidateIngressPlan::from_bytes_with_limits(bytes, limits)
    }
}

#[cfg(test)]
mod tests {
    use super::{
        CandidateCompression, CandidateEntryKind, CandidateIngressCharge, CandidateIngressPhase,
        CandidateIngressPlan,
    };
    use crate::package::{EntryEdit, EntryInsertion};
    use crate::{Limits, SourceCatalog};
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    fn package() -> crate::Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 6000,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        let stream = SnappyStream::compress(&archive.to_bytes()?)?;
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &stream)?;
        writer.write_deflated("Metadata/Properties.plist", b"properties")?;
        writer.finish_to_bytes().map_err(crate::Error::from)
    }

    #[test]
    fn candidate_plan_reports_targets_without_allocating_a_reader_index() -> crate::Result<()> {
        let bytes = package()?;
        let plan = CandidateIngressPlan::from_bytes_with_limits(&bytes, Limits::default())?;
        let entries = plan.entries().collect::<crate::Result<Vec<_>>>()?;
        assert_eq!(entries.len(), 2);
        assert!(entries.iter().any(|entry| {
            entry.kind() == CandidateEntryKind::Iwa
                && entry.compression() == CandidateCompression::Stored
        }));
        assert!(entries.iter().any(|entry| {
            entry.kind() == CandidateEntryKind::Other
                && entry.compression() == CandidateCompression::Deflated
        }));
        assert_eq!(plan.requirements().iwa_entries(), 1);
        Ok(())
    }

    #[test]
    fn candidate_phases_charge_before_decode() -> crate::Result<()> {
        let bytes = package()?;
        let plan = CandidateIngressPlan::from_bytes(&bytes)?;
        let mut phases = Vec::new();
        plan.admit_zip_directory(&mut |charge: CandidateIngressCharge| {
            phases.push(charge.phase());
            Ok::<_, ()>(())
        })
        .map_err(|_| crate::Error::InvalidBundle("zip admission failed".to_owned()))?;
        let entry = plan
            .entries()
            .find_map(|entry| entry.ok().filter(|entry| entry.is_iwa()))
            .ok_or_else(|| crate::Error::InvalidBundle("IWA target missing".to_owned()))?;
        let admitted = plan
            .admit_entry(entry, &mut |charge: CandidateIngressCharge| {
                phases.push(charge.phase());
                Ok::<_, ()>(())
            })
            .map_err(|_| crate::Error::InvalidBundle("entry admission failed".to_owned()))?;
        let snappy = admitted.read()?;
        let requirements = plan
            .admit_snappy_stream(&snappy, &mut |charge: CandidateIngressCharge| {
                phases.push(charge.phase());
                Ok::<_, ()>(())
            })
            .map_err(|_| crate::Error::InvalidBundle("Snappy admission failed".to_owned()))?;
        assert!(requirements.decoded_bytes() > 0);
        let decoded = SnappyStream::decompress(&snappy)?;
        plan.admit_archive_stream(decoded.as_bytes(), &mut |charge: CandidateIngressCharge| {
            phases.push(charge.phase());
            Ok::<_, ()>(())
        })
        .map_err(|_| crate::Error::InvalidBundle("IWA admission failed".to_owned()))?;
        assert_eq!(
            phases,
            vec![
                CandidateIngressPhase::ZipDirectory,
                CandidateIngressPhase::ZipEntryStorage,
                CandidateIngressPhase::SnappyStream,
                CandidateIngressPhase::IwaArchive,
            ]
        );
        Ok(())
    }

    #[test]
    fn candidate_plan_rejects_tight_entry_budget_before_target_read() -> crate::Result<()> {
        let bytes = package()?;
        let limits = Limits::new(
            Limits::MAX_INPUT_BYTES,
            Limits::MAX_ENTRIES,
            1,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            CandidateIngressPlan::from_bytes_with_limits(&bytes, limits),
            Err(crate::Error::Limit {
                kind: crate::LimitKind::EntryBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn reassembly_preparation_is_admitted_before_reader_and_indexes() -> crate::Result<()> {
        let bytes = package()?;
        let catalog = SourceCatalog::from_bytes(&bytes)?;
        let edit = EntryEdit::new("Metadata/Properties.plist", b"edited properties");
        let mut phases = Vec::new();
        let requirements = catalog
            .__admit_reassembly_preparation(
                &[],
                &[edit],
                &[],
                Limits::default(),
                &mut |charge: CandidateIngressCharge| {
                    phases.push(charge.phase());
                    Ok::<_, ()>(())
                },
            )
            .map_err(|_| crate::Error::InvalidBundle("reassembly admission failed".to_owned()))?;
        assert_eq!(phases, vec![CandidateIngressPhase::ReassemblyPreparation]);
        assert_eq!(requirements.edits(), 1);
        assert!(requirements.retained_bytes() > 0);
        assert!(requirements.allocations() > 0);

        let edits = [edit];
        let prepared = catalog.prepare_reassembly(&edits, Limits::default())?;
        assert_eq!(
            prepared.output_bytes(),
            prepared.execution_requirements().output_bytes()
        );
        Ok(())
    }

    #[test]
    fn reassembly_preparation_rejection_happens_before_prepare() -> crate::Result<()> {
        let bytes = package()?;
        let catalog = SourceCatalog::from_bytes(&bytes)?;
        let edit = EntryEdit::new("Metadata/Properties.plist", b"edited properties");
        let result = catalog.__admit_reassembly_preparation(
            &[],
            &[edit],
            &[],
            Limits::default(),
            &mut |_charge: CandidateIngressCharge| Err::<(), _>("ledger full"),
        );
        assert!(matches!(
            result,
            Err(super::CandidateIngressError::Admission("ledger full"))
        ));
        Ok(())
    }

    #[test]
    fn reassembly_preparation_covers_insertions_and_deletions() -> crate::Result<()> {
        let bytes = package()?;
        let catalog = SourceCatalog::from_bytes(&bytes)?;
        let insertions = [EntryInsertion::new("Metadata/New.plist", b"new properties")];
        let edits: [EntryEdit<'_>; 0] = [];
        let deleted_names = ["Metadata/Properties.plist"];
        let mut phases = Vec::new();
        let prepared = catalog
            .__prepare_reassembly_with_changes_admitted(
                &insertions,
                &edits,
                &deleted_names,
                Limits::default(),
                &mut |charge: CandidateIngressCharge| {
                    phases.push(charge.phase());
                    Ok::<_, ()>(())
                },
            )
            .map_err(|_| crate::Error::InvalidBundle("reassembly admission failed".to_owned()))?;
        assert_eq!(phases, vec![CandidateIngressPhase::ReassemblyPreparation]);
        assert_eq!(prepared.execution_requirements().offset_count(), 3);
        Ok(())
    }
}
