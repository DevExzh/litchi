//! Bounded decoded-member splicing for source-backed OPC packages.
//!
//! The format owner proves the logical insertion point and, for XML Parts, the
//! candidate grammar. This module owns only source-backed replay, physical ZIP
//! preservation, and source/sink error precedence. It deliberately retains no
//! complete decoded source or candidate payload.

use super::{
    Chunked, ContextCheckedSink, Counted, OpcError, OpcOperationAccounting, OutputBudgetedSink,
    PartView, Result, SourceArtifact, SourceArtifactFingerprint, SourceBackedPackage,
    SourceCheckedSink, SourceSnapshot, VerifiedDecodedReaderError, finish_source_publication,
    map_execution_error, map_io_error, map_preservation_error, overlay_unavailable,
};
use crate::error::SpliceResource;
use crate::limits::ReadResource;
use crate::packuri::PackURI;
use litchi_core::{ExecutionContext, Reservation, Resource, SourceVersion};
use sha2::{Digest as _, Sha256};
use soapberry_zip::ReaderAt;
use soapberry_zip::{PreservationEntryId, ReplayLimits, ReplayPublicationError, ReplayResource};
use std::fmt;
use std::io::{self, BufRead, Read, Write};
use std::sync::{Arc, Mutex};

const DEFAULT_MAX_FRAGMENT_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_SOURCE_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_CANDIDATE_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSED_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_ARCHIVE_BYTES: u64 = 2 * 1024 * 1024 * 1024;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;
const DEFAULT_MAX_XML_WORKSPACE_BYTES: u64 = 64 * 1024 * 1024;
const DEFAULT_MAX_PRESERVATION_MEMORY_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_REPLAY_MEMORY_BYTES: u64 = 8 * 1024 * 1024;

/// Finite limits for one decoded-member splice.
///
/// These limits bound the logical source, retained insertion fragment,
/// candidate member, generated compressed member, and complete publication.
/// Preservation-index and compressor allocations are separate package-level
/// costs; this type does not claim that those costs are proportional to the
/// insertion fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourcePartSpliceLimits {
    pub max_source_bytes: u64,
    pub max_fragment_bytes: u64,
    pub max_candidate_bytes: u64,
    pub max_compressed_bytes: u64,
    pub max_archive_bytes: u64,
    pub max_output_bytes: u64,
    /// Maximum reserved XML parser or decoded-audit-adapter workspace. XML
    /// targets include parser state; every changed target includes this
    /// module's fixed adapter buffer.
    pub max_xml_workspace_bytes: u64,
    /// Maximum reserved preservation-index metadata workspace.
    pub max_preservation_memory_bytes: u64,
    /// Maximum reserved replay writer and compressor workspace.
    pub max_replay_memory_bytes: u64,
    /// Explicit streaming XML audit profile. The default narrows the
    /// tokenizer to 64 KiB; callers that need a larger OOXML token must opt
    /// into that finite profile explicitly.
    pub xml_audit_limits: xml_minifier::audit::Limits,
}

impl SourcePartSpliceLimits {
    /// Construct explicit finite splice limits.
    pub fn new(
        max_source_bytes: u64,
        max_fragment_bytes: u64,
        max_candidate_bytes: u64,
        max_compressed_bytes: u64,
        max_archive_bytes: u64,
        max_output_bytes: u64,
    ) -> Result<Self> {
        let limits = Self {
            max_source_bytes,
            max_fragment_bytes,
            max_candidate_bytes,
            max_compressed_bytes,
            max_archive_bytes,
            max_output_bytes,
            max_xml_workspace_bytes: DEFAULT_MAX_XML_WORKSPACE_BYTES,
            max_preservation_memory_bytes: DEFAULT_MAX_PRESERVATION_MEMORY_BYTES,
            max_replay_memory_bytes: DEFAULT_MAX_REPLAY_MEMORY_BYTES,
            xml_audit_limits: default_xml_audit_limits(),
        };
        limits.validate()?;
        Ok(limits)
    }

    fn validate(self) -> Result<()> {
        for (resource, value) in [
            (ReadResource::PartBytes, self.max_source_bytes),
            (ReadResource::PartBytes, self.max_fragment_bytes),
            (ReadResource::PartBytes, self.max_candidate_bytes),
            (
                ReadResource::ArchiveCompressedBytes,
                self.max_compressed_bytes,
            ),
            (ReadResource::ArchiveTotalBytes, self.max_archive_bytes),
        ] {
            if value == 0 || value == u64::MAX {
                return Err(OpcError::InvalidReadLimit { resource, value });
            }
        }
        validate_splice_limit(SpliceResource::OutputBytes, self.max_output_bytes)?;
        for (resource, value) in [
            (
                SpliceResource::XmlWorkspaceBytes,
                self.max_xml_workspace_bytes,
            ),
            (
                SpliceResource::PreservationMemoryBytes,
                self.max_preservation_memory_bytes,
            ),
            (
                SpliceResource::ReplayMemoryBytes,
                self.max_replay_memory_bytes,
            ),
        ] {
            validate_splice_limit(resource, value)?;
        }
        Ok(())
    }

    fn archive_limit(self) -> u64 {
        self.max_archive_bytes.min(self.max_output_bytes)
    }

    /// Replace the finite XML audit profile used for source and candidate
    /// verification.
    #[must_use]
    pub const fn with_xml_audit_limits(
        mut self,
        xml_audit_limits: xml_minifier::audit::Limits,
    ) -> Self {
        self.xml_audit_limits = xml_audit_limits;
        self
    }

    /// Bound the reserved XML parser or decoded-audit-adapter workspace
    /// independently of payload size.
    #[must_use]
    pub const fn with_max_xml_workspace_bytes(mut self, maximum: u64) -> Self {
        self.max_xml_workspace_bytes = maximum;
        self
    }

    /// Bound preservation-index metadata workspace independently of payload
    /// and compressor workspace.
    #[must_use]
    pub const fn with_max_preservation_memory_bytes(mut self, maximum: u64) -> Self {
        self.max_preservation_memory_bytes = maximum;
        self
    }

    /// Bound replay writer and compressor workspace independently of source
    /// and candidate bytes.
    #[must_use]
    pub const fn with_max_replay_memory_bytes(mut self, maximum: u64) -> Self {
        self.max_replay_memory_bytes = maximum;
        self
    }
}

impl Default for SourcePartSpliceLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: DEFAULT_MAX_SOURCE_BYTES,
            max_fragment_bytes: DEFAULT_MAX_FRAGMENT_BYTES,
            max_candidate_bytes: DEFAULT_MAX_CANDIDATE_BYTES,
            max_compressed_bytes: DEFAULT_MAX_COMPRESSED_BYTES,
            max_archive_bytes: DEFAULT_MAX_ARCHIVE_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_xml_workspace_bytes: DEFAULT_MAX_XML_WORKSPACE_BYTES,
            max_preservation_memory_bytes: DEFAULT_MAX_PRESERVATION_MEMORY_BYTES,
            max_replay_memory_bytes: DEFAULT_MAX_REPLAY_MEMORY_BYTES,
            xml_audit_limits: default_xml_audit_limits(),
        }
    }
}

fn default_xml_audit_limits() -> xml_minifier::audit::Limits {
    xml_minifier::audit::Limits::default()
        .narrow(xml_minifier::audit::Resource::TokenBytes, 64 * 1024)
}

/// Compact logical proof supplied by a format-owned source/candidate scanner.
///
/// `insertion_offset` and all lengths refer to decoded member bytes. Hashes
/// are SHA-256 over raw decoded bytes, before XML normalization. An insertion
/// plan does not remove bytes, so the candidate length must equal source length
/// plus fragment length.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourcePartSpliceProof {
    pub source_version: SourceVersion,
    pub source_len: u64,
    pub source_sha256: [u8; 32],
    pub insertion_offset: u64,
    pub fragment_len: u64,
    pub fragment_sha256: [u8; 32],
    pub candidate_len: u64,
    pub candidate_sha256: [u8; 32],
}

impl SourcePartSpliceProof {
    /// Return whether this proof describes a zero-byte insertion.
    #[must_use]
    pub const fn is_exact_insertion_noop(self) -> bool {
        self.fragment_len == 0
            && self.candidate_len == self.source_len
            && self.insertion_offset <= self.source_len
    }
}

/// Fixed-length authored storage reserved before allocation by its source
/// package. The mutable slice cannot increase the retained capacity.
///
/// Pass this owner to
/// [`SourceBackedPackage::prepare_source_part_splice_with_fragment`] to transfer
/// its existing memory reservation into the plan without charging it twice.
/// The reservation covers the requested byte-buffer capacity, not allocator
/// bookkeeping or caller-owned encoder state. Storage with a larger reported
/// capacity is refused before it can be authored or retained in a plan.
pub struct SourcePartSpliceFragment<'package> {
    package: &'package SourceBackedPackage,
    bytes: Vec<u8>,
    reservation: Option<Arc<Reservation>>,
}

impl fmt::Debug for SourcePartSpliceFragment<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePartSpliceFragment")
            .field("bytes", &self.bytes.len())
            .finish_non_exhaustive()
    }
}

impl SourcePartSpliceFragment<'_> {
    /// Borrow the fixed-length authored bytes for hashing or validation.
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Fill the preallocated fragment without resizing its storage.
    pub fn as_mut_slice(&mut self) -> &mut [u8] {
        self.bytes.as_mut_slice()
    }
}

/// A prepared source-backed decoded-member splice.
///
/// The plan borrows the immutable source-backed package and retains only the
/// source artifact handle, proof scalars, and bounded immutable fragment. The
/// package's physical ZIP implementation remains private to OPC.
pub struct SourcePartSplicePlan<'package> {
    package: &'package SourceBackedPackage,
    target: usize,
    audit_xml: bool,
    fragment: Arc<Vec<u8>>,
    proof: SourcePartSpliceProof,
    limits: SourcePartSpliceLimits,
    source_artifact: SourceArtifact,
    fragment_memory_reservation: Option<Arc<Reservation>>,
    xml_workspace_reservation: Option<Arc<Reservation>>,
}

impl fmt::Debug for SourcePartSplicePlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePartSplicePlan")
            .field("target", &self.target)
            .field("audit_xml", &self.audit_xml)
            .field("fragment_bytes", &self.fragment.len())
            .field(
                "fragment_memory_reservation",
                &self
                    .fragment_memory_reservation
                    .as_ref()
                    .map(|reservation| reservation.amount()),
            )
            .field(
                "xml_workspace_reservation",
                &self
                    .xml_workspace_reservation
                    .as_ref()
                    .map(|reservation| reservation.amount()),
            )
            .field("proof", &self.proof)
            .field("limits", &self.limits)
            .finish()
    }
}

impl<'package> SourcePartSplicePlan<'package> {
    /// Return the compact source/candidate proof captured by this plan.
    #[must_use]
    pub const fn proof(&self) -> SourcePartSpliceProof {
        self.proof
    }

    /// Return the finite limits captured by this plan.
    #[must_use]
    pub const fn limits(&self) -> SourcePartSpliceLimits {
        self.limits
    }

    /// Return whether this plan copies the exact source artifact.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.proof.is_exact_insertion_noop()
    }

    /// Publish the plan to a sequential sink and retain an exact inverse
    /// authorization for the resulting artifact.
    pub fn write_to_stream<W: Write>(self, writer: W) -> Result<SourcePartSplicePublication> {
        self.write_to_stream_inner(writer, None)
    }

    /// Publish while recording accepted output and low-level ZIP work in an
    /// OPC-owned accounting report.
    pub fn write_to_stream_with_accounting<W: Write>(
        self,
        writer: W,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<SourcePartSplicePublication> {
        self.write_to_stream_inner(writer, Some(accounting))
    }

    fn write_to_stream_inner<W: Write>(
        self,
        writer: W,
        mut accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<SourcePartSplicePublication> {
        self.package.source.ensure_current()?;
        self.package
            .cache
            .check_context()
            .map_err(map_execution_error)?;

        if self.is_noop() {
            // A zero-byte insertion is byte-identical by construction. This
            // branch intentionally bypasses XML and signature checks so a
            // malformed or signed source can still be reproduced exactly.
            limits_check(
                ReadResource::ArchiveTotalBytes,
                self.package.source.length,
                self.limits.archive_limit(),
            )?;
            let mut hashing = HashingSink::new(writer);
            match accounting.as_deref_mut() {
                Some(report) => self
                    .source_artifact
                    .write_to_stream_with_accounting(&mut hashing, report)?,
                None => self.source_artifact.write_to_stream(&mut hashing)?,
            }
            let candidate_artifact = hashing.finish();
            return Ok(SourcePartSplicePublication {
                source_artifact: self.source_artifact,
                candidate_artifact,
                proof: self.proof,
            });
        }

        if self.package.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        if self.package.has_encrypted_entries() {
            return Err(overlay_unavailable(
                "decoded splice refuses encrypted ZIP members",
            ));
        }
        let preservation_memory = preservation_memory_requirement(self.package)?;
        let _preservation_memory_reservation = reserve_memory(
            self.package.source.context.as_ref(),
            preservation_memory,
            self.limits.max_preservation_memory_bytes,
            SpliceResource::PreservationMemoryBytes,
        )?;
        let replay_memory = replay_memory_requirement()?;
        let _replay_memory_reservation = reserve_memory(
            self.package.source.context.as_ref(),
            replay_memory,
            self.limits.max_replay_memory_bytes,
            SpliceResource::ReplayMemoryBytes,
        )?;
        self.package.source.monitor_publication();
        self.package.check_topology_progress()?;

        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC decoded splice preservation index",
                source,
            })?;
        scratch.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let index = self
            .package
            .archive
            .preservation_index_with_limits(&mut scratch, self.package.limits.zip_limits())
            .map_err(map_preservation_error)?;
        self.package.source.ensure_current()?;
        if index.archive_end_offset() != self.package.source.length {
            return Err(overlay_unavailable(
                "source ZIP archive has trailing bytes outside its located archive",
            ));
        }

        let target_name = self.package.parts[self.target].partname.membername();
        let preserved_target = index
            .entries()
            .iter()
            .find(|entry| entry.raw_name_bytes() == target_name.as_bytes())
            .ok_or_else(|| {
                overlay_unavailable(
                    "decoded splice target does not have one canonical UTF-8 source member",
                )
            })?;
        let preservation_target: PreservationEntryId = preserved_target.id();
        let compression = preserved_target.compression_method();
        let replay_limits = ReplayLimits::new(
            self.limits.max_candidate_bytes,
            self.limits.max_compressed_bytes,
            self.limits.archive_limit(),
        )
        .map_err(map_preservation_error)?;

        let source_snapshot = self.package.source.clone();
        let part = PartView {
            package: self.package,
            index: self.target,
        };
        let proof = self.proof;
        let fragment = Arc::clone(&self.fragment);
        let mut hashing = HashingSink::new(writer);
        let mut written = 0_u64;
        let mut accounting_error = None;
        let execution_failure = Arc::new(Mutex::new(None));
        let mut zip_accounting = soapberry_zip::ZipOperationAccounting::default();
        let mut decoded_accounting = OpcOperationAccounting::default();

        let mut callback = |output: &mut dyn Write| -> Result<()> {
            let result = part.with_verified_decoded_reader_with_accounting(
                |reader| {
                    stream_splice(
                        reader,
                        output,
                        proof,
                        fragment.as_slice(),
                        self.limits.xml_audit_limits,
                        part.partname().as_str(),
                        self.audit_xml,
                        &source_snapshot,
                        self.package.source.context.as_ref(),
                    )
                },
                &mut decoded_accounting,
            );
            match result {
                Ok(()) => Ok(()),
                Err(VerifiedDecodedReaderError::Opc { error, .. })
                | Err(VerifiedDecodedReaderError::Callback(error)) => Err(error),
            }
        };

        let result = if let Some(context) = self.package.cache.context() {
            let output_reservation_failures = self
                .package
                .source
                .output_reservation_failures
                .as_ref()
                .ok_or_else(|| {
                    overlay_unavailable("managed source output reservation counter is unavailable")
                })?
                .clone();
            let counted = match accounting.as_deref_mut() {
                Some(report) => Counted::with_accounting(
                    &mut hashing,
                    &mut written,
                    report,
                    &mut accounting_error,
                    false,
                ),
                None => Counted::new(&mut hashing, &mut written),
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: source_snapshot.clone(),
            };
            let cooperative = ContextCheckedSink {
                inner: checked,
                context: Some(context.clone()),
                failure: Arc::clone(&execution_failure),
            };
            let budgeted = OutputBudgetedSink {
                inner: cooperative,
                context: context.clone(),
                failure: Arc::clone(&execution_failure),
                output_reservation_failures,
            };
            match run_replay(
                &index,
                preservation_target,
                compression,
                replay_limits,
                Chunked { inner: budgeted },
                &mut callback,
                &mut zip_accounting,
            ) {
                Ok(mut sink) => sink.flush().map_err(map_io_error),
                Err(error) => Err(map_replay_error(error)),
            }
        } else {
            let counted = match accounting.as_deref_mut() {
                Some(report) => Counted::with_accounting(
                    &mut hashing,
                    &mut written,
                    report,
                    &mut accounting_error,
                    false,
                ),
                None => Counted::new(&mut hashing, &mut written),
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: source_snapshot.clone(),
            };
            let cooperative = ContextCheckedSink {
                inner: checked,
                context: None,
                failure: Arc::clone(&execution_failure),
            };
            match run_replay(
                &index,
                preservation_target,
                compression,
                replay_limits,
                Chunked { inner: cooperative },
                &mut callback,
                &mut zip_accounting,
            ) {
                Ok(mut sink) => sink.flush().map_err(map_io_error),
                Err(error) => Err(map_replay_error(error)),
            }
        };

        let result = execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = match (
            result,
            self.package
                .cache
                .check_context()
                .map_err(map_execution_error),
        ) {
            (Err(error), _) => Err(error),
            (Ok(()), Err(error)) => Err(error),
            (Ok(()), Ok(())) => Ok(()),
        };
        let merge_result = accounting.map_or(Ok(()), |report| {
            let mut first_error = report.merge_from(&decoded_accounting).err();
            if let Err(error) = report.merge_zip(&zip_accounting)
                && first_error.is_none()
            {
                first_error = Some(error);
            }
            first_error.map_or(Ok(()), Err)
        });
        let result = finish_source_publication(result, &self.package.source, written);
        result?;
        if let Some(error) = accounting_error {
            return Err(error);
        }
        merge_result?;
        if hashing.accepted != written {
            return Err(overlay_unavailable(
                "decoded splice sink progress differs from accepted output",
            ));
        }
        let candidate_artifact = hashing.finish();
        Ok(SourcePartSplicePublication {
            source_artifact: self.source_artifact,
            candidate_artifact,
            proof: self.proof,
        })
    }
}

impl SourceBackedPackage {
    /// Reserve and allocate fixed-length fragment storage before authoring.
    ///
    /// The buffer starts filled with zeroes. Its length is checked against
    /// both `maximum` and the package's part limit. Managed packages charge
    /// memory before allocation and work while initializing bounded chunks.
    /// Dropping the buffer releases its reservation. This method performs no
    /// publication or XML admission; the prepared splice verifies its bytes.
    pub fn allocate_source_part_splice_fragment(
        &self,
        length: u64,
        maximum: u64,
    ) -> Result<SourcePartSpliceFragment<'_>> {
        if maximum == 0 || maximum == u64::MAX {
            return Err(OpcError::InvalidReadLimit {
                resource: ReadResource::PartBytes,
                value: maximum,
            });
        }
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        limits_check(ReadResource::PartBytes, length, maximum)?;
        limits_check(
            ReadResource::PartBytes,
            length,
            self.limits.max_part_bytes(),
        )?;
        let capacity = usize::try_from(length)
            .map_err(|_| overlay_unavailable("decoded splice fragment length exceeds usize"))?;
        let reservation = self
            .source
            .context
            .as_ref()
            .filter(|_| length != 0)
            .map(|context| {
                context
                    .reserve(Resource::Memory, length)
                    .map(Arc::new)
                    .map_err(map_execution_error)
            })
            .transpose()?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC authored fragment",
                source,
            })?;
        if bytes.capacity() != capacity {
            return Err(overlay_unavailable(
                "authored fragment allocation reported capacity beyond its reservation",
            ));
        }
        while bytes.len() < capacity {
            self.source.ensure_current()?;
            self.cache.check_context().map_err(map_execution_error)?;
            let count = (capacity - bytes.len()).min(super::SOURCE_PUBLICATION_CHUNK_BYTES);
            if let Some(context) = self.source.context.as_ref() {
                context
                    .consume(Resource::Work, count as u64)
                    .map_err(map_execution_error)?;
            }
            bytes.resize(bytes.len() + count, 0);
        }
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        Ok(SourcePartSpliceFragment {
            package: self,
            bytes,
            reservation,
        })
    }

    /// Prepare an insertion while transferring the fragment's original
    /// preallocation reservation into the returned plan.
    ///
    /// The fragment must have been allocated by this exact package instance;
    /// another instance, even over identical source bytes, cannot substitute
    /// a different execution budget. All ordinary splice proof, XML, source,
    /// and publication checks remain in force.
    pub fn prepare_source_part_splice_with_fragment(
        &self,
        partname: &PackURI,
        proof: SourcePartSpliceProof,
        fragment: SourcePartSpliceFragment<'_>,
        limits: SourcePartSpliceLimits,
    ) -> Result<SourcePartSplicePlan<'_>> {
        if !std::ptr::eq(self, fragment.package) {
            return Err(overlay_unavailable(
                "decoded splice fragment belongs to a different source package",
            ));
        }
        self.prepare_source_part_splice_inner(
            partname,
            proof,
            Arc::new(fragment.bytes),
            limits,
            true,
            fragment.reservation,
        )
    }

    /// Prepare one bounded decoded insertion into an existing Part.
    ///
    /// The format owner must have already validated its source and candidate
    /// grammar and supplied hashes over raw decoded bytes. This method checks
    /// the physical target, finite limits, source freshness, and source proof;
    /// it never materializes the source member. A zero-byte insertion is an
    /// exact no-op and intentionally skips XML and signature policy checks.
    pub fn prepare_source_part_splice(
        &self,
        partname: &PackURI,
        proof: SourcePartSpliceProof,
        fragment: Arc<Vec<u8>>,
        limits: SourcePartSpliceLimits,
    ) -> Result<SourcePartSplicePlan<'_>> {
        self.prepare_source_part_splice_inner(partname, proof, fragment, limits, false, None)
    }

    fn prepare_source_part_splice_inner(
        &self,
        partname: &PackURI,
        proof: SourcePartSpliceProof,
        fragment: Arc<Vec<u8>>,
        limits: SourcePartSpliceLimits,
        fragment_preallocated: bool,
        reservation: Option<Arc<Reservation>>,
    ) -> Result<SourcePartSplicePlan<'_>> {
        limits.validate()?;
        self.source.ensure_current()?;
        self.cache.check_context().map_err(map_execution_error)?;
        let target = self
            .part_index(partname)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
        let audit_xml = xml_minifier::audit::package::is_xml_part(
            partname.as_str(),
            &self.parts[target].content_type,
        );
        let target_entry = self.parts[target].entry_id;
        let source_version = self.source.version();
        if proof.source_version != source_version {
            return Err(OpcError::SourceChanged {
                expected: proof.source_version,
                actual: source_version,
            });
        }
        let declared_source_len = self
            .archive
            .metadata_for(target_entry)
            .map_err(map_preservation_error)?
            .uncompressed_size();
        if proof.source_len != declared_source_len {
            return Err(overlay_unavailable(
                "decoded splice source length does not match its ZIP declaration",
            ));
        }
        self.limits.check(
            ReadResource::PartBytes,
            proof.source_len,
            self.limits.max_part_bytes(),
        )?;
        self.limits.check(
            ReadResource::PartBytes,
            proof.source_len,
            limits.max_source_bytes,
        )?;
        let fragment_len = u64::try_from(fragment.len())
            .map_err(|_| overlay_unavailable("decoded splice fragment length exceeds u64"))?;
        if proof.fragment_len != fragment_len {
            return Err(overlay_unavailable(
                "decoded splice fragment length proof does not match the retained fragment",
            ));
        }
        limits.validate()?;
        limits_check(
            ReadResource::PartBytes,
            fragment_len,
            limits.max_fragment_bytes,
        )?;
        let fragment_capacity = u64::try_from(fragment.capacity())
            .map_err(|_| overlay_unavailable("decoded splice fragment capacity exceeds u64"))?;
        limits_check(
            ReadResource::PartBytes,
            fragment_capacity,
            limits.max_fragment_bytes,
        )?;
        if digest_fragment(&self.source, fragment.as_slice())? != proof.fragment_sha256 {
            return Err(overlay_unavailable(
                "decoded splice fragment hash does not match its proof",
            ));
        }
        let candidate_len = proof
            .source_len
            .checked_add(fragment_len)
            .ok_or_else(|| overlay_unavailable("decoded splice candidate length overflows u64"))?;
        if proof.candidate_len != candidate_len {
            return Err(overlay_unavailable(
                "decoded splice candidate length is not source plus fragment",
            ));
        }
        limits_check(
            ReadResource::PartBytes,
            candidate_len,
            limits.max_candidate_bytes,
        )?;
        if proof.insertion_offset > proof.source_len {
            return Err(overlay_unavailable(
                "decoded splice insertion offset is outside the source member",
            ));
        }
        let candidate_len_usize = usize::try_from(candidate_len).map_err(|_| {
            overlay_unavailable("decoded splice candidate length does not fit this platform")
        })?;
        self.validate_overlay_limits(std::iter::once((target, candidate_len_usize)))?;
        let is_noop = proof.is_exact_insertion_noop();
        if !is_noop {
            if self.has_signature_infrastructure() {
                return Err(OpcError::SignedSourceRequiresExplicitPolicy);
            }
            if self.has_encrypted_entries() {
                return Err(overlay_unavailable(
                    "decoded splice refuses encrypted ZIP members",
                ));
            }
        }
        let fragment_memory_reservation = if fragment_preallocated {
            reservation
        } else {
            self.source
                .context
                .as_ref()
                .filter(|_| fragment_capacity != 0)
                .map(|context| {
                    context
                        .reserve(Resource::Memory, fragment_capacity)
                        .map(Arc::new)
                        .map_err(map_execution_error)
                })
                .transpose()?
        };
        let xml_workspace_reservation = if is_noop {
            None
        } else {
            let workspace = splice_workspace_requirement(limits, audit_xml)?;
            reserve_memory(
                self.source.context.as_ref(),
                workspace,
                limits.max_xml_workspace_bytes,
                SpliceResource::XmlWorkspaceBytes,
            )?
        };
        if is_noop {
            verify_noop_source_proof(&self.part(partname)?, proof)?;
        } else {
            verify_source_proof(
                &self.part(partname)?,
                proof,
                limits.xml_audit_limits,
                audit_xml,
            )?;
            verify_candidate_proof(
                &self.part(partname)?,
                proof,
                fragment.as_slice(),
                limits.xml_audit_limits,
                audit_xml,
            )?;
            self.source.ensure_current()?;
            self.cache.check_context().map_err(map_execution_error)?;
        }
        Ok(SourcePartSplicePlan {
            package: self,
            target,
            audit_xml,
            fragment,
            proof: SourcePartSpliceProof { ..proof },
            limits,
            source_artifact: self.source_artifact(),
            fragment_memory_reservation,
            xml_workspace_reservation,
        })
    }
}

/// Exact publication authorization returned by a successful splice.
#[derive(Clone)]
pub struct SourcePartSplicePublication {
    source_artifact: SourceArtifact,
    candidate_artifact: SourceArtifactFingerprint,
    proof: SourcePartSpliceProof,
}

impl fmt::Debug for SourcePartSplicePublication {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePartSplicePublication")
            .field("candidate_artifact", &self.candidate_artifact)
            .field("proof", &self.proof)
            .finish()
    }
}

impl SourcePartSplicePublication {
    /// Return the proof used to produce this publication.
    #[must_use]
    pub const fn proof(&self) -> SourcePartSpliceProof {
        self.proof
    }

    /// Return the exact candidate artifact fingerprint.
    #[must_use]
    pub const fn candidate_artifact_fingerprint(&self) -> SourceArtifactFingerprint {
        self.candidate_artifact
    }

    /// Return whether this publication used the exact source artifact path.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.proof.is_exact_insertion_noop()
    }

    /// Restore the retained source artifact after authorizing the exact
    /// candidate artifact currently supplied by `current`.
    pub fn write_inverse_to_stream<W: Write>(
        &self,
        current: &SourceBackedPackage,
        writer: W,
    ) -> Result<()> {
        current.source.ensure_current()?;
        current.cache.check_context().map_err(map_execution_error)?;
        let retained = &self.source_artifact.snapshot;
        retained.ensure_current()?;
        let operation_context = current
            .cache
            .context()
            .cloned()
            .or_else(|| retained.context.clone());
        let output_reservation_failures = if current.cache.context().is_some() {
            current.source.output_reservation_failures.clone()
        } else {
            retained.output_reservation_failures.clone()
        };
        if operation_context.is_some() && output_reservation_failures.is_none() {
            return Err(overlay_unavailable(
                "managed source output reservation counter is unavailable",
            ));
        }
        current.source.monitor_publication();
        retained.monitor_publication();
        let _workspace_reservation = operation_context
            .as_ref()
            .map(|context| {
                context
                    .reserve(
                        Resource::Memory,
                        super::SOURCE_PUBLICATION_CHUNK_BYTES as u64,
                    )
                    .map(Arc::new)
                    .map_err(map_execution_error)
            })
            .transpose()?;
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(super::SOURCE_PUBLICATION_CHUNK_BYTES)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC inverse publication buffer",
                source,
            })?;
        buffer.resize(super::SOURCE_PUBLICATION_CHUNK_BYTES, 0);
        let actual = fingerprint_inverse_snapshot(
            current,
            &current.source,
            retained,
            operation_context.as_ref(),
            &mut buffer,
        )?;
        if actual != self.candidate_artifact {
            return Err(overlay_unavailable(
                "decoded splice inverse candidate artifact does not match",
            ));
        }
        let execution_failure = Arc::new(Mutex::new(None));
        let mut written = 0_u64;
        let counted = Counted::new(writer, &mut written);
        let checked = SourceCheckedSink {
            inner: counted,
            snapshot: current.source.clone(),
        };
        let cooperative = ContextCheckedSink {
            inner: checked,
            context: operation_context.clone(),
            failure: Arc::clone(&execution_failure),
        };
        let result = if let Some(context) = operation_context.clone() {
            let mut budgeted = OutputBudgetedSink {
                inner: cooperative,
                context,
                failure: Arc::clone(&execution_failure),
                output_reservation_failures: output_reservation_failures
                    .expect("managed inverse output counter checked above"),
            };
            let result = copy_inverse_snapshot(
                current,
                retained,
                operation_context.as_ref(),
                &mut buffer,
                &mut budgeted,
            );
            result.and_then(|()| budgeted.flush().map_err(map_io_error))
        } else {
            let mut cooperative = cooperative;
            let result =
                copy_inverse_snapshot(current, retained, None, &mut buffer, &mut cooperative);
            result.and_then(|()| cooperative.flush().map_err(map_io_error))
        };
        let result = execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = match (
            result,
            check_inverse_state(current, retained, operation_context.as_ref()),
        ) {
            (_, Err(error @ OpcError::SourceChanged { .. })) => Err(error),
            (Err(error), _) => Err(error),
            (Ok(()), Err(error)) => Err(error),
            (Ok(()), Ok(())) => Ok(()),
        };
        finish_inverse_publication(result, &current.source, retained, written)
    }
}

fn limits_check(resource: ReadResource, actual: u64, maximum: u64) -> Result<()> {
    if actual > maximum {
        return Err(OpcError::ReadLimit {
            resource,
            actual,
            maximum,
        });
    }
    Ok(())
}

fn validate_splice_limit(resource: SpliceResource, value: u64) -> Result<()> {
    if value == 0 || value == u64::MAX {
        return Err(OpcError::InvalidSourcePartSpliceLimit { resource, value });
    }
    Ok(())
}

fn splice_limits_check(resource: SpliceResource, actual: u64, maximum: u64) -> Result<()> {
    if actual > maximum {
        return Err(OpcError::SourcePartSpliceLimit {
            resource,
            actual,
            maximum,
        });
    }
    Ok(())
}

fn finish_inverse_publication(
    result: Result<()>,
    source: &SourceSnapshot,
    retained: &SourceSnapshot,
    written: u64,
) -> Result<()> {
    let freshness = source
        .ensure_current()
        .and_then(|()| retained.ensure_current());
    if let Err(error) = freshness {
        return if written != 0 {
            Err(OpcError::IncompleteOutput {
                written,
                source: Box::new(error),
            })
        } else {
            Err(error)
        };
    }
    match result {
        Err(error @ OpcError::IncompleteOutput { .. }) => Err(error),
        Err(error) if written != 0 => Err(OpcError::IncompleteOutput {
            written,
            source: Box::new(error),
        }),
        other => other,
    }
}

fn check_inverse_state(
    current: &SourceBackedPackage,
    retained: &SourceSnapshot,
    operation_context: Option<&ExecutionContext>,
) -> Result<()> {
    current.source.ensure_current()?;
    retained.ensure_current()?;
    current.cache.check_context().map_err(map_execution_error)?;
    if let Some(context) = retained.context.as_ref() {
        context.check().map_err(map_execution_error)?;
    }
    if let Some(context) = operation_context {
        context.check().map_err(map_execution_error)?;
    }
    Ok(())
}

fn fingerprint_inverse_snapshot(
    current: &SourceBackedPackage,
    snapshot: &SourceSnapshot,
    retained: &SourceSnapshot,
    operation_context: Option<&ExecutionContext>,
    buffer: &mut [u8],
) -> Result<SourceArtifactFingerprint> {
    let mut hasher = Sha256::new();
    let mut offset = 0_u64;
    while offset < snapshot.length {
        check_inverse_state(current, retained, operation_context)?;
        let remaining = usize::try_from((snapshot.length - offset).min(buffer.len() as u64))
            .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
        let read = super::read_source_at_with_context(
            snapshot,
            operation_context,
            offset,
            &mut buffer[..remaining],
            "inverse fingerprinting",
        )?;
        if read == 0 {
            return Err(OpcError::IoError(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "source-backed OPC source ended during inverse fingerprinting",
            )));
        }
        check_inverse_state(current, retained, operation_context)?;
        hasher.update(&buffer[..read]);
        offset = offset
            .checked_add(read as u64)
            .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
    }
    check_inverse_state(current, retained, operation_context)?;
    Ok(SourceArtifactFingerprint::from_sha256(
        hasher.finalize().into(),
    ))
}

fn copy_inverse_snapshot(
    current: &SourceBackedPackage,
    retained: &SourceSnapshot,
    operation_context: Option<&ExecutionContext>,
    buffer: &mut [u8],
    output: &mut dyn Write,
) -> Result<()> {
    let mut offset = 0_u64;
    while offset < retained.length {
        check_inverse_state(current, retained, operation_context)?;
        let remaining = usize::try_from((retained.length - offset).min(buffer.len() as u64))
            .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
        let read = super::read_source_at_with_context(
            retained,
            operation_context,
            offset,
            &mut buffer[..remaining],
            "inverse publication",
        )?;
        if read == 0 {
            return Err(OpcError::IoError(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "source-backed OPC source ended during inverse publication",
            )));
        }
        check_inverse_state(current, retained, operation_context)?;
        output.write_all(&buffer[..read]).map_err(map_io_error)?;
        check_inverse_state(current, retained, operation_context)?;
        offset = offset
            .checked_add(read as u64)
            .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
    }
    Ok(())
}

fn splice_workspace_requirement(limits: SourcePartSpliceLimits, audit_xml: bool) -> Result<u64> {
    let parser = if audit_xml {
        limits
            .xml_audit_limits
            .streaming_memory_upper_bound()
            .ok_or_else(|| overlay_unavailable("XML audit workspace bound overflows usize"))?
    } else {
        0
    };
    u64::try_from(parser)
        .ok()
        .and_then(|parser| parser.checked_add(super::SOURCE_PUBLICATION_CHUNK_BYTES as u64))
        .ok_or_else(|| overlay_unavailable("XML audit workspace bound exceeds u64"))
}

fn preservation_memory_requirement(package: &SourceBackedPackage) -> Result<u64> {
    let ownership = package
        .archive
        .preservation_memory_upper_bound()
        .ok_or_else(|| overlay_unavailable("preservation workspace bound overflows u64"))?;
    ownership
        .checked_add(soapberry_zip::RECOMMENDED_BUFFER_SIZE as u64)
        .ok_or_else(|| overlay_unavailable("preservation workspace bound overflows u64"))
}

fn replay_memory_requirement() -> Result<u64> {
    soapberry_zip::replay_memory_upper_bound()
        .ok_or_else(|| overlay_unavailable("replay workspace bound overflows u64"))
}

fn reserve_memory(
    context: Option<&ExecutionContext>,
    amount: u64,
    maximum: u64,
    resource: SpliceResource,
) -> Result<Option<Arc<Reservation>>> {
    splice_limits_check(resource, amount, maximum)?;
    context
        .filter(|_| amount != 0)
        .map(|context| {
            context
                .reserve(Resource::Memory, amount)
                .map(Arc::new)
                .map_err(map_execution_error)
        })
        .transpose()
}

fn verify_source_proof(
    part: &PartView<'_>,
    proof: SourcePartSpliceProof,
    limits: xml_minifier::audit::Limits,
    audit_xml: bool,
) -> Result<()> {
    let partname = part.partname().as_str().to_owned();
    let observed = part.with_verified_decoded_reader(|reader| {
        if audit_xml {
            let mut sink = io::sink();
            audit_splice(
                reader,
                &mut sink,
                proof,
                &[],
                proof.source_len,
                limits,
                &partname,
                false,
                true,
                Some(&part.package.source),
                part.package.source.context.as_ref(),
            )
        } else {
            hash_reader(reader).map(|(length, hash)| (length, hash, length, hash))
        }
    });
    let observed = match observed {
        Ok(observed) => observed,
        Err(VerifiedDecodedReaderError::Opc { error, .. })
        | Err(VerifiedDecodedReaderError::Callback(error)) => return Err(error),
    };
    if observed.0 != proof.source_len || observed.1 != proof.source_sha256 {
        return Err(overlay_unavailable(
            "decoded splice source hash or length does not match its proof",
        ));
    }
    Ok(())
}

fn verify_noop_source_proof(part: &PartView<'_>, proof: SourcePartSpliceProof) -> Result<()> {
    let observed = part.with_verified_decoded_reader(hash_reader);
    let observed = match observed {
        Ok(observed) => observed,
        Err(VerifiedDecodedReaderError::Opc { error, .. })
        | Err(VerifiedDecodedReaderError::Callback(error)) => return Err(error),
    };
    if observed.0 != proof.source_len
        || observed.1 != proof.source_sha256
        || proof.candidate_sha256 != proof.source_sha256
        || proof.fragment_sha256 != digest_bytes(&[])
    {
        return Err(overlay_unavailable(
            "decoded splice no-op hashes do not describe the exact source",
        ));
    }
    Ok(())
}

fn hash_reader(reader: &mut dyn BufRead) -> Result<(u64, [u8; 32])> {
    let mut hasher = Sha256::new();
    let mut length = 0_u64;
    loop {
        let amount;
        {
            let bytes = reader.fill_buf().map_err(map_io_error)?;
            if bytes.is_empty() {
                break;
            }
            amount = bytes.len();
            hasher.update(bytes);
        }
        length = length
            .checked_add(
                u64::try_from(amount)
                    .map_err(|_| overlay_unavailable("decoded source length exceeds u64"))?,
            )
            .ok_or_else(|| overlay_unavailable("decoded source length overflows u64"))?;
        reader.consume(amount);
    }
    Ok((length, hasher.finalize().into()))
}

fn verify_candidate_proof(
    part: &PartView<'_>,
    proof: SourcePartSpliceProof,
    fragment: &[u8],
    limits: xml_minifier::audit::Limits,
    audit_xml: bool,
) -> Result<()> {
    let partname = part.partname().as_str().to_owned();
    let observed = part.with_verified_decoded_reader(|reader| {
        let mut sink = io::sink();
        audit_splice(
            reader,
            &mut sink,
            proof,
            fragment,
            proof.insertion_offset,
            limits,
            &partname,
            true,
            audit_xml,
            Some(&part.package.source),
            part.package.source.context.as_ref(),
        )
    });
    let observed = match observed {
        Ok(observed) => observed,
        Err(VerifiedDecodedReaderError::Opc { error, .. })
        | Err(VerifiedDecodedReaderError::Callback(error)) => return Err(error),
    };
    if observed.2 != proof.candidate_len || observed.3 != proof.candidate_sha256 {
        return Err(overlay_unavailable(
            "decoded splice candidate hash or length does not match its proof",
        ));
    }
    Ok(())
}

fn digest_bytes(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn digest_fragment(source: &SourceSnapshot, fragment: &[u8]) -> Result<[u8; 32]> {
    let mut hasher = Sha256::new();
    for chunk in fragment.chunks(super::SOURCE_PUBLICATION_CHUNK_BYTES) {
        source.ensure_current()?;
        if let Some(context) = source.context.as_ref() {
            context
                .consume(
                    Resource::Work,
                    u64::try_from(chunk.len())
                        .map_err(|_| overlay_unavailable("splice fragment length exceeds u64"))?,
                )
                .map_err(map_execution_error)?;
        }
        source.ensure_current()?;
        hasher.update(chunk);
    }
    source.ensure_current()?;
    Ok(hasher.finalize().into())
}

fn stream_splice(
    reader: &mut dyn BufRead,
    output: &mut dyn Write,
    proof: SourcePartSpliceProof,
    fragment: &[u8],
    limits: xml_minifier::audit::Limits,
    partname: &str,
    audit_xml: bool,
    source: &SourceSnapshot,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    let observed = audit_splice(
        reader,
        output,
        proof,
        fragment,
        proof.insertion_offset,
        limits,
        partname,
        true,
        audit_xml,
        Some(source),
        context,
    )?;
    if observed.2 != proof.candidate_len || observed.3 != proof.candidate_sha256 {
        return Err(overlay_unavailable(
            "decoded splice candidate hash or length changed during replay",
        ));
    }
    Ok(())
}

fn audit_splice(
    reader: &mut dyn BufRead,
    output: &mut dyn Write,
    proof: SourcePartSpliceProof,
    fragment: &[u8],
    insertion_offset: u64,
    limits: xml_minifier::audit::Limits,
    partname: &str,
    check_candidate: bool,
    audit_xml: bool,
    source: Option<&SourceSnapshot>,
    context: Option<&ExecutionContext>,
) -> Result<(u64, [u8; 32], u64, [u8; 32])> {
    let mut splice_reader = SpliceAuditReader::new(
        reader,
        output,
        proof.source_len,
        insertion_offset,
        fragment,
        source,
        context,
    )?;
    if audit_xml {
        let audit_result = xml_minifier::audit::verify_authored_reader(&mut splice_reader, limits);
        if let Some(error) = splice_reader.take_failure() {
            return Err(error);
        }
        match audit_result {
            Ok(_) => {},
            Err(xml_minifier::audit::StreamError::Input(error)) => {
                return Err(map_io_error(error));
            },
            Err(xml_minifier::audit::StreamError::Audit(source)) => {
                return Err(OpcError::XmlPublication {
                    part: partname.to_string(),
                    source,
                });
            },
            Err(_) => {
                return Err(overlay_unavailable(
                    "decoded splice XML audit returned an unknown stream failure",
                ));
            },
        }
    } else {
        let drain_result = drain_splice_reader(&mut splice_reader);
        if let Some(error) = splice_reader.take_failure() {
            return Err(error);
        }
        drain_result?;
    }
    let observed = splice_reader.finish()?;
    if observed.0 != proof.source_len || observed.1 != proof.source_sha256 {
        return Err(overlay_unavailable(
            "decoded splice source hash or length changed during replay",
        ));
    }
    if check_candidate && observed.2 != proof.candidate_len {
        return Err(overlay_unavailable(
            "decoded splice candidate length changed during replay",
        ));
    }
    Ok(observed)
}

fn drain_splice_reader(reader: &mut SpliceAuditReader<'_, '_, '_, '_, '_>) -> Result<()> {
    loop {
        let amount;
        {
            let bytes = reader.fill_buf().map_err(map_io_error)?;
            if bytes.is_empty() {
                break;
            }
            amount = bytes.len();
        }
        reader.consume(amount);
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SplicePhase {
    Prefix,
    Fragment,
    Suffix,
    Done,
}

/// A bounded logical candidate view used by XML auditing or binary draining.
///
/// `quick_xml` consumes a `BufRead` view, so this adapter copies at most one
/// fixed-size chunk from the verified source into its own buffer. Bytes are
/// written to the replay sink only when the auditor consumes them. That keeps
/// candidate validation and publication on the same byte stream without
/// retaining the complete source or candidate.
struct SpliceAuditReader<'source, 'fragment, 'output, 'snapshot, 'context> {
    source: &'source mut dyn BufRead,
    fragment: &'fragment [u8],
    output: &'output mut dyn Write,
    source_snapshot: Option<&'snapshot SourceSnapshot>,
    context: Option<&'context ExecutionContext>,
    insertion_offset: u64,
    source_length: u64,
    source_position: u64,
    fragment_position: usize,
    phase: SplicePhase,
    buffer: Vec<u8>,
    buffer_start: usize,
    buffer_len: usize,
    buffer_source_bytes: usize,
    source_hash: Sha256,
    candidate_hash: Sha256,
    candidate_length: u64,
    output_accepted: u64,
    pending_error: Option<io::Error>,
    pending_failure: Option<OpcError>,
}

impl<'source, 'fragment, 'output, 'snapshot, 'context>
    SpliceAuditReader<'source, 'fragment, 'output, 'snapshot, 'context>
{
    fn new(
        source: &'source mut dyn BufRead,
        output: &'output mut dyn Write,
        source_length: u64,
        insertion_offset: u64,
        fragment: &'fragment [u8],
        source_snapshot: Option<&'snapshot SourceSnapshot>,
        context: Option<&'context ExecutionContext>,
    ) -> Result<Self> {
        if insertion_offset > source_length {
            return Err(overlay_unavailable(
                "decoded splice insertion offset is outside the source member",
            ));
        }
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(super::SOURCE_PUBLICATION_CHUNK_BYTES)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC decoded splice audit buffer",
                source,
            })?;
        buffer.resize(super::SOURCE_PUBLICATION_CHUNK_BYTES, 0);
        Ok(Self {
            source,
            fragment,
            output,
            source_snapshot,
            context,
            insertion_offset,
            source_length,
            source_position: 0,
            fragment_position: 0,
            phase: SplicePhase::Prefix,
            buffer,
            buffer_start: 0,
            buffer_len: 0,
            buffer_source_bytes: 0,
            source_hash: Sha256::new(),
            candidate_hash: Sha256::new(),
            candidate_length: 0,
            output_accepted: 0,
            pending_error: None,
            pending_failure: None,
        })
    }

    fn set_error(&mut self, error: io::Error) {
        if self.pending_error.is_none() {
            self.pending_error = Some(error);
        }
    }

    fn check_pending(&mut self) -> io::Result<()> {
        if let Some(error) = &self.pending_failure {
            return Err(io::Error::other(error.to_string()));
        }
        if let Some(error) = self.pending_error.take() {
            Err(error)
        } else {
            Ok(())
        }
    }

    fn set_failure(&mut self, error: OpcError) {
        if self.pending_failure.is_none() {
            self.pending_failure = Some(error);
        }
    }

    fn take_failure(&mut self) -> Option<OpcError> {
        self.pending_failure.take()
    }

    fn advance_phase(&mut self) {
        loop {
            let next = match self.phase {
                SplicePhase::Prefix if self.source_position >= self.insertion_offset => {
                    SplicePhase::Fragment
                },
                SplicePhase::Fragment if self.fragment_position >= self.fragment.len() => {
                    SplicePhase::Suffix
                },
                SplicePhase::Suffix if self.source_position >= self.source_length => {
                    SplicePhase::Done
                },
                phase => phase,
            };
            if next == self.phase {
                return;
            }
            self.phase = next;
        }
    }

    fn fill_source_buffer(&mut self, remaining: u64) -> io::Result<&[u8]> {
        let count;
        {
            let available = self.source.fill_buf()?;
            if available.is_empty() {
                return Err(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "decoded splice source ended before its declared length",
                ));
            }
            count = available
                .len()
                .min(usize::try_from(remaining).unwrap_or(usize::MAX))
                .min(self.buffer.len());
            self.buffer[..count].copy_from_slice(&available[..count]);
        }
        self.buffer_start = 0;
        self.buffer_len = count;
        self.buffer_source_bytes = count;
        Ok(&self.buffer[..count])
    }

    fn fill_fragment_buffer(&mut self) -> &[u8] {
        let remaining = self.fragment.len().saturating_sub(self.fragment_position);
        let count = remaining.min(self.buffer.len());
        self.buffer[..count].copy_from_slice(
            &self.fragment[self.fragment_position..self.fragment_position + count],
        );
        self.buffer_start = 0;
        self.buffer_len = count;
        self.buffer_source_bytes = 0;
        &self.buffer[..count]
    }

    fn add_candidate_length(&mut self, amount: usize) {
        match self
            .candidate_length
            .checked_add(u64::try_from(amount).unwrap_or(u64::MAX))
        {
            Some(length) => self.candidate_length = length,
            None => self.set_error(io::Error::new(
                io::ErrorKind::InvalidData,
                "decoded splice candidate length overflows u64",
            )),
        }
    }

    fn write_consumed_range(&mut self, start: usize, amount: usize) {
        let mut offset = 0usize;
        while offset < amount && self.pending_error.is_none() && self.pending_failure.is_none() {
            match self
                .output
                .write(&self.buffer[start + offset..start + amount])
            {
                Ok(0) => self.set_error(io::Error::new(
                    io::ErrorKind::WriteZero,
                    "decoded splice sink accepted no progress",
                )),
                Ok(written) if written > amount - offset => self.set_error(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "decoded splice sink accepted more bytes than supplied",
                )),
                Ok(written) => {
                    self.output_accepted = self
                        .output_accepted
                        .checked_add(u64::try_from(written).unwrap_or(u64::MAX))
                        .unwrap_or(u64::MAX);
                    offset += written;
                },
                Err(error) => self.set_error(error),
            }
        }
    }

    fn finish(mut self) -> Result<(u64, [u8; 32], u64, [u8; 32])> {
        if let Some(error) = self.pending_failure.take() {
            return Err(error);
        }
        if self.buffer_start < self.buffer_len {
            return Err(overlay_unavailable(
                "decoded splice validation returned before consuming its candidate",
            ));
        }
        self.advance_phase();
        if self.phase != SplicePhase::Done {
            return Err(overlay_unavailable(
                "decoded splice validation did not reach candidate end",
            ));
        }
        self.check_pending().map_err(map_io_error)?;
        if self.source_position != self.source_length {
            return Err(overlay_unavailable(
                "decoded splice source length changed during replay",
            ));
        }
        if self.output_accepted != self.candidate_length {
            return Err(overlay_unavailable(
                "decoded splice sink accepted fewer candidate bytes than audited",
            ));
        }
        Ok((
            self.source_position,
            self.source_hash.finalize().into(),
            self.candidate_length,
            self.candidate_hash.finalize().into(),
        ))
    }
}

impl Read for SpliceAuditReader<'_, '_, '_, '_, '_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let amount = available.len().min(output.len());
        output[..amount].copy_from_slice(&available[..amount]);
        self.consume(amount);
        Ok(amount)
    }
}

impl BufRead for SpliceAuditReader<'_, '_, '_, '_, '_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        self.check_pending()?;
        if self.buffer_start < self.buffer_len {
            return Ok(&self.buffer[self.buffer_start..self.buffer_len]);
        }
        self.advance_phase();
        match self.phase {
            SplicePhase::Prefix => {
                let remaining = self.insertion_offset - self.source_position;
                self.fill_source_buffer(remaining)
            },
            SplicePhase::Fragment => Ok(self.fill_fragment_buffer()),
            SplicePhase::Suffix => {
                let remaining = self.source_length - self.source_position;
                self.fill_source_buffer(remaining)
            },
            SplicePhase::Done => Ok(&[]),
        }
    }

    fn consume(&mut self, amount: usize) {
        let available = self.buffer_len.saturating_sub(self.buffer_start);
        let amount = amount.min(available);
        if amount == 0 {
            return;
        }
        let start = self.buffer_start;
        let source_amount = self.buffer_source_bytes.saturating_sub(start).min(amount);
        if self.pending_failure.is_some() || self.pending_error.is_some() {
            return;
        }
        if source_amount == 0 {
            if let Some(source) = self.source_snapshot {
                if let Err(error) = source.ensure_current() {
                    self.set_failure(error);
                    return;
                }
            }
            if let Some(context) = self.context {
                let work = match u64::try_from(amount) {
                    Ok(work) => work,
                    Err(_) => {
                        self.set_failure(overlay_unavailable(
                            "decoded splice fragment length exceeds u64",
                        ));
                        return;
                    },
                };
                if let Err(error) = context
                    .consume(Resource::Work, work)
                    .map_err(map_execution_error)
                {
                    self.set_failure(error);
                    return;
                }
            }
            if let Some(source) = self.source_snapshot {
                if let Err(error) = source.ensure_current() {
                    self.set_failure(error);
                    return;
                }
            }
        }
        if source_amount != 0 {
            self.source_hash
                .update(&self.buffer[start..start + source_amount]);
        }
        self.candidate_hash
            .update(&self.buffer[start..start + amount]);
        self.add_candidate_length(amount);
        self.write_consumed_range(start, amount);
        if source_amount != 0 {
            self.source.consume(source_amount);
            self.source_position = self.source_position.saturating_add(source_amount as u64);
        } else {
            self.fragment_position = self.fragment_position.saturating_add(amount);
        }
        self.buffer_start += amount;
        if self.buffer_start == self.buffer_len {
            self.buffer_start = 0;
            self.buffer_len = 0;
            self.buffer_source_bytes = 0;
            self.advance_phase();
        }
    }
}

fn map_replay_error(error: ReplayPublicationError<OpcError>) -> OpcError {
    match error {
        ReplayPublicationError::Archive { source, .. } => map_preservation_error(source),
        ReplayPublicationError::Callback { source, .. } => source,
        ReplayPublicationError::Limit {
            resource,
            actual,
            maximum,
            ..
        } => OpcError::ReadLimit {
            resource: match resource {
                ReplayResource::DecodedBytes => ReadResource::PartBytes,
                ReplayResource::CompressedBytes => ReadResource::ArchiveCompressedBytes,
                ReplayResource::ArchiveBytes => ReadResource::ArchiveTotalBytes,
            },
            actual,
            maximum,
        },
        ReplayPublicationError::Sink { source, .. } => map_io_error(source),
        ReplayPublicationError::NonDeterministic { .. } => {
            overlay_unavailable("decoded splice replay produced non-deterministic output")
        },
        _ => overlay_unavailable("decoded splice replay returned an unknown failure"),
    }
}

fn run_replay<'archive, R, W, F>(
    index: &soapberry_zip::PreservationIndex<'archive, R>,
    target: PreservationEntryId,
    compression: soapberry_zip::CompressionMethod,
    limits: ReplayLimits,
    sink: W,
    callback: F,
    accounting: &mut soapberry_zip::ZipOperationAccounting,
) -> std::result::Result<W, ReplayPublicationError<OpcError>>
where
    R: ReaderAt,
    W: Write,
    F: FnMut(&mut dyn Write) -> Result<()>,
{
    index.write_replacing_with_replay_with_accounting(
        target,
        compression,
        limits,
        sink,
        accounting,
        callback,
    )
}

struct HashingSink<W> {
    inner: W,
    hasher: Sha256,
    accepted: u64,
}

impl<W> HashingSink<W> {
    fn new(inner: W) -> Self {
        Self {
            inner,
            hasher: Sha256::new(),
            accepted: 0,
        }
    }

    fn finish(self) -> SourceArtifactFingerprint {
        SourceArtifactFingerprint::from_sha256(self.hasher.finalize().into())
    }
}

impl<W: Write> Write for HashingSink<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = self.inner.write(bytes)?;
        if accepted > bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "decoded splice sink accepted more bytes than supplied",
            ));
        }
        self.hasher.update(&bytes[..accepted]);
        self.accepted = self
            .accepted
            .checked_add(u64::try_from(accepted).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidData,
                    "decoded splice sink byte overflow",
                )
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "decoded splice sink overflow")
            })?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}
