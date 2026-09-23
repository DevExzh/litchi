//! Verified compressed transfer of an untouched binary part out of a package
//! opened from an owned source archive (change 0742).
//!
//! A part copied from one package into another is normally published as a
//! new member whose bytes are deflated again from the decoded payload. When
//! the copied part is an untouched, relationship-free, non-XML part of a
//! package that still retains the archive it was opened from, the member's
//! compressed bytes already exist and are already proven to decode to that
//! payload. This module issues a [`CompressedPartTransfer`] for such a part:
//! the ZIP reader proves the member's strict local layout, captures its exact
//! compressed span, decodes the capture, compares every decoded byte with the
//! part's payload and records the actual CRC. A part built with
//! [`BlobPart::with_compressed_transfer`](crate::BlobPart::with_compressed_transfer)
//! then publishes through the targeted writer with fresh known-size framing
//! around those bytes instead of a second Deflate pass.
//!
//! Two questions are kept apart, and both are answered deterministically from
//! the package and its retained archive's bytes:
//!
//! - **Eligibility** ([`OpcPackage::compressed_transfer_size`]) reads
//!   provenance and central-directory metadata only and never decodes a
//!   payload: ownership, identity, content type, relationships, XML-ness,
//!   signature infrastructure, a provable member layout, and a compressed
//!   size bounded by the decoded size.
//! - **Verification** ([`OpcPackage::authorize_compressed_transfer`]) captures
//!   and decodes the member. When the member's own bytes disprove the capture
//!   — a stream that does not consume exactly its declared compressed size, a
//!   checksum, size or descriptor mismatch — the answer is `Ok(None)`: the
//!   caller publishes the part by recompressing it, and the same bytes always
//!   give the same answer. Limits, allocation, I/O and cancellation are typed
//!   errors, never a quiet change of route.

use std::convert::Infallible;
use std::sync::Arc;

use soapberry_zip::office::{
    EntryId, IndexedArchive, VerifiedPrecompressedEntry, VerifiedPrecompressedError,
};

use super::{OpcPackage, SourcePart};
use crate::error::{OpcError, Result, map_io_error};
use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::PackURI;
use crate::part::Part;
use crate::payload::{PartPayload, TransferredPayload};

/// Decoded bytes budgeted per stored-block header when bounding a transferred
/// member's compressed size.
///
/// A Deflate encoder that finds data incompressible frames it in stored
/// blocks of at most 65,535 bytes, each behind a 5-byte header; zlib's default
/// memory level stops at 16 KiB blocks, and smaller memory levels use smaller
/// ones. Budgeting one header per 4 KiB admits every ordinary encoder at four
/// times zlib's default rate, while padding such as runs of empty stored
/// blocks makes a member ineligible, so it is recompressed instead.
const STORED_BLOCK_UNIT: u64 = 4 * 1024;
/// Bytes of one stored-block header.
const STORED_BLOCK_HEADER: u64 = 5;
/// Slack for a final empty block and very short members.
const TRANSFER_SLACK: u64 = 64;

/// The largest compressed size a member of `decoded` bytes may declare and
/// still be transferred, or `None` when the bound overflows.
fn transfer_compressed_bound(decoded: u64) -> Option<u64> {
    decoded
        .div_ceil(STORED_BLOCK_UNIT)
        .checked_mul(STORED_BLOCK_HEADER)?
        .checked_add(decoded)?
        .checked_add(TRANSFER_SLACK)
}

/// A verified compressed representation of one untouched part of a package
/// opened from an owned source archive.
///
/// The value has no public constructor. It is issued only by
/// [`OpcPackage::authorize_compressed_transfer`], after the ZIP reader has
/// proven the member's layout, captured its exact compressed span, decoded
/// that capture and compared every decoded byte with the part's payload. It
/// carries that payload's allocation, the content type the part was opened
/// with, the capture, and the actual CRC; it carries no archive handle, so it
/// retains only the one member's compressed bytes, never the source archive.
///
/// Build the copied part with
/// [`BlobPart::with_compressed_transfer`](crate::BlobPart::with_compressed_transfer).
#[derive(Clone)]
pub struct CompressedPartTransfer {
    payload: Arc<TransferredPayload>,
    content_type: String,
}

impl CompressedPartTransfer {
    /// Content type of the part the transfer was issued for.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Decoded payload bytes the capture was verified against.
    #[must_use]
    pub fn decoded_size(&self) -> usize {
        self.payload.decoded().len()
    }

    /// Exact compressed bytes the capture carries, excluding ZIP framing.
    #[must_use]
    pub fn compressed_size(&self) -> u64 {
        self.payload.compressed().compressed_size()
    }

    pub(crate) fn into_content_type_and_payload(self) -> (String, PartPayload) {
        (self.content_type, PartPayload::Transferred(self.payload))
    }
}

impl std::fmt::Debug for CompressedPartTransfer {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CompressedPartTransfer")
            .field("content_type", &self.content_type)
            .field("payload", &self.payload)
            .finish()
    }
}

/// The retained-archive member a part's transfer is captured from.
struct TransferMember<'package> {
    index: &'package IndexedArchive<Arc<Vec<u8>>>,
    entry_id: EntryId,
    limits: ReadLimits,
}

/// The lazily built ZIP index over an owned package's retained archive.
///
/// `None` records that the archive's own bytes cannot be indexed, which is a
/// deterministic verdict; a failure that is not a property of the bytes
/// (allocation, limits) is returned to the caller and never stored, so a
/// later attempt, or a clone sharing the cell, tries again.
pub(super) type TransferIndexCell = std::sync::OnceLock<Option<Box<IndexedArchive<Arc<Vec<u8>>>>>>;

impl OpcPackage {
    /// The compressed size of this part's source member when
    /// [`Self::authorize_compressed_transfer`] may be asked to capture it,
    /// or `None` when the part is not eligible.
    ///
    /// The answer is a deterministic function of the package's graph, its
    /// provenance and its retained archive's bytes, and it never decodes a
    /// payload. A part is eligible only when all of the following hold:
    ///
    /// - the package retains the owned source archive it was opened from —
    ///   [`Self::from_vec`], [`Self::open`], [`Self::from_reader`], their
    ///   `*_with_limits` forms, [`Self::from_vec_reusing_payloads`] and
    ///   [`Self::from_vec_with_execution`] — with a source member for this
    ///   part;
    /// - the part still holds the payload allocation it was opened with,
    ///   proven by provenance identity rather than by comparing bytes, so a
    ///   payload replaced even with equal bytes is not eligible (a caller
    ///   that wants a verdict about bytes alone reopens the package's
    ///   serialization, in which every part is untouched);
    /// - the part still has the content type it was opened with;
    /// - the part has no relationships;
    /// - the part is not XML by name or content type;
    /// - the package carries no digital-signature infrastructure;
    /// - the member's Store or Deflate layout is provable from the archive's
    ///   headers alone: the local header agrees with the central record, the
    ///   span is bounded and the metadata is not encrypted or unresolved;
    /// - the member's declared compressed size is at most its decoded size
    ///   plus one 5-byte stored-block header per 4 KiB and 64 bytes, so a
    ///   padded member is recompressed rather than published padded.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::PartNotFound`] when no part with `partname`
    /// exists, and an allocation, limit or I/O error that stopped the index
    /// build or the layout proof — never a verdict about the member's bytes.
    pub fn compressed_transfer_size(&self, partname: &PackURI) -> Result<Option<u64>> {
        let part = self.transfer_part(partname)?;
        if self.is_signed() {
            return Ok(None);
        }
        let Some(source_part) = self.transfer_provenance(part) else {
            return Ok(None);
        };
        let Some(member) = self.transfer_member(part, source_part)? else {
            return Ok(None);
        };
        header_verdict(&member)
    }

    /// Capture a verified compressed transfer for one eligible part.
    ///
    /// The part's payload is decoded first if it is still deferred, exactly
    /// as [`Self::get_part`] decodes it. The ZIP reader then captures the
    /// member's exact compressed span (Store or Deflate), decodes the capture,
    /// compares every decoded byte with the part's payload and records the
    /// actual CRC; the member's declared sizes are re-checked against the read
    /// limits the archive was admitted under. The returned value retains the
    /// captured compressed bytes and shares the part's decoded allocation.
    ///
    /// `Ok(None)` means the member's own bytes disprove the capture — for
    /// example a Deflate stream followed by bytes it does not consume, which
    /// the ordinary reader tolerates — so the part must be published by
    /// recompressing it. The same bytes always give the same answer.
    ///
    /// # Errors
    ///
    /// - [`OpcError::PartNotFound`] when no part with `partname` exists;
    /// - [`OpcError::SignedSourceRequiresExplicitPolicy`] for a package with
    ///   digital-signature infrastructure;
    /// - [`OpcError::PreservationUnavailable`] when the part is otherwise not
    ///   eligible (see [`Self::compressed_transfer_size`]);
    /// - the part's own decode refusal when its deferred payload cannot be
    ///   decoded;
    /// - [`OpcError::ReadLimit`] when the member's declared sizes exceed the
    ///   archive's read limits;
    /// - [`OpcError::Allocation`] when the capture cannot be reserved, and an
    ///   I/O error from the retained source.
    pub fn authorize_compressed_transfer(
        &self,
        partname: &PackURI,
    ) -> Result<Option<CompressedPartTransfer>> {
        let part = self.transfer_part(partname)?;
        if self.is_signed() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let Some(source_part) = self.transfer_provenance(part) else {
            return Err(ineligible());
        };
        // Decode through the package's own route, so a deferred payload's
        // refusal is the one every other accessor reports.
        part.ensure_payload()?;
        let decoded = part.blob_arc();
        // A deferred part shares its decode cell with the provenance, so the
        // decode just performed is the provenance's own allocation. Re-prove
        // it on the decoded allocation the capture is compared against.
        if !source_part
            .blob
            .decoded()
            .is_some_and(|source| Arc::ptr_eq(source, &decoded))
        {
            return Err(ineligible());
        }
        let Some(member) = self.transfer_member(part, source_part)? else {
            return Err(ineligible());
        };
        if header_verdict(&member)?.is_none() {
            return Err(ineligible());
        }
        let Some(physical) = verified_capture(&member, &decoded)? else {
            return Ok(None);
        };
        let mut content_type = String::new();
        content_type
            .try_reserve_exact(part.content_type().len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC compressed transfer content type",
                source,
            })?;
        content_type.push_str(part.content_type());
        Ok(Some(CompressedPartTransfer {
            payload: Arc::new(TransferredPayload::new(decoded, physical)),
            content_type,
        }))
    }

    /// Resolve a part for transfer without decoding its payload.
    fn transfer_part(&self, partname: &PackURI) -> Result<&dyn Part> {
        if let Some(part) = self.parts.get(partname) {
            return Ok(&**part);
        }
        self.find_case_insensitive(partname)
            .map(|(_, part)| part)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// The source provenance of a part the graph still holds as opened, or
    /// `None`. Never decodes; the signature, layout and size conditions are
    /// checked by the callers.
    fn transfer_provenance(&self, part: &dyn Part) -> Option<&SourcePart> {
        self.source_archive.as_ref()?;
        let provenance = self.preservation.as_deref()?;
        let source_part = provenance.parts.get(part.partname())?;
        if !source_part.member_present
            || !part.rels().is_empty()
            || part.content_type() != source_part.content_type
            || xml_minifier::audit::package::is_xml_part(
                part.partname().as_str(),
                part.content_type(),
            )
        {
            return None;
        }
        // Provenance identity, never a byte comparison: a payload replaced
        // with equal bytes is a caller's payload, not the source member's.
        let handle = part.payload_handle();
        let untouched = match (handle.payload(), &source_part.blob) {
            (PartPayload::Deferred(current), PartPayload::Deferred(source)) => {
                Arc::ptr_eq(current, source)
            },
            (current, source) => match (current.decoded(), source.decoded()) {
                (Some(current), Some(source)) => Arc::ptr_eq(current, source),
                _ => false,
            },
        };
        untouched.then_some(source_part)
    }

    /// The retained-archive member a part with transfer provenance is
    /// captured from, or `None` when the archive's own bytes cannot index it.
    ///
    /// A deferred part is reached through its provenance payload, which is
    /// the very cell the part holds, and uses the index that payload's decode
    /// builds, after proving that index covers this package's retained
    /// archive. A part the package materialized eagerly, or a deferred part
    /// whose own index was refused, uses the package's transfer index.
    fn transfer_member<'package>(
        &'package self,
        part: &'package dyn Part,
        source_part: &'package SourcePart,
    ) -> Result<Option<TransferMember<'package>>> {
        let Some(source_archive) = self.source_archive.as_ref() else {
            return Ok(None);
        };
        let deferred = source_part.blob.as_deferred();
        if let Some(deferred) = deferred {
            if !Arc::ptr_eq(deferred.source().bytes(), source_archive) {
                return Ok(None);
            }
            // The decode's index caches its refusal (ADR 0030) as an
            // `OpcError` that no longer says whether the bytes caused it, so a
            // refusal is classified again through the transfer index below.
            if let Ok(index) = deferred.source().index() {
                return Ok(index
                    .entry_id(deferred.member())
                    .map(|entry_id| TransferMember {
                        index,
                        entry_id,
                        limits: deferred.source().limits(),
                    }));
            }
        }
        let Some(index) = self.owned_transfer_index(source_archive)? else {
            return Ok(None);
        };
        let member = deferred.map_or_else(
            || part.partname().membername(),
            |deferred| deferred.member(),
        );
        Ok(index.entry_id(member).map(|entry_id| TransferMember {
            index,
            entry_id,
            limits: self.source_limits,
        }))
    }

    /// One index of the retained archive per package open, built on first
    /// use under the limits the archive was admitted under and shared by
    /// clones, or `None` when the archive's own bytes cannot be indexed
    /// ([`soapberry_zip::Error::is_content_fault`]). Any other failure is
    /// returned and not stored, so a later call, or a clone, tries again.
    fn owned_transfer_index(
        &self,
        source_archive: &Arc<Vec<u8>>,
    ) -> Result<Option<&IndexedArchive<Arc<Vec<u8>>>>> {
        let Some(cell) = self.transfer_index.as_ref() else {
            return Ok(None);
        };
        if let Some(state) = cell.get() {
            return Ok(state.as_deref());
        }
        match IndexedArchive::from_reader_with_limits(
            Arc::clone(source_archive),
            source_archive.len() as u64,
            self.source_limits.zip_limits(),
        ) {
            Ok(index) => {
                let _raced = cell.set(Some(Box::new(index)));
            },
            Err(error) if error.is_content_fault() => {
                let _raced = cell.set(None);
            },
            Err(error) => return Err(OpcError::from(error)),
        }
        Ok(cell.get().and_then(Option::as_deref))
    }
}

/// The header-only part of eligibility: a provable layout and a bounded
/// compressed size. Returns the member's compressed size when both hold.
fn header_verdict(member: &TransferMember<'_>) -> Result<Option<u64>> {
    let metadata = match member.index.metadata_for(member.entry_id) {
        Ok(metadata) => metadata,
        Err(error) if error.is_content_fault() => return Ok(None),
        Err(error) => return Err(OpcError::from(error)),
    };
    let within_bound = transfer_compressed_bound(metadata.uncompressed_size())
        .is_some_and(|bound| metadata.compressed_size() <= bound);
    if !within_bound
        || !member
            .index
            .precompressed_layout_provable(member.entry_id)
            .map_err(OpcError::from)?
    {
        return Ok(None);
    }
    Ok(Some(metadata.compressed_size()))
}

/// Capture one member's verified compressed span against `decoded`.
///
/// The member's declared sizes are re-checked against `limits` before any
/// byte is captured; the ZIP reader then captures the exact span, decodes it,
/// compares every decoded byte with `decoded` and records the actual CRC.
/// `Ok(None)` is a content fault of the member's own bytes.
fn verified_capture(
    member: &TransferMember<'_>,
    decoded: &[u8],
) -> Result<Option<VerifiedPrecompressedEntry>> {
    let limits = member.limits;
    let metadata = match member.index.metadata_for(member.entry_id) {
        Ok(metadata) => metadata,
        Err(error) if error.is_content_fault() => return Ok(None),
        Err(error) => return Err(OpcError::from(error)),
    };
    let logical_size = u64::try_from(decoded.len()).map_err(|_| {
        OpcError::ZipError("compressed transfer decoded size exceeds u64".to_owned())
    })?;
    if metadata.uncompressed_size() != logical_size {
        return Ok(None);
    }
    limits.check(
        ReadResource::ArchiveCompressedBytes,
        metadata.compressed_size(),
        limits.max_archive_compressed_bytes(),
    )?;
    limits.check(
        ReadResource::ArchiveEntryBytes,
        metadata.uncompressed_size(),
        limits.max_archive_entry_bytes(),
    )?;
    limits.check(
        ReadResource::PartBytes,
        logical_size,
        limits.max_part_bytes(),
    )?;
    match member
        .index
        .read_entry_precompressed_with_progress(member.entry_id, decoded, |_| {
            Ok::<(), Infallible>(())
        }) {
        Ok(physical) => Ok(Some(physical)),
        Err(VerifiedPrecompressedError::Archive(error)) if error.is_content_fault() => Ok(None),
        Err(error) => Err(map_precompressed_error(error)),
    }
}

fn ineligible() -> OpcError {
    OpcError::PreservationUnavailable {
        reason: "the part is not eligible for a verified compressed transfer".to_owned(),
    }
}

fn map_precompressed_error(error: VerifiedPrecompressedError<Infallible>) -> OpcError {
    match error {
        VerifiedPrecompressedError::Archive(error) => OpcError::from(error),
        VerifiedPrecompressedError::Transport(error) => map_io_error(error),
        VerifiedPrecompressedError::Callback(never) => match never {},
        _ => OpcError::ZipError("unrecognized verified ZIP compressed-transfer failure".to_owned()),
    }
}

#[cfg(test)]
mod tests;
