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
//! The eligibility predicate is a deterministic function of the package: it
//! reads provenance and metadata, never decodes, and never depends on memory
//! pressure. Every failure after a part is found eligible — a decode refusal,
//! a limit, an allocation failure, a layout or checksum mismatch — is a typed
//! error, never a silent fallback to the re-deflating route.

use std::convert::Infallible;
use std::sync::Arc;

use soapberry_zip::office::{IndexedArchive, VerifiedPrecompressedError};

use super::{OpcPackage, SourcePart};
use crate::error::{OpcError, Result, map_io_error};
use crate::limits::ReadResource;
use crate::packuri::PackURI;
use crate::part::Part;
use crate::payload::{PartPayload, TransferredPayload};

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

impl OpcPackage {
    /// Whether [`Self::authorize_compressed_transfer`] may issue a verified
    /// compressed transfer for this part.
    ///
    /// The answer is a deterministic function of the package and never
    /// decodes a payload. A part is eligible only when all of the following
    /// hold:
    ///
    /// - the package retains the owned source archive it was opened from
    ///   ([`Self::from_vec`], [`Self::open`], [`Self::from_reader`] and their
    ///   `*_with_limits` forms), with a source member for this part;
    /// - the part still holds the payload allocation it was opened with,
    ///   proven by provenance identity rather than by comparing bytes, so a
    ///   payload replaced even with equal bytes is not eligible;
    /// - the part still has the content type it was opened with;
    /// - the part has no relationships;
    /// - the part is not XML by name or content type;
    /// - the package carries no digital-signature infrastructure.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::PartNotFound`] when no part with `partname`
    /// exists.
    pub fn compressed_transfer_eligible(&self, partname: &PackURI) -> Result<bool> {
        let part = self.transfer_part(partname)?;
        Ok(self.compressed_transfer_source(part).is_some())
    }

    /// Issue a verified compressed transfer for one eligible part.
    ///
    /// The part's payload is decoded first if it is still deferred, exactly
    /// as [`Self::get_part`] decodes it. The ZIP reader then proves the
    /// member's strict local layout, captures its exact compressed span
    /// (Store or Deflate), decodes the capture, compares every decoded byte
    /// with the part's payload and records the actual CRC; the member's
    /// declared sizes are re-checked against the read limits the archive was
    /// admitted under. The returned value retains the captured compressed
    /// bytes and shares the part's decoded allocation.
    ///
    /// # Errors
    ///
    /// - [`OpcError::PartNotFound`] when no part with `partname` exists;
    /// - [`OpcError::SignedSourceRequiresExplicitPolicy`] for a package with
    ///   digital-signature infrastructure;
    /// - [`OpcError::PreservationUnavailable`] when the part is not eligible
    ///   (see [`Self::compressed_transfer_eligible`]);
    /// - the part's own decode refusal when its deferred payload cannot be
    ///   decoded;
    /// - [`OpcError::ReadLimit`] when the member's declared sizes exceed the
    ///   archive's read limits;
    /// - [`OpcError::Allocation`] when the capture cannot be reserved;
    /// - [`OpcError::ZipError`] for a layout, size, checksum or data-descriptor
    ///   mismatch, or for a capture that does not decode to the payload.
    pub fn authorize_compressed_transfer(
        &self,
        partname: &PackURI,
    ) -> Result<CompressedPartTransfer> {
        let part = self.transfer_part(partname)?;
        if self.is_signed() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let Some(source_part) = self.compressed_transfer_source(part) else {
            return Err(ineligible());
        };
        let Some(source_archive) = self.source_archive.as_ref() else {
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

        let handle = part.payload_handle();
        let built;
        let (index, member, limits) = match handle.payload().as_deferred() {
            Some(deferred) => {
                // The provenance identity above proves this is the package's
                // own deferred source; the archive identity is checked too so
                // a capture can only ever come from `source_archive`.
                if !Arc::ptr_eq(deferred.source().bytes(), source_archive) {
                    return Err(ineligible());
                }
                (
                    deferred.source().index()?,
                    deferred.member(),
                    deferred.source().limits(),
                )
            },
            None => {
                let limits = self.source_limits;
                built = IndexedArchive::from_reader_with_limits(
                    Arc::clone(source_archive),
                    source_archive.len() as u64,
                    limits.zip_limits(),
                )
                .map_err(OpcError::from)?;
                (&built, partname_member(part.partname()), limits)
            },
        };
        let entry_id = index.entry_id(member).ok_or_else(|| {
            OpcError::ZipError(format!(
                "compressed transfer source member {member} is missing from the retained archive"
            ))
        })?;
        let metadata = index.metadata_for(entry_id).map_err(OpcError::from)?;
        let logical_size = u64::try_from(decoded.len()).map_err(|_| {
            OpcError::ZipError("compressed transfer decoded size exceeds u64".to_owned())
        })?;
        if metadata.uncompressed_size() != logical_size {
            return Err(OpcError::ZipError(format!(
                "compressed transfer expected {logical_size} decoded bytes but ZIP metadata declares {}",
                metadata.uncompressed_size()
            )));
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
        let physical = index
            .read_entry_precompressed_with_progress(entry_id, decoded.as_slice(), |_| {
                Ok::<(), Infallible>(())
            })
            .map_err(map_precompressed_error)?;
        let mut content_type = String::new();
        content_type
            .try_reserve_exact(part.content_type().len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC compressed transfer content type",
                source,
            })?;
        content_type.push_str(part.content_type());
        Ok(CompressedPartTransfer {
            payload: Arc::new(TransferredPayload::new(decoded, physical)),
            content_type,
        })
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

    /// The source provenance of an eligible part, or `None` when the part is
    /// not eligible. Never decodes.
    fn compressed_transfer_source(&self, part: &dyn Part) -> Option<&SourcePart> {
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
            || self.is_signed()
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
}

/// The member name the owned reader used for a part it materialized eagerly.
fn partname_member(partname: &PackURI) -> &str {
    partname.membername()
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
