//! Operation-local accounting for the bounded OPC ZIP paths.
//!
//! This report deliberately mirrors the covered low-level ZIP counters without
//! exposing the `soapberry_zip` report type.  A report belongs to the caller;
//! it is never retained by an archive, cache, flight, or publication plan.

use crate::error::OpcError;

/// Checked counters for one OPC part read or sequential source publication.
///
/// The nine ZIP counters retain the low-level payload provenance described by
/// [`soapberry_zip::ZipOperationAccounting`].  `output_bytes_accepted` is an
/// OPC publication counter: it includes every byte accepted by the caller's
/// sink, including ZIP framing and generated records.  On a partial operation
/// all counters retain the work accepted or observed before the error.
///
/// This first propagation slice covers cold single-Part reads, exact source
/// publication, and the singular one-Part overlay publisher.  Ordinary eager
/// package reads, parallel or bulk operations, topology publishers, batch
/// overlays, `PartWriter`, and performance-report schemas intentionally remain
/// outside this report until their operation boundaries are reviewed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OpcOperationAccounting {
    compressed_deflate_payload_bytes_read: u64,
    stored_payload_bytes_read: u64,
    stored_payload_bytes_accepted: u64,
    deflate_bytes_produced: u64,
    deflate_bytes_accepted: u64,
    generated_deflate_payload_bytes_emitted: u64,
    stored_payload_bytes_emitted: u64,
    precompressed_payload_bytes_emitted: u64,
    raw_unchanged_source_bytes_accepted: u64,
    output_bytes_accepted: u64,
}

impl OpcOperationAccounting {
    /// Creates an empty report.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            compressed_deflate_payload_bytes_read: 0,
            stored_payload_bytes_read: 0,
            stored_payload_bytes_accepted: 0,
            deflate_bytes_produced: 0,
            deflate_bytes_accepted: 0,
            generated_deflate_payload_bytes_emitted: 0,
            stored_payload_bytes_emitted: 0,
            precompressed_payload_bytes_emitted: 0,
            raw_unchanged_source_bytes_accepted: 0,
            output_bytes_accepted: 0,
        }
    }

    /// Actual compressed Deflate payload bytes read from a cold ZIP member.
    #[must_use]
    pub const fn compressed_deflate_payload_bytes_read(&self) -> u64 {
        self.compressed_deflate_payload_bytes_read
    }

    /// Actual stored payload bytes read from a cold ZIP member.
    #[must_use]
    pub const fn stored_payload_bytes_read(&self) -> u64 {
        self.stored_payload_bytes_read
    }

    /// Stored payload bytes accepted by a Part destination.
    #[must_use]
    pub const fn stored_payload_bytes_accepted(&self) -> u64 {
        self.stored_payload_bytes_accepted
    }

    /// Bytes produced by a Deflate decoder.
    #[must_use]
    pub const fn deflate_bytes_produced(&self) -> u64 {
        self.deflate_bytes_produced
    }

    /// Decoded Deflate bytes accepted by a Part destination.
    #[must_use]
    pub const fn deflate_bytes_accepted(&self) -> u64 {
        self.deflate_bytes_accepted
    }

    /// Generated Deflate member-payload bytes emitted during preservation.
    #[must_use]
    pub const fn generated_deflate_payload_bytes_emitted(&self) -> u64 {
        self.generated_deflate_payload_bytes_emitted
    }

    /// Stored member-payload bytes emitted during preservation.
    #[must_use]
    pub const fn stored_payload_bytes_emitted(&self) -> u64 {
        self.stored_payload_bytes_emitted
    }

    /// Caller-provided precompressed member-payload bytes emitted.
    #[must_use]
    pub const fn precompressed_payload_bytes_emitted(&self) -> u64 {
        self.precompressed_payload_bytes_emitted
    }

    /// Unchanged source archive bytes accepted by preservation publication.
    ///
    /// This is archive-byte accounting, not payload accounting.  It may
    /// include unchanged local spans, central records, end records, and the
    /// archive comment, depending on the preservation plan.
    #[must_use]
    pub const fn raw_unchanged_source_bytes_accepted(&self) -> u64 {
        self.raw_unchanged_source_bytes_accepted
    }

    /// Total bytes accepted by the caller's sequential output sink.
    ///
    /// Unlike the payload counters, this includes ZIP framing and generated
    /// records.  It is checked independently of the raw/payload counters.
    #[must_use]
    pub const fn output_bytes_accepted(&self) -> u64 {
        self.output_bytes_accepted
    }

    /// Alias for [`Self::compressed_deflate_payload_bytes_read`].
    #[must_use]
    pub const fn deflate_payload_bytes_read(&self) -> u64 {
        self.compressed_deflate_payload_bytes_read()
    }

    /// Alias for [`Self::deflate_bytes_produced`].
    #[must_use]
    pub const fn decompressed_deflate_bytes_produced(&self) -> u64 {
        self.deflate_bytes_produced()
    }

    /// Alias for [`Self::deflate_bytes_accepted`].
    #[must_use]
    pub const fn decompressed_deflate_bytes_accepted(&self) -> u64 {
        self.deflate_bytes_accepted()
    }

    pub(crate) fn add_compressed_deflate_payload_bytes_read(
        &mut self,
        bytes: u64,
    ) -> Result<(), OpcError> {
        checked_add(
            &mut self.compressed_deflate_payload_bytes_read,
            bytes,
            "compressed Deflate payload bytes read",
        )
    }

    pub(crate) fn add_stored_payload_bytes_read(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.stored_payload_bytes_read,
            bytes,
            "stored payload bytes read",
        )
    }

    pub(crate) fn add_stored_payload_bytes_accepted(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.stored_payload_bytes_accepted,
            bytes,
            "stored payload bytes accepted",
        )
    }

    pub(crate) fn add_deflate_bytes_produced(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.deflate_bytes_produced,
            bytes,
            "decompressed Deflate bytes produced",
        )
    }

    pub(crate) fn add_deflate_bytes_accepted(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.deflate_bytes_accepted,
            bytes,
            "decompressed Deflate bytes accepted",
        )
    }

    pub(crate) fn add_generated_deflate_payload_bytes_emitted(
        &mut self,
        bytes: u64,
    ) -> Result<(), OpcError> {
        checked_add(
            &mut self.generated_deflate_payload_bytes_emitted,
            bytes,
            "generated Deflate payload bytes emitted",
        )
    }

    pub(crate) fn add_stored_payload_bytes_emitted(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.stored_payload_bytes_emitted,
            bytes,
            "stored payload bytes emitted",
        )
    }

    pub(crate) fn add_precompressed_payload_bytes_emitted(
        &mut self,
        bytes: u64,
    ) -> Result<(), OpcError> {
        checked_add(
            &mut self.precompressed_payload_bytes_emitted,
            bytes,
            "precompressed payload bytes emitted",
        )
    }

    pub(crate) fn add_raw_unchanged_source_bytes_accepted(
        &mut self,
        bytes: u64,
    ) -> Result<(), OpcError> {
        checked_add(
            &mut self.raw_unchanged_source_bytes_accepted,
            bytes,
            "raw unchanged source bytes accepted",
        )
    }

    pub(crate) fn add_output_bytes_accepted(&mut self, bytes: u64) -> Result<(), OpcError> {
        checked_add(
            &mut self.output_bytes_accepted,
            bytes,
            "OPC output bytes accepted",
        )
    }

    /// Merge one local ZIP report without retaining its low-level type.
    ///
    /// All nine fields are visited in a fixed order.  If one destination
    /// counter is already saturated, representable later fields are still
    /// merged and the first checked overflow is returned afterward.
    pub(crate) fn merge_zip(
        &mut self,
        source: &soapberry_zip::ZipOperationAccounting,
    ) -> Result<(), OpcError> {
        let mut first_error = None;
        macro_rules! merge {
            ($method:ident, $value:expr) => {
                if let Err(error) = self.$method($value) {
                    if first_error.is_none() {
                        first_error = Some(error);
                    }
                }
            };
        }
        merge!(
            add_compressed_deflate_payload_bytes_read,
            source.compressed_deflate_payload_bytes_read()
        );
        merge!(
            add_stored_payload_bytes_read,
            source.stored_payload_bytes_read()
        );
        merge!(
            add_stored_payload_bytes_accepted,
            source.stored_payload_bytes_accepted()
        );
        merge!(add_deflate_bytes_produced, source.deflate_bytes_produced());
        merge!(add_deflate_bytes_accepted, source.deflate_bytes_accepted());
        merge!(
            add_generated_deflate_payload_bytes_emitted,
            source.generated_deflate_payload_bytes_emitted()
        );
        merge!(
            add_stored_payload_bytes_emitted,
            source.stored_payload_bytes_emitted()
        );
        merge!(
            add_precompressed_payload_bytes_emitted,
            source.precompressed_payload_bytes_emitted()
        );
        merge!(
            add_raw_unchanged_source_bytes_accepted,
            source.raw_unchanged_source_bytes_accepted()
        );
        first_error.map_or(Ok(()), Err)
    }
}

fn checked_add(counter: &mut u64, bytes: u64, counter_name: &'static str) -> Result<(), OpcError> {
    *counter = counter
        .checked_add(bytes)
        .ok_or(OpcError::OperationAccountingOverflow {
            counter: counter_name,
        })?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn merge_copies_all_low_level_counters_without_retaining_the_source_type() {
        let mut report = OpcOperationAccounting::default();
        let source = soapberry_zip::ZipOperationAccounting::default();
        assert!(report.merge_zip(&source).is_ok());
        assert_eq!(report, OpcOperationAccounting::default());
    }

    #[test]
    fn output_counter_overflow_is_typed_and_checked() {
        let mut report = OpcOperationAccounting {
            output_bytes_accepted: u64::MAX,
            ..OpcOperationAccounting::default()
        };
        let error = report.add_output_bytes_accepted(1).unwrap_err();
        assert!(matches!(
            error,
            OpcError::OperationAccountingOverflow {
                counter: "OPC output bytes accepted"
            }
        ));
    }
}
