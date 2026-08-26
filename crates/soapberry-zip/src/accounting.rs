//! Operation-local ZIP payload accounting.

use std::io::{self, Read, Write};

use crate::{Error, ErrorKind};

/// Monotonic counters for one ZIP read, write, or preservation operation.
///
/// The value is owned by the caller and is never stored in an archive, reader,
/// writer, cache, or global runtime state.  Payload counters exclude ZIP local
/// headers, data descriptors, central-directory records, EOCD records, and
/// other framing bytes, except that preservation's raw-source counter
/// intentionally includes unchanged source framing bytes. Deflate production
/// counts bytes returned by the decompressor; Deflate acceptance counts bytes
/// accepted by the destination (including a partial destination prefix on an
/// error). This first surface covers single-entry/convenience readers and
/// writers plus preservation; incremental `start_entry` and explicit parallel
/// bulk accounting are deferred.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ZipOperationAccounting {
    compressed_deflate_payload_bytes_read: u64,
    stored_payload_bytes_read: u64,
    stored_payload_bytes_accepted: u64,
    deflate_bytes_produced: u64,
    deflate_bytes_accepted: u64,
    generated_deflate_payload_bytes_emitted: u64,
    stored_payload_bytes_emitted: u64,
    precompressed_payload_bytes_emitted: u64,
    raw_unchanged_source_bytes_accepted: u64,
}

impl ZipOperationAccounting {
    /// Creates an empty accounting value.
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
        }
    }

    /// Actual compressed Deflate payload bytes returned by the source.
    #[must_use]
    pub const fn compressed_deflate_payload_bytes_read(&self) -> u64 {
        self.compressed_deflate_payload_bytes_read
    }

    /// Actual stored payload bytes returned by the source.
    #[must_use]
    pub const fn stored_payload_bytes_read(&self) -> u64 {
        self.stored_payload_bytes_read
    }

    /// Stored payload bytes accepted by a destination sink, including a
    /// partial prefix accepted before a sink, checksum, or size error.
    #[must_use]
    pub const fn stored_payload_bytes_accepted(&self) -> u64 {
        self.stored_payload_bytes_accepted
    }

    /// Bytes returned by a Deflate decoder, including bytes later rejected by
    /// checksum or size validation.
    #[must_use]
    pub const fn deflate_bytes_produced(&self) -> u64 {
        self.deflate_bytes_produced
    }

    /// Deflate bytes accepted by the destination, including a partial prefix
    /// accepted before a sink, checksum, or size error.
    #[must_use]
    pub const fn deflate_bytes_accepted(&self) -> u64 {
        self.deflate_bytes_accepted
    }

    /// Generated Deflate payload bytes accepted by an archive sink.
    #[must_use]
    pub const fn generated_deflate_payload_bytes_emitted(&self) -> u64 {
        self.generated_deflate_payload_bytes_emitted
    }

    /// Stored payload bytes accepted by an archive sink.
    #[must_use]
    pub const fn stored_payload_bytes_emitted(&self) -> u64 {
        self.stored_payload_bytes_emitted
    }

    /// Already-compressed payload bytes accepted by an archive sink.
    #[must_use]
    pub const fn precompressed_payload_bytes_emitted(&self) -> u64 {
        self.precompressed_payload_bytes_emitted
    }

    /// Unchanged source bytes accepted by a preservation publication sink.
    ///
    /// This intentionally includes unchanged local-member spans, unchanged
    /// central-record bytes outside regenerated local offsets, and the source
    /// archive comment. It is not a payload-only counter.
    #[must_use]
    pub const fn raw_unchanged_source_bytes_accepted(&self) -> u64 {
        self.raw_unchanged_source_bytes_accepted
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
    ) -> Result<(), Error> {
        checked_add(
            &mut self.compressed_deflate_payload_bytes_read,
            bytes,
            "compressed Deflate payload bytes read",
        )
    }

    pub(crate) fn add_stored_payload_bytes_read(&mut self, bytes: u64) -> Result<(), Error> {
        checked_add(
            &mut self.stored_payload_bytes_read,
            bytes,
            "stored payload bytes read",
        )
    }

    pub(crate) fn add_stored_payload_bytes_accepted(&mut self, bytes: u64) -> Result<(), Error> {
        checked_add(
            &mut self.stored_payload_bytes_accepted,
            bytes,
            "stored payload bytes accepted",
        )
    }

    pub(crate) fn add_deflate_bytes_produced(&mut self, bytes: u64) -> Result<(), Error> {
        checked_add(
            &mut self.deflate_bytes_produced,
            bytes,
            "decompressed Deflate bytes produced",
        )
    }

    pub(crate) fn add_deflate_bytes_accepted(&mut self, bytes: u64) -> Result<(), Error> {
        checked_add(
            &mut self.deflate_bytes_accepted,
            bytes,
            "decompressed Deflate bytes accepted",
        )
    }

    pub(crate) fn add_generated_deflate_payload_bytes_emitted(
        &mut self,
        bytes: u64,
    ) -> Result<(), Error> {
        checked_add(
            &mut self.generated_deflate_payload_bytes_emitted,
            bytes,
            "generated Deflate payload bytes emitted",
        )
    }

    pub(crate) fn add_stored_payload_bytes_emitted(&mut self, bytes: u64) -> Result<(), Error> {
        checked_add(
            &mut self.stored_payload_bytes_emitted,
            bytes,
            "stored payload bytes emitted",
        )
    }

    pub(crate) fn add_precompressed_payload_bytes_emitted(
        &mut self,
        bytes: u64,
    ) -> Result<(), Error> {
        checked_add(
            &mut self.precompressed_payload_bytes_emitted,
            bytes,
            "precompressed payload bytes emitted",
        )
    }

    pub(crate) fn add_raw_unchanged_source_bytes_accepted(
        &mut self,
        bytes: u64,
    ) -> Result<(), Error> {
        checked_add(
            &mut self.raw_unchanged_source_bytes_accepted,
            bytes,
            "raw unchanged source bytes accepted",
        )
    }
}

fn checked_add(counter: &mut u64, bytes: u64, resource: &'static str) -> Result<(), Error> {
    *counter = counter.checked_add(bytes).ok_or_else(|| {
        Error::from(ErrorKind::InvalidInput {
            msg: format!("ZIP operation accounting overflow: {resource}"),
        })
    })?;
    Ok(())
}

pub(crate) fn usize_to_u64(value: usize, resource: &'static str) -> Result<u64, Error> {
    u64::try_from(value).map_err(|_| accounting_overflow(resource))
}

/// The logical payload kind accepted by a read destination.
#[derive(Debug, Clone, Copy)]
pub(crate) enum AccountingReadKind {
    Stored,
    Deflate,
}

/// The payload class charged by [`write_all_counted`].
#[derive(Debug, Clone, Copy)]
pub(crate) enum AccountingWriteKind {
    GeneratedDeflate,
    Stored,
    Precompressed,
    RawUnchangedSource,
}

impl AccountingWriteKind {
    pub(crate) fn add(
        self,
        accounting: &mut ZipOperationAccounting,
        bytes: u64,
    ) -> Result<(), Error> {
        match self {
            Self::GeneratedDeflate => accounting.add_generated_deflate_payload_bytes_emitted(bytes),
            Self::Stored => accounting.add_stored_payload_bytes_emitted(bytes),
            Self::Precompressed => accounting.add_precompressed_payload_bytes_emitted(bytes),
            Self::RawUnchangedSource => accounting.add_raw_unchanged_source_bytes_accepted(bytes),
        }
    }
}

/// Write all bytes while charging only bytes actually accepted by the sink.
pub(crate) fn write_all_counted<W: Write>(
    sink: &mut W,
    mut buffer: &[u8],
    accounting: &mut ZipOperationAccounting,
    kind: AccountingWriteKind,
) -> Result<(), Error> {
    while !buffer.is_empty() {
        let written = match sink.write(buffer) {
            Ok(written) => written,
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            Err(error) => return Err(error.into()),
        };
        if written == 0 {
            return Err(Error::from(io::Error::new(
                io::ErrorKind::WriteZero,
                "failed to write ZIP output",
            )));
        }
        if written > buffer.len() {
            return Err(ErrorKind::InvalidInput {
                msg: "ZIP output sink returned more bytes than requested".to_string(),
            }
            .into());
        }
        kind.add(
            accounting,
            usize_to_u64(written, "ZIP output bytes accepted")?,
        )?;
        buffer = &buffer[written..];
    }
    Ok(())
}

/// A source reader that records the number of bytes actually returned.
#[derive(Debug)]
pub(crate) struct CountingReader<R> {
    inner: R,
    count: u64,
}

impl<R> CountingReader<R> {
    pub(crate) fn new(inner: R) -> Self {
        Self { inner, count: 0 }
    }

    pub(crate) const fn count(&self) -> u64 {
        self.count
    }
}

impl<R: Read> Read for CountingReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let read = self.inner.read(buffer)?;
        if read > buffer.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP source returned more bytes than requested",
            ));
        }
        self.count = self
            .count
            .checked_add(u64::try_from(read).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "ZIP source byte count overflow")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "ZIP source byte count overflow")
            })?;
        Ok(read)
    }
}

pub(crate) fn accounting_overflow(resource: &'static str) -> Error {
    ErrorKind::InvalidInput {
        msg: format!("ZIP operation accounting overflow: {resource}"),
    }
    .into()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Debug)]
    struct ShortWriter {
        bytes: Vec<u8>,
        maximum: usize,
    }

    impl Write for ShortWriter {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            let written = buffer.len().min(self.maximum);
            self.bytes.extend_from_slice(&buffer[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn counted_writes_charge_only_accepted_bytes() {
        let mut sink = ShortWriter {
            bytes: Vec::new(),
            maximum: 2,
        };
        let mut accounting = ZipOperationAccounting::default();
        write_all_counted(
            &mut sink,
            b"counted payload",
            &mut accounting,
            AccountingWriteKind::RawUnchangedSource,
        )
        .unwrap();
        assert_eq!(sink.bytes, b"counted payload");
        assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), 15);
    }

    #[test]
    fn every_accounting_counter_add_reports_typed_overflow() {
        let adders: [(
            &str,
            fn(&mut ZipOperationAccounting, u64) -> Result<(), Error>,
        ); 9] = [
            (
                "compressed Deflate payload bytes read",
                ZipOperationAccounting::add_compressed_deflate_payload_bytes_read,
            ),
            (
                "stored payload bytes read",
                ZipOperationAccounting::add_stored_payload_bytes_read,
            ),
            (
                "stored payload bytes accepted",
                ZipOperationAccounting::add_stored_payload_bytes_accepted,
            ),
            (
                "decompressed Deflate bytes produced",
                ZipOperationAccounting::add_deflate_bytes_produced,
            ),
            (
                "decompressed Deflate bytes accepted",
                ZipOperationAccounting::add_deflate_bytes_accepted,
            ),
            (
                "generated Deflate payload bytes emitted",
                ZipOperationAccounting::add_generated_deflate_payload_bytes_emitted,
            ),
            (
                "stored payload bytes emitted",
                ZipOperationAccounting::add_stored_payload_bytes_emitted,
            ),
            (
                "precompressed payload bytes emitted",
                ZipOperationAccounting::add_precompressed_payload_bytes_emitted,
            ),
            (
                "raw unchanged source bytes accepted",
                ZipOperationAccounting::add_raw_unchanged_source_bytes_accepted,
            ),
        ];

        for (resource, add) in adders {
            let mut accounting = ZipOperationAccounting::default();
            add(&mut accounting, u64::MAX).unwrap();
            let error = add(&mut accounting, 1).unwrap_err();
            assert!(matches!(
                error.kind(),
                ErrorKind::InvalidInput { msg }
                    if msg.contains("accounting overflow") && msg.contains(resource)
            ));
        }
    }
}
