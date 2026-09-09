//! Operation-scoped payload reads for a source-backed OPC package.
//!
//! The session owns only the unmanaged ZIP decoder state.  Cache admission,
//! source freshness, execution checks, reservations, and publication remain
//! in [`SourceBackedPackage`]'s ordinary Part-read path.

use super::{PartData, PartView, SourceBackedPackage, SourceReader};
use crate::OpcOperationAccounting;
use crate::error::{OpcError, Result};
use soapberry_zip::office::IndexedReadSession;

/// Reusable, package-bound payload reader for one short OPC operation.
///
/// A session does not perform I/O when it is created.  On an unmanaged package
/// the first cold Deflate read lazily creates the ZIP decoder and later reads
/// may reuse it.  Managed packages deliberately keep the existing one-shot
/// decoder path: the existing cache and execution policy account for each
/// temporary decoded allocation, and this session retains no decoder or
/// scratch allocation across managed reads.
///
/// Every read verifies exact package identity before inspecting the supplied
/// [`PartView`].  This prevents a view from another package that happens to
/// use the same source or part index from crossing the session boundary.
#[must_use = "a PartReadSession should be used for the intended package operation"]
pub struct PartReadSession<'package> {
    package: &'package SourceBackedPackage,
    decoder: Option<IndexedReadSession<'package, SourceReader>>,
}

impl SourceBackedPackage {
    /// Begin an operation-scoped payload-read session.
    ///
    /// Construction is metadata-only: it performs no source I/O and does not
    /// allocate a ZIP decoder.  Use [`PartReadSession::read`] or
    /// [`PartReadSession::read_with_accounting`] for payloads from this exact
    /// package.
    pub fn read_session(&self) -> PartReadSession<'_> {
        PartReadSession {
            package: self,
            decoder: None,
        }
    }
}

impl<'package> PartReadSession<'package> {
    /// Read one part through the package cache and source-integrity checks.
    pub fn read(&mut self, part: PartView<'_>) -> Result<PartData> {
        self.read_inner(part, None)
    }

    /// Read one part and merge cold ZIP work into a caller-owned report.
    ///
    /// Cache hits and same-part waiters preserve the ordinary zero-work
    /// accounting behavior.  A failed cold read leaves the report with only
    /// the work accepted before that failure, just like the direct Part API.
    pub fn read_with_accounting(
        &mut self,
        part: PartView<'_>,
        accounting: &mut OpcOperationAccounting,
    ) -> Result<PartData> {
        self.read_inner(part, Some(accounting))
    }

    fn read_inner(
        &mut self,
        part: PartView<'_>,
        accounting: Option<&mut OpcOperationAccounting>,
    ) -> Result<PartData> {
        // Keep this check before touching `part.index`: PartView is a compact
        // metadata handle, and an index from another package is not meaningful
        // even when both packages were opened over equal source bytes.
        if !std::ptr::eq(part.package, self.package) {
            return Err(OpcError::ForeignPartView);
        }
        let index = part.index;

        if self.package.cache.is_managed() {
            // Managed reads intentionally retain the package's one-shot path;
            // it performs the declared-size reservation and releases temporary
            // decoder state at the end of this read.
            self.package.read_part_with_accounting(index, accounting)
        } else {
            let decoder = self
                .decoder
                .get_or_insert_with(|| self.package.archive.read_session());
            self.package
                .read_part_with_session(index, decoder, accounting)
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these tests use direct fixture assertions"
)]
mod tests {
    use super::{OpcError, SourceBackedPackage};
    use crate::{OpcOperationAccounting, PackURI, ReadLimits};
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
        ReadAt, Resource, SourceVersion,
    };
    use std::io;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    const DOCUMENT: &[u8] = b"<document>read-session payload</document>";
    const SECOND: &[u8] = b"<second>read-session payload</second>";
    const ORPHAN: &[u8] = b"stored orphan payload";

    fn root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
    }

    fn mixed_archive() -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer
            .write_deflated("word/document.xml", DOCUMENT)
            .unwrap();
        writer.write_deflated("custom/second.xml", SECOND).unwrap();
        writer.write_stored("custom/orphan.xml", ORPHAN).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn ordinary_archive(document: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[derive(Debug)]
    struct CountingSource {
        bytes: Vec<u8>,
        reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let offset = usize::try_from(offset).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidInput, "offset overflows usize")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x5245_4144,
                self.revision.load(Ordering::SeqCst),
            ))
        }
    }

    fn document_uri() -> PackURI {
        PackURI::new("/word/document.xml").unwrap()
    }

    fn second_uri() -> PackURI {
        PackURI::new("/custom/second.xml").unwrap()
    }

    fn orphan_uri() -> PackURI {
        PackURI::new("/custom/orphan.xml").unwrap()
    }

    fn central_record_start(bytes: &[u8], name: &[u8]) -> (usize, usize) {
        let signature = [0x50, 0x4b, 0x01, 0x02];
        bytes
            .windows(signature.len())
            .enumerate()
            .find_map(|(offset, window)| {
                if window != signature || offset.checked_add(46)? > bytes.len() {
                    return None;
                }
                let name_len =
                    u16::from_le_bytes(bytes[offset + 28..offset + 30].try_into().ok()?) as usize;
                let extra_len =
                    u16::from_le_bytes(bytes[offset + 30..offset + 32].try_into().ok()?) as usize;
                let comment_len =
                    u16::from_le_bytes(bytes[offset + 32..offset + 34].try_into().ok()?) as usize;
                let name_start = offset + 46;
                let name_end = name_start.checked_add(name_len)?;
                let record_end = name_end.checked_add(extra_len)?.checked_add(comment_len)?;
                if record_end <= bytes.len() && &bytes[name_start..name_end] == name {
                    let local_offset =
                        u32::from_le_bytes(bytes[offset + 42..offset + 46].try_into().ok()?);
                    Some((offset, usize::try_from(local_offset).ok()?))
                } else {
                    None
                }
            })
            .expect("central directory entry is present")
    }

    fn mutate_crc(mut bytes: Vec<u8>, name: &[u8]) -> Vec<u8> {
        let (record_start, local_start) = central_record_start(&bytes, name);
        let old = u32::from_le_bytes(
            bytes[record_start + 16..record_start + 20]
                .try_into()
                .unwrap(),
        );
        let wrong = old ^ 1;
        bytes[record_start + 16..record_start + 20].copy_from_slice(&wrong.to_le_bytes());
        bytes[local_start + 14..local_start + 18].copy_from_slice(&wrong.to_le_bytes());
        bytes
    }

    fn mutate_size(mut bytes: Vec<u8>, name: &[u8]) -> Vec<u8> {
        let (record_start, local_start) = central_record_start(&bytes, name);
        let old = u32::from_le_bytes(
            bytes[record_start + 24..record_start + 28]
                .try_into()
                .unwrap(),
        );
        let wrong = old.checked_add(1).unwrap();
        bytes[record_start + 24..record_start + 28].copy_from_slice(&wrong.to_le_bytes());
        bytes[local_start + 22..local_start + 26].copy_from_slice(&wrong.to_le_bytes());
        bytes
    }

    fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "opc-read-session-test",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    #[test]
    fn foreign_part_view_is_rejected_before_source_work() {
        let source = Arc::new(CountingSource::new(mixed_archive()));
        let first = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let second = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let foreign = first.part(&document_uri()).unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);

        let mut session = second.read_session();
        let mut accounting = OpcOperationAccounting::default();
        assert!(matches!(
            session.read_with_accounting(foreign, &mut accounting),
            Err(OpcError::ForeignPartView)
        ));
        assert_eq!(accounting, OpcOperationAccounting::default());
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    }

    #[test]
    fn unmanaged_session_reads_mixed_compression_and_cache_hits_are_free() {
        let package = SourceBackedPackage::from_vec(mixed_archive()).unwrap();
        let deflated = package.part(&document_uri()).unwrap();
        let second = package.part(&second_uri()).unwrap();
        let stored = package.part(&orphan_uri()).unwrap();
        let mut session = package.read_session();
        let mut accounting = OpcOperationAccounting::default();

        assert_eq!(
            session
                .read_with_accounting(deflated, &mut accounting)
                .unwrap()
                .as_bytes(),
            DOCUMENT
        );
        assert_eq!(
            session
                .read_with_accounting(second, &mut accounting)
                .unwrap()
                .as_bytes(),
            SECOND
        );
        assert_eq!(
            session
                .read_with_accounting(stored, &mut accounting)
                .unwrap()
                .as_bytes(),
            ORPHAN
        );
        assert!(accounting.compressed_deflate_payload_bytes_read() > 0);
        assert_eq!(
            accounting.deflate_bytes_produced(),
            (DOCUMENT.len() + SECOND.len()) as u64
        );
        assert_eq!(accounting.stored_payload_bytes_read(), ORPHAN.len() as u64);

        let mut hit_accounting = OpcOperationAccounting::default();
        assert_eq!(
            session
                .read_with_accounting(deflated, &mut hit_accounting)
                .unwrap()
                .as_bytes(),
            DOCUMENT
        );
        assert_eq!(hit_accounting, OpcOperationAccounting::default());
    }

    #[test]
    fn malformed_crc_and_size_failures_can_retry_without_publication() {
        let crc_bytes = mutate_crc(mixed_archive(), b"word/document.xml");
        let size_bytes = mutate_size(mixed_archive(), b"word/document.xml");

        for bytes in [crc_bytes, size_bytes] {
            let source = Arc::new(CountingSource::new(bytes));
            let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
            let bad = package.part(&document_uri()).unwrap();
            let stored = package.part(&orphan_uri()).unwrap();
            let valid_deflate = package.part(&second_uri()).unwrap();
            let mut session = package.read_session();
            assert!(matches!(session.read(bad), Err(OpcError::ZipError(_))));
            let reads_after_first = source.reads.load(Ordering::SeqCst);
            assert_eq!(session.read(stored).unwrap().as_bytes(), ORPHAN);
            assert_eq!(session.read(valid_deflate).unwrap().as_bytes(), SECOND);
            assert!(matches!(session.read(bad), Err(OpcError::ZipError(_))));
            assert!(source.reads.load(Ordering::SeqCst) > reads_after_first);
            let diagnostics = package.cache_diagnostics();
            assert_eq!(diagnostics.retained_entries, 2);
            assert_eq!(diagnostics.failed_loads, 2);
        }
    }

    #[test]
    fn stale_source_is_refused_and_session_can_be_reused_after_freshness_restored() {
        let source = Arc::new(CountingSource::new(ordinary_archive(DOCUMENT)));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let part = package.part(&document_uri()).unwrap();
        let mut session = package.read_session();

        source.revision.store(1, Ordering::SeqCst);
        assert!(matches!(
            session.read(part),
            Err(OpcError::SourceChanged { .. })
        ));
        source.revision.store(0, Ordering::SeqCst);
        assert_eq!(session.read(part).unwrap().as_bytes(), DOCUMENT);
    }

    #[test]
    fn managed_session_cancellation_and_payload_handles_release_budget() {
        let (budget, cancellation_source, context) = managed_context(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(ordinary_archive(DOCUMENT))),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let part = package.part(&document_uri()).unwrap();
        let mut session = package.read_session();
        let data = session.read(part).unwrap();
        assert_eq!(data.as_bytes(), DOCUMENT);
        assert!(budget.used(Resource::Memory) >= DOCUMENT.len() as u64);

        cancellation_source.cancel();
        assert!(matches!(session.read(part), Err(OpcError::Cancelled)));
        drop(data);
        drop(session);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_deflate_quota_refusal_is_preflighted_and_keeps_session_cold() {
        let source = Arc::new(CountingSource::new(mixed_archive()));
        let (budget, _cancellation_source, context) = managed_context((DOCUMENT.len() - 1) as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let part = package.part(&document_uri()).unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let mut session = package.read_session();
        assert!(session.decoder.is_none());

        let error = session.read(part).unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        assert!(session.decoder.is_none());
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Work), 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
    }
}
