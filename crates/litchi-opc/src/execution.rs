//! Explicit, reusable execution adapter for eager OPC package opens.
//!
//! [`OpenSession`] adapts the runtime-neutral policy in
//! [`litchi_core::ExecutionContext`] to soapberry-zip's local bounded read
//! session. It is intentionally separate from ordinary [`crate::OpcPackage`]
//! constructors: those constructors remain synchronous and serial.

use crate::{OpcError, OpcPackage, ReadLimits, Result};
use litchi_core::{AffinityPolicy, ExecutionContext, ExecutionError, Resource};
use soapberry_zip::office::{
    CancellationProbe, LazyArchiveReader, ParallelAffinity, ParallelReadLimits, ParallelReadSession,
};

/// Reusable, explicitly scheduled eager OPC open adapter.
///
/// The session owns a local ZIP worker pool, never uses a global Rayon pool,
/// and retains the caller-selected [`ExecutionContext`] for every open. It is
/// an advanced ingress API; ordinary package CRUD never exposes this type.
#[derive(Debug)]
pub struct OpenSession {
    context: ExecutionContext,
    zip: ParallelReadSession,
}

impl OpenSession {
    /// Creates a reusable OPC open session from an explicit execution context.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::Cancelled`] when the context is already cancelled,
    /// [`OpcError::Execution`] for context failures, or
    /// [`OpcError::ParallelRead`] if the local ZIP session cannot be created.
    pub fn new(context: ExecutionContext) -> Result<Self> {
        context.check().map_err(map_execution_error)?;
        let limits = context.limits();
        let affinity = match limits.affinity() {
            AffinityPolicy::Inherit => ParallelAffinity::Inherit,
            _ => return Err(OpcError::UnsupportedExecutionAffinity),
        };
        let zip_limits = ParallelReadLimits::with_affinity(
            limits.workers(),
            limits.max_in_flight_tasks(),
            limits.max_in_flight_bytes(),
            limits.min_parallel_bytes(),
            affinity,
        )
        .map_err(OpcError::ParallelRead)?;
        let zip = ParallelReadSession::new(zip_limits).map_err(OpcError::ParallelRead)?;
        Ok(Self { context, zip })
    }

    /// Execution context retained by this session.
    #[must_use]
    pub const fn context(&self) -> &ExecutionContext {
        &self.context
    }

    /// Opens a borrowed archive under explicit OPC and execution limits.
    ///
    /// # Errors
    ///
    /// Returns a typed OPC, execution, or local-session error when opening
    /// cannot complete. Cancellation discards the incomplete package.
    pub fn from_bytes(&self, data: &[u8], limits: ReadLimits) -> Result<OpcPackage> {
        OpcPackage::from_bytes_with_open_session(data, limits, self)
    }

    /// Opens an owned archive under explicit OPC and execution limits.
    ///
    /// The successful package retains authorization for exact owned-source
    /// no-op publication, exactly like [`OpcPackage::from_vec_with_limits`].
    ///
    /// # Errors
    ///
    /// Returns a typed OPC, execution, or local-session error when opening
    /// cannot complete. Cancellation discards the incomplete package.
    pub fn from_vec(&self, data: Vec<u8>, limits: ReadLimits) -> Result<OpcPackage> {
        OpcPackage::from_vec_with_open_session(data, limits, self)
    }

    pub(crate) fn check(&self) -> Result<()> {
        self.context.check().map_err(map_execution_error)?;
        Ok(())
    }

    pub(crate) fn charge_input(&self, bytes: u64) -> Result<()> {
        self.context
            .consume(Resource::InputBytes, bytes)
            .map_err(map_execution_error)
    }

    pub(crate) fn read_many<'name>(
        &self,
        archive: &LazyArchiveReader<'_>,
        names: &'name [&'name str],
    ) -> Result<
        Vec<(
            &'name str,
            std::result::Result<Vec<u8>, soapberry_zip::Error>,
        )>,
    > {
        self.context.check().map_err(map_execution_error)?;

        let declared_bytes = names.iter().try_fold(0_u64, |total, name| {
            archive
                .metadata(name)
                .map_err(OpcError::from)
                .and_then(|metadata| {
                    total
                        .checked_add(metadata.uncompressed_size())
                        .ok_or(OpcError::ReadLimit {
                            resource: crate::ReadResource::TotalPartBytes,
                            actual: u64::MAX,
                            maximum: u64::MAX,
                        })
                })
        })?;
        let in_flight = declared_bytes.min(self.context.limits().max_in_flight_bytes().get());
        let reservation = self
            .context
            .reserve(Resource::Memory, in_flight)
            .map_err(map_execution_error)?;
        self.context
            .consume(Resource::Work, declared_bytes)
            .map_err(map_execution_error)?;

        let cancellation = ContextCancellation(&self.context);
        let results = archive
            .read_many_with_session(&self.zip, names, &cancellation)
            .map_err(map_parallel_read_error);
        drop(reservation);
        results
    }
}

struct ContextCancellation<'a>(&'a ExecutionContext);

impl CancellationProbe for ContextCancellation<'_> {
    fn is_cancelled(&self) -> bool {
        self.0.cancellation().is_cancelled()
    }
}

fn map_execution_error(error: ExecutionError) -> OpcError {
    match error {
        ExecutionError::Cancelled => OpcError::Cancelled,
        error => OpcError::Execution(error),
    }
}

fn map_parallel_read_error(error: soapberry_zip::Error) -> OpcError {
    match error.kind() {
        soapberry_zip::ErrorKind::Cancelled => OpcError::Cancelled,
        _ => OpcError::ParallelRead(error),
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::{BlobPart, PackURI, PackageWriter};
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
        Resource,
    };
    use std::num::{NonZeroU64, NonZeroUsize};

    const MIB: u64 = 1024 * 1024;

    fn execution_context(memory: u64) -> (CancellationSource, ExecutionContext) {
        let (source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(4 * 1024).unwrap(),
            0,
        )
        .unwrap();
        let budget = Budget::root(
            "opc-open-test",
            Limits::new(memory, 16 * MIB, 16 * MIB, 10_000, 64, 16 * MIB),
        );
        (source, ExecutionContext::new(budget, cancellation, limits))
    }

    fn archive() -> Vec<u8> {
        let mut package = OpcPackage::new();
        for index in 0_u8..8 {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new(format!("/benchmark/{index}.bin")).unwrap(),
                "application/octet-stream".to_owned(),
                vec![index; 1024],
            )));
        }
        PackageWriter::to_bytes(&package).unwrap()
    }

    fn cache_isolation_archive() -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_deflated("first.xml", b"first payload")
            .unwrap();
        writer
            .write_stored("second.xml", b"second payload")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn package_parts(package: &OpcPackage) -> Vec<(String, Vec<u8>)> {
        let mut parts = package
            .iter_parts()
            .map(|part| (part.partname().to_string(), part.blob().to_vec()))
            .collect::<Vec<_>>();
        parts.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        parts
    }

    #[test]
    fn explicit_opens_match_serial_parts_deterministically() {
        let bytes = archive();
        let expected = package_parts(&OpcPackage::from_bytes(&bytes).unwrap());

        for _ in 0..3 {
            let (_source, context) = execution_context(16 * MIB);
            let session = OpenSession::new(context).unwrap();
            let package = session.from_bytes(&bytes, ReadLimits::default()).unwrap();
            assert_eq!(package_parts(&package), expected);
        }
    }

    #[test]
    fn explicit_open_keeps_ordinary_payloads_out_of_lazy_cache() {
        let bytes = cache_isolation_archive();
        let (_source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();
        let physical = crate::phys_pkg::PhysPkgReader::new(&bytes).unwrap();
        let _reader =
            crate::pkgreader::PackageReader::from_phys_reader_with_session(&physical, &session)
                .unwrap();
        assert_eq!(physical.archive().cache_size(), 0);

        let serial_physical = crate::phys_pkg::PhysPkgReader::new(&bytes).unwrap();
        let _serial_reader =
            crate::pkgreader::PackageReader::from_phys_reader(&serial_physical).unwrap();
        assert_eq!(serial_physical.archive().cache_size(), 2);
    }

    #[test]
    fn explicit_owned_open_retains_exact_source_authorization() {
        let bytes = archive();
        let (_source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();
        let package =
            OpcPackage::from_vec_with_execution(bytes.clone(), ReadLimits::default(), &session)
                .unwrap();

        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), bytes);
    }

    #[test]
    fn explicit_open_observes_pre_cancellation() {
        let bytes = archive();
        let (source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();
        source.cancel();

        assert!(matches!(
            session.from_bytes(&bytes, ReadLimits::default()),
            Err(OpcError::Cancelled)
        ));
    }

    #[test]
    fn explicit_open_refuses_insufficient_in_flight_memory() {
        let bytes = archive();
        let (_source, context) = execution_context(4 * 1024 - 1);
        let session = OpenSession::new(context).unwrap();

        assert!(matches!(
            session.from_bytes(&bytes, ReadLimits::default()),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Memory
        ));
    }

    #[test]
    fn read_limits_reject_before_execution_input_is_charged() {
        let bytes = archive();
        let (_source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            session.from_bytes(&bytes, limits),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::InputBytes,
                ..
            })
        ));
        assert_eq!(session.context().budget().used(Resource::InputBytes), 0);
    }

    #[test]
    fn explicit_owned_open_rejects_malformed_zip_before_input_charge() {
        let (_source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();

        assert!(matches!(
            session.from_vec(b"not an OPC ZIP".to_vec(), ReadLimits::default()),
            Err(OpcError::ZipError(_))
        ));
        assert_eq!(session.context().budget().used(Resource::InputBytes), 0);
    }

    #[test]
    fn explicit_owned_open_rejects_input_limit_before_zip_and_input_charge() {
        let (_source, context) = execution_context(16 * MIB);
        let session = OpenSession::new(context).unwrap();
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            session.from_vec(b"four".to_vec(), limits),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::InputBytes,
                actual: 4,
                maximum: 3,
            })
        ));
        assert_eq!(session.context().budget().used(Resource::InputBytes), 0);
    }
}
