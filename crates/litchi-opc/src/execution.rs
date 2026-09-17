//! Explicit, reusable execution adapter for eager OPC package opens.
//!
//! [`OpenSession`] adapts the runtime-neutral policy in
//! [`litchi_core::ExecutionContext`] to soapberry-zip's local bounded read
//! session. It is intentionally separate from ordinary [`crate::OpcPackage`]
//! constructors: those constructors remain synchronous and serial.

use crate::{OpcError, OpcPackage, ReadLimits, Result};
use litchi_core::{
    AffinityPolicy, ExecutionContext, ExecutionError, Reservation, Resource, ScopedWorkers,
};
use soapberry_zip::office::{
    CancellationProbe, LazyArchiveReader, ParallelAffinity, ParallelReadLimits, ParallelReadSession,
};
use std::num::NonZeroUsize;
use std::sync::{Arc, Mutex};

/// Reusable, explicitly scheduled eager OPC open adapter.
///
/// The session uses a lazy private ZIP worker pool unless the caller attaches
/// a scoped-worker facility, never uses a global Rayon pool, and retains the
/// caller-selected [`ExecutionContext`] for every open. It is an advanced
/// ingress API; ordinary package CRUD never exposes this type.
#[derive(Debug)]
pub struct OpenSession {
    context: ExecutionContext,
    zip: ParallelReadSession,
    read_permits: Mutex<Option<ReadPermits>>,
}

#[derive(Debug)]
struct ReadPermits {
    width: usize,
    _workers: Reservation,
}

struct ReadAdmission {
    width: usize,
    workers: Option<Reservation>,
    io: Reservation,
}

#[derive(Debug)]
struct ZipScopedWorkers(Arc<dyn ScopedWorkers>);

impl soapberry_zip::office::ScopedWorkers for ZipScopedWorkers {
    fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]) {
        self.0.run_all(tasks);
    }
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
        let zip = match context.scoped_workers() {
            Some(workers) => ParallelReadSession::with_scoped_workers(
                zip_limits,
                Arc::new(ZipScopedWorkers(Arc::clone(workers))),
            ),
            None => ParallelReadSession::new(zip_limits).map_err(OpcError::ParallelRead)?,
        };
        Ok(Self {
            context,
            zip,
            read_permits: Mutex::new(None),
        })
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
        if names.is_empty() {
            return Ok(Vec::new());
        }
        let limits = self.context.limits();
        let max_tasks = limits.max_in_flight_tasks().get();
        let max_bytes = limits.max_in_flight_bytes().get();
        let mut batch_tasks = 0usize;
        let mut batch_bytes = 0u64;
        let mut has_parallel_batch = false;
        for name in names {
            let metadata = archive.metadata(name).map_err(OpcError::from)?;
            let bytes = metadata.uncompressed_size();
            if bytes > max_bytes {
                return Err(OpcError::ParallelRead(
                    soapberry_zip::ErrorKind::ParallelReadInFlightBytesExceeded {
                        actual: bytes,
                        maximum: max_bytes,
                    }
                    .into(),
                ));
            }
            let exceeds_bytes = batch_tasks > 0
                && batch_bytes
                    .checked_add(bytes)
                    .is_none_or(|next_bytes| next_bytes > max_bytes);
            if batch_tasks == max_tasks || exceeds_bytes {
                has_parallel_batch |= batch_tasks > 1 && batch_bytes >= limits.min_parallel_bytes();
                batch_tasks = 0;
                batch_bytes = 0;
            }
            batch_tasks += 1;
            batch_bytes = batch_bytes
                .checked_add(bytes)
                .expect("validated ZIP batch byte total fits u64");
        }
        has_parallel_batch |= batch_tasks > 1 && batch_bytes >= limits.min_parallel_bytes();
        let in_flight = declared_bytes.min(limits.max_in_flight_bytes().get());
        let reservation = self
            .context
            .reserve(Resource::Memory, in_flight)
            .map_err(map_execution_error)?;
        self.context
            .consume(Resource::Work, declared_bytes)
            .map_err(map_execution_error)?;

        let candidate = has_parallel_batch
            && declared_bytes >= limits.min_parallel_bytes()
            && (limits.min_task_bytes() == 0
                || names.iter().all(|name| {
                    archive
                        .metadata(name)
                        .map(|metadata| metadata.uncompressed_size() >= limits.min_task_bytes())
                        .unwrap_or(false)
                }));
        let requested_width = limits
            .workers()
            .get()
            .min(limits.max_in_flight_tasks().get())
            .min(names.len());
        let admission = if candidate && requested_width > 1 {
            self.admit_parallel_workers(requested_width, names.len() as u64)?
        } else {
            self.admit_serial_workers(names.len() as u64)?
        };
        let parallel_width = admission.width;
        let _workers = admission.workers;
        let _io = admission.io;
        let parallel_width =
            NonZeroUsize::new(parallel_width).expect("read admission width is non-zero");

        let cancellation = ContextCancellation(&self.context);
        let results = archive
            .read_many_with_session_width(&self.zip, names, parallel_width, &cancellation)
            .map_err(map_parallel_read_error);
        drop(reservation);
        results
    }

    fn consume_read_cpu_tasks(&self, tasks: u64) -> Result<()> {
        self.context.check().map_err(map_execution_error)?;
        self.context
            .consume(Resource::CpuTasks, tasks)
            .map_err(map_execution_error)
    }

    fn admit_serial_workers(&self, cpu_tasks: u64) -> Result<ReadAdmission> {
        let permits = if let Ok(permits) = self.read_permits.lock() {
            permits
        } else {
            return Err(OpcError::Execution(ExecutionError::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: Resource::Workers,
                    observed: 1,
                    limit: 0,
                    scope: Arc::from("opc-open read permits"),
                },
            )));
        };
        if permits.is_some() {
            let io = self
                .context
                .reserve(Resource::IoConcurrency, 1)
                .map_err(map_execution_error)?;
            self.consume_read_cpu_tasks(cpu_tasks)?;
            return Ok(ReadAdmission {
                width: 1,
                workers: None,
                io,
            });
        }
        let workers = self
            .context
            .reserve(Resource::Workers, 1)
            .map_err(map_execution_error)?;
        let io = match self.context.reserve(Resource::IoConcurrency, 1) {
            Ok(io) => io,
            Err(error) => {
                drop(workers);
                return Err(map_execution_error(error));
            },
        };
        self.consume_read_cpu_tasks(cpu_tasks)?;
        Ok(ReadAdmission {
            width: 1,
            workers: Some(workers),
            io,
        })
    }

    fn admit_parallel_workers(&self, requested: usize, cpu_tasks: u64) -> Result<ReadAdmission> {
        let mut permits = if let Ok(permits) = self.read_permits.lock() {
            permits
        } else {
            return Err(OpcError::Execution(ExecutionError::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: Resource::Workers,
                    observed: 1,
                    limit: 0,
                    scope: Arc::from("opc-open read permits"),
                },
            )));
        };
        if let Some(retained) = permits.as_ref() {
            let maximum = retained.width.min(requested);
            let mut last_limit = None;
            for width in (1..=maximum).rev() {
                match self.context.reserve(Resource::IoConcurrency, width as u64) {
                    Ok(io) => {
                        self.consume_read_cpu_tasks(cpu_tasks)?;
                        return Ok(ReadAdmission {
                            width,
                            workers: None,
                            io,
                        });
                    },
                    Err(ExecutionError::ResourceLimit(limit))
                        if limit.resource == Resource::IoConcurrency =>
                    {
                        last_limit = Some(limit);
                    },
                    Err(error) => return Err(map_execution_error(error)),
                }
            }
            if let Some(limit) = last_limit {
                return Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)));
            }
        }

        for width in (2..=requested).rev() {
            let workers = match self.context.reserve(Resource::Workers, width as u64) {
                Ok(workers) => workers,
                Err(ExecutionError::ResourceLimit(limit))
                    if limit.resource == Resource::Workers =>
                {
                    continue;
                },
                Err(error) => return Err(map_execution_error(error)),
            };
            let io = match self.context.reserve(Resource::IoConcurrency, width as u64) {
                Ok(io) => io,
                Err(ExecutionError::ResourceLimit(limit))
                    if limit.resource == Resource::IoConcurrency =>
                {
                    drop(workers);
                    continue;
                },
                Err(error) => {
                    drop(workers);
                    return Err(map_execution_error(error));
                },
            };
            if self.context.scoped_workers().is_some() {
                self.consume_read_cpu_tasks(cpu_tasks)?;
                return Ok(ReadAdmission {
                    width,
                    workers: Some(workers),
                    io,
                });
            }
            let worker_width = NonZeroUsize::new(width).expect("parallel width is non-zero");
            self.consume_read_cpu_tasks(cpu_tasks)?;
            if let Err(error) = self.zip.ensure_local_pool_at_width(worker_width) {
                drop(io);
                drop(workers);
                return Err(OpcError::ParallelRead(error));
            }
            *permits = Some(ReadPermits {
                width,
                _workers: workers,
            });
            return Ok(ReadAdmission {
                width,
                workers: None,
                io,
            });
        }
        let workers = self
            .context
            .reserve(Resource::Workers, 1)
            .map_err(map_execution_error)?;
        let io = match self.context.reserve(Resource::IoConcurrency, 1) {
            Ok(io) => io,
            Err(error) => {
                drop(workers);
                return Err(map_execution_error(error));
            },
        };
        self.consume_read_cpu_tasks(cpu_tasks)?;
        Ok(ReadAdmission {
            width: 1,
            workers: Some(workers),
            io,
        })
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
            .try_iter_parts()
            .map(|part| {
                let part = part.unwrap();
                (part.partname().to_string(), part.blob().to_vec())
            })
            .collect::<Vec<_>>();
        parts.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        parts
    }

    fn direct_read_context(
        budget: Budget,
        workers: usize,
    ) -> (CancellationSource, ExecutionContext) {
        let (source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(workers).unwrap(),
            NonZeroUsize::new(workers.max(4)).unwrap(),
            NonZeroU64::new(16 * 1024).unwrap(),
            0,
        )
        .unwrap();
        (source, ExecutionContext::new(budget, cancellation, limits))
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
    fn explicit_zip_read_refuses_zero_io_before_member_payload_read() {
        let bytes = archive();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let budget = Budget::root(
            "opc-open-zero-io",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 0, u64::MAX),
        );
        let (_source, context) = direct_read_context(budget.clone(), 4);
        let session = OpenSession::new(context).unwrap();
        let names = ["benchmark/0.bin", "benchmark/1.bin"];

        assert!(matches!(
            session.read_many(&reader, &names),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::IoConcurrency
        ));
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::CpuTasks), 0);
    }

    #[test]
    fn explicit_zip_sessions_share_root_worker_admission_and_release_on_drop() {
        let bytes = archive();
        let names = ["benchmark/0.bin", "benchmark/1.bin"];
        let root = Budget::root(
            "opc-open-shared-root",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(3, 3, u64::MAX),
        );
        let first_budget = root.child(
            "first-open",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let second_budget = root.child(
            "second-open",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (_first_source, first_context) = direct_read_context(first_budget, 4);
        let (_second_source, second_context) = direct_read_context(second_budget, 4);
        let first = OpenSession::new(first_context).unwrap();
        let second = OpenSession::new(second_context).unwrap();
        let first_reader = LazyArchiveReader::new(&bytes).unwrap();
        let second_reader = LazyArchiveReader::new(&bytes).unwrap();

        let first_result = first.read_many(&first_reader, &names).unwrap();
        assert_eq!(first.zip.local_pool_worker_count(), 2);
        assert_eq!(root.used(Resource::Workers), 2);
        let second_result = second.read_many(&second_reader, &names).unwrap();
        assert_eq!(first_result.len(), second_result.len());
        // The first session retains width two for its lazy pool; the second
        // deterministically narrows to one under the shared root cap of three
        // and releases its serial admission at return.
        assert!(second.read_permits.lock().unwrap().is_none());
        assert_eq!(second.zip.local_pool_worker_count(), 0);
        assert_eq!(root.used(Resource::Workers), 2);
        assert_eq!(root.used(Resource::IoConcurrency), 0);
        drop(first);
        assert_eq!(root.used(Resource::Workers), 0);
        drop(second);
        assert_eq!(root.used(Resource::Workers), 0);
    }

    #[test]
    fn zero_cpu_budget_refuses_before_retaining_a_zip_pool() {
        let bytes = archive();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let budget = Budget::root(
            "opc-open-zero-cpu",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 4, 0),
        );
        let (_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(16 * 1024).unwrap(),
            0,
        )
        .unwrap();
        let session =
            OpenSession::new(ExecutionContext::new(budget.clone(), cancellation, limits)).unwrap();
        let names = ["benchmark/0.bin", "benchmark/1.bin"];

        assert!(matches!(
            session.read_many(&reader, &names),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::CpuTasks
        ));
        assert_eq!(session.zip.local_pool_worker_count(), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        assert_eq!(budget.used(Resource::CpuTasks), 0);
    }

    #[test]
    fn zip_byte_cap_that_splits_every_batch_does_not_retain_a_pool() {
        let bytes = archive();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let budget = Budget::root(
            "opc-open-serial-batches",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 4, u64::MAX),
        );
        let (_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(1024).unwrap(),
            0,
        )
        .unwrap();
        let session =
            OpenSession::new(ExecutionContext::new(budget.clone(), cancellation, limits)).unwrap();
        let names = ["benchmark/0.bin", "benchmark/1.bin"];
        let results = session.read_many(&reader, &names).unwrap();

        assert!(results.iter().all(|(_, result)| result.is_ok()));
        assert_eq!(session.zip.local_pool_worker_count(), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        assert_eq!(budget.used(Resource::CpuTasks), 2);
    }

    #[test]
    fn explicit_zip_read_preflights_oversized_members_before_cpu_admission() {
        let bytes = archive();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let budget = Budget::root(
            "opc-open-oversized-member",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 4, u64::MAX),
        );
        let (_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(512).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, limits);
        let session = OpenSession::new(context).unwrap();
        let names = ["benchmark/0.bin", "benchmark/1.bin"];

        assert!(matches!(
            session.read_many(&reader, &names),
            Err(OpcError::ParallelRead(error))
                if matches!(
                    error.kind(),
                    soapberry_zip::ErrorKind::ParallelReadInFlightBytesExceeded {
                        actual: 1024,
                        maximum: 512,
                    }
                )
        ));
        assert_eq!(budget.used(Resource::CpuTasks), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
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
