#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic by design"
)]

use crate::{
    CfbValidationError, OleError, OleWriter, SharedOleFile, SharedOleFileLimits, validate_source,
    validate_source_with_limits,
};
use litchi_core::{
    CheckStatus, OwnedSource, ReadAt, SourceVersion, ValidationLimitKind, ValidationLimits,
    ValidationReportError,
};
use std::{
    io::{self, Cursor},
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, AtomicUsize, Ordering},
    },
};

fn sample_file() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Mini"], &[0x41; 128]).unwrap();
    writer.create_stream(&["Regular"], &[0x42; 8_193]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[derive(Debug)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    calls: AtomicUsize,
    read_bytes: AtomicU64,
    version: SourceVersion,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            calls: AtomicUsize::new(0),
            read_bytes: AtomicU64::new(0),
            version: SourceVersion::new(7_001, 0),
        }
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.calls.fetch_add(1, Ordering::Relaxed);
        let Some(input) = usize::try_from(offset)
            .ok()
            .and_then(|start| self.bytes.get(start..))
        else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        self.read_bytes.fetch_add(count as u64, Ordering::Relaxed);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

#[derive(Debug)]
struct ChangingVersionSource {
    bytes: Vec<u8>,
    version_calls: AtomicUsize,
    switch_after: usize,
}

impl ReadAt for ChangingVersionSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let Some(input) = usize::try_from(offset)
            .ok()
            .and_then(|start| self.bytes.get(start..))
        else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let call = self.version_calls.fetch_add(1, Ordering::Relaxed);
        let revision = u64::from(call >= self.switch_after);
        Ok(SourceVersion::new(7_002, revision))
    }
}

#[derive(Debug)]
struct FailingReadSource {
    length: u64,
}

impl ReadAt for FailingReadSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.length)
    }

    fn read_at(&self, _offset: u64, _output: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::other("injected validation read failure"))
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(7_003, 0))
    }
}

#[derive(Debug)]
struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn mutate(&self) {
        let mut bytes = self.bytes.lock().unwrap();
        bytes[0x22] ^= 1;
        self.revision.fetch_add(1, Ordering::Release);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.lock().unwrap();
        let Some(input) = usize::try_from(offset)
            .ok()
            .and_then(|start| bytes.get(start..))
        else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            7_004,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn valid_source_completes_the_single_truthful_ingress_capability() {
    let bytes = sample_file();
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let report = validate_source(source.clone()).unwrap();

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(!report.has_fatal());
    assert_eq!(report.checks().len(), 1);
    assert_eq!(report.checks()[0].id().as_str(), "cfb.container.ingress");
    assert_eq!(report.checks()[0].status(), &CheckStatus::Complete);
    assert!(report.issues().is_empty());
    assert_eq!(source.bytes.as_slice(), bytes.as_slice());

    let calls = source.calls.load(Ordering::Relaxed);
    let read_bytes = source.read_bytes.load(Ordering::Relaxed);
    assert!(calls <= 32, "unexpected CFB ingress call count: {calls}");
    assert!(
        read_bytes <= bytes.len() as u64 * 4,
        "unexpected CFB ingress read volume: {read_bytes} for {} input bytes",
        bytes.len()
    );
}

#[test]
fn fat_topology_rejection_is_a_deterministic_structured_issue() {
    let mut bytes = sample_file();
    let sector_size = 1usize << u16::from_le_bytes([bytes[0x1e], bytes[0x1f]]);
    let first_fat_sector = read_u32(&bytes, 0x4c) as usize;
    let marker_offset = (first_fat_sector + 1) * sector_size + first_fat_sector * 4;
    write_u32(&mut bytes, marker_offset, 0xffff_ffff);

    let first = validate_source(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let second = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();

    assert_eq!(first, second);
    assert!(first.is_complete());
    assert!(first.has_errors());
    assert!(!first.has_fatal());
    assert_eq!(first.issues().len(), 1);
    assert_eq!(first.issues()[0].code(), "cfb.container.corrupted");
    assert_eq!(first.issues()[0].id(), second.issues()[0].id());
    assert_eq!(
        first.issues()[0].locations()[0].part(),
        Some("compound-file")
    );
}

#[test]
fn directory_topology_rejection_is_a_structured_issue() {
    let mut bytes = sample_file();
    let sector_size = 1usize << u16::from_le_bytes([bytes[0x1e], bytes[0x1f]]);
    let directory_sector = read_u32(&bytes, 0x30) as usize;
    let directory_offset = (directory_sector + 1) * sector_size;
    bytes[directory_offset + 66] = 2;

    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();

    assert!(report.is_complete());
    assert!(report.has_errors());
    assert_eq!(report.issues().len(), 1);
    assert_eq!(report.issues()[0].code(), "cfb.container.corrupted");
}

#[test]
fn configured_input_ceiling_blocks_without_calling_the_input_malformed() {
    let bytes = sample_file();
    let source_limit = SharedOleFileLimits::new(bytes.len() as u64 - 1).unwrap();
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        source_limit,
        ValidationLimits::default(),
    )
    .unwrap();

    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(report.issues().is_empty());
    assert!(matches!(
        report.checks()[0].status(),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn report_limits_are_enforced_before_retention() {
    let limits = ValidationLimits::new(0, 4, 2, 2, 128, 256, 128, 128, 128, 1_024);
    let error = validate_source_with_limits(
        Arc::new(OwnedSource::new(sample_file())),
        SharedOleFileLimits::default(),
        limits,
    )
    .unwrap_err();

    assert!(matches!(
        error,
        CfbValidationError::Report(ValidationReportError::Limit {
            kind: ValidationLimitKind::Checks,
            observed: 1,
            limit: 0
        })
    ));
}

#[test]
fn source_instability_and_io_failure_remain_errors() {
    let unstable = Arc::new(ChangingVersionSource {
        bytes: sample_file(),
        version_calls: AtomicUsize::new(0),
        switch_after: 1,
    });
    assert!(matches!(
        validate_source(unstable),
        Err(CfbValidationError::Ingress(OleError::SourceChanged { .. }))
    ));

    let failing = Arc::new(FailingReadSource {
        length: sample_file().len() as u64,
    });
    assert!(matches!(
        validate_source(failing),
        Err(CfbValidationError::Ingress(OleError::Io(_)))
    ));
}

#[test]
fn directory_byte_limit_is_a_blocked_cfb_validation_ingress() {
    let bytes = sample_file();
    let source = Arc::new(MutableSource::new(bytes));
    let limits = SharedOleFileLimits::new(SharedOleFileLimits::MAX_INPUT_BYTES)
        .unwrap()
        .with_max_directory_bytes(511)
        .unwrap();

    let report = validate_source_with_limits(source, limits, ValidationLimits::default())
        .expect("directory ceiling should be represented in the report");
    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        report.checks()[0].status(),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn version_change_between_preflight_and_canonical_ingress_is_an_error() {
    let unstable = Arc::new(ChangingVersionSource {
        bytes: sample_file(),
        version_calls: AtomicUsize::new(0),
        switch_after: 2,
    });

    assert!(matches!(
        validate_source(unstable),
        Err(CfbValidationError::Ingress(OleError::SourceChanged { .. }))
    ));
}

#[test]
fn shared_view_revalidation_refuses_a_changed_source_version() {
    let source = Arc::new(MutableSource::new(sample_file()));
    let shared = SharedOleFile::open(source.clone()).unwrap();
    source.mutate();

    assert!(matches!(
        shared.validate(ValidationLimits::default()),
        Err(CfbValidationError::Ingress(OleError::SourceChanged { .. }))
    ));
}
