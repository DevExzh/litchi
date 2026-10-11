#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "owner acceptance fixtures intentionally fail fast"
)]

use litchi_cfb::{OleFile, OleFileLimits, OleWriter, SharedOleFile, consts::STGTY_STREAM};
use litchi_core::{ReadAt, SourceVersion};
use litchi_xls::package::{Error as PackageError, Package};
use litchi_xls::{OpenOptions, SourceBackedError, SourceBackedLimits, SourceLimits, Writer};
use std::io::{self, Cursor};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

fn fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative),
    )
    .unwrap()
}

fn authored_package() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    writer.write_number(sheet, 4, 0, 1.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn signed_package() -> Vec<u8> {
    let source = authored_package();
    let mut ole = OleFile::open(Cursor::new(source)).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"opaque signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn exact_source_limits(bytes: &[u8]) -> SourceLimits {
    let ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let mut largest_entry = 0_usize;
    let mut aggregate = 0_usize;
    for index in 0..ole.directory_entry_count() {
        let Some(entry) = ole.directory_entry_by_sid(u32::try_from(index).unwrap()) else {
            continue;
        };
        if entry.entry_type != STGTY_STREAM {
            continue;
        }
        let size = usize::try_from(entry.size).unwrap();
        largest_entry = largest_entry.max(size);
        aggregate = aggregate.checked_add(size).unwrap();
    }
    SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .max_entry_bytes(largest_entry)
        .unwrap()
        .max_aggregate_bytes(aggregate)
        .unwrap()
        .build()
        .unwrap()
}

#[derive(Debug)]
struct CountingReadAt {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    revision: AtomicU64,
    change_on_next_read: AtomicBool,
}

impl CountingReadAt {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            change_on_next_read: AtomicBool::new(false),
        }
    }

    fn read_count(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn arm_change_on_next_read(&self) {
        self.change_on_next_read.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("fixture length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if start >= self.bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        self.reads.fetch_add(1, Ordering::SeqCst);
        if self.change_on_next_read.swap(false, Ordering::SeqCst) {
            self.bump();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_535f_4f57_4e52,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct ScratchFile {
    path: PathBuf,
}

impl ScratchFile {
    fn new(label: &str) -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "litchi-xls-package-owner-{label}-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        let _ = std::fs::remove_file(&path);
        Self { path }
    }

    fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for ScratchFile {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.path);
    }
}

#[test]
fn eager_owner_preserves_exact_bytes_and_projection_outlives_package_clone() {
    let bytes = authored_package();
    let projection = {
        let package = Package::from_bytes(&bytes).unwrap();
        assert!(!package.is_source_backed());
        assert_eq!(package.source_bytes().unwrap(), bytes);
        assert_eq!(package.to_bytes().unwrap(), bytes);
        assert_eq!(package.reader_ref().unwrap().sheets().len(), 1);

        let clone = package.clone();
        clone.workbook()
    };

    assert_eq!(projection.worksheet_count().unwrap(), 1);
    let names = projection.worksheet_names().unwrap();
    assert_eq!(names.len(), 1);
    let worksheet = projection.worksheet(0_usize).unwrap().unwrap();
    assert_eq!(worksheet.index().unwrap(), 0);
    assert_eq!(worksheet.name().unwrap(), names[0]);
    assert!(projection.reader_ref().is_ok());
}

#[test]
fn source_adoption_is_lazy_and_retains_the_original_cfb_identity() {
    let source = Arc::new(CountingReadAt::new(fixture("ole/xls/Simple.xls")));
    let cfb = Arc::new(SharedOleFile::open(source.clone()).unwrap());
    let workbook = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
        cfb.clone(),
        SourceBackedLimits::default(),
    )
    .unwrap();
    let reads_before_adoption = source.read_count();

    let package = Package::from_source(workbook).unwrap();
    assert_eq!(source.read_count(), reads_before_adoption);
    let retained = package.shared_source().unwrap();
    assert!(Arc::ptr_eq(&cfb, &retained));
    assert!(package.is_source_backed());

    package.workbook().text().unwrap();
    assert!(source.read_count() > reads_before_adoption);

    package.reader_ref().unwrap();
    source.bump();
    assert!(matches!(
        package.reader_ref(),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
    assert!(matches!(
        package.workbook().worksheet_names(),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
}

#[test]
fn signed_owner_publishes_exactly_but_refuses_editing() {
    let bytes = signed_package();
    let package = Package::from_bytes(&bytes).unwrap();

    assert_eq!(package.source_bytes().unwrap(), bytes);
    assert_eq!(package.to_bytes().unwrap(), bytes);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    assert!(package.edit_snapshot().is_err());
}

#[test]
fn unlocked_encrypted_owner_refuses_ordinary_publication_without_clobbering_destination() {
    let bytes = fixture("poi/test-data/spreadsheet/xor-encryption-abc.xls");
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().password("abc")).unwrap();

    assert_eq!(package.source_bytes().unwrap(), bytes);
    assert!(matches!(
        package.to_bytes(),
        Err(PackageError::EncryptionPolicy)
    ));
    assert!(matches!(
        package.to_plain_bytes(),
        Err(PackageError::EncryptionPolicy)
    ));
    let mut output = Vec::new();
    assert!(matches!(
        package.write_to(&mut output),
        Err(PackageError::EncryptionPolicy)
    ));

    let destination = ScratchFile::new("encrypted");
    let original = b"keep this destination".to_vec();
    std::fs::write(destination.path(), &original).unwrap();
    assert!(matches!(
        package.save(destination.path()),
        Err(PackageError::EncryptionPolicy)
    ));
    assert_eq!(std::fs::read(destination.path()).unwrap(), original);
}

#[test]
fn atomic_source_change_after_staging_leaves_destination_untouched() {
    let source = Arc::new(CountingReadAt::new(fixture("ole/xls/Simple.xls")));
    let cfb = Arc::new(SharedOleFile::open(source.clone()).unwrap());
    let workbook = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
        cfb,
        SourceBackedLimits::default(),
    )
    .unwrap();
    let package = Package::from_source(workbook).unwrap();

    let destination = ScratchFile::new("source-change");
    let original = b"preserve this artifact".to_vec();
    std::fs::write(destination.path(), &original).unwrap();
    source.arm_change_on_next_read();

    assert!(matches!(
        package.save(destination.path()),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
    assert_eq!(std::fs::read(destination.path()).unwrap(), original);
}

#[test]
fn retained_narrow_profile_allows_noop_but_refuses_entry_and_input_growth() {
    let bytes = authored_package();
    let source_limits = exact_source_limits(&bytes);
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(source_limits))
            .unwrap();

    let cfb_limits = OleFileLimits::new(u64::try_from(bytes.len()).unwrap()).unwrap();
    let object_limits = litchi_ole_common::object::Limits::default();
    let snapshot = package
        .edit_snapshot_with_limits(object_limits, cfb_limits)
        .unwrap();
    let noop = snapshot.edit().commit().unwrap();
    assert_eq!(noop.snapshot().bytes(), bytes.as_slice());

    let mut growing = snapshot.edit();
    growing.rename_sheet(0_usize.into(), "Renamed").unwrap();
    assert!(growing.commit().is_err());

    let too_small = OleFileLimits::new(u64::try_from(bytes.len() - 1).unwrap()).unwrap();
    assert!(
        package
            .edit_snapshot_with_limits(object_limits, too_small)
            .is_err()
    );
}

#[test]
fn transformed_source_handoff_retains_original_entry_limits() {
    let bytes = authored_package();
    let limits = exact_source_limits(&bytes);
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits)).unwrap();
    let snapshot = litchi_xls::raw::cell_values_snapshot_from_bytes(&package, bytes).unwrap();
    let mut edit = snapshot.edit();
    edit.rename_sheet(0_usize.into(), "Renamed").unwrap();
    assert!(edit.commit().is_err());
}
