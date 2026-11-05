#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "owner acceptance fixtures intentionally fail fast"
)]

use litchi_cfb::{OleFile, OleFileLimits, OleWriter, SharedOleFile, consts::STGTY_STREAM};
use litchi_core::{ReadAt, SourceVersion};
use litchi_xls::package::{Error as PackageError, Package};
use litchi_xls::{
    Builder, OpenOptions, SourceBackedError, SourceBackedLimits, SourceBackedOptions, SourceLimits,
};
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
    let mut writer = Builder::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    writer.write_number(sheet, 4, 0, 1.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (name, bytes) in streams {
        writer.create_stream(&[*name], bytes).unwrap();
    }
    let mut output = Vec::new();
    writer.write_to(&mut Cursor::new(&mut output)).unwrap();
    output
}

fn signed_package() -> Vec<u8> {
    let source = authored_package();
    let mut ole = OleFile::open(Cursor::new(source), OleFileLimits::default()).unwrap();
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
    let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
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
    version_observations: AtomicUsize,
    bump_at_version: AtomicUsize,
    change_on_next_read: AtomicBool,
    truncate_on_next_read: AtomicBool,
}

impl CountingReadAt {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            version_observations: AtomicUsize::new(0),
            bump_at_version: AtomicUsize::new(0),
            change_on_next_read: AtomicBool::new(false),
            truncate_on_next_read: AtomicBool::new(false),
        }
    }

    fn read_count(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn bump_after_versions(&self, additional: usize) {
        let target = self
            .version_observations
            .load(Ordering::SeqCst)
            .saturating_add(additional);
        self.bump_at_version.store(target, Ordering::SeqCst);
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
        if self.truncate_on_next_read.swap(false, Ordering::SeqCst) {
            self.bump();
            return Ok(0);
        }
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
        let ordinal = self.version_observations.fetch_add(1, Ordering::SeqCst) + 1;
        if self.bump_at_version.load(Ordering::SeqCst) == ordinal {
            self.bump_at_version.store(0, Ordering::SeqCst);
            self.bump();
        }
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
    let cfb = Arc::new(
        SharedOleFile::open(source.clone(), litchi_cfb::SharedOleFileLimits::default()).unwrap(),
    );
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
fn source_reader_ref_separates_physical_and_logical_limits() {
    let original = fixture("poi/test-data/spreadsheet/WithTwoHyperLinks.xls");
    let mut ole = OleFile::open(Cursor::new(original), OleFileLimits::default()).unwrap();
    let workbook_stream = ole.open_stream(&["Workbook"]).unwrap();
    let unrelated = vec![0x5a; 4 * 1024];
    let bytes = cfb_with_streams(&[("Workbook", &workbook_stream), ("Unrelated", &unrelated)]);
    let input_bytes = u64::try_from(bytes.len()).unwrap();
    let worksheet_bytes = 64_u64;
    let limits = SourceBackedLimits::default()
        .with_max_input_bytes(input_bytes)
        .with_max_worksheet_scan_bytes(worksheet_bytes);
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(CountingReadAt::new(bytes.clone())),
        SourceBackedOptions::new().limits(limits),
    )
    .unwrap();
    let package = Package::from_source(workbook).unwrap();

    // The logical worksheet budget does not become a physical CFB stream
    // ceiling: an unrelated 4 KiB stream remains admissible and readable.
    assert_eq!(
        package
            .source()
            .unwrap()
            .open_stream(&["Unrelated"])
            .unwrap()
            .len(),
        4 * 1024
    );

    // The compatibility reader receives the retained physical profile, so a
    // real variable-length HLink payload is still decoded even though the
    // source-backed worksheet traversal budget is only 64 bytes.
    let reader = package.reader_ref().unwrap();
    let source_limits = reader.source_limits();
    assert_eq!(source_limits.max_input_bytes(), bytes.len());
    assert_eq!(
        source_limits.max_entry_bytes(),
        SourceLimits::MAX_ENTRY_BYTES
    );
    assert_eq!(
        source_limits.max_aggregate_bytes(),
        SourceLimits::MAX_AGGREGATE_BYTES
    );
    assert_eq!(reader.xls_worksheet(0).unwrap().hyperlinks().len(), 2);

    // The logical policy remains active on the source-backed owner itself;
    // it is not replaced by the physical profile used above.
    assert!(matches!(
        package
            .source_workbook()
            .unwrap()
            .cell_value_by_index(0, 0, 0),
        Err(SourceBackedError::ResourceLimit {
            resource: "worksheet scan bytes",
            maximum: 64,
            ..
        })
    ));
}

#[test]
fn sheet_catalog_rechecks_direct_source_after_capture() {
    let source = Arc::new(CountingReadAt::new(fixture("ole/xls/Simple.xls")));
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let package = Package::from_source(workbook).unwrap();

    let before = source.version_observations.load(Ordering::SeqCst);
    assert!(!package.sheet_catalog().unwrap().is_empty());
    let after_first = source.version_observations.load(Ordering::SeqCst);
    assert!(
        after_first > before,
        "cold catalog access must fence the source"
    );

    assert!(!package.sheet_catalog().unwrap().is_empty());
    let after_second = source.version_observations.load(Ordering::SeqCst);
    assert!(
        after_second > after_first,
        "a repeated direct catalog access must retain its source fence"
    );

    source.bump();
    assert!(matches!(
        package.sheet_catalog(),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
}

#[test]
fn signed_owner_publishes_exactly_and_allows_identity_editing() {
    let bytes = signed_package();
    let package = Package::from_bytes(&bytes).unwrap();

    assert_eq!(package.source_bytes().unwrap(), bytes);
    assert_eq!(package.to_bytes().unwrap(), bytes);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    let commit = package.edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 0);
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), bytes.as_slice());
    assert_eq!(commit.patch().after(), bytes.as_slice());
    assert_eq!(commit.into_package().to_bytes().unwrap(), bytes);
}

#[test]
fn signed_source_property_edit_reports_freshness_before_policy_refusal() {
    let source = Arc::new(CountingReadAt::new(signed_package()));
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let package = Package::from_source(workbook).unwrap();
    let mut edit = package.edit_property_sets().unwrap();
    edit.before().unwrap();
    source.bump();
    assert!(matches!(
        edit.edit_summary_information(|_| Ok(())),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
}

#[test]
fn property_set_render_rechecks_a_dynamic_source_after_last_read() {
    let source = Arc::new(CountingReadAt::new(fixture("ole/xls/Simple.xls")));
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let package = Package::from_source(workbook).unwrap();
    let mut edit = package.edit_property_sets().unwrap();
    edit.edit_summary_information(|property| property.set_title("changed"))
        .unwrap();
    // Four version observations belong to the leading `check_source`; the
    // next one is the property-set render fence added after editor.finish.
    source.bump_after_versions(5);
    assert!(matches!(
        edit.commit(),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
}

#[test]
fn property_set_publication_rechecks_a_dynamic_source_after_candidate_admission() {
    let source = Arc::new(CountingReadAt::new(fixture("ole/xls/Simple.xls")));
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let package = Package::from_source(workbook).unwrap();
    let mut edit = package.edit_property_sets().unwrap();
    edit.edit_summary_information(|property| property.set_title("changed"))
        .unwrap();
    // The render fence is followed by the two publication fences (before
    // and after admitting the candidate). Bump at that final observation so
    // candidate construction succeeds but publication still refuses a stale
    // original owner.
    source.bump_after_versions(7);
    assert!(matches!(
        edit.commit(),
        Err(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        ))
    ));
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
    let cfb = Arc::new(
        SharedOleFile::open(source.clone(), litchi_cfb::SharedOleFileLimits::default()).unwrap(),
    );
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
fn retained_narrow_profile_allows_noop_but_defers_entry_and_refuses_input_growth() {
    let bytes = authored_package();
    let source_limits = exact_source_limits(&bytes);
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(source_limits))
            .unwrap();

    let cfb_limits = OleFileLimits::new(u64::try_from(bytes.len()).unwrap()).unwrap();
    let options = litchi_xls::package::edit::EditOptions::new().cfb_limits(cfb_limits);
    let noop = package
        .edit_with_options(options)
        .unwrap()
        .commit()
        .unwrap();
    assert_eq!(noop.package().to_bytes().unwrap(), bytes);

    let mut growing = package.edit_with_options(options).unwrap();
    growing.rename_sheet(0_usize, "Renamed").unwrap();
    assert!(growing.commit().is_err());

    let too_small = OleFileLimits::new(u64::try_from(bytes.len() - 1).unwrap()).unwrap();
    let too_small_edit = package
        .edit_with_options(litchi_xls::package::edit::EditOptions::new().cfb_limits(too_small))
        .unwrap();
    assert!(too_small_edit.commit().is_err());
}

#[test]
fn transformed_source_handoff_retains_original_entry_limits() {
    let bytes = authored_package();
    let limits = exact_source_limits(&bytes);
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits)).unwrap();
    let mut edit = package.edit().unwrap();
    edit.rename_sheet(0_usize, "Renamed").unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn primary_commit_updates_projection_and_supports_exact_inverse_and_replay() {
    use litchi_core::sheet::CellValue;
    use litchi_xls::cell_values::Reference;
    let bytes = authored_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_number("Sheet1", reference, 8.5).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(
        commit
            .package()
            .workbook()
            .worksheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell_value(reference)
            .unwrap(),
        Some(CellValue::Float(8.5))
    );
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), bytes);
    let reopened = Package::from_bytes(&bytes).unwrap();
    let applied = reopened.apply(commit.patch()).unwrap();
    assert_eq!(
        applied.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    assert!(applied.package().apply(commit.patch()).is_err());
    let destination = ScratchFile::new("primary-commit");
    applied.package().save(destination.path()).unwrap();
    assert_eq!(
        Package::open(destination.path())
            .unwrap()
            .to_bytes()
            .unwrap(),
        applied.package().to_bytes().unwrap()
    );
    let (package, patch) = commit.into_parts();
    assert_eq!(
        package
            .apply(&patch.inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

#[test]
fn primary_noop_and_dropped_edits_leave_exact_source_unchanged() {
    use litchi_xls::cell_values::Reference;
    let bytes = authored_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let mut dropped = source.edit().unwrap();
    dropped
        .set_number(0_usize, Reference::new(3, 2).unwrap(), 99.0)
        .unwrap();
    drop(dropped);
    let mut noop = source.edit().unwrap();
    noop.set_number(0_usize, Reference::new(3, 2).unwrap(), 4.5)
        .unwrap();
    let commit = noop.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.into_package().to_bytes().unwrap(), bytes);
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn primary_detached_commits_recheck_source_even_for_noop() {
    for change in [false, true] {
        let source = Arc::new(CountingReadAt::new(authored_package()));
        let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
        let package = Package::from_source(workbook).unwrap();
        let mut edit = package.edit().unwrap();
        if change {
            edit.set_number(
                0_usize,
                litchi_xls::cell_values::Reference::new(3, 2).unwrap(),
                8.5,
            )
            .unwrap();
        }
        source.bump();
        assert!(
            edit.commit().is_err(),
            "detached publication must reject source replacement"
        );
    }
}

#[test]
fn primary_retains_narrow_object_profile_through_commit_inverse_and_explicit_widening() {
    use litchi_xls::cell_values::Reference;
    use litchi_xls::package::edit::EditOptions;
    let bytes = authored_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes.clone()), OleFileLimits::default()).unwrap();
    let maximum = ole.open_stream(&["Workbook"]).unwrap().len() as u64;
    let limits = litchi_ole_common::object::Limits {
        max_stream_size: maximum,
        ..Default::default()
    };
    let mut edit = source
        .edit_with_options(EditOptions::new().object_limits(limits))
        .unwrap();
    edit.set_number(0_usize, Reference::new(3, 2).unwrap(), 8.5)
        .unwrap();
    let commit = edit.commit().unwrap();
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), bytes);
    let mut grow = restored
        .package()
        .edit_with_options(EditOptions::default())
        .unwrap();
    grow.rename_sheet(0_usize, "LongerSheetName").unwrap();
    assert!(grow.commit().is_err());

    let mut wide = source.edit().unwrap();
    wide.rename_sheet(0_usize, "LongerSheetName").unwrap();
    let wide = wide.commit().unwrap();
    let narrow = source
        .edit_with_options(EditOptions::new().object_limits(limits))
        .unwrap()
        .commit()
        .unwrap()
        .into_package();
    assert!(
        narrow.apply(wide.patch()).is_err(),
        "recipient limits govern patch replay"
    );
    assert_eq!(narrow.to_bytes().unwrap(), bytes);
}

#[test]
fn primary_structural_edits_preserve_leaf_dependency_checks() {
    use litchi_core::sheet::CellValue;
    use litchi_xls::cell_values::{Reference, Value};
    let bytes = authored_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_cell(0_usize, Reference::new(9, 1).unwrap(), Value::Number(12.0))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(Reference::new(9, 1).unwrap())
            .unwrap(),
        Some(CellValue::Float(12.0))
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    let mut remove = commit.package().edit().unwrap();
    remove
        .remove_cell(0_usize, Reference::new(9, 1).unwrap())
        .unwrap();
    let removed = remove.commit().unwrap().into_package();
    assert_eq!(
        removed
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(Reference::new(9, 1).unwrap())
            .unwrap(),
        None
    );
}

#[test]
fn primary_signed_and_encrypted_sources_are_refused() {
    let signed = Package::from_bytes(&signed_package()).unwrap();
    let unchanged = signed.edit().unwrap();
    assert!(!unchanged.commit().unwrap().diagnostics().changed());
    let mut changed = signed.edit().unwrap();
    changed
        .set_value(
            0_usize,
            litchi_xls::cell_values::Reference::new(3, 2).unwrap(),
            litchi_xls::cell_values::Value::Number(9.0),
        )
        .unwrap();
    assert!(matches!(
        changed.commit(),
        Err(litchi_xls::EditError::CellValues(litchi_xls::Error::Signed))
    ));
    let bytes = fixture("poi/test-data/spreadsheet/xor-encryption-abc.xls");
    let encrypted =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().password("abc")).unwrap();
    assert!(encrypted.edit().is_err());
    assert_eq!(encrypted.source_bytes().unwrap(), bytes);
}

#[test]
fn source_truncation_during_publication_preserves_typed_fence_error() {
    for route in 0..3 {
        let source = Arc::new(CountingReadAt::new(authored_package()));
        let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
        let package = Package::from_source(workbook).unwrap();
        let destination = ScratchFile::new("truncation");
        std::fs::write(destination.path(), b"unchanged destination").unwrap();
        source.truncate_on_next_read.store(true, Ordering::SeqCst);
        let error = match route {
            0 => package.to_bytes().unwrap_err(),
            1 => package.write_to(&mut Vec::new()).unwrap_err(),
            _ => package.save(destination.path()).unwrap_err(),
        };
        assert!(
            matches!(
                error,
                PackageError::Source(SourceBackedError::SourceChanged { .. })
            ),
            "route {route}: {error}"
        );
        assert_eq!(
            std::fs::read(destination.path()).unwrap(),
            b"unchanged destination"
        );
    }
}

#[test]
fn primary_operation_limit_is_typed_and_failed_staging_is_atomic() {
    use litchi_xls::cell_values::Reference;
    use litchi_xls::package::edit::{Error as EditError, LimitKind};
    let mut writer = Builder::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    for row in 0..4097 {
        writer.write_number(sheet, row, 0, 1.0).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let source = Package::from_bytes(&output.into_inner()).unwrap();
    let mut edit = source.edit().unwrap();
    for row in 0..4096 {
        edit.set_number(0_usize, Reference::new(row, 0).unwrap(), 2.0)
            .unwrap();
    }
    let error = edit
        .set_number(0_usize, Reference::new(4096, 0).unwrap(), 2.0)
        .unwrap_err();
    assert!(matches!(
        error,
        EditError::Limit {
            kind: LimitKind::Operations,
            observed: 4097,
            maximum: 4096
        }
    ));
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(Reference::new(4096, 0).unwrap())
            .unwrap(),
        Some(litchi_core::sheet::CellValue::Float(1.0))
    );
}

#[test]
fn primary_default_followup_retains_explicit_wider_profile() {
    use litchi_xls::cell_values::Reference;
    use litchi_xls::package::edit::EditOptions;
    let bytes = authored_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let limits = litchi_ole_common::object::Limits {
        max_stream_size: 256 * 1024 * 1024,
        ..Default::default()
    };
    let committed = source
        .edit_with_options(EditOptions::new().object_limits(limits))
        .unwrap()
        .commit()
        .unwrap()
        .into_package();
    assert_eq!(
        committed
            .edit()
            .unwrap()
            .options()
            .object_limits_value()
            .max_stream_size,
        limits.max_stream_size
    );

    let maximum = 3 * 1024 * 1024 * 1024;
    let cfb = Arc::new(
        SharedOleFile::open(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            litchi_cfb::SharedOleFileLimits::new(maximum).unwrap(),
        )
        .unwrap(),
    );
    let workbook = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
        cfb,
        SourceBackedLimits::new(maximum).unwrap(),
    )
    .unwrap();
    let source = Package::from_source(workbook).unwrap();
    assert_eq!(
        source
            .edit()
            .unwrap()
            .options()
            .cfb_limits_value()
            .max_input_bytes(),
        maximum
    );
    let mut edit = source
        .edit_with_options(EditOptions::new().cfb_limits(OleFileLimits::new(maximum).unwrap()))
        .unwrap();
    edit.set_number(0_usize, Reference::new(3, 2).unwrap(), 8.5)
        .unwrap();
    let committed = edit.commit().unwrap().into_package();
    assert_eq!(
        committed.source_backed_limits().unwrap().max_input_bytes,
        maximum
    );
    assert_eq!(
        committed
            .edit()
            .unwrap()
            .options()
            .cfb_limits_value()
            .max_input_bytes(),
        maximum
    );
}

#[test]
fn primary_transfer_uses_package_donor_and_is_reversible() {
    use litchi_xls::cell_values::Reference;
    let bytes = scalar_transfer_package();
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit().unwrap();
    assert!(
        edit.copy_scalar_cell_from(
            &source,
            0_usize,
            Reference::new(3, 2).unwrap(),
            0_usize,
            Reference::new(9, 0).unwrap()
        )
        .unwrap()
        .is_some()
    );
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(Reference::new(9, 0).unwrap())
            .unwrap(),
        Some(litchi_core::sheet::CellValue::Float(4.5))
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

// The narrow native transfer closure refuses REFRESHALL even when false.
// Build the same dependency-free fixture used by xls_cross_workbook_scalar.
fn scalar_transfer_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(authored_package()), OleFileLimits::default()).unwrap();
    let mut workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut offset = 0;
    let removed = loop {
        let kind = u16::from_le_bytes(workbook[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            workbook[offset + 2..offset + 4].try_into().unwrap(),
        )) + 4;
        if kind == 0x01B7 {
            workbook.drain(offset..offset + length);
            break u32::try_from(length).unwrap();
        }
        assert_ne!(kind, 0x000A);
        offset += length;
    };
    offset = 0;
    loop {
        let kind = u16::from_le_bytes(workbook[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            workbook[offset + 2..offset + 4].try_into().unwrap(),
        )) + 4;
        if kind == 0x0085 {
            let start = u32::from_le_bytes(workbook[offset + 4..offset + 8].try_into().unwrap());
            workbook[offset + 4..offset + 8].copy_from_slice(&(start - removed).to_le_bytes());
        }
        if kind == 0x000A {
            break;
        }
        offset += length;
    }
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn primary_source_commit_retains_query_limits_and_positional_projection() {
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(authored_package())),
        SourceBackedOptions::new().limits(SourceBackedLimits::default().with_max_text_cells(1)),
    )
    .unwrap();
    let source = Package::from_source(workbook).unwrap();
    assert!(source.workbook().text().is_err());
    let mut edit = source.edit().unwrap();
    edit.set_number(
        0_usize,
        litchi_xls::cell_values::Reference::new(3, 2).unwrap(),
        8.5,
    )
    .unwrap();
    let committed = edit.commit().unwrap().into_package();
    assert!(committed.is_source_backed());
    assert!(
        committed
            .source()
            .expect("source-backed commit retains its CFB owner")
            .is_owned_immutable()
    );
    assert!(matches!(
        committed.workbook().text(),
        Err(PackageError::Source(
            SourceBackedError::ResourceLimit { .. }
        ))
    ));
}

#[test]
fn minimal_creation_supports_native_edits_and_exact_inverse() {
    use litchi_xls::cell_values::{Reference, Value};
    let package = Package::new().unwrap();
    assert_eq!(package.workbook().worksheet_names().unwrap(), ["Sheet1"]);
    assert!(package.workbook().text().unwrap().trim().is_empty());
    let original = package.to_bytes().unwrap();
    let reference = Reference::new(0, 0).unwrap();
    let mut edit = package.edit().unwrap();
    edit.insert_cell(0_usize, reference, Value::Number(42.0))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(reference)
            .unwrap(),
        Some(litchi_core::sheet::CellValue::Float(42.0))
    );
    assert_eq!(package.to_bytes().unwrap(), original);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original
    );
}

#[test]
fn first_cell_in_empty_sheet_repairs_later_sheet_offsets() {
    use litchi_xls::cell_values::{Reference, Value};
    let mut writer = Builder::new();
    writer.add_worksheet("Empty").unwrap();
    let second = writer.add_worksheet("Kept").unwrap();
    writer.write_number(second, 2, 3, 9.0).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut bytes).unwrap();
    let package = Package::from_bytes(bytes.get_ref()).unwrap();
    let mut edit = package.edit().unwrap();
    edit.insert_cell(0_usize, Reference::new(0, 0).unwrap(), Value::Number(7.0))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .worksheet(1_usize)
            .unwrap()
            .unwrap()
            .cell_value(Reference::new(2, 3).unwrap())
            .unwrap(),
        Some(litchi_core::sheet::CellValue::Float(9.0))
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        *bytes.get_ref()
    );
}

#[test]
fn transformed_byte_handoff_retains_native_policies_and_source_fences() {
    let bytes = authored_package();
    let limits = exact_source_limits(&bytes);
    let source =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits)).unwrap();
    let target = litchi_xls::raw::package_from_replacement_bytes(&source, bytes.clone()).unwrap();
    assert_eq!(target.source_limits(), source.source_limits());
    assert_eq!(target.leniency(), source.leniency());
    assert_eq!(
        target.compatibility_profile(),
        source.compatibility_profile()
    );
    assert!(!target.is_source_backed());
    assert_eq!(target.to_bytes().unwrap(), bytes);
    let mut oversized = bytes.clone();
    oversized.resize(bytes.len() + 512, 0);
    assert!(litchi_xls::raw::package_from_replacement_bytes(&source, oversized).is_err());

    let changing = Arc::new(CountingReadAt::new(bytes.clone()));
    let cfb = Arc::new(
        SharedOleFile::open(changing.clone(), litchi_cfb::SharedOleFileLimits::default()).unwrap(),
    );
    let read_limits = SourceBackedLimits {
        max_text_cells: 3,
        max_query_index_bytes: 0,
        ..Default::default()
    };
    let workbook =
        litchi_xls::raw::source_backed_workbook_from_shared_ole_file(cfb, read_limits).unwrap();
    let source = Package::from_source(workbook).unwrap();
    let target = litchi_xls::raw::package_from_replacement_bytes(&source, bytes.clone()).unwrap();
    assert!(target.is_source_backed());
    assert_eq!(target.source_backed_limits(), source.source_backed_limits());
    assert_eq!(target.to_bytes().unwrap(), bytes);
    changing.bump();
    assert!(matches!(
        litchi_xls::raw::package_from_replacement_bytes(&source, bytes),
        Err(litchi_xls::EditError::Package(PackageError::Source(
            SourceBackedError::SourceChanged { .. }
        )))
    ));
}

#[test]
fn primary_source_cell_uses_the_admitted_inventory_and_original_state() {
    use litchi_xls::cell_values::{Reference, Value};

    let bytes = authored_package();
    for source_backed in [false, true] {
        let source = Arc::new(CountingReadAt::new(bytes.clone()));
        let package = if source_backed {
            Package::from_source(
                litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap(),
            )
            .unwrap()
        } else {
            Package::from_bytes(&bytes).unwrap()
        };
        let reads_before_edit = source.read_count();
        let mut edit = package.edit().unwrap();
        assert_eq!(
            source.read_count(),
            reads_before_edit,
            "primary edit entry must not materialize the source"
        );
        let original = Reference::new(3, 2).unwrap();
        let missing = Reference::new(10, 2).unwrap();
        assert_eq!(
            edit.source_cell(0usize, original).unwrap().unwrap().value(),
            &Value::Number(4.5)
        );
        let reads_after_first_lookup = source.read_count();
        assert!(
            reads_after_first_lookup >= reads_before_edit,
            "the first actual operation may materialize the deferred source"
        );
        assert!(edit.source_cell("Sheet1", missing).unwrap().is_none());
        edit.set_value(0usize, original, Value::Number(9.0))
            .unwrap();
        edit.insert_cell(0usize, missing, Value::Number(10.0))
            .unwrap();
        assert_eq!(
            edit.source_cell("Sheet1", original)
                .unwrap()
                .unwrap()
                .value(),
            &Value::Number(4.5)
        );
        assert!(edit.source_cell(0usize, missing).unwrap().is_none());
        assert!(edit.source_cell("missing sheet", original).is_err());
        assert_eq!(
            source.read_count(),
            reads_after_first_lookup,
            "source cell lookup must reuse the admitted inventory"
        );
        let commit = edit.commit().unwrap();
        let next = commit.package().edit().unwrap();
        assert_eq!(
            next.source_cell(0usize, original).unwrap().unwrap().value(),
            &Value::Number(9.0)
        );
        assert_eq!(
            next.source_cell(0usize, missing).unwrap().unwrap().value(),
            &Value::Number(10.0)
        );
        assert_eq!(package.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn primary_source_cell_keeps_the_source_fence() {
    use litchi_xls::cell_values::Reference;
    let source = Arc::new(CountingReadAt::new(authored_package()));
    let package = Package::from_source(
        litchi_xls::SourceBackedWorkbook::from_read_at(source.clone()).unwrap(),
    )
    .unwrap();
    let edit = package.edit().unwrap();
    source.bump();
    for sheet in [0usize, 100usize] {
        let error = edit
            .source_cell(sheet, Reference::new(3, 2).unwrap())
            .unwrap_err();
        assert!(
            matches!(
                error,
                litchi_xls::EditError::Package(PackageError::Source(
                    SourceBackedError::SourceChanged { .. }
                ))
            ),
            "source freshness must precede even missing-sheet lookup: {error}"
        );
    }
}
