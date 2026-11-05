#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure"
)]

use litchi_cfb::{OleError, OleFile};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ppt::{
    Error, Leniency, OpenOptions, Package, Presentation, SourceLimitKind, SourceLimits,
};
use std::io::{self, Cursor, Write};
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
};

fn fixture() -> Vec<u8> {
    include_bytes!("../../../test-data/ole/ppt/text_shapes.ppt").to_vec()
}

struct Source {
    bytes: Vec<u8>,
    revision: AtomicU64,
    reads: AtomicUsize,
    change_on_read: AtomicBool,
}

impl Source {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            change_on_read: AtomicBool::new(false),
        }
    }
}
impl ReadAt for Source {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            211,
            self.revision.load(Ordering::Acquire),
        ))
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        if self.change_on_read.swap(false, Ordering::AcqRel) {
            self.revision.fetch_add(1, Ordering::AcqRel);
        }
        let offset =
            usize::try_from(offset).map_err(|_| io::Error::from(io::ErrorKind::InvalidInput))?;
        let data = self.bytes.get(offset..).unwrap_or_default();
        let length = output.len().min(data.len());
        output[..length].copy_from_slice(&data[..length]);
        Ok(length)
    }
}

struct TempDir(std::path::PathBuf);
impl TempDir {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "litchi-ppt-owner-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir(&path).unwrap();
        Self(path)
    }
}
impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn owners_and_projections_are_small_shared_send_sync_handles() {
    fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Package>();
    assert_handle::<Presentation>();
    assert_eq!(size_of::<Package>(), size_of::<usize>());
    assert_eq!(size_of::<Presentation>(), size_of::<usize>());
    let package = Package::from_bytes(&fixture()).unwrap();
    let clone = package.clone();
    assert!(Arc::ptr_eq(
        &package.shared_ole_file_arc(),
        &clone.shared_ole_file_arc()
    ));
    let presentation = package.presentation().unwrap();
    let owned = presentation.clone();
    drop(package);
    drop(presentation);
    std::thread::scope(|scope| {
        scope.spawn(|| {
            assert_eq!(
                owned.text().unwrap(),
                clone.presentation().unwrap().text().unwrap()
            )
        });
    });
}

#[test]
fn eager_reader_and_preparsed_adoption_preserve_exact_bytes() {
    let bytes = fixture();
    let selected = SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new()
        .source_limits(selected)
        .leniency(Leniency::TolerateUserEditChainDefects);
    let packages = [
        Package::from_bytes_with_options(&bytes, options).unwrap(),
        Package::from_reader_with_options(&mut Cursor::new(bytes.clone()), options).unwrap(),
        Package::from_ole_file(
            OleFile::open(
                Cursor::new(bytes.clone()),
                litchi_cfb::OleFileLimits::default(),
            )
            .unwrap(),
            options,
        )
        .unwrap(),
    ];
    for package in packages {
        assert_eq!(package.source_limits(), selected);
        assert_eq!(package.leniency(), Leniency::TolerateUserEditChainDefects);
        assert_eq!(package.to_bytes().unwrap(), bytes);
        let mut sink = Vec::new();
        package.write_to(&mut sink).unwrap();
        assert_eq!(sink, bytes);
        assert_eq!(package.presentation().unwrap().source_limits(), selected);
    }
}

#[test]
fn source_adoption_reuses_index_and_projection_getters_do_not_read_payload() {
    let source = Arc::new(Source::new(fixture()));
    let native = litchi_ppt::SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let index = native.shared_ole_file_arc();
    let package = Package::from_source(native, OpenOptions::new()).unwrap();
    assert!(Arc::ptr_eq(&index, &package.shared_ole_file_arc()));
    let reads = source.reads.load(Ordering::Acquire);
    for _ in 0..10 {
        let _ = package.presentation().unwrap();
    }
    assert_eq!(source.reads.load(Ordering::Acquire), reads);
    assert_eq!(package.to_bytes().unwrap(), source.bytes);
}

fn assert_stale<T>(result: litchi_ppt::package::Result<T>) {
    assert!(matches!(
        result,
        Err(Error::Ole(OleError::SourceChanged { .. }))
    ));
}

#[test]
fn stale_source_is_refused_by_owner_and_retained_projection() {
    let source = Arc::new(Source::new(fixture()));
    let package = Package::from_read_at(source.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    source.revision.fetch_add(1, Ordering::AcqRel);
    assert_stale(package.presentation());
    assert_stale(package.metadata());
    assert_stale(package.to_bytes());
    assert_stale(presentation.text());
    assert_stale(presentation.slides());
}

#[test]
fn atomic_save_preserves_destination_on_source_change_and_removes_temporary_file() {
    let dir = TempDir::new();
    let path = dir.0.join("output.ppt");
    std::fs::write(&path, b"existing destination").unwrap();
    let source = Arc::new(Source::new(fixture()));
    let package = Package::from_read_at(source.clone()).unwrap();
    source.change_on_read.store(true, Ordering::Release);
    assert_stale(package.save(&path));
    assert_eq!(std::fs::read(&path).unwrap(), b"existing destination");
    assert_eq!(std::fs::read_dir(&dir.0).unwrap().count(), 1);
}

#[test]
fn atomic_publication_and_reopen_preserve_exact_source() {
    let dir = TempDir::new();
    let path = dir.0.join("output.ppt");
    let bytes = fixture();
    let package = Package::from_bytes(&bytes).unwrap();
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
    let reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), bytes);
    assert_eq!(
        reopened.presentation().unwrap().text().unwrap(),
        package.presentation().unwrap().text().unwrap()
    );
}

struct BrokenSink;
impl Write for BrokenSink {
    fn write(&mut self, _: &[u8]) -> io::Result<usize> {
        Err(io::Error::from(io::ErrorKind::BrokenPipe))
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
#[test]
fn sequential_publication_preserves_typed_io_failure_and_source_snapshot() {
    let bytes = fixture();
    let package = Package::from_bytes(&bytes).unwrap();
    assert!(
        matches!(package.write_to(&mut BrokenSink), Err(Error::Io(error)) if error.kind() == io::ErrorKind::BrokenPipe)
    );
    assert_eq!(package.to_bytes().unwrap(), bytes);
}

#[test]
fn complete_input_limits_apply_before_owned_or_sequential_admission() {
    let bytes = fixture();
    let maximum = bytes.len() - 1;
    let limits = SourceLimits::builder()
        .max_input_bytes(maximum)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(limits);
    for result in [
        Package::from_bytes_with_options(&bytes, options),
        Package::from_reader_with_options(&mut Cursor::new(bytes.clone()), options),
    ] {
        assert!(
            matches!(result, Err(Error::Limit { kind: SourceLimitKind::InputBytes, observed, maximum: found }) if observed > maximum && found == maximum)
        );
    }
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_owner_retains_profile_and_requires_explicit_ciphertext_publication() {
    use litchi_ppt::EncryptionProfile;
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 0, 0, 100, 100, "encrypted text")
        .unwrap();
    for key_bits in [40, 128] {
        let profile = EncryptionProfile::CryptoApiRc4 { key_bits };
        writer.set_password("password", profile).unwrap();
        let mut buffer = Cursor::new(Vec::new());
        writer.set_encryption_entropy(litchi_ppt::EncryptionEntropy::generate().unwrap());
        litchi_ppt::raw::encode(&mut writer, &mut buffer).unwrap();
        let bytes = buffer.into_inner();
        let package =
            Package::from_bytes_with_options(&bytes, OpenOptions::new().password("password"))
                .unwrap();
        assert_eq!(package.encryption(), Some(profile));
        assert!(package.is_encrypted());
        assert!(matches!(package.to_bytes(), Err(Error::EncryptionPolicy)));
        let mut sink = Vec::new();
        assert!(matches!(
            package.write_to(&mut sink),
            Err(Error::EncryptionPolicy)
        ));
        assert!(sink.is_empty());
        assert_eq!(package.to_encrypted_bytes().unwrap(), bytes);
        assert!(
            package
                .presentation()
                .unwrap()
                .text()
                .unwrap()
                .contains("encrypted text")
        );
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        assert!(
            Package::from_ole_file_or_password_required(ole, OpenOptions::new())
                .unwrap()
                .is_err()
        );
    }
}

#[test]
fn cloned_projections_share_successful_picture_materialization() {
    let mut writer = litchi_ppt::writer::Builder::new();
    writer.add_slide().unwrap();
    let mut picture = vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4];
    picture.resize(6012, 0);
    writer
        .add_picture_data_as(picture, litchi_ppt::writer::PictureKind::Png)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    let bytes = output.into_inner();
    let source = Arc::new(Source::new(bytes));
    let package = Package::from_read_at(source.clone()).unwrap();
    let first = package.presentation().unwrap();
    let second = first.clone();
    let images = first.images().unwrap();
    assert!(!images.is_empty());
    let reads = source.reads.load(Ordering::Acquire);
    assert_eq!(second.images().unwrap().len(), images.len());
    assert_eq!(source.reads.load(Ordering::Acquire), reads);
}

#[test]
fn shared_shapes_reuse_the_native_cache_and_outlive_the_slide() {
    let mut builder = litchi_ppt::Builder::new();
    let position = builder.add_slide().unwrap();
    builder
        .add_textbox(position, 10, 10, 240, 40, "shared text")
        .unwrap();
    let package = builder.build().unwrap();
    let presentation = package.presentation().unwrap();
    let shapes = {
        let slide = presentation.slide_at(0).unwrap().unwrap();
        let first = slide.shared_shapes().unwrap();
        let second = slide.shared_shapes().unwrap();
        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(first.as_ptr(), slide.shapes().unwrap().as_ptr());
        assert!(slide.tolerance_report().unwrap().is_clean());
        first
    };
    drop(presentation);
    drop(package);
    assert_eq!(shapes[0].text().unwrap(), "shared text");
}
