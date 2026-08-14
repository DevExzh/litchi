#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "fixture assertions panic on construction failure by design"
)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ppt::writer::{PictureKind, Writer};
use litchi_ppt::{Error, Package, RecordLimits, SourceBackedPackage};
use std::{
    io::{self, Cursor},
    sync::{
        Arc,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    },
};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt")
            .join(name),
    )
    .expect("PPT fixture")
}

fn serialize_ole(writer: &mut OleWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("OLE fixture serialization");
    output.into_inner()
}

fn picture_fixture() -> Vec<u8> {
    let mut writer = Writer::new();
    let slide = writer.add_slide().expect("PPT picture slide");
    writer
        .add_textbox(slide, 10, 10, 300, 40, "picture")
        .expect("PPT picture textbox");
    let mut picture = vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4];
    picture.extend(std::iter::repeat_n(0u8, 6_000));
    writer
        .add_picture_data_as(picture, PictureKind::Png)
        .expect("PPT picture payload");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("PPT picture serialization");
    output.into_inner()
}

fn malformed_pictures_fixture() -> Vec<u8> {
    let mut source = OleFile::open(Cursor::new(picture_fixture())).expect("PPT fixture OLE");
    let document = source
        .open_stream(&["PowerPoint Document"])
        .expect("PPT document stream");
    let current_user = source
        .open_stream(&["Current User"])
        .expect("PPT Current User stream");
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .expect("PPT document fixture stream");
    writer
        .create_stream(&["Current User"], &current_user)
        .expect("PPT Current User fixture stream");
    writer
        .create_stream(&["Pictures"], &[0; 8])
        .expect("malformed Pictures fixture stream");
    serialize_ole(&mut writer)
}

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    revision: AtomicU64,
    fail_next: AtomicBool,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            fail_next: AtomicBool::new(false),
        }
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        if self.fail_next.swap(false, Ordering::AcqRel) {
            return Err(io::Error::other("transient Pictures read"));
        }
        let start = usize::try_from(offset).map_err(|_error| io::Error::other("offset"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x505054,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn positional_presentation_defers_and_successfully_caches_pictures() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    let after_open = source.reads.load(Ordering::Acquire);

    assert!(presentation.has_pictures());
    assert_eq!(source.reads.load(Ordering::Acquire), after_open);

    let images = presentation.images().unwrap();
    assert!(!images.is_empty());
    let after_first_image_query = source.reads.load(Ordering::Acquire);
    assert!(after_first_image_query > after_open);

    let second = presentation.images().unwrap();
    assert_eq!(
        images
            .iter()
            .map(|image| image.data().unwrap().to_vec())
            .collect::<Vec<_>>(),
        second
            .iter()
            .map(|image| image.data().unwrap().to_vec())
            .collect::<Vec<_>>()
    );
    assert_eq!(
        source.reads.load(Ordering::Acquire),
        after_first_image_query,
        "successful Pictures materialization should be cached"
    );
}

#[test]
fn positional_picture_read_failure_is_not_cached() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    source.fail_next.store(true, Ordering::Release);

    assert!(matches!(presentation.images(), Err(Error::Ole(_))));
    assert!(presentation.images().is_ok());
}

#[test]
fn positional_picture_cache_rechecks_source_version() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    presentation.images().unwrap();
    source.revision.fetch_add(1, Ordering::AcqRel);

    assert!(matches!(presentation.images(), Err(Error::Ole(_))));
}

#[test]
fn positional_presentation_without_pictures_reports_absence_without_loading() {
    let source = Arc::new(CountingSource::new(fixture("empty.ppt")));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    let after_open = source.reads.load(Ordering::Acquire);

    assert!(!presentation.has_pictures());
    assert!(presentation.images().unwrap().is_empty());
    assert_eq!(source.reads.load(Ordering::Acquire), after_open);
}

#[test]
#[cfg(any(unix, windows))]
fn positional_path_open_keeps_the_file_source_positional() {
    let path =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt/empty.ppt");
    let package = SourceBackedPackage::from_path(&path).unwrap();

    assert_eq!(package.len(), std::fs::metadata(&path).unwrap().len());
    assert!(!package.presentation().unwrap().has_pictures());
}

#[test]
fn positional_package_limit_is_typed_before_cfb_payload_reads() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let max_package_bytes = source.bytes.len() - 1;
    let error = match SourceBackedPackage::from_read_at_with_limits(
        source.clone(),
        RecordLimits {
            max_package_bytes,
            ..RecordLimits::default()
        },
    ) {
        Ok(_) => panic!("package limit should reject the source"),
        Err(error) => error,
    };

    assert!(matches!(error, Error::ResourceLimit(message) if message.contains("package size")));
    assert_eq!(source.reads.load(Ordering::Acquire), 0);
}

#[test]
fn positional_presentation_checks_stream_aggregate_before_materialization() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let error = match package.presentation_with_limits(RecordLimits {
        max_aggregate_input_bytes: 1,
        ..RecordLimits::default()
    }) {
        Ok(_) => panic!("aggregate limit should reject the source"),
        Err(error) => error,
    };

    assert!(matches!(error, Error::ResourceLimit(message) if message.contains("aggregate")));
}

#[test]
fn positional_malformed_pictures_report_typed_officeart_failure() {
    let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
        malformed_pictures_fixture(),
    )))
    .unwrap();
    let presentation = package.presentation().unwrap();

    assert!(presentation.has_pictures());
    assert!(matches!(presentation.images(), Err(Error::OfficeArt(_))));
}

#[test]
fn positional_and_eager_picture_results_match() {
    let bytes = picture_fixture();
    let mut eager_package = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    let eager_presentation = eager_package.presentation().unwrap();
    let eager = eager_presentation.images().unwrap();
    let positional_package =
        SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
    let positional_presentation = positional_package.presentation().unwrap();
    let positional = positional_presentation.images().unwrap();

    assert_eq!(eager.len(), positional.len());
    for (left, right) in eager.iter().zip(positional.iter()) {
        assert_eq!(left.kind(), right.kind());
        assert_eq!(left.data().unwrap(), right.data().unwrap());
    }
}
