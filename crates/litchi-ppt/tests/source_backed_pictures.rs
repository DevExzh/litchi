#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "fixture assertions panic on construction failure by design"
)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ppt::writer::{Builder, PictureKind};
use litchi_ppt::{
    Error, Leniency, OpenOptions, SourceBackedPackage, SourceLimitKind, SourceLimits, raw::Reader,
};
use std::{
    io::{self, Cursor, Read, Seek, SeekFrom},
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
    let mut writer = Builder::new();
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
    litchi_ppt::raw::encode(&mut writer, &mut output).expect("PPT picture serialization");
    output.into_inner()
}

fn malformed_pictures_fixture() -> Vec<u8> {
    let mut source = OleFile::open(
        Cursor::new(picture_fixture()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("PPT fixture OLE");
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

struct RejectingSource {
    bytes: Arc<Vec<u8>>,
    ranges: Vec<(u64, u64)>,
    rejected: Arc<AtomicBool>,
}

impl RejectingSource {
    fn new(bytes: Vec<u8>, ranges: Vec<(u64, u64)>) -> (Arc<Self>, Arc<AtomicBool>) {
        let rejected = Arc::new(AtomicBool::new(false));
        let source = Arc::new(Self {
            bytes: Arc::new(bytes),
            ranges,
            rejected: Arc::clone(&rejected),
        });
        (source, rejected)
    }
}

impl ReadAt for RejectingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let end = offset.saturating_add(output.len() as u64);
        if self
            .ranges
            .iter()
            .any(|(start, stop)| offset < *stop && end > *start)
        {
            self.rejected.store(true, Ordering::Release);
            return Err(io::Error::other("stream payload read was not admitted"));
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
        Ok(SourceVersion::new(0x505054, 0))
    }
}

struct RejectingReader {
    inner: Cursor<Vec<u8>>,
    ranges: Vec<(u64, u64)>,
    rejected: Arc<AtomicBool>,
}

impl RejectingReader {
    fn new(bytes: Vec<u8>, ranges: Vec<(u64, u64)>) -> (Self, Arc<AtomicBool>) {
        let rejected = Arc::new(AtomicBool::new(false));
        let reader = Self {
            inner: Cursor::new(bytes),
            ranges,
            rejected: Arc::clone(&rejected),
        };
        (reader, rejected)
    }
}

impl Read for RejectingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let offset = self.inner.position();
        let end = offset.saturating_add(output.len() as u64);
        if self
            .ranges
            .iter()
            .any(|(start, stop)| offset < *stop && end > *start)
        {
            self.rejected.store(true, Ordering::Release);
            return Err(io::Error::other("stream payload read was not admitted"));
        }
        self.inner.read(output)
    }
}

impl Seek for RejectingReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

fn fat_stream_range(bytes: &[u8], name: &str) -> (u64, u64) {
    let ole = OleFile::open(
        Cursor::new(bytes.to_vec()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("fixture OLE");
    let entry = ole
        .list_directory_entries(&[])
        .expect("fixture directory")
        .into_iter()
        .find(|entry| entry.name == name)
        .expect("fixture stream");
    assert!(!entry.is_minifat, "fixture stream must use FAT sectors");
    let sector_size = 1usize << usize::from(u16::from_le_bytes([bytes[30], bytes[31]]));
    let start = (u64::from(entry.start_sector) + 1) * sector_size as u64;
    let end = start + entry.size.div_ceil(sector_size as u64) * sector_size as u64;
    (start, end)
}

fn unrelated_stream_fixture(name: &str, size: usize) -> (Vec<u8>, (u64, u64)) {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &[0; 64])
        .expect("PPT owner stream");
    writer
        .create_stream(&[name], &vec![0xA5; size])
        .expect("unrelated stream");
    let bytes = serialize_ole(&mut writer);
    let range = fat_stream_range(&bytes, name);
    (bytes, range)
}

fn all_streams_fixture() -> (Vec<u8>, Vec<(u64, u64)>) {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &[0; 64])
        .expect("PPT owner stream");
    writer
        .create_stream(&["\u{0005}SummaryInformation"], &vec![0xA5; 8 * 1024])
        .expect("metadata stream");
    writer
        .create_stream(&["UnrelatedPayload"], &vec![0x5A; 8 * 1024])
        .expect("unrelated stream");
    let bytes = serialize_ole(&mut writer);
    let ranges = vec![
        fat_stream_range(&bytes, "\u{0005}SummaryInformation"),
        fat_stream_range(&bytes, "UnrelatedPayload"),
    ];
    (bytes, ranges)
}

fn assert_entry_limit(error: Error) {
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::EntryBytes,
            observed,
            maximum: maximum @ 4096,
        } if observed > maximum
    ));
}

fn assert_aggregate_limit(error: Error) {
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::AggregateBytes,
            observed,
            maximum: maximum @ 10_000,
        } if observed > maximum
    ));
}

#[test]
fn every_declared_metadata_or_unrelated_stream_is_entry_bounded_at_admission() {
    let limits = SourceLimits::builder()
        .max_entry_bytes(4096)
        .unwrap()
        .build()
        .unwrap();

    for stream_name in ["\u{0005}SummaryInformation", "UnrelatedPayload"] {
        let (bytes, range) = unrelated_stream_fixture(stream_name, 8 * 1024);

        let (reader, rejected) = RejectingReader::new(bytes.clone(), vec![range]);
        let error = match Reader::from_reader_with_options(
            reader,
            OpenOptions::new().source_limits(limits),
        ) {
            Ok(_) => panic!("oversized {stream_name} must fail eager admission"),
            Err(error) => error,
        };
        assert_entry_limit(error);
        assert!(
            !rejected.load(Ordering::Acquire),
            "eager admission read {stream_name} payload"
        );

        let (reader, rejected) = RejectingReader::new(bytes.clone(), vec![range]);
        let ole = OleFile::open(reader, litchi_cfb::OleFileLimits::default())
            .expect("directory indexing must not read the payload");
        let error =
            match Reader::from_ole_file_with_options(ole, OpenOptions::new().source_limits(limits))
            {
                Ok(_) => panic!("oversized {stream_name} must fail pre-parsed admission"),
                Err(error) => error,
            };
        assert_entry_limit(error);
        assert!(
            !rejected.load(Ordering::Acquire),
            "pre-parsed admission read {stream_name} payload"
        );

        let (source, rejected) = RejectingSource::new(bytes, vec![range]);
        let error = match SourceBackedPackage::from_read_at_with_options(
            source,
            OpenOptions::new().source_limits(limits),
        ) {
            Ok(_) => panic!("oversized {stream_name} must fail positional admission"),
            Err(error) => error,
        };
        assert_entry_limit(error);
        assert!(
            !rejected.load(Ordering::Acquire),
            "positional admission read {stream_name} payload"
        );
    }
}

#[test]
fn all_declared_streams_contribute_to_aggregate_admission_before_payload_reads() {
    let limits = SourceLimits::builder()
        .max_entry_bytes(16 * 1024)
        .unwrap()
        .max_aggregate_bytes(10_000)
        .unwrap()
        .build()
        .unwrap();
    let (bytes, ranges) = all_streams_fixture();

    let (reader, rejected) = RejectingReader::new(bytes.clone(), ranges.clone());
    let error =
        match Reader::from_reader_with_options(reader, OpenOptions::new().source_limits(limits)) {
            Ok(_) => panic!("aggregate stream overflow must fail eager admission"),
            Err(error) => error,
        };
    assert_aggregate_limit(error);
    assert!(!rejected.load(Ordering::Acquire));

    let (reader, rejected) = RejectingReader::new(bytes.clone(), ranges.clone());
    let ole = OleFile::open(reader, litchi_cfb::OleFileLimits::default())
        .expect("directory indexing must not read payloads");
    let error =
        match Reader::from_ole_file_with_options(ole, OpenOptions::new().source_limits(limits)) {
            Ok(_) => panic!("aggregate stream overflow must fail pre-parsed admission"),
            Err(error) => error,
        };
    assert_aggregate_limit(error);
    assert!(!rejected.load(Ordering::Acquire));

    let (source, rejected) = RejectingSource::new(bytes, ranges);
    let error = match SourceBackedPackage::from_read_at_with_options(
        source,
        OpenOptions::new().source_limits(limits),
    ) {
        Ok(_) => panic!("aggregate stream overflow must fail positional admission"),
        Err(error) => error,
    };
    assert_aggregate_limit(error);
    assert!(!rejected.load(Ordering::Acquire));
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
    let package = SourceBackedPackage::open(&path).unwrap();

    assert_eq!(package.len(), std::fs::metadata(&path).unwrap().len());
    assert!(!package.presentation().unwrap().has_pictures());
}

#[test]
fn positional_package_limit_is_typed_before_cfb_payload_reads() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let max_input_bytes = source.bytes.len() - 1;
    let limits = SourceLimits::builder()
        .max_input_bytes(max_input_bytes)
        .unwrap()
        .build()
        .unwrap();
    let error = match SourceBackedPackage::from_read_at_with_options(
        source.clone(),
        OpenOptions::new().source_limits(limits),
    ) {
        Ok(_) => panic!("package limit should reject the source"),
        Err(error) => error,
    };

    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum,
        } if observed == source.bytes.len() && maximum == max_input_bytes
    ));
    assert_eq!(source.reads.load(Ordering::Acquire), 0);
}

#[test]
fn positional_presentation_checks_stream_aggregate_before_materialization() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let after_open = source.reads.load(Ordering::Acquire);
    let limits = SourceLimits::builder()
        .max_aggregate_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error = match package.presentation_with_options(OpenOptions::new().source_limits(limits)) {
        Ok(_) => panic!("aggregate limit should reject the source"),
        Err(error) => error,
    };

    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::AggregateBytes,
            observed,
            maximum: maximum @ 1,
        } if observed > maximum
    ));
    assert_eq!(
        source.reads.load(Ordering::Acquire),
        after_open,
        "declared aggregate overflow must be rejected before stream payload reads"
    );
}

#[test]
fn positional_source_retains_tighter_parse_policy() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package_limit = SourceLimits::builder()
        .max_events(1)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at_with_options(
        source.clone(),
        OpenOptions::new().source_limits(package_limit),
    )
    .unwrap();
    assert_eq!(package.source_limits(), package_limit);

    let looser = SourceLimits::default();
    let error = match package.presentation_with_options(OpenOptions::new().source_limits(looser)) {
        Ok(_) => panic!("the retained package policy must remain tighter"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::Events,
            maximum: maximum @ 1,
            observed,
        } if observed > maximum
    ));
}

#[test]
fn positional_source_freshness_precedes_stricter_entry_projection() {
    let source = Arc::new(CountingSource::new(picture_fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let tighter = SourceLimits::builder()
        .max_entry_bytes(1)
        .unwrap()
        .build()
        .unwrap();

    let error = match package.presentation_with_options(OpenOptions::new().source_limits(tighter)) {
        Ok(_) => panic!("the stricter entry projection must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::EntryBytes,
            maximum: maximum @ 1,
            observed,
        } if observed > maximum
    ));

    source.revision.fetch_add(1, Ordering::AcqRel);
    let error = match package.presentation_with_options(OpenOptions::new().source_limits(tighter)) {
        Ok(_) => panic!("source freshness must be checked before payload policy"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Ole(litchi_cfb::OleError::SourceChanged { .. })
    ));
}

#[test]
fn open_options_uses_the_checked_source_policy_and_getters() {
    let limits = SourceLimits::builder()
        .max_entry_bytes(123)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new()
        .source_limits(limits)
        .password("secret")
        .leniency(Leniency::TolerateUserEditChainDefects);
    assert_eq!(options.source_limits_value(), limits);
    assert_eq!(options.password_value(), Some("secret"));
    assert_eq!(
        options.leniency_value(),
        Leniency::TolerateUserEditChainDefects
    );
}

#[test]
fn eager_ingress_enforces_input_and_entry_limits_with_exact_boundaries() {
    let bytes = picture_fixture();
    let exact = SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .build()
        .unwrap();
    let mut package = Reader::from_reader_with_options(
        Cursor::new(bytes.clone()),
        OpenOptions::new().source_limits(exact),
    )
    .expect("the complete input boundary is inclusive");
    assert_eq!(package.source_limits().max_input_bytes(), bytes.len());

    let too_small = SourceLimits::builder()
        .max_input_bytes(bytes.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = match Reader::from_reader_with_options(
        Cursor::new(bytes.clone()),
        OpenOptions::new().source_limits(too_small),
    ) {
        Ok(_) => panic!("complete input overflow must be typed"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum,
        } if observed == bytes.len() && maximum == bytes.len() - 1
    ));

    let error = match package.presentation_with_options(OpenOptions::new().source_limits(too_small))
    {
        Ok(_) => panic!("stricter projection input overflow must be typed"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum,
        } if observed == bytes.len() && maximum == bytes.len() - 1
    ));

    let entry_limit = SourceLimits::builder()
        .max_entry_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error =
        match package.presentation_with_options(OpenOptions::new().source_limits(entry_limit)) {
            Ok(_) => panic!("declared stream entry overflow must be typed"),
            Err(error) => error,
        };
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::EntryBytes,
            observed,
            maximum: maximum @ 1,
        } if observed > maximum
    ));
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
    let mut eager_package = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
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
