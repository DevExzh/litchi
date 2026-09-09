use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_ole_common::dataspaces::{MAX_STREAM_BYTES, read_stream, read_stream_with_limit};
use std::io::{self, Cursor, Read, Seek, SeekFrom};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

const STORAGE: &str = "\u{0006}DataSpaces";

struct CountingReader {
    inner: Cursor<Vec<u8>>,
    reads: Arc<AtomicUsize>,
}

impl Read for CountingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if !output.is_empty() {
            self.reads.fetch_add(1, Ordering::Relaxed);
        }
        self.inner.read(output)
    }
}

impl Seek for CountingReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

fn package_with_version(payload: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(&[STORAGE]).unwrap();
    writer
        .create_stream(&[STORAGE, "Version"], payload)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn package_with_oversized_version() -> Vec<u8> {
    let payload = vec![0u8; MAX_STREAM_BYTES + 1];
    package_with_version(&payload)
}

#[test]
fn rejects_oversized_dataspaces_stream_before_payload_io() {
    let reads = Arc::new(AtomicUsize::new(0));
    let source = CountingReader {
        inner: Cursor::new(package_with_oversized_version()),
        reads: Arc::clone(&reads),
    };
    let mut ole = OleFile::open(source).unwrap();
    reads.store(0, Ordering::Relaxed);

    assert!(matches!(
        read_stream(&mut ole, &[STORAGE, "Version"]),
        Err(OleError::LimitExceeded {
            resource: "DataSpaces stream bytes",
            observed,
            maximum,
        }) if observed == u64::try_from(MAX_STREAM_BYTES).unwrap() + 1
            && maximum == u64::try_from(MAX_STREAM_BYTES).unwrap()
    ));
    assert_eq!(reads.load(Ordering::Relaxed), 0);
}

#[test]
fn accepts_exact_and_under_dataspaces_stream_limits() {
    let payload = b"metadata";
    let reads = Arc::new(AtomicUsize::new(0));
    let source = CountingReader {
        inner: Cursor::new(package_with_version(payload)),
        reads: Arc::clone(&reads),
    };
    let mut ole = OleFile::open(source).unwrap();
    reads.store(0, Ordering::Relaxed);

    assert_eq!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], payload.len()).unwrap(),
        payload
    );
    assert!(reads.load(Ordering::Relaxed) > 0);

    reads.store(0, Ordering::Relaxed);
    assert_eq!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], payload.len() + 1).unwrap(),
        payload
    );
    assert!(reads.load(Ordering::Relaxed) > 0);

    reads.store(0, Ordering::Relaxed);
    assert_eq!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], MAX_STREAM_BYTES).unwrap(),
        payload
    );
    assert!(reads.load(Ordering::Relaxed) > 0);

    reads.store(0, Ordering::Relaxed);
    assert!(matches!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], payload.len() - 1),
        Err(OleError::LimitExceeded {
            resource: "DataSpaces stream bytes",
            observed,
            maximum,
        }) if observed == payload.len() as u64 && maximum == (payload.len() - 1) as u64
    ));
    assert_eq!(reads.load(Ordering::Relaxed), 0);
}

#[test]
fn rejects_zero_and_above_ceiling_limits_before_lookup() {
    let reads = Arc::new(AtomicUsize::new(0));
    let source = CountingReader {
        inner: Cursor::new(package_with_version(b"metadata")),
        reads: Arc::clone(&reads),
    };
    let mut ole = OleFile::open(source).unwrap();
    reads.store(0, Ordering::Relaxed);

    assert!(matches!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], 0),
        Err(OleError::InvalidLimit {
            resource: "DataSpaces stream bytes",
            value: 0,
            maximum,
        }) if maximum == MAX_STREAM_BYTES as u64
    ));
    assert_eq!(reads.load(Ordering::Relaxed), 0);

    assert!(matches!(
        read_stream_with_limit(&mut ole, &[STORAGE, "Version"], MAX_STREAM_BYTES + 1),
        Err(OleError::InvalidLimit {
            resource: "DataSpaces stream bytes",
            value,
            maximum,
        }) if value == (MAX_STREAM_BYTES + 1) as u64
            && maximum == MAX_STREAM_BYTES as u64
    ));
    assert_eq!(reads.load(Ordering::Relaxed), 0);
}
