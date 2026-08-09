#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use litchi_cfb::OleFile;
use litchi_doc::writer::{Picture, Writer};
use litchi_doc::{Error, Limits, OpenOptions, Package, PackageOpenOptions, Password, ResourceKind};
use std::io::{self, Cursor, Read, Seek, SeekFrom};
use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

static NEXT_FILE: AtomicU64 = AtomicU64::new(0);

struct GatedReader {
    inner: Cursor<Vec<u8>>,
    fail_reads: Arc<AtomicBool>,
}

struct CountFallbackReader {
    inner: Cursor<Vec<u8>>,
    reject_next_end_seek: bool,
}

impl Read for CountFallbackReader {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        self.inner.read(buffer)
    }
}

impl Seek for CountFallbackReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        if self.reject_next_end_seek && matches!(position, SeekFrom::End(_)) {
            self.reject_next_end_seek = false;
            return Err(io::Error::new(
                io::ErrorKind::Unsupported,
                "injected one-shot end-seek refusal",
            ));
        }
        self.inner.seek(position)
    }
}

impl Read for GatedReader {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        if self.fail_reads.load(Ordering::Relaxed) {
            return Err(io::Error::other("injected DOC payload read failure"));
        }
        self.inner.read(buffer)
    }
}

impl Seek for GatedReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

fn document_bytes() -> Vec<u8> {
    let image =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/images/jpg/abstract4.jpg");
    let mut writer = Writer::new();
    writer.add_paragraph("bounded DOC read").unwrap();
    writer
        .insert_picture(Picture::new(std::fs::read(image).unwrap()).unwrap())
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn exact_limits(bytes: &[u8]) -> Limits {
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let word_size = usize::try_from(ole.stream_len(&["WordDocument"]).unwrap()).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = litchi_doc::parts::fib::FileInformationBlock::parse(&word).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_size = usize::try_from(ole.stream_len(&[table_name]).unwrap()).unwrap();
    let data_size = usize::try_from(ole.stream_len(&["Data"]).unwrap()).unwrap();
    Limits::try_new(
        bytes.len(),
        word_size.max(table_size).max(data_size),
        word_size
            .checked_add(table_size)
            .and_then(|value| value.checked_add(data_size))
            .unwrap(),
    )
    .unwrap()
}

fn temporary_path() -> PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-doc-limits-{}-{}.doc",
        std::process::id(),
        NEXT_FILE.fetch_add(1, Ordering::Relaxed)
    ))
}

#[test]
fn package_limit_accepts_exact_size_and_rejects_one_less() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let path = temporary_path();
    std::fs::write(&path, &bytes).unwrap();

    let mut package = Package::open_with_limits(&path, exact).unwrap();
    assert!(package.document().is_ok());
    let error = Package::open_with_limits(
        &path,
        exact
            .with_max_package_bytes(exact.max_package_bytes() - 1)
            .unwrap(),
    )
    .err()
    .unwrap();
    std::fs::remove_file(path).unwrap();

    let Error::ResourceLimit(limit) = error else {
        panic!("expected package resource limit")
    };
    assert_eq!(limit.resource(), ResourceKind::Package);
    assert_eq!(limit.actual(), u64::try_from(bytes.len()).unwrap());
    assert_eq!(limit.limit(), u64::try_from(bytes.len() - 1).unwrap());
    assert_eq!(limit.path(), None);

    let core_error: litchi_core::Error = Package::from_reader_with_limits(
        Cursor::new(bytes),
        exact
            .with_max_package_bytes(exact.max_package_bytes() - 1)
            .unwrap(),
    )
    .err()
    .unwrap()
    .into();
    let litchi_core::Error::ResourceLimit(core_limit) = core_error else {
        panic!("expected typed core resource limit")
    };
    assert_eq!(core_limit.resource, litchi_core::Resource::InputBytes);
    assert_eq!(
        core_limit.observed,
        u64::try_from(exact.max_package_bytes()).unwrap()
    );
    assert_eq!(
        core_limit.limit,
        u64::try_from(exact.max_package_bytes() - 1).unwrap()
    );
    assert_eq!(core_limit.scope.as_ref(), "DOC package");
}

#[test]
fn default_limits_are_bounded_and_package_open_options_apply_them() {
    let defaults = Limits::default();
    assert!(defaults.max_package_bytes() < Limits::MAX_PACKAGE_BYTES);
    assert!(defaults.max_input_bytes() < Limits::MAX_INPUT_BYTES);
    assert!(defaults.max_aggregate_input_bytes() < Limits::MAX_AGGREGATE_INPUT_BYTES);

    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let path = temporary_path();
    std::fs::write(&path, &bytes).unwrap();
    let package =
        Package::open_with(&path, PackageOpenOptions::default().with_limits(exact)).unwrap();
    std::fs::remove_file(path).unwrap();
    assert_eq!(package.limits(), exact);
}

#[test]
fn typed_password_is_redacted_and_moves_into_open_options() {
    let password = Password::new("not-in-debug-output".to_owned());
    assert_eq!(format!("{password:?}"), "Password([REDACTED])");
    let options = OpenOptions::default().with_password(password);
    assert!(!format!("{options:?}").contains("not-in-debug-output"));
}

#[test]
fn stream_limit_accepts_exact_size_and_failed_read_is_retryable() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();

    let error = package
        .document_with_limits(
            exact
                .with_max_input_bytes(exact.max_input_bytes() - 1)
                .unwrap(),
        )
        .err()
        .unwrap();
    let Error::ResourceLimit(limit) = error else {
        panic!("expected stream resource limit")
    };
    assert_eq!(limit.resource(), ResourceKind::Stream);
    assert_eq!(
        limit.actual(),
        u64::try_from(exact.max_input_bytes()).unwrap()
    );
    assert_eq!(
        limit.limit(),
        u64::try_from(exact.max_input_bytes() - 1).unwrap()
    );
    assert!(limit.path().is_some());

    let document = package
        .document_with_options_and_limits(OpenOptions::default(), exact)
        .unwrap();
    assert!(document.text().unwrap().contains("bounded DOC read"));
}

#[test]
fn aggregate_limit_accepts_exact_size_and_failed_read_is_atomic() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();

    let error = package
        .document_with_limits(
            exact
                .with_max_aggregate_input_bytes(exact.max_aggregate_input_bytes() - 1)
                .unwrap(),
        )
        .err()
        .unwrap();
    let Error::ResourceLimit(limit) = error else {
        panic!("expected aggregate resource limit")
    };
    assert_eq!(limit.resource(), ResourceKind::Aggregate);
    assert_eq!(
        limit.actual(),
        u64::try_from(exact.max_aggregate_input_bytes()).unwrap()
    );
    assert_eq!(
        limit.limit(),
        u64::try_from(exact.max_aggregate_input_bytes() - 1).unwrap()
    );
    assert_eq!(limit.path(), None);

    let document = package.document_with_limits(exact).unwrap();
    assert!(document.text().unwrap().contains("bounded DOC read"));
}

#[test]
fn per_read_limits_cannot_widen_package_limits() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let package_limits = exact
        .with_max_aggregate_input_bytes(exact.max_aggregate_input_bytes() - 1)
        .unwrap();
    let mut package = Package::from_reader_with_limits(Cursor::new(bytes), package_limits).unwrap();

    let error = package
        .document_with_limits(Limits::default())
        .err()
        .unwrap();
    assert!(
        matches!(error, Error::ResourceLimit(limit) if limit.resource() == ResourceKind::Aggregate)
    );
    assert_eq!(package.limits(), package_limits);
}

#[test]
fn payload_io_failure_after_preflight_is_not_reported_as_a_missing_stream() {
    let fail_reads = Arc::new(AtomicBool::new(false));
    let reader = GatedReader {
        inner: Cursor::new(document_bytes()),
        fail_reads: Arc::clone(&fail_reads),
    };
    let mut package = Package::from_reader(reader).unwrap();
    fail_reads.store(true, Ordering::Relaxed);

    let error = package.document().err().unwrap();
    assert!(matches!(error, Error::Ole(litchi_cfb::OleError::Io(_))));
}

#[test]
fn custom_limits_reject_values_above_every_safety_ceiling() {
    let defaults = Limits::default();
    let cases = [
        (
            defaults
                .with_max_package_bytes(Limits::MAX_PACKAGE_BYTES + 1)
                .unwrap_err(),
            ResourceKind::Package,
            Limits::MAX_PACKAGE_BYTES,
        ),
        (
            defaults
                .with_max_input_bytes(Limits::MAX_INPUT_BYTES + 1)
                .unwrap_err(),
            ResourceKind::Stream,
            Limits::MAX_INPUT_BYTES,
        ),
        (
            defaults
                .with_max_aggregate_input_bytes(Limits::MAX_AGGREGATE_INPUT_BYTES + 1)
                .unwrap_err(),
            ResourceKind::Aggregate,
            Limits::MAX_AGGREGATE_INPUT_BYTES,
        ),
    ];

    for (error, resource, maximum) in cases {
        assert_eq!(error.resource(), resource);
        assert_eq!(error.actual(), maximum + 1);
        assert_eq!(error.maximum(), maximum);
    }
}

#[test]
fn already_parsed_ole_enforces_exact_package_limit() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(Package::from_ole_file_with_limits(ole, exact).is_ok());

    let ole = OleFile::open(Cursor::new(&bytes)).unwrap();
    let error = Package::from_ole_file_with_limits(
        ole,
        exact
            .with_max_package_bytes(exact.max_package_bytes() - 1)
            .unwrap(),
    )
    .err()
    .unwrap();
    let Error::ResourceLimit(limit) = error else {
        panic!("expected package resource limit")
    };
    assert_eq!(limit.resource(), ResourceKind::Package);
    assert_eq!(limit.actual(), u64::try_from(bytes.len()).unwrap());
    assert_eq!(limit.limit(), u64::try_from(bytes.len() - 1).unwrap());
}

#[test]
fn fixed_buffer_counting_fallback_accepts_n_and_rejects_n_plus_one() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let exact_reader = CountFallbackReader {
        inner: Cursor::new(bytes.clone()),
        reject_next_end_seek: true,
    };
    assert!(Package::from_reader_with_limits(exact_reader, exact).is_ok());

    let limited = exact
        .with_max_package_bytes(exact.max_package_bytes() - 1)
        .unwrap();
    let oversized_reader = CountFallbackReader {
        inner: Cursor::new(bytes),
        reject_next_end_seek: true,
    };
    let error = Package::from_reader_with_limits(oversized_reader, limited)
        .err()
        .unwrap();
    let Error::ResourceLimit(limit) = error else {
        panic!("expected package resource limit")
    };
    assert_eq!(limit.resource(), ResourceKind::Package);
    assert_eq!(
        limit.actual(),
        u64::try_from(exact.max_package_bytes()).unwrap()
    );
    assert_eq!(
        limit.limit(),
        u64::try_from(exact.max_package_bytes() - 1).unwrap()
    );
}
