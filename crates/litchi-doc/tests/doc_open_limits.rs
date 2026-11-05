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

use litchi_cfb::consts::STGTY_STREAM;
use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::builder::Builder;
use litchi_doc::writer::Picture;
use litchi_doc::{Error, OpenOptions, Password, SourceLimitKind, SourceLimits, raw::Reader};
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
    let mut writer = Builder::new();
    writer.add_paragraph("bounded DOC read").unwrap();
    writer
        .insert_picture(Picture::new(std::fs::read(image).unwrap()).unwrap())
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn exact_limits(bytes: &[u8]) -> SourceLimits {
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut max_entry_bytes = 0_usize;
    let mut aggregate_bytes = 0_usize;
    for index in 0..ole.directory_entry_count() {
        let sid = u32::try_from(index).unwrap();
        let Some(entry) = ole.directory_entry_by_sid(sid) else {
            continue;
        };
        if entry.entry_type != STGTY_STREAM {
            continue;
        }
        let size = usize::try_from(entry.size).unwrap();
        max_entry_bytes = max_entry_bytes.max(size);
        aggregate_bytes = aggregate_bytes.checked_add(size).unwrap();
    }
    SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .max_entry_bytes(max_entry_bytes)
        .unwrap()
        .max_aggregate_bytes(aggregate_bytes)
        .unwrap()
        .build()
        .unwrap()
}

fn document_with_extra_stream(name: &str, payload: &[u8]) -> Vec<u8> {
    let mut source = OleFile::open(
        Cursor::new(document_bytes()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let streams = source
        .list_streams()
        .into_iter()
        .map(|path| {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = source.open_stream(&refs).unwrap();
            (path, data)
        })
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer.create_stream(&[name], payload).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn limits(input_bytes: usize, entry_bytes: usize, aggregate_bytes: usize) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(input_bytes)
        .unwrap()
        .max_entry_bytes(entry_bytes)
        .unwrap()
        .max_aggregate_bytes(aggregate_bytes)
        .unwrap()
        .build()
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

    let mut package =
        Reader::open_with_options(&path, OpenOptions::default().source_limits(exact)).unwrap();
    assert!(package.document().is_ok());
    let error = Reader::open_with_options(
        &path,
        OpenOptions::default().source_limits(limits(
            exact.max_input_bytes() - 1,
            exact.max_entry_bytes(),
            exact.max_aggregate_bytes(),
        )),
    )
    .err()
    .unwrap();
    std::fs::remove_file(path).unwrap();

    let Error::Limit {
        kind,
        observed,
        maximum,
    } = error
    else {
        panic!("expected package resource limit")
    };
    assert_eq!(kind, SourceLimitKind::InputBytes);
    assert_eq!(observed, u64::try_from(bytes.len()).unwrap());
    assert_eq!(maximum, u64::try_from(bytes.len() - 1).unwrap());

    let core_error: litchi_core::Error = Reader::from_reader_with_options(
        Cursor::new(bytes),
        OpenOptions::default().source_limits(limits(
            exact.max_input_bytes() - 1,
            exact.max_entry_bytes(),
            exact.max_aggregate_bytes(),
        )),
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
        u64::try_from(exact.max_input_bytes()).unwrap()
    );
    assert_eq!(
        core_limit.limit,
        u64::try_from(exact.max_input_bytes() - 1).unwrap()
    );
    assert_eq!(core_limit.scope.as_ref(), "DOC input bytes");
}

#[test]
fn default_limits_are_bounded_and_open_options_apply_them() {
    let defaults = SourceLimits::default();
    assert_eq!(defaults.max_input_bytes(), 512 * 1024 * 1024);
    assert_eq!(defaults.max_entry_bytes(), 512 * 1024 * 1024);
    assert_eq!(defaults.max_aggregate_bytes(), 2 * 1024 * 1024 * 1024);
    assert!(defaults.max_input_bytes() <= SourceLimits::MAX_INPUT_BYTES);
    assert!(defaults.max_entry_bytes() <= SourceLimits::MAX_ENTRY_BYTES);
    assert!(defaults.max_aggregate_bytes() <= SourceLimits::MAX_AGGREGATE_BYTES);

    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let path = temporary_path();
    std::fs::write(&path, &bytes).unwrap();
    let package =
        litchi_doc::Package::open_with_options(&path, OpenOptions::default().source_limits(exact))
            .unwrap();
    std::fs::remove_file(path).unwrap();
    assert_eq!(package.source_limits(), exact);
}

#[test]
fn typed_password_is_redacted_and_moves_into_open_options() {
    let password = Password::new("not-in-debug-output".to_owned());
    assert_eq!(format!("{password:?}"), "Password([REDACTED])");
    let options = OpenOptions::default().password(password);
    assert!(!format!("{options:?}").contains("not-in-debug-output"));
}

#[test]
fn stream_limit_accepts_exact_size_and_failed_read_is_retryable() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();

    let error = package
        .document_with_options(OpenOptions::default().source_limits(limits(
            exact.max_input_bytes(),
            exact.max_entry_bytes() - 1,
            exact.max_aggregate_bytes(),
        )))
        .err()
        .unwrap();
    let Error::Limit {
        kind,
        observed,
        maximum,
    } = error
    else {
        panic!("expected stream resource limit")
    };
    assert_eq!(kind, SourceLimitKind::EntryBytes);
    assert_eq!(observed, u64::try_from(exact.max_entry_bytes()).unwrap());
    assert_eq!(maximum, u64::try_from(exact.max_entry_bytes() - 1).unwrap());

    let document = package
        .document_with_options(OpenOptions::default().source_limits(exact))
        .unwrap();
    assert!(document.text().unwrap().contains("bounded DOC read"));
}

#[test]
fn aggregate_limit_accepts_exact_size_and_failed_read_is_atomic() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();

    let error = package
        .document_with_options(OpenOptions::default().source_limits(limits(
            exact.max_input_bytes(),
            exact.max_entry_bytes(),
            exact.max_aggregate_bytes() - 1,
        )))
        .err()
        .unwrap();
    let Error::Limit {
        kind,
        observed,
        maximum,
    } = error
    else {
        panic!("expected aggregate resource limit")
    };
    assert_eq!(kind, SourceLimitKind::AggregateBytes);
    assert_eq!(
        observed,
        u64::try_from(exact.max_aggregate_bytes()).unwrap()
    );
    assert_eq!(
        maximum,
        u64::try_from(exact.max_aggregate_bytes() - 1).unwrap()
    );

    let document = package
        .document_with_options(OpenOptions::default().source_limits(exact))
        .unwrap();
    assert!(document.text().unwrap().contains("bounded DOC read"));
}

#[test]
fn per_read_limits_cannot_widen_package_limits() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let package_limits = exact;
    let mut package = Reader::from_reader_with_options(
        Cursor::new(bytes),
        OpenOptions::default().source_limits(package_limits),
    )
    .unwrap();

    let error = package
        .document_with_options(OpenOptions::default().source_limits(limits(
            exact.max_input_bytes(),
            exact.max_entry_bytes(),
            exact.max_aggregate_bytes() - 1,
        )))
        .err()
        .unwrap();
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::AggregateBytes,
            ..
        }
    ));
    assert_eq!(package.source_limits(), package_limits);
}

#[test]
fn payload_io_failure_after_preflight_is_not_reported_as_a_missing_stream() {
    let fail_reads = Arc::new(AtomicBool::new(false));
    let reader = GatedReader {
        inner: Cursor::new(document_bytes()),
        fail_reads: Arc::clone(&fail_reads),
    };
    let mut package = Reader::from_reader(reader).unwrap();
    fail_reads.store(true, Ordering::Relaxed);

    let error = package.document().err().unwrap();
    assert!(matches!(error, Error::Ole(litchi_cfb::OleError::Io(_))));
}

#[test]
fn custom_limits_reject_values_above_every_safety_ceiling() {
    let cases = [
        (
            SourceLimits::builder()
                .max_input_bytes(SourceLimits::MAX_INPUT_BYTES + 1)
                .unwrap_err(),
            SourceLimitKind::InputBytes,
            SourceLimits::MAX_INPUT_BYTES,
        ),
        (
            SourceLimits::builder()
                .max_entry_bytes(SourceLimits::MAX_ENTRY_BYTES + 1)
                .unwrap_err(),
            SourceLimitKind::EntryBytes,
            SourceLimits::MAX_ENTRY_BYTES,
        ),
        (
            SourceLimits::builder()
                .max_aggregate_bytes(SourceLimits::MAX_AGGREGATE_BYTES + 1)
                .unwrap_err(),
            SourceLimitKind::AggregateBytes,
            SourceLimits::MAX_AGGREGATE_BYTES,
        ),
    ];

    for (error, resource, maximum) in cases {
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: actual_kind,
                value,
                maximum: actual_maximum,
            } if actual_kind == resource && value == maximum + 1 && actual_maximum == maximum
        ));
    }
}

#[test]
fn custom_limits_reject_zero_values() {
    let cases = [
        (
            SourceLimits::builder().max_input_bytes(0),
            SourceLimitKind::InputBytes,
        ),
        (
            SourceLimits::builder().max_entry_bytes(0),
            SourceLimitKind::EntryBytes,
        ),
        (
            SourceLimits::builder().max_aggregate_bytes(0),
            SourceLimitKind::AggregateBytes,
        ),
    ];
    for (result, kind) in cases {
        assert!(matches!(
            result,
            Err(Error::InvalidLimit {
                kind: actual_kind,
                value: 0,
                ..
            }) if actual_kind == kind
        ));
    }
}

#[test]
fn opaque_and_metadata_streams_are_bounded_before_materialization() {
    let payload = vec![0x5a; exact_limits(&document_bytes()).max_entry_bytes() + 4096];
    for name in ["Equation Native", "\u{0005}SummaryInformation"] {
        let bytes = document_with_extra_stream(name, &payload);
        let exact = exact_limits(&bytes);
        let selected = limits(
            exact.max_input_bytes(),
            payload.len() - 1,
            exact.max_aggregate_bytes(),
        );
        let errors = [
            litchi_doc::Package::from_bytes_with_options(
                &bytes,
                OpenOptions::default().source_limits(selected),
            )
            .unwrap_err(),
            litchi_doc::Package::from_ole_file_with_options(
                OleFile::open(
                    Cursor::new(bytes.clone()),
                    litchi_cfb::OleFileLimits::default(),
                )
                .unwrap(),
                OpenOptions::default().source_limits(selected),
            )
            .unwrap_err(),
            Reader::from_ole_file_with_options(
                OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap(),
                OpenOptions::default().source_limits(selected),
            )
            .unwrap_err(),
        ];
        for error in errors {
            assert!(matches!(error, Error::Limit {
                kind: SourceLimitKind::EntryBytes, observed, maximum,
            } if observed == u64::try_from(payload.len()).unwrap()
                && maximum == u64::try_from(payload.len() - 1).unwrap()));
        }
    }
}

#[test]
fn already_parsed_ole_enforces_exact_package_limit() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    assert!(
        Reader::from_ole_file_with_options(ole, OpenOptions::default().source_limits(exact))
            .is_ok()
    );

    let ole = OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let error = Reader::from_ole_file_with_options(
        ole,
        OpenOptions::default().source_limits(limits(
            exact.max_input_bytes() - 1,
            exact.max_entry_bytes(),
            exact.max_aggregate_bytes(),
        )),
    )
    .err()
    .unwrap();
    let Error::Limit {
        kind,
        observed,
        maximum,
    } = error
    else {
        panic!("expected package resource limit")
    };
    assert_eq!(kind, SourceLimitKind::InputBytes);
    assert_eq!(observed, u64::try_from(bytes.len()).unwrap());
    assert_eq!(maximum, u64::try_from(bytes.len() - 1).unwrap());
}

#[test]
fn fixed_buffer_counting_fallback_accepts_n_and_rejects_n_plus_one() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let exact_reader = CountFallbackReader {
        inner: Cursor::new(bytes.clone()),
        reject_next_end_seek: true,
    };
    assert!(
        Reader::from_reader_with_options(exact_reader, OpenOptions::default().source_limits(exact))
            .is_ok()
    );

    let limited = limits(
        exact.max_input_bytes() - 1,
        exact.max_entry_bytes(),
        exact.max_aggregate_bytes(),
    );
    let oversized_reader = CountFallbackReader {
        inner: Cursor::new(bytes),
        reject_next_end_seek: true,
    };
    let error = Reader::from_reader_with_options(
        oversized_reader,
        OpenOptions::default().source_limits(limited),
    )
    .err()
    .unwrap();
    let Error::Limit {
        kind,
        observed,
        maximum,
    } = error
    else {
        panic!("expected package resource limit")
    };
    assert_eq!(kind, SourceLimitKind::InputBytes);
    assert_eq!(observed, u64::try_from(exact.max_input_bytes()).unwrap());
    assert_eq!(maximum, u64::try_from(exact.max_input_bytes() - 1).unwrap());
}

#[test]
fn aggregate_refusal_precedes_any_document_payload_read() {
    let bytes = document_bytes();
    let exact = exact_limits(&bytes);
    let fail_reads = Arc::new(AtomicBool::new(false));
    let reader = GatedReader {
        inner: Cursor::new(bytes),
        fail_reads: Arc::clone(&fail_reads),
    };
    let mut package = Reader::from_reader(reader).unwrap();
    fail_reads.store(true, Ordering::Relaxed);
    let error = package
        .document_with_options(OpenOptions::default().source_limits(limits(
            exact.max_input_bytes(),
            exact.max_entry_bytes(),
            1,
        )))
        .err()
        .unwrap();
    assert!(
        matches!(
            error,
            Error::Limit {
                kind: SourceLimitKind::AggregateBytes,
                maximum: 1,
                ..
            }
        ),
        "the size gate must precede the gated payload I/O: {error:?}"
    );
}
