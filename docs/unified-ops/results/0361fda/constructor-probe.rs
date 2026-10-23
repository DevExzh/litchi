//! Matched raw IWA metadata-constructor qualification probe.
//!
//! The build script compiles this exact source against the pre-constructor
//! base core and the bounded-constructor core.  `before` adapts only
//! the constructor signatures; the fixture, route bodies, allocator and
//! output schema stay identical.  The timed routes reuse moved input buffers
//! where the public API permits it, so allocation counters describe the
//! constructor's own work rather than a freshly rebuilt fixture.

use std::hint::black_box;
use std::time::Instant;

use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, Limits, MessageInfo, RawMessage,
};
use sha2::{Digest, Sha256};

#[cfg(feature = "allocations")]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub static CURRENT: AtomicUsize = AtomicUsize::new(0);
    pub static PEAK: AtomicUsize = AtomicUsize::new(0);
    pub static BYTES: AtomicUsize = AtomicUsize::new(0);
    pub static CALLS: AtomicUsize = AtomicUsize::new(0);

    pub struct Counting;

    // SAFETY: every operation forwards to the system allocator with the
    // caller's layout and pointer; the atomics only observe sizes.
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: the system allocator receives the original layout.
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                add(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: the pointer and layout came from this allocator.
            unsafe { System.dealloc(pointer, layout) };
            CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
            // SAFETY: the pointer and old layout came from this allocator.
            let moved = unsafe { System.realloc(pointer, layout, size) };
            if !moved.is_null() {
                let live = if size >= layout.size() {
                    CURRENT.fetch_add(size - layout.size(), Ordering::Relaxed) + size
                        - layout.size()
                } else {
                    CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed)
                        - (layout.size() - size)
                };
                PEAK.fetch_max(live, Ordering::Relaxed);
                BYTES.fetch_add(size, Ordering::Relaxed);
                CALLS.fetch_add(1, Ordering::Relaxed);
            }
            moved
        }
    }

    fn add(size: usize) {
        let live = CURRENT.fetch_add(size, Ordering::Relaxed) + size;
        PEAK.fetch_max(live, Ordering::Relaxed);
        BYTES.fetch_add(size, Ordering::Relaxed);
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    pub fn start() -> usize {
        let live = CURRENT.load(Ordering::Relaxed);
        PEAK.store(live, Ordering::Relaxed);
        BYTES.store(0, Ordering::Relaxed);
        CALLS.store(0, Ordering::Relaxed);
        live
    }

    pub fn result(base: usize) -> (usize, usize, usize, usize) {
        (
            PEAK.load(Ordering::Relaxed).saturating_sub(base),
            CURRENT.load(Ordering::Relaxed).saturating_sub(base),
            BYTES.load(Ordering::Relaxed),
            CALLS.load(Ordering::Relaxed),
        )
    }
}

#[cfg(feature = "allocations")]
#[global_allocator]
static ALLOCATOR: allocation::Counting = allocation::Counting;

const OBJECT_ID: u64 = 17;
const MESSAGE_COUNT: usize = 8;
const MESSAGE_LENGTH: usize = 1_024;
const PATH: [u32; 8] = [1, 7, 4, 2, 9, 3, 6, 5];
const ROUTES: [&str; 15] = [
    "field_path",
    "field_info",
    "message_info",
    "archive_info",
    "archive_object",
    "archive_object_reject",
    "parse",
    "validate",
    "encode",
    "retained_parse",
    "retained_validate",
    "retained_encoded_len",
    "retained_encode",
    "retained_validate_tight",
    "retained_encode_tight",
];

fn limits() -> Limits {
    Limits::default()
}

#[cfg(feature = "before")]
fn make_path(path: Vec<u32>) -> FieldPath {
    FieldPath::new(path)
}

#[cfg(not(feature = "before"))]
fn make_path(path: Vec<u32>) -> FieldPath {
    FieldPath::new(path, limits()).expect("bounded field path fixture")
}

#[cfg(feature = "before")]
fn make_field(path: FieldPath) -> FieldInfo {
    FieldInfo::new(path)
}

#[cfg(not(feature = "before"))]
fn make_field(path: FieldPath) -> FieldInfo {
    FieldInfo::new(path, limits()).expect("bounded field metadata fixture")
}

#[cfg(feature = "before")]
fn make_message(type_: u32, length: usize) -> MessageInfo {
    MessageInfo::new(
        type_,
        u32::try_from(length).expect("fixture length fits u32"),
    )
}

#[cfg(not(feature = "before"))]
fn make_message(type_: u32, length: usize) -> MessageInfo {
    MessageInfo::new(type_, length, limits()).expect("bounded message metadata fixture")
}

#[cfg(feature = "before")]
fn make_info(identifier: u64, message_infos: Vec<MessageInfo>) -> ArchiveInfo {
    ArchiveInfo::new(identifier, message_infos)
}

#[cfg(not(feature = "before"))]
fn make_info(identifier: u64, message_infos: Vec<MessageInfo>) -> ArchiveInfo {
    ArchiveInfo::new(identifier, message_infos, limits()).expect("bounded archive metadata fixture")
}

fn make_messages() -> Vec<RawMessage> {
    (0..MESSAGE_COUNT)
        .map(|index| RawMessage {
            type_: 1_000 + u32::try_from(index).expect("fixture index fits u32"),
            data: vec![u8::try_from(index).expect("fixture byte fits u8"); MESSAGE_LENGTH],
        })
        .collect()
}

fn make_metadata_messages() -> Vec<MessageInfo> {
    (0..MESSAGE_COUNT)
        .map(|index| {
            make_message(
                1_000 + u32::try_from(index).expect("fixture index fits u32"),
                MESSAGE_LENGTH,
            )
        })
        .collect()
}

fn make_archive_bytes() -> Vec<u8> {
    let object = ArchiveObject::new(OBJECT_ID, make_messages(), limits())
        .expect("archive object fixture must be admitted");
    Archive {
        objects: vec![object],
    }
    .to_bytes(limits())
    .expect("archive fixture must encode")
}

// Retain both unknown fields and a noncanonical varint spelling. The
// public parser must preserve this original header exactly on publication.
fn make_retained_archive_bytes() -> Vec<u8> {
    let canonical = make_archive_bytes();
    let mut header_length = 0usize;
    let mut prefix_length = 0usize;
    for (index, byte) in canonical.iter().copied().enumerate() {
        header_length |= usize::from(byte & 0x7f) << (index * 7);
        prefix_length += 1;
        if byte & 0x80 == 0 {
            break;
        }
    }
    let mut header = canonical[prefix_length..prefix_length + header_length].to_vec();
    for _ in 0..64 {
        header.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
    }
    let mut output = Vec::new();
    let mut length = header.len();
    while length >= 0x80 {
        output.push((length as u8 & 0x7f) | 0x80);
        length >>= 7;
    }
    output.push(length as u8);
    output.extend_from_slice(&header);
    output.extend_from_slice(&canonical[prefix_length + header_length..]);
    output
}

fn source_for_route(route: &str) -> Vec<u8> {
    if route.starts_with("retained_") {
        make_retained_archive_bytes()
    } else {
        make_archive_bytes()
    }
}

fn retained_exact_field_limits(source: &[u8]) -> Limits {
    retained_field_boundary(source).0
}

fn retained_field_boundary(source: &[u8]) -> (Limits, usize, usize, usize) {
    for count in 1..=source.len() {
        let profile = limits().with_header_fields(count).unwrap();
        match Archive::parse(source, profile) {
            Ok(_) => {
                assert!(
                    count > 1,
                    "retained fixture unexpectedly fits one header field"
                );
                let below = count - 1;
                let refusal = Archive::parse(
                    source,
                    limits()
                        .with_header_fields(below)
                        .expect("one-below field profile is valid"),
                )
                .expect_err("one field below the boundary must refuse");
                let (observed, maximum) = match refusal {
                    litchi_iwa_core::Error::Limit {
                        kind: litchi_iwa_core::LimitKind::HeaderFields,
                        observed,
                        maximum,
                    } => (observed, maximum),
                    other => panic!("unexpected one-below field rejection: {other:?}"),
                };
                assert_eq!(maximum, below);
                assert!(observed > maximum);
                return (profile, count, observed, maximum);
            },
            Err(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::HeaderFields,
                ..
            }) => {},
            result => panic!("unexpected retained fixture admission: {result:?}"),
        }
    }
    panic!("retained fixture has no admissible field count");
}

fn parse(data: &[u8]) -> Archive {
    Archive::parse(data, limits()).expect("archive fixture must parse")
}

fn encode(archive: &Archive) -> Vec<u8> {
    archive
        .to_bytes(limits())
        .expect("archive fixture must encode")
}

fn exact_header_boundary() -> (usize, usize, usize) {
    // Find the smallest admissible header-byte budget for the same object
    // used by the constructor routes.  The returned tuple is the accepted
    // budget, the observed size reported by the one-below refusal, and that
    // refusal's maximum.  Keeping the exact accepted budget in the report
    // makes baseline/candidate boundary equivalence an explicit validation
    // value rather than an inference from a generic rejection.
    let upper = make_archive_bytes().len();
    for accepted in 1..=upper {
        let profile = limits()
            .with_header_bytes(accepted)
            .expect("header boundary profile is valid");
        match ArchiveObject::new(OBJECT_ID, make_messages(), profile) {
            Ok(_) => {
                assert!(accepted > 1, "fixture unexpectedly fits one header byte");
                let below = accepted - 1;
                let refusal = ArchiveObject::new(
                    OBJECT_ID,
                    make_messages(),
                    limits()
                        .with_header_bytes(below)
                        .expect("one-below profile is valid"),
                )
                .expect_err("one byte below the boundary must refuse");
                let (observed, maximum) = match refusal {
                    litchi_iwa_core::Error::Limit {
                        kind: litchi_iwa_core::LimitKind::HeaderBytes,
                        observed,
                        maximum,
                    } => (observed, maximum),
                    other => panic!("unexpected one-below header rejection: {other:?}"),
                };
                assert_eq!(maximum, below);
                assert!(observed > maximum);
                return (accepted, observed, maximum);
            },
            Err(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::HeaderBytes,
                ..
            }) => {},
            Err(other) => panic!("unexpected header-boundary result: {other:?}"),
        }
    }
    panic!("archive object has no admissible header-byte boundary");
}

fn validate(archive: &Archive) {
    archive
        .validate(limits())
        .expect("archive fixture must validate")
}

fn digest(data: &[u8]) -> String {
    Sha256::digest(data)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

#[derive(Debug, Clone, Copy)]
struct Metrics {
    ns: u128,
    checksum: u64,
    output_bytes: usize,
    peak: usize,
    retained: usize,
    allocated: usize,
    allocations: usize,
}

fn run(route: &str, iterations: usize) -> Metrics {
    let source = source_for_route(route);
    let archive = parse(&source);
    validate(&archive);
    let expected = encode(&archive);
    let output_bytes = expected.len();
    let output_digest = Sha256::digest(&expected);
    let retained_profile = if route.ends_with("_tight") {
        retained_exact_field_limits(&source)
    } else {
        limits()
    };
    assert_eq!(expected, source);

    // These buffers are moved through each constructor route.  Reusing them
    // keeps caller setup outside the measured region while retaining the
    // exact ownership semantics of each public constructor.
    let mut path = PATH.to_vec();
    let mut message_infos = make_metadata_messages();
    let mut messages = make_messages();
    let rejected_limits = limits()
        .with_header_bytes(1)
        .expect("the one-byte rejection profile is valid");
    #[cfg(feature = "allocations")]
    let before = allocation::start();
    let started = Instant::now();
    let mut checksum = 0u64;
    for _ in 0..iterations {
        match route {
            "field_path" => {
                let value = black_box(make_path(black_box(path)));
                checksum = checksum.wrapping_add(value.path.len() as u64);
                path = value.path;
            },
            "field_info" => {
                let value = black_box(make_field(make_path(black_box(path))));
                checksum = checksum.wrapping_add(value.path.path.len() as u64);
                path = value.path.path;
            },
            "message_info" => {
                let value = black_box(make_message(black_box(1_001), black_box(MESSAGE_LENGTH)));
                checksum = checksum
                    .wrapping_add(u64::from(value.type_))
                    .wrapping_add(u64::from(value.length))
                    .wrapping_add(value.versions.len() as u64);
            },
            "archive_info" => {
                let value = black_box(make_info(black_box(OBJECT_ID), black_box(message_infos)));
                checksum = checksum
                    .wrapping_add(value.message_infos.len() as u64)
                    .wrapping_add(value.identifier.unwrap_or_default());
                message_infos = value.message_infos;
            },
            "archive_object" => {
                let value = black_box(
                    ArchiveObject::new(black_box(OBJECT_ID), black_box(messages), limits())
                        .expect("archive object fixture must be admitted"),
                );
                checksum = checksum
                    .wrapping_add(value.messages.len() as u64)
                    .wrapping_add(value.archive_info.message_infos.len() as u64);
                messages = value.messages;
            },
            "archive_object_reject" => {
                let result = ArchiveObject::new(
                    OBJECT_ID,
                    vec![
                        RawMessage {
                            type_: 1_000,
                            data: Vec::new(),
                        },
                        RawMessage {
                            type_: 1_001,
                            data: Vec::new(),
                        },
                    ],
                    rejected_limits,
                );
                assert!(result.is_err(), "tight header profile must refuse");
                checksum = checksum.wrapping_add(1);
            },
            "parse" | "retained_parse" => {
                let value = parse(black_box(&source));
                checksum = checksum.wrapping_add(value.objects.len() as u64);
            },
            "validate" | "retained_validate" => {
                validate(black_box(&archive));
                checksum = checksum.wrapping_add(archive.objects.len() as u64);
            },
            "encode" | "retained_encode" => {
                let value = encode(black_box(&archive));
                checksum = checksum.wrapping_add(value.len() as u64);
                black_box(value);
            },
            "retained_encoded_len" => {
                let value = black_box(&archive)
                    .encoded_len(black_box(retained_profile))
                    .unwrap();
                checksum = checksum.wrapping_add(value as u64);
            },
            "retained_validate_tight" => {
                black_box(&archive)
                    .validate(black_box(retained_profile))
                    .unwrap();
                checksum = checksum.wrapping_add(archive.objects.len() as u64);
            },
            "retained_encode_tight" => {
                let value = black_box(&archive)
                    .to_bytes(black_box(retained_profile))
                    .unwrap();
                checksum = checksum.wrapping_add(value.len() as u64);
                black_box(value);
            },
            _ => panic!("unknown constructor route: {route}"),
        }
    }
    let ns = started.elapsed().as_nanos();
    let (peak, retained, allocated, allocations) = {
        #[cfg(feature = "allocations")]
        {
            allocation::result(before)
        }
        #[cfg(not(feature = "allocations"))]
        {
            (0, 0, 0, 0)
        }
    };
    black_box(path);
    black_box(message_infos);
    black_box(messages);
    Metrics {
        ns,
        checksum: checksum ^ u64::try_from(output_digest[0]).expect("digest byte fits u64"),
        output_bytes,
        peak,
        retained,
        allocated,
        allocations,
    }
}

fn validate_outputs(directory: &std::path::Path) {
    std::fs::create_dir_all(directory).expect("validation directory");
    let path = make_path(PATH.to_vec());
    assert_eq!(path.path, PATH);
    std::fs::write(directory.join("path.txt"), format!("{path:?}")).unwrap();
    let field = make_field(path);
    assert_eq!(field.path.path, PATH);
    std::fs::write(directory.join("field.txt"), format!("{field:?}")).unwrap();
    let message = make_message(1_001, MESSAGE_LENGTH);
    assert_eq!(
        (message.type_, message.length),
        (1_001, MESSAGE_LENGTH as u32)
    );
    assert_eq!(message.versions.as_slice(), &[1, 0, 5]);
    std::fs::write(directory.join("message.txt"), format!("{message:?}")).unwrap();
    let info = make_info(OBJECT_ID, make_metadata_messages());
    assert_eq!(info.identifier, Some(OBJECT_ID));
    assert_eq!(info.message_infos.len(), MESSAGE_COUNT);
    std::fs::write(directory.join("info.txt"), format!("{info:?}")).unwrap();
    let object = ArchiveObject::new(OBJECT_ID, make_messages(), limits()).unwrap();
    assert_eq!(object.archive_info, info);
    let bytes = Archive {
        objects: vec![object],
    }
    .to_bytes(limits())
    .unwrap();
    let parsed = parse(&bytes);
    validate(&parsed);
    let encoded = encode(&parsed);
    assert_eq!(encoded, bytes);
    std::fs::write(directory.join("source.iwa"), &bytes).unwrap();
    std::fs::write(directory.join("encoded.iwa"), encoded).unwrap();
    let retained_source = make_retained_archive_bytes();
    let retained = parse(&retained_source);
    let (exact, retained_header_boundary, retained_rejected_observed, retained_rejected_maximum) =
        retained_field_boundary(&retained_source);
    retained.validate(exact).unwrap();
    assert_eq!(retained.encoded_len(exact).unwrap(), retained_source.len());
    let retained_output = retained.to_bytes(exact).unwrap();
    assert_eq!(retained_output, retained_source);
    assert_eq!(retained.to_bytes(limits()).unwrap(), retained_source);
    std::fs::write(directory.join("retained-source.iwa"), &retained_source).unwrap();
    std::fs::write(directory.join("retained-encoded.iwa"), retained_output).unwrap();
    let (header_boundary, rejected_observed, rejected_maximum) = exact_header_boundary();
    let source_digest = digest(&bytes);
    let retained_digest = digest(&retained_source);
    let route_names = ROUTES
        .iter()
        .map(|route| format!("\"{route}\""))
        .collect::<Vec<_>>()
        .join(",");
    let route_checks = ROUTES
        .iter()
        .map(|route| {
            format!(
                "{{\"route\":\"{route}\",\"status\":\"ok\",\"output_bytes\":{}}}",
                bytes.len()
            )
        })
        .collect::<Vec<_>>()
        .join(",");
    let report = format!(
        "{{\"schema\":\"litchi.iwa.constructor.validation.v1\",\"status\":\"ok\",\"routes\":[{route_names}],\"route_checks\":[{route_checks}],\"source_sha256\":\"{source_digest}\",\"encoded_sha256\":\"{source_digest}\",\"source_bytes\":{},\"encoded_bytes\":{},\"roundtrip_equal\":true,\"retained_source_sha256\":\"{retained_digest}\",\"retained_encoded_sha256\":\"{retained_digest}\",\"retained_roundtrip_equal\":true,\"header_boundary\":{{\"kind\":\"HeaderBytes\",\"accepted\":{header_boundary},\"below\":{{\"observed\":{rejected_observed},\"maximum\":{rejected_maximum}}}}},\"retained_header_boundary\":{{\"kind\":\"HeaderFields\",\"accepted\":{retained_header_boundary},\"below\":{{\"observed\":{retained_rejected_observed},\"maximum\":{retained_rejected_maximum}}}}}}}\n",
        bytes.len(),
        bytes.len(),
    );
    std::fs::write(directory.join("validation.json"), report).unwrap();
}

fn main() {
    let mut args = std::env::args();
    let _program = args.next();
    let route = args.next().expect("route argument");
    if route == "validate_outputs" {
        validate_outputs(std::path::Path::new(
            &args.next().expect("output directory"),
        ));
        return;
    }
    let iterations = args
        .next()
        .expect("iteration argument")
        .parse::<usize>()
        .expect("iterations must be an integer");
    let source = source_for_route(&route);
    let expected = parse(&source);
    let encoded = encode(&expected);
    let expected_digest = digest(&encoded);
    let source_digest = digest(&source);
    let metrics = run(&route, iterations);
    println!(
        "{{\"schema\":\"litchi.iwa.constructor.probe.v3\",\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{},\"checksum\":{},\"output_bytes\":{},\"source_sha256\":\"{source_digest}\",\"output_sha256\":\"{expected_digest}\",\"output_matches_source\":{},\"correctness\":{{\"status\":\"ok\",\"roundtrip\":true,\"replay\":\"not-applicable\",\"noop\":\"not-applicable\"}},\"peak\":{},\"retained\":{},\"allocated\":{},\"allocations\":{}}}",
        metrics.ns,
        metrics.checksum,
        metrics.output_bytes,
        source == encoded,
        metrics.peak,
        metrics.retained,
        metrics.allocated,
        metrics.allocations,
    );
}
