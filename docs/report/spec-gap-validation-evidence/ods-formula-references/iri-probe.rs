//! Independent allocation probe for the private RFC3987 lexical validator.
//! Compile from the repository root with rustc --edition=2024 --crate-name ods_iri_probe.
#[path = "../../../../crates/litchi-ods/src/codec/formula/reference/iri.rs"]
mod iri;

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::sync::atomic::{AtomicUsize, Ordering};

struct CountingAllocator;
static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

// Probe-only instrumentation; production reference parsing contains no unsafe code.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        unsafe { System.realloc(pointer, layout, size) }
    }
}
#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn main() {
    let cases = [
        ("", true),
        ("#", true),
        ("?q=\u{E000}#part", true),
        ("#\u{E000}", false),
        ("../O'Brien.ods", true),
        ("relative:rootless", true),
        ("1relative:rootless", false),
        ("one/two:three", true),
        ("https://例え.テスト/表.ods", true),
        ("//user:secret@[::1]:/path", true),
        ("//[vAb.a!$&'()*+,;=:_~]/", true),
        ("//[vAb.a@b]/", false),
        ("//[vAb.a%20]/", false),
        ("//[::ffff:192.0.2.1]/", true),
        ("//[::ffff:192.00.2.1]/", false),
        ("//[2001:db8:::1]/", false),
        ("//a@b@c/", false),
        ("//host:655360000/", true),
        ("//host:invalid/", false),
        ("//999.999.999.999/", true),
        ("/a%ff%00", true),
        ("/a%F", false),
        ("/a%XY", false),
        ("/a\\b", false),
        ("/a b", false),
        ("/a\u{A0}b", true),
        ("/\u{FDD0}", false),
        ("/\u{FFEF}", true),
        ("/\u{FFF0}", false),
        ("/\u{E0000}", false),
        ("/\u{E1000}", true),
        ("/\u{EFFFD}", true),
        ("/\u{EFFFE}", false),
        ("?\u{10FFFD}", true),
        ("?\u{10FFFE}", false),
        ("?a?b#c?d", true),
        ("?a#b#c", false),
    ];
    let long = "x".repeat(65_536);
    ALLOCATIONS.store(0, Ordering::Relaxed);
    let mut observations = 0usize;
    for _ in 0..1_000 {
        for &(value, expected) in &cases {
            let actual = iri::is_valid_iri_reference(black_box(value));
            assert_eq!(actual, expected, "IRI fixture {value:?}");
            observations += 1;
            black_box(actual);
        }
    }
    for _ in 0..16 {
        assert!(iri::is_valid_iri_reference(black_box(&long)));
        observations += 1;
    }
    let allocations = ALLOCATIONS.load(Ordering::Relaxed);
    assert_eq!(allocations, 0);
    println!("cases={} observations={observations} long_input_bytes={} allocations={allocations}", cases.len(), long.len());
}
