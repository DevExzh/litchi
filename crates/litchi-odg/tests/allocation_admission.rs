#![allow(
    unsafe_code,
    reason = "the integration test observes allocation admission"
)]

use litchi_odg::{Drawing, FlatDrawing};
use soapberry_zip::office::StreamingArchiveWriter;
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::{
    Mutex,
    atomic::{AtomicBool, AtomicUsize, Ordering},
};

const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const COMMENT: &str = "x";
const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;

static OBSERVE: AtomicBool = AtomicBool::new(false);
static ALLOCATED: AtomicUsize = AtomicUsize::new(0);
static PROBE_LOCK: Mutex<()> = Mutex::new(());

struct AllocationProbe;

#[global_allocator]
static GLOBAL: AllocationProbe = AllocationProbe;

// SAFETY: every operation is delegated to the platform allocator; the probe
// only records requested sizes with atomics while the test window is active.
unsafe impl GlobalAlloc for AllocationProbe {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout, as required by GlobalAlloc.
        let pointer = unsafe { System.alloc(layout) };
        record(layout.size());
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout, as required by GlobalAlloc.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        record(layout.size());
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair is supplied by the allocator caller.
        unsafe { System.dealloc(pointer, layout) };
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        // SAFETY: the pointer/layout pair and requested size are supplied by the allocator caller.
        let pointer = unsafe { System.realloc(pointer, layout, size) };
        record(size.saturating_sub(layout.size()));
        pointer
    }
}

fn record(bytes: usize) {
    if OBSERVE.load(Ordering::Relaxed) {
        ALLOCATED.fetch_add(bytes, Ordering::Relaxed);
    }
}

fn raw_package(content: &str) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("mimetype", MIMETYPE).unwrap();
    writer
        .write_stored("content.xml", content.as_bytes())
        .unwrap();
    writer
        .write_stored("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn flat_document(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body><office:drawing><draw:page draw:name="Page 1">{body}</draw:page></office:drawing></office:body></office:document>"#
    )
    .into_bytes()
}

fn flat_document_with_root_attribute(attribute: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><office:document {attribute} xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body><office:drawing><draw:page draw:name="Page 1"/></office:drawing></office:body></office:document>"#
    )
    .into_bytes()
}

#[test]
fn oversized_ignored_event_does_not_get_owned_before_admission() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let comment = COMMENT.repeat(MAX_TEXT_BYTES + 1);
    let content = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><!--{comment}--><office:drawing><draw:page draw:name="Page 1"/></office:drawing></office:body></office:document-content>"#
    );
    let content_bytes = content.len();
    let package = raw_package(&content);

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = Drawing::from_bytes(package);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized ignored event was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(reason) if reason.contains("event") && reason.contains("limit"))
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < content_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "oversized event allocation probe observed {allocated} bytes for {content_bytes} input bytes"
    );
}

#[test]
fn packaged_public_path_rejects_oversized_unknown_prefix_before_resolution() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let prefix = "p".repeat(MAX_TEXT_BYTES + 1);
    let content = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:drawing><draw:page draw:name="Page 1"><{prefix}:metadata/></draw:page></office:drawing></office:body></office:document-content>"#
    );
    let content_bytes = content.len();
    let package = raw_package(&content);

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = Drawing::from_bytes(package);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized unknown package prefix was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(ref reason) if reason.contains("limit")),
        "unexpected unknown-prefix error: {error}"
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < content_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "unknown-prefix allocation probe observed {allocated} bytes for {content_bytes} input bytes"
    );
}

#[test]
fn packaged_public_path_rejects_oversized_namespace_declaration_before_registration() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let uri = "u".repeat(MAX_TEXT_BYTES + 1);
    let content = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:rect xmlns:oversized="{uri}"/></draw:page></office:drawing></office:body></office:document-content>"#
    );
    let content_bytes = content.len();
    let package = raw_package(&content);

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = Drawing::from_bytes(package);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized namespace declaration was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(ref reason) if reason.contains("limit")),
        "unexpected namespace-declaration error: {error}"
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < content_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "namespace-declaration allocation probe observed {allocated} bytes for {content_bytes} input bytes"
    );
}

#[test]
fn flat_public_path_rejects_oversized_attribute_before_decode_ownership() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let name = "x".repeat(MAX_TEXT_BYTES + 1);
    let input = flat_document(&format!("<draw:rect draw:name=\"{name}\"/>"));
    let input_bytes = input.len();

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = FlatDrawing::from_bytes(input);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized flat attribute was accepted"),
        Err(error) => error,
    };
    assert!(matches!(error, litchi_core::Error::InvalidFormat(reason) if reason.contains("limit")));
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < input_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "flat attribute allocation probe observed {allocated} bytes for {input_bytes} input bytes"
    );
}

#[test]
fn flat_public_path_rejects_oversized_text_before_decode_ownership() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let text = "x".repeat(MAX_TEXT_BYTES + 1);
    let input = flat_document(&format!("<draw:rect><text:p>{text}</text:p></draw:rect>"));
    let input_bytes = input.len();

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = FlatDrawing::from_bytes(input);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized flat text was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(reason) if reason.contains("event") && reason.contains("limit"))
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < input_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "flat text allocation probe observed {allocated} bytes for {input_bytes} input bytes"
    );
}

#[test]
fn flat_public_path_rejects_oversized_unknown_prefix_before_resolution() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let prefix = "p".repeat(MAX_TEXT_BYTES + 1);
    let input = flat_document(&format!("<draw:rect><{prefix}:metadata/></draw:rect>"));
    let input_bytes = input.len();

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = FlatDrawing::from_bytes(input);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized flat unknown prefix was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(ref reason) if reason.contains("limit")),
        "unexpected flat unknown-prefix error: {error}"
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < input_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "flat unknown-prefix allocation probe observed {allocated} bytes for {input_bytes} input bytes"
    );
}

#[test]
fn flat_public_path_rejects_oversized_namespace_declaration_before_registration() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    let uri = "u".repeat(MAX_TEXT_BYTES + 1);
    let input = flat_document_with_root_attribute(&format!(r#"xmlns:oversized="{uri}""#));
    let input_bytes = input.len();

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = FlatDrawing::from_bytes(input);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("oversized flat namespace declaration was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(ref reason) if reason.contains("limit")),
        "unexpected flat namespace-declaration error: {error}"
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated
            < input_bytes
                .saturating_mul(2)
                .saturating_add(8 * 1024 * 1024),
        "flat namespace-declaration allocation probe observed {allocated} bytes for {input_bytes} input bytes"
    );
}

#[test]
fn flat_public_path_rejects_deep_nesting_before_reader_stack_growth() {
    let _probe_guard = PROBE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner);
    const DEEP_LEVELS: usize = 8_000_000;
    let mut body = String::with_capacity(DEEP_LEVELS.saturating_mul(7));
    for _ in 0..DEEP_LEVELS {
        body.push_str("<x>");
    }
    for _ in 0..DEEP_LEVELS {
        body.push_str("</x>");
    }
    let input = flat_document(&body);
    drop(body);
    let input_bytes = input.len();

    ALLOCATED.store(0, Ordering::Relaxed);
    OBSERVE.store(true, Ordering::Relaxed);
    let result = FlatDrawing::from_bytes(input);
    OBSERVE.store(false, Ordering::Relaxed);

    let error = match result {
        Ok(_) => panic!("deep flat ODG nesting was accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(ref reason) if reason.contains("nesting") && reason.contains("limit")),
        "unexpected deep-nesting error: {error}"
    );
    let allocated = ALLOCATED.load(Ordering::Relaxed);
    assert!(
        allocated < 32 * 1024 * 1024,
        "deep-nesting allocation probe observed {allocated} bytes for {input_bytes} input bytes"
    );
}
