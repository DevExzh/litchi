use std::{
    alloc::{GlobalAlloc, Layout, System},
    io::{Cursor, Read, Write},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

use litchi_ods::{
    data_style::{self, NumberBuilder, Patch as MetadataPatch},
    document::{
        DataStyleFamily, DataStyleOwner, DataStyleSelector, Snapshot, StyleGraphExtension,
    },
};

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MANIFEST: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

#[derive(Debug)]
struct Measurement {
    scale: usize,
    operation: &'static str,
    source_bytes: usize,
    result_bytes: usize,
    allocations: u64,
    deallocations: u64,
    requested_bytes: u64,
    released_bytes: u64,
    elapsed_us: u128,
    preserved_opaque: bool,
    exact_noop: bool,
}

fn raw_package(entries: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let mut manifest = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"{MANIFEST}\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>"
    );
    for (path, _, media_type) in entries {
        manifest.push_str("<manifest:file-entry manifest:full-path=\"");
        manifest.push_str(path);
        manifest.push_str("\" manifest:media-type=\"");
        manifest.push_str(media_type);
        manifest.push_str("\"/>");
    }
    manifest.push_str("</manifest:manifest>");
    let mut output = Cursor::new(Vec::new());
    let mut archive = zip::ZipWriter::new(&mut output);
    let stored = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored);
    archive.start_file("mimetype", stored).unwrap();
    archive
        .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .unwrap();
    archive.start_file("META-INF/manifest.xml", stored).unwrap();
    archive.write_all(manifest.as_bytes()).unwrap();
    for (path, bytes, _) in entries {
        archive.start_file(*path, stored).unwrap();
        archive.write_all(bytes).unwrap();
    }
    archive.finish().unwrap();
    output.into_inner()
}

fn package_member(source: &[u8], path: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(source)).unwrap();
    let mut member = archive.by_name(path).unwrap();
    let mut bytes = Vec::new();
    member.read_to_end(&mut bytes).unwrap();
    bytes
}

fn fixture(scale: usize) -> Vec<u8> {
    fixture_with_opaque(scale, true)
}

fn clean_fixture(scale: usize) -> Vec<u8> {
    fixture_with_opaque(scale, false)
}

fn fixture_with_opaque(scale: usize, opaque: bool) -> Vec<u8> {
    let mut automatic = String::new();
    for index in 0..scale {
        use std::fmt::Write as _;
        let opaque_body = if opaque {
            format!(
                "<number:future xmlns:future=\"urn:example:future\"><future:payload><![CDATA[opaque-{index}]]></future:payload></number:future><!--preserve-{index}-->"
            )
        } else {
            String::new()
        };
        writeln!(
            automatic,
            "<number:number-style style:name=\"Style{index:05}\" style:title=\"Original {index}\"><number:fraction number:min-numerator-digits=\"1\"/>{opaque_body}</number:number-style>"
        )
        .unwrap();
    }
    let content = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:style=\"{STYLE}\" xmlns:number=\"{NUMBER}\" xmlns:table=\"{TABLE}\" xmlns:text=\"{TEXT}\" office:version=\"1.3\"><office:automatic-styles><style:style style:name=\"Cell\" style:family=\"table-cell\" style:data-style-name=\"Style00000\"/>{automatic}</office:automatic-styles><office:body><office:spreadsheet><table:table table:name=\"Sheet1\"><table:table-row><table:table-cell table:style-name=\"Cell\"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"
    );
    let styles = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles xmlns:office=\"{OFFICE}\" xmlns:style=\"{STYLE}\" xmlns:number=\"{NUMBER}\" office:version=\"1.3\"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"
    );
    raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn reset_counters() {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
}

fn graph(name: &str) -> StyleGraphExtension {
    let style = NumberBuilder::fraction(name).unwrap().build().unwrap();
    let mut builder = data_style::Graph::builder();
    builder.number_style(style).unwrap();
    builder.build().unwrap()
}

fn measure<F>(scale: usize, operation: &'static str, source: &[u8], work: F) -> Measurement
where
    F: FnOnce(&[u8]) -> (usize, bool, bool),
{
    reset_counters();
    let started = Instant::now();
    let (result_bytes, preserved_opaque, exact_noop) = work(source);
    let elapsed_us = started.elapsed().as_micros();
    Measurement {
        scale,
        operation,
        source_bytes: source.len(),
        result_bytes,
        allocations: ALLOC_CALLS.load(Ordering::Relaxed),
        deallocations: DEALLOC_CALLS.load(Ordering::Relaxed),
        requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
        released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
        elapsed_us,
        preserved_opaque,
        exact_noop,
    }
}

fn print_measurement(value: &Measurement) {
    println!(
        "{{\"scale\":{},\"operation\":\"{}\",\"source_bytes\":{},\"result_bytes\":{},\"allocations\":{},\"deallocations\":{},\"requested_bytes\":{},\"released_bytes\":{},\"elapsed_us\":{},\"preserved_opaque\":{},\"exact_noop\":{}}}",
        value.scale,
        value.operation,
        value.source_bytes,
        value.result_bytes,
        value.allocations,
        value.deallocations,
        value.requested_bytes,
        value.released_bytes,
        value.elapsed_us,
        value.preserved_opaque,
        value.exact_noop,
    );
}

fn main() {
    for scale in [8usize, 512] {
        let source = fixture(scale);
        std::fs::write(format!("/var/tmp/ods-style-profile-fixture-{scale}.ods"), &source)
            .unwrap();
        let name = "Style00000";
        let selector = DataStyleSelector::automatic(name, DataStyleFamily::Number);
        let read = measure(scale, "read_catalog", &source, |bytes| {
            let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
            let entries = snapshot.data_styles(DataStyleOwner::ContentAutomatic).unwrap();
            (entries.len(), false, false)
        });
        print_measurement(&read);

        let metadata = measure(scale, "metadata_patch", &source, |bytes| {
            let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
            let mut edit = snapshot.edit();
            let patch = MetadataPatch::default().set_title("Changed").unwrap();
            edit.patch_data_style(selector, &patch).unwrap();
            let output = edit.as_bytes();
            let content = package_member(output, "content.xml");
            (
                output.len(),
                content.windows(b"opaque-0".len()).any(|part| part == b"opaque-0")
                    && content
                        .windows(b"preserve-0".len())
                        .any(|part| part == b"preserve-0"),
                false,
            )
        });
        print_measurement(&metadata);

        let insert = measure(scale, "graph_insert", &source, |bytes| {
            let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
            let mut edit = snapshot.edit();
            edit.put_extended_style_graph(&graph("AddedScientific"))
                .unwrap();
            (edit.as_bytes().len(), false, false)
        });
        print_measurement(&insert);

        let clean = clean_fixture(scale);
        let replace = measure(scale, "graph_replace", &clean, |bytes| {
            let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
            let mut edit = snapshot.edit();
            edit.replace_extended_style_graph(selector, &graph(name))
                .unwrap();
            (edit.as_bytes().len(), false, false)
        });
        print_measurement(&replace);

        let noop = measure(scale, "exact_noop", &source, |bytes| {
            let snapshot = Snapshot::from_bytes(bytes.to_vec()).unwrap();
            let mut edit = snapshot.edit();
            edit.patch_data_style(selector, &MetadataPatch::default())
                .unwrap();
            let output = edit.as_bytes();
            (output.len(), false, output == bytes)
        });
        print_measurement(&noop);
    }
}
