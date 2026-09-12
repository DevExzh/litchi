use std::{
    alloc::{GlobalAlloc, Layout, System},
    collections::BTreeMap,
    env,
    hint::black_box,
    io::{Cursor, Read, Write},
    path::PathBuf,
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

use litchi_ods::{
    data_style::{self, NumberBuilder, Patch as DataStylePatch},
    document::{DataStyleFamily, DataStyleSelector, Snapshot, StyleGraphExtension},
};

type AnyResult<T> = Result<T, Box<dyn std::error::Error>>;

/// The observer is intentionally process-local.  It reports allocations made by
/// this harness and litchi-ods during one operation; it is not an allocator-
/// independent library contract.
struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let live =
            LIVE_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
        record_peak(live);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let old = layout.size() as u64;
        let new = size as u64;
        let live = if new >= old {
            LIVE_BYTES.fetch_add(new - old, Ordering::Relaxed) + (new - old)
        } else {
            LIVE_BYTES.fetch_sub(old - new, Ordering::Relaxed) - (old - new)
        };
        record_peak(live);
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

const REPEATS: usize = 7;
const SCALES: [usize; 3] = [8, 128, 512];

#[derive(Clone, Debug)]
struct WorkResult {
    result_bytes: usize,
    copy_bytes_observed: u64,
    exact_noop: bool,
}

#[derive(Clone, Debug)]
struct Measurement {
    scale: usize,
    operation: &'static str,
    iteration: usize,
    source_bytes: usize,
    result_bytes: usize,
    allocations: u64,
    deallocations: u64,
    requested_bytes: u64,
    released_bytes: u64,
    live_before_bytes: u64,
    live_after_bytes: u64,
    net_live_bytes_delta: i128,
    peak_live_bytes: u64,
    peak_live_bytes_delta: u64,
    copy_bytes_observed: u64,
    elapsed_ns: u128,
    exact_noop: bool,
}

fn record_peak(value: u64) {
    let mut current = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while value > current {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            current,
            value,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => current = observed,
        }
    }
}

fn reset_observer() -> u64 {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Relaxed);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
    live
}

fn measure_repeated<F>(
    scale: usize,
    operation: &'static str,
    source_bytes: usize,
    repeats: usize,
    mut work: F,
) -> AnyResult<Vec<Measurement>>
where
    F: FnMut() -> AnyResult<WorkResult>,
{
    // The warm-up exercises parser caches and allocator paths but is excluded
    // from receipts.  Correctness assertions are performed by this call.
    let _ = work()?;
    let mut rows = Vec::with_capacity(repeats);
    for iteration in 0..repeats {
        let live_before = reset_observer();
        let started = Instant::now();
        let result = work()?;
        let elapsed_ns = started.elapsed().as_nanos();
        black_box(result.result_bytes);
        let live_after = LIVE_BYTES.load(Ordering::Relaxed);
        let peak_live = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
        let net_live_bytes_delta = i128::from(live_after) - i128::from(live_before);
        rows.push(Measurement {
            scale,
            operation,
            iteration,
            source_bytes,
            result_bytes: result.result_bytes,
            allocations: ALLOC_CALLS.load(Ordering::Relaxed),
            deallocations: DEALLOC_CALLS.load(Ordering::Relaxed),
            requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
            released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
            live_before_bytes: live_before,
            live_after_bytes: live_after,
            net_live_bytes_delta,
            peak_live_bytes: peak_live,
            peak_live_bytes_delta: peak_live.saturating_sub(live_before),
            copy_bytes_observed: result.copy_bytes_observed,
            elapsed_ns,
            exact_noop: result.exact_noop,
        });
    }
    Ok(rows)
}

fn print_measurement(row: &Measurement) {
    println!(
        "{{\"kind\":\"measurement\",\"scale\":{},\"operation\":\"{}\",\"iteration\":{},\"source_bytes\":{},\"result_bytes\":{},\"allocations\":{},\"deallocations\":{},\"requested_bytes\":{},\"released_bytes\":{},\"live_before_bytes\":{},\"live_after_bytes\":{},\"net_live_bytes_delta\":{},\"peak_live_bytes\":{},\"peak_live_bytes_delta\":{},\"copy_bytes_observed\":{},\"elapsed_ns\":{},\"exact_noop\":{}}}",
        row.scale,
        row.operation,
        row.iteration,
        row.source_bytes,
        row.result_bytes,
        row.allocations,
        row.deallocations,
        row.requested_bytes,
        row.released_bytes,
        row.live_before_bytes,
        row.live_after_bytes,
        row.net_live_bytes_delta,
        row.peak_live_bytes,
        row.peak_live_bytes_delta,
        row.copy_bytes_observed,
        row.elapsed_ns,
        row.exact_noop,
    );
}

fn raw_package(entries: &[(&str, &[u8], &str)]) -> AnyResult<Vec<u8>> {
    let mut manifest = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<manifest:manifest xmlns:manifest=\"{MANIFEST}\" manifest:version=\"1.3\">\r\n<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>\r\n"
    );
    for (path, _, media_type) in entries {
        manifest.push_str("<manifest:file-entry manifest:full-path=\"");
        manifest.push_str(path);
        manifest.push_str("\" manifest:media-type=\"");
        manifest.push_str(media_type);
        manifest.push_str("\"/>\r\n");
    }
    manifest.push_str("</manifest:manifest>\r\n");

    let fixed_time = zip::DateTime::from_date_and_time(2020, 1, 2, 3, 4, 5)?;
    let stored = zip::write::SimpleFileOptions::default()
        .last_modified_time(fixed_time)
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o644);
    let deflated = zip::write::SimpleFileOptions::default()
        .last_modified_time(fixed_time)
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o644);
    let mut output = Cursor::new(Vec::new());
    let mut archive = zip::ZipWriter::new(&mut output);
    archive.start_file("mimetype", stored)?;
    archive.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;
    archive.start_file("META-INF/manifest.xml", stored)?;
    archive.write_all(manifest.as_bytes())?;
    for (path, bytes, _) in entries {
        archive.start_file(*path, deflated)?;
        archive.write_all(bytes)?;
    }
    archive.finish()?;
    Ok(output.into_inner())
}

fn style_markup(index: usize, opaque: bool, noncanonical_decimal: bool) -> String {
    let name = format!("Style{index:05}");
    let decimal_places = if noncanonical_decimal && index == 0 {
        "01"
    } else {
        "1"
    };
    // The replacement envelope intentionally accepts the compact typed shape
    // used by the graph fixture.  The preservation fixture keeps formatting
    // and foreign content around the same typed particle to exercise the
    // source-preserving metadata path.
    let mut style = if opaque {
        format!(
            "<number:number-style style:name=\"{name}\" number:title=\"Original {index}\">\r\n<number:fraction number:min-numerator-digits=\"1\"/>\r\n"
        )
    } else {
        format!(
            "<number:number-style style:name=\"{name}\"><number:number number:decimal-places=\"{decimal_places}\"/>"
        )
    };
    if opaque {
        style.push_str(&format!(
            "<number:future xmlns:number-future=\"urn:example:future\"><number-future:payload><![CDATA[opaque-{index}]]></number-future:payload></number:future><!--preserve-{index}-->\r\n"
        ));
    }
    if opaque {
        style.push_str("</number:number-style>\r\n");
    } else {
        style.push_str("</number:number-style>\r\n");
    }
    style
}

fn content_xml(
    scale: usize,
    opaque: bool,
    target_opaque: bool,
    noncanonical_decimal: bool,
) -> Vec<u8> {
    let mut automatic = String::new();
    for index in 0..scale {
        automatic.push_str(&style_markup(
            index,
            if index == 0 { target_opaque } else { opaque },
            noncanonical_decimal,
        ));
    }
    if !target_opaque {
        // This node is intentionally unreferenced and is used for the delete
        // leg of the graph CRUD profile.
        automatic.push_str(
            "<number:number-style style:name=\"DeleteStyle\"><number:number number:decimal-places=\"1\"/></number:number-style>\r\n",
        );
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<office:document-content xmlns:office=\"{OFFICE}\" xmlns:style=\"{STYLE}\" xmlns:number=\"{NUMBER}\" xmlns:table=\"{TABLE}\" xmlns:text=\"{TEXT}\" office:version=\"1.3\">\r\n<office:automatic-styles>\r\n<style:style style:name=\"Cell\" style:family=\"table-cell\" style:data-style-name=\"Style00000\"/>\r\n{automatic}</office:automatic-styles>\r\n<office:body><office:spreadsheet><table:table table:name=\"Sheet1\"><table:table-row><table:table-cell table:style-name=\"Cell\"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>\r\n</office:document-content>\r\n"
    )
    .into_bytes()
}

fn styles_xml() -> Vec<u8> {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<office:document-styles xmlns:office=\"{OFFICE}\" xmlns:style=\"{STYLE}\" xmlns:number=\"{NUMBER}\" office:version=\"1.3\">\r\n<office:styles><number:number-style style:name=\"CommonOnly\"><number:fraction number:min-numerator-digits=\"1\"/></number:number-style></office:styles>\r\n<office:automatic-styles/>\r\n<office:master-styles/>\r\n</office:document-styles>\r\n"
    )
    .into_bytes()
}

fn fixture(scale: usize, target_opaque: bool, noncanonical_decimal: bool) -> AnyResult<Vec<u8>> {
    let content = content_xml(scale, target_opaque, target_opaque, noncanonical_decimal);
    let styles = styles_xml();
    let meta = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<office:document-meta xmlns:office=\"{OFFICE}\" office:version=\"1.3\"><office:meta/></office:document-meta>\r\n"
    )
    .into_bytes();
    let settings = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<office:document-settings xmlns:office=\"{OFFICE}\" office:version=\"1.3\"><office:settings/></office:document-settings>\r\n"
    )
    .into_bytes();
    raw_package(&[
        ("content.xml", &content, "text/xml"),
        ("styles.xml", &styles, "text/xml"),
        ("meta.xml", &meta, "text/xml"),
        ("settings.xml", &settings, "text/xml"),
        (
            "Extras/foreign.bin",
            b"unrelated-payload-v1\0\xff\x7f",
            "application/octet-stream",
        ),
    ])
}

fn member_bytes(source: &[u8], path: &str) -> AnyResult<Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(source))?;
    let mut member = archive.by_name(path)?;
    let mut bytes = Vec::new();
    member.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn source_member_map(source: &[u8]) -> AnyResult<BTreeMap<String, Vec<u8>>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(source))?;
    let mut result = BTreeMap::new();
    for index in 0..archive.len() {
        let mut member = archive.by_index(index)?;
        let name = member.name().to_string();
        let mut bytes = Vec::new();
        member.read_to_end(&mut bytes)?;
        result.insert(name, bytes);
    }
    Ok(result)
}

fn assert_contains(haystack: &[u8], needle: &[u8], label: &str) -> AnyResult<()> {
    if haystack
        .windows(needle.len())
        .any(|window| window == needle)
    {
        Ok(())
    } else {
        Err(format!("{label} is absent").into())
    }
}

fn assert_absent(haystack: &[u8], needle: &[u8], label: &str) -> AnyResult<()> {
    if haystack
        .windows(needle.len())
        .any(|window| window == needle)
    {
        Err(format!("{label} is unexpectedly present").into())
    } else {
        Ok(())
    }
}

fn selector(name: &str) -> DataStyleSelector<'_> {
    DataStyleSelector::automatic(name, DataStyleFamily::Number)
}

fn graph_put() -> AnyResult<StyleGraphExtension> {
    let number = NumberBuilder::scientific("AddedScientific")?
        .decimal_places(3)?
        .build()?;
    let data = data_style::DataBuilder::percentage("AddedPercentage")?
        .decimal_places(2)?
        .build()?;
    let mut builder = data_style::Graph::builder();
    builder.number_style(number)?;
    builder.data_style(data)?;
    Ok(builder.build()?)
}

fn graph_replace() -> AnyResult<StyleGraphExtension> {
    let number = NumberBuilder::decimal("Style00000")?
        .decimal_places(2)?
        .build()?;
    let mut builder = data_style::Graph::builder();
    builder.number_style(number)?;
    Ok(builder.build()?)
}

fn graph_replace_same_value() -> AnyResult<StyleGraphExtension> {
    let number = NumberBuilder::decimal("Style00000")?
        .decimal_places(1)?
        .build()?;
    let mut builder = data_style::Graph::builder();
    builder.number_style(number)?;
    Ok(builder.build()?)
}

fn check_scalar_inverse(source: &[u8], patch: &DataStylePatch) -> AnyResult<()> {
    let target = selector("Style00000");
    let snapshot = Snapshot::from_bytes(source.to_vec())?;
    let mut changed = snapshot.edit();
    changed.patch_data_style(target, patch)?;
    let changed_content = member_bytes(changed.as_bytes(), "content.xml")?;
    assert_contains(
        &changed_content,
        b"number:title=\"Changed title\"",
        "scalar patch title",
    )?;
    assert_contains(&changed_content, b"opaque-0", "scalar patch opaque child")?;
    assert_contains(&changed_content, b"preserve-0", "scalar patch comment")?;

    let changed_snapshot = Snapshot::from_bytes(changed.as_bytes().to_vec())?;
    let reverse = DataStylePatch::default().set_title("Original 0")?;
    let mut restored = changed_snapshot.edit();
    restored.patch_data_style(target, &reverse)?;
    let source_members = source_member_map(source)?;
    let restored_members = source_member_map(restored.as_bytes())?;
    // The package writer is allowed to regenerate manifest XML (including
    // harmless lexical formatting and the root entry's inherited version),
    // so compare every payload member except that generated manifest.  The
    // source-qualified content owner must still be byte-identical after the
    // inverse, and the unrelated binary payload is explicitly checked.
    for (name, source_bytes) in &source_members {
        if name == "META-INF/manifest.xml" {
            continue;
        }
        if restored_members.get(name) != Some(source_bytes) {
            return Err(format!("scalar patch inverse changed package member {name}").into());
        }
    }
    if source_members.keys().collect::<Vec<_>>() != restored_members.keys().collect::<Vec<_>>() {
        return Err("scalar patch inverse changed package member inventory".into());
    }
    Ok(())
}

fn profile_scale(
    scale: usize,
    opaque_source: &[u8],
    graph_source: &[u8],
    graph_same_source: &[u8],
    repeats: usize,
) -> AnyResult<Vec<Measurement>> {
    let target = selector("Style00000");
    let no_op = DataStylePatch::default();
    let changed = DataStylePatch::default().set_title("Changed title")?;
    let put_graph = graph_put()?;
    let replace_graph = graph_replace()?;
    let same_value_graph = graph_replace_same_value()?;
    let remove_names = vec![String::from("DeleteStyle")];
    let mut rows = Vec::new();

    rows.extend(measure_repeated(
        scale,
        "source_query",
        opaque_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(opaque_source.to_vec())?;
            let value = snapshot
                .effective_cell_data_style("Cell")?
                .ok_or("effective cell data-style was absent")?;
            if value.name() != "Style00000" || value.family() != DataStyleFamily::Number {
                return Err(
                    "effective source-qualified data-style query returned the wrong node".into(),
                );
            }
            Ok(WorkResult {
                result_bytes: value.name().len(),
                copy_bytes_observed: opaque_source.len() as u64,
                exact_noop: false,
            })
        },
    )?);

    let clone_source = Snapshot::from_bytes(opaque_source.to_vec())?;
    rows.extend(measure_repeated(
        scale,
        "snapshot_clone",
        opaque_source.len(),
        repeats,
        || {
            let mut checksum = 0usize;
            for _ in 0..10_000 {
                let clone = clone_source.clone();
                checksum ^= clone.as_bytes().len();
                black_box(clone.prepared_index_identity());
            }
            black_box(checksum);
            Ok(WorkResult {
                result_bytes: clone_source.as_bytes().len(),
                copy_bytes_observed: 0,
                exact_noop: false,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "metadata_noop",
        opaque_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(opaque_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.patch_data_style(target, &no_op)?;
            if edit.as_bytes() != opaque_source {
                return Err("metadata no-op changed exact package bytes".into());
            }
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: opaque_source.len() as u64,
                exact_noop: true,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "scalar_patch",
        opaque_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(opaque_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.patch_data_style(target, &changed)?;
            let content = member_bytes(edit.as_bytes(), "content.xml")?;
            assert_contains(
                &content,
                b"number:title=\"Changed title\"",
                "measured scalar title",
            )?;
            assert_contains(&content, b"opaque-0", "measured scalar opaque child")?;
            // The unrelated member is deliberately binary and may be deflated;
            // compare it after extraction rather than searching compressed bytes.
            let unrelated = member_bytes(edit.as_bytes(), "Extras/foreign.bin")?;
            if unrelated != b"unrelated-payload-v1\0\xff\x7f" {
                return Err("measured scalar unrelated payload changed".into());
            }
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: opaque_source.len() as u64,
                exact_noop: false,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "graph_put",
        opaque_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(opaque_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.put_extended_style_graph(&put_graph)?;
            let content = member_bytes(edit.as_bytes(), "content.xml")?;
            assert_contains(&content, b"AddedScientific", "graph put number style")?;
            assert_contains(&content, b"AddedPercentage", "graph put data style")?;
            assert_contains(&content, b"opaque-0", "graph put preserved opaque child")?;
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: opaque_source.len() as u64,
                exact_noop: false,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "graph_replace",
        graph_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(graph_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.replace_extended_style_graph(target, &replace_graph)?;
            let content = member_bytes(edit.as_bytes(), "content.xml")?;
            assert_contains(
                &content,
                b"number:decimal-places=\"2\"",
                "graph replacement body",
            )?;
            assert_contains(
                &content,
                b"style:name=\"Style00001\"",
                "graph replacement preserved unselected style",
            )?;
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: graph_source.len() as u64,
                exact_noop: false,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "graph_replace_same_value",
        graph_same_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(graph_same_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.replace_extended_style_graph(target, &same_value_graph)?;
            let content = member_bytes(edit.as_bytes(), "content.xml")?;
            assert_contains(
                &content,
                b"number:decimal-places=\"1\"",
                "same-value replacement canonical body",
            )?;
            assert_contains(
                &content,
                b"number:decimal-places=\"01\"",
                "same-value replacement preserved noncanonical lexical body",
            )?;
            assert_contains(
                &content,
                b"style:name=\"Style00001\"",
                "same-value preserved unselected style",
            )?;
            if edit.as_bytes() != graph_same_source {
                return Err("same-value replacement changed exact source bytes".into());
            }
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: graph_same_source.len() as u64,
                exact_noop: edit.as_bytes() == graph_same_source,
            })
        },
    )?);

    rows.extend(measure_repeated(
        scale,
        "graph_remove",
        graph_source.len(),
        repeats,
        || {
            let snapshot = Snapshot::from_bytes(graph_source.to_vec())?;
            let mut edit = snapshot.edit();
            edit.remove_automatic_styles(&remove_names)?;
            let content = member_bytes(edit.as_bytes(), "content.xml")?;
            assert_absent(&content, b"DeleteStyle", "graph removal target")?;
            assert_contains(
                &content,
                b"style:name=\"Style00001\"",
                "graph removal preserved unselected style",
            )?;
            Ok(WorkResult {
                result_bytes: edit.as_bytes().len(),
                copy_bytes_observed: graph_source.len() as u64,
                exact_noop: false,
            })
        },
    )?);

    // The warm-up of scalar_patch proves a complete package inverse.  Keep
    // that assertion explicit here as well so future harness edits cannot
    // accidentally remove it from the profile setup.
    check_scalar_inverse(opaque_source, &changed)?;
    Ok(rows)
}

fn output_dir() -> AnyResult<PathBuf> {
    let path = env::var_os("ODS_PROFILE_OUTPUT_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("/var/tmp/ods-data-style-source-profile-output"));
    std::fs::create_dir_all(&path)?;
    Ok(path)
}

fn main() -> AnyResult<()> {
    let output = output_dir()?;
    let repeats = env::var("ODS_PROFILE_REPEATS")
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|value| *value > 0)
        .unwrap_or(REPEATS);
    println!(
        "{{\"kind\":\"run\",\"repeats\":{},\"scales\":[8,128,512],\"output_dir\":\"{}\",\"allocator\":\"process-global-counting\",\"copy_scope\":\"explicit harness byte clones only\"}}",
        repeats,
        output.display()
    );

    for scale in SCALES {
        let opaque_source = fixture(scale, true, false)?;
        let graph_source = fixture(scale, false, false)?;
        let graph_same_source = fixture(scale, false, true)?;
        std::fs::write(
            output.join(format!("fixture-opaque-{scale}.ods")),
            &opaque_source,
        )?;
        std::fs::write(
            output.join(format!("fixture-graph-{scale}.ods")),
            &graph_source,
        )?;
        std::fs::write(
            output.join(format!("fixture-graph-same-value-{scale}.ods")),
            &graph_same_source,
        )?;
        for row in profile_scale(
            scale,
            &opaque_source,
            &graph_source,
            &graph_same_source,
            repeats,
        )? {
            print_measurement(&row);
        }
    }
    Ok(())
}
