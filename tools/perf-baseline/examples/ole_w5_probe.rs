//! Paired native probe for the W5 OLE owner and user-defined-property waves.
//!
//! The probe deliberately keeps the fixture and timed operation boundaries
//! independent of the API revision.  On the pre-migration checkout, pass
//! `--no-default-features --features legacy-owner-api,legacy-property-api`;
//! for the intermediate owner/property pair pass
//! `--no-default-features --features current-owner-api,legacy-property-api`;
//! on the current checkout use the default `current-api` feature. The output
//! is one CSV row per measured sample so an external ABBA runner can own CPU
//! pinning, calibration, and source/binary provenance.

use std::error::Error as StdError;
use std::hint::black_box;
use std::io::Cursor;
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::tracked_revision::{RevisionKind, RevisionMetadata};
use litchi_doc::{Builder as DocBuilder, Package as DocPackage};
use litchi_ole_common::object::{Edit as ObjectEdit, Limits as ObjectLimits, Targets};
use litchi_ole_common::property_set::Value;
use litchi_ole_common::property_set::user_defined::{
    Edit as UserEdit, Hyperlink, Hyperlinks, Limits as UserLimits, LinkBase,
    Properties as UserProperties,
};
use litchi_ole_common::property_set::{CodePage, Section, USER_DEFINED_PROPERTIES_FMTID};
use litchi_xls::Package as XlsPackage;
use litchi_xls::cell_values::{
    Reference as XlsReference, Selector as XlsSelector, Value as XlsValue,
};

/// A deliberately opt-in allocator for the W5 retained-state evidence probe.
///
/// The native ABBA probe does not enable this feature.  Keeping the allocator
/// behind a feature makes the extra atomics and the reporting path impossible
/// to include accidentally in the timing binaries.  Reports are emitted after
/// the before/after snapshots, so their own formatting allocations become the
/// next operation's baseline rather than entering the preceding delta.
#[cfg(feature = "w5-counting-allocator")]
mod w5_counting_allocator {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    static TOTAL_CALLS: AtomicUsize = AtomicUsize::new(0);
    static TOTAL_BYTES: AtomicUsize = AtomicUsize::new(0);
    static LIVE_CALLS: AtomicUsize = AtomicUsize::new(0);
    static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
    static PEAK_CALLS: AtomicUsize = AtomicUsize::new(0);
    static PEAK_BYTES: AtomicUsize = AtomicUsize::new(0);

    pub struct CountingAllocator;

    #[derive(Clone, Copy)]
    pub struct Snapshot {
        pub total_calls: usize,
        pub total_bytes: usize,
        pub live_calls: usize,
        pub live_bytes: usize,
        pub peak_calls: usize,
        pub peak_bytes: usize,
    }

    #[global_allocator]
    static GLOBAL: CountingAllocator = CountingAllocator;

    #[inline]
    fn record_alloc(bytes: usize) {
        TOTAL_CALLS.fetch_add(1, Ordering::Relaxed);
        TOTAL_BYTES.fetch_add(bytes, Ordering::Relaxed);
        let live_calls = LIVE_CALLS.fetch_add(1, Ordering::Relaxed) + 1;
        let live_bytes = LIVE_BYTES.fetch_add(bytes, Ordering::Relaxed) + bytes;
        PEAK_CALLS.fetch_max(live_calls, Ordering::Relaxed);
        PEAK_BYTES.fetch_max(live_bytes, Ordering::Relaxed);
    }

    #[inline]
    fn record_realloc(old_bytes: usize, new_bytes: usize) {
        TOTAL_CALLS.fetch_add(1, Ordering::Relaxed);
        TOTAL_BYTES.fetch_add(new_bytes, Ordering::Relaxed);
        let live_bytes = if new_bytes >= old_bytes {
            LIVE_BYTES.fetch_add(new_bytes - old_bytes, Ordering::Relaxed) + (new_bytes - old_bytes)
        } else {
            LIVE_BYTES.fetch_sub(old_bytes - new_bytes, Ordering::Relaxed) - (old_bytes - new_bytes)
        };
        PEAK_CALLS.fetch_max(LIVE_CALLS.load(Ordering::Relaxed), Ordering::Relaxed);
        PEAK_BYTES.fetch_max(live_bytes, Ordering::Relaxed);
    }

    #[inline]
    fn record_dealloc(bytes: usize) {
        LIVE_CALLS.fetch_sub(1, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(bytes, Ordering::Relaxed);
    }

    // The allocator is used only by this opt-in example binary.  All calls to
    // System are kept in explicit unsafe blocks so the probe remains valid
    // under the workspace's strict unsafe-operation lint settings.
    unsafe impl GlobalAlloc for CountingAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                record_alloc(layout.size());
            }
            pointer
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc_zeroed(layout) };
            if !pointer.is_null() {
                record_alloc(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            unsafe { System.dealloc(pointer, layout) };
            record_dealloc(layout.size());
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            let new_pointer = unsafe { System.realloc(pointer, layout, new_size) };
            if !new_pointer.is_null() {
                record_realloc(layout.size(), new_size);
            }
            new_pointer
        }
    }

    pub fn snapshot() -> Snapshot {
        Snapshot {
            total_calls: TOTAL_CALLS.load(Ordering::Relaxed),
            total_bytes: TOTAL_BYTES.load(Ordering::Relaxed),
            live_calls: LIVE_CALLS.load(Ordering::Relaxed),
            live_bytes: LIVE_BYTES.load(Ordering::Relaxed),
            peak_calls: PEAK_CALLS.load(Ordering::Relaxed),
            peak_bytes: PEAK_BYTES.load(Ordering::Relaxed),
        }
    }

    pub fn begin_operation() -> Snapshot {
        let live_calls = LIVE_CALLS.load(Ordering::Relaxed);
        let live_bytes = LIVE_BYTES.load(Ordering::Relaxed);
        PEAK_CALLS.store(live_calls, Ordering::Relaxed);
        PEAK_BYTES.store(live_bytes, Ordering::Relaxed);
        snapshot()
    }

    pub fn print_header() {
        println!(
            "w5_alloc,phase,route,iteration,delta_calls,delta_bytes,live_calls,live_bytes,peak_calls,peak_bytes,peak_delta_calls,peak_delta_bytes,base_live_calls,base_live_bytes,base_peak_calls,base_peak_bytes"
        );
    }

    pub fn print_delta(
        phase: &str,
        route: &str,
        iteration: usize,
        before: Snapshot,
        after: Snapshot,
    ) {
        println!(
            "w5_alloc,{phase},{route},{iteration},{},{},{},{},{},{},{},{},{},{},{},{}",
            after.total_calls.saturating_sub(before.total_calls),
            after.total_bytes.saturating_sub(before.total_bytes),
            after.live_calls,
            after.live_bytes,
            after.peak_calls,
            after.peak_bytes,
            after.peak_calls.saturating_sub(before.live_calls),
            after.peak_bytes.saturating_sub(before.live_bytes),
            before.live_calls,
            before.live_bytes,
            before.peak_calls,
            before.peak_bytes,
        );
    }
}

#[cfg(feature = "current-owner-api")]
use litchi_ole_common::object::{OpenOptions, Package as ObjectPackage};

#[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
type OwnerSource = litchi_ole_common::object::Snapshot;
#[cfg(feature = "current-owner-api")]
type OwnerSource = litchi_ole_common::object::Package;

const TARGET_NAME: &str = "ole_w5_edit_target.bin";
const ORIGINAL: &[u8] = b"ole-w5-original";
const REPLACEMENT: &[u8] = b"ole-w5-replacement";
const SIBLINGS: usize = 64;
const DOC_HISTORY_STEPS: usize = 128;
const XLS_HISTORY_STEPS: usize = 128;
const DOC_REVISION_TEXT: &str = "ole-w5-tracked-revision ";
const XLS_EDIT_VALUE: f64 = 4242.0;

struct Corpus {
    archive: Vec<u8>,
    limits: ObjectLimits,
    target_path: Vec<String>,
    user_section: Section,
    user_limits: UserLimits,
    link_base: LinkBase,
    hyperlinks: Hyperlinks,
    doc_archive: Vec<u8>,
    xls_archive: Vec<u8>,
}

struct HistoryRun<T, P> {
    output: Vec<u8>,
    packages: Vec<T>,
    patches: Vec<P>,
}

fn main() -> Result<(), Box<dyn StdError>> {
    let (route, warmup, samples, batch) = arguments()?;
    let corpus = build_corpus()?;
    let route_name = route.as_str();
    let mut prepared_owner = None;
    let mut prepared_section = None;
    let mut prepared_doc = None;
    let mut prepared_xls = None;

    if route_name == "begin_edit" {
        prepared_owner = Some(owner_open_source(&corpus)?);
    }
    if route_name == "read" {
        prepared_section = Some(corpus.user_section.clone());
    }
    if matches!(route_name, "doc_tracked_revision" | "doc_history") {
        prepared_doc = Some(DocPackage::from_bytes(&corpus.doc_archive)?);
    }
    if matches!(route_name, "xls_cells" | "xls_cells_stage" | "xls_history") {
        prepared_xls = Some(XlsPackage::from_bytes(&corpus.xls_archive)?);
    }

    let expected_host_output = match route_name {
        "doc_tracked_revision" => Some(doc_tracked_revision_expected(
            prepared_doc
                .as_ref()
                .ok_or("DOC tracked-revision source was not prepared")?,
        )?),
        "xls_cells" | "xls_cells_stage" => Some(xls_cells_expected(
            prepared_xls
                .as_ref()
                .ok_or("XLS cell source was not prepared")?,
        )?),
        "doc_history" => Some(doc_history_expected(
            prepared_doc
                .as_ref()
                .ok_or("DOC history source was not prepared")?,
        )?),
        "xls_history" => Some(xls_history_expected(
            prepared_xls
                .as_ref()
                .ok_or("XLS history source was not prepared")?,
        )?),
        _ => None,
    };

    println!("route,sample,nanoseconds,output_bytes");
    #[cfg(feature = "w5-counting-allocator")]
    w5_counting_allocator::print_header();
    for iteration in 0..warmup.checked_add(samples).ok_or("sample count overflow")? {
        let started = Instant::now();
        #[cfg(feature = "w5-counting-allocator")]
        let allocation_before = w5_counting_allocator::begin_operation();
        let output_bytes = match route_name {
            "open" => {
                let editor = owner_open_editor(&corpus)?;
                black_box(editor.stream(&corpus.target_path));
                corpus.archive.len()
            },
            "begin_edit" => {
                let source = prepared_owner
                    .as_ref()
                    .ok_or("begin_edit source was not prepared")?;
                let editor = owner_start_edit(source)?;
                black_box(editor);
                corpus.archive.len()
            },
            "noop_commit" => {
                let editor = owner_open_editor(&corpus)?;
                let output = owner_commit_bytes(editor)?;
                verify_owner_output(&output, ORIGINAL)?;
                output.len()
            },
            "replace_stream" => {
                let mut editor = owner_open_editor(&corpus)?;
                editor.put_stream(&corpus.target_path, REPLACEMENT.to_vec())?;
                let output = owner_commit_bytes(editor)?;
                verify_owner_output(&output, REPLACEMENT)?;
                output.len()
            },
            "read" => {
                let section = prepared_section
                    .as_ref()
                    .ok_or("read section was not prepared")?;
                let mut output_bytes = 0;
                #[cfg(feature = "w5-counting-allocator")]
                let mut last_properties = None;
                for _ in 0..batch {
                    let properties = user_properties(section, corpus.user_limits)?;
                    black_box((properties.link_base()?, properties.hyperlinks()?));
                    output_bytes = section_size(section);
                    #[cfg(feature = "w5-counting-allocator")]
                    {
                        last_properties = Some(properties);
                    }
                }
                #[cfg(feature = "w5-counting-allocator")]
                black_box(&last_properties);
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                output_bytes
            },
            "edit" => {
                let mut output_bytes = 0;
                let mut final_section = None;
                for _ in 0..batch {
                    let mut section = corpus.user_section.clone();
                    let link_base = corpus.link_base.clone();
                    let hyperlinks = corpus.hyperlinks.clone();
                    let mut edit = user_editor(&mut section, corpus.user_limits)?;
                    edit.set_link_base(link_base)?;
                    edit.set_hyperlinks(hyperlinks)?;
                    black_box(&section);
                    output_bytes = section_size(&section);
                    final_section = Some(section);
                }
                let final_section = final_section.ok_or("property edit batch was empty")?;
                if final_section != corpus.user_section {
                    return Err("property edit output differs from preflight".into());
                }
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                output_bytes
            },
            "doc_tracked_revision" => {
                let source = prepared_doc
                    .as_ref()
                    .ok_or("DOC tracked-revision source was not prepared")?;
                let output = doc_tracked_revision_once(source)?.to_bytes()?;
                let elapsed = started.elapsed().as_nanos();
                let expected = expected_host_output
                    .as_ref()
                    .ok_or("DOC tracked-revision expected output was not prepared")?;
                if output != *expected {
                    return Err("DOC tracked-revision output differs from preflight".into());
                }
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                if iteration >= warmup {
                    println!(
                        "{route_name},{},{elapsed},{}",
                        iteration - warmup,
                        output.len()
                    );
                }
                continue;
            },
            "xls_cells" => {
                let source = prepared_xls
                    .as_ref()
                    .ok_or("XLS cell source was not prepared")?;
                let expected = expected_host_output
                    .as_ref()
                    .ok_or("XLS cell expected output was not prepared")?;
                let mut output_bytes = 0;
                #[cfg(feature = "w5-counting-allocator")]
                let mut last_output = None;
                for _ in 0..batch {
                    let output = xls_cells_once(source)?.to_bytes()?;
                    if output != *expected {
                        return Err("XLS cell output differs from preflight".into());
                    }
                    output_bytes = output.len();
                    #[cfg(feature = "w5-counting-allocator")]
                    {
                        last_output = Some(output);
                    }
                }
                #[cfg(feature = "w5-counting-allocator")]
                black_box(&last_output);
                let elapsed = started.elapsed().as_nanos();
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                if iteration >= warmup {
                    println!(
                        "{route_name},{},{elapsed},{output_bytes}",
                        iteration - warmup,
                    );
                }
                continue;
            },
            "xls_cells_stage" => {
                let source = prepared_xls
                    .as_ref()
                    .ok_or("XLS cell source was not prepared")?;
                let expected = expected_host_output
                    .as_ref()
                    .ok_or("XLS cell expected output was not prepared")?;
                let mut last_commit = None;
                for _ in 0..batch {
                    last_commit = Some(xls_cells_commit_once(source)?);
                }
                let elapsed = started.elapsed().as_nanos();
                let commit = last_commit.ok_or("XLS cell stage batch was empty")?;
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "stage",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                let output = commit.package().to_bytes()?;
                if output != *expected {
                    return Err("XLS cell stage output differs from preflight".into());
                }
                black_box(&commit);
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                if iteration >= warmup {
                    println!(
                        "{route_name},{},{elapsed},{}",
                        iteration - warmup,
                        output.len()
                    );
                }
                continue;
            },
            "doc_history" => {
                let source = prepared_doc
                    .as_ref()
                    .ok_or("DOC history source was not prepared")?;
                let run = doc_history_once(source)?;
                let elapsed = started.elapsed().as_nanos();
                let expected = expected_host_output
                    .as_ref()
                    .ok_or("DOC history expected output was not prepared")?;
                verify_doc_history(&run, source, expected)?;
                let output_bytes = run.output.len();
                black_box((&run.packages, &run.patches));
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                if iteration >= warmup {
                    println!(
                        "{route_name},{},{elapsed},{output_bytes}",
                        iteration - warmup
                    );
                }
                continue;
            },
            "xls_history" => {
                let source = prepared_xls
                    .as_ref()
                    .ok_or("XLS history source was not prepared")?;
                let run = xls_history_once(source)?;
                let elapsed = started.elapsed().as_nanos();
                let expected = expected_host_output
                    .as_ref()
                    .ok_or("XLS history expected output was not prepared")?;
                verify_xls_history(&run, source, expected)?;
                let output_bytes = run.output.len();
                black_box((&run.packages, &run.patches));
                #[cfg(feature = "w5-counting-allocator")]
                if iteration >= warmup {
                    w5_counting_allocator::print_delta(
                        "full",
                        route_name,
                        iteration - warmup,
                        allocation_before,
                        w5_counting_allocator::snapshot(),
                    );
                }
                if iteration >= warmup {
                    println!(
                        "{route_name},{},{elapsed},{output_bytes}",
                        iteration - warmup
                    );
                }
                continue;
            },
            _ => return Err(format!("unknown route {route_name:?}").into()),
        };
        let elapsed = started.elapsed().as_nanos();
        #[cfg(feature = "w5-counting-allocator")]
        if iteration >= warmup && !matches!(route_name, "read" | "edit") {
            w5_counting_allocator::print_delta(
                "full",
                route_name,
                iteration - warmup,
                allocation_before,
                w5_counting_allocator::snapshot(),
            );
        }
        if iteration >= warmup {
            println!(
                "{route_name},{},{elapsed},{output_bytes}",
                iteration - warmup
            );
        }
    }
    Ok(())
}

fn arguments() -> Result<(String, usize, usize, usize), Box<dyn StdError>> {
    let mut route = None;
    let mut warmup = 10usize;
    let mut samples = 100usize;
    let mut batch = 1usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--route" => route = Some(arguments.next().ok_or("--route needs a value")?),
            "--warmup" => warmup = arguments.next().ok_or("--warmup needs a value")?.parse()?,
            "--samples" => samples = arguments.next().ok_or("--samples needs a value")?.parse()?,
            "--batch" => batch = arguments.next().ok_or("--batch needs a value")?.parse()?,
            "--help" | "-h" => {
                println!(
                    "usage: ole_w5_probe --route <open|begin_edit|noop_commit|replace_stream|read|edit|doc_tracked_revision|xls_cells|xls_cells_stage|doc_history|xls_history> [--warmup N] [--samples N] [--batch N]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument {other:?}").into()),
        }
    }
    let route = route.ok_or("--route is required")?;
    if samples == 0 {
        return Err("--samples must be nonzero".into());
    }
    if batch == 0 {
        return Err("--batch must be nonzero".into());
    }
    Ok((route, warmup, samples, batch))
}

fn build_corpus() -> Result<Corpus, Box<dyn StdError>> {
    let mut writer = OleWriter::new();
    for index in 0..SIBLINGS {
        let name = format!("ole_w5_sibling_{index:03}");
        let payload = format!("opaque sibling payload {index:03}");
        writer.create_stream(&[name.as_str()], payload.as_bytes())?;
    }
    writer.create_stream(&[TARGET_NAME], ORIGINAL)?;
    let mut archive = Cursor::new(Vec::new());
    writer.write_to(&mut archive)?;

    let stream_count = u64::try_from(SIBLINGS + 1)?;
    let total_size = u64::try_from(archive.get_ref().len())?;
    let limits = ObjectLimits {
        max_objects: 1,
        max_storage_depth: 1,
        max_streams_per_object: usize::try_from(stream_count)?,
        max_streams: usize::try_from(stream_count)?,
        max_stream_size: total_size.max(64),
        max_object_size: total_size.saturating_add(4096),
        max_total_size: total_size.saturating_add(4096),
    };

    let mut user_section = Section::new(USER_DEFINED_PROPERTIES_FMTID);
    user_section.set_page(CodePage::Utf16Le);
    let user_limits = UserLimits::default();
    let link_base = LinkBase::new("https://example.invalid/base/")?;
    let hyperlinks = Hyperlinks::new(vec![Hyperlink::new(
        7,
        11,
        0,
        "https://example.invalid/target",
        "bookmark",
    )?]);
    {
        let mut edit = user_editor(&mut user_section, user_limits)?;
        edit.set_link_base(link_base.clone())?;
        edit.set_hyperlinks(hyperlinks.clone())?;
    }
    // Keep one unrelated, opaque property so the typed overlay must preserve
    // the section's existing content while replacing only its reserved names.
    user_section.add(42, Value::Blob(vec![0x5a; 128]))?;

    let doc_archive = build_doc_archive()?;
    let xls_archive = build_xls_archive()?;

    Ok(Corpus {
        archive: archive.into_inner(),
        limits,
        target_path: vec![TARGET_NAME.to_owned()],
        user_section,
        user_limits,
        link_base,
        hyperlinks,
        doc_archive,
        xls_archive,
    })
}

fn build_doc_archive() -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut builder = DocBuilder::new();
    for index in 0..256 {
        builder.add_paragraph(&format!("ole-w5 tracked paragraph {index:03}"))?;
    }
    Ok(builder.build()?.to_bytes()?)
}

fn build_xls_archive() -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Cells")?;
    for row in 0..256u32 {
        builder.write_number(sheet, row, 0, f64::from(row))?;
    }
    let mut package = Cursor::new(Vec::new());
    builder.build()?.write_to(&mut package)?;
    let mut source = OleFile::open(
        Cursor::new(package.into_inner()),
        litchi_cfb::OleFileLimits::default(),
    )?;
    let workbook = source.open_stream(&["Workbook"])?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], workbook)?;
    writer.create_stream_owned(&["ole_w5_opaque_payload.bin"], vec![0x5a; 96 * 1024])?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn doc_tracked_revision_once(source: &DocPackage) -> Result<DocPackage, Box<dyn StdError>> {
    let mut edit = source.edit_tracked_revisions()?;
    edit.add_text(
        0,
        DOC_REVISION_TEXT,
        RevisionKind::Insertion,
        RevisionMetadata::new("W5Probe"),
    )?;
    Ok(edit.commit()?.into_package())
}

fn doc_tracked_revision_expected(source: &DocPackage) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut edit = source.edit_tracked_revisions()?;
    edit.add_text(
        0,
        DOC_REVISION_TEXT,
        RevisionKind::Insertion,
        RevisionMetadata::new("W5Probe"),
    )?;
    let commit = edit.commit()?;
    let output = commit.package().to_bytes()?;
    let replayed = source
        .apply_tracked_revisions(commit.patch())?
        .into_package()
        .to_bytes()?;
    if replayed != output {
        return Err("DOC tracked-revision forward replay differs from commit".into());
    }
    let restored = commit
        .package()
        .apply_tracked_revisions(&commit.patch().inverse())?
        .into_package()
        .to_bytes()?;
    if restored != source.to_bytes()? {
        return Err("DOC tracked-revision inverse replay differs from source".into());
    }
    Ok(output)
}

fn xls_cells_once(source: &XlsPackage) -> Result<XlsPackage, Box<dyn StdError>> {
    Ok(xls_cells_commit_once(source)?.into_package())
}

fn xls_cells_commit_once(source: &XlsPackage) -> Result<litchi_xls::Commit, Box<dyn StdError>> {
    let mut edit = source.edit()?;
    edit.set_value(
        XlsSelector::Name("Cells"),
        XlsReference::new(0, 0)?,
        XlsValue::Number(XLS_EDIT_VALUE),
    )?;
    Ok(edit.commit()?)
}

fn xls_cells_expected(source: &XlsPackage) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut edit = source.edit()?;
    edit.set_value(
        XlsSelector::Name("Cells"),
        XlsReference::new(0, 0)?,
        XlsValue::Number(XLS_EDIT_VALUE),
    )?;
    let commit = edit.commit()?;
    let output = commit.package().to_bytes()?;
    let replayed = source.apply(commit.patch())?.package().to_bytes()?;
    if replayed != output {
        return Err("XLS cell forward replay differs from commit".into());
    }
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())?
        .package()
        .to_bytes()?;
    if restored != source.to_bytes()? {
        return Err("XLS cell inverse replay differs from source".into());
    }
    Ok(output)
}

fn doc_history_once(
    source: &DocPackage,
) -> Result<HistoryRun<DocPackage, litchi_doc::tracked_revision::Patch>, Box<dyn StdError>> {
    let mut current = source.clone();
    let mut packages = Vec::with_capacity(DOC_HISTORY_STEPS + 1);
    let mut patches = Vec::with_capacity(DOC_HISTORY_STEPS);
    packages.push(current.clone());
    for index in 0..DOC_HISTORY_STEPS {
        let mut edit = current.edit_tracked_revisions()?;
        if index == 0 {
            edit.add_text(
                0,
                &format!("{DOC_REVISION_TEXT}{index:03} "),
                RevisionKind::Insertion,
                RevisionMetadata::new("W5History"),
            )?;
        } else {
            let revisions = edit.revisions()?;
            if revisions.len() != 1 {
                return Err("DOC history fixture did not retain one revision".into());
            }
            edit.replace(0, RevisionMetadata::new(format!("W5History{index:03}")))?;
        }
        let commit = edit.commit()?;
        current = commit.package().clone();
        patches.push(commit.patch().clone());
        packages.push(current.clone());
    }
    let output = current.to_bytes()?;
    Ok(HistoryRun {
        output,
        packages,
        patches,
    })
}

fn verify_doc_history(
    run: &HistoryRun<DocPackage, litchi_doc::tracked_revision::Patch>,
    source: &DocPackage,
    expected: &[u8],
) -> Result<(), Box<dyn StdError>> {
    if run.output != expected
        || run.packages.len() != DOC_HISTORY_STEPS + 1
        || run.patches.len() != DOC_HISTORY_STEPS
    {
        return Err("DOC history output or retained cardinality differs from preflight".into());
    }
    let mut replayed = source.clone();
    for patch in &run.patches {
        replayed = replayed.apply_tracked_revisions(patch)?.into_package();
    }
    if replayed.to_bytes()? != expected {
        return Err("DOC history forward replay differs from retained output".into());
    }
    for patch in run.patches.iter().rev() {
        replayed = replayed
            .apply_tracked_revisions(&patch.inverse())?
            .into_package();
    }
    if replayed.to_bytes()? != source.to_bytes()? {
        return Err("DOC history inverse replay differs from source".into());
    }
    Ok(())
}

fn doc_history_expected(source: &DocPackage) -> Result<Vec<u8>, Box<dyn StdError>> {
    let run = doc_history_once(source)?;
    let output = run.output.clone();
    verify_doc_history(&run, source, &output)?;
    Ok(output)
}

fn xls_history_once(
    source: &XlsPackage,
) -> Result<HistoryRun<XlsPackage, litchi_xls::Patch>, Box<dyn StdError>> {
    let mut current = source.clone();
    let mut packages = Vec::with_capacity(XLS_HISTORY_STEPS + 1);
    let mut patches = Vec::with_capacity(XLS_HISTORY_STEPS);
    packages.push(current.clone());
    for row in 0..XLS_HISTORY_STEPS {
        let mut edit = current.edit()?;
        edit.set_value(
            XlsSelector::Name("Cells"),
            XlsReference::new(u32::try_from(row)?, 0)?,
            XlsValue::Number(10_000.0 + row as f64),
        )?;
        let commit = edit.commit()?;
        current = commit.package().clone();
        patches.push(commit.patch().clone());
        packages.push(current.clone());
    }
    let output = current.to_bytes()?;
    Ok(HistoryRun {
        output,
        packages,
        patches,
    })
}

fn verify_xls_history(
    run: &HistoryRun<XlsPackage, litchi_xls::Patch>,
    source: &XlsPackage,
    expected: &[u8],
) -> Result<(), Box<dyn StdError>> {
    if run.output != expected
        || run.packages.len() != XLS_HISTORY_STEPS + 1
        || run.patches.len() != XLS_HISTORY_STEPS
    {
        return Err("XLS history output or retained cardinality differs from preflight".into());
    }
    let mut replayed = source.clone();
    for patch in &run.patches {
        replayed = replayed.apply(patch)?.into_package();
    }
    if replayed.to_bytes()? != expected {
        return Err("XLS history forward replay differs from retained output".into());
    }
    for patch in run.patches.iter().rev() {
        replayed = replayed.apply(&patch.inverse())?.into_package();
    }
    if replayed.to_bytes()? != source.to_bytes()? {
        return Err("XLS history inverse replay differs from source".into());
    }
    Ok(())
}

fn xls_history_expected(source: &XlsPackage) -> Result<Vec<u8>, Box<dyn StdError>> {
    let run = xls_history_once(source)?;
    let output = run.output.clone();
    verify_xls_history(&run, source, &output)?;
    Ok(output)
}

fn owner_open_source(corpus: &Corpus) -> Result<OwnerSource, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(ObjectEdit::open(corpus.archive.clone(), Targets::default(), corpus.limits)?.staged())
    }
    #[cfg(feature = "current-owner-api")]
    {
        Ok(ObjectPackage::open_with_options(
            corpus.archive.clone(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(corpus.limits),
        )?)
    }
}

fn owner_open_editor(corpus: &Corpus) -> Result<ObjectEdit, Box<dyn StdError>> {
    let source = owner_open_source(corpus)?;
    owner_start_edit(&source)
}

fn owner_start_edit(source: &OwnerSource) -> Result<ObjectEdit, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(source.edit())
    }
    #[cfg(feature = "current-owner-api")]
    {
        Ok(source.edit()?)
    }
}

fn owner_commit_bytes(editor: ObjectEdit) -> Result<Vec<u8>, Box<dyn StdError>> {
    let package = editor.commit()?.into_parts().0;
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(package.finish()?)
    }
    #[cfg(feature = "current-owner-api")]
    {
        let mut output = Vec::new();
        package.write_to(&mut output)?;
        Ok(output)
    }
}

fn verify_owner_output(output: &[u8], target: &[u8]) -> Result<(), Box<dyn StdError>> {
    let mut file = OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default())?;
    if file.open_stream(&[TARGET_NAME])? != target {
        return Err("owner probe output target differs from expected bytes".into());
    }
    Ok(())
}

fn user_properties<'a>(
    section: &'a Section,
    limits: UserLimits,
) -> Result<UserProperties<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-property-api")))]
    {
        Ok(UserProperties::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-property-api")]
    {
        Ok(UserProperties::new(section, limits)?)
    }
}

fn user_editor<'a>(
    section: &'a mut Section,
    limits: UserLimits,
) -> Result<UserEdit<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-property-api")))]
    {
        Ok(UserEdit::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-property-api")]
    {
        Ok(UserEdit::new(section, limits)?)
    }
}

fn section_size(section: &Section) -> usize {
    section.property_ids().count() + section.named_properties().count()
}
