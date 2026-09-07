//! Opt-in, pinned external QA fixture copy evidence. No native producer claim.
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Resource,
};
use litchi_opc::{ReadLimits, SourceCacheLimits};
use litchi_perf_baseline::pptx_range_source::{PptxRangeSource, PptxRangeSourceConfig};
use litchi_pptx::{SourceBackedPresentation, SourceBackedPresentationEditor};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{PreservationIndex, ZipArchive};
use std::{
    collections::BTreeMap,
    error::Error,
    fs::OpenOptions,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    path::Path,
    sync::Arc,
    time::{Duration, Instant},
};

const FIXTURE_SHA256: &str = "88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44";
const OUTPUT_LIMIT: usize = 1024 * 1024;
const PML: &str = "ppt/presentation.xml";
const RELS: &str = "ppt/_rels/presentation.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";
fn sha(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}
fn context(label: &'static str) -> Result<(Budget, ExecutionContext), Box<dyn Error>> {
    let budget = Budget::root(
        label,
        Limits::new(
            32 * 1024 * 1024,
            64 * 1024 * 1024,
            64 * 1024 * 1024,
            100_000,
            256,
            64 * 1024 * 1024,
        ),
    );
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).ok_or("workers")?,
        NonZeroUsize::new(1).ok_or("tasks")?,
        NonZeroU64::new(64 * 1024 * 1024).ok_or("work")?,
        0,
    )?;
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    Ok((budget, context))
}
struct Sink {
    bytes: Vec<u8>,
    calls: u64,
    maximum_request: usize,
}
impl Write for Sink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.len() > OUTPUT_LIMIT - self.bytes.len() {
            return Err(io::Error::other("external copy output limit"));
        }
        self.bytes.extend_from_slice(bytes);
        self.calls += 1;
        self.maximum_request = self.maximum_request.max(bytes.len());
        Ok(bytes.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
#[derive(PartialEq, Eq)]
struct Member {
    payload: Vec<u8>,
    local: Vec<u8>,
    central: Vec<u8>,
}
fn members(bytes: &[u8]) -> Result<BTreeMap<String, Member>, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?.into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut scratch)?;
    let mut result = BTreeMap::new();
    for entry in index.entries() {
        let name = std::str::from_utf8(entry.raw_name_bytes())?.to_owned();
        let local = entry.local_span();
        let central = entry.central_record();
        let mut record =
            bytes[usize::try_from(central.start)?..usize::try_from(central.end)?].to_vec();
        record
            .get_mut(42..46)
            .ok_or("short central record")?
            .fill(0);
        let member = Member {
            payload: ArchiveReader::new(bytes)?.read(&name)?,
            local: bytes[usize::try_from(local.start)?..usize::try_from(local.end)?].to_vec(),
            central: record,
        };
        if result.insert(name, member).is_some() {
            return Err("duplicate member".into());
        }
    }
    Ok(result)
}
fn check_output(
    source: &[u8],
    output: &[u8],
    before: &BTreeMap<String, Member>,
) -> Result<Value, Box<dyn Error>> {
    let after = members(output)?;
    for (name, member) in before {
        if !matches!(name.as_str(), PML | RELS | CONTENT_TYPES) && after.get(name) != Some(member) {
            return Err(format!("untouched member changed: {name}").into());
        }
    }
    let added = after
        .keys()
        .filter(|n| !before.contains_key(*n))
        .cloned()
        .collect::<Vec<_>>();
    if added.len() != 3 {
        return Err("expected one slide, slide relationships and image addition".into());
    }
    let copied = added
        .iter()
        .find(|n| n.starts_with("ppt/slides/slide") && n.ends_with(".xml"))
        .ok_or("copied slide")?;
    let image = added
        .iter()
        .find(|n| n.starts_with("ppt/media/"))
        .ok_or("copied image")?;
    if after[copied].payload != before["ppt/slides/slide1.xml"].payload
        || after[image].payload != before["ppt/media/image1.png"].payload
    {
        return Err("copied slide or image payload changed".into());
    }
    let original_rels = &before[RELS].payload;
    let closing = std::str::from_utf8(original_rels)?
        .rfind("</Relationships>")
        .ok_or("fixture root close")?;
    if !after[RELS].payload.starts_with(&original_rels[..closing])
        || !after[RELS].payload.ends_with(&original_rels[closing..])
    {
        return Err("existing relationship lexical bytes changed".into());
    }
    for (member, close) in [(PML, "</p:sldIdLst>"), (CONTENT_TYPES, "</Types>")] {
        let original = &before[member].payload;
        let mut boundary = std::str::from_utf8(original)?
            .rfind(close)
            .ok_or("fixture insertion boundary")?;
        // The content-types writer inserts before the root-closing event,
        // including its preceding XML whitespace. Those original bytes must
        // remain in the suffix, exactly as supplied by this pinned fixture.
        if member == CONTENT_TYPES {
            while boundary > 0 && matches!(original[boundary - 1], b' ' | b'\t' | b'\r' | b'\n') {
                boundary -= 1;
            }
        }
        if !after[member].payload.starts_with(&original[..boundary])
            || !after[member].payload.ends_with(&original[boundary..])
        {
            return Err(format!("original metadata lexical bytes changed: {member}").into());
        }
    }
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output.to_vec())))?;
    let original =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(source.to_vec())))?;
    let expected = original.slide(0).ok_or("source slide")?.text_and_name()?;
    if reopened.slide_count() != 2 || !expected.1.is_empty() {
        return Err("slide count or name mismatch".into());
    }
    for (position, slide) in reopened.slides().enumerate() {
        let copied_image = slide.read_image(0)?;
        let expected_image_member = if position == 0 {
            "ppt/media/image1.png"
        } else {
            image.as_str()
        };
        if copied_image
            .descriptor()
            .target()
            .part_uri()
            .map(|uri| uri.membername())
            != Some(expected_image_member)
        {
            return Err("image relationship targets the wrong member".into());
        }
        if slide.text_and_name()? != expected
            || copied_image.bytes() != before["ppt/media/image1.png"].payload
        {
            return Err("copied semantic readback mismatch".into());
        }
    }
    let eager = litchi_pptx::Package::from_vec(output.to_vec())?;
    let opened = eager.opened_presentation()?;
    if opened.slides().len() != 2 || opened.slides().iter().any(|slide| !slide.name().is_empty()) {
        return Err("eager slide catalog mismatch".into());
    }
    for slide in eager.presentation()?.slides()? {
        let text = slide.text()?;
        if !text.contains("Hello") || !text.contains("Radekski :-)") {
            return Err("eager slide text mismatch".into());
        }
    }
    Ok(
        json!({"original_members":before.len(),"output_members":after.len(),"added_members":added,"source_slide_sha256":sha(&before["ppt/slides/slide1.xml"].payload),"image_sha256":sha(&before["ppt/media/image1.png"].payload),"untouched_raw_records_checked":true,"exact_slide_and_image_checked":true,"relationship_lexical_prefix_suffix_checked":true,"presentation_and_content_types_prefix_suffix_checked":true,"copied_image_relationship_target_checked":true,"semantic_reopen_checked":true,"eager_reopen_checked":true}),
    )
}
fn gauges(budget: &Budget) -> Value {
    json!({"memory":budget.used(Resource::Memory),"objects":budget.used(Resource::Objects),"depth":budget.used(Resource::Depth),"input_bytes":budget.used(Resource::InputBytes),"output_bytes":budget.used(Resource::OutputBytes),"work":budget.used(Resource::Work)})
}
fn write_output_artifact(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut output = OpenOptions::new().write(true).create_new(true).open(path)?;
    output.write_all(bytes)?;
    output.sync_all()?;
    Ok(())
}
fn iteration(
    fixture: &[u8],
    before: &BTreeMap<String, Member>,
    range: bool,
    output_artifact: Option<&Path>,
) -> Result<Value, Box<dyn Error>> {
    let (source_budget, source_context) = context("external-copy-source")?;
    let (destination_budget, destination_context) = context("external-copy-destination")?;
    let config = PptxRangeSourceConfig {
        max_returned_bytes: range.then_some(256),
        fixed_delay: range.then_some(Duration::from_micros(100)),
        ..PptxRangeSourceConfig::default()
    };
    let source_adapter = Arc::new(PptxRangeSource::new(
        Arc::new(OwnedSource::new(fixture.to_vec())),
        config,
    ));
    let destination_adapter = Arc::new(PptxRangeSource::new(
        Arc::new(OwnedSource::new(fixture.to_vec())),
        config,
    ));
    let cache = SourceCacheLimits::new(1024 * 1024, 128)?;
    let output_reservation = destination_budget.reserve(Resource::Memory, OUTPUT_LIMIT as u64)?;
    let mut sink = Sink {
        bytes: Vec::with_capacity(OUTPUT_LIMIT),
        calls: 0,
        maximum_request: 0,
    };
    let started = Instant::now();
    let source =
        SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source_adapter.clone(),
            ReadLimits::default(),
            cache,
            source_context,
        )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(destination_adapter.clone(), ReadLimits::default(), cache, destination_context)?;
    let open_ns = u64::try_from(started.elapsed().as_nanos())?;
    let opened =
        json!({"source":source_adapter.snapshot()?,"destination":destination_adapter.snapshot()?});
    let started = Instant::now();
    let plan = editor.plan_cross_slide_copy(&source, 0, 0, 1)?;
    let plan_ns = u64::try_from(started.elapsed().as_nanos())?;
    let planned = json!({"source":source_adapter.snapshot()?,"destination":destination_adapter.snapshot()?,"source_budget":gauges(&source_budget),"destination_budget":gauges(&destination_budget)});
    let started = Instant::now();
    let published = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
    let publication_ns = u64::try_from(started.elapsed().as_nanos())?;
    let published_reads = json!({
        "source": source_adapter.snapshot()?,
        "destination": destination_adapter.snapshot()?,
        "source_budget": gauges(&source_budget),
        "destination_budget": gauges(&destination_budget),
    });
    if !published.name().is_empty() || published.destination_slide_count() != 2 {
        return Err("published metadata mismatch".into());
    }
    let oracle = check_output(fixture, &sink.bytes, before)?;
    let output_sha256 = sha(&sink.bytes);
    let output_bytes = sink.bytes.len();
    let write_calls = sink.calls;
    let maximum_write_request = sink.maximum_request;
    if let Some(path) = output_artifact {
        write_output_artifact(path, &sink.bytes)?;
    }
    drop(published);
    drop(plan);
    drop(source);
    drop(sink);
    drop(output_reservation);
    for budget in [&source_budget, &destination_budget] {
        for resource in [Resource::Memory, Resource::Objects, Resource::Depth] {
            if budget.used(resource) != 0 {
                return Err("retained budget after drop".into());
            }
        }
    }
    Ok(
        json!({"open_ns":open_ns,"plan_ns":plan_ns,"publication_ns":publication_ns,"api_sum_ns":open_ns.checked_add(plan_ns).and_then(|v|v.checked_add(publication_ns)).ok_or("timing overflow")?,"opened":opened,"planned":planned,"published":published_reads,"output_sha256":output_sha256,"output_bytes":output_bytes,"write_calls":write_calls,"maximum_write_request":maximum_write_request,"oracle":oracle,"final_source_budget":gauges(&source_budget),"final_destination_budget":gauges(&destination_budget),"final_memory_objects_depth_zero_checked":true}),
    )
}
fn main() -> Result<(), Box<dyn Error>> {
    let args = std::env::args().skip(1).collect::<Vec<_>>();
    if args.len() != 6 {
        return Err(
            "expected fixture output-json samples warmups source-revision bytes|range".into(),
        );
    }
    let samples = args[2].parse::<usize>()?;
    let warmups = args[3].parse::<usize>()?;
    if !matches!(samples, 1 | 30)
        || !matches!(warmups, 0 | 3)
        || !matches!(args[5].as_str(), "bytes" | "range")
    {
        return Err("unsupported capture protocol".into());
    }
    let fixture = std::fs::read(&args[0])?;
    if fixture.len() != 29956 || sha(&fixture) != FIXTURE_SHA256 {
        return Err("pinned external fixture mismatch".into());
    }
    let before = members(&fixture)?;
    let output_artifact = Path::new(&args[1]).with_extension("pptx");
    let mut rows = Vec::with_capacity(samples);
    let mut expected = None;
    for i in 0..warmups + samples {
        let row = iteration(
            &fixture,
            &before,
            args[5] == "range",
            (i == warmups).then_some(output_artifact.as_path()),
        )?;
        let identity = (row["output_sha256"].clone(), row["output_bytes"].clone());
        if expected.as_ref().is_some_and(|v| *v != identity) {
            return Err("nondeterministic output".into());
        }
        expected = Some(identity);
        if i >= warmups {
            rows.push(row);
        }
    }
    let exe = std::env::current_exe()?;
    let binary = std::fs::read(&exe)?;
    let output_artifact_bytes = std::fs::read(&output_artifact)?;
    let first = rows.first().ok_or("no retained samples")?;
    let expected_output_sha256 = first["output_sha256"].as_str().ok_or("output hash")?;
    let expected_output_bytes = first["output_bytes"].as_u64().ok_or("output bytes")?;
    if output_artifact_bytes.len() != usize::try_from(expected_output_bytes)?
        || sha(&output_artifact_bytes) != expected_output_sha256
    {
        return Err("retained output artifact identity mismatch".into());
    }
    let report = json!({"schema":"pptx-external-cross-copy-v1","source_revision":args[4],"current_exe":exe,"binary_sha256":sha(&binary),"binary_bytes":binary.len(),"fixture_path":args[0],"fixture_sha256":FIXTURE_SHA256,"fixture_bytes":fixture.len(),"output_artifact":{"path":output_artifact.to_string_lossy().into_owned(),"bytes":output_artifact_bytes.len(),"sha256":sha(&output_artifact_bytes)},"scope":"Unmodified LibreOffice QA fixture independently reopened as source/destination; no known native producer/save chain or native application run; full output retained within 1 MiB sink cap.","timing_scope":"Only source/destination open, plan and publication APIs; fixture load, adapters, sink allocation, diagnostics, exhaustive output oracles and drops excluded.","allocator":"Rust system allocator","instrumentation":"none","provider":args[5],"range_scope":"Logical caller ReadAt; optional 256-byte cap and 100us per-call sleep, no physical network/cold I/O claim.","memory_budget_per_owner":32*1024*1024,"cache_bytes":1024*1024,"cache_entries":128,"output_limit":OUTPUT_LIMIT,"samples":samples,"warmups":warmups,"samples_raw":rows});
    let output = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&args[1])?;
    serde_json::to_writer_pretty(output, &report)?;
    println!("external copy capture passed: {samples} samples");
    Ok(())
}
