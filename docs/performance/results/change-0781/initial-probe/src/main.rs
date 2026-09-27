//! Standalone public-workflow probe for change 0781.
//!
//! This executable measures the legacy public `litchi_ppt::writer::Writer`
//! path.  Fixture generation is deterministic and happens before any timed
//! operation.  The `write` mode constructs and populates a writer before the
//! clock and measures only `Writer::write_to`; `lifecycle` includes writer
//! construction, slide and text-box authoring, and publication in one public
//! operation.  Reopening and semantic readback happen after the clock.

mod allocation_metrics;

#[cfg(feature = "allocator-metrics")]
mod counting_allocator;

use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::time::Instant;

use litchi_ppt::Package;
use litchi_ppt::writer::Writer;
use litchi_ppt::writer::text_format::{Paragraph, TextRun};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.ppt.borrowed-text-probe.v1";
const TOOL: &str = "ppt-borrowed-text-probe-0781";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;
const PAYLOAD_UNITS: usize = 40_000;

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Write,
    Lifecycle,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "write" => Ok(Self::Write),
            "lifecycle" => Ok(Self::Lifecycle),
            _ => Err(format!("--mode must be one of write or lifecycle (got {value:?})").into()),
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Write => {
                "Writer::write_to only; Writer construction and slide/text-box authoring are outside the clock"
            },
            Self::Lifecycle => {
                "Writer::new, add_slide, add_textbox/add_rich_textbox, and Writer::write_to"
            },
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Shape {
    Tiny,
    Many,
    Payload,
    Unicode,
    Rich,
}

impl Shape {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "tiny" => Ok(Self::Tiny),
            "many" => Ok(Self::Many),
            "payload" => Ok(Self::Payload),
            "unicode" => Ok(Self::Unicode),
            "rich" => Ok(Self::Rich),
            _ => Err(format!(
                "--shape must be one of tiny, many, payload, unicode, rich (got {value:?})"
            )
            .into()),
        }
    }

    const fn dimensions(self) -> (usize, usize) {
        match self {
            Self::Tiny => (2, 3),
            Self::Many => (100, 10),
            Self::Payload | Self::Unicode => (16, 4),
            Self::Rich => (2, 2),
        }
    }
}

#[derive(Debug)]
struct Config {
    mode: Mode,
    shape: Shape,
    samples: usize,
    warmup: usize,
    output: PathBuf,
}

#[derive(Debug, Serialize, Clone)]
struct Identity {
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct AllocatorIdentity {
    binary: String,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Debug, Serialize)]
struct Verification {
    reopened: bool,
    slide_count: usize,
    expected_slide_count: usize,
    slide_count_match: bool,
    semantic_check: bool,
    expected_semantic_text_bytes: u64,
    expected_semantic_text_sha256: String,
    semantic_text_bytes: u64,
    semantic_text_sha256: String,
    exact_text_match: bool,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    metrics: BTreeMap<String, u64>,
    source_sha256: String,
    output: Identity,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<allocation_metrics::Sample>,
    verification: Verification,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    mode: Mode,
    shape: Shape,
    slides: usize,
    boxes_per_slide: usize,
    text_boxes: usize,
    timing_scope: &'static str,
    source: Identity,
    expected_semantic_text_bytes: u64,
    warmup: usize,
    samples_requested: usize,
    samples: Vec<SampleRecord>,
    allocator: AllocatorIdentity,
}

#[derive(Clone, Debug)]
struct Fixture {
    slides: Vec<FixtureSlide>,
    source: Identity,
    expected_text: String,
    text_bytes: usize,
}

#[derive(Clone, Debug)]
struct FixtureSlide {
    boxes: Vec<FixtureBox>,
}

#[derive(Clone, Debug)]
enum FixtureBox {
    Plain(String),
    Rich(Vec<Paragraph>),
}

impl FixtureBox {
    fn text(&self) -> String {
        match self {
            Self::Plain(text) => text.clone(),
            Self::Rich(paragraphs) => {
                let mut result = String::new();
                for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
                    if paragraph_index != 0 {
                        result.push('\r');
                    }
                    for run in &paragraph.runs {
                        result.push_str(&run.text);
                    }
                }
                result
            },
        }
    }

    fn append_canonical(&self, output: &mut Vec<u8>) {
        match self {
            Self::Plain(text) => {
                output.push(b'P');
                append_len_prefixed(output, text.as_bytes());
            },
            Self::Rich(paragraphs) => {
                output.extend_from_slice(b"RICH-FORMAT-V1");
                output.push(0);
                output.extend_from_slice(
                    &u64::try_from(paragraphs.len())
                        .unwrap_or(u64::MAX)
                        .to_le_bytes(),
                );
                for paragraph in paragraphs {
                    output.extend_from_slice(
                        &u64::try_from(paragraph.runs.len())
                            .unwrap_or(u64::MAX)
                            .to_le_bytes(),
                    );
                    for run in &paragraph.runs {
                        // The fixture uses this compact explicit marker to
                        // bind the rich control's formatting to its source
                        // identity without serializing an implementation
                        // debug representation.
                        let style = u8::from(run.style.bold)
                            | (u8::from(run.style.italic) << 1)
                            | (u8::from(run.style.underline) << 2);
                        output.push(style);
                        append_len_prefixed(output, run.text.as_bytes());
                    }
                }
            },
        }
    }
}

#[derive(Debug)]
struct OperationResult {
    elapsed_ns: u64,
    output: Vec<u8>,
    verification: Verification,
    allocation: Option<allocation_metrics::Sample>,
    extra_metrics: BTreeMap<String, u64>,
}

fn main() -> Result<(), Box<dyn Error>> {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();

    let config = parse_args(env::args_os().skip(1))?;
    run(config)
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    if let Some(parent) = config.output.parent()
        && !parent.as_os_str().is_empty()
        && !parent.is_dir()
    {
        return Err(format!("output parent is not a directory: {}", parent.display()).into());
    }
    if config.output.exists() {
        return Err(format!("output already exists: {}", config.output.display()).into());
    }

    // Corpus construction, including every deterministic payload string and
    // rich-format model, is outside all sample clocks.
    let fixture = build_fixture(config.shape)?;
    let mut samples = Vec::with_capacity(config.samples);
    let mut deterministic_output: Option<Identity> = None;

    for _ in 0..config.warmup {
        let result = run_one(config.mode, &fixture)?;
        observe_output(&mut deterministic_output, &result.output)?;
        black_box(result);
    }
    for index in 0..config.samples {
        let result = run_one(config.mode, &fixture)?;
        observe_output(&mut deterministic_output, &result.output)?;
        let output = identity(&result.output)?;
        let mut metrics = result.extra_metrics;
        metrics.insert("elapsed_ns".to_owned(), result.elapsed_ns);
        metrics.insert("slides".to_owned(), fixture.slides.len() as u64);
        metrics.insert(
            "boxes_per_slide".to_owned(),
            fixture.boxes_per_slide() as u64,
        );
        samples.push(SampleRecord {
            index,
            elapsed_ns: result.elapsed_ns,
            metrics,
            source_sha256: fixture.source.sha256.clone(),
            output,
            allocation: result.allocation,
            verification: result.verification,
        });
    }

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        mode: config.mode,
        shape: config.shape,
        slides: fixture.slides.len(),
        boxes_per_slide: fixture.boxes_per_slide(),
        text_boxes: fixture.text_boxes(),
        timing_scope: config.mode.timing_scope(),
        source: fixture.source.clone(),
        expected_semantic_text_bytes: u64::try_from(fixture.text_bytes)?,
        warmup: config.warmup,
        samples_requested: config.samples,
        samples,
        allocator: AllocatorIdentity {
            binary: executable_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
    };
    fs::write(&config.output, serde_json::to_vec_pretty(&report)?)?;
    Ok(())
}

fn run_one(mode: Mode, fixture: &Fixture) -> Result<OperationResult, Box<dyn Error>> {
    match mode {
        Mode::Write => run_write(fixture),
        Mode::Lifecycle => run_lifecycle(fixture),
    }
}

/// Measure only the public `Writer::write_to` operation after authoring is
/// complete.  The writer and destination remain live until `finish()` so the
/// region includes the complete write's allocation lifetime.
#[inline(never)]
fn run_write(fixture: &Fixture) -> Result<OperationResult, Box<dyn Error>> {
    let mut writer = populate_writer(fixture)?;
    let mut output = Cursor::new(Vec::new());
    let region = allocation_metrics::begin();
    let started = Instant::now();
    writer.write_to(&mut output)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    black_box(&writer);
    black_box(&output);
    let allocation = region.finish();
    let output = output.into_inner();
    let verification = verify_output(&output, fixture)?;
    Ok(OperationResult {
        elapsed_ns,
        output,
        verification,
        allocation,
        extra_metrics: BTreeMap::new(),
    })
}

/// Measure construction, authoring, and `Writer::write_to` together.  The
/// fixture strings themselves were generated before entering this function.
#[inline(never)]
fn run_lifecycle(fixture: &Fixture) -> Result<OperationResult, Box<dyn Error>> {
    let region = allocation_metrics::begin();
    let started = Instant::now();
    let mut writer = populate_writer(fixture)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    black_box(&writer);
    black_box(&output);
    let allocation = region.finish();
    let output = output.into_inner();
    let verification = verify_output(&output, fixture)?;
    Ok(OperationResult {
        elapsed_ns,
        output,
        verification,
        allocation,
        extra_metrics: BTreeMap::new(),
    })
}

fn populate_writer(fixture: &Fixture) -> Result<Writer, Box<dyn Error>> {
    let mut writer = Writer::new();
    for (slide_index, slide_fixture) in fixture.slides.iter().enumerate() {
        let slide = writer.add_slide()?;
        for (box_index, fixture_box) in slide_fixture.boxes.iter().enumerate() {
            let x = 36 + i32::try_from(box_index % 4)? * 180;
            let y = 36 + i32::try_from(box_index / 4)? * 90;
            match fixture_box {
                FixtureBox::Plain(text) => {
                    writer.add_textbox(slide, x, y, 144, 54, text)?;
                },
                FixtureBox::Rich(paragraphs) => {
                    writer.add_rich_textbox(slide, x, y, 144, 54, paragraphs.clone())?;
                },
            }
        }
        // Keep the slide index in the loop's observable path.  It documents
        // the source order and prevents future fixture changes from silently
        // dropping a slide while preserving the public API call sequence.
        debug_assert_eq!(slide, slide_index);
    }
    Ok(writer)
}

fn verify_output(output: &[u8], fixture: &Fixture) -> Result<Verification, Box<dyn Error>> {
    let expected_text = &fixture.expected_text;
    let expected_text_identity = identity(expected_text.as_bytes())?;
    let mut package = Package::from_reader(Cursor::new(output.to_vec()))?;
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let actual_text = presentation.text()?;
    let slide_count_match = slides.len() == fixture.slides.len();
    let exact_slide_texts = slide_count_match
        && slides
            .iter()
            .zip(fixture.slides.iter())
            .all(|(slide, expected_slide)| {
                let expected = expected_slide_text(expected_slide);
                slide
                    .text()
                    .map(|actual| actual == expected.as_str())
                    .unwrap_or(false)
            });
    let exact_text_match = actual_text == *expected_text && exact_slide_texts;
    let actual_identity = identity(actual_text.as_bytes())?;
    if !exact_text_match {
        return Err(format!(
            "PPT readback differed from the exact fixture text or slide order; expected {} bytes / {}, got {} bytes / {}",
            expected_text_identity.bytes,
            expected_text_identity.sha256,
            actual_identity.bytes,
            actual_identity.sha256,
        )
        .into());
    }
    Ok(Verification {
        reopened: true,
        slide_count: slides.len(),
        expected_slide_count: fixture.slides.len(),
        slide_count_match,
        semantic_check: exact_text_match,
        expected_semantic_text_bytes: expected_text_identity.bytes,
        expected_semantic_text_sha256: expected_text_identity.sha256,
        semantic_text_bytes: actual_identity.bytes,
        semantic_text_sha256: actual_identity.sha256,
        exact_text_match,
    })
}

fn build_fixture(shape: Shape) -> Result<Fixture, Box<dyn Error>> {
    let (slide_count, boxes_per_slide) = shape.dimensions();
    let mut slides = Vec::with_capacity(slide_count);
    for slide in 0..slide_count {
        let mut boxes = Vec::with_capacity(boxes_per_slide);
        for box_index in 0..boxes_per_slide {
            let fixture_box = match shape {
                Shape::Tiny | Shape::Many => FixtureBox::Plain(short_text(shape, slide, box_index)),
                Shape::Payload => FixtureBox::Plain(ascii_payload(slide, box_index)),
                Shape::Unicode => FixtureBox::Plain(unicode_payload(slide, box_index)),
                Shape::Rich => FixtureBox::Rich(rich_text(slide, box_index)),
            };
            boxes.push(fixture_box);
        }
        slides.push(FixtureSlide { boxes });
    }

    let mut canonical = Vec::new();
    canonical.extend_from_slice(b"litchi-ppt-0781-fixture-v1\0");
    canonical.push(shape as u8);
    canonical.extend_from_slice(
        &u64::try_from(slides.len())
            .unwrap_or(u64::MAX)
            .to_le_bytes(),
    );
    for slide in &slides {
        canonical.extend_from_slice(
            &u64::try_from(slide.boxes.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        for fixture_box in &slide.boxes {
            fixture_box.append_canonical(&mut canonical);
        }
    }

    let expected_slide_texts: Vec<String> = slides.iter().map(expected_slide_text).collect();
    let expected_text = expected_slide_texts.join("\n\n");
    let text_bytes = expected_text.len();
    Ok(Fixture {
        slides,
        source: identity(&canonical)?,
        expected_text,
        text_bytes,
    })
}

fn expected_slide_text(slide: &FixtureSlide) -> String {
    slide
        .boxes
        .iter()
        .map(FixtureBox::text)
        .collect::<Vec<_>>()
        .join("\n")
}

fn short_text(shape: Shape, slide: usize, box_index: usize) -> String {
    let label = match shape {
        Shape::Tiny => "tiny-ascii",
        Shape::Many => "many-short-ascii",
        Shape::Payload => "payload-ascii",
        Shape::Unicode => "payload-unicode",
        Shape::Rich => "rich-control",
    };
    format!("litchi-perf-0781-{label}-{slide:03}-{box_index:02}")
}

fn ascii_payload(slide: usize, box_index: usize) -> String {
    repeat_to_bytes(
        &format!("litchi-perf-0781-payload-ascii-{slide:03}-{box_index:02} "),
        PAYLOAD_UNITS,
    )
}

fn unicode_payload(slide: usize, box_index: usize) -> String {
    repeat_to_utf16_units(
        &format!("litchi-perf-0781-payload-unicode-{slide:03}-{box_index:02} BMP漢字😀🚀 "),
        PAYLOAD_UNITS,
    )
}

fn rich_text(slide: usize, box_index: usize) -> Vec<Paragraph> {
    vec![Paragraph::with_runs(vec![
        TextRun::new(format!("litchi-perf-0781-rich-{slide:02}-{box_index:02}-")).bold(),
        TextRun::new("bold").italic(),
        TextRun::new("-control"),
    ])]
}

fn repeat_to_bytes(seed: &str, target: usize) -> String {
    let mut value = String::with_capacity(target);
    while value.len() + seed.len() <= target {
        value.push_str(seed);
    }
    if value.len() < target {
        value.push_str(&"A".repeat(target - value.len()));
    }
    value
}

fn repeat_to_utf16_units(seed: &str, target: usize) -> String {
    let mut value = String::with_capacity(target);
    let seed_units = seed.encode_utf16().count();
    let mut units = 0usize;
    while units + seed_units <= target {
        value.push_str(seed);
        units += seed_units;
    }
    while units < target {
        value.push('X');
        units += 1;
    }
    value
}

fn append_len_prefixed(output: &mut Vec<u8>, bytes: &[u8]) {
    output.extend_from_slice(&u64::try_from(bytes.len()).unwrap_or(u64::MAX).to_le_bytes());
    output.extend_from_slice(bytes);
}

fn executable_identity() -> String {
    env::current_exe()
        .ok()
        .and_then(|path| {
            path.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .unwrap_or_else(|| allocation_metrics::binary_identity().to_owned())
}

fn observe_output(expected: &mut Option<Identity>, output: &[u8]) -> Result<(), Box<dyn Error>> {
    let current = identity(output)?;
    match expected {
        None => *expected = Some(current),
        Some(previous) if previous.bytes == current.bytes && previous.sha256 == current.sha256 => {
        },
        Some(previous) => {
            return Err(format!(
                "non-deterministic publication: {} bytes / {} followed {} bytes / {}",
                previous.bytes, previous.sha256, current.bytes, current.sha256
            )
            .into());
        },
    }
    Ok(())
}

impl Fixture {
    fn boxes_per_slide(&self) -> usize {
        self.slides.first().map_or(0, |slide| slide.boxes.len())
    }

    fn text_boxes(&self) -> usize {
        self.slides.iter().map(|slide| slide.boxes.len()).sum()
    }
}

fn identity(bytes: &[u8]) -> Result<Identity, Box<dyn Error>> {
    Ok(Identity {
        bytes: u64::try_from(bytes.len())?,
        sha256: sha256_hex(bytes),
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos())
        .map_err(|_| "elapsed duration overflows u64 nanoseconds".into())
}

fn parse_args<I>(arguments: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut mode = None;
    let mut shape = None;
    let mut samples = DEFAULT_SAMPLES;
    let mut warmup = DEFAULT_WARMUP;
    let mut output = None;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let argument = argument.to_string_lossy();
        match argument.as_ref() {
            "--mode" => mode = Some(Mode::parse(&next_string(&mut arguments, "--mode")?)?),
            "--shape" => shape = Some(Shape::parse(&next_string(&mut arguments, "--shape")?)?),
            "--samples" => {
                samples = parse_count(
                    &next_string(&mut arguments, "--samples")?,
                    "samples",
                    MAX_SAMPLES,
                )?
            },
            "--warmup" => {
                warmup = parse_count(
                    &next_string(&mut arguments, "--warmup")?,
                    "warmup",
                    MAX_WARMUP,
                )?
            },
            "--output" => output = Some(next_path(&mut arguments, "--output")?),
            "--help" | "-h" => return Err(usage().into()),
            value => return Err(format!("unknown argument {value:?}\n{}", usage()).into()),
        }
    }
    let mode = mode.ok_or_else(|| format!("--mode is required\n{}", usage()))?;
    let shape = shape.ok_or_else(|| format!("--shape is required\n{}", usage()))?;
    let output = output.ok_or_else(|| format!("--output is required\n{}", usage()))?;
    if samples == 0 {
        return Err("--samples must be at least 1".into());
    }
    Ok(Config {
        mode,
        shape,
        samples,
        warmup,
        output,
    })
}

fn next_path<I>(arguments: &mut I, option: &str) -> Result<PathBuf, Box<dyn Error>>
where
    I: Iterator<Item = std::ffi::OsString>,
{
    arguments
        .next()
        .map(PathBuf::from)
        .ok_or_else(|| format!("{option} requires a path").into())
}

fn next_string<I>(arguments: &mut I, option: &str) -> Result<String, Box<dyn Error>>
where
    I: Iterator<Item = std::ffi::OsString>,
{
    arguments
        .next()
        .map(|value| value.to_string_lossy().into_owned())
        .ok_or_else(|| format!("{option} requires a value").into())
}

fn parse_count(value: &str, name: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let count = value
        .parse::<usize>()
        .map_err(|error| format!("--{name} must be a non-negative integer: {error}"))?;
    if count > maximum {
        return Err(format!("--{name} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn usage() -> &'static str {
    "usage: ppt-borrowed-text-probe --mode write|lifecycle --shape tiny|many|payload|unicode|rich --samples N --warmup N --output PATH"
}
