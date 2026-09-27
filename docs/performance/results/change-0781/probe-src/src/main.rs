//! Standalone public-workflow probe for change 0781.
//!
//! This executable measures the legacy public `litchi_ppt::writer::Writer`
//! path.  Fixture generation is deterministic and happens before any timed
//! operation.  The `write` mode constructs and populates a writer before the
//! clock and measures only `Writer::write_to`; `lifecycle` includes writer
//! construction, slide and text-box authoring, and publication in one public
//! operation.  Reopening, semantic readback, and raw text-atom verification
//! happen after the clock.

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

use litchi_ppt::odraw::{Record as OfficeArtRecord, RecordKind};
use litchi_ppt::records::Record as PptRecord;
use litchi_ppt::writer::Writer;
use litchi_ppt::writer::text_format::{Paragraph, TextRun};
use litchi_ppt::{EscherTextboxWrapper, Package, RecordType};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.ppt.borrowed-text-probe.v1";
const TOOL: &str = "ppt-borrowed-text-probe-0781";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;
const PAYLOAD_UNITS: usize = 40_000;
const MAX_RAW_RECORDS: usize = 1_000_000;
const MAX_RAW_DEPTH: usize = 256;

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
    raw_text_check: bool,
    expected_raw_text_bytes: u64,
    expected_raw_text_sha256: String,
    raw_text_bytes: u64,
    raw_text_sha256: String,
    raw_text_box_count: usize,
    expected_raw_text_box_count: usize,
    raw_text_atom_count: usize,
    raw_text_boxes_match: bool,
    raw_text_match: bool,
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
    expected_raw_text_bytes: u64,
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
    semantic_text: String,
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

    fn semantic_text(&self) -> String {
        // `odraw::text_from_ppt_records` trims each decoded text atom before
        // appending it.  Keep this normalization separate from `text()`: the
        // latter is the authored source used by the raw atom oracle.
        self.text().trim().to_owned()
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
        expected_semantic_text_bytes: u64::try_from(fixture.semantic_text.len())?,
        expected_raw_text_bytes: u64::try_from(fixture.expected_text.len())?,
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
    let expected_semantic_text = &fixture.semantic_text;
    let expected_semantic_identity = identity(expected_semantic_text.as_bytes())?;
    let expected_raw_text = &fixture.expected_text;
    let expected_raw_identity = identity(expected_raw_text.as_bytes())?;
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
                let expected = semantic_slide_text(expected_slide);
                slide
                    .text()
                    .map(|actual| actual == expected.as_str())
                    .unwrap_or(false)
            });
    let exact_text_match = actual_text == *expected_semantic_text && exact_slide_texts;
    let actual_identity = identity(actual_text.as_bytes())?;
    if !exact_text_match {
        return Err(format!(
            "PPT semantic readback differed from the normalized fixture text or slide order; expected {} bytes / {}, got {} bytes / {}",
            expected_semantic_identity.bytes,
            expected_semantic_identity.sha256,
            actual_identity.bytes,
            actual_identity.sha256,
        )
        .into());
    }

    let mut raw_slide_texts = Vec::with_capacity(slides.len());
    let mut raw_text_box_count = 0usize;
    let mut raw_text_atom_count = 0usize;
    let mut raw_text_boxes = Vec::new();
    for slide in &slides {
        let boxes = raw_textboxes_for_slide(slide)?;
        raw_text_box_count = raw_text_box_count
            .checked_add(boxes.len())
            .ok_or("raw text-box count overflow")?;
        raw_text_atom_count = raw_text_atom_count
            .checked_add(
                boxes
                    .iter()
                    .map(|textbox| textbox.atom_count)
                    .sum::<usize>(),
            )
            .ok_or("raw text-atom count overflow")?;
        raw_slide_texts.push(
            boxes
                .iter()
                .map(|textbox| textbox.text.as_str())
                .collect::<Vec<_>>()
                .join("\n"),
        );
        raw_text_boxes.extend(boxes.into_iter().map(|textbox| textbox.text));
    }
    let raw_text = raw_slide_texts.join("\n\n");
    let expected_raw_text_boxes = fixture
        .slides
        .iter()
        .flat_map(|slide| slide.boxes.iter().map(FixtureBox::text))
        .collect::<Vec<_>>();
    let raw_text_boxes_match = raw_text_boxes == expected_raw_text_boxes;
    let raw_text_match = raw_text == *expected_raw_text && raw_text_boxes_match;
    let raw_identity = identity(raw_text.as_bytes())?;
    if !raw_text_match {
        return Err(format!(
            "PPT raw text atoms differed from the authored fixture or slide order; expected {} bytes / {}, got {} bytes / {} ({} boxes / {} atoms)",
            expected_raw_identity.bytes,
            expected_raw_identity.sha256,
            raw_identity.bytes,
            raw_identity.sha256,
            raw_text_box_count,
            raw_text_atom_count,
        )
        .into());
    }

    Ok(Verification {
        reopened: true,
        slide_count: slides.len(),
        expected_slide_count: fixture.slides.len(),
        slide_count_match,
        semantic_check: exact_text_match,
        expected_semantic_text_bytes: expected_semantic_identity.bytes,
        expected_semantic_text_sha256: expected_semantic_identity.sha256,
        semantic_text_bytes: actual_identity.bytes,
        semantic_text_sha256: actual_identity.sha256,
        exact_text_match,
        raw_text_check: raw_text_match,
        expected_raw_text_bytes: expected_raw_identity.bytes,
        expected_raw_text_sha256: expected_raw_identity.sha256,
        raw_text_bytes: raw_identity.bytes,
        raw_text_sha256: raw_identity.sha256,
        raw_text_box_count,
        expected_raw_text_box_count: expected_raw_text_boxes.len(),
        raw_text_atom_count,
        raw_text_boxes_match,
        raw_text_match,
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
    let semantic_slide_texts: Vec<String> = slides.iter().map(semantic_slide_text).collect();
    let semantic_text = semantic_slide_texts.join("\n\n");
    Ok(Fixture {
        slides,
        source: identity(&canonical)?,
        expected_text,
        semantic_text,
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

fn semantic_slide_text(slide: &FixtureSlide) -> String {
    slide
        .boxes
        .iter()
        .map(FixtureBox::semantic_text)
        .filter(|text| !text.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

#[derive(Debug)]
struct RawTextbox {
    text: String,
    atom_count: usize,
}

fn raw_textboxes_for_slide(
    slide: &litchi_ppt::slide::Slide<'_>,
) -> Result<Vec<RawTextbox>, Box<dyn Error>> {
    let Some(ppdrawing) = slide.record().find_child(RecordType::PPDrawing) else {
        return Ok(Vec::new());
    };

    // `Record` from the PPT layer treats PPDrawing as an opaque payload.  Parse
    // that payload through the public OfficeArt record API so this oracle sees
    // the published ClientTextbox and text-atom bytes directly.
    let data = ppdrawing.data.as_ref();
    let mut textboxes = Vec::new();
    let mut visited = 0usize;
    let mut offset = 0usize;
    while offset < data.len() {
        let (record, consumed) = OfficeArtRecord::parse(data, offset).map_err(|error| {
            raw_error(format!(
                "OfficeArt record parse failed at offset {offset}: {error}"
            ))
        })?;
        if consumed == 0 {
            return Err(raw_error("OfficeArt parser made no progress"));
        }
        collect_officeart_textboxes(&record, 0, &mut visited, &mut textboxes)?;
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| raw_error("OfficeArt record offset overflow"))?;
    }
    Ok(textboxes)
}

fn collect_officeart_textboxes(
    record: &OfficeArtRecord<'_>,
    depth: usize,
    visited: &mut usize,
    textboxes: &mut Vec<RawTextbox>,
) -> Result<(), Box<dyn Error>> {
    *visited = visited
        .checked_add(1)
        .ok_or_else(|| raw_error("OfficeArt record count overflow"))?;
    if *visited > MAX_RAW_RECORDS {
        return Err(raw_error(
            "OfficeArt raw verification record limit exceeded",
        ));
    }

    if record.kind() == RecordKind::ClientTextbox {
        textboxes.push(raw_textbox(record.data())?);
        return Ok(());
    }
    if !record.is_container() {
        return Ok(());
    }
    if depth >= MAX_RAW_DEPTH {
        return Err(raw_error(
            "OfficeArt raw verification nesting limit exceeded",
        ));
    }

    let data = record.data();
    let mut offset = 0usize;
    while offset < data.len() {
        let (child, consumed) = OfficeArtRecord::parse(data, offset).map_err(|error| {
            raw_error(format!(
                "OfficeArt child parse failed at offset {offset}: {error}"
            ))
        })?;
        if consumed == 0 {
            return Err(raw_error("OfficeArt child parser made no progress"));
        }
        collect_officeart_textboxes(&child, depth + 1, visited, textboxes)?;
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| raw_error("OfficeArt child offset overflow"))?;
    }
    Ok(())
}

fn raw_textbox(data: &[u8]) -> Result<RawTextbox, Box<dyn Error>> {
    // Raw text is concatenated exactly as emitted.  `FixtureBox::text` models
    // explicit CR separators between rich paragraphs; the writer's final
    // paragraph break is implicit and is therefore absent from both sides.
    let wrapper = EscherTextboxWrapper::new(data.to_vec())
        .map_err(|error| raw_error(format!("ClientTextbox parse failed: {error}")))?;
    let mut text = String::new();
    let mut atom_count = 0usize;
    let mut visited = 0usize;
    for record in wrapper.child_records() {
        collect_ppt_text_atoms(record, 0, &mut visited, &mut atom_count, &mut text)?;
    }
    if atom_count == 0 {
        return Err(raw_error("ClientTextbox contains no text atom"));
    }
    Ok(RawTextbox { text, atom_count })
}

fn collect_ppt_text_atoms(
    record: &PptRecord,
    depth: usize,
    visited: &mut usize,
    atom_count: &mut usize,
    text: &mut String,
) -> Result<(), Box<dyn Error>> {
    *visited = visited
        .checked_add(1)
        .ok_or_else(|| raw_error("PPT text record count overflow"))?;
    if *visited > MAX_RAW_RECORDS {
        return Err(raw_error("PPT raw text verification record limit exceeded"));
    }

    match record.record_type {
        RecordType::TextCharsAtom => {
            let value = decode_raw_utf16le(record.data.as_ref())?;
            *atom_count = atom_count
                .checked_add(1)
                .ok_or_else(|| raw_error("PPT text atom count overflow"))?;
            text.push_str(&value);
        },
        RecordType::TextBytesAtom => {
            let value = decode_raw_ascii(record.data.as_ref())?;
            *atom_count = atom_count
                .checked_add(1)
                .ok_or_else(|| raw_error("PPT text atom count overflow"))?;
            text.push_str(&value);
        },
        _ => {
            if depth >= MAX_RAW_DEPTH && !record.children.is_empty() {
                return Err(raw_error(
                    "PPT raw text verification nesting limit exceeded",
                ));
            }
            for child in &record.children {
                collect_ppt_text_atoms(child, depth + 1, visited, atom_count, text)?;
            }
        },
    }
    Ok(())
}

fn decode_raw_utf16le(data: &[u8]) -> Result<String, Box<dyn Error>> {
    let mut units = Vec::with_capacity(data.len() / 2);
    let mut chunks = data.chunks_exact(2);
    for chunk in &mut chunks {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    if !chunks.remainder().is_empty() {
        return Err(raw_error("TextCharsAtom has an odd byte length"));
    }
    String::from_utf16(&units).map_err(|_| raw_error("TextCharsAtom contains invalid UTF-16"))
}

fn decode_raw_ascii(data: &[u8]) -> Result<String, Box<dyn Error>> {
    if let Some(&byte) = data.iter().find(|&&byte| !byte.is_ascii()) {
        return Err(raw_error(format!(
            "TextBytesAtom contains non-ASCII byte 0x{:02X}",
            byte
        )));
    }
    Ok(data.iter().map(|byte| char::from(*byte)).collect())
}

fn raw_error(message: impl Into<String>) -> Box<dyn Error> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn semantic_oracle_preserves_authored_fixture_separately() {
        let fixture = FixtureBox::Plain(" authored text ".to_owned());
        assert_eq!(fixture.text(), " authored text ");
        assert_eq!(fixture.semantic_text(), "authored text");
    }

    #[test]
    fn raw_utf16_decoder_rejects_odd_payloads() {
        assert!(decode_raw_utf16le(&[0x41]).is_err());
    }

    #[test]
    fn raw_utf16_decoder_rejects_lone_surrogates() {
        assert!(decode_raw_utf16le(&[0x00, 0xD8]).is_err());
        assert!(decode_raw_utf16le(&[0x00, 0xDC]).is_err());
    }

    #[test]
    fn raw_ascii_decoder_rejects_non_ascii_payloads() {
        assert!(decode_raw_ascii(&[0x80]).is_err());
    }

    #[test]
    fn raw_officeart_parser_rejects_truncated_headers_and_payloads() {
        assert!(OfficeArtRecord::parse(&[0; 7], 0).is_err());

        let mut truncated = vec![0; 8];
        truncated[2..4].copy_from_slice(&RecordKind::Sp.raw().to_le_bytes());
        truncated[4..8].copy_from_slice(&1u32.to_le_bytes());
        assert!(OfficeArtRecord::parse(&truncated, 0).is_err());
    }
}
