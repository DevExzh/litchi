//! Real source-backed DrawingML adapters for the paired raw qualification.
//!
//! This endpoint exercises the shared `litchi-drawingml` codecs directly and
//! leaves package placement to the owning format crates. Every qualified row
//! carries an outside-clock fidelity receipt. A missing adapter or typed
//! refusal is a hard error; it is never converted into a timing claim.

#![forbid(unsafe_code)]

use std::env;
use std::fs;
use std::hint::black_box;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use litchi_drawingml as dml;
use quick_xml::Reader;
use quick_xml::events::Event;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

const LANES: &[&str] = &[
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
];

const OPERATIONS: &[&str] = &[
    "entry_read",
    "read_shared_when_supported",
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
];

#[derive(Debug)]
struct Arguments {
    lane: String,
    role: String,
    mode: String,
    fixture: PathBuf,
    source_revision: String,
    iterations: usize,
    warmups: usize,
    blocks: usize,
    samples: usize,
    target_ms: u64,
    minimum_ms: u64,
    output: PathBuf,
}

#[derive(Debug, Clone, Deserialize)]
struct FixtureLane {
    root: String,
    mutation: String,
    malformed: String,
}

#[derive(Debug, Deserialize)]
struct FixtureFile {
    lanes: std::collections::BTreeMap<String, FixtureLane>,
}

#[derive(Debug, Clone, Copy)]
pub struct AllocationToken {
    pub allocation_calls: u64,
    pub deallocation_calls: u64,
    pub reallocation_calls: u64,
    pub failed_calls: u64,
    pub direct_bytes: u64,
    pub deallocated_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub live_bytes: u64,
    pub peak_live_bytes: u64,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct AllocationMetrics {
    pub allocation_calls: u64,
    pub deallocation_calls: u64,
    pub reallocation_calls: u64,
    pub failed_calls: u64,
    pub direct_bytes: u64,
    pub deallocated_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub live_before: u64,
    pub live_after: u64,
    pub peak_live_delta: u64,
    pub conservation_ok: bool,
    pub elapsed_claimable: bool,
}

/// A safe observation hook implemented only by the allocator executable.
pub trait AllocationObserver {
    fn begin(&self) -> AllocationToken;
    fn finish(&self, before: AllocationToken) -> AllocationMetrics;
}

#[derive(Debug, Clone, Copy)]
enum PartKind {
    Chart,
    FormatCode,
    Diagram,
    TextBody,
    SvgBlip,
    Model3d,
    Blip,
    Theme,
    ThemeFamily,
    InkActions,
    Geometry,
    Color,
    Transform,
}

#[derive(Debug, Clone)]
struct Part {
    kind: PartKind,
    source: Vec<u8>,
    changed: Vec<u8>,
    malformed: Vec<u8>,
    over: Vec<u8>,
    limit_resource: &'static str,
}

#[derive(Debug, Clone)]
struct LaneCase {
    parts: Vec<Part>,
    stale: Vec<u8>,
}

#[derive(Debug, Clone, Serialize)]
struct RefusalReceipt {
    kind: &'static str,
    resource: Option<&'static str>,
    limit: Option<usize>,
    observed: usize,
    message: String,
}

#[derive(Debug)]
struct ProbeFailure {
    message: String,
    refusal: Option<RefusalReceipt>,
}

#[derive(Debug, Clone, Serialize)]
struct SinkReceipt {
    writes: usize,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Clone, Serialize)]
struct SampleReceipt {
    elapsed_ns: u128,
    iterations: usize,
}

#[derive(Debug, Clone, Serialize)]
struct Row {
    operation: String,
    status: &'static str,
    elapsed_ns: u128,
    fidelity_checked: bool,
    typed_refusal_checked: bool,
    sink_writes: usize,
    sink: Option<SinkReceipt>,
    samples: Vec<SampleReceipt>,
    details: serde_json::Value,
}

#[derive(Debug, Clone, Serialize)]
struct SizeClassReceipt {
    class: &'static str,
    bytes: usize,
    accepted: bool,
    output_sha256: Option<String>,
    refusal: Option<RefusalReceipt>,
}

#[derive(Debug, Clone, Serialize)]
struct ResourceMetrics {
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
}

#[derive(Debug, Clone, Serialize)]
struct AllocatorReport {
    elapsed_claimable: bool,
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_calls: u64,
    direct_bytes: u64,
    deallocated_bytes: u64,
    realloc_old_bytes: u64,
    realloc_new_bytes: u64,
    live_before: u64,
    live_after: u64,
    live_balance: i128,
    peak_live_delta: u64,
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
    conservation_ok: bool,
}

#[derive(Debug, Clone, Serialize)]
struct Protocol {
    warmups: usize,
    samples_per_leg: usize,
    blocks: usize,
    target_ms: u64,
    minimum_ms: u64,
    fresh_process: bool,
    abba_order: [&'static str; 4],
}

#[derive(Debug, Clone, Serialize)]
struct Report {
    schema: &'static str,
    lane: String,
    role: String,
    mode: String,
    source_revision: String,
    status: &'static str,
    qualification_claim: bool,
    fixture_bytes: usize,
    fixture_sha256: String,
    fixture_root_sha256: String,
    fixture_mutation: String,
    size_classes: Vec<SizeClassReceipt>,
    protocol: Protocol,
    allocator_metrics: Option<AllocatorReport>,
    resource_metrics: ResourceMetrics,
    rows: Vec<Row>,
    iterations: usize,
}

fn usage() -> &'static str {
    "usage: drawingml-raw-{native,allocator} --lane LANE --role ROLE \
     --mode MODE --fixture PATH --source-revision SHA --iterations N --warmups N \
     --blocks N --samples N --target-ms N --minimum-ms N --json PATH"
}

fn invalid_argument(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

fn argument_value(arguments: &mut impl Iterator<Item = String>, name: &str) -> io::Result<String> {
    arguments
        .next()
        .ok_or_else(|| invalid_argument(format!("missing value for {name}")))
}

fn parse(arguments: impl Iterator<Item = String>, expected_mode: &str) -> io::Result<Arguments> {
    let mut lane = None;
    let mut role = None;
    let mut mode = None;
    let mut fixture = None;
    let mut source_revision = None;
    let mut iterations = None;
    let mut warmups = None;
    let mut blocks = None;
    let mut samples = None;
    let mut target_ms = None;
    let mut minimum_ms = None;
    let mut output = None;
    let mut arguments = arguments;
    while let Some(argument) = arguments.next() {
        let slot = match argument.as_str() {
            "--lane" => &mut lane,
            "--role" => &mut role,
            "--mode" => &mut mode,
            "--fixture" => &mut fixture,
            "--source-revision" => &mut source_revision,
            "--iterations" => &mut iterations,
            "--warmups" => &mut warmups,
            "--blocks" => &mut blocks,
            "--samples" => &mut samples,
            "--target-ms" => &mut target_ms,
            "--minimum-ms" => &mut minimum_ms,
            "--json" => &mut output,
            _ => return Err(invalid_argument(format!("unknown argument {argument:?}"))),
        };
        if slot.is_some() {
            return Err(invalid_argument(format!("duplicate argument {argument}")));
        }
        *slot = Some(argument_value(&mut arguments, &argument)?);
    }

    let lane = lane.ok_or_else(|| invalid_argument("--lane is required"))?;
    if !LANES.contains(&lane.as_str()) {
        return Err(invalid_argument(format!(
            "unsupported DrawingML lane {lane:?}"
        )));
    }
    let role = role.ok_or_else(|| invalid_argument("--role is required"))?;
    if !matches!(role.as_str(), "baseline" | "candidate") {
        return Err(invalid_argument(format!("unsupported role {role:?}")));
    }
    let mode = mode.ok_or_else(|| invalid_argument("--mode is required"))?;
    if mode != expected_mode {
        return Err(invalid_argument(format!(
            "{expected_mode} binary requires --mode {expected_mode}"
        )));
    }
    let fixture = PathBuf::from(fixture.ok_or_else(|| invalid_argument("--fixture is required"))?);
    if fixture.is_symlink() || !fixture.is_file() {
        return Err(invalid_argument(format!(
            "fixture is not a regular file: {fixture:?}"
        )));
    }
    let source_revision =
        source_revision.ok_or_else(|| invalid_argument("--source-revision is required"))?;
    if source_revision.len() != 40
        || !source_revision
            .bytes()
            .all(|byte| byte.is_ascii_hexdigit() && !byte.is_ascii_uppercase())
    {
        return Err(invalid_argument(
            "--source-revision must be a lowercase 40-character SHA",
        ));
    }
    let parse_usize = |value: Option<String>, name: &str| -> io::Result<usize> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<usize>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let parse_u64 = |value: Option<String>, name: &str| -> io::Result<u64> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<u64>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let iterations = parse_usize(iterations, "--iterations")?;
    let warmups = parse_usize(warmups, "--warmups")?;
    let blocks = parse_usize(blocks, "--blocks")?;
    let samples = parse_usize(samples, "--samples")?;
    let target_ms = parse_u64(target_ms, "--target-ms")?;
    let minimum_ms = parse_u64(minimum_ms, "--minimum-ms")?;
    if iterations == 0 || warmups != 3 || blocks != 5 || samples != 20 {
        return Err(invalid_argument(
            "iterations must be positive, warmups=3, blocks=5, samples=20",
        ));
    }
    if target_ms < minimum_ms || minimum_ms < 150 {
        return Err(invalid_argument(
            "target-ms must be at least minimum-ms and minimum-ms must be >= 150",
        ));
    }
    let output = PathBuf::from(output.ok_or_else(|| invalid_argument("--json is required"))?);
    if output.as_os_str().is_empty() || output.is_symlink() || output.exists() {
        return Err(invalid_argument(format!(
            "JSON output must be a fresh regular-file path: {output:?}"
        )));
    }
    Ok(Arguments {
        lane,
        role,
        mode,
        fixture,
        source_revision,
        iterations,
        warmups,
        blocks,
        samples,
        target_ms,
        minimum_ms,
        output,
    })
}

fn sha256(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    format!("{:x}", digest.finalize())
}

fn load_fixture(path: &Path, lane: &str) -> io::Result<(Vec<u8>, FixtureLane)> {
    let bytes = fs::read(path)?;
    let value: FixtureFile = serde_json::from_slice(&bytes).map_err(|error| {
        invalid_argument(format!("fixture descriptor is invalid JSON: {error}"))
    })?;
    let lane_value = value
        .lanes
        .get(lane)
        .cloned()
        .ok_or_else(|| invalid_argument(format!("fixture descriptor has no lane {lane}")))?;
    Ok((bytes, lane_value))
}

fn fail(message: impl Into<String>) -> ProbeFailure {
    ProbeFailure {
        message: message.into(),
        refusal: None,
    }
}

fn classify(error: dml::Error, observed: usize) -> ProbeFailure {
    let message = error.to_string();
    let refusal = match error {
        dml::Error::Limit { resource, limit } => Some(RefusalReceipt {
            kind: "limit",
            resource: Some(resource),
            limit: Some(limit),
            observed,
            message: message.clone(),
        }),
        _ => Some(RefusalReceipt {
            kind: "invalid_input",
            resource: None,
            limit: None,
            observed,
            message: message.clone(),
        }),
    };
    ProbeFailure { message, refusal }
}

fn io_failure(error: io::Error, observed: usize) -> ProbeFailure {
    ProbeFailure {
        message: error.to_string(),
        refusal: Some(RefusalReceipt {
            kind: "io_error",
            resource: None,
            limit: None,
            observed,
            message: error.to_string(),
        }),
    }
}

fn read_chart(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let chart = dml::chart::reader::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    let mut output = Vec::new();
    dml::chart::writer::write(&mut output, &chart)
        .map_err(|error| io_failure(error, bytes.len()))?;
    Ok(output)
}

fn read_format_code(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::chart::extension::formatcode2::read(bytes)
        .map_err(|error| classify(error, bytes.len()))?;
    dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_diagram(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::Outline::read(bytes)
        .map(|_| bytes.to_vec())
        .map_err(|error| classify(error, bytes.len()))
}

fn body_roundtrip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let body =
        dml::text::body::reader::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    let mut contents = String::new();
    dml::text::body::writer::write(&mut contents, &body)
        .map_err(|error| classify(error, bytes.len()))?;
    let mut output = String::with_capacity(contents.len() + 32);
    output.push_str("<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
    output.push_str(&contents);
    output.push_str("</a:txBody>");
    Ok(output.into_bytes())
}

fn read_svg(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::svg_blip::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_model3d(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::model3d::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::model3d::codec::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_blip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let _ = dml::blip::find_first_embed(bytes).map_err(|error| classify(error, bytes.len()))?;
    Ok(bytes.to_vec())
}

fn read_theme(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let theme = dml::theme::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::theme::codec::encode_part(&theme.name, &theme.colors, &theme.fonts)
        .map_err(|error| classify(error, bytes.len()))
}

fn read_family(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let family = dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::theme::family::write(&family).map_err(|error| classify(error, bytes.len()))
}

fn read_ink(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let profile =
        dml::ink::actions::read_profile(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::ink::actions::write_profile(&profile).map_err(|error| classify(error, bytes.len()))
}

fn read_geometry(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let geometry = parse_geometry(bytes)?;
    let mut output = String::new();
    dml::geometry::writer::write_custom_geometry(&mut output, &geometry);
    Ok(output.into_bytes())
}

fn read_color(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::color::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::color::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_transform(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::transform::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::transform::codec::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn roundtrip(kind: PartKind, bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    match kind {
        PartKind::Chart => read_chart(bytes),
        PartKind::FormatCode => read_format_code(bytes),
        PartKind::Diagram => read_diagram(bytes),
        PartKind::TextBody => body_roundtrip(bytes),
        PartKind::SvgBlip => read_svg(bytes),
        PartKind::Model3d => read_model3d(bytes),
        PartKind::Blip => read_blip(bytes),
        PartKind::Theme => read_theme(bytes),
        PartKind::ThemeFamily => read_family(bytes),
        PartKind::InkActions => read_ink(bytes),
        PartKind::Geometry => read_geometry(bytes),
        PartKind::Color => read_color(bytes),
        PartKind::Transform => read_transform(bytes),
    }
}

fn write_kind_to<W: Write>(
    kind: PartKind,
    bytes: &[u8],
    writer: &mut W,
) -> Result<(), ProbeFailure> {
    match kind {
        PartKind::Chart => {
            let chart =
                dml::chart::reader::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::chart::writer::write(writer, &chart)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::FormatCode => {
            let value = dml::chart::extension::formatcode2::read(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            dml::chart::extension::formatcode2::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Diagram => {
            let _ = dml::diagram::edit::Outline::read(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(bytes)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::TextBody => {
            let body = dml::text::body::reader::read(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            let mut contents = String::new();
            dml::text::body::writer::write(&mut contents, &body)
                .map_err(|error| classify(error, bytes.len()))?;
            let mut output = String::with_capacity(contents.len() + 32);
            output.push_str(
                "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
            );
            output.push_str(&contents);
            output.push_str("</a:txBody>");
            writer
                .write_all(output.as_bytes())
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::SvgBlip => {
            let value = dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::svg_blip::write_to(writer, &value).map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Model3d => {
            let value =
                dml::model3d::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::model3d::codec::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Blip => {
            let _ =
                dml::blip::find_first_embed(bytes).map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(bytes)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::Theme => {
            let value =
                dml::theme::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            let output = dml::theme::codec::encode_part(&value.name, &value.colors, &value.fonts)
                .map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::ThemeFamily => {
            let value =
                dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::theme::family::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::InkActions => {
            let value = dml::ink::actions::read_profile(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            dml::ink::actions::write_profile_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Geometry => {
            let geometry = parse_geometry(bytes)?;
            let mut output = String::new();
            dml::geometry::writer::write_custom_geometry(&mut output, &geometry);
            writer
                .write_all(output.as_bytes())
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::Color => {
            let value = dml::color::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            let output = dml::color::write(&value).map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::Transform => {
            let value =
                dml::transform::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::transform::codec::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
    }
}

struct CountingSink {
    cap: usize,
    bytes: Vec<u8>,
    writes: usize,
}

impl CountingSink {
    fn new(cap: usize) -> Self {
        Self {
            cap,
            bytes: Vec::new(),
            writes: 0,
        }
    }
}

impl Write for CountingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len().saturating_add(bytes.len()) > self.cap {
            return Err(io::Error::new(io::ErrorKind::WriteZero, "bounded sink cap"));
        }
        self.writes = self.writes.saturating_add(1);
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn write_receipt(sink: CountingSink) -> SinkReceipt {
    SinkReceipt {
        writes: sink.writes,
        bytes: sink.bytes.len(),
        sha256: sha256(&sink.bytes),
    }
}

fn fixture_seed(fixture: &FixtureLane) -> String {
    let mut seed = fixture.root.clone();
    seed.truncate(seed.len().min(32));
    seed.chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .collect()
}

fn chart_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::chart::authoring::{Kind, Series, Spec, Variant};
    let seed = fixture_seed(fixture);
    let spec = Spec::new(Variant::new(Kind::Column))
        .with_title(format!("DrawingML {seed}"))
        .with_categories(["Q1", "Q2", "Q3"])
        .with_series(Series::new("East", vec![1.0, 2.0, 3.0]));
    let chart = spec.to_chart().map_err(|error| classify(error, 0))?;
    let mut source_wire = Vec::new();
    dml::chart::writer::write(&mut source_wire, &chart).map_err(|error| io_failure(error, 0))?;
    let source = read_chart(&source_wire)?;
    let changed_spec = Spec::new(Variant::new(Kind::Column))
        .with_title(format!("{} changed", fixture.mutation))
        .with_categories(["Q1", "Q2", "Q3"])
        .with_series(Series::new("East", vec![1.0, 2.0, 4.0]));
    let changed_chart = changed_spec
        .to_chart()
        .map_err(|error| classify(error, 0))?;
    let mut changed_wire = Vec::new();
    dml::chart::writer::write(&mut changed_wire, &changed_chart)
        .map_err(|error| io_failure(error, 0))?;
    let changed = read_chart(&changed_wire)?;
    Ok(vec![
        Part {
            kind: PartKind::Chart,
            source,
            changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::chart::extension::formatcode2::MAX_XML_BYTES + 1],
            limit_resource: "formatcode2 XML bytes",
        },
        format_code_part(fixture)?,
    ])
}

fn format_code_part(fixture: &FixtureLane) -> Result<Part, ProbeFailure> {
    let mut value = dml::chart::extension::formatcode2::Element::new("0.00")
        .map_err(|error| classify(error, 0))?;
    let source =
        dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, 0))?;
    value
        .set_value(format!("0.00;{}", fixture.mutation.len()))
        .map_err(|error| classify(error, 0))?;
    let changed =
        dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, 0))?;
    Ok(Part {
        kind: PartKind::FormatCode,
        source,
        changed,
        malformed: fixture.malformed.as_bytes().to_vec(),
        over: vec![b' '; dml::chart::extension::formatcode2::MAX_XML_BYTES + 1],
        limit_resource: "formatcode2 XML bytes",
    })
}

fn diagram_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::diagram::{DiagramType, SmartArtBuilder, generate_smartart_data_xml};
    let diagram = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("relationship")
        .add_items([fixture_seed(fixture), "second"])
        .build();
    let source = generate_smartart_data_xml(&diagram).into_bytes();
    let outline = dml::diagram::edit::Outline::read(&source)
        .map_err(|error| classify(error, source.len()))?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| fail("diagram authoring produced no content entry"))?;
    let text_changed =
        dml::diagram::edit::set_text(&source, &entry.id, &format!("{} changed", fixture.mutation))
            .map_err(|error| classify(error, source.len()))?;
    let changed = dml::diagram::edit::set_drawing_reference(&text_changed, Some("rIdDrawing"))
        .map_err(|error| classify(error, text_changed.len()))?;
    Ok(vec![Part {
        kind: PartKind::Diagram,
        source,
        changed,
        malformed: fixture.malformed.as_bytes().to_vec(),
        over: vec![b' '; dml::diagram::edit::MAX_PART_BYTES + 1],
        limit_resource: "diagram data-model bytes",
    }])
}

fn text_source() -> Result<Vec<u8>, ProbeFailure> {
    use dml::text::body::{Body, Paragraph, Run};
    let body = Body {
        paragraphs: vec![Paragraph {
            runs: vec![Run::plain("seed")],
            ..Paragraph::default()
        }],
        ..Body::default()
    };
    let mut content = String::new();
    dml::text::body::writer::write(&mut content, &body).map_err(|error| classify(error, 0))?;
    Ok(format!(
        "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{content}</a:txBody>"
    )
    .into_bytes())
}

fn text_media_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    let source = text_source()?;
    let mut body =
        dml::text::body::reader::read(&source).map_err(|error| classify(error, source.len()))?;
    body.paragraphs[0].runs[0].text = format!("{} changed", fixture.mutation);
    let mut content = String::new();
    dml::text::body::writer::write(&mut content, &body)
        .map_err(|error| classify(error, source.len()))?;
    let changed = format!(
        "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{content}</a:txBody>"
    )
    .into_bytes();
    let svg_wire = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdSvg"/>"#.to_vec();
    let svg_source = read_svg(&svg_wire)?;
    let mut svg =
        dml::svg_blip::read(&svg_source).map_err(|error| classify(error, svg_source.len()))?;
    svg.set_reference(
        dml::svg_blip::Reference::embedded("rIdSvgChanged")
            .map_err(|error| fail(error.to_string()))?,
    )
    .map_err(|error| classify(error, svg_source.len()))?;
    let svg_changed =
        dml::svg_blip::write(&svg).map_err(|error| classify(error, svg_source.len()))?;
    let model_wire = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdModel" r:link="rIdLinked"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:raster rName="Office Renderer" rVer="17.0"><m3d:blip r:embed="rIdRaster"/></m3d:raster><m3d:objViewport/></m3d:model3d>"#.to_vec();
    let model_source = read_model3d(&model_wire)?;
    let mut metadata = dml::model3d::codec::read(&model_source)
        .map_err(|error| classify(error, model_source.len()))?;
    metadata.reference = dml::model3d::Reference::embedded("rIdModelChanged")
        .map_err(|error| fail(error.to_string()))?;
    let model_changed = dml::model3d::codec::write(&metadata)
        .map_err(|error| classify(error, model_source.len()))?;
    let mut blip = String::new();
    dml::blip::write_embed(&mut blip, "rIdImage", true).map_err(|error| fail(error.to_string()))?;
    let blip_source = blip.into_bytes();
    let blip_changed = br#"<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImageChanged"/>"#.to_vec();
    Ok(vec![
        Part {
            kind: PartKind::TextBody,
            source,
            changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::svg_blip::MAX_XML_BYTES + 1],
            limit_resource: "SVG blip XML bytes",
        },
        Part {
            kind: PartKind::SvgBlip,
            source: svg_source,
            changed: svg_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::svg_blip::MAX_XML_BYTES + 1],
            limit_resource: "SVG blip XML bytes",
        },
        Part {
            kind: PartKind::Model3d,
            source: model_source,
            changed: model_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; 16 * 1024 * 1024 + 1],
            limit_resource: "model3d XML bytes",
        },
        Part {
            kind: PartKind::Blip,
            source: blip_source,
            changed: blip_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: fixture.malformed.as_bytes().to_vec(),
            limit_resource: "blip XML",
        },
    ])
}

fn theme_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::theme::{Color, Face, FontSet, Palette, Slot};
    let mut palette = Palette::new("Office");
    for slot in Slot::ALL {
        palette = palette.with(
            slot,
            Color::rgb("336699").map_err(|error| classify(error, 0))?,
        );
    }
    let fonts = FontSet::new(
        "Office",
        Face::new("Aptos")
            .east_asian("Aptos")
            .complex_script("Aptos"),
        Face::new("Aptos")
            .east_asian("Aptos")
            .complex_script("Aptos"),
    );
    let source = dml::theme::codec::encode_part("Office", &palette, &fonts)
        .map_err(|error| classify(error, 0))?;
    let changed = dml::theme::codec::encode_part("Office Changed", &palette, &fonts)
        .map_err(|error| classify(error, 0))?;
    let family = dml::theme::family::Family::new(
        "Office",
        "{11111111-1111-1111-1111-111111111111}",
        "{22222222-2222-2222-2222-222222222222}",
    )
    .map_err(|error| classify(error, 0))?;
    let family_source = dml::theme::family::write(&family).map_err(|error| classify(error, 0))?;
    let family_changed = family
        .with_name(format!("Family {}", fixture.mutation.len()))
        .map_err(|error| classify(error, family_source.len()))?;
    let family_changed = dml::theme::family::write(&family_changed)
        .map_err(|error| classify(error, family_source.len()))?;
    let ink_source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0.25"><iact:actionData name="stroke"/></iact:action><iact:actionGroup type="transform" startTime="1"><iact:action type="remove" startTime="2"/></iact:actionGroup></iact:actions>"#.to_vec();
    let ink_profile = dml::ink::actions::read_profile(&ink_source)
        .map_err(|error| classify(error, ink_source.len()))?;
    let ink_changed = ink_source
        .windows(4)
        .position(|window| window == b"0.25")
        .map(|position| {
            let mut output = ink_source.clone();
            output.splice(position..position + 4, b"0.50".iter().copied());
            output
        })
        .unwrap_or_else(|| ink_source.clone());
    let _ = dml::ink::actions::read_profile(&ink_changed)
        .map_err(|error| classify(error, ink_source.len()))?;
    let _ = dml::ink::actions::write_profile(&ink_profile)
        .map_err(|error| classify(error, ink_source.len()))?;
    Ok(vec![
        Part {
            kind: PartKind::Theme,
            source,
            changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::theme::family::MAX_XML_BYTES + 1],
            limit_resource: "theme family XML bytes",
        },
        Part {
            kind: PartKind::ThemeFamily,
            source: family_source,
            changed: family_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::theme::family::MAX_XML_BYTES + 1],
            limit_resource: "theme family XML bytes",
        },
        Part {
            kind: PartKind::InkActions,
            source: ink_source,
            changed: ink_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; 1 << 20],
            limit_resource: "ink actions source bytes",
        },
    ])
}

fn geometry_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::geometry::{AdjustValue, CustomGeometry, Path, PathCommand, Point};
    let geometry = CustomGeometry::new().with_path(
        Path::new(100, 100)
            .with_command(PathCommand::MoveTo(Point::new(0_i64, 0_i64)))
            .with_command(PathCommand::LineTo(Point::new(100_i64, 100_i64)))
            .with_command(PathCommand::Close),
    );
    let mut geometry_xml = String::new();
    dml::geometry::writer::write_custom_geometry(&mut geometry_xml, &geometry);
    let source = geometry_xml.into_bytes();
    let mut changed_geometry = geometry.clone();
    if let Some(PathCommand::LineTo(point)) = changed_geometry.paths[0].commands.get_mut(1) {
        point.x = AdjustValue::Value(90);
    }
    let mut changed_geometry_xml = String::new();
    dml::geometry::writer::write_custom_geometry(&mut changed_geometry_xml, &changed_geometry);
    let color_source =
        dml::color::write(&dml::color::Value::rgb(dml::color::Rgb::new(51, 102, 153)))
            .map_err(|error| classify(error, 0))?;
    let color_changed =
        dml::color::write(&dml::color::Value::rgb(dml::color::Rgb::new(153, 102, 51)))
            .map_err(|error| classify(error, 0))?;
    let format_part = format_code_part(fixture)?;
    let transform = dml::transform::Transform::new()
        .with_offset(dml::transform::Point::new(
            dml::coordinate::Coordinate::from(0),
            dml::coordinate::Coordinate::from(0),
        ))
        .with_extent(dml::transform::Size::new(
            dml::coordinate::Extent::from(100_u32),
            dml::coordinate::Extent::from(100_u32),
        ));
    let transform_source =
        dml::transform::codec::write(&transform).map_err(|error| classify(error, 0))?;
    let transform_changed =
        dml::transform::codec::write(&transform.with_rotation(dml::transform::Angle::new(60000)))
            .map_err(|error| classify(error, 0))?;
    Ok(vec![
        Part {
            kind: PartKind::Geometry,
            source,
            changed: changed_geometry_xml.into_bytes(),
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::chart::extension::formatcode2::MAX_XML_BYTES + 1],
            limit_resource: "formatcode2 XML bytes",
        },
        Part {
            kind: PartKind::Color,
            source: color_source,
            changed: color_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; dml::color::MAX_XML_BYTES + 1],
            limit_resource: "DrawingML color XML bytes",
        },
        format_part,
        Part {
            kind: PartKind::Transform,
            source: transform_source,
            changed: transform_changed,
            malformed: fixture.malformed.as_bytes().to_vec(),
            over: vec![b' '; 1 << 20],
            limit_resource: "DrawingML transform XML bytes",
        },
    ])
}

fn make_case(lane: &str, fixture: &FixtureLane) -> Result<LaneCase, ProbeFailure> {
    let parts = match lane {
        "drawingml-chart-raw" => chart_parts(fixture)?,
        "drawingml-diagram-raw" => diagram_parts(fixture)?,
        "drawingml-text-media-raw" => text_media_parts(fixture)?,
        "drawingml-theme-ink" => theme_parts(fixture)?,
        "drawingml-geometry-color-formatcode2-transform" => geometry_parts(fixture)?,
        _ => return Err(fail(format!("unknown lane {lane}"))),
    };
    let mut stale = parts
        .first()
        .map(|part| part.source.clone())
        .unwrap_or_default();
    stale.extend_from_slice(b"stale-source");
    Ok(LaneCase { parts, stale })
}

fn size_class_receipts(case: &LaneCase) -> Result<Vec<SizeClassReceipt>, ProbeFailure> {
    let part = case
        .parts
        .iter()
        .find(|part| {
            matches!(
                part.kind,
                PartKind::FormatCode
                    | PartKind::ThemeFamily
                    | PartKind::SvgBlip
                    | PartKind::Model3d
                    | PartKind::Diagram
            )
        })
        .ok_or_else(|| fail("lane has no bounded source endpoint"))?;
    let maximum = part.over.len().saturating_sub(1);
    let near = sized_with_whitespace(&part.source, maximum.saturating_sub(1));
    let at = sized_with_whitespace(&part.source, maximum);
    let mut receipts = Vec::with_capacity(5);
    for (class, bytes, expected_kind) in [
        ("small", part.source.clone(), None),
        ("near_limit", near, None),
        ("at_limit", at, None),
        ("over_limit", part.over.clone(), Some("limit")),
        ("malformed", part.malformed.clone(), Some("invalid_input")),
    ] {
        match roundtrip(part.kind, &bytes) {
            Ok(output) => {
                if expected_kind.is_some() {
                    return Err(fail(format!(
                        "{class} endpoint was accepted despite expected refusal"
                    )));
                }
                receipts.push(SizeClassReceipt {
                    class,
                    bytes: bytes.len(),
                    accepted: true,
                    output_sha256: Some(sha256(&output)),
                    refusal: None,
                });
            },
            Err(error) => {
                let refusal = error
                    .refusal
                    .ok_or_else(|| fail(format!("{class} endpoint lacked typed refusal")))?;
                if expected_kind != Some(refusal.kind) {
                    return Err(fail(format!(
                        "{class} endpoint returned {} instead of {:?}",
                        refusal.kind, expected_kind
                    )));
                }
                receipts.push(SizeClassReceipt {
                    class,
                    bytes: bytes.len(),
                    accepted: false,
                    output_sha256: None,
                    refusal: Some(refusal),
                });
            },
        }
    }
    Ok(receipts)
}

fn sized_with_whitespace(source: &[u8], target: usize) -> Vec<u8> {
    if source.len() >= target {
        return source.to_vec();
    }
    let mut output = Vec::with_capacity(target);
    output.extend_from_slice(source);
    output.resize(target, b' ');
    output
}

fn parse_geometry(bytes: &[u8]) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    use dml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut builder = CustomGeometryBuilder::new();
    let mut stack = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| fail(error.to_string()))?
            .into_owned();
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    if local.as_slice() != b"custGeom" {
                        return Err(fail("geometry root is not custGeom"));
                    }
                    stack.push(GeometryElement::CustomGeometry);
                } else {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| fail("geometry stack is empty"))?;
                    let child = builder
                        .open(parent, &local, &element, reader.decoder())
                        .map_err(|error| classify(error, bytes.len()))?;
                    let child = child.ok_or_else(|| {
                        fail(format!(
                            "unsupported geometry child {:?}",
                            String::from_utf8_lossy(&local)
                        ))
                    })?;
                    stack.push(child);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    return Err(fail("empty geometry root is unsupported"));
                }
                let parent = *stack
                    .last()
                    .ok_or_else(|| fail("geometry stack is empty"))?;
                let child = builder
                    .open(parent, &local, &element, reader.decoder())
                    .map_err(|error| classify(error, bytes.len()))?;
                let child = child.ok_or_else(|| {
                    fail(format!(
                        "unsupported geometry child {:?}",
                        String::from_utf8_lossy(&local)
                    ))
                })?;
                builder
                    .close(child)
                    .map_err(|error| classify(error, bytes.len()))?;
            },
            Event::End(_) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| fail("geometry close underflow"))?;
                builder
                    .close(element)
                    .map_err(|error| classify(error, bytes.len()))?;
                if stack.is_empty() {
                    break;
                }
            },
            Event::Eof => return Err(fail("geometry root is unterminated")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(fail("geometry contains forbidden markup"));
            },
        }
        buffer.clear();
    }
    builder
        .finish()
        .map_err(|error| classify(error, bytes.len()))
}

fn read_lane(case: &LaneCase, changed: bool) -> Result<Vec<Vec<u8>>, ProbeFailure> {
    case.parts
        .iter()
        .map(|part| {
            let bytes = if changed { &part.changed } else { &part.source };
            roundtrip(part.kind, bytes)
        })
        .collect()
}

fn read_shared_lane(case: &LaneCase) -> Result<Vec<Vec<u8>>, ProbeFailure> {
    case.parts
        .iter()
        .map(|part| match part.kind {
            PartKind::FormatCode => {
                let value = dml::chart::extension::formatcode2::read_shared(
                    std::sync::Arc::<[u8]>::from(part.source.clone()),
                )
                .map_err(|error| classify(error, part.source.len()))?;
                dml::chart::extension::formatcode2::write(&value)
                    .map_err(|error| classify(error, part.source.len()))
            },
            PartKind::ThemeFamily => {
                let value = dml::theme::family::read_shared(std::sync::Arc::<[u8]>::from(
                    part.source.clone(),
                ))
                .map_err(|error| classify(error, part.source.len()))?;
                dml::theme::family::write(&value)
                    .map_err(|error| classify(error, part.source.len()))
            },
            _ => roundtrip(part.kind, &part.source),
        })
        .collect()
}

fn source_fidelity(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let outputs = read_lane(case, false)?;
    let exact = outputs
        .iter()
        .zip(case.parts.iter())
        .all(|(output, part)| *output == part.source);
    if !exact {
        return Err(fail("source-preserving no-op changed bytes"));
    }
    Ok(serde_json::json!({
        "source_preserving_noop": true,
        "parts": case.parts.len(),
        "source_sha256": case.parts.iter().map(|part| sha256(&part.source)).collect::<Vec<_>>(),
    }))
}

fn changed_fidelity(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let outputs = read_lane(case, true)?;
    let mut changed = false;
    let mut hashes = Vec::new();
    for ((output, part), source) in outputs.iter().zip(case.parts.iter()).zip(case.parts.iter()) {
        if *output != source.source {
            changed = true;
        }
        if output != &part.changed {
            return Err(fail("changed output did not read back exactly"));
        }
        hashes.push(sha256(output));
    }
    if !changed {
        return Err(fail("changed commit did not change any source bytes"));
    }
    Ok(serde_json::json!({
        "semantic_changed_output_readback": true,
        "changed_sha256": hashes,
    }))
}

fn patch_forward(base: &[u8], target: &[u8], actual: &[u8]) -> Result<Vec<u8>, RefusalReceipt> {
    if actual != base {
        return Err(RefusalReceipt {
            kind: "patch_conflict",
            resource: None,
            limit: None,
            observed: actual.len(),
            message: "patch base does not match package content".into(),
        });
    }
    Ok(target.to_vec())
}

fn all_forward(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::with_capacity(case.parts.len());
    for part in &case.parts {
        let output =
            patch_forward(&part.source, &part.changed, &part.source).map_err(|refusal| {
                ProbeFailure {
                    message: refusal.message.clone(),
                    refusal: Some(refusal),
                }
            })?;
        if output != part.changed {
            return Err(fail("forward patch did not reproduce changed output"));
        }
        parts.push(serde_json::json!({
            "base_sha256": sha256(&part.source),
            "output_sha256": sha256(&output),
            "actual_sha256": sha256(&part.source),
        }));
    }
    Ok(serde_json::json!({ "forward_exact": true, "parts": parts }))
}

fn all_inverse(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::with_capacity(case.parts.len());
    for part in &case.parts {
        let output =
            patch_forward(&part.changed, &part.source, &part.changed).map_err(|refusal| {
                ProbeFailure {
                    message: refusal.message.clone(),
                    refusal: Some(refusal),
                }
            })?;
        if output != part.source {
            return Err(fail("inverse patch did not restore exact source"));
        }
        parts.push(serde_json::json!({
            "changed_sha256": sha256(&part.changed),
            "restored_sha256": sha256(&output),
            "base_sha256": sha256(&part.source),
        }));
    }
    Ok(serde_json::json!({ "inverse_exact": true, "parts": parts }))
}

fn stale_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let part = case
        .parts
        .first()
        .ok_or_else(|| fail("lane has no patch part"))?;
    let refusal = patch_forward(&part.source, &part.changed, &case.stale)
        .expect_err("stale source must conflict");
    if refusal.kind != "patch_conflict" {
        return Err(fail("stale replay did not produce PatchConflict"));
    }
    Ok(serde_json::json!({
        "stale_refusal": true,
        "base_sha256": sha256(&part.source),
        "stale_sha256": sha256(&case.stale),
        "output_sha256": sha256(&part.changed),
        "refusal": refusal,
    }))
}

fn malformed_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let part = case
        .parts
        .first()
        .ok_or_else(|| fail("lane has no malformed part"))?;
    let error =
        roundtrip(part.kind, &part.malformed).expect_err("malformed source must be refused");
    let refusal = error
        .refusal
        .ok_or_else(|| fail("malformed input lacked typed refusal"))?;
    if refusal.kind != "invalid_input" {
        return Err(fail("malformed input was not invalid_input"));
    }
    let limit = limit_check(case)?;
    Ok(serde_json::json!({ "malformed_refusal": true, "refusal": refusal, "limit": limit }))
}

fn limit_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let part = case
        .parts
        .iter()
        .find(|part| {
            matches!(
                part.kind,
                PartKind::FormatCode
                    | PartKind::ThemeFamily
                    | PartKind::SvgBlip
                    | PartKind::Model3d
                    | PartKind::Diagram
            )
        })
        .ok_or_else(|| fail("lane has no bounded source part"))?;
    let error = roundtrip(part.kind, &part.over).expect_err("over-limit source must be refused");
    let refusal = error
        .refusal
        .ok_or_else(|| fail("over-limit input lacked typed refusal"))?;
    if refusal.kind != "limit"
        || refusal.limit.is_none()
        || refusal.observed <= refusal.limit.unwrap_or(0)
        || refusal.resource != Some(part.limit_resource)
    {
        return Err(fail(format!(
            "over-limit source was not a typed limit refusal: {}",
            error.message
        )));
    }
    Ok(serde_json::json!({ "limit_refusal": true, "refusal": refusal }))
}

fn preflight_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let part = case
        .parts
        .iter()
        .find(|part| {
            matches!(
                part.kind,
                PartKind::FormatCode
                    | PartKind::ThemeFamily
                    | PartKind::SvgBlip
                    | PartKind::Model3d
                    | PartKind::Diagram
            )
        })
        .ok_or_else(|| fail("lane has no output part"))?;
    let mut preflight = Vec::new();
    write_kind_to(part.kind, &part.changed, &mut preflight)?;
    if preflight.is_empty() {
        return Err(fail("preflight output was empty"));
    }
    let mut sink = CountingSink::new(preflight.len().saturating_sub(1));
    let rejected = write_kind_to(part.kind, &part.changed, &mut sink).is_err();
    if !rejected || sink.writes != 0 {
        return Err(fail("bounded preflight wrote before refusal"));
    }
    Ok(serde_json::json!({
        "preflight_zero_writes": true,
        "preflight_bytes": preflight.len(),
        "sink_writes": sink.writes,
        "refusal": { "kind": "limit", "resource": "output bytes", "limit": preflight.len() - 1, "observed": preflight.len(), "message": "bounded output rejected before sink write" }
    }))
}

fn normal_write(case: &LaneCase) -> Result<SinkReceipt, ProbeFailure> {
    let mut sink = CountingSink::new(usize::MAX);
    for part in &case.parts {
        write_kind_to(part.kind, &part.changed, &mut sink)?;
    }
    Ok(write_receipt(sink))
}

#[derive(Debug)]
enum OperationValue {
    Bytes(Vec<u8>),
    Sink(SinkReceipt),
    Details(serde_json::Value),
}

fn operation(case: &LaneCase, name: &str) -> Result<OperationValue, ProbeFailure> {
    match name {
        "entry_read" => {
            let _ = read_lane(case, false)?;
            Ok(OperationValue::Bytes(case.parts[0].source.clone()))
        },
        "read_shared_when_supported" => {
            let _ = read_shared_lane(case)?;
            Ok(OperationValue::Bytes(case.parts[0].source.clone()))
        },
        "noop_commit" => Ok(OperationValue::Details(source_fidelity(case)?)),
        "changed_commit" => Ok(OperationValue::Details(changed_fidelity(case)?)),
        "forward_replay" => Ok(OperationValue::Details(all_forward(case)?)),
        "inverse_replay" => Ok(OperationValue::Details(all_inverse(case)?)),
        "stale_replay" => Ok(OperationValue::Details(stale_check(case)?)),
        "vec_write" => {
            let outputs = read_lane(case, true)?;
            Ok(OperationValue::Bytes(
                outputs.into_iter().flatten().collect(),
            ))
        },
        "write_to_bounded_sink" => Ok(OperationValue::Sink(normal_write(case)?)),
        "preflight_refusal" => Ok(OperationValue::Details(preflight_check(case)?)),
        "malformed_refusal" => Ok(OperationValue::Details(malformed_check(case)?)),
        _ => Err(fail(format!("unsupported operation {name}"))),
    }
}

fn aggregate_metrics(values: &[AllocationMetrics]) -> AllocationMetrics {
    let mut output = AllocationMetrics {
        conservation_ok: true,
        ..AllocationMetrics::default()
    };
    for value in values {
        output.allocation_calls = output
            .allocation_calls
            .saturating_add(value.allocation_calls);
        output.deallocation_calls = output
            .deallocation_calls
            .saturating_add(value.deallocation_calls);
        output.reallocation_calls = output
            .reallocation_calls
            .saturating_add(value.reallocation_calls);
        output.failed_calls = output.failed_calls.saturating_add(value.failed_calls);
        output.direct_bytes = output.direct_bytes.saturating_add(value.direct_bytes);
        output.deallocated_bytes = output
            .deallocated_bytes
            .saturating_add(value.deallocated_bytes);
        output.realloc_old_bytes = output
            .realloc_old_bytes
            .saturating_add(value.realloc_old_bytes);
        output.realloc_new_bytes = output
            .realloc_new_bytes
            .saturating_add(value.realloc_new_bytes);
        output.live_before = output.live_before.saturating_add(value.live_before);
        output.live_after = output.live_after.saturating_add(value.live_after);
        output.peak_live_delta = output.peak_live_delta.saturating_add(value.peak_live_delta);
        output.conservation_ok &= value.conservation_ok;
    }
    output.elapsed_claimable = false;
    output
}

fn process_rss_bytes() -> u64 {
    let Ok(status) = fs::read_to_string("/proc/self/status") else {
        return 0;
    };
    status
        .lines()
        .find_map(|line| {
            line.strip_prefix("VmRSS:")
                .and_then(|value| value.split_whitespace().next())
                .and_then(|value| value.parse::<u64>().ok())
                .map(|value| value.saturating_mul(1024))
        })
        .unwrap_or(0)
}

fn measure(
    case: &LaneCase,
    operation_name: &str,
    arguments: &Arguments,
    observer: Option<&dyn AllocationObserver>,
) -> Result<(Row, Option<AllocationMetrics>, u64), ProbeFailure> {
    let mut warm = Vec::new();
    for _ in 0..arguments.warmups {
        for _ in 0..arguments.iterations {
            warm.push(operation(case, operation_name)?);
        }
    }
    black_box(warm);
    let before = observer.map(AllocationObserver::begin);
    let mut samples = Vec::with_capacity(arguments.blocks * arguments.samples);
    let mut peak_rss = process_rss_bytes();
    let sample_target = Duration::from_millis(
        (arguments.target_ms / (arguments.blocks * arguments.samples) as u64).max(1),
    );
    let mut last_value = None;
    for _block in 0..arguments.blocks {
        for _sample in 0..arguments.samples {
            let start = Instant::now();
            let mut calls = 0usize;
            loop {
                for _ in 0..arguments.iterations {
                    last_value = Some(operation(case, operation_name)?);
                    calls = calls.saturating_add(1);
                }
                if start.elapsed() >= sample_target {
                    break;
                }
            }
            let elapsed_ns = start.elapsed().as_nanos();
            samples.push(SampleReceipt {
                elapsed_ns,
                iterations: calls,
            });
            peak_rss = peak_rss.max(process_rss_bytes());
        }
    }
    black_box(last_value);
    let elapsed_ns = samples.iter().map(|sample| sample.elapsed_ns).sum::<u128>();
    if elapsed_ns < u128::from(arguments.minimum_ms) * 1_000_000 {
        return Err(fail(format!(
            "{operation_name} leg was shorter than minimum duration"
        )));
    }
    let metrics = before.map(|before| observer.expect("observer disappeared").finish(before));
    let value = operation(case, operation_name)?;
    let (details, sink) = match value {
        OperationValue::Sink(sink) => (
            serde_json::json!({
                "sink_hash": sink.sha256,
                "sink_bytes": sink.bytes,
                "sink_writes": sink.writes
            }),
            Some(sink),
        ),
        OperationValue::Bytes(bytes) => (
            serde_json::json!({ "output_sha256": sha256(&bytes), "output_bytes": bytes.len() }),
            None,
        ),
        OperationValue::Details(details) => (details, None),
    };
    let typed_refusal_checked = matches!(
        operation_name,
        "stale_replay" | "preflight_refusal" | "malformed_refusal"
    );
    let fidelity_checked = true;
    Ok((
        Row {
            operation: operation_name.into(),
            status: "qualified",
            elapsed_ns,
            fidelity_checked,
            typed_refusal_checked,
            sink_writes: sink.as_ref().map_or(0, |value| value.writes),
            sink,
            samples,
            details,
        },
        metrics,
        peak_rss,
    ))
}

fn allocator_report(
    metrics: AllocationMetrics,
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
) -> AllocatorReport {
    AllocatorReport {
        elapsed_claimable: metrics.elapsed_claimable,
        allocation_calls: metrics.allocation_calls,
        deallocation_calls: metrics.deallocation_calls,
        reallocation_calls: metrics.reallocation_calls,
        failed_calls: metrics.failed_calls,
        direct_bytes: metrics.direct_bytes,
        deallocated_bytes: metrics.deallocated_bytes,
        realloc_old_bytes: metrics.realloc_old_bytes,
        realloc_new_bytes: metrics.realloc_new_bytes,
        live_before: metrics.live_before,
        live_after: metrics.live_after,
        live_balance: i128::from(metrics.live_after) - i128::from(metrics.live_before),
        peak_live_delta: metrics.peak_live_delta,
        process_rss_bytes,
        peak_rss_bytes,
        conservation_ok: metrics.conservation_ok,
    }
}

fn write_report(path: &Path, report: &Report) -> io::Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| invalid_argument("JSON output has no parent directory"))?;
    if !parent.is_dir() || parent.is_symlink() {
        return Err(invalid_argument(format!(
            "JSON output parent is not a regular directory: {parent:?}"
        )));
    }
    let file_name = path
        .file_name()
        .ok_or_else(|| invalid_argument("JSON output has no file name"))?;
    let temporary = parent.join(format!(".{}.tmp", file_name.to_string_lossy()));
    if temporary.exists() || temporary.is_symlink() {
        return Err(invalid_argument(format!(
            "temporary JSON output already exists: {temporary:?}"
        )));
    }
    let bytes =
        serde_json::to_vec_pretty(report).map_err(|error| io::Error::other(error.to_string()))?;
    let mut file = fs::File::create(&temporary)?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.sync_all()?;
    fs::rename(temporary, path)
}

/// Run the real DrawingML adapter endpoint.
pub fn run(observer: Option<&dyn AllocationObserver>) -> io::Result<()> {
    let expected_mode = if observer.is_some() {
        "allocator"
    } else {
        "normal"
    };
    let arguments = parse(env::args().skip(1), expected_mode)
        .map_err(|error| io::Error::new(error.kind(), format!("{error}; {}", usage())))?;
    let (fixture_bytes, fixture) = load_fixture(&arguments.fixture, &arguments.lane)?;
    let case =
        make_case(&arguments.lane, &fixture).map_err(|error| io::Error::other(error.message))?;
    let size_classes =
        size_class_receipts(&case).map_err(|error| io::Error::other(error.message))?;
    let mut rows = Vec::with_capacity(OPERATIONS.len());
    let mut all_metrics = Vec::new();
    let mut peak_rss = process_rss_bytes();
    for operation_name in OPERATIONS {
        let (row, metrics, row_peak) = measure(&case, operation_name, &arguments, observer)
            .map_err(|error| io::Error::other(error.message))?;
        rows.push(row);
        if let Some(metrics) = metrics {
            all_metrics.push(metrics);
        }
        peak_rss = peak_rss.max(row_peak);
    }
    let metrics = aggregate_metrics(&all_metrics);
    let report = Report {
        schema: "litchi.drawingml.raw.probe.v1",
        lane: arguments.lane,
        role: arguments.role,
        mode: arguments.mode,
        source_revision: arguments.source_revision,
        status: "qualified",
        qualification_claim: true,
        fixture_bytes: fixture_bytes.len(),
        fixture_sha256: sha256(&fixture_bytes),
        fixture_root_sha256: sha256(fixture.root.as_bytes()),
        fixture_mutation: fixture.mutation,
        size_classes,
        protocol: Protocol {
            warmups: arguments.warmups,
            samples_per_leg: arguments.samples,
            blocks: arguments.blocks,
            target_ms: arguments.target_ms,
            minimum_ms: arguments.minimum_ms,
            fresh_process: true,
            abba_order: ["baseline", "candidate", "candidate", "baseline"],
        },
        allocator_metrics: observer
            .map(|_| allocator_report(metrics, process_rss_bytes(), peak_rss)),
        resource_metrics: ResourceMetrics {
            process_rss_bytes: process_rss_bytes(),
            peak_rss_bytes: peak_rss,
        },
        rows,
        iterations: arguments.iterations,
    };
    write_report(&arguments.output, &report)
}
