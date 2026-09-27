//! Standalone public-workflow probe for change 0785.
//!
//! The corpus and the public PPTX calls are copied from the deterministic
//! 0760 probe.  This executable is deliberately independent of the workspace
//! harness: the build driver materializes its manifest, and the optional
//! allocator target installs instrumentation only in this process.
//!
//! The timed regions are deliberately small and explicit:
//!
//! * `capture` measures `Package::opened_presentation`.
//! * `commit` stages one edit before the clock and measures `commit`.
//! * `lifecycle` measures capture, edit, commit, publication, and
//!   serialization as one public operation.
//!
//! Package serialization, reopening, marker checks, and destruction happen
//! after the timed region unless they are explicitly part of `lifecycle`.
//! Owners remain live until the allocation region has finished.

// Shared allocation diagnostics expose APIs unused by this census-only binary.
#[allow(dead_code)]
mod allocation_metrics;

#[cfg(feature = "allocator-metrics")]
mod counting_allocator;

use std::collections::BTreeMap;
#[cfg(feature = "attribute-census")]
use std::collections::BTreeSet;
use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::path::PathBuf;
use std::time::Instant;

use litchi_pptx::Package;
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.pptx.attribute-census-probe.v1";
const TOOL: &str = "attribute-census-probe-0798";
const MARKER: &str = "litchi-perf-0780-static-mce-capabilities";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;

const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const KNOWN_URI_NOTES: &[(&str, &str)] = &[
    (
        "P",
        "http://schemas.openxmlformats.org/presentationml/2006/main",
    ),
    ("PS", "http://purl.oclc.org/ooxml/presentationml/main"),
    ("A", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    ("AS", "http://purl.oclc.org/ooxml/drawingml/main"),
    (
        "R",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "RS",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ),
];
const VENDOR_PREFIXES: &[&str] = &["vP", "vPS", "vA", "vAS", "vR", "vRS"];
const ATTRIBUTE_LOCAL_NAMES: &[&str] = &[
    "probeP", "probePS", "probeA", "probeAS", "probeR", "probeRS",
];

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Capture,
    Commit,
    Lifecycle,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "capture" => Ok(Self::Capture),
            "commit" => Ok(Self::Commit),
            "lifecycle" => Ok(Self::Lifecycle),
            _ => Err(
                format!("--mode must be one of capture, commit, lifecycle (got {value:?})").into(),
            ),
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Capture => "Package::opened_presentation only",
            Self::Commit => {
                "Transaction::commit only; package capture and one set_shape_text staging are outside the clock"
            },
            Self::Lifecycle => {
                "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes"
            },
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Shape {
    Tiny,
    Medium,
    Large,
    Vendor,
    UnicodeVendor,
}

impl Shape {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "tiny" => Ok(Self::Tiny),
            "medium" => Ok(Self::Medium),
            "large" => Ok(Self::Large),
            "vendor" => Ok(Self::Vendor),
            "unicode-vendor" => Ok(Self::UnicodeVendor),
            _ => Err(format!(
                "--shape must be one of tiny, medium, large, vendor, unicode-vendor (got {value:?})"
            )
            .into()),
        }
    }

    const fn dimensions(self) -> (usize, usize) {
        match self {
            Self::Tiny => (3, 4),
            Self::Medium => (12, 8),
            Self::Large => (100, 100),
            Self::Vendor | Self::UnicodeVendor => (12, 8),
        }
    }

    const fn has_unknown_uris(self) -> bool {
        matches!(self, Self::Vendor | Self::UnicodeVendor)
    }

    const fn injection_name(self) -> &'static str {
        match self {
            Self::Tiny | Self::Medium | Self::Large => "none",
            Self::Vendor => "same-length-known-uri-near-misses",
            Self::UnicodeVendor => "same-length-valid-utf8-unknown-uris",
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
    semantic_check: bool,
    reopened: bool,
    expected_text: Option<String>,
    actual_text: Option<String>,
    semantic_text_bytes: Option<u64>,
    semantic_text_sha256: Option<String>,
    readback_bytes: Option<u64>,
    readback_sha256: Option<String>,
    marker_matches: Option<bool>,
    unknown_namespace_check: Option<bool>,
    unknown_namespace_occurrences: Option<u64>,
}

#[derive(Debug, Serialize, Clone)]
struct FixtureMetadata {
    injection: &'static str,
    slide_parts: usize,
    replaced_text_tags: usize,
    namespace_declarations: usize,
    namespaced_attributes: usize,
    namespace_uris: Vec<String>,
    attribute_names: Vec<String>,
}

#[derive(Debug)]
struct Fixture {
    bytes: Vec<u8>,
    metadata: FixtureMetadata,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    metrics: BTreeMap<String, u64>,
    source_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    output: Option<Identity>,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    census: Option<AttributeCensusReport0798>,
    verification: Verification,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    mode: Mode,
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
    timing_scope: &'static str,
    marker: &'static str,
    source: Identity,
    fixture: FixtureMetadata,
    warmup: usize,
    samples_requested: usize,
    samples: Vec<SampleRecord>,
    allocator: AllocatorIdentity,
}

#[derive(Debug)]
struct OperationResult {
    elapsed_ns: u64,
    output: Option<Vec<u8>>,
    verification: Verification,
    allocation: Option<allocation_metrics::Sample>,
    census: Option<AttributeCensusReport0798>,
    extra_metrics: BTreeMap<String, u64>,
}

/// Lossless serialization of the OPC-only checked-iterator census.  The OPC
/// crate deliberately keeps this diagnostic API serde-free; this probe owns
/// the evidence schema and copies every counter and raw tag byte here.
#[derive(Debug, Serialize)]
struct AttributeCensusReport0798 {
    owner: &'static str,
    iterator_starts: u64,
    iterator_clones: u64,
    iterator_drops: u64,
    live_instances_at_finish: u64,
    counter_saturated: bool,
    raw_instance_rows: u64,
    instance_identity_qualified: bool,
    aggregation_conserved: bool,
    rows: Vec<AttributeCensusRow0798>,
}

#[derive(Debug, Serialize)]
struct AttributeCensusRow0798 {
    frequency: u64,
    is_clone: bool,
    starting_successful_yields: u64,
    tag_bytes: Vec<u8>,
    element_name_bytes: Vec<u8>,
    next_calls: u64,
    successful_yields: u64,
    error_yields: u64,
    end_yields: u64,
    lexical_attribute_count: u64,
    lexical_item_count: u64,
    lexical_error_count: u64,
    lexical_scan_completed: bool,
    termination: &'static str,
    dropped: bool,
    early_drop: bool,
    partial_consumption: bool,
    live_at_finish: bool,
    counter_saturated: bool,
}

#[cfg(feature = "attribute-census")]
#[derive(Debug, Eq, Ord, PartialEq, PartialOrd)]
struct AttributeCensusAggregateKey0798 {
    is_clone: bool,
    starting_successful_yields: u64,
    tag_bytes: Vec<u8>,
    element_name_bytes: Vec<u8>,
    next_calls: u64,
    successful_yields: u64,
    error_yields: u64,
    end_yields: u64,
    lexical_attribute_count: u64,
    lexical_item_count: u64,
    lexical_error_count: u64,
    lexical_scan_completed: bool,
    termination: u8,
    dropped: bool,
    early_drop: bool,
    partial_consumption: bool,
    live_at_finish: bool,
    counter_saturated: bool,
}

#[cfg(feature = "attribute-census")]
struct AttributeCensusScope0798 {
    active: bool,
}

#[cfg(feature = "attribute-census")]
impl AttributeCensusScope0798 {
    fn begin() -> Self {
        litchi_opc::xml_attributes::begin_attribute_census_0798();
        Self { active: true }
    }

    fn finish(mut self) -> Option<AttributeCensusReport0798> {
        self.active = false;
        Some(convert_attribute_census_0798(
            litchi_opc::xml_attributes::finish_attribute_census_0798(),
        ))
    }
}

#[cfg(feature = "attribute-census")]
impl Drop for AttributeCensusScope0798 {
    fn drop(&mut self) {
        if self.active {
            // Error paths cannot publish a sample.  Finish and discard the
            // report so a stale enabled session cannot affect the next call.
            let _ = litchi_opc::xml_attributes::finish_attribute_census_0798();
            self.active = false;
        }
    }
}

#[cfg(not(feature = "attribute-census"))]
struct AttributeCensusScope0798;

#[cfg(not(feature = "attribute-census"))]
impl AttributeCensusScope0798 {
    const fn begin() -> Self {
        Self
    }

    fn finish(self) -> Option<AttributeCensusReport0798> {
        None
    }
}

#[cfg(feature = "attribute-census")]
fn convert_attribute_census_0798(
    report: litchi_opc::xml_attributes::AttributeCensusReport0798,
) -> AttributeCensusReport0798 {
    let raw_instance_rows = u64::try_from(report.rows.len()).unwrap_or(u64::MAX);
    let (starts, clones) = report
        .rows
        .iter()
        .fold((0u64, 0u64), |(starts, clones), row| {
            if row.is_clone {
                (starts, clones.saturating_add(1))
            } else {
                (starts.saturating_add(1), clones)
            }
        });
    let instance_identity_qualified = starts == report.iterator_starts
        && clones == report.iterator_clones
        && starts.checked_add(clones) == Some(raw_instance_rows)
        && report
            .iterator_drops
            .checked_add(report.live_instances_at_finish)
            == Some(raw_instance_rows)
        && qualify_attribute_census_instances_0798(&report.rows);
    let mut counter_saturated = report.counter_saturated;
    let mut grouped = BTreeMap::<AttributeCensusAggregateKey0798, u64>::new();
    for row in report.rows {
        let termination = match row.termination {
            litchi_opc::xml_attributes::AttributeCensusTermination0798::Active => 0,
            litchi_opc::xml_attributes::AttributeCensusTermination0798::Error => 1,
            litchi_opc::xml_attributes::AttributeCensusTermination0798::Exhausted => 2,
            litchi_opc::xml_attributes::AttributeCensusTermination0798::Dropped => 3,
            litchi_opc::xml_attributes::AttributeCensusTermination0798::LiveAtFinish => 4,
        };
        let key = AttributeCensusAggregateKey0798 {
            is_clone: row.is_clone,
            starting_successful_yields: row.starting_successful_yields,
            tag_bytes: row.source,
            element_name_bytes: row.element_name,
            next_calls: row.next_calls,
            successful_yields: row.successful_yields,
            error_yields: row.error_yields,
            end_yields: row.end_yields,
            lexical_attribute_count: row.lexical_attribute_count,
            lexical_item_count: row.lexical_item_count,
            lexical_error_count: row.lexical_error_count,
            lexical_scan_completed: row.lexical_scan_completed,
            termination,
            dropped: row.dropped,
            early_drop: row.early_drop,
            partial_consumption: row.partial_consumption,
            live_at_finish: row.live_at_finish,
            counter_saturated: row.counter_saturated,
        };
        let frequency = grouped.entry(key).or_insert(0);
        if let Some(next) = frequency.checked_add(1) {
            *frequency = next;
        } else {
            *frequency = u64::MAX;
            counter_saturated = true;
        }
    }
    let mut aggregate_rows = 0u64;
    let mut aggregation_conserved = true;
    for frequency in grouped.values() {
        if let Some(next) = aggregate_rows.checked_add(*frequency) {
            aggregate_rows = next;
        } else {
            aggregation_conserved = false;
            break;
        }
    }
    aggregation_conserved &= aggregate_rows == raw_instance_rows;
    AttributeCensusReport0798 {
        owner: "litchi-opc::xml_attributes::CheckedAttributes",
        iterator_starts: report.iterator_starts,
        iterator_clones: report.iterator_clones,
        iterator_drops: report.iterator_drops,
        live_instances_at_finish: report.live_instances_at_finish,
        counter_saturated,
        raw_instance_rows,
        instance_identity_qualified,
        aggregation_conserved,
        rows: grouped
            .into_iter()
            .map(|(row, frequency)| AttributeCensusRow0798 {
                frequency,
                is_clone: row.is_clone,
                starting_successful_yields: row.starting_successful_yields,
                tag_bytes: row.tag_bytes,
                element_name_bytes: row.element_name_bytes,
                next_calls: row.next_calls,
                successful_yields: row.successful_yields,
                error_yields: row.error_yields,
                end_yields: row.end_yields,
                lexical_attribute_count: row.lexical_attribute_count,
                lexical_item_count: row.lexical_item_count,
                lexical_error_count: row.lexical_error_count,
                lexical_scan_completed: row.lexical_scan_completed,
                termination: match row.termination {
                    0 => "active",
                    1 => "error",
                    2 => "exhausted",
                    3 => "dropped",
                    4 => "live-at-finish",
                    _ => unreachable!("invalid OPC census termination code"),
                },
                dropped: row.dropped,
                early_drop: row.early_drop,
                partial_consumption: row.partial_consumption,
                live_at_finish: row.live_at_finish,
                counter_saturated: row.counter_saturated,
            })
            .collect(),
    }
}

#[cfg(feature = "attribute-census")]
fn qualify_attribute_census_instances_0798(
    rows: &[litchi_opc::xml_attributes::AttributeCensusRow0798],
) -> bool {
    let mut instance_ids = BTreeSet::new();
    let mut lineage_origins = BTreeMap::<u64, (Vec<u8>, Vec<u8>)>::new();
    for row in rows {
        if row.instance_id == 0
            || row.lineage_id == 0
            || !instance_ids.insert(row.instance_id)
            || row.successful_yields > row.lexical_attribute_count
            || row
                .starting_successful_yields
                .checked_add(row.successful_yields)
                .is_none_or(|consumed| consumed > row.lexical_attribute_count)
        {
            return false;
        }
        if row.is_clone {
            let Some((source, element_name)) = lineage_origins.get(&row.lineage_id) else {
                return false;
            };
            if source != &row.source || element_name != &row.element_name {
                return false;
            }
            if row.starting_successful_yields > row.lexical_attribute_count {
                return false;
            }
        } else if row.starting_successful_yields != 0
            || lineage_origins
                .insert(
                    row.lineage_id,
                    (row.source.clone(), row.element_name.clone()),
                )
                .is_some()
        {
            return false;
        }
    }
    true
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

    let (slides, shapes_per_slide) = config.shape.dimensions();
    let fixture = build(config.shape)?;
    let source_bytes = &fixture.bytes;
    let source = identity(source_bytes)?;
    let source_sha256 = source.sha256.clone();

    let mut samples = Vec::with_capacity(config.samples);
    let mut deterministic_output = None;
    for _ in 0..config.warmup {
        let result = run_one(
            &config,
            source_bytes,
            config.shape,
            slides,
            shapes_per_slide,
            0,
        )?;
        observe_output(&mut deterministic_output, result.output.as_deref())?;
        black_box(result);
    }
    for index in 0..config.samples {
        let result = run_one(
            &config,
            source_bytes,
            config.shape,
            slides,
            shapes_per_slide,
            index,
        )?;
        observe_output(&mut deterministic_output, result.output.as_deref())?;
        let output = result.output.as_deref().map(identity).transpose()?;
        let mut metrics = result.extra_metrics;
        metrics.insert("elapsed_ns".to_owned(), result.elapsed_ns);
        metrics.insert("slides".to_owned(), slides as u64);
        metrics.insert("shapes_per_slide".to_owned(), shapes_per_slide as u64);
        samples.push(SampleRecord {
            index,
            elapsed_ns: result.elapsed_ns,
            metrics,
            source_sha256: source_sha256.clone(),
            output,
            allocation: result.allocation,
            census: result.census,
            verification: result.verification,
        });
    }

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        mode: config.mode,
        shape: config.shape,
        slides,
        shapes_per_slide,
        timing_scope: config.mode.timing_scope(),
        marker: MARKER,
        source,
        fixture: fixture.metadata,
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

fn run_one(
    config: &Config,
    source_bytes: &[u8],
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
    _index: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    match config.mode {
        Mode::Capture => run_capture(source_bytes, shape, slides, shapes_per_slide),
        Mode::Commit => run_commit(source_bytes, shape, slides, shapes_per_slide),
        Mode::Lifecycle => run_lifecycle(source_bytes, shape, slides, shapes_per_slide),
    }
}

fn run_capture(
    source_bytes: &[u8],
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    let mut package = Package::from_bytes(source_bytes)?;
    let region = allocation_metrics::begin();
    let census = AttributeCensusScope0798::begin();
    let started = Instant::now();
    #[cfg(feature = "capture-profile")]
    let snapshot = capture_region_0793(&package)?;
    #[cfg(not(feature = "capture-profile"))]
    let snapshot = package.opened_presentation()?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    let census = census.finish();
    black_box(&snapshot);

    if snapshot.slides().len() != slides {
        return Err(format!(
            "capture readback found {} slides, expected {slides}",
            snapshot.slides().len()
        )
        .into());
    }
    let output = package.to_bytes()?;
    let verification = verify_output(&output, shape, slides, shapes_per_slide, false)?;
    let mut extra_metrics = BTreeMap::new();
    extra_metrics.insert("captured_slides".to_owned(), snapshot.slides().len() as u64);
    extra_metrics.insert(
        "captured_shapes_per_slide".to_owned(),
        shapes_per_slide as u64,
    );
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        census,
        extra_metrics,
    })
}

/// Keep the public capture operation in a distinct, non-inlined owner for the
/// profile variant.  The opaque use of the result is inside this wrapper so
/// the call cannot be reduced to a tail call into `Package::opened_presentation`.
#[cfg(feature = "capture-profile")]
#[inline(never)]
fn capture_region_0793(package: &Package) -> litchi_pptx::Result<litchi_pptx::opened::Snapshot> {
    let result = package.opened_presentation();
    black_box(&result);
    result
}

fn run_commit(
    source_bytes: &[u8],
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    let mut package = Package::from_bytes(source_bytes)?;
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    if !edit.set_shape_text(0, 0, MARKER)? {
        return Err("commit setup did not stage a changed marker".into());
    }

    let region = allocation_metrics::begin();
    let census = AttributeCensusScope0798::begin();
    let started = Instant::now();
    let commit = edit.commit()?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    let census = census.finish();
    black_box(&commit);
    if !commit.is_changed() {
        return Err("commit produced an empty patch".into());
    }

    package.apply_opened_presentation_commit(commit)?;
    let output = package.to_bytes()?;
    let verification = verify_output(&output, shape, slides, shapes_per_slide, true)?;
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        census,
        extra_metrics: BTreeMap::new(),
    })
}

fn run_lifecycle(
    source_bytes: &[u8],
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    // Package ownership and source ingress are intentionally outside this
    // clock.  The public capture/edit/publication route is the measured unit.
    let mut package = Package::from_bytes(source_bytes)?;
    let region = allocation_metrics::begin();
    let census = AttributeCensusScope0798::begin();
    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    if !edit.set_shape_text(0, 0, MARKER)? {
        return Err("lifecycle setup did not stage a changed marker".into());
    }
    let commit = edit.commit()?;
    package.apply_opened_presentation_commit(commit)?;
    let output = package.to_bytes()?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    let census = census.finish();
    black_box(&package);

    let verification = verify_output(&output, shape, slides, shapes_per_slide, true)?;
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        census,
        extra_metrics: BTreeMap::new(),
    })
}

fn verify_output(
    output: &[u8],
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
    marker_operation: bool,
) -> Result<Verification, Box<dyn Error>> {
    let mut reopened = Package::from_bytes(output)?;
    let actual_text = reopened.presentation()?.text()?;
    let expected_full_text = expected_text(slides, shapes_per_slide, marker_operation);
    let semantic_check = actual_text == expected_full_text;
    let first_expected = if marker_operation {
        MARKER.to_owned()
    } else {
        source_text(0, 0)
    };
    if !semantic_check {
        return Err(format!(
            "PPTX readback differed from the exact generated text; expected {} bytes, got {} bytes",
            expected_full_text.len(),
            actual_text.len()
        )
        .into());
    }
    let unknown_namespace_occurrences = if shape.has_unknown_uris() {
        Some(u64::try_from(verify_unknown_namespace_payload(
            &mut reopened,
            shape,
            slides,
            shapes_per_slide,
        )?)?)
    } else {
        None
    };
    Ok(Verification {
        semantic_check,
        reopened: true,
        expected_text: Some(first_expected.clone()),
        actual_text: Some(first_expected),
        semantic_text_bytes: Some(u64::try_from(actual_text.len())?),
        semantic_text_sha256: Some(sha256_hex(actual_text.as_bytes())),
        // The readback identity binds the semantic check to the exact
        // reopened package that the sample published.  The selected text is
        // retained separately in `actual_text`; the full semantic text is
        // deliberately represented by a digest rather than copied into every
        // large-sample report.
        readback_bytes: Some(u64::try_from(output.len())?),
        readback_sha256: Some(sha256_hex(output)),
        marker_matches: marker_operation.then_some(true),
        unknown_namespace_check: shape.has_unknown_uris().then_some(true),
        unknown_namespace_occurrences,
    })
}

#[derive(Debug)]
struct InjectionEntry {
    uri: String,
    declaration: String,
    attribute_name: String,
    attribute: String,
}

fn injection_entries(shape: Shape) -> Result<Vec<InjectionEntry>, Box<dyn Error>> {
    if !shape.has_unknown_uris() {
        return Ok(Vec::new());
    }
    KNOWN_URI_NOTES
        .iter()
        .zip(VENDOR_PREFIXES.iter().copied())
        .zip(ATTRIBUTE_LOCAL_NAMES.iter().copied())
        .map(|(((note, known_uri), prefix), local_name)| {
            let uri = match shape {
                Shape::Vendor => near_miss_uri(known_uri)?,
                Shape::UnicodeVendor => unicode_vendor_uri(known_uri.len())?,
                Shape::Tiny | Shape::Medium | Shape::Large => unreachable!(),
            };
            let declaration = format!(r#"xmlns:{prefix}="{uri}""#);
            let attribute_name = format!("{prefix}:{local_name}");
            let attribute = format!(r#"{attribute_name}="litchi-perf-0785-{note}""#);
            Ok(InjectionEntry {
                uri,
                declaration,
                attribute_name,
                attribute,
            })
        })
        .collect()
}

fn near_miss_uri(known_uri: &str) -> Result<String, Box<dyn Error>> {
    let mut bytes = known_uri.as_bytes().to_vec();
    if bytes.first() != Some(&b'h') {
        return Err(format!("known namespace URI does not start with h: {known_uri}").into());
    }
    bytes[0] = b'X';
    Ok(String::from_utf8(bytes)?)
}

fn unicode_vendor_uri(target_bytes: usize) -> Result<String, Box<dyn Error>> {
    const PREFIX: &str = "urn:vendor:é";
    if target_bytes <= PREFIX.len() {
        return Err(
            format!("known URI is too short for Unicode vendor prefix: {target_bytes}").into(),
        );
    }
    let uri = format!("{PREFIX}{}", "u".repeat(target_bytes - PREFIX.len()));
    if uri.len() != target_bytes {
        return Err(format!(
            "Unicode vendor URI has {} bytes, expected {target_bytes}",
            uri.len()
        )
        .into());
    }
    Ok(uri)
}

fn fixture_metadata(
    shape: Shape,
    slides: usize,
    replaced_text_tags: usize,
    entries: &[InjectionEntry],
) -> FixtureMetadata {
    FixtureMetadata {
        injection: shape.injection_name(),
        slide_parts: if entries.is_empty() { 0 } else { slides },
        replaced_text_tags,
        namespace_declarations: entries.len(),
        namespaced_attributes: entries.len(),
        namespace_uris: entries.iter().map(|entry| entry.uri.clone()).collect(),
        attribute_names: entries
            .iter()
            .map(|entry| entry.attribute_name.clone())
            .collect(),
    }
}

fn verify_unknown_namespace_payload(
    package: &mut Package,
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
) -> Result<usize, Box<dyn Error>> {
    let entries = injection_entries(shape)?;
    let checks = package.edit_opc(|opc| {
        let part_names = opc
            .iter_parts()
            .filter(|part| {
                part.content_type() == SLIDE_CONTENT_TYPE
                    && part.partname().as_str().starts_with("/ppt/slides/slide")
            })
            .map(|part| part.partname().clone())
            .collect::<Vec<_>>();
        let mut checks = Vec::with_capacity(part_names.len());
        for part_name in part_names {
            let part = opc.get_part_mut(&part_name)?;
            let blob = part.blob();
            let text_tags = count_text_tags(blob);
            let declarations = entries
                .iter()
                .map(|entry| count_occurrences(blob, entry.declaration.as_bytes()))
                .collect::<Vec<_>>();
            let attributes = entries
                .iter()
                .map(|entry| count_occurrences(blob, entry.attribute.as_bytes()))
                .collect::<Vec<_>>();
            checks.push((text_tags, declarations, attributes));
        }
        Ok(checks)
    })?;

    if checks.len() != slides {
        return Err(format!(
            "unknown-URI fixture exposed {} slide parts, expected {slides}",
            checks.len()
        )
        .into());
    }
    let mut total_text_tags = 0usize;
    for (slide_index, (text_tags, declarations, attributes)) in checks.iter().enumerate() {
        if *text_tags != shapes_per_slide {
            return Err(format!(
                "slide {slide_index} has {text_tags} a:t tags, expected {shapes_per_slide}"
            )
            .into());
        }
        if declarations.iter().any(|count| *count != 1)
            || attributes.iter().any(|count| *count != *text_tags)
        {
            return Err(format!(
                "slide {slide_index} did not retain every namespace declaration and attribute"
            )
            .into());
        }
        total_text_tags += *text_tags;
    }
    Ok(total_text_tags)
}

fn count_occurrences(haystack: &[u8], needle: &[u8]) -> usize {
    if needle.is_empty() || haystack.len() < needle.len() {
        return 0;
    }
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

fn expected_text(slides: usize, shapes_per_slide: usize, marker_operation: bool) -> String {
    let mut values = Vec::with_capacity(slides.saturating_mul(shapes_per_slide));
    for slide in 0..slides {
        for shape in 0..shapes_per_slide {
            values.push(if marker_operation && slide == 0 && shape == 0 {
                MARKER.to_owned()
            } else {
                source_text(slide, shape)
            });
        }
    }
    values.join("\n")
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

fn observe_output(
    expected: &mut Option<Identity>,
    output: Option<&[u8]>,
) -> Result<(), Box<dyn Error>> {
    let Some(output) = output else {
        return Ok(());
    };
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

// ---------------------------------------------------------------------------
// Deterministic corpus copied from change 0760.
// ---------------------------------------------------------------------------

fn source_text(slide: usize, shape: usize) -> String {
    format!("litchi-perf-baseline-pptx-semantic-v1-source-{slide:03}-{shape:03}")
}

fn build(shape: Shape) -> Result<Fixture, Box<dyn Error>> {
    let (slides, boxes) = shape.dimensions();
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for slide_index in 0..slides {
            let slide = presentation.add_slide()?;
            for shape_index in 0..boxes {
                slide.add_text_box(
                    &source_text(slide_index, shape_index),
                    36 + i64::try_from(shape_index % 4)? * 180,
                    36 + i64::try_from(shape_index / 4)? * 90,
                    144,
                    54,
                );
            }
        }
    }
    let regular_source = package.to_bytes()?;
    let entries = injection_entries(shape)?;
    let (source_bytes, replaced_text_tags) = if entries.is_empty() {
        (regular_source, 0)
    } else {
        let mut reopened = Package::from_bytes(&regular_source)?;
        let replaced_text_tags = reopened.edit_opc(|opc| {
            let part_names = opc
                .iter_parts()
                .filter(|part| {
                    part.content_type() == SLIDE_CONTENT_TYPE
                        && part.partname().as_str().starts_with("/ppt/slides/slide")
                })
                .map(|part| part.partname().clone())
                .collect::<Vec<_>>();
            let mut replaced_text_tags = 0usize;
            for part_name in part_names {
                let part = opc.get_part_mut(&part_name)?;
                let (blob, count) = rewrite_slide_blob(part.blob(), &entries);
                part.set_blob(blob);
                replaced_text_tags += count;
            }
            Ok(replaced_text_tags)
        })?;
        (reopened.to_bytes()?, replaced_text_tags)
    };
    let expected_text_tags = slides.saturating_mul(boxes);
    if !entries.is_empty() && replaced_text_tags != expected_text_tags {
        return Err(format!(
            "namespace fixture replaced {replaced_text_tags} a:t tags, expected {expected_text_tags}"
        )
        .into());
    }
    let metadata = fixture_metadata(shape, slides, replaced_text_tags, &entries);
    let _ = verify_output(&source_bytes, shape, slides, boxes, false)?;
    Ok(Fixture {
        bytes: source_bytes,
        metadata,
    })
}

fn rewrite_slide_blob(blob: &[u8], entries: &[InjectionEntry]) -> (Vec<u8>, usize) {
    let declarations = entries
        .iter()
        .map(|entry| entry.declaration.as_str())
        .collect::<Vec<_>>();
    let with_declarations = insert_attributes(blob, b"<p:sld", &declarations);
    rewrite_text_tags(&with_declarations, entries)
}

fn insert_attributes(blob: &[u8], element: &[u8], attributes: &[&str]) -> Vec<u8> {
    let mut search_cursor = 0usize;
    while let Some(relative_start) = find_bytes(&blob[search_cursor..], element) {
        let start = search_cursor + relative_start;
        let after_name = start + element.len();
        if !is_text_tag_start(blob, after_name) {
            search_cursor = after_name;
            continue;
        }
        let Some(relative_end) = blob[after_name..].iter().position(|byte| *byte == b'>') else {
            return blob.to_vec();
        };
        let end = after_name + relative_end;
        let insert_at = if end > start && blob[end - 1] == b'/' {
            end - 1
        } else {
            end
        };
        let added_bytes: usize = attributes.iter().map(|attribute| attribute.len() + 1).sum();
        let mut rewritten = Vec::with_capacity(blob.len() + added_bytes);
        rewritten.extend_from_slice(&blob[..insert_at]);
        for attribute in attributes {
            rewritten.push(b' ');
            rewritten.extend_from_slice(attribute.as_bytes());
        }
        rewritten.extend_from_slice(&blob[insert_at..]);
        return rewritten;
    }
    blob.to_vec()
}

fn rewrite_text_tags(blob: &[u8], entries: &[InjectionEntry]) -> (Vec<u8>, usize) {
    let added_bytes = entries
        .iter()
        .map(|entry| entry.attribute.len() + 1)
        .sum::<usize>();
    let mut rewritten = Vec::with_capacity(blob.len() + added_bytes);
    let mut search_cursor = 0usize;
    let mut emitted_cursor = 0usize;
    let mut replaced = 0usize;
    while let Some(relative_start) = find_bytes(&blob[search_cursor..], b"<a:t") {
        let start = search_cursor + relative_start;
        let after_name = start + b"<a:t".len();
        if !is_text_tag_start(blob, after_name) {
            search_cursor = after_name;
            continue;
        }
        let Some(relative_end) = blob[after_name..].iter().position(|byte| *byte == b'>') else {
            break;
        };
        let end = after_name + relative_end;
        let insert_at = if end > start && blob[end - 1] == b'/' {
            end - 1
        } else {
            end
        };
        rewritten.extend_from_slice(&blob[emitted_cursor..insert_at]);
        for entry in entries {
            rewritten.push(b' ');
            rewritten.extend_from_slice(entry.attribute.as_bytes());
        }
        rewritten.extend_from_slice(&blob[insert_at..=end]);
        emitted_cursor = end + 1;
        search_cursor = emitted_cursor;
        replaced += 1;
    }
    rewritten.extend_from_slice(&blob[emitted_cursor..]);
    (rewritten, replaced)
}

fn count_text_tags(blob: &[u8]) -> usize {
    let mut cursor = 0usize;
    let mut count = 0usize;
    while let Some(relative_start) = find_bytes(&blob[cursor..], b"<a:t") {
        let start = cursor + relative_start;
        let after_name = start + b"<a:t".len();
        if is_text_tag_start(blob, after_name) {
            count += 1;
        }
        cursor = after_name;
    }
    count
}

fn is_text_tag_start(blob: &[u8], offset: usize) -> bool {
    blob.get(offset)
        .is_some_and(|byte| *byte == b'>' || *byte == b'/' || byte.is_ascii_whitespace())
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || haystack.len() < needle.len() {
        return None;
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn namespace_injection_keeps_non_text_tags_and_places_declarations_on_root() {
        let entries = injection_entries(Shape::Vendor).expect("vendor entries");
        let source = br#"<p:sld xmlns:p="urn:p"><p:cSld><p:spTree><a:txBody><a:p><a:r><a:t>one</a:t></a:r></a:p></a:txBody><a:t>two</a:t></p:spTree></p:cSld></p:sld>"#;
        let (rewritten, count) = rewrite_slide_blob(source, &entries);
        assert_eq!(count, 2);
        assert_eq!(count_text_tags(&rewritten), 2);
        assert!(rewritten.starts_with(b"<p:sld xmlns:p=\"urn:p\" xmlns:vP=\"X"));
        assert!(
            rewritten
                .windows(b"<a:txBody>".len())
                .any(|window| window == b"<a:txBody>")
        );
        for entry in entries {
            assert_eq!(
                count_occurrences(&rewritten, entry.declaration.as_bytes()),
                1
            );
            assert_eq!(count_occurrences(&rewritten, entry.attribute.as_bytes()), 2);
        }
    }

    #[test]
    fn unicode_uris_are_valid_utf8_and_preserve_known_uri_byte_lengths() {
        let entries = injection_entries(Shape::UnicodeVendor).expect("Unicode entries");
        assert_eq!(entries.len(), KNOWN_URI_NOTES.len());
        for ((_, known), entry) in KNOWN_URI_NOTES.iter().zip(entries) {
            assert_eq!(entry.uri.len(), known.len());
            assert!(entry.uri.starts_with("urn:vendor:é"));
            assert!(std::str::from_utf8(entry.uri.as_bytes()).is_ok());
        }
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
    "usage: namespace-uri-probe --mode capture|commit|lifecycle --shape tiny|medium|large|vendor|unicode-vendor --samples N --warmup N --output PATH"
}
