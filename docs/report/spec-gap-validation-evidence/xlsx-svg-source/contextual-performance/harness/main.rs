//! Process-isolated allocation evidence for the borrowed XLSX drawing source
//! scanner.  This harness deliberately stops at the source projection and
//! shared SVG codec boundary; it does not open or mutate an XLSX package.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the evidence binary owns bounded fixtures and a process-local allocator observer"
)]

mod support;

use std::env;
use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::hint::black_box;
use std::time::Instant;

use litchi_drawingml::svg_blip::{self, NamespaceContext, Reference};
use litchi_xlsx::drawing::{SourceDrawing, SvgOwnerState};

use support::{AllocDelta, AllocSnapshot};

type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SMALL_OUTPUT_CAP: usize = 128;

const LANES: &[&str] = &[
    "contextual_read_p16_n0",
    "contextual_read_p16_n32",
    "contextual_read_p16_n128",
    "contextual_read_p32_n0",
    "contextual_read_p32_n32",
    "contextual_read_p32_n128",
    "contextual_read_p128_n0",
    "contextual_read_p128_n32",
    "contextual_read_p128_n128",
    "standalone_export_p16_n0",
    "standalone_export_p16_n32",
    "standalone_export_p16_n128",
    "standalone_export_p32_n0",
    "standalone_export_p32_n32",
    "standalone_export_p32_n128",
    "standalone_export_p128_n0",
    "standalone_export_p128_n32",
    "standalone_export_p128_n128",
    "scalar_reference_edit_p16_n0",
    "scalar_reference_edit_p16_n32",
    "scalar_reference_edit_p16_n128",
    "scalar_reference_edit_p32_n0",
    "scalar_reference_edit_p32_n32",
    "scalar_reference_edit_p32_n128",
    "scalar_reference_edit_p128_n0",
    "scalar_reference_edit_p128_n32",
    "scalar_reference_edit_p128_n128",
    "small_cap_refusal_p16_n0",
    "small_cap_refusal_p16_n32",
    "small_cap_refusal_p16_n128",
    "small_cap_refusal_p32_n0",
    "small_cap_refusal_p32_n32",
    "small_cap_refusal_p32_n128",
    "small_cap_refusal_p128_n0",
    "small_cap_refusal_p128_n32",
    "small_cap_refusal_p128_n128",
    "contextual_read_original_149433",
    "standalone_export_original_149433",
    "small_cap_refusal_original_149433",
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Scenario {
    Read,
    Export,
    Edit,
    Refusal,
}

#[derive(Clone, Copy, Debug)]
struct LaneSpec {
    scenario: Scenario,
    pictures: usize,
    namespaces: usize,
    original: bool,
}

impl LaneSpec {
    fn expected_success(self) -> bool {
        self.scenario != Scenario::Refusal
    }
}

#[derive(Debug)]
struct Fixture {
    source: Vec<u8>,
    pictures: usize,
    namespaces: usize,
    original: bool,
}

#[derive(Clone, Copy, Debug, Default)]
struct Observation {
    retained_raw_source_bytes: u64,
    source_none_count: usize,
    context_present_count: usize,
    context_distinct_count: usize,
    shared_context_identity: bool,
    context_binding_count_max: usize,
}

#[derive(Clone, Copy, Debug, Default)]
struct Execution {
    actual_success: bool,
    semantic_ok: bool,
    output_exact: bool,
    retained_raw_source_bytes: u64,
    source_none_count: usize,
    context_present_count: usize,
    context_distinct_count: usize,
    shared_context_identity: bool,
    context_binding_count_max: usize,
    export_bytes: u64,
    readback_count: usize,
    qname_preserved: bool,
    exact_embedded_reference: bool,
    linked_reference_count: usize,
    refusal_count: usize,
}

#[derive(Debug)]
struct Failure {
    message: String,
}

impl std::fmt::Display for Failure {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl StdError for Failure {}

fn main() -> std::process::ExitCode {
    let arguments = match parse_args() {
        Ok(value) => value,
        Err(error) if error.to_string() == "help" => return std::process::ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            return std::process::ExitCode::from(2);
        },
    };
    match run(&arguments.0, arguments.1, arguments.2) {
        Ok(receipt) => {
            println!("{receipt}");
            std::process::ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("contextual profile failed: {error}");
            std::process::ExitCode::from(1)
        },
    }
}

fn parse_args() -> Result<(String, usize, usize)> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: xlsx-svg-contextual-profile --lane NAME [--warmup N] [--samples N]\\n\\n{} source-only lanes; no XLSX package or lifecycle operation is included",
                    LANES.len()
                );
                return Err("help".into());
            },
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a value")?
                    .parse()
                    .map_err(|_| "--warmup must be an integer")?;
            },
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a value")?
                    .parse()
                    .map_err(|_| "--samples must be an integer")?;
            },
            unknown => return Err(format!("unknown argument: {unknown}").into()),
        }
    }
    let lane = lane.ok_or("--lane is required")?;
    if !LANES.contains(&lane.as_str()) {
        return Err(format!("unknown lane: {lane}").into());
    }
    if samples == 0 {
        return Err("--samples must be nonzero".into());
    }
    Ok((lane, warmup, samples))
}

fn run(lane: &str, warmup: usize, samples: usize) -> Result<String> {
    let spec = lane_spec(lane)?;
    let fixture = fixture(spec)?;
    for _ in 0..warmup {
        execute(spec, &fixture).map_err(|error| Failure {
            message: format!("warmup failed: {error}"),
        })?;
    }
    let mut sample_receipts = Vec::with_capacity(samples);
    for _ in 0..samples {
        support::reset_counters();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let result = execute(spec, &fixture);
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let allocation = before.delta(after);
        sample_receipts.push(sample_json(spec, elapsed_ns, allocation, result));
    }
    Ok(receipt_json(
        lane,
        spec,
        &fixture,
        warmup,
        samples,
        &sample_receipts,
    ))
}

fn lane_spec(lane: &str) -> Result<LaneSpec> {
    for prefix in [
        ("contextual_read_", Scenario::Read),
        ("standalone_export_", Scenario::Export),
        ("scalar_reference_edit_", Scenario::Edit),
        ("small_cap_refusal_", Scenario::Refusal),
    ] {
        if let Some(rest) = lane.strip_prefix(prefix.0) {
            if rest == "original_149433" {
                return Ok(LaneSpec {
                    scenario: prefix.1,
                    pictures: 32,
                    namespaces: 128,
                    original: true,
                });
            }
            let (pictures, namespaces) = rest
                .strip_prefix('p')
                .ok_or("lane picture count is missing")?
                .split_once("_n")
                .ok_or("lane namespace count is missing")?;
            let pictures = pictures.parse::<usize>()?;
            let namespaces = namespaces.parse::<usize>()?;
            if !matches!(pictures, 16 | 32 | 128) || !matches!(namespaces, 0 | 32 | 128) {
                return Err("lane count is outside the bounded corpus".into());
            }
            return Ok(LaneSpec {
                scenario: prefix.1,
                pictures,
                namespaces,
                original: false,
            });
        }
    }
    Err(format!("unknown lane: {lane}").into())
}

fn fixture(spec: LaneSpec) -> Result<Fixture> {
    if spec.original {
        let source = original_source();
        if source.len() != 149_433 {
            return Err(format!(
                "original corpus changed: expected 149433 bytes, got {}",
                source.len()
            )
            .into());
        }
        return Ok(Fixture {
            source,
            pictures: 32,
            namespaces: 128,
            original: true,
        });
    }
    Ok(Fixture {
        source: drawing_xml(spec.pictures, spec.namespaces),
        pictures: spec.pictures,
        namespaces: spec.namespaces,
        original: false,
    })
}

fn drawing_xml(pictures: usize, namespaces: usize) -> Vec<u8> {
    let mut root = format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}" xmlns:future="urn:litchi:future" xmlns:q="urn:litchi:qname""#
    );
    let payload = "x".repeat(1_000);
    for index in 0..namespaces {
        write!(
            root,
            " xmlns:n{index:03}=\"urn:litchi:context:{index:03}:{payload}\""
        )
        .expect("String cannot fail");
    }
    root.push('>');
    for index in 0..pictures {
        root.push_str(&picture_xml(index));
    }
    root.push_str("</xdr:wsDr>");
    root.into_bytes()
}

fn picture_xml(index: usize) -> String {
    format!(
        r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg{index}" future:qname="q:Opaque"><future:opaque>q:QName</future:opaque></asvg:svgBlip></a:ext></a:extLst></a:blip></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#,
        id = index.saturating_add(1),
    )
}

fn original_source() -> Vec<u8> {
    let mut xml = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main""#,
    );
    for index in 0..128 {
        write!(xml, " xmlns:p{index}=\"urn:opaque:{}\"", "x".repeat(1_000))
            .expect("String cannot fail");
    }
    xml.push('>');
    for index in 0usize..32 {
        write!(
            xml,
            r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg{index}"/></a:ext></a:extLst></a:blip></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#,
            index.saturating_add(1)
        )
        .expect("String cannot fail");
    }
    xml.push_str("</xdr:wsDr>");
    xml.into_bytes()
}

fn execute(spec: LaneSpec, fixture: &Fixture) -> Result<Execution> {
    let drawing = SourceDrawing::scan(&fixture.source)?;
    let observation = observe(&drawing, fixture)?;
    let common = Execution {
        actual_success: spec.expected_success(),
        semantic_ok: observation.context_distinct_count <= fixture.pictures,
        output_exact: true,
        retained_raw_source_bytes: observation.retained_raw_source_bytes,
        source_none_count: observation.source_none_count,
        context_present_count: observation.context_present_count,
        context_distinct_count: observation.context_distinct_count,
        shared_context_identity: observation.shared_context_identity,
        context_binding_count_max: observation.context_binding_count_max,
        ..Execution::default()
    };
    if drawing.pictures().len() != fixture.pictures {
        return Ok(Execution {
            semantic_ok: false,
            ..common
        });
    }
    match spec.scenario {
        Scenario::Read => Ok(common),
        Scenario::Export => execute_export(common, &drawing, fixture),
        Scenario::Edit => execute_edit(common, &drawing, fixture),
        Scenario::Refusal => execute_refusal(common, &drawing, fixture),
    }
}

fn observe(drawing: &SourceDrawing<'_>, fixture: &Fixture) -> Result<Observation> {
    let mut observation = Observation::default();
    // `namespace_context()` returns a borrowed view, so wrapper-pointer
    // identity is not a useful sharing check.  Count distinct immutable
    // backing nodes through the public `shares_storage` predicate instead.
    let mut contexts: Vec<&NamespaceContext> = Vec::new();
    for picture in drawing.pictures() {
        let owner = match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner,
            _ => {
                return Err(Failure {
                    message: "fixture SVG owner was not embedded".into(),
                }
                .into());
            },
        };
        let value = owner.value();
        let raw = value.raw_source().ok_or_else(|| Failure {
            message: "SVG value has no retained raw source".into(),
        })?;
        observation.retained_raw_source_bytes = observation
            .retained_raw_source_bytes
            .saturating_add(u64::try_from(raw.len()).unwrap_or(u64::MAX));
        if value.source().is_none() {
            observation.source_none_count = observation.source_none_count.saturating_add(1);
        }
        if let Some(context) = value.namespace_context() {
            observation.context_present_count = observation.context_present_count.saturating_add(1);
            if !contexts.iter().any(|known| known.shares_storage(context)) {
                contexts.push(context);
            }
            observation.context_binding_count_max = observation
                .context_binding_count_max
                .max(context.binding_count());
        }
    }
    observation.context_distinct_count = contexts.len();
    observation.shared_context_identity = observation.context_present_count == fixture.pictures
        && !contexts.is_empty()
        && contexts.len() == 1;
    Ok(observation)
}

fn execute_export(
    mut execution: Execution,
    drawing: &SourceDrawing<'_>,
    fixture: &Fixture,
) -> Result<Execution> {
    let mut export_bytes = 0_u64;
    let mut readback_count = 0_usize;
    let mut qname_preserved = true;
    let mut exact_embedded_reference = true;
    let mut linked_reference_count = 0_usize;
    for (index, picture) in drawing.pictures().iter().enumerate() {
        let owner = embedded_owner(picture)?;
        let value = owner.value();
        let output = export_owner(owner, value, &fixture.source, usize::MAX)?;
        export_bytes = export_bytes.saturating_add(u64::try_from(output.len()).unwrap_or(u64::MAX));
        let readback = svg_blip::read(&output)?;
        if readback
            .embedded()
            .is_some_and(|id| matches_svg_reference(id.as_str(), index))
        {
            readback_count = readback_count.saturating_add(1);
        }
        if readback.linked().is_some() {
            linked_reference_count = linked_reference_count.saturating_add(1);
        }
        exact_embedded_reference &= readback
            .embedded()
            .is_some_and(|id| matches_svg_reference(id.as_str(), index))
            && readback.linked().is_none();
        qname_preserved &= fixture.original
            || (output
                .windows(b"q:Opaque".len())
                .any(|window| window == b"q:Opaque")
                && output
                    .windows(b"q:QName".len())
                    .any(|window| window == b"q:QName"));
        drop(readback);
        drop(output);
    }
    execution.export_bytes = export_bytes;
    execution.readback_count = readback_count;
    execution.qname_preserved = qname_preserved;
    execution.exact_embedded_reference = exact_embedded_reference;
    execution.linked_reference_count = linked_reference_count;
    execution.semantic_ok &= readback_count == fixture.pictures
        && qname_preserved
        && exact_embedded_reference
        && linked_reference_count == 0;
    execution.output_exact = execution.semantic_ok;
    Ok(execution)
}

fn matches_svg_reference(value: &str, index: usize) -> bool {
    value
        .strip_prefix("rIdSvg")
        .and_then(|suffix| suffix.parse::<usize>().ok())
        == Some(index)
}

fn execute_edit(
    mut execution: Execution,
    drawing: &SourceDrawing<'_>,
    fixture: &Fixture,
) -> Result<Execution> {
    let mut export_bytes = 0_u64;
    let mut readback_count = 0_usize;
    let mut qname_preserved = true;
    let mut exact_embedded_reference = true;
    let mut linked_reference_count = 0_usize;
    for picture in drawing.pictures() {
        let owner = embedded_owner(picture)?;
        let mut value = owner.value().clone();
        value.set_reference(Reference::embedded("rIdEdited")?)?;
        let output = if value.namespace_context().is_some() {
            svg_blip::write_contextual(&value, usize::MAX)?
        } else {
            svg_blip::write(&value)?
        };
        export_bytes = export_bytes.saturating_add(u64::try_from(output.len()).unwrap_or(u64::MAX));
        let readback = svg_blip::read(&output)?;
        if readback
            .embedded()
            .is_some_and(|id| id.as_str() == "rIdEdited")
        {
            readback_count = readback_count.saturating_add(1);
        }
        if readback.linked().is_some() {
            linked_reference_count = linked_reference_count.saturating_add(1);
        }
        exact_embedded_reference &= readback
            .embedded()
            .is_some_and(|id| id.as_str() == "rIdEdited")
            && readback.linked().is_none();
        qname_preserved &= fixture.original
            || (output
                .windows(b"q:Opaque".len())
                .any(|window| window == b"q:Opaque")
                && output
                    .windows(b"q:QName".len())
                    .any(|window| window == b"q:QName"));
        drop(readback);
        drop(output);
        let _ = black_box(value);
    }
    execution.actual_success = true;
    execution.export_bytes = export_bytes;
    execution.readback_count = readback_count;
    execution.qname_preserved = qname_preserved;
    execution.exact_embedded_reference = exact_embedded_reference;
    execution.linked_reference_count = linked_reference_count;
    execution.semantic_ok &= readback_count == fixture.pictures
        && qname_preserved
        && exact_embedded_reference
        && linked_reference_count == 0;
    execution.output_exact = execution.semantic_ok;
    Ok(execution)
}

fn execute_refusal(
    mut execution: Execution,
    drawing: &SourceDrawing<'_>,
    fixture: &Fixture,
) -> Result<Execution> {
    let mut refusal_count = 0_usize;
    for picture in drawing.pictures() {
        let owner = embedded_owner(picture)?;
        let value = owner.value();
        let result = export_owner(owner, value, &fixture.source, SMALL_OUTPUT_CAP);
        if result.is_err() {
            refusal_count = refusal_count.saturating_add(1);
        }
    }
    execution.actual_success = false;
    execution.refusal_count = refusal_count;
    execution.semantic_ok &= refusal_count == fixture.pictures;
    execution.output_exact = execution.semantic_ok;
    Ok(execution)
}

fn embedded_owner<'a>(
    picture: &'a litchi_xlsx::drawing::PictureSource<'a>,
) -> Result<&'a litchi_xlsx::drawing::SvgOwner<'a>> {
    match picture.svg_owner() {
        SvgOwnerState::Embedded(owner) => Ok(owner),
        _ => Err(Failure {
            message: "fixture SVG owner was not embedded".into(),
        }
        .into()),
    }
}

fn export_owner(
    owner: &litchi_xlsx::drawing::SvgOwner<'_>,
    value: &litchi_drawingml::svg_blip::SvgBlip,
    source: &[u8],
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    if value.namespace_context().is_some() {
        Ok(svg_blip::write_contextual(value, max_output_bytes)?)
    } else {
        Ok(owner.namespace_complete(source, max_output_bytes)?)
    }
}

fn sample_json(
    spec: LaneSpec,
    elapsed_ns: u64,
    allocation: AllocDelta,
    result: Result<Execution>,
) -> String {
    let (execution, error) = match result {
        Ok(execution) => (execution, None),
        Err(error) => (Execution::default(), Some(error.to_string())),
    };
    let error_json = error
        .as_deref()
        .map(|value| format!("\"{}\"", support::json_escape(value)))
        .unwrap_or_else(|| "null".into());
    format!(
        "{{\"elapsed_ns\":{elapsed_ns},\"alloc_calls\":{},\"realloc_calls\":{},\"dealloc_calls\":{},\"direct_allocated_bytes\":{},\"realloc_new_bytes\":{},\"realloc_old_bytes\":{},\"deallocated_bytes\":{},\"requested_alloc_bytes\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"alloc_failed\":{},\"alloc_invalid\":{},\"alloc_balance_ok\":{},\"expected_success\":{},\"actual_success\":{},\"semantic_ok\":{},\"output_exact\":{},\"retained_raw_source_bytes\":{},\"source_none_count\":{},\"context_present_count\":{},\"context_distinct_count\":{},\"shared_context_identity\":{},\"context_binding_count_max\":{},\"export_bytes\":{},\"readback_count\":{},\"qname_preserved\":{},\"exact_embedded_reference\":{},\"linked_reference_count\":{},\"refusal_count\":{},\"error\":{error_json}}}",
        allocation.calls,
        allocation.realloc_calls,
        allocation.dealloc_calls,
        allocation.direct,
        allocation.realloc_new,
        allocation.realloc_old,
        allocation.deallocated,
        allocation.requested(),
        allocation.live_before,
        allocation.live_after,
        allocation.peak_delta,
        allocation.failed,
        allocation.invalid,
        allocation.balanced(),
        spec.expected_success(),
        execution.actual_success,
        execution.semantic_ok,
        execution.output_exact,
        execution.retained_raw_source_bytes,
        execution.source_none_count,
        execution.context_present_count,
        execution.context_distinct_count,
        execution.shared_context_identity,
        execution.context_binding_count_max,
        execution.export_bytes,
        execution.readback_count,
        execution.qname_preserved,
        execution.exact_embedded_reference,
        execution.linked_reference_count,
        execution.refusal_count,
    )
}

fn receipt_json(
    lane: &str,
    spec: LaneSpec,
    fixture: &Fixture,
    warmup: usize,
    samples: usize,
    samples_json: &[String],
) -> String {
    let mut receipt = String::new();
    write!(
        receipt,
        "{{\"schema\":\"xlsx-svg-contextual-profile-v1\",\"lane\":\"{lane}\",\"input_bytes\":{},\"input_hash_fnv1a64\":{},\"input_hash_sha256\":\"{}\",\"picture_count\":{},\"namespace_declarations\":{},\"original_corpus\":{},\"warmup\":{warmup},\"sample_count\":{samples},\"expected_success\":{},\"samples\":[",
        fixture.source.len(),
        support::fnv1a64(&fixture.source),
        support::sha256_hex(&fixture.source),
        fixture.pictures,
        fixture.namespaces,
        fixture.original,
        spec.expected_success(),
    )
    .expect("String cannot fail");
    for (index, sample) in samples_json.iter().enumerate() {
        if index != 0 {
            receipt.push(',');
        }
        receipt.push_str(sample);
    }
    receipt.push_str("]}");
    receipt
}
