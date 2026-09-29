//! Small root-lock-aligned probe for the `flate2`/`zlib-rs` runtime-dispatch
//! question.
//!
//! The executable deliberately names only the public OPC and ZIP APIs.  It
//! has no direct `flate2` dependency, so a baseline and a candidate build can
//! differ only in the resolved feature set of the same production dependency
//! path.  Fixture construction, expected output construction, archive
//! readback, and all hashes used as oracles are outside the measured clocks
//! except for the decompression work in `read`.

#![forbid(unsafe_code)]

use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::hint::black_box;
use std::io::Write as IoWrite;
use std::path::{Path, PathBuf};
use std::time::Instant;

use litchi_opc::{
    BlobPart, OpcPackage, PackURI, PackageWriter, SourceBackedPackage, SourceTopologyPlan,
};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;

const SCHEMA: &str = "litchi.compression-probe-0842.v1";
const TOOL: &str = "compression-probe-0842";
const TINY_BYTES: usize = 64;
const LARGE_BYTES: usize = 1024 * 1024;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 10_000;
const TARGET_NAME: &str = "/custom/payload.bin";

const PARTS: [PartSpec; 4] = [
    PartSpec {
        name: "/word/document.xml",
        archive_name: "word/document.xml",
        content_type: "application/xml",
        role: "DOCX-like XML part",
    },
    PartSpec {
        name: "/xl/workbook.xml",
        archive_name: "xl/workbook.xml",
        content_type: "application/xml",
        role: "XLSX-like XML part",
    },
    PartSpec {
        name: "/ppt/presentation.xml",
        archive_name: "ppt/presentation.xml",
        content_type: "application/xml",
        role: "PPTX-like XML part",
    },
    PartSpec {
        name: TARGET_NAME,
        archive_name: "custom/payload.bin",
        content_type: "application/octet-stream",
        role: "opaque payload part",
    },
];

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Write,
    Read,
    Changed,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "write" => Ok(Self::Write),
            "read" => Ok(Self::Read),
            "changed" => Ok(Self::Changed),
            _ => {
                Err(format!("--mode must be one of write, read, or changed (got {value:?})").into())
            },
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Write => "litchi_opc::PackageWriter::to_bytes on a prepared OpcPackage",
            Self::Read => {
                "ArchiveReader::new plus indexed reads and Deflate decompression of every member"
            },
            Self::Changed => {
                "SourceBackedPackage::from_vec, one replacement plan, and source-backed publication"
            },
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Shape {
    Tiny,
    Xml,
    Random,
    Repeated,
}

impl Shape {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "tiny" => Ok(Self::Tiny),
            "xml" => Ok(Self::Xml),
            "random" => Ok(Self::Random),
            "repeated" => Ok(Self::Repeated),
            _ => Err(format!(
                "--shape must be one of tiny, xml, random, or repeated (got {value:?})"
            )
            .into()),
        }
    }

    const fn payload_bytes(self) -> usize {
        match self {
            Self::Tiny => TINY_BYTES,
            Self::Xml | Self::Random | Self::Repeated => LARGE_BYTES,
        }
    }

    const fn description(self) -> &'static str {
        match self {
            Self::Tiny => "four 64-byte members",
            Self::Xml => "four 1 MiB members with varying valid XML or textual payloads",
            Self::Random => "three 1 MiB valid XML members and one xorshift payload",
            Self::Repeated => {
                "three 1 MiB valid XML members and one repeated deterministic byte pattern"
            },
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PartSpec {
    name: &'static str,
    archive_name: &'static str,
    content_type: &'static str,
    role: &'static str,
}

#[derive(Debug)]
struct Config {
    mode: Mode,
    shape: Shape,
    samples: usize,
    warmup: usize,
    output: PathBuf,
    artifact: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct Identity {
    bytes: u64,
    sha256: String,
}

#[derive(Debug)]
struct Fixture {
    package: OpcPackage,
    source: Vec<u8>,
    source_payloads: Vec<Vec<u8>>,
    changed_payload: Vec<u8>,
    changed_payloads: Vec<Vec<u8>>,
    expected_changed: Vec<u8>,
    source_identity: Identity,
    expected_logical: Identity,
    expected_changed_logical: Identity,
    target_uri: PackURI,
}

#[derive(Debug)]
struct OperationResult {
    elapsed_ns: u64,
    archive: Option<Vec<u8>>,
    archive_identity: Identity,
    logical: Identity,
    archive_exact_match: bool,
    logical_exact_match: bool,
}

#[derive(Debug, Serialize)]
struct Sample {
    index: usize,
    elapsed_ns: u64,
    archive: Identity,
    logical: Identity,
    archive_exact_match: bool,
    logical_exact_match: bool,
    archive_role: &'static str,
    input_archive_sha256: String,
    expected_logical_sha256: String,
    uncompressed_bytes: u64,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    mode: Mode,
    shape: Shape,
    shape_description: &'static str,
    timing_scope: &'static str,
    part_count: usize,
    parts: Vec<PartSpec>,
    target_part: &'static str,
    member_payload_bytes: usize,
    warmup: usize,
    samples_requested: usize,
    source: Identity,
    expected_archive: Identity,
    expected_logical: Identity,
    expected_changed_archive: Identity,
    expected_changed_logical: Identity,
    samples: Vec<Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    artifact: Option<Identity>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let config = parse_args(std::env::args().skip(1))?;
    run(config)
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    validate_output_path(&config.output, "--output")?;
    if let Some(artifact) = &config.artifact {
        validate_output_path(artifact, "--artifact")?;
        if artifact == &config.output {
            return Err("--artifact must differ from --output".into());
        }
    }

    // The package, source archive, changed archive, and their independent
    // logical-byte identities are all prepared before the first warmup.
    let fixture = build_fixture(config.shape)?;
    let expected_archive = match config.mode {
        Mode::Write | Mode::Read => fixture.source_identity.clone(),
        Mode::Changed => identity(&fixture.expected_changed),
    };
    let expected_logical = match config.mode {
        Mode::Write | Mode::Read => fixture.expected_logical.clone(),
        Mode::Changed => fixture.expected_changed_logical.clone(),
    };

    for _ in 0..config.warmup {
        let result = run_one(config.mode, &fixture, &expected_archive, &expected_logical)?;
        require_oracle(&result)?;
        black_box(result);
    }

    let mut samples = Vec::with_capacity(config.samples);
    let mut final_archive = None;
    for index in 0..config.samples {
        let result = run_one(config.mode, &fixture, &expected_archive, &expected_logical)?;
        require_oracle(&result)?;
        let archive_identity = result.archive_identity.clone();
        let logical = result.logical.clone();
        let archive_exact_match = result.archive_exact_match;
        let logical_exact_match = result.logical_exact_match;
        if config.artifact.is_some() && index + 1 == config.samples {
            final_archive = result.archive;
        }
        samples.push(Sample {
            index,
            elapsed_ns: result.elapsed_ns,
            archive: archive_identity,
            logical,
            archive_exact_match,
            logical_exact_match,
            archive_role: match config.mode {
                Mode::Read => "input-source-archive",
                Mode::Write | Mode::Changed => "operation-output-archive",
            },
            input_archive_sha256: fixture.source_identity.sha256.clone(),
            expected_logical_sha256: expected_logical.sha256.clone(),
            uncompressed_bytes: expected_logical.bytes,
        });
    }

    let artifact_identity = if let Some(path) = &config.artifact {
        let bytes = match config.mode {
            Mode::Read => fixture.source.as_slice(),
            Mode::Write | Mode::Changed => final_archive
                .as_deref()
                .ok_or("an artifact requires at least one retained final operation output")?,
        };
        let artifact_identity = identity(bytes);
        write_new(path, bytes)?;
        Some(artifact_identity)
    } else {
        None
    };

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        mode: config.mode,
        shape: config.shape,
        shape_description: config.shape.description(),
        timing_scope: config.mode.timing_scope(),
        part_count: PARTS.len(),
        parts: PARTS.to_vec(),
        target_part: TARGET_NAME,
        member_payload_bytes: config.shape.payload_bytes(),
        warmup: config.warmup,
        samples_requested: config.samples,
        source: fixture.source_identity,
        expected_archive,
        expected_logical,
        expected_changed_archive: identity(&fixture.expected_changed),
        expected_changed_logical: fixture.expected_changed_logical,
        samples,
        artifact: artifact_identity,
    };
    let report_bytes = serde_json::to_vec_pretty(&report)?;
    write_new(&config.output, &report_bytes)
}

fn run_one(
    mode: Mode,
    fixture: &Fixture,
    expected_archive: &Identity,
    expected_logical: &Identity,
) -> Result<OperationResult, Box<dyn Error>> {
    match mode {
        Mode::Write => run_write(fixture, expected_archive, expected_logical),
        Mode::Read => run_read(fixture, expected_archive, expected_logical),
        Mode::Changed => run_changed(fixture, expected_archive, expected_logical),
    }
}

fn run_write(
    fixture: &Fixture,
    expected_archive: &Identity,
    expected_logical: &Identity,
) -> Result<OperationResult, Box<dyn Error>> {
    let started = Instant::now();
    let output = PackageWriter::to_bytes(&fixture.package)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    black_box(&output);

    let archive_identity = identity(&output);
    let archive_exact_match = archive_identity == *expected_archive && output == fixture.source;
    let logical = archive_logical_identity(&output)?;
    let logical = verify_archive_payloads(&output, fixture, false, logical)?;
    let logical_exact_match = logical == *expected_logical;
    Ok(OperationResult {
        elapsed_ns,
        archive: Some(output),
        archive_identity,
        logical,
        archive_exact_match,
        logical_exact_match,
    })
}

fn run_read(
    fixture: &Fixture,
    expected_archive: &Identity,
    expected_logical: &Identity,
) -> Result<OperationResult, Box<dyn Error>> {
    let started = Instant::now();
    let reader = ArchiveReader::new(&fixture.source)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut payloads = Vec::with_capacity(names.len());
    for name in &names {
        let payload = reader.read(name)?;
        black_box(payload.len());
        payloads.push(payload);
    }
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    drop(reader);

    let logical = verify_member_payloads(&names, &payloads, fixture, false)?;
    let archive_identity = fixture.source_identity.clone();
    let archive_exact_match = archive_identity == *expected_archive;
    let logical_exact_match = logical == *expected_logical;
    Ok(OperationResult {
        elapsed_ns,
        archive: None,
        archive_identity,
        logical,
        archive_exact_match,
        logical_exact_match,
    })
}

fn run_changed(
    fixture: &Fixture,
    expected_archive: &Identity,
    expected_logical: &Identity,
) -> Result<OperationResult, Box<dyn Error>> {
    // Cloning the already prepared source and replacement is fixture setup;
    // the clock begins at the source-backed open and includes the actual
    // replacement plan plus preservation publication.
    let source = fixture.source.clone();
    let replacement = fixture.changed_payload.clone();
    let target = fixture.target_uri.clone();
    let started = Instant::now();
    let package = SourceBackedPackage::from_vec(source)?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(target, replacement)?;
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    black_box(&output);

    let archive_identity = identity(&output);
    let archive_exact_match =
        archive_identity == *expected_archive && output == fixture.expected_changed;
    let logical = archive_logical_identity(&output)?;
    let logical = verify_archive_payloads(&output, fixture, true, logical)?;
    let logical_exact_match = logical == *expected_logical;
    Ok(OperationResult {
        elapsed_ns,
        archive: Some(output),
        archive_identity,
        logical,
        archive_exact_match,
        logical_exact_match,
    })
}

fn build_fixture(shape: Shape) -> Result<Fixture, Box<dyn Error>> {
    let mut package = OpcPackage::new();
    let mut source_payloads = Vec::with_capacity(PARTS.len());
    for (index, part) in PARTS.iter().enumerate() {
        let payload = payload_for(shape, index, false)?;
        source_payloads.push(payload.clone());
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(part.name)?,
            part.content_type.to_owned(),
            payload,
        )));
    }
    let source = PackageWriter::to_bytes(&package)?;
    let changed_payload = payload_for(shape, PARTS.len() - 1, true)?;
    let mut changed_payloads = source_payloads.clone();
    changed_payloads[PARTS.len() - 1] = changed_payload.clone();
    let expected_changed = publish_changed(&source, &changed_payload)?;
    let source_identity = identity(&source);
    let expected_logical = archive_logical_identity(&source)?;
    let expected_changed_logical = archive_logical_identity(&expected_changed)?;
    Ok(Fixture {
        package,
        source,
        source_payloads,
        changed_payload,
        changed_payloads,
        expected_changed,
        source_identity,
        expected_logical,
        expected_changed_logical,
        target_uri: PackURI::new(TARGET_NAME)?,
    })
}

fn publish_changed(source: &[u8], replacement: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    let package = SourceBackedPackage::from_vec(source.to_vec())?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(PackURI::new(TARGET_NAME)?, replacement.to_vec())?;
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    Ok(output)
}

fn payload_for(shape: Shape, part_index: usize, changed: bool) -> Result<Vec<u8>, Box<dyn Error>> {
    let size = shape.payload_bytes();
    if part_index < PARTS.len() - 1 {
        return xml_payload(size, shape, part_index, changed);
    }
    Ok(match shape {
        Shape::Random => random_payload(size, part_index, changed),
        Shape::Repeated | Shape::Xml | Shape::Tiny => repeated_payload(size, part_index, changed),
    })
}

fn xml_payload(
    size: usize,
    shape: Shape,
    part_index: usize,
    changed: bool,
) -> Result<Vec<u8>, Box<dyn Error>> {
    let shape_name = match shape {
        Shape::Tiny => "tiny",
        Shape::Xml => "xml",
        Shape::Random => "random-control",
        Shape::Repeated => "repeated-control",
    };
    let root = format!("<root shape=\"{shape_name}\" part=\"{part_index}\" changed=\"{changed}\">");
    let tail_open = b"<tail>";
    let tail_close = b"</tail>";
    let close = b"</root>";
    let fixed = root
        .len()
        .checked_add(tail_open.len())
        .and_then(|value| value.checked_add(tail_close.len()))
        .and_then(|value| value.checked_add(close.len()))
        .ok_or("XML fixture framing length overflow")?;
    if size < fixed {
        return Err(format!("XML fixture size {size} is too small for its framing").into());
    }

    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(root.as_bytes());
    let mut item = 0usize;
    loop {
        let mut item_bytes = String::new();
        write!(
            &mut item_bytes,
            "<item id=\"{item}\">value-{shape_name}-{part_index}-{item}</item>"
        )?;
        if output.len() + item_bytes.len() + fixed > size {
            break;
        }
        output.extend_from_slice(item_bytes.as_bytes());
        item = item
            .checked_add(1)
            .ok_or("XML fixture item counter overflow")?;
    }
    output.extend_from_slice(tail_open);
    let filler = size
        .checked_sub(output.len())
        .and_then(|value| value.checked_sub(tail_close.len()))
        .and_then(|value| value.checked_sub(close.len()))
        .ok_or("XML fixture filler length underflow")?;
    for offset in 0..filler {
        output.push(b'a' + u8::try_from((offset + part_index) % 26)?);
    }
    output.extend_from_slice(tail_close);
    output.extend_from_slice(close);
    debug_assert_eq!(output.len(), size);
    Ok(output)
}

fn repeated_payload(size: usize, part_index: usize, changed: bool) -> Vec<u8> {
    const PATTERN: &[u8] = b"litchi-0842-repeated-deflate-pattern-0123456789abcdef";
    let mut output = Vec::with_capacity(size);
    for offset in 0..size {
        let marker = if changed { 0x5a } else { 0xa5 };
        output.push(PATTERN[(offset + part_index * 7) % PATTERN.len()] ^ marker);
    }
    output
}

fn random_payload(size: usize, part_index: usize, changed: bool) -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32
        .wrapping_add((part_index as u32).wrapping_mul(0x6d2b_79f5))
        .wrapping_add(if changed { 0x1357_9bdf } else { 0 });
    let mut output = Vec::with_capacity(size);
    for _ in 0..size {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        output.push((state >> 24) as u8);
    }
    output
}

fn archive_logical_identity(bytes: &[u8]) -> Result<Identity, Box<dyn Error>> {
    let reader = ArchiveReader::new(bytes)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut payloads = Vec::with_capacity(names.len());
    for name in &names {
        payloads.push(reader.read(name)?);
    }
    logical_identity_from_members(&names, &payloads)
}

fn verify_archive_payloads(
    bytes: &[u8],
    fixture: &Fixture,
    changed: bool,
    logical: Identity,
) -> Result<Identity, Box<dyn Error>> {
    let reader = ArchiveReader::new(bytes)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut payloads = Vec::with_capacity(names.len());
    for name in &names {
        payloads.push(reader.read(name)?);
    }
    let checked = verify_member_payloads(&names, &payloads, fixture, changed)?;
    if checked != logical {
        return Err(
            "archive logical identity changed during independent payload verification".into(),
        );
    }
    Ok(checked)
}

fn verify_member_payloads(
    names: &[String],
    payloads: &[Vec<u8>],
    fixture: &Fixture,
    changed: bool,
) -> Result<Identity, Box<dyn Error>> {
    if names.len() != PARTS.len() + 1 || payloads.len() != names.len() {
        return Err("OPC archive member count differs from the deterministic fixture".into());
    }
    if names
        .iter()
        .filter(|name| name.as_str() == "[Content_Types].xml")
        .count()
        != 1
    {
        return Err("OPC archive did not retain exactly one content-types member".into());
    }
    let expected_payloads = if changed {
        &fixture.changed_payloads
    } else {
        &fixture.source_payloads
    };
    for (index, part) in PARTS.iter().enumerate() {
        let Some(name_index) = names.iter().position(|name| name == part.archive_name) else {
            return Err(format!("OPC archive omitted member {}", part.archive_name).into());
        };
        if payloads[name_index] != expected_payloads[index] {
            return Err(format!(
                "OPC member {} differs from its independently generated payload",
                part.archive_name
            )
            .into());
        }
    }
    logical_identity_from_members(names, payloads)
}

fn require_oracle(result: &OperationResult) -> Result<(), Box<dyn Error>> {
    if !result.archive_exact_match || !result.logical_exact_match {
        return Err(format!(
            "operation oracle failed: archive_exact_match={} logical_exact_match={}",
            result.archive_exact_match, result.logical_exact_match
        )
        .into());
    }
    Ok(())
}

fn logical_identity_from_members(
    names: &[String],
    payloads: &[Vec<u8>],
) -> Result<Identity, Box<dyn Error>> {
    if names.len() != payloads.len() {
        return Err("logical identity member/name count mismatch".into());
    }
    let mut digest = Sha256::new();
    let mut total = 0_u64;
    for (name, payload) in names.iter().zip(payloads) {
        digest.update(u64::try_from(name.len())?.to_le_bytes());
        digest.update(name.as_bytes());
        digest.update(u64::try_from(payload.len())?.to_le_bytes());
        digest.update(payload);
        total = total
            .checked_add(u64::try_from(payload.len())?)
            .ok_or("logical member byte count overflow")?;
    }
    Ok(Identity {
        bytes: total,
        sha256: hex_digest(digest.finalize()),
    })
}

fn identity(bytes: &[u8]) -> Identity {
    Identity {
        bytes: bytes.len() as u64,
        sha256: hex_digest(Sha256::digest(bytes)),
    }
}

fn hex_digest(bytes: impl AsRef<[u8]>) -> String {
    let mut output = String::with_capacity(bytes.as_ref().len() * 2);
    for byte in bytes.as_ref() {
        write!(&mut output, "{byte:02x}").expect("writing into String cannot fail");
    }
    output
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos()).map_err(|_| "elapsed nanoseconds overflow".into())
}

fn validate_output_path(path: &Path, flag: &str) -> Result<(), Box<dyn Error>> {
    if path.as_os_str().is_empty() {
        return Err(format!("{flag} must not be empty").into());
    }
    if path.exists() {
        return Err(format!("{flag} already exists: {}", path.display()).into());
    }
    let Some(parent) = path.parent() else {
        return Err(format!("{flag} has no parent directory").into());
    };
    if !parent.as_os_str().is_empty() && !parent.is_dir() {
        return Err(format!("{flag} parent is not a directory: {}", parent.display()).into());
    }
    Ok(())
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)?;
    file.write_all(bytes)?;
    file.flush()?;
    Ok(())
}

fn parse_args<I>(args: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = String>,
{
    let mut args = args.into_iter();
    let mut mode = None;
    let mut shape = None;
    let mut samples = 30;
    let mut warmup = 3;
    let mut output = None;
    let mut artifact = None;

    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--mode" => mode = Some(Mode::parse(&required_value("--mode", &mut args)?)?),
            "--shape" => shape = Some(Shape::parse(&required_value("--shape", &mut args)?)?),
            "--samples" => {
                samples = parse_count(
                    "--samples",
                    &required_value("--samples", &mut args)?,
                    MAX_SAMPLES,
                )?;
            },
            "--warmup" => {
                warmup = parse_count(
                    "--warmup",
                    &required_value("--warmup", &mut args)?,
                    MAX_WARMUP,
                )?;
            },
            "--output" => output = Some(PathBuf::from(required_value("--output", &mut args)?)),
            "--artifact" => {
                artifact = Some(PathBuf::from(required_value("--artifact", &mut args)?));
            },
            _ => return Err(format!("unknown argument {argument:?}").into()),
        }
    }

    Ok(Config {
        mode: mode.ok_or("--mode is required")?,
        shape: shape.ok_or("--shape is required")?,
        samples,
        warmup,
        output: output.ok_or("--output is required")?,
        artifact,
    })
}

fn required_value<I>(flag: &str, args: &mut I) -> Result<String, Box<dyn Error>>
where
    I: Iterator<Item = String>,
{
    args.next()
        .ok_or_else(|| format!("{flag} requires a value").into())
}

fn parse_count(flag: &str, value: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} must be an unsigned decimal count"))?;
    if parsed > maximum {
        return Err(format!("{flag} exceeds maximum {maximum}").into());
    }
    if flag == "--samples" && parsed == 0 {
        return Err("--samples must be at least 1".into());
    }
    Ok(parsed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shapes_have_the_declared_member_sizes() {
        for shape in [Shape::Tiny, Shape::Xml, Shape::Random, Shape::Repeated] {
            let fixture = build_fixture(shape).expect("deterministic OPC fixture");
            assert_eq!(
                shape.payload_bytes(),
                match shape {
                    Shape::Tiny => TINY_BYTES,
                    Shape::Xml | Shape::Random | Shape::Repeated => LARGE_BYTES,
                }
            );
            assert_eq!(fixture.changed_payload.len(), shape.payload_bytes());
            assert_eq!(fixture.source_identity, identity(&fixture.source));
        }
    }

    #[test]
    fn changed_publication_preserves_unchanged_member_logical_bytes() {
        let fixture = build_fixture(Shape::Repeated).expect("fixture");
        let source_reader = ArchiveReader::new(&fixture.source).expect("source archive");
        let changed_reader =
            ArchiveReader::new(&fixture.expected_changed).expect("changed archive");
        for part in PARTS.iter().take(3) {
            assert_eq!(
                source_reader
                    .read(part.archive_name)
                    .expect("source member"),
                changed_reader
                    .read(part.archive_name)
                    .expect("changed member")
            );
        }
        assert_ne!(
            source_reader
                .read("custom/payload.bin")
                .expect("source target"),
            changed_reader
                .read("custom/payload.bin")
                .expect("changed target")
        );
    }

    #[test]
    fn cli_rejects_missing_and_out_of_bound_counts() {
        assert!(parse_args(Vec::<String>::new()).is_err());
        assert!(parse_count("--samples", "0", MAX_SAMPLES).is_err());
        assert!(parse_count("--samples", "10001", MAX_SAMPLES).is_err());
        assert!(parse_count("--warmup", "10001", MAX_WARMUP).is_err());
    }
}
