//! Standalone many-small-part OPC preservation profile.
//!
//! The probe owns the deterministic fixture and all correctness oracles.  The
//! only operation called by its timed path is the public source-backed OPC
//! topology publisher.  Archive reads, hashes, member checks, and result drops
//! happen after each clock stops.

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
use soapberry_zip::office::{ArchiveReader, Metadata};

pub const SCHEMA: &str = "litchi.opc-preservation-profile-0843.v1";
const TOOL: &str = "opc-preservation-probe-0843";
const PAYLOAD_BYTES: usize = 256;
const MIN_SAMPLES: usize = 1;
const MAX_SAMPLES: usize = 1_000;
const MAX_WARMUP: usize = 100;
const CONTENT_TYPES_MEMBER: &str = "[Content_Types].xml";
const PACKAGE_RELATIONSHIPS_MEMBER: &str = "_rels/.rels";

/// The two supported clocks deliberately differ in what they include.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Mode {
    Publish,
    Lifecycle,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "publish" => Ok(Self::Publish),
            "lifecycle" => Ok(Self::Lifecycle),
            _ => Err(format!("--mode must be one of publish or lifecycle (got {value:?})").into()),
        }
    }

    const fn timing_description(self) -> &'static str {
        match self {
            Self::Publish => {
                "SourceBackedPackage::prepare_topology validation and write_topology_to_stream via profile_publish on a prepared source package, replacement plan, and empty output Vec"
            },
            Self::Lifecycle => {
                "SourceBackedPackage::from_vec, SourceTopologyPlan::try_replace_part, prepare_topology validation, and write_topology_to_stream via profile_publish; source/replacement/target clones are prepared before the clock"
            },
        }
    }
}

#[derive(Debug, Clone)]
struct Config {
    parts: usize,
    mode: Mode,
    samples: usize,
    warmup: usize,
    output: PathBuf,
    artifact: Option<PathBuf>,
}

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct Identity {
    pub bytes: u64,
    pub sha256: String,
}

#[derive(Debug, Clone, Serialize)]
struct GeneratorDescription {
    part_count: usize,
    member_count: usize,
    payload_bytes: usize,
    part_name_pattern: &'static str,
    payload_formula: &'static str,
    target_part_index: usize,
    target_part_name: String,
    replacement: &'static str,
    metadata_members: [&'static str; 2],
}

#[derive(Debug)]
struct Fixture {
    parts: usize,
    source: Vec<u8>,
    expected_output: Vec<u8>,
    source_payloads: Vec<Vec<u8>>,
    replacement: Vec<u8>,
    target_uri: PackURI,
    source_content_types: Vec<u8>,
    source_package_relationships: Vec<u8>,
    source_content_types_metadata: Metadata,
    source_package_relationships_metadata: Metadata,
    source_identity: Identity,
    expected_identity: Identity,
    source_logical: Identity,
    expected_logical: Identity,
}

#[derive(Debug)]
struct TimedOutput {
    elapsed_ns: u64,
    output: Vec<u8>,
}

#[derive(Debug)]
struct Observation {
    output: Identity,
    logical: Identity,
    member_count: usize,
    output_exact: bool,
    logical_exact: bool,
}

#[derive(Debug, Serialize)]
struct HashWitness {
    output_sha256: String,
    logical_sha256: String,
    output_exact_expected: bool,
    logical_exact_expected: bool,
}

#[derive(Debug, Serialize)]
struct Sample {
    index: usize,
    elapsed_ns: u64,
    output: Identity,
    logical: Identity,
    member_count: usize,
    hash_witness: HashWitness,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    mode: Mode,
    timing_description: &'static str,
    generator: GeneratorDescription,
    warmup: usize,
    samples_requested: usize,
    source: Identity,
    source_logical: Identity,
    expected_output: Identity,
    expected_output_logical: Identity,
    samples: Vec<Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    artifact: Option<Identity>,
    #[serde(skip_serializing_if = "Option::is_none")]
    artifact_role: Option<&'static str>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let config = parse_args(std::env::args().skip(1))?;
    run(config)
}

/// The sole timed owner call.  Keeping this function public gives the runner
/// an unambiguous symbol for allocation/profiling attribution.
#[inline(never)]
pub fn profile_publish(
    package: SourceBackedPackage,
    plan: SourceTopologyPlan,
    output: &mut Vec<u8>,
) -> litchi_opc::Result<()> {
    let result = package.write_topology_to_stream(output, plan);
    let result = black_box(result);
    black_box(output.as_slice());
    result
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    validate_output_path(&config.output, "--output")?;
    if let Some(artifact) = &config.artifact {
        validate_output_path(artifact, "--artifact")?;
        if artifact == &config.output {
            return Err("--artifact must differ from --output".into());
        }
    }

    // Fixture construction and the direct expected publication are outside
    // every measured operation, and never call profile_publish.
    let fixture = build_fixture(config.parts)?;

    for _ in 0..config.warmup {
        let timed = run_one(config.mode, &fixture)?;
        let observation = observe_output(&timed.output, &fixture)?;
        require_oracle(&observation)?;
        black_box(timed);
    }

    let mut samples = Vec::with_capacity(config.samples);
    let mut final_output = None;
    for index in 0..config.samples {
        let timed = run_one(config.mode, &fixture)?;
        let observation = observe_output(&timed.output, &fixture)?;
        require_oracle(&observation)?;
        if config.artifact.is_some() && index + 1 == config.samples {
            // --artifact is intentionally the source ZIP.  The source is the
            // stable input needed by the independent local-span verifier.
            final_output = Some(timed.output);
        }
        samples.push(Sample {
            index,
            elapsed_ns: timed.elapsed_ns,
            output: observation.output.clone(),
            logical: observation.logical.clone(),
            member_count: observation.member_count,
            hash_witness: HashWitness {
                output_sha256: observation.output.sha256,
                logical_sha256: observation.logical.sha256,
                output_exact_expected: observation.output_exact,
                logical_exact_expected: observation.logical_exact,
            },
        });
        if final_output.is_none() {
            // The output is dropped only after all post-clock checks complete.
            black_box(timed.output);
        }
    }

    let artifact_identity = if let Some(path) = &config.artifact {
        let source = fixture.source.as_slice();
        let identity = identity(source);
        write_new(path, source)?;
        Some(identity)
    } else {
        None
    };
    // Keep this binding alive through report construction so the requested
    // final output is retained exactly once when an artifact was requested.
    black_box(final_output);

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        mode: config.mode,
        timing_description: config.mode.timing_description(),
        generator: GeneratorDescription {
            part_count: fixture.parts,
            member_count: fixture.parts + 2,
            payload_bytes: PAYLOAD_BYTES,
            part_name_pattern: "/custom/p{index:05}.bin",
            payload_formula: "byte[i] = (part_index * 17 + i * 13) % 251",
            target_part_index: fixture.parts - 1,
            target_part_name: part_name(fixture.parts - 1),
            replacement: "target payload with byte[0] XOR 1",
            metadata_members: [CONTENT_TYPES_MEMBER, PACKAGE_RELATIONSHIPS_MEMBER],
        },
        warmup: config.warmup,
        samples_requested: config.samples,
        source: fixture.source_identity,
        source_logical: fixture.source_logical,
        expected_output: fixture.expected_identity,
        expected_output_logical: fixture.expected_logical,
        samples,
        artifact: artifact_identity,
        artifact_role: config.artifact.map(|_| "source-input-archive"),
    };
    let report_bytes = serde_json::to_vec_pretty(&report)?;
    write_new(&config.output, &report_bytes)
}

fn run_one(mode: Mode, fixture: &Fixture) -> Result<TimedOutput, Box<dyn Error>> {
    match mode {
        Mode::Publish => run_publish(fixture),
        Mode::Lifecycle => run_lifecycle(fixture),
    }
}

fn run_publish(fixture: &Fixture) -> Result<TimedOutput, Box<dyn Error>> {
    // All preparation, including the source-backed open, plan construction,
    // clones, and output allocation, is before the measured call.
    let source = fixture.source.clone();
    let replacement = fixture.replacement.clone();
    let target = fixture.target_uri.clone();
    let package = SourceBackedPackage::from_vec(source)?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(target, replacement)?;
    let mut output = Vec::new();
    let started = Instant::now();
    profile_publish(package, plan, &mut output)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    Ok(TimedOutput { elapsed_ns, output })
}

fn run_lifecycle(fixture: &Fixture) -> Result<TimedOutput, Box<dyn Error>> {
    // The source/replacement/target clones are deliberately outside the
    // lifecycle clock; opening, planning, and publication remain inside it.
    let source = fixture.source.clone();
    let replacement = fixture.replacement.clone();
    let target = fixture.target_uri.clone();
    let mut output = Vec::new();
    let started = Instant::now();
    let package = SourceBackedPackage::from_vec(source)?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(target, replacement)?;
    profile_publish(package, plan, &mut output)?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    Ok(TimedOutput { elapsed_ns, output })
}

fn build_fixture(parts: usize) -> Result<Fixture, Box<dyn Error>> {
    validate_parts(parts)?;
    let mut package = OpcPackage::new();
    let mut source_payloads = Vec::with_capacity(parts);
    for index in 0..parts {
        let payload = payload_for(index, false);
        source_payloads.push(payload.clone());
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(part_name(index))?,
            "application/octet-stream".to_owned(),
            payload,
        )));
    }
    let source = PackageWriter::to_bytes(&package)?;
    let target_uri = PackURI::new(part_name(parts - 1))?;
    let replacement = payload_for(parts - 1, true);

    // This is intentionally a direct source-backed publication.  The owner
    // function is reserved for measured warmups/samples.
    let expected_output = direct_expected_publish(&source, &target_uri, &replacement)?;
    let source_snapshot = snapshot_source(&source, parts, &source_payloads)?;
    let expected_snapshot = verify_archive(
        &expected_output,
        parts,
        &source_payloads,
        &replacement,
        &source_snapshot,
        true,
        Some(&expected_output),
    )?;

    Ok(Fixture {
        parts,
        source_identity: identity(&source),
        source_logical: source_snapshot.logical,
        expected_identity: identity(&expected_output),
        expected_logical: expected_snapshot.logical,
        source,
        expected_output,
        source_payloads,
        replacement,
        target_uri,
        source_content_types: source_snapshot.content_types,
        source_package_relationships: source_snapshot.package_relationships,
        source_content_types_metadata: source_snapshot.content_types_metadata,
        source_package_relationships_metadata: source_snapshot.package_relationships_metadata,
    })
}

fn direct_expected_publish(
    source: &[u8],
    target: &PackURI,
    replacement: &[u8],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let package = SourceBackedPackage::from_vec(source.to_vec())?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(target.clone(), replacement.to_vec())?;
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    Ok(output)
}

#[derive(Debug)]
struct SourceSnapshot {
    logical: Identity,
    content_types: Vec<u8>,
    package_relationships: Vec<u8>,
    content_types_metadata: Metadata,
    package_relationships_metadata: Metadata,
}

fn snapshot_source(
    source: &[u8],
    parts: usize,
    expected_payloads: &[Vec<u8>],
) -> Result<SourceSnapshot, Box<dyn Error>> {
    let reader = ArchiveReader::new(source)?;
    let names = member_names(&reader);
    verify_member_set(&names, parts)?;
    let content_types = reader.read(CONTENT_TYPES_MEMBER)?;
    let package_relationships = reader.read(PACKAGE_RELATIONSHIPS_MEMBER)?;
    let content_types_metadata = reader.metadata(CONTENT_TYPES_MEMBER)?;
    let package_relationships_metadata = reader.metadata(PACKAGE_RELATIONSHIPS_MEMBER)?;
    let logical = verify_payloads_and_logical(
        &reader,
        &names,
        parts,
        expected_payloads,
        expected_payloads,
        &content_types,
        &package_relationships,
        content_types_metadata,
        package_relationships_metadata,
    )?;
    Ok(SourceSnapshot {
        logical,
        content_types,
        package_relationships,
        content_types_metadata,
        package_relationships_metadata,
    })
}

fn observe_output(bytes: &[u8], fixture: &Fixture) -> Result<Observation, Box<dyn Error>> {
    let reader = ArchiveReader::new(bytes)?;
    let names = member_names(&reader);
    let logical = verify_payloads_and_logical(
        &reader,
        &names,
        fixture.parts,
        &fixture.source_payloads,
        &fixture.replacement,
        &fixture.source_content_types,
        &fixture.source_package_relationships,
        fixture.source_content_types_metadata,
        fixture.source_package_relationships_metadata,
    )?;
    let output = identity(bytes);
    Ok(Observation {
        output_exact: bytes == fixture.expected_output,
        logical_exact: logical == fixture.expected_logical,
        output,
        logical,
        member_count: names.len(),
    })
}

fn verify_archive(
    bytes: &[u8],
    parts: usize,
    source_payloads: &[Vec<u8>],
    replacement: &[u8],
    source: &SourceSnapshot,
    changed: bool,
    exact: Option<&[u8]>,
) -> Result<Observation, Box<dyn Error>> {
    let reader = ArchiveReader::new(bytes)?;
    let names = member_names(&reader);
    let logical = verify_payloads_and_logical(
        &reader,
        &names,
        parts,
        source_payloads,
        replacement,
        &source.content_types,
        &source.package_relationships,
        source.content_types_metadata,
        source.package_relationships_metadata,
    )?;
    let expected_logical = if changed {
        logical
    } else {
        source.logical.clone()
    };
    Ok(Observation {
        output: identity(bytes),
        logical: logical.clone(),
        member_count: names.len(),
        output_exact: exact.is_some_and(|expected| bytes == expected),
        logical_exact: logical == expected_logical,
    })
}

fn verify_payloads_and_logical(
    reader: &ArchiveReader<'_>,
    names: &[String],
    parts: usize,
    source_payloads: &[Vec<u8>],
    replacement: &[u8],
    source_content_types: &[u8],
    source_package_relationships: &[u8],
    source_content_types_metadata: Metadata,
    source_package_relationships_metadata: Metadata,
) -> Result<Identity, Box<dyn Error>> {
    verify_member_set(names, parts)?;
    if source_payloads.len() != parts {
        return Err("source payload count differs from requested part count".into());
    }
    let content_types = reader.read(CONTENT_TYPES_MEMBER)?;
    let package_relationships = reader.read(PACKAGE_RELATIONSHIPS_MEMBER)?;
    if content_types != source_content_types {
        return Err("[Content_Types].xml differs from the source metadata member".into());
    }
    if package_relationships != source_package_relationships {
        return Err("_rels/.rels differs from the source metadata member".into());
    }
    if reader.metadata(CONTENT_TYPES_MEMBER)? != source_content_types_metadata {
        return Err("[Content_Types].xml central metadata differs from source".into());
    }
    if reader.metadata(PACKAGE_RELATIONSHIPS_MEMBER)? != source_package_relationships_metadata {
        return Err("_rels/.rels central metadata differs from source".into());
    }
    if reader.metadata(CONTENT_TYPES_MEMBER)?.uncompressed_size()
        != u64::try_from(content_types.len())?
        || reader
            .metadata(PACKAGE_RELATIONSHIPS_MEMBER)?
            .uncompressed_size()
            != u64::try_from(package_relationships.len())?
    {
        return Err("metadata member declared size differs from CRC-checked payload".into());
    }

    let mut ordered_payloads = Vec::with_capacity(names.len());
    for name in names {
        let payload = reader.read(name)?;
        let metadata = reader.metadata(name)?;
        if metadata.is_directory() || metadata.uncompressed_size() != u64::try_from(payload.len())?
        {
            return Err(format!("member metadata is inconsistent for {name}").into());
        }
        ordered_payloads.push(payload);
    }

    for index in 0..parts {
        let expected = if index + 1 == parts {
            replacement
        } else {
            &source_payloads[index]
        };
        let payload = reader.read(&part_member_name(index))?;
        if payload != expected {
            return Err(format!("payload mismatch for part index {index}").into());
        }
    }

    Ok(logical_identity(names, &ordered_payloads)?)
}

fn member_names(reader: &ArchiveReader<'_>) -> Vec<String> {
    reader.file_names().map(str::to_owned).collect()
}

fn verify_member_set(names: &[String], parts: usize) -> Result<(), Box<dyn Error>> {
    if names.len() != parts + 2 {
        return Err(format!(
            "archive has {} members; expected {} (= parts + metadata)",
            names.len(),
            parts + 2
        )
        .into());
    }
    let mut actual = names.to_vec();
    actual.sort_unstable();
    let mut expected = expected_member_names(parts);
    expected.sort_unstable();
    if actual != expected {
        return Err("archive member set differs from the deterministic OPC set".into());
    }
    Ok(())
}

fn logical_identity(names: &[String], payloads: &[Vec<u8>]) -> Result<Identity, Box<dyn Error>> {
    if names.len() != payloads.len() {
        return Err("logical identity name/payload count mismatch".into());
    }
    let mut digest = Sha256::new();
    let mut bytes = 0_u64;
    for (name, payload) in names.iter().zip(payloads) {
        digest.update(u64::try_from(name.len())?.to_le_bytes());
        digest.update(name.as_bytes());
        digest.update(u64::try_from(payload.len())?.to_le_bytes());
        digest.update(payload);
        bytes = bytes
            .checked_add(u64::try_from(payload.len())?)
            .ok_or("logical payload byte count overflow")?;
    }
    Ok(Identity {
        bytes,
        sha256: hex_digest(digest.finalize()),
    })
}

fn payload_for(part_index: usize, changed: bool) -> Vec<u8> {
    let mut payload = Vec::with_capacity(PAYLOAD_BYTES);
    for index in 0..PAYLOAD_BYTES {
        let value = (part_index * 17 + index * 13) % 251;
        payload.push(value as u8);
    }
    if changed {
        payload[0] ^= 1;
    }
    payload
}

fn part_name(index: usize) -> String {
    format!("/custom/p{index:05}.bin")
}

fn part_member_name(index: usize) -> String {
    format!("custom/p{index:05}.bin")
}

fn expected_member_names(parts: usize) -> Vec<String> {
    let mut names = Vec::with_capacity(parts + 2);
    names.push(CONTENT_TYPES_MEMBER.to_owned());
    names.push(PACKAGE_RELATIONSHIPS_MEMBER.to_owned());
    names.extend((0..parts).map(part_member_name));
    names
}

fn require_oracle(observation: &Observation) -> Result<(), Box<dyn Error>> {
    if !observation.output_exact || !observation.logical_exact {
        return Err(format!(
            "publication oracle failed: output_exact={} logical_exact={} members={}",
            observation.output_exact, observation.logical_exact, observation.member_count
        )
        .into());
    }
    Ok(())
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
        write!(&mut output, "{byte:02x}").expect("String formatting cannot fail");
    }
    output
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos()).map_err(|_| "elapsed nanoseconds overflow".into())
}

fn validate_parts(parts: usize) -> Result<(), Box<dyn Error>> {
    match parts {
        4 | 64 | 1_024 | 4_096 => Ok(()),
        _ => Err("--parts must be one of 4, 64, 1024, or 4096".into()),
    }
}

fn parse_count(
    flag: &str,
    value: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} must be an integer (got {value:?})"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("{flag} must be in {minimum}..={maximum} (got {parsed})").into());
    }
    Ok(parsed)
}

fn required_value<I>(flag: &str, args: &mut I) -> Result<String, Box<dyn Error>>
where
    I: Iterator<Item = String>,
{
    args.next()
        .ok_or_else(|| format!("{flag} requires a value").into())
}

fn parse_args<I>(args: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = String>,
{
    let mut args = args.into_iter();
    let mut parts = None;
    let mut mode = None;
    let mut samples = None;
    let mut warmup = None;
    let mut output = None;
    let mut artifact = None;

    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--parts" => {
                let value = required_value("--parts", &mut args)?;
                let value = value
                    .parse::<usize>()
                    .map_err(|_| format!("--parts must be an integer (got {value:?})"))?;
                validate_parts(value)?;
                parts = Some(value);
            },
            "--mode" => mode = Some(Mode::parse(&required_value("--mode", &mut args)?)?),
            "--samples" => {
                samples = Some(parse_count(
                    "--samples",
                    &required_value("--samples", &mut args)?,
                    MIN_SAMPLES,
                    MAX_SAMPLES,
                )?);
            },
            "--warmup" => {
                warmup = Some(parse_count(
                    "--warmup",
                    &required_value("--warmup", &mut args)?,
                    0,
                    MAX_WARMUP,
                )?);
            },
            "--output" => output = Some(PathBuf::from(required_value("--output", &mut args)?)),
            "--artifact" => {
                artifact = Some(PathBuf::from(required_value("--artifact", &mut args)?));
            },
            _ => return Err(format!("unknown argument {argument:?}").into()),
        }
    }

    Ok(Config {
        parts: parts.ok_or("--parts is required")?,
        mode: mode.ok_or("--mode is required")?,
        samples: samples.ok_or("--samples is required")?,
        warmup: warmup.ok_or("--warmup is required")?,
        output: output.ok_or("--output is required")?,
        artifact,
    })
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn all_supported_part_dimensions_have_n_plus_two_members() {
        for parts in [4, 64, 1_024, 4_096] {
            let fixture = build_fixture(parts).expect("fixture");
            let source = snapshot_source(&fixture.source, parts, &fixture.source_payloads)
                .expect("source oracle");
            assert_eq!(source.logical, fixture.source_logical);
            let expected = verify_archive(
                &fixture.expected_output,
                parts,
                &fixture.source_payloads,
                &fixture.replacement,
                &source,
                true,
                Some(&fixture.expected_output),
            )
            .expect("changed oracle");
            assert_eq!(expected.member_count, parts + 2);
            assert_eq!(expected.logical, fixture.expected_logical);
        }
    }

    #[test]
    fn only_the_last_part_payload_changes() {
        let fixture = build_fixture(64).expect("fixture");
        let reader = ArchiveReader::new(&fixture.expected_output).expect("reader");
        for index in 0..fixture.parts {
            let actual = reader.read(&part_member_name(index)).expect("part");
            let expected = if index + 1 == fixture.parts {
                &fixture.replacement
            } else {
                &fixture.source_payloads[index]
            };
            assert_eq!(&actual, expected, "part {index}");
            if index + 1 != fixture.parts {
                assert_eq!(actual, payload_for(index, false));
            }
        }
        assert_ne!(
            reader
                .read(&part_member_name(fixture.parts - 1))
                .expect("target"),
            fixture.source_payloads[fixture.parts - 1]
        );
    }

    #[test]
    fn metadata_members_and_member_set_are_exact() {
        let fixture = build_fixture(4).expect("fixture");
        let reader = ArchiveReader::new(&fixture.expected_output).expect("reader");
        let names = member_names(&reader);
        verify_member_set(&names, 4).expect("member set");
        assert_eq!(
            reader.read(CONTENT_TYPES_MEMBER).expect("types"),
            fixture.source_content_types
        );
        assert_eq!(
            reader.read(PACKAGE_RELATIONSHIPS_MEMBER).expect("rels"),
            fixture.source_package_relationships
        );
    }

    #[test]
    fn both_modes_small_case_use_the_exact_oracle() {
        let fixture = build_fixture(4).expect("fixture");
        for mode in [Mode::Publish, Mode::Lifecycle] {
            let timed = run_one(mode, &fixture).expect("timed");
            assert!(timed.elapsed_ns > 0);
            let observation = observe_output(&timed.output, &fixture).expect("observation");
            require_oracle(&observation).expect("oracle");
        }
    }

    #[test]
    fn cli_bounds_are_closed_and_dimensions_are_restricted() {
        let valid = |extra: &[&str]| {
            let mut args = vec![
                "--parts",
                "4",
                "--mode",
                "publish",
                "--samples",
                "1",
                "--warmup",
                "0",
                "--output",
                "report.json",
            ];
            args.extend_from_slice(extra);
            parse_args(args.into_iter().map(str::to_owned)).expect("valid args")
        };
        assert_eq!(valid(&[]).parts, 4);
        assert!(
            parse_args(
                [
                    "--parts",
                    "5",
                    "--mode",
                    "publish",
                    "--samples",
                    "1",
                    "--warmup",
                    "0",
                    "--output",
                    "x"
                ]
                .into_iter()
                .map(str::to_owned)
            )
            .is_err()
        );
        assert!(
            parse_args(
                [
                    "--parts",
                    "4",
                    "--mode",
                    "publish",
                    "--samples",
                    "0",
                    "--warmup",
                    "0",
                    "--output",
                    "x"
                ]
                .into_iter()
                .map(str::to_owned)
            )
            .is_err()
        );
        assert!(
            parse_args(
                [
                    "--parts",
                    "4",
                    "--mode",
                    "publish",
                    "--samples",
                    "1001",
                    "--warmup",
                    "0",
                    "--output",
                    "x"
                ]
                .into_iter()
                .map(str::to_owned)
            )
            .is_err()
        );
        assert!(
            parse_args(
                [
                    "--parts",
                    "4",
                    "--mode",
                    "publish",
                    "--samples",
                    "1",
                    "--warmup",
                    "101",
                    "--output",
                    "x"
                ]
                .into_iter()
                .map(str::to_owned)
            )
            .is_err()
        );
        assert_eq!(
            valid(&["--artifact", "source.zip"]).artifact,
            Some(PathBuf::from("source.zip"))
        );
    }
}
