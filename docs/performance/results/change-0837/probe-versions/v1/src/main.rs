//! Small, self-contained publication probe for the OPC writer.
//!
//! The timed operation in `run` is exactly `PackageWriter::to_bytes` on one
//! already-prepared `OpcPackage`.  Package construction, input parsing,
//! mutation, the expected-output oracle, and all validation happen outside
//! the timer.  The probe is deliberately independent of the main workspace;
//! its only production dependencies are `litchi-opc` and the two hash/JSON
//! crates needed for its evidence record.

use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::process;
use std::time::Instant;

use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, XmlPart};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

type ProbeResult<T> = Result<T, Box<dyn Error>>;

const FIXTURE_NAME: &str = "fixture.xml.zip";
const FIXTURE_MANIFEST_NAME: &str = "fixture-manifest.json";
const FIXTURE_SHA_NAME: &str = "fixture.xml.zip.sha256";
const XML_PART: &str = "/custom/data.xml";
const CUSTOM_REL: &str = "http://schemas.litchi.example/relationships/custom-data";
const XML_CONTENT_TYPE: &str = "application/xml";
const TARGET_XML_BYTES: usize = 2 * 1024 * 1024;
const MANY_PARTS: usize = 256;
const MANY_PART_BYTES: usize = 4 * 1024;
const RANDOM_BYTES: usize = 4 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Case {
    FreshSmall,
    FreshXml,
    FreshMany,
    FreshRandom,
    BorrowedEdit,
    OwnedEdit,
}

impl Case {
    fn parse(value: &str) -> ProbeResult<Self> {
        match value {
            "fresh-small" => Ok(Self::FreshSmall),
            "fresh-xml" => Ok(Self::FreshXml),
            "fresh-many" => Ok(Self::FreshMany),
            "fresh-random" => Ok(Self::FreshRandom),
            "borrowed-edit" => Ok(Self::BorrowedEdit),
            "owned-edit" => Ok(Self::OwnedEdit),
            other => Err(format!("unknown case {other:?}").into()),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::FreshSmall => "fresh-small",
            Self::FreshXml => "fresh-xml",
            Self::FreshMany => "fresh-many",
            Self::FreshRandom => "fresh-random",
            Self::BorrowedEdit => "borrowed-edit",
            Self::OwnedEdit => "owned-edit",
        }
    }

    fn is_edit(self) -> bool {
        matches!(self, Self::BorrowedEdit | Self::OwnedEdit)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipFact {
    id: String,
    reltype: String,
    target: String,
    base: String,
    source: Option<String>,
    mode: &'static str,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartFact {
    name: String,
    content_type: String,
    bytes: usize,
    sha256: String,
    relationships: Vec<RelationshipFact>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SemanticOracle {
    package_relationships: Vec<RelationshipFact>,
    parts: Vec<PartFact>,
    sha256: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct MemberFact {
    name: String,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct MemberOracle {
    members: Vec<MemberFact>,
    canonical_sha256: String,
    output_bytes: usize,
    output_sha256: String,
}

#[derive(Clone, Debug)]
struct Prepared {
    package: OpcPackage,
    semantic: SemanticOracle,
    member_oracle: MemberOracle,
    fixture_sha256: Option<String>,
    selected_part: Option<String>,
}

fn main() -> ProbeResult<()> {
    let mut args = env::args().skip(1);
    let command = args.next().ok_or("missing command")?;
    match command.as_str() {
        "prepare" => {
            let directory = PathBuf::from(args.next().ok_or("prepare requires DIR")?);
            if args.next().is_some() {
                return Err("prepare accepts exactly one DIR".into());
            }
            prepare(&directory)
        },
        "run" => {
            let case = Case::parse(&args.next().ok_or("run requires CASE")?)?;
            let samples: usize = args
                .next()
                .ok_or("run requires SAMPLES")?
                .parse()
                .map_err(|_| "SAMPLES must be a non-negative integer")?;
            let warmup: usize = args
                .next()
                .ok_or("run requires WARMUP")?
                .parse()
                .map_err(|_| "WARMUP must be a non-negative integer")?;
            let fixture_directory = PathBuf::from(args.next().ok_or("run requires FIXTUREDIR")?);
            let output = PathBuf::from(args.next().ok_or("run requires OUTPUTJSON")?);
            if args.next().is_some() {
                return Err("run accepts CASE SAMPLES WARMUP FIXTUREDIR OUTPUTJSON".into());
            }
            run(case, samples, warmup, &fixture_directory, &output)
        },
        "qualify" => {
            let fixture_directory =
                PathBuf::from(args.next().ok_or("qualify requires FIXTUREDIR")?);
            let output = PathBuf::from(args.next().ok_or("qualify requires OUTPUTJSON")?);
            if args.next().is_some() {
                return Err("qualify accepts FIXTUREDIR OUTPUTJSON".into());
            }
            qualify(&fixture_directory, &output)
        },
        other => {
            Err(format!("unknown command {other:?}; expected prepare, run, or qualify").into())
        },
    }
}

fn prepare(directory: &Path) -> ProbeResult<()> {
    fs::create_dir_all(directory)?;
    let package = fresh_package(Case::FreshXml)?;
    let semantic = semantic_oracle(&package)?;
    let bytes = PackageWriter::to_bytes(&package)?;
    let members = member_oracle(&bytes)?;

    let fixture_path = directory.join(FIXTURE_NAME);
    fs::write(&fixture_path, &bytes)?;
    let raw_sha256 = sha256_hex(&bytes);
    fs::write(
        directory.join(FIXTURE_SHA_NAME),
        format!("{raw_sha256}  {FIXTURE_NAME}\n"),
    )?;
    let manifest = json!({
        "schema": "fresh-publication-fixture-v1",
        "case": "fresh-xml",
        "fixture": FIXTURE_NAME,
        "raw_fixture_sha256": raw_sha256,
        "fixture_bytes": bytes.len(),
        "input_semantic_sha256": semantic.sha256,
        "member_canonical_sha256": members.canonical_sha256,
        "selected_xml_part": XML_PART,
    });
    write_json(&directory.join(FIXTURE_MANIFEST_NAME), &manifest)?;

    println!(
        "{}",
        serde_json::to_string(&json!({
            "command": "prepare",
            "case": "fresh-xml",
            "fixture": fixture_path,
            "raw_fixture_sha256": raw_sha256,
            "fixture_bytes": bytes.len(),
            "input_semantic_sha256": semantic.sha256,
            "member_canonical_sha256": members.canonical_sha256,
        }))?
    );
    Ok(())
}

fn run(
    case: Case,
    samples: usize,
    warmup: usize,
    fixture_directory: &Path,
    output: &Path,
) -> ProbeResult<()> {
    if samples == 0 {
        return Err("SAMPLES must be positive".into());
    }
    let prepared = prepare_case(case, fixture_directory)?;
    let expected_sha256 = prepared.member_oracle.output_sha256.clone();
    let expected_canonical_sha256 = prepared.member_oracle.canonical_sha256.clone();
    let expected_semantic_sha256 = prepared.semantic.sha256.clone();
    let expected_bytes = prepared.member_oracle.output_bytes;

    let mut warmup_elapsed_ns = Vec::with_capacity(warmup);
    let mut sample_elapsed_ns = Vec::with_capacity(samples);
    let mut all_elapsed_ns = Vec::with_capacity(warmup.saturating_add(samples));
    let mut first_output_sha256: Option<String> = None;
    let mut output_bytes: Option<usize> = None;
    let mut all_validated = true;

    for index in 0..warmup.saturating_add(samples) {
        // The package graph and all input payloads were prepared before this
        // loop. The clock begins immediately before the one operation under
        // test and stops immediately after it returns.
        let package = black_box(&prepared.package);
        let started = Instant::now();
        let bytes = PackageWriter::to_bytes(package)?;
        let elapsed_ns = started.elapsed().as_nanos();
        let facts = validate_output(&bytes, &prepared.semantic, &prepared.member_oracle)?;
        let output_sha256 = facts.output_sha256;
        if let Some(first) = &first_output_sha256 {
            if first != &output_sha256 {
                return Err(format!(
                    "output SHA changed within one process: first {first}, iteration {index} {output_sha256}"
                )
                .into());
            }
        } else {
            first_output_sha256 = Some(output_sha256);
        }
        if let Some(previous_bytes) = output_bytes {
            if previous_bytes != bytes.len() {
                return Err(format!(
                    "output length changed within one process: first {previous_bytes}, iteration {} {}",
                    bytes.len(),
                )
                .into());
            }
        } else {
            output_bytes = Some(bytes.len());
        }
        all_validated &= facts.validated;
        all_elapsed_ns.push(elapsed_ns);
        if index < warmup {
            warmup_elapsed_ns.push(elapsed_ns);
        } else {
            sample_elapsed_ns.push(elapsed_ns);
        }
    }

    let first_output_sha256 = first_output_sha256.ok_or("no output was produced")?;
    let report = json!({
        "schema": "fresh-publication-v1",
        "command": "run",
        "case": case.as_str(),
        "samples": samples,
        "warmup": warmup,
        "elapsed_ns": sample_elapsed_ns,
        "sample_elapsed_ns": sample_elapsed_ns,
        "warmup_elapsed_ns": warmup_elapsed_ns,
        "all_elapsed_ns": all_elapsed_ns,
        "pid": process::id(),
        "vm_hwm_bytes": vm_hwm_bytes(),
        "input_semantic_sha256": expected_semantic_sha256,
        "output_sha256": first_output_sha256,
        "output_bytes": output_bytes.unwrap_or(expected_bytes),
        "expected_output_sha256": expected_sha256,
        "member_canonical_sha256": expected_canonical_sha256,
        "members": member_facts_json(&prepared.member_oracle.members),
        "fixture_raw_sha256": prepared.fixture_sha256,
        "selected_xml_part": prepared.selected_part,
        "output_stable_within_process": true,
        "all_outputs_validated": all_validated,
    });
    write_json(output, &report)?;
    println!("{}", serde_json::to_string(&report)?);
    Ok(())
}

fn prepare_case(case: Case, fixture_directory: &Path) -> ProbeResult<Prepared> {
    if !case.is_edit() {
        let package = fresh_package(case)?;
        let semantic = semantic_oracle(&package)?;
        let expected = PackageWriter::to_bytes(&package)?;
        let member_oracle = member_oracle(&expected)?;
        return Ok(Prepared {
            package,
            semantic,
            member_oracle,
            fixture_sha256: None,
            selected_part: None,
        });
    }

    let fixture_path = fixture_directory.join(FIXTURE_NAME);
    let fixture = fs::read(&fixture_path).map_err(|error| {
        format!(
            "cannot read pinned baseline fixture {}: {error}",
            fixture_path.display()
        )
    })?;
    let fixture_sha256 = sha256_hex(&fixture);
    verify_fixture_manifest(fixture_directory, &fixture_sha256)?;
    let mut package = match case {
        Case::BorrowedEdit => OpcPackage::from_bytes(&fixture)?,
        Case::OwnedEdit => OpcPackage::from_vec(fixture.clone())?,
        _ => unreachable!("fresh cases returned before edit setup"),
    };
    let selected_part = select_xml_part(&package)?;
    let original_len = package.get_part(&selected_part)?.blob().len();
    let replacement = deterministic_replacement_xml(original_len);
    package.get_part_mut(&selected_part)?.set_blob(replacement);
    let semantic = semantic_oracle(&package)?;
    let expected = PackageWriter::to_bytes(&package)?;
    let member_oracle = member_oracle(&expected)?;
    Ok(Prepared {
        package,
        semantic,
        member_oracle,
        fixture_sha256: Some(fixture_sha256),
        selected_part: Some(selected_part.as_str().to_owned()),
    })
}

fn fresh_package(case: Case) -> ProbeResult<OpcPackage> {
    let mut package = OpcPackage::new();
    match case {
        Case::FreshSmall => {
            add_xml_part(&mut package, XML_PART, small_xml().into_bytes())?;
            package.relate_to(XML_PART.trim_start_matches('/'), CUSTOM_REL);
        },
        Case::FreshXml => {
            add_xml_part(&mut package, XML_PART, repeated_xml(TARGET_XML_BYTES, 7))?;
            package.relate_to(XML_PART.trim_start_matches('/'), CUSTOM_REL);
        },
        Case::FreshMany => {
            for index in 0..MANY_PARTS {
                let name = format!("/custom/part-{index:03}.xml");
                add_xml_part(
                    &mut package,
                    &name,
                    fixed_xml(MANY_PART_BYTES, index as u64),
                )?;
            }
            package.relate_to("custom/part-000.xml", CUSTOM_REL);
        },
        Case::FreshRandom => {
            let name = PackURI::new("/custom/random.bin")?;
            package.try_add_part(Box::new(BlobPart::new(
                name,
                "application/octet-stream".to_owned(),
                deterministic_bytes(RANDOM_BYTES, 0x8370_5eed),
            )))?;
            package.relate_to("custom/random.bin", CUSTOM_REL);
        },
        Case::BorrowedEdit | Case::OwnedEdit => {
            return Err("edit cases require the prepared fixture".into());
        },
    }
    Ok(package)
}

fn add_xml_part(package: &mut OpcPackage, name: &str, bytes: Vec<u8>) -> ProbeResult<()> {
    package.try_add_part(Box::new(XmlPart::new(
        PackURI::new(name)?,
        XML_CONTENT_TYPE.to_owned(),
        bytes,
    )))?;
    Ok(())
}

fn small_xml() -> String {
    "<root><item id=\"0\" value=\"small\"/><item id=\"1\" value=\"stable\"/></root>".to_owned()
}

fn repeated_xml(target: usize, seed: u64) -> Vec<u8> {
    let mut xml = String::with_capacity(target);
    xml.push_str("<root>");
    let mut index = 0_u64;
    while xml.len().saturating_add(64) < target {
        let a = index.wrapping_mul(17).wrapping_add(seed) % 1_000_003;
        let b = index.wrapping_mul(31).wrapping_add(seed * 3) % 10_000_019;
        let c = index.wrapping_mul(47).wrapping_add(seed * 11) % 100_000_007;
        // Keep the input minified and vary numeric attributes so the corpus
        // is highly repetitive but not a single repeated byte string.
        use std::fmt::Write as _;
        let _ = write!(xml, "<item a=\"{a}\" b=\"{b}\" c=\"{c}\"/>");
        index = index.wrapping_add(1);
    }
    xml.push_str("</root>");
    xml.into_bytes()
}

fn fixed_xml(target: usize, seed: u64) -> Vec<u8> {
    repeated_xml(target, seed)
}

fn deterministic_replacement_xml(target: usize) -> Vec<u8> {
    let target = target.max(256);
    repeated_xml(target, 0x8370_7bad)
}

fn deterministic_bytes(length: usize, mut state: u64) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(length);
    for _ in 0..length {
        // A fixed, dependency-free generator is sufficient for an
        // incompressible deterministic control; it is not a cryptographic RNG.
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        bytes.push((state >> 24) as u8);
    }
    bytes
}

fn select_xml_part(package: &OpcPackage) -> ProbeResult<PackURI> {
    let mut names: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| {
            part.content_type().contains("xml") || part.partname().as_str().ends_with(".xml")
        })
        .map(|part| part.partname().clone())
        .collect();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    names
        .into_iter()
        .next()
        .ok_or_else(|| "pinned fixture has no XML part".into())
}

fn verify_fixture_manifest(directory: &Path, actual_sha256: &str) -> ProbeResult<()> {
    let manifest_path = directory.join(FIXTURE_MANIFEST_NAME);
    if !manifest_path.exists() {
        return Err(format!(
            "missing pinned fixture manifest {}",
            manifest_path.display()
        )
        .into());
    }
    let manifest: Value = serde_json::from_slice(&fs::read(&manifest_path)?)?;
    let declared = manifest
        .get("raw_fixture_sha256")
        .and_then(Value::as_str)
        .ok_or("fixture manifest has no raw_fixture_sha256")?;
    if declared != actual_sha256 {
        return Err(format!(
            "pinned fixture SHA mismatch: manifest {declared}, actual {actual_sha256}"
        )
        .into());
    }
    Ok(())
}

fn semantic_oracle(package: &OpcPackage) -> ProbeResult<SemanticOracle> {
    let mut package_relationships = relationship_facts(package.rels());
    package_relationships.sort_by(|left, right| left.id.cmp(&right.id));

    let mut parts = Vec::new();
    for metadata in package.iter_parts() {
        let part = package.get_part(metadata.partname())?;
        let mut relationships = relationship_facts(part.rels());
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        parts.push(PartFact {
            name: part.partname().as_str().to_owned(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob().len(),
            sha256: sha256_hex(part.blob()),
            relationships,
        });
    }
    parts.sort_by(|left, right| left.name.cmp(&right.name));
    let sha256 = semantic_sha256(&package_relationships, &parts);
    Ok(SemanticOracle {
        package_relationships,
        parts,
        sha256,
    })
}

fn relationship_facts(relationships: &litchi_opc::Relationships) -> Vec<RelationshipFact> {
    relationships
        .iter()
        .map(|relationship| RelationshipFact {
            id: relationship.r_id().to_owned(),
            reltype: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            base: relationship.base_uri().to_owned(),
            source: relationship.source_uri().map(str::to_owned),
            mode: if relationship.is_external() {
                "external"
            } else {
                "internal"
            },
        })
        .collect()
}

fn semantic_sha256(package_relationships: &[RelationshipFact], parts: &[PartFact]) -> String {
    let mut hasher = Sha256::new();
    hash_tag(&mut hasher, b"opc-semantic-v1");
    hash_relationships(&mut hasher, package_relationships);
    hash_u64(&mut hasher, parts.len() as u64);
    for part in parts {
        hash_string(&mut hasher, &part.name);
        hash_string(&mut hasher, &part.content_type);
        hash_u64(&mut hasher, part.bytes as u64);
        hash_string(&mut hasher, &part.sha256);
        hash_relationships(&mut hasher, &part.relationships);
    }
    hex_bytes(&hasher.finalize())
}

fn hash_relationships(hasher: &mut Sha256, relationships: &[RelationshipFact]) {
    hash_u64(hasher, relationships.len() as u64);
    for relationship in relationships {
        hash_string(hasher, &relationship.id);
        hash_string(hasher, &relationship.reltype);
        hash_string(hasher, &relationship.target);
        hash_string(hasher, &relationship.base);
        hash_string(hasher, relationship.source.as_deref().unwrap_or(""));
        hash_string(hasher, relationship.mode);
    }
}

fn member_oracle(bytes: &[u8]) -> ProbeResult<MemberOracle> {
    let archive = soapberry_zip::office::ArchiveReader::new(bytes)?;
    let mut names: Vec<String> = archive.file_names().map(str::to_owned).collect();
    names.sort();
    if names.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err("archive contains duplicate member names".into());
    }
    let mut members = Vec::with_capacity(names.len());
    let mut canonical = Sha256::new();
    hash_tag(&mut canonical, b"opc-members-decompressed-v1");
    hash_u64(&mut canonical, names.len() as u64);
    for name in names {
        let payload = archive.read(&name)?;
        hash_string(&mut canonical, &name);
        hash_u64(&mut canonical, payload.len() as u64);
        canonical.update(&payload);
        members.push(MemberFact {
            name,
            bytes: payload.len(),
            sha256: sha256_hex(&payload),
        });
    }
    let canonical_sha256 = hex_bytes(&canonical.finalize());
    Ok(MemberOracle {
        members,
        canonical_sha256,
        output_bytes: bytes.len(),
        output_sha256: sha256_hex(bytes),
    })
}

#[derive(Clone, Debug)]
struct OutputFacts {
    output_sha256: String,
    validated: bool,
}

fn validate_output(
    bytes: &[u8],
    expected_semantic: &SemanticOracle,
    expected_members: &MemberOracle,
) -> ProbeResult<OutputFacts> {
    let actual_members = member_oracle(bytes)?;
    if actual_members.members != expected_members.members
        || actual_members.canonical_sha256 != expected_members.canonical_sha256
    {
        return Err("published ZIP member names or decompressed payload hashes changed".into());
    }
    let reopened = OpcPackage::from_bytes(bytes)?;
    let actual_semantic = semantic_oracle(&reopened)?;
    if actual_semantic != *expected_semantic {
        return Err(format!(
            "published OPC semantic oracle changed: expected {}, got {}",
            expected_semantic.sha256, actual_semantic.sha256
        )
        .into());
    }
    Ok(OutputFacts {
        output_sha256: sha256_hex(bytes),
        validated: true,
    })
}

fn member_facts_json(members: &[MemberFact]) -> Value {
    Value::Array(
        members
            .iter()
            .map(|member| {
                json!({
                    "name": member.name,
                    "bytes": member.bytes,
                    "sha256": member.sha256,
                })
            })
            .collect(),
    )
}

fn qualify(fixture_directory: &Path, output: &Path) -> ProbeResult<()> {
    let prepared = prepare_case(Case::OwnedEdit, fixture_directory)?;
    let expected = PackageWriter::to_bytes(&prepared.package)?;

    let mut ordinary = Vec::new();
    PackageWriter::write_to_stream(&mut ordinary, &prepared.package)?;

    let mut short = ShortInterruptedWriter::new(37, 0);
    PackageWriter::write_to_stream(&mut short, &prepared.package)?;

    let mut interrupted = ShortInterruptedWriter::new(37, 3);
    PackageWriter::write_to_stream(&mut interrupted, &prepared.package)?;

    let expected_sha256 = sha256_hex(&expected);
    let ordinary_sha256 = sha256_hex(&ordinary);
    let short_sha256 = sha256_hex(&short.bytes);
    let interrupted_sha256 = sha256_hex(&interrupted.bytes);
    let report = json!({
        "schema": "fresh-publication-stream-qualification-v1",
        "command": "qualify",
        "case": "owned-edit",
        "expected_sha256": expected_sha256,
        "ordinary_sha256": ordinary_sha256,
        "short_sha256": short_sha256,
        "interrupted_sha256": interrupted_sha256,
        "ordinary_parity": ordinary == expected,
        "short_parity": short.bytes == expected,
        "interrupted_parity": interrupted.bytes == expected,
        "short_write_calls": short.write_calls,
        "interrupted_write_calls": interrupted.write_calls,
        "interrupted_retries": interrupted.interruptions_seen,
    });
    if report["ordinary_parity"] != true
        || report["short_parity"] != true
        || report["interrupted_parity"] != true
    {
        return Err("write_to_stream parity qualification failed".into());
    }
    write_json(output, &report)?;
    println!("{}", serde_json::to_string(&report)?);
    Ok(())
}

struct ShortInterruptedWriter {
    bytes: Vec<u8>,
    max_write: usize,
    interruptions_remaining: usize,
    interruptions_seen: usize,
    write_calls: usize,
}

impl ShortInterruptedWriter {
    fn new(max_write: usize, interruptions_remaining: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_write,
            interruptions_remaining,
            interruptions_seen: 0,
            write_calls: 0,
        }
    }
}

impl Write for ShortInterruptedWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.write_calls = self.write_calls.saturating_add(1);
        if self.interruptions_remaining != 0 {
            self.interruptions_remaining -= 1;
            self.interruptions_seen = self.interruptions_seen.saturating_add(1);
            return Err(io::Error::new(io::ErrorKind::Interrupted, "probe retry"));
        }
        let count = bytes.len().min(self.max_write);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn hash_tag(hasher: &mut Sha256, tag: &[u8]) {
    hasher.update((tag.len() as u64).to_le_bytes());
    hasher.update(tag);
}

fn hash_string(hasher: &mut Sha256, value: &str) {
    hash_tag(hasher, value.as_bytes());
}

fn hash_u64(hasher: &mut Sha256, value: u64) {
    hasher.update(value.to_le_bytes());
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    hex_bytes(&digest)
}

fn hex_bytes(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn vm_hwm_bytes() -> Option<u64> {
    let status = fs::read_to_string("/proc/self/status").ok()?;
    let value = status
        .lines()
        .find_map(|line| line.strip_prefix("VmHWM:"))?
        .split_whitespace()
        .next()?
        .parse::<u64>()?;
    value.checked_mul(1024)
}

fn write_json(path: &Path, value: &Value) -> ProbeResult<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)?;
    }
    let mut bytes = serde_json::to_vec_pretty(value)?;
    bytes.push(b'\n');
    fs::write(path, bytes)?;
    Ok(())
}
