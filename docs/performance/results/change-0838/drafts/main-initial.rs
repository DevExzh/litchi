//! Diagnostic for the exact inverse contract of the source-backed plain
//! paragraph DOCX operations.
//!
//! This is deliberately a diagnostic rather than a benchmark. It takes one
//! source archive and one output directory, retains the original source and
//! two deterministic physical controls, and exercises copy(0 -> 2) and
//! removal(0) against all three source variants through both
//! retained-publication and durable-inverse paths. Every forward, durable
//! inverse, and retained-publication inverse archive is retained, together
//! with each durable wire, and complete decompressed member facts are recorded
//! in the report.

use std::env;
use std::error::Error;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{OwnedSource, Position, ReadAt};
use litchi_docx::source_backed;
use litchi_docx::source_backed::paragraph_copy::{Error as CopyError, Patch as CopyPatch};
use litchi_docx::source_backed::paragraph_remove::{Error as RemovalError, Patch as RemovalPatch};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, Part};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::CompressionMethod;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

type DiagnosticResult<T> = Result<T, Box<dyn Error>>;

const MAIN_MEMBER: &str = "word/document.xml";
const REPORT_NAME: &str = "report.json";

#[derive(Clone, Debug, PartialEq, Eq)]
struct MemberFact {
    name: String,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug)]
struct ArchiveFacts {
    bytes: usize,
    sha256: String,
    member_canonical_sha256: String,
    members: Vec<MemberFact>,
}

struct SourceVariant {
    label: &'static str,
    archive_name: &'static str,
    bytes: Vec<u8>,
    facts: ArchiveFacts,
}

fn main() -> DiagnosticResult<()> {
    let mut args = env::args().skip(1);
    let source_path = PathBuf::from(
        args.next()
            .ok_or("usage: source-inverse-diagnostic SOURCE_ZIP OUTPUT_DIRECTORY")?,
    );
    let output_directory = PathBuf::from(
        args.next()
            .ok_or("usage: source-inverse-diagnostic SOURCE_ZIP OUTPUT_DIRECTORY")?,
    );
    if args.next().is_some() {
        return Err("usage: source-inverse-diagnostic SOURCE_ZIP OUTPUT_DIRECTORY".into());
    }
    fs::create_dir_all(&output_directory)?;

    let source = fs::read(&source_path).map_err(|error| {
        format!(
            "cannot read source archive {}: {error}",
            source_path.display()
        )
    })?;
    let source_facts = archive_facts(&source)?;
    let borrowed_control = borrowed_control(&source)?;
    let owned_staged_control = owned_staged_control(&source)?;
    let borrowed_facts = archive_facts(&borrowed_control)?;
    let owned_staged_facts = archive_facts(&owned_staged_control)?;

    if borrowed_facts.members != source_facts.members {
        return Err("borrowed OPC control changed decompressed member bytes".into());
    }
    if owned_staged_facts.members != source_facts.members {
        return Err("owned staged control changed decompressed member bytes".into());
    }

    let variants = [
        SourceVariant {
            label: "original-retained",
            archive_name: "source-original-retained.zip",
            bytes: source.clone(),
            facts: source_facts.clone(),
        },
        SourceVariant {
            label: "borrowed-regenerated",
            archive_name: "source-borrowed-regenerated.zip",
            bytes: borrowed_control.clone(),
            facts: borrowed_facts.clone(),
        },
        SourceVariant {
            label: "current-owned-staged",
            archive_name: "source-current-owned-staged.zip",
            bytes: owned_staged_control.clone(),
            facts: owned_staged_facts.clone(),
        },
    ];
    for variant in &variants {
        write_archive(&output_directory.join(variant.archive_name), &variant.bytes)?;
    }
    let source_label = source_path
        .file_stem()
        .and_then(|value| value.to_str())
        .map_or_else(|| "source".to_owned(), str::to_owned);
    let source_file_name = source_path
        .file_name()
        .and_then(|value| value.to_str())
        .map_or_else(|| "source.zip".to_owned(), str::to_owned);
    let mut operations = Vec::with_capacity(variants.len() * 2);
    for variant in &variants {
        operations.push(run_copy(
            variant.label,
            &variant.bytes,
            &variant.facts,
            &output_directory,
        )?);
        operations.push(run_removal(
            variant.label,
            &variant.bytes,
            &variant.facts,
            &output_directory,
        )?);
    }

    let report = json!({
        "schema": "source-backed-docx-inverse-diagnostic-v1",
        "command": "diagnose",
        "pid": std::process::id(),
        "case": "plain-paragraph-copy-0-to-2-and-removal-0",
        "source_label": source_label,
        "source_path": source_file_name,
        "source": archive_value("source-original-retained.zip", &source_facts),
        "controls": {
            "borrowed_regenerated": archive_value("source-borrowed-regenerated.zip", &borrowed_facts),
            "current_owned_staged": archive_value("source-current-owned-staged.zip", &owned_staged_facts),
            "borrowed_logical_equal_source": borrowed_facts.members == source_facts.members,
            "owned_staged_logical_equal_source": owned_staged_facts.members == source_facts.members,
        },
        "sources": variants.iter().map(|variant| {
            archive_value(variant.archive_name, &variant.facts)
        }).collect::<Vec<_>>(),
        "operations": operations,
    });
    write_json(&output_directory.join(REPORT_NAME), &report)?;
    Ok(())
}

fn borrowed_control(source: &[u8]) -> DiagnosticResult<Vec<u8>> {
    let package = OpcPackage::from_bytes(source)?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn owned_staged_control(source: &[u8]) -> DiagnosticResult<Vec<u8>> {
    let archive = ArchiveReader::new(source)?;
    let mut names: Vec<String> = archive.file_names().map(str::to_owned).collect();
    names.sort();
    if names.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err("source archive contains duplicate member names".into());
    }

    let mut writer = StreamingArchiveWriter::new();
    for name in names {
        let payload = archive.read(&name)?;
        let mut entry = writer
            .start_entry(&name, CompressionMethod::Deflate)
            .map_err(|failure| failure.into_error())?;
        entry.write_all(&payload)?;
        writer = entry.finish().map_err(|failure| failure.into_error())?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn open_source(source: &[u8]) -> DiagnosticResult<source_backed::Package> {
    let read_at: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source.to_vec()));
    Ok(source_backed::Package::from_read_at(read_at)?)
}

fn run_copy(
    source_label: &str,
    source: &[u8],
    source_facts: &ArchiveFacts,
    output_directory: &Path,
) -> DiagnosticResult<Value> {
    let package = open_source(source)?;
    let mut edit = package.edit_plain_paragraph_copy()?;
    edit.copy_plain_paragraph(Position::new(0), Position::new(2))?;
    let commit = edit.commit();
    let projected_xml = commit.projected().xml_bytes().to_vec();

    let mut forward = Vec::new();
    let publication = package.publish_plain_paragraph_copy_to_stream(&mut forward, &commit)?;
    let inverse_wire = publication.inverse_patch().to_bytes()?;
    let durable_inverse = CopyPatch::from_bytes(&inverse_wire)?;
    if durable_inverse.to_bytes()? != inverse_wire {
        return Err("copy durable inverse did not re-encode canonically".into());
    }

    let mut durable_restored = Vec::new();
    open_source(&forward)?
        .publish_plain_paragraph_copy_patch_to_stream(&mut durable_restored, &durable_inverse)?;

    let mut immediate_restored = Vec::new();
    open_source(&forward)?
        .publish_plain_paragraph_copy_inverse_to_stream(&mut immediate_restored, &publication)?;

    let foreign_forward = foreign_archive(&forward)?;
    let mut stale_sink = Vec::new();
    let stale_result = open_source(&foreign_forward)?
        .publish_plain_paragraph_copy_patch_to_stream(&mut stale_sink, &durable_inverse);
    let stale_refused = matches!(stale_result, Err(CopyError::StaleSource));
    if !stale_refused || !stale_sink.is_empty() {
        return Err("copy durable inverse accepted a stale source or wrote to its sink".into());
    }

    let forward_facts = archive_facts(&forward)?;
    let durable_facts = archive_facts(&durable_restored)?;
    let immediate_facts = archive_facts(&immediate_restored)?;
    let expected_forward = expected_members(source_facts, &projected_xml)?;
    let forward_logical_equal = forward_facts.members == expected_forward;
    let durable_logical_equal = durable_facts.members == source_facts.members;
    let immediate_logical_equal = immediate_facts.members == source_facts.members;
    let durable_exact = durable_restored == source;
    let immediate_exact = immediate_restored == source;
    if !forward_logical_equal || !durable_logical_equal || !immediate_logical_equal {
        return Err("copy publication changed logical member bytes".into());
    }
    if !immediate_exact {
        return Err("copy retained-publication inverse did not restore exact bytes".into());
    }

    let forward_name = format!("{source_label}-copy-forward.zip");
    let durable_name = format!("{source_label}-copy-durable-inverse.zip");
    let immediate_name = format!("{source_label}-copy-immediate-inverse.zip");
    let wire_name = format!("{source_label}-copy-inverse.patch");
    write_archive(&output_directory.join(&forward_name), &forward)?;
    write_archive(&output_directory.join(&durable_name), &durable_restored)?;
    write_archive(&output_directory.join(&immediate_name), &immediate_restored)?;
    write_archive(&output_directory.join(&wire_name), &inverse_wire)?;

    Ok(json!({
        "case": "plain-paragraph-copy-0-to-2",
        "source_variant": source_label,
        "source_member_count": source_facts.members.len(),
        "forward": archive_value_with_logical(
            forward_name, &forward_facts, forward_logical_equal,
        ),
        "durable_inverse": archive_value_with_logical(
            durable_name, &durable_facts, durable_logical_equal,
        ),
        "immediate_inverse": archive_value_with_logical(
            immediate_name, &immediate_facts, immediate_logical_equal,
        ),
        "durable_wire": {
            "path": wire_name,
            "bytes": inverse_wire.len(),
            "sha256": sha256_hex(&inverse_wire),
        },
        "logical_equal": forward_logical_equal && durable_logical_equal && immediate_logical_equal,
        "immediate_inverse_exact": immediate_exact,
        "durable_inverse_exact": durable_exact,
        "stale_refused": stale_refused,
        "stale_sink_empty": stale_sink.is_empty(),
    }))
}

fn run_removal(
    source_label: &str,
    source: &[u8],
    source_facts: &ArchiveFacts,
    output_directory: &Path,
) -> DiagnosticResult<Value> {
    let package = open_source(source)?;
    let mut edit = package.edit_plain_paragraph_removal()?;
    edit.remove_plain_paragraph(Position::new(0))?;
    let commit = edit.commit();
    let projected_xml = commit.projected().xml_bytes().to_vec();

    let mut forward = Vec::new();
    let publication = package.publish_plain_paragraph_removal_to_stream(&mut forward, &commit)?;
    let inverse_wire = publication.inverse_patch().to_bytes()?;
    let durable_inverse = RemovalPatch::from_bytes(&inverse_wire)?;
    if durable_inverse.to_bytes()? != inverse_wire {
        return Err("removal durable inverse did not re-encode canonically".into());
    }

    let mut durable_restored = Vec::new();
    open_source(&forward)?
        .publish_plain_paragraph_removal_patch_to_stream(&mut durable_restored, &durable_inverse)?;

    let mut immediate_restored = Vec::new();
    open_source(&forward)?
        .publish_plain_paragraph_removal_inverse_to_stream(&mut immediate_restored, &publication)?;

    let foreign_forward = foreign_archive(&forward)?;
    let mut stale_sink = Vec::new();
    let stale_result = open_source(&foreign_forward)?
        .publish_plain_paragraph_removal_patch_to_stream(&mut stale_sink, &durable_inverse);
    let stale_refused = matches!(stale_result, Err(RemovalError::StaleSource));
    if !stale_refused || !stale_sink.is_empty() {
        return Err("removal durable inverse accepted a stale source or wrote to its sink".into());
    }

    let forward_facts = archive_facts(&forward)?;
    let durable_facts = archive_facts(&durable_restored)?;
    let immediate_facts = archive_facts(&immediate_restored)?;
    let expected_forward = expected_members(source_facts, &projected_xml)?;
    let forward_logical_equal = forward_facts.members == expected_forward;
    let durable_logical_equal = durable_facts.members == source_facts.members;
    let immediate_logical_equal = immediate_facts.members == source_facts.members;
    let durable_exact = durable_restored == source;
    let immediate_exact = immediate_restored == source;
    if !forward_logical_equal || !durable_logical_equal || !immediate_logical_equal {
        return Err("removal publication changed logical member bytes".into());
    }
    if !immediate_exact {
        return Err("removal retained-publication inverse did not restore exact bytes".into());
    }

    let forward_name = format!("{source_label}-removal-forward.zip");
    let durable_name = format!("{source_label}-removal-durable-inverse.zip");
    let immediate_name = format!("{source_label}-removal-immediate-inverse.zip");
    let wire_name = format!("{source_label}-removal-inverse.patch");
    write_archive(&output_directory.join(&forward_name), &forward)?;
    write_archive(&output_directory.join(&durable_name), &durable_restored)?;
    write_archive(&output_directory.join(&immediate_name), &immediate_restored)?;
    write_archive(&output_directory.join(&wire_name), &inverse_wire)?;

    Ok(json!({
        "case": "plain-paragraph-removal-0",
        "source_variant": source_label,
        "source_member_count": source_facts.members.len(),
        "forward": archive_value_with_logical(
            forward_name, &forward_facts, forward_logical_equal,
        ),
        "durable_inverse": archive_value_with_logical(
            durable_name, &durable_facts, durable_logical_equal,
        ),
        "immediate_inverse": archive_value_with_logical(
            immediate_name, &immediate_facts, immediate_logical_equal,
        ),
        "durable_wire": {
            "path": wire_name,
            "bytes": inverse_wire.len(),
            "sha256": sha256_hex(&inverse_wire),
        },
        "logical_equal": forward_logical_equal && durable_logical_equal && immediate_logical_equal,
        "immediate_inverse_exact": immediate_exact,
        "durable_inverse_exact": durable_exact,
        "stale_refused": stale_refused,
        "stale_sink_empty": stale_sink.is_empty(),
    }))
}

fn foreign_archive(bytes: &[u8]) -> DiagnosticResult<Vec<u8>> {
    let mut package = OpcPackage::from_bytes(bytes)?;
    let name = package
        .iter_parts()
        .map(|part| part.partname().as_str().to_owned())
        .find(|name| name != "/word/document.xml")
        .ok_or("archive has no non-main member to mutate")?;
    let part_name = PackURI::new(name)?;
    let part = package.get_part_mut(&part_name)?;
    let mut changed = part.blob().to_vec();
    changed.extend_from_slice(b"\0stale-diagnostic");
    part.set_blob(changed);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn expected_members(
    source: &ArchiveFacts,
    projected_xml: &[u8],
) -> DiagnosticResult<Vec<MemberFact>> {
    let mut members = source.members.clone();
    let main = members
        .iter_mut()
        .find(|member| member.name == MAIN_MEMBER)
        .ok_or("source archive has no word/document.xml member")?;
    main.bytes = projected_xml.len();
    main.sha256 = sha256_hex(projected_xml);
    Ok(members)
}

fn archive_facts(bytes: &[u8]) -> DiagnosticResult<ArchiveFacts> {
    let archive = ArchiveReader::new(bytes)?;
    let mut names: Vec<String> = archive.file_names().map(str::to_owned).collect();
    names.sort();
    if names.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err("archive contains duplicate member names".into());
    }

    let mut members = Vec::with_capacity(names.len());
    let mut canonical = Sha256::new();
    hash_tag(&mut canonical, b"docx-diagnostic-members-v1");
    hash_u64(&mut canonical, names.len() as u64);
    for name in names {
        let payload = archive.read(&name)?;
        let payload_sha256 = sha256_hex(&payload);
        hash_string(&mut canonical, &name);
        hash_u64(&mut canonical, payload.len() as u64);
        canonical.update(&payload);
        members.push(MemberFact {
            name,
            bytes: payload.len(),
            sha256: payload_sha256,
        });
    }

    Ok(ArchiveFacts {
        bytes: bytes.len(),
        sha256: sha256_hex(bytes),
        member_canonical_sha256: hex_bytes(&canonical.finalize()),
        members,
    })
}

fn archive_value(path: &str, facts: &ArchiveFacts) -> Value {
    archive_value_with_logical(path, facts, true)
}

fn archive_value_with_logical(
    path: impl Into<Value>,
    facts: &ArchiveFacts,
    logical_equal: bool,
) -> Value {
    json!({
        "archive": path.into(),
        "bytes": facts.bytes,
        "raw_sha256": facts.sha256,
        "member_canonical_sha256": facts.member_canonical_sha256,
        "members": facts.members.iter().map(|member| json!({
            "name": member.name,
            "bytes": member.bytes,
            "sha256": member.sha256,
        })).collect::<Vec<_>>(),
        "logical_equal": logical_equal,
    })
}

fn write_archive(path: &Path, bytes: &[u8]) -> DiagnosticResult<()> {
    fs::write(path, bytes)
        .map_err(|error| format!("cannot write archive {}: {error}", path.display()).into())
}

fn write_json(path: &Path, value: &Value) -> DiagnosticResult<()> {
    let bytes = serde_json::to_vec_pretty(value)?;
    fs::write(path, bytes)
        .map_err(|error| format!("cannot write report {}: {error}", path.display()).into())
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex_bytes(&hasher.finalize())
}

fn hash_tag(hasher: &mut Sha256, tag: &[u8]) {
    hash_bytes(hasher, tag);
}

fn hash_string(hasher: &mut Sha256, value: &str) {
    hash_bytes(hasher, value.as_bytes());
}

fn hash_bytes(hasher: &mut Sha256, value: &[u8]) {
    hash_u64(hasher, value.len() as u64);
    hasher.update(value);
}

fn hash_u64(hasher: &mut Sha256, value: u64) {
    hasher.update(value.to_le_bytes());
}

fn hex_bytes(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}
