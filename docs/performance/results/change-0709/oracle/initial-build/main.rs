//! Independent semantic oracle for the 0709 ordinary DOCX save baseline.
//!
//! The baseline harness proves that its reference publication is deterministic,
//! but that is not a semantic edit oracle.  This small, opt-in binary uses only
//! the public `litchi_docx::Package` API to check one admitted real fixture and
//! one intentionally refused real fixture.  It is kept outside the workspace
//! and outside the timed harness so its checks cannot enter a timing region.

use litchi_docx::Package;
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::ZipArchive;
use std::{
    collections::BTreeMap,
    env,
    error::Error,
    fs,
    io::Cursor,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

const SUCCESS_FIXTURE: &str =
    "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx";
const REFUSAL_FIXTURE: &str =
    "test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx";
const MAIN_MEMBER: &str = "word/document.xml";
const DEFAULT_MARKER: &str = "litchi-perf-0709-appended-paragraph";
const BODY_SECTION_REFUSAL: &str = "body-final section properties are not the final body child";

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    probe: &'static str,
    revision: Option<String>,
    repository_root: String,
    marker: String,
    successful_fixture: SuccessFixtureReport,
    refusal_fixture: RefusalFixtureReport,
    oracle_pass: bool,
}

#[derive(Debug, Serialize)]
struct FixtureIdentity {
    relative_path: &'static str,
    absolute_path: String,
    source_bytes: usize,
    source_sha256: String,
    source_member_count: usize,
    source_main_member_compressed_payload_sha256: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct SemanticSnapshot {
    document_text: Option<String>,
    document_text_error: Option<String>,
    paragraph_text: Option<Vec<String>>,
    paragraph_text_error: Option<String>,
}

#[derive(Debug, Serialize)]
struct MemberComparison {
    output_member_count: usize,
    changed_non_main_compressed_members: Vec<String>,
    missing_non_main_members: Vec<String>,
    added_non_main_members: Vec<String>,
    main_member_compressed_payload_changed: bool,
    unchanged_compressed_payloads_except_main: bool,
}

#[derive(Debug, Serialize)]
struct SuccessChecks {
    source_marker_occurrences: usize,
    output_marker_occurrences: usize,
    marker_is_last_paragraph: bool,
    marker_exactly_once: bool,
    prior_paragraphs_exact_prefix: bool,
    prior_document_text_is_prefix: bool,
    output_reopened: bool,
    semantic_edit_verified: bool,
}

#[derive(Debug, Serialize)]
struct SuccessRouteReport {
    route: &'static str,
    output_bytes: usize,
    output_sha256: String,
    before: SemanticSnapshot,
    after: SemanticSnapshot,
    checks: SuccessChecks,
    members: MemberComparison,
}

#[derive(Debug, Serialize)]
struct SuccessFixtureReport {
    identity: FixtureIdentity,
    source: SemanticSnapshot,
    edit_admitted: bool,
    edit_error: Option<String>,
    routes: Vec<SuccessRouteReport>,
    fixture_pass: bool,
}

#[derive(Debug, Serialize)]
struct RefusalChecks {
    edit_admitted: bool,
    refusal_contains_current_structural_reason: bool,
    output_reopened: bool,
    source_and_output_text_equal: Option<bool>,
    no_marker_added: bool,
    no_non_main_compressed_payload_changed: bool,
    refusal_round_trip_verified: bool,
}

#[derive(Debug, Serialize)]
struct RefusalRouteReport {
    route: &'static str,
    output_bytes: usize,
    output_sha256: String,
    before: SemanticSnapshot,
    after: SemanticSnapshot,
    edit_error: Option<String>,
    checks: RefusalChecks,
    members: MemberComparison,
}

#[derive(Debug, Serialize)]
struct RefusalFixtureReport {
    identity: FixtureIdentity,
    source: SemanticSnapshot,
    historical_0638_outcome: &'static str,
    routes: Vec<RefusalRouteReport>,
    fixture_pass: bool,
}

#[derive(Clone, Debug)]
struct MemberFacts {
    method: String,
    compressed: Vec<u8>,
    uncompressed_size: u64,
}

#[derive(Debug)]
struct Args {
    repository_root: PathBuf,
    marker: String,
    output: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse()?;
    if args.marker.is_empty() {
        return Err("--marker must not be empty".into());
    }

    let success_path = args.repository_root.join(SUCCESS_FIXTURE);
    let refusal_path = args.repository_root.join(REFUSAL_FIXTURE);
    let success_bytes = fs::read(&success_path)?;
    let refusal_bytes = fs::read(&refusal_path)?;

    let success_members = member_facts(&success_bytes)?;
    let refusal_members = member_facts(&refusal_bytes)?;
    let success_identity = fixture_identity(
        SUCCESS_FIXTURE,
        &success_path,
        &success_bytes,
        &success_members,
    );
    let refusal_identity = fixture_identity(
        REFUSAL_FIXTURE,
        &refusal_path,
        &refusal_bytes,
        &refusal_members,
    );

    let mut success_source_package = Package::open(&success_path)?;
    let success_source = semantic_snapshot(&success_source_package);
    let (edit_admitted, edit_error) = match success_source_package.document_mut() {
        Ok(document) => {
            document.add_paragraph_with_text(&args.marker);
            (true, None)
        },
        Err(error) => (false, Some(error.to_string())),
    };
    drop(success_source_package);

    let success_routes = vec![
        run_success_stream(
            &success_path,
            &success_source,
            &success_members,
            &args.marker,
        )?,
        run_success_save(
            &success_path,
            &success_source,
            &success_members,
            &args.marker,
        )?,
    ];
    let success_fixture_pass = edit_admitted
        && edit_error.is_none()
        && success_routes
            .iter()
            .all(|route| route.checks.semantic_edit_verified);

    let mut refusal_source_package = Package::open(&refusal_path)?;
    let refusal_source = semantic_snapshot(&refusal_source_package);
    let refusal_error = match refusal_source_package.document_mut() {
        Ok(_) => None,
        Err(error) => Some(error.to_string()),
    };
    drop(refusal_source_package);

    let refusal_routes = vec![
        run_refusal_stream(&refusal_path, &refusal_source, &refusal_members)?,
        run_refusal_save(&refusal_path, &refusal_source, &refusal_members)?,
    ];
    let refusal_fixture_pass = refusal_error
        .as_deref()
        .is_some_and(|error| error.contains(BODY_SECTION_REFUSAL))
        && refusal_routes
            .iter()
            .all(|route| route.checks.refusal_round_trip_verified);

    let report = Report {
        schema: "litchi.docx-ordinary-save-oracle-0709.v1",
        probe: "public-package-semantic-append-and-refusal-round-trip",
        revision: env::var("LITCHI_GIT_REV").ok(),
        repository_root: args.repository_root.display().to_string(),
        marker: args.marker,
        successful_fixture: SuccessFixtureReport {
            identity: success_identity,
            source: success_source,
            edit_admitted,
            edit_error,
            routes: success_routes,
            fixture_pass: success_fixture_pass,
        },
        refusal_fixture: RefusalFixtureReport {
            identity: refusal_identity,
            source: refusal_source,
            historical_0638_outcome: "0638 recorded a syntax refusal; current source reaches the same contract as a body-final section-placement refusal",
            routes: refusal_routes,
            fixture_pass: refusal_fixture_pass,
        },
        oracle_pass: success_fixture_pass && refusal_fixture_pass,
    };

    let encoded = serde_json::to_vec_pretty(&report)?;
    if let Some(path) = args.output {
        fs::write(&path, &encoded)?;
        println!("wrote {}", path.display());
    } else {
        println!("{}", String::from_utf8(encoded)?);
    }
    if !report.oracle_pass {
        return Err("0709 DOCX semantic oracle failed".into());
    }
    Ok(())
}

impl Args {
    fn parse() -> Result<Self, Box<dyn Error>> {
        let mut repository_root = default_repository_root()?;
        let mut marker = DEFAULT_MARKER.to_owned();
        let mut output = None;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--repo-root" => {
                    repository_root =
                        PathBuf::from(arguments.next().ok_or("--repo-root requires a path")?);
                },
                "--marker" => {
                    marker = arguments.next().ok_or("--marker requires text")?;
                },
                "--output" => {
                    output = Some(PathBuf::from(
                        arguments.next().ok_or("--output requires a path")?,
                    ));
                },
                "--help" | "-h" => {
                    println!(
                        "usage: docx-ordinary-save-oracle-0709 [--repo-root PATH] [--marker TEXT] [--output PATH]"
                    );
                    std::process::exit(0);
                },
                other => return Err(format!("unknown argument {other:?}").into()),
            }
        }
        let repository_root = repository_root.canonicalize()?;
        Ok(Self {
            repository_root,
            marker,
            output,
        })
    }
}

fn default_repository_root() -> Result<PathBuf, Box<dyn Error>> {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../../../")
        .canonicalize()
        .map_err(Into::into)
}

fn fixture_identity(
    relative_path: &'static str,
    path: &Path,
    source: &[u8],
    members: &BTreeMap<String, MemberFacts>,
) -> FixtureIdentity {
    FixtureIdentity {
        relative_path,
        absolute_path: path.display().to_string(),
        source_bytes: source.len(),
        source_sha256: sha256_hex(source),
        source_member_count: members.len(),
        source_main_member_compressed_payload_sha256: members
            .get(MAIN_MEMBER)
            .map(|member| sha256_hex(&member.compressed)),
    }
}

fn semantic_snapshot(package: &Package) -> SemanticSnapshot {
    let document = match package.document() {
        Ok(document) => document,
        Err(error) => {
            let message = error.to_string();
            return SemanticSnapshot {
                document_text: None,
                document_text_error: Some(message.clone()),
                paragraph_text: None,
                paragraph_text_error: Some(message),
            };
        },
    };

    let (document_text, document_text_error) = match document.text() {
        Ok(text) => (Some(text), None),
        Err(error) => (None, Some(error.to_string())),
    };
    let (paragraph_text, paragraph_text_error) = match document.paragraphs() {
        Ok(paragraphs) => {
            let mut text = Vec::with_capacity(paragraphs.len());
            for paragraph in paragraphs {
                match paragraph.text() {
                    Ok(value) => text.push(value),
                    Err(error) => {
                        return SemanticSnapshot {
                            document_text,
                            document_text_error,
                            paragraph_text: None,
                            paragraph_text_error: Some(error.to_string()),
                        };
                    },
                }
            }
            (Some(text), None)
        },
        Err(error) => (None, Some(error.to_string())),
    };
    SemanticSnapshot {
        document_text,
        document_text_error,
        paragraph_text,
        paragraph_text_error,
    }
}

fn run_success_stream(
    source_path: &Path,
    source_snapshot: &SemanticSnapshot,
    source_members: &BTreeMap<String, MemberFacts>,
    marker: &str,
) -> Result<SuccessRouteReport, Box<dyn Error>> {
    let mut package = Package::open(source_path)?;
    append_marker(&mut package, marker)?;
    let mut output = Vec::new();
    package.to_stream(&mut output)?;
    let output_sha256 = sha256_hex(&output);
    let reopened = Package::from_reader(Cursor::new(output.as_slice()))?;
    let after = semantic_snapshot(&reopened);
    let checks = success_checks(source_snapshot, &after, marker, true);
    let members = compare_members(source_members, &member_facts(&output)?);
    Ok(SuccessRouteReport {
        route: "to_stream",
        output_bytes: output.len(),
        output_sha256,
        before: source_snapshot.clone(),
        after,
        checks,
        members,
    })
}

fn run_success_save(
    source_path: &Path,
    source_snapshot: &SemanticSnapshot,
    source_members: &BTreeMap<String, MemberFacts>,
    marker: &str,
) -> Result<SuccessRouteReport, Box<dyn Error>> {
    let output_path = temporary_output_path("success-save")?;
    let mut package = Package::open(source_path)?;
    append_marker(&mut package, marker)?;
    package.save(&output_path)?;
    let output = fs::read(&output_path)?;
    let output_sha256 = sha256_hex(&output);
    let reopened = Package::open(&output_path)?;
    let after = semantic_snapshot(&reopened);
    let _ = fs::remove_file(&output_path);
    let checks = success_checks(source_snapshot, &after, marker, true);
    let members = compare_members(source_members, &member_facts(&output)?);
    Ok(SuccessRouteReport {
        route: "save",
        output_bytes: output.len(),
        output_sha256,
        before: source_snapshot.clone(),
        after,
        checks,
        members,
    })
}

fn run_refusal_stream(
    source_path: &Path,
    source_snapshot: &SemanticSnapshot,
    source_members: &BTreeMap<String, MemberFacts>,
) -> Result<RefusalRouteReport, Box<dyn Error>> {
    let mut package = Package::open(source_path)?;
    let edit_error = refusal_error(&mut package);
    let edit_admitted = edit_error.is_none();
    let mut output = Vec::new();
    package.to_stream(&mut output)?;
    let output_sha256 = sha256_hex(&output);
    let reopened = Package::from_reader(Cursor::new(output.as_slice()))?;
    let after = semantic_snapshot(&mut reopened);
    let members = compare_members(source_members, &member_facts(&output)?);
    let checks = refusal_checks(
        edit_admitted,
        edit_error.as_deref(),
        source_snapshot,
        &after,
        &members,
    );
    Ok(RefusalRouteReport {
        route: "to_stream",
        output_bytes: output.len(),
        output_sha256,
        before: source_snapshot.clone(),
        after,
        edit_error,
        checks,
        members,
    })
}

fn run_refusal_save(
    source_path: &Path,
    source_snapshot: &SemanticSnapshot,
    source_members: &BTreeMap<String, MemberFacts>,
) -> Result<RefusalRouteReport, Box<dyn Error>> {
    let output_path = temporary_output_path("refusal-save")?;
    let mut package = Package::open(source_path)?;
    let edit_error = refusal_error(&mut package);
    let edit_admitted = edit_error.is_none();
    package.save(&output_path)?;
    let output = fs::read(&output_path)?;
    let output_sha256 = sha256_hex(&output);
    let reopened = Package::open(&output_path)?;
    let after = semantic_snapshot(&reopened);
    let _ = fs::remove_file(&output_path);
    let members = compare_members(source_members, &member_facts(&output)?);
    let checks = refusal_checks(
        edit_admitted,
        edit_error.as_deref(),
        source_snapshot,
        &after,
        &members,
    );
    Ok(RefusalRouteReport {
        route: "save",
        output_bytes: output.len(),
        output_sha256,
        before: source_snapshot.clone(),
        after,
        edit_error,
        checks,
        members,
    })
}

fn append_marker(package: &mut Package, marker: &str) -> Result<(), Box<dyn Error>> {
    let document = package.document_mut()?;
    document.add_paragraph_with_text(marker);
    Ok(())
}

fn refusal_error(package: &mut Package) -> Option<String> {
    match package.document_mut() {
        Ok(_) => None,
        Err(error) => Some(error.to_string()),
    }
}

fn success_checks(
    before: &SemanticSnapshot,
    after: &SemanticSnapshot,
    marker: &str,
    output_reopened: bool,
) -> SuccessChecks {
    let before_paragraphs = before.paragraph_text.as_deref().unwrap_or(&[]);
    let after_paragraphs = after.paragraph_text.as_deref().unwrap_or(&[]);
    let source_marker_occurrences = before_paragraphs
        .iter()
        .map(|text| text.match_indices(marker).count())
        .sum();
    let output_marker_occurrences = after_paragraphs
        .iter()
        .map(|text| text.match_indices(marker).count())
        .sum();
    let marker_is_last_paragraph = after_paragraphs.last().is_some_and(|text| text == marker);
    let marker_exactly_once = source_marker_occurrences == 0 && output_marker_occurrences == 1;
    let prior_paragraphs_exact_prefix = after_paragraphs.len() >= before_paragraphs.len()
        && after_paragraphs[..before_paragraphs.len()] == before_paragraphs[..];
    let prior_document_text_is_prefix = before
        .document_text
        .as_deref()
        .zip(after.document_text.as_deref())
        .is_some_and(|(source, output)| output.starts_with(source));
    let semantic_edit_verified = output_reopened
        && before.paragraph_text.is_some()
        && after.paragraph_text.is_some()
        && marker_exactly_once
        && marker_is_last_paragraph
        && prior_paragraphs_exact_prefix;
    SuccessChecks {
        source_marker_occurrences,
        output_marker_occurrences,
        marker_is_last_paragraph,
        marker_exactly_once,
        prior_paragraphs_exact_prefix,
        prior_document_text_is_prefix,
        output_reopened,
        semantic_edit_verified,
    }
}

fn refusal_checks(
    edit_admitted: bool,
    edit_error: Option<&str>,
    before: &SemanticSnapshot,
    after: &SemanticSnapshot,
    members: &MemberComparison,
) -> RefusalChecks {
    let source_text = before.document_text.as_deref();
    let output_text = after.document_text.as_deref();
    let source_and_output_text_equal = source_text.zip(output_text).map(|(a, b)| a == b);
    let no_marker_added = source_text
        .zip(output_text)
        .is_none_or(|(source, output)| source == output);
    let refusal_contains_current_structural_reason =
        edit_error.is_some_and(|error| error.contains(BODY_SECTION_REFUSAL));
    let no_non_main_compressed_payload_changed = members.unchanged_compressed_payloads_except_main
        && !members.main_member_compressed_payload_changed;
    let refusal_round_trip_verified = !edit_admitted
        && refusal_contains_current_structural_reason
        && source_and_output_text_equal.unwrap_or(true)
        && no_marker_added
        && no_non_main_compressed_payload_changed;
    RefusalChecks {
        edit_admitted,
        refusal_contains_current_structural_reason,
        output_reopened: true,
        source_and_output_text_equal,
        no_marker_added,
        no_non_main_compressed_payload_changed,
        refusal_round_trip_verified,
    }
}

fn temporary_output_path(label: &str) -> Result<PathBuf, Box<dyn Error>> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    Ok(env::temp_dir().join(format!(
        "litchi-docx-ordinary-save-oracle-{}-{label}-{nanos}.docx",
        std::process::id()
    )))
}

fn member_facts(bytes: &[u8]) -> Result<BTreeMap<String, MemberFacts>, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut members = BTreeMap::new();
    for header in archive.entries() {
        let header = header?;
        if header.is_dir() {
            continue;
        }
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let method = format!("{:?}", header.compression_method());
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        let start = usize::try_from(start)?;
        let end = usize::try_from(end)?;
        let compressed = bytes
            .get(start..end)
            .ok_or("ZIP member payload range is out of bounds")?
            .to_vec();
        let old = members.insert(
            name.clone(),
            MemberFacts {
                method,
                compressed,
                uncompressed_size: header.uncompressed_size_hint(),
            },
        );
        if old.is_some() {
            return Err(format!("duplicate ZIP member {name:?}").into());
        }
    }
    Ok(members)
}

fn compare_members(
    source: &BTreeMap<String, MemberFacts>,
    output: &BTreeMap<String, MemberFacts>,
) -> MemberComparison {
    let mut changed_non_main_compressed_members = Vec::new();
    let mut missing_non_main_members = Vec::new();
    let mut added_non_main_members = Vec::new();
    for (name, source_member) in source {
        let Some(output_member) = output.get(name) else {
            if name != MAIN_MEMBER {
                missing_non_main_members.push(name.clone());
            }
            continue;
        };
        let same = source_member.method == output_member.method
            && source_member.compressed == output_member.compressed
            && source_member.uncompressed_size == output_member.uncompressed_size;
        if !same && name != MAIN_MEMBER {
            changed_non_main_compressed_members.push(name.clone());
        }
    }
    for name in output.keys() {
        if !source.contains_key(name) && name != MAIN_MEMBER {
            added_non_main_members.push(name.clone());
        }
    }
    let main_member_compressed_payload_changed = source
        .get(MAIN_MEMBER)
        .zip(output.get(MAIN_MEMBER))
        .is_some_and(|(source_member, output_member)| {
            source_member.method != output_member.method
                || source_member.compressed != output_member.compressed
        });
    let unchanged_compressed_payloads_except_main = changed_non_main_compressed_members.is_empty()
        && missing_non_main_members.is_empty()
        && added_non_main_members.is_empty();
    MemberComparison {
        output_member_count: output.len(),
        changed_non_main_compressed_members,
        missing_non_main_members,
        added_non_main_members,
        main_member_compressed_payload_changed,
        unchanged_compressed_payloads_except_main,
    }
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}
