//! Deterministic diagnostic probe for DOCX active-offset selection.
//!
//! This is an observation tool, rather than a benchmark.  It compares the
//! public `alt::active` results for all synthetic Word block starts, for only
//! `altChunk` starts, and for no offsets.  It also sends each XML document
//! through the public mutable-package facade so a zero-anchor `alt::scan`
//! result cannot hide a refusal from the writer's range selection.

use litchi_docx::{
    Package,
    alt::{active, scan},
};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{env, error::Error, fs, path::PathBuf};

const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const OFFSET_LIMIT: usize = 1_000_000;

const BLOCK_PREFIXES: &[&[u8]] = &[
    b"<w:p",
    b"<w:tbl",
    b"<w:altChunk",
    b"<s:p",
    b"<s:tbl",
    b"<s:altChunk",
];
const ALT_PREFIXES: &[&[u8]] = &[b"<w:altChunk", b"<s:altChunk"];

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    oracle: &'static str,
    case_count: usize,
    cases: Vec<CaseResult>,
    packages: Vec<serde_json::Value>,
}

#[derive(Debug, Serialize)]
struct CaseResult {
    name: String,
    xml_sha256: String,
    xml_bytes: usize,
    ids: Vec<String>,
    known_body_block_starts: Vec<u32>,
    alt_chunk_starts: Vec<u32>,
    note: String,
    scan: ScanOutcome,
    active_full: ActiveProbe,
    active_all_alt: ActiveProbe,
    active_empty: ActiveProbe,
    document_mut: FacadeOutcome,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum ScanOutcome {
    Success { chunks: Vec<ChunkResult> },
    Error { display: String, debug: String },
}

#[derive(Debug, Serialize)]
struct ChunkResult {
    offset: u32,
    relationship: String,
    match_source: Option<bool>,
    debug: String,
}

#[derive(Debug, Serialize)]
struct ActiveProbe {
    input_count: usize,
    input_sha256: String,
    input_offsets: Option<Vec<u32>>,
    outcome: ActiveOutcome,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum ActiveOutcome {
    Success { offsets: Vec<u32> },
    Error { display: String, debug: String },
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum FacadeOutcome {
    Success {
        alt_count: usize,
    },
    Error {
        stage: &'static str,
        display: String,
        debug: String,
    },
}

struct CaseInput {
    name: &'static str,
    xml: Vec<u8>,
    ids: Vec<&'static str>,
    note: &'static str,
    full_override: Option<Vec<u32>>,
}

#[derive(Debug)]
struct Args {
    output: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse()?;
    let cases = build_cases();
    let packages = package_cases()?;
    let results = cases.iter().map(run_case).collect::<Vec<_>>();
    let report = Report {
        schema: "litchi.docx-active-offset-oracle-0721.v1",
        oracle: "public-alt-active-and-document-mut-diagnostic-matrix",
        case_count: results.len(),
        cases: results,
        packages,
    };
    let encoded = serde_json::to_vec_pretty(&report)?;
    if let Some(path) = args.output {
        if let Some(parent) = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
        {
            fs::create_dir_all(parent)?;
        }
        fs::write(&path, encoded)?;
        println!("wrote {}", path.display());
    } else {
        println!("{}", String::from_utf8(encoded)?);
    }
    Ok(())
}

impl Args {
    fn parse() -> Result<Self, Box<dyn Error>> {
        let mut output = None;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--output" => {
                    output = Some(PathBuf::from(
                        arguments.next().ok_or("--output requires a path")?,
                    ));
                },
                "--help" | "-h" => {
                    println!("usage: docx-active-offset-oracle-0721 [--output PATH]");
                    std::process::exit(0);
                },
                other => return Err(format!("unknown argument {other:?}").into()),
            }
        }
        Ok(Self { output })
    }
}

fn run_case(case: &CaseInput) -> CaseResult {
    eprintln!("CASE0721 {}", case.name);
    let known_body_block_starts = tag_offsets(&case.xml, BLOCK_PREFIXES);
    let alt_chunk_starts = tag_offsets(&case.xml, ALT_PREFIXES);
    let full_offsets = case
        .full_override
        .clone()
        .unwrap_or_else(|| known_body_block_starts.clone());
    CaseResult {
        name: case.name.to_owned(),
        xml_sha256: sha256_hex(&case.xml),
        xml_bytes: case.xml.len(),
        ids: case.ids.iter().map(|id| (*id).to_owned()).collect(),
        known_body_block_starts,
        alt_chunk_starts: alt_chunk_starts.clone(),
        note: case.note.to_owned(),
        scan: scan_outcome(&case.xml),
        active_full: active_probe(&case.xml, full_offsets),
        active_all_alt: active_probe(&case.xml, alt_chunk_starts),
        active_empty: active_probe(&case.xml, Vec::new()),
        document_mut: document_mut_outcome(&case.xml),
    }
}

fn scan_outcome(xml: &[u8]) -> ScanOutcome {
    match scan(xml) {
        Ok(chunks) => ScanOutcome::Success {
            chunks: chunks
                .into_iter()
                .map(|(offset, chunk)| ChunkResult {
                    offset,
                    relationship: chunk.relationship().as_str().to_owned(),
                    match_source: chunk.match_source(),
                    debug: format!("{chunk:?}"),
                })
                .collect(),
        },
        Err(error) => ScanOutcome::Error {
            display: error.to_string(),
            debug: format!("{error:?}"),
        },
    }
}

fn active_probe(xml: &[u8], offsets: Vec<u32>) -> ActiveProbe {
    let input_count = offsets.len();
    let input_sha256 = offsets_sha256(&offsets);
    let input_offsets = (input_count <= 64).then_some(offsets.clone());
    let outcome = match active(xml, &offsets) {
        Ok(offsets) => ActiveOutcome::Success { offsets },
        Err(error) => ActiveOutcome::Error {
            display: error.to_string(),
            debug: format!("{error:?}"),
        },
    };
    ActiveProbe {
        input_count,
        input_sha256,
        input_offsets,
        outcome,
    }
}

fn document_mut_outcome(xml: &[u8]) -> FacadeOutcome {
    let document_uri = match PackURI::new("/word/document.xml") {
        Ok(uri) => uri,
        Err(error) => {
            return FacadeOutcome::Error {
                stage: "construct_document_part",
                display: error.to_string(),
                debug: format!("{error:?}"),
            };
        },
    };
    let document = BlobPart::new(document_uri, ct::WML_DOCUMENT_MAIN.to_owned(), xml.to_vec());
    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(document));
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rId1".to_owned(),
        false,
    );
    let mut package = match Package::from_opc_package(opc) {
        Ok(package) => package,
        Err(error) => {
            return FacadeOutcome::Error {
                stage: "from_opc_package",
                display: error.to_string(),
                debug: format!("{error:?}"),
            };
        },
    };
    match package.document_mut() {
        Ok(document) => FacadeOutcome::Success {
            alt_count: document.alts().len(),
        },
        Err(error) => FacadeOutcome::Error {
            stage: "document_mut",
            display: error.to_string(),
            debug: format!("{error:?}"),
        },
    }
}

fn build_cases() -> Vec<CaseInput> {
    let mut cases = vec![
        CaseInput {
            name: "transitional-no-alt-plain-blocks",
            xml: transitional("<w:p/><w:tbl/><w:p/>"),
            ids: vec![],
            note: "No altChunk exists: scan is empty, while active still receives p/tbl starts.",
            full_override: None,
        },
        CaseInput {
            name: "strict-no-alt-plain-blocks",
            xml: strict("<s:p/><s:tbl/><s:p/>"),
            ids: vec![],
            note: "Strict namespace spelling with no altChunk.",
            full_override: None,
        },
        CaseInput {
            name: "valid-fallback-no-alt-blocks",
            xml: mce(
                r#"<mc:AlternateContent><mc:Choice Requires="u"><w:p/><w:tbl/></mc:Choice><mc:Fallback><w:p/><w:tbl/></mc:Fallback></mc:AlternateContent><w:p/>"#,
            ),
            ids: vec![],
            note: "A valid MCE fallback selects only its p/tbl starts even though scan has zero chunks.",
            full_override: None,
        },
        CaseInput {
            name: "malformed-alternate-content-no-choice-no-alt",
            xml: mce(r#"<mc:AlternateContent><w:p/><w:tbl/></mc:AlternateContent><w:p/>"#),
            ids: vec![],
            note: "Structurally valid XML with malformed MCE: no Choice. Empty offsets fast-path; full offsets must diagnose it.",
            full_override: None,
        },
        CaseInput {
            name: "unknown-must-understand-no-alt",
            xml: mce_with_root_attrs(r#" mc:MustUnderstand="u""#, r#"<w:p/><w:tbl/>"#),
            ids: vec![],
            note: "Unknown MustUnderstand namespace with no altChunk; the writer still has p/tbl ranges to select.",
            full_override: None,
        },
        CaseInput {
            name: "valid-fallback-with-inactive-anchor",
            xml: mce(
                r#"<mc:AlternateContent><mc:Choice Requires="u"><w:altChunk r:id="inactive-choice"/><w:p/></mc:Choice><mc:Fallback><w:altChunk r:id="active-fallback"/><w:tbl/></mc:Fallback></mc:AlternateContent><w:p/>"#,
            ),
            ids: vec!["inactive-choice", "active-fallback"],
            note: "Unsupported Choice is suppressed; fallback anchor and blocks remain active.",
            full_override: None,
        },
        CaseInput {
            name: "valid-choice-with-inactive-fallback",
            xml: mce(
                r#"<mc:AlternateContent><mc:Choice Requires="w"><w:altChunk r:id="active-choice"/><w:p/></mc:Choice><mc:Fallback><w:altChunk r:id="inactive-fallback"/><w:tbl/></mc:Fallback></mc:AlternateContent><w:p/>"#,
            ),
            ids: vec!["active-choice", "inactive-fallback"],
            note: "The baseline Word namespace satisfies Requires=w; the fallback anchor is inactive.",
            full_override: None,
        },
        CaseInput {
            name: "strict-valid-fallback-with-inactive-anchor",
            xml: strict_mce(
                r#"<mc:AlternateContent><mc:Choice Requires="u"><s:altChunk r:id="strict-inactive"/><s:p/></mc:Choice><mc:Fallback><s:altChunk r:id="strict-active"/><s:tbl/></mc:Fallback></mc:AlternateContent>"#,
            ),
            ids: vec!["strict-inactive", "strict-active"],
            note: "Strict Word and relationship namespaces use the same active fallback semantics.",
            full_override: None,
        },
        CaseInput {
            name: "nested-alt-in-paragraph",
            xml: transitional(
                r#"<w:p><w:altChunk r:id="nested-paragraph"/></w:p><w:altChunk r:id="top-level"/>"#,
            ),
            ids: vec!["nested-paragraph", "top-level"],
            note: "active sees both source starts; the writer's range scanner captures the outer p and suppresses nested target ranges.",
            full_override: None,
        },
        CaseInput {
            name: "nested-alt-in-table",
            xml: transitional(
                r#"<w:tbl><w:tr><w:tc><w:p/><w:altChunk r:id="nested-table"/></w:tc></w:tr></w:tbl><w:altChunk r:id="top-level"/>"#,
            ),
            ids: vec!["nested-table", "top-level"],
            note: "active sees nested table content; the writer's range scanner captures the outer tbl and suppresses nested targets.",
            full_override: None,
        },
        CaseInput {
            name: "active-offset-count-overflow",
            xml: transitional("<w:p/><w:tbl/>"),
            ids: vec![],
            note: "The full probe intentionally supplies OFFSET_LIMIT+1 repeated offsets; the public constant is crate-private, so the observed boundary is recorded explicitly.",
            full_override: Some(vec![0; OFFSET_LIMIT + 1]),
        },
    ];
    extra_cases(&mut cases);
    cases
}

fn transitional(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn strict(body: &str) -> Vec<u8> {
    format!(
        r#"<s:document xmlns:s="{STRICT_WORD}" xmlns:r="{STRICT_RELATIONSHIPS}"><s:body>{body}</s:body></s:document>"#
    )
    .into_bytes()
}

fn mce(body: &str) -> Vec<u8> {
    mce_with_root_attrs("", body)
}

fn mce_with_root_attrs(attributes: &str, body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}" xmlns:mc="{MCE_NAMESPACE}" xmlns:u="urn:unsupported"{attributes}><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn strict_mce(body: &str) -> Vec<u8> {
    format!(
        r#"<s:document xmlns:s="{STRICT_WORD}" xmlns:r="{STRICT_RELATIONSHIPS}" xmlns:mc="{MCE_NAMESPACE}" xmlns:u="urn:unsupported"><s:body>{body}</s:body></s:document>"#
    )
    .into_bytes()
}

fn tag_offsets(xml: &[u8], prefixes: &[&[u8]]) -> Vec<u32> {
    let mut offsets = Vec::new();
    for prefix in prefixes {
        if prefix.is_empty() {
            continue;
        }
        for start in 0..=xml.len().saturating_sub(prefix.len()) {
            if !xml[start..].starts_with(prefix) {
                continue;
            }
            let boundary = xml.get(start + prefix.len()).copied();
            if matches!(
                boundary,
                Some(b'>') | Some(b'/') | Some(b' ') | Some(b'\t') | Some(b'\r') | Some(b'\n')
            ) {
                offsets.push(u32::try_from(start).expect("synthetic XML offset fits u32"));
            }
        }
    }
    offsets.sort_unstable();
    offsets
}

fn offsets_sha256(offsets: &[u32]) -> String {
    let mut bytes = Vec::with_capacity(offsets.len().saturating_mul(4));
    for offset in offsets {
        bytes.extend_from_slice(&offset.to_le_bytes());
    }
    sha256_hex(&bytes)
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn extra_cases(cases: &mut Vec<CaseInput>) {
    let mut add = |name, xml, note| {
        cases.push(CaseInput {
            name,
            xml,
            note,
            ids: vec![],
            full_override: None,
        })
    };
    add(
        "bom-plain",
        [b"\xef\xbb\xbf".as_slice(), &transitional("<w:p/>")].concat(),
        "Offsets must address the BOM-stripped body input.",
    );
    add(
        "marker-text",
        transitional("<w:p><w:r><w:t>litchi-mce-active-0000000000000000-00:</w:t></w:r></w:p>"),
        "Marker-like source text must survive selection.",
    );
    add(
        "foreign-lookalikes",
        transitional(
            r#"<x:p xmlns:x="urn:foreign"><w:p/></x:p><x:altChunk xmlns:x="urn:foreign"/>"#,
        ),
        "Foreign names do not become Word targets.",
    );
    add(
        "range-depth-only",
        transitional(&format!(
            "{}<w:p/>{}",
            "<w:custom>".repeat(127),
            "</w:custom>".repeat(127)
        )),
        "Range depth 128 fails before alt depth 256.",
    );
    add(
        "range-depth-before-bad-alt",
        transitional(&format!(
            "{}{}<w:altChunk/>",
            "<w:custom>".repeat(127),
            "</w:custom>".repeat(127)
        )),
        "Later alt relationship error must precede earlier range depth error.",
    );
    add(
        "empty-at-depth-256",
        transitional(&format!(
            "{}<w:custom/>{}",
            "<w:custom>".repeat(254),
            "</w:custom>".repeat(254)
        )),
        "Alt counts Empty depth; range scanner only counts Start nesting.",
    );
    add(
        "unbound-fragment",
        b"<document><body><p/><altChunk/></body></document>".to_vec(),
        "Range fragment heuristic differs from alt URI matching.",
    );
    add(
        "malformed-tail",
        transitional("<w:p>").into_iter().collect(),
        "Malformed XML error remains observable.",
    );
}

fn package_cases() -> Result<Vec<serde_json::Value>, Box<dyn Error>> {
    use std::io::Cursor;
    eprintln!("CASE0721 generated-setup");
    let mut package = Package::new()?;
    let document = package.document_mut()?;
    for index in 0..200 {
        document.add_paragraph_with_text(&format!(
            "litchi-perf-baseline-docx-semantic-v1-source-{index:05}"
        ));
    }
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output)?;
    let generated = output.into_inner();
    let real =
        fs::read("test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx")?;
    let mut results = Vec::new();
    for (name, bytes) in [("generated-medium", generated), ("numbered-list", real)] {
        eprintln!("CASE0721 {name}");
        let mut package = Package::from_reader(Cursor::new(bytes.clone()))?;
        let outcome = match package.document_mut() {
            Ok(doc) => {
                let count = doc.alts().len();
                doc.add_paragraph_with_text("litchi-0721-fusion-publication-marker");
                format!("Ok(alts={count})")
            },
            Err(error) => format!("Err({error:?})"),
        };
        let mut publication = Cursor::new(Vec::new());
        package.to_stream(&mut publication)?;
        let publication = publication.into_inner();
        results.push(serde_json::json!({"name":name,"archive_bytes":bytes.len(),"archive_sha256":sha256_hex(&bytes),"outcome":outcome,"publication_bytes":publication.len(),"publication_sha256":sha256_hex(&publication)}));
    }
    Ok(results)
}
