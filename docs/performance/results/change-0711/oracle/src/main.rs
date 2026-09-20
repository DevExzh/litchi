//! Standalone differential oracle for the public `litchi_docx::alt::scan` API.
//!
//! The binary deliberately records scanner errors as exact outcomes.  It does
//! not catch panics: a panic is a failed run and cannot become a passing case.

use litchi_docx::alt::{MAX_CHUNKS, MAX_XML_BYTES, MAX_XML_DEPTH, scan};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::{
    env,
    error::Error,
    fmt::Write as _,
    fs,
    path::{Path, PathBuf},
};

const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAIN_MEMBER: &str = "word/document.xml";
const NUMBERED_LIST_FIXTURE: &str =
    "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx";
const ALT_CHUNK_FIXTURE: &str =
    "test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx";

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    oracle: &'static str,
    case_count: usize,
    cases: Vec<CaseResult>,
}

#[derive(Debug, Serialize)]
struct CaseResult {
    name: String,
    source: String,
    source_sha256: String,
    input_sha256: String,
    input_bytes: usize,
    outcome: Outcome,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum Outcome {
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

struct CaseInput {
    name: String,
    source: String,
    source_sha256: String,
    input: Vec<u8>,
}

#[derive(Debug)]
struct Args {
    repository_root: PathBuf,
    output: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse()?;
    let cases = build_cases(&args.repository_root)?;
    let results = cases.iter().map(run_case).collect::<Vec<_>>();
    let report = Report {
        schema: "litchi.docx-alt-scan-oracle-0711.v1",
        oracle: "public-alt-scan-deterministic-differential-matrix",
        case_count: results.len(),
        cases: results,
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
        let mut repository_root = default_repository_root()?;
        let mut output = None;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--repo-root" => {
                    repository_root =
                        PathBuf::from(arguments.next().ok_or("--repo-root requires a path")?);
                },
                "--output" => {
                    output = Some(PathBuf::from(
                        arguments.next().ok_or("--output requires a path")?,
                    ));
                },
                "--help" | "-h" => {
                    println!("usage: docx-alt-scan-oracle-0711 [--repo-root PATH] [--output PATH]");
                    std::process::exit(0);
                },
                other => return Err(format!("unknown argument {other:?}").into()),
            }
        }
        Ok(Self {
            repository_root: repository_root.canonicalize()?,
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

fn run_case(case: &CaseInput) -> CaseResult {
    let outcome = match scan(&case.input) {
        Ok(chunks) => Outcome::Success {
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
        Err(error) => Outcome::Error {
            display: error.to_string(),
            debug: format!("{error:?}"),
        },
    };
    CaseResult {
        name: case.name.clone(),
        source: case.source.clone(),
        source_sha256: case.source_sha256.clone(),
        input_sha256: sha256_hex(&case.input),
        input_bytes: case.input.len(),
        outcome,
    }
}

fn build_cases(repository_root: &Path) -> Result<Vec<CaseInput>, Box<dyn Error>> {
    let mut cases = Vec::new();

    push_inline(
        &mut cases,
        "transitional-basic",
        transitional_document(r#"<w:altChunk r:id="rIdBasic"/>"#),
    );
    push_inline(
        &mut cases,
        "strict-basic",
        strict_document(r#"<w:altChunk r:id="rIdStrict"/>"#),
    );
    push_inline(
        &mut cases,
        "custom-prefixes",
        format!(
            r#"<wd:document xmlns:wd="{TRANSITIONAL_WORD}" xmlns:rel="{TRANSITIONAL_RELATIONSHIPS}"><wd:body><wd:altChunk rel:id="rIdCustom"/></wd:body></wd:document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "default-word-namespace",
        format!(
            r#"<document xmlns="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><body><altChunk r:id="rIdDefault"/></body></document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "namespace-rebinding",
        transitional_document(
            r#"<w:altChunk r:id="rIdBefore"/><w:container xmlns:w="urn:rebound"><w:altChunk r:id="rIdHidden"/></w:container><w:altChunk r:id="rIdAfter"/>"#,
        ),
    );
    push_inline(
        &mut cases,
        "temporary-prefix-default-rebinding-after-empty-and-end",
        transitional_document(
            r#"<w:altChunk r:id="rIdBefore"/><w:shadow xmlns:w="urn:shadow"><w:empty/></w:shadow><foreign xmlns="urn:foreign"><empty/></foreign><w:altChunk r:id="rIdAfter"/>"#,
        ),
    );
    push_inline(
        &mut cases,
        "relationship-namespace-rebinding",
        transitional_document(r#"<w:altChunk xmlns:r="urn:wrong" r:id="rIdWrong"/>"#),
    );
    push_inline(
        &mut cases,
        "matchsrc-transitional-values",
        transitional_document(
            r#"<w:altChunk r:id="rIdTrue"><w:altChunkPr><w:matchSrc/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdOne"><w:altChunkPr><w:matchSrc w:val="1"/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdFalse"><w:altChunkPr><w:matchSrc val="false"/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdOff"><w:altChunkPr><w:matchSrc w:val="off"/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "matchsrc-strict-values",
        strict_document(
            r#"<w:altChunk r:id="rIdTrue"><w:altChunkPr><w:matchSrc/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdOne"><w:altChunkPr><w:matchSrc w:val="1"/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdFalse"><w:altChunkPr><w:matchSrc val="false"/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdZero"><w:altChunkPr><w:matchSrc w:val="0"/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "matchsrc-legacy-values-rejected-strict",
        strict_document(
            r#"<w:altChunk r:id="rIdOn"><w:altChunkPr><w:matchSrc w:val="on"/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "altchunkpr-empty-is-valid",
        transitional_document(r#"<w:altChunk r:id="rIdEmpty"><w:altChunkPr/></w:altChunk>"#),
    );
    push_inline(
        &mut cases,
        "opaque-non-word-child",
        transitional_document(
            r#"<w:altChunk r:id="rIdOpaque"><x:payload xmlns:x="urn:opaque"><x:item><![CDATA[foreign <xml> bytes]]></x:item><!-- retained as opaque --></x:payload></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "multiple-source-order",
        transitional_document(
            r#"<w:altChunk r:id="rIdThird"/><w:altChunk r:id="rIdFirst"><w:altChunkPr><w:matchSrc w:val="0"/></w:altChunkPr></w:altChunk><w:altChunk r:id="rIdSecond"/>"#,
        ),
    );
    push_inline(
        &mut cases,
        "mce-fallback-selected",
        mce_document(
            r#"<mc:AlternateContent><mc:Choice Requires="u"><w:altChunk r:id="rIdInactive"/></mc:Choice><mc:Fallback><w:altChunk r:id="rIdFallback"/></mc:Fallback></mc:AlternateContent>"#,
        ),
    );
    push_inline(
        &mut cases,
        "mce-choice-selected",
        mce_document(
            r#"<mc:AlternateContent><mc:Choice Requires="w"><w:altChunk r:id="rIdChoice"/></mc:Choice><mc:Fallback><w:altChunk r:id="rIdInactiveFallback"/></mc:Fallback></mc:AlternateContent>"#,
        ),
    );
    push_inline(
        &mut cases,
        "mce-order-with-inactive-branches",
        mce_document(
            r#"<mc:AlternateContent><mc:Choice Requires="u"><w:altChunk r:id="rIdInactiveA"/></mc:Choice><mc:Fallback><w:altChunk r:id="rIdFallbackA"/></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="w"><w:altChunk r:id="rIdChoiceB"/></mc:Choice><mc:Fallback><w:altChunk r:id="rIdInactiveB"/></mc:Fallback></mc:AlternateContent>"#,
        ),
    );
    push_inline(
        &mut cases,
        "events-outside-anchor",
        format!(
            r#"<?xml version="1.0"?><!DOCTYPE w:document><w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>text &amp; <![CDATA[CDATA outside]]><!-- comment --><?probe value?><w:altChunk r:id="rIdEvents"/><!-- tail --></w:body></w:document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "events-outside-anchor-without-doctype",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>text &amp; <![CDATA[CDATA outside]]><!-- comment --><?probe value?><w:altChunk r:id="rIdEvents"/><!-- tail --></w:body></w:document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "comments-and-pi-inside-anchor",
        transitional_document(
            r#"<w:altChunk r:id="rIdEvents"><!-- before --><w:altChunkPr><?inside?><w:matchSrc><!-- value --></w:matchSrc></w:altChunkPr><?after?></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "whitespace-text-inside-anchor",
        transitional_document(
            "<w:altChunk r:id=\"rIdWhitespace\"> \n\t<w:altChunkPr> \n<w:matchSrc/> \n</w:altChunkPr> \n</w:altChunk>",
        ),
    );
    push_inline(
        &mut cases,
        "non-whitespace-text-inside-anchor",
        transitional_document(r#"<w:altChunk r:id="rIdText">unexpected</w:altChunk>"#),
    );
    push_inline(
        &mut cases,
        "cdata-inside-anchor",
        transitional_document(r#"<w:altChunk r:id="rIdCdata"><![CDATA[foreign]]></w:altChunk>"#),
    );
    push_inline(
        &mut cases,
        "general-reference-inside-anchor",
        transitional_document(r#"<w:altChunk r:id="rIdRef">&amp;</w:altChunk>"#),
    );
    push_inline(
        &mut cases,
        "missing-relationship-id",
        transitional_document(r#"<w:altChunk/>"#),
    );
    push_inline(
        &mut cases,
        "relationship-id-in-wrong-namespace",
        format!(
            r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:bad="urn:bad"><w:body><w:altChunk bad:id="rIdWrong"/></w:body></w:document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "duplicate-relationship-id",
        format!(
            r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}" xmlns:q="{STRICT_RELATIONSHIPS}"><w:body><w:altChunk r:id="rIdA" q:id="rIdB"/></w:body></w:document>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "empty-relationship-id",
        transitional_document(r#"<w:altChunk r:id=""/>"#),
    );
    push_inline(
        &mut cases,
        "unsafe-relationship-id",
        transitional_document(r#"<w:altChunk r:id="bad&amp;id"/>"#),
    );
    push_inline(
        &mut cases,
        "duplicate-altchunk-properties",
        transitional_document(
            r#"<w:altChunk r:id="rIdDuplicatePr"><w:altChunkPr/><w:altChunkPr/></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "duplicate-matchsrc-value",
        transitional_document(
            r#"<w:altChunk r:id="rIdDuplicateValue"><w:altChunkPr><w:matchSrc w:val="1" val="0"/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "duplicate-matchsrc-element",
        transitional_document(
            r#"<w:altChunk r:id="rIdDuplicateMatch"><w:altChunkPr><w:matchSrc/><w:matchSrc/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "invalid-matchsrc-value",
        transitional_document(
            r#"<w:altChunk r:id="rIdInvalidValue"><w:altChunkPr><w:matchSrc w:val="maybe"/></w:altChunkPr></w:altChunk>"#,
        ),
    );
    push_inline(
        &mut cases,
        "word-child-in-invalid-position",
        transitional_document(r#"<w:altChunk r:id="rIdInvalidChild"><w:unexpected/></w:altChunk>"#),
    );
    push_inline(
        &mut cases,
        "matchsrc-outside-properties",
        transitional_document(r#"<w:altChunk r:id="rIdWrongPosition"><w:matchSrc/></w:altChunk>"#),
    );
    push_inline(&mut cases, "junk-xml-text", b"junk XML text".to_vec());
    push_inline(
        &mut cases,
        "truncated-document",
        format!(
            r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body><w:altChunk r:id="rIdTruncated"/>"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "truncated-altchunk",
        format!(
            r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body><w:altChunk r:id="rIdTruncated">"#
        )
        .into_bytes(),
    );
    push_inline(
        &mut cases,
        "mismatched-tags",
        format!(
            r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body><w:altChunk r:id="rIdMismatch"></w:body></w:document>"#
        )
        .into_bytes(),
    );
    push_inline(&mut cases, "invalid-utf8", {
        let mut bytes = transitional_document(r#"<w:altChunk r:id="rIdUtf8"/>"#);
        bytes.insert(20, 0xff);
        bytes
    });
    push_inline(&mut cases, "empty-input", Vec::new());
    push_inline(
        &mut cases,
        "multiple-root-elements",
        [
            transitional_document(r#"<w:altChunk r:id="rIdFirstRoot"/>"#),
            transitional_document(r#"<w:altChunk r:id="rIdSecondRoot"/>"#),
        ]
        .concat(),
    );
    push_inline(&mut cases, "utf8-bom", {
        let mut bytes = b"\xef\xbb\xbf".to_vec();
        bytes.extend(transitional_document(r#"<w:altChunk r:id="rIdBom"/>"#));
        bytes
    });
    push_inline(
        &mut cases,
        "xml-depth-boundary",
        nested_document(MAX_XML_DEPTH.saturating_sub(3)),
    );
    push_inline(
        &mut cases,
        "xml-depth-overflow",
        nested_document(MAX_XML_DEPTH.saturating_sub(2)),
    );
    push_inline(
        &mut cases,
        "anchor-count-boundary",
        anchors_document(MAX_CHUNKS),
    );
    push_inline(
        &mut cases,
        "anchor-count-overflow",
        anchors_document(MAX_CHUNKS + 1),
    );
    push_inline(
        &mut cases,
        "xml-byte-limit-overflow",
        vec![b'x'; MAX_XML_BYTES + 1],
    );

    push_fixture(
        &mut cases,
        repository_root,
        NUMBERED_LIST_FIXTURE,
        "fixture-numbered-list-main",
    )?;
    push_fixture(
        &mut cases,
        repository_root,
        ALT_CHUNK_FIXTURE,
        "fixture-alt-chunk-main",
    )?;

    Ok(cases)
}

fn push_inline(cases: &mut Vec<CaseInput>, name: &str, input: Vec<u8>) {
    cases.push(CaseInput {
        name: name.to_owned(),
        source: format!("inline:{name}"),
        source_sha256: sha256_hex(&input),
        input,
    });
}

fn push_fixture(
    cases: &mut Vec<CaseInput>,
    repository_root: &Path,
    relative_path: &str,
    name: &str,
) -> Result<(), Box<dyn Error>> {
    let path = repository_root.join(relative_path);
    let source_bytes = fs::read(&path)?;
    let source_sha256 = sha256_hex(&source_bytes);
    let input = ArchiveReader::new(&source_bytes)?.read(MAIN_MEMBER)?;
    cases.push(CaseInput {
        name: name.to_owned(),
        source: format!("fixture:{relative_path}"),
        source_sha256,
        input,
    });
    Ok(())
}

fn transitional_document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn strict_document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{STRICT_WORD}" xmlns:r="{STRICT_RELATIONSHIPS}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn mce_document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}" xmlns:mc="{MCE_NAMESPACE}" xmlns:u="urn:unsupported"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn nested_document(levels: usize) -> Vec<u8> {
    let mut xml = format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>"#
    );
    for index in 0..levels {
        let _ = write!(xml, "<n{index}>");
    }
    xml.push_str(r#"<w:altChunk r:id="rIdDepth"/>"#);
    for index in (0..levels).rev() {
        let _ = write!(xml, "</n{index}>");
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn anchors_document(count: usize) -> Vec<u8> {
    let mut xml = format!(
        r#"<w:document xmlns:w="{TRANSITIONAL_WORD}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS}"><w:body>"#
    );
    for index in 0..count {
        let _ = write!(xml, r#"<w:altChunk r:id="rId{index}"/>"#);
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}
