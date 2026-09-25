//! Record 0764: hostile-attribute timings and a before/after differential.
//!
//! The same source builds against the base and the changed library, so it
//! uses only APIs both expose.
//!
//! * `adversarial --case NAME --samples N --warmup W --json PATH` builds one
//!   hostile input once, then times `W + N` runs of the operation it attacks
//!   and records each run's wall time and outcome.
//! * `differential --json PATH ROOT...` runs the MCE processor, the MCE
//!   stream, the publication audits and the DOCX style reader over every XML
//!   member of every OOXML package under the roots, and over generated MCE
//!   documents, and records a digest of each outcome, so two builds can be
//!   compared member by member. `--aliasing N` (record 0771) adds `N`
//!   generated documents whose prefixes alias one another and the fixed
//!   namespaces, through both MCE processors.

#![forbid(unsafe_code)]

use std::collections::BTreeMap;
use std::convert::Infallible;
use std::error::Error;
use std::fmt::Write as _;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::time::Instant;

use litchi_ooxml_common::mce::{
    Capabilities, Limits as MceLimits, SemanticEvent, StreamLimits, process_markup_compatibility,
    process_markup_compatibility_stream, process_markup_compatibility_stream_with_observers,
};
use litchi_opc::{BlobPart, PackURI};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use xml_minifier::audit;

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("adversarial") => adversarial(&args[1..]),
        Some("differential") => differential(&args[1..]),
        _ => Err("usage: xml_attribute_bounds adversarial|differential ...".into()),
    }
}

fn option<'a>(args: &'a [String], name: &str) -> Option<&'a str> {
    args.iter()
        .position(|arg| arg == name)
        .and_then(|index| args.get(index + 1))
        .map(String::as_str)
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

// --------------------------------------------------------------------------
// Adversarial cases
// --------------------------------------------------------------------------

/// One hostile input and the operation it attacks, returning an outcome
/// string that is identical across runs.
struct Case {
    input: Vec<u8>,
    run: fn(&[u8]) -> String,
}

fn adversarial(args: &[String]) -> Result<(), Box<dyn Error>> {
    let name = option(args, "--case").ok_or("--case is required")?;
    let samples: usize = option(args, "--samples").unwrap_or("15").parse()?;
    let warmup: usize = option(args, "--warmup").unwrap_or("3").parse()?;
    let json = option(args, "--json").ok_or("--json is required")?;
    let case = build_case(name)?;
    let mut outcomes = Vec::new();
    let mut elapsed = Vec::new();
    for index in 0..warmup + samples {
        let start = Instant::now();
        let outcome = (case.run)(&case.input);
        let nanos = start.elapsed().as_nanos();
        if index >= warmup {
            elapsed.push(u64::try_from(nanos).unwrap_or(u64::MAX));
            outcomes.push(outcome);
        }
    }
    outcomes.dedup();
    let mut sorted = elapsed.clone();
    sorted.sort_unstable();
    let report = serde_json::json!({
        "case": name,
        "input_bytes": case.input.len(),
        "input_sha256": sha256_hex(&case.input),
        "samples": samples,
        "warmup": warmup,
        "elapsed_ns": elapsed,
        "p50_ns": sorted.get(sorted.len() / 2),
        "outcomes": outcomes,
    });
    std::fs::write(json, serde_json::to_vec_pretty(&report)?)?;
    Ok(())
}

fn build_case(name: &str) -> Result<Case, Box<dyn Error>> {
    Ok(match name {
        // One element with 20,000 namespace declarations in a part that names
        // the MCE namespace.
        "mce_declaration_flood" => Case {
            input: declaration_flood(20_000).into_bytes(),
            run: run_mce_codec,
        },
        // 32 nested elements re-declaring 1,000 prefixes each, around 50
        // elements of 1,000 attributes each in a namespace bound at the root.
        "mce_shadowed_chain" => Case {
            input: shadowed_chain(32, 1_000, 1_000, 50).into_bytes(),
            run: run_mce_codec,
        },
        // 2,000 sibling elements that each declare the same 1,000 prefixes,
        // which both builds admit.
        "mce_declarations_admitted" => Case {
            input: stream_declarations(2_000, 1_000).into_bytes(),
            run: run_mce_codec,
        },
        // Ten dropped wrappers re-declaring 100 prefixes each, around 2,000
        // emitted children that re-declare the innermost bindings.
        "mce_hoisting" => Case {
            input: hoisting(5, 100, 2_000).into_bytes(),
            run: run_mce_codec,
        },
        // 500 elements with 1,000 namespace declarations each through the
        // streaming processor.
        "mce_stream_declarations" => Case {
            input: stream_declarations(500, 1_000).into_bytes(),
            run: run_mce_stream,
        },
        // A DOCX style element with 40,000 distinct attributes followed by
        // 40,000 repeats of the last one, read by the lenient style reader.
        "docx_styles_duplicates" => Case {
            input: styles_duplicates(40_000, 40_000).into_bytes(),
            run: run_docx_styles,
        },
        // One tag with 200,000 attributes through the source audit.
        "audit_attribute_flood" => Case {
            input: attribute_flood(200_000).into_bytes(),
            run: run_audit_source,
        },
        // 1,000 attributes in a namespace whose URI is 4 MiB long, declared on
        // the root: plain, ignorable, and ignorable with a preservation
        // wildcard; and 4 x 1,000 such attributes under a 1 MiB URI (the
        // stream's name limit) through the MCE stream.
        "mce_long_uri_plain" => Case {
            input: long_uri(4 << 20, "", 1, 1_000).into_bytes(),
            run: run_mce_codec,
        },
        "mce_long_uri_ignorable" => Case {
            input: long_uri(4 << 20, r#" mc:Ignorable="z""#, 1, 1_000).into_bytes(),
            run: run_mce_codec,
        },
        "mce_long_uri_preserved" => Case {
            input: long_uri(
                4 << 20,
                r#" mc:Ignorable="z" mc:PreserveAttributes="z:*""#,
                1,
                1_000,
            )
            .into_bytes(),
            run: run_mce_codec,
        },
        "mce_stream_long_uri" => Case {
            input: long_uri((1 << 20) - 64, "", 4, 1_000).into_bytes(),
            run: run_mce_stream_count,
        },
        "mce_stream_long_uri_ignorable" => Case {
            input: long_uri((1 << 20) - 64, r#" mc:Ignorable="z""#, 4, 1_000).into_bytes(),
            run: run_mce_stream_count,
        },
        // The review's probe inputs, as its probe built them: attributes in a
        // namespace with a long URI beside a short ignorable namespace, with
        // the baseline capabilities; 1,000 under a 4 MiB URI through the
        // processor, 4 x 1,000 under a 1,040,000-byte URI through the stream
        // with an observer that only counts events.
        "mce_review_long_uri" => Case {
            input: review_long_uri(4 << 20, 1_000).into_bytes(),
            run: run_mce_codec_baseline,
        },
        "mce_stream_review_long_uri" => Case {
            input: review_long_uri(1_040_000, 4_000).into_bytes(),
            run: run_mce_stream_review,
        },
        // Record 0771: element names and directive tokens in a namespace
        // whose URI is at the stream's name limit: 4,000 empty elements in
        // `z`, emitted or skipped as ignorable, and 4,000 elements that each
        // make `z` ignorable and preserve its elements.
        "mce_stream_long_uri_elements" => Case {
            input: long_uri_elements((1 << 20) - 64, "", 4_000).into_bytes(),
            run: run_mce_stream_count,
        },
        "mce_stream_long_uri_skipped" => Case {
            input: long_uri_elements((1 << 20) - 64, r#" mc:Ignorable="z""#, 4_000).into_bytes(),
            run: run_mce_stream_count,
        },
        "mce_stream_long_uri_tokens" => Case {
            input: long_uri_tokens((1 << 20) - 64, 4_000).into_bytes(),
            run: run_mce_stream_count,
        },
        // Benign controls: real producer parts that name the MCE namespace,
        // read from the repository's fixtures (run from the worktree root).
        "mce_benign_worksheet" => Case {
            input: fixture_member(BENIGN_WORKBOOK, BENIGN_WORKSHEET)?,
            run: run_mce_codec_baseline,
        },
        "mce_benign_document" => Case {
            input: fixture_member(BENIGN_DOCUMENT, "word/document.xml")?,
            run: run_mce_codec_baseline,
        },
        "mce_stream_benign_worksheet" => Case {
            input: fixture_member(BENIGN_WORKBOOK, BENIGN_WORKSHEET)?,
            run: run_mce_stream_baseline,
        },
        // Record 0771: the stream on the real document part, and on both
        // parts with observers that only count, so the timing is the
        // stream's own work.
        "mce_stream_benign_document" => Case {
            input: fixture_member(BENIGN_DOCUMENT, "word/document.xml")?,
            run: run_mce_stream_baseline,
        },
        "mce_stream_count_worksheet" => Case {
            input: fixture_member(BENIGN_WORKBOOK, BENIGN_WORKSHEET)?,
            run: run_mce_stream_count_baseline,
        },
        "mce_stream_count_document" => Case {
            input: fixture_member(BENIGN_DOCUMENT, "word/document.xml")?,
            run: run_mce_stream_count_baseline,
        },
        "audit_benign_worksheet" => Case {
            input: fixture_member(BENIGN_WORKBOOK, BENIGN_WORKSHEET)?,
            run: run_audit_source,
        },
        "docx_styles_benign" => Case {
            input: fixture_member(BENIGN_STYLES, "word/styles.xml")?,
            run: run_docx_styles,
        },
        _ => return Err(format!("unknown case {name}").into()),
    })
}

const BENIGN_WORKBOOK: &str = "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx";
const BENIGN_WORKSHEET: &str = "xl/worksheets/sheet3.xml";
const BENIGN_DOCUMENT: &str = "test-data/ooxml/docx/drawing.docx";
const BENIGN_STYLES: &str =
    "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx";

fn fixture_member(package: &str, member: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let bytes = std::fs::read(package)?;
    let archive = ArchiveReader::new(&bytes)?;
    Ok(archive.read(member)?)
}

/// A root declaring `z` with a URI of `uri_bytes` bytes and `directives`,
/// around `elements` elements of `attributes` attributes in `z` each.
fn long_uri(uri_bytes: usize, directives: &str, elements: usize, attributes: usize) -> String {
    let uri = format!("urn:{}", "u".repeat(uri_bytes.saturating_sub(4)));
    let mut element = String::from("<e");
    for index in 0..attributes {
        let _ = write!(element, r#" z:a{index}="""#);
    }
    element.push_str("/>");
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{uri}"{directives}>{}</r>"#,
        element.repeat(elements)
    )
}

/// A root declaring `z` with a URI of `uri_bytes` bytes and `directives`,
/// around `elements` empty elements in `z`.
fn long_uri_elements(uri_bytes: usize, directives: &str, elements: usize) -> String {
    let uri = format!("urn:{}", "u".repeat(uri_bytes.saturating_sub(4)));
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{uri}"{directives}>{}</r>"#,
        "<z:e/>".repeat(elements)
    )
}

/// A root declaring `z` with a URI of `uri_bytes` bytes, around `elements`
/// empty elements that each make `z` ignorable and preserve every element in
/// it: one `Ignorable` and one `PreserveElements` token per element.
fn long_uri_tokens(uri_bytes: usize, elements: usize) -> String {
    let uri = format!("urn:{}", "u".repeat(uri_bytes.saturating_sub(4)));
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{uri}">{}</r>"#,
        r#"<e mc:Ignorable="z" mc:PreserveElements="z:*"/>"#.repeat(elements)
    )
}

/// The review's long-URI probe input: `p` bound to `urn:` followed by
/// `u_count` `u`s, beside an ignorable `x`, around elements of at most 1,000
/// attributes in `p` until `attributes` are written.
fn review_long_uri(u_count: usize, attributes: usize) -> String {
    let uri = format!("urn:{}", "u".repeat(u_count));
    let mut xml =
        format!(r#"<r xmlns:mc="{MC}" xmlns:x="urn:x" xmlns:p="{uri}" mc:Ignorable="x">"#);
    let per = attributes.min(1_000);
    let mut written = 0;
    while written < attributes {
        xml.push_str("<e");
        for index in 0..per {
            let _ = write!(xml, r#" p:a{index}="""#);
        }
        xml.push_str("/>");
        written += per;
    }
    xml.push_str("</r>");
    xml
}

fn declaration_flood(count: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{MC}"><t"#);
    for index in 0..count {
        let _ = write!(xml, r#" xmlns:p{index}="urn:{index}""#);
    }
    xml.push_str("/></r>");
    xml
}

fn shadowed_chain(depth: usize, width: usize, attributes: usize, elements: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:z="urn:z">"#);
    for level in 0..depth {
        xml.push_str("<s");
        for prefix in 0..width {
            let _ = write!(xml, r#" xmlns:q{prefix}="urn:{level}:{prefix}""#);
        }
        xml.push('>');
    }
    let mut element = String::from("<z:e");
    for index in 0..attributes {
        let _ = write!(element, r#" z:a{index}="""#);
    }
    element.push_str("/>");
    xml.push_str(&element.repeat(elements));
    for _ in 0..depth {
        xml.push_str("</s>");
    }
    xml.push_str("</r>");
    xml
}

fn hoisting(levels: usize, width: usize, children: usize) -> String {
    let declarations = |layer: usize| -> String {
        let mut text = String::new();
        for prefix in 0..width {
            let _ = write!(text, r#" xmlns:p{prefix}="urn:{layer}:{prefix}""#);
        }
        text
    };
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:i="urn:i" xmlns:w="urn:w">"#);
    for level in 0..levels {
        let _ = write!(
            xml,
            r#"<mc:AlternateContent{}><mc:Choice Requires="i"/><mc:Fallback{}>"#,
            declarations(2 * level),
            declarations(2 * level + 1)
        );
    }
    for _ in 0..children {
        xml.push_str("<w:x/>");
    }
    for _ in 0..levels {
        xml.push_str("</mc:Fallback></mc:AlternateContent>");
    }
    xml.push_str("</r>");
    xml
}

fn stream_declarations(elements: usize, declarations: usize) -> String {
    let mut element = String::from("<e");
    for index in 0..declarations {
        let _ = write!(element, r#" xmlns:p{index}="urn:{index}""#);
    }
    element.push_str("/>");
    format!(r#"<r xmlns:mc="{MC}">{}</r>"#, element.repeat(elements))
}

fn styles_duplicates(distinct: usize, repeats: usize) -> String {
    let mut xml = format!(r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="S""#);
    for index in 0..distinct {
        let _ = write!(xml, r#" a{index:06}="""#);
    }
    let last = format!(r#" a{:06}="""#, distinct - 1);
    for _ in 0..repeats {
        xml.push_str(&last);
    }
    xml.push_str(r#"><w:name w:val="Style"/></w:style></w:styles>"#);
    xml
}

fn attribute_flood(count: usize) -> String {
    let mut xml = String::from("<root");
    for index in 0..count {
        let _ = write!(xml, r#" a{index}="""#);
    }
    xml.push_str("/>");
    xml
}

fn run_mce_codec(input: &[u8]) -> String {
    match process_markup_compatibility(input, &Capabilities::new(), &MceLimits::default()) {
        Ok(output) => format!("ok:{}", sha256_hex(output.xml.as_ref())),
        Err(error) => format!("err:{error}"),
    }
}

fn run_mce_stream(input: &[u8]) -> String {
    stream_digest(input, &Capabilities::new())
}

/// The MCE stream with observers that count events and attributes without
/// reading any name, so the timing is the stream's own work.
fn run_mce_stream_count(input: &[u8]) -> String {
    stream_count(input, &Capabilities::new())
}

/// [`run_mce_stream_count`] with the OOXML baseline capabilities.
fn run_mce_stream_count_baseline(input: &[u8]) -> String {
    stream_count(input, &Capabilities::ooxml_baseline())
}

fn stream_count(input: &[u8], capabilities: &Capabilities) -> String {
    let mut raw = (0usize, 0usize);
    let mut semantic = (0usize, 0usize);
    let mut cursor = Cursor::new(input);
    let result = process_markup_compatibility_stream_with_observers(
        &mut cursor,
        capabilities,
        &StreamLimits::default(),
        |element| {
            raw.0 += 1;
            raw.1 += element.attrs().len();
            Ok::<(), Infallible>(())
        },
        |event| {
            semantic.0 += 1;
            if let SemanticEvent::Start(element) | SemanticEvent::Empty(element) = &event {
                semantic.1 += element.attrs().len();
            }
            Ok::<(), Infallible>(())
        },
    );
    match result {
        Ok(report) => format!("ok:{raw:?}:{semantic:?}:{report:?}"),
        Err(error) => format!("err:{error}:{raw:?}:{semantic:?}"),
    }
}

/// The MCE stream as the review's probe ran it: baseline capabilities and an
/// active observer that only counts events.
fn run_mce_stream_review(input: &[u8]) -> String {
    let mut events = 0usize;
    let mut cursor = Cursor::new(input);
    let result = process_markup_compatibility_stream(
        &mut cursor,
        &Capabilities::ooxml_baseline(),
        &StreamLimits::default(),
        |_event| {
            events += 1;
            Ok::<(), Infallible>(())
        },
    );
    match result {
        Ok(report) => format!("ok:{events}:{report:?}"),
        Err(error) => format!("err:{error}:{events}"),
    }
}

fn run_mce_codec_baseline(input: &[u8]) -> String {
    match process_markup_compatibility(
        input,
        &Capabilities::ooxml_baseline(),
        &MceLimits::default(),
    ) {
        Ok(output) => format!("ok:{}", sha256_hex(output.xml.as_ref())),
        Err(error) => format!("err:{error}"),
    }
}

fn run_mce_stream_baseline(input: &[u8]) -> String {
    stream_digest(input, &Capabilities::ooxml_baseline())
}

fn run_docx_styles(input: &[u8]) -> String {
    styles_digest(input)
}

fn run_audit_source(input: &[u8]) -> String {
    match audit::verify_source(input, audit::Limits::default()) {
        Ok(report) => format!("ok:{report:?}"),
        Err(error) => format!("err:{error}"),
    }
}

// --------------------------------------------------------------------------
// Differential
// --------------------------------------------------------------------------

fn stream_digest(input: &[u8], capabilities: &Capabilities) -> String {
    let mut hasher = Sha256::new();
    let mut raw_hasher = Sha256::new();
    let mut cursor = Cursor::new(input);
    let result = process_markup_compatibility_stream_with_observers(
        &mut cursor,
        capabilities,
        &StreamLimits::default(),
        |element| {
            let mut line = format!(
                "{:?}|{}",
                element.kind,
                String::from_utf8_lossy(element.name())
            );
            let _ = write!(
                line,
                "|{}|{}",
                element.expanded_name.namespace, element.expanded_name.local_name
            );
            for attribute in element.attrs() {
                let _ = write!(
                    line,
                    "|{}={}|{}:{}",
                    String::from_utf8_lossy(attribute.name()),
                    attribute.decoded(),
                    attribute.expanded_name.namespace,
                    attribute.expanded_name.local_name
                );
            }
            raw_hasher.update(line.as_bytes());
            raw_hasher.update(b"\n");
            Ok::<(), Infallible>(())
        },
        |event| {
            hasher.update(event_signature(&event).as_bytes());
            hasher.update(b"\n");
            Ok::<(), Infallible>(())
        },
    );
    let semantic = sha256_hex(&hasher.finalize());
    let raw = sha256_hex(&raw_hasher.finalize());
    match result {
        Ok(report) => format!("ok:{semantic}:{raw}:{report:?}"),
        Err(error) => format!("err:{error}:{semantic}:{raw}"),
    }
}

fn event_signature(event: &SemanticEvent<'_>) -> String {
    match event {
        SemanticEvent::Start(element) | SemanticEvent::Empty(element) => {
            let kind = if matches!(event, SemanticEvent::Start(_)) {
                "start"
            } else {
                "empty"
            };
            let mut line = format!(
                "{kind}|{}|{}|{}",
                String::from_utf8_lossy(element.name()),
                element.expanded_name.namespace,
                element.expanded_name.local_name
            );
            for attribute in element.attrs() {
                let _ = write!(
                    line,
                    "|{}={}|{}:{}",
                    String::from_utf8_lossy(attribute.name()),
                    attribute.value(),
                    attribute.expanded_name.namespace,
                    attribute.expanded_name.local_name
                );
            }
            line
        },
        SemanticEvent::End(element) => format!(
            "end|{}|{}|{}",
            String::from_utf8_lossy(element.name()),
            element.expanded_name.namespace,
            element.expanded_name.local_name
        ),
        SemanticEvent::Text(text) => format!("text|{}", text.text()),
        SemanticEvent::CData(text) => format!("cdata|{}", text.text()),
        SemanticEvent::Comment(text) => format!("comment|{}", text.text()),
        SemanticEvent::Decl(decl) => format!("decl|{}", String::from_utf8_lossy(decl.raw.as_ref())),
        SemanticEvent::GeneralRef(reference) => {
            format!("ref|{}", String::from_utf8_lossy(reference.name.as_ref()))
        },
        _ => "other".to_owned(),
    }
}

fn styles_digest(input: &[u8]) -> String {
    let Ok(partname) = PackURI::new("/word/styles.xml") else {
        return "err:partname".to_owned();
    };
    let part = BlobPart::new(partname, "application/xml".to_owned(), input.to_vec());
    let mut styles = litchi_docx::styles::Styles::from_part(&part);
    match styles.iter() {
        Ok(iter) => {
            let mut hasher = Sha256::new();
            let mut count = 0usize;
            for style in iter {
                count += 1;
                let line = format!(
                    "{}|{:?}|{:?}|{}|{:?}|{:?}|{}|{}|{:?}",
                    style.style_id(),
                    style.name(),
                    style.style_type(),
                    style.is_default(),
                    style.based_on(),
                    style.priority(),
                    style.is_custom(),
                    style.is_hidden(),
                    style.numbering(),
                );
                hasher.update(line.as_bytes());
                hasher.update(b"\n");
            }
            format!("ok:{count}:{}", sha256_hex(&hasher.finalize()))
        },
        Err(error) => format!("err:{error}"),
    }
}

fn audit_digest(input: &[u8]) -> [String; 3] {
    let limits = audit::Limits::default();
    let render = |result: Result<audit::Report, audit::Error>| match result {
        Ok(report) => format!("ok:{report:?}"),
        Err(error) => format!("err:{error}"),
    };
    [
        render(audit::verify_source(input, limits)),
        render(audit::verify_authored(input, limits)),
        render(audit::verify(input, limits)),
    ]
}

fn differential(args: &[String]) -> Result<(), Box<dyn Error>> {
    let json = option(args, "--json").ok_or("--json is required")?;
    let generated: usize = option(args, "--generated").unwrap_or("20000").parse()?;
    let mut roots = Vec::new();
    let mut skip = false;
    for arg in args {
        if skip {
            skip = false;
            continue;
        }
        if arg.starts_with("--") {
            skip = true;
            continue;
        }
        roots.push(PathBuf::from(arg));
    }
    let mut results = BTreeMap::new();
    let baseline = Capabilities::ooxml_baseline();
    let bare = Capabilities::new();
    let mut packages = Vec::new();
    for root in &roots {
        collect_packages(root, &mut packages)?;
    }
    packages.sort();
    let mut members = 0usize;
    for path in &packages {
        let Ok(bytes) = std::fs::read(path) else {
            continue;
        };
        let Ok(archive) = ArchiveReader::new(&bytes) else {
            continue;
        };
        let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        for name in names {
            let lower = name.to_ascii_lowercase();
            if !(lower.ends_with(".xml") || lower.ends_with(".rels") || lower.ends_with(".vml")) {
                continue;
            }
            let Ok(member) = archive.read(&name) else {
                continue;
            };
            members += 1;
            let key = format!("{}::{name}", path.display());
            let mce = |capabilities: &Capabilities| match process_markup_compatibility(
                &member,
                capabilities,
                &MceLimits::default(),
            ) {
                Ok(output) => format!("ok:{}:{:?}", sha256_hex(output.xml.as_ref()), output.report),
                Err(error) => format!("err:{error}"),
            };
            results.insert(format!("{key}::mce_baseline"), mce(&baseline));
            results.insert(format!("{key}::mce_bare"), mce(&bare));
            results.insert(
                format!("{key}::stream_baseline"),
                stream_digest(&member, &baseline),
            );
            let [source, authored, compact] = audit_digest(&member);
            results.insert(format!("{key}::audit_source"), source);
            results.insert(format!("{key}::audit_authored"), authored);
            results.insert(format!("{key}::audit_compact"), compact);
            if lower.ends_with("styles.xml") {
                results.insert(format!("{key}::docx_styles"), styles_digest(&member));
            }
        }
    }
    for index in 0..generated {
        let document = generated_document(index as u64);
        let key = format!("generated::{index:06}");
        for (label, capabilities) in [("bare", &bare), ("understood", &understanding_w())] {
            let codec = match process_markup_compatibility(
                document.as_bytes(),
                capabilities,
                &MceLimits::default(),
            ) {
                Ok(output) => format!("ok:{}:{:?}", sha256_hex(output.xml.as_ref()), output.report),
                Err(error) => format!("err:{error}"),
            };
            results.insert(format!("{key}::mce_{label}"), codec);
            results.insert(
                format!("{key}::stream_{label}"),
                stream_digest(document.as_bytes(), capabilities),
            );
        }
    }
    // Record 0771: documents whose prefixes alias one another and the fixed
    // namespaces, under three capability sets.
    let aliasing: usize = option(args, "--aliasing").unwrap_or("0").parse()?;
    let understood_a = understanding("urn:a");
    for index in 0..aliasing {
        let document = aliasing_document(index as u64);
        let key = format!("aliasing::{index:06}");
        for (label, capabilities) in [
            ("bare", &bare),
            ("baseline", &baseline),
            ("understood", &understood_a),
        ] {
            let codec = match process_markup_compatibility(
                document.as_bytes(),
                capabilities,
                &MceLimits::default(),
            ) {
                Ok(output) => format!("ok:{}:{:?}", sha256_hex(output.xml.as_ref()), output.report),
                Err(error) => format!("err:{error}"),
            };
            results.insert(format!("{key}::mce_{label}"), codec);
            results.insert(
                format!("{key}::stream_{label}"),
                stream_digest(document.as_bytes(), capabilities),
            );
        }
    }
    let summary = serde_json::json!({
        "packages": packages.len(),
        "members": members,
        "generated": generated,
        "aliasing": aliasing,
        "results": results,
    });
    std::fs::write(json, serde_json::to_vec(&summary)?)?;
    Ok(())
}

fn understanding_w() -> Capabilities {
    understanding("urn:w")
}

fn understanding(namespace: &str) -> Capabilities {
    let mut capabilities = Capabilities::new();
    capabilities.understand_namespace(namespace);
    capabilities
}

fn collect_packages(root: &Path, packages: &mut Vec<PathBuf>) -> Result<(), Box<dyn Error>> {
    let metadata = std::fs::symlink_metadata(root)?;
    if metadata.is_dir() {
        let mut entries: Vec<_> = std::fs::read_dir(root)?.filter_map(Result::ok).collect();
        entries.sort_by_key(std::fs::DirEntry::path);
        for entry in entries {
            collect_packages(&entry.path(), packages)?;
        }
    } else if metadata.is_file() {
        let Ok(bytes) = std::fs::read(root) else {
            return Ok(());
        };
        if bytes.starts_with(b"PK\x03\x04")
            && ArchiveReader::new(&bytes)
                .is_ok_and(|archive| archive.contains("[Content_Types].xml"))
        {
            packages.push(root.to_owned());
        }
    }
    Ok(())
}

/// A deterministic pseudo-random MCE document exercising namespace scopes,
/// shadowing, default-namespace resets, alternate content, ignorable,
/// preserved and unwrapped elements, and directive targets.
fn generated_document(seed: u64) -> String {
    let mut random = Xorshift(seed.wrapping_mul(0x9e37_79b9_7f4a_7c15) | 1);
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:w="urn:w" xmlns:i="urn:i" xmlns:j="urn:j""#);
    match random.below(4) {
        0 => xml.push_str(r#" mc:Ignorable="i j" mc:ProcessContent="i:u j:*" mc:PreserveElements="i:keep" mc:PreserveAttributes="i:*""#),
        1 => xml.push_str(r#" mc:Ignorable="i" mc:PreserveAttributes="i:k""#),
        2 => xml.push_str(r#" mc:Ignorable="j""#),
        _ => {},
    }
    xml.push('>');
    children(&mut random, &mut xml, 0);
    xml.push_str("</r>");
    xml
}

struct Xorshift(u64);

impl Xorshift {
    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value << 13;
        value ^= value >> 7;
        value ^= value << 17;
        self.0 = value;
        value
    }

    fn below(&mut self, bound: u64) -> u64 {
        self.next() % bound
    }
}

const PREFIXES: [&str; 6] = ["p0", "p1", "p2", "i", "j", "w"];

fn declarations(random: &mut Xorshift, xml: &mut String, depth: usize) {
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let prefix = PREFIXES[random.below(PREFIXES.len() as u64) as usize];
        if used.contains(&prefix) {
            continue;
        }
        used.push(prefix);
        let namespace = match random.below(4) {
            0 => "urn:i".to_owned(),
            1 => "urn:w".to_owned(),
            2 => "urn:j".to_owned(),
            _ => format!("urn:{prefix}:{depth}"),
        };
        let _ = write!(xml, r#" xmlns:{prefix}="{namespace}""#);
    }
    match random.below(8) {
        0 => xml.push_str(r#" xmlns="urn:d""#),
        1 => xml.push_str(r#" xmlns="""#),
        _ => {},
    }
}

fn attributes(random: &mut Xorshift, xml: &mut String) {
    let names = ["w:a", "i:b", "i:k", "j:c", "d", "p0:e", "p1:f", "xml:space"];
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let name = names[random.below(names.len() as u64) as usize];
        if used.contains(&name) {
            continue;
        }
        used.push(name);
        let value = if name == "xml:space" {
            "preserve"
        } else {
            "v&amp;1"
        };
        let _ = write!(xml, r#" {name}="{value}""#);
    }
}

fn children(random: &mut Xorshift, xml: &mut String, depth: usize) {
    if depth >= 6 {
        return;
    }
    let count = random.below(4);
    for _ in 0..count {
        match random.below(9) {
            0 | 1 => {
                xml.push_str("<w:x");
                declarations(random, xml, depth);
                attributes(random, xml);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</w:x>");
            },
            2 => {
                xml.push_str("<mc:AlternateContent");
                declarations(random, xml, depth);
                xml.push('>');
                let requires = ["w", "i", "j", "p0"][random.below(4) as usize];
                let _ = write!(xml, r#"<mc:Choice Requires="{requires}""#);
                declarations(random, xml, depth);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</mc:Choice>");
                if random.below(2) == 0 {
                    xml.push_str("<mc:Fallback");
                    declarations(random, xml, depth);
                    xml.push('>');
                    children(random, xml, depth + 1);
                    xml.push_str("</mc:Fallback>");
                }
                xml.push_str("</mc:AlternateContent>");
            },
            3 => {
                xml.push_str("<i:u");
                if random.below(3) == 0 {
                    declarations(random, xml, depth);
                }
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:u>");
            },
            4 => {
                xml.push_str("<i:keep");
                attributes(random, xml);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:keep>");
            },
            5 => {
                xml.push_str("<i:other");
                declarations(random, xml, depth);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:other>");
            },
            6 => {
                xml.push_str("<j:y");
                attributes(random, xml);
                xml.push_str("/>");
            },
            7 => xml.push_str("text &lt; t"),
            _ => {
                xml.push_str("<w:e");
                declarations(random, xml, depth);
                attributes(random, xml);
                xml.push_str("/>");
            },
        }
    }
}

// --------------------------------------------------------------------------
// Record 0771: aliased namespaces
// --------------------------------------------------------------------------

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// A deterministic pseudo-random document whose prefixes alias one another
/// and the fixed namespaces: `b` may be bound to `a`'s URI, `x` to the `xml`
/// namespace, `n` to the `xmlns` namespace and `m` to the markup
/// compatibility namespace, at the root or again below it, beside
/// default-namespace resets. Its attributes, elements and directive tokens
/// name those namespaces through either prefix, so two names or targets that
/// differ only in their prefix are the same name.
fn aliasing_document(seed: u64) -> String {
    let mut random = Xorshift(seed.wrapping_mul(0xd1b5_4a32_d192_ed03) | 1);
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:a="urn:a" xmlns:w="urn:w""#);
    aliased_bindings(&mut random, &mut xml, true);
    aliased_directives(&mut random, &mut xml);
    aliased_attributes(&mut random, &mut xml);
    xml.push('>');
    aliased_children(&mut random, &mut xml, 0);
    xml.push_str("</r>");
    xml
}

/// Bind `b`, `x`, `n` and `m`, each to an alias or to a URI of its own: all
/// four at the root, some of them on a descendant.
fn aliased_bindings(random: &mut Xorshift, xml: &mut String, root: bool) {
    let choices: [(&str, &str, &str); 4] = [
        ("b", "urn:a", "urn:b"),
        ("x", XML_NAMESPACE, "urn:x"),
        ("n", XMLNS_NAMESPACE, "urn:n"),
        ("m", MC, "urn:m"),
    ];
    for (prefix, alias, own) in choices {
        if !root && random.below(3) != 0 {
            continue;
        }
        let namespace = if random.below(2) == 0 { alias } else { own };
        let _ = write!(xml, r#" xmlns:{prefix}="{namespace}""#);
    }
    if !root {
        match random.below(6) {
            0 => xml.push_str(r#" xmlns="""#),
            1 => xml.push_str(r#" xmlns="urn:a""#),
            2 => {
                let _ = write!(xml, r#" xmlns="{XML_NAMESPACE}""#);
            },
            _ => {},
        }
    }
}

/// Between one and `most` distinct items of `pool`.
fn pick<'a>(random: &mut Xorshift, pool: &[&'a str], most: u64) -> Vec<&'a str> {
    let count = 1 + random.below(most);
    let mut chosen: Vec<&str> = Vec::new();
    for _ in 0..count {
        let item = pool[random.below(pool.len() as u64) as usize];
        if !chosen.contains(&item) {
            chosen.push(item);
        }
    }
    chosen
}

/// The prefix the root may bind to the same namespace as `prefix`.
fn partner(prefix: &str) -> &str {
    match prefix {
        "a" => "b",
        "b" => "a",
        "x" => "xml",
        "xml" => "x",
        other => other,
    }
}

/// Compatibility directives whose targets name an ignorable prefix, its
/// alias, or now and then any prefix.
fn aliased_directives(random: &mut Xorshift, xml: &mut String) {
    const PREFIXES: [&str; 6] = ["a", "b", "x", "n", "xml", "w"];
    let directive = if random.below(3) == 0 { "m" } else { "mc" };
    if random.below(4) != 0 {
        let ignorable = pick(random, &PREFIXES, 3);
        let _ = write!(xml, r#" {directive}:Ignorable="{}""#, ignorable.join(" "));
        let target = |random: &mut Xorshift, locals: &[&str]| -> String {
            let chosen = ignorable[random.below(ignorable.len() as u64) as usize];
            let prefix = match random.below(6) {
                0 | 1 => partner(chosen),
                2 => PREFIXES[random.below(PREFIXES.len() as u64) as usize],
                _ => chosen,
            };
            let local = locals[random.below(locals.len() as u64) as usize];
            format!("{prefix}:{local}")
        };
        for (name, locals, probability) in [
            (
                "PreserveAttributes",
                &["*", "k", "lang", "space", "q"][..],
                2,
            ),
            ("PreserveElements", &["keep", "*"][..], 3),
            ("ProcessContent", &["u", "*"][..], 3),
        ] {
            if random.below(probability) != 0 {
                continue;
            }
            let count = 1 + random.below(2);
            let targets: Vec<String> = (0..count).map(|_| target(random, locals)).collect();
            let _ = write!(xml, r#" {directive}:{name}="{}""#, targets.join(" "));
        }
    }
    if random.below(12) == 0 {
        let tokens = pick(random, &["a", "b", "w", "x", "xml"], 2);
        let _ = write!(xml, r#" {directive}:MustUnderstand="{}""#, tokens.join(" "));
    }
}

fn aliased_attributes(random: &mut Xorshift, xml: &mut String) {
    let names = [
        "a:k",
        "b:k",
        "x:lang",
        "xml:lang",
        "xml:space",
        "n:q",
        "d",
        "w:v",
        "b:v",
    ];
    // Two names that differ only in aliased prefixes are the same name, so
    // most lists avoid the pair and a few keep it.
    let alias = |name: &str| match name {
        "a:k" => "b:k",
        "b:k" => "a:k",
        "x:lang" => "xml:lang",
        "xml:lang" => "x:lang",
        _ => "",
    };
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let name = names[random.below(names.len() as u64) as usize];
        if used.contains(&name) || (used.contains(&alias(name)) && random.below(4) != 0) {
            continue;
        }
        used.push(name);
        let value = if name.ends_with(":space") {
            "preserve"
        } else {
            "v"
        };
        let _ = write!(xml, r#" {name}="{value}""#);
    }
}

fn aliased_children(random: &mut Xorshift, xml: &mut String, depth: usize) {
    if depth >= 4 {
        return;
    }
    let count = random.below(4);
    for _ in 0..count {
        let name = match random.below(12) {
            0 => "a:u",
            1 => "b:u",
            2 => "x:u",
            3 => "a:keep",
            4 => "b:keep",
            5 => "n:e",
            6 => "w:e",
            7 => "e",
            8 => "x:keep",
            9 => {
                aliased_alternate(random, xml, depth);
                continue;
            },
            10 => {
                xml.push_str("t &amp; u");
                continue;
            },
            _ => "b:other",
        };
        let _ = write!(xml, "<{name}");
        if random.below(3) == 0 {
            aliased_bindings(random, xml, false);
        }
        if random.below(4) == 0 {
            aliased_directives(random, xml);
        }
        aliased_attributes(random, xml);
        if random.below(3) == 0 {
            xml.push_str("/>");
        } else {
            xml.push('>');
            aliased_children(random, xml, depth + 1);
            let _ = write!(xml, "</{name}>");
        }
    }
}

/// An `AlternateContent` whose markup names the compatibility namespace
/// through `mc` or `m`, with `Requires` naming aliased prefixes.
fn aliased_alternate(random: &mut Xorshift, xml: &mut String, depth: usize) {
    let container = if random.below(2) == 0 { "mc" } else { "m" };
    let choice = if random.below(2) == 0 { "mc" } else { "m" };
    let requires = pick(random, &["a", "b", "w", "x", "xml", "n"], 2).join(" ");
    let _ = write!(
        xml,
        r#"<{container}:AlternateContent><{choice}:Choice Requires="{requires}">"#
    );
    aliased_children(random, xml, depth + 1);
    let _ = write!(xml, "</{choice}:Choice>");
    if random.below(2) == 0 {
        let fallback = if random.below(2) == 0 { "mc" } else { "m" };
        let _ = write!(xml, "<{fallback}:Fallback>");
        aliased_children(random, xml, depth + 1);
        let _ = write!(xml, "</{fallback}:Fallback>");
    }
    let _ = write!(xml, "</{container}:AlternateContent>");
}
