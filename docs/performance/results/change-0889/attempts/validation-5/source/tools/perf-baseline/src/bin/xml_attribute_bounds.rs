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
    Capabilities, Error as MceError, Limits as MceLimits, SemanticEvent, StreamError, StreamLimits,
    process_markup_compatibility, process_markup_compatibility_stream,
    process_markup_compatibility_stream_with_observers,
};
use litchi_opc::{BlobPart, PackURI};
use serde_json::{Value, json};
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
    digest_hex(Sha256::digest(bytes).as_ref())
}

fn digest_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(64);
    for byte in bytes {
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
        // `mce_stream_long_uri` with a 64-byte URI: the same attributes,
        // elements and events, so the two differ only in the URI's length.
        "mce_stream_short_uri" => Case {
            input: long_uri(64, "", 4, 1_000).into_bytes(),
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
// Qualification oracles
// --------------------------------------------------------------------------

/// The timing wrappers above intentionally keep their small, historical
/// outcome strings.  Differential qualification uses the records below so a
/// result is independently checkable: every byte-bearing field is length
/// framed before hashing, reports are represented by named counters, and
/// failures retain their public error category instead of only a Display
/// string.  This avoids treating a timing digest as a complete semantic
/// oracle.
fn mce_limits_value(limits: &MceLimits) -> Value {
    json!({
        "max_input_bytes": limits.max_input_bytes,
        "max_output_bytes": limits.max_output_bytes,
        "max_depth": limits.max_depth,
        "max_namespace_bindings": limits.max_namespace_bindings,
        "max_directive_tokens": limits.max_directive_tokens,
        "max_choices_per_alternate": limits.max_choices_per_alternate,
        "max_attributes_per_element": limits.max_attributes_per_element,
    })
}

fn stream_limits_value(limits: &StreamLimits) -> Value {
    json!({
        "processing": mce_limits_value(&limits.processing),
        "max_events": limits.max_events,
        "max_event_bytes": limits.max_event_bytes,
        "max_attributes_per_event": limits.max_attributes_per_event,
        "max_attribute_bytes_per_event": limits.max_attribute_bytes_per_event,
        "max_context_bytes": limits.max_context_bytes,
        "max_name_bytes": limits.max_name_bytes,
    })
}

fn mce_report_value(report: &litchi_ooxml_common::mce::Report) -> Value {
    json!({
        "alternate_content_count": report.alternate_content_count,
        "selected_choices": report.selected_choices,
        "selected_fallbacks": report.selected_fallbacks,
        "ignored_elements": report.ignored_elements,
        "ignored_attributes": report.ignored_attributes,
        "preserved_elements": report.preserved_elements,
        "preserved_attributes": report.preserved_attributes,
        "unwrapped_elements": report.unwrapped_elements,
    })
}

fn stream_report_value(report: &litchi_ooxml_common::mce::StreamReport) -> Value {
    json!({
        "alternate_content_count": report.alternate_content_count,
        "selected_choices": report.selected_choices,
        "selected_fallbacks": report.selected_fallbacks,
        "ignored_elements": report.ignored_elements,
        "ignored_attributes": report.ignored_attributes,
        "preserved_elements": report.preserved_elements,
        "preserved_attributes": report.preserved_attributes,
        "unwrapped_elements": report.unwrapped_elements,
        "events": report.events,
    })
}

fn mce_error_value(error: &MceError) -> Value {
    match error {
        MceError::NonConformant(message) => {
            json!({"kind": "non_conformant", "message": message})
        },
        MceError::MustUnderstand(message) => {
            json!({"kind": "must_understand", "message": message})
        },
        MceError::LimitExceeded(message) => {
            json!({"kind": "limit_exceeded", "message": message})
        },
        MceError::Xml(message) => json!({"kind": "xml", "message": message}),
        MceError::Allocation { resource, source } => json!({
            "kind": "allocation",
            "resource": resource,
            "message": source.to_string(),
        }),
        _ => json!({
            "kind": "unknown",
            "message": error.to_string(),
        }),
    }
}

fn stream_error_value(error: &StreamError<Infallible, Infallible>) -> Value {
    match error {
        StreamError::Input {
            error,
            prior_mce_error,
            raw_error,
            active_error,
        } => json!({
            "kind": "input",
            "error_kind": format!("{:?}", error.kind()),
            "message": error.to_string(),
            "prior_mce_error": prior_mce_error.as_ref().map(mce_error_value),
            "raw_callback_error": raw_error.is_some(),
            "active_callback_error": active_error.is_some(),
        }),
        StreamError::Mce {
            error,
            prior_mce_error,
            raw_error,
            active_error,
        } => json!({
            "kind": "mce",
            "error": mce_error_value(error),
            "prior_mce_error": prior_mce_error.as_ref().map(mce_error_value),
            "raw_callback_error": raw_error.is_some(),
            "active_callback_error": active_error.is_some(),
        }),
        StreamError::Callback {
            raw_error,
            active_error,
        } => json!({
            "kind": "callback",
            "raw_callback_error": raw_error.is_some(),
            "active_callback_error": active_error.is_some(),
        }),
        _ => json!({
            "kind": "unknown",
            "message": format!("{error}"),
        }),
    }
}

fn error_value_is_complete(value: &Value) -> bool {
    let Some(object) = value.as_object() else {
        return false;
    };
    if object.get("kind").and_then(Value::as_str) == Some("unknown") {
        return false;
    }
    ["error", "prior_mce_error"].into_iter().all(|field| {
        object
            .get(field)
            .is_none_or(|nested| nested.is_null() || error_value_is_complete(nested))
    })
}

/// Add one unambiguous field to a digest.  A delimiter-concatenated textual
/// signature can make `("a", "bc")` and `(“ab”, “c”)` collide; the length
/// prefix makes every event field and every record boundary observable.
fn hash_field(hasher: &mut Sha256, bytes: &[u8]) {
    hasher.update((bytes.len() as u64).to_be_bytes());
    hasher.update(bytes);
}

fn hash_tag(hasher: &mut Sha256, tag: &'static [u8]) {
    hash_field(hasher, tag);
}

fn hash_count(hasher: &mut Sha256, count: usize) {
    hasher.update((count as u64).to_be_bytes());
}

fn hash_expanded_name(
    hasher: &mut Sha256,
    qualified_name: &[u8],
    namespace: &str,
    local_name: &str,
) {
    hash_field(hasher, qualified_name);
    hash_field(hasher, namespace.as_bytes());
    hash_field(hasher, local_name.as_bytes());
}

fn hash_raw_element(
    hasher: &mut Sha256,
    element: &litchi_ooxml_common::mce::RawElement<'_>,
) -> bool {
    let (kind, unknown) = match element.kind {
        litchi_ooxml_common::mce::RawElementKind::Start => (b"start" as &[u8], false),
        litchi_ooxml_common::mce::RawElementKind::Empty => (b"empty" as &[u8], false),
        _ => (b"unknown" as &[u8], true),
    };
    hash_tag(hasher, kind);
    hash_expanded_name(
        hasher,
        element.name(),
        element.expanded_name.namespace.as_ref(),
        &element.expanded_name.local_name,
    );
    hash_count(hasher, element.attrs().len());
    for attribute in element.attrs() {
        hash_expanded_name(
            hasher,
            attribute.name(),
            attribute.expanded_name.namespace.as_ref(),
            &attribute.expanded_name.local_name,
        );
        hash_field(hasher, attribute.value());
        hash_field(hasher, attribute.decoded().as_bytes());
    }
    unknown
}

fn hash_semantic_element(
    hasher: &mut Sha256,
    element: &litchi_ooxml_common::mce::SemanticElement<'_>,
) {
    hash_expanded_name(
        hasher,
        element.name(),
        element.expanded_name.namespace.as_ref(),
        &element.expanded_name.local_name,
    );
    hash_count(hasher, element.attrs().len());
    for attribute in element.attrs() {
        hash_expanded_name(
            hasher,
            attribute.name(),
            attribute.expanded_name.namespace.as_ref(),
            &attribute.expanded_name.local_name,
        );
        hash_field(hasher, attribute.raw_value.as_ref());
        hash_field(hasher, attribute.value().as_bytes());
    }
}

fn hash_semantic_event(hasher: &mut Sha256, event: &SemanticEvent<'_>) -> bool {
    match event {
        SemanticEvent::Start(element) => {
            hash_tag(hasher, b"start");
            hash_semantic_element(hasher, element);
        },
        SemanticEvent::Empty(element) => {
            hash_tag(hasher, b"empty");
            hash_semantic_element(hasher, element);
        },
        SemanticEvent::End(element) => {
            hash_tag(hasher, b"end");
            hash_expanded_name(
                hasher,
                element.name(),
                element.expanded_name.namespace.as_ref(),
                &element.expanded_name.local_name,
            );
        },
        SemanticEvent::Text(text) => {
            hash_tag(hasher, b"text");
            hash_field(hasher, text.raw.as_ref());
            hash_field(hasher, text.text().as_bytes());
        },
        SemanticEvent::CData(text) => {
            hash_tag(hasher, b"cdata");
            hash_field(hasher, text.raw.as_ref());
            hash_field(hasher, text.text().as_bytes());
        },
        SemanticEvent::Comment(text) => {
            hash_tag(hasher, b"comment");
            hash_field(hasher, text.raw.as_ref());
            hash_field(hasher, text.text().as_bytes());
        },
        SemanticEvent::Decl(decl) => {
            hash_tag(hasher, b"decl");
            hash_field(hasher, decl.raw.as_ref());
        },
        SemanticEvent::GeneralRef(reference) => {
            hash_tag(hasher, b"general_ref");
            hash_field(hasher, reference.name.as_ref());
        },
        // SemanticEvent is non-exhaustive.  Keep a visible count and a stable
        // record marker if a future library adds a variant before this
        // harness is rebuilt to understand its fields.
        _ => {
            hash_tag(hasher, b"unknown");
            return true;
        },
    }
    false
}

fn codec_oracle(input: &[u8], capabilities: &Capabilities, capability_id: &str) -> (String, Value) {
    codec_oracle_with_limits(input, capabilities, capability_id, MceLimits::default())
}

fn codec_oracle_with_limits(
    input: &[u8],
    capabilities: &Capabilities,
    capability_id: &str,
    limits: MceLimits,
) -> (String, Value) {
    let input_sha256 = sha256_hex(input);
    match process_markup_compatibility(input, capabilities, &limits) {
        Ok(output) => {
            let output_sha256 = sha256_hex(output.xml.as_ref());
            let legacy = format!("ok:{output_sha256}:{:?}", output.report);
            let record = json!({
                "kind": "mce_codec",
                "capabilities": capability_id,
                "limits": mce_limits_value(&limits),
                "input": {"bytes": input.len(), "sha256": input_sha256},
                "success": true,
                "qualification_complete": true,
                "output": {"bytes": output.xml.len(), "sha256": output_sha256},
                "report": mce_report_value(&output.report),
                "error": Value::Null,
            });
            (legacy, record)
        },
        Err(error) => {
            let typed_error = mce_error_value(&error);
            let legacy = format!("err:{error}");
            let qualification_complete = error_value_is_complete(&typed_error);
            let record = json!({
                "kind": "mce_codec",
                "capabilities": capability_id,
                "limits": mce_limits_value(&limits),
                "input": {"bytes": input.len(), "sha256": input_sha256},
                "success": false,
                "output": Value::Null,
                "report": Value::Null,
                "error": typed_error,
                "qualification_complete": qualification_complete,
            });
            (legacy, record)
        },
    }
}

fn stream_oracle(
    input: &[u8],
    capabilities: &Capabilities,
    capability_id: &str,
) -> (String, Value) {
    stream_oracle_with_limits(input, capabilities, capability_id, StreamLimits::default())
}

fn stream_oracle_with_limits(
    input: &[u8],
    capabilities: &Capabilities,
    capability_id: &str,
    limits: StreamLimits,
) -> (String, Value) {
    let mut raw_hasher = Sha256::new();
    let mut semantic_hasher = Sha256::new();
    let mut legacy_raw_hasher = Sha256::new();
    let mut legacy_semantic_hasher = Sha256::new();
    let mut raw_events = 0usize;
    let mut semantic_events = 0usize;
    let mut unknown_semantic_events = 0usize;
    let mut unknown_raw_kinds = 0usize;
    let mut cursor = Cursor::new(input);
    let result = process_markup_compatibility_stream_with_observers(
        &mut cursor,
        capabilities,
        &limits,
        |element| {
            raw_events += 1;
            if hash_raw_element(&mut raw_hasher, &element) {
                unknown_raw_kinds += 1;
            }
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
            legacy_raw_hasher.update(line.as_bytes());
            legacy_raw_hasher.update(b"\n");
            Ok::<(), Infallible>(())
        },
        |event| {
            semantic_events += 1;
            if hash_semantic_event(&mut semantic_hasher, &event) {
                unknown_semantic_events += 1;
            }
            legacy_semantic_hasher.update(event_signature(&event).as_bytes());
            legacy_semantic_hasher.update(b"\n");
            Ok::<(), Infallible>(())
        },
    );
    let raw_sha256 = digest_hex(raw_hasher.finalize().as_ref());
    let semantic_sha256 = digest_hex(semantic_hasher.finalize().as_ref());
    // Preserve the historical timing-result strings exactly.  `stream_digest`
    // predates the structured oracle and intentionally hashes the 32-byte
    // digest once more; only the new framed fields use the direct digest.
    let legacy_raw_sha256 = sha256_hex(legacy_raw_hasher.finalize().as_ref());
    let legacy_semantic_sha256 = sha256_hex(legacy_semantic_hasher.finalize().as_ref());
    let legacy = match &result {
        Ok(report) => format!("ok:{legacy_semantic_sha256}:{legacy_raw_sha256}:{report:?}"),
        Err(error) => format!("err:{error}:{legacy_semantic_sha256}:{legacy_raw_sha256}"),
    };
    let error_value = result.as_ref().err().map(stream_error_value);
    let qualification_complete = unknown_raw_kinds == 0
        && unknown_semantic_events == 0
        && error_value.as_ref().is_none_or(error_value_is_complete);
    let record = json!({
        "kind": "mce_stream",
        "capabilities": capability_id,
        "limits": stream_limits_value(&limits),
        "input": {"bytes": input.len(), "sha256": sha256_hex(input)},
        "success": result.is_ok(),
        "qualification_complete": qualification_complete,
        "raw": {"events": raw_events, "sha256": raw_sha256},
        "raw_unknown_kinds": unknown_raw_kinds,
        "semantic": {
            "events": semantic_events,
            "unknown_events": unknown_semantic_events,
            "sha256": semantic_sha256,
        },
        "report": result.as_ref().ok().map(stream_report_value),
        "error": error_value,
    });
    (legacy, record)
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

struct PackageCandidate {
    path: PathBuf,
    first_scan_bytes: usize,
    first_scan_sha256: String,
}

impl PackageCandidate {
    fn identity_matches(&self, bytes: &[u8]) -> bool {
        bytes.len() == self.first_scan_bytes && sha256_hex(bytes) == self.first_scan_sha256
    }
}

struct PackageInput {
    path: String,
    bytes: usize,
    sha256: String,
    archive_member_count: usize,
    selected_member_count: usize,
    read_member_count: usize,
    members: Vec<Value>,
    member_receipts: Vec<Value>,
}

#[derive(Default)]
struct InputAccounting {
    packages: Vec<PackageInput>,
    failures: Vec<Value>,
    selected_member_count: usize,
    read_member_count: usize,
    seen_packages: std::collections::BTreeSet<PathBuf>,
    scanned_file_count: usize,
    excluded_files: Vec<Value>,
}

impl InputAccounting {
    fn record_failure(
        &mut self,
        stage: &'static str,
        path: &Path,
        member: Option<&str>,
        error: impl std::fmt::Display,
    ) {
        self.failures.push(json!({
            "stage": stage,
            "path": path.display().to_string(),
            "member": member,
            "error": error.to_string(),
        }));
    }

    fn value(&self, roots: &[PathBuf]) -> Value {
        let mut failure_counts = BTreeMap::<String, usize>::new();
        for failure in &self.failures {
            if let Some(stage) = failure.get("stage").and_then(Value::as_str) {
                *failure_counts.entry(stage.to_owned()).or_default() += 1;
            }
        }
        json!({
            "roots": roots.iter().map(|root| root.display().to_string()).collect::<Vec<_>>(),
            "packages": self.packages.iter().map(|package| json!({
                "path": package.path,
                "bytes": package.bytes,
                "sha256": package.sha256,
                "archive_member_count": package.archive_member_count,
                "selected_member_count": package.selected_member_count,
                "read_member_count": package.read_member_count,
                "members": package.members,
                "member_receipts": package.member_receipts,
            })).collect::<Vec<_>>(),
            "package_count": self.packages.len(),
            "scanned_file_count": self.scanned_file_count,
            "excluded_file_count": self.excluded_files.len(),
            "excluded_files": self.excluded_files,
            "selected_member_count": self.selected_member_count,
            "read_member_count": self.read_member_count,
            "read_failure_count": self.failures.len(),
            "failure_counts": failure_counts,
            "complete": self.failures.is_empty(),
            "failures": self.failures,
        })
    }

    fn record_excluded(&mut self, path: &Path, bytes: &[u8], reason: &'static str) {
        self.excluded_files.push(json!({
            "path": path.display().to_string(),
            "bytes": bytes.len(),
            "sha256": sha256_hex(bytes),
            "reason": reason,
        }));
    }
}

fn is_xml_member(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower.ends_with(".xml") || lower.ends_with(".rels") || lower.ends_with(".vml")
}

fn root_has_package(root: &Path, packages: &[PackageCandidate]) -> bool {
    packages
        .iter()
        .any(|package| package.path == root || package.path.starts_with(root))
}

fn attribute_boundary_document(attribute_count: usize) -> String {
    if attribute_count == 0 {
        return "<r/>".to_owned();
    }
    let mut xml = format!(r#"<r xmlns:mc="{MC}""#);
    for index in 0..attribute_count.saturating_sub(1) {
        let _ = write!(xml, r#" a{index}="v""#);
    }
    xml.push_str("/>");
    xml
}

fn semantic_control_document() -> String {
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:i="urn:i" xmlns:w="urn:w"
        mc:Ignorable="i" mc:ProcessContent="i:u"
        mc:PreserveElements="i:keep" mc:PreserveAttributes="i:keep">
        <mc:AlternateContent>
          <mc:Choice Requires="i"><i:keep i:keep="v"/></mc:Choice>
          <mc:Fallback><w:fallback/></mc:Fallback>
        </mc:AlternateContent>
        <i:u><i:child/></i:u>
        <i:keep i:keep="v"/>
        </r>"#
    )
}

fn qualification_controls() -> (Vec<Value>, usize) {
    let mut controls = Vec::new();
    let mut failures = 0usize;
    let attribute_cases = [
        ("no_mce_zero_attributes", 0usize, 1024usize, "success"),
        ("attributes_exact_one", 1, 1024, "success"),
        ("attributes_exact_four", 4, 1024, "success"),
        ("attributes_exact_five", 5, 1024, "success"),
        ("attributes_exact_1024", 1024, 1024, "success"),
        ("attributes_over_1024", 1025, 1024, "limit"),
        ("attributes_exact_4096", 4096, 4096, "success"),
        ("attributes_over_4096", 4097, 4096, "limit"),
    ];
    for (name, attribute_count, configured_limit, expectation) in attribute_cases {
        let document = attribute_boundary_document(attribute_count);
        let processing = MceLimits {
            max_attributes_per_element: configured_limit,
            ..MceLimits::default()
        };
        let (codec_legacy, codec) = codec_oracle_with_limits(
            document.as_bytes(),
            &Capabilities::new(),
            "bare",
            processing.clone(),
        );
        let stream_limits = StreamLimits::new(processing);
        let (stream_legacy, stream) = stream_oracle_with_limits(
            document.as_bytes(),
            &Capabilities::new(),
            "bare",
            stream_limits,
        );
        let codec_expected = match expectation {
            "success" => codec.get("success") == Some(&Value::Bool(true)),
            "limit" => {
                codec.get("success") == Some(&Value::Bool(false))
                    && codec.get("error").and_then(|error| error.get("kind"))
                        == Some(&Value::String("limit_exceeded".to_owned()))
            },
            _ => false,
        };
        let stream_expected = match expectation {
            "success" => stream.get("success") == Some(&Value::Bool(true)),
            "limit" => {
                stream.get("success") == Some(&Value::Bool(false))
                    && stream.get("error").and_then(|error| error.get("kind"))
                        == Some(&Value::String("mce".to_owned()))
                    && stream
                        .get("error")
                        .and_then(|error| error.get("error"))
                        .and_then(|error| error.get("kind"))
                        == Some(&Value::String("limit_exceeded".to_owned()))
            },
            _ => false,
        };
        let expected_codec_error_message = if expectation == "limit" {
            Some("attributes per element")
        } else {
            None
        };
        let expected_stream_error_message = if expectation == "limit" {
            Some(if configured_limit == 4096 {
                "stream attributes per event"
            } else {
                "attributes per element"
            })
        } else {
            None
        };
        let codec_error_message = codec
            .get("error")
            .and_then(|error| error.get("message"))
            .and_then(Value::as_str);
        let stream_error_message = stream
            .get("error")
            .and_then(|error| error.get("error"))
            .and_then(|error| error.get("message"))
            .and_then(Value::as_str);
        let expectation_met = codec_expected
            && stream_expected
            && codec.get("qualification_complete") == Some(&Value::Bool(true))
            && stream.get("qualification_complete") == Some(&Value::Bool(true))
            && codec_error_message == expected_codec_error_message
            && stream_error_message == expected_stream_error_message;
        if !expectation_met {
            failures += 1;
        }
        controls.push(json!({
            "name": name,
            "family": "attribute_cardinality",
            "attribute_count_including_xmlns": attribute_count,
            "configured_limit": configured_limit,
            "expectation": expectation,
            "expectation_met": expectation_met,
            "codec_legacy": codec_legacy,
            "stream_legacy": stream_legacy,
            "codec": codec,
            "stream": stream,
        }));
    }

    let semantic_document = semantic_control_document();
    for (name, capabilities, capability_id) in [
        ("semantic_bare", Capabilities::new(), "bare"),
        (
            "semantic_understood_extension",
            understanding("urn:i"),
            "understood:urn:i",
        ),
    ] {
        let (codec_legacy, codec) =
            codec_oracle(semantic_document.as_bytes(), &capabilities, capability_id);
        let (stream_legacy, stream) =
            stream_oracle(semantic_document.as_bytes(), &capabilities, capability_id);
        let understood = capability_id == "understood:urn:i";
        let selected_choices = if understood { 1 } else { 0 };
        let selected_fallbacks = if understood { 0 } else { 1 };
        let ignored_elements = if understood { 0 } else { 1 };
        let preserved_elements = if understood { 0 } else { 2 };
        let unwrapped_elements = if understood { 0 } else { 1 };
        let expectation_met = codec.get("qualification_complete") == Some(&Value::Bool(true))
            && stream.get("qualification_complete") == Some(&Value::Bool(true))
            && codec
                .get("report")
                .and_then(|report| report.get("selected_choices"))
                == Some(&json!(selected_choices))
            && codec
                .get("report")
                .and_then(|report| report.get("selected_fallbacks"))
                == Some(&json!(selected_fallbacks))
            && codec
                .get("report")
                .and_then(|report| report.get("ignored_elements"))
                == Some(&json!(ignored_elements))
            && codec
                .get("report")
                .and_then(|report| report.get("preserved_elements"))
                == Some(&json!(preserved_elements))
            && codec
                .get("report")
                .and_then(|report| report.get("unwrapped_elements"))
                == Some(&json!(unwrapped_elements))
            && stream
                .get("report")
                .and_then(|report| report.get("selected_choices"))
                == Some(&json!(selected_choices))
            && stream
                .get("report")
                .and_then(|report| report.get("selected_fallbacks"))
                == Some(&json!(selected_fallbacks))
            && stream
                .get("report")
                .and_then(|report| report.get("ignored_elements"))
                == Some(&json!(ignored_elements))
            && stream
                .get("report")
                .and_then(|report| report.get("preserved_elements"))
                == Some(&json!(preserved_elements))
            && stream
                .get("report")
                .and_then(|report| report.get("unwrapped_elements"))
                == Some(&json!(unwrapped_elements));
        if !expectation_met {
            failures += 1;
        }
        controls.push(json!({
            "name": name,
            "family": "semantic_selection",
            "expectation": "path_specific_codec_and_stream_records",
            "expectation_met": expectation_met,
            "codec_legacy": codec_legacy,
            "stream_legacy": stream_legacy,
            "codec": codec,
            "stream": stream,
        }));
    }
    (controls, failures)
}

fn validate_differential_args(args: &[String]) -> Result<(), Box<dyn Error>> {
    let mut index = 0usize;
    let mut seen_options = std::collections::BTreeSet::new();
    while index < args.len() {
        let arg = &args[index];
        if !arg.starts_with("--") {
            index += 1;
            continue;
        }
        if !matches!(arg.as_str(), "--json" | "--generated" | "--aliasing") {
            return Err(format!("unknown differential option {arg}").into());
        }
        if !seen_options.insert(arg.as_str()) {
            return Err(format!("duplicate differential option {arg}").into());
        }
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("differential option {arg} requires a value"))?;
        if value.starts_with("--") {
            return Err(format!("differential option {arg} requires a value").into());
        }
        index += 2;
    }
    Ok(())
}

fn differential(args: &[String]) -> Result<(), Box<dyn Error>> {
    validate_differential_args(args)?;
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
    let mut input = InputAccounting::default();
    for root in &roots {
        collect_packages(root, &mut packages, &mut input)?;
    }
    for root in &roots {
        if !root_has_package(root, &packages) {
            input.record_failure(
                "no_ooxml_package",
                root,
                None,
                "root contained no OOXML package",
            );
        }
    }
    packages.sort_by(|left, right| left.path.cmp(&right.path));
    let mut members = 0usize;
    let mut oracles = BTreeMap::new();
    for package in packages {
        let path = package.path.clone();
        let bytes = match std::fs::read(&path) {
            Ok(bytes) => bytes,
            Err(error) => {
                input.record_failure("package_read", &path, None, error);
                continue;
            },
        };
        if !package.identity_matches(&bytes) {
            input.record_failure(
                "package_changed",
                &path,
                None,
                format!(
                    "package content changed after discovery (first scan {} bytes / {}, current {} bytes / {})",
                    package.first_scan_bytes,
                    package.first_scan_sha256,
                    bytes.len(),
                    sha256_hex(&bytes),
                ),
            );
            continue;
        }
        let archive = match ArchiveReader::new(&bytes) {
            Ok(archive) => archive,
            Err(error) => {
                input.record_failure("package_open", &path, None, error);
                continue;
            },
        };
        let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        let mut package_input = PackageInput {
            path: path.display().to_string(),
            bytes: bytes.len(),
            sha256: sha256_hex(&bytes),
            archive_member_count: names.len(),
            selected_member_count: 0,
            read_member_count: 0,
            members: Vec::new(),
            member_receipts: Vec::new(),
        };
        for name in names {
            if !is_xml_member(&name) {
                continue;
            }
            package_input.selected_member_count += 1;
            input.selected_member_count += 1;
            let declared = match archive.metadata(&name) {
                Ok(metadata) => Some(json!({
                    "compressed_bytes": metadata.compressed_size(),
                    "uncompressed_bytes": metadata.uncompressed_size(),
                })),
                Err(error) => {
                    input.record_failure("member_metadata_read", &path, Some(&name), error);
                    None
                },
            };
            let member = match archive.read(&name) {
                Ok(member) => member,
                Err(error) => {
                    package_input.member_receipts.push(json!({
                        "name": name.clone(),
                        "declared": declared,
                        "read": false,
                        "bytes": Value::Null,
                        "sha256": Value::Null,
                        "error": error.to_string(),
                    }));
                    input.record_failure("member_read", &path, Some(&name), error);
                    continue;
                },
            };
            members += 1;
            package_input.read_member_count += 1;
            input.read_member_count += 1;
            package_input.members.push(json!({
                "name": name.clone(),
                "bytes": member.len(),
                "sha256": sha256_hex(&member),
            }));
            package_input.member_receipts.push(json!({
                "name": name.clone(),
                "declared": declared,
                "read": true,
                "bytes": member.len(),
                "sha256": sha256_hex(&member),
            }));
            let key = format!("{}::{name}", path.display());
            let (legacy, oracle) = codec_oracle(&member, &baseline, "ooxml_baseline");
            results.insert(format!("{key}::mce_baseline"), legacy);
            oracles.insert(format!("{key}::mce_baseline"), oracle);
            let (legacy, oracle) = codec_oracle(&member, &bare, "bare");
            results.insert(format!("{key}::mce_bare"), legacy);
            oracles.insert(format!("{key}::mce_bare"), oracle);
            let (legacy, oracle) = stream_oracle(&member, &baseline, "ooxml_baseline");
            results.insert(format!("{key}::stream_baseline"), legacy);
            oracles.insert(format!("{key}::stream_baseline"), oracle);
            let [source, authored, compact] = audit_digest(&member);
            results.insert(format!("{key}::audit_source"), source);
            results.insert(format!("{key}::audit_authored"), authored);
            results.insert(format!("{key}::audit_compact"), compact);
            if name.to_ascii_lowercase().ends_with("styles.xml") {
                results.insert(format!("{key}::docx_styles"), styles_digest(&member));
            }
        }
        package_input.members.sort_by(|left, right| {
            left.get("name")
                .and_then(Value::as_str)
                .cmp(&right.get("name").and_then(Value::as_str))
        });
        package_input.member_receipts.sort_by(|left, right| {
            left.get("name")
                .and_then(Value::as_str)
                .cmp(&right.get("name").and_then(Value::as_str))
        });
        input.packages.push(package_input);
    }
    for index in 0..generated {
        let document = generated_document(index as u64);
        let key = format!("generated::{index:06}");
        let understood = understanding_w();
        for (label, capability_id, capabilities) in [
            ("bare", "bare", &bare),
            ("understood", "understood:urn:w", &understood),
        ] {
            let (legacy, oracle) = codec_oracle(document.as_bytes(), capabilities, capability_id);
            results.insert(format!("{key}::mce_{label}"), legacy);
            oracles.insert(format!("{key}::mce_{label}"), oracle);
            let (legacy, oracle) = stream_oracle(document.as_bytes(), capabilities, capability_id);
            results.insert(format!("{key}::stream_{label}"), legacy);
            oracles.insert(format!("{key}::stream_{label}"), oracle);
        }
    }
    // Record 0771: documents whose prefixes alias one another and the fixed
    // namespaces, under three capability sets.
    let aliasing: usize = option(args, "--aliasing").unwrap_or("0").parse()?;
    let understood_a = understanding("urn:a");
    for index in 0..aliasing {
        let document = aliasing_document(index as u64);
        let key = format!("aliasing::{index:06}");
        for (label, capability_id, capabilities) in [
            ("bare", "bare", &bare),
            ("baseline", "ooxml_baseline", &baseline),
            ("understood", "understood:urn:a", &understood_a),
        ] {
            let (legacy, oracle) = codec_oracle(document.as_bytes(), capabilities, capability_id);
            results.insert(format!("{key}::mce_{label}"), legacy);
            oracles.insert(format!("{key}::mce_{label}"), oracle);
            let (legacy, oracle) = stream_oracle(document.as_bytes(), capabilities, capability_id);
            results.insert(format!("{key}::stream_{label}"), legacy);
            oracles.insert(format!("{key}::stream_{label}"), oracle);
        }
    }
    let (controls, control_failure_count) = qualification_controls();
    let incomplete_oracle_count = oracles
        .values()
        .filter(|oracle| oracle.get("qualification_complete") == Some(&Value::Bool(false)))
        .count();
    let summary = serde_json::json!({
        "schema": "litchi.xml-attribute-bounds.differential.v2",
        "packages": input.packages.len(),
        "members": members,
        "generated": generated,
        "aliasing": aliasing,
        "results": results,
        "oracles": oracles,
        "input_inventory": input.value(&roots),
        "controls": controls,
        "qualification_complete": input.failures.is_empty()
            && incomplete_oracle_count == 0
            && control_failure_count == 0,
        "incomplete_oracle_count": incomplete_oracle_count,
        "control_failure_count": control_failure_count,
    });
    std::fs::write(json, serde_json::to_vec(&summary)?)?;
    if !input.failures.is_empty() || incomplete_oracle_count != 0 || control_failure_count != 0 {
        return Err(format!(
            "differential qualification incomplete: {} input failures, {} incomplete oracles, {} control failures (see JSON report)",
            input.failures.len(),
            incomplete_oracle_count,
            control_failure_count,
        )
        .into());
    }
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

fn collect_packages(
    root: &Path,
    packages: &mut Vec<PackageCandidate>,
    input: &mut InputAccounting,
) -> Result<(), Box<dyn Error>> {
    let metadata = match std::fs::symlink_metadata(root) {
        Ok(metadata) => metadata,
        Err(error) => {
            input.record_failure("root_stat", root, None, error);
            return Ok(());
        },
    };
    if metadata.is_dir() {
        let mut entries = Vec::new();
        let read_dir = match std::fs::read_dir(root) {
            Ok(read_dir) => read_dir,
            Err(error) => {
                input.record_failure("directory_read", root, None, error);
                return Ok(());
            },
        };
        for entry in read_dir {
            match entry {
                Ok(entry) => entries.push(entry),
                Err(error) => input.record_failure("directory_entry_read", root, None, error),
            }
        }
        entries.sort_by_key(std::fs::DirEntry::path);
        for entry in entries {
            collect_packages(&entry.path(), packages, input)?;
        }
    } else if metadata.is_file() {
        input.scanned_file_count += 1;
        let bytes = match std::fs::read(root) {
            Ok(bytes) => bytes,
            Err(error) => {
                input.record_failure("package_scan_read", root, None, error);
                return Ok(());
            },
        };
        if bytes.starts_with(b"PK\x03\x04") {
            match ArchiveReader::new(&bytes) {
                Ok(archive) if archive.contains("[Content_Types].xml") => {
                    let identity = match std::fs::canonicalize(root) {
                        Ok(identity) => identity,
                        Err(error) => {
                            input.record_failure("package_identity", root, None, error);
                            root.to_owned()
                        },
                    };
                    if input.seen_packages.insert(identity) {
                        packages.push(PackageCandidate {
                            path: root.to_owned(),
                            first_scan_bytes: bytes.len(),
                            first_scan_sha256: sha256_hex(&bytes),
                        });
                    } else {
                        input.record_failure(
                            "duplicate_package",
                            root,
                            None,
                            "package reached through overlapping input roots",
                        );
                    }
                },
                Ok(_) => input.record_excluded(root, &bytes, "zip_without_content_types"),
                Err(error) => input.record_failure("package_open", root, None, error),
            }
        } else {
            input.record_excluded(root, &bytes, "not_ooxml_zip");
        }
    } else {
        input.record_failure(
            "unsupported_filesystem_entry",
            root,
            None,
            "input entry is neither a directory nor a regular file",
        );
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

#[cfg(test)]
#[path = "xml_attribute_bounds/qualification_tests.rs"]
mod qualification_tests;

#[cfg(test)]
#[path = "xml_attribute_bounds/boundary_tests.rs"]
mod boundary_tests;

#[cfg(test)]
#[path = "xml_attribute_bounds/input_tests.rs"]
mod input_tests;
