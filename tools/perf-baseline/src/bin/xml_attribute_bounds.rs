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
//!   compared member by member.

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
    process_markup_compatibility_stream_with_observers,
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
        // 32 nested elements re-declaring 1,000 prefixes each, around an
        // element of 1,000 attributes in a namespace bound at the root.
        "mce_shadowed_chain" => Case {
            input: shadowed_chain(32, 1_000, 1_000).into_bytes(),
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
        _ => return Err(format!("unknown case {name}").into()),
    })
}

fn declaration_flood(count: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{MC}"><t"#);
    for index in 0..count {
        let _ = write!(xml, r#" xmlns:p{index}="urn:{index}""#);
    }
    xml.push_str("/></r>");
    xml
}

fn shadowed_chain(depth: usize, width: usize, attributes: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:z="urn:z">"#);
    for level in 0..depth {
        xml.push_str("<s");
        for prefix in 0..width {
            let _ = write!(xml, r#" xmlns:q{prefix}="urn:{level}:{prefix}""#);
        }
        xml.push('>');
    }
    xml.push_str("<z:e");
    for index in 0..attributes {
        let _ = write!(xml, r#" z:a{index}="""#);
    }
    xml.push_str("/>");
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
    let summary = serde_json::json!({
        "packages": packages.len(),
        "members": members,
        "generated": generated,
        "results": results,
    });
    std::fs::write(json, serde_json::to_vec(&summary)?)?;
    Ok(())
}

fn understanding_w() -> Capabilities {
    let mut capabilities = Capabilities::new();
    capabilities.understand_namespace("urn:w");
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
