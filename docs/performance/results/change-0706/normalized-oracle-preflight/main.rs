//! Standalone, measurement-free differential probe for the public XML auditor.
//!
//! The probe deliberately lives outside the workspace package graph. The
//! caller can build it once against the baseline xml-minifier source, then
//! replace only that path dependency's source and build it again. It emits a
//! deterministic JSON object containing every normalized public result and a
//! SHA-256 over the complete ordered result stream. It performs no timing,
//! allocation, or process measurements.

#![deny(unsafe_code)]
#![deny(warnings)]

use std::fmt::Write as _;
use std::hint::black_box;
use std::io::{self, BufRead, Read};
use std::panic::{self, AssertUnwindSafe};
use std::time::Instant;

use sha2::{Digest as _, Sha256};
use xml_minifier::audit::{self, Limits};

const SCHEMA: &str = "litchi.xml-minifier-public-oracle.v1";
const GENERATED_CASES: usize = 4_096;

const CHUNKS_ONE: &[usize] = &[1];
const CHUNKS_MIXED: &[usize] = &[2, 1, 3, 5, 8, 13];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum LimitProfile {
    Normal,
    BytesBelow,
    AttributesZero,
    AttributesOne,
    DepthZero,
    EventsOne,
    TextZero,
    TokenBelow,
}

impl LimitProfile {
    const fn name(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::BytesBelow => "bytes_below",
            Self::AttributesZero => "attributes_zero",
            Self::AttributesOne => "attributes_one",
            Self::DepthZero => "depth_zero",
            Self::EventsOne => "events_one",
            Self::TextZero => "text_zero",
            Self::TokenBelow => "token_below",
        }
    }
}

struct Case {
    id: usize,
    name: String,
    input: Vec<u8>,
    limits: LimitProfile,
}

struct Chunked<'a> {
    input: &'a [u8],
    chunks: &'a [usize],
    position: usize,
    chunk_index: usize,
}

impl<'a> Chunked<'a> {
    const fn new(input: &'a [u8], chunks: &'a [usize]) -> Self {
        Self {
            input,
            chunks,
            position: 0,
            chunk_index: 0,
        }
    }

    fn available_len(&self) -> usize {
        if self.position >= self.input.len() {
            return 0;
        }
        let requested = self.chunks[self.chunk_index % self.chunks.len()];
        requested.max(1).min(self.input.len() - self.position)
    }
}

impl Read for Chunked<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for Chunked<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        let length = self.available_len();
        Ok(&self.input[self.position..self.position + length])
    }

    fn consume(&mut self, amount: usize) {
        let count = amount.min(self.available_len());
        self.position += count;
        if count != 0 {
            self.chunk_index = self.chunk_index.wrapping_add(1);
        }
    }
}

fn main() {
    // A malformed candidate must become a deterministic outcome rather than
    // truncate the corpus and accidentally look equivalent.
    panic::set_hook(Box::new(|_| {}));

    if std::env::args().nth(1).as_deref() == Some("--bench") {
        run_bench();
        return;
    }

    let cases = build_cases();
    let mut digest = Sha256::new();
    digest.update(SCHEMA.as_bytes());
    digest.update(b"\n");

    let mut output = String::new();
    output.push('{');
    write_json_key_value(&mut output, "schema", SCHEMA, true);
    write!(output, ",\"casecount\":{}", cases.len()).expect("writing String cannot fail");
    output.push_str(",\"chunk_patterns\":[\"one_byte\",\"mixed\"]");
    output.push_str(",\"outcomes\":[");

    let mut call_count = 0usize;
    for (case_index, case) in cases.iter().enumerate() {
        if case_index != 0 {
            output.push(',');
        }
        let input_digest = hex_digest(&case.input);
        write!(
            output,
            "{{\"id\":{},\"name\":{},\"bytes\":{},\"input_sha256\":{},\"limit_profile\":{},\"results\":[",
            case.id,
            json_string(&case.name),
            case.input.len(),
            json_string(&input_digest),
            json_string(case.limits.name()),
        )
        .expect("writing String cannot fail");

        let mut result_index = 0usize;
        for (policy, result) in slice_results(case) {
            append_result(
                &mut output,
                &mut digest,
                &mut result_index,
                &mut call_count,
                case,
                policy,
                "whole",
                result,
            );
        }
        for (chunk_name, chunks) in [("one_byte", CHUNKS_ONE), ("mixed", CHUNKS_MIXED)] {
            for (policy, result) in stream_results(case, chunks) {
                append_result(
                    &mut output,
                    &mut digest,
                    &mut result_index,
                    &mut call_count,
                    case,
                    policy,
                    chunk_name,
                    result,
                );
            }
        }
        output.push_str("]}");
    }
    output.push(']');

    let results_digest = hex_bytes(&digest.finalize());
    write!(
        output,
        ",\"callcount\":{},\"results_digest\":{} }}\n",
        call_count,
        json_string(&results_digest),
    )
    .expect("writing String cannot fail");
    print!("{output}");
}

type SliceResult = (&'static str, String);

fn slice_results(case: &Case) -> [SliceResult; 3] {
    [
        (
            "slice/compact",
            capture_result(|| audit::verify(&case.input, limits_for(case))),
        ),
        (
            "slice/authored",
            capture_result(|| audit::verify_authored(&case.input, limits_for(case))),
        ),
        (
            "slice/source",
            capture_result(|| audit::verify_source(&case.input, limits_for(case))),
        ),
    ]
}

fn stream_results(case: &Case, chunks: &'static [usize]) -> [SliceResult; 2] {
    [
        (
            "stream/compact",
            capture_result(|| {
                audit::verify_reader(Chunked::new(&case.input, chunks), limits_for(case))
            }),
        ),
        (
            "stream/authored",
            capture_result(|| {
                audit::verify_authored_reader(Chunked::new(&case.input, chunks), limits_for(case))
            }),
        ),
    ]
}

fn capture_result<T, F>(operation: F) -> String
where
    T: std::fmt::Debug,
    F: FnOnce() -> T,
{
    match panic::catch_unwind(AssertUnwindSafe(operation)) {
        Ok(result) => normalize_debug(&format!("{result:?}")),
        Err(_) => "PANIC".to_owned(),
    }
}

fn append_result(
    output: &mut String,
    digest: &mut Sha256,
    result_index: &mut usize,
    call_count: &mut usize,
    case: &Case,
    policy: &str,
    chunk: &str,
    result: String,
) {
    if *result_index != 0 {
        output.push(',');
    }
    write!(
        output,
        "{{\"policy\":{},\"chunk\":{},\"debug\":{}}}",
        json_string(policy),
        json_string(chunk),
        json_string(&result),
    )
    .expect("writing String cannot fail");

    // Include every identity and invocation dimension in the digest. The
    // complete normalized Debug result carries report counters, error details,
    // and all observed offsets instead of reducing outcomes to a category.
    let mut canonical = String::new();
    write!(
        canonical,
        "case={}\tname={}\tbytes={}\tinput_sha256={}\tlimits={}\tpolicy={}\tchunk={}\tdebug={}\n",
        case.id,
        case.name,
        case.input.len(),
        hex_digest(&case.input),
        case.limits.name(),
        policy,
        chunk,
        result,
    )
    .expect("writing String cannot fail");
    digest.update(canonical.as_bytes());

    *result_index += 1;
    *call_count += 1;
}

fn normalize_debug(value: &str) -> String {
    let mut normalized = String::with_capacity(value.len());
    let mut pending_space = false;
    for character in value.chars() {
        if character.is_ascii_whitespace() {
            pending_space = !normalized.is_empty();
            continue;
        }
        if pending_space {
            normalized.push(' ');
            pending_space = false;
        }
        normalized.push(character);
    }
    normalized
}

fn limits_for(case: &Case) -> Limits {
    let length = case.input.len();
    let base = Limits::new(length, 64, 4_096, 512, length.max(1), length)
        .expect("probe limit profile must stay below immutable ceilings");
    let (bytes, depth, events, attributes, token_bytes, text_bytes) = match case.limits {
        LimitProfile::Normal => (
            base.max_bytes(),
            base.max_depth(),
            base.max_events(),
            base.max_attributes(),
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::BytesBelow => (
            length.saturating_sub(1),
            base.max_depth(),
            base.max_events(),
            base.max_attributes(),
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::AttributesZero => (
            base.max_bytes(),
            base.max_depth(),
            base.max_events(),
            0,
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::AttributesOne => (
            base.max_bytes(),
            base.max_depth(),
            base.max_events(),
            1,
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::DepthZero => (
            base.max_bytes(),
            0,
            base.max_events(),
            base.max_attributes(),
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::EventsOne => (
            base.max_bytes(),
            base.max_depth(),
            1,
            base.max_attributes(),
            base.max_token_bytes(),
            base.max_text_bytes(),
        ),
        LimitProfile::TextZero => (
            base.max_bytes(),
            base.max_depth(),
            base.max_events(),
            base.max_attributes(),
            base.max_token_bytes(),
            0,
        ),
        LimitProfile::TokenBelow => (
            base.max_bytes(),
            base.max_depth(),
            base.max_events(),
            base.max_attributes(),
            length.saturating_sub(1),
            base.max_text_bytes(),
        ),
    };
    Limits::new(bytes, depth, events, attributes, token_bytes, text_bytes)
        .expect("probe limit profile must stay below immutable ceilings")
}

fn build_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    let explicit: &[(&str, &[u8], LimitProfile)] = &[
        ("empty", b"", LimitProfile::Normal),
        (
            "zero-attributes-empty-element",
            b"<a/>",
            LimitProfile::Normal,
        ),
        (
            "one-attribute-double-quote",
            b"<a x=\"1\"/>",
            LimitProfile::Normal,
        ),
        (
            "one-attribute-single-quote",
            b"<a x='1'/>",
            LimitProfile::Normal,
        ),
        (
            "two-attributes",
            b"<a a=\"1\" b=\"2\"/>",
            LimitProfile::Normal,
        ),
        (
            "many-attributes",
            b"<a a=\"1\" b=\"2\" c=\"3\" d=\"4\" e=\"5\" f=\"6\"/>",
            LimitProfile::Normal,
        ),
        (
            "duplicate-attribute",
            b"<a a=\"1\" a=\"2\"/>",
            LimitProfile::Normal,
        ),
        (
            "duplicate-attribute-with-two-distinct",
            b"<a a=\"1\" b=\"2\" a=\"3\"/>",
            LimitProfile::AttributesOne,
        ),
        (
            "attribute-separator-double-space",
            b"<a  x=\"1\"/>",
            LimitProfile::Normal,
        ),
        (
            "attribute-separator-tab",
            b"<a\tx=\"1\"/>",
            LimitProfile::Normal,
        ),
        (
            "attribute-space-around-equals",
            b"<a x = \"1\"/>",
            LimitProfile::Normal,
        ),
        (
            "attribute-quote-in-single-quoted-value",
            b"<a x='a\"b'/>",
            LimitProfile::Normal,
        ),
        (
            "attribute-apostrophe-in-double-quoted-value",
            b"<a x=\"a'b\"/>",
            LimitProfile::Normal,
        ),
        (
            "attribute-entities-in-value",
            b"<a x=\"&quot;&apos;&amp;&lt;&gt;\"/>",
            LimitProfile::Normal,
        ),
        ("utf8-text", "<a>café</a>".as_bytes(), LimitProfile::Normal),
        (
            "utf8-attribute",
            "<a x=\"雪\"/>".as_bytes(),
            LimitProfile::Normal,
        ),
        ("leading-bom", b"\xef\xbb\xbf<a/>", LimitProfile::Normal),
        (
            "leading-bom-with-declaration",
            b"\xef\xbb\xbf<?xml version=\"1.0\"?><a/>",
            LimitProfile::Normal,
        ),
        (
            "double-leading-bom",
            b"\xef\xbb\xbf\xef\xbb\xbf<a/>",
            LimitProfile::Normal,
        ),
        (
            "xml-space-preserve",
            b"<a xml:space=\"preserve\"> </a>",
            LimitProfile::Normal,
        ),
        (
            "xml-space-default-newline",
            b"<a xml:space=\"default\">\n</a>",
            LimitProfile::Normal,
        ),
        (
            "xml-space-invalid",
            b"<a xml:space=\"keep\">x</a>",
            LimitProfile::Normal,
        ),
        (
            "entity-and-character-references",
            b"<a>&amp;&#32;&#x20;</a>",
            LimitProfile::Normal,
        ),
        ("unknown-entity", b"<a>&unknown;</a>", LimitProfile::Normal),
        ("malformed-entity", b"<a>&#xZZ;</a>", LimitProfile::Normal),
        (
            "cdata-inside",
            b"<a><![CDATA[x]]></a>",
            LimitProfile::Normal,
        ),
        ("doctype", b"<!DOCTYPE a><a/>", LimitProfile::Normal),
        ("cdata-outside", b"<![CDATA[x]]><a/>", LimitProfile::Normal),
        ("leading-formatting-space", b" <a/>", LimitProfile::Normal),
        ("trailing-formatting-space", b"<a/> ", LimitProfile::Normal),
        ("ambiguous-space-text", b"<a> </a>", LimitProfile::Normal),
        (
            "structural-newline-text",
            b"<a>\n</a>",
            LimitProfile::Normal,
        ),
        ("space-before-empty-close", b"<a />", LimitProfile::Normal),
        ("space-before-end-close", b"<a></a >", LimitProfile::Normal),
        ("unclosed-element", b"<a>", LimitProfile::Normal),
        ("mismatched-end-element", b"<a></b>", LimitProfile::Normal),
        ("multiple-roots", b"<a/><b/>", LimitProfile::Normal),
        ("unexpected-end-element", b"</a>", LimitProfile::Normal),
        (
            "unterminated-comment",
            b"<a><!--x</a>",
            LimitProfile::Normal,
        ),
        (
            "unterminated-quoted-attribute",
            b"<a x=\"1/>",
            LimitProfile::Normal,
        ),
        ("invalid-utf8-text", b"<a>\xff</a>", LimitProfile::Normal),
        (
            "invalid-utf8-attribute",
            b"<a x=\"\xff\"/>",
            LimitProfile::Normal,
        ),
        (
            "bytes-limit-below",
            b"<a x=\"1\"/>",
            LimitProfile::BytesBelow,
        ),
        (
            "attributes-limit-zero",
            b"<a x=\"1\"/>",
            LimitProfile::AttributesZero,
        ),
        (
            "attributes-limit-one",
            b"<a a=\"1\" b=\"2\"/>",
            LimitProfile::AttributesOne,
        ),
        ("depth-limit-zero", b"<a/>", LimitProfile::DepthZero),
        ("events-limit-one", b"<a/>", LimitProfile::EventsOne),
        ("text-limit-zero", b"<a>x</a>", LimitProfile::TextZero),
        (
            "token-limit-below",
            b"<a x=\"1\"/>",
            LimitProfile::TokenBelow,
        ),
        ("bytes-limit-exact", b"<a/>", LimitProfile::Normal),
        (
            "attributes-limit-exact-one",
            b"<a x=\"1\"/>",
            LimitProfile::AttributesOne,
        ),
        ("max-depth-one", b"<a><b/></a>", LimitProfile::Normal),
    ];
    for (name, input, limits) in explicit {
        push_case(&mut cases, (*name).to_owned(), input.to_vec(), *limits);
    }

    for index in 0..GENERATED_CASES {
        let (input, attributes, variant) = generated_attributes(index);
        let limits = match index % 16 {
            0 => LimitProfile::AttributesZero,
            1 => LimitProfile::AttributesOne,
            2 => LimitProfile::BytesBelow,
            3 => LimitProfile::DepthZero,
            4 => LimitProfile::EventsOne,
            5 => LimitProfile::TextZero,
            6 => LimitProfile::TokenBelow,
            _ => LimitProfile::Normal,
        };
        let name = format!("generated-{index:04}-attributes-{attributes:02}-variant-{variant:02}");
        push_case(&mut cases, name, input, limits);
    }
    cases
}

fn push_case(cases: &mut Vec<Case>, name: String, input: Vec<u8>, limits: LimitProfile) {
    let id = cases.len();
    cases.push(Case {
        id,
        name,
        input,
        limits,
    });
}

fn generated_attributes(index: usize) -> (Vec<u8>, usize, usize) {
    let mut output = Vec::with_capacity(320);
    let attributes = index % 13;
    let variant = (index / 13) % 32;
    output.extend_from_slice(b"<r");

    for attribute in 0..attributes {
        let duplicate = variant & 1 != 0 && attribute % 3 == 2;
        let name = if duplicate {
            "dup"
        } else if variant & 2 != 0 && attribute == 0 {
            "xml:space"
        } else {
            "a"
        };
        if variant & 4 != 0 && attribute % 2 == 1 {
            output.push(b'\t');
        } else if variant & 8 != 0 && attribute % 2 == 1 {
            output.extend_from_slice(b"  ");
        } else {
            output.push(b' ');
        }
        output.extend_from_slice(name.as_bytes());
        if variant & 16 != 0 && attribute == 0 {
            output.extend_from_slice(b" = ");
        } else {
            output.push(b'=');
        }
        let quote = if variant & 2 != 0 && attribute % 2 == 1 {
            b'\''
        } else {
            b'\"'
        };
        output.push(quote);
        if name == "xml:space" {
            if variant & 1 != 0 {
                output.extend_from_slice(b"default");
            } else {
                output.extend_from_slice(b"preserve");
            }
        } else if variant & 8 != 0 && attribute % 2 == 0 {
            output.extend_from_slice("雪".as_bytes());
        } else if variant & 16 != 0 && attribute % 2 == 0 {
            output.extend_from_slice(b"&amp;&quot;");
        } else {
            let value = format!("v{index:04}{attribute:02}");
            output.extend_from_slice(value.as_bytes());
        }
        output.push(quote);
    }

    if variant & 4 != 0 {
        output.push(b' ');
    }
    if variant & 16 != 0 && attributes != 0 {
        // A useful malformed/refusal ordering case: an extra attribute starts
        // but is never closed.
        output.extend_from_slice(b" broken=\"unterminated");
    }
    output.extend_from_slice(b"/>");

    if variant & 1 != 0 && attributes != 0 {
        output.extend_from_slice(b"<!--stable-comment-->");
    }
    (output, attributes, variant)
}

fn hex_digest(input: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(input);
    hex_bytes(&digest.finalize())
}

fn hex_bytes(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut result = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        result.push(HEX[(byte >> 4) as usize] as char);
        result.push(HEX[(byte & 0x0f) as usize] as char);
    }
    result
}

fn json_string(value: &str) -> String {
    let mut result = String::with_capacity(value.len() + 2);
    result.push('\"');
    for character in value.chars() {
        match character {
            '\"' => result.push_str("\\\""),
            '\\' => result.push_str("\\\\"),
            '\n' => result.push_str("\\n"),
            '\r' => result.push_str("\\r"),
            '\t' => result.push_str("\\t"),
            character if character.is_control() => {
                write!(result, "\\u{:04x}", character as u32).expect("writing String cannot fail");
            },
            character => result.push(character),
        }
    }
    result.push('\"');
    result
}

fn write_json_key_value(output: &mut String, key: &str, value: &str, first: bool) {
    if !first {
        output.push(',');
    }
    write!(output, "{}:{}", json_string(key), json_string(value))
        .expect("writing String cannot fail");
}

#[derive(Clone, Copy)]
enum BenchPolicy {
    SliceCompact,
    SliceAuthored,
    SliceSource,
    StreamCompact,
    StreamAuthored,
}

impl BenchPolicy {
    const ALL: [Self; 5] = [
        Self::SliceCompact,
        Self::SliceAuthored,
        Self::SliceSource,
        Self::StreamCompact,
        Self::StreamAuthored,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::SliceCompact => "slice/compact",
            Self::SliceAuthored => "slice/authored",
            Self::SliceSource => "slice/source",
            Self::StreamCompact => "stream/compact",
            Self::StreamAuthored => "stream/authored",
        }
    }

    const fn chunk(self) -> &'static str {
        match self {
            Self::SliceCompact | Self::SliceAuthored | Self::SliceSource => "whole",
            Self::StreamCompact | Self::StreamAuthored => "mixed",
        }
    }
}

struct BenchCase {
    name: &'static str,
    input: Vec<u8>,
}

fn run_bench() {
    const WARMUPS: usize = 5;
    const SAMPLES: usize = 100;
    const BATCH: usize = 100;

    let cases = [
        BenchCase {
            name: "zero-attributes",
            input: b"<a/>".to_vec(),
        },
        BenchCase {
            name: "one-attribute",
            input: b"<a x=\"1\"/>".to_vec(),
        },
        BenchCase {
            name: "two-attributes",
            input: b"<a a=\"1\" b=\"2\"/>".to_vec(),
        },
        BenchCase {
            name: "sixty-four-attributes",
            input: sixty_four_attributes(),
        },
        BenchCase {
            name: "xmlspace-preserve",
            input: b"<a xml:space=\"preserve\"> </a>".to_vec(),
        },
        BenchCase {
            name: "source-noncompact",
            input: b"<a  x=\"1\"/>".to_vec(),
        },
    ];

    let mut output = String::new();
    output.push('{');
    write_json_key_value(&mut output, "schema", "litchi.xml-minifier-bench.v1", true);
    write!(
        output,
        ",\"warmups\":{},\"samples\":{},\"batch_calls\":{},\"rows\":[",
        WARMUPS, SAMPLES, BATCH,
    )
    .expect("writing String cannot fail");

    let mut row_index = 0usize;
    for case in &cases {
        for policy in BenchPolicy::ALL {
            if row_index != 0 {
                output.push(',');
            }
            let limits = bench_limits(&case.input);
            let semantic = bench_semantic(policy, &case.input, limits);
            for _ in 0..WARMUPS {
                for _ in 0..BATCH {
                    bench_call(policy, &case.input, limits);
                }
            }

            let mut samples = Vec::with_capacity(SAMPLES);
            for _ in 0..SAMPLES {
                let start = Instant::now();
                for _ in 0..BATCH {
                    bench_call(policy, &case.input, limits);
                }
                samples.push(start.elapsed().as_nanos());
            }

            write!(
                output,
                "{{\"case\":{},\"policy\":{},\"chunk\":{},\"semantic_digest\":{},\"sample_ns\":[",
                json_string(case.name),
                json_string(policy.name()),
                json_string(policy.chunk()),
                json_string(&hex_digest(semantic.as_bytes())),
            )
            .expect("writing String cannot fail");
            for (sample_index, sample) in samples.iter().enumerate() {
                if sample_index != 0 {
                    output.push(',');
                }
                write!(output, "{sample}").expect("writing String cannot fail");
            }
            output.push_str("]}");
            row_index += 1;
        }
    }
    output.push_str("]}\n");
    print!("{output}");
}

fn bench_limits(input: &[u8]) -> Limits {
    Limits::new(
        input.len(),
        64,
        4_096,
        1_024,
        input.len().max(1),
        input.len(),
    )
    .expect("benchmark limit profile must stay below immutable ceilings")
}

fn bench_semantic(policy: BenchPolicy, input: &[u8], limits: Limits) -> String {
    match policy {
        BenchPolicy::SliceCompact => capture_result(|| audit::verify(input, limits)),
        BenchPolicy::SliceAuthored => capture_result(|| audit::verify_authored(input, limits)),
        BenchPolicy::SliceSource => capture_result(|| audit::verify_source(input, limits)),
        BenchPolicy::StreamCompact => {
            capture_result(|| audit::verify_reader(Chunked::new(input, CHUNKS_MIXED), limits))
        },
        BenchPolicy::StreamAuthored => capture_result(|| {
            audit::verify_authored_reader(Chunked::new(input, CHUNKS_MIXED), limits)
        }),
    }
}

fn bench_call(policy: BenchPolicy, input: &[u8], limits: Limits) {
    match policy {
        BenchPolicy::SliceCompact => {
            let _ = black_box(audit::verify(input, limits));
        },
        BenchPolicy::SliceAuthored => {
            let _ = black_box(audit::verify_authored(input, limits));
        },
        BenchPolicy::SliceSource => {
            let _ = black_box(audit::verify_source(input, limits));
        },
        BenchPolicy::StreamCompact => {
            let _ = black_box(audit::verify_reader(
                Chunked::new(input, CHUNKS_MIXED),
                limits,
            ));
        },
        BenchPolicy::StreamAuthored => {
            let _ = black_box(audit::verify_authored_reader(
                Chunked::new(input, CHUNKS_MIXED),
                limits,
            ));
        },
    }
}

fn sixty_four_attributes() -> Vec<u8> {
    let mut input = Vec::with_capacity(900);
    input.extend_from_slice(b"<r");
    for attribute in 0..64 {
        let value = format!(" a{attribute}=\"v{attribute}\"");
        input.extend_from_slice(value.as_bytes());
    }
    input.extend_from_slice(b"/>");
    input
}
