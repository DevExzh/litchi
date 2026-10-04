//! Deterministic semantic oracle for the attribute-summary experiment.
//!
//! This example deliberately does not assert a timing result.  It emits a
//! JSON-lines record for every audit invocation so that a binary built from
//! the unchanged tree and a binary built from a candidate tree can be
//! compared byte for byte.  The records retain the exact `Debug` rendering
//! of each public result, its stable `Display` text, and a digest of those two
//! strings.

#![allow(
    clippy::format_push_string,
    clippy::print_stdout,
    clippy::too_many_lines,
    reason = "this standalone example is an intentionally exhaustive oracle"
)]

use std::cmp::min;
use std::io::{self, BufRead, BufWriter, Write};

use xml_minifier::audit::{
    self, Error, Limits, ReplacementError, Report, Resource, StreamError, verify, verify_authored,
    verify_authored_reader, verify_reader, verify_source, verify_source_replacement,
};

const SCHEMA: &str = "litchi.xml-attribute-summary-parity-0905.v1";
const CASE_COUNT: usize = 11_520;
const CARDINALITIES: [usize; 6] = [0, 1, 2, 31, 32, 33];
const PROFILE_COUNT: usize = 15;
const STYLE_COUNT: usize = 128;

/// A caller-owned source that exposes at most `width` bytes per `fill_buf`.
/// `quick-xml` must therefore observe the same XML through one-byte, uneven,
/// and nearly-whole-input buffers.
struct Chunked {
    bytes: Vec<u8>,
    position: usize,
    width: usize,
}

impl Chunked {
    fn new(bytes: &[u8], width: usize) -> Self {
        Self {
            bytes: bytes.to_vec(),
            position: 0,
            width: width.max(1),
        }
    }
}

impl std::io::Read for Chunked {
    fn read(&mut self, destination: &mut [u8]) -> io::Result<usize> {
        if self.position == self.bytes.len() || destination.is_empty() {
            return Ok(0);
        }
        let end = min(
            self.bytes.len(),
            self.position
                .saturating_add(min(self.width, destination.len())),
        );
        let amount = end - self.position;
        destination[..amount].copy_from_slice(&self.bytes[self.position..end]);
        self.position = end;
        Ok(amount)
    }
}

impl BufRead for Chunked {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        let end = min(self.bytes.len(), self.position.saturating_add(self.width));
        Ok(&self.bytes[self.position..end])
    }

    fn consume(&mut self, amount: usize) {
        self.position = min(self.bytes.len(), self.position.saturating_add(amount));
    }
}

#[derive(Clone, Copy)]
enum Profile {
    Default,
    AttributesExact,
    AttributesUnder,
    ElementAttributesExact,
    ElementAttributesUnder,
    DepthExact,
    DepthUnder,
    EventsExact,
    EventsUnder,
    TextExact,
    TextUnder,
    BytesExact,
    BytesUnder,
    TokenExact,
    TokenUnder,
}

impl Profile {
    fn from_index(index: usize) -> Self {
        match index % PROFILE_COUNT {
            0 => Self::Default,
            1 => Self::AttributesExact,
            2 => Self::AttributesUnder,
            3 => Self::ElementAttributesExact,
            4 => Self::ElementAttributesUnder,
            5 => Self::DepthExact,
            6 => Self::DepthUnder,
            7 => Self::EventsExact,
            8 => Self::EventsUnder,
            9 => Self::TextExact,
            10 => Self::TextUnder,
            11 => Self::BytesExact,
            12 => Self::BytesUnder,
            13 => Self::TokenExact,
            _ => Self::TokenUnder,
        }
    }

    fn is_default(self) -> bool {
        matches!(self, Self::Default)
    }

    fn name(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::AttributesExact => "attributes-exact",
            Self::AttributesUnder => "attributes-under",
            Self::ElementAttributesExact => "element-attributes-exact",
            Self::ElementAttributesUnder => "element-attributes-under",
            Self::DepthExact => "depth-exact",
            Self::DepthUnder => "depth-under",
            Self::EventsExact => "events-exact",
            Self::EventsUnder => "events-under",
            Self::TextExact => "text-exact",
            Self::TextUnder => "text-under",
            Self::BytesExact => "bytes-exact",
            Self::BytesUnder => "bytes-under",
            Self::TokenExact => "token-exact",
            Self::TokenUnder => "token-under",
        }
    }

    fn limits(self, baseline: Option<&Report>, case: &Case) -> Limits {
        let report_value = |read: fn(Report) -> usize, fallback: usize| {
            baseline.map_or(fallback, |report| read(*report))
        };
        let narrow = |resource: Resource, value: usize| Limits::default().narrow(resource, value);
        match self {
            Self::Default => Limits::default(),
            Self::AttributesExact => narrow(
                Resource::Attributes,
                report_value(Report::attributes, case.attributes),
            ),
            Self::AttributesUnder => narrow(
                Resource::Attributes,
                report_value(Report::attributes, case.attributes).saturating_sub(1),
            ),
            Self::ElementAttributesExact => narrow(Resource::ElementAttributes, case.attributes),
            Self::ElementAttributesUnder => narrow(
                Resource::ElementAttributes,
                case.attributes.saturating_sub(1),
            ),
            Self::DepthExact => {
                narrow(Resource::Depth, report_value(Report::max_depth, case.depth))
            },
            Self::DepthUnder => narrow(
                Resource::Depth,
                report_value(Report::max_depth, case.depth).saturating_sub(1),
            ),
            Self::EventsExact => {
                narrow(Resource::Events, report_value(Report::events, case.events))
            },
            Self::EventsUnder => narrow(
                Resource::Events,
                report_value(Report::events, case.events).saturating_sub(1),
            ),
            Self::TextExact => narrow(
                Resource::TextBytes,
                report_value(Report::text_bytes, case.text_bytes),
            ),
            Self::TextUnder => narrow(
                Resource::TextBytes,
                report_value(Report::text_bytes, case.text_bytes).saturating_sub(1),
            ),
            Self::BytesExact => narrow(Resource::Bytes, report_value(Report::bytes, case.bytes)),
            Self::BytesUnder => narrow(
                Resource::Bytes,
                report_value(Report::bytes, case.bytes).saturating_sub(1),
            ),
            Self::TokenExact => narrow(Resource::TokenBytes, case.token_bytes),
            Self::TokenUnder => narrow(Resource::TokenBytes, case.token_bytes.saturating_sub(1)),
        }
    }
}

#[derive(Clone)]
struct Case {
    id: usize,
    cardinality: usize,
    style: usize,
    original: Vec<u8>,
    replacement: Vec<u8>,
    replacement_kind: &'static str,
    attributes: usize,
    bytes: usize,
    depth: usize,
    events: usize,
    text_bytes: usize,
    token_bytes: usize,
}

struct Outcome {
    debug: String,
    text: String,
    hash: String,
}

impl Outcome {
    fn new(debug: String, text: String) -> Self {
        let mut material = Vec::with_capacity(debug.len() + text.len() + 1);
        material.extend_from_slice(debug.as_bytes());
        material.push(0);
        material.extend_from_slice(text.as_bytes());
        Self {
            debug,
            text,
            hash: fnv64(&material),
        }
    }
}

fn report_text(report: Report) -> String {
    format!(
        "ok attributes={} bytes={} events={} max_depth={} text_bytes={}",
        report.attributes(),
        report.bytes(),
        report.events(),
        report.max_depth(),
        report.text_bytes()
    )
}

fn report_outcome(result: Result<Report, Error>) -> Outcome {
    let debug = format!("{result:?}");
    let text = match result {
        Ok(report) => report_text(report),
        Err(error) => error.to_string(),
    };
    Outcome::new(debug, text)
}

fn stream_outcome(result: Result<Report, StreamError>) -> Outcome {
    let debug = format!("{result:?}");
    let text = match result {
        Ok(report) => report_text(report),
        Err(error) => error.to_string(),
    };
    Outcome::new(debug, text)
}

fn replacement_outcome(result: Result<audit::ReplacementProof, ReplacementError>) -> Outcome {
    let debug = format!("{result:?}");
    let text = match result {
        Ok(proof) => format!("ok {proof:?}"),
        Err(error) => error.to_string(),
    };
    Outcome::new(debug, text)
}

fn json_quote(value: &str) -> String {
    let mut quoted = String::with_capacity(value.len() + 2);
    quoted.push('"');
    for character in value.chars() {
        match character {
            '"' => quoted.push_str("\\\""),
            '\\' => quoted.push_str("\\\\"),
            '\n' => quoted.push_str("\\n"),
            '\r' => quoted.push_str("\\r"),
            '\t' => quoted.push_str("\\t"),
            character if character.is_control() => {
                quoted.push_str(&format!("\\u{:04x}", character as u32));
            },
            character => quoted.push(character),
        }
    }
    quoted.push('"');
    quoted
}

fn fnv64(bytes: &[u8]) -> String {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for &byte in bytes {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    format!("{hash:016x}")
}

fn emit<W: Write>(
    writer: &mut W,
    case: &Case,
    label: &str,
    api: &str,
    limits: Limits,
    input: &[u8],
    replacement: Option<&[u8]>,
    outcome: Outcome,
) -> io::Result<()> {
    let limits_debug = format!("{limits:?}");
    let input_hash = fnv64(input);
    let replacement_field =
        replacement.map_or_else(|| "null".to_owned(), |bytes| json_quote(&fnv64(bytes)));
    writeln!(
        writer,
        "{{\"schema\":{},\"case\":{},\"case_label\":{},\"api\":{},\"limits\":{},\"input_hash\":{},\"replacement_hash\":{},\"debug\":{},\"text\":{},\"hash\":{}}}",
        json_quote(SCHEMA),
        case.id,
        json_quote(label),
        json_quote(api),
        json_quote(&limits_debug),
        json_quote(&input_hash),
        replacement_field,
        json_quote(&outcome.debug),
        json_quote(&outcome.text),
        json_quote(&outcome.hash),
    )
}

fn separator(mode: usize) -> &'static str {
    match mode {
        8 => "\t",
        9 => "\u{a0}",
        10 => "  ",
        _ => " ",
    }
}

fn quote(style: usize) -> char {
    if style & 1 == 0 { '"' } else { '\'' }
}

fn push_attribute(attributes: &mut String, separator: &str, key: &str, value: &str, mark: char) {
    attributes.push_str(separator);
    attributes.push_str(key);
    attributes.push('=');
    attributes.push(mark);
    attributes.push_str(value);
    attributes.push(mark);
}

fn make_attributes(count: usize, style: usize) -> String {
    if count == 0 {
        return String::new();
    }
    let mode = style & 0x0f;
    let mark = quote(style);
    let spacing = separator(mode);
    let mut attributes = String::new();
    for index in 0..count {
        let ordinary_key = format!("a{index}");
        let (key, mut value) = match mode {
            2 => {
                if index == 0 {
                    ("xml:space".to_owned(), "preserve".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            3 => {
                if index == 0 {
                    ("xml:space".to_owned(), "default".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            4 => {
                if index == 0 {
                    let value = if style & 2 == 0 {
                        "maybe"
                    } else {
                        "pre&#120;erve"
                    };
                    ("xml:space".to_owned(), value.to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            5 => {
                if index == 0 {
                    ("xmlns:x".to_owned(), "urn:x".to_owned())
                } else if index == 1 {
                    ("x:k".to_owned(), "1".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            6 => {
                if index == 0 {
                    ("x:k".to_owned(), "1".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            7 => {
                if index == 0 {
                    ("=bad".to_owned(), "1".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            11 => {
                if index == 0 {
                    (ordinary_key, "unterminated".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            12 => {
                if index == 0 {
                    (ordinary_key, "<bad>".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            13 => {
                if index == 0 {
                    (ordinary_key, "&bogus;".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            14 => {
                if index + 1 == count {
                    ("a0".to_owned(), "2".to_owned())
                } else {
                    (ordinary_key, "1".to_owned())
                }
            },
            15 => {
                if index == 0 {
                    let key = if style & 2 == 0 { "naïve" } else { "xml:lang" };
                    (key.to_owned(), "a&amp;b&gt;c".to_owned())
                } else {
                    (format!("a.{index}"), "é".to_owned())
                }
            },
            _ => (ordinary_key, "1".to_owned()),
        };

        if mode == 0 && index == 0 && style & 2 != 0 {
            value.clear();
        }

        if mode == 10 && index == 0 {
            attributes.push_str(spacing);
            attributes.push_str(&key);
            if style & 2 == 0 {
                attributes.push_str(" = ");
                attributes.push(mark);
                attributes.push_str(&value);
                attributes.push(mark);
            } else {
                attributes.push_str("=unquoted");
            }
        } else if mode == 11 && index == 0 {
            attributes.push_str(spacing);
            attributes.push_str(&key);
            attributes.push('=');
            attributes.push(mark);
            attributes.push_str(&value);
            // Deliberately omit the closing quote; the root close becomes
            // part of the value and the parser reports the unterminated tag.
        } else {
            push_attribute(&mut attributes, spacing, &key, &value, mark);
        }
    }
    attributes
}

fn body(style: usize, id: usize) -> Vec<u8> {
    match (style >> 4) & 0x07 {
        1 => format!("é中{id}").into_bytes(),
        2 => b" ".to_vec(),
        3 => b"\n\t".to_vec(),
        4 => b"a&amp;b".to_vec(),
        5 => b"<![CDATA[c]]>".to_vec(),
        6 => b"]]>".to_vec(),
        7 => b"<bad".to_vec(),
        _ => format!("v{id}").into_bytes(),
    }
}

fn document(
    id: usize,
    style: usize,
    cardinality: usize,
) -> (Vec<u8>, usize, usize, usize, usize, usize) {
    let attributes = make_attributes(cardinality, style);
    let empty_root = style % 16 == 0;
    let declaration = if style & 0x80 != 0 {
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>".to_vec()
    } else {
        Vec::new()
    };
    let body = body(style, id);
    let root_open = if attributes.is_empty() {
        b"<root>".to_vec()
    } else {
        format!("<root{attributes}>").into_bytes()
    };
    let mut bytes = Vec::with_capacity(declaration.len() + 64);
    if id % 53 == 0 {
        bytes.extend_from_slice(b"\xef\xbb\xbf");
    }
    bytes.extend_from_slice(&declaration);
    bytes.extend_from_slice(&root_open);
    if empty_root {
        // Replace the normal `>` with `/>` for an empty-element root.
        let last = bytes.len().saturating_sub(1);
        bytes.insert(last, b'/');
    } else {
        bytes.extend_from_slice(b"<cell>");
        bytes.extend_from_slice(&body);
        bytes.extend_from_slice(b"</cell></root>");
    }

    // A non-ASCII or syntactically malformed attribute mode is enough to
    // exercise the parser's raw-key and tail behavior.  Keep the explicit
    // invalid UTF-8 cases separate so they cannot be repaired by a text path.
    if id % 211 == 0 && !empty_root {
        let insertion = bytes.len().saturating_sub(14);
        bytes.insert(insertion, 0xff);
    }

    let text_bytes = body.len();
    let depth = if empty_root { 1 } else { 2 };
    let events = if empty_root {
        2
    } else if body.is_empty() {
        4
    } else {
        5
    };
    let malformed = (style & 0x0f) >= 9 || ((style >> 4) & 0x07) >= 6;
    let token_bytes = if malformed {
        bytes.len()
    } else {
        let declaration_bytes = if style & 0x80 != 0 {
            b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>".len()
        } else {
            0
        };
        declaration_bytes
            .max(root_open.len())
            .max(body.len())
            .max(7)
            .max(14)
    };
    (bytes, cardinality, text_bytes, depth, events, token_bytes)
}

fn replace_once(bytes: &[u8], needle: &[u8], replacement: &[u8]) -> Option<Vec<u8>> {
    let start = bytes
        .windows(needle.len())
        .position(|window| window == needle)?;
    let mut result = Vec::with_capacity(bytes.len() + replacement.len());
    result.extend_from_slice(&bytes[..start]);
    result.extend_from_slice(replacement);
    result.extend_from_slice(&bytes[start + needle.len()..]);
    Some(result)
}

fn make_case(id: usize) -> Case {
    let cardinality = CARDINALITIES[(id / PROFILE_COUNT) % CARDINALITIES.len()];
    let style = (id / (PROFILE_COUNT * CARDINALITIES.len())) % STYLE_COUNT;
    let (mut original, attributes, text_bytes, depth, events, token_bytes) =
        document(id, style, cardinality);
    let normal_replacement =
        replace_once(&original, b"<cell>", b"<cell>w").unwrap_or_else(|| original.clone());
    let (replacement, replacement_kind) = match (id / PROFILE_COUNT) % 8 {
        0 => (original.clone(), "identical"),
        1 => (
            replace_once(&original, b"<cell>", b"<cell><inner>")
                .and_then(|value| replace_once(&value, b"</cell>", b"</inner></cell>"))
                .unwrap_or_else(|| normal_replacement.clone()),
            "window-nested",
        ),
        2 => (normal_replacement, "window-text"),
        3 => (
            replace_once(&original, b"<cell>", b"<cell>&bogus;")
                .unwrap_or_else(|| original.clone()),
            "replacement-failure",
        ),
        4 => (
            replace_once(&original, b"<cell>", b"<cell extra=\"1\">")
                .unwrap_or_else(|| original.clone()),
            "window-retotal-attributes",
        ),
        5 => (
            b"<other><different>replacement</different></other>".to_vec(),
            "complete-fallback",
        ),
        6 => {
            let valid = original.clone();
            original = b"<!DOCTYPE root><root><cell>broken</root>".to_vec();
            (valid, "original-failure")
        },
        _ => {
            let valid = original.clone();
            original = b"<root><cell>broken</root>".to_vec();
            (valid, "original-malformed")
        },
    };
    let bytes = original.len();
    let token_bytes = token_bytes.max(1);
    Case {
        id,
        cardinality,
        style,
        original,
        replacement,
        replacement_kind,
        attributes,
        bytes,
        depth,
        events,
        text_bytes,
        token_bytes,
    }
}

fn case_label(case: &Case, profile: Profile) -> String {
    format!(
        "cardinality={};style={};profile={};replacement={}",
        case.cardinality,
        case.style,
        profile.name(),
        case.replacement_kind
    )
}

fn main() -> io::Result<()> {
    let stdout = io::stdout();
    let mut writer = BufWriter::new(stdout.lock());
    writeln!(
        writer,
        "{{\"schema\":{},\"cases\":{},\"profiles\":{},\"styles\":{},\"cardinalities\":[0,1,2,31,32,33],\"reader_chunks\":[\"one\",\"uneven\",\"prefix\"]}}",
        json_quote(SCHEMA),
        CASE_COUNT,
        PROFILE_COUNT,
        STYLE_COUNT,
    )?;

    let mut records = 0usize;
    for id in 0..CASE_COUNT {
        let case = make_case(id);
        let default_limits = Limits::default();
        let default_result = verify_source(&case.original, default_limits);
        let profile = Profile::from_index(id);
        let limits = profile.limits(default_result.as_ref().ok(), &case);
        let label = case_label(&case, profile);

        emit(
            &mut writer,
            &case,
            &label,
            "slice/source/default",
            default_limits,
            &case.original,
            None,
            report_outcome(default_result),
        )?;
        records += 1;

        if !profile.is_default() {
            emit(
                &mut writer,
                &case,
                &label,
                "slice/source/profile",
                limits,
                &case.original,
                None,
                report_outcome(verify_source(&case.original, limits)),
            )?;
            records += 1;
        }

        emit(
            &mut writer,
            &case,
            &label,
            "slice/authored",
            limits,
            &case.original,
            None,
            report_outcome(verify_authored(&case.original, limits)),
        )?;
        emit(
            &mut writer,
            &case,
            &label,
            "slice/compact",
            limits,
            &case.original,
            None,
            report_outcome(verify(&case.original, limits)),
        )?;
        records += 2;

        let widths = [
            ("one", 1usize),
            ("uneven", 2 + id % 7),
            ("prefix", case.original.len().saturating_sub(1).max(1)),
        ];
        for (chunk_name, width) in widths {
            emit(
                &mut writer,
                &case,
                &label,
                &format!("reader/authored/{chunk_name}"),
                limits,
                &case.original,
                None,
                stream_outcome(verify_authored_reader(
                    Chunked::new(&case.original, width),
                    limits,
                )),
            )?;
            emit(
                &mut writer,
                &case,
                &label,
                &format!("reader/compact/{chunk_name}"),
                limits,
                &case.original,
                None,
                stream_outcome(verify_reader(Chunked::new(&case.original, width), limits)),
            )?;
            records += 2;
        }

        emit(
            &mut writer,
            &case,
            &label,
            "replacement/source",
            limits,
            &case.original,
            Some(&case.replacement),
            replacement_outcome(verify_source_replacement(
                &case.original,
                &case.replacement,
                limits,
            )),
        )?;
        records += 1;
    }

    writeln!(
        writer,
        "{{\"schema\":{},\"complete\":true,\"cases\":{},\"records\":{}}}",
        json_quote(SCHEMA),
        CASE_COUNT,
        records,
    )?;
    writer.flush()
}
