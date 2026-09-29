//! Isolated fresh-writer probe for CFB emission order and its DOC producer.
//!
//! The corpus generator is intentionally copied in small form from the
//! tools/perf-baseline writer generator. All text and payload bytes are
//! prepared before the operation interval. A timed sample includes creation
//! of the public writer, registration of the already-prepared inputs, and the
//! public write_to call. It excludes output verification, hashing, and
//! dropping the resulting byte vector.

#![forbid(unsafe_code)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::{Package, writer::Writer};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::error::Error;
use std::fmt;
use std::fmt::Write as _;
use std::fs;
use std::io::{self, Cursor, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::time::Instant;

const SCHEMA: &str = "litchi.execution-cfb-emission.v1";
const GENERATOR: &str = "tools/perf-baseline/src/lib.rs::write_fresh_doc+cfb-emission-v1";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 10_000;
const MAX_OBSERVER_EVENTS: usize = 4_096;
const LARGE_BYTES: usize = 4 * 1024 * 1024;
const DIFAT_BYTES: usize = 8 * 1024 * 1024;

type AnyResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum Case {
    DocTiny,
    DocLarge,
    DocPayload,
    CfbTiny,
    CfbLarge,
    CfbLargeOnly,
    CfbMiniOnly,
    CfbV4,
    CfbDifat,
}

impl Case {
    const ALL: [Self; 9] = [
        Self::DocTiny,
        Self::DocLarge,
        Self::DocPayload,
        Self::CfbTiny,
        Self::CfbLarge,
        Self::CfbLargeOnly,
        Self::CfbMiniOnly,
        Self::CfbV4,
        Self::CfbDifat,
    ];

    pub(crate) const fn name(self) -> &'static str {
        match self {
            Self::DocTiny => "doc-tiny",
            Self::DocLarge => "doc-large",
            Self::DocPayload => "doc-payload",
            Self::CfbTiny => "cfb-tiny",
            Self::CfbLarge => "cfb-large",
            Self::CfbLargeOnly => "cfb-large-only",
            Self::CfbMiniOnly => "cfb-mini-only",
            Self::CfbV4 => "cfb-v4",
            Self::CfbDifat => "cfb-difat",
        }
    }

    const fn is_doc(self) -> bool {
        matches!(self, Self::DocTiny | Self::DocLarge | Self::DocPayload)
    }
}

impl std::str::FromStr for Case {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::ALL
            .into_iter()
            .find(|case| case.name() == value)
            .ok_or_else(|| {
                format!(
                    "--case must be one of {}; got {value:?}",
                    Self::ALL.map(Self::name).join(", ")
                )
            })
    }
}

#[derive(Debug, Clone)]
struct Config {
    case: Case,
    samples: usize,
    warmup: usize,
    output: PathBuf,
    artifact: Option<PathBuf>,
    observe: bool,
}

#[derive(Debug, Clone)]
pub(crate) struct DocInput {
    paragraphs: Vec<String>,
}

#[derive(Debug, Clone)]
struct CfbStream {
    name: String,
    bytes: Vec<u8>,
}

#[derive(Debug, Clone)]
pub(crate) struct CfbInput {
    sector_size: usize,
    streams: Vec<CfbStream>,
}

#[derive(Debug, Clone)]
pub(crate) enum Prepared {
    Doc(DocInput),
    Cfb(CfbInput),
}

#[derive(Debug, Clone, Serialize)]
struct InputMember {
    name: String,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct InputIdentity {
    generator: &'static str,
    case: &'static str,
    format: &'static str,
    sector_size: Option<usize>,
    paragraph_count: Option<usize>,
    stream_count: Option<usize>,
    logical_input_bytes: usize,
    input_sha256: String,
    members: Vec<InputMember>,
    preparation_boundary: &'static str,
}

#[derive(Debug, Clone, Serialize)]
struct ReportConfig {
    case: &'static str,
    samples: usize,
    warmup: usize,
    measured_operation: &'static str,
    output_verification: &'static str,
}

#[derive(Debug, Clone, Serialize)]
struct Verification {
    valid_output: bool,
    output_bytes: usize,
    output_sha256: String,
    semantic_sha256: String,
    logical_output_bytes: usize,
    member_or_paragraph_count: usize,
    exact_input_match: bool,
}

#[derive(Debug, Clone, Serialize)]
struct TimedSample {
    sample: usize,
    wall_ns: u128,
    output_bytes: usize,
    output_sha256: String,
    verification: Verification,
}

#[derive(Debug, Serialize)]
struct TimedReport {
    schema: &'static str,
    mode: &'static str,
    config: ReportConfig,
    input: InputIdentity,
    samples: Vec<TimedSample>,
    final_verification: Verification,
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum ObserverEvent {
    Seek {
        from: u64,
        to: u64,
    },
    Write {
        at: u64,
        bytes: usize,
        gap_before: u64,
    },
}

#[derive(Debug, Serialize)]
struct ObserverSummary {
    write_calls: u64,
    write_bytes: u64,
    seek_calls: u64,
    backward_seek_calls: u64,
    backward_seek_bytes: u64,
    gap_events: u64,
    zero_filled_gap_bytes: u64,
    largest_gap_bytes: u64,
    output_bytes: usize,
    output_sha256: String,
    events_recorded: usize,
    events_truncated: bool,
    events: Vec<ObserverEvent>,
}

#[derive(Debug, Serialize)]
struct ObserverReport {
    schema: &'static str,
    mode: &'static str,
    config: ReportConfig,
    input: InputIdentity,
    verification: Verification,
    observer: ObserverSummary,
}

#[derive(Debug)]
struct ProbeError(String);

impl fmt::Display for ProbeError {
    fn fmt(&self, output: &mut fmt::Formatter<'_>) -> fmt::Result {
        output.write_str(&self.0)
    }
}

impl Error for ProbeError {}

fn error(message: impl Into<String>) -> Box<dyn Error + Send + Sync> {
    Box::new(ProbeError(message.into()))
}

fn usage() -> &'static str {
    "usage: cfb-emission-probe [--observe] --case CASE --samples N --warmup N --output PATH [--artifact PATH]"
}

fn parse_bounded(name: &str, value: &str, maximum: usize) -> Result<usize, String> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{name} must be a nonnegative integer, got {value:?}"))?;
    if parsed > maximum {
        return Err(format!("{name} must be at most {maximum}, got {parsed}"));
    }
    Ok(parsed)
}

fn parse_args_from<I>(args: I) -> Result<Config, String>
where
    I: IntoIterator<Item = String>,
{
    let mut args = args.into_iter();
    let mut case = None;
    let mut samples = None;
    let mut warmup = None;
    let mut output = None;
    let mut artifact = None;
    let mut observe = false;

    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--observe" if !observe => observe = true,
            "--case" if case.is_none() => {
                let value = args.next().ok_or("--case requires a value")?;
                case = Some(value.parse()?);
            },
            "--samples" if samples.is_none() => {
                let value = args.next().ok_or("--samples requires a value")?;
                samples = Some(parse_bounded("--samples", &value, MAX_SAMPLES)?);
            },
            "--warmup" if warmup.is_none() => {
                let value = args.next().ok_or("--warmup requires a value")?;
                warmup = Some(parse_bounded("--warmup", &value, MAX_WARMUP)?);
            },
            "--output" if output.is_none() => {
                output = Some(PathBuf::from(
                    args.next().ok_or("--output requires a value")?,
                ));
            },
            "--artifact" if artifact.is_none() => {
                artifact = Some(PathBuf::from(
                    args.next().ok_or("--artifact requires a value")?,
                ));
            },
            "--help" | "-h" => return Err(usage().to_string()),
            other => return Err(format!("unknown or duplicate argument {other:?}")),
        }
    }

    let samples = samples.ok_or("missing --samples")?;
    if samples == 0 {
        return Err("--samples must be at least 1".to_string());
    }
    let warmup = warmup.unwrap_or(0);
    if observe && (samples != 1 || warmup != 0) {
        return Err("--observe requires exactly --samples 1 and --warmup 0".to_string());
    }
    Ok(Config {
        case: case.ok_or("missing --case")?,
        samples,
        warmup,
        output: output.ok_or("missing --output")?,
        artifact,
        observe,
    })
}

fn parse_args() -> Result<Config, String> {
    parse_args_from(std::env::args().skip(1))
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn digest_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn sequence_sha256<'a, I>(items: I) -> String
where
    I: IntoIterator<Item = (&'a str, &'a [u8])>,
{
    let mut digest = Sha256::new();
    for (name, bytes) in items {
        digest.update(name.as_bytes());
        digest.update([0]);
        digest.update(bytes);
        digest.update([0xff]);
    }
    let mut output = String::with_capacity(64);
    for byte in digest.finalize() {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn writer_text(kind: &str, first: usize, second: usize, third: usize) -> String {
    format!(
        "litchi-perf-baseline-{kind}-v1-{first:03}-{second:05}-{third:03} deterministic payload"
    )
}

fn writer_payload_text(
    kind: &str,
    first: usize,
    second: usize,
    third: usize,
    length: usize,
) -> String {
    const REPEATED_TEXT: &str = "litchi-perf-baseline-payload-heavy-v1 ";
    let mut text = writer_text(kind, first, second, third);
    while text.len() < length {
        text.push_str(REPEATED_TEXT);
    }
    text.truncate(length);
    text
}

fn payload_bytes(length: usize, seed: usize) -> Vec<u8> {
    let mut state = (seed as u64)
        .wrapping_mul(0x9e37_79b9_7f4a_7c15)
        .wrapping_add(0xd1b5_4a32_d192_ed03);
    let mut bytes = Vec::with_capacity(length);
    for _ in 0..length {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        bytes.push((state >> 24) as u8);
    }
    bytes
}

fn small_payload(length: usize, seed: usize) -> Vec<u8> {
    const BLOCK: &[u8] = b"litchi-perf-cfb-emission-mini-payload-v1\n";
    (0..length)
        .map(|offset| BLOCK[(offset + seed) % BLOCK.len()])
        .collect()
}

fn prepared(case: Case) -> AnyResult<Prepared> {
    if case.is_doc() {
        let (count, payload_length) = match case {
            Case::DocTiny => (3, None),
            Case::DocLarge => (512, None),
            Case::DocPayload => (128, Some(20_000)),
            _ => return Err(error("non-DOC case reached DOC generator")),
        };
        let paragraphs = (0..count)
            .map(|index| {
                payload_length.map_or_else(
                    || writer_text("doc", 0, index, 0),
                    |length| writer_payload_text("doc", 0, index, 0, length),
                )
            })
            .collect();
        return Ok(Prepared::Doc(DocInput { paragraphs }));
    }

    let (sector_size, streams) = match case {
        Case::CfbTiny => (
            512,
            vec![
                CfbStream {
                    name: "MiniPayload".to_string(),
                    bytes: small_payload(512, 0),
                },
                CfbStream {
                    name: "LargePayload".to_string(),
                    bytes: payload_bytes(8 * 1024, 1),
                },
            ],
        ),
        Case::CfbLarge | Case::CfbV4 => (
            if case == Case::CfbV4 { 4096 } else { 512 },
            vec![
                CfbStream {
                    name: "MiniPayload".to_string(),
                    bytes: small_payload(2 * 1024, 0),
                },
                CfbStream {
                    name: "LargePayload".to_string(),
                    bytes: payload_bytes(LARGE_BYTES, 1),
                },
            ],
        ),
        Case::CfbLargeOnly => (
            512,
            vec![CfbStream {
                name: "LargePayload".to_string(),
                bytes: payload_bytes(LARGE_BYTES, 1),
            }],
        ),
        Case::CfbMiniOnly => (
            512,
            vec![
                CfbStream {
                    name: "MiniPayloadA".to_string(),
                    bytes: small_payload(256, 0),
                },
                CfbStream {
                    name: "MiniPayloadB".to_string(),
                    bytes: small_payload(1536, 1),
                },
                CfbStream {
                    name: "MiniPayloadC".to_string(),
                    bytes: small_payload(3072, 2),
                },
            ],
        ),
        Case::CfbDifat => (
            512,
            vec![CfbStream {
                name: "DifatPayload".to_string(),
                bytes: payload_bytes(DIFAT_BYTES, 7),
            }],
        ),
        Case::DocTiny | Case::DocLarge | Case::DocPayload => {
            return Err(error("DOC case reached CFB generator"));
        },
    };
    Ok(Prepared::Cfb(CfbInput {
        sector_size,
        streams,
    }))
}

fn input_identity(case: Case, input: &Prepared) -> InputIdentity {
    match input {
        Prepared::Doc(doc) => {
            let members = doc
                .paragraphs
                .iter()
                .enumerate()
                .map(|(index, text)| InputMember {
                    name: format!("paragraph:{index:05}"),
                    bytes: text.len(),
                    sha256: sha256_hex(text.as_bytes()),
                })
                .collect::<Vec<_>>();
            let logical_input_bytes = doc.paragraphs.iter().map(String::len).sum();
            let mut digest = Sha256::new();
            for (index, text) in doc.paragraphs.iter().enumerate() {
                digest.update(format!("paragraph:{index:05}").as_bytes());
                digest.update([0]);
                digest.update(text.as_bytes());
                digest.update([0xff]);
            }
            let input_sha256 = digest_hex(&digest.finalize());
            InputIdentity {
                generator: GENERATOR,
                case: case.name(),
                format: "DOC/CFB",
                sector_size: None,
                paragraph_count: Some(doc.paragraphs.len()),
                stream_count: None,
                logical_input_bytes,
                input_sha256,
                members,
                preparation_boundary: "paragraph text generation and hashing occur before each measured operation",
            }
        },
        Prepared::Cfb(cfb) => {
            let members = cfb
                .streams
                .iter()
                .map(|stream| InputMember {
                    name: stream.name.clone(),
                    bytes: stream.bytes.len(),
                    sha256: sha256_hex(&stream.bytes),
                })
                .collect::<Vec<_>>();
            let logical_input_bytes = cfb.streams.iter().map(|stream| stream.bytes.len()).sum();
            let input_sha256 = sequence_sha256(
                cfb.streams
                    .iter()
                    .map(|stream| (stream.name.as_str(), stream.bytes.as_slice())),
            );
            InputIdentity {
                generator: GENERATOR,
                case: case.name(),
                format: "CFB/OLE2",
                sector_size: Some(cfb.sector_size),
                paragraph_count: None,
                stream_count: Some(cfb.streams.len()),
                logical_input_bytes,
                input_sha256,
                members,
                preparation_boundary: "stream payload generation and hashing occur before each measured operation",
            }
        },
    }
}

#[allow(dead_code, reason = "used by the separate allocation companion binary")]
pub(crate) fn identity_for_case(case: Case, input: &Prepared) -> InputIdentity {
    input_identity(case, input)
}

fn write_prepared<W: Write + Seek>(input: &Prepared, output: &mut W) -> AnyResult<()> {
    match input {
        Prepared::Doc(doc) => {
            let mut writer = Writer::new();
            for paragraph in &doc.paragraphs {
                writer.add_paragraph(paragraph)?;
            }
            writer.write_to(output)?;
        },
        Prepared::Cfb(cfb) => {
            let mut writer = if cfb.sector_size == 512 {
                OleWriter::new()
            } else {
                OleWriter::with_sector_size(cfb.sector_size)?
            };
            for stream in &cfb.streams {
                writer.create_stream(&[stream.name.as_str()], &stream.bytes)?;
            }
            writer.write_to(output)?;
        },
    }
    Ok(())
}

/// Prepare one bounded corpus before an allocation region starts.
#[allow(dead_code, reason = "used by the separate allocation companion binary")]
pub(crate) fn prepare_case(case: Case) -> AnyResult<Prepared> {
    prepared(case)
}

/// Execute one fresh public-writer operation and return its owned output.
#[allow(dead_code, reason = "used by the separate allocation companion binary")]
pub(crate) fn write_case(input: &Prepared) -> AnyResult<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    write_prepared(input, &mut output)?;
    Ok(output.into_inner())
}

/// Run the untimed semantic and byte verification for an owned output.
#[allow(dead_code, reason = "used by the separate allocation companion binary")]
pub(crate) fn verify_case(input: &Prepared, bytes: &[u8]) -> AnyResult<()> {
    let verification = verify_output(input, bytes)?;
    if !verification.valid_output || !verification.exact_input_match {
        return Err(error(
            "fresh-writer verification did not establish exact input equality",
        ));
    }
    Ok(())
}

fn timed_write(input: &Prepared) -> AnyResult<(Vec<u8>, u128)> {
    let mut output = Cursor::new(Vec::new());
    let started = Instant::now();
    write_prepared(input, &mut output)?;
    let wall_ns = started.elapsed().as_nanos();
    Ok((output.into_inner(), wall_ns))
}

fn verify_output(input: &Prepared, bytes: &[u8]) -> AnyResult<Verification> {
    match input {
        Prepared::Doc(doc) => {
            let mut package = Package::from_reader(Cursor::new(bytes.to_vec()))?;
            let document = package.document()?;
            let paragraphs = document.paragraphs()?;
            if paragraphs.len() != doc.paragraphs.len() {
                return Err(error(format!(
                    "DOC paragraph count mismatch: expected {}, got {}",
                    doc.paragraphs.len(),
                    paragraphs.len()
                )));
            }
            let mut full_text = String::new();
            for (index, paragraph) in paragraphs.iter().enumerate() {
                let expected = &doc.paragraphs[index];
                if paragraph.text()? != expected.as_str() {
                    return Err(error(format!("DOC paragraph {index} differs from input")));
                }
                full_text.push_str(expected);
                full_text.push('\r');
            }
            if document.text()? != full_text {
                return Err(error("DOC public full-text projection differs from input"));
            }
            Ok(Verification {
                valid_output: true,
                output_bytes: bytes.len(),
                output_sha256: sha256_hex(bytes),
                semantic_sha256: sha256_hex(full_text.as_bytes()),
                logical_output_bytes: doc.paragraphs.iter().map(String::len).sum(),
                member_or_paragraph_count: paragraphs.len(),
                exact_input_match: true,
            })
        },
        Prepared::Cfb(cfb) => {
            let mut ole = OleFile::open(Cursor::new(bytes.to_vec()))?;
            let mut expected_paths = cfb
                .streams
                .iter()
                .map(|stream| vec![stream.name.clone()])
                .collect::<Vec<_>>();
            expected_paths.sort();
            let mut actual_paths = ole.list_streams();
            actual_paths.sort();
            if actual_paths != expected_paths {
                return Err(error(format!(
                    "CFB stream inventory differs: expected {expected_paths:?}, got {actual_paths:?}"
                )));
            }
            let mut logical_output_bytes = 0usize;
            let mut actual_sequence = Vec::with_capacity(cfb.streams.len());
            for stream in &cfb.streams {
                let refs = [stream.name.as_str()];
                let actual = ole.open_stream(&refs)?;
                if actual != stream.bytes {
                    return Err(error(format!(
                        "CFB stream {} differs from input",
                        stream.name
                    )));
                }
                logical_output_bytes += actual.len();
                actual_sequence.push((stream.name.as_str(), actual));
            }
            let semantic_sha256 = sequence_sha256(
                actual_sequence
                    .iter()
                    .map(|(name, bytes)| (*name, bytes.as_slice())),
            );
            Ok(Verification {
                valid_output: true,
                output_bytes: bytes.len(),
                output_sha256: sha256_hex(bytes),
                semantic_sha256,
                logical_output_bytes,
                member_or_paragraph_count: actual_paths.len(),
                exact_input_match: true,
            })
        },
    }
}

fn report_config(config: &Config) -> ReportConfig {
    ReportConfig {
        case: config.case.name(),
        samples: config.samples,
        warmup: config.warmup,
        measured_operation: "fresh public writer construction, prepared-input registration, and write_to",
        output_verification: "outside wall timer: reopen, semantic projection, hashes, and inventory",
    }
}

fn write_json(path: &Path, value: &impl Serialize) -> AnyResult<()> {
    let mut bytes = serde_json::to_vec_pretty(value)?;
    bytes.push(b'\n');
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)?;
    file.write_all(&bytes)?;
    file.flush()?;
    Ok(())
}

fn write_artifact(path: Option<&Path>, bytes: &[u8]) -> AnyResult<()> {
    if let Some(path) = path {
        let mut file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(path)?;
        file.write_all(bytes)?;
        file.flush()?;
    }
    Ok(())
}

fn run_timed(config: &Config, input: &Prepared, identity: InputIdentity) -> AnyResult<()> {
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or_else(|| error("warmup plus samples overflowed"))?;
    let mut samples = Vec::with_capacity(config.samples);
    let mut final_bytes = None;
    let mut final_verification = None;
    for iteration in 0..total {
        let (bytes, wall_ns) = timed_write(input)?;
        let verification = verify_output(input, &bytes)?;
        let is_final = iteration + 1 == total;
        if iteration >= config.warmup {
            samples.push(TimedSample {
                sample: iteration - config.warmup,
                wall_ns,
                output_bytes: bytes.len(),
                output_sha256: verification.output_sha256.clone(),
                verification: verification.clone(),
            });
        }
        if is_final {
            final_bytes = Some(bytes);
            final_verification = Some(verification);
        }
    }
    let final_bytes = final_bytes.ok_or_else(|| error("no timed output was produced"))?;
    let final_verification = final_verification.ok_or_else(|| error("no final verification"))?;
    write_artifact(config.artifact.as_deref(), &final_bytes)?;
    write_json(
        &config.output,
        &TimedReport {
            schema: SCHEMA,
            mode: "timed",
            config: report_config(config),
            input: identity,
            samples,
            final_verification,
        },
    )
}

struct TrackingSink {
    inner: Cursor<Vec<u8>>,
    write_calls: u64,
    write_bytes: u64,
    seek_calls: u64,
    backward_seek_calls: u64,
    backward_seek_bytes: u64,
    gap_events: u64,
    zero_filled_gap_bytes: u64,
    largest_gap_bytes: u64,
    events: Vec<ObserverEvent>,
    events_truncated: bool,
}

impl TrackingSink {
    fn new() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            write_calls: 0,
            write_bytes: 0,
            seek_calls: 0,
            backward_seek_calls: 0,
            backward_seek_bytes: 0,
            gap_events: 0,
            zero_filled_gap_bytes: 0,
            largest_gap_bytes: 0,
            events: Vec::new(),
            events_truncated: false,
        }
    }

    fn event(&mut self, event: ObserverEvent) {
        if self.events.len() < MAX_OBSERVER_EVENTS {
            self.events.push(event);
        } else {
            self.events_truncated = true;
        }
    }

    fn into_parts(self) -> (Vec<u8>, ObserverSummary) {
        let bytes = self.inner.into_inner();
        let summary = ObserverSummary {
            write_calls: self.write_calls,
            write_bytes: self.write_bytes,
            seek_calls: self.seek_calls,
            backward_seek_calls: self.backward_seek_calls,
            backward_seek_bytes: self.backward_seek_bytes,
            gap_events: self.gap_events,
            zero_filled_gap_bytes: self.zero_filled_gap_bytes,
            largest_gap_bytes: self.largest_gap_bytes,
            output_bytes: bytes.len(),
            output_sha256: sha256_hex(&bytes),
            events_recorded: self.events.len(),
            events_truncated: self.events_truncated,
            events: self.events,
        };
        (bytes, summary)
    }
}

impl Write for TrackingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let at = self.inner.stream_position()?;
        let current_length = u64::try_from(self.inner.get_ref().len()).unwrap_or(u64::MAX);
        let gap = at.saturating_sub(current_length);
        self.write_calls = self.write_calls.saturating_add(1);
        let written = self.inner.write(bytes)?;
        self.write_bytes = self
            .write_bytes
            .saturating_add(u64::try_from(written).unwrap_or(u64::MAX));
        if gap > 0 {
            self.gap_events = self.gap_events.saturating_add(1);
            self.zero_filled_gap_bytes = self.zero_filled_gap_bytes.saturating_add(gap);
            self.largest_gap_bytes = self.largest_gap_bytes.max(gap);
        }
        self.event(ObserverEvent::Write {
            at,
            bytes: written,
            gap_before: gap,
        });
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for TrackingSink {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let from = self.inner.stream_position()?;
        let to = self.inner.seek(position)?;
        self.seek_calls = self.seek_calls.saturating_add(1);
        if to < from {
            self.backward_seek_calls = self.backward_seek_calls.saturating_add(1);
            self.backward_seek_bytes = self
                .backward_seek_bytes
                .saturating_add(from.saturating_sub(to));
        }
        self.event(ObserverEvent::Seek { from, to });
        Ok(to)
    }
}

fn run_observe(config: &Config, input: &Prepared, identity: InputIdentity) -> AnyResult<()> {
    let mut sink = TrackingSink::new();
    write_prepared(input, &mut sink)?;
    let (bytes, observer) = sink.into_parts();
    let verification = verify_output(input, &bytes)?;
    write_artifact(config.artifact.as_deref(), &bytes)?;
    write_json(
        &config.output,
        &ObserverReport {
            schema: SCHEMA,
            mode: "observe",
            config: report_config(config),
            input: identity,
            verification,
            observer,
        },
    )
}

pub fn run_cli() -> AnyResult<()> {
    let config = parse_args().map_err(error)?;
    let input = prepared(config.case)?;
    let identity = input_identity(config.case, &input);
    if config.observe {
        run_observe(&config, &input, identity)
    } else {
        run_timed(&config, &input, identity)
    }
}

fn main() {
    if let Err(error) = run_cli() {
        eprintln!("cfb-emission-probe: {error}");
        std::process::exit(2);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parser_accepts_all_cases_and_bounds_samples() {
        for case in Case::ALL {
            let config = parse_args_from([
                "--case".to_string(),
                case.name().to_string(),
                "--samples".to_string(),
                "1".to_string(),
                "--warmup".to_string(),
                "0".to_string(),
                "--output".to_string(),
                "out.json".to_string(),
            ])
            .unwrap();
            assert_eq!(config.case, case);
        }
        assert!(
            parse_args_from([
                "--case".to_string(),
                "cfb-tiny".to_string(),
                "--samples".to_string(),
                (MAX_SAMPLES + 1).to_string(),
                "--output".to_string(),
                "out.json".to_string(),
            ])
            .is_err()
        );
    }

    #[test]
    fn generation_is_deterministic_and_has_expected_shape_counts() {
        let tiny = prepared(Case::DocTiny).unwrap();
        let again = prepared(Case::DocTiny).unwrap();
        assert!(matches!(&tiny, Prepared::Doc(doc) if doc.paragraphs.len() == 3));
        assert_eq!(
            input_identity(Case::DocTiny, &tiny).input_sha256,
            input_identity(Case::DocTiny, &again).input_sha256
        );

        let payload = prepared(Case::DocPayload).unwrap();
        assert!(
            matches!(&payload, Prepared::Doc(doc) if doc.paragraphs.len() == 128 && doc.paragraphs.iter().all(|text| text.len() == 20_000))
        );

        let large = prepared(Case::CfbLarge).unwrap();
        assert!(
            matches!(&large, Prepared::Cfb(cfb) if cfb.sector_size == 512 && cfb.streams.len() == 2 && cfb.streams[1].bytes.len() == LARGE_BYTES)
        );
        let difat = prepared(Case::CfbDifat).unwrap();
        assert!(
            matches!(&difat, Prepared::Cfb(cfb) if cfb.streams.len() == 1 && cfb.streams[0].bytes.len() == DIFAT_BYTES)
        );
    }

    #[test]
    fn observer_records_gap_and_backward_seek() {
        let mut sink = TrackingSink::new();
        sink.seek(SeekFrom::Start(16)).unwrap();
        sink.write_all(b"abc").unwrap();
        sink.seek(SeekFrom::Start(0)).unwrap();
        sink.write_all(b"z").unwrap();
        let (bytes, summary) = sink.into_parts();
        assert_eq!(bytes.len(), 19);
        assert_eq!(summary.zero_filled_gap_bytes, 16);
        assert_eq!(summary.gap_events, 1);
        assert_eq!(summary.backward_seek_calls, 1);
        assert_eq!(summary.backward_seek_bytes, 19);
    }

    #[test]
    fn observer_propagates_sink_failure_without_timing() {
        struct FailingSink {
            inner: Cursor<Vec<u8>>,
        }
        impl Write for FailingSink {
            fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
                Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "synthetic failure",
                ))
            }
            fn flush(&mut self) -> io::Result<()> {
                Ok(())
            }
        }
        impl Seek for FailingSink {
            fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
                self.inner.seek(position)
            }
        }
        let mut sink = FailingSink {
            inner: Cursor::new(Vec::new()),
        };
        assert!(write_prepared(&prepared(Case::CfbTiny).unwrap(), &mut sink).is_err());
    }

    #[test]
    fn tiny_observer_reopens_and_preserves_all_streams() {
        let input = prepared(Case::CfbTiny).unwrap();
        let mut sink = TrackingSink::new();
        write_prepared(&input, &mut sink).unwrap();
        let (bytes, summary) = sink.into_parts();
        let verification = verify_output(&input, &bytes).unwrap();
        assert!(verification.valid_output);
        assert_eq!(verification.output_sha256, summary.output_sha256);
        assert!(summary.write_calls > 0);
        assert!(summary.seek_calls > 0);
    }
}
