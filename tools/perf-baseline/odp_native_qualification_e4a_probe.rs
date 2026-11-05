//! External-fixture protocol adapter for the original e4a ODP owner.
//!
//! This file is copied into the pinned e4a checkout's
//! `crates/litchi-odp/examples/` directory for the paired qualification
//! build. It intentionally uses that checkout's `SourceBackedPresentation`,
//! `tail_append::Edit::plan`, and materialize/write APIs. It does not compile
//! against the current candidate API and it does not invent Patch replay.

#![forbid(unsafe_code)]

use std::{
    error::Error,
    fs,
    io::{self, Write},
    path::PathBuf,
    sync::Arc,
    time::Instant,
};

use litchi_core::{EvidenceDigest, OwnedSource, ReadAt};
use litchi_odf_common::core::{SourceContentPublicationOptions, SourceContentPublicationReport};
use litchi_odp::SourceBackedPresentation;
use soapberry_zip::office::ArchiveReader;

const SCHEMA: &str = "litchi.odp.native-qualification.legacy-probe.v1";
const ROUTES: &[&str] = &["open", "read", "write", "append"];
const DEFAULT_TITLE: &str = "Qualification appended title";
const DEFAULT_BODY: &str = "Qualification appended body";
const DEFAULT_MAX_ITERATIONS: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Mode {
    Verify,
    Calibrate,
    Timed,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "verify" => Ok(Self::Verify),
            "calibrate" => Ok(Self::Calibrate),
            "timed" => Ok(Self::Timed),
            other => Err(format!("unknown mode {other:?}")),
        }
    }
}

#[derive(Debug)]
struct Args {
    mode: Mode,
    fixture: PathBuf,
    route: Option<String>,
    title: String,
    body: String,
    iterations: usize,
    samples: usize,
    warmups: usize,
    target_ns: u128,
    max_iterations: usize,
}

impl Args {
    fn parse() -> Result<Self, String> {
        let mut args = std::env::args();
        let program = args
            .next()
            .unwrap_or_else(|| "odp_native_qualification_e4a_probe".to_owned());
        let mut mode = None;
        let mut fixture = None;
        let mut route = None;
        let mut title = DEFAULT_TITLE.to_owned();
        let mut body = DEFAULT_BODY.to_owned();
        let mut iterations = 1_usize;
        let mut samples = 1_usize;
        let mut warmups = 0_usize;
        let mut target_ns = 500_000_000_u128;
        let mut max_iterations = DEFAULT_MAX_ITERATIONS;

        while let Some(argument) = args.next() {
            match argument.as_str() {
                "--mode" => mode = Some(Mode::parse(&next(&mut args, &program, "--mode")?)?),
                "--fixture" => {
                    fixture = Some(PathBuf::from(next(&mut args, &program, "--fixture")?))
                },
                "--route" => route = Some(next(&mut args, &program, "--route")?),
                "--title" => title = next(&mut args, &program, "--title")?,
                "--body" => body = next(&mut args, &program, "--body")?,
                "--iterations" => {
                    iterations =
                        parse_positive(&next(&mut args, &program, "--iterations")?, "iterations")?
                },
                "--samples" => {
                    samples = parse_positive(&next(&mut args, &program, "--samples")?, "samples")?
                },
                "--warmups" => {
                    warmups =
                        parse_nonnegative(&next(&mut args, &program, "--warmups")?, "warmups")?
                },
                "--target-ns" => {
                    target_ns = parse_positive_u128(
                        &next(&mut args, &program, "--target-ns")?,
                        "target-ns",
                    )?
                },
                "--max-iterations" => {
                    max_iterations = parse_positive(
                        &next(&mut args, &program, "--max-iterations")?,
                        "max-iterations",
                    )?
                },
                "--help" | "-h" => return Err(usage(&program)),
                other => return Err(format!("{program:?}: unknown argument {other:?}")),
            }
        }

        let mode = mode.ok_or_else(|| format!("{program:?}: --mode is required"))?;
        let fixture = fixture.ok_or_else(|| format!("{program:?}: --fixture is required"))?;
        if let Some(route) = route.as_deref() {
            validate_route(route)?;
        }
        if matches!(mode, Mode::Calibrate | Mode::Timed) && route.is_none() {
            return Err(format!("{program:?}: {mode:?} requires --route"));
        }
        Ok(Self {
            mode,
            fixture,
            route,
            title,
            body,
            iterations,
            samples,
            warmups,
            target_ns,
            max_iterations,
        })
    }
}

fn next(
    args: &mut impl Iterator<Item = String>,
    program: &str,
    name: &str,
) -> Result<String, String> {
    args.next()
        .ok_or_else(|| format!("{program:?}: {name} requires a value"))
}

fn usage(program: &str) -> String {
    format!(
        "usage: {program} --mode verify|calibrate|timed --fixture PATH [--route ROUTE] \
         [--title TEXT --body TEXT] [--iterations N --samples N --warmups N] \
         [--target-ns N --max-iterations N]"
    )
}

fn parse_positive(value: &str, name: &str) -> Result<usize, String> {
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("{name} must be a positive integer: {error}"))?;
    if parsed == 0 {
        return Err(format!("{name} must be positive"));
    }
    Ok(parsed)
}

fn parse_nonnegative(value: &str, name: &str) -> Result<usize, String> {
    value
        .parse::<usize>()
        .map_err(|error| format!("{name} must be a nonnegative integer: {error}"))
}

fn parse_positive_u128(value: &str, name: &str) -> Result<u128, String> {
    let parsed = value
        .parse::<u128>()
        .map_err(|error| format!("{name} must be a positive integer: {error}"))?;
    if parsed == 0 {
        return Err(format!("{name} must be positive"));
    }
    Ok(parsed)
}

fn validate_route(route: &str) -> Result<(), String> {
    if ROUTES.contains(&route) {
        Ok(())
    } else {
        Err(format!(
            "legacy e4a adapter does not expose {route:?}; expected one of {}",
            ROUTES.join(", ")
        ))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Digest {
    bytes: usize,
    sha256: String,
}

impl Digest {
    fn from_bytes(bytes: &[u8]) -> Self {
        Self {
            bytes: bytes.len(),
            sha256: EvidenceDigest::of(bytes).to_string(),
        }
    }

    fn json(&self) -> String {
        format!(r#"{{"bytes":{},"sha256":"{}"}}"#, self.bytes, self.sha256)
    }
}

#[derive(Debug, Default)]
struct DigestSink {
    bytes: usize,
}

impl Write for DigestSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes = self
            .bytes
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("qualification sink byte count overflow"))?;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn source_for_fixture(bytes: &[u8]) -> Arc<dyn ReadAt> {
    Arc::new(OwnedSource::new(bytes.to_vec()))
}

fn owner(source: Arc<dyn ReadAt>) -> Result<SourceBackedPresentation, Box<dyn Error>> {
    Ok(SourceBackedPresentation::from_read_at(source)?)
}

fn fixture_content_digest(fixture: &[u8]) -> Result<Digest, Box<dyn Error>> {
    let archive = ArchiveReader::new(fixture, soapberry_zip::office::ArchiveLimits::default())?;
    Ok(Digest::from_bytes(&archive.read("content.xml")?))
}

fn write_materialized(
    source: Arc<dyn ReadAt>,
    sink: &mut impl Write,
) -> Result<usize, Box<dyn Error>> {
    let package = owner(source)?.materialize()?;
    package.write_to(&mut *sink)?;
    Ok(0)
}

fn materialized_bytes(fixture: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    let _ = write_materialized(source_for_fixture(fixture), &mut bytes)?;
    Ok(bytes)
}

fn write_append(
    source: Arc<dyn ReadAt>,
    title: &str,
    body: &str,
    sink: &mut impl Write,
) -> Result<SourceContentPublicationReport, Box<dyn Error>> {
    let edit = litchi_odp::tail_append::Edit::from_read_at(source, title, body)?;
    let options = SourceContentPublicationOptions::default();
    let plan = edit.plan(&options)?;
    Ok(plan.write_to(sink, options)?)
}

fn write_route(
    route: &str,
    source: Arc<dyn ReadAt>,
    title: &str,
    body: &str,
    sink: &mut impl Write,
) -> Result<usize, Box<dyn Error>> {
    match route {
        "open" => {
            std::hint::black_box(owner(source)?);
            Ok(0)
        },
        "read" => {
            let presentation = owner(source)?;
            let count = presentation.slide_count()?;
            let text = presentation.text()?;
            std::hint::black_box((count, text.len()));
            Ok(0)
        },
        "write" => {
            let mut sink = CountingSink::new(sink);
            let _ = write_materialized(source, &mut sink)?;
            Ok(sink.bytes)
        },
        "append" => {
            let mut sink = CountingSink::new(sink);
            let report = write_append(source, title, body, &mut sink)?;
            let bytes = usize::try_from(report.bytes())
                .map_err(|_| io::Error::other("publication byte count exceeds usize"))?;
            Ok(bytes)
        },
        other => Err(io::Error::other(format!("unknown shared route {other:?}")).into()),
    }
}

struct CountingSink<'a, W: Write + ?Sized> {
    output: &'a mut W,
    bytes: usize,
}

impl<'a, W: Write + ?Sized> CountingSink<'a, W> {
    fn new(output: &'a mut W) -> Self {
        Self { output, bytes: 0 }
    }
}

impl<W: Write + ?Sized> Write for CountingSink<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let written = self.output.write(bytes)?;
        self.bytes = self
            .bytes
            .checked_add(written)
            .ok_or_else(|| io::Error::other("qualification sink byte count overflow"))?;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.output.flush()
    }
}

fn append_bytes(fixture: &[u8], title: &str, body: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    let _ = write_append(source_for_fixture(fixture), title, body, &mut bytes)?;
    Ok(bytes)
}

fn read_receipt(fixture: &[u8]) -> Result<(usize, bool, Digest), Box<dyn Error>> {
    let presentation = owner(source_for_fixture(fixture))?;
    let slide_count = presentation.slide_count()?;
    let first_slide_present = presentation.slide(0)?.is_some();
    let text = presentation.text()?;
    Ok((
        slide_count,
        first_slide_present,
        Digest::from_bytes(text.as_bytes()),
    ))
}

fn proof(route: &str, fixture: &[u8], title: &str, body: &str) -> Result<String, Box<dyn Error>> {
    match route {
        "open" => Ok(format!(
            r#"{{"kind":"open","source":{}}}"#,
            Digest::from_bytes(fixture).json()
        )),
        "read" => {
            let (slide_count, _, text) = read_receipt(fixture)?;
            Ok(format!(
                r#"{{"kind":"read","slide_count":{},"text":{}}}"#,
                slide_count,
                text.json()
            ))
        },
        "write" => Ok(format!(
            r#"{{"kind":"package","output":{}}}"#,
            Digest::from_bytes(&materialized_bytes(fixture)?).json()
        )),
        "append" => Ok(format!(
            r#"{{"kind":"package","output":{}}}"#,
            Digest::from_bytes(&append_bytes(fixture, title, body)?).json()
        )),
        other => Err(io::Error::other(format!("unknown shared route {other:?}")).into()),
    }
}

fn timed_batch(
    route: &str,
    fixture: &[u8],
    title: &str,
    body: &str,
    iterations: usize,
) -> Result<(u128, usize), Box<dyn Error>> {
    let mut elapsed_ns = 0_u128;
    let mut output_bytes = None;
    for _ in 0..iterations {
        let source = source_for_fixture(fixture);
        let mut sink = DigestSink::default();
        let started = Instant::now();
        let observed = write_route(route, source, title, body, &mut sink)?;
        elapsed_ns = elapsed_ns
            .checked_add(started.elapsed().as_nanos())
            .ok_or_else(|| io::Error::other("qualification timing overflow"))?;
        match output_bytes {
            Some(expected) if expected != observed => {
                return Err(io::Error::other(format!(
                    "{route} output length changed within a timed batch: {expected} versus {observed}"
                ))
                .into());
            },
            None => output_bytes = Some(observed),
            _ => {},
        }
    }
    Ok((elapsed_ns, output_bytes.unwrap_or(0)))
}

fn verify(fixture: &[u8], title: &str, body: &str) -> Result<String, Box<dyn Error>> {
    let source = Digest::from_bytes(fixture);
    let content = fixture_content_digest(fixture)?;
    let (slide_count, first_slide_present, text) = read_receipt(fixture)?;
    let no_op = Digest::from_bytes(&materialized_bytes(fixture)?);
    let append = append_bytes(fixture, title, body)?;
    let append_digest = Digest::from_bytes(&append);
    let repeated_digest = Digest::from_bytes(&append_bytes(fixture, title, body)?);
    let patch_source = {
        let edit =
            litchi_odp::tail_append::Edit::from_read_at(source_for_fixture(fixture), title, body)?;
        let options = SourceContentPublicationOptions::default();
        edit.plan(&options)?.source_content_sha256()
    };
    let patch_target = {
        let edit =
            litchi_odp::tail_append::Edit::from_read_at(source_for_fixture(fixture), title, body)?;
        let options = SourceContentPublicationOptions::default();
        edit.plan(&options)?.target_content_sha256()
    };
    let source_content_sha256 = hex_bytes(&patch_source);
    let target_content_sha256 = hex_bytes(&patch_target);
    let checks = [
        ("no_op_exact_source", no_op == source),
        ("read_semantics", slide_count > 0 && first_slide_present),
        (
            "append_changed",
            source_content_sha256 != target_content_sha256 && append_digest != source,
        ),
        (
            "patch_source_matches_fixture",
            source_content_sha256 == content.sha256,
        ),
        ("append_output_changed", append_digest != source),
        (
            "independent_plan_repeat_exact",
            repeated_digest == append_digest,
        ),
    ];
    if checks.iter().any(|(_, value)| !value) {
        return Err(io::Error::other("legacy ODP correctness preflight failed").into());
    }
    let checks_json = checks
        .iter()
        .map(|(name, value)| format!(r#""{name}":{value}"#))
        .collect::<Vec<_>>()
        .join(",");
    Ok(format!(
        concat!(
            r#"{{"schema":"{}","mode":"verify","capabilities":{{"patch_replay":false,"inverse_replay":false,"scoped_replay":false,"timed_routes":["open","read","write","append"],"replay_scope":"independent_plan_not_patch"}},"#,
            r#""fixture":{{"bytes":{},"sha256":"{}","content":{}}},"#,
            r#""checks":{{{}}},"#,
            r#""read":{{"slide_count":{},"first_slide_present":{},"text_bytes":{},"text_sha256":"{}"}},"#,
            r#""patch":{{"is_noop":false,"source_content_sha256":"{}","target_content_sha256":"{}"}},"#,
            r#""outputs":{{"source":{},"no_op":{},"append":{},"independent_append":{}}}}}"#
        ),
        SCHEMA,
        source.bytes,
        source.sha256,
        content.json(),
        checks_json,
        slide_count,
        first_slide_present,
        text.bytes,
        text.sha256,
        source_content_sha256,
        target_content_sha256,
        source.json(),
        no_op.json(),
        append_digest.json(),
        repeated_digest.json(),
    ))
}

fn calibrate(args: &Args, fixture: &[u8]) -> Result<String, Box<dyn Error>> {
    let route = args.route.as_deref().ok_or("calibrate route is missing")?;
    let proof = proof(route, fixture, &args.title, &args.body)?;
    let mut iterations = 1_usize;
    let mut attempts = Vec::new();
    loop {
        let (elapsed_ns, output_bytes) =
            timed_batch(route, fixture, &args.title, &args.body, iterations)?;
        attempts.push(format!(
            r#"{{"iterations":{},"elapsed_ns":{},"output_bytes":{}}}"#,
            iterations, elapsed_ns, output_bytes
        ));
        if elapsed_ns >= args.target_ns {
            return Ok(format!(
                r#"{{"schema":"{SCHEMA}","mode":"calibrate","route":"{route}","proof":{proof},"target_ns":{},"iterations":{},"elapsed_ns":{},"output_bytes":{},"attempts":[{}]}}"#,
                args.target_ns,
                iterations,
                elapsed_ns,
                output_bytes,
                attempts.join(",")
            ));
        }
        if iterations >= args.max_iterations {
            return Err(io::Error::other("legacy calibration reached max-iterations").into());
        }
        iterations = iterations.saturating_mul(2).min(args.max_iterations);
    }
}

fn timed(args: &Args, fixture: &[u8]) -> Result<String, Box<dyn Error>> {
    let route = args.route.as_deref().ok_or("timed route is missing")?;
    let proof = proof(route, fixture, &args.title, &args.body)?;
    for _ in 0..args.warmups {
        let _ = timed_batch(route, fixture, &args.title, &args.body, args.iterations)?;
    }
    let mut samples = Vec::with_capacity(args.samples);
    let mut output_bytes = None;
    for _ in 0..args.samples {
        let (elapsed_ns, bytes) =
            timed_batch(route, fixture, &args.title, &args.body, args.iterations)?;
        if elapsed_ns == 0 {
            return Err(io::Error::other("legacy timing returned zero nanoseconds").into());
        }
        if let Some(expected) = output_bytes {
            if expected != bytes {
                return Err(io::Error::other("legacy timed output length changed").into());
            }
        } else {
            output_bytes = Some(bytes);
        }
        samples.push(elapsed_ns);
    }
    Ok(format!(
        r#"{{"schema":"{SCHEMA}","mode":"timed","route":"{route}","proof":{proof},"iterations":{},"warmups":{},"samples":[{}],"output_bytes":{}}}"#,
        args.iterations,
        args.warmups,
        samples
            .iter()
            .map(u128::to_string)
            .collect::<Vec<_>>()
            .join(","),
        output_bytes.unwrap_or(0),
    ))
}

fn hex_bytes(bytes: &[u8; 32]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn run(args: Args) -> Result<String, Box<dyn Error>> {
    let fixture = fs::read(&args.fixture)?;
    match args.mode {
        Mode::Verify => verify(&fixture, &args.title, &args.body),
        Mode::Calibrate => calibrate(&args, &fixture),
        Mode::Timed => timed(&args, &fixture),
    }
}

fn main() {
    let result = match Args::parse() {
        Ok(args) => run(args),
        Err(error) => Err(io::Error::other(error).into()),
    };
    match result {
        Ok(report) => println!("{report}"),
        Err(error) => {
            eprintln!("legacy ODP qualification probe failed: {error}");
            std::process::exit(1);
        },
    }
}
