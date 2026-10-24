//! Matched native ODP qualification probe.
//!
//! The probe deliberately keeps fixture preparation outside each timed
//! lifecycle.  Every timed iteration creates a fresh positional source from
//! the frozen fixture bytes, then measures only the requested ODP operation:
//! opening, semantic reading, unchanged writing, appending, or
//! content-matching replay.  Replay is a candidate-only correctness route;
//! the external ABBA runner times only the four routes shared with the
//! pre-Patch baseline.  The `verify` mode is untimed and must pass before the
//! runner is allowed to collect timings.

#![forbid(unsafe_code)]

use std::{
    error::Error,
    fs,
    io::{self, Write},
    path::PathBuf,
    sync::Arc,
    time::Instant,
};

use litchi_core::{OwnedSource, ReadAt};
use litchi_odp::SourceBackedPresentation;
use serde_json::{Value, json};
use sha2::{Digest as Sha2Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const SCHEMA: &str = "litchi.odp.native-qualification.probe.v1";
const ROUTES: &[&str] = &["open", "read", "write", "append", "replay"];
const DEFAULT_TITLE: &str = "Qualification appended title";
const DEFAULT_BODY: &str = "Qualification appended body";
const DEFAULT_MAX_ITERATIONS: usize = 1_000_000;
const SCOPED_REPLAY_FALLBACK_MEMBER: &str = "Qualification/unrelated.bin";
const SCOPED_REPLAY_MARKER: &[u8] =
    b"\n<!-- litchi native qualification unrelated-member variant -->\n";
const SCOPED_REPLAY_PROVENANCE: &str = "deterministic generated recipient: one unrelated archive member changed; content.xml unchanged";

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
        let mut args = std::env::args_os();
        let program = args
            .next()
            .unwrap_or_else(|| "odp_native_qualification_probe".into());
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
            let argument = argument
                .into_string()
                .map_err(|_| format!("{program:?}: arguments must be UTF-8"))?;
            let mut value = |name: &str| {
                args.next()
                    .ok_or_else(|| format!("{program:?}: {name} requires a value"))
                    .and_then(|value| {
                        value
                            .into_string()
                            .map_err(|_| format!("{program:?}: {name} must be UTF-8"))
                    })
            };
            match argument.as_str() {
                "--mode" => mode = Some(Mode::parse(&value("--mode")?)?),
                "--fixture" => fixture = Some(PathBuf::from(value("--fixture")?)),
                "--route" => route = Some(value("--route")?),
                "--title" => title = value("--title")?,
                "--body" => body = value("--body")?,
                "--iterations" => {
                    iterations = parse_positive(&value("--iterations")?, "iterations")?
                },
                "--samples" => samples = parse_positive(&value("--samples")?, "samples")?,
                "--warmups" => warmups = parse_nonnegative(&value("--warmups")?, "warmups")?,
                "--target-ns" => {
                    target_ns = parse_positive_u128(&value("--target-ns")?, "target-ns")?
                },
                "--max-iterations" => {
                    max_iterations = parse_positive(&value("--max-iterations")?, "max-iterations")?
                },
                "--help" | "-h" => return Err(usage(&program.to_string_lossy())),
                other => return Err(format!("{program:?}: unknown argument {other:?}")),
            }
        }

        let mode = mode.ok_or_else(|| format!("{program:?}: --mode is required"))?;
        let fixture = fixture.ok_or_else(|| format!("{program:?}: --fixture is required"))?;
        if max_iterations == 0 {
            return Err("max-iterations must be positive".to_owned());
        }
        if matches!(mode, Mode::Verify) && (iterations != 1 || samples != 1 || warmups != 0) {
            return Err("verify does not accept timing controls".to_owned());
        }
        if matches!(mode, Mode::Calibrate) && route.is_none() {
            return Err("calibrate requires --route".to_owned());
        }
        if matches!(mode, Mode::Timed) && route.is_none() {
            return Err("timed requires --route".to_owned());
        }
        if let Some(route) = route.as_deref() {
            validate_route(route)?;
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

fn usage(program: &str) -> String {
    format!(
        "usage: {program} --mode verify|calibrate|timed --fixture PATH [--route ROUTE] \\
         [--title TEXT --body TEXT] [--iterations N --samples N --warmups N] \\
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
            "unknown route {route:?}; expected one of {}",
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
            sha256: sha256_hex(bytes),
        }
    }

    fn json(&self) -> Value {
        json!({"bytes": self.bytes, "sha256": self.sha256})
    }
}

#[derive(Debug, Default)]
struct CountingSink {
    bytes: usize,
}

impl Write for CountingSink {
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

/// A bounded-output sink used by the untimed correctness preflight.
///
/// Verification needs exact output identity, but retaining a complete package
/// would turn a correctness receipt into an avoidable memory benchmark.  The
/// digest and byte count are sufficient to prove equality with the frozen
/// fixture and with every replay result.
#[derive(Debug)]
struct DigestSink {
    digest: Sha256,
    bytes: usize,
}

impl DigestSink {
    fn new() -> Self {
        Self {
            digest: Sha256::new(),
            bytes: 0,
        }
    }

    fn finish(self) -> Digest {
        Digest {
            bytes: self.bytes,
            sha256: hex_bytes(&self.digest.finalize()),
        }
    }
}

impl Write for DigestSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes = self
            .bytes
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("qualification digest byte count overflow"))?;
        self.digest.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn fixture_content_digest(fixture: &[u8]) -> Result<Digest, Box<dyn Error>> {
    let archive = ArchiveReader::new(fixture, soapberry_zip::office::ArchiveLimits::default())?;
    let content = archive.read("content.xml")?;
    Ok(Digest::from_bytes(&content))
}

fn member_digest(bytes: &[u8], name: &str) -> Result<Digest, Box<dyn Error>> {
    let archive = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())?;
    let member = archive.read(name)?;
    Ok(Digest::from_bytes(&member))
}

fn scoped_recipient_fixture(
    fixture: &[u8],
) -> Result<(Vec<u8>, String, Option<Digest>), Box<dyn Error>> {
    let reader = ArchiveReader::new(fixture, soapberry_zip::office::ArchiveLimits::default())?;
    let mut names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    names.retain(|name| !name.ends_with('/'));
    let target = names
        .iter()
        .find(|name| name.as_str() == "styles.xml")
        .or_else(|| {
            names.iter().find(|name| {
                name.as_str() != "mimetype"
                    && name.as_str() != "content.xml"
                    && name.as_str() != "META-INF/manifest.xml"
            })
        })
        .cloned()
        .unwrap_or_else(|| SCOPED_REPLAY_FALLBACK_MEMBER.to_owned());
    let source_member = names
        .iter()
        .find(|name| name.as_str() == target)
        .map(|name| member_digest(fixture, name.as_str()))
        .transpose()?;

    let mut ordered = Vec::with_capacity(names.len());
    if names.iter().any(|name| name == "mimetype") {
        ordered.push("mimetype".to_owned());
    }
    ordered.extend(names.into_iter().filter(|name| name != "mimetype"));
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    let mut changed_existing = false;
    for name in ordered {
        let mut bytes = reader.read(&name)?;
        if name == target {
            bytes.extend_from_slice(SCOPED_REPLAY_MARKER);
            changed_existing = true;
        }
        if name == "mimetype" {
            writer.write_stored(&name, &bytes)?;
        } else {
            writer.write_deflated(&name, &bytes)?;
        }
    }
    if !changed_existing {
        writer.write_stored(SCOPED_REPLAY_FALLBACK_MEMBER, SCOPED_REPLAY_MARKER)?;
    }
    Ok((writer.finish_to_bytes()?, target, source_member))
}

fn source_for_fixture(bytes: &[u8]) -> Arc<dyn ReadAt> {
    Arc::new(OwnedSource::new(bytes.to_vec()))
}

fn owner(bytes: &[u8]) -> Result<SourceBackedPresentation, Box<dyn Error>> {
    owner_source(source_for_fixture(bytes))
}

fn owner_source(source: Arc<dyn ReadAt>) -> Result<SourceBackedPresentation, Box<dyn Error>> {
    Ok(SourceBackedPresentation::from_read_at(source)?)
}

fn write_owner(presentation: &SourceBackedPresentation) -> Result<Digest, Box<dyn Error>> {
    let mut output = DigestSink::new();
    presentation.write_to(&mut output)?;
    Ok(output.finish())
}

fn write_owner_bytes(presentation: &SourceBackedPresentation) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut output = Vec::new();
    presentation.write_to(&mut output)?;
    Ok(output)
}

fn append_commit(
    presentation: &SourceBackedPresentation,
    title: &str,
    body: &str,
) -> Result<litchi_odp::tail_append::Commit, Box<dyn Error>> {
    Ok(presentation.edit_tail_append(title, body)?.commit()?)
}

fn route_proof(
    route: &str,
    fixture: &[u8],
    title: &str,
    body: &str,
) -> Result<Value, Box<dyn Error>> {
    match route {
        "open" => {
            let presentation = owner(fixture)?;
            std::hint::black_box(presentation);
            Ok(json!({"kind": "open", "source": Digest::from_bytes(fixture).json()}))
        },
        "read" => {
            let presentation = owner(fixture)?;
            let slide_count = presentation.slide_count()?;
            let text = presentation.text()?;
            Ok(json!({
                "kind": "read",
                "slide_count": slide_count,
                "text": Digest::from_bytes(text.as_bytes()).json(),
            }))
        },
        "write" => {
            let presentation = owner(fixture)?;
            Ok(json!({"kind": "package", "output": write_owner(&presentation)?.json()}))
        },
        "append" => {
            let presentation = owner(fixture)?;
            let commit = append_commit(&presentation, title, body)?;
            Ok(json!({
                "kind": "package",
                "output": write_owner(commit.package())?.json(),
            }))
        },
        "replay" => {
            let presentation = owner(fixture)?;
            let commit = append_commit(&presentation, title, body)?;
            let independent = owner(fixture)?;
            let replayed = independent.apply_tail_append(commit.patch())?;
            Ok(json!({
                "kind": "package",
                "output": write_owner(replayed.package())?.json(),
            }))
        },
        other => Err(io::Error::other(format!("unknown route {other:?}")).into()),
    }
}

/// Run one complete lifecycle and return the number of bytes accepted by its
/// output sink.  Source setup is intentionally outside this function's timer.
fn run_route(
    route: &str,
    source: Arc<dyn ReadAt>,
    replay_source: Option<Arc<dyn ReadAt>>,
    title: &str,
    body: &str,
) -> Result<usize, Box<dyn Error>> {
    let bytes = match route {
        "open" => {
            let presentation = owner_source(source)?;
            std::hint::black_box(presentation);
            0
        },
        "read" => {
            let presentation = owner_source(source)?;
            let count = presentation.slide_count()?;
            let text = presentation.text()?;
            std::hint::black_box((count, text.len()));
            0
        },
        "write" => {
            let presentation = owner_source(source)?;
            let mut sink = CountingSink::default();
            presentation.write_to(&mut sink)?;
            sink.bytes
        },
        "append" => {
            let presentation = owner_source(source)?;
            let commit = append_commit(&presentation, title, body)?;
            let mut sink = CountingSink::default();
            commit.package().write_to(&mut sink)?;
            sink.bytes
        },
        "replay" => {
            let presentation = owner_source(source)?;
            let commit = append_commit(&presentation, title, body)?;
            let independent = owner_source(
                replay_source
                    .ok_or_else(|| io::Error::other("replay source preparation is missing"))?,
            )?;
            let replayed = independent.apply_tail_append(commit.patch())?;
            let mut sink = CountingSink::default();
            replayed.package().write_to(&mut sink)?;
            sink.bytes
        },
        other => return Err(io::Error::other(format!("unknown route {other:?}")).into()),
    };
    Ok(bytes)
}

fn timed_batch(
    route: &str,
    fixture: &[u8],
    title: &str,
    body: &str,
    iterations: usize,
) -> Result<(u128, usize), Box<dyn Error>> {
    // Match the accepted ODP probe boundary: every lifecycle receives fresh
    // positional sources, but source-byte allocation is setup work outside
    // that lifecycle's measured operation.  Summing per-lifecycle spans gives
    // a batch without retaining one fixture clone per iteration.
    let mut elapsed_ns = 0_u128;
    let mut output_bytes = None;
    for _ in 0..iterations {
        let source = source_for_fixture(fixture);
        let replay_source = (route == "replay").then(|| source_for_fixture(fixture));
        let started = Instant::now();
        let observed = run_route(
            route,
            Arc::clone(&source),
            replay_source.as_ref().map(Arc::clone),
            title,
            body,
        )?;
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

fn timed(args: &Args, fixture: &[u8]) -> Result<Value, Box<dyn Error>> {
    let route = args
        .route
        .as_deref()
        .ok_or_else(|| io::Error::other("timed route is missing"))?;
    let proof = route_proof(route, fixture, &args.title, &args.body)?;
    for _ in 0..args.warmups {
        let _ = timed_batch(route, fixture, &args.title, &args.body, args.iterations)?;
    }
    let mut elapsed_ns = Vec::with_capacity(args.samples);
    let mut output_bytes = None;
    for _ in 0..args.samples {
        let (elapsed, bytes) =
            timed_batch(route, fixture, &args.title, &args.body, args.iterations)?;
        if elapsed == 0 {
            return Err(
                io::Error::other(format!("{route} timing returned zero nanoseconds")).into(),
            );
        }
        if let Some(expected) = output_bytes {
            if expected != bytes {
                return Err(io::Error::other(format!(
                    "{route} output length changed across timed samples: {expected} versus {bytes}"
                ))
                .into());
            }
        } else {
            output_bytes = Some(bytes);
        }
        elapsed_ns.push(elapsed);
    }
    Ok(json!({
        "schema": SCHEMA,
        "mode": "timed",
        "route": route,
        "proof": proof,
        "iterations": args.iterations,
        "warmups": args.warmups,
        "samples": elapsed_ns,
        "output_bytes": output_bytes.unwrap_or(0),
    }))
}

fn calibrate(args: &Args, fixture: &[u8]) -> Result<Value, Box<dyn Error>> {
    let route = args
        .route
        .as_deref()
        .ok_or_else(|| io::Error::other("calibrate route is missing"))?;
    let proof = route_proof(route, fixture, &args.title, &args.body)?;
    let mut iterations = 1_usize;
    let mut attempts = Vec::new();
    loop {
        let (elapsed, output_bytes) =
            timed_batch(route, fixture, &args.title, &args.body, iterations)?;
        attempts.push(json!({
            "iterations": iterations,
            "elapsed_ns": elapsed,
            "output_bytes": output_bytes,
        }));
        if elapsed >= args.target_ns {
            return Ok(json!({
                "schema": SCHEMA,
                "mode": "calibrate",
                "route": route,
                "proof": proof,
                "target_ns": args.target_ns,
                "iterations": iterations,
                "elapsed_ns": elapsed,
                "output_bytes": output_bytes,
                "attempts": attempts,
            }));
        }
        if iterations >= args.max_iterations {
            return Err(io::Error::other(format!(
                "{route} calibration did not reach target {} ns within {} iterations",
                args.target_ns, args.max_iterations
            ))
            .into());
        }
        iterations = iterations.saturating_mul(2).min(args.max_iterations);
    }
}

fn verify(fixture: &[u8], title: &str, body: &str) -> Result<Value, Box<dyn Error>> {
    let source_digest = Digest::from_bytes(fixture);
    let content_digest = fixture_content_digest(fixture)?;
    let presentation = owner(fixture)?;

    let no_op_digest = write_owner(&presentation)?;
    let no_op_exact_source = no_op_digest == source_digest;

    let slide_count = presentation.slide_count()?;
    if slide_count == 0 {
        return Err(
            io::Error::other("qualification fixture must contain at least one slide").into(),
        );
    }
    let first_slide_present = presentation.slide(0)?.is_some();
    let text = presentation.text()?;

    let commit = append_commit(&presentation, title, body)?;
    let patch = commit.patch();
    // Keep one published copy only so inverse replay can reopen it as an
    // independent recipient.  All other correctness publications use the
    // bounded digest sink above.
    let append = write_owner_bytes(commit.package())?;
    let append_digest = Digest::from_bytes(&append);
    let same_owner = presentation.apply_tail_append(patch)?;
    let same_owner_digest = write_owner(same_owner.package())?;

    let independent = owner(fixture)?;
    let matching = independent.apply_tail_append(patch)?;
    let matching_digest = write_owner(matching.package())?;

    // Rebase onto a deterministic recipient whose content.xml is identical
    // but one unrelated archive member differs.  The member is deliberately
    // inspected through the published bytes so this proves content-scoped
    // replay and preservation of the recipient's opaque archive state.
    let (scoped_recipient_bytes, scoped_member_name, scoped_source_member) =
        scoped_recipient_fixture(fixture)?;
    let scoped_recipient_archive = Digest::from_bytes(&scoped_recipient_bytes);
    let scoped_recipient_content = member_digest(&scoped_recipient_bytes, "content.xml")?;
    let scoped_recipient_member = member_digest(&scoped_recipient_bytes, &scoped_member_name)?;
    let scoped_recipient = owner(&scoped_recipient_bytes)?;
    let scoped_forward = scoped_recipient.apply_tail_append(patch)?;
    let scoped_forward_bytes = write_owner_bytes(scoped_forward.package())?;
    let scoped_forward_content = member_digest(&scoped_forward_bytes, "content.xml")?;
    let scoped_forward_member = member_digest(&scoped_forward_bytes, &scoped_member_name)?;
    let scoped_forward_owner = owner(&scoped_forward_bytes)?;
    let scoped_inverse = scoped_forward_owner.apply_tail_append(&patch.inverse())?;
    let scoped_inverse_bytes = write_owner_bytes(scoped_inverse.package())?;
    let scoped_inverse_content = member_digest(&scoped_inverse_bytes, "content.xml")?;
    let scoped_inverse_member = member_digest(&scoped_inverse_bytes, &scoped_member_name)?;

    // Apply the inverse directly to the committed owner as a separate
    // whole-owner proof.  This recipient still carries the commit's original
    // archive entries, so its publication must restore the complete source
    // artifact, not only content.xml.
    let whole_owner_inverse = commit.package().apply_tail_append(&patch.inverse())?;
    let whole_owner_inverse_digest = write_owner(whole_owner_inverse.package())?;

    // Reopen the published output as an independent recipient as well.  Its
    // content scope must restore exactly, while ZIP entry metadata may be
    // regenerated by the publication.
    let output_owner = owner(&append)?;
    let inverse = output_owner.apply_tail_append(&patch.inverse())?;
    let inverse_digest = write_owner(inverse.package())?;
    let inverse_content_digest =
        member_digest(&write_owner_bytes(inverse.package())?, "content.xml")?;

    let changed = append_commit(&presentation, "Qualification different title", body)?;
    let changed_bytes = write_owner_bytes(changed.package())?;
    let changed_recipient = owner(&changed_bytes)?;
    let changed_recipient_refused = matches!(
        changed_recipient.apply_tail_append(patch),
        Err(litchi_odp::tail_append::Error::PatchConflict { .. })
    );

    let patch_changed = patch.source_content_sha256() != patch.target_content_sha256();
    let patch_source_matches_fixture =
        hex_bytes(&patch.source_content_sha256()) == content_digest.sha256;
    let append_output_changed = append_digest != source_digest;
    let scoped_recipient_content_matches = scoped_recipient_content == content_digest;
    let scoped_recipient_member_changed = match scoped_source_member.as_ref() {
        Some(source_member) => source_member != &scoped_recipient_member,
        None => true,
    };
    let scoped_forward_content_matches_patch =
        hex_bytes(&patch.target_content_sha256()) == scoped_forward_content.sha256;
    let scoped_forward_preserves_unrelated_member =
        scoped_forward_member == scoped_recipient_member;
    let scoped_inverse_preserves_unrelated_member =
        scoped_inverse_member == scoped_recipient_member;
    let scoped_inverse_restores_content_scope = scoped_inverse_content == content_digest;
    let whole_owner_inverse_exact_source = whole_owner_inverse_digest == source_digest;
    let inverse_content_exact_source = inverse_content_digest == content_digest;
    let checks = json!({
        "no_op_exact_source": no_op_exact_source,
        "read_semantics": slide_count > 0 && first_slide_present,
        "append_changed": patch_changed && patch_source_matches_fixture && append_output_changed,
        "patch_source_matches_fixture": patch_source_matches_fixture,
        "append_output_changed": append_output_changed,
        "scoped_recipient_content_matches": scoped_recipient_content_matches,
        "scoped_recipient_member_changed": scoped_recipient_member_changed,
        "scoped_forward_content_matches_patch": scoped_forward_content_matches_patch,
        "scoped_forward_preserves_unrelated_member": scoped_forward_preserves_unrelated_member,
        "scoped_inverse_preserves_unrelated_member": scoped_inverse_preserves_unrelated_member,
        "scoped_inverse_restores_content_scope": scoped_inverse_restores_content_scope,
        "same_owner_replay_exact": same_owner_digest == append_digest,
        "content_matching_forward_exact": matching_digest == append_digest,
        "whole_owner_inverse_exact_source": whole_owner_inverse_exact_source,
        "inverse_content_exact_source": inverse_content_exact_source,
        "changed_recipient_refused": changed_recipient_refused,
    });
    if checks
        .as_object()
        .into_iter()
        .flat_map(|object| object.values())
        .any(|value| value != &Value::Bool(true))
    {
        return Err(
            io::Error::other("ODP correctness preflight did not satisfy every check").into(),
        );
    }

    Ok(json!({
        "schema": SCHEMA,
        "mode": "verify",
        "fixture": {
            "bytes": source_digest.bytes,
            "sha256": source_digest.sha256,
            "content": content_digest.json(),
        },
        "checks": checks,
        "read": {
            "slide_count": slide_count,
            "first_slide_present": first_slide_present,
            "text_bytes": text.len(),
            "text_sha256": sha256_hex(text.as_bytes()),
        },
        "patch": {
            "is_noop": patch.is_noop(),
            "source_content_sha256": hex_bytes(&patch.source_content_sha256()),
            "target_content_sha256": hex_bytes(&patch.target_content_sha256()),
        },
        "scoped_replay": {
            "provenance": SCOPED_REPLAY_PROVENANCE,
            "member_name": scoped_member_name,
            "recipient_archive": scoped_recipient_archive.json(),
            "source_member": scoped_source_member
                .as_ref()
                .map(Digest::json)
                .unwrap_or(Value::Null),
            "recipient_member": scoped_recipient_member.json(),
            "forward_member": scoped_forward_member.json(),
            "inverse_member": scoped_inverse_member.json(),
            "recipient_content": scoped_recipient_content.json(),
            "forward_content": scoped_forward_content.json(),
            "inverse_content": scoped_inverse_content.json(),
        },
        "outputs": {
            "source": source_digest.json(),
            "no_op": no_op_digest.json(),
            "append": append_digest.json(),
            "same_owner_replay": same_owner_digest.json(),
            "content_matching_replay": matching_digest.json(),
            "whole_owner_inverse": whole_owner_inverse_digest.json(),
            "inverse": inverse_digest.json(),
            "inverse_content": inverse_content_digest.json(),
        },
    }))
}

fn hex_bytes(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn run(args: Args) -> Result<Value, Box<dyn Error>> {
    let fixture = fs::read(&args.fixture)?;
    match args.mode {
        Mode::Verify => verify(&fixture, &args.title, &args.body),
        Mode::Calibrate => calibrate(&args, &fixture),
        Mode::Timed => timed(&args, &fixture),
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse().map_err(io::Error::other)?;
    let report = run(args)?;
    println!("{}", serde_json::to_string(&report)?);
    Ok(())
}
