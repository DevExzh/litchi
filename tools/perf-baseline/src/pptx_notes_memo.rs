//! Opt-in PPTX notes-root memo probe.
//!
//! This probe is deliberately kept outside the normal `Case` catalog.  It is
//! a small, fixed workload for the notes-owning three-slide deck used by the
//! slide-root memo tests.  The report keeps the existing harness
//! `elapsed_ns` statistics shape, while the surrounding schema is local to
//! this probe so adding it cannot silently change the default matrix.
//!
//! The timed regions are:
//!
//! * `cold_capture`: `Package::from_bytes` and `opened_presentation`;
//! * `warm_notes_commit`: a real notes edit and commit with all slide roots
//!   unchanged, so the opened snapshot's memo is eligible for every root;
//! * `changed_slide_commit`: a real title edit on slide zero and commit, so
//!   exactly one slide root is changed while the other two remain eligible.
//!
//! Fixture creation, source setup for the two edit rows, serialization,
//! semantic fidelity checks, and source immutability checks are outside the
//! clocks.  The allocator binary enters the same operation-scoped region as
//! the normal binary; the latter reports an explicit unavailable sample.

use std::{error::Error, ffi::OsString, fs::OpenOptions, io::Write, path::PathBuf, time::Instant};

use litchi_pptx::{Package, builder::Builder, notes::Limits as NotesLimits};
use serde::Serialize;

const SCHEMA: &str = "pptx_notes_memo_probe_v1";
const GENERATOR: &str = "pptx-notes-memo-three-slide-v1";
const FIXTURE_SLIDES: usize = 3;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 1_000;
const SOURCE_REVISION_BYTES: usize = 40;

const COLD_CASE: &str = "pptx_notes_memo_cold_capture";
const WARM_CASE: &str = "pptx_notes_memo_warm_notes_commit";
const CHANGED_CASE: &str = "pptx_notes_memo_changed_slide_commit";

const COLD_SCOPE: &str = "Package::from_bytes plus opened_presentation; fixture construction and all fidelity checks are outside";
const WARM_SCOPE: &str = "opened Snapshot::edit, set_notes_text(1), and commit; source package setup, serialization, and fidelity checks are outside";
const CHANGED_SCOPE: &str = "opened Snapshot::edit, set_shape_text(0, 0), and commit; source package setup, serialization, and fidelity checks are outside";

#[derive(Debug)]
struct Config {
    output: PathBuf,
    samples: usize,
    warmup: usize,
    source_revision: String,
}

#[derive(Debug)]
struct Fixture {
    bytes: Vec<u8>,
    sha256: String,
}

#[derive(Clone, Copy, Debug)]
enum Operation {
    Cold,
    WarmNotes,
    ChangedSlide,
}

impl Operation {
    const ALL: [Self; 3] = [Self::Cold, Self::WarmNotes, Self::ChangedSlide];

    const fn case(self) -> &'static str {
        match self {
            Self::Cold => COLD_CASE,
            Self::WarmNotes => WARM_CASE,
            Self::ChangedSlide => CHANGED_CASE,
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Cold => COLD_SCOPE,
            Self::WarmNotes => WARM_SCOPE,
            Self::ChangedSlide => CHANGED_SCOPE,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Cold => "cold_capture",
            Self::WarmNotes => "warm_notes_commit",
            Self::ChangedSlide => "changed_slide_commit",
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct FixtureIdentity {
    generator: &'static str,
    slides: usize,
    notes_slides: usize,
    authored_shapes_per_slide: usize,
    bytes: usize,
    sha256: String,
    recipe: &'static str,
}

#[derive(Clone, Debug, Serialize)]
struct Fidelity {
    verified: bool,
    slide_count: usize,
    notes_slide_count: usize,
    slide_text_sha256: String,
    notes_text_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_immutable: Option<bool>,
    changed_commit_reported: bool,
    changed_slide_index: Option<usize>,
    untouched_slide_root_count: usize,
}

#[derive(Clone, Debug, Serialize)]
struct Observation {
    sample_index: usize,
    elapsed_ns: u64,
    output_bytes: usize,
    output_sha256: String,
    fidelity: Fidelity,
    allocation: crate::allocation_metrics::Sample,
}

#[derive(Serialize)]
struct ResultRow {
    case: &'static str,
    operation: &'static str,
    timing_scope: &'static str,
    elapsed_ns: crate::Statistics,
    observations: Vec<Observation>,
}

#[derive(Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    binary_identity: crate::BinaryIdentity,
    source_revision: String,
    instrumentation: &'static str,
    allocator: &'static str,
    allocator_counter_revision: Option<&'static str>,
    timing_claim: &'static str,
    allocation_scope: &'static str,
    fixture: FixtureIdentity,
    samples: usize,
    warmup: usize,
    checked_iteration_count: usize,
    memo_scope: &'static str,
    fidelity_scope: &'static str,
    results: Vec<ResultRow>,
}

fn parse_config<I>(args: I) -> Result<Option<Config>, Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let mut args = args.into_iter();
    let mut output = None;
    let mut samples = 20;
    let mut warmup = 3;
    let mut source_revision = None;
    while let Some(argument) = args.next() {
        let argument = argument
            .into_string()
            .map_err(|_| "notes memo probe arguments must be UTF-8")?;
        match argument.as_str() {
            "--help" | "-h" => return Ok(None),
            "--output" => {
                let value = args.next().ok_or("--output requires a path")?;
                output = Some(PathBuf::from(value));
            },
            "--samples" => {
                samples = parse_bounded_usize(&mut args, "--samples", 1, MAX_SAMPLES)?;
            },
            "--warmup" => {
                warmup = parse_bounded_usize(&mut args, "--warmup", 0, MAX_WARMUP)?;
            },
            "--source-revision" => {
                source_revision = Some(
                    args.next()
                        .ok_or("--source-revision requires a 40-character Git revision")?
                        .into_string()
                        .map_err(|_| "--source-revision must be UTF-8")?,
                );
            },
            other => return Err(format!("unknown notes memo probe argument: {other}").into()),
        }
    }
    let output = output.ok_or("--output is required")?;
    let source_revision = source_revision.ok_or("--source-revision is required")?;
    if source_revision.len() != SOURCE_REVISION_BYTES
        || !source_revision.bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err("--source-revision must be a 40-character hexadecimal Git revision".into());
    }
    Ok(Some(Config {
        output,
        samples,
        warmup,
        source_revision,
    }))
}

fn parse_bounded_usize<I>(
    args: &mut I,
    flag: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>>
where
    I: Iterator<Item = OsString>,
{
    let value = args
        .next()
        .ok_or_else(|| format!("{flag} requires a value"))?;
    let value = value
        .into_string()
        .map_err(|_| format!("{flag} must be UTF-8"))?;
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} must be an integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("{flag} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn fixture() -> Result<Fixture, Box<dyn Error>> {
    let mut builder = Builder::new()?;
    for index in 0..FIXTURE_SLIDES {
        let slide = builder.presentation()?.add_slide()?;
        slide.set_title(&format!("notes slide {index}"));
        slide.add_text_box(&format!("slide body {index}"), 36, 36, 914_400, 457_200);
        slide.set_notes(&format!("speaker notes {index}"));
    }
    let bytes = builder.build()?.to_bytes()?;
    if bytes.is_empty() {
        return Err("notes memo fixture serialized to an empty package".into());
    }
    let sha256 = crate::sha256_hex(&bytes);
    Ok(Fixture { bytes, sha256 })
}

fn expected_slide_text(index: usize, changed: bool) -> String {
    if changed && index == 0 {
        "rewritten slide title\nslide body 0".to_owned()
    } else {
        format!("notes slide {index}\nslide body {index}")
    }
}

fn expected_note_text(index: usize, warm_notes: bool) -> String {
    if warm_notes && index == 1 {
        "rewritten notes".to_owned()
    } else {
        format!("speaker notes {index}")
    }
}

fn check_fidelity(
    package: &Package,
    operation: Operation,
    source_immutable: Option<bool>,
    changed_commit_reported: bool,
) -> Result<Fidelity, Box<dyn Error>> {
    let changed = matches!(operation, Operation::ChangedSlide);
    let warm_notes = matches!(operation, Operation::WarmNotes);
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    if slides.len() != FIXTURE_SLIDES {
        return Err(format!("expected {FIXTURE_SLIDES} slides, got {}", slides.len()).into());
    }
    let mut slide_texts = Vec::with_capacity(slides.len());
    for (index, slide) in slides.iter().enumerate() {
        let text = slide.text()?;
        if text != expected_slide_text(index, changed) {
            return Err(format!("slide {index} text fidelity mismatch").into());
        }
        slide_texts.push(text);
    }
    let notes = package
        .notes()?
        .ok_or("notes memo fixture lost its notes graph")?;
    if notes.slides().len() != FIXTURE_SLIDES {
        return Err(format!(
            "expected {FIXTURE_SLIDES} notes slides, got {}",
            notes.slides().len()
        )
        .into());
    }
    let mut note_texts = Vec::with_capacity(notes.slides().len());
    for (index, note) in notes.slides().iter().enumerate() {
        let text = note
            .text(NotesLimits::DEFAULT)?
            .ok_or_else(|| format!("notes slide {index} has no text"))?;
        if text != expected_note_text(index, warm_notes) {
            return Err(format!("notes slide {index} text fidelity mismatch").into());
        }
        note_texts.push(text);
    }
    Ok(Fidelity {
        verified: true,
        slide_count: slides.len(),
        notes_slide_count: notes.slides().len(),
        slide_text_sha256: crate::sha256_hex(slide_texts.join("\u{1f}").as_bytes()),
        notes_text_sha256: crate::sha256_hex(note_texts.join("\u{1f}").as_bytes()),
        source_immutable,
        changed_commit_reported,
        changed_slide_index: changed.then_some(0),
        untouched_slide_root_count: match operation {
            Operation::Cold => 0,
            Operation::WarmNotes => FIXTURE_SLIDES,
            Operation::ChangedSlide => FIXTURE_SLIDES - 1,
        },
    })
}

fn commit_opened(
    opened: litchi_pptx::opened::Snapshot,
    operation: Operation,
) -> Result<(Package, bool), Box<dyn Error>> {
    let mut edit = opened.edit();
    let changed = match operation {
        Operation::WarmNotes => edit.set_notes_text(1_usize, "rewritten notes")?,
        Operation::ChangedSlide => {
            edit.set_shape_text(0_usize, 0_usize, "rewritten slide title")?
        },
        Operation::Cold => return Err("cold capture cannot use commit_opened".into()),
    };
    if !changed {
        return Err(format!("{} unexpectedly produced a no-op edit", operation.name()).into());
    }
    let commit = edit.commit()?;
    let diagnostics_changed = commit.diagnostics().changed();
    if !diagnostics_changed {
        return Err(format!("{} commit reported no change", operation.name()).into());
    }
    let package = commit.into_package();
    Ok((package, diagnostics_changed))
}

struct TimedOperation {
    package: Package,
    changed_commit_reported: bool,
    elapsed_ns: u64,
    allocation: crate::allocation_metrics::Sample,
    source_immutable: Option<bool>,
}

fn run_cold(source: &[u8]) -> Result<TimedOperation, Box<dyn Error>> {
    let region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let result: Result<Package, Box<dyn Error>> = (|| {
        let package = Package::from_bytes(source)?;
        let _opened = package.opened_presentation()?;
        Ok(package)
    })();
    let elapsed_ns = crate::elapsed_ns(started.elapsed())?;
    let allocation = region
        .finish()
        .unwrap_or_else(crate::allocation_metrics::unavailable_sample);
    Ok(TimedOperation {
        package: result?,
        changed_commit_reported: false,
        elapsed_ns,
        allocation,
        source_immutable: None,
    })
}

fn run_edit(source: &[u8], operation: Operation) -> Result<TimedOperation, Box<dyn Error>> {
    let source_package = Package::from_bytes(source)?;
    let source_before = source_package.to_bytes()?;
    let opened = source_package.opened_presentation()?;

    // Only the detached edit and commit are timed.  Parsing, opened-root
    // capture, and the source immutability oracle remain outside this clock.
    let region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let result = commit_opened(opened, operation);
    let elapsed_ns = crate::elapsed_ns(started.elapsed())?;
    let allocation = region
        .finish()
        .unwrap_or_else(crate::allocation_metrics::unavailable_sample);
    let (package, changed_commit_reported) = result?;
    let source_immutable = source_package.to_bytes()? == source_before;
    if !source_immutable {
        return Err("source package changed while publishing a detached commit".into());
    }
    Ok(TimedOperation {
        package,
        changed_commit_reported,
        elapsed_ns,
        allocation,
        source_immutable: Some(true),
    })
}

fn one_iteration(
    source: &[u8],
    operation: Operation,
    sample_index: usize,
) -> Result<Observation, Box<dyn Error>> {
    let timed = match operation {
        Operation::Cold => run_cold(source)?,
        Operation::WarmNotes | Operation::ChangedSlide => run_edit(source, operation)?,
    };
    let fidelity = check_fidelity(
        &timed.package,
        operation,
        timed.source_immutable,
        timed.changed_commit_reported,
    )?;
    let output = timed.package.to_bytes()?;
    let output_sha256 = crate::sha256_hex(&output);
    Ok(Observation {
        sample_index,
        elapsed_ns: timed.elapsed_ns,
        output_bytes: output.len(),
        output_sha256,
        fidelity,
        allocation: timed.allocation,
    })
}

fn run_row(
    config: &Config,
    fixture: &Fixture,
    operation: Operation,
) -> Result<ResultRow, Box<dyn Error>> {
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup plus samples overflows usize")?;
    let mut elapsed = Vec::with_capacity(config.samples);
    let mut observations = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let observation = one_iteration(
            &fixture.bytes,
            operation,
            iteration.saturating_sub(config.warmup),
        )?;
        if iteration >= config.warmup {
            elapsed.push(observation.elapsed_ns);
            observations.push(observation);
        }
    }
    Ok(ResultRow {
        case: operation.case(),
        operation: operation.name(),
        timing_scope: operation.timing_scope(),
        elapsed_ns: crate::statistics(elapsed),
        observations,
    })
}

fn usage() {
    eprintln!(
        "usage: litchi-perf-baseline pptx-notes-memo --output PATH --source-revision HEX40 [--samples N] [--warmup N]"
    );
}

/// Run the fixed notes-memo probe after the executable has consumed its
/// selector.  The destination is opened with `create_new`, so an old report
/// cannot be replaced accidentally.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let Some(config) = parse_config(args)? else {
        usage();
        return Ok(());
    };
    let fixture = fixture()?;
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup plus samples overflows usize")?;
    let results = Operation::ALL
        .into_iter()
        .map(|operation| run_row(&config, &fixture, operation))
        .collect::<Result<Vec<_>, _>>()?;
    let binary_identity = crate::current_executable_identity()?;
    let report = Report {
        schema: SCHEMA,
        tool: env!("CARGO_PKG_NAME"),
        binary_identity,
        source_revision: config.source_revision,
        instrumentation: crate::allocation_metrics::instrumentation_identity(),
        allocator: crate::allocation_metrics::allocator_identity(),
        allocator_counter_revision: crate::allocation_metrics::counter_revision(),
        timing_claim: if crate::allocation_metrics::instrumentation_identity() == "none" {
            "native_wall_clock_samples; compare only under the approved quiet ABBA protocol"
        } else {
            "allocator-instrumented wall-clock samples; not a latency claim"
        },
        allocation_scope: "one non-nested operation-scoped global-allocator region around each timed API operation",
        fixture: FixtureIdentity {
            generator: GENERATOR,
            slides: FIXTURE_SLIDES,
            notes_slides: FIXTURE_SLIDES,
            authored_shapes_per_slide: 1,
            bytes: fixture.bytes.len(),
            sha256: fixture.sha256,
            recipe: "three authored slides; title notes slide {i}, one body text box, speaker notes {i}",
        },
        samples: config.samples,
        warmup: config.warmup,
        checked_iteration_count: total,
        memo_scope: "source-audited SlideRootMemo eligibility; runtime proved-root counters remain internal and are covered by the focused crate test",
        fidelity_scope: "untimed exact three-slide text and three-notes-slide text oracle after every retained operation",
        results,
    };
    let bytes = serde_json::to_vec_pretty(&report)?;
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(config.output)?;
    output.write_all(&bytes)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parser_requires_source_identity_and_output() {
        assert!(parse_config(Vec::<OsString>::new()).is_err());
        let valid = vec![
            OsString::from("--output"),
            OsString::from("memo.json"),
            OsString::from("--source-revision"),
            OsString::from("0123456789abcdef0123456789abcdef01234567"),
        ];
        assert!(parse_config(valid).unwrap().is_some());
    }

    #[test]
    fn parser_rejects_unbounded_or_non_hex_controls() {
        let base = vec![
            OsString::from("--output"),
            OsString::from("memo.json"),
            OsString::from("--source-revision"),
            OsString::from("not-a-revision"),
        ];
        assert!(parse_config(base).is_err());
    }

    #[test]
    fn operation_names_are_distinct_and_scopes_are_explicit() {
        assert_eq!(
            Operation::ALL.map(Operation::case),
            [COLD_CASE, WARM_CASE, CHANGED_CASE]
        );
        assert!(
            Operation::WarmNotes
                .timing_scope()
                .contains("set_notes_text")
        );
        assert!(
            Operation::ChangedSlide
                .timing_scope()
                .contains("set_shape_text")
        );
    }
}
