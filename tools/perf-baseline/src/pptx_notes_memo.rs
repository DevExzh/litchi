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
//! * `cold_capture`: `opened_presentation` over a pre-parsed package;
//! * `warm_notes_commit`: a real notes edit and commit with all slide roots
//!   unchanged, so the opened snapshot's memo is eligible for every root;
//! * `changed_slide_commit`: a real title edit on slide zero and commit, so
//!   exactly one slide root is changed while the other two remain eligible.
//!
//! Fixture creation, source setup for the two edit rows, serialization,
//! semantic reopen and raw ZIP fidelity checks, and source immutability checks
//! are outside the clocks.  The allocator binary enters the same
//! operation-scoped region as the normal binary; the latter reports an
//! explicit unavailable sample.

use std::{
    collections::{BTreeMap, BTreeSet},
    error::Error,
    ffi::OsString,
    fs::OpenOptions,
    io::Write,
    path::PathBuf,
    time::Instant,
};

use litchi_pptx::{Package, builder::Builder, notes::Limits as NotesLimits};
use serde::Serialize;
use soapberry_zip::office::{ArchiveLimits, ArchiveReader};
use soapberry_zip::{PreservationIndex, ZipArchive};

const SCHEMA: &str = "pptx_notes_memo_probe_v1";
const GENERATOR: &str = "pptx-notes-memo-three-slide-v1";
const FIXTURE_SLIDES: usize = 3;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 1_000;
const SOURCE_REVISION_BYTES: usize = 40;

const COLD_CASE: &str = "pptx_notes_memo_cold_capture";
const WARM_CASE: &str = "pptx_notes_memo_warm_notes_commit";
const CHANGED_CASE: &str = "pptx_notes_memo_changed_slide_commit";

const COLD_SCOPE: &str = "opened_presentation over a pre-parsed package; Package::from_bytes fixture setup and all fidelity checks are outside";
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

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
struct SemanticFidelity {
    verified: bool,
    slide_count: usize,
    notes_slide_count: usize,
    slide_text_sha256: String,
    notes_text_sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct RawOracle {
    source_member_count: usize,
    output_member_count: usize,
    source_payload_sha256: String,
    output_payload_sha256: String,
    member_names_equal: bool,
    untouched_payloads_equal: bool,
    relationship_member_count: usize,
    relationship_names_equal: bool,
    relationship_payloads_equal: bool,
    master_theme_member_count: usize,
    master_theme_payloads_equal: bool,
    changed_members: Vec<String>,
    expected_changed_members: Vec<String>,
    exact_changed_member_set: bool,
    changed_payloads_nonempty: bool,
    slide_root_count: usize,
    notes_root_count: usize,
    changed_slide_index: Option<usize>,
    changed_notes_slide_index: Option<usize>,
    untouched_slide_root_count: usize,
}

#[derive(Clone, Debug, Serialize)]
struct Fidelity {
    semantic: SemanticFidelity,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_immutable: Option<bool>,
    changed_commit_reported: bool,
    output_reopened: bool,
    changed_slide_index: Option<usize>,
    untouched_slide_root_count: usize,
    raw: RawOracle,
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

#[derive(Debug)]
struct RawArchive {
    payloads: BTreeMap<String, Vec<u8>>,
}

fn raw_archive(bytes: &[u8]) -> Result<RawArchive, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?.into_zip_archive();
    let limits = ArchiveLimits::default();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut scratch, limits)?;
    let reader = ArchiveReader::new(bytes, limits)?;
    let mut payloads = BTreeMap::new();
    for entry in index.entries() {
        let name = std::str::from_utf8(entry.raw_name_bytes())?.to_owned();
        if payloads.insert(name.clone(), reader.read(&name)?).is_some() {
            return Err("notes memo archive contains duplicate member names".into());
        }
    }
    if payloads.is_empty() {
        return Err("notes memo archive has no members".into());
    }
    Ok(RawArchive { payloads })
}

fn payload_digest(payloads: &BTreeMap<String, Vec<u8>>) -> String {
    let mut input = Vec::new();
    for (name, payload) in payloads {
        input.extend_from_slice(&(name.len() as u64).to_le_bytes());
        input.extend_from_slice(name.as_bytes());
        input.extend_from_slice(&(payload.len() as u64).to_le_bytes());
        input.extend_from_slice(payload);
    }
    crate::sha256_hex(&input)
}

fn member_set_with_suffix(payloads: &BTreeMap<String, Vec<u8>>, suffix: &str) -> BTreeSet<String> {
    payloads
        .keys()
        .filter(|name| name.ends_with(suffix))
        .cloned()
        .collect()
}

fn slide_root_names(payloads: &BTreeMap<String, Vec<u8>>) -> BTreeSet<String> {
    payloads
        .keys()
        .filter(|name| {
            name.strip_prefix("ppt/slides/slide")
                .and_then(|name| name.strip_suffix(".xml"))
                .is_some_and(|index| {
                    !index.is_empty() && index.bytes().all(|byte| byte.is_ascii_digit())
                })
        })
        .cloned()
        .collect()
}

fn notes_root_names(payloads: &BTreeMap<String, Vec<u8>>) -> BTreeSet<String> {
    payloads
        .keys()
        .filter(|name| {
            name.strip_prefix("ppt/notesSlides/notesSlide")
                .and_then(|name| name.strip_suffix(".xml"))
                .is_some_and(|index| {
                    !index.is_empty() && index.bytes().all(|byte| byte.is_ascii_digit())
                })
        })
        .cloned()
        .collect()
}

fn slide_index(name: &str) -> Option<usize> {
    name.strip_prefix("ppt/slides/slide")
        .and_then(|name| name.strip_suffix(".xml"))
        .and_then(|index| index.parse::<usize>().ok())
        .and_then(|index| index.checked_sub(1))
}

fn notes_slide_index(name: &str) -> Option<usize> {
    name.strip_prefix("ppt/notesSlides/notesSlide")
        .and_then(|name| name.strip_suffix(".xml"))
        .and_then(|index| index.parse::<usize>().ok())
        .and_then(|index| index.checked_sub(1))
}

fn expected_changed_members(operation: Operation) -> BTreeSet<String> {
    match operation {
        Operation::Cold => BTreeSet::new(),
        Operation::WarmNotes => BTreeSet::from(["ppt/notesSlides/notesSlide2.xml".to_owned()]),
        Operation::ChangedSlide => BTreeSet::from(["ppt/slides/slide1.xml".to_owned()]),
    }
}

fn check_raw_oracle(
    source: &[u8],
    output: &[u8],
    operation: Operation,
) -> Result<RawOracle, Box<dyn Error>> {
    let source = raw_archive(source)?.payloads;
    let output = raw_archive(output)?.payloads;
    let source_payload_sha256 = payload_digest(&source);
    let output_payload_sha256 = payload_digest(&output);
    let digests_match = source_payload_sha256 == output_payload_sha256;
    if matches!(operation, Operation::Cold) != digests_match {
        return Err(format!(
            "notes memo source/output payload digest relationship is wrong for {}",
            operation.name()
        )
        .into());
    }
    let source_names = source.keys().cloned().collect::<BTreeSet<_>>();
    let output_names = output.keys().cloned().collect::<BTreeSet<_>>();
    let member_names_equal = source_names == output_names;
    if !member_names_equal {
        return Err("notes memo output changed the ZIP member identity set".into());
    }
    let changed_members = source_names
        .iter()
        .filter(|name| source[*name] != output[*name])
        .cloned()
        .collect::<BTreeSet<_>>();
    let expected = expected_changed_members(operation);
    let exact_changed_member_set = changed_members == expected;
    if !exact_changed_member_set {
        return Err(format!(
            "notes memo changed members differ: actual={changed_members:?} expected={expected:?}"
        )
        .into());
    }
    let untouched_payloads_equal = source_names
        .difference(&changed_members)
        .all(|name| source[name] == output[name]);
    if !untouched_payloads_equal {
        return Err("notes memo changed an untouched payload".into());
    }
    let changed_payloads_nonempty = changed_members
        .iter()
        .all(|name| !source[name].is_empty() && !output[name].is_empty());
    if !changed_payloads_nonempty && !changed_members.is_empty() {
        return Err("notes memo changed payload evidence is empty".into());
    }

    let source_relationships = member_set_with_suffix(&source, ".rels");
    let output_relationships = member_set_with_suffix(&output, ".rels");
    let relationship_names_equal = source_relationships == output_relationships;
    let relationship_payloads_equal = relationship_names_equal
        && source_relationships
            .iter()
            .all(|name| source[name] == output[name]);
    if !relationship_payloads_equal {
        return Err("notes memo changed relationship identities or targets".into());
    }

    let source_masters = source
        .keys()
        .filter(|name| {
            name.contains("/notesMasters/")
                || name.contains("/slideMasters/")
                || name.contains("/slideLayouts/")
                || name.contains("/theme/")
        })
        .cloned()
        .collect::<BTreeSet<_>>();
    let output_masters = output
        .keys()
        .filter(|name| {
            name.contains("/notesMasters/")
                || name.contains("/slideMasters/")
                || name.contains("/slideLayouts/")
                || name.contains("/theme/")
        })
        .cloned()
        .collect::<BTreeSet<_>>();
    let master_theme_payloads_equal = source_masters == output_masters
        && source_masters
            .iter()
            .all(|name| source[name] == output[name]);
    if !master_theme_payloads_equal {
        return Err("notes memo changed a master, layout, or theme payload".into());
    }

    let slide_roots = slide_root_names(&source);
    let notes_roots = notes_root_names(&source);
    if slide_roots.len() != FIXTURE_SLIDES || notes_roots.len() != FIXTURE_SLIDES {
        return Err("notes memo fixture root inventory is not three-by-three".into());
    }
    let changed_slide_indices = changed_members
        .iter()
        .filter_map(|name| slide_index(name))
        .collect::<Vec<_>>();
    let changed_notes_indices = changed_members
        .iter()
        .filter_map(|name| notes_slide_index(name))
        .collect::<Vec<_>>();
    if changed_slide_indices.len() > 1 || changed_notes_indices.len() > 1 {
        return Err("notes memo changed more than one slide or notes root".into());
    }
    let changed_slide_count = changed_slide_indices.len();
    Ok(RawOracle {
        source_member_count: source.len(),
        output_member_count: output.len(),
        source_payload_sha256,
        output_payload_sha256,
        member_names_equal,
        untouched_payloads_equal,
        relationship_member_count: source_relationships.len(),
        relationship_names_equal,
        relationship_payloads_equal,
        master_theme_member_count: source_masters.len(),
        master_theme_payloads_equal,
        changed_members: changed_members.into_iter().collect(),
        expected_changed_members: expected.into_iter().collect(),
        exact_changed_member_set,
        changed_payloads_nonempty,
        slide_root_count: slide_roots.len(),
        notes_root_count: notes_roots.len(),
        changed_slide_index: changed_slide_indices.first().copied(),
        changed_notes_slide_index: changed_notes_indices.first().copied(),
        untouched_slide_root_count: slide_roots.len() - changed_slide_count,
    })
}

fn check_semantic_fidelity(
    package: &Package,
    operation: Operation,
) -> Result<SemanticFidelity, Box<dyn Error>> {
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
    Ok(SemanticFidelity {
        verified: true,
        slide_count: slides.len(),
        notes_slide_count: notes.slides().len(),
        slide_text_sha256: crate::sha256_hex(slide_texts.join("\u{1f}").as_bytes()),
        notes_text_sha256: crate::sha256_hex(note_texts.join("\u{1f}").as_bytes()),
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
    // Package ingestion is fixture setup.  The cold row measures only the
    // opened-presentation capture, which is the physical notes-root scan
    // under comparison.
    let package = Package::from_bytes(source)?;
    let region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let result = package.opened_presentation();
    let elapsed_ns = crate::elapsed_ns(started.elapsed())?;
    let allocation = region
        .finish()
        .unwrap_or_else(crate::allocation_metrics::unavailable_sample);
    result?;
    Ok(TimedOperation {
        package,
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
    let semantic = check_semantic_fidelity(&timed.package, operation)?;
    let output = timed.package.to_bytes()?;
    let output_sha256 = crate::sha256_hex(&output);
    let reopened = Package::from_bytes(&output)?;
    let reopened_semantic = check_semantic_fidelity(&reopened, operation)?;
    if reopened_semantic != semantic {
        return Err("notes memo output failed the independent semantic reopen oracle".into());
    }
    let raw = check_raw_oracle(source, &output, operation)?;
    let fidelity = Fidelity {
        semantic,
        source_immutable: timed.source_immutable,
        changed_commit_reported: timed.changed_commit_reported,
        output_reopened: true,
        changed_slide_index: raw.changed_slide_index,
        untouched_slide_root_count: raw.untouched_slide_root_count,
        raw,
    };
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
        fidelity_scope: "untimed semantic reopen plus exact three-slide/notes and raw ZIP member, relationship, master/theme, and root oracle after every retained operation",
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
