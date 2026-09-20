//! Candidate-only observer for the bounded opened-PresentationML MCE memo.
//!
//! This binary intentionally has no timer and no counting allocator.  It
//! exercises the public retention observations over two repeats for a real
//! LibreOffice deck and the bounded marker-free generated fixture.  Every
//! semantic and revision assertion is made through the normal PPTX API.

use litchi_pptx::Package;
use litchi_pptx::opened::Limits;
use sha2::{Digest, Sha256};
use std::env;
use std::error::Error;

const PROBE_ID: &str = "0704-retention-observer";
const EDIT_MARKER: &str = "litchi-perf-0691-opened-phase";
const GENERATED_SLIDES: usize = 12;
const GENERATED_SHAPES: usize = 8;
const MAX_REPEATS: usize = 2;
const MAX_SOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_OBSERVED_SLIDES: usize = 64;
const MAX_OBSERVED_SHAPES: usize = 64;
const DEFAULT_MCE_BYTES: usize = 1024 * 1024;
const TINY_MCE_BYTES: usize = 1;

type Fallible<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Target {
    slide: usize,
    shape: usize,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct SemanticDigest {
    digest: String,
    payload_bytes: usize,
    slides: usize,
    shapes: usize,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct Outcome {
    source: String,
    repeat: usize,
    budget: &'static str,
    budget_bytes: usize,
    target1: Target,
    target2: Target,
    before_revision: String,
    after_revision: String,
    commit_revision: String,
    before_semantic: SemanticDigest,
    after_semantic: SemanticDigest,
    source_charge: usize,
    clone_charge: usize,
    clone_after_release: usize,
    source_after_clone_release: usize,
    transaction_charge: usize,
    transaction_after_release: usize,
    transaction_after_first_edit: usize,
    transaction_after_second_edit: usize,
    commit_charge: usize,
    commit_snapshot_charge: usize,
    release_commit_charge: usize,
    commit_after_release: usize,
    commit_reference_charge: usize,
    published_charge: usize,
    patch_bytes: Vec<u8>,
    patch_changed: bool,
    first_changed: bool,
    second_changed: bool,
    target1_text_sha256: String,
    target2_text_sha256: String,
}

fn parse_bounded(value: &str, label: &str, maximum: usize) -> Fallible<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 || parsed > maximum {
        return Err(format!("{label} must be between 1 and {maximum}, got {parsed}").into());
    }
    Ok(parsed)
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut result = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        result.push(char::from(DIGITS[(byte >> 4) as usize]));
        result.push(char::from(DIGITS[(byte & 0x0f) as usize]));
    }
    result
}

fn sha256(bytes: &[u8]) -> String {
    hex(Sha256::digest(bytes).as_ref())
}

fn update_field(hasher: &mut Sha256, bytes: &mut usize, label: &str, value: &[u8]) {
    hasher.update((label.len() as u64).to_le_bytes());
    hasher.update(label.as_bytes());
    hasher.update((value.len() as u64).to_le_bytes());
    hasher.update(value);
    *bytes = bytes.saturating_add(value.len());
}

fn semantic_digest(package: &Package) -> Fallible<SemanticDigest> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    if slides.len() > MAX_OBSERVED_SLIDES {
        return Err(format!(
            "semantic observer admits at most {MAX_OBSERVED_SLIDES} slides, got {}",
            slides.len()
        )
        .into());
    }
    let mut hasher = Sha256::new();
    let mut payload_bytes = 0usize;
    let mut shape_count = 0usize;
    for (slide_index, slide) in slides.iter().enumerate() {
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-index",
            &(slide_index as u64).to_le_bytes(),
        );
        let name = slide.name()?;
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-name",
            name.as_bytes(),
        );
        let text = slide.text()?;
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-text",
            text.as_bytes(),
        );
        let scene = slide.shapes()?;
        if scene.len() > MAX_OBSERVED_SHAPES {
            return Err(format!(
                "semantic observer admits at most {MAX_OBSERVED_SHAPES} shapes per slide, got {}",
                scene.len()
            )
            .into());
        }
        shape_count = shape_count.saturating_add(scene.len());
        for (shape_index, shape) in scene.iter().enumerate() {
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-index",
                &(shape_index as u64).to_le_bytes(),
            );
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-name",
                shape.common().name().unwrap_or("").as_bytes(),
            );
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-text",
                shape.common().text().unwrap_or("").as_bytes(),
            );
        }
    }
    Ok(SemanticDigest {
        digest: hex(hasher.finalize().as_ref()),
        payload_bytes,
        slides: slides.len(),
        shapes: shape_count,
    })
}

fn shape_text(package: &Package, target: Target) -> Fallible<String> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let slide = slides
        .get(target.slide)
        .ok_or_else(|| format!("slide {} is out of range", target.slide))?;
    let scene = slide.shapes()?;
    let shape = scene
        .at(target.shape)
        .map_err(|error| format!("target shape lookup failed: {error:?}"))?;
    Ok(shape.common().text().unwrap_or("").to_owned())
}

fn generated_bytes() -> Fallible<Vec<u8>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for slide_index in 0..GENERATED_SLIDES {
            let slide = presentation.add_slide()?;
            for shape_index in 0..GENERATED_SHAPES {
                slide.add_text_box(
                    &format!(
                        "litchi-perf-baseline-pptx-semantic-v1-source-{slide_index:05}-{shape_index:05}"
                    ),
                    36 + i64::try_from(shape_index % 4)? * 180,
                    36 + i64::try_from(shape_index / 4)? * 90,
                    144,
                    54,
                );
            }
        }
    }
    Ok(package.to_bytes()?)
}

fn source_bytes(source: &str) -> Fallible<Vec<u8>> {
    if source == "generated:12x8" {
        return generated_bytes();
    }
    if source.starts_with("generated:") {
        return Err("only the bounded generated:12x8 fixture is admitted".into());
    }
    let bytes = std::fs::read(source)?;
    if bytes.len() > MAX_SOURCE_BYTES {
        return Err(format!(
            "retention observer admits at most {MAX_SOURCE_BYTES} source bytes, got {}",
            bytes.len()
        )
        .into());
    }
    Ok(bytes)
}

fn source_targets(source: &str) -> [Target; 2] {
    if source == "generated:12x8" {
        [Target { slide: 0, shape: 0 }, Target { slide: 1, shape: 0 }]
    } else {
        [Target { slide: 1, shape: 1 }, Target { slide: 2, shape: 0 }]
    }
}

fn revision(snapshot: &litchi_pptx::opened::Snapshot) -> String {
    hex(&snapshot.revision())
}

fn limits_for(budget: &'static str) -> (Limits, usize) {
    match budget {
        "default" => (
            Limits::default().with_max_retained_mce_bytes(DEFAULT_MCE_BYTES),
            DEFAULT_MCE_BYTES,
        ),
        "off" => (Limits::default().with_max_retained_mce_bytes(0), 0),
        "tiny" => (
            Limits::default().with_max_retained_mce_bytes(TINY_MCE_BYTES),
            TINY_MCE_BYTES,
        ),
        _ => unreachable!("budget names are static and closed"),
    }
}

fn run_once(
    source: &str,
    archive: &[u8],
    repeat: usize,
    budget: &'static str,
) -> Fallible<Outcome> {
    let targets = source_targets(source);
    let (limits, budget_bytes) = limits_for(budget);
    let mut package = Package::from_bytes(archive)?;
    let before_semantic = semantic_digest(&package)?;
    let snapshot = package.opened_presentation_with_limits(limits)?;
    let before_revision = revision(&snapshot);
    let source_charge = snapshot.retained_mce_bytes();

    let mut clone = snapshot.clone();
    let clone_charge = clone.retained_mce_bytes();
    clone.release_retained_mce();
    let clone_after_release = clone.retained_mce_bytes();
    let source_after_clone_release = snapshot.retained_mce_bytes();
    if source_after_clone_release != source_charge {
        return Err("releasing one snapshot clone changed the other snapshot".into());
    }
    drop(clone);

    let mut transaction_observer = snapshot.edit();
    let transaction_charge = transaction_observer.retained_mce_bytes();
    transaction_observer.release_retained_mce();
    let transaction_after_release = transaction_observer.retained_mce_bytes();
    if snapshot.retained_mce_bytes() != source_charge {
        return Err("releasing a transaction changed its source snapshot".into());
    }
    drop(transaction_observer);

    let mut transaction = snapshot.edit();
    let first_changed =
        transaction.set_shape_text(targets[0].slide, targets[0].shape, EDIT_MARKER)?;
    let transaction_after_first_edit = transaction.retained_mce_bytes();
    let second_changed =
        transaction.set_shape_text(targets[1].slide, targets[1].shape, EDIT_MARKER)?;
    let transaction_after_second_edit = transaction.retained_mce_bytes();
    if !first_changed || !second_changed {
        return Err("the bounded observer target edit did not change both shapes".into());
    }

    // A separate commit exercises the value-preserving release operation.  A
    // clone of its snapshot proves that the release drops this owner's
    // reference only; the revision must remain exact.
    let mut release_transaction = snapshot.edit();
    release_transaction.set_shape_text(targets[0].slide, targets[0].shape, EDIT_MARKER)?;
    let mut release_commit = release_transaction.commit()?;
    let release_revision = revision(release_commit.snapshot());
    let release_reference = release_commit.snapshot().clone();
    let release_charge = release_commit.retained_mce_bytes();
    release_commit.release_retained_mce();
    let release_after = release_commit.retained_mce_bytes();
    if release_after != 0
        || revision(release_commit.snapshot()) != release_revision
        || release_reference.revision() != release_commit.snapshot().revision()
    {
        return Err("commit retention release changed its revision or charge".into());
    }
    let commit_reference_charge = release_reference.retained_mce_bytes();
    if release_charge > 0 && commit_reference_charge == 0 {
        return Err("releasing a commit dropped its independent snapshot reference".into());
    }
    drop(release_reference);

    let commit = transaction.commit()?;
    let commit_charge = commit.retained_mce_bytes();
    let commit_snapshot_charge = commit.snapshot().retained_mce_bytes();
    if commit_charge != commit_snapshot_charge {
        return Err("commit and commit snapshot retention observations disagree".into());
    }
    let commit_revision = revision(commit.snapshot());
    let patch_changed = commit.is_changed();
    if !patch_changed {
        return Err("two target edits produced an unchanged patch".into());
    }
    let patch_bytes = commit.patch().to_bytes()?;
    let published = package.apply_opened_presentation_commit(commit)?;
    let after_revision = revision(&published);
    if after_revision != commit_revision {
        return Err("publication changed the committed revision".into());
    }
    let after_semantic = semantic_digest(&package)?;
    let target1_text = shape_text(&package, targets[0])?;
    let target2_text = shape_text(&package, targets[1])?;
    if target1_text != EDIT_MARKER || target2_text != EDIT_MARKER {
        return Err("published target text did not round-trip through the public API".into());
    }

    Ok(Outcome {
        source: source.to_owned(),
        repeat,
        budget,
        budget_bytes,
        target1: targets[0],
        target2: targets[1],
        before_revision,
        after_revision,
        commit_revision,
        before_semantic,
        after_semantic,
        source_charge,
        clone_charge,
        clone_after_release,
        source_after_clone_release,
        transaction_charge,
        transaction_after_release,
        transaction_after_first_edit,
        transaction_after_second_edit,
        commit_charge,
        commit_snapshot_charge,
        release_commit_charge: release_charge,
        commit_after_release: release_after,
        commit_reference_charge,
        published_charge: published.retained_mce_bytes(),
        patch_bytes,
        patch_changed,
        first_changed,
        second_changed,
        target1_text_sha256: sha256(target1_text.as_bytes()),
        target2_text_sha256: sha256(target2_text.as_bytes()),
    })
}

fn print_semantic(prefix: &str, semantic: &SemanticDigest) {
    println!("{prefix}_semantic_sha256\t{}", semantic.digest);
    println!(
        "{prefix}_semantic_payload_bytes\t{}",
        semantic.payload_bytes
    );
    println!("{prefix}_slides\t{}", semantic.slides);
    println!("{prefix}_shapes\t{}", semantic.shapes);
}

fn print_outcome(outcome: &Outcome) {
    println!(
        "record\t{}\t{}\t{}",
        outcome.source, outcome.budget, outcome.repeat
    );
    println!("budget_bytes\t{}", outcome.budget_bytes);
    println!(
        "target1\tslide:{}\tshape:{}",
        outcome.target1.slide, outcome.target1.shape
    );
    println!(
        "target2\tslide:{}\tshape:{}",
        outcome.target2.slide, outcome.target2.shape
    );
    println!("before_revision\t{}", outcome.before_revision);
    println!("commit_revision\t{}", outcome.commit_revision);
    println!("after_revision\t{}", outcome.after_revision);
    print_semantic("before", &outcome.before_semantic);
    print_semantic("after", &outcome.after_semantic);
    println!("source_retained_mce_bytes\t{}", outcome.source_charge);
    println!("clone_retained_mce_bytes\t{}", outcome.clone_charge);
    println!("clone_after_release_bytes\t{}", outcome.clone_after_release);
    println!(
        "source_after_clone_release_bytes\t{}",
        outcome.source_after_clone_release
    );
    println!(
        "transaction_retained_mce_bytes\t{}",
        outcome.transaction_charge
    );
    println!(
        "transaction_after_release_bytes\t{}",
        outcome.transaction_after_release
    );
    println!(
        "transaction_after_first_edit_bytes\t{}",
        outcome.transaction_after_first_edit
    );
    println!(
        "transaction_after_second_edit_bytes\t{}",
        outcome.transaction_after_second_edit
    );
    println!("commit_retained_mce_bytes\t{}", outcome.commit_charge);
    println!(
        "commit_snapshot_retained_mce_bytes\t{}",
        outcome.commit_snapshot_charge
    );
    println!(
        "release_commit_retained_mce_bytes\t{}",
        outcome.release_commit_charge
    );
    println!(
        "commit_after_release_bytes\t{}",
        outcome.commit_after_release
    );
    println!(
        "commit_reference_retained_mce_bytes\t{}",
        outcome.commit_reference_charge
    );
    println!("published_retained_mce_bytes\t{}", outcome.published_charge);
    println!("patch_bytes\t{}", outcome.patch_bytes.len());
    println!("patch_sha256\t{}", sha256(&outcome.patch_bytes));
    println!("first_changed\t{}", outcome.first_changed);
    println!("second_changed\t{}", outcome.second_changed);
    println!("patch_changed\t{}", outcome.patch_changed);
    println!("target1_text_sha256\t{}", outcome.target1_text_sha256);
    println!("target2_text_sha256\t{}", outcome.target2_text_sha256);
}

fn compare_parity(default: &Outcome, other: &Outcome) -> Fallible<()> {
    if default.before_revision != other.before_revision
        || default.after_revision != other.after_revision
        || default.commit_revision != other.commit_revision
        || default.before_semantic.digest != other.before_semantic.digest
        || default.after_semantic.digest != other.after_semantic.digest
        || default.before_semantic.payload_bytes != other.before_semantic.payload_bytes
        || default.after_semantic.payload_bytes != other.after_semantic.payload_bytes
        || default.before_semantic.slides != other.before_semantic.slides
        || default.after_semantic.slides != other.after_semantic.slides
        || default.before_semantic.shapes != other.before_semantic.shapes
        || default.after_semantic.shapes != other.after_semantic.shapes
        || default.patch_bytes != other.patch_bytes
        || default.patch_changed != other.patch_changed
        || default.first_changed != other.first_changed
        || default.second_changed != other.second_changed
        || default.target1_text_sha256 != other.target1_text_sha256
        || default.target2_text_sha256 != other.target2_text_sha256
    {
        return Err(format!(
            "semantic/revision/patch parity failed for {} repeat {} against {}",
            default.source, default.repeat, other.budget
        )
        .into());
    }
    Ok(())
}

fn compare_fresh_repeat(first: &Outcome, second: &Outcome) -> Fallible<()> {
    let mut first_observed = first.clone();
    first_observed.repeat = second.repeat;
    if first_observed != *second {
        return Err(format!(
            "fresh repeat changed raw retention or semantic evidence for {} {}",
            first.source, first.budget
        )
        .into());
    }
    Ok(())
}

fn run(source: &str, repeats: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let mut outcomes = Vec::new();
    for repeat in 0..repeats {
        let default = run_once(source, &archive, repeat, "default")?;
        let off = run_once(source, &archive, repeat, "off")?;
        let tiny = run_once(source, &archive, repeat, "tiny")?;
        compare_parity(&default, &off)?;
        compare_parity(&default, &tiny)?;
        if source == "generated:12x8" {
            if default.source_charge != 0 || default.published_charge != 0 {
                return Err("marker-free generated fixture retained MCE bytes".into());
            }
        } else if default.source_charge == 0 || default.published_charge == 0 {
            return Err(
                "real marker-bearing fixture retained no MCE bytes at capture or publication"
                    .into(),
            );
        }
        outcomes.push(default);
        outcomes.push(off);
        outcomes.push(tiny);
    }
    for repeat in 1..repeats {
        for budget in 0..3 {
            compare_fresh_repeat(&outcomes[budget], &outcomes[repeat * 3 + budget])?;
        }
    }
    println!("probe\t{PROBE_ID}");
    println!("source\t{source}");
    println!("repeats\t{repeats}");
    println!("archive_bytes\t{}", archive.len());
    println!("max_source_bytes\t{MAX_SOURCE_BYTES}");
    println!("max_observed_slides\t{MAX_OBSERVED_SLIDES}");
    println!("max_observed_shapes\t{MAX_OBSERVED_SHAPES}");
    println!("edit_marker\t{EDIT_MARKER}");
    for outcome in &outcomes {
        print_outcome(outcome);
    }
    println!("parity\tdefault_vs_off\ttrue");
    println!("parity\tdefault_vs_tiny\ttrue");
    println!("run_complete\ttrue");
    Ok(())
}

fn usage() -> &'static str {
    "usage: retention-observer <pptx-path|generated:12x8> [repeats<=2]"
}

fn main() -> Fallible<()> {
    let mut args = env::args().skip(1);
    let source = args.next().ok_or(usage())?;
    let repeats = args
        .next()
        .map(|value| parse_bounded(&value, "repeat count", MAX_REPEATS))
        .transpose()?
        .unwrap_or(MAX_REPEATS);
    if args.next().is_some() {
        return Err(usage().into());
    }
    run(&source, repeats)
}
