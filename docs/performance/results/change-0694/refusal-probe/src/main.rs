//! Retained native refusal probe for opened-PresentationML capture.
//!
//! Every fixture is prepared before the timed loop.  The loop calls only
//! `Package::opened_presentation` on an immutable prepared package; snapshot
//! validation, error formatting, and snapshot destruction happen after the
//! timer stops.  The probe is a guardrail for error precedence and refusal
//! cost, not a production performance claim.

use litchi_pptx::{Error, Package};
use sha2::{Digest, Sha256};
use std::env;
use std::error::Error as StdError;
use std::hint::black_box;
use std::time::Instant;

const PROBE_ID: &str = "0693-refusal";
const DEFAULT_SAMPLES: usize = 100;
const DEFAULT_WARMUPS: usize = 5;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUPS: usize = 10_000;
const GENERATED_SLIDES: usize = 12;
const GENERATED_SHAPES: usize = 8;
const COMPLEX_SLIDES: usize = 3;
const PRESENTATION_PART: &str = "/ppt/presentation.xml";
const FIRST_SLIDE_PART: &str = "/ppt/slides/slide1.xml";
const SECOND_SLIDE_PART: &str = "/ppt/slides/slide2.xml";
const LAST_SLIDE_PART: &str = "/ppt/slides/slide3.xml";
const NOTES_SLIDE_PART: &str = "/ppt/notesSlides/notesSlide1.xml";
const TRANSITIONAL_PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MCE_ROOT_MARKER: &[u8] = b"<p:sld xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"urn:probe:0693\" mc:Ignorable=\"p14\" ";

type Fallible<T> = Result<T, Box<dyn StdError + Send + Sync>>;

struct Case {
    name: &'static str,
    package: Package,
    expected: Expected,
    authoring_base_archive_bytes: usize,
    authoring_base_archive_sha256: String,
    prepared_input_graph_sha256: String,
}

struct Authored {
    package: Package,
    archive_bytes: usize,
    archive_sha256: String,
}

enum Expected {
    Success {
        slides: usize,
        first_name: &'static str,
    },
    Exact(&'static str),
    MissingRelationship(String),
    DuplicateName(String),
}

impl Expected {
    fn description(&self) -> String {
        match self {
            Self::Success { slides, first_name } => {
                format!("ok (slides={slides}, first_name={first_name:?})")
            },
            Self::Exact(debug) => (*debug).to_owned(),
            Self::MissingRelationship(id) => format!(
                "Relationship(\"presentation slide reference is missing relationship '{id}'\")"
            ),
            Self::DuplicateName(debug) => debug.clone(),
        }
    }
}

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn parse_count(value: &str, label: &str, maximum: usize) -> Fallible<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed > maximum {
        return Err(format!("{label} must be at most {maximum}, got {parsed}").into());
    }
    Ok(parsed)
}

fn replace_once(source: &[u8], needle: &[u8], replacement: &[u8]) -> Fallible<Vec<u8>> {
    let Some(offset) = source
        .windows(needle.len())
        .position(|window| window == needle)
    else {
        return Err(format!(
            "fixture mutation did not find {:?}",
            String::from_utf8_lossy(needle)
        )
        .into());
    };
    let mut output = Vec::with_capacity(
        source
            .len()
            .saturating_sub(needle.len())
            .saturating_add(replacement.len()),
    );
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset + needle.len()..]);
    Ok(output)
}

fn authored_package(slides: usize, shapes: usize, notes_on_first: bool) -> Fallible<Authored> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for slide_index in 0..slides {
            let slide = presentation.add_slide()?;
            slide.set_title(&format!("Refusal fixture slide {slide_index}"));
            for shape_index in 0..shapes {
                slide.add_text_box(
                    &format!("fixture-{slide_index:02}-{shape_index:02}"),
                    36 + (shape_index % 4) as i64 * 180,
                    36 + (shape_index / 4) as i64 * 90,
                    144,
                    54,
                );
            }
            if notes_on_first && slide_index == 0 {
                slide.set_notes("complex first slide notes for refusal timing");
            }
        }
    }
    let bytes = package.to_bytes()?;
    Ok(Authored {
        package: Package::from_bytes(&bytes)?,
        archive_bytes: bytes.len(),
        archive_sha256: sha256(&bytes),
    })
}

fn rewrite_part(
    package: &Package,
    wanted: &str,
    rewrite: impl FnOnce(&[u8]) -> Fallible<Vec<u8>>,
) -> Fallible<Package> {
    let mut opc = package.opc()?.clone();
    let part_name = opc
        .iter_parts()
        .find(|part| part.partname().as_str() == wanted)
        .map(|part| part.partname().clone())
        .ok_or_else(|| format!("fixture part not found: {wanted}"))?;
    let source = opc.get_part(&part_name)?.blob().to_vec();
    let replacement = rewrite(&source)?;
    opc.get_part_mut(&part_name)?.set_blob(replacement);
    Ok(Package::from_opc_package(opc)?)
}

fn rewrite_notes_part(
    package: &Package,
    rewrite: impl FnOnce(&[u8]) -> Fallible<Vec<u8>>,
) -> Fallible<Package> {
    rewrite_part(package, NOTES_SLIDE_PART, rewrite)
}

fn add_mce_markers(package: &Package) -> Fallible<Package> {
    let first = rewrite_part(package, FIRST_SLIDE_PART, |source| {
        replace_once(source, b"<p:sld ", MCE_ROOT_MARKER)
    })?;
    rewrite_part(&first, SECOND_SLIDE_PART, |source| {
        replace_once(source, b"<p:sld ", MCE_ROOT_MARKER)
    })
}

fn remove_last_slide_relationship(package: &Package) -> Fallible<(Package, String)> {
    let mut opc = package.opc()?.clone();
    let presentation_name = opc
        .iter_parts()
        .find(|part| part.partname().as_str() == PRESENTATION_PART)
        .map(|part| part.partname().clone())
        .ok_or("presentation part is missing")?;
    let relationship_id = opc
        .get_part(&presentation_name)?
        .rels()
        .iter()
        .find(|relationship| relationship.target_ref().ends_with("slide3.xml"))
        .map(|relationship| relationship.r_id().to_owned())
        .ok_or("last slide relationship is missing before mutation")?;
    opc.get_part_mut(&presentation_name)?
        .rels_mut()
        .remove(&relationship_id);
    Ok((Package::from_opc_package(opc)?, relationship_id))
}

fn graph_digest(package: &Package) -> Fallible<String> {
    let opc = package.opc()?;
    let mut parts = Vec::new();
    for part in opc.try_iter_parts() {
        let part = part?;
        let mut relationships = part
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    format!("{:?}", relationship.target_mode()),
                )
            })
            .collect::<Vec<_>>();
        relationships.sort();
        parts.push((
            part.partname().as_str().to_owned(),
            part.content_type().to_owned(),
            sha256(part.blob()),
            relationships,
        ));
    }
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    let mut root_relationships = opc
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                format!("{:?}", relationship.target_mode()),
            )
        })
        .collect::<Vec<_>>();
    root_relationships.sort();
    Ok(sha256(
        format!("parts={parts:?};root_relationships={root_relationships:?}").as_bytes(),
    ))
}

fn case(
    name: &'static str,
    package: Package,
    expected: Expected,
    authoring_base_archive_bytes: usize,
    authoring_base_archive_sha256: &str,
) -> Fallible<Case> {
    let prepared_input_graph_sha256 = graph_digest(&package)?;
    Ok(Case {
        name,
        package,
        expected,
        authoring_base_archive_bytes,
        authoring_base_archive_sha256: authoring_base_archive_sha256.to_owned(),
        prepared_input_graph_sha256,
    })
}

fn duplicate_name_debug(package: &Package) -> Fallible<String> {
    let error = match package.opened_presentation() {
        Ok(_) => return Err("duplicate-name fixture unexpectedly captured".into()),
        Err(error) => error,
    };
    let debug = format!("{error:?}");
    if !matches!(error, Error::MarkupCompatibility(_) | Error::Decode(_))
        || (!debug.contains("duplicated attribute")
            && !debug.contains("duplicate XML attribute 'name'"))
    {
        return Err(format!("unexpected duplicate-name fixture error: {debug}").into());
    }
    Ok(debug)
}

fn build_cases() -> Fallible<Vec<Case>> {
    let small = authored_package(2, 2, true)?;
    let generated = authored_package(GENERATED_SLIDES, GENERATED_SHAPES, false)?;
    let complex = authored_package(COMPLEX_SLIDES, 8, true)?;
    let complex_archive_bytes = complex.archive_bytes;
    let complex_archive_sha256 = complex.archive_sha256.clone();
    let complex_package = complex.package;
    let mce_complex_package = add_mce_markers(&complex_package)?;

    let early_name = rewrite_part(&complex_package, FIRST_SLIDE_PART, |source| {
        replace_once(
            source,
            b" name=\"Slide 256\"",
            b" name=\"Slide 256\" name=\"duplicate\"",
        )
    })?;
    let early_name_debug = duplicate_name_debug(&early_name)?;
    let late_root = rewrite_part(&complex_package, LAST_SLIDE_PART, |source| {
        replace_once(source, b"<p:sld ", b"<p:wrong ")
    })?;
    let (late_missing_relationship, missing_id) = remove_last_slide_relationship(&complex_package)?;
    let late_root_mce = rewrite_part(&mce_complex_package, LAST_SLIDE_PART, |source| {
        replace_once(source, b"<p:sld ", b"<p:wrong ")
    })?;
    let (late_missing_relationship_mce, missing_id_mce) =
        remove_last_slide_relationship(&mce_complex_package)?;
    let notes_invalid_tail = rewrite_notes_part(&complex_package, |source| {
        replace_once(source, b"</p:notes>", b"<p:broken")
    })?;
    let mixed_conformance = rewrite_part(&complex_package, SECOND_SLIDE_PART, |source| {
        replace_once(source, TRANSITIONAL_PML, STRICT_PML)
    })?;
    let slide_raw_overlimit = rewrite_part(&complex_package, FIRST_SLIDE_PART, |source| {
        let mut oversized = source.to_vec();
        oversized.resize(16 * 1024 * 1024 + 1, b' ');
        Ok(oversized)
    })?;

    let Authored {
        package: small_package,
        archive_bytes: small_archive_bytes,
        archive_sha256: small_archive_sha256,
    } = small;
    let Authored {
        package: generated_package,
        archive_bytes: generated_archive_bytes,
        archive_sha256: generated_archive_sha256,
    } = generated;

    Ok(vec![
        case(
            "small-valid",
            small_package,
            Expected::Success {
                slides: 2,
                first_name: "Slide 256",
            },
            small_archive_bytes,
            &small_archive_sha256,
        )?,
        case(
            "generated-12x8-valid",
            generated_package,
            Expected::Success {
                slides: GENERATED_SLIDES,
                first_name: "Slide 256",
            },
            generated_archive_bytes,
            &generated_archive_sha256,
        )?,
        case(
            "early-name-error",
            early_name,
            Expected::DuplicateName(early_name_debug),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "late-root-error",
            late_root,
            Expected::Exact("Invalid(\"slide part does not have a p:sld root\")"),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "late-root-error-mce",
            late_root_mce,
            Expected::Exact("Invalid(\"slide part does not have a p:sld root\")"),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "late-missing-relationship",
            late_missing_relationship,
            Expected::MissingRelationship(missing_id),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "late-missing-relationship-mce",
            late_missing_relationship_mce,
            Expected::MissingRelationship(missing_id_mce),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "notes-invalid-tail",
            notes_invalid_tail,
            Expected::Exact(
                "Xml(\"syntax error: tag not closed: `>` not found before end of input\")",
            ),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "mixed-conformance",
            mixed_conformance,
            Expected::Exact("Invalid(\"slide conformance differs from presentation\")"),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
        case(
            "slide-raw-overlimit-16m-to-64m",
            slide_raw_overlimit,
            Expected::Exact("Invalid(\"invalid sld root or namespace\")"),
            complex_archive_bytes,
            &complex_archive_sha256,
        )?,
    ])
}

fn assert_capture(
    expected: &Expected,
    result: litchi_pptx::Result<litchi_pptx::opened::Snapshot>,
) -> Fallible<String> {
    match (expected, result) {
        (
            Expected::Success {
                slides: expected_slides,
                first_name: expected_first_name,
            },
            Ok(snapshot),
        ) => {
            if snapshot.slides().len() != *expected_slides {
                return Err(format!(
                    "expected {expected_slides} slides, got {}",
                    snapshot.slides().len()
                )
                .into());
            }
            let actual_first_name = snapshot.slides().first().map(|slide| slide.name());
            if actual_first_name != Some(*expected_first_name) {
                return Err(format!(
                    "expected first slide name {expected_first_name:?}, got {actual_first_name:?}"
                )
                .into());
            }
            black_box(snapshot);
            Ok("ok".to_owned())
        },
        (Expected::Success { .. }, Err(error)) => {
            Err(format!("expected success, got {error:?}").into())
        },
        (Expected::Exact(wanted), Err(error)) => {
            let debug = format!("{error:?}");
            if debug != *wanted {
                return Err(format!("expected {wanted}, got {debug}").into());
            }
            Ok(debug)
        },
        (Expected::Exact(wanted), Ok(snapshot)) => {
            black_box(snapshot);
            Err(format!("expected {wanted}, got success").into())
        },
        (Expected::MissingRelationship(id), Err(error)) => {
            let debug = format!("{error:?}");
            let wanted = format!(
                "Relationship(\"presentation slide reference is missing relationship '{id}'\")"
            );
            if debug != wanted {
                return Err(format!("expected {wanted}, got {debug}").into());
            }
            Ok(debug)
        },
        (Expected::MissingRelationship(id), Ok(snapshot)) => {
            black_box(snapshot);
            Err(format!("expected missing relationship for {id}, got success").into())
        },
        (Expected::DuplicateName(wanted), Err(error)) => {
            let debug = format!("{error:?}");
            if debug != *wanted {
                return Err(format!("expected {wanted}, got {debug}").into());
            }
            Ok(debug)
        },
        (Expected::DuplicateName(wanted), Ok(snapshot)) => {
            black_box(snapshot);
            Err(format!("expected {wanted}, got success").into())
        },
    }
}

fn run_case(case: &Case, samples: usize, warmups: usize) -> Fallible<()> {
    let mut first_error = None;
    for _ in 0..warmups {
        let result = case.package.opened_presentation();
        let observed = assert_capture(&case.expected, result)?;
        if observed != "ok" && first_error.is_none() {
            first_error = Some(observed);
        }
    }

    let mut timings = Vec::with_capacity(samples);
    for _ in 0..samples {
        let started = Instant::now();
        let result = case.package.opened_presentation();
        let elapsed_ns = started.elapsed().as_nanos();
        // The result is inspected only after the timer, and a successful
        // snapshot is dropped in this helper after the timer has stopped.
        let observed = assert_capture(&case.expected, result)?;
        if observed != "ok" && first_error.is_none() {
            first_error = Some(observed);
        }
        timings.push(elapsed_ns);
    }

    println!("case\t{}", case.name);
    println!(
        "authoring_base_archive_bytes\t{}",
        case.authoring_base_archive_bytes
    );
    println!(
        "authoring_base_archive_sha256\t{}",
        case.authoring_base_archive_sha256
    );
    println!(
        "prepared_input_graph_sha256\t{}",
        case.prepared_input_graph_sha256
    );
    println!("expected_error_debug\t{}", case.expected.description());
    if let Some(error) = first_error {
        println!("observed_error_debug\t{error}");
    }
    println!("sample_ns");
    for (index, elapsed_ns) in timings.into_iter().enumerate() {
        println!("{index}\t{elapsed_ns}");
    }
    Ok(())
}

fn usage() -> &'static str {
    "usage: probe0694 matrix [samples [warmups]]"
}

fn matrix(samples: usize, warmups: usize) -> Fallible<()> {
    let cases = build_cases()?;
    println!("probe\t{PROBE_ID}");
    println!("mode\tmatrix");
    println!("samples\t{samples}");
    println!("warmups\t{warmups}");
    println!("cases\t{}", cases.len());
    for case in &cases {
        run_case(case, samples, warmups)?;
    }
    println!("all_iterations_passed\ttrue");
    Ok(())
}

fn main() -> Fallible<()> {
    let arguments: Vec<String> = env::args().skip(1).collect();
    match arguments.as_slice() {
        [mode] if mode == "matrix" => matrix(DEFAULT_SAMPLES, DEFAULT_WARMUPS),
        [mode, samples] if mode == "matrix" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            matrix(samples, DEFAULT_WARMUPS)
        },
        [mode, samples, warmups] if mode == "matrix" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            let warmups = parse_count(warmups, "warmups", MAX_WARMUPS)?;
            matrix(samples, warmups)
        },
        _ => Err(usage().into()),
    }
}
