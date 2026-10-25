//! Paired find/replace timing (record 0273): the unified `File::set` with
//! `find=`/`replace=` against the best hand-written leaf equivalent.
//!
//! - DOCX: a body of 10,000 two-run paragraphs. The leaf leg reads the
//!   main story (`Edit::story`, the substitution primitive find shares with
//!   template merge since record 0285), finds `needle` in each text segment
//!   with `str::match_indices` and writes every match in one
//!   `Edit::substitute` call. It knows the fixture never splits a match
//!   across segments, so it skips grouping and cross-segment matching: a
//!   lower bound.
//! - PPTX: 1,000 slides with one text box each. The leaf leg reads each
//!   slide's shape scene, and stages one run-text edit per matching shape,
//!   knowing each body is one paragraph of one run: a lower bound too.
//!
//! - DOCX paragraph scope (record 0285): the same body, `needle` in two
//!   paragraphs, and a find over the middle one (`/body/p[5001]`). The leaf
//!   leg reads that paragraph's story (`Edit::story_within`) and substitutes
//!   its match.
//!
//! `M` paragraphs or slides hold the word `needle` (1 or 1,000). Owners are
//! opened eagerly once; fixture construction and output checks are outside
//! the clock; every sample edits the same immutable owner and commits
//! without publishing. Two warm-up rounds, then ABBA rounds.
//!
//! Run with `cargo run --release -p litchi-ops --example find_replace_bench
//! --features docx,pptx` under `taskset`.

#[cfg(all(feature = "docx", feature = "pptx"))]
use std::hint::black_box;
#[cfg(all(feature = "docx", feature = "pptx"))]
use std::time::Instant;

#[cfg(all(feature = "docx", feature = "pptx"))]
use serde_json::json;

#[cfg(all(feature = "docx", feature = "pptx"))]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(all(feature = "docx", feature = "pptx"))]
const ROUNDS: usize = 9;
#[cfg(all(feature = "docx", feature = "pptx"))]
const WARMUP: usize = 2;

#[cfg(all(feature = "docx", feature = "pptx"))]
fn main() -> Outcome<()> {
    let paragraphs: usize =
        std::env::var("FIND_DOCX_PARAGRAPHS").map_or(Ok(10_000), |value| value.parse())?;
    let slides: usize =
        std::env::var("FIND_PPTX_SLIDES").map_or(Ok(1_000), |value| value.parse())?;
    // `FIND_BENCH_ONLY=docx-paragraph` runs that case alone (for a
    // profiler).
    let only = std::env::var("FIND_BENCH_ONLY").ok();
    let mut cases = Vec::new();
    if only.is_none() {
        for matches in [1, 1_000] {
            cases.push(docx(paragraphs, matches)?);
            cases.push(pptx(slides, matches)?);
        }
    }
    if only.as_deref().is_none_or(|only| only == "docx-paragraph") {
        cases.push(docx_paragraph(paragraphs)?);
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": "find_replace_benchmark.v1",
            "rounds": ROUNDS,
            "warmup_rounds": WARMUP,
            "order": "ABBA",
            "timing": "eager owners opened once; each sample finds, replaces and commits from the same immutable owner; publication and output checks excluded",
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(all(feature = "docx", feature = "pptx"))]
fn docx(paragraphs: usize, matches: usize) -> Outcome<serde_json::Value> {
    use litchi_docx::document::substitute::Replacement;
    use litchi_ops::{File, OpenOptions, SourceMode, props};

    let bytes = docx_fixture(paragraphs, matches)?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let leaf = || -> Outcome<litchi_docx::Package> {
        let mut edit = package.edit()?;
        let story = edit.story()?;
        let mut replacements = Vec::new();
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            let text = paragraph.text();
            for segment in paragraph.segments() {
                let range = segment.range();
                for (at, _) in text[range.clone()].match_indices("needle") {
                    let start = range.start + at;
                    replacements.push(Replacement::new(index, start..start + 6, "pin"));
                }
            }
        }
        edit.substitute(&story, &replacements)?;
        Ok(edit.commit()?.into_package())
    };
    let unified = || -> Outcome<File> {
        Ok(file.set("/", props! { "find" => "needle", "replace" => "pin" })?)
    };
    let leaf_package = leaf()?;
    let leaf_text = leaf_package.document()?.text()?;
    if leaf_text.matches("pin").count() != matches
        || leaf_package.to_bytes()? != unified()?.to_bytes()?
    {
        return Err("DOCX: the legs disagree".into());
    }
    paired(
        "docx",
        paragraphs,
        matches,
        bytes.len(),
        &mut || {
            black_box(leaf()?);
            Ok(())
        },
        &mut || {
            black_box(unified()?);
            Ok(())
        },
    )
}

/// A find scoped to one paragraph in the middle of the body.
#[cfg(all(feature = "docx", feature = "pptx"))]
fn docx_paragraph(paragraphs: usize) -> Outcome<serde_json::Value> {
    use litchi_core::Position;
    use litchi_docx::document::substitute::Replacement;
    use litchi_docx::document::{ElementPath, ElementStep};
    use litchi_ops::{File, OpenOptions, SourceMode, props};

    let bytes = docx_fixture(paragraphs, 2)?;
    let target = paragraphs / 2;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let leaf = || -> Outcome<litchi_docx::Package> {
        let mut edit = package.edit()?;
        let path = ElementPath::new([ElementStep::Paragraph(Position::new(target))]);
        let story = edit.story_within(&[path])?;
        let mut replacements = Vec::new();
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            let text = paragraph.text();
            for segment in paragraph.segments() {
                let range = segment.range();
                for (at, _) in text[range.clone()].match_indices("needle") {
                    let start = range.start + at;
                    replacements.push(Replacement::new(index, start..start + 6, "pin"));
                }
            }
        }
        edit.substitute(&story, &replacements)?;
        Ok(edit.commit()?.into_package())
    };
    let path = format!("/body/p[{}]", target + 1);
    let unified = || -> Outcome<File> {
        Ok(file.set(
            path.as_str(),
            props! { "find" => "needle", "replace" => "pin" },
        )?)
    };
    let leaf_package = leaf()?;
    if leaf_package.document()?.text()?.matches("pin").count() != 1
        || leaf_package.to_bytes()? != unified()?.to_bytes()?
    {
        return Err("DOCX paragraph scope: the legs disagree".into());
    }
    paired(
        "docx-paragraph",
        1,
        1,
        bytes.len(),
        &mut || {
            black_box(leaf()?);
            Ok(())
        },
        &mut || {
            black_box(unified()?);
            Ok(())
        },
    )
}

#[cfg(all(feature = "docx", feature = "pptx"))]
fn pptx(slides: usize, matches: usize) -> Outcome<serde_json::Value> {
    use litchi_core::Position;
    use litchi_ops::{File, OpenOptions, SourceMode, props};
    use litchi_pptx::opened::{TextAssignment, TextProperties, TextTarget};

    let bytes = pptx_fixture(slides, matches)?;
    let package = litchi_pptx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let opc = package.opc()?;
    let leaf = || -> Outcome<litchi_pptx::Package> {
        let mut edit = package.edit()?;
        let parts: Vec<_> = edit
            .slides()?
            .iter()
            .map(|slide| slide.part_name().clone())
            .collect();
        for (slide, part) in parts.iter().enumerate() {
            let part = opc.get_part(part)?;
            let scene =
                litchi_pptx::shape::Scene::read(part.blob(), litchi_pptx::shape::Limits::DEFAULT)?;
            let mut texts = Vec::new();
            for index in 0..scene.len() {
                if let Some(text) = scene.at(index)?.text()
                    && text.contains("needle")
                {
                    texts.push((index, text.replace("needle", "pin")));
                }
            }
            let edits: Vec<TextAssignment<'_>> = texts
                .iter()
                .map(|(index, text)| {
                    TextAssignment::new(
                        TextTarget::Run {
                            shape: Position::new(*index),
                            paragraph: Position::new(0),
                            run: Position::new(0),
                        },
                        TextProperties::default().with_text(text),
                    )
                })
                .collect();
            if !edits.is_empty() {
                edit.set_text_properties(slide, &edits)?;
            }
        }
        Ok(edit.commit()?.into_package())
    };
    let unified = || -> Outcome<File> {
        Ok(file.set("/", props! { "find" => "needle", "replace" => "pin" })?)
    };
    let leaf_bytes = leaf()?.to_bytes()?;
    let unified_bytes = unified()?.to_bytes()?;
    if leaf_bytes != unified_bytes {
        return Err("PPTX: the legs publish different bytes".into());
    }
    paired(
        "pptx",
        slides,
        matches,
        bytes.len(),
        &mut || {
            black_box(leaf()?);
            Ok(())
        },
        &mut || {
            black_box(unified()?);
            Ok(())
        },
    )
}

#[cfg(all(feature = "docx", feature = "pptx"))]
fn paired(
    format: &str,
    size: usize,
    matches: usize,
    fixture_bytes: usize,
    leaf: &mut dyn FnMut() -> Outcome<()>,
    unified: &mut dyn FnMut() -> Outcome<()>,
) -> Outcome<serde_json::Value> {
    let sample = |leg: &mut dyn FnMut() -> Outcome<()>| -> Outcome<f64> {
        let start = Instant::now();
        leg()?;
        Ok(start.elapsed().as_nanos() as f64)
    };
    for _ in 0..WARMUP {
        sample(leaf)?;
        sample(unified)?;
    }
    let mut leaf_samples = Vec::new();
    let mut unified_samples = Vec::new();
    for round in 0..ROUNDS {
        if round % 2 == 0 {
            leaf_samples.push(sample(leaf)?);
            unified_samples.push(sample(unified)?);
            unified_samples.push(sample(unified)?);
            leaf_samples.push(sample(leaf)?);
        } else {
            unified_samples.push(sample(unified)?);
            leaf_samples.push(sample(leaf)?);
            leaf_samples.push(sample(leaf)?);
            unified_samples.push(sample(unified)?);
        }
    }
    let leaf_median = median(&mut leaf_samples);
    let unified_median = median(&mut unified_samples);
    Ok(json!({
        "format": format,
        "containers": size,
        "matches": matches,
        "fixture_bytes": fixture_bytes,
        "leaf_median_ns": leaf_median,
        "unified_median_ns": unified_median,
        "ratio": unified_median / leaf_median,
        "unified_containers_per_second": size as f64 / (unified_median / 1e9),
        "leaf_samples_ns": leaf_samples,
        "unified_samples_ns": unified_samples,
    }))
}

#[cfg(all(feature = "docx", feature = "pptx"))]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        (samples[middle - 1] + samples[middle]) / 2.0
    } else {
        samples[middle]
    }
}

/// `size` two-run body paragraphs; the first `matches` (spread evenly)
/// hold `needle`.
#[cfg(all(feature = "docx", feature = "pptx"))]
fn docx_fixture(size: usize, matches: usize) -> Outcome<Vec<u8>> {
    use std::io::Write as _;
    let stride = (size / matches.max(1)).max(1);
    let mut body = String::new();
    for index in 0..size {
        let word = if index % stride == 0 && index / stride < matches {
            "needle"
        } else {
            "thread"
        };
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">Paragraph {index} has a </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{word} in its ordinary text.</w:t></w:r></w:p>"#
        ));
    }
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options)?;
        zip.write_all(contents.as_bytes())?;
    }
    Ok(zip.finish()?.into_inner())
}

/// `size` slides of one text box each; `matches` of them hold `needle`.
#[cfg(all(feature = "docx", feature = "pptx"))]
fn pptx_fixture(size: usize, matches: usize) -> Outcome<Vec<u8>> {
    let stride = (size / matches.max(1)).max(1);
    let mut builder = litchi_pptx::builder::Builder::new()?;
    let presentation = builder.presentation()?;
    for index in 0..size {
        let word = if index % stride == 0 && index / stride < matches {
            "needle"
        } else {
            "thread"
        };
        presentation.add_slide()?.add_text_box(
            &format!("Slide {index} has a {word} in its text"),
            0,
            0,
            1_828_800,
            914_400,
        );
    }
    Ok(builder.build()?.to_bytes()?)
}

#[cfg(not(all(feature = "docx", feature = "pptx")))]
fn main() {
    eprintln!("find_replace_bench needs the docx and pptx features");
}
