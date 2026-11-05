//! Paired DOCX paragraph-text fanout timing (record 0251f, after 0089).
//!
//! For N body paragraphs of one run each, the leaf leg stages
//! `Edit::replace_body_paragraph_texts` over all N (or `replace_paragraph_text`
//! for N = 1) and commits; the unified leg runs `File::set` on `/body > p`
//! (or `/body/p[1]` for N = 1) with the same text, which holds a tab and a
//! line feed as 0089's input does. Both legs use eager owners opened once;
//! fixture construction and output checks are outside the clock. Rounds
//! alternate ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example docx_fanout_bench
//! --features docx` (pin the CPU with `taskset`). `DOCX_FANOUT_SIZES` picks
//! the sizes (default `1,100,1000,10000`); `DOCX_FANOUT_PROFILE=leaf` or
//! `unified` loops one leg instead, for a profiler (`DOCX_FANOUT_LOOPS` sets
//! the loop count).

#[cfg(feature = "docx")]
use std::hint::black_box;
#[cfg(feature = "docx")]
use std::time::Instant;

#[cfg(feature = "docx")]
use litchi_core::Position;
#[cfg(feature = "docx")]
use litchi_docx::document::ParagraphTextReplacement;
#[cfg(feature = "docx")]
use litchi_ops::{File, OpenOptions, SourceMode, props};
#[cfg(feature = "docx")]
use serde_json::json;

#[cfg(feature = "docx")]
const ROUNDS: usize = 9;
#[cfg(feature = "docx")]
const WARMUP: usize = 2;
#[cfg(feature = "docx")]
const TEXT: &str = "Replaced\ttext\nwith a break";

#[cfg(feature = "docx")]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(feature = "docx")]
fn main() -> Outcome<()> {
    let sizes: Vec<usize> = std::env::var("DOCX_FANOUT_SIZES")
        .unwrap_or_else(|_| "1,100,1000,10000".to_owned())
        .split(',')
        .map(|size| size.trim().parse())
        .collect::<Result<_, _>>()?;
    let mut cases = Vec::new();
    for size in sizes {
        cases.push(measure(size)?);
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": "docx_fanout_benchmark.v1",
            "rounds": ROUNDS,
            "warmup_rounds": WARMUP,
            "order": "ABBA",
            "text": TEXT,
            "timing": "eager owners opened once; each sample stages and commits one fanout from the same immutable owner; output checks excluded",
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(feature = "docx")]
fn measure(size: usize) -> Outcome<serde_json::Value> {
    let bytes = fixture(size)?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let replacements: Vec<ParagraphTextReplacement> = (0..size)
        .map(|index| ParagraphTextReplacement::new(Position::new(index), TEXT.to_owned()))
        .collect();
    let target = if size == 1 { "/body/p[1]" } else { "/body > p" };
    let leaf = || -> Outcome<litchi_docx::Package> {
        let mut edit = package.edit()?;
        if size == 1 {
            edit.replace_paragraph_text(Position::new(0), TEXT)?;
        } else {
            edit.replace_body_paragraph_texts(&replacements)?;
        }
        Ok(edit.commit()?.into_package())
    };
    let unified = || -> Outcome<File> { Ok(file.set(target, props! { "text" => TEXT })?) };

    // Output checks, outside the clock.
    let leaf_text = leaf()?.document()?.text()?;
    let unified_text = match unified()? {
        File::Docx(handle) => handle.text()?,
        _ => return Err("the unified leg returned a non-DOCX handle".into()),
    };
    if leaf_text.replace('\n', "") != unified_text.replace('\n', "") {
        return Err(format!("fanout {size}: the legs disagree").into());
    }

    // `DOCX_FANOUT_PROFILE=leaf|unified` loops one leg for a profiler.
    if let Ok(leg) = std::env::var("DOCX_FANOUT_PROFILE") {
        let loops = match std::env::var("DOCX_FANOUT_LOOPS") {
            Ok(loops) => loops.parse()?,
            Err(_) => (2_000_000 / size).max(1),
        };
        for _ in 0..loops {
            if leg == "leaf" {
                black_box(leaf()?);
            } else {
                black_box(unified()?);
            }
        }
        return Ok(json!({ "paragraphs": size, "profiled": leg, "loops": loops }));
    }
    let iterations = (2_000 / size).clamp(1, 64);
    let sample = |leg: &mut dyn FnMut() -> Outcome<()>| -> Outcome<f64> {
        let start = Instant::now();
        for _ in 0..iterations {
            leg()?;
        }
        Ok(start.elapsed().as_nanos() as f64 / iterations as f64)
    };
    let mut leaf_leg = || -> Outcome<()> {
        black_box(leaf()?);
        Ok(())
    };
    let mut unified_leg = || -> Outcome<()> {
        black_box(unified()?);
        Ok(())
    };
    for _ in 0..WARMUP {
        sample(&mut leaf_leg)?;
        sample(&mut unified_leg)?;
    }
    let mut leaf_samples = Vec::new();
    let mut unified_samples = Vec::new();
    for round in 0..ROUNDS {
        if round % 2 == 0 {
            leaf_samples.push(sample(&mut leaf_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
        } else {
            unified_samples.push(sample(&mut unified_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
        }
    }
    let leaf_median = median(&mut leaf_samples);
    let unified_median = median(&mut unified_samples);
    Ok(json!({
        "paragraphs": size,
        "iterations_per_sample": iterations,
        "fixture_bytes": bytes.len(),
        "leaf_median_ns": leaf_median,
        "unified_median_ns": unified_median,
        "ratio": unified_median / leaf_median,
        "leaf_samples_ns": leaf_samples,
        "unified_samples_ns": unified_samples,
    }))
}

#[cfg(feature = "docx")]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        (samples[middle - 1] + samples[middle]) / 2.0
    } else {
        samples[middle]
    }
}

/// A package of `size` one-run body paragraphs.
#[cfg(feature = "docx")]
fn fixture(size: usize) -> Outcome<Vec<u8>> {
    use std::io::Write as _;
    let mut body = String::new();
    for index in 0..size {
        body.push_str(&format!(
            "<w:p><w:r><w:t>Paragraph {index} with some ordinary text.</w:t></w:r></w:p>"
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

#[cfg(not(feature = "docx"))]
fn main() {
    eprintln!("docx_fanout_bench needs the docx feature");
}
