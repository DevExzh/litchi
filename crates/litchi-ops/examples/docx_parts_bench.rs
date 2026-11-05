//! Paired DOCX timing for the part elements of record 0319: styles,
//! numbering levels, headers and notes, and the tab stop and break marks.
//!
//! Each write case times one edit against the same immutable opened owner:
//! the leaf leg stages the `litchi_docx::document::Edit` catalog verb and
//! commits; the unified leg runs the `File` verb on the matching path,
//! including path resolution, property admission and the read-back of the
//! returned path. Each read case times a warm `get` and its `format`
//! against the leaf catalog read the node answers from. Fixture
//! construction and output checks are outside the clock; rounds alternate
//! ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example docx_parts_bench
//! --features docx` (pin the CPU with `taskset`). `DOCX_BENCH_CASE`
//! restricts a run to one case.

#[cfg(feature = "docx")]
use std::hint::black_box;
#[cfg(feature = "docx")]
use std::time::Instant;

#[cfg(feature = "docx")]
use litchi_core::Position;
#[cfg(feature = "docx")]
use litchi_docx::document::{ElementPath, ElementStep, NewNote, NewStory, NoteKind};
#[cfg(feature = "docx")]
use litchi_docx::header_footer::{Kind as StoryKind, Role};
#[cfg(feature = "docx")]
use litchi_docx::numbering::ListKind;
#[cfg(feature = "docx")]
use litchi_docx::numbering::catalog::LevelChange;
#[cfg(feature = "docx")]
use litchi_docx::properties::Twips;
#[cfg(feature = "docx")]
use litchi_docx::styles::Type;
#[cfg(feature = "docx")]
use litchi_docx::styles::catalog::{NewStyle, StyleChange};
#[cfg(feature = "docx")]
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};
#[cfg(feature = "docx")]
use serde_json::json;

#[cfg(feature = "docx")]
const ROUNDS: usize = 9;
#[cfg(feature = "docx")]
const WARMUP: usize = 2;
#[cfg(feature = "docx")]
const ITERATIONS: usize = 64;

#[cfg(feature = "docx")]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(feature = "docx")]
fn main() -> Outcome<()> {
    let bytes = fixture()?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let header = package.header_footer_entries()?[0].name().to_owned();
    let definition = package
        .numbering_catalog()?
        .ok_or("the fixture has numbering")?
        .abstract_nums()[0]
        .id();
    let mut cases = Vec::new();

    cases.push(write(
        "style_set_bold",
        || {
            let catalog = package.style_catalog()?.ok_or("styles")?;
            let mut run = catalog.run_properties(catalog.get("Heading1").ok_or("Heading1")?)?;
            run.bold = Some(true);
            let mut edit = package.edit()?;
            edit.set_style("Heading1", &StyleChange::default().with_run(run))?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/styles/Heading1", props! { "bold" => true })?),
    )?);
    cases.push(write(
        "style_add",
        || {
            let mut edit = package.edit()?;
            edit.add_style(
                &NewStyle::new("Accent", Type::Paragraph)
                    .with_custom(true)
                    .with_change(StyleChange::default().with_name("Accent")),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add("/styles", NewNode::new("style").prop("id", "Accent"))?
                .into_file())
        },
    )?);
    cases.push(write(
        "level_set_format",
        || {
            let mut edit = package.edit()?;
            edit.set_level(
                definition,
                0,
                &LevelChange::default().with_format(Some("upperRoman".into())),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/numbering/abstractNum[1]/level[1]",
                props! { "format" => "upperRoman" },
            )?)
        },
    )?);
    cases.push(write(
        "header_text",
        || {
            let mut edit = package.edit()?;
            edit.set_story_text(&header, "Revised header")?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/header[1]", props! { "text" => "Revised header" })?),
    )?);
    cases.push(write(
        "footnote_add",
        || {
            let mut edit = package.edit()?;
            edit.add_note_at(
                NoteKind::Footnote,
                &ElementPath::new([ElementStep::Paragraph(Position::new(150))]),
                None,
                &NewNote::new("A note"),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[151]",
                    NewNode::new("footnote").prop("text", "A note"),
                )?
                .into_file())
        },
    )?);
    cases.push(write(
        "paragraph_style_by_name",
        || {
            let mut edit = package.edit()?;
            edit.format_paragraph(Position::new(150), |paragraph| {
                paragraph.style = Some("Heading1".to_owned());
            })?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[151]", props! { "style" => "heading 1" })?),
    )?);

    cases.push(write(
        "tab_set",
        || {
            let mut edit = package.edit()?;
            edit.format_paragraph(Position::new(4), |paragraph| {
                if let Some(stop) = paragraph.tabs.as_mut().and_then(|tabs| tabs.first_mut()) {
                    stop.position = Twips(4320);
                }
            })?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[5]/tab[1]", props! { "pos" => 4320 })?),
    )?);
    cases.push(write(
        "break_add",
        || {
            let mut edit = package.edit()?;
            edit.insert_break_at(
                &ElementPath::new([ElementStep::Paragraph(Position::new(9))]),
                None,
                litchi_docx::RunBreakType::Page,
            )?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.add("/body/p[10]", NewNode::new("break"))?.into_file()),
    )?);

    // Warm reads: the handle has read its parts once.
    let catalog = package.style_catalog()?.ok_or("styles")?;
    let _ = file.get("/styles/Heading1")?.view().format()?.count();
    cases.push(read(
        "style_get_format",
        || {
            let entry = catalog.get("Heading1").ok_or("Heading1")?;
            black_box(catalog.paragraph_properties(entry)?);
            black_box(catalog.run_properties(entry)?);
            Ok(())
        },
        || {
            let node = file.get("/styles/Heading1")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let numbering = package.numbering_catalog()?.ok_or("numbering")?;
    let _ = file.get("/body/p[3]")?.view().format()?.count();
    cases.push(read(
        "paragraph_list_format",
        || {
            let document = package.document()?;
            let paragraph = document.body_paragraph(Position::new(2))?.ok_or("p[3]")?;
            let properties = paragraph.direct_properties()?;
            let numbering_properties = properties.numbering.ok_or("numbered")?;
            black_box(numbering.format_of(
                numbering_properties.id.ok_or("numId")?,
                numbering_properties.level.unwrap_or(0),
            ));
            black_box(
                catalog
                    .get(properties.style.as_deref().unwrap_or("Normal"))
                    .and_then(|entry| entry.name()),
            );
            Ok(())
        },
        || {
            let node = file.get("/body/p[3]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let _ = file.get("/body/p[5]/tab[1]")?.view().format()?.count();
    cases.push(read(
        "tab_get_format",
        || {
            let document = package.document()?;
            let paragraph = document.body_paragraph(Position::new(4))?.ok_or("p[5]")?;
            let tabs = paragraph.direct_properties()?.tabs.ok_or("tabs")?;
            black_box(tabs.first().copied());
            Ok(())
        },
        || {
            let node = file.get("/body/p[5]/tab[1]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "fixture": "Builder document: 200 paragraphs, table-alignment.docx styles, a list, a tab stop, a header",
            "rounds": ROUNDS,
            "iterations": ITERATIONS,
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(feature = "docx")]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    samples[samples.len() / 2]
}

#[cfg(feature = "docx")]
fn paired(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<()>,
    mut unified: impl FnMut() -> Outcome<()>,
) -> Outcome<serde_json::Value> {
    let sample = |leg: &mut dyn FnMut() -> Outcome<()>| -> Outcome<f64> {
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            leg()?;
        }
        Ok(start.elapsed().as_nanos() as f64 / ITERATIONS as f64)
    };
    for _ in 0..WARMUP {
        sample(&mut leaf)?;
        sample(&mut unified)?;
    }
    let mut leaf_samples = Vec::new();
    let mut unified_samples = Vec::new();
    for round in 0..ROUNDS {
        if round % 2 == 0 {
            leaf_samples.push(sample(&mut leaf)?);
            unified_samples.push(sample(&mut unified)?);
            unified_samples.push(sample(&mut unified)?);
            leaf_samples.push(sample(&mut leaf)?);
        } else {
            unified_samples.push(sample(&mut unified)?);
            leaf_samples.push(sample(&mut leaf)?);
            leaf_samples.push(sample(&mut leaf)?);
            unified_samples.push(sample(&mut unified)?);
        }
    }
    let leaf_median = median(&mut leaf_samples);
    let unified_median = median(&mut unified_samples);
    Ok(json!({
        "case": name,
        "leaf_median_ns": leaf_median,
        "unified_median_ns": unified_median,
        "ratio": unified_median / leaf_median,
        "leaf_samples_ns": leaf_samples,
        "unified_samples_ns": unified_samples,
    }))
}

#[cfg(feature = "docx")]
fn skipped(name: &str) -> bool {
    std::env::var("DOCX_BENCH_CASE").is_ok_and(|only| only != name)
}

#[cfg(feature = "docx")]
fn write(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<litchi_docx::Package>,
    mut unified: impl FnMut() -> Outcome<File>,
) -> Outcome<serde_json::Value> {
    if skipped(name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    // Output check, outside the clock: both legs publish the same parts.
    let leaf_bytes = leaf()?.to_bytes()?;
    let unified_bytes = unified()?.to_bytes()?;
    let leaf_parts = parts(&leaf_bytes)?;
    let unified_parts = parts(&unified_bytes)?;
    if leaf_parts != unified_parts {
        return Err(format!("{name}: the legs publish different parts").into());
    }
    paired(
        name,
        || {
            black_box(leaf()?);
            Ok(())
        },
        || {
            black_box(unified()?);
            Ok(())
        },
    )
}

#[cfg(feature = "docx")]
fn read(
    name: &str,
    leaf: impl FnMut() -> Outcome<()>,
    unified: impl FnMut() -> Outcome<()>,
) -> Outcome<serde_json::Value> {
    if skipped(name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    paired(name, leaf, unified)
}

/// The members of a package with their uncompressed payloads.
#[cfg(feature = "docx")]
fn parts(bytes: &[u8]) -> Outcome<std::collections::BTreeMap<String, Vec<u8>>> {
    use std::io::Read as _;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
    let mut members = std::collections::BTreeMap::new();
    for index in 0..archive.len() {
        let mut member = archive.by_index(index)?;
        let mut payload = Vec::new();
        member.read_to_end(&mut payload)?;
        members.insert(member.name().to_owned(), payload);
    }
    Ok(members)
}

/// A Builder document of 200 paragraphs (the third in a list) with the
/// styles of `test-data/ooxml/docx/table-alignment.docx` and one header.
#[cfg(feature = "docx")]
fn fixture() -> Outcome<Vec<u8>> {
    use std::io::Read as _;
    let mut builder = litchi_docx::Builder::new()?;
    for index in 0..200 {
        builder
            .document()?
            .add_paragraph_with_text(&format!("Paragraph {index} with some ordinary text."));
    }
    let package = builder.build()?;
    let mut edit = package.edit()?;
    let list = edit.add_list(ListKind::Decimal, None)?;
    edit.format_paragraph(Position::new(2), |paragraph| {
        let mut numbering = litchi_docx::properties::Numbering::default();
        numbering.id = Some(list.num_id());
        numbering.level = Some(0);
        paragraph.numbering = Some(numbering);
    })?;
    edit.format_paragraph(Position::new(4), |paragraph| {
        paragraph.tabs = Some(vec![litchi_docx::properties::TabStop::new(
            litchi_docx::properties::TabKind::Left,
            Twips(720),
        )]);
    })?;
    edit.add_header_footer(Role::Header, StoryKind::Primary, &NewStory::new("Header"))?;
    let source = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ooxml/docx/table-alignment.docx"),
    )?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(source))?;
    let mut styles = Vec::new();
    archive
        .by_name("word/styles.xml")?
        .read_to_end(&mut styles)?;
    match edit.style_catalog()? {
        Some(_) => {
            let (name, _) = edit
                .related_part(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                )?
                .ok_or("styles relationship")?;
            edit.replace_part(&name, styles)?;
        },
        None => return Err("the Builder document has no styles part".into()),
    }
    let package = edit.commit()?.into_package();
    Ok(package.to_bytes()?)
}

#[cfg(not(feature = "docx"))]
fn main() {
    eprintln!("enable the docx feature");
}
