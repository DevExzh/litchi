//! Paired DOC timing for the parts and marks of record 0324: paragraph and
//! run styles, hyperlinks, bookmarks, tab stops and section page setup, and
//! the style, list-paragraph, hyperlink and section reads.
//!
//! Each write case times one edit against the same immutable opened owner:
//! the leaf leg stages the `litchi_doc::Edit` verb and commits; the unified
//! leg runs the `File` verb on the matching path, including path
//! resolution, style-name resolution, property admission and the returned
//! path. Both legs must publish the same bytes, checked outside the clock.
//! Each read case times a warm `get` and its `format` against the leaf
//! reads the node answers from. Rounds alternate ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example doc_parts_bench
//! --features doc` (pin the CPU with `taskset`). `DOC_BENCH_CASE` restricts
//! a run to one case.

#[cfg(feature = "doc")]
use std::hint::black_box;
#[cfg(feature = "doc")]
use std::time::Instant;

#[cfg(feature = "doc")]
use litchi_core::Position;
#[cfg(feature = "doc")]
use litchi_doc::body_text::{CharacterValue, FieldPart, HyperlinkInsert, SectionValue, TextTarget};
#[cfg(feature = "doc")]
use litchi_doc::parts::pap::{TabAlignment, TabLeader, TabStop};
#[cfg(feature = "doc")]
use litchi_ops::{File, NewNode, props};
#[cfg(feature = "doc")]
use serde_json::json;

#[cfg(feature = "doc")]
const ROUNDS: usize = 9;
#[cfg(feature = "doc")]
const WARMUP: usize = 2;
#[cfg(feature = "doc")]
const ITERATIONS: usize = 64;

#[cfg(feature = "doc")]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(feature = "doc")]
fn main() -> Outcome<()> {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc/Lists.doc"),
    )?;
    let package = litchi_doc::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes(bytes.clone())?;
    let document = package.document();
    let sheet = document.stylesheet().ok_or("stylesheet")?;
    let index_of = |name: &str| -> Outcome<u16> {
        (0..u16::try_from(sheet.styles().len())?)
            .find(|index| sheet.get(*index).is_some_and(|style| style.name == name))
            .ok_or_else(|| format!("style {name}").into())
    };
    let heading = index_of("Heading 1")?;
    let heading_char = index_of("Heading 1 Char")?;
    let paragraph_end = u32::try_from(
        document
            .body_paragraph_text(Position::new(1))?
            .ok_or("p[2]")?
            .encode_utf16()
            .count(),
    )?;
    let run = file
        .get("/body/p[2]/r[1]")?
        .as_doc()
        .and_then(|node| {
            node.text()
                .ok()
                .flatten()
                .map(|text| text.encode_utf16().count())
        })
        .ok_or("run")?;
    let run_end = u32::try_from(run)?.min(paragraph_end);
    // A package and a handle that already hold one hyperlink.
    let link = HyperlinkInsert {
        instruction: r#"HYPERLINK "http://example.com/""#.to_string(),
        text: "Example".to_string(),
    };
    let linked_package = {
        let mut edit = package.edit()?;
        edit.insert_hyperlink(Position::new(1), paragraph_end, &link)?;
        edit.commit()?.into_package()
    };
    let linked = litchi_ops::open_bytes(linked_package.to_bytes()?)?;
    let field_start = linked_package
        .document()
        .fields()?
        .iter()
        .find(|field| field.field.field_type == litchi_doc::FieldType::Hyperlink)
        .ok_or("link")?
        .field
        .start_cp;
    let mut cases = Vec::new();

    cases.push(write(
        "paragraph_style_set",
        || {
            let mut edit = package.edit()?;
            edit.set_paragraph_style(Position::new(1), heading)?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[2]", props! { "style" => "Heading 1" })?),
    )?);
    cases.push(write(
        "run_style_set",
        || {
            let mut edit = package.edit()?;
            edit.set_character_value(
                TextTarget::BodySpan {
                    paragraph: Position::new(1),
                    start: 0,
                    end: run_end,
                },
                CharacterValue::Style(heading_char),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[2]/r[1]", props! { "rStyle" => "Heading1Char" })?),
    )?);
    cases.push(write(
        "hyperlink_add",
        || {
            let mut edit = package.edit()?;
            edit.insert_hyperlink(Position::new(1), paragraph_end, &link)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[2]",
                    NewNode::new("hyperlink")
                        .prop("url", "http://example.com/")
                        .prop("text", "Example"),
                )?
                .file)
        },
    )?);
    cases.push(write(
        "hyperlink_set_url",
        || {
            let mut edit = linked_package.edit()?;
            edit.replace_text(
                TextTarget::Field {
                    start: field_start,
                    part: FieldPart::Instruction,
                },
                r#" HYPERLINK "http://example.org/" "#,
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(linked.set(
                "/body/p[2]/hyperlink[1]",
                props! { "url" => "http://example.org/" },
            )?)
        },
    )?);
    cases.push(write(
        "bookmark_add",
        || {
            let mut edit = package.edit()?;
            edit.add_bookmark("Probe", Position::new(1), 0, 4)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[2]",
                    NewNode::new("bookmark")
                        .prop("name", "Probe")
                        .prop("text", "This"),
                )?
                .file)
        },
    )?);
    cases.push(write(
        "tab_add",
        || {
            let mut edit = package.edit()?;
            let mut stops = edit.paragraph_tabs(Position::new(1))?;
            stops.push(TabStop {
                position: 2880,
                alignment: TabAlignment::Right,
                leader: TabLeader::Dots,
            });
            edit.set_paragraph_tabs(Position::new(1), &stops)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[2]",
                    NewNode::new("tab")
                        .prop("pos", 2880)
                        .prop("val", "right")
                        .prop("leader", "dot"),
                )?
                .file)
        },
    )?);
    cases.push(write(
        "section_set",
        || {
            let mut edit = package.edit()?;
            edit.set_section_values(
                Position::new(0),
                &[SectionValue::Landscape(true), SectionValue::Columns(2)],
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/section[1]",
                props! { "orientation" => "landscape", "columns" => 2 },
            )?)
        },
    )?);

    let reader = package.document();
    let _ = file.get("/styles/Heading1")?.view().format()?.count();
    cases.push(read(
        "style_get_format",
        || {
            let sheet = reader.stylesheet().ok_or("stylesheet")?;
            let style = sheet.get(heading).ok_or("Heading 1")?;
            let character = style
                .character_properties()
                .map(litchi_doc::parts::chp::CharacterProperties::from_sprm)
                .transpose()?;
            let paragraph = style
                .paragraph_properties()
                .and_then(|upx| upx.get(2..))
                .map(litchi_doc::parts::pap::ParagraphProperties::from_sprm)
                .transpose()?;
            black_box((character, paragraph));
            Ok(())
        },
        || {
            let node = file.get("/styles/Heading1")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let _ = file.get("/body/p[28]")?.view().format()?.count();
    cases.push(read(
        "paragraph_list_format",
        || {
            let paragraph = reader.body_paragraph(Position::new(27))?.ok_or("p[28]")?;
            let binding = reader.paragraph_list_binding(&paragraph).ok_or("listed")?;
            black_box((
                binding.effective_level().number_format,
                binding.effective_start_at(),
            ));
            let istd = paragraph.properties().style_index.ok_or("styled")?;
            black_box(
                reader
                    .stylesheet()
                    .and_then(|sheet| sheet.get(istd))
                    .map(|style| style.name.len()),
            );
            Ok(())
        },
        || {
            let node = file.get("/body/p[28]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let link_reader = linked_package.document();
    let _ = linked
        .get("/body/p[2]/hyperlink[1]")?
        .view()
        .format()?
        .count();
    cases.push(read(
        "hyperlink_get_format",
        || {
            let paragraph = link_reader
                .body_paragraph(Position::new(1))?
                .ok_or("p[2]")?;
            let range = paragraph.cp_range().ok_or("range")?;
            let field = link_reader
                .fields()?
                .into_iter()
                .find(|field| {
                    field.field.start_cp >= range.start() && field.field.start_cp < range.end()
                })
                .ok_or("field")?;
            black_box(field.hyperlink_field());
            black_box(paragraph.runs()?.len());
            Ok(())
        },
        || {
            let node = linked.get("/body/p[2]/hyperlink[1]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let _ = file.get("/section[1]")?.view().format()?.count();
    cases.push(read(
        "section_get_format",
        || {
            black_box(reader.sections().first().ok_or("section")?.page);
            Ok(())
        },
        || {
            let node = file.get("/section[1]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "fixture": "test-data/ole/doc/Lists.doc (40 paragraphs, 7 styles, 5 lists)",
            "rounds": ROUNDS,
            "iterations": ITERATIONS,
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(feature = "doc")]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    samples[samples.len() / 2]
}

#[cfg(feature = "doc")]
#[allow(clippy::cast_precision_loss, reason = "nanosecond medians")]
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

#[cfg(feature = "doc")]
fn skipped(name: &str) -> bool {
    std::env::var("DOC_BENCH_CASE").is_ok_and(|only| only != name)
}

#[cfg(feature = "doc")]
fn write(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<litchi_doc::Package>,
    mut unified: impl FnMut() -> Outcome<File>,
) -> Outcome<serde_json::Value> {
    if skipped(name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    // Output check, outside the clock: both legs publish the same bytes.
    if leaf()?.to_bytes()? != unified()?.to_bytes()? {
        return Err(format!("{name}: the legs publish different bytes").into());
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

#[cfg(feature = "doc")]
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

#[cfg(not(feature = "doc"))]
fn main() {
    eprintln!("enable the `doc` feature");
}
