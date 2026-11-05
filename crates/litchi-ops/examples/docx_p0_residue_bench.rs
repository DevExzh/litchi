//! Paired DOCX timing for the P0 residue of record 0335: picture set and
//! remove, comment set, SmartArt node text set and chart series set.
//!
//! Each case times one edit against the same immutable opened eager owner:
//! the leaf leg stages the `litchi_docx::document::Edit` verb (with the
//! chart or diagram editor work the unified call also does) and commits;
//! the unified leg runs `File::set` or `File::remove` on the canonical path,
//! with path resolution and property admission. Both legs must publish the
//! same parts (checked outside the clock); rounds alternate ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example
//! docx_p0_residue_bench --features docx` (pin the CPU with `taskset`).
//! `DOCX_BENCH_CASE` restricts a run to one case.

#[cfg(feature = "docx")]
use std::hint::black_box;
#[cfg(feature = "docx")]
use std::time::Instant;

#[cfg(feature = "docx")]
use litchi_core::Position;
#[cfg(feature = "docx")]
use litchi_docx::document::{CommentChange, ElementPath, ElementStep, PictureChange};
#[cfg(feature = "docx")]
use litchi_drawingml::diagram::authoring::Spec;
#[cfg(feature = "docx")]
use litchi_drawingml::diagram::edit::Outline;
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
const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

#[cfg(feature = "docx")]
fn paragraph(index: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(index))])
}

#[cfg(feature = "docx")]
fn main() -> Outcome<()> {
    let bytes = fixture()?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let comment = package.document()?.comments()?[0].id();
    let diagram = package.smart_arts()?[0].data_relationship_id.clone();
    let chart = package.document()?.paragraphs()?[179].charts()?[0]
        .relationship_id()
        .to_owned();
    let mut cases = Vec::new();

    cases.push(write(
        "picture_set_size",
        || {
            let mut edit = package.edit()?;
            edit.set_picture_at(
                &paragraph(99),
                Position::new(0),
                &PictureChange::new().with_size(914_400, 457_200),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/body/p[100]/picture[1]",
                props! { "width" => "1in", "height" => "0.5in" },
            )?)
        },
    )?);
    cases.push(write(
        "picture_set_alt",
        || {
            let mut edit = package.edit()?;
            edit.set_picture_at(
                &paragraph(99),
                Position::new(0),
                &PictureChange::new().with_description("The logo"),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[100]/picture[1]", props! { "alt" => "The logo" })?),
    )?);
    cases.push(write(
        "picture_remove",
        || {
            let mut edit = package.edit()?;
            edit.remove_picture_at(&paragraph(99), Position::new(0))?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.remove("/body/p[100]/picture[1]")?),
    )?);
    cases.push(write(
        "comment_set_text",
        || {
            let mut edit = package.edit()?;
            edit.set_comment(comment, &CommentChange::new().with_text("Revised"))?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/comments/comment[1]", props! { "text" => "Revised" })?),
    )?);
    cases.push(write(
        "smartart_node_text",
        || {
            let mut edit = package.edit()?;
            let parts = edit.smartart_parts(&diagram)?.ok_or("the graphic")?;
            let limits = litchi_drawingml::diagram::Limits::default();
            let outline = Outline::read(&parts.data, &limits)?;
            let mut spec = Spec::recognize(&outline, &parts.data, &parts.layout, &limits)?
                .ok_or("authored")?;
            spec.nodes_mut()[0].text = "Design".to_owned();
            let relationship = outline.drawing_relationship.as_deref().unwrap_or("rId0");
            let mut next = parts.clone();
            next.data = spec.data_xml(Some(relationship), &limits)?.into_bytes();
            next.layout = spec.layout_xml(&limits)?.into_bytes();
            next.drawing = Some(
                spec.drawing_xml(5_760_000, 3_240_000, &limits)?
                    .into_bytes(),
            );
            edit.set_smartart(&diagram, next)?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/smartart[1]/node[1]", props! { "text" => "Design" })?),
    )?);
    cases.push(write(
        "chart_series_color",
        || {
            let mut edit = package.edit()?;
            let (xml, _) = edit.chart_parts(&chart)?.ok_or("the chart")?;
            let snapshot = litchi_drawingml::chart::edit::Package::from_xml(xml)?;
            let mut chart_edit = snapshot.edit();
            chart_edit.set_series_color(0, Some("70AD47"))?;
            let xml = chart_edit.commit()?.into_package().xml_bytes().to_vec();
            edit.set_chart(&chart, xml, None)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/body/p[180]/chart[1]/series[1]",
                props! { "color" => "70AD47" },
            )?)
        },
    )?);
    cases.push(write(
        "chart_axis_title",
        || {
            let mut edit = package.edit()?;
            let (xml, _) = edit.chart_parts(&chart)?.ok_or("the chart")?;
            let snapshot = litchi_drawingml::chart::edit::Package::from_xml(xml)?;
            let axis = litchi_drawingml::chart::reader::read(
                snapshot.xml_bytes(),
                &litchi_drawingml::chart::Limits::DEFAULT,
            )?
            .plot_area
            .axes
            .iter()
            .find_map(|axis| match axis {
                litchi_drawingml::chart::Axis::Value(axis) => Some(axis.common.axis_id),
                _ => None,
            })
            .ok_or("a value axis")?;
            let mut chart_edit = snapshot.edit();
            chart_edit.set_axis_title(
                axis,
                Some(litchi_drawingml::chart::data::TitleText::from_string(
                    "Units".to_owned(),
                )),
            )?;
            let xml = chart_edit.commit()?.into_package().xml_bytes().to_vec();
            edit.set_chart(&chart, xml, None)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/body/p[180]/chart[1]/axis[@role=value]",
                props! { "title" => "Units" },
            )?)
        },
    )?);
    // The open of the cold cases alone, as their control.
    cases.push(write(
        "open_control",
        || Ok(litchi_docx::Package::from_bytes(&bytes)?),
        || {
            Ok(litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )?)
        },
    )?);
    // Cold handles: both legs open the bytes first, so the unified leg's
    // SmartArt inventory and chart projection are read on this handle.
    cases.push(write(
        "smartart_node_text_cold",
        || {
            let package = litchi_docx::Package::from_bytes(&bytes)?;
            let mut edit = package.edit()?;
            let parts = edit.smartart_parts(&diagram)?.ok_or("the graphic")?;
            let limits = litchi_drawingml::diagram::Limits::default();
            let outline = Outline::read(&parts.data, &limits)?;
            let mut spec = Spec::recognize(&outline, &parts.data, &parts.layout, &limits)?
                .ok_or("authored")?;
            spec.nodes_mut()[0].text = "Design".to_owned();
            let relationship = outline.drawing_relationship.as_deref().unwrap_or("rId0");
            let mut next = parts.clone();
            next.data = spec.data_xml(Some(relationship), &limits)?.into_bytes();
            next.layout = spec.layout_xml(&limits)?.into_bytes();
            next.drawing = Some(
                spec.drawing_xml(5_760_000, 3_240_000, &limits)?
                    .into_bytes(),
            );
            edit.set_smartart(&diagram, next)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            let file = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )?;
            Ok(file.set("/smartart[1]/node[1]", props! { "text" => "Design" })?)
        },
    )?);
    cases.push(write(
        "chart_series_color_cold",
        || {
            let package = litchi_docx::Package::from_bytes(&bytes)?;
            let mut edit = package.edit()?;
            let (xml, _) = edit.chart_parts(&chart)?.ok_or("the chart")?;
            let snapshot = litchi_drawingml::chart::edit::Package::from_xml(xml)?;
            let mut chart_edit = snapshot.edit();
            chart_edit.set_series_color(0, Some("70AD47"))?;
            let xml = chart_edit.commit()?.into_package().xml_bytes().to_vec();
            edit.set_chart(&chart, xml, None)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            let file = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )?;
            Ok(file.set(
                "/body/p[180]/chart[1]/series[1]",
                props! { "color" => "70AD47" },
            )?)
        },
    )?);
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "fixture": "create(): 200 paragraphs, a comment on p[50], a picture in p[100], a SmartArt in p[150], a column chart in p[180]",
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
fn write(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<litchi_docx::Package>,
    mut unified: impl FnMut() -> Outcome<File>,
) -> Outcome<serde_json::Value> {
    if std::env::var("DOCX_BENCH_CASE").is_ok_and(|only| only != name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    // Output check, outside the clock: both legs publish the same parts.
    let leaf_parts = parts(&leaf()?.to_bytes()?)?;
    let unified_parts = parts(&unified()?.to_bytes()?)?;
    if leaf_parts != unified_parts {
        let differing: Vec<&String> = leaf_parts
            .keys()
            .chain(unified_parts.keys())
            .filter(|name| leaf_parts.get(*name) != unified_parts.get(*name))
            .collect();
        return Err(format!("{name}: the legs publish different parts: {differing:?}").into());
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

/// A created document of 200 paragraphs with one comment, picture,
/// SmartArt graphic and chart.
#[cfg(feature = "docx")]
fn fixture() -> Outcome<Vec<u8>> {
    let mut file = litchi_ops::create(litchi_core::FileFormat::Docx)?;
    for index in 0..200 {
        file = file
            .add(
                "/body",
                NewNode::new("paragraph").prop(
                    "text",
                    format!("Paragraph {index} with some ordinary text."),
                ),
            )?
            .into_file();
    }
    let file = file
        .add(
            "/body/p[50]",
            NewNode::new("comment")
                .prop("author", "Ada")
                .prop("text", "Draft"),
        )?
        .into_file()
        .add("/body/p[100]", NewNode::new("picture").prop("src", PNG))?
        .into_file()
        .add(
            "/body/p[150]",
            NewNode::new("smartart")
                .prop("layout", "process")
                .prop("text", "Plan\nBuild\nShip"),
        )?
        .into_file()
        .add(
            "/body/p[180]",
            NewNode::new("chart")
                .prop("chartType", "column")
                .prop("categories", "Q1,Q2,Q3")
                .prop("data", "East:1,2,3;West:4,5,6"),
        )?
        .into_file();
    Ok(file.to_bytes()?)
}

#[cfg(not(feature = "docx"))]
fn main() {
    eprintln!("enable the docx feature");
}
