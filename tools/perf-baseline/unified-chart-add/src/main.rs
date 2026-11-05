//! Paired chart add (record 0270): the native leaf calls versus `File::add`.
//!
//! PPTX: native `opened::Edit::add_chart` with an authored
//! `litchi_drawingml::chart::authoring::Spec` plus one commit, versus
//! unified `File::add("/slide[1]", chart)` with the same chart as OfficeCLI
//! properties. XLSX: native `Spec::to_xml` plus `Package::edit_charts`,
//! `Edit::add` and one commit, versus unified `File::add("/Sheet1", chart)`.
//! Both legs start from the same eager owner, which already holds `N`
//! charts. Corpus generation, opening, warm-up, serialization and the
//! equivalence check are outside the timed interval; each timed sample is one
//! chart add. Samples alternate ABBA; medians are the upper middle.
//!
//! DOCX (record 0283): native `document::Edit::insert_chart_at` on the
//! first paragraph plus one commit, versus `File::add("/body/p[1]", chart)`.
//!
//! `--compare set` (record 0283) pairs the native chart retitle (the
//! source-preserving chart editor's `set_title` and the host's chart-part
//! replacement, one commit) with `File::set(<chart[1]>, title)`, on owners
//! holding `max(N, 1)` charts. Routes `{pptx,xlsx,docx}_chart_set_v1`.
//!
//! `--compare dotted` (record 0271) pairs two unified legs instead: the
//! chart given as `data=` plus the `colors` palette, versus the same chart
//! given as OfficeCLI's dotted per-series keys (`series1.name`,
//! `series1.values`, `series1.color`, ...). Both legs publish identical
//! bytes. Routes `pptx_chart_add_dotted_v1`, `xlsx_chart_add_dotted_v1`.

use litchi_core::FileFormat;
use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
use litchi_drawingml::chart::types::LegendPosition;
use litchi_ops::{File, NewNode, OpenOptions, SourceMode};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 3] = [0, 10, 100];
const DEFAULT_ROUNDS: usize = 9;
const BOUNDS: (i64, i64, i64, i64) = (914_400, 914_400, 5_486_400, 3_657_600);
const CATEGORIES: &str = "Q1,Q2,Q3,Q4";
const DATA: &str = "East:120,135,148,162;West:95,108,115,128";

fn spec() -> Spec {
    Spec::new(Variant::new(Kind::Column))
        .with_title("Sales")
        .with_categories(["Q1", "Q2", "Q3", "Q4"])
        .with_series(Series::new("East", vec![120.0, 135.0, 148.0, 162.0]))
        .with_series(Series::new("West", vec![95.0, 108.0, 115.0, 128.0]))
        .with_legend(Some(LegendPosition::Bottom))
}

fn node(format: FileFormat) -> NewNode<'static> {
    let node = NewNode::new("chart")
        .prop("chartType", "column")
        .prop("title", "Sales")
        .prop("legend", "bottom")
        .prop("categories", CATEGORIES)
        .prop("data", DATA);
    placed(node, format)
}

/// Which keys a unified leg names its series with.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Keys {
    /// `data=` and the `colors` palette.
    Data,
    /// `series{N}.name`, `.values` and `.color`.
    Dotted,
}

fn colored_node(format: FileFormat, keys: Keys) -> NewNode<'static> {
    let node = NewNode::new("chart")
        .prop("chartType", "column")
        .prop("title", "Sales")
        .prop("legend", "bottom")
        .prop("categories", CATEGORIES);
    let node = match keys {
        Keys::Data => node.prop("data", DATA).prop("colors", "4472C4,ED7D31"),
        Keys::Dotted => node
            .prop("series1.name", "East")
            .prop("series1.values", "120,135,148,162")
            .prop("series1.color", "4472C4")
            .prop("series2.name", "West")
            .prop("series2.values", "95,108,115,128")
            .prop("series2.color", "ED7D31"),
    };
    placed(node, format)
}

fn placed(node: NewNode<'static>, format: FileFormat) -> NewNode<'static> {
    match format {
        FileFormat::Pptx => node
            .prop("x", "1in")
            .prop("y", "1in")
            .prop("width", "6in")
            .prop("height", "4in"),
        FileFormat::Docx => node,
        _ => node.prop("x", 0).prop("width", 8).prop("height", 15),
    }
}

/// A deck or workbook that already holds `count` charts.
fn corpus(format: FileFormat, count: usize) -> Result<Vec<u8>> {
    match format {
        FileFormat::Pptx => {
            let package = litchi_pptx::Package::new()?;
            let mut edit = package.edit()?;
            edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
            for index in 0..count {
                edit.add_chart(0_usize, &spec(), &format!("Chart {}", index + 1), BOUNDS)?;
            }
            let mut package = edit.commit()?.into_package();
            Ok(package.to_bytes()?)
        },
        FileFormat::Docx => {
            let mut builder = litchi_docx::Builder::new()?;
            builder.document()?.add_paragraph_with_text("Charts");
            let package = builder.build()?;
            let mut edit = package.edit()?;
            for _ in 0..count {
                edit.insert_chart_at(&first_paragraph(), None, &docx_chart())?;
            }
            Ok(edit.commit()?.into_package().to_bytes()?)
        },
        _ => {
            let package = litchi_xlsx::Package::new()?;
            let xml = spec().to_xml(None)?;
            let mut edit = package.edit_charts()?;
            for index in 0..count {
                let row = u32::try_from(index * 16)?;
                edit.add(
                    0,
                    &litchi_xlsx::builder::chart::Anchor::new(0, row, 8, row + 15),
                    &format!("Chart {}", index + 1),
                    &xml,
                )?;
            }
            Ok(edit.commit()?.into_package().to_bytes()?)
        },
    }
}

enum Native {
    Pptx(litchi_pptx::Package),
    Xlsx(litchi_xlsx::Package),
    Docx(litchi_docx::Package),
}

fn first_paragraph() -> litchi_docx::document::ElementPath {
    litchi_docx::document::ElementPath::new([litchi_docx::document::ElementStep::Paragraph(
        litchi_core::Position::new(0),
    )])
}

/// The unified DOCX add names an unnamed chart after its title.
fn docx_chart() -> litchi_docx::document::NewChart {
    litchi_docx::document::NewChart::new(spec()).with_name("Sales")
}

fn route(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Pptx => "pptx",
        FileFormat::Docx => "docx",
        _ => "xlsx",
    }
}

fn open_native(format: FileFormat, bytes: &[u8]) -> Result<Native> {
    Ok(match format {
        FileFormat::Pptx => Native::Pptx(litchi_pptx::Package::from_bytes(bytes)?),
        FileFormat::Docx => Native::Docx(litchi_docx::Package::from_bytes(bytes)?),
        _ => Native::Xlsx(litchi_xlsx::Package::from_bytes(bytes)?),
    })
}

fn native_once(owner: &Native, count: usize) -> Result<Native> {
    Ok(match owner {
        Native::Pptx(base) => {
            let mut edit = base.edit()?;
            edit.add_chart(0_usize, &spec(), &format!("Chart {}", count + 1), BOUNDS)?;
            Native::Pptx(edit.commit()?.into_package())
        },
        Native::Xlsx(base) => {
            let xml = spec().to_xml(None)?;
            let row = u32::try_from(count * 16)?;
            let mut edit = base.edit_charts()?;
            edit.add(
                0,
                &litchi_xlsx::builder::chart::Anchor::new(0, row, 8, row + 15),
                "Sales",
                &xml,
            )?;
            Native::Xlsx(edit.commit()?.into_package())
        },
        Native::Docx(base) => {
            let mut edit = base.edit()?;
            edit.insert_chart_at(&first_paragraph(), None, &docx_chart())?;
            Native::Docx(edit.commit()?.into_package())
        },
    })
}

/// The native retitle of the first chart.
fn native_set(owner: &Native) -> Result<Native> {
    use litchi_drawingml::chart::TitleText;
    let retitle = |xml: Vec<u8>| -> Result<Vec<u8>> {
        let snapshot = litchi_drawingml::chart::edit::Package::from_xml(xml)?;
        let mut edit = snapshot.edit();
        edit.set_title(Some(TitleText::from_string("Revenue")))?;
        Ok(edit.commit()?.into_package().xml_bytes().to_vec())
    };
    Ok(match owner {
        Native::Pptx(base) => {
            let mut edit = base.edit()?;
            let (xml, _) = edit.chart(0_usize, 0_usize)?;
            edit.set_chart(0_usize, 0_usize, retitle(xml)?, None)?;
            Native::Pptx(edit.commit()?.into_package())
        },
        Native::Xlsx(base) => {
            let mut edit = base.edit_charts()?;
            let xml = edit.chart(0, 0)?;
            edit.set(0, 0, retitle(xml)?)?;
            Native::Xlsx(edit.commit()?.into_package())
        },
        Native::Docx(base) => {
            let charts = base.document()?.paragraphs()?[0].charts()?;
            let id = charts[0].relationship_id().to_owned();
            let mut edit = base.edit()?;
            let (xml, _) = edit.chart_parts(&id)?.ok_or("no chart")?;
            edit.set_chart(&id, retitle(xml)?, None)?;
            Native::Docx(edit.commit()?.into_package())
        },
    })
}

fn chart_path(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Pptx => "/slide[1]/chart[1]",
        FileFormat::Docx => "/body/p[1]/chart[1]",
        _ => "/Sheet1/chart[1]",
    }
}

fn unified_set(file: &File, format: FileFormat) -> Result<File> {
    Ok(file.set(
        chart_path(format),
        litchi_ops::props! { "title" => "Revenue" },
    )?)
}

fn native_bytes(owner: &mut Native) -> Result<Vec<u8>> {
    Ok(match owner {
        Native::Pptx(package) => package.to_bytes()?,
        Native::Xlsx(package) => package.to_bytes()?,
        Native::Docx(package) => package.to_bytes()?,
    })
}

fn unified_once(file: &File, format: FileFormat, count: usize) -> Result<File> {
    unified_with(file, format, count, node(format))
}

fn unified_with(file: &File, format: FileFormat, count: usize, node: NewNode<'_>) -> Result<File> {
    let parent = match format {
        FileFormat::Pptx => "/slide[1]",
        FileFormat::Docx => "/body/p[1]",
        _ => "/Sheet1",
    };
    let node = match format {
        FileFormat::Pptx | FileFormat::Docx => node,
        _ => node.prop("y", i64::try_from(count * 16)?),
    };
    Ok(file.add(parent, node)?.file)
}

/// `--compare dotted`: `data=` (A) versus dotted series keys (B), both
/// unified, on the same eager owner.
fn case_dotted(format: FileFormat, count: usize, rounds: usize) -> Result<Value> {
    let bytes = corpus(format, count)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let leg = |keys: Keys| unified_with(&file, format, count, colored_node(format, keys));
    if leg(Keys::Data)?.to_bytes()? != leg(Keys::Dotted)?.to_bytes()? {
        return Err(format!("{format:?} with {count} charts: the legs differ").into());
    }
    let mut data_samples = Vec::with_capacity(rounds * 2);
    let mut dotted_samples = Vec::with_capacity(rounds * 2);
    for round in 0..rounds * 2 {
        let data_first = (round / 2) % 2 == 0;
        for step in 0..2 {
            let keys = if (step == 0) == data_first {
                Keys::Data
            } else {
                Keys::Dotted
            };
            let start = Instant::now();
            black_box(leg(keys)?);
            let elapsed = start.elapsed().as_nanos();
            match keys {
                Keys::Data => data_samples.push(elapsed),
                Keys::Dotted => dotted_samples.push(elapsed),
            }
        }
    }
    let data_median = median(&mut data_samples.clone());
    let dotted_median = median(&mut dotted_samples.clone());
    Ok(json!({
        "route": format!("{}_chart_add_dotted_v1", if format == FileFormat::Pptx { "pptx" } else { "xlsx" }),
        "owner": "eager",
        "existing_charts": count,
        "corpus_bytes": bytes.len(),
        "corpus_sha256": sha256(&bytes),
        "equivalence": "exact_bytes",
        "data_median_ns": data_median,
        "dotted_median_ns": dotted_median,
        "dotted_over_data": dotted_median as f64 / data_median as f64,
        "data_min_ns": data_samples.iter().min(),
        "dotted_min_ns": dotted_samples.iter().min(),
        "data_samples_ns": data_samples,
        "dotted_samples_ns": dotted_samples,
    }))
}

/// Every archive member, decoded, by name.
fn members(bytes: &[u8]) -> Result<std::collections::BTreeMap<String, Vec<u8>>> {
    let archive = soapberry_zip::office::ArchiveReader::new(
        bytes,
        soapberry_zip::office::ArchiveLimits::default(),
    )?;
    let mut members = std::collections::BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn median(samples: &mut [u128]) -> u128 {
    samples.sort_unstable();
    samples[samples.len() / 2]
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn case(format: FileFormat, count: usize, rounds: usize) -> Result<Value> {
    let bytes = corpus(format, count)?;
    let native = open_native(format, &bytes)?;
    let unified = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    // Warm-up and equivalence: both legs publish the same chart.
    let mut native_after = native_once(&native, count)?;
    let unified_after = unified_once(&unified, format, count)?;
    let native_out = native_bytes(&mut native_after)?;
    let unified_out = unified_after.to_bytes()?;
    if let Ok(dir) = std::env::var("CHART_ADD_DUMP") {
        std::fs::write(format!("{dir}/native-{format:?}-{count}.bin"), &native_out)?;
        std::fs::write(
            format!("{dir}/unified-{format:?}-{count}.bin"),
            &unified_out,
        )?;
    }
    let equivalence = if native_out == unified_out {
        "exact_bytes"
    } else if members(&native_out)? == members(&unified_out)? {
        "decoded_members"
    } else {
        return Err(
            format!("{format:?} with {count} charts: the legs publish different packages").into(),
        );
    };
    let mut native_samples = Vec::with_capacity(rounds * 2);
    let mut unified_samples = Vec::with_capacity(rounds * 2);
    for round in 0..rounds * 2 {
        // ABBA: native first on even pairs, unified first on odd pairs.
        let native_first = (round / 2) % 2 == 0;
        for leg in 0..2 {
            let run_native = (leg == 0) == native_first;
            let start = Instant::now();
            if run_native {
                black_box(native_once(&native, count)?);
                native_samples.push(start.elapsed().as_nanos());
            } else {
                black_box(unified_once(&unified, format, count)?);
                unified_samples.push(start.elapsed().as_nanos());
            }
        }
    }
    let native_median = median(&mut native_samples.clone());
    let unified_median = median(&mut unified_samples.clone());
    Ok(json!({
        "route": format!("{}_chart_add_v1", route(format)),
        "owner": "eager",
        "existing_charts": count,
        "corpus_bytes": bytes.len(),
        "corpus_sha256": sha256(&bytes),
        "equivalence": equivalence,
        "native_median_ns": native_median,
        "unified_median_ns": unified_median,
        "unified_over_native": unified_median as f64 / native_median as f64,
        "native_samples_ns": native_samples,
        "unified_samples_ns": unified_samples,
    }))
}

/// `--compare set`: the native retitle (A) versus `File::set` (B).
fn case_set(format: FileFormat, count: usize, rounds: usize) -> Result<Value> {
    let count = count.max(1);
    let bytes = corpus(format, count)?;
    let native = open_native(format, &bytes)?;
    let unified = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let native_out = native_bytes(&mut native_set(&native)?)?;
    let unified_out = unified_set(&unified, format)?.to_bytes()?;
    let equivalence = if native_out == unified_out {
        "exact_bytes"
    } else if members(&native_out)? == members(&unified_out)? {
        "decoded_members"
    } else {
        return Err(format!("{format:?} set with {count} charts: the legs differ").into());
    };
    let mut native_samples = Vec::with_capacity(rounds * 2);
    let mut unified_samples = Vec::with_capacity(rounds * 2);
    for round in 0..rounds * 2 {
        let native_first = (round / 2) % 2 == 0;
        for leg in 0..2 {
            let start = Instant::now();
            if (leg == 0) == native_first {
                black_box(native_set(&native)?);
                native_samples.push(start.elapsed().as_nanos());
            } else {
                black_box(unified_set(&unified, format)?);
                unified_samples.push(start.elapsed().as_nanos());
            }
        }
    }
    let native_median = median(&mut native_samples.clone());
    let unified_median = median(&mut unified_samples.clone());
    Ok(json!({
        "route": format!("{}_chart_set_v1", route(format)),
        "owner": "eager",
        "existing_charts": count,
        "corpus_bytes": bytes.len(),
        "corpus_sha256": sha256(&bytes),
        "equivalence": equivalence,
        "native_median_ns": native_median,
        "unified_median_ns": unified_median,
        "unified_over_native": unified_median as f64 / native_median as f64,
        "native_min_ns": native_samples.iter().min(),
        "unified_min_ns": unified_samples.iter().min(),
        "native_samples_ns": native_samples,
        "unified_samples_ns": unified_samples,
    }))
}

fn main() -> Result<()> {
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut rounds = DEFAULT_ROUNDS;
    let mut dotted = false;
    let mut set = false;
    let mut formats = vec![FileFormat::Pptx, FileFormat::Xlsx];
    let mut args = std::env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--counts" => {
                counts = args
                    .next()
                    .ok_or("--counts needs a value")?
                    .split(',')
                    .map(str::parse)
                    .collect::<std::result::Result<_, _>>()?;
            },
            "--rounds" => rounds = args.next().ok_or("--rounds needs a value")?.parse()?,
            "--compare" => match args.next().as_deref() {
                Some("native") => dotted = false,
                Some("dotted") => dotted = true,
                Some("set") => set = true,
                _ => return Err("--compare takes native, dotted or set".into()),
            },
            "--formats" => {
                formats = args
                    .next()
                    .ok_or("--formats needs a value")?
                    .split(',')
                    .map(|name| match name {
                        "pptx" => Ok(FileFormat::Pptx),
                        "xlsx" => Ok(FileFormat::Xlsx),
                        "docx" => Ok(FileFormat::Docx),
                        _ => Err(format!("unknown format {name}")),
                    })
                    .collect::<std::result::Result<_, _>>()?;
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    if rounds < 3 {
        return Err("at least three rounds".into());
    }
    let mut results = Vec::new();
    for &format in &formats {
        for &count in &counts {
            results.push(if set {
                case_set(format, count, rounds)?
            } else if dotted {
                case_dotted(format, count, rounds)?
            } else {
                case(format, count, rounds)?
            });
        }
    }
    let report = json!({
        "schema": 1,
        "harness": "unified-chart-add-baseline",
        "compare": if set { "set" } else if dotted { "dotted" } else { "native" },
        "rounds": rounds,
        "provenance": {
            "commit": std::env::var("LITCHI_SOURCE_COMMIT").ok(),
            "rustc": std::env::var("LITCHI_RUSTC_VERSION").ok(),
            "target": std::env::var("LITCHI_TARGET").ok(),
            "build": std::env::var("LITCHI_BUILD_COMMAND").ok(),
        },
        "results": results,
    });
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}
