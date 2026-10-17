//! Paired reads and writes of the record 0321 PPTX elements: the native
//! leaf calls versus the unified verbs, on eager owners of a deck with `N`
//! slides, each holding a text box with a run hyperlink and a 3×3 table
//! (the equation route reads a tracked fixture).
//!
//! Writes (one operation and one commit per sample):
//! - `pptx_theme_set_v1`: `Edit::{theme, set_theme_colors}` versus
//!   `File::set("/theme", {accent1})`;
//! - `pptx_master_name_v1`: `Edit::set_master_name` versus
//!   `File::set("/slidemaster[1]", {name})`;
//! - `pptx_column_width_v1`: `Edit::set_table_column_width` versus
//!   `File::set(".../col[2]", {width})`;
//! - `pptx_column_add_v1`: `Edit::insert_table_column` versus
//!   `File::add(".../table[1]", col{width, text})`;
//! - `pptx_hyperlink_set_v1`: `Edit::set_shape_hyperlink` versus
//!   `File::add(".../shape[1]", hyperlink{link})`.
//!
//! Reads (one read per sample, on the last slide):
//! - `pptx_hyperlink_get_v1`: the run's `Click::from_text` and
//!   `Slide::hyperlink` versus `File::get(".../run[1]/hyperlink")` and its
//!   `format`;
//! - `pptx_equation_get_v1`: `Slide::equations` on the first slide versus
//!   `File::get("/slide[1]/shape[1]/equation[1]")` and its `format`;
//! - `pptx_theme_get_v1`: `Package::theme` versus `File::get("/theme")` and
//!   its `format`.
//!
//! Corpus generation, opening, warm-up and the equivalence check (the
//! write legs publish the same decoded slide, theme or master part; the
//! read legs read the same value) are outside the timed interval. Legs
//! alternate ABBA; medians are the upper middle.

use litchi_core::Position;
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};
use litchi_pptx::hyperlinks::{Click, Link, Target};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 2] = [1, 50];
const DEFAULT_ROUNDS: usize = 9;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Route {
    ThemeSet,
    MasterName,
    ColumnWidth,
    ColumnAdd,
    HyperlinkSet,
    HyperlinkGet,
    EquationGet,
    ThemeGet,
}

const ROUTES: [Route; 8] = [
    Route::ThemeSet,
    Route::MasterName,
    Route::ColumnWidth,
    Route::ColumnAdd,
    Route::HyperlinkSet,
    Route::HyperlinkGet,
    Route::EquationGet,
    Route::ThemeGet,
];

impl Route {
    fn name(self) -> &'static str {
        match self {
            Self::ThemeSet => "pptx_theme_set_v1",
            Self::MasterName => "pptx_master_name_v1",
            Self::ColumnWidth => "pptx_column_width_v1",
            Self::ColumnAdd => "pptx_column_add_v1",
            Self::HyperlinkSet => "pptx_hyperlink_set_v1",
            Self::HyperlinkGet => "pptx_hyperlink_get_v1",
            Self::EquationGet => "pptx_equation_get_v1",
            Self::ThemeGet => "pptx_theme_get_v1",
        }
    }

    fn is_read(self) -> bool {
        matches!(
            self,
            Self::HyperlinkGet | Self::EquationGet | Self::ThemeGet
        )
    }
}

/// A deck holding one Office Math equation in a markup-compatibility
/// choice (LibreOffice `oox/qa/unit/data/import-mce.pptx`, MPL-2.0, tracked
/// in `test-data`).
const EQUATION_DECK: &[u8] =
    include_bytes!("../../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx");

/// A deck of `count` slides, each with a linked text box and a 3×3 table.
fn corpus(count: usize) -> Result<Vec<u8>> {
    let package = litchi_pptx::Package::new()?;
    let mut edit = package.edit()?;
    let cells: Vec<Vec<String>> = (0..3)
        .map(|row| (0..3).map(|col| format!("r{row}c{col}")).collect())
        .collect();
    for index in 0..count {
        let slide = edit.insert_slide(Position::new(index), "Blank")?;
        edit.add_text_box(slide.name(), "Docs", (0, 0, 1_828_800, 914_400))?;
        edit.add_table(slide.name(), &cells, (0, 1_828_800, 2_743_200, 914_400))?;
        let link = Link {
            target: Target::Url(format!("https://example.com/{index}")),
            tooltip: None,
        };
        edit.set_run_hyperlink(
            index,
            0_usize,
            Position::new(0),
            Position::new(0),
            Some(&link),
        )?;
    }
    let mut package = edit.commit()?.into_package();
    Ok(package.to_bytes()?)
}

fn native_write(
    base: &litchi_pptx::Package,
    route: Route,
    last: usize,
) -> Result<litchi_pptx::Package> {
    let mut edit = base.edit()?;
    match route {
        Route::ThemeSet => {
            let theme = edit.theme()?.ok_or("no theme")?;
            let colors = theme.colors.with(
                litchi_drawingml::theme::Slot::Accent1,
                litchi_drawingml::theme::Color::rgb("FF0000")?,
            );
            edit.set_theme_colors(&colors)?;
        },
        Route::MasterName => {
            edit.set_master_name(Position::new(0), "Brand")?;
        },
        Route::ColumnWidth => {
            edit.set_table_column_width(last, 1_usize, Position::new(1), 1_371_600)?;
        },
        Route::ColumnAdd => {
            edit.insert_table_column(last, 1_usize, Position::new(3), Some(914_400), "new")?;
        },
        Route::HyperlinkSet => {
            let link = Link {
                target: Target::Url("https://example.org/".into()),
                tooltip: None,
            };
            edit.set_shape_hyperlink(last, 0_usize, Some(&link))?;
        },
        _ => unreachable!("write routes"),
    }
    Ok(edit.commit()?.into_package())
}

fn unified_write(file: &File, route: Route, slide: usize) -> Result<File> {
    Ok(match route {
        Route::ThemeSet => file.set("/theme", props! { "accent1" => "FF0000" })?,
        Route::MasterName => file.set("/slidemaster[1]", props! { "name" => "Brand" })?,
        Route::ColumnWidth => file.set(
            format!("/slide[{slide}]/table[1]/col[2]").as_str(),
            props! { "width" => "1.5in" },
        )?,
        Route::ColumnAdd => {
            file.add(
                format!("/slide[{slide}]/table[1]").as_str(),
                NewNode::new("col").prop("width", "1in").prop("text", "new"),
            )?
            .file
        },
        Route::HyperlinkSet => {
            file.add(
                format!("/slide[{slide}]/shape[1]").as_str(),
                NewNode::new("hyperlink").prop("link", "https://example.org/"),
            )?
            .file
        },
        _ => unreachable!("write routes"),
    })
}

/// A native read; returns the value both legs compare.
fn native_read(base: &litchi_pptx::Package, route: Route, last: usize) -> Result<String> {
    let presentation = base.presentation()?;
    Ok(match route {
        Route::HyperlinkGet => {
            let slide = presentation.slide(last)?.ok_or("slide")?;
            let scene = slide.scene()?;
            let handle = scene.shape(0_usize)?.ok_or("shape")?;
            let body = handle.view().text_body()?.ok_or("body")?;
            let link = body.paragraphs[0].runs[0]
                .properties
                .hyperlink
                .as_ref()
                .ok_or("link")?;
            match slide
                .hyperlink(&Click::from_text(link))?
                .ok_or("resolved")?
                .target
            {
                Target::Url(url) => url,
                _ => return Err("not a URL".into()),
            }
        },
        Route::EquationGet => {
            let slide = presentation.slide(0)?.ok_or("slide")?;
            slide
                .equations()?
                .first()
                .ok_or("equation")?
                .formula()
                .to_owned()
        },
        Route::ThemeGet => {
            let theme = base.theme()?.ok_or("theme")?;
            theme.fonts.minor().latin.clone()
        },
        _ => unreachable!("read routes"),
    })
}

fn unified_read(file: &File, route: Route, slide: usize) -> Result<String> {
    let (path, key) = match route {
        Route::HyperlinkGet => (
            format!("/slide[{slide}]/shape[1]/paragraph[1]/run[1]/hyperlink"),
            "link",
        ),
        Route::EquationGet => ("/slide[1]/shape[1]/equation[1]".to_owned(), "formula"),
        Route::ThemeGet => ("/theme".to_owned(), "bodyFont"),
        _ => unreachable!("read routes"),
    };
    let node = file.get(path.as_str())?;
    let value = node
        .view()
        .format()?
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
        .ok_or("no value")?;
    Ok(value)
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

fn case(route: Route, count: usize, rounds: usize) -> Result<Value> {
    let bytes = if route == Route::EquationGet {
        EQUATION_DECK.to_vec()
    } else {
        corpus(count)?
    };
    let count = if route == Route::EquationGet {
        1
    } else {
        count
    };
    let native = litchi_pptx::Package::from_bytes(&bytes)?;
    let unified = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let last = count - 1;
    let equivalence = if route.is_read() {
        let left = native_read(&native, route, last)?;
        let right = unified_read(&unified, route, count)?;
        if left != right {
            return Err(format!("{}: {left:?} != {right:?}", route.name()).into());
        }
        "equal_value"
    } else {
        let mut left = native_write(&native, route, last)?;
        let right = unified_write(&unified, route, count)?;
        let left = left.to_bytes()?;
        if left == right.to_bytes()? {
            "exact_bytes"
        } else {
            "same_operation"
        }
    };
    let mut native_samples = Vec::with_capacity(rounds * 2);
    let mut unified_samples = Vec::with_capacity(rounds * 2);
    for round in 0..rounds * 2 {
        let native_first = (round / 2) % 2 == 0;
        for leg in 0..2 {
            let run_native = (leg == 0) == native_first;
            let start = Instant::now();
            if run_native {
                if route.is_read() {
                    black_box(native_read(&native, route, last)?);
                } else {
                    black_box(native_write(&native, route, last)?);
                }
                native_samples.push(start.elapsed().as_nanos());
            } else {
                if route.is_read() {
                    black_box(unified_read(&unified, route, count)?);
                } else {
                    black_box(unified_write(&unified, route, count)?);
                }
                unified_samples.push(start.elapsed().as_nanos());
            }
        }
    }
    let native_median = median(&mut native_samples.clone());
    let unified_median = median(&mut unified_samples.clone());
    Ok(json!({
        "route": route.name(),
        "owner": "eager",
        "slides": count,
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
    // Record 0326: `--only ROUTE` runs one route (for callgrind).
    let mut only: Option<String> = None;
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
            "--only" => only = Some(args.next().ok_or("--only needs a route")?),
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    if rounds < 3 || counts.contains(&0) {
        return Err("at least three rounds and one slide".into());
    }
    let mut results = Vec::new();
    for route in ROUTES {
        if only.as_deref().is_some_and(|only| only != route.name()) {
            continue;
        }
        for (at, &count) in counts.iter().enumerate() {
            // The equation route reads one fixed deck.
            if route == Route::EquationGet && at > 0 {
                continue;
            }
            results.push(case(route, count, rounds)?);
        }
    }
    let report = json!({
        "schema": 1,
        "harness": "unified-pptx-elements-baseline",
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
