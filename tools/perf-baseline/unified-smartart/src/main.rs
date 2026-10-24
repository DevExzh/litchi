//! Paired SmartArt writes (record 0284): the native leaf calls versus the
//! unified verbs, on eager owners that already hold `N` SmartArt graphics.
//!
//! - `pptx_smartart_add_v1`: native `opened::Edit::add_smartart` with a
//!   `litchi_drawingml::diagram::authoring::Spec` plus one commit, versus
//!   `File::add("/slide[1]", smartart{layout, text, ...})`.
//! - `pptx_smartart_node_text_v1`: the hand-written leaf edit of the first
//!   graphic's second node — `Edit::smartart`, `Spec::from_parts`, the text
//!   change, `Spec::{data_xml, drawing_xml}`, `Edit::set_smartart` and one
//!   commit — versus `File::set("/slide[1]/smartart[1]/node[2]", {text})`.
//! - `pptx_smartart_node_add_v1`: the same leaf sequence appending a node,
//!   versus `File::add("/slide[1]/smartart[1]", node{text})`.
//!
//! Corpus generation, opening, warm-up, serialization and the equivalence
//! check (exact bytes, else equal decoded archive members: the eager
//! owner's writer may order members differently) are outside the timed interval; each sample is one operation. Legs
//! alternate ABBA; medians are the upper middle.

use litchi_drawingml::diagram::authoring::{Layout, Node, Spec};
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};
use litchi_pptx::opened::SmartArtParts;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 3] = [0, 10, 50];
const DEFAULT_ROUNDS: usize = 9;
const BOUNDS: (i64, i64, i64, i64) = (914_400, 914_400, 5_486_400, 3_657_600);
const OUTLINE: &str = "Plan\n\tScope\n\tBudget\nBuild\nTest\nShip";

#[derive(Clone, Copy, PartialEq, Eq)]
enum Route {
    Add,
    NodeText,
    NodeAdd,
}

impl Route {
    fn name(self) -> &'static str {
        match self {
            Self::Add => "pptx_smartart_add_v1",
            Self::NodeText => "pptx_smartart_node_text_v1",
            Self::NodeAdd => "pptx_smartart_node_add_v1",
        }
    }
}

fn spec() -> Result<Spec> {
    Ok(Spec::parse_outline(
        Layout::Process,
        OUTLINE,
        &litchi_drawingml::diagram::Limits::default(),
    )?)
}

/// A deck with one blank slide holding `count` SmartArt graphics (at least
/// one for the node routes).
fn corpus(count: usize) -> Result<Vec<u8>> {
    let package = litchi_pptx::Package::new()?;
    let mut edit = package.edit()?;
    edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    for index in 0..count {
        edit.add_smartart(0_usize, &spec()?, &format!("Diagram {}", index + 1), BOUNDS)?;
    }
    let package = edit.commit()?.into_package();
    Ok(package.to_bytes()?)
}

/// The hand-written leaf change of the first graphic's parts.
fn leaf_change(parts: &SmartArtParts, route: Route) -> Result<SmartArtParts> {
    let limits = litchi_drawingml::diagram::Limits::default();
    let mut spec = Spec::from_parts(&parts.data, &parts.layout, &limits)?.ok_or("not authored")?;
    match route {
        Route::NodeText => spec.nodes_mut()[1].text = "Make".to_owned(),
        _ => spec.nodes_mut().push(Node::new("Learn")),
    }
    let outline = litchi_drawingml::diagram::edit::Outline::read(
        &parts.data,
        &litchi_drawingml::diagram::Limits::default(),
    )?;
    Ok(SmartArtParts {
        data: spec
            .data_xml(outline.drawing_relationship.as_deref(), &limits)?
            .into_bytes(),
        layout: parts.layout.clone(),
        quick_style: parts.quick_style.clone(),
        colors: parts.colors.clone(),
        drawing: Some(spec.drawing_xml(BOUNDS.2, BOUNDS.3, &limits)?.into_bytes()),
    })
}

fn native_once(
    base: &litchi_pptx::Package,
    route: Route,
    count: usize,
) -> Result<litchi_pptx::Package> {
    let mut edit = base.edit()?;
    match route {
        Route::Add => {
            edit.add_smartart(0_usize, &spec()?, &format!("Diagram {}", count + 1), BOUNDS)?;
        },
        _ => {
            let parts = edit.smartart(0_usize, "Diagram 1")?;
            let next = leaf_change(&parts, route)?;
            edit.set_smartart(0_usize, "Diagram 1", next)?;
        },
    }
    Ok(edit.commit()?.into_package())
}

fn unified_once(file: &File, route: Route) -> Result<File> {
    Ok(match route {
        Route::Add => {
            file.add(
                "/slide[1]",
                NewNode::new("smartart")
                    .prop("layout", "process")
                    .prop("text", OUTLINE)
                    .prop("x", "1in")
                    .prop("y", "1in")
                    .prop("width", "6in")
                    .prop("height", "4in"),
            )?
            .file
        },
        Route::NodeText => {
            file.set("/slide[1]/smartart[1]/node[2]", props! { "text" => "Make" })?
        },
        Route::NodeAdd => {
            file.add(
                "/slide[1]/smartart[1]",
                NewNode::new("node").prop("text", "Learn"),
            )?
            .file
        },
    })
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

fn case(route: Route, count: usize, rounds: usize) -> Result<Value> {
    let count = if route == Route::Add {
        count
    } else {
        count.max(1)
    };
    let bytes = corpus(count)?;
    let native = litchi_pptx::Package::from_bytes(&bytes)?;
    let unified = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    // Warm-up and equivalence: both legs publish the same package.
    let native_out = native_once(&native, route, count)?.to_bytes()?;
    let unified_out = unified_once(&unified, route)?.to_bytes()?;
    let equivalence = if native_out == unified_out {
        "exact_bytes"
    } else if members(&native_out)? == members(&unified_out)? {
        "decoded_members"
    } else {
        return Err(format!("{}: the legs publish different packages", route.name()).into());
    };
    let mut native_samples = Vec::with_capacity(rounds * 2);
    let mut unified_samples = Vec::with_capacity(rounds * 2);
    for round in 0..rounds * 2 {
        let native_first = (round / 2) % 2 == 0;
        for leg in 0..2 {
            let run_native = (leg == 0) == native_first;
            let start = Instant::now();
            if run_native {
                black_box(native_once(&native, route, count)?);
                native_samples.push(start.elapsed().as_nanos());
            } else {
                black_box(unified_once(&unified, route)?);
                unified_samples.push(start.elapsed().as_nanos());
            }
        }
    }
    let native_median = median(&mut native_samples.clone());
    let unified_median = median(&mut unified_samples.clone());
    Ok(json!({
        "route": route.name(),
        "owner": "eager",
        "existing_smartart": count,
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
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    if rounds < 3 {
        return Err("at least three rounds".into());
    }
    let mut results = Vec::new();
    for route in [Route::Add, Route::NodeText, Route::NodeAdd] {
        for &count in &counts {
            results.push(case(route, count, rounds)?);
        }
    }
    let report = json!({
        "schema": 1,
        "harness": "unified-smartart-baseline",
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
