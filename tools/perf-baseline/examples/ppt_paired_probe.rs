//! Paired native-versus-unified PPT write probe (record 0260).
//!
//! Every scenario builds its deterministic deck outside the timer and checks
//! that both legs publish identical bytes before timing. Samples alternate in
//! ABBA order (native, unified, unified, native) so slow drift affects both
//! legs equally. Each sample times one complete immutable write: the unified
//! leg includes path parsing, resolution, native staging, publication and
//! adoption; the native leg is the equivalent hand-written `Package::edit`.
//!
//! ```sh
//! cargo build --release --offline --manifest-path tools/perf-baseline/Cargo.toml \
//!   --example ppt_paired_probe
//! taskset -c 3 target/release/examples/ppt_paired_probe single 31 > single.json
//! taskset -c 3 target/release/examples/ppt_paired_probe batch 11 > batch.json
//! taskset -c 3 target/release/examples/ppt_paired_probe phases 5 > phases.json
//! ```

use std::error::Error;
use std::time::Instant;

use litchi_ops::transaction::SemanticLimits;
use litchi_ppt::Package;
use litchi_ppt::package::edit::{Position, Target};
use litchi_ppt::slide_order::{MAX_SHAPE_TEXT_REPLACEMENTS, ShapeTextReplacement};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

fn source_text(index: usize) -> String {
    format!("litchi-ppt-probe-source-{index:05}")
}

fn updated_text(index: usize) -> String {
    format!("litchi-ppt-probe-updated-{index:05}")
}

/// One slide with `count` text boxes, as the 0218–0221 text-batch corpus.
fn deck(count: usize) -> Result<Vec<u8>> {
    let mut builder = litchi_ppt::writer::Builder::new();
    let slide = builder.add_slide()?;
    for index in 0..count {
        let x = 24 + i32::try_from(index % 16)? * 120;
        let y = 24 + i32::try_from(index / 16)? * 72;
        builder.add_textbox(slide, x, y, 96, 48, &source_text(index))?;
    }
    Ok(builder.build()?.to_bytes()?)
}

fn median(values: &mut [f64]) -> f64 {
    values.sort_by(f64::total_cmp);
    let middle = values.len() / 2;
    if values.len().is_multiple_of(2) {
        (values[middle - 1] + values[middle]) / 2.0
    } else {
        values[middle]
    }
}

struct Paired {
    native_us: Vec<f64>,
    unified_us: Vec<f64>,
}

impl Paired {
    /// ABBA rounds: `rounds` native and `rounds` unified samples.
    fn run(
        rounds: usize,
        mut native: impl FnMut() -> Result<()>,
        mut unified: impl FnMut() -> Result<()>,
    ) -> Result<Self> {
        native()?;
        unified()?;
        let mut native_us = Vec::with_capacity(rounds);
        let mut unified_us = Vec::with_capacity(rounds);
        let time = |leg: &mut dyn FnMut() -> Result<()>| -> Result<f64> {
            let start = Instant::now();
            leg()?;
            Ok(start.elapsed().as_secs_f64() * 1e6)
        };
        for round in 0..rounds {
            if round % 2 == 0 {
                native_us.push(time(&mut native)?);
                unified_us.push(time(&mut unified)?);
            } else {
                unified_us.push(time(&mut unified)?);
                native_us.push(time(&mut native)?);
            }
        }
        Ok(Self {
            native_us,
            unified_us,
        })
    }

    fn json(mut self, scenario: &str, size: usize) -> String {
        let mut ratios = self
            .native_us
            .iter()
            .zip(&self.unified_us)
            .map(|(native, unified)| unified / native)
            .collect::<Vec<_>>();
        let raw_native = self.native_us.clone();
        let raw_unified = self.unified_us.clone();
        format!(
            "{{\"scenario\":\"{scenario}\",\"size\":{size},\"samples\":{},\"native_median_us\":{:.3},\"unified_median_us\":{:.3},\"pair_ratio_median\":{:.4},\"native_us\":{:?},\"unified_us\":{:?}}}",
            raw_native.len(),
            median(&mut self.native_us),
            median(&mut self.unified_us),
            median(&mut ratios),
            raw_native
                .iter()
                .map(|value| (value * 1000.0).round() / 1000.0)
                .collect::<Vec<_>>(),
            raw_unified
                .iter()
                .map(|value| (value * 1000.0).round() / 1000.0)
                .collect::<Vec<_>>(),
        )
    }
}

/// One text set on the last shape of a one-slide deck (record 0216).
fn single(rounds: usize) -> Result<Vec<String>> {
    let mut lines = Vec::new();
    for shapes in [1usize, 100, 1000] {
        let bytes = deck(shapes)?;
        let native = Package::from_bytes(&bytes)?;
        let unified = litchi::open_bytes(bytes.clone())?;
        let index = shapes - 1;
        let path =
            litchi_ops::path::Path::parse(&format!("/slide[1]/shape[{shapes}]"))?.into_owned();
        let text = updated_text(index);
        let native_leg = || -> Result<Package> {
            let mut edit = native.edit()?;
            edit.set_shape_text(
                Target::new(Position::new(0), Position::new(index)),
                text.as_str(),
            )?;
            Ok(edit.commit()?.into_package())
        };
        let unified_leg = || -> Result<litchi::File> {
            Ok(unified.set(&path, litchi_ops::props! { "text" => text.as_str() })?)
        };
        if native_leg()?.to_bytes()? != unified_leg()?.to_bytes()? {
            return Err("single-set legs publish different bytes".into());
        }
        let paired = Paired::run(
            rounds,
            || native_leg().map(drop),
            || unified_leg().map(drop),
        )?;
        lines.push(paired.json("single_set_eager", shapes));
    }
    Ok(lines)
}

/// One text set on a freshly opened deck: every sample opens new native and
/// unified owners outside the timer, so no parse is reused across samples.
fn single_cold(rounds: usize) -> Result<Vec<String>> {
    let mut lines = Vec::new();
    for shapes in [1usize, 100, 1000] {
        let bytes = deck(shapes)?;
        let index = shapes - 1;
        let path =
            litchi_ops::path::Path::parse(&format!("/slide[1]/shape[{shapes}]"))?.into_owned();
        let text = updated_text(index);
        let target = Target::new(Position::new(0), Position::new(index));
        let native_leg = |native: &Package| -> Result<Package> {
            let mut edit = native.edit()?;
            edit.set_shape_text(target, text.as_str())?;
            Ok(edit.commit()?.into_package())
        };
        let unified_leg = |unified: &litchi::File| -> Result<litchi::File> {
            Ok(unified.set(&path, litchi_ops::props! { "text" => text.as_str() })?)
        };
        if native_leg(&Package::from_bytes(&bytes)?)?.to_bytes()?
            != unified_leg(&litchi::open_bytes(bytes.clone())?)?.to_bytes()?
        {
            return Err("cold single-set legs publish different bytes".into());
        }
        let mut native_us = Vec::with_capacity(rounds);
        let mut unified_us = Vec::with_capacity(rounds);
        for round in 0..rounds {
            let native = Package::from_bytes(&bytes)?;
            let unified = litchi::open_bytes(bytes.clone())?;
            let time_native = || -> Result<f64> {
                let start = Instant::now();
                drop(native_leg(&native)?);
                Ok(start.elapsed().as_secs_f64() * 1e6)
            };
            let time_unified = || -> Result<f64> {
                let start = Instant::now();
                drop(unified_leg(&unified)?);
                Ok(start.elapsed().as_secs_f64() * 1e6)
            };
            // Each owner is timed once (cold); ABBA order alternates per round.
            if round % 2 == 0 {
                native_us.push(time_native()?);
                unified_us.push(time_unified()?);
            } else {
                unified_us.push(time_unified()?);
                native_us.push(time_native()?);
            }
        }
        lines.push(
            Paired {
                native_us,
                unified_us,
            }
            .json("single_set_eager_cold", shapes),
        );
    }
    Ok(lines)
}

fn fixture(relative: &str) -> Result<Vec<u8>> {
    Ok(std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative),
    )?)
}

/// Cold paired timing of one operation: fresh owners per sample, ABBA order.
fn cold_pair(
    scenario: &str,
    bytes: &[u8],
    rounds: usize,
    native_leg: impl Fn(&Package) -> Result<Package>,
    unified_leg: impl Fn(&litchi::File) -> Result<litchi::File>,
) -> Result<String> {
    if native_leg(&Package::from_bytes(bytes)?)?.to_bytes()?
        != unified_leg(&litchi::open_bytes(bytes.to_vec())?)?.to_bytes()?
    {
        return Err(format!("{scenario}: legs publish different bytes").into());
    }
    let mut native_us = Vec::with_capacity(rounds);
    let mut unified_us = Vec::with_capacity(rounds);
    for round in 0..rounds {
        let native = Package::from_bytes(bytes)?;
        let unified = litchi::open_bytes(bytes.to_vec())?;
        let time_native = || -> Result<f64> {
            let start = Instant::now();
            drop(native_leg(&native)?);
            Ok(start.elapsed().as_secs_f64() * 1e6)
        };
        let time_unified = || -> Result<f64> {
            let start = Instant::now();
            drop(unified_leg(&unified)?);
            Ok(start.elapsed().as_secs_f64() * 1e6)
        };
        if round % 2 == 0 {
            native_us.push(time_native()?);
            unified_us.push(time_unified()?);
        } else {
            unified_us.push(time_unified()?);
            native_us.push(time_native()?);
        }
    }
    Ok(Paired {
        native_us,
        unified_us,
    }
    .json(scenario, bytes.len()))
}

/// The record 0260 write paths on PowerPoint-authored fixtures.
fn paths(rounds: usize) -> Result<Vec<String>> {
    let sample = fixture("ole/ppt/SampleShow.ppt")?;
    let notes = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt")?;
    let long = fixture("poi/test-data/slideshow/41246-1.ppt")?;
    let title = Target::new(Position::new(0), Position::new(0));
    let mut lines = Vec::new();
    lines.push(cold_pair(
        "outline_title_set_sampleshow",
        &sample,
        rounds,
        |native| {
            let mut edit = native.edit()?;
            edit.set_shape_text(title, "A longer replacement title")?;
            Ok(edit.commit()?.into_package())
        },
        |unified| {
            Ok(unified.set(
                "/slide[1]/shape[1]",
                litchi_ops::props! { "text" => "A longer replacement title" },
            )?)
        },
    )?);
    lines.push(cold_pair(
        "outline_body_set_41246",
        &long,
        rounds,
        |native| {
            let mut edit = native.edit()?;
            edit.set_shape_text(
                Target::new(Position::new(7), Position::new(1)),
                "First paragraph\rSecond paragraph",
            )?;
            Ok(edit.commit()?.into_package())
        },
        |unified| {
            Ok(unified.set(
                "/slide[8]/shape[2]",
                litchi_ops::props! { "text" => "First paragraph\rSecond paragraph" },
            )?)
        },
    )?);
    lines.push(cold_pair(
        "paragraph_set_basic",
        &notes,
        rounds,
        |native| {
            let mut edit = native.edit()?;
            edit.set_shape_text(
                Target::new(Position::new(0), Position::new(1)),
                "This is a test subtitle\rRewritten line",
            )?;
            Ok(edit.commit()?.into_package())
        },
        |unified| {
            Ok(unified.set(
                "/slide[1]/shape[2]/paragraph[2]",
                litchi_ops::props! { "text" => "Rewritten line" },
            )?)
        },
    )?);
    lines.push(cold_pair(
        "notes_set_basic",
        &notes,
        rounds,
        |native| {
            let mut edit = native.edit()?;
            edit.set_notes_text(Position::new(0), "Rewritten speaker notes")?;
            Ok(edit.commit()?.into_package())
        },
        |unified| {
            Ok(unified.set(
                "/slide[1]/notes",
                litchi_ops::props! { "text" => "Rewritten speaker notes" },
            )?)
        },
    )?);
    for (scenario, bytes) in [
        ("slide_add_sampleshow", &sample),
        ("slide_add_41246", &long),
    ] {
        lines.push(cold_pair(
            scenario,
            bytes,
            rounds,
            |native| {
                let count = native.presentation()?.slide_count();
                let mut edit = native.edit()?;
                edit.insert_slide(Position::new(count))?;
                Ok(edit.commit()?.into_package())
            },
            |unified| Ok(unified.add("/", litchi_ops::NewNode::new("slide"))?.file),
        )?);
    }
    Ok(lines)
}

/// `count` distinct direct text sets on one slide in one atomic edit
/// (records 0218–0221).
fn batch(rounds: usize) -> Result<Vec<String>> {
    let mut lines = Vec::new();
    for count in [1000usize, 4096, 10_000] {
        let bytes = deck(count)?;
        let native = Package::from_bytes(&bytes)?;
        let unified = litchi::open_bytes(bytes.clone())?;
        let replacements = (0..count)
            .map(|index| {
                ShapeTextReplacement::new(
                    Target::new(Position::new(0), Position::new(index)),
                    updated_text(index),
                )
            })
            .collect::<Vec<_>>();
        let paths = (0..count)
            .map(|index| {
                litchi_ops::path::Path::parse(&format!("/slide[1]/shape[{}]", index + 1))
                    .map(litchi_ops::path::Path::into_owned)
            })
            .collect::<std::result::Result<Vec<_>, _>>()?;
        let texts = (0..count).map(updated_text).collect::<Vec<_>>();
        let limits = SemanticLimits::builder()
            .max_snapshot_bytes(1 << 30)?
            .max_staged_bytes(1 << 30)?
            .max_operations(1 << 20)?
            .max_expanded_targets(1 << 20)?
            .build()?;
        let native_leg = || -> Result<Package> {
            let mut package = native.clone();
            for operations in replacements.chunks(litchi_ppt::package::edit::MAX_OPERATIONS) {
                let chunks = operations
                    .chunks(MAX_SHAPE_TEXT_REPLACEMENTS)
                    .collect::<Vec<_>>();
                let mut edit = package.edit()?;
                edit.set_shape_text_chunks(&chunks)?;
                package = edit.commit()?.into_package();
            }
            Ok(package)
        };
        let unified_leg = || -> Result<litchi::File> {
            let commit = unified.edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| {
                    for (path, text) in paths.iter().zip(&texts) {
                        edit.set(path, litchi_ops::props! { "text" => text.as_str() })?;
                    }
                    Ok(())
                },
            )?;
            Ok(commit.into_file())
        };
        if native_leg()?.to_bytes()? != unified_leg()?.to_bytes()? {
            return Err("batch legs publish different bytes".into());
        }
        let paired = Paired::run(
            rounds,
            || native_leg().map(drop),
            || unified_leg().map(drop),
        )?;
        lines.push(paired.json("text_batch_eager", count));
    }
    Ok(lines)
}

/// Native phase times of the batch scenario, for attribution.
fn phases(rounds: usize) -> Result<Vec<String>> {
    let mut lines = Vec::new();
    for count in [1000usize, 4096, 10_000] {
        let bytes = deck(count)?;
        let native = Package::from_bytes(&bytes)?;
        let replacements = (0..count)
            .map(|index| {
                ShapeTextReplacement::new(
                    Target::new(Position::new(0), Position::new(index)),
                    updated_text(index),
                )
            })
            .collect::<Vec<_>>();
        let mut open = Vec::new();
        let mut stage = Vec::new();
        let mut commit = Vec::new();
        for _ in 0..rounds {
            let mut package = native.clone();
            let (mut o, mut s, mut c) = (0.0, 0.0, 0.0);
            for operations in replacements.chunks(litchi_ppt::package::edit::MAX_OPERATIONS) {
                let chunks = operations
                    .chunks(MAX_SHAPE_TEXT_REPLACEMENTS)
                    .collect::<Vec<_>>();
                let start = Instant::now();
                let mut edit = package.edit()?;
                o += start.elapsed().as_secs_f64() * 1e6;
                let start = Instant::now();
                edit.set_shape_text_chunks(&chunks)?;
                s += start.elapsed().as_secs_f64() * 1e6;
                let start = Instant::now();
                package = edit.commit()?.into_package();
                c += start.elapsed().as_secs_f64() * 1e6;
            }
            open.push(o);
            stage.push(s);
            commit.push(c);
        }
        lines.push(format!(
            "{{\"scenario\":\"native_phases\",\"size\":{count},\"edit_open_median_us\":{:.1},\"stage_median_us\":{:.1},\"commit_median_us\":{:.1}}}",
            median(&mut open),
            median(&mut stage),
            median(&mut commit)
        ));
    }
    Ok(lines)
}

fn main() -> Result<()> {
    let mut args = std::env::args().skip(1);
    let mode = args.next().unwrap_or_else(|| "single".to_owned());
    let rounds = args.next().map_or(Ok(11), |value| value.parse())?;
    let lines = match mode.as_str() {
        "single" => single(rounds)?,
        "single-cold" => single_cold(rounds)?,
        "paths" => paths(rounds)?,
        "batch" => batch(rounds)?,
        "phases" => phases(rounds)?,
        "profile-batch" => {
            // A long native-only loop for sampling profilers.
            let bytes = deck(10_000)?;
            let native = Package::from_bytes(&bytes)?;
            let replacements = (0..10_000)
                .map(|index| {
                    ShapeTextReplacement::new(
                        Target::new(Position::new(0), Position::new(index)),
                        updated_text(index),
                    )
                })
                .collect::<Vec<_>>();
            for _ in 0..rounds {
                let mut package = native.clone();
                for operations in replacements.chunks(litchi_ppt::package::edit::MAX_OPERATIONS) {
                    let chunks = operations
                        .chunks(MAX_SHAPE_TEXT_REPLACEMENTS)
                        .collect::<Vec<_>>();
                    let mut edit = package.edit()?;
                    edit.set_shape_text_chunks(&chunks)?;
                    package = edit.commit()?.into_package();
                }
            }
            Vec::new()
        },
        "profile-batch-unified" => {
            // A long unified-only 10k loop for sampling profilers (record 0264).
            let bytes = deck(10_000)?;
            let unified = litchi::open_bytes(bytes)?;
            let paths = (0..10_000)
                .map(|index| {
                    litchi_ops::path::Path::parse(&format!("/slide[1]/shape[{}]", index + 1))
                        .map(litchi_ops::path::Path::into_owned)
                })
                .collect::<std::result::Result<Vec<_>, _>>()?;
            let texts = (0..10_000).map(updated_text).collect::<Vec<_>>();
            let limits = SemanticLimits::builder()
                .max_snapshot_bytes(1 << 30)?
                .max_staged_bytes(1 << 30)?
                .max_operations(1 << 20)?
                .max_expanded_targets(1 << 20)?
                .build()?;
            for _ in 0..rounds {
                let commit = unified.edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                    |edit| {
                        for (path, text) in paths.iter().zip(&texts) {
                            edit.set(path, litchi_ops::props! { "text" => text.as_str() })?;
                        }
                        Ok(())
                    },
                )?;
                drop(commit);
            }
            Vec::new()
        },
        "profile-single" => {
            let bytes = deck(1000)?;
            let unified = litchi::open_bytes(bytes)?;
            let path = litchi_ops::path::Path::parse("/slide[1]/shape[1000]")?.into_owned();
            for _ in 0..rounds {
                drop(unified.set(&path, litchi_ops::props! { "text" => "changed" })?);
            }
            Vec::new()
        },
        "profile-single-native" => {
            let bytes = deck(1000)?;
            let native = Package::from_bytes(&bytes)?;
            for _ in 0..rounds {
                let mut edit = native.edit()?;
                edit.set_shape_text(Target::new(Position::new(0), Position::new(999)), "changed")?;
                drop(edit.commit()?.into_package());
            }
            Vec::new()
        },
        other => return Err(format!("unknown mode {other}").into()),
    };
    println!("[\n{}\n]", lines.join(",\n"));
    Ok(())
}
