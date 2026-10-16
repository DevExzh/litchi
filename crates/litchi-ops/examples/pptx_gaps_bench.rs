//! Unified-versus-native and before/after timing for record 0303.
//!
//! Paired scenarios (native source-backed `litchi_pptx::source::Package`
//! edit and commit, against the unified `File` call that wraps it):
//!
//! - `pptx_slide_remove_owned`: remove slide 1 of POI `45545_Comment.pptx`,
//!   which owns a notes slide, comments, a VML drawing and an OLE object;
//! - `pptx_ole_remove`: remove the OLE frame (markup compatibility enclosed
//!   in its `p:graphicFrame`) of LibreOffice `ole.pptx`;
//! - `pptx_section_move`: move slide 1 after slide 2 of LibreOffice
//!   `slide-sections.pptx`, which keeps its p14 sections;
//! - `pptx_slide_remove_plain`, `pptx_slide_move_plain` and
//!   `pptx_shape_text_set`: a dependency-free slide removal, a slide move and
//!   a shape-text set on `shapes.pptx`, paths the record changed (the
//!   owned-Part fast path, the section check, the snapshot admission); these
//!   also run on a build of the base tree for a before/after comparison.
//!
//! Each scenario prints one JSON line with the median and minimum of
//! `ROUNDS` samples (after `WARMUP`), each sample the mean of `REPEATS`
//! operations; paired legs alternate ABBA. A scenario the build refuses
//! prints a `skipped` line. Build one binary per tree and run them
//! alternately, pinned:
//! `cargo build --release -p litchi-ops --features pptx --example
//! pptx_gaps_bench`, then `taskset -c 7 <binary>`.
//! `PPTX_GAPS_BENCH_SCENARIO=<name>` runs one scenario (for callgrind).

#[cfg(feature = "pptx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::path::Path;
    use std::sync::Arc;
    use std::time::Instant;

    use litchi_core::{OwnedSource, Position};
    use litchi_ops::{Placement, open_bytes};
    use litchi_pptx::source::Package;

    const ROUNDS: usize = 21;
    const WARMUP: usize = 3;

    let only = std::env::var("PPTX_GAPS_BENCH_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);

    fn sample(repeats: usize, mut run: impl FnMut()) -> f64 {
        let started = Instant::now();
        for _ in 0..repeats {
            run();
        }
        started.elapsed().as_secs_f64() * 1e6 / repeats as f64
    }

    fn summary(mut samples: Vec<f64>) -> (f64, f64) {
        samples.sort_by(f64::total_cmp);
        (samples[samples.len() / 2], samples[0])
    }

    fn paired(name: &str, repeats: usize, mut native: impl FnMut(), mut unified: impl FnMut()) {
        let mut native_samples = Vec::new();
        let mut unified_samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let (n, u) = if round % 2 == 0 {
                let n = sample(repeats, &mut native);
                (n, sample(repeats, &mut unified))
            } else {
                let u = sample(repeats, &mut unified);
                (sample(repeats, &mut native), u)
            };
            if round >= WARMUP {
                native_samples.push(n);
                unified_samples.push(u);
            }
        }
        let (native_median, native_min) = summary(native_samples);
        let (unified_median, unified_min) = summary(unified_samples);
        println!(
            "{}",
            serde_json::json!({
                "scenario": name,
                "native_median_us": native_median, "native_min_us": native_min,
                "unified_median_us": unified_median, "unified_min_us": unified_min,
                "ratio": unified_median / native_median,
                "rounds": ROUNDS, "repeats": repeats, "order": "ABBA"
            })
        );
    }

    fn skipped(name: &str, error: &dyn std::fmt::Display) {
        println!(
            "{}",
            serde_json::json!({ "scenario": name, "skipped": error.to_string() })
        );
    }

    let read = |relative: &str| -> std::io::Result<Vec<u8>> {
        std::fs::read(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(relative),
        )
    };
    let native = |bytes: &[u8]| -> litchi_pptx::Result<Package> {
        Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
    };

    // One scenario: a native edit and the unified call, checked to publish
    // the same bytes before timing.
    macro_rules! scenario {
        ($name:expr, $repeats:expr, $bytes:expr, |$edit:ident| $native_edit:expr, |$file:ident| $unified:expr) => {{
            let bytes = $bytes;
            let package = native(&bytes)?;
            let unified_file = open_bytes(bytes.clone())?;
            let native_once = || -> litchi_pptx::Result<Package> {
                let mut $edit = package.edit()?;
                $native_edit?;
                Ok($edit.commit()?.into_package())
            };
            let unified_once = || {
                let $file = &unified_file;
                $unified
            };
            match (native_once(), unified_once()) {
                (Ok(native_result), Ok(unified_result)) => {
                    assert_eq!(
                        native_result.to_bytes()?,
                        unified_result.to_bytes()?,
                        "both legs publish the same deck"
                    );
                    paired(
                        $name,
                        $repeats,
                        || {
                            black_box(native_once().expect("native"));
                        },
                        || {
                            black_box(unified_once().expect("unified"));
                        },
                    );
                },
                (Err(error), _) => skipped($name, &error),
                (_, Err(error)) => skipped($name, &error),
            }
        }};
    }

    if wanted("pptx_slide_remove_owned") {
        scenario!(
            "pptx_slide_remove_owned",
            5,
            read("poi/test-data/slideshow/45545_Comment.pptx")?,
            |edit| edit.remove_slide(Position::new(0)),
            |file| file.remove("/slide[1]")
        );
    }
    if wanted("pptx_ole_remove") {
        let bytes = read("libreoffice-core/sd/qa/unit/data/pptx/ole.pptx")?;
        let path = open_bytes(bytes.clone())?
            .query("ole")?
            .next()
            .ok_or("no OLE object")??
            .path()
            .to_string();
        scenario!(
            "pptx_ole_remove",
            20,
            bytes,
            |edit| edit.remove_shape(0, 0_usize),
            |file| file.remove(path.as_str())
        );
    }
    if wanted("pptx_section_move") {
        scenario!(
            "pptx_section_move",
            20,
            read("libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx")?,
            |edit| edit.move_slide(Position::new(0), Position::new(1)),
            |file| file.move_to("/slide[1]", Placement::after("/slide[2]")?)
        );
    }
    let shapes = read("ooxml/pptx/shapes.pptx")?;
    if wanted("pptx_slide_remove_plain") {
        scenario!(
            "pptx_slide_remove_plain",
            20,
            shapes.clone(),
            |edit| edit.remove_slide(Position::new(5)),
            |file| file.remove("/slide[6]")
        );
    }
    if wanted("pptx_slide_move_plain") {
        scenario!(
            "pptx_slide_move_plain",
            50,
            shapes.clone(),
            |edit| edit.move_slide(Position::new(0), Position::new(1)),
            |file| file.move_to("/slide[1]", Placement::after("/slide[2]")?)
        );
    }
    if wanted("pptx_shape_text_set") {
        scenario!(
            "pptx_shape_text_set",
            50,
            shapes.clone(),
            |edit| edit.set_shape_text(1, 0_usize, "Benchmark title"),
            |file| file.set(
                "/slide[2]/shape[1]",
                litchi_ops::props! { "text" => "Benchmark title" }
            )
        );
    }
    Ok(())
}

#[cfg(not(feature = "pptx"))]
fn main() {
    eprintln!("build with --features pptx");
}
