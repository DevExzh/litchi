//! Before/after timing for the hot paths record 0293 touches.
//!
//! - `xls_complete_open`: the leaf's complete XLS reader
//!   (`litchi_xls::raw::Reader::new`) over a 20,000-cell workbook and over
//!   the corpus `FormulaEvalTestData.xls`; its worksheet loop now skips
//!   embedded substreams;
//! - `xls_row_insert`: insert one row at the top of a 10,000-cell Litchi
//!   workbook, natively (`Package::edit` + `insert_rows` + `commit`) and
//!   through `File::add`; the shift now closes defined names, chart links and
//!   3-D formulas and keeps producer row extents;
//! - `rtf_move`: move the first of 1,000 plain paragraphs to the end,
//!   natively and through `File::move_to`; the commit now falls back to the
//!   producer splice only when the plain-body paths refuse.
//!
//! Every scenario prints one JSON line with the median and minimum of
//! `ROUNDS` samples (after `WARMUP`), each sample the mean of `REPEATS`
//! operations. Paired legs alternate ABBA. Build one binary per tree and run
//! them alternately, pinned:
//! `cargo build --release -p litchi-ops --features xls,rtf --example
//! producer_structure_bench`, then `taskset -c 7 <binary>`.
//! `PRODUCER_BENCH_SCENARIO=<name>` runs one scenario (for callgrind).

#[cfg(all(feature = "xls", feature = "rtf"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode};
    use litchi_xls::cell_values::Selector;
    use litchi_xls::writer::Builder;

    const ROUNDS: usize = 15;
    const WARMUP: usize = 3;

    let only = std::env::var("PRODUCER_BENCH_SCENARIO").ok();
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

    fn single(name: &str, repeats: usize, mut run: impl FnMut()) {
        let mut samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let value = sample(repeats, &mut run);
            if round >= WARMUP {
                samples.push(value);
            }
        }
        let (median, min) = summary(samples);
        println!(
            "{}",
            serde_json::json!({ "scenario": name, "median_us": median, "min_us": min, "rounds": ROUNDS, "repeats": repeats })
        );
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

    let workbook = |cells: usize| -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut builder = Builder::new();
        let data = builder.add_worksheet("Data")?;
        builder.add_worksheet("Archive")?;
        for row in 0..cells {
            builder.write_number(data, row as u32, 0, (row + 1) as f64)?;
        }
        Ok(builder.build()?.to_bytes()?)
    };

    if wanted("xls_complete_open") {
        let generated = workbook(20_000)?;
        single("xls_complete_open_20k", 20, || {
            black_box(
                litchi_xls::raw::Reader::new(std::io::Cursor::new(generated.clone()))
                    .expect("reader"),
            );
        });
        let corpus = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/xls/FormulaEvalTestData.xls"),
        )?;
        single("xls_complete_open_formula_eval", 20, || {
            black_box(
                litchi_xls::raw::Reader::new(std::io::Cursor::new(corpus.clone())).expect("reader"),
            );
        });
    }

    if wanted("xls_row_insert") {
        let bytes = workbook(10_000)?;
        let native = litchi_xls::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )?;
        let native_once = || {
            let mut edit = native.edit().expect("edit");
            edit.insert_rows(Selector::Position(Position::new(0)), 0, 1)
                .expect("insert");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            unified
                .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
                .expect("add")
                .file
        };
        assert_eq!(
            native_once().to_bytes()?,
            unified_once().to_bytes()?,
            "both legs publish the same workbook"
        );
        paired(
            "xls_row_insert_10k",
            5,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }

    if wanted("rtf_move") {
        let texts: Vec<_> = (0..1_000).map(|i| format!("paragraph {i:05}")).collect();
        let bytes = format!(r"{{\rtf1\ansi {}}}", texts.join(r"\par ")).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes(bytes.as_slice())?;
        let native_once = || {
            let mut edit = native.edit();
            edit.move_paragraph(Position::new(0), Position::new(999))
                .expect("move");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || unified.move_to("/body/p[1]", Placement::End).expect("move");
        assert_eq!(native_once().to_bytes()?, unified_once().to_bytes()?);
        paired(
            "rtf_move_1k",
            10,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "xls", feature = "rtf")))]
fn main() {
    eprintln!("build with --features xls,rtf");
}
