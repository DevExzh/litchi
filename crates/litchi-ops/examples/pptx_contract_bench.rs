//! Before/after timing of the PPTX routes record 0356 re-measures (record
//! 0351 §5, wave W2): open, a single `set`, a 1k batch, `to_bytes` and
//! `to_bytes` twice on one package, on the leaf `litchi_pptx::Package` and
//! the unified `File`.
//!
//! The source uses only calls that compile on both sides of the record 0356
//! API change, so one copy of it builds on the base tree and on the branch;
//! run the two binaries alternately (A, B, B, A), pinned. Inputs:
//! `45545_Comment.pptx` (Apache POI, 308 KB: notes, comments, VML, OLE) and
//! generated one-slide decks of 20 and 1,000 text boxes.
//!
//! Each scenario prints one JSON line: the median and minimum of `ROUNDS`
//! samples (after `WARMUP`) in µs, each sample one call.
//! `CONTRACT_SCENARIO=<prefix>` runs the scenarios whose name starts with it.
//! `cargo build --release -p litchi-ops --features pptx --example
//! pptx_contract_bench`, then `taskset -c 7 <binary>`.

#[cfg(feature = "pptx")]
#[allow(
    unused_mut,
    reason = "the base tree's outputs take &mut self; the branch's take &self"
)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::{BatchItem, BatchMode, OpenOptions, Props, SourceMode, props};

    const WARMUP: usize = 3;
    const ROUNDS: usize = 31;
    const COMMENTS: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/slideshow/45545_Comment.pptx");

    let only = std::env::var("CONTRACT_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| name.starts_with(only));

    fn measure(name: &str, rounds: usize, mut run: impl FnMut()) {
        let mut samples = Vec::with_capacity(rounds);
        for round in 0..WARMUP + rounds {
            let started = Instant::now();
            run();
            let elapsed = started.elapsed().as_secs_f64() * 1e6;
            if round >= WARMUP {
                samples.push(elapsed);
            }
        }
        samples.sort_by(f64::total_cmp);
        println!(
            "{}",
            serde_json::json!({
                "scenario": name,
                "median_us": samples[samples.len() / 2],
                "min_us": samples[0],
            })
        );
    }

    fn deck(shapes: usize) -> Result<Vec<u8>, litchi_pptx::Error> {
        let mut builder = litchi_pptx::builder::Builder::new()?;
        let slide = builder.presentation()?.add_slide()?;
        for index in 0..shapes {
            slide.add_text_box(
                &format!("text-{index}"),
                i64::try_from(index).unwrap_or(0) * 100,
                0,
                914_400,
                914_400,
            );
        }
        builder.build()?.to_bytes()
    }

    fn eager() -> OpenOptions {
        OpenOptions::new().source_mode(SourceMode::Eager)
    }

    let small = deck(20)?;
    let large = deck(1_000)?;
    for (label, bytes) in [("comments", COMMENTS.to_vec()), ("deck20", small.clone())] {
        let open_leaf = format!("open/leaf/{label}");
        if wanted(&open_leaf) {
            measure(&open_leaf, ROUNDS, || {
                drop(black_box(
                    litchi_pptx::Package::from_bytes(&bytes).expect("open"),
                ));
            });
        }
        let open_unified = format!("open/unified/{label}");
        if wanted(&open_unified) {
            measure(&open_unified, ROUNDS, || {
                drop(black_box(
                    litchi_ops::open_bytes_with(bytes.clone(), &eager()).expect("open"),
                ));
            });
        }
        let mut package = litchi_pptx::Package::from_bytes(&bytes)?;
        let save_leaf = format!("save/leaf/{label}");
        if wanted(&save_leaf) {
            measure(&save_leaf, ROUNDS, || {
                drop(black_box(package.to_bytes().expect("save")));
            });
        }
        let save_twice_leaf = format!("save_twice/leaf/{label}");
        if wanted(&save_twice_leaf) {
            measure(&save_twice_leaf, ROUNDS, || {
                let mut fresh = litchi_pptx::Package::from_bytes(&bytes).expect("open");
                drop(black_box(fresh.to_bytes().expect("first save")));
                drop(black_box(fresh.to_bytes().expect("second save")));
            });
        }
        let file = litchi_ops::open_bytes_with(bytes.clone(), &eager())?;
        let save_unified = format!("save/unified/{label}");
        if wanted(&save_unified) {
            measure(&save_unified, ROUNDS, || {
                drop(black_box(file.to_bytes().expect("save")));
            });
        }
        let save_twice_unified = format!("save_twice/unified/{label}");
        if wanted(&save_twice_unified) {
            measure(&save_twice_unified, ROUNDS, || {
                let fresh = litchi_ops::open_bytes_with(bytes.clone(), &eager()).expect("open");
                drop(black_box(fresh.to_bytes().expect("first save")));
                drop(black_box(fresh.to_bytes().expect("second save")));
            });
        }
    }

    let file = litchi_ops::open_bytes_with(small, &eager())?;
    if wanted("set/unified/deck20") {
        measure("set/unified/deck20", ROUNDS, || {
            drop(black_box(
                file.set("/slide[1]/shape[10]", props! { "text" => "changed" })
                    .expect("set"),
            ));
        });
    }
    if wanted("set_save/unified/deck20") {
        measure("set_save/unified/deck20", ROUNDS, || {
            let changed = file
                .set("/slide[1]/shape[10]", props! { "text" => "changed" })
                .expect("set");
            drop(black_box(changed.to_bytes().expect("save")));
        });
    }

    fn set_item(path: String, props: Props<'static>) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(props);
        item
    }
    let file = litchi_ops::open_bytes_with(large, &eager())?;
    let items: Vec<_> = (1..=1_000)
        .map(|index| {
            set_item(
                format!("/slide[1]/shape[{index}]/paragraph[1]/run[1]"),
                Props::new().with("bold", true).with("size", "20pt"),
            )
        })
        .collect();
    if wanted("batch1k/unified/run_format") {
        measure("batch1k/unified/run_format", 9, || {
            drop(black_box(
                file.batch(items.clone(), BatchMode::Atomic).expect("batch"),
            ));
        });
    }
    let texts: Vec<_> = (1..=1_000)
        .map(|index| {
            set_item(
                format!("/slide[1]/shape[{index}]"),
                Props::new().with("text", "changed"),
            )
        })
        .collect();
    if wanted("batch1k/unified/shape_text") {
        measure("batch1k/unified/shape_text", 9, || {
            drop(black_box(
                file.batch(texts.clone(), BatchMode::Atomic).expect("batch"),
            ));
        });
    }
    Ok(())
}

#[cfg(not(feature = "pptx"))]
fn main() {}
