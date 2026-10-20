//! Unified-against-leaf timing for the paths record 0308 adds.
//!
//! - `odp_run_format`: bold and size on one run of a 200-slide
//!   presentation, natively (`Package::edit` + `splice_scene` with
//!   `Change::SetTextProperties` + `commit`) and through `File::set`;
//! - `rtf_producer_insert`: add a paragraph at the start of 1,000 producer
//!   paragraphs (`\pard\plain` each), natively (`insert_paragraph` +
//!   `commit`, which falls back to the producer splice) and through
//!   `File::add`;
//! - `rtf_producer_replay`: move the first of 1,000 paragraphs that inherit
//!   root-level state (the value-exact replay path);
//! - `xls_row_insert_unindexed`: insert a row at the top of a worksheet
//!   with no INDEX record (`WithExtendedStyles.xls`).
//!
//! Output and method as `office_gaps_bench`. Build:
//! `cargo build --release -p litchi-ops --features xls,rtf,odp --example
//! office_gaps_new_bench`.

#[cfg(all(feature = "xls", feature = "rtf", feature = "odp"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_odp::scene::{Change, ObjectRef, StoryRef, TextPropertyUpdate};
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode, props};
    use litchi_xls::cell_values::Selector;

    const ROUNDS: usize = 15;
    const WARMUP: usize = 3;

    let only = std::env::var("OFFICE_GAPS_SCENARIO").ok();
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

    if wanted("odp_run_format") {
        // A 200-slide deck made through the unified layer.
        let mut deck = litchi_ops::create(litchi_core::FileFormat::Odp)?;
        for index in 0..200 {
            deck = deck.add("/", NewNode::new("slide"))?.file;
            let slide = format!("/slide[{}]", index + 1);
            deck = deck
                .add(
                    slide.as_str(),
                    NewNode::new("shape")
                        .prop("text", format!("Slide {index} text"))
                        .prop("width", "8cm")
                        .prop("height", "2cm"),
                )?
                .file;
        }
        let bytes = deck.to_bytes()?;
        let native = litchi_odp::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes(bytes.clone())?;
        let path = "/slide[100]/shape[1]/paragraph[1]/run[1]";
        let mut properties = TextPropertyUpdate::new();
        properties.bold = Some(true);
        properties.size = Some("24pt".to_owned());
        let changes = [Change::SetTextProperties {
            story: StoryRef::object(ObjectRef::new(99, 0)),
            paragraph: 0,
            run: Some(0),
            properties,
        }];
        let native_once = || {
            let mut edit = native.edit().expect("edit");
            edit.splice_scene(&changes).expect("splice");
            edit.commit().expect("commit")
        };
        let unified_once = || {
            unified
                .set(path, props! { "bold" => true, "size" => "24pt" })
                .expect("set")
        };
        assert_eq!(
            native_once().into_parts().0.as_bytes(),
            unified_once().to_bytes()?.as_slice()
        );
        paired(
            "odp_run_format_200",
            5,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }

    let producer = |inherit: bool| {
        let texts: Vec<_> = (0..1_000).map(|i| format!("paragraph {i:05}")).collect();
        let separator = if inherit {
            "\\par\r\n"
        } else {
            "\\par\r\n\\pard\\plain\\f0 "
        };
        format!(
            r"{{\rtf1\ansi{{\fonttbl{{\f0 Times;}}}}\pard\plain\qc\b\f0 {}\par}}",
            texts.join(separator)
        )
        .into_bytes()
    };

    if wanted("rtf_producer_insert") {
        let bytes = producer(false);
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes(bytes.as_slice())?;
        let native_once = || {
            let mut edit = native.edit();
            edit.insert_paragraph(Position::new(0), "Litchi")
                .expect("insert");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            unified
                .add(
                    "/body",
                    NewNode::new("paragraph")
                        .prop("text", "Litchi")
                        .at(Placement::Start),
                )
                .expect("add")
                .file
        };
        assert_eq!(native_once().to_bytes()?, unified_once().to_bytes()?);
        paired(
            "rtf_producer_insert_1k",
            10,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }

    if wanted("rtf_producer_replay") {
        let bytes = producer(true);
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
            "rtf_producer_replay_1k",
            10,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }

    if wanted("xls_row_insert_unindexed") {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/xls/WithExtendedStyles.xls"),
        )?;
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
        assert_eq!(native_once().to_bytes()?, unified_once().to_bytes()?);
        paired(
            "xls_row_insert_unindexed",
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

#[cfg(not(all(feature = "xls", feature = "rtf", feature = "odp")))]
fn main() {
    eprintln!("build with --features xls,rtf,odp");
}
