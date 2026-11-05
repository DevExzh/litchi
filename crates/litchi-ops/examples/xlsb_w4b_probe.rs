//! Record 0360b (W4 XLSB contract): the XLSB routes of record 0351 §5, timed
//! on one build so the same source runs on the wave's base and head
//! (binary-level A-B-B-A).
//!
//! The leaf legs use only XLSB calls whose signatures the wave keeps
//! (`Package::{new, from_bytes, edit_cell_values, to_bytes}`,
//! `cell_values::Edit::{set_scalar, commit}`, `Commit::into_package`); the
//! unified legs use only the `litchi_ops` API. Inputs are built outside the
//! clock:
//!
//! - `leaf/open`: `Package::from_bytes` of the 100×10 workbook;
//! - `leaf/set`, `leaf/set_save`: one `edit_cell_values` commit of one
//!   scalar (then `to_bytes`);
//! - `leaf/batch1k`: one `edit_cell_values` commit of 1,000 scalars, the
//!   reference of the batch budget;
//! - `<owner>/open_get`: `open_bytes` + one `get` (`source` is the default
//!   source-backed open, `eager` is `SourceMode::Eager`);
//! - `<owner>/set`, `<owner>/set_save`: one `File::set` of a cell value
//!   (then `to_bytes`);
//! - `<owner>/format`: one `File::set` of `font.bold`;
//! - `<owner>/metadata`: one `File::set` of the `/metadata` title;
//! - `<owner>/batch1k`: one `File::edit` with 1,000 cell sets;
//! - `eager/comment`, `eager/hyperlink`, `eager/name`, `eager/merge`: one
//!   `File::add` of a note, a hyperlink and a defined name, one merge `set`;
//! - `eager/chain`: two `File::set`s, the second on the first's result.
//!
//! `LITCHI_PROBE_ONLY=<prefix>` runs only the scenarios starting with it.
//!
//! Output: one JSON array of `{scenario, samples, median_us, min_us}` on
//! stdout. Run pinned, release:
//! `taskset -c 29 cargo run --release -p litchi-ops --features xlsb --example xlsb_w4b_probe -- [rounds]`.

#[cfg(feature = "xlsb")]
mod probe {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::{NewNode, OpenOptions, SourceMode, props};
    use litchi_xlsb::cell_values::{Reference, Scalar};

    const ROWS: u32 = 100;
    const COLUMNS: u32 = 10;
    const LETTERS: [&str; 10] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];

    /// A one-sheet workbook of 100×10 numbers, built with the leaf edit.
    fn workbook() -> Vec<u8> {
        let package = litchi_xlsb::Package::new().expect("new package");
        let mut edit = package
            .edit_cell_values(0_usize)
            .expect("cell edit")
            .resolve()
            .expect("worksheet resolution")
            .expect("worksheet");
        for row in 0..ROWS {
            for column in 0..COLUMNS {
                edit.set_scalar(
                    Reference::new(row, column).expect("reference"),
                    Scalar::Number(f64::from(row * COLUMNS + column)),
                )
                .expect("scalar");
            }
        }
        edit.commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes")
    }

    /// Whether `scenario` runs: every one, or those starting with
    /// `LITCHI_PROBE_ONLY`.
    fn selected(scenario: &str) -> bool {
        std::env::var("LITCHI_PROBE_ONLY").map_or(true, |prefix| scenario.starts_with(&prefix))
    }

    struct Probe {
        rounds: usize,
        rows: Vec<serde_json::Value>,
    }

    impl Probe {
        fn time(&mut self, scenario: &str, calls: usize, mut run: impl FnMut()) {
            if !selected(scenario) {
                return;
            }
            for _ in 0..2 {
                run();
            }
            let mut samples = Vec::with_capacity(self.rounds);
            for _ in 0..self.rounds {
                let start = Instant::now();
                for _ in 0..calls {
                    run();
                }
                samples.push(start.elapsed().as_secs_f64() * 1e6 / calls as f64);
            }
            samples.sort_by(f64::total_cmp);
            let median = samples[samples.len() / 2];
            eprintln!("{scenario:>24} {median:>12.1} µs");
            self.rows.push(serde_json::json!({
                "scenario": scenario,
                "samples": samples.len(),
                "median_us": median,
                "min_us": samples[0],
            }));
        }
    }

    pub(super) fn run() {
        let rounds = std::env::args()
            .nth(1)
            .and_then(|value| value.parse().ok())
            .unwrap_or(15);
        let mut probe = Probe {
            rounds,
            rows: Vec::new(),
        };
        let bytes = workbook();
        let cells: Vec<String> = (1..=ROWS)
            .flat_map(|row| {
                LETTERS
                    .into_iter()
                    .map(move |column| format!("/Sheet1/{column}{row}"))
            })
            .collect();
        let references: Vec<Reference> = (0..ROWS)
            .flat_map(|row| {
                (0..COLUMNS).map(move |column| Reference::new(row, column).expect("reference"))
            })
            .collect();

        // Leaf legs: the XLSB package API the unified routes wrap.
        probe.time("leaf/open", 5, || {
            black_box(litchi_xlsb::Package::from_bytes(&bytes).expect("open"));
        });
        let package = litchi_xlsb::Package::from_bytes(&bytes).expect("open");
        let c5 = Reference::new(4, 2).expect("reference");
        let leaf_set = || {
            let mut edit = package
                .edit_cell_values(0_usize)
                .expect("cell edit")
                .resolve()
                .expect("worksheet resolution")
                .expect("worksheet");
            edit.set_scalar(c5, Scalar::Number(42.0)).expect("scalar");
            edit.commit().expect("commit").into_package()
        };
        probe.time("leaf/set", 5, || {
            black_box(leaf_set());
        });
        probe.time("leaf/set_save", 5, || {
            black_box(leaf_set().to_bytes().expect("bytes"));
        });
        probe.time("leaf/batch1k", 1, || {
            let mut edit = package
                .edit_cell_values(0_usize)
                .expect("cell edit")
                .resolve()
                .expect("worksheet resolution")
                .expect("worksheet");
            for reference in &references {
                edit.set_scalar(*reference, Scalar::Number(7.0))
                    .expect("scalar");
            }
            black_box(edit.commit().expect("commit").into_package());
        });

        for owner in ["source", "eager"] {
            let options = if owner == "eager" {
                OpenOptions::new().source_mode(SourceMode::Eager)
            } else {
                OpenOptions::new()
            };
            probe.time(&format!("{owner}/open_get"), 5, || {
                let file = litchi_ops::open_bytes_with(bytes.clone(), &options).expect("open");
                black_box(
                    file.get("/Sheet1/C5")
                        .expect("get")
                        .path()
                        .to_string()
                        .len(),
                );
            });
            let file = litchi_ops::open_bytes_with(bytes.clone(), &options).expect("open");
            probe.time(&format!("{owner}/set"), 5, || {
                black_box(
                    file.set("/Sheet1/C5", props! { "value" => 42.0 })
                        .expect("set"),
                );
            });
            probe.time(&format!("{owner}/set_save"), 5, || {
                let after = file
                    .set("/Sheet1/C5", props! { "value" => 42.0 })
                    .expect("set");
                black_box(after.to_bytes().expect("bytes"));
            });
            probe.time(&format!("{owner}/format"), 5, || {
                black_box(
                    file.set("/Sheet1/C5", props! { "font.bold" => true })
                        .expect("format"),
                );
            });
            probe.time(&format!("{owner}/metadata"), 5, || {
                black_box(
                    file.set("/metadata", props! { "title" => "Probe" })
                        .expect("metadata"),
                );
            });
            probe.time(&format!("{owner}/batch1k"), 1, || {
                let commit = file
                    .edit(|tx| {
                        for cell in &cells {
                            tx.set(cell.as_str(), props! { "value" => 7.0 })?;
                        }
                        Ok(())
                    })
                    .expect("edit");
                black_box(commit.into_file());
            });
        }

        let file = litchi_ops::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .expect("open");
        probe.time("eager/comment", 5, || {
            black_box(
                file.add(
                    "/Sheet1",
                    NewNode::new("comment")
                        .prop("ref", "B2")
                        .prop("text", "Check")
                        .prop("author", "Ann"),
                )
                .expect("comment"),
            );
        });
        probe.time("eager/hyperlink", 5, || {
            black_box(
                file.add(
                    "/Sheet1/A9",
                    NewNode::new("hyperlink").prop("url", "https://example.com/"),
                )
                .expect("hyperlink"),
            );
        });
        probe.time("eager/name", 5, || {
            black_box(
                file.add(
                    "/",
                    NewNode::new("namedrange")
                        .prop("name", "Totals")
                        .prop("ref", "Sheet1!$B$1:$B$5"),
                )
                .expect("name"),
            );
        });
        probe.time("eager/merge", 5, || {
            black_box(
                file.set("/Sheet1/C1:D2", props! { "merge" => true })
                    .expect("merge"),
            );
        });
        probe.time("eager/chain", 5, || {
            let first = file
                .set("/Sheet1/C5", props! { "value" => 1.0 })
                .expect("first");
            black_box(
                first
                    .set("/Sheet1/C6", props! { "value" => 2.0 })
                    .expect("second"),
            );
        });

        println!(
            "{}",
            serde_json::to_string_pretty(&probe.rows).expect("json")
        );
    }
}

#[cfg(feature = "xlsb")]
fn main() {
    probe::run();
}

#[cfg(not(feature = "xlsb"))]
fn main() {}
