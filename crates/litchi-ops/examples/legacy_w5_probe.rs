//! Record 0359 (W5 DOC/XLS/RTF contract): the unified routes of record 0351
//! §5 for the wave, timed on one build so the same source can run on the
//! wave's base and head (binary-level A-B-B-A).
//!
//! Every scenario uses only the `litchi_ops` API (stable across the wave)
//! and inputs built outside the clock with the leaf builders or plain RTF
//! source:
//!
//! - `doc/open_get`, `xls/open_get`, `rtf/open_get`: `open_bytes` + one `get`;
//! - `doc/get`: one `File::get` on an open file;
//! - `<fmt>/set`: one `File::set` on an open file (publishes a new file);
//! - `rtf/add`: one `File::add` of a paragraph on an open file;
//! - `doc/batch256`: one `File::edit` with 256 sets (the DOC body
//!   transaction's default operation bound);
//! - `xls/batch1k`, `rtf/batch1k`: one `File::edit` with 1,000 sets.
//!
//! `LITCHI_PROBE_ONLY=<prefix>` runs only the scenarios starting with it
//! (record 0359a, for profiling one route).
//!
//! Output: one JSON array of `{scenario, samples, median_us, min_us}` on
//! stdout. Run pinned, release:
//! `taskset -c 29 cargo run --release -p litchi-ops --features doc,xls,rtf --example legacy_w5_probe -- [rounds]`.

#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
mod probe {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::{NewNode, props};

    fn doc(paragraphs: usize) -> Vec<u8> {
        let mut builder = litchi_doc::Builder::new();
        for index in 0..paragraphs {
            builder
                .add_paragraph(&format!(
                    "Paragraph {index:05} holds a few ordinary words of body text."
                ))
                .expect("paragraph");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn xls(rows: usize, columns: usize) -> Vec<u8> {
        let mut builder = litchi_xls::writer::Builder::new();
        let data = builder.add_worksheet("Data").expect("sheet");
        for row in 0..rows {
            for column in 0..columns {
                builder
                    .write_number(
                        data,
                        row as u32,
                        column as u16,
                        (row * columns + column) as f64,
                    )
                    .expect("number");
            }
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn rtf(paragraphs: usize) -> Vec<u8> {
        let body: Vec<String> = (0..paragraphs)
            .map(|index| format!("Paragraph {index:05} holds a few ordinary words of body text."))
            .collect();
        format!(r"{{\rtf1\ansi {}}}", body.join(r"\par ")).into_bytes()
    }

    /// Whether `scenario` runs: every one, or those starting with
    /// `LITCHI_PROBE_ONLY` (for profiling one route; record 0359a).
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
            eprintln!("{scenario:>20} {median:>12.1} µs");
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
        let paragraphs: Vec<String> = (1..=1_000)
            .map(|index| format!("/body/p[{index}]"))
            .collect();

        let bytes = doc(1_000);
        probe.time("doc/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/body/p[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("doc/get", 50, || {
            black_box(
                file.get("/body/p[500]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        probe.time("doc/set", 5, || {
            black_box(
                file.set(
                    "/body/p[500]",
                    props! { "text" => "Replaced paragraph text" },
                )
                .expect("set"),
            );
        });
        probe.time("doc/batch256", 1, || {
            let commit = file
                .edit(|tx| {
                    for path in &paragraphs[..256] {
                        tx.set(path.as_str(), props! { "bold" => true })?;
                    }
                    Ok(())
                })
                .expect("edit");
            black_box(commit.file());
        });

        let bytes = xls(100, 10);
        probe.time("xls/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(file.get("/Data/C5").expect("get").path().to_string().len());
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("xls/set", 5, || {
            black_box(
                file.set("/Data/C5", props! { "value" => 42.0 })
                    .expect("set"),
            );
        });
        let cells: Vec<String> = (1..=100)
            .flat_map(|row| {
                ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
                    .into_iter()
                    .map(move |column| format!("/Data/{column}{row}"))
            })
            .collect();
        probe.time("xls/batch1k", 1, || {
            let commit = file
                .edit(|tx| {
                    for cell in &cells {
                        tx.set(cell.as_str(), props! { "value" => 7.0 })?;
                    }
                    Ok(())
                })
                .expect("edit");
            black_box(commit.file());
        });

        let bytes = rtf(1_000);
        probe.time("rtf/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/body/p[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("rtf/set", 5, || {
            black_box(
                file.set(
                    "/body/p[500]",
                    props! { "text" => "Replaced paragraph text" },
                )
                .expect("set"),
            );
        });
        probe.time("rtf/add", 5, || {
            black_box(
                file.add(
                    "/body",
                    NewNode::new("paragraph").prop("text", "Appended paragraph text"),
                )
                .expect("add"),
            );
        });
        probe.time("rtf/batch1k", 1, || {
            let commit = file
                .edit(|tx| {
                    for path in &paragraphs {
                        tx.set(
                            path.as_str(),
                            props! { "text" => "Batch replaced paragraph" },
                        )?;
                    }
                    Ok(())
                })
                .expect("edit");
            black_box(commit.file());
        });

        println!(
            "{}",
            serde_json::to_string_pretty(&probe.rows).expect("json")
        );
    }
}

fn main() {
    #[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
    probe::run();
    #[cfg(not(all(feature = "doc", feature = "xls", feature = "rtf")))]
    eprintln!("legacy_w5_probe needs the doc, xls and rtf features");
}
