//! Record 0359a (W5 follow-up): the routes record 0359a changed, timed so
//! the same workloads run on the wave's base (record 0359's tip) and head
//! (binary-level A-B-B-A).
//!
//! Source-backed publishers (F2) — each call opens the source, stages, commits
//! and writes the complete output to a `Vec`; only the calls into the leaf
//! crate differ between the base and head builds (`mod routes`):
//!
//! - `doc/source_set`: one same-width paragraph replacement (1k paragraphs);
//! - `xls/source_cells1`, `xls/source_cells100`: numeric value fields
//!   (100×10 sheet), on the base through `commit_source_backed` and
//!   `commit_source_backed_plan` (`*_plan`);
//! - `xls/source_comments1`, `xls/source_comments256`: same-length comment
//!   replacements (256 comments);
//! - `xls/source_hide7`: seven `BoundSheet8` visibility splices (8 sheets);
//! - `rtf/source_append`: one paragraph appended at the proven tail (1k
//!   paragraphs; the document is parsed outside the clock).
//!
//! DOC body batch (the 0359 follow-up): `doc/batch1k_ops` is one
//! `File::edit_with_options` of 1,000 `bold` sets, `doc/batch1k_leaf` the
//! leaf `Package::edit_with_options` with the transaction's operation bound
//! raised to 1,000 and 1,000 `set_paragraph_bold` calls, and
//! `doc/batch1k_leaf_write` repeats that leaf route through complete
//! `Package::write_to` publication. `doc/leaf_write` is the corresponding
//! one-toggle commit-and-write route; all run
//! `LITCHI_PROBE_BATCH_ROUNDS` rounds (default 7).
//!
//! `LITCHI_PROBE_ONLY=<prefix>` runs only the scenarios starting with it.
//! `LITCHI_PROBE_OUTPUT_DIR=<directory>` additionally writes each selected
//! scenario's output as `<scenario-with-slashes-replaced>.bin`.
//!
//! Output: one JSON array of `{scenario, samples, median_us, min_us}` on
//! stdout. Run pinned, release:
//! `taskset -c 29 cargo run --release -p litchi-ops --features doc,xls,rtf --example legacy_w5_source_probe -- [rounds]`.

#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
mod probe {
    use std::hint::black_box;
    use std::path::PathBuf;
    use std::time::Instant;

    use litchi_ops::props;
    use litchi_ops::transaction::SemanticLimits;

    pub(super) fn doc(paragraphs: usize) -> Vec<u8> {
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

    fn xls_comments(count: usize) -> Vec<u8> {
        let mut builder = litchi_xls::writer::Builder::new();
        let sheet = builder.add_worksheet("Comments").expect("sheet");
        for index in 0..count {
            builder
                .add_comment(
                    sheet,
                    index as u32,
                    1,
                    &format!("Source {index:03}"),
                    &format!("comment source {index:03}"),
                )
                .expect("comment");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn xls_sheets(count: usize) -> Vec<u8> {
        let mut builder = litchi_xls::writer::Builder::new();
        for index in 0..count {
            let sheet = builder
                .add_worksheet(&format!("Sheet{index}"))
                .expect("sheet");
            builder
                .write_number(sheet, 0, 0, index as f64)
                .expect("number");
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
        output_dir: Option<PathBuf>,
        rows: Vec<serde_json::Value>,
    }

    impl Probe {
        /// Record the SHA-256 of the scenario's output on its row, so the
        /// base and head outputs can be compared byte for byte.
        fn output(&mut self, bytes: &[u8]) {
            let scenario = self
                .rows
                .last()
                .and_then(|row| row["scenario"].as_str())
                .map(str::to_owned);
            if let (Some(output_dir), Some(scenario)) = (&self.output_dir, scenario) {
                let filename = format!("{}.bin", scenario.replace('/', "_"));
                std::fs::write(output_dir.join(filename), bytes).expect("probe output");
            }
            if let Some(row) = self.rows.last_mut() {
                row["output_sha256"] = litchi_core::BlobId::of(bytes).as_hex().into();
            }
        }

        fn time(&mut self, scenario: &str, calls: usize, run: impl FnMut()) {
            self.time_rounds(scenario, calls, self.rounds, run);
        }

        fn time_rounds(
            &mut self,
            scenario: &str,
            calls: usize,
            rounds: usize,
            mut run: impl FnMut(),
        ) {
            if !selected(scenario) {
                return;
            }
            for _ in 0..2 {
                run();
            }
            let mut samples = Vec::with_capacity(rounds);
            for _ in 0..rounds {
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
        let output_dir = std::env::var_os("LITCHI_PROBE_OUTPUT_DIR").map(PathBuf::from);
        if let Some(output_dir) = &output_dir {
            std::fs::create_dir_all(output_dir).expect("probe output directory");
        }
        let mut probe = Probe {
            rounds,
            output_dir,
            rows: Vec::new(),
        };

        let bytes = doc(1_000);
        if selected("doc/source_set") {
            let expected = super::routes::doc_source_set(&bytes);
            probe.time("doc/source_set", 5, || {
                assert_eq!(super::routes::doc_source_set(&bytes), expected);
            });
            probe.output(&expected);
        }

        let batch_rounds = std::env::var("LITCHI_PROBE_BATCH_ROUNDS")
            .ok()
            .and_then(|value| value.parse().ok())
            .unwrap_or(7);

        if selected("doc/batch1k_ops") {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            let paragraphs: Vec<String> = (1..=1_000)
                .map(|index| format!("/body/p[{index}]"))
                .collect();
            // The base build retains one complete snapshot per isolated `bold`
            // step, past the default 256 MiB retention bound: it raises that
            // bound (`routes::DOC_BATCH_SNAPSHOT_BYTES`) and runs fewer rounds
            // (`LITCHI_PROBE_BATCH_ROUNDS`, default 7).
            let limits = SemanticLimits::builder()
                .max_operations(1_000)
                .expect("bound")
                .max_snapshot_bytes(super::routes::DOC_BATCH_SNAPSHOT_BYTES)
                .expect("bound")
                .build()
                .expect("limits");
            probe.time_rounds("doc/batch1k_ops", 1, batch_rounds, || {
                let commit = file
                    .edit_with_options(
                        litchi_ops::EditOptions::new().semantic_limits(limits),
                        |tx| {
                            for path in &paragraphs {
                                tx.set(path.as_str(), props! { "bold" => true })?;
                            }
                            Ok(())
                        },
                    )
                    .expect("edit");
                black_box(commit.file());
            });
        }

        if selected("doc/batch1k_leaf") {
            let package = litchi_doc::Package::from_bytes(&bytes).expect("package");
            probe.time_rounds("doc/batch1k_leaf", 1, batch_rounds, || {
                let options = litchi_doc::EditOptions::default().transaction_limits(
                    litchi_doc::body_text::TransactionLimits::new(1_000, usize::MAX, usize::MAX),
                );
                let mut edit = package.edit_with_options(options).expect("edit");
                for index in 0..1_000 {
                    edit.set_paragraph_bold(litchi_core::Position::new(index), true)
                        .expect("bold");
                }
                black_box(edit.commit().expect("commit").into_package());
            });
        }
        if selected("doc/batch1k_leaf_write") {
            let expected_leaf_write = super::routes::doc_batch_leaf_write(&bytes, 1_000);
            probe.time_rounds("doc/batch1k_leaf_write", 1, batch_rounds, || {
                assert_eq!(
                    super::routes::doc_batch_leaf_write(&bytes, 1_000),
                    expected_leaf_write
                );
            });
            probe.output(&expected_leaf_write);
        }

        if selected("doc/leaf_write") {
            let expected_leaf = super::routes::doc_leaf_write(&bytes);
            probe.time_rounds("doc/leaf_write", 1, batch_rounds, || {
                assert_eq!(super::routes::doc_leaf_write(&bytes), expected_leaf);
            });
            probe.output(&expected_leaf);
        }

        let bytes = xls(100, 10);
        for (scenario, cells) in [("xls/source_cells1", 1), ("xls/source_cells100", 100)] {
            let plan_scenario = format!("{scenario}_plan");
            if selected(scenario) || (super::routes::PLAN_ONLY_ROUTE && selected(&plan_scenario)) {
                let expected = super::routes::xls_source_cells(&bytes, cells, false);
                if selected(scenario) {
                    probe.time(scenario, 5, || {
                        assert_eq!(
                            super::routes::xls_source_cells(&bytes, cells, false),
                            expected
                        );
                    });
                    probe.output(&expected);
                }
                if super::routes::PLAN_ONLY_ROUTE && selected(&plan_scenario) {
                    probe.time(&plan_scenario, 5, || {
                        assert_eq!(
                            super::routes::xls_source_cells(&bytes, cells, true),
                            expected
                        );
                    });
                    probe.output(&expected);
                }
            }
        }

        let bytes = xls_comments(256);
        for (scenario, count) in [("xls/source_comments1", 1), ("xls/source_comments256", 256)] {
            if selected(scenario) {
                let expected = super::routes::xls_source_comments(&bytes, count);
                probe.time(scenario, 5, || {
                    assert_eq!(super::routes::xls_source_comments(&bytes, count), expected);
                });
                probe.output(&expected);
            }
        }

        let bytes = xls_sheets(8);
        if selected("xls/source_hide7") {
            let expected = super::routes::xls_source_hide(&bytes, 7);
            probe.time("xls/source_hide7", 5, || {
                assert_eq!(super::routes::xls_source_hide(&bytes, 7), expected);
            });
            probe.output(&expected);
        }

        let bytes = rtf(1_000);
        if selected("rtf/source_append") {
            let document = litchi_rtf::Package::from_bytes(&bytes).expect("document");
            let expected = super::routes::rtf_source_append(&document, &bytes);
            probe.time("rtf/source_append", 5, || {
                assert_eq!(
                    super::routes::rtf_source_append(&document, &bytes),
                    expected
                );
            });
            probe.output(&expected);
        }

        println!(
            "{}",
            serde_json::to_string_pretty(&probe.rows).expect("json")
        );
    }
}

/// The leaf calls of each source-backed route on this build (record 0359a
/// head: the deferred source-backed families).
#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
mod routes {
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;

    use litchi_core::{OwnedSource, Position, ReadAt};
    use litchi_xls::cell_values::{Reference, Selector};

    /// Whether this build has the separate plan-only XLS cell route.
    pub(super) const PLAN_ONLY_ROUTE: bool = false;

    /// The ops retention bound of the DOC 1k batch (the default).
    pub(super) const DOC_BATCH_SNAPSHOT_BYTES: usize = 256 * 1024 * 1024;

    fn same_width(text: &str) -> String {
        text.replacen("holds", "keeps", 1)
    }

    pub(super) fn doc_source_set(bytes: &[u8]) -> Vec<u8> {
        let package = litchi_doc::source_backed::Package::from_bytes(bytes.to_vec()).expect("open");
        let mut edit = package.edit_paragraph(Position::new(500)).expect("edit");
        let text = same_width(edit.text().expect("resolve paragraph"));
        edit.set_text(text).expect("set");
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.package().write_to(&mut output).expect("write");
        output
    }

    /// Run the 1k leaf body-toggle workload through complete DOC publication.
    /// This route includes the final `write_to` so the performance follow-up
    /// measures the user-visible commit-and-write path as well as staging.
    pub(super) fn doc_batch_leaf_write(bytes: &[u8], paragraphs: usize) -> Vec<u8> {
        let package = litchi_doc::Package::from_bytes(bytes).expect("package");
        let options = litchi_doc::EditOptions::default().transaction_limits(
            litchi_doc::body_text::TransactionLimits::new(paragraphs, usize::MAX, usize::MAX),
        );
        let mut edit = package.edit_with_options(options).expect("edit");
        for index in 0..paragraphs {
            edit.set_paragraph_bold(Position::new(index), true)
                .expect("bold");
        }
        let package = edit.commit().expect("commit").into_package();
        let mut output = Vec::new();
        package.write_to(&mut output).expect("write");
        output
    }

    /// Run one body toggle through complete DOC publication. This is the
    /// ordinary leaf control for the 1k coalescing measurement.
    pub(super) fn doc_leaf_write(bytes: &[u8]) -> Vec<u8> {
        let package = litchi_doc::Package::from_bytes(bytes).expect("package");
        let mut edit = package.edit().expect("edit");
        edit.set_paragraph_bold(Position::new(500), true)
            .expect("bold");
        let package = edit.commit().expect("commit").into_package();
        let mut output = Vec::new();
        package.write_to(&mut output).expect("write");
        output
    }

    fn workbook(bytes: &[u8]) -> litchi_xls::SourceBackedWorkbook {
        litchi_xls::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
            .expect("open")
    }

    pub(super) fn xls_source_cells(bytes: &[u8], cells: usize, _plan_only: bool) -> Vec<u8> {
        let mut edit = workbook(bytes).edit_cells().expect("edit");
        for index in 0..cells {
            let reference = Reference::new((index / 10) as u32, (index % 10) as u32).expect("ref");
            edit.set_numeric(Selector::Name("Data"), reference, 7.0 + index as f64)
                .expect("set");
        }
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.package().write_to(&mut output).expect("write");
        output
    }

    fn comment(index: usize) -> litchi_xls::comments::Value {
        litchi_xls::comments::Value::new(
            format!("Target {index:03}"),
            format!("comment target {index:03}"),
        )
        .expect("value")
    }

    pub(super) fn xls_source_comments(bytes: &[u8], count: usize) -> Vec<u8> {
        let mut edit = workbook(bytes).edit_comments().expect("edit");
        let updates = (0..count).map(|index| {
            litchi_xls::comments::Update::new(
                Reference::new(index as u32, 1).expect("ref"),
                comment(index),
            )
        });
        edit.replace_many(Selector::Position(Position::new(0)), updates)
            .expect("replace");
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.package().write_to(&mut output).expect("write");
        output
    }

    pub(super) fn xls_source_hide(bytes: &[u8], count: usize) -> Vec<u8> {
        let mut edit = workbook(bytes).edit_sheet_visibility().expect("edit");
        for position in 1..=count {
            edit.hide(position.into()).expect("hide");
        }
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.package().write_to(&mut output).expect("write");
        output
    }

    fn context() -> litchi_core::ExecutionContext {
        let budget = litchi_core::Budget::root(
            "legacy-w5-source-probe",
            litchi_core::Limits::new(1 << 30, 1 << 30, 1 << 30, 1 << 30, 64, 1 << 40),
        );
        let (_source, token) = litchi_core::CancellationSource::pair();
        let limits = litchi_core::ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one"),
            NonZeroUsize::new(1).expect("one"),
            NonZeroU64::new(64 * 1024).expect("window"),
            0,
        )
        .expect("limits");
        litchi_core::ExecutionContext::new(budget, token, limits)
    }

    pub(super) fn rtf_source_append(document: &litchi_rtf::Package, bytes: &[u8]) -> Vec<u8> {
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.to_vec()));
        let package =
            litchi_rtf::source_backed::Package::open(document, source, context()).expect("open");
        let mut edit = package.edit_tail_append().expect("edit");
        edit.append_text_paragraphs(&["Appended paragraph."])
            .expect("append");
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.package().write_to(&mut output).expect("write");
        output
    }
}

fn main() {
    #[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
    probe::run();
}
