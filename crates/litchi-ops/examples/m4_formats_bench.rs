//! Paired timing of the record 0295 engines on more formats: the unified
//! call against the best hand-written use of the same leaf API.
//!
//! - `odt_find_M`: a body of N one-run paragraphs, `needle` in M of them
//!   (1, or every one). Leaf leg: walk the body index, rewrite each
//!   character-data run that holds the needle with `Edit::set_text`, commit.
//!   Unified leg: `File::set("/", {find: needle, replace: done})` (the
//!   match count is checked through `File::apply` outside the clock).
//! - `odt_merge`: the same body with one `{{key}}` placeholder per
//!   paragraph. Leaf leg: the same walk with a plain `{{key}}` replace.
//!   Unified leg: `File::merge`.
//! - `doc_find_M`: N paragraphs, the needle in M. Leaf leg: scan the warm
//!   reader's main-story paragraph texts (`Document::body_paragraphs`) and
//!   stage one `replace_paragraph` per matching paragraph (whole paragraphs
//!   of one run: a lower bound for run-level rewrites). Unified leg: `set /`
//!   with `find=`.
//! - `ods_import`: N rows of four fields (text, number, boolean, text) into
//!   an empty sheet. Leaf leg: split the CSV by hand and stage the typed
//!   cells with one `set_cells` per 1,024 changes. Unified leg:
//!   `File::import`.
//!
//! Inputs are built outside the clock; every sample edits the same
//! immutable owner and commits without publishing. Two warm-up rounds,
//! then ABBA rounds; medians in µs. Run pinned, release:
//! `taskset -c 7 cargo run --release -p litchi-ops --example m4_formats_bench
//! --features odt,doc,ods`.

#[cfg(all(feature = "odt", feature = "doc", feature = "ods"))]
mod bench {
    use std::hint::black_box;
    use std::io::Cursor;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::merge::{Options as MergeOptions, Record};
    use litchi_ops::{BatchItem, File, Props, open_bytes};

    const ROUNDS: usize = 9;
    const WARMUP: usize = 2;

    fn median(mut samples: Vec<f64>) -> f64 {
        samples.sort_by(f64::total_cmp);
        samples[samples.len() / 2]
    }

    /// Time `leaf` and `unified` in ABBA order; the medians in µs.
    fn paired(mut leaf: impl FnMut(), mut unified: impl FnMut()) -> (f64, f64, Vec<f64>, Vec<f64>) {
        let time = |run: &mut dyn FnMut()| {
            let start = Instant::now();
            run();
            start.elapsed().as_secs_f64() * 1e6
        };
        for _ in 0..WARMUP {
            leaf();
            unified();
        }
        let (mut leaf_samples, mut unified_samples) = (Vec::new(), Vec::new());
        for round in 0..ROUNDS {
            if round % 2 == 0 {
                leaf_samples.push(time(&mut leaf));
                unified_samples.push(time(&mut unified));
            } else {
                unified_samples.push(time(&mut unified));
                leaf_samples.push(time(&mut leaf));
            }
        }
        (
            median(leaf_samples.clone()),
            median(unified_samples.clone()),
            leaf_samples,
            unified_samples,
        )
    }

    fn odt(paragraphs: usize, line: impl Fn(usize) -> String) -> Vec<u8> {
        let mut builder = litchi_odt::Builder::new();
        for index in 0..paragraphs {
            builder.add_paragraph(&line(index)).expect("paragraph");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn doc(paragraphs: usize, line: impl Fn(usize) -> String) -> Vec<u8> {
        let mut builder = litchi_doc::Builder::new();
        for index in 0..paragraphs {
            builder.add_paragraph(&line(index)).expect("paragraph");
        }
        let mut output = Cursor::new(Vec::new());
        builder
            .build()
            .expect("build")
            .write_to(&mut output)
            .expect("write");
        output.into_inner()
    }

    /// The leaf ODT rewrite: every character-data run holding `from`.
    fn odt_leaf(package: &litchi_odt::Package, from: &str, to: &str) -> litchi_odt::Package {
        let document = package.document().expect("document");
        let body = document.body().expect("body");
        let mut edit = package.edit();
        for node in body.nodes() {
            if node.kind() != litchi_odt::body::Kind::Text {
                continue;
            }
            let text = node.text().expect("text");
            if text.contains(from) {
                edit.set_text(node.id(), text.replace(from, to))
                    .expect("set text");
            }
        }
        edit.commit().expect("commit").into_package()
    }

    fn find_item(pattern: &str, replacement: &str) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some("/".into());
        item.props = Some(
            Props::new()
                .with("find", pattern.to_owned())
                .with("replace", replacement.to_owned()),
        );
        item
    }

    pub(super) fn run() -> Result<(), Box<dyn std::error::Error>> {
        let mut rows = Vec::new();
        let mut report = |name: String, size: usize, result: (f64, f64, Vec<f64>, Vec<f64>)| {
            let (leaf, unified, leaf_samples, unified_samples) = result;
            let ratio = unified / leaf;
            println!(
                "{name:>16} {size:>6} leaf {leaf:>12.1} µs  unified {unified:>12.1} µs  ratio {ratio:>6.3}"
            );
            rows.push(serde_json::json!({
                "scenario": name,
                "size": size,
                "leaf_median_us": leaf,
                "unified_median_us": unified,
                "ratio": ratio,
                "leaf_us": leaf_samples,
                "unified_us": unified_samples,
            }));
        };

        // For a profiler: `M4_ONLY=doc*` runs one family, `M4_ONLY=doc_find_1`
        // one scenario; `M4_SIZE=5000` one size.
        let only = std::env::var("M4_ONLY").ok();
        let size = std::env::var("M4_SIZE")
            .ok()
            .and_then(|size| size.parse::<usize>().ok());
        let wanted = |name: &str| {
            only.as_deref().is_none_or(|only| {
                only == name
                    || only
                        .strip_suffix('*')
                        .is_some_and(|prefix| name.starts_with(prefix))
            })
        };
        let sized = |value: usize| size.is_none_or(|size| size == value);
        for paragraphs in [1_000usize, 10_000] {
            if !wanted("odt_find_1") && !wanted("odt_merge") && !wanted("odt*")
                || !sized(paragraphs)
            {
                continue;
            }
            for matches in [1usize, paragraphs] {
                if !wanted(&format!("odt_find_{matches}")) && !wanted("odt*") {
                    continue;
                }
                let every = paragraphs / matches;
                let bytes = odt(paragraphs, |index| {
                    if index % every == 0 {
                        format!("Paragraph {index} has the needle in it.")
                    } else {
                        format!("Paragraph {index} holds ordinary words.")
                    }
                });
                let package = litchi_odt::Package::from_bytes(&bytes)?;
                let file = open_bytes(bytes)?;
                let item = find_item("needle", "done");
                assert_eq!(file.apply(&item)?.value().matched(), Some(matches));
                let props = Props::new().with("find", "needle").with("replace", "done");
                let result = paired(
                    || {
                        black_box(odt_leaf(&package, "needle", "done"));
                    },
                    || {
                        black_box(file.set("/", props.clone()).expect("unified find"));
                    },
                );
                report(format!("odt_find_{matches}"), paragraphs, result);
            }

            let bytes = odt(paragraphs, |index| {
                format!("Dear {{{{key}}}}, item {index}.")
            });
            let package = litchi_odt::Package::from_bytes(&bytes)?;
            let file = open_bytes(bytes)?;
            let record: Record = [("key", "Ada")].into_iter().collect();
            let result = paired(
                || {
                    black_box(odt_leaf(&package, "{{key}}", "Ada"));
                },
                || {
                    let commit = file
                        .merge(&record, &MergeOptions::new())
                        .expect("unified merge");
                    assert_eq!(commit.value().replacements(), paragraphs);
                    black_box(commit.file());
                },
            );
            report("odt_merge".to_owned(), paragraphs, result);
        }

        for paragraphs in [1_000usize, 5_000] {
            if !sized(paragraphs) {
                continue;
            }
            for matches in [1usize, 100] {
                if !wanted(&format!("doc_find_{matches}")) {
                    continue;
                }
                let every = paragraphs / matches;
                let line = |index: usize| {
                    if index % every == 0 {
                        format!("Paragraph {index:05} has the needle in it.")
                    } else {
                        format!("Paragraph {index:05} holds ordinary words.")
                    }
                };
                let bytes = doc(paragraphs, line);
                let package = litchi_doc::Package::from_bytes(&bytes)?;
                let reader = package.document();
                let file: File = open_bytes(bytes)?;
                let item = find_item("needle", "done");
                assert_eq!(file.apply(&item)?.value().matched(), Some(matches));
                let props = Props::new().with("find", "needle").with("replace", "done");
                let result = paired(
                    || {
                        let mut targets = Vec::new();
                        for (index, text) in reader.body_paragraphs().enumerate() {
                            let text = text.expect("paragraph");
                            if text.contains("needle") {
                                targets.push((index, text.replace("needle", "done")));
                            }
                        }
                        let mut edit = package.edit().expect("leaf edit");
                        for (index, text) in targets.iter().rev() {
                            edit.replace_paragraph(Position::new(*index), text)
                                .expect("replace");
                        }
                        black_box(edit.commit().expect("commit").into_package());
                    },
                    || {
                        black_box(file.set("/", props.clone()).expect("unified find"));
                    },
                );
                report(format!("doc_find_{matches}"), paragraphs, result);
            }
        }

        for rows_count in [1_000usize, 10_000] {
            if !wanted("ods_import") || !sized(rows_count) {
                continue;
            }
            let mut csv = String::new();
            for row in 0..rows_count {
                csv.push_str(&format!("name{row},{row}.5,TRUE,note {row}\n"));
            }
            let file = litchi_ops::create(litchi_core::FileFormat::Ods)?;
            let bytes = file.to_bytes()?;
            let package = litchi_ods::Package::from_bytes(&bytes)?;
            let options = litchi_ops::import::Options::csv();
            let result = paired(
                || {
                    use litchi_ods::worksheet::{CellChange, MAX_CELL_CHANGES, Selector};
                    use litchi_ods::{Cell, CellValue};
                    let mut changes = Vec::with_capacity(rows_count * 4);
                    for (row, line) in csv.lines().enumerate() {
                        for (column, field) in line.split(',').enumerate() {
                            let value = if let Ok(number) = field.parse::<f64>() {
                                CellValue::Number(number)
                            } else if field.eq_ignore_ascii_case("true") {
                                CellValue::Boolean(true)
                            } else {
                                CellValue::Text(field.to_owned())
                            };
                            changes.push(CellChange::new(row, column, Cell::new(value, field)));
                        }
                    }
                    let mut edit = package.edit().expect("leaf edit");
                    edit.worksheets(|worksheets| {
                        let mut rest = changes;
                        while !rest.is_empty() {
                            let tail = rest.split_off(rest.len().min(MAX_CELL_CHANGES));
                            worksheets.set_cells(Selector::Position(Position::new(0)), rest)?;
                            rest = tail;
                        }
                        Ok(())
                    })
                    .expect("leaf cells");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    let commit = file
                        .import("/Sheet1", csv.as_bytes(), &options)
                        .expect("unified import");
                    assert_eq!(commit.value().cells(), rows_count * 4);
                    black_box(commit.file());
                },
            );
            report("ods_import".to_owned(), rows_count, result);
        }

        println!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "schema": "litchi.unified-ops.m4-formats.v1",
                "rounds": ROUNDS,
                "order": "ABBA",
                "scenarios": rows,
            }))?
        );
        Ok(())
    }
}

#[cfg(all(feature = "odt", feature = "doc", feature = "ods"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    bench::run()
}

#[cfg(not(all(feature = "odt", feature = "doc", feature = "ods")))]
fn main() {}
