//! Paired ODT timing: the leaf `litchi_odt` API against the unified
//! `litchi_ops` API on the same generated document (record 0276).
//!
//! Scenarios, each timed as leaf and unified legs in ABBA order:
//!
//! - `open_first_get`: open from bytes and read the first paragraph's text
//!   (the body index is built inside the clock on both sides);
//! - `query_all`: every paragraph's text from a warm handle (leaf: the body
//!   index walk; unified: `query("p")` with a path per match);
//! - `query_all_owned` (record 0280): the same unified leg against a fair
//!   leaf equivalent that also keeps each match's handle and positional
//!   address;
//! - `single_set`: replace one mid-document paragraph's text and commit;
//! - `batch_bold_N`: set bold on N paragraphs (leaf: one `Edit` with N
//!   `set_text_properties`; unified: one `File::edit` with N direct-path
//!   sets, which the ODT run coalesces into one leaf commit).
//!
//! Inputs are built and cloned outside the clock; outputs are verified
//! outside the clock. Run pinned, release:
//! `taskset -c 7 cargo run --release -p litchi-ops --features odt --example odt_paired_bench`.

#[cfg(feature = "odt")]
use std::hint::black_box;
#[cfg(feature = "odt")]
use std::time::Instant;

#[cfg(feature = "odt")]
const ROUNDS: usize = 9;
#[cfg(feature = "odt")]
const WARMUP: usize = 2;

#[cfg(feature = "odt")]
fn document(paragraphs: usize) -> Vec<u8> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..paragraphs {
        builder
            .add_paragraph(&format!(
                "Paragraph {index} holds a few ordinary words of body text."
            ))
            .expect("paragraph");
    }
    builder.build().expect("build").to_bytes().expect("bytes")
}

#[cfg(feature = "odt")]
fn median(mut samples: Vec<f64>) -> f64 {
    samples.sort_by(f64::total_cmp);
    samples[samples.len() / 2]
}

/// Time `leaf` and `unified` in ABBA order; returns the medians in µs.
#[cfg(feature = "odt")]
fn paired(mut leaf: impl FnMut(), mut unified: impl FnMut()) -> (f64, f64) {
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
    (median(leaf_samples), median(unified_samples))
}

#[cfg(feature = "odt")]
fn first_paragraph_id(package: &litchi_odt::Package, index: usize) -> litchi_odt::body::NodeId {
    let document = package.document().expect("document");
    let body = document.body().expect("body");
    body.root()
        .children()
        .filter(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
        .nth(index)
        .expect("paragraph")
        .id()
}

#[cfg(feature = "odt")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_ops::props;

    let mut rows = Vec::new();
    let mut report = |name: &str, paragraphs: usize, (leaf, unified): (f64, f64)| {
        let ratio = unified / leaf;
        println!(
            "{name:>18} {paragraphs:>6} leaf {leaf:>12.1} µs  unified {unified:>12.1} µs  ratio {ratio:>6.3}"
        );
        rows.push(serde_json::json!({
            "scenario": name,
            "paragraphs": paragraphs,
            "leaf_us": leaf,
            "unified_us": unified,
            "ratio": ratio,
        }));
    };

    for paragraphs in [100usize, 1_000, 10_000] {
        let bytes = document(paragraphs);

        // open + first get
        let result = paired(
            || {
                let package = litchi_odt::Package::from_bytes(&bytes).expect("leaf open");
                let document = package.document().expect("document");
                let body = document.body().expect("body");
                let first = body
                    .root()
                    .children()
                    .find(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
                    .expect("paragraph");
                black_box(first.text().expect("text").len());
            },
            || {
                let file = litchi_ops::open_bytes(black_box(bytes.clone())).expect("open");
                let node = file.get("/body/p[1]").expect("get");
                black_box(node.view().text().expect("text").map(|text| text.len()));
            },
        );
        report("open_first_get", paragraphs, result);

        // full query on warm handles
        let package = litchi_odt::Package::from_bytes(&bytes)?;
        let document = package.document()?;
        let file = litchi_ops::open_bytes(bytes.clone())?;
        let _ = file.query("p")?.count();
        let _ = document.body()?;
        let result = paired(
            || {
                let body = document.body().expect("body");
                let mut total = 0usize;
                for node in body.nodes() {
                    if node.kind() == litchi_odt::body::Kind::Paragraph {
                        total += node.text().expect("text").len();
                    }
                }
                black_box(total);
            },
            || {
                let mut total = 0usize;
                for node in file.query("p").expect("query") {
                    let node = node.expect("node");
                    total += node
                        .view()
                        .text()
                        .expect("text")
                        .map_or(0, |text| text.len());
                    black_box(node.path());
                }
                black_box(total);
            },
        );
        report("query_all", paragraphs, result);

        // Record 0280: the fair leaf equivalent of the same query. Each
        // match is kept as a retainable handle with its positional address
        // (the parent and the ordinal among its paragraph siblings), which
        // is what a unified node and its path carry.
        let result = paired(
            || {
                let body = document.body().expect("body");
                let mut total = 0usize;
                let mut ordinals: Vec<(Option<litchi_odt::body::NodeId>, u32)> = Vec::new();
                for node in body.nodes() {
                    if node.kind() == litchi_odt::body::Kind::Paragraph {
                        total += node.text().expect("text").len();
                        let parent = node.parent().map(|parent| parent.id());
                        let ordinal = match ordinals.iter_mut().find(|(at, _)| *at == parent) {
                            Some((_, count)) => {
                                *count += 1;
                                *count
                            },
                            None => {
                                ordinals.push((parent, 1));
                                1
                            },
                        };
                        black_box((node, parent, ordinal));
                    }
                }
                black_box(total);
            },
            || {
                let mut total = 0usize;
                for node in file.query("p").expect("query") {
                    let node = node.expect("node");
                    total += node
                        .view()
                        .text()
                        .expect("text")
                        .map_or(0, |text| text.len());
                    black_box(node.path());
                }
                black_box(total);
            },
        );
        report("query_all_owned", paragraphs, result);

        // single set
        let middle = paragraphs / 2;
        let id = first_paragraph_id(&package, middle);
        let path = format!("/body/p[{}]", middle + 1);
        let result = paired(
            || {
                let mut edit = package.edit();
                edit.set_text(id, "Replaced paragraph text").expect("stage");
                black_box(edit.commit().expect("commit").into_package());
            },
            || {
                black_box(
                    file.set(
                        path.as_str(),
                        props! { "text" => "Replaced paragraph text" },
                    )
                    .expect("set"),
                );
            },
        );
        report("single_set", paragraphs, result);

        // batches
        let ids: Vec<_> = {
            let body = document.body()?;
            body.root()
                .children()
                .filter(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
                .map(|node| node.id())
                .collect()
        };
        let paths: Vec<String> = (1..=paragraphs)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        for count in [1usize, 10, 100, 1_000, 10_000] {
            if count > paragraphs {
                continue;
            }
            let mut bold = litchi_odt::body::TextPropertyUpdate::new();
            bold.bold = Some(true);
            let result = paired(
                || {
                    let mut edit = package.edit();
                    for id in &ids[..count] {
                        edit.set_text_properties(*id, &bold).expect("stage");
                    }
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    let commit = file
                        .edit(|tx| {
                            for path in &paths[..count] {
                                tx.set(path.as_str(), props! { "bold" => true })?;
                            }
                            Ok(())
                        })
                        .expect("edit");
                    black_box(commit.file());
                },
            );
            report(&format!("batch_bold_{count}"), paragraphs, result);
        }
        // Verify the unified batch equals the leaf batch outside the clock.
        let leaf_bytes = {
            let mut edit = package.edit();
            let mut bold = litchi_odt::body::TextPropertyUpdate::new();
            bold.bold = Some(true);
            for id in &ids {
                edit.set_text_properties(*id, &bold)?;
            }
            edit.commit()?.into_package().to_bytes()?
        };
        let unified_bytes = file
            .edit(|tx| {
                for path in &paths {
                    tx.set(path.as_str(), props! { "bold" => true })?;
                }
                Ok(())
            })?
            .file()
            .to_bytes()?;
        assert_eq!(
            leaf_bytes, unified_bytes,
            "unified batch equals the leaf batch"
        );
    }
    let json = serde_json::json!({
        "schema": "odt_paired_benchmark.v1",
        "rounds": ROUNDS,
        "warmup_rounds": WARMUP,
        "order": "ABBA",
        "rows": rows,
    });
    if let Some(path) = std::env::args().nth(1) {
        std::fs::write(path, serde_json::to_string_pretty(&json)?)?;
    }
    Ok(())
}

#[cfg(not(feature = "odt"))]
fn main() {
    eprintln!("enable the odt feature");
}
