//! Paired leaf-vs-unified measurements of Pages writes (record 0299).
//!
//! Scenarios, each a leaf/unified pair over the same bytes in one process:
//!
//! - `open`, `open_get`, `set`: the record 0296 shapes on
//!   `test-data/iwork/pages/basic.pages` (open; open and read the first
//!   paragraph; set the first paragraph's text and commit);
//! - `add`, `remove`, `move`: one structural paragraph operation on the
//!   same fixture (leaf: `Package::edit_section_paragraphs` with one
//!   `insert`/`remove`/`relocate`; unified: `File::{add, remove, move_to}`);
//! - `batch_N` for N in 1, 10, 100, 1000 and 10000: N paragraph text sets on
//!   a 10 000-paragraph document built with `litchi_pages::builder` (fixed
//!   identity). Leaf: one `edit_section_paragraphs` edit with N `set`s and
//!   one commit; unified: `File::edit` with N `tx.set("/body/p[i]", text)`
//!   calls, which coalesce into one leaf paragraph transaction.
//! - record 0314: `batch_1000_parsed` and `batch_10000_parsed`, the same
//!   batches with every path parsed once outside the clock (the unified leg
//!   passes `&TargetArg`).
//! - record 0309: `paragraph_font` (bold on the first paragraph of
//!   `basic.pages`; leaf `edit_section_text_style` + `set` of its span;
//!   unified `set` of `font.bold`) and `table_cell` (a text value in the
//!   first cell of `body-table-catalog-native.pages`; leaf
//!   `edit_body_table_values`; unified `set` of `text`). Their outputs are
//!   compared byte for byte before timing.
//! - record 0329: `table_add`, a 3 x 3 table at the end of
//!   `body-table-catalog-native.pages` (leaf: `edit_section_paragraphs`
//!   inserting an empty last paragraph, then `edit_body_table_creation().insert(..)` at its
//!   start; unified `File::add("/body", tbl)` with `rows` and `cols`);
//!   outputs compared byte for byte before timing.
//!
//! Inputs (opened packages and files) are prepared outside the timed
//! region; the unified legs include building the object tree on their first
//! read. Each pair runs in ABBA order for `--rounds` rounds of
//! `--iterations` timed calls per leg (batches of 1000 and more use a
//! quarter of the iterations, at least 3); every leg's per-call median is
//! printed as one JSON line and a summary line gives the median unified/leaf
//! ratio over the rounds. Pin the process with `taskset`.
//!
//! ```text
//! cargo run --release -p litchi-ops --features pages --example pages_bench \
//!     -- [--rounds 8] [--iterations 30] [--only SCENARIO]
//! ```

use std::hint::black_box;
use std::time::{Duration, Instant};

use litchi_core::Position;
use litchi_ops::{NewNode, Placement, props};
use litchi_pages::Package;
use litchi_pages::builder::{Builder, Identity};

fn test_data() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

fn median(mut samples: Vec<Duration>) -> Duration {
    samples.sort_unstable();
    samples[samples.len() / 2]
}

/// Times `iterations` calls of `call`, each on a fresh input from `input`
/// prepared outside the timed region; returns the per-call median.
fn leg<I>(iterations: usize, mut input: impl FnMut() -> I, mut call: impl FnMut(I)) -> Duration {
    let mut samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let value = input();
        let start = Instant::now();
        call(value);
        samples.push(start.elapsed());
    }
    median(samples)
}

fn argument(name: &str) -> Option<String> {
    let arguments: Vec<String> = std::env::args().collect();
    arguments
        .iter()
        .position(|argument| argument == name)
        .and_then(|index| arguments.get(index + 1))
        .cloned()
}

fn number(name: &str, default: usize) -> usize {
    argument(name)
        .and_then(|value| value.parse().ok())
        .unwrap_or(default)
}

/// One paired scenario: the leaf leg and the unified leg.
struct Scenario<'a> {
    name: String,
    iterations: usize,
    leaf: Box<dyn FnMut(usize) -> Duration + 'a>,
    unified: Box<dyn FnMut(usize) -> Duration + 'a>,
}

fn paragraph_edit(
    package: &Package,
    stage: impl FnOnce(&mut litchi_pages::section::paragraphs::Edit<'_>),
) {
    let mut edit = package.edit_section_paragraphs(0_usize).unwrap();
    stage(&mut edit);
    drop(black_box(edit.commit().unwrap()));
}

fn main() {
    let rounds = number("--rounds", 8);
    let iterations = number("--iterations", 30);
    let only = argument("--only");
    let basic = std::fs::read(test_data().join("iwork/pages/basic.pages")).unwrap();
    let tables =
        std::fs::read(test_data().join("iwork/pages/body-table-catalog-native.pages")).unwrap();
    let large_text = (0..10_000)
        .map(|index| format!("Paragraph {index}"))
        .collect::<Vec<_>>()
        .join("\n");
    let large = Builder::new()
        .identity(Identity::from_bytes([0x29; 16]))
        .body_text(large_text)
        .build()
        .unwrap()
        .to_bytes()
        .unwrap();

    let mut scenarios: Vec<Scenario<'_>> = Vec::new();
    let basic_ref = &basic;
    scenarios.push(Scenario {
        name: "open".to_owned(),
        iterations,
        leaf: Box::new(move |n| {
            leg(
                n,
                || (),
                |()| {
                    black_box(Package::from_bytes(basic_ref).unwrap());
                },
            )
        }),
        unified: Box::new(move |n| {
            leg(
                n,
                || basic_ref.clone(),
                |owned| {
                    black_box(litchi_ops::open_bytes(owned).unwrap());
                },
            )
        }),
    });
    scenarios.push(Scenario {
        name: "open_get".to_owned(),
        iterations,
        leaf: Box::new(move |n| {
            leg(
                n,
                || (),
                |()| {
                    let package = Package::from_bytes(basic_ref).unwrap();
                    black_box(package.section_paragraphs(0_usize).unwrap()[0].to_owned());
                },
            )
        }),
        unified: Box::new(move |n| {
            leg(
                n,
                || basic_ref.clone(),
                |owned| {
                    let file = litchi_ops::open_bytes(owned).unwrap();
                    let node = file.get("/body/p[1]").unwrap();
                    let node = node.as_iwork().unwrap();
                    black_box(node.text().unwrap().map(std::borrow::Cow::into_owned));
                },
            )
        }),
    });
    type Leaf = fn(&mut litchi_pages::section::paragraphs::Edit<'_>);
    type Unified = fn(&litchi_ops::File);
    let structural: [(&str, Leaf, Unified); 4] = [
        (
            "set",
            |edit| {
                edit.set(Position::new(0), "Benchmark paragraph").unwrap();
            },
            |file| {
                black_box(
                    file.set("/body/p[1]", props! { "text" => "Benchmark paragraph" })
                        .unwrap(),
                );
            },
        ),
        (
            "add",
            |edit| {
                edit.insert(Position::new(3), "Appended").unwrap();
            },
            |file| {
                black_box(
                    file.add("/body", NewNode::new("p").prop("text", "Appended"))
                        .unwrap(),
                );
            },
        ),
        (
            "remove",
            |edit| {
                edit.remove(Position::new(1)).unwrap();
            },
            |file| {
                black_box(file.remove("/body/p[2]").unwrap());
            },
        ),
        (
            "move",
            |edit| {
                edit.relocate(Position::new(0), Position::new(2)).unwrap();
            },
            |file| {
                black_box(file.move_to("/body/p[1]", Placement::End).unwrap());
            },
        ),
    ];
    for (name, leaf, unified) in structural {
        scenarios.push(Scenario {
            name: name.to_owned(),
            iterations,
            leaf: Box::new(move |n| {
                leg(
                    n,
                    || Package::from_bytes(basic_ref).unwrap(),
                    |package| paragraph_edit(&package, leaf),
                )
            }),
            unified: Box::new(move |n| {
                leg(
                    n,
                    || litchi_ops::open_bytes(basic_ref.clone()).unwrap(),
                    |file| unified(&file),
                )
            }),
        });
    }
    // Record 0309: character properties and table cell values.
    let first = Package::from_bytes(basic_ref)
        .unwrap()
        .section_paragraphs(0_usize)
        .unwrap()[0]
        .encode_utf16()
        .count();
    let leaf_font = move |package: &Package| {
        let mut bold = litchi_pages::section::text_style::CharacterProperties::new();
        bold.bold = Some(true);
        let mut edit = package.edit_section_text_style(0_usize).unwrap();
        edit.set(
            litchi_pages::TextSpan::from_utf16_indexes(0, first).unwrap(),
            bold,
        )
        .unwrap();
        edit.commit().unwrap().into_package()
    };
    let unified_font = |file: &litchi_ops::File| {
        file.set("/body/p[1]", props! { "font.bold" => true })
            .unwrap()
    };
    assert!(
        leaf_font(&Package::from_bytes(basic_ref).unwrap())
            .to_bytes()
            .unwrap()
            == unified_font(&litchi_ops::open_bytes(basic.clone()).unwrap())
                .to_bytes()
                .unwrap(),
        "paragraph_font outputs differ"
    );
    scenarios.push(Scenario {
        name: "paragraph_font".to_owned(),
        iterations,
        leaf: Box::new(move |n| {
            leg(
                n,
                || Package::from_bytes(basic_ref).unwrap(),
                |package| drop(black_box(leaf_font(&package))),
            )
        }),
        unified: Box::new(move |n| {
            leg(
                n,
                || litchi_ops::open_bytes(basic_ref.clone()).unwrap(),
                |file| drop(black_box(unified_font(&file))),
            )
        }),
    });
    let tables_ref = &tables;
    let leaf_cell = |package: &Package| {
        let mut edit = package
            .edit_body_table_values(litchi_pages::BodyTableSelector::index(0))
            .unwrap();
        edit.set(
            0,
            0,
            litchi_pages::table::values::Value::Text("Bench".into()),
        )
        .unwrap();
        edit.commit().unwrap().into_package()
    };
    let unified_cell = |file: &litchi_ops::File| {
        file.set("/body/tbl[1]/tr[1]/tc[1]", props! { "text" => "Bench" })
            .unwrap()
    };
    assert!(
        leaf_cell(&Package::from_bytes(tables_ref).unwrap())
            .to_bytes()
            .unwrap()
            == unified_cell(&litchi_ops::open_bytes(tables.clone()).unwrap())
                .to_bytes()
                .unwrap(),
        "table_cell outputs differ"
    );
    scenarios.push(Scenario {
        name: "table_cell".to_owned(),
        iterations,
        leaf: Box::new(move |n| {
            leg(
                n,
                || Package::from_bytes(tables_ref).unwrap(),
                |package| drop(black_box(leaf_cell(&package))),
            )
        }),
        unified: Box::new(move |n| {
            leg(
                n,
                || litchi_ops::open_bytes(tables_ref.clone()).unwrap(),
                |file| drop(black_box(unified_cell(&file))),
            )
        }),
    });
    // Record 0329: a new table in a new last paragraph.
    let table_anchor = {
        let package = Package::from_bytes(tables_ref).unwrap();
        let sections = package.sections().len();
        let total: usize = (0..sections)
            .map(|index| package.section_text(index).unwrap().encode_utf16().count())
            .sum::<usize>()
            + sections
            - 1;
        let lines = package
            .section_text(sections - 1)
            .unwrap()
            .split('\n')
            .count();
        (sections - 1, lines, total + 1)
    };
    let leaf_table = move |package: &Package| {
        let (section, line, anchor) = table_anchor;
        let mut edit = package.edit_section_paragraphs(section).unwrap();
        edit.insert(Position::new(line), "").unwrap();
        let with_paragraph = edit.commit().unwrap().into_package();
        with_paragraph
            .edit_body_table_creation()
            .insert(
                Position::new(anchor),
                litchi_pages::table::creation::NewTable::new(3, 3),
            )
            .and_then(|edit| edit.commit())
            .unwrap()
            .into_package()
    };
    let unified_table = |file: &litchi_ops::File| {
        file.add("/body", NewNode::new("tbl").prop("rows", 3).prop("cols", 3))
            .unwrap()
            .file
    };
    assert!(
        leaf_table(&Package::from_bytes(tables_ref).unwrap())
            .to_bytes()
            .unwrap()
            == unified_table(&litchi_ops::open_bytes(tables.clone()).unwrap())
                .to_bytes()
                .unwrap(),
        "table_add outputs differ"
    );
    scenarios.push(Scenario {
        name: "table_add".to_owned(),
        iterations,
        leaf: Box::new(move |n| {
            leg(
                n,
                || Package::from_bytes(tables_ref).unwrap(),
                |package| drop(black_box(leaf_table(&package))),
            )
        }),
        unified: Box::new(move |n| {
            leg(
                n,
                || litchi_ops::open_bytes(tables_ref.clone()).unwrap(),
                |file| drop(black_box(unified_table(&file))),
            )
        }),
    });
    let large_ref = &large;
    for count in [1_usize, 10, 100, 1_000, 10_000] {
        let paths: Vec<String> = (1..=count)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        let texts: Vec<String> = (0..count).map(|index| format!("Edited {index}")).collect();
        let (paths, texts) = (std::rc::Rc::new(paths), std::rc::Rc::new(texts));
        let (leaf_texts, unified_texts, unified_paths) =
            (texts.clone(), texts.clone(), paths.clone());
        scenarios.push(Scenario {
            name: format!("batch_{count}"),
            iterations: if count >= 1_000 {
                (iterations / 4).max(3)
            } else {
                iterations
            },
            leaf: Box::new(move |n| {
                let texts = leaf_texts.clone();
                leg(
                    n,
                    || Package::from_bytes(large_ref).unwrap(),
                    |package| {
                        paragraph_edit(&package, |edit| {
                            for (index, text) in texts.iter().enumerate() {
                                edit.set(Position::new(index), text).unwrap();
                            }
                        });
                    },
                )
            }),
            unified: Box::new(move |n| {
                let (paths, texts) = (unified_paths.clone(), unified_texts.clone());
                leg(
                    n,
                    || litchi_ops::open_bytes(large_ref.clone()).unwrap(),
                    |file| {
                        let commit = file
                            .edit(|tx| {
                                for (path, text) in paths.iter().zip(texts.iter()) {
                                    tx.set(path.as_str(), props! { "text" => text.as_str() })?;
                                }
                                Ok(())
                            })
                            .unwrap();
                        assert_eq!(commit.diagnostics().commits(), 1);
                        black_box(commit);
                    },
                )
            }),
        });
    }
    // Record 0314: the same batches with each path parsed once, outside the
    // clock (`TargetArg::parse`; `&TargetArg` reruns without parsing).
    for count in [1_000_usize, 10_000] {
        let targets: Vec<litchi_ops::TargetArg<'static>> = (1..=count)
            .map(|index| {
                let text: &'static str = Box::leak(format!("/body/p[{index}]").into_boxed_str());
                litchi_ops::TargetArg::parse(text).unwrap()
            })
            .collect();
        let texts: Vec<String> = (0..count).map(|index| format!("Edited {index}")).collect();
        let (targets, texts) = (std::rc::Rc::new(targets), std::rc::Rc::new(texts));
        let (leaf_texts, unified_texts) = (texts.clone(), texts.clone());
        scenarios.push(Scenario {
            name: format!("batch_{count}_parsed"),
            iterations: (iterations / 4).max(3),
            leaf: Box::new(move |n| {
                let texts = leaf_texts.clone();
                leg(
                    n,
                    || Package::from_bytes(large_ref).unwrap(),
                    |package| {
                        paragraph_edit(&package, |edit| {
                            for (index, text) in texts.iter().enumerate() {
                                edit.set(Position::new(index), text).unwrap();
                            }
                        });
                    },
                )
            }),
            unified: Box::new(move |n| {
                let (targets, texts) = (targets.clone(), unified_texts.clone());
                leg(
                    n,
                    || litchi_ops::open_bytes(large_ref.clone()).unwrap(),
                    |file| {
                        let commit = file
                            .edit(|tx| {
                                for (target, text) in targets.iter().zip(texts.iter()) {
                                    tx.set(target, props! { "text" => text.as_str() })?;
                                }
                                Ok(())
                            })
                            .unwrap();
                        assert_eq!(commit.diagnostics().commits(), 1);
                        black_box(commit);
                    },
                )
            }),
        });
    }

    for scenario in &mut scenarios {
        if only.as_deref().is_some_and(|only| only != scenario.name) {
            continue;
        }
        let mut ratios = Vec::with_capacity(rounds);
        let mut leaf_medians = Vec::new();
        let mut unified_medians = Vec::new();
        for round in 0..rounds {
            // ABBA: leaf, unified, unified, leaf.
            for (position, leg_name) in ["leaf", "unified", "unified", "leaf"].iter().enumerate() {
                let median = if *leg_name == "leaf" {
                    (scenario.leaf)(scenario.iterations)
                } else {
                    (scenario.unified)(scenario.iterations)
                };
                println!(
                    "{{\"format\":\"pages\",\"scenario\":\"{}\",\"round\":{round},\"position\":{position},\"leg\":\"{leg_name}\",\"median_ns\":{}}}",
                    scenario.name,
                    median.as_nanos()
                );
                if *leg_name == "leaf" {
                    leaf_medians.push(median);
                } else {
                    unified_medians.push(median);
                }
            }
            let leaf = leaf_medians[leaf_medians.len() - 2..]
                .iter()
                .sum::<Duration>()
                .as_secs_f64();
            let unified = unified_medians[unified_medians.len() - 2..]
                .iter()
                .sum::<Duration>()
                .as_secs_f64();
            ratios.push(unified / leaf);
        }
        ratios.sort_by(f64::total_cmp);
        let leaf = median(leaf_medians);
        let unified = median(unified_medians);
        println!(
            "{{\"summary\":true,\"format\":\"pages\",\"scenario\":\"{}\",\"leaf_median_ns\":{},\"unified_median_ns\":{},\"ratio_median\":{:.4},\"ratio_min\":{:.4},\"ratio_max\":{:.4},\"delta_ns\":{}}}",
            scenario.name,
            leaf.as_nanos(),
            unified.as_nanos(),
            ratios[ratios.len() / 2],
            ratios[0],
            ratios[ratios.len() - 1],
            unified.as_nanos() as i128 - leaf.as_nanos() as i128,
        );
    }
}
