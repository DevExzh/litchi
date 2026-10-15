//! Paired leaf-vs-unified measurements of Keynote reads and writes (record
//! 0298), extending `iwork_bench` (record 0296) for Keynote.
//!
//! Scenarios, each a leaf/unified pair over
//! `test-data/iwork/keynote/multi-slide.key` (five slides) in one process:
//!
//! - `open`: open only;
//! - `open_get`: open and read the second slide's title;
//! - `set`: set the first slide's title (leaf: `edit_slide_text` + `set`);
//! - `notes`: set the first slide's speaker notes (leaf:
//!   `edit_slide_notes` + `set`);
//! - `frame`: set the first slide's text box `x` (leaf:
//!   `edit_slide_drawables` + `set_geometry`);
//! - `batch_1000`: 1 000 sets in one `File::edit`, alternating the three
//!   body paragraphs of slide 2 and the `x` of slide 1's text box; the
//!   unified run coalesces them into one slide-drawable commit and one
//!   slide-text commit. The leaf leg hand-writes the same leaf work: one
//!   `edit_slide_text` replacing the changed span with the final body text
//!   and one `edit_slide_drawables` with the final frame;
//! - `duplicate`, `add_slide`, `add_textbox`, `remove_slide`: one
//!   structural operation (leaf: `edit_slide_creation` with
//!   `duplicate_slide` / `add_slide` from the default layout,
//!   `edit_slide_drawables` with `add_text_box`, `edit_slide_deletion`;
//!   unified: `File::add` / `File::remove`).
//!
//! Inputs (opened packages and files) are prepared outside the timed
//! region; the unified legs include building the object tree on their first
//! read. Each pair runs in ABBA order for `--rounds` rounds of
//! `--iterations` timed calls per leg; every leg's per-call median is
//! printed as one JSON line and a summary line gives the median
//! unified/leaf ratio over the rounds. Pin the process with `taskset`.
//!
//! ```text
//! cargo run --release -p litchi-ops --features keynote --example keynote_bench \
//!     -- [--rounds 8] [--iterations 30] [--only SCENARIO]
//! ```

use std::hint::black_box;
use std::time::{Duration, Instant};

use litchi_keynote::slide::drawable::Geometry;
use litchi_keynote::slide::placeholder::Kind as Role;
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector, TextPosition, TextSpan};
use litchi_ops::{NewNode, props};

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

/// The smallest UTF-16 span of `before` whose replacement gives `after`.
fn span<'a>(before: &str, after: &'a str) -> (TextSpan, &'a str) {
    let prefix = before
        .char_indices()
        .zip(after.chars())
        .find(|((_, left), right)| left != right)
        .map_or_else(|| before.len().min(after.len()), |((index, _), _)| index);
    let (before_rest, after_rest) = (&before[prefix..], &after[prefix..]);
    let suffix = before_rest
        .chars()
        .rev()
        .zip(after_rest.chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(left, _)| left.len_utf8())
        .sum::<usize>();
    let removed = &before_rest[..before_rest.len() - suffix];
    let start = before[..prefix].encode_utf16().count();
    let end = start + removed.encode_utf16().count();
    let position = |units| TextPosition::from_utf16_index(units).unwrap();
    (
        TextSpan::new(position(start), position(end)).unwrap(),
        &after_rest[..after_rest.len() - suffix],
    )
}

/// One paired scenario: the leaf leg and the unified leg.
struct Scenario<'a> {
    name: &'static str,
    leaf: Box<dyn FnMut(usize) -> Duration + 'a>,
    unified: Box<dyn FnMut(usize) -> Duration + 'a>,
}

const OPERATIONS: usize = 1000;

#[allow(clippy::too_many_lines, reason = "one scenario table")]
fn main() {
    let rounds = number("--rounds", 8);
    let iterations = number("--iterations", 30);
    let only = argument("--only");
    let bytes = std::fs::read(test_data().join("iwork/keynote/multi-slide.key")).unwrap();
    let bytes = &bytes;
    let package = || Package::from_bytes(bytes).unwrap();
    let file = || litchi_ops::open_bytes(bytes.clone()).unwrap();

    // The batch: rounds of the three body paragraphs of slide 2 and the
    // frame of slide 1's text box.
    let body = "/slide[2]/placeholder[body]";
    let mut batch = Vec::with_capacity(OPERATIONS);
    for index in 0..OPERATIONS {
        if index % 4 == 3 {
            batch.push((
                "/slide[1]/shape[1]".to_owned(),
                "x",
                format!("{}pt", index / 4),
            ));
        } else {
            let line = index % 4 + 1;
            batch.push((
                format!("{body}/p[{line}]"),
                "text",
                format!("op {index} topic {line}"),
            ));
        }
    }
    let batch = &batch;
    let final_body = (1..=3)
        .map(|line| {
            batch
                .iter()
                .rev()
                .find(|(path, _, _)| path.ends_with(&format!("p[{line}]")))
                .map(|(_, _, text)| text.clone())
                .unwrap()
        })
        .collect::<Vec<_>>()
        .join("\n");
    let final_x = batch
        .iter()
        .rev()
        .find(|(_, key, _)| *key == "x")
        .map(|(_, _, value)| value.trim_end_matches("pt").parse::<f32>().unwrap())
        .unwrap();
    let final_body = &final_body;
    let frame = package()
        .slide_drawable_list(SlideSelector::index(0))
        .unwrap()[0]
        .geometry()
        .unwrap();

    let scenarios: Vec<Scenario<'_>> = vec![
        Scenario {
            name: "open",
            leaf: Box::new(move |n| {
                leg(
                    n,
                    || (),
                    |()| {
                        black_box(Package::from_bytes(bytes).unwrap());
                    },
                )
            }),
            unified: Box::new(move |n| {
                leg(
                    n,
                    || bytes.clone(),
                    |owned| {
                        black_box(litchi_ops::open_bytes(owned).unwrap());
                    },
                )
            }),
        },
        Scenario {
            name: "open_get",
            leaf: Box::new(move |n| {
                leg(
                    n,
                    || (),
                    |()| {
                        let package = Package::from_bytes(bytes).unwrap();
                        black_box(package.slide_title(SlideSelector::index(1)).unwrap());
                    },
                )
            }),
            unified: Box::new(move |n| {
                leg(
                    n,
                    || bytes.clone(),
                    |owned| {
                        let file = litchi_ops::open_bytes(owned).unwrap();
                        let node = file.get("/slide[2]/placeholder[title]").unwrap();
                        let node = node.as_iwork().unwrap();
                        black_box(node.text().unwrap().map(std::borrow::Cow::into_owned));
                    },
                )
            }),
        },
        Scenario {
            name: "set",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package
                        .edit_slide_text(SlideSelector::index(0), Role::Title)
                        .unwrap();
                    edit.set("Benchmark title").unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.set(
                            "/slide[1]/placeholder[title]",
                            props! { "text" => "Benchmark title" },
                        )
                        .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "notes",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_notes(SlideSelector::index(0)).unwrap();
                    edit.set("Benchmark notes").unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.set("/slide[1]/notes", props! { "text" => "Benchmark notes" })
                            .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "frame",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.set_geometry(
                        SlideSelector::index(0),
                        DrawableSelector::index(0),
                        Geometry { x: 72.0, ..frame },
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.set("/slide[1]/shape[1]", props! { "x" => "72pt" })
                            .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "batch_1000",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let commit = {
                        let mut edit = package
                            .edit_slide_text(SlideSelector::index(1), Role::Body)
                            .unwrap();
                        let (span, replacement) = span(edit.text(), final_body);
                        edit.replace(span, replacement).unwrap();
                        edit.commit().unwrap()
                    };
                    let package = commit.into_package();
                    let mut edit = package.edit_slide_drawables();
                    edit.set_geometry(
                        SlideSelector::index(0),
                        DrawableSelector::index(0),
                        Geometry {
                            x: final_x,
                            ..frame
                        },
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    let commit = file
                        .edit(|tx| {
                            for (path, key, value) in batch {
                                tx.set(path.as_str(), props! { *key => value.as_str() })?;
                            }
                            Ok(())
                        })
                        .unwrap();
                    assert_eq!(commit.diagnostics.commits(), 1);
                    black_box(commit);
                })
            }),
        },
        Scenario {
            name: "duplicate",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_creation();
                    edit.duplicate_slide(SlideSelector::index(1), Position::new(2))
                        .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.add(
                            "/",
                            NewNode::new("slide")
                                .prop("from", "/slide[2]")
                                .at(litchi_ops::Placement::after("/slide[2]").unwrap()),
                        )
                        .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "add_slide",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let layout = package.default_slide_layout().unwrap().unwrap();
                    let mut edit = package.edit_slide_creation();
                    edit.add_slide(&layout, Position::new(5)).unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(file.add("/", NewNode::new("slide")).unwrap());
                })
            }),
        },
        Scenario {
            name: "add_textbox",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.add_text_box(
                        SlideSelector::index(0),
                        Geometry::new(100.0, 100.0, 400.0, 100.0),
                        "Box",
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.add("/slide[1]", NewNode::new("textbox").prop("text", "Box"))
                            .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "remove_slide",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_deletion();
                    edit.remove_slide(SlideSelector::index(2)).unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(file.remove("/slide[3]").unwrap());
                })
            }),
        },
    ];

    for mut scenario in scenarios {
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
                    (scenario.leaf)(iterations)
                } else {
                    (scenario.unified)(iterations)
                };
                println!(
                    "{{\"format\":\"keynote\",\"scenario\":\"{}\",\"round\":{round},\"position\":{position},\"leg\":\"{leg_name}\",\"median_ns\":{}}}",
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
            "{{\"summary\":true,\"format\":\"keynote\",\"scenario\":\"{}\",\"leaf_median_ns\":{},\"unified_median_ns\":{},\"ratio_median\":{:.4},\"ratio_min\":{:.4},\"ratio_max\":{:.4},\"delta_ns\":{}}}",
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
