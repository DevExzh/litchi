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
//! - record 0314: `batch_1000_parsed`, the same batch with every path
//!   parsed once outside the clock (the unified leg passes `&TargetArg`);
//! - `duplicate`, `add_slide`, `add_textbox`, `remove_slide`: one
//!   structural operation (leaf: `edit_slide_creation` with
//!   `duplicate_slide` / `add_slide` from the default layout,
//!   `edit_slide_drawables` with `add_text_box`, `edit_slide_deletion`;
//!   unified: `File::add` / `File::remove`).
//! - record 0309: `run_style` (bold and 24 pt on the first run of slide 1's
//!   text box; leaf `edit_slide_drawables` + `set_text_style`),
//!   `placeholder_hide` (slide 2's body placeholder; leaf
//!   `edit_slide_placeholder_visibility` + `hide`; unified `File::remove`),
//!   `table_cell` (the first cell of `table-discovery.key`; leaf
//!   `set_table_cell`; unified `set` of `text`) and `media_remove` (the
//!   first clip of `media-lifecycle-audio-remove-native.key`; leaf
//!   `edit_slide_media_lifecycle().remove(..)`; unified `File::remove`). Their outputs are
//!   compared byte for byte before timing.
//! - record 0329: `table_add`, a 3 x 3 table on the table slide of
//!   `table-discovery.key` (leaf `edit_slide_drawables` + `add_table`;
//!   unified `File::add` of `tbl` with `rows` and `cols`); outputs compared
//!   byte for byte before timing; `notes_run_style` (bold on the first
//!   run of slide 2's speaker notes, written outside the clock; leaf
//!   `edit_slide_drawables` + `set_notes_text_style`; unified `set` of the
//!   notes run's `bold`).
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

use litchi_keynote::slide::drawable::{
    Geometry, Kind as DrawableKind, NewTable, Style, TableCellValue, TextSpan as DrawableSpan,
};
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
    let parsed: Vec<(litchi_ops::TargetArg<'_>, &str, String)> = batch
        .iter()
        .map(|(path, key, value)| {
            (
                litchi_ops::TargetArg::parse(path.as_str()).unwrap(),
                *key,
                value.clone(),
            )
        })
        .collect();
    let parsed = &parsed;
    let frame = package()
        .slide_drawable_list(SlideSelector::index(0))
        .unwrap()[0]
        .geometry()
        .unwrap();

    // Record 0309 inputs; each pair's outputs are compared once here.
    const RUN: &str = "/slide[1]/shape[1]/p[1]/run[1]";
    let run_length = file()
        .get(RUN)
        .unwrap()
        .as_iwork()
        .unwrap()
        .text()
        .unwrap()
        .map_or(0, |text| text.chars().count());
    let leaf_bytes = |package: &Package| {
        let mut out = Vec::new();
        package.write_to(&mut out).unwrap();
        out
    };
    {
        let source = package();
        let mut edit = source.edit_slide_drawables();
        edit.set_text_style(
            SlideSelector::index(0),
            DrawableSelector::index(0),
            DrawableSpan::new(0, run_length),
            Style::new().with_bold(true).with_font_size(24.0),
        )
        .unwrap();
        let leaf = leaf_bytes(edit.commit().unwrap().package());
        let unified = file()
            .set(RUN, props! { "bold" => true, "size" => "24pt" })
            .unwrap()
            .to_bytes()
            .unwrap();
        assert!(leaf == unified, "run_style outputs differ");
        let leaf = leaf_bytes(
            package()
                .edit_slide_placeholder_visibility(SlideSelector::index(1), Role::Body)
                .unwrap()
                .hide()
                .commit()
                .unwrap()
                .package(),
        );
        assert!(
            leaf == file().remove(body).unwrap().to_bytes().unwrap(),
            "placeholder_hide outputs differ"
        );
    }
    let table_bytes = std::fs::read(test_data().join("iwork/keynote/table-discovery.key")).unwrap();
    let table_bytes = &table_bytes;
    let table_package = || Package::from_bytes(table_bytes).unwrap();
    let table_file = || litchi_ops::open_bytes(table_bytes.clone()).unwrap();
    let cell = table_file()
        .query("tc")
        .unwrap()
        .next()
        .unwrap()
        .unwrap()
        .path()
        .to_string();
    let cell = &*Box::leak(cell.into_boxed_str());
    let slide_of = |path: &str| -> usize {
        let start = path.find('[').unwrap() + 1;
        let end = path.find(']').unwrap();
        path[start..end].parse::<usize>().unwrap() - 1
    };
    let table_slide = slide_of(cell);
    let table_drawable = table_package()
        .slide_drawable_list(SlideSelector::index(table_slide))
        .unwrap()
        .iter()
        .position(|drawable| drawable.kind() == DrawableKind::Table)
        .unwrap();
    {
        let source = table_package();
        let mut edit = source.edit_slide_drawables();
        edit.set_table_cell(
            SlideSelector::index(table_slide),
            DrawableSelector::index(table_drawable),
            0,
            0,
            TableCellValue::Text("Bench".into()),
        )
        .unwrap();
        let leaf = leaf_bytes(edit.commit().unwrap().package());
        let unified = table_file()
            .set(cell, props! { "text" => "Bench" })
            .unwrap()
            .to_bytes()
            .unwrap();
        assert!(leaf == unified, "table_cell outputs differ");
    }
    // Record 0329: the unified default frame of a new table.
    let new_table = || NewTable::new(3, 3, Geometry::new(100.0, 100.0, 600.0, 200.0));
    let table_slide_path = &*Box::leak(format!("/slide[{}]", table_slide + 1).into_boxed_str());
    {
        let source = table_package();
        let mut edit = source.edit_slide_drawables();
        edit.add_table(SlideSelector::index(table_slide), new_table())
            .unwrap();
        let leaf = leaf_bytes(edit.commit().unwrap().package());
        let unified = table_file()
            .add(
                table_slide_path,
                NewNode::new("tbl").prop("rows", 3).prop("cols", 3),
            )
            .unwrap()
            .file
            .to_bytes()
            .unwrap();
        assert!(leaf == unified, "table_add outputs differ");
    }
    // Record 0329: speaker notes on slide 2, written outside the clock.
    let notes_bytes = {
        let source = package();
        let mut edit = source.edit_slide_notes(SlideSelector::index(1)).unwrap();
        edit.set("Remember the second point").unwrap();
        leaf_bytes(edit.commit().unwrap().package())
    };
    let notes_bytes = &notes_bytes;
    let notes_package = || Package::from_bytes(notes_bytes).unwrap();
    let notes_file = || litchi_ops::open_bytes(notes_bytes.clone()).unwrap();
    const NOTES_RUN: &str = "/slide[2]/notes/p[1]/run[1]";
    let notes_length = "Remember the second point".chars().count();
    {
        let source = notes_package();
        let mut edit = source.edit_slide_drawables();
        edit.set_notes_text_style(
            SlideSelector::index(1),
            DrawableSpan::new(0, notes_length),
            Style::new().with_bold(true),
        )
        .unwrap();
        let leaf = leaf_bytes(edit.commit().unwrap().package());
        let unified = notes_file()
            .set(NOTES_RUN, props! { "bold" => true })
            .unwrap()
            .to_bytes()
            .unwrap();
        assert!(leaf == unified, "notes_run_style outputs differ");
    }
    let media_bytes =
        std::fs::read(test_data().join("iwork/keynote/media-lifecycle-audio-remove-native.key"))
            .unwrap();
    let media_bytes = &media_bytes;
    let media_package = || Package::from_bytes(media_bytes).unwrap();
    let media_file = || litchi_ops::open_bytes(media_bytes.clone()).unwrap();
    let video = media_file()
        .query("video")
        .unwrap()
        .next()
        .unwrap()
        .unwrap()
        .path()
        .to_string();
    let video = &*Box::leak(video.into_boxed_str());
    let media_slide = slide_of(video);
    {
        let leaf = leaf_bytes(
            media_package()
                .edit_slide_media_lifecycle()
                .remove(
                    SlideSelector::index(media_slide),
                    litchi_keynote::MovieSelector::index(0),
                )
                .and_then(|edit| edit.commit())
                .unwrap()
                .package(),
        );
        assert!(
            leaf == media_file().remove(video).unwrap().to_bytes().unwrap(),
            "media_remove outputs differ"
        );
    }

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
                        let (span, replacement) = span(edit.text().unwrap(), final_body);
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
                    assert_eq!(commit.diagnostics().commits(), 1);
                    black_box(commit);
                })
            }),
        },
        // Record 0314: the same batch with each path parsed once, outside
        // the clock (`&TargetArg` reruns without parsing).
        Scenario {
            name: "batch_1000_parsed",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let commit = {
                        let mut edit = package
                            .edit_slide_text(SlideSelector::index(1), Role::Body)
                            .unwrap();
                        let (span, replacement) = span(edit.text().unwrap(), final_body);
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
                            for ((target, key, value), _) in parsed.iter().zip(batch) {
                                tx.set(target, props! { *key => value.as_str() })?;
                            }
                            Ok(())
                        })
                        .unwrap();
                    assert_eq!(commit.diagnostics().commits(), 1);
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
        Scenario {
            name: "run_style",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.set_text_style(
                        SlideSelector::index(0),
                        DrawableSelector::index(0),
                        DrawableSpan::new(0, run_length),
                        Style::new().with_bold(true).with_font_size(24.0),
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(
                        file.set(RUN, props! { "bold" => true, "size" => "24pt" })
                            .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "placeholder_hide",
            leaf: Box::new(move |n| {
                leg(n, package, |package| {
                    let edit = package
                        .edit_slide_placeholder_visibility(SlideSelector::index(1), Role::Body)
                        .unwrap();
                    drop(black_box(edit.hide().commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, file, |file| {
                    black_box(file.remove(body).unwrap());
                })
            }),
        },
        Scenario {
            name: "table_cell",
            leaf: Box::new(move |n| {
                leg(n, table_package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.set_table_cell(
                        SlideSelector::index(table_slide),
                        DrawableSelector::index(table_drawable),
                        0,
                        0,
                        TableCellValue::Text("Bench".into()),
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, table_file, |file| {
                    black_box(file.set(cell, props! { "text" => "Bench" }).unwrap());
                })
            }),
        },
        Scenario {
            name: "table_add",
            leaf: Box::new(move |n| {
                leg(n, table_package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.add_table(SlideSelector::index(table_slide), new_table())
                        .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, table_file, |file| {
                    black_box(
                        file.add(
                            table_slide_path,
                            NewNode::new("tbl").prop("rows", 3).prop("cols", 3),
                        )
                        .unwrap(),
                    );
                })
            }),
        },
        Scenario {
            name: "notes_run_style",
            leaf: Box::new(move |n| {
                leg(n, notes_package, |package| {
                    let mut edit = package.edit_slide_drawables();
                    edit.set_notes_text_style(
                        SlideSelector::index(1),
                        DrawableSpan::new(0, notes_length),
                        Style::new().with_bold(true),
                    )
                    .unwrap();
                    drop(black_box(edit.commit().unwrap()));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, notes_file, |file| {
                    black_box(file.set(NOTES_RUN, props! { "bold" => true }).unwrap());
                })
            }),
        },
        Scenario {
            name: "media_remove",
            leaf: Box::new(move |n| {
                leg(n, media_package, |package| {
                    drop(black_box(
                        package
                            .edit_slide_media_lifecycle()
                            .remove(
                                SlideSelector::index(media_slide),
                                litchi_keynote::MovieSelector::index(0),
                            )
                            .and_then(|edit| edit.commit())
                            .unwrap(),
                    ));
                })
            }),
            unified: Box::new(move |n| {
                leg(n, media_file, |file| {
                    black_box(file.remove(video).unwrap());
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
