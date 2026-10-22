//! Paired leaf-vs-unified measurements of the PPTX P0 residue (record
//! 0333).
//!
//! Inputs are made here: one-slide decks of N text boxes (N = 1, 10, 100,
//! 1,000, 10,000) for the batch scenarios, and a deck with a text box, a
//! 3 × 3 table, a picture and two comments for the single operations. Each
//! scenario times the leaf call and the unified call doing the same work on
//! the same owner, ABBA within one process (leaf, unified, unified, leaf per
//! round pair), and checks outside the clock that both legs publish the same
//! bytes:
//!
//! - `run_format_batch/<owner>/<N>`: one atomic `File::batch` of N
//!   `set /slide[1]/shape[i]/paragraph[1]/run[1] {bold, size, color}`
//!   against one native edit staging the same N `TextAssignment`s in 256-edit
//!   chunks through one `set_text_property_chunks` call and one commit (the
//!   1k-op batch budget: within 2×);
//! - `shape_text_batch/<owner>/<N>`: N `set /slide[1]/shape[i] {text}`
//!   against `set_shape_text_chunks`;
//! - single operations on the eager owner (budget 10%): `paragraph_add`,
//!   `run_add`, `paragraph_remove`, `row_add`, `row_remove`, `cell_fill`
//!   (`accent1`), `comment_remove`, `picture_frame_set`.
//!
//! `P0_SCENARIO=<prefix>` runs the scenarios whose name starts with it;
//! `P0_MAX=<N>` caps the batch sizes. Each prints one JSON line. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example pptx_p0_bench
//! --features pptx`.

#[cfg(feature = "pptx")]
#[allow(
    clippy::too_many_lines,
    reason = "one measurement script, one scenario after another"
)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::sync::Arc;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::{BatchItem, BatchMode, NewNode, OpenOptions, Props, SourceMode, props};
    use litchi_pptx::opened::{
        CellFill, Rgb, ShapeTextReplacement, TextAssignment, TextBody, TextProperties, TextTarget,
    };

    const WARMUP: usize = 2;

    let only = std::env::var("P0_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| name.starts_with(only));
    let max: usize = std::env::var("P0_MAX")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(10_000);

    fn stats(mut samples: Vec<f64>) -> (f64, f64) {
        samples.sort_by(f64::total_cmp);
        (samples[samples.len() / 2], samples[0])
    }

    /// Microseconds per call of each leg, ABBA over `rounds` round pairs.
    fn paired(
        rounds: usize,
        mut leaf: impl FnMut(),
        mut unified: impl FnMut(),
    ) -> (Vec<f64>, Vec<f64>) {
        let time = |run: &mut dyn FnMut()| {
            let started = Instant::now();
            run();
            started.elapsed().as_secs_f64() * 1e6
        };
        let (mut a, mut b) = (Vec::new(), Vec::new());
        for round in 0..WARMUP + rounds {
            let first = time(&mut leaf);
            let second = time(&mut unified);
            let third = time(&mut unified);
            let fourth = time(&mut leaf);
            if round >= WARMUP {
                a.extend([first, fourth]);
                b.extend([second, third]);
            }
        }
        (a, b)
    }

    fn report(name: &str, same: bool, (leaf, unified): (Vec<f64>, Vec<f64>)) {
        let (leaf_median, leaf_min) = stats(leaf);
        let (unified_median, unified_min) = stats(unified);
        println!(
            "{}",
            serde_json::json!({
                "scenario": name,
                "leaf_median_us": leaf_median,
                "leaf_min_us": leaf_min,
                "unified_median_us": unified_median,
                "unified_min_us": unified_min,
                "ratio": unified_median / leaf_median,
                "same_bytes": same,
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

    fn set_item(path: String, props: Props<'static>) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(props);
        item
    }

    let format = TextProperties::default()
        .with_bold(Some(true))
        .with_size(Some(2_000))
        .with_color(Some(Rgb([0x1F, 0x4E, 0x79])));
    let rounds_for = |n: usize| match n {
        0..=100 => 10,
        101..=1_000 => 6,
        _ => 3,
    };

    for n in [1_usize, 10, 100, 1_000, 10_000]
        .into_iter()
        .filter(|n| *n <= max)
    {
        let bytes = deck(n)?;
        for eager in [true, false] {
            let owner = if eager { "eager" } else { "source" };
            let file = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(if eager {
                    SourceMode::Eager
                } else {
                    SourceMode::Auto
                }),
            )?;
            let eager_package = litchi_pptx::Package::from_bytes(&bytes)?;
            let source_package = litchi_pptx::source::Package::from_read_at(Arc::new(
                litchi_core::OwnedSource::new(bytes.clone()),
            ))?;

            let name = format!("run_format_batch/{owner}/{n}");
            if wanted(&name) {
                let edits: Vec<TextAssignment<'_>> = (0..n)
                    .map(|index| {
                        TextAssignment::new(
                            TextTarget::Run {
                                shape: Position::new(index),
                                paragraph: Position::new(0),
                                run: Position::new(0),
                            },
                            format,
                        )
                    })
                    .collect();
                let items: Vec<BatchItem<'static>> = (1..=n)
                    .map(|index| {
                        set_item(
                            format!("/slide[1]/shape[{index}]/paragraph[1]/run[1]"),
                            Props::new()
                                .with("bold", true)
                                .with("size", "20pt")
                                .with("color", "1F4E79"),
                        )
                    })
                    .collect();
                let leaf = || -> Vec<u8> {
                    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(256).collect();
                    if eager {
                        let mut edit = eager_package.edit().expect("edit");
                        edit.set_text_property_chunks(0_usize, &chunks)
                            .expect("stage");
                        let mut package = edit.commit().expect("commit").into_package();
                        black_box(&mut package);
                        Vec::new()
                    } else {
                        let mut edit = source_package.edit().expect("edit");
                        edit.set_text_property_chunks(Position::new(0), &chunks)
                            .expect("stage");
                        drop(black_box(edit.commit().expect("commit")));
                        Vec::new()
                    }
                };
                let unified = || {
                    file.batch(items.clone(), BatchMode::Atomic)
                        .expect("batch")
                        .into_file()
                };
                let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(256).collect();
                let leaf_bytes = if eager {
                    let mut edit = eager_package.edit()?;
                    edit.set_text_property_chunks(0_usize, &chunks)?;
                    edit.commit()?.into_package().to_bytes()?
                } else {
                    let mut edit = source_package.edit()?;
                    edit.set_text_property_chunks(Position::new(0), &chunks)?;
                    edit.commit()?.into_package().to_bytes()?
                };
                let same = leaf_bytes == unified().to_bytes()?;
                report(
                    &name,
                    same,
                    paired(
                        rounds_for(n),
                        || drop(black_box(leaf())),
                        || drop(black_box(unified())),
                    ),
                );
            }

            let name = format!("shape_text_batch/{owner}/{n}");
            if wanted(&name) {
                let texts: Vec<String> = (0..n).map(|index| format!("batch-{index}")).collect();
                let replacements: Vec<ShapeTextReplacement<'_>> = texts
                    .iter()
                    .enumerate()
                    .map(|(index, text)| ShapeTextReplacement::new(index.into(), text))
                    .collect();
                let chunks: Vec<&[ShapeTextReplacement<'_>]> = replacements.chunks(256).collect();
                let items: Vec<BatchItem<'static>> = (1..=n)
                    .map(|index| {
                        set_item(
                            format!("/slide[1]/shape[{index}]"),
                            Props::new().with("text", format!("batch-{}", index - 1)),
                        )
                    })
                    .collect();
                let leaf_once = || -> Result<Vec<u8>, litchi_pptx::Error> {
                    if eager {
                        let mut edit = eager_package.edit()?;
                        edit.set_shape_text_chunks(0_usize, &chunks)?;
                        edit.commit()?.into_package().to_bytes()
                    } else {
                        let mut edit = source_package.edit()?;
                        edit.set_shape_text_chunks(Position::new(0), &chunks)?;
                        edit.commit()?.into_package().to_bytes()
                    }
                };
                let leaf = || {
                    if eager {
                        let mut edit = eager_package.edit().expect("edit");
                        edit.set_shape_text_chunks(0_usize, &chunks).expect("stage");
                        drop(black_box(edit.commit().expect("commit")));
                    } else {
                        let mut edit = source_package.edit().expect("edit");
                        edit.set_shape_text_chunks(Position::new(0), &chunks)
                            .expect("stage");
                        drop(black_box(edit.commit().expect("commit")));
                    }
                };
                let unified = || {
                    file.batch(items.clone(), BatchMode::Atomic)
                        .expect("batch")
                        .into_file()
                };
                let same = leaf_once()? == unified().to_bytes()?;
                report(
                    &name,
                    same,
                    paired(rounds_for(n), leaf, || drop(black_box(unified()))),
                );
            }
        }
    }

    // ------------------------------------------------------------------
    // Single operations, eager owner
    // ------------------------------------------------------------------
    let mut seed = litchi_pptx::Package::from_bytes(&deck(20)?)?;
    {
        let mut edit = seed.edit()?;
        let cells: Vec<Vec<String>> = (0..3)
            .map(|row| (0..3).map(|column| format!("r{row}c{column}")).collect())
            .collect();
        edit.add_table(0_usize, &cells, (0, 1_000_000, 3_000_000, 900_000))?;
        let png = b"\x89PNG\r\n\x1a\nbench-image".to_vec();
        edit.add_picture(
            0_usize,
            "Picture",
            &litchi_pptx::media_parts::Resource::new("/ppt/media/image1.png", "image/png", png),
            (0, 2_000_000, 914_400, 914_400),
        )?;
        for text in ["first", "second"] {
            edit.add_comment_text(0_usize, "Ada", "A", text, (0, 0), None)?;
        }
        seed = edit.commit()?.into_package();
    }
    let bytes = seed.to_bytes()?;
    let package = litchi_pptx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    // The 20 text boxes come first in the scene, then the table and the
    // picture.
    let table_shape = 20_usize;
    let picture_shape = 21_usize;
    let shape10 = Position::new(9);
    let body = TextBody::Shape { shape: shape10 };

    macro_rules! single {
        ($name:literal, $leaf:expr, $unified:expr) => {
            if wanted($name) {
                let leaf = $leaf;
                let unified = $unified;
                let same = leaf().to_bytes()? == unified().to_bytes()?;
                report(
                    $name,
                    same,
                    paired(
                        20,
                        || drop(black_box(leaf())),
                        || drop(black_box(unified())),
                    ),
                );
            }
        };
    }

    single!(
        "paragraph_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.insert_paragraph(
                0_usize,
                body,
                Position::new(1),
                TextProperties::default()
                    .with_text("added")
                    .with_bold(Some(true)),
            )
            .expect("insert");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]/shape[10]",
                NewNode::new("paragraph")
                    .prop("text", "added")
                    .prop("bold", true),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "run_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.insert_run(
                0_usize,
                body,
                Position::new(0),
                Position::new(1),
                TextProperties::default().with_text(" more"),
            )
            .expect("insert");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]/shape[10]/paragraph[1]",
                NewNode::new("run").prop("text", " more"),
            )
            .expect("add")
            .file
        }
    );
    let two = file
        .add(
            "/slide[1]/shape[10]",
            NewNode::new("paragraph").prop("text", "second"),
        )?
        .file;
    let two_package = litchi_pptx::Package::from_bytes(&two.to_bytes()?)?;
    single!(
        "paragraph_remove",
        || {
            let mut edit = two_package.edit().expect("edit");
            edit.remove_paragraph(0_usize, body, Position::new(1))
                .expect("remove");
            edit.commit().expect("commit").into_package()
        },
        || {
            two.remove("/slide[1]/shape[10]/paragraph[2]")
                .expect("remove")
        }
    );
    single!(
        "row_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.insert_table_row(0_usize, table_shape, Position::new(3), None, &["x", "y"])
                .expect("insert");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]/table[1]",
                NewNode::new("tr").prop("c1", "x").prop("c2", "y"),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "row_remove",
        || {
            let mut edit = package.edit().expect("edit");
            edit.remove_table_row(0_usize, table_shape, Position::new(1))
                .expect("remove");
            edit.commit().expect("commit").into_package()
        },
        || file.remove("/slide[1]/table[1]/tr[2]").expect("remove")
    );
    single!(
        "cell_fill",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_text_properties(
                0_usize,
                &[TextAssignment::new(
                    TextTarget::Cell {
                        shape: Position::new(table_shape),
                        row: Position::new(1),
                        column: Position::new(1),
                    },
                    TextProperties::default().with_fill(Some(CellFill::Scheme(
                        litchi_drawingml::color::Scheme::Accent1,
                    ))),
                )],
            )
            .expect("fill");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set(
                "/slide[1]/table[1]/tr[2]/tc[2]",
                props! { "fill" => "accent1" },
            )
            .expect("set")
        }
    );
    single!(
        "comment_remove",
        || {
            let mut edit = package.edit().expect("edit");
            edit.remove_comment_at(0_usize, Position::new(0))
                .expect("remove");
            edit.commit().expect("commit").into_package()
        },
        || file.remove("/slide[1]/comment[1]").expect("remove")
    );
    single!(
        "picture_frame_set",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_shape_properties(
                0_usize,
                &[
                    litchi_pptx::opened::ShapePropertiesReplacement::new(Position::new(
                        picture_shape,
                    ))
                    .with_bounds((914_400, 2_000_000, 1_828_800, 914_400)),
                ],
            )
            .expect("bounds");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set(
                "/slide[1]/picture[1]",
                props! { "x" => "1in", "width" => "2in" },
            )
            .expect("set")
        }
    );
    Ok(())
}

#[cfg(not(feature = "pptx"))]
fn main() {}
