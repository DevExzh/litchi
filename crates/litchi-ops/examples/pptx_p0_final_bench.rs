//! Paired leaf-vs-unified measurements of the last PPTX P0 rows (record
//! 0346) and of the PPTX batch routes (record 0348).
//!
//! Inputs are made here. Each scenario times the leaf call and the unified
//! call doing the same work on the same owner, ABBA within one process
//! (leaf, unified, unified, leaf per round pair), and checks outside the
//! clock that both legs publish the same bytes:
//!
//! - `cell_format_batch/<owner>/<N>`: one atomic `File::batch` of N
//!   `set /slide[1]/table[1]/tr[R]/tc[C] {bold, fill}` against one native
//!   edit staging the same N `TextAssignment`s in 256-edit chunks through one
//!   `set_text_property_chunks` call and one commit (a table of up to 100
//!   columns; the 1k-op batch budget: within 2×);
//! - `cell_paragraph_batch/<owner>/<N>`: N `set …/tc[C]/paragraph[1]
//!   {align}` against the same native call;
//! - `paragraph_text_batch/<owner>/<N>`: N `set /slide[1]/shape[i]/
//!   paragraph[1] {text}` over N text boxes;
//! - `notes_format_batch/eager/<N>`: N `set /slide[1]/notes/paragraph[i]
//!   {align}` against one `set_notes_text_property_chunks` call staging
//!   256-edit chunks in one notes-part pass (record 0348; record 0346 made
//!   one `set_notes_text_properties` call per chunk), `cold` on a fresh
//!   handle per batch and `warm` on one (record 0348);
//! - single operations on the eager owner (budget 10%): `comment_date_set`,
//!   `comment_author_set`, `modern_comment_add`, `modern_comment_set`,
//!   `modern_reply_add`, `placeholder_add`, `notes_add`,
//!   `slide_add_title_text`, `slide_layout_set`.
//!
//! `P0_SCENARIO=<prefix>` runs the scenarios whose name starts with it;
//! `P0_MAX=<N>` caps the batch sizes. Each prints one JSON line. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example
//! pptx_p0_final_bench --features pptx`.

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
        CellFill, ModernCommentChange, Rgb, TextAssignment, TextProperties, TextTarget,
    };

    const WARMUP: usize = 2;
    const CREATED: &str = "2026-10-19T09:30:00Z";

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

    /// Whether two packages hold the same parts with the same bytes (the
    /// leaf and the facade may order the archive members differently).
    fn same_parts(left: &[u8], right: &[u8]) -> bool {
        use std::io::Read as _;
        let parts = |bytes: &[u8]| -> Option<Vec<(String, Vec<u8>)>> {
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).ok()?;
            let mut parts = Vec::new();
            for index in 0..archive.len() {
                let mut member = archive.by_index(index).ok()?;
                let mut data = Vec::new();
                member.read_to_end(&mut data).ok()?;
                parts.push((member.name().to_owned(), data));
            }
            parts.sort();
            Some(parts)
        };
        parts(left).is_some_and(|left| parts(right) == Some(left))
    }

    fn set_item(path: String, props: Props<'static>) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(props);
        item
    }

    fn eager(bytes: &[u8]) -> Result<litchi_ops::File, litchi_ops::Error> {
        litchi_ops::open_bytes_with(
            bytes.to_vec(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
    }

    /// The builder's default layouts and no slides.
    fn layouts() -> Result<Vec<u8>, litchi_pptx::Error> {
        litchi_pptx::builder::Builder::new()?.build()?.to_bytes()
    }

    /// A one-slide deck of N text boxes.
    fn boxes(shapes: usize) -> Result<Vec<u8>, litchi_pptx::Error> {
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

    let rounds_for = |n: usize| match n {
        0..=100 => 10,
        101..=1_000 => 6,
        _ => 3,
    };
    let sizes: Vec<usize> = [1_usize, 10, 100, 1_000, 10_000]
        .into_iter()
        .filter(|n| *n <= max)
        .collect();

    // ------------------------------------------------------------------
    // Cell batches: a table of up to 100 columns holding N cells.
    // ------------------------------------------------------------------
    for &n in &sizes {
        let columns = n.min(100);
        let rows = n.div_ceil(columns);
        let mut seed = litchi_pptx::Package::from_bytes(&boxes(0)?)?;
        {
            let mut edit = seed.edit()?;
            let cells: Vec<Vec<String>> = (0..rows)
                .map(|row| {
                    (0..columns)
                        .map(|column| format!("r{row}c{column}"))
                        .collect()
                })
                .collect();
            edit.add_table(0_usize, &cells, (0, 0, 9_000_000, 6_000_000))?;
            seed = edit.commit()?.into_package();
        }
        let bytes = seed.to_bytes()?;
        let cells: Vec<(usize, usize)> = (0..n).map(|i| (i / columns, i % columns)).collect();
        for eager_owner in [true, false] {
            let owner = if eager_owner { "eager" } else { "source" };
            let file = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(if eager_owner {
                    SourceMode::Eager
                } else {
                    SourceMode::Auto
                }),
            )?;
            let eager_package = litchi_pptx::Package::from_bytes(&bytes)?;
            let source_package = litchi_pptx::source::Package::from_read_at(Arc::new(
                litchi_core::OwnedSource::new(bytes.clone()),
            ))?;
            for (scenario, paragraph) in
                [("cell_format_batch", false), ("cell_paragraph_batch", true)]
            {
                let name = format!("{scenario}/{owner}/{n}");
                if !wanted(&name) {
                    continue;
                }
                let properties = if paragraph {
                    TextProperties::default()
                        .with_align(Some(litchi_drawingml::text::Alignment::Center))
                } else {
                    TextProperties::default()
                        .with_bold(Some(true))
                        .with_fill(Some(CellFill::Rgb(Rgb([0xFF, 0xEE, 0xCC]))))
                };
                let edits: Vec<TextAssignment<'_>> = cells
                    .iter()
                    .map(|&(row, column)| {
                        let (row, column) = (Position::new(row), Position::new(column));
                        let target = if paragraph {
                            TextTarget::CellParagraph {
                                shape: Position::new(0),
                                row,
                                column,
                                paragraph: Position::new(0),
                            }
                        } else {
                            TextTarget::Cell {
                                shape: Position::new(0),
                                row,
                                column,
                            }
                        };
                        TextAssignment::new(target, properties)
                    })
                    .collect();
                let items: Vec<BatchItem<'static>> = cells
                    .iter()
                    .map(|&(row, column)| {
                        let path = format!("/slide[1]/table[1]/tr[{}]/tc[{}]", row + 1, column + 1);
                        if paragraph {
                            set_item(
                                format!("{path}/paragraph[1]"),
                                Props::new().with("align", "center"),
                            )
                        } else {
                            set_item(path, Props::new().with("bold", true).with("fill", "FFEECC"))
                        }
                    })
                    .collect();
                let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(256).collect();
                let leaf_bytes = || -> Result<Vec<u8>, litchi_pptx::Error> {
                    if eager_owner {
                        let mut edit = eager_package.edit()?;
                        edit.set_text_property_chunks(0_usize, &chunks)?;
                        edit.commit()?.into_package().to_bytes()
                    } else {
                        let mut edit = source_package.edit()?;
                        edit.set_text_property_chunks(Position::new(0), &chunks)?;
                        edit.commit()?.into_package().to_bytes()
                    }
                };
                let leaf = || {
                    if eager_owner {
                        let mut edit = eager_package.edit().expect("edit");
                        edit.set_text_property_chunks(0_usize, &chunks)
                            .expect("stage");
                        drop(black_box(edit.commit().expect("commit")));
                    } else {
                        let mut edit = source_package.edit().expect("edit");
                        edit.set_text_property_chunks(Position::new(0), &chunks)
                            .expect("stage");
                        drop(black_box(edit.commit().expect("commit")));
                    }
                };
                let unified = || {
                    file.batch(items.clone(), BatchMode::Atomic)
                        .expect("batch")
                        .into_file()
                };
                let same = same_parts(&leaf_bytes()?, &unified().to_bytes()?);
                // A handle memoizes the table a batch reads (record 0346):
                // the cold leg batches on a fresh handle each time, opened
                // outside the clock; the warm leg reuses one handle.
                let rounds = rounds_for(n);
                let mut fresh: Vec<litchi_ops::File> = (0..2 * (WARMUP + rounds))
                    .map(|_| {
                        litchi_ops::open_bytes_with(
                            bytes.clone(),
                            &OpenOptions::new().source_mode(if eager_owner {
                                SourceMode::Eager
                            } else {
                                SourceMode::Auto
                            }),
                        )
                    })
                    .collect::<Result<_, _>>()?;
                report(
                    &format!("{name}/cold"),
                    same,
                    paired(rounds, leaf, || {
                        let file = fresh.pop().expect("a fresh handle");
                        drop(black_box(
                            file.batch(items.clone(), BatchMode::Atomic)
                                .expect("batch")
                                .into_file(),
                        ));
                    }),
                );
                report(
                    &format!("{name}/warm"),
                    same,
                    paired(rounds, leaf, || drop(black_box(unified()))),
                );
            }
        }
    }

    // ------------------------------------------------------------------
    // Paragraph text batches over N text boxes.
    // ------------------------------------------------------------------
    for &n in &sizes {
        let bytes = boxes(n)?;
        for eager_owner in [true, false] {
            let owner = if eager_owner { "eager" } else { "source" };
            let name = format!("paragraph_text_batch/{owner}/{n}");
            if !wanted(&name) {
                continue;
            }
            let file = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(if eager_owner {
                    SourceMode::Eager
                } else {
                    SourceMode::Auto
                }),
            )?;
            let eager_package = litchi_pptx::Package::from_bytes(&bytes)?;
            let source_package = litchi_pptx::source::Package::from_read_at(Arc::new(
                litchi_core::OwnedSource::new(bytes.clone()),
            ))?;
            let texts: Vec<String> = (0..n).map(|index| format!("para-{index}")).collect();
            let edits: Vec<TextAssignment<'_>> = texts
                .iter()
                .enumerate()
                .map(|(index, text)| {
                    TextAssignment::new(
                        TextTarget::Paragraph {
                            shape: Position::new(index),
                            paragraph: Position::new(0),
                        },
                        TextProperties::default().with_text(text),
                    )
                })
                .collect();
            let items: Vec<BatchItem<'static>> = (0..n)
                .map(|index| {
                    set_item(
                        format!("/slide[1]/shape[{}]/paragraph[1]", index + 1),
                        Props::new().with("text", format!("para-{index}")),
                    )
                })
                .collect();
            let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(256).collect();
            let leaf_bytes = || -> Result<Vec<u8>, litchi_pptx::Error> {
                if eager_owner {
                    let mut edit = eager_package.edit()?;
                    edit.set_text_property_chunks(0_usize, &chunks)?;
                    edit.commit()?.into_package().to_bytes()
                } else {
                    let mut edit = source_package.edit()?;
                    edit.set_text_property_chunks(Position::new(0), &chunks)?;
                    edit.commit()?.into_package().to_bytes()
                }
            };
            let leaf = || {
                if eager_owner {
                    let mut edit = eager_package.edit().expect("edit");
                    edit.set_text_property_chunks(0_usize, &chunks)
                        .expect("stage");
                    drop(black_box(edit.commit().expect("commit")));
                } else {
                    let mut edit = source_package.edit().expect("edit");
                    edit.set_text_property_chunks(Position::new(0), &chunks)
                        .expect("stage");
                    drop(black_box(edit.commit().expect("commit")));
                }
            };
            let unified = || {
                file.batch(items.clone(), BatchMode::Atomic)
                    .expect("batch")
                    .into_file()
            };
            let same = same_parts(&leaf_bytes()?, &unified().to_bytes()?);
            report(
                &name,
                same,
                paired(rounds_for(n), leaf, || drop(black_box(unified()))),
            );
        }
    }

    // ------------------------------------------------------------------
    // Notes batches (eager owner): N notes paragraphs.
    // ------------------------------------------------------------------
    for &n in &sizes {
        let name = format!("notes_format_batch/eager/{n}");
        if !wanted(&name) {
            continue;
        }
        let deck = eager(&boxes(1)?)?;
        let lines: Vec<String> = (0..n).map(|index| format!("note-{index}")).collect();
        let deck = deck
            .add(
                "/slide[1]",
                NewNode::new("notes").prop("text", lines[0].clone()),
            )?
            .file;
        let mut deck = deck;
        for line in &lines[1..] {
            deck = deck
                .add(
                    "/slide[1]/notes",
                    NewNode::new("paragraph").prop("text", line.clone()),
                )?
                .file;
        }
        let bytes = deck.to_bytes()?;
        let file = eager(&bytes)?;
        let package = litchi_pptx::Package::from_bytes(&bytes)?;
        // The notes page's body placeholder: its positional scene index.
        let notes_shape = {
            use std::io::Read as _;
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes.clone()))?;
            let mut xml = Vec::new();
            archive
                .by_name("ppt/notesSlides/notesSlide1.xml")?
                .read_to_end(&mut xml)?;
            let scene = litchi_pptx::shape::Scene::read(&xml, litchi_pptx::shape::Limits::DEFAULT)?;
            scene
                .iter()
                .position(|shape| {
                    shape
                        .placeholder()
                        .is_some_and(|placeholder| placeholder.kind() == Some("body"))
                })
                .ok_or("the notes page has a body")?
        };
        let edits: Vec<TextAssignment<'_>> = (0..n)
            .map(|index| {
                TextAssignment::new(
                    TextTarget::Paragraph {
                        shape: Position::new(notes_shape),
                        paragraph: Position::new(index),
                    },
                    TextProperties::default()
                        .with_align(Some(litchi_drawingml::text::Alignment::Right)),
                )
            })
            .collect();
        let items: Vec<BatchItem<'static>> = (1..=n)
            .map(|index| {
                set_item(
                    format!("/slide[1]/notes/paragraph[{index}]"),
                    Props::new().with("align", "right"),
                )
            })
            .collect();
        let leaf_once = || -> Result<litchi_pptx::Package, litchi_pptx::Error> {
            // One notes-part pass for every chunk (record 0348).
            let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(256).collect();
            let mut edit = package.edit()?;
            edit.set_notes_text_property_chunks(0_usize, &chunks)?;
            Ok(edit.commit()?.into_package())
        };
        let unified = || {
            file.batch(items.clone(), BatchMode::Atomic)
                .expect("batch")
                .into_file()
        };
        let same = same_parts(&leaf_once()?.to_bytes()?, &unified().to_bytes()?);
        // A handle memoizes the notes body a batch reads (record 0348): the
        // cold leg batches on a fresh handle each time, opened outside the
        // clock; the warm leg reuses one handle.
        let rounds = rounds_for(n);
        let mut fresh: Vec<litchi_ops::File> = (0..2 * (WARMUP + rounds))
            .map(|_| eager(&bytes))
            .collect::<Result<_, _>>()?;
        report(
            &format!("{name}/cold"),
            same,
            paired(
                rounds,
                || drop(black_box(leaf_once().expect("leaf"))),
                || {
                    let file = fresh.pop().expect("a fresh handle");
                    drop(black_box(
                        file.batch(items.clone(), BatchMode::Atomic)
                            .expect("batch")
                            .into_file(),
                    ));
                },
            ),
        );
        report(
            &format!("{name}/warm"),
            same,
            paired(
                rounds,
                || drop(black_box(leaf_once().expect("leaf"))),
                || drop(black_box(unified())),
            ),
        );
    }

    // ------------------------------------------------------------------
    // Single operations, eager owner
    // ------------------------------------------------------------------
    let seed = eager(&layouts()?)?
        .add("/", NewNode::new("slide").prop("layout", "Blank"))?
        .file;
    let mut seed = seed;
    for text in ["first", "second"] {
        seed = seed
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada")
                    .prop("text", text),
            )?
            .file;
    }
    seed = seed
        .add(
            "/slide[1]",
            NewNode::new("modernComment")
                .prop("author", "Grace")
                .prop("text", "Thread")
                .prop("created", CREATED),
        )?
        .file;
    let bytes = seed.to_bytes()?;
    let package = litchi_pptx::Package::from_bytes(&bytes)?;
    let file = eager(&bytes)?;

    macro_rules! single {
        ($name:literal, $leaf:expr, $unified:expr) => {
            if wanted($name) {
                let leaf = $leaf;
                let unified = $unified;
                let same = same_parts(&leaf().to_bytes()?, &unified().to_bytes()?);
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
        "comment_date_set",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_comment_date_at(0_usize, Position::new(0), "2026-10-19T08:00:00Z")
                .expect("date");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set(
                "/slide[1]/comment[1]",
                props! { "date" => "2026-10-19T08:00:00Z" },
            )
            .expect("set")
        }
    );
    single!(
        "comment_author_set",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_comment_author_at(0_usize, Position::new(0), Some("Ada King"), None)
                .expect("author");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set("/slide[1]/comment[1]", props! { "author" => "Ada King" })
                .expect("set")
        }
    );
    single!(
        "modern_comment_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.add_modern_comment_text(0_usize, "Ada", Some("A"), "Check", CREATED, false)
                .expect("add");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]",
                NewNode::new("modernComment")
                    .prop("author", "Ada")
                    .prop("text", "Check")
                    .prop("created", CREATED),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "modern_comment_set",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_modern_comment_at(
                0_usize,
                Position::new(0),
                None,
                &ModernCommentChange {
                    text: Some("Reworded"),
                    resolved: Some(true),
                    ..ModernCommentChange::default()
                },
            )
            .expect("set");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set(
                "/slide[1]/moderncomment[1]",
                props! { "text" => "Reworded", "resolved" => true },
            )
            .expect("set")
        }
    );
    single!(
        "modern_reply_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.add_modern_comment_reply_text(
                0_usize,
                Position::new(0),
                "Ada",
                Some("A"),
                "Agreed",
                CREATED,
            )
            .expect("reply");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]",
                NewNode::new("modernComment")
                    .prop("author", "Ada")
                    .prop("text", "Agreed")
                    .prop("created", CREATED)
                    .prop("parent", "/slide[1]/moderncomment[1]"),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "placeholder_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.add_placeholder(0_usize, litchi_pptx::PlaceholderKind::Title, "Heading")
                .expect("add");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]",
                NewNode::new("placeholder")
                    .prop("phType", "title")
                    .prop("text", "Heading"),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "notes_add",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_notes_text(0_usize, "Speaker notes")
                .expect("notes");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/slide[1]",
                NewNode::new("notes").prop("text", "Speaker notes"),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "slide_add_title_text",
        || {
            let mut edit = package.edit().expect("edit");
            edit.insert_slide(Position::new(1), "Title and Content")
                .expect("insert");
            edit.put_placeholder_text(1_usize, litchi_pptx::PlaceholderKind::Title, None, "Review")
                .expect("title");
            edit.put_placeholder_text(
                1_usize,
                litchi_pptx::PlaceholderKind::Body,
                Some(1),
                "Revenue",
            )
            .expect("body");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Title and Content")
                    .prop("title", "Review")
                    .prop("text", "Revenue"),
            )
            .expect("add")
            .file
        }
    );
    single!(
        "slide_layout_set",
        || {
            let mut edit = package.edit().expect("edit");
            edit.set_slide_layout(0_usize, "Title Only")
                .expect("layout");
            edit.commit().expect("commit").into_package()
        },
        || {
            file.set("/slide[1]", props! { "layout" => "Title Only" })
                .expect("set")
        }
    );
    Ok(())
}

#[cfg(not(feature = "pptx"))]
fn main() {}
