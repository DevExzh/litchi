//! Paired leaf and unified ODP reads and writes (record 0278).
//!
//! Run with `taskset -c <cpu> cargo run --release -p litchi-ops --features odp
//! --example odp_unified_bench -- [rounds]`. Each scenario times the leaf call
//! (`litchi_odp::Package`, `scene`, `Edit::splice_scene`) and the unified call
//! doing the same work, alternating A B B A within every round; fixture
//! construction, path parsing and the leaf/unified byte-equality checks are
//! outside timing. Prints one TSV row per sample, then per-scenario medians.
//! This probe is not frozen-harness qualification.

#[cfg(feature = "odp")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_odp::scene::{Change, ObjectRef, Scene, StoryRef};
    use litchi_ops::path::Path;

    /// A deck of `slides` slides, each with one text box of `lines`
    /// paragraphs, authored through the unified API.
    fn deck(slides: usize, lines: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(litchi_core::FileFormat::Odp).expect("create");
        let text = (0..lines)
            .map(|line| format!("line {line} of the slide body"))
            .collect::<Vec<_>>()
            .join("\n");
        for slide in 0..slides {
            if slide > 0 {
                file = file
                    .add("/", litchi_ops::NewNode::new("slide"))
                    .expect("slide")
                    .file;
            }
            file = file
                .add(
                    format!("/slide[{}]", slide + 1).as_str(),
                    litchi_ops::NewNode::new("shape")
                        .prop("text", text.as_str())
                        .prop("width", "20cm")
                        .prop("height", "12cm"),
                )
                .expect("text box")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    fn time<T>(repeats: usize, mut operation: impl FnMut() -> T) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            drop(black_box(operation()));
        }
        start.elapsed().as_nanos() / repeats as u128
    }

    struct Scenario<'a> {
        name: String,
        repeats: usize,
        leaf: Box<dyn FnMut() -> usize + 'a>,
        unified: Box<dyn FnMut() -> usize + 'a>,
    }

    fn median(values: &mut [u128]) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    /// The `(slide, paragraph)` targets of `count` distinct paragraph sets.
    fn targets(lines: usize, count: usize) -> Vec<(usize, usize)> {
        (0..count)
            .map(|index| (index / lines, index % lines))
            .collect()
    }

    fn leaf_batch(package: &litchi_odp::Package, targets: &[(usize, usize)]) -> Vec<u8> {
        let changes: Vec<Change> = targets
            .iter()
            .map(|&(slide, paragraph)| Change::SetParagraphText {
                story: StoryRef::object(ObjectRef::new(slide, 0)),
                paragraph,
                text: format!("set {slide} {paragraph}"),
            })
            .collect();
        let mut edit = package.edit().expect("leaf edit");
        edit.splice_scene(&changes).expect("leaf splice");
        edit.commit()
            .expect("leaf commit")
            .into_package()
            .as_bytes()
            .to_vec()
    }

    fn unified_batch(
        file: &litchi_ops::File,
        paths: &[(Path<'static>, String)],
    ) -> litchi_ops::File {
        file.edit(|tx| {
            for (path, text) in paths {
                tx.set(path, litchi_ops::props! { "text" => text.as_str() })?;
            }
            Ok(())
        })
        .expect("unified batch")
        .file()
        .clone()
    }

    pub(super) fn run(rounds: usize) {
        let small = deck(50, 20);
        let large = deck(100, 100);
        let package = litchi_odp::Package::from_bytes(&small).expect("leaf open");
        let file = litchi_ops::open_bytes(small.clone()).expect("unified open");
        let large_package = litchi_odp::Package::from_bytes(&large).expect("leaf open");
        let large_file = litchi_ops::open_bytes(large.clone()).expect("unified open");
        let first = Path::parse("/slide[1]/shape[1]/paragraph[1]").expect("path");
        let one = Path::parse("/slide[2]/shape[1]/paragraph[3]").expect("path");
        // Warm both scene indexes so the query rows compare the walks.
        let leaf_scene =
            Scene::parse(package.presentation().expect("projection").content_xml()).expect("scene");
        let _ = file.get(&first).expect("warm");

        let mut scenarios: Vec<Scenario<'_>> = Vec::new();
        scenarios.push(Scenario {
            name: "open+first get".into(),
            repeats: 20,
            leaf: Box::new(|| {
                let package = litchi_odp::Package::from_bytes(&small).expect("open");
                let presentation = package.presentation().expect("projection");
                let scene = Scene::parse(presentation.content_xml()).expect("scene");
                scene.pages()[0].objects()[0]
                    .story()
                    .expect("story")
                    .paragraphs()[0]
                    .text()
                    .len()
            }),
            unified: Box::new(|| {
                let file = litchi_ops::open_bytes(small.clone()).expect("open");
                let node = file.get(&first).expect("get");
                node.view()
                    .text()
                    .expect("text")
                    .map_or(0, |text| text.len())
            }),
        });
        scenarios.push(Scenario {
            name: "query paragraph (1000)".into(),
            repeats: 20,
            leaf: Box::new(|| {
                let mut total = 0;
                for page in leaf_scene.pages() {
                    for object in page.objects() {
                        if let Some(story) = object.story() {
                            for paragraph in story.paragraphs() {
                                total += paragraph.text().len();
                            }
                        }
                    }
                }
                total
            }),
            unified: Box::new(|| {
                file.query("paragraph")
                    .expect("query")
                    .map(|node| {
                        node.expect("node")
                            .view()
                            .text()
                            .expect("text")
                            .map_or(0, |text| text.len())
                    })
                    .sum()
            }),
        });
        scenarios.push(Scenario {
            name: "single set".into(),
            repeats: 20,
            leaf: Box::new(|| leaf_batch(&package, &[(1, 2)]).len()),
            unified: Box::new(|| {
                file.set(&one, litchi_ops::props! { "text" => "set 1 2" })
                    .expect("set")
                    .to_bytes()
                    .expect("bytes")
                    .len()
            }),
        });
        for (count, deck_package, deck_file, lines) in [
            (1, &package, &file, 20),
            (10, &package, &file, 20),
            (100, &package, &file, 20),
            (1_000, &package, &file, 20),
            (10_000, &large_package, &large_file, 100),
        ] {
            let targets = targets(lines, count);
            let paths: Vec<(Path<'static>, String)> = targets
                .iter()
                .map(|&(slide, paragraph)| {
                    (
                        Path::parse(&format!(
                            "/slide[{}]/shape[1]/paragraph[{}]",
                            slide + 1,
                            paragraph + 1
                        ))
                        .expect("path")
                        .into_owned(),
                        format!("set {slide} {paragraph}"),
                    )
                })
                .collect();
            // The two routes publish the same bytes.
            assert_eq!(
                leaf_batch(deck_package, &targets),
                unified_batch(deck_file, &paths).to_bytes().expect("bytes"),
                "batch {count}"
            );
            let repeats = if count >= 1_000 { 3 } else { 10 };
            let leaf_targets = targets.clone();
            scenarios.push(Scenario {
                name: format!("batch {count}"),
                repeats,
                leaf: Box::new(move || leaf_batch(deck_package, &leaf_targets).len()),
                unified: Box::new(move || {
                    unified_batch(deck_file, &paths)
                        .to_bytes()
                        .expect("bytes")
                        .len()
                }),
            });
        }
        assert_eq!(
            leaf_batch(&package, &[(1, 2)]),
            file.set(&one, litchi_ops::props! { "text" => "set 1 2" })
                .expect("set")
                .to_bytes()
                .expect("bytes")
        );

        println!("scenario\tround\tleaf_ns\tunified_ns");
        let mut summary = Vec::new();
        for scenario in &mut scenarios {
            for _ in 0..2 {
                // Warm-up.
                (scenario.leaf)();
                (scenario.unified)();
            }
            let (mut leaf, mut unified) = (Vec::new(), Vec::new());
            for round in 0..rounds {
                let (a, b) = if round % 2 == 0 {
                    let a = time(scenario.repeats, &mut scenario.leaf);
                    let b = time(scenario.repeats, &mut scenario.unified);
                    (a, b)
                } else {
                    let b = time(scenario.repeats, &mut scenario.unified);
                    let a = time(scenario.repeats, &mut scenario.leaf);
                    (a, b)
                };
                println!("{}\t{round}\t{a}\t{b}", scenario.name);
                leaf.push(a);
                unified.push(b);
            }
            summary.push((
                scenario.name.clone(),
                median(&mut leaf),
                median(&mut unified),
            ));
        }
        println!();
        println!("scenario\tleaf_median_us\tunified_median_us\toverhead");
        for (name, leaf, unified) in summary {
            println!(
                "{name}\t{:.1}\t{:.1}\t{:+.1}%",
                leaf as f64 / 1_000.0,
                unified as f64 / 1_000.0,
                (unified as f64 / leaf as f64 - 1.0) * 100.0
            );
        }
    }
}

fn main() {
    #[cfg(feature = "odp")]
    enabled::run(
        std::env::args()
            .nth(1)
            .and_then(|rounds| rounds.parse().ok())
            .unwrap_or(12),
    );
}
