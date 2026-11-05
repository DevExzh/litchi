//! Paired leaf and unified timings for the record 0269 writes.
//!
//! PPTX (eager owner, generated one-slide decks of 1, 100 and 1,000 text
//! boxes): adding a formatted, filled, preset-geometry text box; setting a
//! shape's fill and geometry; setting a slide transition. The leaf leg is the
//! native `opened::Edit` calls plus one commit; the unified leg is one
//! `File::add` or `File::set`. XLSX: one cell `set` through the command
//! channel under `Unsupported::Refuse` and `Unsupported::Warn` (no property
//! refused), and a 100-item cell batch under both policies (Warn does not
//! coalesce, so each item retains its own snapshot: about 850 XLSX items
//! reach the default 256 MiB snapshot budget, as best effort does). Fixture generation, opening and payload construction are
//! outside the timed regions; the first pass of each pair checks that both
//! legs produce identical bytes.
//!
//! ```text
//! cargo run --release -p litchi-ops --features pptx,xlsx --example corpus_parity_bench
//! ```
//!
//! `LITCHI_PARITY_BENCH_ROUNDS` sets the ABBA rounds (default 15); each round
//! times every leg twice. Output is one JSON object per scenario.

#[cfg(all(feature = "pptx", feature = "xlsx"))]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::{FileFormat, Position};
    use litchi_drawingml::geom::Preset;
    use litchi_ops::batch::{Options, Unsupported};
    use litchi_ops::{BatchItem, File, NewNode, OpenOptions, Props, SourceMode, props};
    use litchi_pptx::opened::{Rgb, ShapeFill, TextAssignment, TextProperties, TextTarget};
    use litchi_pptx::transition::{Kind, Side, Transition};

    const DECKS: &[usize] = &[1, 100, 1_000];

    fn rounds() -> usize {
        std::env::var("LITCHI_PARITY_BENCH_ROUNDS")
            .ok()
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 0)
            .unwrap_or(15)
    }

    fn median(values: &mut [u128]) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    /// ABBA pairs of `a` and `b`; each returns elapsed nanoseconds.
    fn pair(name: &str, size: usize, mut a: impl FnMut() -> u128, mut b: impl FnMut() -> u128) {
        let (mut left, mut right) = (Vec::new(), Vec::new());
        for _ in 0..rounds() {
            left.push(a());
            right.push(b());
            right.push(b());
            left.push(a());
        }
        let (l, r) = (median(&mut left), median(&mut right));
        #[allow(clippy::cast_precision_loss, reason = "a reporting ratio")]
        let ratio = r as f64 / l as f64;
        println!(
            "{{\"scenario\":\"{name}\",\"size\":{size},\"samples\":{},\"leaf_or_a_median_ns\":{l},\"unified_or_b_median_ns\":{r},\"ratio\":{ratio:.3}}}",
            left.len()
        );
    }

    fn deck(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_pptx::builder::Builder::new().expect("builder");
        let slide = builder
            .presentation()
            .expect("presentation")
            .add_slide()
            .expect("slide");
        for index in 0..shapes {
            slide.add_text_box(&format!("box {index}"), 0, 0, 914_400, 457_200);
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn eager(bytes: &[u8]) -> File {
        litchi_ops::open_bytes_with(
            bytes.to_vec(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .expect("open")
    }

    fn format() -> TextProperties<'static> {
        TextProperties::default()
            .with_size(Some(2_400))
            .with_bold(Some(true))
            .with_color(Some(Rgb([0xFF, 0, 0])))
            .with_shape_fill(Some(ShapeFill::Solid(Rgb([0x1E, 0x3A, 0x5F]))))
            .with_geometry(Preset::Ellipse)
    }

    fn pptx(shapes: usize) {
        let bytes = deck(shapes);
        let native = litchi_pptx::Package::from_bytes(&bytes).expect("native");
        let file = eager(&bytes);
        let node = NewNode::new("shape")
            .with_props(props! {
                "text" => "Hello", "x" => 360_000, "y" => 360_000,
                "width" => 3_600_000, "height" => 720_000,
                "size" => 24, "bold" => true, "color" => "FF0000",
                "fill" => "1E3A5F", "geometry" => "ellipse",
            })
            .expect("node");

        let leaf_add = || {
            let mut edit = native.edit().expect("edit");
            edit.add_text_box(0_usize, "Hello", (360_000, 360_000, 3_600_000, 720_000))
                .expect("add");
            let shape = Position::new(shapes);
            edit.set_text_properties(
                0_usize,
                &[TextAssignment::new(TextTarget::Shape { shape }, format())],
            )
            .expect("format");
            edit.commit().expect("commit").into_package()
        };
        let check = leaf_add();
        let unified = file.add("/slide[1]", node.clone()).expect("unified add");
        assert_eq!(
            check.to_bytes().expect("leaf bytes"),
            unified.file.to_bytes().expect("unified bytes"),
            "add legs differ"
        );
        pair(
            "pptx_add_formatted_text_box",
            shapes,
            || {
                let start = Instant::now();
                black_box(leaf_add());
                start.elapsed().as_nanos()
            },
            || {
                let node = node.clone();
                let start = Instant::now();
                black_box(file.add("/slide[1]", node).expect("add"));
                start.elapsed().as_nanos()
            },
        );

        // The same text box without formatting: the add path's own overhead
        // before record 0269 (no body formatting edit).
        let plain = NewNode::new("shape")
            .with_props(props! {
                "text" => "Hello", "x" => 360_000, "y" => 360_000,
                "width" => 3_600_000, "height" => 720_000,
            })
            .expect("node");
        let leaf_plain = || {
            let mut edit = native.edit().expect("edit");
            edit.add_text_box(0_usize, "Hello", (360_000, 360_000, 3_600_000, 720_000))
                .expect("add");
            edit.commit().expect("commit").into_package()
        };
        let check = leaf_plain();
        let unified = file
            .add("/slide[1]", plain.clone())
            .expect("unified plain add");
        assert_eq!(
            check.to_bytes().expect("leaf bytes"),
            unified.file.to_bytes().expect("unified bytes"),
            "plain add legs differ"
        );
        pair(
            "pptx_add_plain_text_box",
            shapes,
            || {
                let start = Instant::now();
                black_box(leaf_plain());
                start.elapsed().as_nanos()
            },
            || {
                let node = plain.clone();
                let start = Instant::now();
                black_box(file.add("/slide[1]", node).expect("add"));
                start.elapsed().as_nanos()
            },
        );

        let fill = TextProperties::default()
            .with_shape_fill(Some(ShapeFill::Solid(Rgb([0xFF, 0, 0]))))
            .with_geometry(Preset::RoundRect);
        let leaf_fill = || {
            let mut edit = native.edit().expect("edit");
            edit.set_text_properties(
                0_usize,
                &[TextAssignment::new(
                    TextTarget::Shape {
                        shape: Position::new(0),
                    },
                    fill,
                )],
            )
            .expect("fill");
            edit.commit().expect("commit").into_package()
        };
        let check = leaf_fill();
        let unified = file
            .set(
                "/slide[1]/shape[1]",
                props! { "fill" => "FF0000", "geometry" => "roundRect" },
            )
            .expect("unified fill");
        assert_eq!(
            check.to_bytes().expect("leaf bytes"),
            unified.to_bytes().expect("unified bytes"),
            "fill legs differ"
        );
        pair(
            "pptx_set_shape_fill_geometry",
            shapes,
            || {
                let start = Instant::now();
                black_box(leaf_fill());
                start.elapsed().as_nanos()
            },
            || {
                let start = Instant::now();
                black_box(
                    file.set(
                        "/slide[1]/shape[1]",
                        props! { "fill" => "FF0000", "geometry" => "roundRect" },
                    )
                    .expect("set"),
                );
                start.elapsed().as_nanos()
            },
        );

        let push = Transition::new(Kind::Push(Side::Right));
        let leaf_transition = || {
            let mut edit = native.edit().expect("edit");
            edit.set_slide_transition(0_usize, Some(&push))
                .expect("transition");
            edit.commit().expect("commit").into_package()
        };
        let check = leaf_transition();
        let unified = file
            .set("/slide[1]", props! { "transition" => "push-right" })
            .expect("unified transition");
        assert_eq!(
            check.to_bytes().expect("leaf bytes"),
            unified.to_bytes().expect("unified bytes"),
            "transition legs differ"
        );
        pair(
            "pptx_set_slide_transition",
            shapes,
            || {
                let start = Instant::now();
                black_box(leaf_transition());
                start.elapsed().as_nanos()
            },
            || {
                let start = Instant::now();
                black_box(
                    file.set("/slide[1]", props! { "transition" => "push-right" })
                        .expect("set"),
                );
                start.elapsed().as_nanos()
            },
        );
    }

    fn cell(path: &'static str, value: &'static str) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(Props::new().with("value", value));
        item
    }

    fn xlsx() {
        let file = litchi_ops::create(FileFormat::Xlsx).expect("create");
        let refuse = Options::new();
        let warn = Options::new().unsupported(Unsupported::Warn);
        let item = cell("/Sheet1/B2", "42");
        assert_eq!(
            file.apply_with(&item, &refuse)
                .unwrap()
                .file()
                .to_bytes()
                .unwrap(),
            file.apply_with(&item, &warn)
                .unwrap()
                .file()
                .to_bytes()
                .unwrap()
        );
        pair(
            "xlsx_apply_refuse_vs_warn",
            1,
            || {
                let start = Instant::now();
                black_box(file.apply_with(&item, &refuse).expect("refuse"));
                start.elapsed().as_nanos()
            },
            || {
                let start = Instant::now();
                black_box(file.apply_with(&item, &warn).expect("warn"));
                start.elapsed().as_nanos()
            },
        );
        let paths: Vec<&'static str> = (1..=100)
            .map(|row| &*Box::leak(format!("/Sheet1/A{row}").into_boxed_str()))
            .collect();
        let items: Vec<BatchItem<'static>> = paths.iter().map(|path| cell(path, "x")).collect();
        assert_eq!(
            file.batch_with(&items, &refuse)
                .unwrap()
                .file()
                .to_bytes()
                .unwrap(),
            file.batch_with(&items, &warn)
                .unwrap()
                .file()
                .to_bytes()
                .unwrap()
        );
        pair(
            "xlsx_batch_100_refuse_vs_warn",
            100,
            || {
                let start = Instant::now();
                black_box(file.batch_with(&items, &refuse).expect("refuse"));
                start.elapsed().as_nanos()
            },
            || {
                let start = Instant::now();
                black_box(file.batch_with(&items, &warn).expect("warn"));
                start.elapsed().as_nanos()
            },
        );
        // Warn with one refused property: the retry cost against the subset.
        let mut refused = cell("/Sheet1/B2", "42");
        refused.props = Some(Props::new().with("value", "42").with("sparkle", true));
        pair(
            "xlsx_apply_subset_vs_warn_one_refused",
            1,
            || {
                let start = Instant::now();
                black_box(file.apply_with(&item, &refuse).expect("subset"));
                start.elapsed().as_nanos()
            },
            || {
                let start = Instant::now();
                black_box(file.apply_with(&refused, &warn).expect("warn"));
                start.elapsed().as_nanos()
            },
        );
    }

    pub(super) fn run() {
        xlsx();
        for &shapes in DECKS {
            pptx(shapes);
        }
    }
}

fn main() {
    #[cfg(all(feature = "pptx", feature = "xlsx"))]
    enabled::run();
    #[cfg(not(all(feature = "pptx", feature = "xlsx")))]
    eprintln!("corpus_parity_bench requires the litchi-ops `pptx` and `xlsx` features");
}
