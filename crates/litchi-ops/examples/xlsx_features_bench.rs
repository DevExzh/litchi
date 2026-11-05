//! Paired leaf-vs-unified measurements of the XLSX worksheet features of
//! record 0320.
//!
//! The input is a workbook made here: 40 notes, 40 conditional formats, 40
//! validations and 40 shapes on `Sheet1`. Each scenario times the leaf call
//! and the unified call doing the same work, alternating them ABBA within one
//! process (leaf, unified, unified, leaf per round pair):
//!
//! - `cf_get`: the 20th conditional format and its formatting (leaf:
//!   `Worksheet::conditional_formattings` and the differential formats;
//!   unified: `get("/Sheet1/cf[20]")` and its format map);
//! - `comment_get`: the 20th note (leaf: `Worksheet::comments`; unified:
//!   `get` and its format map);
//! - `comment_add`: one note (leaf: `edit_comments().set().commit()`;
//!   unified: `add`);
//! - `cf_add`: one cellIs rule with a fill (leaf: `add_differential`,
//!   `set_formats`, `commit`; unified: `add`);
//! - `link_set`: one cell hyperlink (leaf: `edit().put_hyperlink().commit()`;
//!   unified: `set link`);
//! - `shape_add`: one shape with text (leaf: `edit_drawings().add_shape()`;
//!   unified: `add`).
//!
//! `FEATURE_SCENARIO=<name>` runs one scenario. Each prints one JSON line:
//! the median and minimum microseconds of both legs and the ratio of the
//! medians. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example
//! xlsx_features_bench --features xlsx`.

#[cfg(feature = "xlsx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::{NewNode, props};
    use litchi_xlsx as xlsx;

    const ROUNDS: usize = 20;
    const WARMUP: usize = 2;
    const COUNT: usize = 40;

    let only = std::env::var("FEATURE_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);

    // The input workbook, built through the unified API.
    let mut file = litchi_ops::create(FileFormat::Xlsx)?;
    for index in 0..COUNT {
        let row = index + 2;
        file = file
            .add(
                "/Sheet1",
                NewNode::new("comment")
                    .prop("ref", format!("B{row}"))
                    .prop("text", format!("note {index}")),
            )?
            .file
            .add(
                "/Sheet1",
                NewNode::new("cf")
                    .prop("type", "cellIs")
                    .prop("ref", format!("C{row}"))
                    .prop("value", "5")
                    .prop("fill", "FFC7CE"),
            )?
            .file
            .add(
                "/Sheet1",
                NewNode::new("validation")
                    .prop("type", "list")
                    .prop("ref", format!("D{row}"))
                    .prop("formula1", "a,b"),
            )?
            .file
            .add(
                "/Sheet1",
                NewNode::new("shape").prop("text", format!("shape {index}")),
            )?
            .file;
    }
    let bytes = file.to_bytes()?;
    let unified = litchi_ops::open_bytes_with(
        bytes.clone(),
        &litchi_ops::OpenOptions::new().source_mode(litchi_ops::SourceMode::Eager),
    )?;
    let package = xlsx::Package::from_bytes(&bytes)?;

    fn stats(mut samples: Vec<f64>) -> (f64, f64) {
        samples.sort_by(f64::total_cmp);
        (samples[samples.len() / 2], samples[0])
    }

    /// Microseconds per call of each leg, `repeat` calls per sample, ABBA.
    fn paired(
        repeat: u32,
        mut leaf: impl FnMut(),
        mut unified: impl FnMut(),
    ) -> (Vec<f64>, Vec<f64>) {
        let time = |run: &mut dyn FnMut()| {
            let started = Instant::now();
            for _ in 0..repeat {
                run();
            }
            started.elapsed().as_secs_f64() * 1e6 / f64::from(repeat)
        };
        let (mut a, mut b) = (Vec::new(), Vec::new());
        for round in 0..WARMUP + ROUNDS / 2 {
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

    fn report(name: &str, (leaf, unified): (Vec<f64>, Vec<f64>)) {
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
            })
        );
    }

    let first_format =
        |view: litchi_ops::node::View<'_, '_>| -> usize { view.format().expect("format").count() };

    if wanted("cf_get") {
        report(
            "cf_get",
            paired(
                20,
                || {
                    let sheet = package.workbook().sheet(0).expect("sheet").expect("Sheet1");
                    let formats = sheet.conditional_formattings().expect("formats");
                    let dxfs = package.workbook().differential_formats().expect("dxfs");
                    let rule = &formats[19].rules[0];
                    if let Some(xlsx::conditional_formatting::DifferentialRef::StylesIndex(index)) =
                        rule.differential_format
                    {
                        black_box(
                            xlsx::conditional_formatting::edit::DifferentialStyle::read(
                                &dxfs[index as usize],
                            )
                            .expect("style"),
                        );
                    }
                    black_box(&formats);
                },
                || {
                    let node = unified.get("/Sheet1/cf[20]").expect("cf");
                    black_box(first_format(node.view()));
                },
            ),
        );
    }
    if wanted("comment_get") {
        report(
            "comment_get",
            paired(
                20,
                || {
                    let sheet = package.workbook().sheet(0).expect("sheet").expect("Sheet1");
                    let comments = sheet.comments().expect("comments").expect("part");
                    black_box(comments.get("B21"));
                },
                || {
                    let node = unified.get("/Sheet1/comment[20]").expect("comment");
                    black_box(first_format(node.view()));
                },
            ),
        );
    }
    if wanted("comment_add") {
        report(
            "comment_add",
            paired(
                5,
                || {
                    let mut edit = package.edit_comments().expect("edit");
                    edit.set(0, "Z99", "Author", "added").expect("set");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add(
                                "/Sheet1",
                                NewNode::new("comment")
                                    .prop("ref", "Z99")
                                    .prop("text", "added"),
                            )
                            .expect("add"),
                    );
                },
            ),
        );
    }
    if wanted("cf_add") {
        report(
            "cf_add",
            paired(
                5,
                || {
                    use xlsx::conditional_formatting::{
                        DifferentialRef, Formatting, Kind, Operator, Range, Rule,
                    };
                    let mut edit = package.edit_conditional_formats().expect("edit");
                    let mut formats = edit.formats(0).expect("formats");
                    let priority = formats
                        .iter()
                        .flat_map(|format| format.rules.iter().filter_map(|rule| rule.priority))
                        .max()
                        .unwrap_or(0)
                        + 1;
                    let dxf = edit
                        .add_differential(&xlsx::conditional_formatting::edit::DifferentialStyle {
                            fill: Some(xlsx::Rgb::from_hex("FFC7CE").expect("rgb")),
                            ..Default::default()
                        })
                        .expect("dxf");
                    let mut rule = Rule::new(Kind::CellIs, priority).expect("rule");
                    rule.operator = Some(Operator::GreaterThan);
                    rule.push_formula("5").expect("formula");
                    rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
                    formats.push(
                        Formatting::new(vec![Range::new("Z1:Z9").expect("range")], vec![rule])
                            .expect("format"),
                    );
                    edit.set_formats(0, formats).expect("set");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add(
                                "/Sheet1",
                                NewNode::new("cf")
                                    .prop("type", "cellIs")
                                    .prop("ref", "Z1:Z9")
                                    .prop("value", "5")
                                    .prop("fill", "FFC7CE"),
                            )
                            .expect("add"),
                    );
                },
            ),
        );
    }
    if wanted("link_set") {
        report(
            "link_set",
            paired(
                5,
                || {
                    let mut edit = package.edit().expect("edit");
                    let link = xlsx::hyperlinks::Hyperlink::external(
                        xlsx::hyperlinks::HyperlinkReference::new("Z1").expect("ref"),
                        "https://example.test/",
                    )
                    .expect("link");
                    edit.put_hyperlink(0, link).expect("put");
                    black_box(edit.commit().expect("commit"));
                },
                || {
                    black_box(
                        unified
                            .set("/Sheet1/Z1", props! { "link" => "https://example.test/" })
                            .expect("set"),
                    );
                },
            ),
        );
    }
    if wanted("shape_add") {
        report(
            "shape_add",
            paired(
                5,
                || {
                    use xlsx::drawing::edit::{Marker, Placement, ShapeSpec};
                    let mut edit = package.edit_drawings().expect("edit");
                    let count = edit
                        .elements(0)
                        .expect("elements")
                        .iter()
                        .filter(|element| element.kind == xlsx::drawing::edit::ElementKind::Shape)
                        .count();
                    let mut spec = ShapeSpec::new(
                        Placement::TwoCell {
                            from: Marker::cell(1, 1),
                            to: Marker::cell(6, 4),
                        },
                        format!("Shape {}", count + 1),
                    );
                    spec.text = Some("added".to_owned());
                    edit.add_shape(0, &spec).expect("add");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add("/Sheet1", NewNode::new("shape").prop("text", "added"))
                            .expect("add"),
                    );
                },
            ),
        );
    }
    Ok(())
}

#[cfg(not(feature = "xlsx"))]
fn main() {
    eprintln!("enable the `xlsx` feature");
}
