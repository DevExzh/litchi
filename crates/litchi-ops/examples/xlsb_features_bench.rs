//! Paired leaf-vs-unified measurements of the XLSB worksheet features of
//! record 0322.
//!
//! The input is a workbook made here: 40 notes, 40 conditional formats, 40
//! validations, 40 hyperlinks and 40 shapes on `Sheet1`. Each scenario times
//! the leaf call on an eager `litchi_xlsb::Workbook` and the unified call on
//! an eager handle doing the same work, alternating them ABBA within one
//! process (leaf, unified, unified, leaf per round pair):
//!
//! - `cf_get`: the 20th conditional format and its formatting (leaf:
//!   `Workbook::sheet_features` and `differential_formats`; unified:
//!   `get("/Sheet1/cf[20]")` and its format map);
//! - `comment_get`: the 20th note in row-major order (leaf:
//!   `Workbook::sheet_comments`; unified: `get` and its format map);
//! - `comment_add`, `cf_add` (cellIs with a fill), `link_add`,
//!   `validation_add`, `shape_add`, `rowbreak_add`: one feature (leaf: the
//!   family's `Workbook::edit_*`, the change, `commit`; unified: `add`).
//!
//! `FEATURE_SCENARIO=<name>` runs one scenario. Each prints one JSON line:
//! the median and minimum microseconds of both legs and the ratio of the
//! medians. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example
//! xlsb_features_bench --features xlsb`.

#[cfg(feature = "xlsb")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::NewNode;
    use litchi_xlsb as xlsb;

    const ROUNDS: usize = 20;
    const WARMUP: usize = 2;
    const COUNT: usize = 40;

    let only = std::env::var("FEATURE_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);

    // The input workbook, built through the unified API.
    let mut file = litchi_ops::create(FileFormat::Xlsb)?;
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
                NewNode::new("hyperlink")
                    .prop("ref", format!("E{row}"))
                    .prop("url", format!("https://example.test/{index}")),
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
    let package = xlsb::Package::from_bytes(&bytes)?;
    let workbook = package.workbook()?;

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
                    let features = workbook
                        .sheet_features(0, xlsb::sheet::FeatureSet::FORMATS)
                        .expect("features");
                    let dxfs = workbook.differential_formats().expect("dxfs");
                    let rule = &features.formats[19].rules[0];
                    black_box(rule.dxf_id.and_then(|index| dxfs.get(index as usize)));
                    black_box(&features);
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
                    let mut notes = workbook.sheet_comments(0).expect("comments");
                    notes.sort_by_key(|note| (note.row, note.col));
                    black_box(&notes[19]);
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
                    edit.add(0, 98, 25, "Author", "added").expect("add");
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
                    use xlsb::conditional_formatting::{Formatting, Rule, RuleType};
                    let priority = workbook
                        .sheet_features(0, xlsb::sheet::FeatureSet::FORMATS)
                        .expect("features")
                        .formats
                        .iter()
                        .flat_map(|format| format.rules.iter().map(|rule| rule.priority))
                        .max()
                        .unwrap_or(0)
                        + 1;
                    let mut edit = package.edit_conditional_formats().expect("edit");
                    let dxf = edit
                        .add_differential(&xlsb::conditional_formatting::edit::DifferentialStyle {
                            fill: Some(0x00FF_C7CE),
                            ..Default::default()
                        })
                        .expect("dxf");
                    let mut rule = Rule::new(RuleType::CellIs, priority);
                    rule.operator = Some(5);
                    rule.parameter = 5;
                    rule.formula_texts = vec!["5".to_owned()];
                    rule.dxf_id = Some(dxf);
                    let mut format = Formatting::new(vec!["Z1:Z9".to_owned()]);
                    format.add_rule(rule);
                    edit.add(0, format).expect("add");
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
    if wanted("link_add") {
        report(
            "link_add",
            paired(
                5,
                || {
                    let mut edit = package.edit_hyperlinks().expect("edit");
                    let links = edit.hyperlinks(0).expect("links");
                    black_box(&links);
                    edit.add(
                        0,
                        xlsb::hyperlinks::Hyperlink::new_external(
                            0,
                            0,
                            25,
                            25,
                            "https://example.test/".to_owned(),
                        ),
                    )
                    .expect("add");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add(
                                "/Sheet1/Z1",
                                NewNode::new("hyperlink").prop("url", "https://example.test/"),
                            )
                            .expect("add"),
                    );
                },
            ),
        );
    }
    if wanted("validation_add") {
        report(
            "validation_add",
            paired(
                5,
                || {
                    let mut edit = package.edit_data_validations().expect("edit");
                    let existing = edit.validations(0).expect("validations");
                    black_box(&existing);
                    let mut validation =
                        xlsb::package::data_validation::Validation::new(3, "Z1:Z9".to_owned());
                    validation.formula1 = Some("\"a,b\"".to_owned());
                    edit.add(0, validation).expect("add");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add(
                                "/Sheet1",
                                NewNode::new("validation")
                                    .prop("type", "list")
                                    .prop("ref", "Z1:Z9")
                                    .prop("formula1", "a,b"),
                            )
                            .expect("add"),
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
                    use xlsb::drawing::edit::{ElementKind, Marker, Placement, ShapeSpec};
                    let mut edit = package.edit_drawings().expect("edit");
                    let count = edit
                        .elements(0)
                        .expect("elements")
                        .iter()
                        .filter(|element| element.kind == ElementKind::Shape)
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
    if wanted("rowbreak_add") {
        report(
            "rowbreak_add",
            paired(
                5,
                || {
                    let mut edit = package.edit_page_breaks().expect("edit");
                    let mut breaks = edit.page_breaks(0).expect("breaks");
                    breaks.rows.push(xlsb::page_breaks::Break::row(20));
                    edit.set_page_breaks(0, &breaks).expect("set");
                    black_box(edit.commit().expect("commit").into_package());
                },
                || {
                    black_box(
                        unified
                            .add("/Sheet1", NewNode::new("rowbreak").prop("row", 20))
                            .expect("add"),
                    );
                },
            ),
        );
    }
    Ok(())
}

#[cfg(not(feature = "xlsb"))]
fn main() {
    eprintln!("enable the `xlsb` feature");
}
