//! Before/after timing for the existing paths record 0316 touches.
//!
//! The file compiles against the record's base tree too, so one binary per
//! tree runs the same scenarios:
//!
//! - `xls_complete_open`: the leaf's complete XLS reader over a 20,000-cell
//!   workbook and over the corpus `FormulaEvalTestData.xls` (Formula flags
//!   are now decoded with their defect, and WINDOW2 zoom defects drained);
//! - `xls_source_visit`: a source-backed walk of the 20,000 cells (each
//!   frame is now checked for an embedded substream);
//! - `xls_row_insert`: insert one row at the top of a 10,000-cell Litchi
//!   workbook and of the corpus `SimpleWithImages.xls`, natively
//!   (`Package::edit` + `insert_rows` + `commit`) and through `File::add`
//!   (rule records, formula anchors and drawing sizes are now visited);
//! - `odp_scene`: `Scene::parse` over a 200-slide presentation with 400
//!   automatic styles (the scan now counts each run's children);
//! - `odp_run_read`: the `bold` of every run of that deck through the
//!   unified layer (reads now resolve the style chain);
//! - `odp_run_format`: bold and size on one run of a 200-slide deck,
//!   natively and through `File::set`.
//!
//! Each scenario prints one JSON line with the median and minimum of
//! `ROUNDS` samples (after `WARMUP`), each sample the mean of `REPEATS`
//! operations; paired legs alternate ABBA. Build:
//! `cargo build --release -p litchi-ops --features xls,odp --example
//! xls_odp_followups_bench`, then run each tree's binary alternately with
//! `taskset -c <core>`. `FOLLOWUPS_SCENARIO=<name>` runs one scenario.

#[cfg(all(feature = "xls", feature = "odp"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::sync::Arc;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_odp::scene::{Change, ObjectRef, StoryRef, TextPropertyUpdate};
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode, props};
    use litchi_xls::cell_values::Selector;
    use litchi_xls::writer::Builder;

    const ROUNDS: usize = 15;
    const WARMUP: usize = 3;

    let only = std::env::var("FOLLOWUPS_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);

    fn sample(repeats: usize, mut run: impl FnMut()) -> f64 {
        let started = Instant::now();
        for _ in 0..repeats {
            run();
        }
        started.elapsed().as_secs_f64() * 1e6 / repeats as f64
    }

    fn summary(mut samples: Vec<f64>) -> (f64, f64) {
        samples.sort_by(f64::total_cmp);
        (samples[samples.len() / 2], samples[0])
    }

    fn single(name: &str, repeats: usize, mut run: impl FnMut()) {
        let mut samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let value = sample(repeats, &mut run);
            if round >= WARMUP {
                samples.push(value);
            }
        }
        let (median, min) = summary(samples);
        println!(
            "{}",
            serde_json::json!({ "scenario": name, "median_us": median, "min_us": min, "rounds": ROUNDS, "repeats": repeats })
        );
    }

    fn paired(name: &str, repeats: usize, mut native: impl FnMut(), mut unified: impl FnMut()) {
        let mut native_samples = Vec::new();
        let mut unified_samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let (n, u) = if round % 2 == 0 {
                let n = sample(repeats, &mut native);
                (n, sample(repeats, &mut unified))
            } else {
                let u = sample(repeats, &mut unified);
                (sample(repeats, &mut native), u)
            };
            if round >= WARMUP {
                native_samples.push(n);
                unified_samples.push(u);
            }
        }
        let (native_median, native_min) = summary(native_samples);
        let (unified_median, unified_min) = summary(unified_samples);
        println!(
            "{}",
            serde_json::json!({
                "scenario": name,
                "native_median_us": native_median, "native_min_us": native_min,
                "unified_median_us": unified_median, "unified_min_us": unified_min,
                "ratio": unified_median / native_median,
                "rounds": ROUNDS, "repeats": repeats, "order": "ABBA"
            })
        );
    }

    let fixture = |relative: &str| {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(relative),
        )
    };

    let workbook = |cells: usize| -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut builder = Builder::new();
        let data = builder.add_worksheet("Data")?;
        builder.add_worksheet("Archive")?;
        for row in 0..cells {
            builder.write_number(data, row as u32, 0, (row + 1) as f64)?;
        }
        Ok(builder.build()?.to_bytes()?)
    };

    if wanted("xls_complete_open") {
        let generated = workbook(20_000)?;
        single("xls_complete_open_20k", 20, || {
            black_box(
                litchi_xls::raw::Reader::new(std::io::Cursor::new(generated.clone()))
                    .expect("reader"),
            );
        });
        let corpus = fixture("ole/xls/FormulaEvalTestData.xls")?;
        single("xls_complete_open_formula_eval", 20, || {
            black_box(
                litchi_xls::raw::Reader::new(std::io::Cursor::new(corpus.clone())).expect("reader"),
            );
        });
    }

    if wanted("xls_source_visit") {
        let generated = workbook(20_000)?;
        let owner = litchi_xls::SourceBackedWorkbook::from_read_at(Arc::new(
            litchi_core::OwnedSource::new(generated),
        ))?;
        let sheet = owner.worksheet_by_index(0)?.expect("sheet");
        single("xls_source_visit_20k", 20, || {
            let mut cells = 0_usize;
            sheet
                .visit_cells(|_cell| {
                    cells += 1;
                    Ok(())
                })
                .expect("visit");
            black_box(cells);
        });
    }

    if wanted("xls_row_insert") {
        for (name, bytes, repeats) in [
            ("xls_row_insert_10k", workbook(10_000)?, 5),
            (
                "xls_row_insert_images",
                fixture("ole/xls/SimpleWithImages.xls")?,
                20,
            ),
        ] {
            let native = litchi_xls::Package::from_bytes(&bytes)?;
            let unified = litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )?;
            let native_once = || {
                let mut edit = native.edit().expect("edit");
                edit.insert_rows(Selector::Position(Position::new(0)), 0, 1)
                    .expect("insert");
                edit.commit().expect("commit").into_package()
            };
            let unified_once = || {
                unified
                    .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
                    .expect("add")
                    .file
            };
            assert_eq!(native_once().to_bytes()?, unified_once().to_bytes()?);
            paired(
                name,
                repeats,
                || {
                    black_box(native_once());
                },
                || {
                    black_box(unified_once());
                },
            );
        }
    }

    if wanted("odp_scene") || wanted("odp_run_read") {
        let mut styles = String::new();
        for index in 0..200 {
            styles.push_str(&format!(
                r#"<style:style style:name="P{index}" style:family="paragraph"><style:paragraph-properties fo:margin-left="0cm"/></style:style><style:style style:name="T{index}" style:family="text"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>"#
            ));
        }
        let mut pages = String::new();
        for index in 0..200 {
            pages.push_str(&format!(
                r#"<draw:page draw:name="page{index}"><draw:frame draw:name="Box"><draw:text-box><text:p text:style-name="P{index}">Slide {index} <text:span text:style-name="T{index}">bold</text:span> text</text:p></draw:text-box></draw:frame></draw:page>"#
            ));
        }
        let content = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
                r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
                r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
                r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
                r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
                r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
                r#"<office:automatic-styles>{}</office:automatic-styles>"#,
                r#"<office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#
            ),
            styles, pages
        );
        if wanted("odp_scene") {
            single("odp_scene_parse_200", 20, || {
                black_box(litchi_odp::scene::Scene::parse(&content).expect("scene"));
            });
        }
        if wanted("odp_run_read") {
            // The same scene as a flat document, opened once; every sample
            // reads all runs' `bold` (styles are indexed on first use).
            let flat = content
                .replacen("office:document-content", "office:document", 1)
                .replacen(
                    "xmlns:office=",
                    r#"office:mimetype="application/vnd.oasis.opendocument.presentation" xmlns:office="#,
                    1,
                )
                .replace(
                    "</office:document-content>",
                    "</office:document>",
                );
            let file = litchi_ops::open_bytes(flat.into_bytes())?;
            single("odp_run_read_200", 10, || {
                let mut bold = 0_usize;
                for node in file.query("run").expect("query") {
                    let node = node.expect("run");
                    let format = node.view().format().expect("format");
                    bold += format
                        .into_iter()
                        .filter(|prop| prop.key() == "bold" && prop.to_string() == "true")
                        .count();
                }
                assert_eq!(bold, 200);
                black_box(bold);
            });
        }
    }

    if wanted("odp_run_format") {
        let mut deck = litchi_ops::create(litchi_core::FileFormat::Odp)?;
        for index in 0..200 {
            deck = deck.add("/", NewNode::new("slide"))?.file;
            let slide = format!("/slide[{}]", index + 1);
            deck = deck
                .add(
                    slide.as_str(),
                    NewNode::new("shape")
                        .prop("text", format!("Slide {index} text"))
                        .prop("width", "8cm")
                        .prop("height", "2cm"),
                )?
                .file;
        }
        let bytes = deck.to_bytes()?;
        let native = litchi_odp::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes(bytes.clone())?;
        let path = "/slide[100]/shape[1]/paragraph[1]/run[1]";
        let mut properties = TextPropertyUpdate::new();
        properties.bold = Some(true);
        properties.size = Some("24pt".to_owned());
        let changes = [Change::SetTextProperties {
            story: StoryRef::object(ObjectRef::new(99, 0)),
            paragraph: 0,
            run: Some(0),
            properties,
        }];
        let native_once = || {
            let mut edit = native.edit().expect("edit");
            edit.splice_scene(&changes).expect("splice");
            edit.commit().expect("commit")
        };
        let unified_once = || {
            unified
                .set(path, props! { "bold" => true, "size" => "24pt" })
                .expect("set")
        };
        assert_eq!(
            native_once().into_parts().0.as_bytes(),
            unified_once().to_bytes()?.as_slice()
        );
        paired(
            "odp_run_format_200",
            5,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "xls", feature = "odp")))]
fn main() {
    eprintln!("build with --features xls,odp");
}
