//! Before/after timing for the existing paths record 0308 touches.
//!
//! The file compiles against the record's base tree too, so one binary per
//! tree runs the same scenarios:
//!
//! - `xls_complete_open`: the leaf's complete XLS reader over a 20,000-cell
//!   workbook and over the corpus `FormulaEvalTestData.xls` (Formula flags
//!   now mask the reserved bits);
//! - `xls_row_insert`: insert one row at the top of a 10,000-cell Litchi
//!   workbook, natively (`Package::edit` + `insert_rows` + `commit`) and
//!   through `File::add` (the shift now locates the worksheet's substream,
//!   its top-level records, the drawing stream and the stored extents);
//! - `rtf_move`: move the first of 1,000 plain paragraphs to the end, and
//!   the first of 1,000 producer paragraphs (`\pard\plain` each, the exact
//!   splice), natively and through `File::move_to`;
//! - `odp_scene`: `Scene::parse` over a 200-slide presentation with 400
//!   automatic styles (the scan now indexes them);
//! - `doc_text_set`: replace the text of one paragraph of a DOC file,
//!   natively and through `File::set` (the edit now checks range
//!   boundaries);
//! - `doc_package_text_set` (record 0311): the same set against the leaf
//!   `Package::edit` + `replace_paragraph` + `commit` it wraps.
//!
//! Every scenario prints one JSON line with the median and minimum of
//! `ROUNDS` samples (after `WARMUP`), each sample the mean of `REPEATS`
//! operations; paired legs alternate ABBA. Build:
//! `cargo build --release -p litchi-ops --features xls,rtf,odp,doc --example
//! office_gaps_bench`, then run each tree's binary alternately with
//! `taskset -c <core>`. `OFFICE_GAPS_SCENARIO=<name>` runs one scenario.

#[cfg(all(feature = "xls", feature = "rtf", feature = "odp", feature = "doc"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode, props};
    use litchi_xls::cell_values::Selector;
    use litchi_xls::writer::Builder;

    const ROUNDS: usize = 15;
    const WARMUP: usize = 3;

    let only = std::env::var("OFFICE_GAPS_SCENARIO").ok();
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

    if wanted("xls_row_insert") {
        let bytes = workbook(10_000)?;
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
            "xls_row_insert_10k",
            5,
            || {
                black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }

    if wanted("rtf_move") {
        for (name, separator, prefix) in [
            ("rtf_move_plain_1k", r"\par ", ""),
            (
                "rtf_move_producer_1k",
                "\\par\r\n\\pard\\plain\\f0 ",
                "\\pard\\plain\\f0 ",
            ),
        ] {
            let texts: Vec<_> = (0..1_000).map(|i| format!("paragraph {i:05}")).collect();
            let bytes = format!(
                r"{{\rtf1\ansi{{\fonttbl{{\f0 Times;}}}}{prefix}{}\par}}",
                texts.join(separator)
            )
            .into_bytes();
            let native = litchi_rtf::Package::from_bytes(&bytes)?;
            let unified = litchi_ops::open_bytes(bytes.as_slice())?;
            let native_once = || {
                let mut edit = native.edit();
                edit.move_paragraph(Position::new(0), Position::new(999))
                    .expect("move");
                edit.commit().expect("commit").into_package()
            };
            let unified_once = || unified.move_to("/body/p[1]", Placement::End).expect("move");
            assert_eq!(native_once().to_bytes()?, unified_once().to_bytes()?);
            paired(
                name,
                10,
                || {
                    black_box(native_once());
                },
                || {
                    black_box(unified_once());
                },
            );
        }
    }

    if wanted("odp_scene") {
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
        single("odp_scene_parse_200", 20, || {
            black_box(litchi_odp::scene::Scene::parse(&content).expect("scene"));
        });
    }

    if wanted("doc_text_set") {
        let bytes = fixture("ole/doc/NoHeadFoot.doc")?;
        // Record 0359a: the body-text engine is crate-private; its leaf leg is
        // `Package::edit`, the route it backs.
        let source = litchi_doc::Package::from_bytes(&bytes)?;
        let target = source
            .body_paragraphs(litchi_doc::body_text::Projection::All)?
            .into_iter()
            .find(|paragraph| !paragraph.text().is_empty())
            .expect("a paragraph with text")
            .position();
        let unified = litchi_ops::open_bytes(bytes.clone())?;
        let path = format!("/body/p[{}]", target.get() + 1);
        let native_once = || {
            let mut edit = source.edit().expect("edit");
            edit.replace_paragraph(target, "Replaced paragraph text")
                .expect("replace");
            edit.commit().expect("commit")
        };
        let unified_once = || {
            unified
                .set(
                    path.as_str(),
                    props! { "text" => "Replaced paragraph text" },
                )
                .expect("set")
        };
        paired(
            "doc_text_set",
            10,
            || {
                let _commit = black_box(native_once());
            },
            || {
                black_box(unified_once());
            },
        );
    }
    if wanted("doc_package_text_set") {
        // Record 0311: the same set against the leaf call it wraps,
        // `Package::edit` + `replace_paragraph` + `commit`, whose commit also
        // projects the published document (the `doc_text_set` leaf leg is the
        // body-text snapshot API and publishes no readable package).
        let bytes = fixture("ole/doc/NoHeadFoot.doc")?;
        let package = litchi_doc::Package::from_bytes(&bytes)?;
        let target = package
            .body_paragraphs(litchi_doc::body_text::Projection::All)?
            .into_iter()
            .find(|paragraph| !paragraph.text().is_empty())
            .expect("a paragraph with text")
            .position();
        let unified = litchi_ops::open_bytes(bytes.clone())?;
        let path = format!("/body/p[{}]", target.get() + 1);
        paired(
            "doc_package_text_set",
            10,
            || {
                let mut edit = package.edit().expect("edit");
                edit.replace_paragraph(target, "Replaced paragraph text")
                    .expect("replace");
                black_box(edit.commit().expect("commit").into_package());
            },
            || {
                black_box(
                    unified
                        .set(
                            path.as_str(),
                            props! { "text" => "Replaced paragraph text" },
                        )
                        .expect("set"),
                );
            },
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "xls", feature = "rtf", feature = "odp", feature = "doc")))]
fn main() {
    eprintln!("build with --features xls,rtf,odp,doc");
}
