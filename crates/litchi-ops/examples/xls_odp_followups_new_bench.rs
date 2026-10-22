//! Unified-against-leaf timing for the paths record 0316 adds.
//!
//! - `xls_row_insert_owners`: insert one row at the top of a 10,000-cell
//!   worksheet that also holds an array formula and a data table;
//! - `xls_row_insert_rules`: the same on the corpus workbooks whose
//!   conditional formats and validations now shift
//!   (`WithConditionalFormatting.xls`, `validation.xls`,
//!   `NewStyleConditionalFormattings.xls`) and whose comment boxes are
//!   position-locked (`SimpleWithComments.xls`);
//! - `odp_shape_format`: bold on a whole shape of a 200-slide deck,
//!   natively (`Package::edit` + `splice_scene` with one
//!   `Change::SetTextProperties` per paragraph + `commit`) and through
//!   `File::set`;
//! - `odp_shape_read`: the shared `bold` of every shape of that deck through
//!   the unified layer, against the same resolution done with the leaf
//!   scene API.
//!
//! Every leg pair asserts identical output first. Output and method as
//! `xls_odp_followups_bench`. Build: `cargo build --release -p litchi-ops
//! --features xls,odp --example xls_odp_followups_new_bench`.

#[cfg(all(feature = "xls", feature = "odp"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_odp::scene::{Change, CommonStyles, ObjectRef, Scene, StoryRef, TextPropertyUpdate};
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode, props};
    use litchi_xls::cell_values::Selector;
    use litchi_xls::formula_metadata::Range;
    use litchi_xls::writer::Builder;
    use litchi_xls::{DataTable, DataTableInputCell, DataTableRange};

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

    let row_insert =
        |name: &str, bytes: Vec<u8>, repeats: usize| -> Result<(), Box<dyn std::error::Error>> {
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
            Ok(())
        };

    if wanted("xls_row_insert_owners") {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Model")?;
        for row in 0..10_000_u32 {
            builder.write_number(sheet, row + 20, 0, f64::from(row))?;
        }
        builder.write_number(sheet, 0, 6, 1.0)?;
        builder.write_array_formula(sheet, Range::try_new(6, 1, 7, 2)?, "A21+7")?;
        let table = DataTable::one_variable(
            DataTableRange::new(12, 16, 3, 5)?,
            true,
            DataTableInputCell::Present { row: 0, col: 6 },
        );
        builder.add_data_table(sheet, 11, 2, table)?;
        row_insert("xls_row_insert_owners_10k", builder.build()?.to_bytes()?, 5)?;
    }

    if wanted("xls_row_insert_rules") {
        for (name, relative) in [
            (
                "xls_row_insert_conditional_formats",
                "ole/xls/WithConditionalFormatting.xls",
            ),
            (
                "xls_row_insert_validations",
                "libreoffice-core/sc/qa/unit/data/xls/validation.xls",
            ),
            (
                "xls_row_insert_cf12",
                "poi/test-data/spreadsheet/NewStyleConditionalFormattings.xls",
            ),
            (
                "xls_row_insert_locked_anchors",
                "ole/xls/SimpleWithComments.xls",
            ),
        ] {
            row_insert(name, fixture(relative)?, 20)?;
        }
    }

    if wanted("odp_shape_format") || wanted("odp_shape_read") {
        let mut deck = litchi_ops::create(litchi_core::FileFormat::Odp)?;
        for index in 0..200 {
            deck = deck.add("/", NewNode::new("slide"))?.file;
            let slide = format!("/slide[{}]", index + 1);
            deck = deck
                .add(
                    slide.as_str(),
                    NewNode::new("shape")
                        .prop("text", format!("Slide {index} text\nSecond line"))
                        .prop("width", "8cm")
                        .prop("height", "2cm"),
                )?
                .file;
        }
        let bytes = deck.to_bytes()?;
        if wanted("odp_shape_format") {
            let native = litchi_odp::Package::from_bytes(&bytes)?;
            let unified = litchi_ops::open_bytes(bytes.clone())?;
            let mut properties = TextPropertyUpdate::new();
            properties.bold = Some(true);
            let changes: Vec<Change> = (0..2)
                .map(|paragraph| Change::SetTextProperties {
                    story: StoryRef::object(ObjectRef::new(99, 0)),
                    paragraph,
                    run: None,
                    properties: properties.clone(),
                })
                .collect();
            let native_once = || {
                let mut edit = native.edit().expect("edit");
                edit.splice_scene(&changes).expect("splice");
                edit.commit().expect("commit")
            };
            let unified_once = || {
                unified
                    .set("/slide[100]/shape[1]", props! { "bold" => true })
                    .expect("set")
            };
            assert_eq!(
                native_once().into_parts().0.as_bytes(),
                unified_once().to_bytes()?.as_slice()
            );
            paired(
                "odp_shape_format_200",
                5,
                || {
                    black_box(native_once());
                },
                || {
                    black_box(unified_once());
                },
            );
        }
        if wanted("odp_shape_read") {
            let formatted = litchi_ops::open_bytes(bytes.clone())?
                .edit(|tx| {
                    for slide in 1..=200 {
                        tx.set(
                            format!("/slide[{slide}]/shape[1]").as_str(),
                            props! { "bold" => true },
                        )?;
                    }
                    Ok(())
                })?
                .file()
                .to_bytes()?;
            // Both legs open the package and read every shape's shared
            // `bold`: the leaf leg through the scene and style API.
            let native_once = || {
                let package = litchi_odp::Package::from_bytes(&formatted).expect("package");
                let presentation = package.presentation().expect("presentation");
                let content = presentation.content_xml();
                let scene = Scene::parse(content).expect("scene");
                let styles = scene.automatic_styles(content).expect("styles");
                let common = presentation
                    .styles_xml()
                    .map_or_else(|| Ok(CommonStyles::default()), CommonStyles::parse)
                    .expect("common");
                let mut bold = 0_usize;
                for page in scene.pages() {
                    for object in page.objects() {
                        let Some(story) = object.story() else {
                            continue;
                        };
                        let all = story.paragraphs().iter().all(|paragraph| {
                            let base = styles.resolve_paragraph(&common, paragraph);
                            paragraph.runs().iter().all(|run| {
                                let mut resolved = styles.resolve_run(&common, run);
                                resolved.inherit(&base);
                                resolved.bold() == Some(true)
                            })
                        });
                        bold += usize::from(all);
                    }
                }
                bold
            };
            let unified_once = || {
                let file = litchi_ops::open_bytes(formatted.as_slice()).expect("open");
                file.query("shape")
                    .expect("query")
                    .filter_map(Result::ok)
                    .filter(|node| {
                        node.view().format().is_ok_and(|format| {
                            format
                                .into_iter()
                                .any(|prop| prop.key() == "bold" && prop.to_string() == "true")
                        })
                    })
                    .count()
            };
            assert_eq!(native_once(), 200);
            assert_eq!(unified_once(), 200);
            paired(
                "odp_shape_read_200",
                5,
                || {
                    black_box(native_once());
                },
                || {
                    black_box(unified_once());
                },
            );
        }
    }
    Ok(())
}

#[cfg(not(all(feature = "xls", feature = "odp")))]
fn main() {
    eprintln!("build with --features xls,odp");
}
