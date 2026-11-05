//! Paired timing of the ODF follow-ups of record 0288: each leaf crate's own
//! API against the unified `litchi_ops` API on the same input.
//!
//! Scenarios, each timed as leaf and unified legs in ABBA order (per-call
//! means of the calls per sample, medians over `ROUNDS`):
//!
//! - `batch_N` (ODG shape text, ODB query command, formula token text): `N`
//!   sets in one leaf `Edit` and one commit, against one unified
//!   `File::edit` closure of `N` sets (coalesced into one leaf commit);
//! - `metadata_set` (ODG, ODS, ODP): one title through the metadata family
//!   against `set("/metadata", ..)`;
//! - `insert_page`, `move_page` (ODG, a LibreOffice drawing): one page
//!   inserted first, or the first page moved last;
//! - `title_set` (ODC, a typed chart), `geometry_set` (ODI, a fresh image).
//!
//! `--single` skips the batch scenarios. Inputs are built outside the clock; the leaf and unified outputs of every
//! write scenario are checked byte-identical outside it. Run pinned,
//! release:
//! `taskset -c 9 cargo run --release -p litchi-ops --features ods,odp,odg,odc,odi,odb,odf-formula --example odf_gaps_bench`.

#[cfg(all(
    feature = "ods",
    feature = "odp",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "odf-formula"
))]
mod bench {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_odf_common::metadata::Field;
    use litchi_ops::{File, NewNode, Placement, props};

    const ROUNDS: usize = 12;
    const WARMUP: usize = 2;

    fn median(mut samples: Vec<f64>) -> f64 {
        samples.sort_by(f64::total_cmp);
        samples[samples.len() / 2]
    }

    /// Per-call means of `calls` calls, sampled in ABBA order over the legs.
    fn paired(calls: usize, legs: &mut [&mut dyn FnMut()]) -> Vec<f64> {
        let time = |run: &mut dyn FnMut()| {
            let start = Instant::now();
            for _ in 0..calls {
                run();
            }
            start.elapsed().as_secs_f64() * 1e6 / calls as f64
        };
        for _ in 0..WARMUP {
            for leg in legs.iter_mut() {
                leg();
            }
        }
        let mut samples = vec![Vec::new(); legs.len()];
        for round in 0..ROUNDS {
            let order: Vec<usize> = if round % 2 == 0 {
                (0..legs.len()).collect()
            } else {
                (0..legs.len()).rev().collect()
            };
            for index in order {
                samples[index].push(time(&mut *legs[index]));
            }
        }
        samples.into_iter().map(median).collect()
    }

    fn fixture(path: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(path),
        )
        .expect("fixture")
    }

    fn row(rows: &mut Vec<serde_json::Value>, format: &str, scenario: &str, times: &[f64]) {
        let (leaf, unified) = (times[0], times[1]);
        println!(
            "{format:>4} {scenario:>13}  leaf {leaf:>11.1} µs  unified {unified:>11.1} µs  ratio {:.3}",
            unified / leaf
        );
        rows.push(serde_json::json!({
            "format": format,
            "scenario": scenario,
            "leaf_us": leaf,
            "unified_us": unified,
            "ratio_unified_leaf": unified / leaf,
        }));
    }

    fn open(bytes: &[u8]) -> File {
        litchi_ops::open_bytes(bytes.to_vec()).expect("unified open")
    }

    /// A drawing with one page of `count` named text rectangles.
    fn drawing(count: usize) -> Vec<u8> {
        let package = litchi_odg::Package::new().expect("new");
        let mut edit = package.edit();
        edit.add_page(litchi_odg::page::Page::new("page1"))
            .expect("page");
        let package = edit.commit().expect("commit").into_package();
        let mut edit = package.edit();
        for index in 0..count {
            edit.add_shape(
                0,
                litchi_odg::shape::Shape::new(litchi_odg::shape::ShapeKind::Rectangle)
                    .with_geometry("1cm", "1cm", "2cm", "1cm")
                    .with_text(format!("shape {index}"))
                    .with_name(format!("n{index}")),
            )
            .expect("shape");
        }
        edit.commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes")
    }

    fn odg_batches(rows: &mut Vec<serde_json::Value>) {
        let bytes = drawing(1_000);
        let package = litchi_odg::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        for count in [1usize, 10, 100, 1_000] {
            let texts: Vec<String> = (0..count).map(|index| format!("edited {index}")).collect();
            let paths: Vec<String> = (1..=count)
                .map(|index| format!("/page[1]/shape[{index}]"))
                .collect();
            let leaf_once = || {
                let mut edit = package.edit();
                for (index, text) in texts.iter().enumerate() {
                    edit.set_shape_text(0, index, text.as_str()).expect("stage");
                }
                edit.commit().expect("commit").into_package()
            };
            let unified_once = || {
                file.edit(|tx| {
                    for (path, text) in paths.iter().zip(&texts) {
                        tx.set(path.as_str(), props! { "text" => text.as_str() })?;
                    }
                    Ok(())
                })
                .expect("unified")
            };
            let commit = unified_once();
            assert_eq!(commit.diagnostics().commits(), 1);
            assert_eq!(
                commit.file().to_bytes().expect("bytes"),
                leaf_once().as_bytes(),
                "ODG batch outputs differ"
            );
            let mut leaf = || {
                black_box(leaf_once());
            };
            let mut unified = || {
                black_box(unified_once());
            };
            let calls = (1_000 / count).clamp(1, 50);
            let times = paired(calls, &mut [&mut leaf, &mut unified]);
            row(rows, "odg", &format!("batch_{count}"), &times);
        }
    }

    fn odb_batches(rows: &mut Vec<serde_json::Value>) {
        let package = litchi_odb::Package::new().expect("new");
        let mut edit = package.edit();
        for index in 0..1_000 {
            edit.add_query(litchi_odb::Query::new(
                format!("q{index}"),
                format!("SELECT {index} FROM t"),
            ))
            .expect("query");
        }
        let bytes = edit
            .commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes");
        let package = litchi_odb::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        for count in [1usize, 10, 100, 1_000] {
            let names: Vec<String> = (0..count).map(|index| format!("q{index}")).collect();
            let commands: Vec<String> = (0..count)
                .map(|index| format!("SELECT {index} FROM u"))
                .collect();
            let paths: Vec<String> = (1..=count)
                .map(|index| format!("/query[{index}]"))
                .collect();
            let leaf_once = || {
                let mut edit = package.edit();
                for (name, command) in names.iter().zip(&commands) {
                    edit.set_query_command(name, command.as_str())
                        .expect("stage");
                }
                edit.commit().expect("commit").into_package()
            };
            let unified_once = || {
                file.edit(|tx| {
                    for (path, command) in paths.iter().zip(&commands) {
                        tx.set(path.as_str(), props! { "command" => command.as_str() })?;
                    }
                    Ok(())
                })
                .expect("unified")
            };
            let commit = unified_once();
            assert_eq!(commit.diagnostics().commits(), 1);
            assert_eq!(
                commit.file().to_bytes().expect("bytes"),
                leaf_once().as_bytes(),
                "ODB batch outputs differ"
            );
            let mut leaf = || {
                black_box(leaf_once());
            };
            let mut unified = || {
                black_box(unified_once());
            };
            let calls = (1_000 / count).clamp(1, 50);
            let times = paired(calls, &mut [&mut leaf, &mut unified]);
            row(rows, "odb", &format!("batch_{count}"), &times);
        }
    }

    fn formula_batches(rows: &mut Vec<serde_json::Value>) {
        let mut mathml = String::from("<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>");
        for index in 0..1_000 {
            if index > 0 {
                mathml.push_str("<mo>+</mo>");
            }
            mathml.push_str(&format!("<mi>a{index}</mi>"));
        }
        mathml.push_str("</mrow></math>");
        let bytes = litchi_odf_formula::Builder::from_mathml(mathml)
            .build()
            .expect("build")
            .to_bytes()
            .expect("bytes");
        let package = litchi_odf_formula::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        for count in [1usize, 10, 100, 1_000] {
            let texts: Vec<String> = (0..count).map(|index| format!("b{index}")).collect();
            let routes: Vec<litchi_odf_formula::NodePath> = (0..count)
                .map(|index| litchi_odf_formula::NodePath::new(vec![0, index * 2]))
                .collect();
            let paths: Vec<String> = (0..count)
                .map(|index| format!("/element[1]/element[{}]", index * 2 + 1))
                .collect();
            let leaf_once = || {
                let mut edit = package.edit();
                for (route, text) in routes.iter().zip(&texts) {
                    edit.set_text(route, text).expect("stage");
                }
                edit.commit().expect("commit").into_package()
            };
            let unified_once = || {
                file.edit(|tx| {
                    for (path, text) in paths.iter().zip(&texts) {
                        tx.set(path.as_str(), props! { "text" => text.as_str() })?;
                    }
                    Ok(())
                })
                .expect("unified")
            };
            let commit = unified_once();
            assert_eq!(commit.diagnostics().commits(), 1);
            assert_eq!(
                commit.file().to_bytes().expect("bytes"),
                leaf_once().as_bytes(),
                "formula batch outputs differ"
            );
            let mut leaf = || {
                black_box(leaf_once());
            };
            let mut unified = || {
                black_box(unified_once());
            };
            let calls = (1_000 / count).clamp(1, 50);
            let times = paired(calls, &mut [&mut leaf, &mut unified]);
            row(rows, "odf", &format!("batch_{count}"), &times);
        }
    }

    fn odg_single(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("libreoffice-core/desktop/qa/data/3page.odg");
        let package = litchi_odg::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);

        let leaf_metadata = || {
            let mut edit = package.edit_metadata();
            edit.set(Field::Title, Some("Benchmark")).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_metadata = || {
            file.set("/metadata", props! { "title" => "Benchmark" })
                .expect("unified")
        };
        assert_eq!(
            unified_metadata().to_bytes().expect("bytes"),
            leaf_metadata().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_metadata());
        };
        let mut unified = || {
            black_box(unified_metadata());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odg", "metadata_set", &times);

        let leaf_insert = || {
            let mut edit = package.edit();
            let mut page = litchi_odg::page::Page::new("inserted");
            let last = package.drawing().pages().last().cloned().expect("page");
            if let Some(master) = last.master_page_name() {
                page = page.with_master_page_name(master);
            }
            if let Some(style) = last.style_name() {
                page = page.with_style_name(style);
            }
            edit.insert_page(0, page).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_insert = || {
            file.add(
                "/",
                NewNode::new("page")
                    .prop("name", "inserted")
                    .at(Placement::Start),
            )
            .expect("unified")
            .file
        };
        assert_eq!(
            unified_insert().to_bytes().expect("bytes"),
            leaf_insert().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_insert());
        };
        let mut unified = || {
            black_box(unified_insert());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odg", "insert_page", &times);

        let leaf_move = || {
            let mut edit = package.edit();
            edit.move_page(0, 2).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_move = || file.move_to("/page[1]", Placement::End).expect("unified");
        assert_eq!(
            unified_move().to_bytes().expect("bytes"),
            leaf_move().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_move());
        };
        let mut unified = || {
            black_box(unified_move());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odg", "move_page", &times);
    }

    fn family_metadata(rows: &mut Vec<serde_json::Value>) {
        let ods = fixture("odf/corpus/calc-two-sheets.ods");
        let package = litchi_ods::Package::from_bytes(&ods).expect("leaf");
        let file = open(&ods);
        let leaf_once = || {
            let mut edit = package.edit_metadata();
            edit.set(Field::Title, Some("Benchmark")).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            file.set("/metadata", props! { "title" => "Benchmark" })
                .expect("unified")
        };
        assert_eq!(
            unified_once().to_bytes().expect("bytes"),
            leaf_once().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_once());
        };
        let mut unified = || {
            black_box(unified_once());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "ods", "metadata_set", &times);

        let odp = fixture("odf/corpus/impress-basic.odp");
        let package = litchi_odp::Package::from_bytes(&odp).expect("leaf");
        let file = open(&odp);
        let leaf_once = || {
            let mut edit = package.edit_metadata();
            edit.set(Field::Title, Some("Benchmark")).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            file.set("/metadata", props! { "title" => "Benchmark" })
                .expect("unified")
        };
        assert_eq!(
            unified_once().to_bytes().expect("bytes"),
            leaf_once().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_once());
        };
        let mut unified = || {
            black_box(unified_once());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odp", "metadata_set", &times);
    }

    fn odc_and_odi(rows: &mut Vec<serde_json::Value>) {
        // A typed chart: two axes and three series.
        let package = litchi_odc::Package::new().expect("new");
        let mut definition = package.chart().definition().expect("definition");
        for dimension in [
            litchi_odc::chart::Dimension::X,
            litchi_odc::chart::Dimension::Y,
        ] {
            definition
                .plot_area
                .axes
                .push(litchi_odc::AxisSpec::new(dimension));
        }
        for index in 0..3 {
            definition.plot_area.series.push(litchi_odc::SeriesSpec {
                values_cell_range_address: Some(format!("local-table.B{}:B9", index + 1)),
                ..litchi_odc::SeriesSpec::default()
            });
        }
        let mut edit = package.edit();
        edit.replace_chart(&definition).expect("replace");
        let bytes = edit
            .commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes");
        let package = litchi_odc::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        let leaf_once = || {
            let mut definition = package.chart().definition().expect("definition");
            definition.title = Some(litchi_odc::Text::new("Benchmark"));
            let mut edit = package.edit();
            edit.replace_chart(&definition).expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            file.set("/", props! { "title" => "Benchmark" })
                .expect("unified")
        };
        assert_eq!(
            unified_once().to_bytes().expect("bytes"),
            leaf_once().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_once());
        };
        let mut unified = || {
            black_box(unified_once());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odc", "title_set", &times);

        let package = litchi_odi::Package::new().expect("new");
        let bytes = package.to_bytes().expect("bytes");
        let package = litchi_odi::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        let leaf_once = || {
            let mut edit = package.edit();
            edit.set_geometry(
                0,
                Some("1cm".into()),
                Some("2cm".into()),
                Some("3cm".into()),
                Some("4cm".into()),
            )
            .expect("stage");
            edit.commit().expect("commit").into_package()
        };
        let unified_once = || {
            file.set(
                "/frame[1]",
                props! { "x" => "1cm", "y" => "2cm", "width" => "3cm", "height" => "4cm" },
            )
            .expect("unified")
        };
        assert_eq!(
            unified_once().to_bytes().expect("bytes"),
            leaf_once().as_bytes()
        );
        let mut leaf = || {
            black_box(leaf_once());
        };
        let mut unified = || {
            black_box(unified_once());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "odi", "geometry_set", &times);
    }

    pub(super) fn run() {
        let mut rows = Vec::new();
        // `--single` runs only the single-operation scenarios.
        if !std::env::args().any(|argument| argument == "--single") {
            odg_batches(&mut rows);
            odb_batches(&mut rows);
            formula_batches(&mut rows);
        }
        odg_single(&mut rows);
        family_metadata(&mut rows);
        odc_and_odi(&mut rows);
        println!(
            "{}",
            serde_json::to_string(&serde_json::json!({ "rows": rows })).expect("json")
        );
    }
}

fn main() {
    #[cfg(all(
        feature = "ods",
        feature = "odp",
        feature = "odg",
        feature = "odc",
        feature = "odi",
        feature = "odb",
        feature = "odf-formula"
    ))]
    bench::run();
}
