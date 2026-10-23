//! Paired timing of the remaining ODF formats (record 0282): each leaf
//! crate's own API against the unified `litchi_ops` API on the same input.
//!
//! Scenarios per format, each timed as leaf and unified legs in ABBA order
//! (per-call means of `CALLS` calls per sample, medians over `ROUNDS`):
//!
//! - `open_first_get`: open from bytes and read the first element the format
//!   is about (paragraph text, shape name, axis name, frame name, table
//!   name, first MathML token);
//! - `single_set`: change one value and commit (paragraph text, shape name,
//!   axis name, frame name, query command, token text).
//!
//! OTH and ODM are read and written through the ODT body engine
//! (`litchi_odt`), so they get a third leg, `engine`: the same operation
//! written by hand against `litchi_odt`, which is the unified layer's own
//! overhead. Inputs are built outside the clock and outputs are checked
//! outside it. Run pinned, release:
//! `taskset -c 9 cargo run --release -p litchi-ops --features oth,odm,odg,odc,odi,odb,odf-formula --example odf_leaves_bench`.

#[cfg(all(
    feature = "oth",
    feature = "odm",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "odf-formula"
))]
mod bench {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::props;

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

    fn row(rows: &mut Vec<serde_json::Value>, format: &str, scenario: &str, legs: &[(&str, f64)]) {
        let leaf = legs[0].1;
        let mut line = format!("{format:>4} {scenario:>15}");
        let mut value = serde_json::json!({ "format": format, "scenario": scenario });
        for (name, micros) in legs {
            line.push_str(&format!("  {name} {micros:>10.1} µs"));
            value[format!("{name}_us")] = serde_json::json!(micros);
        }
        let unified = legs
            .iter()
            .find(|(name, _)| *name == "unified")
            .map_or(0.0, |leg| leg.1);
        line.push_str(&format!("  ratio {:.3}", unified / leaf));
        value["ratio_unified_leaf"] = serde_json::json!(unified / leaf);
        if let Some((_, engine)) = legs.iter().find(|(name, _)| *name == "engine") {
            line.push_str(&format!("  unified/engine {:.3}", unified / engine));
            value["ratio_unified_engine"] = serde_json::json!(unified / engine);
        }
        println!("{line}");
        rows.push(value);
    }

    fn oth(rows: &mut Vec<serde_json::Value>) {
        let mut builder = litchi_oth::Builder::new();
        for index in 0..1_000 {
            builder = builder.paragraph(litchi_oth::paragraph::Paragraph::new(format!(
                "Paragraph {index} holds a few ordinary words of body text."
            )));
        }
        let bytes = builder.build().expect("build").to_bytes().expect("bytes");
        let mut leaf = || {
            let package = litchi_oth::Package::from_bytes(&bytes).expect("leaf");
            let body = package.document().text_body().expect("body");
            black_box(body.paragraphs()[0].text().len());
        };
        let mut engine = || {
            let package = litchi_odt::Package::from_bytes(&bytes).expect("engine");
            let document = package.document().expect("document");
            let body = document.body().expect("body");
            let first = body.root().children().next().expect("paragraph");
            black_box(first.text().expect("text").len());
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            let node = file.get("/body/p[1]").expect("get");
            black_box(node.view().text().expect("text").map(|text| text.len()));
        };
        let times = paired(20, &mut [&mut leaf, &mut engine, &mut unified]);
        row(
            rows,
            "oth",
            "open_first_get",
            &[
                ("leaf", times[0]),
                ("engine", times[1]),
                ("unified", times[2]),
            ],
        );

        let leaf_package = litchi_oth::Package::from_bytes(&bytes).expect("leaf");
        let engine_package = litchi_odt::Package::from_bytes(&bytes).expect("engine");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut leaf = || {
            let mut edit = leaf_package.edit();
            edit.set_paragraph_text(Position::new(499), "Replaced text")
                .expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut engine = || {
            let document = engine_package.document().expect("document");
            let body = document.body().expect("body");
            let id = body.root().children().nth(499).expect("paragraph").id();
            let mut edit = engine_package.edit();
            edit.set_text(id, "Replaced text").expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/body/p[500]", props! { "text" => "Replaced text" })
                    .expect("set"),
            );
        };
        let times = paired(10, &mut [&mut leaf, &mut engine, &mut unified]);
        row(
            rows,
            "oth",
            "single_set",
            &[
                ("leaf", times[0]),
                ("engine", times[1]),
                ("unified", times[2]),
            ],
        );
    }

    fn odm(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("libreoffice-core/vcl/qa/cppunit/pdfexport/data/master.odm");
        let mut leaf = || {
            let package = litchi_odm::Package::from_bytes(&bytes).expect("leaf");
            black_box(package.document().content_xml().len());
        };
        let mut engine = || {
            let package = litchi_odt::Package::from_bytes(&bytes).expect("engine");
            let document = package.document().expect("document");
            let body = document.body().expect("body");
            let first = body
                .nodes()
                .find(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
                .expect("paragraph");
            black_box(first.text().expect("text").len());
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            let node = file.get("/body/p[1]").expect("get");
            black_box(node.view().text().expect("text").map(|text| text.len()));
        };
        let times = paired(50, &mut [&mut leaf, &mut engine, &mut unified]);
        row(
            rows,
            "odm",
            "open_first_get",
            &[
                ("leaf", times[0]),
                ("engine", times[1]),
                ("unified", times[2]),
            ],
        );

        let engine_package = litchi_odt::Package::from_bytes(&bytes).expect("engine");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut engine = || {
            let document = engine_package.document().expect("document");
            let body = document.body().expect("body");
            let id = body
                .root()
                .children()
                .find(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
                .expect("paragraph")
                .id();
            let mut edit = engine_package.edit();
            edit.set_text(id, "Replaced text").expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/body/p[1]", props! { "text" => "Replaced text" })
                    .expect("set"),
            );
        };
        // The ODM leaf has no paragraph text edit: the engine leg is the
        // leaf the adapter writes through.
        let times = paired(20, &mut [&mut engine, &mut unified]);
        row(
            rows,
            "odm",
            "single_set",
            &[("engine", times[0]), ("unified", times[1])],
        );
    }

    fn odg(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("odf/native-resave/source/rhbz1870501.odg");
        let mut leaf = || {
            let package = litchi_odg::Package::from_bytes(&bytes).expect("leaf");
            black_box(
                package.drawing().pages()[0].shapes()[0]
                    .name()
                    .map(str::len),
            );
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            let node = file.get("/page[1]/shape[1]").expect("get");
            black_box(node.path().to_string().len());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odg",
            "open_first_get",
            &[("leaf", times[0]), ("unified", times[1])],
        );

        let package = litchi_odg::Package::from_bytes(&bytes).expect("leaf");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut leaf = || {
            let mut edit = package.edit();
            edit.set_shape_name(0, 0, "Renamed").expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/page[1]/shape[1]", props! { "name" => "Renamed" })
                    .expect("set"),
            );
        };
        let times = paired(10, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odg",
            "single_set",
            &[("leaf", times[0]), ("unified", times[1])],
        );
    }

    fn odc(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("odf/odc-producer-evidence/odfdom-created.odc");
        let mut leaf = || {
            let package = litchi_odc::Package::from_bytes(&bytes).expect("leaf");
            let axis = package
                .chart()
                .plot_area()
                .expect("plot")
                .axes()
                .next()
                .expect("axis");
            black_box(axis.name().map(str::len));
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            black_box(file.get("/axis[1]").expect("get").path().to_string().len());
        };
        let times = paired(50, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odc",
            "open_first_get",
            &[("leaf", times[0]), ("unified", times[1])],
        );

        let package = litchi_odc::Package::from_bytes(&bytes).expect("leaf");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut leaf = || {
            let mut edit = package.edit();
            edit.update_axis(0, litchi_odc::AxisUpdate::named("horizontal"))
                .expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/axis[1]", props! { "name" => "horizontal" })
                    .expect("set"),
            );
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odc",
            "single_set",
            &[("leaf", times[0]), ("unified", times[1])],
        );
    }

    fn odi(rows: &mut Vec<serde_json::Value>) {
        let bytes = litchi_odi::Package::new()
            .expect("new")
            .to_bytes()
            .expect("bytes");
        let mut leaf = || {
            let package = litchi_odi::Package::from_bytes(&bytes).expect("leaf");
            black_box(package.image().frames()[0].name().map(str::len));
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            black_box(file.get("/frame[1]").expect("get").path().to_string().len());
        };
        let times = paired(50, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odi",
            "open_first_get",
            &[("leaf", times[0]), ("unified", times[1])],
        );

        let package = litchi_odi::Package::from_bytes(&bytes).expect("leaf");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut leaf = || {
            let mut edit = package.edit();
            edit.set_frame_name(0, Some("Logo".to_owned()))
                .expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/frame[1]", props! { "name" => "Logo" })
                    .expect("set"),
            );
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odi",
            "single_set",
            &[("leaf", times[0]), ("unified", times[1])],
        );
    }

    fn odb(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb");
        let mut leaf = || {
            let package = litchi_odb::Package::from_bytes(&bytes).expect("leaf");
            let catalog = package.database().catalog().expect("catalog");
            black_box(catalog.tables()[0].name().len());
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            black_box(file.get("/table[1]").expect("get").path().to_string().len());
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odb",
            "open_first_get",
            &[("leaf", times[0]), ("unified", times[1])],
        );

        let package = litchi_odb::Package::from_bytes(&bytes).expect("leaf");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let mut leaf = || {
            let mut edit = package.edit();
            edit.set_query_command("AliasTest", "SELECT 1")
                .expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(
                file.set("/query[1]", props! { "command" => "SELECT 1" })
                    .expect("set"),
            );
        };
        let times = paired(10, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odb",
            "single_set",
            &[("leaf", times[0]), ("unified", times[1])],
        );
    }

    fn formula(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("odf/native-resave/source/font-styles.odf");
        let path = [0_usize, 0, 0];
        let mut leaf = || {
            let package = litchi_odf_formula::Package::from_bytes(&bytes).expect("leaf");
            let root = package.formula().root();
            let first = root.children().next().and_then(|row| row.children().next());
            black_box(first.map(litchi_odf_formula::Element::all_text));
        };
        let mut unified = || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
            black_box(
                file.get("/element[1]/element[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        };
        let times = paired(50, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odf",
            "open_first_get",
            &[("leaf", times[0]), ("unified", times[1])],
        );

        let package = litchi_odf_formula::Package::from_bytes(&bytes).expect("leaf");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified");
        let target = "/element[1]/element[1]/element[1]";
        let mut leaf = || {
            let mut edit = package.edit();
            edit.set_text(&litchi_odf_formula::NodePath::new(path), "g")
                .expect("stage");
            let _ = black_box(edit.commit().expect("commit"));
        };
        let mut unified = || {
            black_box(file.set(target, props! { "text" => "g" }).expect("set"));
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(
            rows,
            "odf",
            "single_set",
            &[("leaf", times[0]), ("unified", times[1])],
        );
    }

    pub(crate) fn run() {
        let mut rows = Vec::new();
        oth(&mut rows);
        odm(&mut rows);
        odg(&mut rows);
        odc(&mut rows);
        odi(&mut rows);
        odb(&mut rows);
        formula(&mut rows);
        if let Ok(path) = std::env::var("ODF_LEAVES_BENCH_JSON") {
            std::fs::write(path, serde_json::to_string_pretty(&rows).expect("json"))
                .expect("write");
        }
    }
}

fn main() {
    #[cfg(all(
        feature = "oth",
        feature = "odm",
        feature = "odg",
        feature = "odc",
        feature = "odi",
        feature = "odb",
        feature = "odf-formula"
    ))]
    bench::run();
}
