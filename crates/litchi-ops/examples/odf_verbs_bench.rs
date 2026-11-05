//! Paired timing of record 0292: each leaf crate's own API against the
//! unified `litchi_ops` API on the same input.
//!
//! Scenarios, each timed as leaf and unified legs in ABBA order (per-call
//! means, medians over `ROUNDS`):
//!
//! - `odg cross_page_move`: the first of 100 shapes moved into a second
//!   page (`Edit::move_shape_to_page` against `File::move_to(.., Into)`);
//! - `odb query_move`: the first of 100 queries moved last;
//! - `odc series_swap`: two of 10 series of a typed chart exchanged;
//! - `odf element_move`: the first of 50 MathML tokens moved last;
//! - `fodt open`: a 54 KB flat text document opened (flat parse, packaged
//!   projection, document projection) against `open_bytes`;
//! - `fodt paragraph_set`: one paragraph's text through the packaged
//!   projection's edit, commit and flat republication against
//!   `File::set`.
//!
//! Inputs are built outside the clock; every write scenario checks outside
//! it that the unified and leaf outputs are byte-identical. Run pinned,
//! release:
//! `taskset -c 9 cargo run --release -p litchi-ops --features odt,odg,odc,odb,odf-formula --example odf_verbs_bench`.

#[cfg(all(
    feature = "odt",
    feature = "odg",
    feature = "odc",
    feature = "odb",
    feature = "odf-formula"
))]
mod bench {
    use std::hint::black_box;
    use std::time::Instant;

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

    fn row(rows: &mut Vec<serde_json::Value>, format: &str, scenario: &str, times: &[f64]) {
        let (leaf, unified) = (times[0], times[1]);
        println!(
            "{format:>5} {scenario:>17}  leaf {leaf:>10.1} µs  unified {unified:>10.1} µs  ratio {:.3}",
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

    fn into(path: &str) -> Placement<'_> {
        Placement::Into(litchi_ops::IntoPath::into_path(path).expect("path"), None)
    }

    /// Time `leaf` against `unified` after checking they publish the same
    /// bytes.
    fn write_pair(
        rows: &mut Vec<serde_json::Value>,
        format: &str,
        scenario: &str,
        calls: usize,
        leaf: &dyn Fn() -> Vec<u8>,
        unified: &dyn Fn() -> File,
    ) {
        assert_eq!(
            unified().to_bytes().expect("bytes"),
            leaf(),
            "{format} {scenario}: unified and leaf outputs differ"
        );
        let mut leaf_leg = || {
            black_box(leaf());
        };
        let mut unified_leg = || {
            black_box(unified());
        };
        let times = paired(calls, &mut [&mut leaf_leg, &mut unified_leg]);
        row(rows, format, scenario, &times);
    }

    fn odg(rows: &mut Vec<serde_json::Value>) {
        let package = litchi_odg::Package::new().expect("new");
        let mut edit = package.edit();
        edit.add_page(litchi_odg::page::Page::new("one"))
            .expect("page");
        edit.add_page(litchi_odg::page::Page::new("two"))
            .expect("page");
        let package = edit.commit().expect("commit").into_package();
        let mut edit = package.edit();
        for index in 0..100 {
            edit.add_shape(
                0,
                litchi_odg::shape::Shape::new(litchi_odg::shape::ShapeKind::Rectangle)
                    .with_geometry("1cm", "1cm", "2cm", "1cm")
                    .with_text(format!("shape {index}")),
            )
            .expect("shape");
        }
        let bytes = edit
            .commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes");
        let package = litchi_odg::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        let leaf = || {
            let mut edit = package.edit();
            edit.move_shape_to_page(0, 0, 1, 0).expect("stage");
            edit.commit()
                .expect("commit")
                .into_package()
                .as_bytes()
                .to_vec()
        };
        let unified = || {
            file.move_to("/page[1]/shape[1]", into("/page[2]"))
                .expect("unified")
        };
        write_pair(rows, "odg", "cross_page_move", 20, &leaf, &unified);
    }

    fn odb(rows: &mut Vec<serde_json::Value>) {
        let package = litchi_odb::Package::new().expect("new");
        let mut edit = package.edit();
        for index in 0..100 {
            edit.add_query(litchi_odb::Query::new(
                format!("q{index}"),
                format!("SELECT {index}"),
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
        let leaf = || {
            let mut edit = package.edit();
            edit.move_query("q0", 99).expect("stage");
            edit.commit()
                .expect("commit")
                .into_package()
                .as_bytes()
                .to_vec()
        };
        let unified = || file.move_to("/query[1]", Placement::End).expect("unified");
        write_pair(rows, "odb", "query_move", 20, &leaf, &unified);
    }

    fn odc(rows: &mut Vec<serde_json::Value>) {
        let mut file = litchi_ops::create(litchi_core::FileFormat::Odc).expect("create");
        for index in 0..10 {
            file = file
                .add(
                    "/",
                    NewNode::new("series").prop(
                        "values",
                        format!("t.{0}2:{0}9", char::from(b'B' + index)).as_str(),
                    ),
                )
                .expect("series")
                .file;
        }
        let bytes = file.to_bytes().expect("bytes");
        let package = litchi_odc::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        let leaf = || {
            let mut definition = package.chart().definition().expect("definition");
            definition.plot_area.series.swap(2, 7);
            let mut edit = package.edit();
            edit.replace_chart(&definition).expect("stage");
            edit.commit()
                .expect("commit")
                .into_package()
                .as_bytes()
                .to_vec()
        };
        let unified = || file.swap("/series[3]", "/series[8]").expect("unified");
        write_pair(rows, "odc", "series_swap", 20, &leaf, &unified);
    }

    fn formula(rows: &mut Vec<serde_json::Value>) {
        let mut file = litchi_ops::create(litchi_core::FileFormat::Odf).expect("create");
        let added = file
            .add("/", NewNode::new("element").prop("tag", "mrow"))
            .expect("row");
        let row_path = added.path.to_string();
        file = added.file;
        file = file
            .edit(|tx| {
                for index in 0..50 {
                    tx.add(
                        row_path.as_str(),
                        NewNode::new("element")
                            .prop("tag", "mi")
                            .prop("text", format!("x{index}").as_str()),
                    )?;
                }
                Ok(())
            })
            .expect("tokens")
            .file()
            .clone();
        let bytes = file.to_bytes().expect("bytes");
        let package = litchi_odf_formula::Package::from_bytes(&bytes).expect("leaf");
        let file = open(&bytes);
        let row_route = file.get(row_path.as_str()).expect("row").path().to_string();
        let route: Vec<usize> = {
            // The row is the last child of the root.
            let count = package.formula().root().children().count();
            vec![count - 1]
        };
        let first = format!("{row_route}/element[1]");
        let leaf = || {
            let mut edit = package.edit();
            let mut token = route.clone();
            token.push(0);
            let element = edit
                .remove(&litchi_odf_formula::NodePath::new(token))
                .expect("remove");
            edit.insert_child(
                &litchi_odf_formula::NodePath::new(route.clone()),
                49,
                element,
            )
            .expect("insert");
            edit.commit()
                .expect("commit")
                .into_package()
                .as_bytes()
                .to_vec()
        };
        let unified = || {
            file.move_to(first.as_str(), Placement::End)
                .expect("unified")
        };
        write_pair(rows, "odf", "element_move", 20, &leaf, &unified);
    }

    fn fixture(path: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(path),
        )
        .expect("fixture")
    }

    fn fodt(rows: &mut Vec<serde_json::Value>) {
        let bytes = fixture("odf/odt/libreoffice-page-columns-compact.fodt");
        let mut leaf = || {
            let flat = litchi_odt::flat::Document::from_bytes(&bytes).expect("flat");
            let package = flat.to_package().expect("projection");
            black_box(package.document().expect("document"));
            black_box(flat);
        };
        let mut unified = || {
            black_box(open(&bytes));
        };
        let times = paired(20, &mut [&mut leaf, &mut unified]);
        row(rows, "fodt", "open", &times);

        let flat = litchi_odt::flat::Document::from_bytes(&bytes).expect("flat");
        let package = flat.to_package().expect("projection");
        let document = package.document().expect("document");
        let body = document.body().expect("body");
        let (ordinal, id) = body
            .nodes()
            .filter(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
            .enumerate()
            .find(|(_, node)| !node.text().unwrap_or_default().is_empty())
            .map(|(ordinal, node)| (ordinal, node.id()))
            .expect("a paragraph with text");
        let file = open(&bytes);
        let path = file
            .query("p")
            .expect("query")
            .nth(ordinal)
            .expect("paragraph")
            .expect("node")
            .path()
            .to_string();
        let leaf = || {
            let mut edit = package.edit();
            edit.set_text(id, "Edited paragraph").expect("stage");
            let edited = edit.commit().expect("commit").into_package();
            flat.republish(&edited)
                .expect("republish")
                .as_bytes()
                .to_vec()
        };
        let unified = || {
            file.set(path.as_str(), props! { "text" => "Edited paragraph" })
                .expect("unified")
        };
        write_pair(rows, "fodt", "paragraph_set", 20, &leaf, &unified);
    }

    pub(super) fn run() {
        let mut rows = Vec::new();
        odg(&mut rows);
        odb(&mut rows);
        odc(&mut rows);
        formula(&mut rows);
        fodt(&mut rows);
        println!(
            "{}",
            serde_json::to_string(&serde_json::json!({ "rows": rows })).expect("json")
        );
    }
}

fn main() {
    #[cfg(all(
        feature = "odt",
        feature = "odg",
        feature = "odc",
        feature = "odb",
        feature = "odf-formula"
    ))]
    bench::run();
}
