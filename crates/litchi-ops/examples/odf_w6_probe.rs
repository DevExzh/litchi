//! Record 0358 (W6 ODF contract): the unified routes the wave touched, timed
//! on one build so the same source can run on the wave's base and head
//! (binary-level A-B-B-A).
//!
//! Every scenario uses only the `litchi_ops` API (stable across the wave)
//! and inputs built outside the clock with the leaf builders (unchanged by
//! the wave) or read from `test-data/`:
//!
//! - `<fmt>/open_get`: `open_bytes` + one `get` (ODT, ODS, ODP, ODG, ODC);
//! - `<fmt>/set`: one `File::set` on an open file (publishes a new file);
//! - `odt/batch1k`, `ods/batch1k`: one `File::edit` with 1,000 sets;
//! - `<flat>/open_get`: `open_bytes` + one `get` on a flat document
//!   (`.fodt`, `.fods`).
//!
//! Output: one JSON array of `{scenario, samples, median_us, min_us}` on
//! stdout. Run pinned, release:
//! `taskset -c 29 cargo run --release -p litchi-ops --features odt,ods,odp,odg,odc --example odf_w6_probe -- [rounds]`.

#[cfg(all(
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "odg",
    feature = "odc"
))]
mod probe {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::props;

    fn fixture(path: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(path),
        )
        .expect("fixture")
    }

    fn odt(paragraphs: usize) -> Vec<u8> {
        let mut builder = litchi_odt::Builder::new();
        for index in 0..paragraphs {
            builder
                .add_paragraph(&format!(
                    "Paragraph {index} holds a few ordinary words of body text."
                ))
                .expect("paragraph");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn ods(rows: usize, columns: usize) -> Vec<u8> {
        let mut sheet = litchi_ods::Sheet::new("Data").expect("sheet");
        for row in 0..rows {
            let mut cells = litchi_ods::Row::new();
            for column in 0..columns {
                let value = (row * columns + column) as f64;
                cells
                    .push_cell(litchi_ods::Cell::new(
                        litchi_ods::CellValue::Number(value),
                        value.to_string(),
                    ))
                    .expect("cell");
            }
            sheet.push_row(cells).expect("row");
        }
        let mut builder = litchi_ods::Builder::new();
        builder.add_sheet(sheet).expect("add sheet");
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn odp(slides: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(litchi_core::FileFormat::Odp).expect("create");
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
                        .prop("text", "slide body text")
                        .prop("width", "20cm")
                        .prop("height", "12cm"),
                )
                .expect("text box")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    fn fodt(paragraphs: usize) -> Vec<u8> {
        let mut body = String::new();
        for index in 0..paragraphs {
            body.push_str(&format!(
                "<text:p>Paragraph {index} holds a few ordinary words of body text.</text:p>"
            ));
        }
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text"><office:body><office:text>{body}</office:text></office:body></office:document>"#
        )
        .into_bytes()
    }

    fn fods(rows: usize, columns: usize) -> Vec<u8> {
        let mut body = String::new();
        for row in 0..rows {
            body.push_str("<table:table-row>");
            for column in 0..columns {
                let value = row * columns + column;
                body.push_str(&format!(
                    r#"<table:table-cell office:value-type="float" office:value="{value}"><text:p>{value}</text:p></table:table-cell>"#
                ));
            }
            body.push_str("</table:table-row>");
        }
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet><table:table table:name="Data">{body}</table:table></office:spreadsheet></office:body></office:document>"#
        )
        .into_bytes()
    }

    struct Probe {
        rounds: usize,
        rows: Vec<serde_json::Value>,
    }

    impl Probe {
        fn time(&mut self, scenario: &str, calls: usize, mut run: impl FnMut()) {
            for _ in 0..2 {
                run();
            }
            let mut samples = Vec::with_capacity(self.rounds);
            for _ in 0..self.rounds {
                let start = Instant::now();
                for _ in 0..calls {
                    run();
                }
                samples.push(start.elapsed().as_secs_f64() * 1e6 / calls as f64);
            }
            samples.sort_by(f64::total_cmp);
            let median = samples[samples.len() / 2];
            eprintln!("{scenario:>20} {median:>12.1} µs");
            self.rows.push(serde_json::json!({
                "scenario": scenario,
                "samples": samples.len(),
                "median_us": median,
                "min_us": samples[0],
            }));
        }
    }

    pub(super) fn run() {
        let rounds = std::env::args()
            .nth(1)
            .and_then(|value| value.parse().ok())
            .unwrap_or(15);
        let mut probe = Probe {
            rounds,
            rows: Vec::new(),
        };

        let bytes = odt(1_000);
        probe.time("odt/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/body/p[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("odt/set", 5, || {
            black_box(
                file.set("/body/p[500]", props! { "text" => "Replaced text" })
                    .expect("set"),
            );
        });
        let paths: Vec<String> = (1..=1_000)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        probe.time("odt/batch1k", 1, || {
            let commit = file
                .edit(|tx| {
                    for path in &paths {
                        tx.set(path.as_str(), props! { "bold" => true })?;
                    }
                    Ok(())
                })
                .expect("edit");
            black_box(commit.file());
        });

        let bytes = ods(100, 10);
        probe.time("ods/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(file.get("/Data/C5").expect("get").path().to_string().len());
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("ods/set", 5, || {
            black_box(
                file.set("/Data/C5", props! { "value" => 42.0 })
                    .expect("set"),
            );
        });
        let cells: Vec<String> = (1..=100)
            .flat_map(|row| {
                ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
                    .into_iter()
                    .map(move |column| format!("/Data/{column}{row}"))
            })
            .collect();
        probe.time("ods/batch1k", 1, || {
            let commit = file
                .edit(|tx| {
                    for cell in &cells {
                        tx.set(cell.as_str(), props! { "value" => 7.0 })?;
                    }
                    Ok(())
                })
                .expect("edit");
            black_box(commit.file());
        });

        let bytes = odp(20);
        probe.time("odp/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/slide[1]/shape[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("odp/set", 5, || {
            black_box(
                file.set("/slide[10]/shape[1]", props! { "text" => "Replaced" })
                    .expect("set"),
            );
        });

        let bytes = fixture("odf/native-resave/source/rhbz1870501.odg");
        probe.time("odg/open_get", 20, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/page[1]/shape[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("odg/set", 10, || {
            black_box(
                file.set("/page[1]/shape[1]", props! { "name" => "Renamed" })
                    .expect("set"),
            );
        });

        let bytes = fixture("odf/odc-producer-evidence/odfdom-created.odc");
        probe.time("odc/open_get", 50, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(file.get("/axis[1]").expect("get").path().to_string().len());
        });
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        probe.time("odc/set", 20, || {
            black_box(
                file.set("/axis[1]", props! { "name" => "horizontal" })
                    .expect("set"),
            );
        });

        let bytes = fodt(1_000);
        probe.time("fodt/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(
                file.get("/body/p[1]")
                    .expect("get")
                    .path()
                    .to_string()
                    .len(),
            );
        });
        let bytes = fods(100, 10);
        probe.time("fods/open_get", 5, || {
            let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
            black_box(file.get("/Data/C5").expect("get").path().to_string().len());
        });

        println!(
            "{}",
            serde_json::to_string_pretty(&probe.rows).expect("json")
        );
    }
}

fn main() {
    #[cfg(all(
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "odg",
        feature = "odc"
    ))]
    probe::run();
    #[cfg(not(all(
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "odg",
        feature = "odc"
    )))]
    eprintln!("enable the odt, ods, odp, odg and odc features");
}
