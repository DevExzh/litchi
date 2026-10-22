//! Record 0362 (W8, the unified `File`): the `File` routes of record 0351
//! §5 for every format family, timed on one build so the same source runs
//! on the wave's base and head (binary-level A-B-B-A).
//!
//! The probe uses only `litchi_ops` API that is the same before and after
//! the wave (`open_bytes`, `get`, `set`, `add`, `batch`, `import`, `merge`;
//! commits are dropped, never read), and builds its inputs outside the
//! clock with the leaf builders or the unified `add`/`set`:
//!
//! - `<fmt>/open_get`: `open_bytes` + one `get`;
//! - `<fmt>/set`: one `File::set` on an open file;
//! - `<fmt>/batch1k`: one atomic `File::batch` of 1,000 `set` items (DOC:
//!   256, its body transaction's default operation bound);
//! - `<fmt>/import`: `File::import` of a 1,000-row, four-column CSV
//!   (workbook formats; XLS: 100 rows);
//! - `<fmt>/merge`: `File::merge` of one record into a one-placeholder
//!   template.
//!
//! `LITCHI_PROBE_ONLY=<prefix>` runs only the scenarios starting with it.
//! Output: one JSON array of `{scenario, samples, median_us, min_us}` on
//! stdout. Run pinned, release:
//! `taskset -c 29 cargo run --release -p litchi-ops --features
//! docx,xlsx,xlsb,pptx,doc,xls,ppt,rtf,odt,ods,odp --example file_w8_probe -- [rounds]`.

#[cfg(all(
    feature = "docx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "pptx",
    feature = "doc",
    feature = "xls",
    feature = "ppt",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
mod probe {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::merge::{Options as MergeOptions, Record};
    use litchi_ops::{BatchItem, BatchMode, File, NewNode, Props, open_bytes, props};

    const PARAGRAPHS: usize = 1_000;
    const SHAPES: usize = 100;

    fn selected(scenario: &str) -> bool {
        std::env::var("LITCHI_PROBE_ONLY").map_or(true, |prefix| scenario.starts_with(&prefix))
    }

    struct Probe {
        rounds: usize,
        rows: Vec<serde_json::Value>,
    }

    impl Probe {
        fn time(&mut self, scenario: &str, calls: usize, mut run: impl FnMut()) {
            if !selected(scenario) {
                return;
            }
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
            eprintln!("{scenario:>24} {median:>12.1} µs");
            self.rows.push(serde_json::json!({
                "scenario": scenario,
                "samples": samples.len(),
                "median_us": median,
                "min_us": samples[0],
            }));
        }
    }

    fn line(index: usize) -> String {
        format!("Paragraph {index:05} holds a few ordinary words of body text.")
    }

    /// A word-processing file of `count` paragraphs through the unified
    /// `add` (DOCX, ODT); the first paragraph holds the merge placeholder.
    fn unified_words(format: FileFormat, count: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(format).expect("create");
        for index in 0..count {
            let text = if index == 0 {
                "Dear {{name}},".to_owned()
            } else {
                line(index)
            };
            file = file
                .add(
                    "/body",
                    NewNode::new("paragraph").prop("text", text.as_str()),
                )
                .expect("paragraph")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    fn doc(count: usize) -> Vec<u8> {
        let mut builder = litchi_doc::Builder::new();
        builder.add_paragraph("Dear {{name}},").expect("paragraph");
        for index in 1..count {
            builder.add_paragraph(&line(index)).expect("paragraph");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn rtf(count: usize) -> Vec<u8> {
        let mut body = vec![r"Dear \{\{name\}\},".to_owned()];
        body.extend((1..count).map(line));
        format!(r"{{\rtf1\ansi {}}}", body.join(r"\par ")).into_bytes()
    }

    /// A workbook whose first sheet holds A1:J100 and the placeholder in
    /// L1, through the unified `set`.
    fn workbook(format: FileFormat) -> Vec<u8> {
        litchi_ops::create(format)
            .expect("create")
            .set("/sheet[1]/A1:J100", props! { "value" => 7 })
            .expect("range")
            .set("/sheet[1]/L1", props! { "value" => "Invoice for {{name}}" })
            .expect("placeholder")
            .to_bytes()
            .expect("bytes")
    }

    fn pptx(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_pptx::builder::Builder::new().expect("builder");
        let slide = builder
            .presentation()
            .expect("presentation")
            .add_slide()
            .expect("slide");
        for index in 0..shapes {
            let text = if index == 0 {
                "Dear {{name}}".to_owned()
            } else {
                format!("text-{index}")
            };
            slide.add_text_box(
                &text,
                i64::try_from(index).unwrap_or(0) * 100,
                0,
                914_400,
                914_400,
            );
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn ppt(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_ppt::writer::Builder::new();
        let slide = builder.add_slide().expect("slide");
        for index in 0..shapes {
            let text = if index == 0 {
                "Dear {{name}}"
            } else {
                "before"
            };
            builder
                .add_textbox(slide, 10 + (index % 100) as i32, 10, 240, 40, text)
                .expect("textbox");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn odp(shapes: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(FileFormat::Odp).expect("create");
        for index in 0..shapes {
            let text = if index == 0 {
                "Dear {{name}}".to_owned()
            } else {
                format!("text-{index}")
            };
            file = file
                .add(
                    "/slide[1]",
                    NewNode::new("shape")
                        .prop("text", text.as_str())
                        .prop("width", "2cm")
                        .prop("height", "1cm"),
                )
                .expect("shape")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    fn sets(paths: &[String], count: usize, props: &Props<'static>) -> Vec<BatchItem<'static>> {
        (0..count)
            .map(|index| {
                let mut item = BatchItem::new("set");
                item.path = Some(paths[index % paths.len()].clone().into());
                item.props = Some(props.clone());
                item
            })
            .collect()
    }

    fn csv(rows: usize) -> Vec<u8> {
        let mut text = String::from("name,amount,paid,note\n");
        for row in 0..rows {
            text.push_str(&format!(
                "row {row},{}.5,{},n{row}\n",
                row * 3,
                row % 2 == 0
            ));
        }
        text.into_bytes()
    }

    struct Family {
        name: &'static str,
        bytes: Vec<u8>,
        paths: Vec<String>,
        set_props: Props<'static>,
        batch: usize,
        /// CSV rows of the import route; zero for no import.
        import: usize,
    }

    fn families() -> Vec<Family> {
        let paragraphs: Vec<String> = (1..=PARAGRAPHS)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        let cells: Vec<String> = (1..=100)
            .flat_map(|row| {
                ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
                    .into_iter()
                    .map(move |column| format!("/sheet[1]/{column}{row}"))
            })
            .collect();
        let shapes: Vec<String> = (1..=SHAPES)
            .map(|index| format!("/slide[1]/shape[{index}]"))
            .collect();
        let text = Props::new().with("text", "Replaced text");
        let value = Props::new().with("value", 42);
        let word = |name, bytes, batch| Family {
            name,
            bytes,
            paths: paragraphs.clone(),
            set_props: text.clone(),
            batch,
            import: 0,
        };
        let sheet = |name, format, import| Family {
            name,
            bytes: workbook(format),
            paths: cells.clone(),
            set_props: value.clone(),
            batch: 1_000,
            import,
        };
        let slides = |name, bytes| Family {
            name,
            bytes,
            paths: shapes.clone(),
            set_props: text.clone(),
            batch: 1_000,
            import: 0,
        };
        vec![
            word("docx", unified_words(FileFormat::Docx, PARAGRAPHS), 1_000),
            word("odt", unified_words(FileFormat::Odt, PARAGRAPHS), 1_000),
            word("doc", doc(PARAGRAPHS), 256),
            word("rtf", rtf(PARAGRAPHS), 1_000),
            sheet("xlsx", FileFormat::Xlsx, 1_000),
            sheet("xlsb", FileFormat::Xlsb, 1_000),
            // The XLS import route takes ~13 s for 1,000 rows at base and
            // head alike (record 0362 §5), so it is timed on 100 rows.
            sheet("xls", FileFormat::Xls, 100),
            sheet("ods", FileFormat::Ods, 1_000),
            slides("pptx", pptx(SHAPES)),
            slides("ppt", ppt(SHAPES)),
            slides("odp", odp(SHAPES)),
        ]
    }

    fn scenario(family: &Family, route: &str) -> String {
        format!("{}/{route}", family.name)
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
        let mut record = Record::new();
        record.insert("name", "Acme");
        for family in families() {
            let first = family.paths[0].as_str();
            let middle = family.paths[family.paths.len() / 2].as_str();
            probe.time(&scenario(&family, "open_get"), 5, || {
                let file = open_bytes(family.bytes.clone()).expect("open");
                black_box(file.get(first).expect("get").path().to_string().len());
            });
            let file: File = open_bytes(family.bytes.clone()).expect("open");
            probe.time(&scenario(&family, "set"), 5, || {
                black_box(file.set(middle, family.set_props.clone()).expect("set"));
            });
            let items = sets(&family.paths, family.batch, &family.set_props);
            probe.time(&scenario(&family, "batch1k"), 1, || {
                black_box(file.batch(items.clone(), BatchMode::Atomic).expect("batch"));
            });
            if family.import > 0 {
                let csv = csv(family.import);
                let options = litchi_ops::import::Options::csv();
                probe.time(&scenario(&family, "import"), 1, || {
                    black_box(file.import("/sheet[1]/N1", &csv, &options).expect("import"));
                });
            }
            let options = MergeOptions::new();
            probe.time(&scenario(&family, "merge"), 1, || {
                black_box(file.merge(&record, &options).expect("merge"));
            });
        }
        serde_json::to_writer(std::io::stdout().lock(), &probe.rows).expect("write JSON");
        println!();
    }
}

fn main() {
    #[cfg(all(
        feature = "docx",
        feature = "xlsx",
        feature = "xlsb",
        feature = "pptx",
        feature = "doc",
        feature = "xls",
        feature = "ppt",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    ))]
    probe::run();
}
