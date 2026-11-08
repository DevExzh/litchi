//! Text-set timing for record 0304: the unified `set … text=` on the
//! containers whose text set now keeps or checks the non-text content
//! inside them, and the leaf ODT call it wraps.
//!
//! Each scenario times one call per iteration, `INNER` iterations per
//! sample, nine samples after two warm-ups, and reports the median in µs
//! per call. ODT scenarios also time the leaf `Edit::set_text` + commit in
//! ABBA order against the unified call. Built at two revisions, the same
//! harness gives the before/after comparison of the record.
//!
//! - `odt_plain`: a 1,000-paragraph builder document, the middle paragraph
//!   (character data only: the keep plan's fast path);
//! - `odt_frames`: `writer-images-frames.odt` `/body/p[2]`, which holds a
//!   picture frame the set keeps;
//! - `pptx_shape`, `pptx_paragraph`, `pptx_cell`: a builder deck's text
//!   box, its first paragraph and a table cell;
//! - `ppt_shape`: `text_shapes.ppt` `/slide[1]/shape[3]`;
//! - `docx_cell`: a two-paragraph table cell;
//! - `odp_shape`: `impress-basic.odp` `/slide[1]/shape[1]`.
//!
//! Run pinned, release:
//! `taskset -c 9 cargo run --release -p litchi-ops --features odt,pptx,ppt,docx,odp --example text_set_bench -- out.json`.

#[cfg(all(
    feature = "odt",
    feature = "pptx",
    feature = "ppt",
    feature = "docx",
    feature = "odp"
))]
mod bench {
    use std::hint::black_box;
    use std::io::{Read as _, Write as _};
    use std::time::Instant;

    use litchi_ops::{File, props};

    const ROUNDS: usize = 9;
    const WARMUP: usize = 2;
    const INNER: usize = 20;

    fn median(mut samples: Vec<f64>) -> f64 {
        samples.sort_by(f64::total_cmp);
        samples[samples.len() / 2]
    }

    fn sample(run: &mut dyn FnMut()) -> f64 {
        let start = Instant::now();
        for _ in 0..INNER {
            run();
        }
        start.elapsed().as_secs_f64() * 1e6 / INNER as f64
    }

    /// Medians of `leaf` and `unified` timed in ABBA order.
    fn paired(mut leaf: impl FnMut(), mut unified: impl FnMut()) -> (f64, f64) {
        for _ in 0..WARMUP {
            leaf();
            unified();
        }
        let (mut a, mut b) = (Vec::new(), Vec::new());
        for round in 0..ROUNDS {
            if round % 2 == 0 {
                a.push(sample(&mut leaf));
                b.push(sample(&mut unified));
            } else {
                b.push(sample(&mut unified));
                a.push(sample(&mut leaf));
            }
        }
        (median(a), median(b))
    }

    fn single(mut unified: impl FnMut()) -> f64 {
        for _ in 0..WARMUP {
            unified();
        }
        median((0..ROUNDS).map(|_| sample(&mut unified)).collect())
    }

    fn set_text(file: &File, path: &str, text: &str) {
        black_box(file.set(path, props! { "text" => text }).expect("set"));
    }

    fn data(relative: &str) -> Vec<u8> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        std::fs::read(root.join(relative)).expect("fixture")
    }

    fn odt(bytes: Vec<u8>, paragraph: usize, text: &str) -> (f64, f64) {
        let package = litchi_odt::Package::from_bytes(&bytes).expect("leaf open");
        let id = {
            let document = package.document().expect("document");
            let body = document.body().expect("body");
            body.root()
                .children()
                .filter(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
                .nth(paragraph)
                .expect("paragraph")
                .id()
        };
        let file = litchi_ops::open_bytes(bytes).expect("open");
        let path = format!("/body/p[{}]", paragraph + 1);
        paired(
            || {
                let mut edit = package.edit();
                edit.set_text(id, text).expect("stage");
                black_box(edit.commit().expect("commit").into_package());
            },
            || set_text(&file, &path, text),
        )
    }

    fn deck() -> Vec<u8> {
        let mut builder = litchi_pptx::builder::Builder::new().expect("builder");
        builder
            .presentation()
            .expect("presentation")
            .add_slide()
            .expect("slide")
            .add_text_box("alpha", 0, 0, 1_828_800, 914_400);
        let package = builder.build().expect("build");
        let mut edit = package.edit().expect("edit");
        edit.add_table(
            0usize,
            &[
                vec!["a1".to_owned(), "b1".to_owned()],
                vec!["a2".to_owned(), "b2".to_owned()],
            ],
            (0, 1_000_000, 2_000_000, 600_000),
        )
        .expect("table");
        edit.commit()
            .expect("commit")
            .into_package()
            .to_bytes()
            .expect("bytes")
    }

    /// `comment.docx` with a table whose cell holds two paragraphs.
    fn docx_cell() -> Vec<u8> {
        let bytes = data("ooxml/docx/comment.docx");
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("zip");
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).expect("entry");
            let name = entry.name().expect("valid ZIP member name").into_owned();
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).expect("read");
            if name == "word/document.xml" {
                let xml = String::from_utf8(contents).expect("UTF-8");
                let start = xml.find("<w:body>").expect("body") + "<w:body>".len();
                let end = xml.find("<w:sectPr>").expect("sectPr");
                let cell = "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>\
                    <w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>\
                    </w:tc></w:tr></w:tbl><w:p/>";
                contents = format!("{}{cell}{}", &xml[..start], &xml[end..]).into_bytes();
            }
            zip.start_file(name, options).expect("member");
            zip.write_all(&contents).expect("write");
        }
        zip.finish().expect("finish").into_inner()
    }

    pub(super) fn run() -> Result<(), Box<dyn std::error::Error>> {
        let mut rows = Vec::new();
        let mut report = |name: &str, leaf: Option<f64>, unified: f64| {
            match leaf {
                Some(leaf) => println!(
                    "{name:>16} leaf {leaf:>9.2} µs  unified {unified:>9.2} µs  ratio {:>6.3}",
                    unified / leaf
                ),
                None => println!("{name:>16}                     unified {unified:>9.2} µs"),
            }
            rows.push(serde_json::json!({
                "scenario": name,
                "leaf_us": leaf,
                "unified_us": unified,
            }));
        };

        let mut builder = litchi_odt::Builder::new();
        for index in 0..1_000 {
            builder.add_paragraph(&format!(
                "Paragraph {index} holds a few ordinary words of body text."
            ))?;
        }
        let (leaf, unified) = odt(builder.build()?.to_bytes()?, 500, "Replaced text");
        report("odt_plain", Some(leaf), unified);
        let (leaf, unified) = odt(
            data("odf/corpus/writer-images-frames.odt"),
            1,
            "Replaced text",
        );
        report("odt_frames", Some(leaf), unified);

        let deck = litchi_ops::open_bytes(deck())?;
        for (name, path) in [
            ("pptx_shape", "/slide[1]/shape[1]"),
            ("pptx_paragraph", "/slide[1]/shape[1]/paragraph[1]"),
            ("pptx_cell", "/slide[1]/table[1]/tr[2]/tc[1]"),
        ] {
            report(name, None, single(|| set_text(&deck, path, "omega")));
        }
        let ppt = litchi_ops::open_bytes(data("ole/ppt/text_shapes.ppt"))?;
        report(
            "ppt_shape",
            None,
            single(|| set_text(&ppt, "/slide[1]/shape[3]", "Square")),
        );
        let docx = litchi_ops::open_bytes(docx_cell())?;
        report(
            "docx_cell",
            None,
            single(|| set_text(&docx, "/body/tbl[1]/tr[1]/tc[1]", "New")),
        );
        let odp = litchi_ops::open_bytes(data("odf/corpus/impress-basic.odp"))?;
        report(
            "odp_shape",
            None,
            single(|| set_text(&odp, "/slide[1]/shape[1]", "New title")),
        );

        if let Some(out) = std::env::args().nth(1) {
            std::fs::write(out, serde_json::to_string_pretty(&rows)?)?;
        }
        Ok(())
    }
}

#[cfg(all(
    feature = "odt",
    feature = "pptx",
    feature = "ppt",
    feature = "docx",
    feature = "odp"
))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    bench::run()
}

#[cfg(not(all(
    feature = "odt",
    feature = "pptx",
    feature = "ppt",
    feature = "docx",
    feature = "odp"
)))]
fn main() {}
