//! DOCX set/add timing for record 0307: the paths run-level markup
//! compatibility touched (the text-owner and story scanners, the paragraph
//! compactor), against the leaf calls they wrap, and the newly admitted
//! documents.
//!
//! Each scenario times one call per iteration, `INNER` iterations per
//! sample, nine samples after two warm-ups, and reports the median in µs
//! per call; leaf and unified legs alternate in ABBA order. Built at two
//! revisions, the same harness gives the record's before/after comparison.
//!
//! - `text_set`: the middle paragraph of a 1,000-paragraph document, the
//!   leaf `Edit::replace_paragraph_text` + commit against the unified
//!   `set … text=`;
//! - `add_paragraph`: a paragraph inserted at the middle, the leaf
//!   `Edit::insert_paragraph` + commit against the unified `add`;
//! - `run_bold`: the unified run property set on the middle paragraph's
//!   first run (the changed paragraph is compacted);
//! - `find`: a find/replace over the whole body (the story scanner);
//! - `mc_text_set`, `mc_add`: the same unified calls on
//!   `do-not-capture-draw-objs-on-page-draw-wrap-none.docx`, which holds a
//!   run-level text box (refused before record 0307, so after only).
//!
//! Run pinned, release:
//! `taskset -c 9 cargo run --release -p litchi-ops --features docx --example docx_mc_bench -- out.json`.

#[cfg(feature = "docx")]
mod bench {
    use std::hint::black_box;
    use std::io::Write as _;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ops::{BatchItem, File, NewNode, Placement, Props, props};

    const ROUNDS: usize = 9;
    const WARMUP: usize = 2;
    const INNER: usize = 20;
    const PARAGRAPHS: usize = 1_000;
    const MIDDLE: usize = 500;

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

    /// A package of `PARAGRAPHS` paragraphs of two runs each.
    fn document() -> Vec<u8> {
        let mut body = String::new();
        for index in 0..PARAGRAPHS {
            body.push_str(&format!(
                "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Paragraph {index} holds </w:t></w:r><w:r><w:t>a few ordinary words of body text.</w:t></w:r></w:p>"
            ));
        }
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
        );
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
        let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, contents) in [
            ("[Content_Types].xml", content_types),
            ("_rels/.rels", root),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, options).expect("member");
            zip.write_all(contents.as_bytes()).expect("write");
        }
        zip.finish().expect("finish").into_inner()
    }

    fn data(relative: &str) -> Vec<u8> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        std::fs::read(root.join(relative)).expect("fixture")
    }

    fn add(file: &File, position: usize) {
        black_box(
            file.add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Inserted")
                    .at(Placement::Position(Position::new(position))),
            )
            .expect("add"),
        );
    }

    pub(super) fn run() -> Result<(), Box<dyn std::error::Error>> {
        let mut rows = Vec::new();
        let mut report = |name: &str, leaf: Option<f64>, unified: f64| {
            match leaf {
                Some(leaf) => println!(
                    "{name:>14} leaf {leaf:>9.2} µs  unified {unified:>9.2} µs  ratio {:>6.3}",
                    unified / leaf
                ),
                None => println!("{name:>14}                     unified {unified:>9.2} µs"),
            }
            rows.push(serde_json::json!({
                "scenario": name,
                "leaf_us": leaf,
                "unified_us": unified,
            }));
        };

        let bytes = document();
        let package = litchi_docx::Package::from_reader(std::io::Cursor::new(bytes.clone()))?;
        let file = litchi_ops::open_bytes(bytes)?;
        let path = format!("/body/p[{}]", MIDDLE + 1);

        let (leaf, unified) = paired(
            || {
                let mut edit = package.edit().expect("edit");
                edit.replace_paragraph_text(Position::new(MIDDLE), "Replaced text")
                    .expect("stage");
                black_box(edit.commit().expect("commit").into_package());
            },
            || {
                black_box(file.set(path.as_str(), props! { "text" => "Replaced text" }))
                    .map(drop)
                    .expect("set")
            },
        );
        report("text_set", Some(leaf), unified);

        let (leaf, unified) = paired(
            || {
                let mut edit = package.edit().expect("edit");
                edit.insert_paragraph(Position::new(MIDDLE), "Inserted")
                    .expect("stage");
                black_box(edit.commit().expect("commit").into_package());
            },
            || add(&file, MIDDLE),
        );
        report("add_paragraph", Some(leaf), unified);

        let run = format!("{path}/r[1]");
        report(
            "run_bold",
            None,
            single(|| {
                black_box(
                    file.set(run.as_str(), props! { "bold" => true })
                        .expect("set"),
                );
            }),
        );

        let mut item = BatchItem::new("set");
        item.path = Some("/".into());
        item.props = Some(
            Props::new()
                .with("find", "ordinary")
                .with("replace", "plain"),
        );
        report(
            "find",
            None,
            single(|| {
                black_box(file.apply(&item).expect("find"));
            }),
        );

        // Refused before record 0307: timed only where it is admitted.
        let alternate = litchi_ops::open_bytes(data(
            "libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx",
        ))?;
        if alternate
            .set("/body/p[28]", props! { "text" => "Changed" })
            .is_ok()
        {
            report(
                "mc_text_set",
                None,
                single(|| {
                    black_box(
                        alternate
                            .set("/body/p[28]", props! { "text" => "Changed" })
                            .expect("set"),
                    );
                }),
            );
            report("mc_add", None, single(|| add(&alternate, 0)));
        }

        if let Some(out) = std::env::args().nth(1) {
            std::fs::write(out, serde_json::to_string_pretty(&rows)?)?;
        }
        Ok(())
    }
}

#[cfg(feature = "docx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    bench::run()
}

#[cfg(not(feature = "docx"))]
fn main() {}
