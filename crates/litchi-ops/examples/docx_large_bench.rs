//! Large-document DOCX single-op timing (record 0305): the middle paragraph
//! of a 1,000-paragraph document (record 0307's fixture shape), the leaf
//! `Edit::replace_paragraph_text` / `insert_paragraph` + commit against the
//! unified `set` / `add` on a default (`SourceMode::Auto`, source-backed)
//! owner and on an eager owner, and (record 0311, `committed_*`) on a handle
//! already committed from the default owner, and (record 0313, `fresh_*`) the
//! first add on a handle opened for that call, so the overlay audits the
//! original and the insertion as a pair. Each leg: two warm-ups, nine samples of
//! `DOCX_LARGE_LOOPS` calls (default 20), the median in µs per call.
//! `DOCX_LARGE_ONLY=<leg>` runs one leg; `DOCX_LARGE_PROFILE=1` loops it
//! once for a profiler.
//!
//! `taskset -c 30 cargo run --release -p litchi-ops --features docx --example docx_large_bench`

#[cfg(feature = "docx")]
mod bench {
    use litchi_core::Position;
    use litchi_ops::{NewNode, OpenOptions, Placement, SourceMode, props};
    use std::hint::black_box;
    use std::io::Write as _;
    use std::time::Instant;

    fn document(n: usize) -> Vec<u8> {
        let mut body = String::new();
        for index in 0..n {
            body.push_str(&format!("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Paragraph {index} holds </w:t></w:r><w:r><w:t>a few ordinary words of body text.</w:t></w:r></w:p>"));
        }
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
        );
        let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
        let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let o = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, c) in [
            ("[Content_Types].xml", ct),
            ("_rels/.rels", root),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, o).unwrap();
            zip.write_all(c.as_bytes()).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }
    fn time(n: usize, f: &mut dyn FnMut()) -> f64 {
        for _ in 0..2 {
            f();
        }
        let mut v: Vec<f64> = (0..9)
            .map(|_| {
                let s = Instant::now();
                for _ in 0..n {
                    f();
                }
                s.elapsed().as_secs_f64() * 1e6 / n as f64
            })
            .collect();
        v.sort_by(f64::total_cmp);
        v[4]
    }
    pub(super) fn main() {
        let only = std::env::var("DOCX_LARGE_ONLY").ok();
        let loops: usize = std::env::var("DOCX_LARGE_LOOPS")
            .ok()
            .map_or(20, |v| v.parse().unwrap());
        let bytes = document(1000);
        let package =
            litchi_docx::Package::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
        let auto = litchi_ops::open_bytes(bytes.clone()).unwrap();
        let eager = litchi_ops::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .unwrap();
        let path = "/body/p[501]";
        // A handle already committed from the default owner (record 0311):
        // its next edit reads the effective revision, not the physical one.
        let committed = auto
            .set("/body/p[2]", props! {"text" => "Committed earlier"})
            .unwrap();
        let mut legs: Vec<(&str, Box<dyn FnMut()>)> = vec![
            (
                "leaf_set",
                Box::new(|| {
                    let mut e = package.edit().unwrap();
                    e.replace_paragraph_text(Position::new(500), "Replaced text")
                        .unwrap();
                    black_box(e.commit().unwrap().into_package());
                }),
            ),
            (
                "auto_set",
                Box::new(|| {
                    black_box(auto.set(path, props! {"text" => "Replaced text"}).unwrap());
                }),
            ),
            (
                "eager_set",
                Box::new(|| {
                    black_box(eager.set(path, props! {"text" => "Replaced text"}).unwrap());
                }),
            ),
            (
                "committed_set",
                Box::new(|| {
                    black_box(
                        committed
                            .set(path, props! {"text" => "Replaced text"})
                            .unwrap(),
                    );
                }),
            ),
            (
                "leaf_add",
                Box::new(|| {
                    let mut e = package.edit().unwrap();
                    e.insert_paragraph(Position::new(500), "Inserted").unwrap();
                    black_box(e.commit().unwrap().into_package());
                }),
            ),
            (
                "auto_add",
                Box::new(|| {
                    black_box(
                        auto.add(
                            "/body",
                            NewNode::new("paragraph")
                                .prop("text", "Inserted")
                                .at(Placement::Position(Position::new(500))),
                        )
                        .unwrap(),
                    );
                }),
            ),
            (
                "eager_add",
                Box::new(|| {
                    black_box(
                        eager
                            .add(
                                "/body",
                                NewNode::new("paragraph")
                                    .prop("text", "Inserted")
                                    .at(Placement::Position(Position::new(500))),
                            )
                            .unwrap(),
                    );
                }),
            ),
            (
                "committed_add",
                Box::new(|| {
                    black_box(
                        committed
                            .add(
                                "/body",
                                NewNode::new("paragraph")
                                    .prop("text", "Inserted")
                                    .at(Placement::Position(Position::new(500))),
                            )
                            .unwrap(),
                    );
                }),
            ),
            (
                "fresh_auto_add",
                Box::new(|| {
                    let file = litchi_ops::open_bytes(bytes.clone()).unwrap();
                    black_box(
                        file.add(
                            "/body",
                            NewNode::new("paragraph")
                                .prop("text", "Inserted")
                                .at(Placement::Position(Position::new(500))),
                        )
                        .unwrap(),
                    );
                }),
            ),
            (
                "fresh_eager_add",
                Box::new(|| {
                    let file = litchi_ops::open_bytes_with(
                        bytes.clone(),
                        &OpenOptions::new().source_mode(SourceMode::Eager),
                    )
                    .unwrap();
                    black_box(
                        file.add(
                            "/body",
                            NewNode::new("paragraph")
                                .prop("text", "Inserted")
                                .at(Placement::Position(Position::new(500))),
                        )
                        .unwrap(),
                    );
                }),
            ),
        ];
        for (name, f) in legs.iter_mut() {
            if only.as_deref().is_some_and(|o| o != *name) {
                continue;
            }
            if std::env::var("DOCX_LARGE_PROFILE").is_ok() {
                for _ in 0..loops {
                    f();
                }
                println!("{name} looped {loops}");
                continue;
            }
            println!("{name} {:.1} us", time(loops, f.as_mut()));
        }
    }
}

fn main() {
    #[cfg(feature = "docx")]
    bench::main();
    #[cfg(not(feature = "docx"))]
    eprintln!("docx_large_bench needs the docx feature");
}
