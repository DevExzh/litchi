//! Paired leaf-vs-unified measurements for record 0279: field refresh.
//!
//! ```text
//! cargo build --release -p litchi-ops --features docx,xlsx,eval --example refresh_bench
//! taskset -c 5 target/release/examples/refresh_bench [headings] [rounds] > refresh.json
//! ```
//!
//! Each scenario runs its two legs in ABBA order for `rounds` rounds
//! (default 7) and reports every sample, the median and the minimum in
//! microseconds, and the unified/leaf ratio of the medians. Both legs'
//! results are checked equal outside the clock.
//!
//! * `toc-<headings>`: a DOCX holding OfficeCLI's placeholder table of
//!   contents (`TOC \o "1-2" \h`) and `headings` (default 1,000) headings
//!   alternating `Heading1` and `Heading2`; one refresh builds every entry
//!   and bookmarks every heading, without serialization. The leaf leg opens
//!   the document as a `litchi_docx::Package` once and, per round, edits it,
//!   reads the field index, plans the entries knowing the fixture (style id
//!   `HeadingN` is level N) and commits. The unified leg is `File::refresh`
//!   on an eager handle (it also reads the style part and plans every
//!   field).
//! * `toc-<headings>-current`: the same refresh on the refreshed document,
//!   an exact no-op for both legs (the leaf's rewrite finds its result
//!   unchanged; the unified planner finds the entries current and makes no
//!   edit).
//! * `xlsx-cache-6k`: a workbook with 2,000 values and 6,000 uncached
//!   formulas; one refresh with `write_cache` writes every result. The leaf
//!   leg is `File::recalculate` (the engine both legs use) and a hand-written
//!   `litchi_xlsx` edit calling `set_cached` per formula and committing;
//!   the unified leg is `File::refresh`.
//!
//! `REFRESH_BENCH_PROFILE=1` runs only the unified `toc-<headings>-current`
//! leg, `rounds` times, and prints nothing: a profiler target.

use std::io::{Cursor, Write as _};
use std::time::Instant;

use litchi_docx::document::fields::{TocEntry, Update};
use litchi_ops::refresh::Options;
use litchi_ops::{File, OpenOptions, SourceMode, props};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn docx(headings: usize) -> Vec<u8> {
    let mut body = String::from(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-2" \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Update field to see table of contents</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    for index in 0..headings {
        let level = 1 + index % 2;
        body.push_str(&format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r><w:t>Section {index}</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body text of section {index}, long enough to be a paragraph.</w:t></w:r></w:p>"#
        ));
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>"#
    );
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", relationships),
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", styles.as_str()),
    ] {
        zip.start_file(name, options).expect("member");
        zip.write_all(contents.as_bytes()).expect("write");
    }
    zip.finish().expect("finish").into_inner()
}

/// The hand-written leaf refresh: the fixture's headings are the
/// paragraphs styled `HeadingN`; every heading gets a `_Toc` bookmark.
fn leaf_toc(package: &litchi_docx::Package) -> litchi_docx::Package {
    let mut edit = package.edit().expect("edit");
    let fields = edit.fields().expect("fields");
    let width = fields.text_width().unwrap_or(9360);
    let mut updates = Vec::new();
    let mut entries = Vec::new();
    let mut name = 100_000_000u64;
    let anchors: std::collections::HashMap<usize, &str> = fields
        .bookmarks()
        .iter()
        .filter(|bookmark| bookmark.name().starts_with("_Toc"))
        .filter_map(|bookmark| Some((bookmark.paragraph()?, bookmark.name())))
        .collect();
    for (index, paragraph) in fields.paragraphs().iter().enumerate() {
        let Some(level) = paragraph
            .style()
            .and_then(|style| style.strip_prefix("Heading"))
            .and_then(|level| level.parse::<u8>().ok())
        else {
            continue;
        };
        let anchor = if let Some(existing) = anchors.get(&index) {
            (*existing).to_owned()
        } else {
            name += 1;
            let anchor = format!("_Toc{name}");
            updates.push(Update::Bookmark {
                paragraph: index,
                name: anchor.clone(),
            });
            anchor
        };
        entries.push(
            TocEntry::new(level, paragraph.text())
                .anchor(anchor)
                .hyperlink(true)
                .page("")
                .tab(width),
        );
    }
    updates.push(Update::Toc {
        field: 0,
        entries,
        dirty: true,
    });
    edit.refresh_fields(&fields, &updates).expect("refresh");
    edit.commit().expect("commit").into_package()
}

fn workbook() -> Vec<u8> {
    let file = litchi_ops::create(litchi_core::FileFormat::Xlsx).expect("create");
    let commit = file
        .edit(|edit| {
            for row in 1..=2_000i64 {
                edit.set(
                    format!("/Sheet1/A{row}").as_str(),
                    props! { "value" => row },
                )?;
                for (column, formula) in [
                    ("B", format!("A{row}*2")),
                    ("C", format!("B{row}+1")),
                    ("D", format!("SUM(A{row},C{row})")),
                ] {
                    edit.set(
                        format!("/Sheet1/{column}{row}").as_str(),
                        props! { "formula" => formula.as_str() },
                    )?;
                }
            }
            Ok(())
        })
        .expect("workbook");
    commit.file().to_bytes().expect("bytes")
}

fn time(mut leg: impl FnMut()) -> f64 {
    let start = Instant::now();
    leg();
    start.elapsed().as_secs_f64() * 1e6
}

fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        f64::midpoint(samples[middle - 1], samples[middle])
    } else {
        samples[middle]
    }
}

fn abba(rounds: usize, items: usize, mut leaf: impl FnMut(), mut unified: impl FnMut()) -> Value {
    leaf();
    unified();
    let (mut a, mut b) = (Vec::new(), Vec::new());
    for _ in 0..rounds {
        a.push(time(&mut leaf));
        b.push(time(&mut unified));
        b.push(time(&mut unified));
        a.push(time(&mut leaf));
    }
    let (leaf_samples, unified_samples) = (a.clone(), b.clone());
    let (leaf_median, unified_median) = (median(&mut a), median(&mut b));
    #[allow(clippy::cast_precision_loss, reason = "item counts are small")]
    let per_second = |micros: f64| items as f64 / (micros / 1e6);
    json!({
        "leaf_us": {"median": leaf_median, "min": a[0], "samples": leaf_samples},
        "unified_us": {"median": unified_median, "min": b[0], "samples": unified_samples},
        "leaf_per_second": per_second(leaf_median),
        "unified_per_second": per_second(unified_median),
        "ratio_of_medians": unified_median / leaf_median,
    })
}

fn document_xml(file: &File) -> Vec<u8> {
    let bytes = file.to_bytes().expect("bytes");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let mut member = archive.by_name("word/document.xml").expect("document");
    let mut xml = Vec::new();
    std::io::copy(&mut member, &mut xml).expect("read");
    xml
}

fn main() {
    let mut args = std::env::args().skip(1);
    let headings: usize = args
        .next()
        .map_or(1_000, |value| value.parse().expect("headings"));
    let rounds: usize = args
        .next()
        .map_or(7, |value| value.parse().expect("rounds"));
    let mut results = serde_json::Map::new();
    let eager = OpenOptions::new().source_mode(SourceMode::Eager);

    // A table of contents built over `headings` headings.
    let source = docx(headings);
    let package = litchi_docx::Package::from_bytes(&source).expect("leaf document");
    let file = litchi_ops::open_bytes_with(source.clone(), &eager).expect("document");
    let options = Options::new();
    {
        let leaf = leaf_toc(&package);
        let unified = file.refresh(&options).expect("refresh");
        let leaf_file = litchi_ops::open_bytes_with(leaf.to_bytes().expect("leaf bytes"), &eager)
            .expect("leaf file");
        assert_eq!(
            document_xml(&leaf_file),
            document_xml(unified.file()),
            "both legs write the same document"
        );
    }
    results.insert(
        format!("toc-{headings}"),
        abba(
            rounds,
            headings,
            || {
                std::hint::black_box(leaf_toc(&package));
            },
            || {
                std::hint::black_box(file.refresh(&options).expect("refresh"));
            },
        ),
    );

    // The refreshed table again: an exact no-op.
    let current = file.refresh(&options).expect("refresh").into_file();
    let current =
        litchi_ops::open_bytes_with(current.to_bytes().expect("bytes"), &eager).expect("current");
    assert!(
        current
            .refresh(&options)
            .expect("refresh")
            .patch()
            .is_noop()
    );
    if std::env::var_os("REFRESH_BENCH_PROFILE").is_some() {
        for _ in 0..rounds {
            std::hint::black_box(current.refresh(&options).expect("refresh"));
        }
        return;
    }
    let current_package = litchi_docx::Package::from_bytes(&current.to_bytes().expect("bytes"))
        .expect("leaf current");
    results.insert(
        format!("toc-{headings}-current"),
        abba(
            rounds,
            headings,
            || {
                std::hint::black_box(leaf_toc(&current_package));
            },
            || {
                std::hint::black_box(current.refresh(&options).expect("refresh"));
            },
        ),
    );

    // 6,000 formula caches written back.
    let book = workbook();
    let file = litchi_ops::open_bytes_with(book.clone(), &eager).expect("workbook");
    let leaf_package = litchi_xlsx::Package::from_bytes(&book).expect("leaf workbook");
    let write = Options::new().write_cache(true);
    let leaf_cache = || {
        let report = file
            .recalculate(&litchi_ops::calc::Options::new())
            .expect("recalculate");
        let mut edit = leaf_package.edit().expect("edit");
        let mut sheet = edit.sheet("Sheet1").expect("sheet").expect("worksheet");
        for formula in report.formulas() {
            let Some(litchi_ops::calc::Value::Number(number)) = formula.computed() else {
                panic!("the fixture computes numbers");
            };
            let address = format!(
                "{}{}",
                char::from(b'A' + u8::try_from(formula.column()).expect("column")),
                formula.row() + 1
            );
            let value = litchi_xlsx::Number::new(format!("{number}")).expect("number");
            sheet.set_cached(address.as_str(), value).expect("cache");
        }
        edit.commit().expect("commit").into_package()
    };
    {
        let leaf = leaf_cache();
        let unified = file.refresh(&write).expect("refresh");
        assert_eq!(unified.value().refreshed(), 6_000);
        assert_eq!(
            leaf.to_bytes().expect("leaf bytes"),
            unified.file().to_bytes().expect("unified bytes"),
            "both legs write the same workbook"
        );
    }
    results.insert(
        "xlsx-cache-6k".to_owned(),
        abba(
            rounds,
            6_000,
            || {
                std::hint::black_box(leaf_cache());
            },
            || {
                std::hint::black_box(file.refresh(&write).expect("refresh"));
            },
        ),
    );
    println!(
        "{}",
        serde_json::to_string_pretty(&Value::Object(results)).expect("json")
    );
}
