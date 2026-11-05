//! Paired leaf-vs-unified measurements for record 0274: template merge.
//!
//! ```text
//! cargo build --release -p litchi-ops --features docx --example merge_bench
//! taskset -c 5 target/release/examples/merge_bench [records] [rounds] > merge.json
//! ```
//!
//! Each scenario runs its two legs in ABBA order for `rounds` rounds (default
//! 7) and reports every sample, the median and the minimum in microseconds,
//! and the unified/leaf ratio of the medians.
//!
//! * `records/<mode>`: `records` (default 1,000) records merged one by one
//!   into a small DOCX template (six paragraphs, eight placeholders, one of
//!   them split across two runs), each output serialized to bytes. The leaf
//!   leg opens the template as a `litchi_docx::Package` once and, per
//!   record, edits it, reads the main story, finds `{{key}}` placeholders
//!   with a hand-written scan, substitutes, commits and serializes. The
//!   unified leg opens the template with `litchi_ops::open_bytes` (`auto`:
//!   source-backed; `eager`: `SourceMode::Eager`) once and calls
//!   `File::merge` and `File::to_bytes` per record.
//! * `placeholders-10k`: one record merged into a 2,000-paragraph DOCX
//!   holding 10,000 placeholders (every fifth split across two runs),
//!   without serialization: the leaf leg's edit, story, plan, substitution
//!   and commit against `File::merge` on an eager handle.
//!
//! `MERGE_BENCH_PROFILE=<scenario>/<leaf|unified>` runs only that leg,
//! `rounds` times, and prints nothing: a profiler target.

use std::collections::HashMap;
use std::io::{Cursor, Write as _};
use std::time::Instant;

use litchi_docx::document::substitute::Replacement;
use litchi_ops::merge::{Options, Record};
use litchi_ops::{File, OpenOptions, SourceMode};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn docx(body: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
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

fn small_template() -> Vec<u8> {
    docx(concat!(
        r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Invoice {{number}}</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t xml:space="preserve">Dear {{first}} </w:t></w:r><w:r><w:t>{{last}},</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t xml:space="preserve">Your order {{order}} of {{date}} totals </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>{{to</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>tal}}</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>Thank you for your business.</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>{{company}}</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:t>{{city}}</w:t></w:r></w:p>"#,
    ))
}

fn large_template() -> Vec<u8> {
    let mut body = String::new();
    for index in 0..2_000 {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">{index}: {{{{k{a}}}}} and {{{{k{b}}}}}, {{{{k{c}}}}}; {{{{k{d}}}}} then </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>{{{{k</w:t></w:r><w:r><w:t>{e}}}}} end</w:t></w:r></w:p>"#,
            a = index % 50,
            b = (index + 1) % 50,
            c = (index + 2) % 50,
            d = (index + 3) % 50,
            e = (index + 4) % 50,
        ));
    }
    docx(&body)
}

fn small_records(count: usize) -> Vec<Vec<(String, String)>> {
    (0..count)
        .map(|index| {
            vec![
                ("number".to_owned(), format!("INV-{index:05}")),
                ("first".to_owned(), format!("First{index}")),
                ("last".to_owned(), format!("Last{index}")),
                ("order".to_owned(), format!("{}", 100_000 + index)),
                ("date".to_owned(), "2026-10-13".to_owned()),
                (
                    "total".to_owned(),
                    format!("{}.{:02}", index * 3, index % 100),
                ),
                ("company".to_owned(), "Acme & Sons <Ltd>".to_owned()),
                ("city".to_owned(), "London".to_owned()),
            ]
        })
        .collect()
}

fn large_record() -> Vec<(String, String)> {
    (0..50)
        .map(|key| (format!("k{key}"), format!("value number {key}")))
        .collect()
}

/// A hand-written leaf merge: exact `{{key}}` placeholders, no white space,
/// over the main story.
fn leaf_merge(
    package: &litchi_docx::Package,
    values: &HashMap<&str, &str>,
) -> litchi_docx::Package {
    let mut edit = package.edit().expect("edit");
    let story = edit.story().expect("story");
    let mut replacements = Vec::new();
    for (index, paragraph) in story.paragraphs().iter().enumerate() {
        let text = paragraph.text();
        let mut from = 0;
        while let Some(start) = text[from..].find("{{").map(|offset| from + offset) {
            let Some(end) = text[start..].find("}}").map(|offset| start + offset + 2) else {
                break;
            };
            if let Some(value) = values.get(&text[start + 2..end - 2]) {
                if paragraph.is_text(start..end) {
                    replacements.push(Replacement::new(index, start..end, *value));
                }
            }
            from = end;
        }
    }
    edit.substitute(&story, &replacements).expect("substitute");
    edit.commit().expect("commit").into_package()
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

fn abba(
    name: &str,
    rounds: usize,
    items: usize,
    mut leaf: impl FnMut(),
    mut unified: impl FnMut(),
) -> Value {
    if let Ok(profile) = std::env::var("MERGE_BENCH_PROFILE") {
        if !profile.starts_with(&format!("{name}/")) {
            return Value::Null;
        }
        let leg: &mut dyn FnMut() = if profile.ends_with("/leaf") {
            &mut leaf
        } else {
            &mut unified
        };
        for _ in 0..rounds {
            leg();
        }
        return Value::Null;
    }
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

fn main() {
    let mut args = std::env::args().skip(1);
    let records: usize = args
        .next()
        .map_or(1_000, |value| value.parse().expect("records"));
    let rounds: usize = args
        .next()
        .map_or(7, |value| value.parse().expect("rounds"));
    let mut results = serde_json::Map::new();

    // Records into a small template, serialized.
    let template = small_template();
    let rows = small_records(records);
    let leaf_values: Vec<HashMap<&str, &str>> = rows
        .iter()
        .map(|row| {
            row.iter()
                .map(|(key, value)| (key.as_str(), value.as_str()))
                .collect()
        })
        .collect();
    let unified_records: Vec<Record> = rows
        .iter()
        .map(|row| {
            row.iter()
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect()
        })
        .collect();
    let package = litchi_docx::Package::from_bytes(&template).expect("leaf template");
    let leaf_leg = || {
        for values in &leaf_values {
            let merged = leaf_merge(&package, values);
            let bytes = merged.to_bytes().expect("leaf bytes");
            std::hint::black_box(bytes);
        }
    };
    {
        let probe = leaf_merge(&package, &leaf_values[0]);
        let unified = litchi_ops::open_bytes(template.clone()).expect("unified template");
        let merged = unified
            .merge(&unified_records[0], &Options::new())
            .expect("merge");
        assert_eq!(
            probe
                .document()
                .expect("leaf document")
                .text()
                .expect("leaf text"),
            match merged.file() {
                File::Docx(_) => merged
                    .file()
                    .text()
                    .expect("unified text")
                    .replace('\n', ""),
                _ => unreachable!(),
            },
            "both legs produce the same text"
        );
    }
    for (mode, options) in [
        ("auto", OpenOptions::new()),
        ("eager", OpenOptions::new().source_mode(SourceMode::Eager)),
    ] {
        let file = litchi_ops::open_bytes_with(template.clone(), &options).expect("template");
        let unified_leg = || {
            for record in &unified_records {
                let commit = file.merge(record, &Options::new()).expect("merge");
                let bytes = commit.file().to_bytes().expect("unified bytes");
                std::hint::black_box(bytes);
            }
        };
        let name = format!("records/{mode}");
        let result = abba(&name, rounds, records, leaf_leg, unified_leg);
        results.insert(name, result);
    }

    // 10,000 placeholders in one document, without serialization.
    let template = large_template();
    let row = large_record();
    let values: HashMap<&str, &str> = row
        .iter()
        .map(|(key, value)| (key.as_str(), value.as_str()))
        .collect();
    let record: Record = row.iter().cloned().collect();
    let package = litchi_docx::Package::from_bytes(&template).expect("leaf template");
    let file = litchi_ops::open_bytes_with(
        template.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("template");
    let summary = file
        .merge(&record, &Options::new())
        .expect("merge")
        .into_value();
    assert_eq!(summary.replacements(), 10_000);
    results.insert(
        "placeholders-10k".to_owned(),
        abba(
            "placeholders-10k",
            rounds,
            10_000,
            || {
                std::hint::black_box(leaf_merge(&package, &values));
            },
            || {
                std::hint::black_box(file.merge(&record, &Options::new()).expect("merge"));
            },
        ),
    );
    let output = json!({
        "records": records,
        "rounds": rounds,
        "template_bytes": small_template().len(),
        "large_template_bytes": template.len(),
        "scenarios": results,
    });
    println!("{}", serde_json::to_string_pretty(&output).expect("json"));
}
