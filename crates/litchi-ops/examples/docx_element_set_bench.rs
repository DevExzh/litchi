//! Paired DOCX timing for path-addressed element writes (record 0251a).
//!
//! Each case times one write against the same immutable opened owner: the
//! leaf leg stages the `litchi_docx::document::Edit` element operation and
//! commits; the unified leg runs the `File` verb on the matching path,
//! including path resolution, position proofs and property admission.
//! Fixture construction, path parsing and output checks are outside the
//! clock. Rounds alternate ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example
//! docx_element_set_bench --features docx` (pin the CPU with `taskset`).

#[cfg(feature = "docx")]
use std::hint::black_box;
#[cfg(feature = "docx")]
use std::time::Instant;

#[cfg(feature = "docx")]
use litchi_core::Position;
#[cfg(feature = "docx")]
use litchi_docx::document::{ElementAnchor, ElementPath, ElementStep, NewTable};
#[cfg(feature = "docx")]
use litchi_docx::properties::Justification;
#[cfg(feature = "docx")]
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};
#[cfg(feature = "docx")]
use serde_json::json;

#[cfg(feature = "docx")]
const ROUNDS: usize = 9;
#[cfg(feature = "docx")]
const WARMUP: usize = 2;
#[cfg(feature = "docx")]
const ITERATIONS: usize = 64;

#[cfg(feature = "docx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = fixture(200, 20)?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )?;
    let cell_paragraph = ElementPath::new([
        ElementStep::Table(Position::new(0)),
        ElementStep::Row(Position::new(10)),
        ElementStep::Cell(Position::new(2)),
        ElementStep::Paragraph(Position::new(0)),
    ]);
    let cell = ElementPath::new([
        ElementStep::Table(Position::new(0)),
        ElementStep::Row(Position::new(10)),
        ElementStep::Cell(Position::new(2)),
    ]);
    let hyperlink_run = ElementPath::new([
        ElementStep::Paragraph(Position::new(100)),
        ElementStep::Hyperlink(Position::new(0)),
        ElementStep::Run(Position::new(0)),
    ]);
    let table = NewTable::new(3, 3)?;

    let mut cases = Vec::new();
    cases.push(measure(
        "cell_paragraph_align",
        || {
            let mut edit = package.edit()?;
            edit.format_paragraph_at(&cell_paragraph, |paragraph| {
                paragraph.justification = Some(Justification::Center);
            })?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/body/tbl[1]/tr[11]/tc[3]/p[1]",
                props! { "align" => "center" },
            )?)
        },
    )?);
    cases.push(measure(
        "cell_text",
        || {
            let mut edit = package.edit()?;
            edit.replace_cell_text_at(&cell, "revised")?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/tbl[1]/tr[11]/tc[3]", props! { "text" => "revised" })?),
    )?);
    cases.push(measure(
        "hyperlink_run_bold",
        || {
            let mut edit = package.edit()?;
            edit.format_run_at(&hyperlink_run, |run| run.bold = Some(true))?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/body/p[101]/r[2]", props! { "bold" => true })?),
    )?);
    cases.push(measure(
        "table_add_3x3",
        || {
            let mut edit = package.edit()?;
            edit.insert_table_at(&ElementPath::body(), ElementAnchor::End, &table)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body",
                    NewNode::new("table").prop("rows", 3).prop("cols", 3),
                )?
                .into_file())
        },
    )?);
    let body_paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(150))]);
    let link_run = litchi_docx::document::NewRun::new("site");
    cases.push(measure(
        "hyperlink_add_url",
        || {
            let mut edit = package.edit()?;
            edit.insert_hyperlink_at(
                &body_paragraph,
                None,
                &litchi_docx::document::HyperlinkTarget::Url("https://example.test/".to_owned()),
                &link_run,
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[151]",
                    NewNode::new("hyperlink")
                        .prop("url", "https://example.test/")
                        .prop("text", "site")
                        .prop("rStyle", "Hyperlink"),
                )?
                .into_file())
        },
    )?);
    let note = litchi_docx::document::NewComment::new("Ada", "Check this");
    cases.push(measure(
        "comment_add_paragraph",
        || {
            let mut edit = package.edit()?;
            edit.add_comment_at(&body_paragraph, None, &note)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    "/body/p[151]",
                    NewNode::new("comment")
                        .prop("author", "Ada")
                        .prop("text", "Check this"),
                )?
                .into_file())
        },
    )?);
    // A 4x2 PNG, as bytes for the leaf and as a data URI for the unified call.
    let png_uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC";
    let png = {
        use base64::Engine as _;
        base64::engine::general_purpose::STANDARD.decode(png_uri.split_once(',').ok_or("uri")?.1)?
    };
    cases.push(measure(
        "picture_add_png",
        || {
            let mut edit = package.edit()?;
            edit.insert_picture_at(
                &body_paragraph,
                None,
                &litchi_docx::document::NewPicture::new(png.clone()),
            )?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add("/body/p[151]", NewNode::new("picture").prop("src", png_uri))?
                .into_file())
        },
    )?);
    cases.push(measure(
        "list_style_bullet",
        || {
            let mut edit = package.edit()?;
            let list = edit.add_list(litchi_docx::numbering::ListKind::Bullet, None)?;
            edit.format_paragraph_at(&body_paragraph, |properties| {
                let mut numbering = litchi_docx::properties::Numbering::default();
                numbering.id = Some(list.num_id());
                numbering.level = Some(0);
                properties.numbering = Some(numbering);
            })?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file.set(
                "/body/p[151]",
                litchi_ops::props! { "listStyle" => "bullet" },
            )?)
        },
    )?);
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": "docx_element_set_benchmark.v1",
            "fixture": {
                "body_paragraphs": 200,
                "table_rows": 20,
                "table_columns": 5,
                "bytes": bytes.len(),
            },
            "rounds": ROUNDS,
            "warmup_rounds": WARMUP,
            "iterations_per_sample": ITERATIONS,
            "order": "ABBA",
            "timing": "eager owners opened once; each sample stages and commits one write from the same immutable owner; output checks excluded",
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(feature = "docx")]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(feature = "docx")]
fn measure(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<litchi_docx::Package>,
    mut unified: impl FnMut() -> Outcome<File>,
) -> Outcome<serde_json::Value> {
    // `DOCX_BENCH_CASE` restricts a run to one case, for profiling.
    if std::env::var("DOCX_BENCH_CASE").is_ok_and(|only| only != name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    // Output checks, outside the clock: both legs publish the same text.
    let leaf_text = leaf()?.document()?.text()?;
    let unified_text = match unified()? {
        File::Docx(handle) => handle.text()?,
        _ => return Err("the unified leg returned a non-DOCX handle".into()),
    };
    if leaf_text.replace('\n', "") != unified_text.replace('\n', "") {
        return Err(format!("{name}: the legs disagree").into());
    }
    let sample = |leg: &mut dyn FnMut() -> Outcome<()>| -> Outcome<f64> {
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            leg()?;
        }
        Ok(start.elapsed().as_nanos() as f64 / ITERATIONS as f64)
    };
    let mut leaf_leg = || -> Outcome<()> {
        black_box(leaf()?);
        Ok(())
    };
    let mut unified_leg = || -> Outcome<()> {
        black_box(unified()?);
        Ok(())
    };
    for _ in 0..WARMUP {
        sample(&mut leaf_leg)?;
        sample(&mut unified_leg)?;
    }
    let mut leaf_samples = Vec::new();
    let mut unified_samples = Vec::new();
    for round in 0..ROUNDS {
        if round % 2 == 0 {
            leaf_samples.push(sample(&mut leaf_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
        } else {
            unified_samples.push(sample(&mut unified_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
            leaf_samples.push(sample(&mut leaf_leg)?);
            unified_samples.push(sample(&mut unified_leg)?);
        }
    }
    let leaf_median = median(&mut leaf_samples);
    let unified_median = median(&mut unified_samples);
    Ok(json!({
        "case": name,
        "leaf_median_ns": leaf_median,
        "unified_median_ns": unified_median,
        "ratio": unified_median / leaf_median,
        "leaf_samples_ns": leaf_samples,
        "unified_samples_ns": unified_samples,
    }))
}

#[cfg(feature = "docx")]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        (samples[middle - 1] + samples[middle]) / 2.0
    } else {
        samples[middle]
    }
}

/// A body of `paragraphs` paragraphs (the middle one with a hyperlink) and
/// one `rows` x 5 table after the first paragraph.
#[cfg(feature = "docx")]
fn fixture(paragraphs: usize, rows: usize) -> Outcome<Vec<u8>> {
    use std::io::Write as _;
    let mut body = String::new();
    for index in 0..paragraphs {
        if index == 100 {
            body.push_str(r#"<w:p><w:r><w:t>Lead </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>"#);
        } else {
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index} with some ordinary text.</w:t></w:r></w:p>"
            ));
        }
        if index == 0 {
            body.push_str(r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>"#);
            for row in 0..rows {
                body.push_str("<w:tr>");
                for column in 0..5 {
                    body.push_str(&format!(
                        "<w:tc><w:p><w:r><w:t>R{row}C{column}</w:t></w:r></w:p></w:tc>"
                    ));
                }
                body.push_str("</w:tr>");
            }
            body.push_str("</w:tbl>");
        }
    }
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/" TargetMode="External"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", document_rels),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options)?;
        zip.write_all(contents.as_bytes())?;
    }
    Ok(zip.finish()?.into_inner())
}

#[cfg(not(feature = "docx"))]
fn main() {
    eprintln!("enable the docx feature");
}
