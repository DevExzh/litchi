//! Before/after timing for the hot paths record 0289's fixes touch.
//!
//! - `docx_edit_bold_1k`: one `File::edit` with 1,000 direct-path run
//!   `bold` sets, the coalesced DOCX formatting run whose plan now lets
//!   chart targets and hyperlink targets through to the ordinary setter;
//! - `doc_edit_text_200`: one `File::edit` with 200 direct-path DOC
//!   paragraph text sets (the DOC body transaction's operation bound keeps
//!   the generated input below 1,000 paragraphs), the coalesced DOC run whose plan now passes
//!   non-paragraph targets on;
//! - `odt_metadata_set_publish`: a LibreOffice ODT (`writer-definition-
//!   lists.odt`) `set /metadata author` plus `to_bytes`, the ODF rewrite
//!   that now copies manifest directory entries in path order.
//!
//! `PHASE_G_SCENARIO=<name>` runs one scenario (for callgrind). Inputs are
//! built outside the clock. Each scenario prints the median and
//! minimum of `ROUNDS` timed rounds (after `WARMUP`) as one JSON line.
//! Build one binary per tree and alternate them ABBA, pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example phase_g_bench
//! --features docx,doc,odt`.

#[cfg(all(feature = "docx", feature = "doc", feature = "odt"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::{NewNode, props};

    const ROUNDS: usize = 15;
    const WARMUP: usize = 3;

    let only = std::env::var("PHASE_G_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);

    fn report(name: &str, mut samples: Vec<f64>) {
        samples.sort_by(f64::total_cmp);
        let median = samples[samples.len() / 2];
        println!(
            "{}",
            serde_json::json!({ "scenario": name, "median_us": median, "min_us": samples[0], "rounds": samples.len() })
        );
    }

    fn time(mut run: impl FnMut()) -> Vec<f64> {
        let mut samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let started = Instant::now();
            run();
            if round >= WARMUP {
                samples.push(started.elapsed().as_secs_f64() * 1e6);
            }
        }
        samples
    }

    if wanted("docx_edit_bold_1k") {
        // DOCX: 1,000 paragraphs, each with one run.
        let mut docx = litchi_ops::create(FileFormat::Docx)?;
        docx = docx
            .edit(|tx| {
                for index in 0..1000 {
                    tx.add(
                        "/body",
                        NewNode::new("paragraph").prop("text", format!("Paragraph {index}")),
                    )?;
                }
                Ok(())
            })?
            .file()
            .clone();
        let docx = litchi_ops::open_bytes(docx.to_bytes()?)?;
        let paths: Vec<String> = (1..=1000).map(|n| format!("/body/p[{n}]/r[1]")).collect();
        report(
            "docx_edit_bold_1k",
            time(|| {
                let commit = docx
                    .edit(|tx| {
                        for path in &paths {
                            tx.set(path.as_str(), props! { "bold" => true })?;
                        }
                        Ok(())
                    })
                    .expect("docx edit");
                black_box(commit);
            }),
        );
    }
    if wanted("doc_edit_text_200") {
        // DOC: 200 paragraphs.
        let mut doc = litchi_ops::create(FileFormat::Doc)?;
        doc = doc
            .edit(|tx| {
                for index in 0..200 {
                    tx.add(
                        "/body",
                        NewNode::new("paragraph").prop("text", format!("Paragraph {index}")),
                    )?;
                }
                Ok(())
            })?
            .file()
            .clone();
        let doc = litchi_ops::open_bytes(doc.to_bytes()?)?;
        let paths: Vec<String> = (1..=200).map(|n| format!("/body/p[{n}]")).collect();
        report(
            "doc_edit_text_200",
            time(|| {
                let commit = doc
                    .edit(|tx| {
                        for path in &paths {
                            tx.set(path.as_str(), props! { "text" => "Changed" })?;
                        }
                        Ok(())
                    })
                    .expect("doc edit");
                black_box(commit);
            }),
        );
    }
    if wanted("odt_metadata_set_publish") {
        // ODT: a LibreOffice package with manifest directory entries.
        let odt = litchi_ops::open_bytes(std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/odf/corpus/writer-definition-lists.odt"
        ))?)?;
        report(
            "odt_metadata_set_publish",
            time(|| {
                let bytes = odt
                    .set("/metadata", props! { "author" => "Litchi" })
                    .expect("odt set")
                    .to_bytes()
                    .expect("odt publish");
                black_box(bytes);
            }),
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "docx", feature = "doc", feature = "odt")))]
fn main() {
    eprintln!("enable the docx, doc and odt features");
}
