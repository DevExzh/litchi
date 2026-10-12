//! Paired RTF run insertion timing; exact bytes and semantic checks are untimed.
#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::{Position, selector::Path};
    use litchi_ops::{NewNode, Placement, open_bytes};
    use litchi_rtf::edit::RunFormatPatch;
    use std::hint::black_box;
    use std::time::Instant;

    let mut cases = Vec::new();
    for shape in ["paragraphs", "dense_runs"] {
        for count in [0usize, 1, 10, 100, 1_000, 10_000] {
            let groups: Vec<_> = (0..count)
                .map(|i| format!(r"{{\b0 plain{i:05}}}{{\i italic{i:05}}}{{\b0 tail{i:05}}}"))
                .collect();
            let body = if count == 0 {
                r"\par".to_owned()
            } else {
                groups.join(if shape == "paragraphs" { r"\par " } else { "" })
            };
            let bytes = format!(r"{{\rtf1\ansi {body}}}").into_bytes();
            let native = litchi_rtf::Package::from_bytes(&bytes)?;
            let unified = open_bytes(bytes.as_slice())?;
            let paragraphs = if shape == "paragraphs" {
                count.max(1)
            } else {
                1
            };
            let paragraph = paragraphs - 1;
            let run_count = if count == 0 {
                0
            } else if shape == "paragraphs" {
                3
            } else {
                count * 3
            };
            let path = Path::parse(&format!("/body/p[{paragraphs}]"))?.into_owned();
            let patch = RunFormatPatch::new()
                .with_bold(true)
                .with_italic(true)
                .with_size(std::num::NonZeroU16::new(28).ok_or("zero size")?);
            let limits = litchi_rtf::SemanticLimits::builder()
                .max_operations(10_000)?
                .max_replacement_bytes(16 * 1024 * 1024)?
                .build()?;
            let options = litchi_rtf::EditOptions::new().semantic_limits(limits);
            for (route, position) in [("start", 0), ("middle", run_count / 2), ("end", run_count)] {
                let native_once = || -> Result<litchi_rtf::Package, litchi_rtf::edit::Error> {
                    let mut edit = native.edit_with_options(options);
                    edit.insert_run(
                        Position::new(paragraph),
                        Position::new(position),
                        "inserted",
                        patch,
                    )?;
                    Ok(edit.commit()?.into_package())
                };
                let expected = native_once()?.to_bytes()?;
                let mut native_ns = Vec::new();
                let mut unified_ns = Vec::new();
                for round in 0..11 {
                    for facade in [false, true, true, false] {
                        let node = NewNode::new("r")
                            .prop("text", "inserted")
                            .prop("bold", true)
                            .prop("italic", true)
                            .prop("size", "14pt")
                            .at(Placement::Position(Position::new(position)));
                        let start = Instant::now();
                        let (elapsed, output) = if facade {
                            let output = black_box(unified.add(&path, node)?);
                            let elapsed = start.elapsed().as_nanos();
                            assert_eq!(
                                output.path.to_string(),
                                format!("/body/p[{paragraphs}]/r[{}]", position + 1)
                            );
                            assert_eq!(
                                output.file.get(&output.path)?.view().text()?.as_deref(),
                                Some("inserted")
                            );
                            (elapsed, output.to_bytes()?)
                        } else {
                            let output = black_box(native_once()?);
                            let elapsed = start.elapsed().as_nanos();
                            (elapsed, output.to_bytes()?)
                        };
                        assert_eq!(output, expected);
                        let reopened = litchi_rtf::Package::from_bytes(&output)?;
                        assert_eq!(reopened.paragraph_count(), paragraphs);
                        for (p, para) in reopened.body().paragraphs().enumerate() {
                            let mut texts: Vec<String> = if count == 0 {
                                Vec::new()
                            } else {
                                let units = if shape == "paragraphs" {
                                    p..p + 1
                                } else {
                                    0..count
                                };
                                units
                                    .flat_map(|i| {
                                        [
                                            format!("plain{i:05}"),
                                            format!("italic{i:05}"),
                                            format!("tail{i:05}"),
                                        ]
                                    })
                                    .collect()
                            };
                            if p == paragraph {
                                texts.insert(position, "inserted".to_owned());
                            }
                            let runs: Vec<_> = para.runs().collect();
                            assert_eq!(runs.len(), texts.len());
                            for (run, text) in runs.iter().zip(&texts) {
                                assert_eq!(run.text(), text);
                                let inserted = text == "inserted";
                                assert_eq!(run.format().bold(), inserted);
                                assert_eq!(
                                    run.format().italic(),
                                    inserted || text.starts_with("italic")
                                );
                                assert_eq!(
                                    run.format().size().get(),
                                    if inserted { 28 } else { 24 }
                                );
                            }
                        }
                        if round >= 2 {
                            if facade {
                                unified_ns.push(elapsed);
                            } else {
                                native_ns.push(elapsed);
                            }
                        }
                    }
                }
                assert_eq!(native.to_bytes()?, bytes);
                assert_eq!(unified.to_bytes()?, bytes);
                let median = |values: &[u128]| {
                    let mut sorted = values.to_vec();
                    sorted.sort_unstable();
                    (sorted[sorted.len() / 2 - 1] + sorted[sorted.len() / 2]) as f64 / 2.0
                };
                cases.push(serde_json::json!({"shape":shape,"groups":count,"paragraphs":paragraphs,
                    "source_runs":count*3,"route":route,"native_ns":native_ns,"unified_ns":unified_ns,
                    "native_median_ns":median(&native_ns),"unified_median_ns":median(&unified_ns),
                    "ratio":median(&unified_ns)/median(&native_ns),"exact_bytes_created_path_and_reopened_runs":true}));
            }
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema":"rtf_run_add_benchmark.v1","rounds":9,"warmup_rounds":2,"order":"ABBA",
            "timing":"warmed immutable owners; selectors, props and native patch prepared outside timing; staging and commit timed; serialization, output validation and drop excluded",
            "cases":cases
        }))?
    );
    Ok(())
}
#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
