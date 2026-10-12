//! Paired RTF run-removal timing with exact native/facade checks outside timing.
#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::Position;
    use litchi_core::selector::{Path, Scope};
    use litchi_ops::open_bytes;
    use litchi_ops::query::Query;

    use litchi_rtf::edit::RunTextReplacement;
    use std::hint::black_box;
    use std::time::Instant;

    let opaque = std::env::var_os("LITCHI_RTF_BENCH_OPAQUE").is_some();
    let mut cases = Vec::new();
    for shape in ["paragraphs", "dense_runs"] {
        for count in [1usize, 10, 100, 1_000, 10_000] {
            let texts: Vec<_> = (0..count)
                .map(|i| {
                    let anchor = if opaque {
                        format!(r"{{\*\vendor anchor{i:05}}}")
                    } else {
                        String::new()
                    };
                    format!(r"{{\b0 plain{i:05}}}{{\i italic{i:05}}}{anchor}{{\b0 tail{i:05}}}")
                })
                .collect();
            let bytes = format!(
                r"{{\rtf1\ansi {}}}",
                texts.join(if shape == "paragraphs" { r"\par " } else { "" })
            )
            .into_bytes();
            let native = litchi_rtf::Package::from_bytes(&bytes)?;
            let unified = open_bytes(bytes.as_slice())?;
            let paragraphs = if shape == "paragraphs" { count } else { 1 };
            let last_run = if shape == "paragraphs" {
                2
            } else {
                (count - 1) * 3 + 2
            };
            let path = Path::parse(&format!("/body/p[{paragraphs}]/r[{last_run}]"))?.into_owned();
            let query = Query::parse("r[italic=true]")?
                .with_scope(Scope::Path(Box::new(Path::parse("/body")?)));
            let limits = litchi_rtf::SemanticLimits::builder()
                .max_operations(10_000)?
                .max_replacement_bytes(16 * 1024 * 1024)?
                .build()?;
            let options = litchi_rtf::EditOptions::new().semantic_limits(limits);
            for all in [false, true] {
                let positions: Vec<_> = if all {
                    (0..count).collect()
                } else {
                    vec![count - 1]
                };
                let updates: Vec<_> = positions
                    .iter()
                    .map(|&p| {
                        RunTextReplacement::new(
                            Position::new(if shape == "paragraphs" { p } else { 0 }),
                            Position::new(if shape == "paragraphs" { 1 } else { p * 3 + 1 }),
                            "",
                        )
                    })
                    .collect();
                let native_once = || -> Result<litchi_rtf::Package, litchi_rtf::edit::Error> {
                    let mut edit = native.edit_with_options(options);
                    edit.replace_body_run_texts(&updates)?;
                    Ok(edit.commit()?.into_package())
                };
                let expected = native_once()?.to_bytes()?;
                let mut native_ns = Vec::new();
                let mut unified_ns = Vec::new();
                for round in 0..11 {
                    for route in [false, true, true, false] {
                        let selector = query.clone();
                        let start = Instant::now();
                        let (elapsed, output) = if route {
                            let output = black_box(if all {
                                unified.remove(selector)?
                            } else {
                                unified.remove(&path)?
                            });
                            let elapsed = start.elapsed().as_nanos();
                            (elapsed, output.to_bytes()?)
                        } else {
                            let output = black_box(native_once()?);
                            let elapsed = start.elapsed().as_nanos();
                            (elapsed, output.to_bytes()?)
                        };
                        assert_eq!(output, expected);
                        let reopened = litchi_rtf::Package::from_bytes(&output)?;
                        assert_eq!(reopened.paragraph_count(), paragraphs);
                        for (i, paragraph) in reopened.body().paragraphs().enumerate() {
                            let unit_text = |j| {
                                let middle = if all || j == count - 1 {
                                    String::new()
                                } else {
                                    format!("italic{j:05}")
                                };
                                format!("plain{j:05}{middle}tail{j:05}")
                            };
                            let expected_text = if shape == "paragraphs" {
                                unit_text(i)
                            } else {
                                (0..count).map(unit_text).collect::<String>()
                            };
                            assert_eq!(paragraph.to_text(), expected_text);
                            let runs: Vec<_> = paragraph.runs().collect();
                            let removed = if shape == "paragraphs" {
                                usize::from(all || i == count - 1)
                            } else {
                                updates.len()
                            };
                            assert_eq!(
                                runs.len(),
                                if shape == "paragraphs" {
                                    3 - removed
                                } else {
                                    count * 3 - removed
                                }
                            );
                            for run in runs {
                                let format = run.format();
                                assert_eq!(format.italic(), run.text().starts_with("italic"));
                                assert!(!format.bold());
                                assert!(!format.strike());
                                assert_eq!(format.size().get(), 24);
                                assert_eq!(format.underline(), litchi_rtf::text::Underline::None);
                            }
                        }
                        if opaque {
                            let nodes: Vec<_> = reopened.opaque().iter().collect();
                            assert_eq!(nodes.len(), count);
                            let mut offset = 0;
                            for (unit, node) in nodes.iter().enumerate() {
                                let middle_len = if all || unit == count - 1 { 0 } else { 11 };
                                offset += 10 + middle_len;
                                assert_eq!(node.anchor(), litchi_rtf::opaque::Anchor::Body(offset));
                                assert_eq!(
                                    node.source(),
                                    format!(r"{{\*\vendor anchor{unit:05}}}").as_bytes()
                                );
                                offset += 9 + usize::from(shape == "paragraphs");
                            }
                        }
                        if round >= 2 {
                            if route {
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
                cases.push(serde_json::json!({
                "shape":shape,"paragraphs":paragraphs,"runs":count*3,"units":count,"route":if all {"query_all"} else {"last_path"},"targets":updates.len(),
                "native_ns":native_ns,"unified_ns":unified_ns,
                "native_median_ns":median(&native_ns),"unified_median_ns":median(&unified_ns),
                "ratio":median(&unified_ns)/median(&native_ns),"exact_bytes_and_reopened_run_format":true
            }));
            }
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema":"rtf_run_remove_benchmark.v1","rounds":9,"warmup_rounds":2,"order":"ABBA",
            "timing":"warmed immutable owners; native updates and facade selectors prepared outside timing; staging and commit timed; serialization, readback and output drop excluded",
            "opaque_body_anchors":opaque,"cases":cases
        }))?
    );
    Ok(())
}
#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
