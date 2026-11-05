//! Paired native/unified borrowed RTF run reads and predicate traversal.
#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::selector::Path;
    use litchi_ops::query::{Plan, Query};
    use std::hint::black_box;
    use std::time::Instant;

    let mut cases = Vec::new();
    for count in [1usize, 10, 100, 1_000, 10_000] {
        let paragraphs: Vec<_> = (0..count)
            .map(|i| format!(r"\b0 plain{i:05}\b bold{i:05}\b0 tail{i:05}"))
            .collect();
        let bytes = format!(r"{{\rtf1\ansi {}}}", paragraphs.join(r"\par ")).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = litchi_ops::open_bytes(bytes.as_slice())?;
        let path = Path::parse(&format!("/body/p[{count}]/r[3]"))?.into_owned();
        let query = Plan::compile(&Query::parse("r[bold=true]")?, litchi_core::FileFormat::Rtf)?;
        for route in ["last_path", "query_all", "query_first"] {
            let native_once = || {
                if route == "last_path" {
                    let run = black_box(&native)
                        .body()
                        .paragraphs()
                        .nth(count - 1)
                        .and_then(|p| p.runs().nth(2));
                    run.map_or((0, 0), |run| (1, run.text().len()))
                } else {
                    black_box(&native)
                        .body()
                        .paragraphs()
                        .flat_map(|p| p.runs())
                        .filter(|run| run.format().bold())
                        .take(if route == "query_first" {
                            1
                        } else {
                            usize::MAX
                        })
                        .fold((0, 0), |(n, bytes), run| (n + 1, bytes + run.text().len()))
                }
            };
            let expected = native_once();
            let expected_texts: Vec<_> = native
                .body()
                .paragraphs()
                .flat_map(|p| p.runs())
                .filter(|run| run.format().bold())
                .map(|run| run.text())
                .collect();
            let actual_texts = unified
                .query(&query)?
                .map(|node| {
                    node.and_then(|node| {
                        node.view()
                            .text()
                            .map(|text| text.map(|text| text.into_owned()))
                    })
                })
                .collect::<litchi_ops::Result<Vec<_>>>()?;
            assert_eq!(
                actual_texts,
                expected_texts
                    .iter()
                    .map(|text| Some((*text).to_owned()))
                    .collect::<Vec<_>>()
            );
            let actual_last = unified
                .get(&path)?
                .view()
                .text()?
                .map(|text| text.into_owned());
            assert_eq!(actual_last, Some(format!("tail{:05}", count - 1)));
            let repetitions = if route == "query_first" || count <= 10 {
                1_000
            } else if count <= 100 {
                100
            } else {
                10
            };
            let mut native_ns = Vec::new();
            let mut unified_ns = Vec::new();
            for round in 0..11 {
                for facade in [false, true, true, false] {
                    let start = Instant::now();
                    let mut total = (0usize, 0usize);
                    for _ in 0..repetitions {
                        let value = if facade {
                            if route == "last_path" {
                                let node = black_box(&unified).get(black_box(&path))?;
                                node.view().text()?.map_or((0, 0), |text| (1, text.len()))
                            } else {
                                let mut value = (0usize, 0usize);
                                for node in black_box(&unified).query(black_box(&query))?.take(
                                    if route == "query_first" {
                                        1
                                    } else {
                                        usize::MAX
                                    },
                                ) {
                                    let node = node?;
                                    value.0 += 1;
                                    value.1 += node.view().text()?.map_or(0, |text| text.len());
                                }
                                value
                            }
                        } else {
                            native_once()
                        };
                        let value = black_box(value);
                        total.0 += value.0;
                        total.1 += value.1;
                    }
                    let elapsed = start.elapsed().as_nanos();
                    assert_eq!(total, (expected.0 * repetitions, expected.1 * repetitions));
                    if round >= 2 {
                        if facade {
                            unified_ns.push(elapsed);
                        } else {
                            native_ns.push(elapsed);
                        }
                    }
                }
            }
            let median = |values: &[u128]| {
                let mut values = values.to_vec();
                values.sort_unstable();
                (values[values.len() / 2 - 1] + values[values.len() / 2]) as f64 / 2.0
            };
            cases.push(serde_json::json!({
                "paragraphs": count, "runs": count * 3, "route": route,
                "repetitions_per_sample": repetitions,
                "native_batch_ns": native_ns, "unified_batch_ns": unified_ns,
                "native_median_ns":median(&native_ns) / repetitions as f64, "unified_median_ns":median(&unified_ns) / repetitions as f64,
                "ratio":median(&unified_ns)/median(&native_ns),
                "matched_runs":expected.0, "matched_text_bytes":expected.1,
                "exact_text_and_predicate_oracle":true
            }));
        }
        assert_eq!(native.to_bytes()?, bytes);
        assert_eq!(unified.to_bytes()?, bytes);
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema":"rtf_runs_benchmark.v1", "rounds":9, "warmup_rounds":2, "order":"ABBA",
            "timing":"warmed owners, prepared path and compiled query plan; native/facade lookup, borrowed text length and bold predicates; per-route count/byte checksum; full text and source-byte checks outside timing",
            "cases":cases
        }))?
    );
    Ok(())
}
#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
