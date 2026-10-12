//! Paired RTF paragraph move timing; setup and exact output checks are untimed.
#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::{Position, selector::Path};
    use litchi_ops::{Placement, open_bytes};
    use std::hint::black_box;
    use std::time::Instant;

    let mut cases = Vec::new();
    for count in [1usize, 10, 100, 1_000, 10_000] {
        let texts: Vec<_> = (0..count).map(|i| format!("paragraph {i:05}")).collect();
        let bytes = format!(r"{{\rtf1\ansi {}}}", texts.join(r"\par ")).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = open_bytes(bytes.as_slice())?;
        let path = Path::parse("/body/p[1]")?;
        let limits = litchi_rtf::SemanticLimits::builder()
            .max_operations(10_000)?
            .max_replacement_bytes(16 * 1024 * 1024)?
            .build()?;
        let options = litchi_rtf::EditOptions::new().semantic_limits(limits);
        let native_once = || -> Result<litchi_rtf::Package, litchi_rtf::edit::Error> {
            let mut edit = native.edit_with_options(options);
            edit.move_paragraph(Position::new(0), Position::new(count - 1))?;
            Ok(edit.commit()?.into_package())
        };
        let expected = native_once()?.to_bytes()?;
        let mut expected_text = texts.clone();
        expected_text.rotate_left(1);
        let mut native_ns = Vec::new();
        let mut unified_ns = Vec::new();
        // Two warm-up rounds followed by nine ABBA rounds: 18 samples per route.
        for round in 0..11 {
            for route in [false, true, true, false] {
                let start = Instant::now();
                let (elapsed, output) = if route {
                    let output = black_box(unified.move_to(&path, Placement::End)?);
                    let elapsed = start.elapsed().as_nanos();
                    (elapsed, output.to_bytes()?)
                } else {
                    let output = black_box(native_once()?);
                    let elapsed = start.elapsed().as_nanos();
                    (elapsed, output.to_bytes()?)
                };
                assert_eq!(output, expected);
                let reopened = litchi_rtf::Package::from_bytes(&output)?;
                let actual: Vec<_> = reopened.body().paragraphs().map(|p| p.to_text()).collect();
                assert_eq!(actual, expected_text);
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
            "paragraphs": count, "native_ns": native_ns, "unified_ns": unified_ns,
            "native_median_ns": median(&native_ns), "unified_median_ns": median(&unified_ns),
            "ratio": median(&unified_ns) / median(&native_ns),
            "exact_bytes_and_reopened_order": true
        }));
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema": "rtf_move_benchmark.v1", "placement": "first_to_end",
            "timing": "warmed immutable owner; staging and commit; setup, serialization, output validation and drop excluded",
            "rounds": 9, "warmup_rounds": 2, "order": "ABBA", "cases": cases
        }))?
    );
    Ok(())
}

#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
