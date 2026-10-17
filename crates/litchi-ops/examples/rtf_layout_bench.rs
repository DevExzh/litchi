//! Paired RTF layout timing with exact native/facade checks outside timing.
//!
//! Record 0315 adds the route `last_path_typed`: the `last_path` set with
//! every length passed as a typed [`litchi_ops::value::Length`] instead of
//! the text `"6pt"`.
#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::Position;
    use litchi_core::selector::{Path, Scope};
    use litchi_ops::query::Query;
    use litchi_ops::value::Length;
    use litchi_ops::{Props, open_bytes};

    use litchi_rtf::edit::{ParagraphLayoutPatch, ParagraphLayoutUpdate};
    use std::hint::black_box;
    use std::time::Instant;

    let mut cases = Vec::new();
    for count in [1usize, 10, 100, 1_000, 10_000] {
        let texts: Vec<_> = (0..count).map(|i| format!("paragraph {i:05}")).collect();
        let bytes = format!(r"{{\rtf1\ansi {}}}", texts.join(r"\par ")).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = open_bytes(bytes.as_slice())?;
        let path = Path::parse(&format!("/body/p[{count}]"))?.into_owned();
        let query = Query::parse("p")?.with_scope(Scope::Path(Box::new(Path::parse("/body")?)));
        let limits = litchi_rtf::SemanticLimits::builder()
            .max_operations(10_000)?
            .max_replacement_bytes(16 * 1024 * 1024)?
            .build()?;
        let options = litchi_rtf::EditOptions::new().semantic_limits(limits);
        // (every paragraph, typed lengths): `last_path`, `query_all`, then
        // record 0315's `last_path_typed`.
        for (all, typed) in [(false, false), (true, false), (false, true)] {
            let positions: Vec<_> = if all {
                (0..count).collect()
            } else {
                vec![count - 1]
            };
            let delta = ParagraphLayoutPatch::new()
                .with_space_before(120)
                .with_space_after(60)
                .with_left_indent(240)
                .with_right_indent(80)
                .with_first_line_indent(-40)
                .with_keep_together(true)
                .with_keep_with_next(true)
                .with_page_break_before(true);
            let updates: Vec<_> = positions
                .iter()
                .map(|&p| ParagraphLayoutUpdate::new(Position::new(p), delta))
                .collect();
            let native_once = || -> Result<litchi_rtf::Package, litchi_rtf::edit::Error> {
                let mut edit = native.edit_with_options(options);
                edit.patch_body_paragraph_layouts(&updates)?;
                Ok(edit.commit()?.into_package())
            };
            let expected = native_once()?.to_bytes()?;
            let mut native_ns = Vec::new();
            let mut unified_ns = Vec::new();
            for round in 0..11 {
                for route in [false, true, true, false] {
                    let props = if typed {
                        Props::new()
                            .with("spaceBefore", Length::from_points(6))
                            .with("spaceAfter", Length::from_points(3))
                            .with("indent", Length::from_points(12))
                            .with("rightIndent", Length::from_points(4))
                            .with("firstLineIndent", Length::from_points(-2))
                    } else {
                        Props::new()
                            .with("spaceBefore", "6pt")
                            .with("spaceAfter", "3pt")
                            .with("indent", "12pt")
                            .with("rightIndent", "4pt")
                            .with("firstLineIndent", "-2pt")
                    }
                    .with("keepLines", true)
                    .with("keepNext", true)
                    .with("pageBreakBefore", true);
                    let selector = query.clone();
                    let start = Instant::now();
                    let (elapsed, output) = if route {
                        let output = black_box(if all {
                            unified.set(selector, props)?
                        } else {
                            unified.set(&path, props)?
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
                    assert_eq!(reopened.paragraph_count(), count);
                    for (i, paragraph) in reopened.body().paragraphs().enumerate() {
                        assert_eq!(paragraph.to_text(), texts[i]);
                        let layout = paragraph.format().layout();
                        let selected = all || i == count - 1;
                        assert_eq!(layout.space_before(), if selected { 120 } else { 0 });
                        assert_eq!(layout.space_after(), if selected { 60 } else { 0 });
                        assert_eq!(layout.left_indent(), if selected { 240 } else { 0 });
                        assert_eq!(layout.right_indent(), if selected { 80 } else { 0 });
                        assert_eq!(layout.first_line_indent(), if selected { -40 } else { 0 });
                        assert_eq!(layout.keep_together(), selected);
                        assert_eq!(layout.keep_with_next(), selected);
                        assert_eq!(layout.page_break_before(), selected);
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
                "paragraphs":count,"route":match (all, typed) {
                    (true, _) => "query_all",
                    (false, false) => "last_path",
                    (false, true) => "last_path_typed",
                },"targets":updates.len(),
                "native_ns":native_ns,"unified_ns":unified_ns,
                "native_median_ns":median(&native_ns),"unified_median_ns":median(&unified_ns),
                "ratio":median(&unified_ns)/median(&native_ns),"exact_bytes_and_reopened_layout":true
            }));
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "schema":"rtf_layout_benchmark.v1","rounds":9,"warmup_rounds":2,"order":"ABBA",
            "timing":"warmed immutable owners; native updates and facade selectors/props prepared outside timing; staging and commit timed; serialization, readback and output drop excluded",
            "cases":cases
        }))?
    );
    Ok(())
}
#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
