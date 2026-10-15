//! Before/after timing for the hot paths record 0290's fixes touch.
//!
//! - `pptx_query_shape`: `query("shape")` over `ooxml/pptx/shapes.pptx`, the
//!   walker whose structural tests gained the placeholder test (F1);
//! - `pptx_chart_view`: the node view of a chart with ordinary axis IDs
//!   (`tdf114821.pptx`), the DrawingML axis-ID read that now accepts
//!   negative producer spellings (F2);
//! - `odt_metadata_set_publish`: a LibreOffice ODT `set /metadata author`
//!   plus `to_bytes`, the ODT body-edit publication that now creates a
//!   missing `meta.xml` (F6) with the same single part lookup;
//! - `odp_metadata_set_publish`: the same on a LibreOffice ODP, through the
//!   shared ODF metadata engine, whose commit now looks `meta.xml` up once;
//! - `odp_metadata_read`: `File::metadata()` on a LibreOffice ODP, the shared
//!   `meta.xml` conversion that now carries the language.
//!
//! `CHECKLIST_SCENARIO=<name>` runs one scenario. Inputs are opened outside
//! the clock. Each scenario prints the median and minimum of `ROUNDS` timed
//! rounds (after `WARMUP`) as one JSON line. Build one binary per tree and
//! alternate them ABBA, pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example
//! checklist_fixes_bench --features pptx,odt,odp`.

#[cfg(all(feature = "pptx", feature = "odt", feature = "odp"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::props;

    const ROUNDS: usize = 21;
    const WARMUP: usize = 3;

    let only = std::env::var("CHECKLIST_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);
    let fixture = |relative: &str| {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(relative),
        )
    };

    fn report(name: &str, mut samples: Vec<f64>) {
        samples.sort_by(f64::total_cmp);
        let median = samples[samples.len() / 2];
        println!(
            "{}",
            serde_json::json!({ "scenario": name, "median_us": median, "min_us": samples[0], "rounds": samples.len() })
        );
    }

    /// Microseconds per call of `run`, `repeat` calls per round.
    fn time(repeat: u32, mut run: impl FnMut()) -> Vec<f64> {
        let mut samples = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let started = Instant::now();
            for _ in 0..repeat {
                run();
            }
            if round >= WARMUP {
                samples.push(started.elapsed().as_secs_f64() * 1e6 / f64::from(repeat));
            }
        }
        samples
    }

    if wanted("pptx_query_shape") {
        let file = litchi_ops::open_bytes(fixture("ooxml/pptx/shapes.pptx")?)?;
        report(
            "pptx_query_shape",
            time(50, || {
                black_box(file.query("shape").expect("query").count());
            }),
        );
    }
    if wanted("pptx_chart_view") {
        let file = litchi_ops::open_bytes(fixture(
            "libreoffice-core/sd/qa/unit/data/pptx/tdf114821.pptx",
        )?)?;
        report(
            "pptx_chart_view",
            time(20, || {
                let node = file.get("/slide[1]/chart[1]").expect("chart");
                black_box(node.view().format().expect("chart view"));
            }),
        );
    }
    if wanted("odt_metadata_set_publish") {
        let odt = litchi_ops::open_bytes(fixture("odf/corpus/writer-definition-lists.odt")?)?;
        report(
            "odt_metadata_set_publish",
            time(5, || {
                let bytes = odt
                    .set("/metadata", props! { "author" => "Litchi" })
                    .expect("odt set")
                    .to_bytes()
                    .expect("odt publish");
                black_box(bytes);
            }),
        );
    }
    if wanted("odp_metadata_set_publish") {
        let odp = litchi_ops::open_bytes(fixture("odf/corpus/impress-master-layouts.odp")?)?;
        report(
            "odp_metadata_set_publish",
            time(5, || {
                let bytes = odp
                    .set("/metadata", props! { "author" => "Litchi" })
                    .expect("odp set")
                    .to_bytes()
                    .expect("odp publish");
                black_box(bytes);
            }),
        );
    }
    if wanted("odp_metadata_read") {
        let odp = litchi_ops::open_bytes(fixture("odf/corpus/impress-master-layouts.odp")?)?;
        report(
            "odp_metadata_read",
            time(50, || {
                black_box(odp.metadata().expect("odp metadata"));
            }),
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "pptx", feature = "odt", feature = "odp")))]
fn main() {
    eprintln!("checklist_fixes_bench needs the pptx, odt and odp features");
}
