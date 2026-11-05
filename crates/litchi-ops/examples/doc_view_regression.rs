//! Before/after timing of the DOC paths record 0324 touches, through the
//! public `File` API only, so the same source builds on the base commit and
//! on the branch: paragraph `format` reads (now answering style and list
//! keys), a whole-body query with every paragraph's format, the root and
//! body views, and the paragraph text and run formatting writes (whose
//! dispatch now first checks for a mark path).
//!
//! Run with `cargo run --release -p litchi-ops --example doc_view_regression
//! --features doc [fixture]`; the fixture defaults to
//! `test-data/ole/doc/Lists.doc`. Alternate the two builds' runs.

#[cfg(feature = "doc")]
use std::hint::black_box;
#[cfg(feature = "doc")]
use std::time::Instant;

#[cfg(feature = "doc")]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(feature = "doc")]
const ROUNDS: usize = 15;
#[cfg(feature = "doc")]
const WARMUP: usize = 3;
#[cfg(feature = "doc")]
const ITERATIONS: usize = 64;

#[cfg(feature = "doc")]
#[allow(clippy::cast_precision_loss, reason = "nanosecond medians")]
fn time(mut leg: impl FnMut() -> Outcome<()>) -> Outcome<f64> {
    for _ in 0..WARMUP {
        leg()?;
    }
    let mut samples = Vec::with_capacity(ROUNDS);
    for _ in 0..ROUNDS {
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            leg()?;
        }
        samples.push(start.elapsed().as_nanos() as f64 / ITERATIONS as f64);
    }
    samples.sort_by(f64::total_cmp);
    Ok(samples[samples.len() / 2])
}

#[cfg(feature = "doc")]
fn main() -> Outcome<()> {
    use litchi_ops::props;
    let path = std::env::args().nth(1).map_or_else(
        || {
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/doc/Lists.doc")
        },
        std::path::PathBuf::from,
    );
    let file = litchi_ops::open_bytes(std::fs::read(&path)?)?;
    let format = |path: &'static str| {
        let file = &file;
        move || -> Outcome<()> {
            black_box(file.get(path)?.view().format()?.count());
            Ok(())
        }
    };
    let cases: Vec<(&str, f64)> = vec![
        ("p2_format", time(format("/body/p[2]"))?),
        ("p28_format", time(format("/body/p[28]"))?),
        ("p1_r1_format", time(format("/body/p[1]/r[1]"))?),
        (
            "query_p_formats",
            time(|| {
                for node in file.query("/body > p")? {
                    black_box(node?.view().format()?.count());
                }
                Ok(())
            })?,
        ),
        (
            "root_children",
            time(|| {
                let root = file.get("/")?;
                black_box(root.view().children().count());
                Ok(())
            })?,
        ),
        (
            "body_children",
            time(|| {
                let body = file.get("/body")?;
                black_box(body.view().children().count());
                Ok(())
            })?,
        ),
        (
            "view_text",
            time(|| {
                black_box(file.view(litchi_ops::view::Mode::Text)?.len());
                Ok(())
            })?,
        ),
        (
            "paragraph_text_set",
            time(|| {
                black_box(file.set("/body/p[2]", props! { "text" => "Replaced text" })?);
                Ok(())
            })?,
        ),
        (
            "run_bold_set",
            time(|| {
                black_box(file.set("/body/p[2]/r[1]", props! { "bold" => true })?);
                Ok(())
            })?,
        ),
    ];
    for (name, median) in cases {
        println!("{name}\t{median:.0}");
    }
    Ok(())
}

#[cfg(not(feature = "doc"))]
fn main() {
    eprintln!("enable the `doc` feature");
}
