//! Before/after timing of the ODT, ODP and RTF paths record 0328 touches,
//! through the public `File` API only, so the same source builds on the
//! base commit and on the branch: element `format` reads, the root, body
//! and slide views (whose child lists now hold or skip the new kinds), a
//! query over every paragraph, the text view, and a paragraph text write.
//!
//! Run with `cargo run --release -p litchi-ops --example
//! odf_rtf_view_regression --features odt,odp,rtf`. Alternate the two
//! builds' runs.

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use std::hint::black_box;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use std::time::Instant;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const ROUNDS: usize = 15;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const WARMUP: usize = 3;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const ITERATIONS: usize = 64;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
#[allow(clippy::cast_precision_loss, reason = "nanosecond medians")]
fn time(mut leg: impl FnMut() -> Outcome<()>) -> Outcome<f64> {
    // `VR_REPEAT=N` runs each leg N times untimed (for callgrind with
    // `VR_CASE`).
    if let Some(repeat) = std::env::var("VR_REPEAT")
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
    {
        for _ in 0..repeat {
            leg()?;
        }
        return Ok(0.0);
    }
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

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn open(path: &str) -> Outcome<litchi_ops::File> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    Ok(litchi_ops::open_bytes(std::fs::read(path)?)?)
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn format(file: &litchi_ops::File, path: &str) -> Outcome<()> {
    black_box(file.get(path)?.view().format()?.count());
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn children(file: &litchi_ops::File, path: &str) -> Outcome<()> {
    black_box(file.get(path)?.view().children().count());
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn query(file: &litchi_ops::File, selector: &str) -> Outcome<()> {
    for node in file.query(selector)? {
        black_box(node?.view().format()?.count());
    }
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn text_view(file: &litchi_ops::File) -> Outcome<()> {
    black_box(file.view(litchi_ops::view::Mode::Text)?.len());
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn main() -> Outcome<()> {
    use litchi_ops::props;
    let odt = open("odf/corpus/writer-paragraph-styles.odt")?;
    let odp = open("odf/corpus/impress-basic.odp")?;
    let rtf = open(
        &std::env::var("RTF_FIXTURE")
            .unwrap_or_else(|_| "rtf/testDefaultEncodingParse.rtf".to_owned()),
    )?;
    #[allow(clippy::type_complexity, reason = "a list of named legs")]
    let mut cases: Vec<(&str, Box<dyn FnMut() -> Outcome<()> + '_>)> = vec![
        ("odt_p1_format", Box::new(|| format(&odt, "/body/p[1]"))),
        ("odt_root_children", Box::new(|| children(&odt, "/"))),
        ("odt_body_children", Box::new(|| children(&odt, "/body"))),
        ("odt_p1_children", Box::new(|| children(&odt, "/body/p[1]"))),
        ("odt_query_p_formats", Box::new(|| query(&odt, "p"))),
        ("odt_view_text", Box::new(|| text_view(&odt))),
        (
            "odt_p1_text_set",
            Box::new(|| {
                black_box(odt.set("/body/p[1]", props! { "text" => "Replaced text" })?);
                Ok(())
            }),
        ),
        (
            "odp_slide_children",
            Box::new(|| children(&odp, "/slide[1]")),
        ),
        (
            "odp_shape_children",
            Box::new(|| children(&odp, "/slide[1]/shape[1]")),
        ),
        (
            "odp_paragraph_children",
            Box::new(|| children(&odp, "/slide[1]/shape[1]/paragraph[1]")),
        ),
        (
            "odp_shape_format",
            Box::new(|| format(&odp, "/slide[1]/shape[1]")),
        ),
        ("odp_query_shape_formats", Box::new(|| query(&odp, "shape"))),
        ("odp_view_text", Box::new(|| text_view(&odp))),
        ("rtf_p1_format", Box::new(|| format(&rtf, "/body/p[1]"))),
        ("rtf_root_children", Box::new(|| children(&rtf, "/"))),
        ("rtf_body_children", Box::new(|| children(&rtf, "/body"))),
        ("rtf_p1_children", Box::new(|| children(&rtf, "/body/p[1]"))),
        ("rtf_query_p_formats", Box::new(|| query(&rtf, "p"))),
        ("rtf_view_text", Box::new(|| text_view(&rtf))),
        (
            "rtf_p1_text_set",
            Box::new(|| {
                black_box(rtf.set("/body/p[1]", props! { "text" => "Replaced text" })?);
                Ok(())
            }),
        ),
    ];
    if std::env::var("VR_SIZES").is_ok() {
        let root = odt.get("/")?;
        let slide = odp.get("/slide[1]")?;
        let body = rtf.get("/body")?;
        println!(
            "sizes\tnode={} odt={} odp={} rtf={}",
            size_of_val(&root),
            size_of_val(root.as_odt().ok_or("odt")?),
            size_of_val(slide.as_odp().ok_or("odp")?),
            size_of_val(body.as_rtf().ok_or("rtf")?),
        );
        return Ok(());
    }
    let only = std::env::var("VR_CASE").ok();
    for (name, leg) in &mut cases {
        if only.as_deref().is_some_and(|only| only != *name) {
            continue;
        }
        let median = time(leg)?;
        println!("{name}\t{median:.0}");
    }
    Ok(())
}

#[cfg(not(all(feature = "odt", feature = "odp", feature = "rtf")))]
fn main() {
    eprintln!("enable the `odt`, `odp` and `rtf` features");
}
