//! The leaf leg of record 0304's `odt_plain` scenario alone in its binary:
//! `Edit::set_text` + `commit` of the middle paragraph of a 1,000-paragraph
//! builder document, so a build of another tree's `litchi-odt` (record 0311)
//! differs in nothing else. Prints the median and minimum of `ODT_ROUNDS`
//! (default 15) samples of 20 calls after 3 warm-ups, in µs per call.
//! `ODT_LOOPS=<n>` instead runs `n` calls for a profiler.
//!
//! `taskset -c 30 cargo run --release -p litchi-odt --example plain_set_bench`

use std::hint::black_box;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..1_000 {
        builder.add_paragraph(&format!(
            "Paragraph {index} holds a few ordinary words of body text."
        ))?;
    }
    let package = litchi_odt::Package::from_bytes(&builder.build()?.to_bytes()?)?;
    let id = {
        let document = package.document()?;
        let body = document.body()?;
        body.root()
            .children()
            .filter(|node| node.kind() == litchi_odt::body::Kind::Paragraph)
            .nth(500)
            .expect("paragraph")
            .id()
    };
    let set = || {
        let mut edit = package.edit();
        edit.set_text(id, "Replaced text").expect("stage");
        black_box(edit.commit().expect("commit").into_package());
    };
    if let Some(loops) = std::env::var("ODT_LOOPS")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    {
        for _ in 0..loops {
            set();
        }
        return Ok(());
    }
    let rounds = std::env::var("ODT_ROUNDS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(15usize);
    let mut samples = Vec::new();
    for round in 0..3 + rounds {
        let started = Instant::now();
        for _ in 0..20 {
            set();
        }
        if round >= 3 {
            samples.push(started.elapsed().as_secs_f64() * 1e6 / 20.0);
        }
    }
    samples.sort_by(f64::total_cmp);
    println!(
        "odt_plain_set median_us={:.1} min_us={:.1}",
        samples[samples.len() / 2],
        samples[0]
    );
    Ok(())
}
