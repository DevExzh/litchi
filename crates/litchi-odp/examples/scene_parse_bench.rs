//! `Scene::parse` over record 0308's scene (200 slides, 400 automatic
//! styles), alone in its binary so a build of another tree's scene module
//! (record 0311) differs in nothing else. Prints the median and minimum of
//! `SCENE_ROUNDS` (default 15) samples of 20 parses after 3 warm-ups, in µs
//! per parse. `SCENE_LOOPS=<n>` instead parses `n` times for a profiler.
//!
//! `taskset -c 30 cargo run --release -p litchi-odp --example scene_parse_bench`

use std::hint::black_box;
use std::time::Instant;

fn content() -> String {
    let mut styles = String::new();
    for index in 0..200 {
        styles.push_str(&format!(
            r#"<style:style style:name="P{index}" style:family="paragraph"><style:paragraph-properties fo:margin-left="0cm"/></style:style><style:style style:name="T{index}" style:family="text"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>"#
        ));
    }
    let mut pages = String::new();
    for index in 0..200 {
        pages.push_str(&format!(
            r#"<draw:page draw:name="page{index}"><draw:frame draw:name="Box"><draw:text-box><text:p text:style-name="P{index}">Slide {index} <text:span text:style-name="T{index}">bold</text:span> text</text:p></draw:text-box></draw:frame></draw:page>"#
        ));
    }
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
            r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
            r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
            r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
            r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
            r#"<office:automatic-styles>{}</office:automatic-styles>"#,
            r#"<office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#
        ),
        styles, pages
    )
}

fn main() {
    let content = content();
    let parse = || black_box(litchi_odp::scene::Scene::parse(black_box(&content)).expect("scene"));
    if let Some(loops) = std::env::var("SCENE_LOOPS")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    {
        for _ in 0..loops {
            parse();
        }
        return;
    }
    let rounds = std::env::var("SCENE_ROUNDS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(15usize);
    let mut samples = Vec::new();
    for round in 0..3 + rounds {
        let started = Instant::now();
        for _ in 0..20 {
            parse();
        }
        if round >= 3 {
            samples.push(started.elapsed().as_secs_f64() * 1e6 / 20.0);
        }
    }
    samples.sort_by(f64::total_cmp);
    println!(
        "scene_parse_200 median_us={:.1} min_us={:.1}",
        samples[samples.len() / 2],
        samples[0]
    );
}
