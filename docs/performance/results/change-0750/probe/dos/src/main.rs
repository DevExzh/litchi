//! Change 0750 follow-up: times `verify_source` (and one window proof) on
//! inputs whose namespace names are megabytes long and declared by ancestors
//! of many tags or attributes: the reviewer's four inputs and further
//! adversarial families for every namespace check. One process per case.
//!
//! Usage: litchi-dos-probe-0750 <case> [--repeat N] [--budget-seconds S]
//! Prints one JSON line: the case, its size, the verdict, and the seconds of
//! each audit run (up to N runs, stopping once S seconds have been spent).

use std::hint::black_box;
use std::time::Instant;

use xml_minifier::audit::{self, Limits};

/// Three nested declarations of names of `length` bytes that differ only in
/// their last byte (the first two equal when `alias`), around `body`: the
/// reviewer's shape.
fn nested(length: usize, alias: bool, body: &str) -> Vec<u8> {
    let base = "u".repeat(length - 1);
    let second = if alias { '1' } else { '2' };
    format!(
        r#"<r xmlns:a="{base}1"><s xmlns:b="{base}{second}"><t xmlns:c="{base}3">{body}</t></s></r>"#
    )
    .into_bytes()
}

/// Two nested declarations, as in the reviewer's per-tag input.
fn two(length: usize, alias: bool, body: &str) -> Vec<u8> {
    let base = "u".repeat(length - 1);
    let second = if alias { '1' } else { '2' };
    format!(r#"<r xmlns:a="{base}1"><s xmlns:b="{base}{second}">{body}</s></r>"#).into_bytes()
}

fn many_attributes(count: usize) -> String {
    let attributes: Vec<String> = (0..count)
        .map(|index| format!(r#"{}:k{index}="""#, ["a", "b", "c"][index % 3]))
        .collect();
    format!("<x {}/>", attributes.join(" "))
}

fn input(case: &str) -> Vec<u8> {
    const U: usize = 3_900_000;
    match case {
        // The reviewer's inputs.
        "r1-one-tag-249990-attributes" => nested(U, false, &many_attributes(249_990)),
        "r2-one-tag-60000-attributes" => nested(U, false, &many_attributes(60_000)),
        "r3-one-tag-60000-attributes-aliased" => nested(U, true, &many_attributes(60_000)),
        "r4-50000-tags-two-names" => two(U, false, &r#"<x a:k="" b:k=""/>"#.repeat(50_000)),
        // The aliased per-tag path the fix adds: two prefixes bound to one
        // name on every tag, with distinct local names.
        "f1-50000-tags-aliased" => two(U, true, &r#"<x a:k="" b:j=""/>"#.repeat(50_000)),
        // A name redeclared at each of 200 nested levels: each declaration
        // is compared once with the name in scope.
        "f2-200-levels-redeclaring-100k-name" => {
            let name = "v".repeat(100_000);
            let mut document = String::new();
            for level in 0..200 {
                document.push_str(&format!(r#"<e{level} xmlns:a="{name}">"#));
            }
            document.push_str(&r#"<x a:k="" a:j=""/>"#.repeat(10_000));
            for level in (0..200).rev() {
                document.push_str(&format!("</e{level}>"));
            }
            document.into_bytes()
        },
        // 100,000 distinct prefixes in scope, each used once.
        "f3-100000-prefixes" => {
            let mut document = String::from("<r");
            for index in 0..100_000 {
                document.push_str(&format!(r#" xmlns:p{index}="urn:{index}""#));
            }
            document.push('>');
            for index in 0..100_000usize {
                let prefix = index.wrapping_mul(7_919) % 100_000;
                document.push_str(&format!(r#"<x p{prefix}:k=""/>"#));
            }
            document.push_str("</r>");
            document.into_bytes()
        },
        // Six nested elements each binding one more prefix to the same 4 MB
        // name (a start tag may hold 4 MiB), then 20,000 tags that use all six
        // prefixes with distinct local names.
        "f4-6-aliases-of-4mb-name" => {
            let name = "w".repeat(4_000_000);
            let mut document = String::new();
            for index in 0..6 {
                document.push_str(&format!(r#"<e{index} xmlns:q{index}="{name}">"#));
            }
            let tag: Vec<String> = (0..6).map(|index| format!(r#"q{index}:k{index}="""#)).collect();
            document.push_str(&format!("<x {}/>", tag.join(" ")).repeat(20_000));
            for index in (0..6).rev() {
                document.push_str(&format!("</e{index}>"));
            }
            document.into_bytes()
        },
        other => panic!("unknown case {other}"),
    }
}

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    let case = arguments[1].as_str();
    let option = |flag: &str, default: f64| {
        arguments
            .iter()
            .position(|value| value == flag)
            .map_or(default, |index| arguments[index + 1].parse().unwrap())
    };
    let repeat = option("--repeat", 5.0) as usize;
    let budget = option("--budget-seconds", 60.0);
    let limits = Limits::default();
    if case == "w1-window-in-50000-aliased-tags" {
        let original = input("f1-50000-tags-aliased");
        let at = original.len() / 2;
        let edit = at + original[at..].windows(2).position(|pair| pair == b"j=").unwrap();
        let mut replacement = original.clone();
        replacement[edit] = b'i';
        let mut seconds = Vec::new();
        let mut verdict = String::new();
        let spent = Instant::now();
        while seconds.len() < repeat && (seconds.is_empty() || spent.elapsed().as_secs_f64() < budget) {
            let started = Instant::now();
            let result = audit::verify_source_replacement(
                black_box(&original),
                black_box(&replacement),
                limits,
            );
            seconds.push(started.elapsed().as_secs_f64());
            verdict = format!("{result:?}");
        }
        report(case, original.len(), &verdict, &seconds);
        return;
    }
    let document = input(case);
    let mut seconds = Vec::new();
    let mut verdict = String::new();
    let spent = Instant::now();
    while seconds.len() < repeat && (seconds.is_empty() || spent.elapsed().as_secs_f64() < budget) {
        let started = Instant::now();
        let result = audit::verify_source(black_box(&document), limits);
        seconds.push(started.elapsed().as_secs_f64());
        verdict = match result {
            Ok(_) => "Ok".to_owned(),
            Err(error) => format!("{error}"),
        };
    }
    report(case, document.len(), &verdict, &seconds);
}

fn report(case: &str, bytes: usize, verdict: &str, seconds: &[f64]) {
    let verdict = verdict.chars().take(160).collect::<String>().replace('"', "'");
    let list = seconds.iter().map(|value| format!("{value:.6}")).collect::<Vec<_>>().join(",");
    println!(r#"{{"case":"{case}","bytes":{bytes},"verdict":"{verdict}","seconds":[{list}]}}"#);
}
