//! Path-text parse cost (record 0326).
//!
//! Times the three ways path text is parsed on the way into a verb:
//! `litchi_core::selector::Path::parse` (the grammar), `PathArg::parse`
//! (`get`, `find`) and `TargetArg::parse` (`set`, `add`, `remove` and the
//! transaction verbs, which record 0314 measured at about 0.96 K
//! instructions per `/body/p[N]`). Each case parses one path form; the
//! printed figure is the median nanoseconds per parse over `--rounds`
//! rounds of `--iterations` parses. No leaf leg: the parse has no leaf
//! counterpart, and the batch harnesses (`pages_bench` `batch_*` against
//! `batch_*_parsed`) measure what it costs a verb.
//!
//! `profile CASE N` parses one case `N` times without timing, for
//! callgrind (`--toggle-collect='*profiled*'`).
//!
//! ```text
//! cargo run --release -p litchi-ops --example path_parse_bench -- [--rounds 9] [--iterations 200000]
//! cargo run --release -p litchi-ops --example path_parse_bench -- profile target_body_p 10000
//! ```

use std::hint::black_box;
use std::time::Instant;

use litchi_core::selector::Path;
use litchi_ops::TargetArg;
use litchi_ops::path::PathArg;

/// The path forms: what the batch harnesses and the OfficeCLI examples use.
const FORMS: &[(&str, &str)] = &[
    ("body_p", "/body/p[1234]"),
    ("table_cell_run", "/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]"),
    ("slide_shape", "/slide[3]/shape[5]"),
    ("sheet_cell", "/Sheet1/B12"),
    ("style_name", "/styles/Heading1"),
    ("para_id", "/body/p[@paraId=1A2B3C4D]"),
    ("shape_id", "/slide[1]/shape[@id=4]/paragraph[2]"),
    ("last", "/body/p[last()]"),
];

#[derive(Clone, Copy)]
enum Parser {
    Grammar,
    PathArg,
    Target,
}

impl Parser {
    const ALL: [(Self, &'static str); 3] = [
        (Self::Grammar, "grammar"),
        (Self::PathArg, "path"),
        (Self::Target, "target"),
    ];

    #[inline(never)]
    fn parse(self, text: &str) -> usize {
        match self {
            Self::Grammar => Path::parse(black_box(text)).map_or(0, |path| path.steps().len()),
            Self::PathArg => {
                PathArg::parse(black_box(text)).map_or(0, |path| path.path().steps().len())
            },
            Self::Target => TargetArg::parse(black_box(text))
                .ok()
                .and_then(|target| target.path().map(|path| path.path().steps().len()))
                .unwrap_or(0),
        }
    }
}

fn argument(name: &str) -> Option<String> {
    let arguments: Vec<String> = std::env::args().collect();
    arguments
        .iter()
        .position(|argument| argument == name)
        .and_then(|index| arguments.get(index + 1).cloned())
}

#[inline(never)]
fn profiled(parser: Parser, text: &str, calls: usize) -> usize {
    let mut steps = 0;
    for _ in 0..calls {
        steps += parser.parse(text);
    }
    steps
}

fn case(name: &str) -> Option<(Parser, &'static str)> {
    let (prefix, form) = name.split_once('_')?;
    let parser = Parser::ALL
        .iter()
        .find(|(_, label)| *label == prefix)
        .map(|(parser, _)| *parser)?;
    let text = FORMS
        .iter()
        .find(|(label, _)| *label == form)
        .map(|(_, text)| *text)?;
    Some((parser, text))
}

fn main() {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    if arguments.first().map(String::as_str) == Some("profile") {
        let name = arguments.get(1).expect("a case such as target_body_p");
        let calls = arguments
            .get(2)
            .and_then(|calls| calls.parse().ok())
            .unwrap_or(10_000);
        let (parser, text) = case(name).expect("a known case");
        black_box(parser.parse(text));
        println!("{}", black_box(profiled(parser, text, calls)));
        return;
    }
    let rounds: usize = argument("--rounds")
        .and_then(|value| value.parse().ok())
        .unwrap_or(9);
    let iterations: usize = argument("--iterations")
        .and_then(|value| value.parse().ok())
        .unwrap_or(200_000);
    for (parser, label) in Parser::ALL {
        for (form, text) in FORMS {
            let steps = parser.parse(text);
            assert!(steps > 0, "{label}_{form}: {text} must parse");
            let mut samples: Vec<f64> = (0..rounds)
                .map(|_| {
                    let start = Instant::now();
                    black_box(profiled(parser, text, iterations));
                    start.elapsed().as_nanos() as f64 / iterations as f64
                })
                .collect();
            samples.sort_by(f64::total_cmp);
            println!(
                "{{\"case\":\"{label}_{form}\",\"path\":\"{text}\",\"steps\":{steps},\"rounds\":{rounds},\"iterations\":{iterations},\"median_ns\":{:.2},\"min_ns\":{:.2}}}",
                samples[rounds / 2],
                samples[0]
            );
        }
    }
}
