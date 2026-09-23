//! Change 0747 review follow-up probe (scratch; never committed as an example).
//!
//! 1. Times the pair API against two `verify_source` calls on constructed
//!    ~64 KiB documents: a one-element edit (the window path), a valid
//!    replacement whose window fails a window-only check late (the last token
//!    of the window is character data ending in `>`), and an invalid
//!    replacement whose defect sits at the end of a window of about half the
//!    document.
//! 2. Prints `verify_source`'s verdict on the pre-existing auditor gaps the
//!    adversarial review listed, and checks that the pair agrees.
#![allow(clippy::unwrap_used, clippy::expect_used, clippy::print_stdout)]

use std::hint::black_box;
use std::time::Instant;

use xml_minifier::audit::{Limits, ReplacementProof, verify_source, verify_source_replacement};

const BATCHES: usize = 41;
const TARGET_BATCH_NS: u128 = 20_000_000;
const ITEMS: usize = 1_900;
const ITEM: &str = "<i a=\"1\">text</i>";

fn time<F: FnMut() -> usize>(mut f: F) -> f64 {
    let mut reps = 1usize;
    loop {
        let start = Instant::now();
        for _ in 0..reps {
            black_box(f());
        }
        if start.elapsed().as_nanos() >= TARGET_BATCH_NS / 4 {
            break;
        }
        reps *= 2;
    }
    let mut per_call = Vec::with_capacity(BATCHES);
    for _ in 0..BATCHES {
        let start = Instant::now();
        for _ in 0..reps {
            black_box(f());
        }
        per_call.push(start.elapsed().as_nanos() as f64 / reps as f64);
    }
    per_call.sort_by(|a, b| a.partial_cmp(b).unwrap());
    per_call[BATCHES / 2]
}

fn items(count: usize, first: &str, last: &str) -> String {
    let mut out = String::new();
    for index in 0..count {
        out.push_str(match index {
            0 => first,
            _ if index == count - 1 => last,
            _ => ITEM,
        });
    }
    out
}

fn document(w_first: &str, w_last: &str, after_w: &str) -> Vec<u8> {
    format!(
        "<root><pre>{}</pre><w>{}</w>{after_w}<z/></root>",
        // A slightly longer sibling keeps a window over all of `<w>` within
        // half the document.
        items(ITEMS + 20, ITEM, ITEM),
        items(ITEMS, w_first, w_last)
    )
    .into_bytes()
}

fn two_audits(original: &[u8], replacement: &[u8]) -> usize {
    usize::from(verify_source(original, Limits::default()).is_ok())
        + usize::from(verify_source(replacement, Limits::default()).is_ok())
}

fn pair(original: &[u8], replacement: &[u8]) -> usize {
    match verify_source_replacement(original, replacement, Limits::default()) {
        Ok(ReplacementProof::Window { .. }) => 2,
        Ok(_) => 1,
        Err(_) => 0,
    }
}

fn main() {
    let original = document(ITEM, ITEM, "");
    let cases = [
        (
            "one-element edit (window proof)",
            document(ITEM, "<i a=\"1\">text, edited</i>", ""),
        ),
        (
            "valid replacement, window fails its last-token check",
            document(ITEM, ITEM, "x>"),
        ),
        (
            "invalid replacement, defect at the end of a half-size window",
            document("<i a=\"2\">text</i>", "<i xml:space=\"bad\"/>", ""),
        ),
    ];
    println!("case,bytes,pair_result,two_verify_source_ns,verify_source_replacement_ns,ratio");
    for (label, replacement) in &cases {
        let expected = (
            verify_source(&original, Limits::default()).is_ok(),
            verify_source(replacement, Limits::default()).is_ok(),
        );
        let result = verify_source_replacement(&original, replacement, Limits::default());
        assert_eq!(result.is_ok(), expected.0 && expected.1, "{label}");
        let before = time(|| two_audits(&original, replacement));
        let after = time(|| pair(&original, replacement));
        println!(
            "{label},{},{:?},{before:.0},{after:.0},{:.4}",
            replacement.len(),
            result.map_err(|error| error.to_string()),
            after / before
        );
    }
    println!();
    println!("gap_input,verify_source,pair_with_edited_neighbour_agrees");
    for input in [
        "<r/>& ;",
        "<r>&;</r>",
        "<r>&a b;</r>",
        "<r>&#xZZ;&#0;</r>",
        "<r><?xml version=\"1.0\"?></r>",
        "<r>]]></r>",
        "<p:r/>",
    ] {
        let verdict = verify_source(input.as_bytes(), Limits::default());
        // The same bytes as a replacement of a document differing in one
        // element: the pair must reach the same verdict.
        let base = format!("<q>{input}<s>0</s></q>");
        let edited = format!("<q>{input}<s>1</s></q>");
        let pair = verify_source_replacement(base.as_bytes(), edited.as_bytes(), Limits::default());
        let complete = (
            verify_source(base.as_bytes(), Limits::default()).is_ok(),
            verify_source(edited.as_bytes(), Limits::default()).is_ok(),
        );
        println!(
            "{input:?},{},{}",
            match &verdict {
                Ok(_) => "accepted".to_string(),
                Err(error) => format!("refused: {error}"),
            },
            pair.is_ok() == (complete.0 && complete.1)
        );
    }
}
