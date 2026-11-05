//! Record 0280: early-stop and full-query parity for the document, slide
//! and ODF formats (record 0267's `query_early_stop.rs` covers the workbook
//! formats on generated sheets).
//!
//! For every fixture of each format in `test-data` (a bounded sample per
//! format), and for each selector that compiles for it, `limit(n)` returns
//! exactly the full query's first `n` matches and `first()` its first one,
//! each compared by kind, path and text; a second full query on the same,
//! now warm, handle returns the same matches as the first.

#![cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "odt",
    feature = "odp",
    feature = "ods",
    feature = "rtf",
    feature = "doc",
    feature = "ppt"
))]

use std::path::{Path, PathBuf};

use litchi_ops::File;

/// The most fixtures read per format, so the test stays fast in debug.
const PER_FORMAT: usize = 25;

/// Selectors; each format keeps those that compile for it.
const SELECTORS: &[&str] = &[
    "p",
    "r",
    "table",
    "tr",
    "tc",
    "cell",
    "row",
    "shape",
    "slide",
    "picture",
    "p:contains(e)",
    "p[2]",
    "p[last()]",
    "shape[2]",
    "cell[3]",
    "table p",
    "slide > shape",
    "p:empty",
];

const LIMITS: &[usize] = &[1, 2, 3, 7];

fn collect(directory: &Path, extensions: &[&str], out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    let mut entries: Vec<_> = entries.map(|entry| entry.expect("entry").path()).collect();
    entries.sort();
    for path in entries {
        if path.is_dir() {
            collect(&path, extensions, out);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extensions.contains(&extension))
        {
            out.push(path);
        }
    }
}

/// A match as the tests compare it.
fn describe(node: &litchi_ops::node::Node<'_>) -> String {
    let text = node
        .view()
        .text()
        .ok()
        .flatten()
        .map(|text| text.into_owned())
        .unwrap_or_default();
    format!("{} {} {text:?}", node.kind(), node.path())
}

/// The full query's matches, or `None` when the selector does not compile
/// or the query fails (its early-stopping forms are then not compared).
fn full(file: &File, selector: &str) -> Option<Vec<String>> {
    let matches = file.query(selector).ok()?;
    matches
        .map(|node| node.map(|node| describe(&node)))
        .collect::<litchi_ops::Result<Vec<_>>>()
        .ok()
}

fn check_format(extensions: &[&str]) -> usize {
    let mut paths = Vec::new();
    collect(
        Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data")),
        extensions,
        &mut paths,
    );
    let mut checked = 0;
    for path in paths {
        if checked == PER_FORMAT {
            break;
        }
        let Ok(file) = litchi_ops::open(&path) else {
            continue;
        };
        checked += 1;
        for selector in SELECTORS {
            let Some(expected) = full(&file, selector) else {
                continue;
            };
            let label = format!("{} `{selector}`", path.display());
            for &limit in LIMITS {
                let limited: Vec<String> = file
                    .query(*selector)
                    .expect("compiles")
                    .limit(limit)
                    .map(|node| describe(&node.expect("a leading match")))
                    .collect();
                assert_eq!(
                    limited,
                    expected[..expected.len().min(limit)],
                    "{label} limit {limit}"
                );
            }
            let first = file
                .query(*selector)
                .expect("compiles")
                .first()
                .expect("a first match")
                .map(|node| describe(&node));
            assert_eq!(first.as_ref(), expected.first(), "{label} first");
            assert_eq!(
                full(&file, selector).as_ref(),
                Some(&expected),
                "{label} warm"
            );
        }
    }
    checked
}

#[cfg(feature = "docx")]
#[test]
fn docx_early_stop_matches_the_full_query() {
    assert!(check_format(&["docx"]) > 5);
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_early_stop_matches_the_full_query() {
    assert!(check_format(&["pptx"]) > 5);
}

#[cfg(feature = "odt")]
#[test]
fn odt_early_stop_matches_the_full_query() {
    assert!(check_format(&["odt", "fodt"]) > 5);
}

#[cfg(feature = "odp")]
#[test]
fn odp_early_stop_matches_the_full_query() {
    assert!(check_format(&["odp", "fodp"]) > 2);
}

#[cfg(feature = "ods")]
#[test]
fn ods_early_stop_matches_the_full_query() {
    assert!(check_format(&["ods", "fods"]) > 2);
}

#[cfg(feature = "rtf")]
#[test]
fn rtf_early_stop_matches_the_full_query() {
    assert!(check_format(&["rtf"]) > 5);
}

#[cfg(feature = "doc")]
#[test]
fn doc_early_stop_matches_the_full_query() {
    assert!(check_format(&["doc"]) > 5);
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_early_stop_matches_the_full_query() {
    assert!(check_format(&["ppt"]) > 5);
}
