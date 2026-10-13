//! Differential evidence for records 0262 and 0267: representative queries
//! over the XLSX, DOCX, PPTX, XLS and XLSB corpora return exactly what they
//! returned before the query-walker optimizations (typed comparison, small
//! nodes, direct iteration, walker shortcuts) and the streaming cell cursors.
//!
//! `tests/golden/query_corpus/digests.json` holds the digests the trunk
//! build before record 0262 printed (`examples/query_corpus_digest.rs`, base
//! `949224c319`; re-recorded over the larger corpus after the merge into
//! trunk, where the pre-0262 build `ab7e4b7480` and the merged build
//! `80915b091b` printed byte-identical digests). The XLS and XLSB digests
//! were added by record 0267, printed by trunk `4d1afe92bc` before its
//! changes. The ODS and flat ODS digests were added by record 0277, printed
//! by its read adapter before the ODS writers landed; the ODT digests were
//! added by record 0276 the same way. Each digest covers,
//! file by file, every match's kind,
//! path, positional path and text, and every error, so a changed result set
//! on any file changes it. Set `LITCHI_BLESS_GOLDEN=1` to write the current
//! build's digests instead, after showing that a change of results is
//! intended.

#![cfg(all(
    feature = "xlsx",
    feature = "docx",
    feature = "pptx",
    feature = "xls",
    feature = "xlsb",
    feature = "ods",
    feature = "odt"
))]

use std::collections::BTreeMap;
use std::path::PathBuf;

#[path = "support/query_corpus.rs"]
mod query_corpus;

const GOLDEN: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/golden/query_corpus/digests.json"
);

#[test]
fn corpus_queries_return_what_they_returned_before() {
    let root = PathBuf::from(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data"));
    let actual = query_corpus::digests(&root);
    if std::env::var_os("LITCHI_BLESS_GOLDEN").is_some() {
        let text = serde_json::to_string_pretty(&actual).expect("digests serialize");
        std::fs::write(GOLDEN, format!("{text}\n")).expect("writes the golden");
        return;
    }
    let golden: BTreeMap<String, String> =
        serde_json::from_str(&std::fs::read_to_string(GOLDEN).expect("reads the golden"))
            .expect("parses the golden");
    let mut differences = Vec::new();
    for (key, expected) in &golden {
        match actual.get(key) {
            Some(found) if found == expected => {},
            found => differences.push(format!("{key}: expected {expected}, found {found:?}")),
        }
    }
    for key in actual.keys().filter(|key| !golden.contains_key(*key)) {
        differences.push(format!("{key}: not in the golden"));
    }
    assert!(
        differences.is_empty(),
        "{} of {} digests differ:\n{}",
        differences.len(),
        golden.len(),
        differences.join("\n")
    );
    // The selectors cover the three formats and find matches.
    for (extension, selectors) in [
        ("xlsx", query_corpus::XLSX),
        ("docx", query_corpus::DOCX),
        ("pptx", query_corpus::PPTX),
        ("xls", query_corpus::XLS),
        ("xlsb", query_corpus::XLSB),
        ("ods", query_corpus::ODS),
        ("fods", query_corpus::ODS),
        ("odt", query_corpus::ODT),
    ] {
        let matched = selectors
            .iter()
            .filter(|selector| {
                actual[&format!("{extension} {selector}")]
                    .split(' ')
                    .next()
                    .is_some_and(|count| count != "0")
            })
            .count();
        assert!(
            matched * 2 >= selectors.len(),
            "{extension}: only {matched} of {} selectors match anything",
            selectors.len()
        );
    }
}
