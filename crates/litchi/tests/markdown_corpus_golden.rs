//! `File` Markdown over the fixture corpus equals the Markdown the facade's
//! `Document` and `Presentation` types produced.
//!
//! `tests/golden/markdown_corpus.tsv` records, for every DOCX, DOC, RTF,
//! ODT, PPTX, PPT and ODP fixture under `test-data`, under three option sets
//! and on the path and byte routes, the length and FNV-1a digest of the
//! facade types' Markdown, or the facade's refusal. M0a step S8c recorded it
//! from the facade types, and step S9 deleted those types together with
//! `markdown_file_parity.rs`, which pinned the golden to them. This test
//! carries that parity past S9: `File` produces the same bytes where the
//! facade produced Markdown, and refuses where the facade refused. The golden
//! is fixed reference data now: a difference is a behaviour change to review,
//! not output to record again.

#![cfg(all(
    feature = "markdown",
    feature = "docx",
    feature = "doc",
    feature = "rtf",
    feature = "odt",
    feature = "pptx",
    feature = "ppt",
    feature = "odp"
))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::path::{Path, PathBuf};

use litchi::markdown::{MarkdownOptions, TableStyle, ToMarkdown};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// FNV-1a over 64 bits, as the golden records it.
fn fnv1a64(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// The option sets of the golden's columns, in column order.
fn option_sets() -> [(&'static str, MarkdownOptions); 3] {
    [
        ("default", MarkdownOptions::new()),
        ("plain", MarkdownOptions::new().with_styles(false)),
        (
            "metadata+html",
            MarkdownOptions::new()
                .with_metadata(true)
                .with_table_style(TableStyle::StyledHtml),
        ),
    ]
}

/// `File`'s Markdown of the fixture at `path`, on the byte route or the path
/// route.
fn file_markdown(path: &Path, bytes: bool, options: &MarkdownOptions) -> litchi::Result<String> {
    let file = if bytes {
        litchi::open_bytes(std::fs::read(path).unwrap())
    } else {
        litchi::open(path)
    };
    file.and_then(|file| file.to_markdown_with_options(options))
}

#[test]
fn file_markdown_matches_the_recorded_facade_markdown_over_the_corpus() {
    let golden = std::fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/golden/markdown_corpus.tsv"),
    )
    .expect("the Markdown corpus golden");
    let option_sets = option_sets();

    let mut fixtures = 0usize;
    let mut identical = 0usize;
    let mut refused = 0usize;
    let mut differences = Vec::new();
    for line in golden.lines().filter(|line| !line.starts_with('#')) {
        let mut columns = line.split('\t');
        let relative = columns.next().expect("a fixture column");
        let path = fixture(relative);
        assert!(path.is_file(), "the corpus lost {relative}");
        fixtures += 1;

        let mut previous: Option<&str> = None;
        for (name, options) in &option_sets {
            for bytes in [false, true] {
                let route = if bytes { "bytes" } else { "path" };
                let column = columns
                    .next()
                    .unwrap_or_else(|| panic!("{relative}: missing the {name}/{route} column"));
                let expected = if column == "=" {
                    previous.expect("`=` repeats an earlier column")
                } else {
                    column
                };
                previous = Some(expected);

                let actual = file_markdown(&path, bytes, options);
                match (expected.strip_prefix('!'), &actual) {
                    // The facade refused it: `File` must refuse it too.
                    (Some(_), Err(_)) => refused += 1,
                    (None, Ok(markdown))
                        if expected
                            == format!(
                                "{}:{:016x}",
                                markdown.len(),
                                fnv1a64(markdown.as_bytes())
                            ) =>
                    {
                        identical += 1;
                    },
                    _ => differences.push(format!(
                        "{relative} [{name}, {route}]: facade {expected} / File {}",
                        match &actual {
                            Ok(markdown) =>
                                format!("{}:{:016x}", markdown.len(), fnv1a64(markdown.as_bytes())),
                            Err(error) => format!("refused ({})", error.code()),
                        }
                    )),
                }
            }
        }
        assert!(columns.next().is_none(), "{relative}: extra columns");
    }
    eprintln!(
        "Markdown golden over {fixtures} fixtures x {} option sets x 2 routes: {identical} identical, {refused} refused by both, {} differences",
        option_sets.len(),
        differences.len()
    );
    assert!(
        fixtures > 300,
        "the corpus golden shrank: {fixtures} fixtures"
    );
    assert!(
        differences.is_empty(),
        "File Markdown differs from the facade types' recorded Markdown:\n{}",
        differences.join("\n")
    );
}
