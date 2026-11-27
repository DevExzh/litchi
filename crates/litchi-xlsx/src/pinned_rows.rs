//! Test support for characterization tables: each row pins the outcome a
//! reader gives for one input.

/// Render `(input, outcome)` rows as the Rust table a test pins.
fn render_rows(rows: &[(&str, String)]) -> String {
    rows.iter()
        .map(|(input, outcome)| format!("    (r#\"{input}\"#, {outcome:?}),\n"))
        .collect()
}

/// Require `outcome` to answer every input of `expected` with its pinned
/// outcome; on a difference, print the whole table as it is now.
#[track_caller]
pub(crate) fn assert_rows(expected: &[(&str, &str)], outcome: impl Fn(&str) -> String) {
    let actual = expected
        .iter()
        .map(|(input, _)| (*input, outcome(input)))
        .collect::<Vec<_>>();
    let matches = actual
        .iter()
        .zip(expected)
        .all(|((_, got), (_, want))| got == want);
    assert!(matches, "outcomes changed:\n{}", render_rows(&actual));
}
