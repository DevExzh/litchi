//! Every path and selector string harvested from OfficeCLI's examples, docs,
//! and help schemas parses, and its canonical display parses back to an equal
//! value.

use litchi_core::selector::{Path, Query};

fn corpus(text: &str) -> impl Iterator<Item = (usize, &str)> {
    text.lines()
        .enumerate()
        .map(|(index, line)| (index + 1, line))
        .filter(|(_, line)| !line.trim().is_empty() && !line.starts_with('#'))
}

#[test]
fn officecli_paths_parse_and_round_trip() {
    let text = include_str!("fixtures/officecli_paths.txt");
    let mut count = 0;
    for (line, input) in corpus(text) {
        let path =
            Path::parse(input).unwrap_or_else(|error| panic!("line {line}: {input:?}: {error}"));
        let shown = path.to_string();
        let again = Path::parse(&shown)
            .unwrap_or_else(|error| panic!("line {line}: display {shown:?}: {error}"));
        assert_eq!(path, again, "line {line}: {input:?} -> {shown:?}");
        // Owned parsing agrees with borrowed parsing.
        let owned: Path<'static> = input.parse().expect("owned parse");
        assert_eq!(owned, path);
        // A corpus path whose last segment is an element type or cell range
        // reads as a query too: the query's scope is the path prefix and its
        // last step is the subject. (A trailing free-form sheet name such as
        // `/1-Basics & Quartile` is a path, not a selector compound.)
        let last_is_type = path.last().is_some_and(|step| {
            step.cell_range().is_some()
                || step
                    .name()
                    .bytes()
                    .all(|b| b.is_ascii_alphanumeric() || b == b'_' || b == b'-')
        });
        if last_is_type {
            Query::parse(input)
                .unwrap_or_else(|error| panic!("line {line}: as query {input:?}: {error}"));
        }
        count += 1;
    }
    assert!(count > 900, "harvested path corpus shrank to {count}");
}

#[test]
fn officecli_selectors_parse_and_round_trip() {
    let text = include_str!("fixtures/officecli_selectors.txt");
    let mut count = 0;
    for (line, input) in corpus(text) {
        let query =
            Query::parse(input).unwrap_or_else(|error| panic!("line {line}: {input:?}: {error}"));
        let shown = query.to_string();
        let again = Query::parse(&shown)
            .unwrap_or_else(|error| panic!("line {line}: display {shown:?}: {error}"));
        assert_eq!(query, again, "line {line}: {input:?} -> {shown:?}");
        let owned: Query<'static> = input.parse().expect("owned parse");
        assert_eq!(owned, query);
        count += 1;
    }
    assert!(count > 120, "harvested selector corpus shrank to {count}");
}

#[test]
fn canonical_display_of_harvested_paths_is_stable() {
    // Display is canonical: rendering a rendered path changes nothing.
    let text = include_str!("fixtures/officecli_paths.txt");
    for (_, input) in corpus(text) {
        let once = Path::parse(input).expect("parses").to_string();
        let twice = Path::parse(&once).expect("parses").to_string();
        assert_eq!(once, twice);
    }
}
