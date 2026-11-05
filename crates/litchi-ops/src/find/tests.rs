#![allow(clippy::single_range_in_vec_init, reason = "one mark per rewrite")]

use super::{Matcher, Piece, Spec, plan, requested};
use crate::Code;
use crate::props::Props;

const MAX: usize = 1 << 20;

fn matches(pattern: &str, match_case: bool, whole_word: bool, text: &str) -> Vec<(usize, usize)> {
    Matcher::new(pattern, match_case, whole_word)
        .find_iter(text)
        .map(|range| (range.start, range.end))
        .collect()
}

/// Apply a plan to piece texts, returning the new texts.
fn replace(texts: &[(&str, usize)], pattern: &str, replacement: &str) -> (usize, Vec<String>) {
    let pieces: Vec<Piece<'_>> = texts
        .iter()
        .map(|&(text, group)| Piece::new(text, group))
        .collect();
    let matcher = Matcher::new(pattern, true, false);
    let planned = plan(&pieces, &matcher, Some(replacement), false, MAX).unwrap();
    let mut output: Vec<String> = texts.iter().map(|(text, _)| (*text).to_owned()).collect();
    for rewrite in &planned.rewrites {
        output[rewrite.piece] = rewrite
            .text
            .clone()
            .unwrap_or_else(|| texts[rewrite.piece].0.to_owned());
        assert!(rewrite.marks.is_empty());
    }
    (planned.matches, output)
}

#[test]
fn literal_matches_are_leftmost_and_non_overlapping() {
    assert_eq!(matches("aa", true, false, "aaaa"), [(0, 2), (2, 4)]);
    assert_eq!(matches("aa", true, false, "aaa"), [(0, 2)]);
    assert_eq!(matches("x", true, false, "abc"), []);
    assert_eq!(matches("ab", true, false, "AB ab"), [(3, 5)]);
}

#[test]
fn case_folding_is_simple_and_unicode_aware() {
    assert_eq!(
        matches("straße", false, false, "STRAẞE straße"),
        [(0, 8), (9, 16)]
    );
    assert_eq!(matches("ΣΟΦΟΣ", false, false, "σοφος"), [(0, 10)]);
    assert_eq!(matches("Hello", false, false, "hELLO"), [(0, 5)]);
    // A one-to-many lowercase (U+0130) compares as itself.
    assert_eq!(matches("İ", false, false, "i İ"), [(2, 4)]);
    // Folding never expands: `ß` does not match `ss`.
    assert!(matches("ß", false, false, "ss").is_empty());
}

#[test]
fn whole_words_retry_inside_rejected_candidates() {
    assert_eq!(
        matches("cat", true, true, "cat concat cat_ cat."),
        [(0, 3), (16, 19)]
    );
    assert_eq!(matches("aa", true, true, "baaa aa"), [(5, 7)]);
    assert_eq!(matches("é", true, true, "éé é"), [(5, 7)]);
}

#[test]
fn a_cross_run_match_collapses_into_the_first_run() {
    let (count, texts) = replace(&[("Hel", 0), ("lo wor", 0), ("ld", 0)], "llo w", "X");
    assert_eq!(count, 1);
    assert_eq!(texts, ["HeX", "or", "ld"]);
    let (count, texts) = replace(&[("ab", 0), ("cd", 0), ("ef", 0)], "bcde", "Z");
    assert_eq!(count, 1);
    assert_eq!(texts, ["aZ", "", "f"]);
}

#[test]
fn a_match_at_a_run_boundary_takes_the_following_run() {
    let (_, texts) = replace(&[("ab", 0), ("cd", 0)], "cd", "Y");
    assert_eq!(texts, ["ab", "Y"]);
    let (_, texts) = replace(&[("ab", 0), ("cd", 0)], "ab", "Y");
    assert_eq!(texts, ["Y", "cd"]);
}

#[test]
fn groups_are_barriers() {
    let (count, texts) = replace(&[("ab", 0), ("cd", 1)], "bc", "Y");
    assert_eq!(count, 0);
    assert_eq!(texts, ["ab", "cd"]);
}

#[test]
fn several_matches_in_one_run_and_across_runs() {
    let (count, texts) = replace(&[("a-a-", 0), ("a", 0)], "a", "bb");
    assert_eq!(count, 3);
    assert_eq!(texts, ["bb-bb-", "bb"]);
    let (count, texts) = replace(&[("xa", 0), ("bx", 0), ("ab", 0)], "ab", "");
    assert_eq!(count, 2);
    assert_eq!(texts, ["x", "x", ""]);
}

#[test]
fn empty_pieces_are_never_rewritten() {
    let (count, texts) = replace(&[("a", 0), ("", 0), ("b", 0)], "ab", "Q");
    assert_eq!(count, 1);
    assert_eq!(texts, ["Q", "", ""]);
}

#[test]
fn untouched_pieces_have_no_rewrite() {
    let pieces = [
        Piece::new("keep", 0),
        Piece::new("old", 1),
        Piece::new("same", 2),
    ];
    let planned = plan(
        &pieces,
        &Matcher::new("old", true, false),
        Some("new"),
        false,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.rewrites.len(), 1);
    assert_eq!(planned.rewrites[0].piece, 1);
    // Replacing a match with itself counts it but rewrites nothing.
    let planned = plan(
        &pieces,
        &Matcher::new("old", true, false),
        Some("old"),
        false,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.matches, 1);
    assert!(planned.rewrites.is_empty());
}

#[test]
fn replacement_marks_cover_only_the_replacement() {
    let pieces = [Piece::new("say hi", 0), Piece::new(" there", 0)];
    let planned = plan(
        &pieces,
        &Matcher::new("hi th", true, false),
        Some("HEY"),
        true,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.rewrites.len(), 2);
    assert_eq!(planned.rewrites[0].text.as_deref(), Some("say HEY"));
    assert_eq!(planned.rewrites[0].marks, [4..7]);
    assert_eq!(planned.rewrites[1].text.as_deref(), Some("ere"));
    assert!(planned.rewrites[1].marks.is_empty());
}

#[test]
fn format_only_marks_every_matched_slice() {
    let pieces = [Piece::new("say hi", 0), Piece::new(" there", 0)];
    let planned = plan(
        &pieces,
        &Matcher::new("hi th", true, false),
        None,
        true,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.matches, 1);
    assert_eq!(planned.rewrites.len(), 2);
    assert!(
        planned
            .rewrites
            .iter()
            .all(|rewrite| rewrite.text.is_none())
    );
    assert_eq!(planned.rewrites[0].marks, [4..6]);
    assert_eq!(planned.rewrites[1].marks, [0..3]);
}

#[test]
fn unscoped_pieces_filter_matches_but_give_context() {
    let mut pieces = [
        Piece::new("ab", 0),
        Piece::new("ab", 0),
        Piece::new("ab", 0),
    ];
    pieces[0].scoped = false;
    pieces[2].scoped = false;
    let planned = plan(
        &pieces,
        &Matcher::new("ba", true, false),
        Some("X"),
        false,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.matches, 0);
    let planned = plan(
        &pieces,
        &Matcher::new("ab", true, false),
        Some("X"),
        false,
        MAX,
    )
    .unwrap();
    assert_eq!(planned.matches, 1);
    assert_eq!(planned.rewrites[0].piece, 1);
}

#[test]
fn staged_text_is_bounded() {
    let pieces = [Piece::new("aaaa", 0)];
    let error = plan(
        &pieces,
        &Matcher::new("a", true, false),
        Some("bbbb"),
        false,
        8,
    )
    .err()
    .unwrap();
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn unicode_offsets_stay_on_character_boundaries() {
    let (count, texts) = replace(&[("日本", 0), ("語です", 0)], "本語", "🙂");
    assert_eq!(count, 1);
    assert_eq!(texts, ["日🙂", "です"]);
}

fn spec(props: Props<'_>) -> crate::Result<Spec<'_>> {
    Spec::admit(&props)
}

#[test]
fn admission_follows_officecli() {
    assert!(requested(&Props::new().with("FIND", "x")));
    assert!(!requested(&Props::new().with("replace", "x")));
    let props = Props::new().with("find", "a").with("replace", "b");
    let admitted = spec(props.clone()).unwrap();
    assert_eq!(admitted.replacement(), Some("b"));
    assert!(!admitted.formats());
    let props = Props::new().with("find", "a").with("bold", true);
    assert!(spec(props.clone()).unwrap().formats());
    // A bare find does nothing: OfficeCLI refuses it.
    let error = spec(Props::new().with("find", "a")).unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    assert_eq!(error.property(), Some("replace"));
    // An empty replacement deletes.
    let props = Props::new().with("find", "a").with("replace", "");
    assert_eq!(spec(props.clone()).unwrap().replacement(), Some(""));
}

#[test]
fn admission_refusals() {
    let codes = [
        (
            Props::new().with("find", "").with("replace", "b"),
            Code::InvalidValue,
        ),
        (
            Props::new().with("find", "a\tb").with("replace", "b"),
            Code::InvalidValue,
        ),
        (
            Props::new().with("find", "a").with("replace", "b\nc"),
            Code::InvalidValue,
        ),
        (
            Props::new()
                .with("find", "a")
                .with("find", "b")
                .with("replace", "c"),
            Code::InvalidArgument,
        ),
        (
            Props::new()
                .with("find", "a")
                .with("replace", "b")
                .with("regex", true),
            Code::UnsupportedCapability,
        ),
        (
            Props::new().with("find", r#"r"\d+""#).with("replace", "b"),
            Code::UnsupportedCapability,
        ),
        (
            Props::new()
                .with("find", "a")
                .with("replace", "b")
                .with("matchCase", "perhaps"),
            Code::InvalidValue,
        ),
        (
            Props::new().with("find", true).with("replace", "b"),
            Code::InvalidValue,
        ),
        (
            Props::new()
                .with("find", "a".repeat(super::MAX_TEXT_BYTES + 1))
                .with("replace", "b"),
            Code::LimitExceeded,
        ),
    ];
    for (props, code) in codes {
        assert_eq!(spec(props).unwrap_err().code(), code);
    }
    // `regex=false` is accepted, and integers are literal text.
    let props = Props::new()
        .with("find", 2024)
        .with("replace", "x")
        .with("regex", "false");
    assert!(spec(props).is_ok());
}
