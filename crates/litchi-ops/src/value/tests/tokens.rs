//! Closed vocabularies.

use super::super::{ErrorKind, Tokens, What};
use crate::Code;

static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &["single", "double", "thick", "dotted", "none"],
    &[("dbl", 1), ("false", 4), ("sgl", 0), ("true", 0)],
);

#[test]
fn canonical_values_match_in_any_case_and_trimmed() {
    assert_eq!(UNDERLINE.parse("single"), Ok("single"));
    assert_eq!(UNDERLINE.parse("DOUBLE"), Ok("double"));
    assert_eq!(UNDERLINE.parse("  Thick\t"), Ok("thick"));
    assert_eq!(UNDERLINE.get("dotted"), Some("dotted"));
    assert_eq!(UNDERLINE.get("dotted line"), None);
}

#[test]
fn aliases_resolve_to_their_canonical_value() {
    assert_eq!(UNDERLINE.parse("true"), Ok("single"));
    assert_eq!(UNDERLINE.parse("SGL"), Ok("single"));
    assert_eq!(UNDERLINE.parse("dbl"), Ok("double"));
    assert_eq!(UNDERLINE.parse("False"), Ok("none"));
}

#[test]
fn unknown_words_list_the_values_and_suggest() {
    let error = UNDERLINE.parse("doubel").unwrap_err();
    assert_eq!(error.kind(), ErrorKind::UnknownToken);
    assert_eq!(error.what(), What::Token(&UNDERLINE));
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(
        error.to_string(),
        "invalid underline: expected one of single, double, thick, dotted, none"
    );
    assert_eq!(
        error.suggestion().as_deref(),
        Some("did you mean 'double'?")
    );
    // An alias close to the input suggests the canonical value it means.
    let error = UNDERLINE.parse("sgll").unwrap_err();
    assert_eq!(
        error.suggestion().as_deref(),
        Some("did you mean 'single'?")
    );
    assert_eq!(UNDERLINE.parse("zigzagzig").unwrap_err().suggestion(), None);
    assert_eq!(UNDERLINE.parse(" ").unwrap_err().kind(), ErrorKind::Empty);
    let unified = crate::Error::from(UNDERLINE.parse("doubel").unwrap_err());
    assert_eq!(unified.code(), Code::InvalidValue);
    assert_eq!(unified.suggestion(), Some("did you mean 'double'?"));
}

#[test]
fn accessors_expose_the_vocabulary() {
    assert_eq!(UNDERLINE.name(), "underline");
    assert_eq!(UNDERLINE.values().len(), 5);
}

#[test]
#[should_panic(expected = "aliases must be sorted")]
fn unsorted_aliases_are_refused() {
    let _ = Tokens::new("x", &["a", "b"], &[("zz", 0), ("aa", 1)]);
}

#[test]
#[should_panic(expected = "lower-case ASCII")]
fn upper_case_aliases_are_refused() {
    let _ = Tokens::new("x", &["a"], &[("Aa", 0)]);
}

#[test]
#[should_panic(expected = "alias index out of range")]
fn out_of_range_aliases_are_refused() {
    let _ = Tokens::new("x", &["a"], &[("b", 1)]);
}

#[test]
#[should_panic(expected = "at least one value")]
fn empty_vocabularies_are_refused() {
    let _ = Tokens::new("x", &[], &[]);
}
