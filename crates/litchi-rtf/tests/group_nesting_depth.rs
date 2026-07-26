//! Group nesting must stay bounded.
//!
//! RTF groups nest arbitrarily and the parser walks them recursively, so an
//! unbounded nesting depth previously let a corrupt file exhaust the call stack
//! and abort the process instead of returning a recoverable error. These tests
//! pin the bound and the surrounding behaviour.

use litchi_rtf::RtfDocument;

/// Mirrors `parser::MAX_GROUP_NESTING_DEPTH`, which is a private implementation
/// detail rather than part of the public surface.
const MAX_GROUP_NESTING_DEPTH: usize = 32;

/// Build `{\rtf1\ansi` + `depth` nested groups wrapping `text`.
fn nested(depth: usize, text: &str) -> String {
    let mut rtf = String::from(r"{\rtf1\ansi");
    rtf.push_str(&"{".repeat(depth));
    rtf.push_str(text);
    rtf.push_str(&"}".repeat(depth));
    rtf.push('}');
    rtf
}

#[test]
fn accepts_nesting_within_the_supported_depth() {
    // The deepest nesting seen anywhere in the real-world compatibility corpus
    // is 15 levels, so documents of that shape must keep parsing.
    for depth in [1, 8, 15, MAX_GROUP_NESTING_DEPTH - 3] {
        let document = RtfDocument::parse(&nested(depth, "content"))
            .unwrap_or_else(|error| panic!("depth {depth} rejected: {error}"));
        assert_eq!(document.text(), "content", "depth {depth} lost its text");
    }
}

#[test]
fn rejects_nesting_beyond_the_supported_depth() {
    for depth in [MAX_GROUP_NESTING_DEPTH + 1, 512, 20_000] {
        let Err(error) = RtfDocument::parse(&nested(depth, "content")) else {
            panic!("depth {depth} was accepted");
        };
        assert!(
            error.to_string().contains("group nesting depth"),
            "depth {depth} reported an unrelated error: {error}"
        );
    }
}

#[test]
fn pathological_nesting_reports_an_error_instead_of_aborting() {
    // A quarter of a million opening braces must still come back as a typed
    // error rather than unwinding the stack.
    let mut rtf = String::from(r"{\rtf1\ansi");
    rtf.push_str(&"{".repeat(250_000));
    assert!(RtfDocument::parse(&rtf).is_err());
}

#[test]
fn corrupt_binary_input_does_not_exhaust_the_stack() {
    // This 155-byte fixture is byte soup, but its interleaved braces and
    // single-letter control words used to drive `parse_group` deep enough to
    // abort the process. Parsing must now terminate with an error.
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/rtf/binary-garbage-nested-groups.rtf");
    let bytes = std::fs::read(&path)
        .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
    assert!(RtfDocument::from_bytes(&bytes).is_err());
}
