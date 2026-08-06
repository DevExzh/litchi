use super::super::rewrite;
use super::ppt_record;
use crate::package::Error;

#[test]
fn recursively_replaces_only_external_object_list() {
    let unknown = ppt_record(0, 0x7777, b"unknown");
    let old = ppt_record(0x000F, rewrite::external_object_list(), b"old");
    let mut children = unknown.clone();
    children.extend_from_slice(&old);
    let document = ppt_record(0x000F, 1000, &children);
    let replacement = ppt_record(0x000F, rewrite::external_object_list(), b"new-list");
    let rewritten =
        rewrite::replace_nested_record(&document, rewrite::external_object_list(), &replacement)
            .unwrap();
    assert!(
        rewritten
            .windows(unknown.len())
            .any(|value| value == unknown)
    );
    assert!(
        rewritten
            .windows(replacement.len())
            .any(|value| value == replacement)
    );
    assert!(!rewritten.windows(old.len()).any(|value| value == old));
}

#[test]
fn rejects_excessive_nested_records_without_stack_exhaustion() {
    let mut document = ppt_record(0x000F, rewrite::external_object_list(), b"old");
    for _ in 0..=rewrite::MAX_NESTED_RECORD_DEPTH {
        document = ppt_record(0x000F, 1000, &document);
    }

    let error = rewrite::replace_nested_record(
        &document,
        rewrite::external_object_list(),
        &ppt_record(0x000F, rewrite::external_object_list(), b"new"),
    )
    .unwrap_err();

    assert!(matches!(
        error,
        Error::Corrupted(message)
            if message == "PPT record nesting exceeds the safety limit"
    ));
}
