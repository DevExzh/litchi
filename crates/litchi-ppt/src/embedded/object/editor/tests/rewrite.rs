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

#[test]
fn inserts_a_missing_external_object_list_after_the_document_atom() {
    let atom = ppt_record(1, 1001, &[0; 40]);
    let text_info = ppt_record(0, 1010, b"text");
    let end = ppt_record(0, 1002, b"");
    let mut children = atom.clone();
    children.extend_from_slice(&text_info);
    children.extend_from_slice(&end);
    let document = ppt_record(0x000F, 1000, &children);
    let list = ppt_record(0x000F, rewrite::external_object_list(), b"new-list");

    let rewritten = rewrite::replace_or_insert_object_list(&document, &list).unwrap();
    let mut expected = atom;
    expected.extend_from_slice(&list);
    expected.extend_from_slice(&text_info);
    expected.extend_from_slice(&end);
    assert_eq!(rewritten, ppt_record(0x000F, 1000, &expected));

    // An existing list is replaced in place, not duplicated.
    let replacement = ppt_record(0x000F, rewrite::external_object_list(), b"other");
    let replaced = rewrite::replace_or_insert_object_list(&rewritten, &replacement).unwrap();
    assert_eq!(
        replaced.len(),
        rewritten.len() - list.len() + replacement.len()
    );
    assert!(!replaced.windows(list.len()).any(|value| value == list));

    // A record that is not a DocumentContainer starting with its atom is
    // refused.
    assert!(rewrite::replace_or_insert_object_list(&text_info, &list).is_err());
    let no_atom = ppt_record(0x000F, 1000, &end);
    assert!(rewrite::replace_or_insert_object_list(&no_atom, &list).is_err());
}
