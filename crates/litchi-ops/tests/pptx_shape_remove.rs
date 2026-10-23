//! Unified PPTX top-level shape removal through source and eager owners.
#![cfg(feature = "pptx")]

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    for (index, text) in ["A", "B", "C", "D"].into_iter().enumerate() {
        slide.add_text_box(
            text,
            i64::try_from(index).unwrap() * 100,
            0,
            914_400,
            914_400,
        );
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

fn shape_texts(file: &File) -> Vec<String> {
    (1..=4)
        .filter_map(|index| {
            let path = format!("/slide[1]/shape[{index}]");
            file.get(&path)
                .ok()
                .map(|node| node.view().text().unwrap().unwrap().into_owned())
        })
        .collect()
}

fn scoped_shape_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .unwrap()
        .with_scope(Scope::Path(Box::new(SelectorPath::parse("/").unwrap())))
}

#[test]
fn path_removal_reindexes_surviving_shape_ordinals_for_both_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();

        let removed = source.remove("/slide[1]/shape[2]").unwrap();
        assert!(matches!(&removed, File::Pptx(handle) if handle.is_source_backed() != eager));
        assert_eq!(shape_texts(&removed), ["A", "C", "D"]);
        assert_eq!(text(&removed, "/slide[1]/shape[2]"), "C");
        assert!(removed.get("/slide[1]/shape[4]").is_err());

        let removed_again = removed.remove("/slide[1]/shape[2]").unwrap();
        assert_eq!(shape_texts(&removed_again), ["A", "D"]);
        assert_eq!(text(&removed_again, "/slide[1]/shape[2]"), "D");
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn disjoint_shape_selector_removal_is_atomic_for_both_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let selected = scoped_shape_query("slide[1] > shape[2], slide[1] > shape[4]");

        let removed = source.remove(selected).unwrap();
        assert_eq!(shape_texts(&removed), ["A", "C"]);
        assert_eq!(text(&removed, "/slide[1]/shape[1]"), "A");
        assert_eq!(text(&removed, "/slide[1]/shape[2]"), "C");
        assert!(removed.get("/slide[1]/shape[3]").is_err());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn native_shape_removal_charges_the_retained_snapshot_before_publication() {
    let source = open(&fixture(), false);
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.remove("/slide[1]/shape[2]"),
        )
        .expect_err("the native proof must not bypass the snapshot budget");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(shape_texts(&source), ["A", "B", "C", "D"]);
}

#[test]
fn coalesced_shape_write_charges_the_native_patch_before_publication() {
    let source = open(&fixture(), false);
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set("/slide[1]/shape[2]", props! { "text" => "updated" }),
        )
        .expect_err("the coalesced native patch must not bypass the snapshot budget");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(shape_texts(&source), ["A", "B", "C", "D"]);
}

#[test]
fn mixed_shape_and_slide_query_is_refused_without_partial_removal() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let error = source
            .remove(scoped_shape_query("slide[1], slide[1] > shape[2]"))
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(source.to_bytes().unwrap(), before);

        let error = source.remove("/slide[1]/shape[99]").unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(source.to_bytes().unwrap(), before);

        let error = source
            .edit(|edit| {
                edit.remove("/slide[1]/shape[2]")?;
                edit.remove("/slide[1]/shape[99]")
            })
            .unwrap_err();
        assert_eq!(error.item(), Some(1));
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn shape_removal_transaction_replays_inverts_and_keeps_source_immutable() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let commit = source
            .edit(|edit| {
                edit.remove(scoped_shape_query(
                    "slide[1] > shape[2], slide[1] > shape[4]",
                ))
            })
            .unwrap();
        assert_eq!(shape_texts(commit.file()), ["A", "C"]);
        assert_eq!(source.to_bytes().unwrap(), before);

        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            before
        );

        let reopened = litchi_ops::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
        assert_eq!(shape_texts(&reopened), ["A", "C"]);
    }
}

#[test]
fn shape_removal_can_follow_with_a_surviving_ordinal_edit() {
    let source = open(&fixture(), false);
    let removed = source.remove("/slide[1]/shape[2]").unwrap();
    let edited = removed
        .set(
            "/slide[1]/shape[2]",
            litchi_ops::props! { "text" => "surviving C" },
        )
        .unwrap();
    assert_eq!(text(&edited, "/slide[1]/shape[1]"), "A");
    assert_eq!(text(&edited, "/slide[1]/shape[2]"), "surviving C");
    assert_eq!(text(&edited, "/slide[1]/shape[3]"), "D");
}

#[test]
fn disjoint_slide_and_shape_removal_keeps_snapshot_ordinals() {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for value in ["first", "second", "third"] {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box(value, 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let commit = source
            .edit(|edit| edit.remove(scoped_shape_query("slide[1], slide[3] > shape[1]")))
            .unwrap();
        assert_eq!(text(commit.file(), "/slide[1]/shape[1]"), "second");
        assert!(commit.file().get("/slide[2]").is_ok());
        assert!(commit.file().get("/slide[2]/shape[1]").is_err());
        assert!(commit.file().get("/slide[3]").is_err());
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source.to_bytes().unwrap()
        );
    }
}
