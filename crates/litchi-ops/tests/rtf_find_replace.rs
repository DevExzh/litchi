//! RTF find/replace through `set` with `find=` and `replace=` (record
//! 0273): one native run-text replacement per changed run, each run keeping
//! its character formatting; a cross-run match collapses into the run where
//! it starts.
#![cfg(feature = "rtf")]

use litchi_ops::{BatchItem, Code, File, Props, open_bytes};

const INPUT: &[u8] =
    br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}\f0 draft one, dra{\b ft two}\par Draft tail}";

fn find(file: &File, path: &str, props: Props<'_>) -> litchi_ops::Result<(File, Option<usize>)> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    let commit = file.apply(&item)?;
    Ok((commit.file().clone(), commit.value().matched()))
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

#[test]
fn matches_collapse_into_their_first_run() {
    let file = open_bytes(INPUT).unwrap();
    assert_eq!(text(&file, "/body/p[1]/r[2]"), "ft two");
    let (after, matched) = find(
        &file,
        "/",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap();
    assert_eq!(matched, Some(2));
    assert_eq!(text(&after, "/body/p[1]"), "final one, final two");
    assert_eq!(text(&after, "/body/p[1]/r[1]"), "final one, final");
    assert_eq!(text(&after, "/body/p[1]/r[2]"), " two");
    assert_eq!(text(&after, "/body/p[2]"), "Draft tail");
    let reopened = open_bytes(after.to_bytes().unwrap()).unwrap();
    assert_eq!(text(&reopened, "/body/p[1]"), "final one, final two");
    // Run scope, case folding and zero matches.
    let (_, matched) = find(
        &file,
        "/body/p[1]/r[2]",
        Props::new().with("find", "draft").with("replace", "x"),
    )
    .unwrap();
    assert_eq!(matched, Some(0));
    let (_, matched) = find(
        &file,
        "/body",
        Props::new()
            .with("find", "DRAFT")
            .with("replace", "x")
            .with("matchCase", false),
    )
    .unwrap();
    assert_eq!(matched, Some(3));
    let before = file.to_bytes().unwrap();
    let (after, matched) = find(
        &file,
        "/",
        Props::new().with("find", "absent").with("replace", "x"),
    )
    .unwrap();
    assert_eq!(matched, Some(0));
    assert_eq!(after.to_bytes().unwrap(), before);
}

#[test]
fn formatting_is_refused() {
    let file = open_bytes(INPUT).unwrap();
    let error = find(
        &file,
        "/",
        Props::new()
            .with("find", "draft")
            .with("replace", "x")
            .with("bold", true),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(error.property(), Some("bold"));
}
