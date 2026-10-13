//! The command channel's unsupported-property policy (record 0269):
//! `batch::Unsupported::Warn` applies a command's supported properties,
//! atomically, and reports each refused one; `Refuse`, the default, refuses
//! the whole command. Invalid values and every other failure still fail.

#![cfg(any(feature = "xlsx", feature = "docx"))]
#![cfg_attr(
    not(feature = "xlsx"),
    allow(
        dead_code,
        unused_imports,
        reason = "the batch helpers serve the XLSX tests"
    )
)]

use litchi_core::FileFormat;
use litchi_ops::batch::{self, Options, Unsupported};
use litchi_ops::{BatchItem, BatchMode, Code, ItemResult, Props, create};

fn warn() -> Options {
    Options::new().unsupported(Unsupported::Warn)
}

fn set<'a>(path: &'a str, props: Props<'a>) -> BatchItem<'a> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

#[cfg(feature = "xlsx")]
#[test]
fn warn_applies_the_supported_subset_exactly() {
    let file = create(FileFormat::Xlsx).unwrap();
    let item = set(
        "/Sheet1/A1",
        Props::new()
            .with("value", "Total")
            .with("sparkle", true)
            .with("font.bold", true)
            .with("glow", "red"),
    );
    // The default refuses everything and names the first refused property.
    let error = file.apply(&item).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(error.property(), Some("sparkle"));

    let commit = file.apply_with(&item, &warn()).unwrap();
    let skipped: Vec<_> = commit
        .value
        .warnings
        .iter()
        .map(|warning| (warning.code(), warning.property(), warning.item()))
        .collect();
    assert_eq!(
        skipped,
        [
            (Code::UnsupportedProperty, Some("sparkle"), Some(0)),
            (Code::UnsupportedProperty, Some("glow"), Some(0)),
        ]
    );
    // The result is byte-identical to applying only the supported subset.
    let subset = file
        .apply(set(
            "/Sheet1/A1",
            Props::new().with("value", "Total").with("font.bold", true),
        ))
        .unwrap();
    assert_eq!(
        commit.file.to_bytes().unwrap(),
        subset.file.to_bytes().unwrap()
    );
    assert!(commit.diagnostics.changed());
    // Refuse is the default of the options too.
    assert_eq!(
        file.apply_with(&item, &Options::new()).unwrap_err().code(),
        Code::UnsupportedProperty
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn warn_keeps_other_failures_and_noops_exact() {
    let file = create(FileFormat::Xlsx).unwrap();
    let before = file.to_bytes().unwrap();
    // Every property refused: a no-op that reports them all.
    let commit = file
        .apply_with(
            set(
                "/Sheet1/A1",
                Props::new().with("sparkle", 1).with("glow", 2),
            ),
            &warn(),
        )
        .unwrap();
    assert!(!commit.diagnostics.changed());
    assert!(!commit.value.changed);
    assert_eq!(commit.value.warnings.len(), 2);
    assert_eq!(commit.file.to_bytes().unwrap(), before);
    // An invalid value is never skipped, and nothing applies.
    let error = file
        .apply_with(
            set(
                "/Sheet1/A1",
                Props::new()
                    .with("value", "kept?")
                    .with("font.bold", "maybe")
                    .with("sparkle", 1),
            ),
            &warn(),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(file.to_bytes().unwrap(), before);
    // A missing path is still a path error.
    assert_eq!(
        file.apply_with(set("/Nope/A1", Props::new().with("value", 1)), &warn())
            .unwrap_err()
            .code(),
        Code::PathNotFound
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn warn_batches_report_per_item_and_stay_atomic() {
    let file = create(FileFormat::Xlsx).unwrap();
    let items = batch::from_bytes(
        br#"[
          {"command":"set","path":"/Sheet1/A1","props":{"value":"a","sparkle":1}},
          {"command":"set","path":"/Sheet1/A2","props":{"value":"b"}},
          {"command":"set","path":"/Sheet1/A3","props":{"glow":1,"value":"c"}}
        ]"#,
    )
    .unwrap();
    let commit = file.batch_with(&items, &warn()).unwrap();
    let warnings: Vec<Vec<(Option<&str>, Option<usize>)>> = commit
        .results
        .iter()
        .map(|result| match result {
            ItemResult::Applied(applied) => applied
                .warnings
                .iter()
                .map(|warning| (warning.property(), warning.item()))
                .collect(),
            other => panic!("{other:?}"),
        })
        .collect();
    assert_eq!(
        warnings,
        [
            vec![(Some("sparkle"), Some(0))],
            vec![],
            vec![(Some("glow"), Some(2))]
        ]
    );
    // An item failing for another reason still rolls back the whole batch.
    let items = batch::from_bytes(
        br#"[
          {"command":"set","path":"/Sheet1/A1","props":{"value":"a","sparkle":1}},
          {"command":"set","path":"/Sheet1/A2","props":{"font.bold":"maybe"}}
        ]"#,
    )
    .unwrap();
    let error = file.batch_with(&items, &warn()).unwrap_err();
    assert_eq!((error.code(), error.item()), (Code::InvalidValue, Some(1)));
    // Best effort keeps the warned item and reports the failed one.
    let commit = file
        .batch_with(&items, &warn().mode(BatchMode::BestEffort))
        .unwrap();
    assert!(
        matches!(&commit.results[0], ItemResult::Applied(applied) if applied.warnings.len() == 1)
    );
    assert!(matches!(&commit.results[1], ItemResult::Failed(_)));
}

#[cfg(feature = "docx")]
#[test]
fn warn_adds_the_element_without_the_refused_properties() {
    let file = create(FileFormat::Docx).unwrap();
    let mut item = BatchItem::new("add");
    item.parent = Some("/body".into());
    item.kind = Some("paragraph".into());
    item.props = Some(Props::new().with("text", "Hello").with("em", "dot"));
    assert_eq!(
        file.apply(&item).unwrap_err().code(),
        Code::UnsupportedProperty
    );
    let commit = file.apply_with(&item, &warn()).unwrap();
    assert_eq!(
        commit
            .value
            .path
            .as_deref()
            .map(ToString::to_string)
            .as_deref(),
        Some("/body/p[1]")
    );
    assert_eq!(commit.value.warnings[0].property(), Some("em"));
    let text = commit
        .file
        .get("/body/p[1]")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned);
    assert_eq!(text.as_deref(), Some("Hello"));
}
