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
        .value()
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
        commit.file().to_bytes().unwrap(),
        subset.file().to_bytes().unwrap()
    );
    assert!(commit.diagnostics().changed());
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
    assert!(!commit.diagnostics().changed());
    assert!(!commit.value().changed);
    assert_eq!(commit.value().warnings.len(), 2);
    assert_eq!(commit.file().to_bytes().unwrap(), before);
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
        &batch::SourceLimits::default(),
    )
    .unwrap();
    let commit = file.batch_with(&items, &warn()).unwrap();
    let warnings: Vec<Vec<(Option<&str>, Option<usize>)>> = commit
        .value()
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
        &batch::SourceLimits::default(),
    )
    .unwrap();
    let error = file.batch_with(&items, &warn()).unwrap_err();
    assert_eq!((error.code(), error.item()), (Code::InvalidValue, Some(1)));
    // Best effort keeps the warned item and reports the failed one.
    let commit = file
        .batch_with(&items, &warn().mode(BatchMode::BestEffort))
        .unwrap();
    assert!(
        matches!(&commit.value()[0], ItemResult::Applied(applied) if applied.warnings.len() == 1)
    );
    assert!(matches!(commit.value()[1], ItemResult::Failed(_)));
}

#[cfg(feature = "xlsx")]
#[test]
fn warn_batches_keep_the_native_runs() {
    // Record 0281: an atomic warn batch coalesces like a refusing one. A
    // homogeneous batch is one native commit under both policies.
    let file = create(FileFormat::Xlsx).unwrap();
    let paths: Vec<String> = (1..=100).map(|row| format!("/Sheet1/A{row}")).collect();
    let items: Vec<BatchItem<'_>> = paths
        .iter()
        .map(|path| set(path, Props::new().with("value", 7)))
        .collect();
    let refuse = file.batch_with(&items, &Options::new()).unwrap();
    let warned = file.batch_with(&items, &warn()).unwrap();
    assert_eq!(refuse.diagnostics().commits(), 1);
    assert_eq!(warned.diagnostics().commits(), 1);
    assert_eq!(warned.diagnostics().operations(), 100);
    assert_eq!(
        warned.file().to_bytes().unwrap(),
        refuse.file().to_bytes().unwrap()
    );

    // A refused item closes the run, applies its supported subset alone and
    // reports its own warnings; the runs on either side stay native. The
    // result equals applying each item in turn.
    let items = batch::from_bytes(
        br#"[
          {"command":"set","path":"/Sheet1/A1","props":{"value":"a"}},
          {"command":"set","path":"/Sheet1/A2","props":{"value":"b"}},
          {"command":"set","path":"/Sheet1/A3","props":{"value":"c","sparkle":1}},
          {"command":"set","path":"/Sheet1/A4","props":{"value":"d"}},
          {"command":"set","path":"/Sheet1/A5","props":{"glow":1}},
          {"command":"set","path":"/Sheet1/A6","props":{"value":"f"}},
          {"command":"set","path":"/Sheet1/A7","props":{"value":"g"}}
        ]"#,
        &batch::SourceLimits::default(),
    )
    .unwrap();
    let commit = file.batch_with(&items, &warn()).unwrap();
    let mut sequential = file.clone();
    for (index, item) in items.iter().enumerate() {
        let step = sequential.apply_with(item, &warn()).unwrap();
        let ItemResult::Applied(applied) = &commit.value()[index] else {
            panic!("item {index} failed");
        };
        assert_eq!(applied.changed, step.value().changed, "item {index}");
        let properties = |warnings: &[litchi_ops::Error]| -> Vec<Option<String>> {
            warnings
                .iter()
                .map(|warning| warning.property().map(str::to_owned))
                .collect()
        };
        assert_eq!(
            properties(&applied.warnings),
            properties(&step.value().warnings),
            "item {index}"
        );
        // Batch warnings carry the wire index.
        assert!(applied.warnings.iter().all(|w| w.item() == Some(index)));
        sequential = step.file().clone();
    }
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        sequential.to_bytes().unwrap()
    );
    // Runs: A1-A2, A3 alone, A4, A6-A7. A5 skips its only property: a
    // no-op that dispatches nothing.
    assert_eq!(commit.diagnostics().operations(), 6);
    assert!(
        commit.diagnostics().commits() <= 4,
        "{:?}",
        commit.diagnostics()
    );

    // A failure other than a refusal still rolls back the whole batch and is
    // attributed to its own item, after a pending run.
    let items = batch::from_bytes(
        br#"[
          {"command":"set","path":"/Sheet1/A1","props":{"value":"a"}},
          {"command":"set","path":"/Sheet1/A2","props":{"value":"b","sparkle":1}},
          {"command":"set","path":"/Sheet1/A3","props":{"value":"c"}},
          {"command":"set","path":"/Sheet1/A4","props":{"font.bold":"maybe"}}
        ]"#,
        &batch::SourceLimits::default(),
    )
    .unwrap();
    let error = file.batch_with(&items, &warn()).unwrap_err();
    assert_eq!((error.code(), error.item()), (Code::InvalidValue, Some(3)));
}

#[cfg(feature = "docx")]
#[test]
fn warn_docx_formatting_runs_match_sequential_application() {
    let file = create(FileFormat::Docx).unwrap();
    let mut seed = Vec::new();
    for text in ["one", "two", "three", "four"] {
        let mut item = BatchItem::new("add");
        item.parent = Some("/body".into());
        item.kind = Some("paragraph".into());
        item.props = Some(Props::new().with("text", text));
        seed.push(item);
    }
    let file = file.batch(&seed, BatchMode::Atomic).unwrap().into_file();
    let items = batch::from_bytes(
        br#"[
          {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
          {"command":"set","path":"/body/p[2]/r[1]","props":{"italic":true,"sparkle":1}},
          {"command":"set","path":"/body/p[3]/r[1]","props":{"bold":true}},
          {"command":"set","path":"/body/p[4]/r[1]","props":{"underline":"single"}}
        ]"#,
        &batch::SourceLimits::default(),
    )
    .unwrap();
    let commit = file.batch_with(&items, &warn()).unwrap();
    let mut sequential = file.clone();
    for item in &items {
        sequential = sequential.apply_with(item, &warn()).unwrap().into_file();
    }
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        sequential.to_bytes().unwrap()
    );
    let ItemResult::Applied(applied) = &commit.value()[1] else {
        panic!("item 1 failed");
    };
    assert_eq!(applied.warnings.len(), 1);
    assert_eq!(applied.warnings[0].property(), Some("sparkle"));
    assert_eq!(applied.warnings[0].item(), Some(1));
    assert!(
        commit.diagnostics().commits() < 4,
        "{:?}",
        commit.diagnostics()
    );
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
            .value()
            .path
            .as_deref()
            .map(ToString::to_string)
            .as_deref(),
        Some("/body/p[1]")
    );
    assert_eq!(commit.value().warnings[0].property(), Some("em"));
    let text = commit
        .file()
        .get("/body/p[1]")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned);
    assert_eq!(text.as_deref(), Some("Hello"));
}

/// A cell `set` carrying a hyperlink takes the ordinary cell-link path in a
/// batch too: the coalesced cell-payload run declines it instead of refusing
/// `link` and `tooltip`, which the registry lists as settable (record 0330,
/// the OfficeCLI corpus line `excel/cell-formatting.sh:158`).
#[cfg(feature = "xlsx")]
#[test]
fn cell_links_in_a_batch_are_applied_not_skipped() {
    let file = create(FileFormat::Xlsx).unwrap();
    let item = set(
        "/Sheet1/A9",
        Props::new()
            .with("value", "OfficeCLI on GitHub")
            .with("link", "https://github.com/iOfficeAI/OfficeCLI")
            .with("tooltip", "Open the repo"),
    );
    for options in [Options::new(), warn()] {
        let commit = file.apply_with(&item, &options).unwrap();
        assert!(
            commit.value().warnings.is_empty(),
            "{:?}",
            commit.value().warnings
        );
        let link = commit.file().get("/Sheet1/hyperlink[1]").unwrap();
        let format: Vec<_> = link
            .view()
            .format()
            .unwrap()
            .map(|prop| (prop.key().to_owned(), prop.value().to_string()))
            .collect();
        assert!(
            format.contains(&("tooltip".to_owned(), "Open the repo".to_owned())),
            "{format:?}"
        );
        // Identical to the direct, non-batch set.
        let direct = file
            .set(
                "/Sheet1/A9",
                litchi_ops::props! {
                    "value" => "OfficeCLI on GitHub",
                    "link" => "https://github.com/iOfficeAI/OfficeCLI",
                    "tooltip" => "Open the repo",
                },
            )
            .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            direct.to_bytes().unwrap()
        );
    }
    // In a batch of several cell sets, the link still lands.
    let other = set("/Sheet1/B1", Props::new().with("value", 2));
    let batch = file
        .batch([other.clone(), item.clone(), other], BatchMode::Atomic)
        .unwrap();
    assert!(batch.file().get("/Sheet1/hyperlink[1]").is_ok());
}
