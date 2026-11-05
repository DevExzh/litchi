//! PPTX formatting batches coalesce into one native commit (record 0333):
//! run, paragraph and shape-body formatting of direct paths share the
//! shape-text run, report per-item changes, and keep sequential semantics.
#![cfg(feature = "pptx")]

use litchi_ops::{BatchItem, BatchMode, File, ItemResult, OpenOptions, Props, SourceMode};

fn fixture(shapes: usize) -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    for index in 0..shapes {
        slide.add_text_box(
            &format!("text-{index}"),
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

fn item(path: String, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

fn run(index: usize) -> String {
    format!("/slide[1]/shape[{index}]/paragraph[1]/run[1]")
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn changes(results: &[ItemResult]) -> Vec<bool> {
    results
        .iter()
        .map(|result| match result {
            ItemResult::Applied(applied) => applied.changed,
            other => panic!("unexpected result {other:?}"),
        })
        .collect()
}

/// The same items applied one `set` at a time.
fn sequential(file: &File, items: &[(String, Props<'static>)]) -> Vec<u8> {
    let mut current = file.clone();
    for (path, props) in items {
        current = current.set(path.as_str(), props.clone()).unwrap();
    }
    current.to_bytes().unwrap()
}

#[test]
fn a_thousand_run_formats_are_one_commit() {
    let bytes = fixture(1_000);
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let items: Vec<_> = (1..=1_000)
            .map(|index| {
                item(
                    run(index),
                    Props::new()
                        .with("bold", true)
                        .with("size", "20pt")
                        .with("color", "1F4E79"),
                )
            })
            .collect();
        let commit = file.batch(items.clone(), BatchMode::Atomic).unwrap();
        assert_eq!(commit.diagnostics().commits(), 1, "{eager}");
        assert_eq!(commit.diagnostics().operations(), 1_000);
        assert!(changes(commit.value()).iter().all(|changed| *changed));
        assert_eq!(
            prop(commit.file(), &run(1), "bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            prop(commit.file(), &run(1_000), "size").as_deref(),
            Some("20pt")
        );
        // Replaying the batch on its result changes nothing, item by item.
        let again = commit.file().batch(items, BatchMode::Atomic).unwrap();
        assert!(changes(again.value()).iter().all(|changed| !*changed));
        assert!(!again.diagnostics().changed());
    }
}

#[test]
fn mixed_text_and_formatting_keep_sequential_semantics() {
    let bytes = fixture(3);
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let plan: Vec<(String, Props<'static>)> = vec![
            (
                "/slide[1]/shape[1]".into(),
                Props::new().with("text", "one"),
            ),
            (
                run(2),
                Props::new().with("italic", true).with("text", "two"),
            ),
            (
                "/slide[1]/shape[3]/paragraph[1]".into(),
                Props::new().with("align", "center").with("bold", true),
            ),
            // The same run again: closes the run and sees the italic.
            (run(2), Props::new().with("italic", true)),
            // The shape's text after its run's formatting.
            ("/slide[1]/shape[2]".into(), Props::new().with("text", "2")),
            (
                "/slide[1]/shape[1]".into(),
                Props::new().with("underline", "single"),
            ),
        ];
        let commit = file
            .batch(
                plan.iter()
                    .map(|(path, props)| item(path.clone(), props.clone()))
                    .collect::<Vec<_>>(),
                BatchMode::Atomic,
            )
            .unwrap();
        assert_eq!(
            changes(commit.value()),
            [true, true, true, false, true, true],
            "{eager}"
        );
        assert!(commit.diagnostics().commits() < plan.len());
        assert_eq!(commit.file().to_bytes().unwrap(), sequential(&file, &plan));
        assert_eq!(
            prop(commit.file(), &run(1), "underline").as_deref(),
            Some("single")
        );
    }
}

#[test]
fn a_failing_format_keeps_its_wire_index_and_rolls_back() {
    let bytes = fixture(3);
    let file = open(&bytes, false);
    let before = file.to_bytes().unwrap();
    let items = vec![
        item(run(1), Props::new().with("bold", true)),
        item(run(2), Props::new().with("bold", true)),
        // Past the shape's runs: the ordinary writer reports it.
        item(
            "/slide[1]/shape[3]/paragraph[1]/run[9]".into(),
            Props::new().with("bold", true),
        ),
    ];
    let error = file.batch(items, BatchMode::Atomic).unwrap_err();
    assert_eq!(error.item(), Some(2), "{error}");
    assert_eq!(file.to_bytes().unwrap(), before);
}
