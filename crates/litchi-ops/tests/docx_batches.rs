//! Command-channel semantics over the same immutable transaction coordinator.
#![cfg(feature = "docx")]
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, NewNode, OpenOptions, Placement, SourceMode,
    props,
};
use std::io::{Cursor, Write as _};

const XML: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>FIRST_PRIVATE</w:t></w:r><w:r><w:t>tail</w:t></w:r></w:p><!--between--><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;

fn bytes() -> Vec<u8> {
    bytes_with_xml(XML)
}

fn bytes_with_xml(xml: &str) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", xml),
        ("opaque.xml", "<opaque>UNTOUCHED_PRIVATE</opaque>"),
    ] {
        writer.start_file(name, options).unwrap();
        writer.write_all(text.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn open(eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}
fn text(file: &File) -> String {
    file.get("/body/p[1]/r[1]")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}
fn items(json: &str) -> Vec<BatchItem<'_>> {
    litchi_ops::batch::from_bytes(json.as_bytes(), &litchi_ops::batch::SourceLimits::default())
        .unwrap()
}
fn applied(result: &ItemResult) -> &litchi_ops::Applied {
    match result {
        ItemResult::Applied(value) => value,
        ItemResult::Failed(error) => panic!("unexpected item failure {error:?}"),
        _ => panic!("unknown result"),
    }
}
fn failed(result: &ItemResult, index: usize, code: Code) {
    let ItemResult::Failed(error) = result else {
        panic!("expected item failure")
    };
    assert_eq!(error.item(), Some(index));
    assert_eq!(error.code(), code);
}

#[test]
fn ordered_five_command_batches_match_single_operation_bytes_and_invert() {
    let input = items(
        r#"[
      {"command":"meta","dumpVersion":2},
      {"command":"ADD","parent":"/body","type":"paragraph","props":{"text":"third"}},
      {"command":"set","path":"/body/p[3]/r[1]","props":{"text":"last"}},
      {"command":"move","path":"/body/p[3]","index":0},
      {"command":"swap","path":"/body/p[1]","path2":"/body/p[2]"},
      {"command":"remove","path":"/body/p[3]"}
    ]"#,
    );
    for eager in [false, true] {
        let file = open(eager);
        let oracle = file
            .add("/body", NewNode::new("paragraph").prop("text", "third"))
            .unwrap()
            .file
            .set("/body/p[3]/r[1]", props! {"text"=>"last"})
            .unwrap()
            .move_to("/body/p[3]", Placement::Start)
            .unwrap()
            .swap("/body/p[1]", "/body/p[2]")
            .unwrap()
            .remove("/body/p[3]")
            .unwrap();
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let commit = file.batch(&input, mode).unwrap();
            assert_eq!(commit.value().len(), 6);
            assert!(!applied(&commit.value()[0]).changed);
            assert_eq!(
                applied(&commit.value()[1])
                    .path
                    .as_ref()
                    .unwrap()
                    .to_string(),
                "/body/p[3]"
            );
            assert!(
                commit.value()[1..]
                    .iter()
                    .all(|result| applied(result).changed)
            );
            assert_eq!(commit.diagnostics().operations(), 5);
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                oracle.to_bytes().unwrap()
            );
            assert_eq!(
                commit.patch().apply(&file).unwrap().to_bytes().unwrap(),
                oracle.to_bytes().unwrap()
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
            assert_eq!(file.to_bytes().unwrap(), bytes());
        }
    }
}

#[test]
fn atomic_stops_at_first_failure_with_original_wire_index() {
    let file = open(false);
    let input = items(
        r#"[
      {"command":"meta","dumpVersion":2},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"staged"}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"partial","wrong-property":true}},
      {"command":"remove","path":"/body/p[1]"}
    ]"#,
    );
    let pulls = std::cell::Cell::new(0);
    let error = file
        .batch(
            input.into_iter().inspect(|_| pulls.set(pulls.get() + 1)),
            BatchMode::Atomic,
        )
        .unwrap_err();
    assert_eq!(error.item(), Some(2));
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(pulls.get(), 3);
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn best_effort_discards_failed_candidates_and_keeps_original_item_indices() {
    for eager in [false, true] {
        let file = open(eager);
        let commit = file.batch(items(r#"[
          {"command":"unknown_PRIVATE"},
          {"command":"meta","dumpVersion":2},
          {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"staged"}},
          {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"partial","wrong-property":true}},
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"third"}},
          {"command":"remove","path":"/body/p[99]"},
          {"command":"set","path":"/body/p[3]/r[1]","props":{"bold":true}}
        ]"#), BatchMode::BestEffort).unwrap();
        assert_eq!(commit.value().len(), 7);
        failed(&commit.value()[0], 0, Code::UnsupportedCommand);
        failed(&commit.value()[3], 3, Code::UnsupportedProperty);
        failed(&commit.value()[5], 5, Code::PathNotFound);
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(text(commit.file()), "staged");
        let oracle = file
            .set("/body/p[1]/r[1]", props! {"text"=>"staged"})
            .unwrap()
            .add("/body", NewNode::new("paragraph").prop("text", "third"))
            .unwrap()
            .file
            .set("/body/p[3]/r[1]", props! {"bold"=>true})
            .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
        assert!(!format!("{commit:?} {:?}", commit.value()).contains("PRIVATE"));
    }
}

#[test]
fn single_apply_preserves_path_value_and_has_the_same_strict_admission() {
    let file = open(false);
    let mut add = BatchItem::new("add");
    add.path = Some("/body".into());
    add.kind = Some("paragraph".into());
    let commit = file.apply(&add).unwrap();
    assert_eq!(
        commit.value().path.as_ref().unwrap().to_string(),
        "/body/p[3]"
    );
    assert!(commit.value().changed);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        bytes()
    );
    let mut set = BatchItem::new("set");
    set.path = Some("/body/p[1]".into());
    let error = file.apply(set).unwrap_err();
    assert_eq!(error.code(), Code::MissingArgument);
    assert_eq!(error.item(), Some(0));
}

#[test]
fn strict_fields_aliases_and_placement_conflicts_never_drop_data() {
    let file = open(false);
    for json in [
        r#"[{"command":"set","path":"/body/p[1]","props":{"bold":true},"text":"unused"}]"#,
        r#"[{"command":"remove","path":"/body/p[1]","props":{}}]"#,
        r#"[{"command":"add","parent":"/body","path":"/body","type":"paragraph"}]"#,
        r#"[{"command":"add","parent":"/body","type":"paragraph","from":"/body/p[1]"}]"#,
        r#"[{"command":"add","parent":"/body","type":"paragraph","index":0,"after":"/body/p[1]"}]"#,
        r#"[{"command":"move","path":"/body/p[1]","before":"/body/p[2]","after":"/body/p[2]"}]"#,
        r#"[{"command":"swap","path":"/body/p[1]","path2":"/body/p[2]","to":"/body/p[2]"}]"#,
        r#"[{"command":"set","path":"/body/p[1]","props":{}}]"#,
        r#"[{"command":"move","path":"/body/p[1]","index":-1}]"#,
    ] {
        let error = file.batch(items(json), BatchMode::Atomic).unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument, "{json}");
        assert_eq!(error.item(), Some(0));
    }
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn placement_and_swap_aliases_use_current_positional_state() {
    let file = open(false);
    let input = items(
        r#"[
      {"op":"add","path":"/body","type":"paragraph","after":"/body/p[1]","props":{"text":"middle"}},
      {"command":"move","path":"/body/p[3]/r[1]","to":"/body/p[2]","index":0},
      {"command":"swap","path":"/body/p[1]","to":"/body/p[2]"}
    ]"#,
    );
    let commit = file.batch(input, BatchMode::Atomic).unwrap();
    let oracle = file
        .add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "middle")
                .at(Placement::after("/body/p[1]").unwrap()),
        )
        .unwrap()
        .file
        .move_to(
            "/body/p[3]/r[1]",
            Placement::inside_at("/body/p[2]", litchi_core::Position::new(0)).unwrap(),
        )
        .unwrap()
        .swap("/body/p[1]", "/body/p[2]")
        .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        oracle.to_bytes().unwrap()
    );
}

#[test]
fn every_known_pending_command_and_transfer_option_is_explicit() {
    let file = open(false);
    for command in [
        "get", "query", "view", "raw", "raw-set", "add-part", "validate", "import",
    ] {
        assert_eq!(
            file.apply(BatchItem::new(command)).unwrap_err().code(),
            Code::UnsupportedCapability
        );
    }
    for json in [
        r#"[{"command":"add","parent":"/body","from":"/body/p[1]"}]"#,
        r#"[{"command":"move","path":"/body/p[1]","props":{"text":"ignored"}}]"#,
        r#"[{"command":"move","path":"/body/p[1]","props":{}}]"#,
        r#"[{"command":"move","path":"/body/p[1]","to":"/body","after":"/body/p[2]"}]"#,
    ] {
        assert_eq!(
            file.batch(items(json), BatchMode::Atomic)
                .unwrap_err()
                .code(),
            Code::UnsupportedCapability
        );
    }
    assert_eq!(
        file.apply(BatchItem::new("move")).unwrap_err().code(),
        Code::MissingArgument
    );
}

#[test]
fn metadata_outcomes_preserve_indices_and_unsupported_versions_abort_both_modes() {
    let file = open(false);
    let input = items(r#"[{"command":"meta","dumpVersion":2},{"command":"meta","dumpVersion":2}]"#);
    let commit = file.batch(input, BatchMode::Atomic).unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.value().len(), 2);
    assert_eq!(commit.diagnostics().operations(), 0);
    for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
        for version in [None, Some(-1), Some(0), Some(1), Some(3)] {
            let mut input =
                items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"staged"}}]"#);
            let mut meta = BatchItem::new("meta");
            meta.dump_version = version;
            input.push(meta);
            let error = file.batch(input, mode).unwrap_err();
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert_eq!(error.item(), Some(1));
        }
    }
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn all_failed_empty_and_cancelling_batches_preserve_exact_input() {
    let file = open(false);
    let empty = file
        .batch(Vec::<BatchItem<'_>>::new(), BatchMode::Atomic)
        .unwrap();
    assert!(empty.value().is_empty());
    assert!(empty.patch().is_noop());
    let failed = file
        .batch(
            [BatchItem::new("bad"), BatchItem::new("get")],
            BatchMode::BestEffort,
        )
        .unwrap();
    assert_eq!(failed.diagnostics().operations(), 0);
    assert!(failed.patch().is_noop());
    let cancel = file
        .batch(
            items(
                r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"staged"}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"FIRST_PRIVATE"}}
    ]"#,
            ),
            BatchMode::Atomic,
        )
        .unwrap();
    assert!(cancel.value().iter().all(|r| applied(r).changed));
    assert!(cancel.patch().is_noop());
    assert!(!cancel.diagnostics().changed());
    for commit in [empty, failed, cancel] {
        assert_eq!(commit.file().to_bytes().unwrap(), bytes());
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
    }
}

#[test]
fn input_bound_counts_failures_and_metadata_and_stops_unbounded_iterators() {
    let file = open(false);
    let limits = SemanticLimits::builder()
        .max_operations(2)
        .unwrap()
        .build()
        .unwrap();
    for command in ["bad", "meta"] {
        let pulls = std::cell::Cell::new(0);
        let input = std::iter::from_fn(|| {
            pulls.set(pulls.get() + 1);
            let mut item = BatchItem::new(command);
            if command == "meta" {
                item.dump_version = Some(2);
            }
            Some(item)
        });
        let error = file
            .batch_with_options(
                input,
                BatchMode::BestEffort,
                litchi_ops::EditOptions::new().semantic_limits(limits),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(error.item(), Some(2));
        assert_eq!(pulls.get(), 3);
    }
}

#[test]
fn retained_byte_budget_is_aggregate_and_fatal_in_best_effort() {
    let file = open(false);
    let input =
        items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"FIRST_PRIVATE"}}]"#);
    let charge = file
        .apply(&input[0])
        .unwrap()
        .diagnostics()
        .snapshot_bytes();
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(charge)
        .unwrap()
        .build()
        .unwrap();
    let pulls = std::cell::Cell::new(0);
    let commands = std::iter::repeat(input[0].clone()).inspect(|_| pulls.set(pulls.get() + 1));
    let error = file
        .batch_with_options(
            commands,
            BatchMode::BestEffort,
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.item(), Some(1));
    assert_eq!(pulls.get(), 2);
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn source_changes_during_iterator_execution_are_fatal_and_not_skipped() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, bytes()).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    let mut input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"staged"}},
      {"command":"remove","path":"/body/p[2]"},
      {"command":"add","parent":"/body","type":"paragraph"}
    ]"#,
    )
    .into_iter();
    let pulls = std::cell::Cell::new(0);
    let iter = std::iter::from_fn(|| {
        let item = input.next()?;
        pulls.set(pulls.get() + 1);
        if pulls.get() == 2 {
            std::fs::write(&path, b"changed source").unwrap();
        }
        Some(item)
    });
    let error = file.batch(iter, BatchMode::BestEffort).unwrap_err();
    assert_eq!(error.code(), Code::IoError);
    assert_eq!(pulls.get(), 2);
}

#[test]
fn selectors_resolve_against_prior_successes_and_failed_fanout_is_atomic() {
    for eager in [false, true] {
        let file = open(eager);
        let commit=file.batch(items(r#"[
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Drop"}},
          {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Drop"}},
          {"command":"set","path":"/body/p[text='Drop']","props":{"align":"center","wrong-property":true}},
          {"command":"remove","path":"/body/p[text='Drop']"},
          {"command":"set","path":"r","props":{"bold":true}}
        ]"#),BatchMode::BestEffort).unwrap();
        failed(&commit.value()[2], 2, Code::UnsupportedProperty);
        failed(&commit.value()[4], 4, Code::BareSelectorRejected);
        assert!(applied(&commit.value()[3]).changed);
        let oracle = file
            .add("/body", NewNode::new("paragraph").prop("text", "Drop"))
            .unwrap()
            .file
            .add("/body", NewNode::new("paragraph").prop("text", "Drop"))
            .unwrap()
            .file
            .remove("/body/p[text='Drop']")
            .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
    }
}

#[test]
fn protection_and_read_only_admission_precede_iterator_consumption() {
    let signed_bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    ));
    let signed = litchi_ops::open_bytes(signed_bytes.to_vec()).unwrap();
    let readonly =
        litchi_ops::open_bytes_with(bytes(), &OpenOptions::new().retain_source(false)).unwrap();
    for (file, code) in [
        (signed, Code::DocumentProtected),
        (readonly, Code::UnsupportedCapability),
    ] {
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let iter = std::iter::from_fn(|| -> Option<BatchItem<'static>> {
                panic!("protected batch polled iterator")
            });
            assert_eq!(file.batch(iter, mode).unwrap_err().code(), code);
        }
    }
}

#[test]
fn iterator_panics_cannot_mutate_the_original() {
    let file = open(false);
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let mut input = items(r#"[{"command":"remove","path":"/body/p[1]"}]"#).into_iter();
        let iter = std::iter::from_fn(|| Some(input.next().expect("user iterator panic")));
        let _ = file.batch(iter, BatchMode::BestEffort);
    }));
    assert!(result.is_err());
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn invalid_meta_fields_abort_even_best_effort_and_keep_input_untouched() {
    let file = open(false);
    let error = file
        .batch(
            items(
                r#"[
      {"command":"remove","path":"/body/p[1]"},
      {"command":"meta","dumpVersion":2,"text":"must not be ignored"}
    ]"#,
            ),
            BatchMode::BestEffort,
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    assert_eq!(error.item(), Some(1));
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn command_run_text_controls_follow_literal_policy_with_metadata() {
    let file = open(false);
    for text in ["a\nb", "a\rb", "a\u{b}b"] {
        for meta in [false, true] {
            let mut commands = Vec::new();
            if meta {
                let mut item = BatchItem::new("meta");
                item.dump_version = Some(2);
                commands.push(item);
            }
            let mut set = BatchItem::new("set");
            set.path = Some("/body/p[1]/r[1]".into());
            set.props = Some(litchi_ops::Props::new().with("TEXT", text));
            commands.push(set);
            if text.contains('\u{b}') {
                let error = file.batch(commands, BatchMode::Atomic).unwrap_err();
                assert_eq!(error.item(), Some(usize::from(meta)));
            } else {
                let commit = file.batch(commands, BatchMode::Atomic).unwrap();
                assert_eq!(
                    commit
                        .file()
                        .get("/body/p[1]/r[1]")
                        .unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .as_deref(),
                    Some(text)
                );
                assert_eq!(
                    commit
                        .patch()
                        .inverse()
                        .apply(commit.file())
                        .unwrap()
                        .to_bytes()
                        .unwrap(),
                    bytes()
                );
            }
        }
    }
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn mutation_results_without_created_paths_stay_compact() {
    // Results are retained for every command. An absent path must not embed
    // the selector grammar's large inline path buffer for every simple set.
    assert!(size_of::<ItemResult>() <= 4 * size_of::<usize>());
}

#[test]
fn atomic_formatting_runs_match_sequential_bytes_and_preserve_item_changes() {
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true,"italic":true}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[2]","props":{"align":"center"}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":false,"bold":true}}
    ]"#,
    );
    for eager in [false, true] {
        let file = open(eager);
        let oracle = file.batch(&input, BatchMode::BestEffort).unwrap();
        let grouped = file.batch(&input, BatchMode::Atomic).unwrap();
        assert_eq!(grouped.diagnostics().operations(), 4);
        assert_eq!(grouped.diagnostics().commits(), 1);
        assert_eq!(oracle.diagnostics().commits(), 4);
        assert_eq!(
            grouped
                .value()
                .iter()
                .map(|r| applied(r).changed)
                .collect::<Vec<_>>(),
            vec![true, false, true, false]
        );
        assert_eq!(
            grouped.file().to_bytes().unwrap(),
            oracle.file().to_bytes().unwrap()
        );
        assert_eq!(
            grouped
                .patch()
                .inverse()
                .apply(grouped.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
    }
}

#[test]
fn text_metadata_and_structural_commands_close_formatting_runs() {
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[2]","props":{"align":"center"}},
      {"command":"meta","dumpVersion":2},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"text":"changed"}},
      {"command":"remove","path":"/body/p[1]"},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}}
    ]"#,
    );
    for eager in [false, true] {
        let file = open(eager);
        let grouped = file.batch(&input, BatchMode::Atomic).unwrap();
        let oracle = file.batch(&input, BatchMode::BestEffort).unwrap();
        assert_eq!(grouped.diagnostics().commits(), 5);
        assert_eq!(grouped.diagnostics().operations(), 6);
        assert_eq!(
            grouped.file().to_bytes().unwrap(),
            oracle.file().to_bytes().unwrap()
        );
    }
}

#[test]
fn formatting_failure_discards_pending_run_without_consuming_later_items() {
    let input = items(
        r#"[
      {"command":"meta","dumpVersion":2},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[2]","props":{"align":"center","wrong-property":true}},
      {"command":"remove","path":"/body/p[1]"}
    ]"#,
    );
    let file = open(true);
    let pulls = std::cell::Cell::new(0);
    let error = file
        .batch(
            input.iter().inspect(|_| pulls.set(pulls.get() + 1)),
            BatchMode::Atomic,
        )
        .unwrap_err();
    assert_eq!(error.item(), Some(2));
    assert_eq!(pulls.get(), 3);
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn formatting_flushes_before_native_operation_capacity_and_charges_one_snapshot_pair() {
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":false}}
    ]"#,
    );
    let file = open(true);
    // One more than the native per-edit operation bound (65,536 since
    // record 0251f), which is above the default batch bound.
    let wide = SemanticLimits::builder()
        .max_operations(65_537)
        .unwrap()
        .build()
        .unwrap();
    let grouped = file
        .batch_with_options(
            input.iter().cycle().take(65_537),
            BatchMode::Atomic,
            litchi_ops::EditOptions::new().semantic_limits(wide),
        )
        .unwrap();
    assert_eq!(grouped.diagnostics().operations(), 65_537);
    assert_eq!(grouped.diagnostics().commits(), 2);
    let expected = file.set("/body/p[1]/r[1]", props! {"bold"=>true}).unwrap();
    assert_eq!(
        grouped.file().to_bytes().unwrap(),
        expected.to_bytes().unwrap()
    );
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(2048)
        .unwrap()
        .build()
        .unwrap();
    let compact = file
        .batch_with_options(
            input.iter().cycle().take(100),
            BatchMode::Atomic,
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .unwrap();
    assert_eq!(compact.diagnostics().commits(), 1);
    assert_eq!(
        file.batch_with_options(
            input.iter().cycle().take(100),
            BatchMode::BestEffort,
            litchi_ops::EditOptions::new().semantic_limits(limits)
        )
        .unwrap_err()
        .code(),
        Code::LimitExceeded
    );
}

#[test]
fn source_change_at_end_of_pending_formatting_run_refuses_publication() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, bytes()).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    let mut input =
        items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}}]"#).into_iter();
    let iter = std::iter::from_fn(|| {
        let next = input.next();
        if next.is_none() {
            std::fs::write(&path, b"source changed at EOF").unwrap();
        }
        next
    });
    assert_eq!(
        file.batch(iter, BatchMode::Atomic).unwrap_err().code(),
        Code::IoError
    );
}

#[test]
fn formatting_preserves_opaque_properties_and_uncompacted_text_markup() {
    let xml = XML.replace("<w:r><w:t>FIRST_PRIVATE</w:t>",
        "<w:r><w:rPr><!--opaque--><x:custom xmlns:x=\"urn:test\"/></w:rPr><w:t>FIRST</w:t><w:t>_PRIVATE</w:t>");
    let original = bytes_with_xml(&xml);
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"italic":true}}
    ]"#,
    );
    for mode in [SourceMode::Eager, SourceMode::Auto] {
        let file =
            litchi_ops::open_bytes_with(original.clone(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let grouped = file.batch(&input, BatchMode::Atomic).unwrap();
        let oracle = file.batch(&input, BatchMode::BestEffort).unwrap();
        assert_eq!(
            grouped.file().to_bytes().unwrap(),
            oracle.file().to_bytes().unwrap()
        );
        assert_eq!(
            grouped
                .patch()
                .inverse()
                .apply(grouped.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            original
        );
    }
}

#[test]
fn cached_formatting_target_does_not_bypass_source_revision_fences() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, bytes()).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":false}},
      {"command":"remove","path":"/body/p[2]"}
    ]"#,
    );
    let pulls = std::cell::Cell::new(0);
    let iter = input.into_iter().inspect(|_| {
        pulls.set(pulls.get() + 1);
        if pulls.get() == 2 {
            std::fs::write(&path, b"source changed before cached target").unwrap();
        }
    });
    assert_eq!(
        file.batch(iter, BatchMode::Atomic).unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(pulls.get(), 2);
}

#[test]
fn command_run_add_normalizes_inline_controls_without_splitting_paragraphs() {
    use litchi_docx::{RunBreakType, RunContent};
    for eager in [false, true] {
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let file = open(eager);
            let mut add = BatchItem::new("add");
            add.parent = Some("/body/p[1]".into());
            add.kind = Some("RuN".into());
            add.props = Some(
                litchi_ops::Props::from_props(
                    props! { "TEXT" => "α\r\nβ\rγ\nδ\u{b}ε\t\\n", "bold" => true },
                )
                .unwrap(),
            );
            let commit = file.batch([add], mode).unwrap();
            assert!(
                matches!(&commit.value()[0], ItemResult::Applied(result) if result.path.as_ref().unwrap().to_string() == "/body/p[1]/r[3]")
            );
            let saved = commit.file().to_bytes().unwrap();
            let package = litchi_docx::Package::from_bytes(&saved).unwrap();
            let paragraphs = package.document().unwrap().paragraphs().unwrap();
            assert_eq!(paragraphs.len(), 2);
            let runs = paragraphs[0].runs().unwrap();
            assert_eq!(runs.len(), 3);
            assert_eq!(runs[2].bold().unwrap(), Some(true));
            let contents = runs[2].contents().unwrap();
            assert_eq!(contents.len(), 11);
            for index in [1, 3, 5, 7] {
                assert!(
                    matches!(contents[index], RunContent::Break(value) if value.break_type == RunBreakType::TextWrapping)
                );
            }
            assert!(matches!(contents[9], RunContent::Tab));
            assert!(matches!(&contents[10], RunContent::Text(value) if value == "\\n"));
            assert_eq!(runs[2].text().unwrap(), "α\nβ\nγ\nδ\nε\t\\n");
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
            assert_eq!(file.to_bytes().unwrap(), bytes());
        }
    }
}

#[test]
fn run_add_controls_preserve_atomic_and_best_effort_failure_boundaries() {
    for eager in [false, true] {
        let file = open(eager);
        let mut add = BatchItem::new("add");
        add.parent = Some("/body/p[1]".into());
        add.kind = Some("r".into());
        add.props = Some(litchi_ops::Props::new().with("text", "a\u{b}b"));
        let mut invalid = add.clone();
        invalid.props = Some(litchi_ops::Props::new().with("text", "bad\u{1}"));
        let error = file
            .batch([add.clone(), invalid.clone()], BatchMode::Atomic)
            .unwrap_err();
        assert_eq!(error.item(), Some(1));
        let commit = file.batch([add, invalid], BatchMode::BestEffort).unwrap();
        assert!(matches!(commit.value()[0], ItemResult::Applied(_)));
        assert!(matches!(&commit.value()[1], ItemResult::Failed(error) if error.item() == Some(1)));
        assert_eq!(
            commit
                .file()
                .get("/body/p[1]/r[3]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("a\nb")
        );
        assert!(commit.file().get("/body/p[1]/r[4]").is_err());
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn wire_add_accepts_v2_run_and_paragraph_controls() {
    let file = open(false);
    let commands = items(
        r#"[{"command":"meta","dumpVersion":2},{"command":"add","parent":"/body/p[1]","type":"r","props":{"text":"earlier\r","TEXT":"final\r\n\u000b"}}]"#,
    );
    let batch = file.batch(&commands, BatchMode::Atomic).unwrap();
    let single = file.apply(&commands[1]).unwrap();
    assert_eq!(
        batch.file().to_bytes().unwrap(),
        single.file().to_bytes().unwrap()
    );
    assert_eq!(
        batch
            .file()
            .get("/body/p[1]/r[3]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("final\n\n")
    );
    for kind in ["p", "paragraph"] {
        let mut add = BatchItem::new("add");
        add.parent = Some("/body".into());
        add.kind = Some(kind.into());
        add.props = Some(litchi_ops::Props::new().with("text", "a\nb"));
        assert_eq!(
            file.apply(add)
                .unwrap()
                .into_value()
                .path
                .unwrap()
                .to_string(),
            "/body/p[3]"
        );
    }
}

#[test]
fn direct_paragraph_command_controls_normalize_and_preserve_transaction_boundaries() {
    for eager in [false, true] {
        let file = open(eager);
        let commands = items(
            r#"[{"command":"set","path":"/body/p[2]","props":{"align":"center"}},{"command":"set","path":"/body/Paragraph[2]","props":{"TEXT":"a\r\nb\rc\nd\u000be\tend"}}]"#,
        );
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let commit = file.batch(&commands, mode).unwrap();
            assert_eq!(
                commit
                    .file()
                    .get("/body/p[2]")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("a\nb\nc\nd\ne\tend")
            );
            assert!(
                !commit
                    .file()
                    .apply(&commands[1])
                    .unwrap()
                    .diagnostics()
                    .changed()
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
            for target in ["/body/p, r", "/body"] {
                let mut invalid = commands[1].clone();
                invalid.path = Some(target.into());
                let error = file
                    .batch([commands[1].clone(), invalid.clone()], BatchMode::Atomic)
                    .unwrap_err();
                assert_eq!(error.code(), Code::UnsupportedCapability);
                assert_eq!(error.item(), Some(1));
                let partial = file
                    .batch([commands[1].clone(), invalid], BatchMode::BestEffort)
                    .unwrap();
                assert!(
                    matches!(&partial.value()[1], ItemResult::Failed(error) if error.item() == Some(1))
                );
            }
        }
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn paragraph_selector_commands_normalize_controls_and_preserve_inverse() {
    for eager in [false, true] {
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let file = open(eager);
            let commands = items(
                r#"[{"command":"set","path":"/body/p[1]","props":{"align":"center"}},{"command":"set","path":"/body/p[text!='']","props":{"TEXT":"a\r\nb\rc\u000bd\tend"}}]"#,
            );
            let commit = file.batch(&commands, mode).unwrap();
            for index in [1, 2] {
                assert_eq!(
                    commit
                        .file()
                        .get(format!("/body/p[{index}]").as_str())
                        .unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .as_deref(),
                    Some("a\nb\nc\nd\tend")
                );
            }
            assert!(
                !commit
                    .file()
                    .apply(&commands[1])
                    .unwrap()
                    .diagnostics()
                    .changed()
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
            assert_eq!(file.to_bytes().unwrap(), bytes());
        }
    }
}

#[test]
fn command_text_selector_admission_keeps_diagnostics_and_target_scope() {
    for eager in [false, true] {
        let file = open(eager);
        for (target, code) in [
            ("p", Code::BareSelectorRejected),
            ("/body/p[text='missing']", Code::ZeroMatches),
            ("/body/p[not(unknown=secret)]", Code::InvalidSelector),
            ("/body/p, r", Code::UnsupportedCapability),
        ] {
            let mut set = BatchItem::new("set");
            set.path = Some(target.into());
            set.props = Some(litchi_ops::Props::new().with("text", "a\u{b}b"));
            let error = file.apply(set).unwrap_err();
            assert_eq!(error.code(), code, "{target}");
            assert_eq!(error.item(), Some(0));
        }
        let commands = items(
            r#"[{"command":"set","path":"/body/p[2]","props":{"text":"selected"}},{"command":"set","path":"/body/p[text='selected']","props":{"text":"new\u000btext"}}]"#,
        );
        let commit = file.batch(commands, BatchMode::Atomic).unwrap();
        assert_eq!(
            commit
                .file()
                .get("/body/p[2]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("new\ntext")
        );
        assert_eq!(text(commit.file()), "FIRST_PRIVATE");
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn command_run_text_replaces_first_literal_node_and_preserves_other_content() {
    let source = bytes_with_xml(&XML.replace(
        "<w:t>FIRST_PRIVATE</w:t>",
        "<w:t xml:space=\"preserve\">FIRST_PRIVATE</w:t><w:t>suffix</w:t><w:tab/><w:br/>",
    ));
    let authored = " A\r\nB\rC\nD\tE\u{2011}\u{ad}&< ";
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let mut command = BatchItem::new("set");
        command.path = Some("/body/p[1]/r[1]".into());
        command.props = Some(
            litchi_ops::Props::new()
                .with("TEXT", authored)
                .with("bold", true),
        );
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let commit = file.batch([command.clone()], mode).unwrap();
            assert_eq!(text(commit.file()), format!("{authored}suffix\t\n"));
            assert!(
                !commit
                    .file()
                    .apply(&command)
                    .unwrap()
                    .diagnostics()
                    .changed()
            );
            assert_eq!(
                commit.patch().apply(&file).unwrap().to_bytes().unwrap(),
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
                source
            );
            let reopened = litchi_ops::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
            assert_eq!(text(&reopened), format!("{authored}suffix\t\n"));
        }
        let native = file
            .set("/body/p[1]/r[1]", props! {"text" => "native\ntext"})
            .unwrap();
        assert_eq!(text(&native), "native\ntext");
        assert_eq!(file.to_bytes().unwrap(), source);
    }
}

#[test]
fn command_run_literal_text_obeys_selector_and_failure_boundaries() {
    for eager in [false, true] {
        let file = open(eager);
        let commands = items(
            r#"[{"command":"set","path":"/body/p/r[text!='']","props":{"text":"literal\r\nline\tend"}},{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"invalid\u000btext"}},{"command":"set","path":"/body/p[2]/r[1]","props":{"text":"last"}}]"#,
        );
        let error = file.batch(&commands, BatchMode::Atomic).unwrap_err();
        assert_eq!(error.item(), Some(1));
        let commit = file.batch(&commands, BatchMode::BestEffort).unwrap();
        assert_eq!(text(commit.file()), "literal\r\nline\tend");
        assert!(matches!(&commit.value()[1], ItemResult::Failed(error) if error.item() == Some(1)));
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn command_run_text_refuses_special_content_and_empty_run_is_noop() {
    for eager in [false, true] {
        for child in [
            "<w:ptab/>",
            "<w:fldChar w:fldCharType=\"begin\"/>",
            "<w:instrText>PAGE</w:instrText>",
            "<w:drawing/>",
            "<w:pict/>",
            "<w:object/>",
        ] {
            let source = bytes_with_xml(&XML.replace(
                "<w:t>FIRST_PRIVATE</w:t>",
                &format!("<w:t>FIRST_PRIVATE</w:t>{child}"),
            ));
            let file = litchi_ops::open_bytes_with(
                source.clone(),
                &OpenOptions::new().source_mode(if eager {
                    SourceMode::Eager
                } else {
                    SourceMode::Auto
                }),
            )
            .unwrap();
            let command =
                items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"literal"}}]"#);
            assert!(file.apply(&command[0]).is_err(), "accepted {child}");
            assert_eq!(file.to_bytes().unwrap(), source);
        }
        let source = bytes_with_xml(&XML.replace("<w:t>FIRST_PRIVATE</w:t>", ""));
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let command =
            items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"literal"}}]"#);
        let commit = file.apply(&command[0]).unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.file().to_bytes().unwrap(), source);
    }
}

#[test]
fn command_run_literal_text_keeps_paragraph_local_namespace_context() {
    let xml = XML
        .replacen(
            "<w:p>",
            "<w:p xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
            1,
        )
        .replacen(
            "<w:r><w:t>FIRST_PRIVATE</w:t></w:r>",
            "<q:r><q:t>FIRST_PRIVATE</q:t><t xmlns=\"\">foreign</t></q:r>",
            1,
        );
    let source = bytes_with_xml(&xml);
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let command =
            items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"literal\r\n"}}]"#);
        let commit = file.apply(&command[0]).unwrap();
        assert_eq!(text(commit.file()), "literal\r\n");
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source
        );
    }
}

#[test]
fn command_run_literal_text_expands_an_empty_text_node() {
    let source = bytes_with_xml(&XML.replace(
        "<w:t>FIRST_PRIVATE</w:t>",
        "<w:t xml:space=\"preserve\"/><w:t>tail-node</w:t>",
    ));
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let command =
            items(r#"[{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"literal\r\n"}}]"#);
        let commit = file.apply(&command[0]).unwrap();
        assert_eq!(text(commit.file()), "literal\r\ntail-node");
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source
        );
    }
}

#[test]
fn paragraph_add_commands_split_blocks_copy_formatting_and_return_first_path() {
    let commands = items(
        r#"[{"command":"add","parent":"/body","type":"paragraph","after":"/body/p[1]","props":{"text":"\r\nA\rB\nC\u000bD\t\n","bold":true,"align":"center"}}]"#,
    );
    let lines = ["", "A", "B", "C\nD\t", ""];
    for eager in [false, true] {
        let file = open(eager);
        let mut oracle = file.clone();
        for (index, line) in lines.iter().enumerate() {
            oracle = oracle
                .add(
                    "/body",
                    NewNode::new("paragraph")
                        .prop("text", *line)
                        .prop("bold", true)
                        .prop("align", "center")
                        .at(Placement::after(&format!("/body/p[{}]", index + 1)).unwrap()),
                )
                .unwrap()
                .file;
        }
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let commit = file.batch(&commands, mode).unwrap();
            assert_eq!(
                applied(&commit.value()[0])
                    .path
                    .as_ref()
                    .unwrap()
                    .to_string(),
                "/body/p[2]"
            );
            assert_eq!(commit.diagnostics().operations(), 1);
            assert_eq!(commit.diagnostics().leaf_operations(), lines.len());
            assert_eq!(commit.diagnostics().commits(), 1);
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                oracle.to_bytes().unwrap()
            );
            assert_eq!(
                commit.patch().apply(&file).unwrap().to_bytes().unwrap(),
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
                bytes()
            );
            let reopened = litchi_ops::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
            for (index, line) in lines.iter().enumerate() {
                assert_eq!(
                    reopened
                        .get(&format!("/body/p[{}]", index + 2))
                        .unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .as_deref(),
                    Some(*line)
                );
            }
        }
        let native = file
            .add("/body", NewNode::new("paragraph").prop("text", "a\nb"))
            .unwrap();
        assert_eq!(
            native
                .file
                .get("/body/p[3]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("a\nb")
        );
        assert!(native.file.get("/body/p[4]").is_err());
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn paragraph_split_is_one_atomic_item_and_later_paths_see_all_insertions() {
    for eager in [false, true] {
        let file = open(eager);
        let commands = items(
            r#"[{"command":"add","parent":"/body","type":"p","before":"/body/p[2]","props":{"text":"first\nsecond"}},{"command":"set","path":"/body/p[3]","props":{"text":"changed"}}]"#,
        );
        let commit = file.batch(&commands, BatchMode::Atomic).unwrap();
        assert_eq!(
            commit
                .file()
                .get("/body/p[3]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("changed")
        );
        assert_eq!(
            commit
                .file()
                .get("/body/p[4]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("second")
        );
        let invalid = items(
            r#"[{"command":"add","parent":"/body","type":"p","props":{"text":"valid\ninvalid\u0001"}},{"command":"set","path":"/body/p[1]/r[1]","props":{"text":"survives"}}]"#,
        );
        assert_eq!(
            file.batch(&invalid, BatchMode::Atomic).unwrap_err().item(),
            Some(0)
        );
        let partial = file.batch(&invalid, BatchMode::BestEffort).unwrap();
        assert!(
            matches!(&partial.value()[0], ItemResult::Failed(error) if error.item() == Some(0))
        );
        assert!(partial.file().get("/body/p[3]").is_err());
        assert_eq!(text(partial.file()), "survives");
        assert_eq!(
            partial
                .patch()
                .inverse()
                .apply(partial.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
    }
}

#[test]
fn paragraph_split_refuses_native_capacity_before_staging() {
    let file = open(false);
    let mut add = BatchItem::new("add");
    add.parent = Some("/body".into());
    add.kind = Some("p".into());
    // 65,537 paragraphs: one more than the native operation bound.
    add.props = Some(litchi_ops::Props::new().with("text", "\n".repeat(65_536)));
    for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
        let error = file.batch([add.clone()], mode).unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(error.item(), Some(0));
    }
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

fn main_document_xml(file: &File) -> String {
    use std::io::Read as _;
    let bytes = file.to_bytes().unwrap();
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("word/document.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

fn cached_field_bytes() -> Vec<u8> {
    bytes_with_xml(&XML.replace("<w:r><w:t>FIRST_PRIVATE</w:t></w:r>", r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="false"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#))
}

#[test]
fn command_cached_run_text_marks_owner_dirty_even_for_unchanged_text() {
    let source = cached_field_bytes();
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        for value in ["cached", "new\r\nresult"] {
            let mut command = BatchItem::new("set");
            command.path = Some("/body/p[1]/r[4]".into());
            command.props = Some(litchi_ops::Props::new().with("text", value));
            for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
                let commit = file.batch([command.clone()], mode).unwrap();
                assert!(commit.diagnostics().changed());
                assert_eq!(commit.diagnostics().operations(), 1);
                assert_eq!(commit.diagnostics().commits(), 1);
                assert_eq!(
                    commit.diagnostics().leaf_operations(),
                    if value == "cached" { 1 } else { 2 }
                );
                let xml = main_document_xml(commit.file());
                assert!(xml.contains("w:dirty=\"true\"") || xml.contains("w:dirty=\"1\""));
                assert!(xml.contains("<w:instrText>PAGE</w:instrText>"));
                assert_eq!(
                    commit
                        .file()
                        .get("/body/p[1]/r[4]")
                        .unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .as_deref(),
                    Some(value)
                );
                assert!(
                    !commit
                        .file()
                        .apply(&command)
                        .unwrap()
                        .diagnostics()
                        .changed()
                );
                assert_eq!(
                    commit.patch().apply(&file).unwrap().to_bytes().unwrap(),
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
                    source
                );
                let reopened = litchi_ops::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
                assert_eq!(main_document_xml(&reopened), xml);
            }
        }
        let native = file
            .set("/body/p[1]/r[4]", props! {"text" => "native"})
            .unwrap();
        assert!(main_document_xml(&native).contains("w:dirty=\"false\""));
        assert_eq!(file.to_bytes().unwrap(), source);
    }
}

#[test]
fn cached_field_marker_and_text_rollback_together() {
    let source = cached_field_bytes();
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let commands = items(
            r#"[{"command":"set","path":"/body/p[1]/r[4]","props":{"text":"changed"}},{"command":"set","path":"/body/p[1]/r[4]","props":{"text":"bad\u000b"}}]"#,
        );
        assert_eq!(
            file.batch(&commands, BatchMode::Atomic).unwrap_err().item(),
            Some(1)
        );
        let partial = file.batch(&commands, BatchMode::BestEffort).unwrap();
        assert!(
            matches!(&partial.value()[1], ItemResult::Failed(error) if error.item() == Some(1))
        );
        assert_eq!(
            partial
                .file()
                .get("/body/p[1]/r[4]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("changed")
        );
        assert_eq!(
            partial
                .patch()
                .inverse()
                .apply(partial.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source
        );
        assert_eq!(file.to_bytes().unwrap(), source);
    }
}

#[test]
fn cached_run_command_skips_closed_nested_field_and_stops_at_end() {
    let source = bytes_with_xml(&XML.replace("<w:r><w:t>FIRST_PRIVATE</w:t></w:r>", r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="false"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="begin" w:dirty="false"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>inner</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>outer</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#));
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            source.clone(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let outer =
            items(r#"[{"command":"set","path":"/body/p[1]/r[7]","props":{"text":"changed"}}]"#);
        let commit = file.apply(&outer[0]).unwrap();
        let xml = main_document_xml(commit.file());
        let begins: Vec<_> = xml
            .split("<w:fldChar")
            .skip(1)
            .map(|s| s.split("/>").next().unwrap())
            .filter(|s| s.contains("\"begin\""))
            .collect();
        assert_eq!(begins.len(), 2);
        assert!(begins[0].contains("w:dirty=\"true\"") || begins[0].contains("w:dirty=\"1\""));
        assert!(begins[1].contains("w:dirty=\"false\""));
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source
        );
        let inner = items(
            r#"[{"command":"set","path":"/body/p[1]/r[5]","props":{"text":"changed inner"}}]"#,
        );
        let inner_commit = file.apply(&inner[0]).unwrap();
        let inner_xml = main_document_xml(inner_commit.file());
        let inner_begins: Vec<_> = inner_xml
            .split("<w:fldChar")
            .skip(1)
            .map(|s| s.split("/>").next().unwrap())
            .filter(|s| s.contains("\"begin\""))
            .collect();
        assert!(inner_begins[0].contains("w:dirty=\"false\""));
        assert!(
            inner_begins[1].contains("w:dirty=\"true\"")
                || inner_begins[1].contains("w:dirty=\"1\"")
        );
        assert_eq!(
            inner_commit
                .patch()
                .inverse()
                .apply(inner_commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source
        );
        let outside =
            items(r#"[{"command":"set","path":"/body/p[1]/r[9]","props":{"text":"outside"}}]"#);
        let commit = file.apply(&outside[0]).unwrap();
        assert_eq!(commit.diagnostics().leaf_operations(), 1);
        assert_eq!(
            main_document_xml(commit.file())
                .matches("w:dirty=\"false\"")
                .count(),
            2
        );
    }
}

/// Record 0355a: `text` on direct body paragraphs joins the coalesced run.
/// A batch still equals its items applied one at a time, including run
/// paths inside a paragraph whose text the run replaced (a barrier), a
/// repeated paragraph, formatting between texts, both owners, both modes,
/// and the closure edit; a refused item keeps its wire index.
#[test]
fn coalesced_paragraph_text_sets_match_single_operations() {
    let input = items(
        r#"[
      {"command":"set","path":"/body/p[1]","props":{"text":"one"}},
      {"command":"set","path":"/body/p[2]","props":{"align":"right"}},
      {"command":"set","path":"/body/p[2]","props":{"text":"two"}},
      {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
      {"command":"set","path":"/body/p[1]","props":{"text":"again"}},
      {"command":"set","path":"/body/p[2]/r[1]","props":{"text":"run"}},
      {"command":"set","path":"/body/p[1]","props":{"align":"center"}},
      {"command":"set","path":"/body/p[2]","props":{"text":"last"}}
    ]"#,
    );
    for eager in [false, true] {
        let file = open(eager);
        let mut oracle = file.clone();
        for item in &input {
            let props = item.props.clone().unwrap();
            oracle = oracle.set(item.path.as_deref().unwrap(), props).unwrap();
        }
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let commit = file.batch(&input, mode).unwrap();
            assert!(commit.value().iter().all(|result| applied(result).changed));
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                oracle.to_bytes().unwrap()
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
        }
        let edited = file
            .edit(|tx| {
                for item in &input {
                    tx.set(item.path.as_deref().unwrap(), item.props.clone().unwrap())?;
                }
                Ok(())
            })
            .unwrap()
            .into_file();
        assert_eq!(edited.to_bytes().unwrap(), oracle.to_bytes().unwrap());

        // A refused text keeps its own wire index; best effort keeps the rest.
        let refused = items(
            r#"[
          {"command":"set","path":"/body/p[1]","props":{"text":"kept"}},
          {"command":"set","path":"/body/p[2]","props":{"text":"bad\u000bvalue"}},
          {"command":"set","path":"/body/p[2]","props":{"text":"after"}}
        ]"#,
        );
        let error = file.batch(&refused, BatchMode::Atomic).unwrap_err();
        assert_eq!(error.item(), Some(1));
        let code = file
            .set("/body/p[2]", props! {"text"=>"bad\u{b}value"})
            .unwrap_err()
            .code();
        assert_eq!(error.code(), code);
        let commit = file.batch(&refused, BatchMode::BestEffort).unwrap();
        applied(&commit.value()[0]);
        failed(&commit.value()[1], 1, code);
        applied(&commit.value()[2]);
        let expected = file
            .set("/body/p[1]", props! {"text"=>"kept"})
            .unwrap()
            .set("/body/p[2]", props! {"text"=>"after"})
            .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            expected.to_bytes().unwrap()
        );
    }
}
