//! Public transaction coverage for consecutive direct legacy-PPT shape sets.
//!
//! The direct path form is intentionally narrow: the transaction coordinator
//! may retain a native edit for distinct top-level shapes, while a query,
//! read, slide operation, repeated target, or structural command closes that
//! run.  These tests exercise the public `File`/`Edit`/`batch` APIs for both
//! eager and positional owners.

#![cfg(feature = "ppt")]

use std::path::PathBuf;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::node::Value;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Props, SourceMode, props};
use litchi_ppt::writer::Builder;

const FIRST: &str = "/slide[1]/shape[1]";
const SECOND: &str = "/slide[1]/shape[2]";
const DIRECT_COUNT: usize = 1_000;
const NATIVE_ROLLOVER: usize = 16_384;

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new();
    let first = builder.add_slide().expect("first slide");
    builder
        .set_slide_timing(first, litchi_ppt::writer::SlideTiming::default())
        .expect("first timing");
    builder
        .add_textbox(first, 10, 10, 240, 40, "first")
        .expect("first textbox");
    builder
        .add_textbox(first, 10, 70, 240, 40, "second")
        .expect("second textbox");

    let second = builder.add_slide().expect("second slide");
    builder
        .set_slide_timing(second, litchi_ppt::writer::SlideTiming::default())
        .expect("second timing");
    builder
        .add_textbox(second, 10, 10, 240, 40, "other slide")
        .expect("other slide textbox");

    builder
        .build()
        .expect("build PPT fixture")
        .to_bytes()
        .expect("serialize PPT fixture")
}

fn many_shapes(count: usize) -> Vec<u8> {
    let mut builder = Builder::new();
    let slide = builder.add_slide().expect("slide");
    builder
        .set_slide_timing(slide, litchi_ppt::writer::SlideTiming::default())
        .expect("timing");
    for index in 0..count {
        builder
            .add_textbox(slide, 10 + (index % 100) as i32, 10, 240, 40, "before")
            .expect("textbox");
    }
    builder
        .build()
        .expect("build many-shape fixture")
        .to_bytes()
        .expect("serialize many-shape fixture")
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn positional(bytes: &[u8]) -> (tempfile::TempDir, PathBuf, File) {
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("source.ppt");
    std::fs::write(&path, bytes).expect("write PPT source");
    let file = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
        .expect("open positional PPT");
    (directory, path, file)
}

fn each_owner(bytes: &[u8], mut test: impl FnMut(File)) {
    test(eager(bytes));
    let (_directory, _path, file) = positional(bytes);
    test(file);
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .expect("text-bearing shape")
        .into_owned()
}

fn length(file: &File, path: &str, key: &str) -> i64 {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"));
    let property = node
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .unwrap_or_else(|| panic!("missing {key} on {path}"));
    let Value::Length(value) = property.value() else {
        panic!("{key} is not a length")
    };
    value.emu()
}

fn scoped_shapes() -> Query<'static> {
    Query::parse("shape")
        .expect("shape query")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/slide[1]").expect("slide scope"),
        )))
}

fn batch_set(
    path: impl Into<std::borrow::Cow<'static, str>>,
    props: Props<'static>,
) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

fn direct_items(count: usize) -> Vec<BatchItem<'static>> {
    (1..=count)
        .map(|index| {
            batch_set(
                format!("/slide[1]/shape[{index}]"),
                Props::new().with("text", format!("after-{index}")),
            )
        })
        .collect()
}

#[test]
fn edit_and_batch_coalesce_one_thousand_distinct_direct_text_sets_for_both_owners() {
    let bytes = many_shapes(DIRECT_COUNT);
    let native = litchi_ppt::Package::from_bytes(&bytes).expect("native source");
    let replacements = (0..DIRECT_COUNT)
        .map(|index| {
            litchi_ppt::slide_order::ShapeTextReplacement::new(
                litchi_ppt::text_edit::Target::new(
                    litchi_core::Position::new(0),
                    litchi_core::Position::new(index),
                ),
                format!("after-{}", index + 1),
            )
        })
        .collect::<Vec<_>>();
    let chunks = replacements
        .chunks(litchi_ppt::slide_order::MAX_SHAPE_TEXT_REPLACEMENTS)
        .collect::<Vec<_>>();
    let mut native_edit = native.edit().expect("native edit");
    assert_eq!(
        native_edit
            .set_shape_text_chunks(&chunks)
            .expect("one native publication"),
        DIRECT_COUNT
    );
    let native_output = native_edit
        .commit()
        .expect("native commit")
        .package()
        .to_bytes()
        .expect("native bytes");

    let run_edit = |source: File| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .edit(|edit| {
                for index in 1..=DIRECT_COUNT {
                    edit.set(
                        format!("/slide[1]/shape[{index}]").as_str(),
                        Props::new().with("text", format!("after-{index}")),
                    )?;
                }
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("one-thousand direct edit");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), DIRECT_COUNT);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(text(commit.file(), "/slide[1]/shape[1]"), "after-1");
        assert_eq!(
            text(commit.file(), &format!("/slide[1]/shape[{DIRECT_COUNT}]")),
            format!("after-{DIRECT_COUNT}")
        );
        assert_eq!(source.to_bytes().expect("immutable source"), before);
        assert_eq!(
            commit.file().to_bytes().expect("facade bytes"),
            native_output
        );
        assert_eq!(
            commit
                .patch()
                .apply(&source)
                .expect("forward replay")
                .to_bytes()
                .expect("replayed bytes"),
            native_output
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse replay")
                .to_bytes()
                .expect("restored bytes"),
            before
        );
        commit.into_file()
    };

    let edit_eager = run_edit(eager(&bytes));
    assert_eq!(text(&edit_eager, "/slide[1]/shape[500]"), "after-500");

    let (_directory, path, source) = positional(&bytes);
    let before_path = std::fs::read(&path).expect("source file bytes");
    let edit_positional = run_edit(source);
    assert_eq!(text(&edit_positional, "/slide[1]/shape[500]"), "after-500");
    assert_eq!(
        std::fs::read(path).expect("unchanged source file"),
        before_path
    );

    let run_batch = |source: File| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .batch(direct_items(DIRECT_COUNT), BatchMode::Atomic)
            .expect("one-thousand direct batch");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), DIRECT_COUNT);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(commit.value().len(), DIRECT_COUNT);
        assert!(commit.value().iter().all(
            |result| matches!(result, litchi_ops::ItemResult::Applied(applied) if applied.changed)
        ));
        assert_eq!(text(commit.file(), "/slide[1]/shape[1]"), "after-1");
        assert_eq!(
            text(commit.file(), &format!("/slide[1]/shape[{DIRECT_COUNT}]")),
            format!("after-{DIRECT_COUNT}")
        );
        assert_eq!(source.to_bytes().expect("immutable batch source"), before);
        commit.into_file()
    };

    let batch_eager = run_batch(eager(&bytes));
    assert_eq!(text(&batch_eager, "/slide[1]/shape[500]"), "after-500");

    let (_directory, path, source) = positional(&bytes);
    let before_path = std::fs::read(&path).expect("batch source file bytes");
    let batch_positional = run_batch(source);
    assert_eq!(text(&batch_positional, "/slide[1]/shape[500]"), "after-500");
    assert_eq!(
        std::fs::read(path).expect("unchanged batch source file"),
        before_path
    );
}

#[test]
fn mixed_text_and_geometry_stays_in_one_edit_for_eager_and_positional_owners() {
    each_owner(&fixture(), |source| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "text" => "changed" })?;
                edit.set(
                    SECOND,
                    props! {
                        "x" => "1in",
                        "y" => "2in",
                        "width" => "3in",
                        "height" => "4in",
                    },
                )
            })
            .expect("mixed direct edit");

        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(text(commit.file(), FIRST), "changed");
        assert_eq!(length(commit.file(), SECOND, "x"), 914_400);
        assert_eq!(length(commit.file(), SECOND, "height"), 3_657_600);
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
        assert_eq!(source.to_bytes().expect("immutable source"), before);
    });
}

#[test]
fn repeated_direct_target_closes_run_and_replans_partial_geometry() {
    each_owner(&fixture(), |source| {
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "x" => "1in" })?;
                edit.set(FIRST, props! { "width" => "2in", "height" => "3in" })?;
                edit.set(SECOND, props! { "text" => "after barrier" })
            })
            .expect("repeated target edit");

        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 2);
        assert_eq!(length(commit.file(), FIRST, "x"), 914_400);
        assert_eq!(length(commit.file(), FIRST, "width"), 1_828_800);
        assert_eq!(length(commit.file(), FIRST, "height"), 2_743_200);
        assert_eq!(text(commit.file(), SECOND), "after barrier");
    });
}

#[test]
fn query_read_slide_and_structural_commands_are_public_coalescing_barriers() {
    let bytes = fixture();

    let read = eager(&bytes)
        .edit(|edit| {
            edit.set(FIRST, props! { "text" => "before read" })?;
            assert_eq!(
                edit.get(FIRST)?.view().text()?.as_deref(),
                Some("before read")
            );
            edit.set(SECOND, props! { "text" => "after read" })
        })
        .expect("read barrier");
    assert_eq!(read.diagnostics().commits(), 2);
    assert_eq!(text(read.file(), SECOND), "after read");

    let query = eager(&bytes)
        .edit(|edit| {
            edit.set(FIRST, props! { "text" => "before query" })?;
            edit.set(scoped_shapes(), props! { "text" => "query result" })?;
            edit.set(SECOND, props! { "text" => "after query" })
        })
        .expect("query barrier");
    assert_eq!(query.diagnostics().commits(), 3);
    assert_eq!(text(query.file(), FIRST), "query result");
    assert_eq!(text(query.file(), SECOND), "after query");

    let slide = eager(&bytes)
        .edit(|edit| {
            edit.set(FIRST, props! { "text" => "before slide" })?;
            edit.set("/slide[1]", props! { "hidden" => true })?;
            edit.set(SECOND, props! { "text" => "after slide" })
        })
        .expect("slide barrier");
    assert_eq!(slide.diagnostics().commits(), 3);
    assert_eq!(text(slide.file(), SECOND), "after slide");

    let structural = eager(&bytes)
        .edit(|edit| {
            edit.set(FIRST, props! { "text" => "before remove" })?;
            edit.remove("/slide[2]")?;
            edit.set(FIRST, props! { "text" => "after remove" })
        })
        .expect("structural barrier");
    assert_eq!(structural.diagnostics().commits(), 3);
    assert!(structural.file().get("/slide[2]").is_err());
    assert_eq!(text(structural.file(), FIRST), "after remove");
}

#[test]
fn explicit_raw_ordinal_paths_close_direct_runs() {
    each_owner(&fixture(), |source| {
        let raw = litchi_ops::path::raw::path(SECOND).expect("raw ordinal path");
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "text" => "before raw" })?;
                edit.set(raw, props! { "text" => "raw target" })?;
                edit.set("/slide[2]/shape[1]", props! { "text" => "after raw" })
            })
            .expect("raw path barrier");
        assert_eq!(commit.diagnostics().commits(), 3);
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(text(commit.file(), FIRST), "before raw");
        assert_eq!(text(commit.file(), SECOND), "raw target");
        assert_eq!(text(commit.file(), "/slide[2]/shape[1]"), "after raw");
        assert_eq!(text(&source, SECOND), "second");
    });
}

#[test]
fn batch_reports_exact_noop_and_changed_results_without_mutating_input() {
    each_owner(&fixture(), |source| {
        let before = source.to_bytes().expect("source bytes");
        let mut changed = BatchItem::new("set");
        changed.path = Some(SECOND.into());
        changed.props = Some(Props::new().with("text", "changed"));
        let commit = source
            .batch(
                [
                    batch_set(FIRST, Props::new().with("text", "first")),
                    changed,
                ],
                BatchMode::Atomic,
            )
            .expect("no-op plus changed batch");

        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert!(matches!(
            &commit.value()[0],
            litchi_ops::ItemResult::Applied(applied) if !applied.changed
        ));
        assert!(matches!(
            &commit.value()[1],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(text(commit.file(), FIRST), "first");
        assert_eq!(text(commit.file(), SECOND), "changed");
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn atomic_and_best_effort_batches_keep_their_public_failure_contract() {
    each_owner(&fixture(), |source| {
        let make_items = || {
            [
                batch_set(FIRST, Props::new().with("text", "atomic first")),
                batch_set(FIRST, Props::new().with("id", 1)),
                batch_set(SECOND, Props::new().with("text", "atomic third")),
            ]
        };
        let before = source.to_bytes().expect("source bytes");
        let error = source
            .batch(make_items(), BatchMode::Atomic)
            .expect_err("atomic invalid item");
        assert_eq!(error.item(), Some(1));
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(source.to_bytes().expect("atomic rollback"), before);

        let commit = source
            .batch(make_items(), BatchMode::BestEffort)
            .expect("best effort invalid item");
        assert_eq!(commit.value().len(), 3);
        assert!(matches!(
            &commit.value()[0],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &commit.value()[1],
            litchi_ops::ItemResult::Failed(error) if error.item() == Some(1)
                && error.code() == Code::UnsupportedProperty
        ));
        assert!(matches!(
            &commit.value()[2],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(commit.diagnostics().commits(), 2);
        assert_eq!(text(commit.file(), FIRST), "atomic first");
        assert_eq!(text(commit.file(), SECOND), "atomic third");
        assert_eq!(source.to_bytes().expect("best effort source"), before);
    });
}

#[test]
fn poisoned_edits_retain_the_first_wire_index_and_discard_staged_changes() {
    each_owner(&fixture(), |source| {
        let before = source.to_bytes().expect("source bytes");
        let result = source.edit(|edit| {
            edit.set(FIRST, props! { "text" => "discarded" })?;
            let first_error = edit
                .set(FIRST, props! { "id" => 1 })
                .expect_err("unsupported property");
            assert_eq!(first_error.item(), Some(1));
            assert_eq!(first_error.code(), Code::UnsupportedProperty);
            let repeated_error = edit
                .set(SECOND, props! { "text" => "also discarded" })
                .expect_err("poisoned transaction");
            assert_eq!(repeated_error.item(), Some(1));
            assert_eq!(repeated_error.code(), Code::UnsupportedProperty);
            Ok::<_, litchi_ops::Error>(())
        });
        let error = result.expect_err("poisoned edit");
        assert_eq!(error.item(), Some(1));
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(source.to_bytes().expect("poison rollback"), before);
    });

    let limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("operation limit")
        .build()
        .expect("limits");
    let source = eager(&fixture());
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.set(FIRST, props! { "text" => "wire zero" })?;
                edit.set(SECOND, props! { "text" => "wire one" })
            },
        )
        .expect_err("second wire item exceeds operation limit");
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn explicit_limits_and_native_operation_rollover_preserve_atomicity() {
    let source = eager(&fixture());
    let before = source.to_bytes().expect("source bytes");
    for limits in [
        SemanticLimits::builder()
            .max_staged_bytes(1)
            .expect("staging limit")
            .build()
            .expect("limits"),
        SemanticLimits::builder()
            .max_snapshot_bytes(1)
            .expect("snapshot limit")
            .build()
            .expect("limits"),
    ] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.set(FIRST, props! { "text" => "bounded" }),
            )
            .expect_err("explicit PPT limit");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().expect("limit rollback"), before);
    }

    // One past the native primary-edit bound (16384 since record 0264)
    // exceeds the default unified operation bound, so select wider limits.
    let bytes = many_shapes(NATIVE_ROLLOVER + 1);
    let source = eager(&bytes);
    let before = source.to_bytes().expect("rollover source");
    let wide = SemanticLimits::builder()
        .max_operations(1 << 20)
        .expect("operation limit")
        .max_snapshot_bytes(1 << 30)
        .expect("snapshot limit")
        .max_staged_bytes(1 << 30)
        .expect("staging limit")
        .build()
        .expect("limits");
    let commit = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(wide),
            |edit| {
                for index in 1..=NATIVE_ROLLOVER + 1 {
                    edit.set(
                        format!("/slide[1]/shape[{index}]").as_str(),
                        Props::new().with("text", format!("rollover-{index}")),
                    )?;
                }
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect("native operation rollover");
    assert_eq!(commit.diagnostics().operations(), NATIVE_ROLLOVER + 1);
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(text(commit.file(), "/slide[1]/shape[1]"), "rollover-1");
    assert_eq!(
        text(
            commit.file(),
            &format!("/slide[1]/shape[{}]", NATIVE_ROLLOVER + 1)
        ),
        format!("rollover-{}", NATIVE_ROLLOVER + 1)
    );
    assert_eq!(
        source.to_bytes().expect("rollover source unchanged"),
        before
    );
}
