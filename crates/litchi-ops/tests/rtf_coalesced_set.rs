//! Public transaction coverage for direct RTF `set` coalescing.
//!
//! RTF has one source-backed native transaction for an ordered run of direct
//! paragraph or run assignments.  The public coordinator still has to close
//! that run at every semantic boundary, and it must preserve the result that
//! the same commands would have produced one at a time.

#![cfg(feature = "rtf")]

use std::path::PathBuf;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, OpenOptions, Placement, Props, SourceMode, open_bytes,
    open_bytes_with, props,
};

const RICH_INPUT: &[u8] = br#"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}\b First\b0 plain\par\qc Second\par\sb120\keep\i Third\i0 tail\par Fourth}"#;

const INPUT: &[u8] = br"{\rtf1\ansi Firstplain\par Second\par Thirdtail\par Fourth}";

fn eager(bytes: &[u8]) -> File {
    open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager RTF")
}

fn positional(bytes: &[u8]) -> (tempfile::TempDir, PathBuf, File) {
    let directory = tempfile::tempdir().expect("RTF source directory");
    let path = directory.path().join("source.rtf");
    std::fs::write(&path, bytes).expect("write RTF source");
    let file = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
        .expect("open positional RTF");
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
        .expect("text-bearing RTF node")
        .into_owned()
}

fn property(file: &File, path: &str, key: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .unwrap_or_else(|| panic!("missing {key} on {path}"))
        .to_string()
}

fn query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("valid RTF query")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/body").expect("body scope"),
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

#[test]
fn ordered_direct_paragraph_text_sets_share_one_publication_for_both_owners() {
    let bytes = INPUT;
    each_owner(bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let oracle = source
            .set("/body/p[1]", props! {"text" => "First changed"})
            .expect("paragraph one")
            .set("/body/p[2]", props! {"text" => "Second changed"})
            .expect("paragraph two")
            .set("/body/p[3]", props! {"text" => "Third changed"})
            .expect("paragraph three")
            .set("/body/p[4]", props! {"text" => "Fourth changed"})
            .expect("paragraph four");

        let commit = source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"text" => "First changed"})?;
                edit.set("/body/p[2]", props! {"text" => "Second changed"})?;
                edit.set("/body/p[3]", props! {"text" => "Third changed"})?;
                edit.set("/body/p[4]", props! {"text" => "Fourth changed"})
            })
            .expect("ordered paragraph edit");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            commit.file().to_bytes().expect("coalesced bytes"),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(text(commit.file(), "/body/p[1]"), "First changed");
        assert_eq!(text(commit.file(), "/body/p[2]"), "Second changed");
        assert_eq!(text(commit.file(), "/body/p[3]"), "Third changed");
        assert_eq!(text(commit.file(), "/body/p[4]"), "Fourth changed");
        assert_eq!(source.to_bytes().expect("source remains immutable"), before);
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
    });
}

#[test]
fn ordered_direct_paragraph_format_sets_share_one_publication() {
    each_owner(INPUT, |source| {
        let before = source.to_bytes().unwrap();
        let oracle = source
            .set("/body/p[1]", props! {"align" => "right"})
            .unwrap()
            .set("/body/p[2]", props! {"spaceBefore" => "6pt"})
            .unwrap()
            .set("/body/p[3]", props! {"keepNext" => true})
            .unwrap()
            .set("/body/p[4]", props! {"align" => "center"})
            .unwrap();
        let commit = source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"align" => "right"})?;
                edit.set("/body/p[2]", props! {"spaceBefore" => "6pt"})?;
                edit.set("/body/p[3]", props! {"keepNext" => true})?;
                edit.set("/body/p[4]", props! {"align" => "center"})
            })
            .expect("ordered paragraph formatting edit");
        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        // Sequential native formatting publication retains redundant control
        // prefixes. Compare semantics with it and exact bytes with one native
        // bulk transaction, which is the coalesced publication contract.
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.set_body_paragraph_alignments(&[
            litchi_rtf::edit::ParagraphAlignmentUpdate::new(
                litchi_core::Position::new(0),
                litchi_rtf::Alignment::Right,
            ),
            litchi_rtf::edit::ParagraphAlignmentUpdate::new(
                litchi_core::Position::new(3),
                litchi_rtf::Alignment::Center,
            ),
        ])
        .unwrap();
        edit.patch_body_paragraph_layouts(&[
            litchi_rtf::edit::ParagraphLayoutUpdate::new(
                litchi_core::Position::new(1),
                litchi_rtf::edit::ParagraphLayoutPatch::new().with_space_before(120),
            ),
            litchi_rtf::edit::ParagraphLayoutUpdate::new(
                litchi_core::Position::new(2),
                litchi_rtf::edit::ParagraphLayoutPatch::new().with_keep_with_next(true),
            ),
        ])
        .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            edit.commit().unwrap().package().to_bytes().unwrap()
        );
        for (path, key) in [
            ("/body/p[1]", "align"),
            ("/body/p[2]", "spaceBefore"),
            ("/body/p[3]", "keepNext"),
            ("/body/p[4]", "align"),
        ] {
            assert_eq!(
                property(commit.file(), path, key),
                property(&oracle, path, key)
            );
        }
        assert_eq!(property(commit.file(), "/body/p[1]", "align"), "right");
        assert_eq!(property(commit.file(), "/body/p[2]", "spaceBefore"), "6pt");
        assert_eq!(property(commit.file(), "/body/p[3]", "keepNext"), "true");
        assert_eq!(property(commit.file(), "/body/p[4]", "align"), "center");
        assert_eq!(source.to_bytes().unwrap(), before);
    });
}

#[test]
fn paragraph_format_then_text_preserves_native_failure_boundary() {
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let error = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"align" => "right"})?;
            edit.set("/body/p[2]", props! {"text" => "text after format"})
        })
        .expect_err("native paragraph format/text closure");
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn ordered_direct_run_text_and_format_sets_share_one_publication() {
    each_owner(RICH_INPUT, |source| {
        let before = source.to_bytes().expect("source bytes");
        let oracle = source
            .set(
                "/body/p[1]/r[1]",
                props! {"text" => "First changed", "bold" => false},
            )
            .unwrap()
            .set(
                "/body/p[2]/r[1]",
                props! {"text" => "Second changed", "italic" => true},
            )
            .unwrap()
            .set(
                "/body/p[3]/r[1]",
                props! {"text" => "Third changed", "underline" => "single"},
            )
            .unwrap()
            .set(
                "/body/p[4]/r[1]",
                props! {"text" => "Fourth changed", "bold" => true},
            )
            .unwrap();

        let commit = source
            .edit(|edit| {
                edit.set(
                    "/body/p[1]/r[1]",
                    props! {"text" => "First changed", "bold" => false},
                )?;
                edit.set(
                    "/body/p[2]/r[1]",
                    props! {"text" => "Second changed", "italic" => true},
                )?;
                edit.set(
                    "/body/p[3]/r[1]",
                    props! {"text" => "Third changed", "underline" => "single"},
                )?;
                edit.set(
                    "/body/p[4]/r[1]",
                    props! {"text" => "Fourth changed", "bold" => true},
                )
            })
            .expect("ordered run edit");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(text(commit.file(), "/body/p[1]/r[1]"), "First changed");
        assert_eq!(text(commit.file(), "/body/p[2]/r[1]"), "Second changed");
        assert_eq!(text(commit.file(), "/body/p[3]/r[1]"), "Third changed");
        assert_eq!(text(commit.file(), "/body/p[4]/r[1]"), "Fourth changed");
        assert_eq!(property(commit.file(), "/body/p[2]/r[1]", "italic"), "true");
        assert_eq!(
            property(commit.file(), "/body/p[3]/r[1]", "underline"),
            "single"
        );
        assert_eq!(property(commit.file(), "/body/p[4]/r[1]", "bold"), "true");
        assert_eq!(source.to_bytes().unwrap(), before);
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
    });
}

#[test]
fn repeated_and_descending_direct_targets_close_runs_and_keep_sequential_results() {
    each_owner(INPUT, |source| {
        let repeated = source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"text" => "first staged"})?;
                edit.set("/body/p[1]", props! {"text" => "first final"})?;
                edit.set("/body/p[2]", props! {"text" => "second staged"})
            })
            .expect("repeated paragraph target");
        let repeated_oracle = source
            .set("/body/p[1]", props! {"text" => "first staged"})
            .unwrap()
            .set("/body/p[1]", props! {"text" => "first final"})
            .unwrap()
            .set("/body/p[2]", props! {"text" => "second staged"})
            .unwrap();
        assert_eq!(repeated.diagnostics().operations(), 3);
        assert_eq!(repeated.diagnostics().commits(), 2);
        assert_eq!(
            repeated.file().to_bytes().unwrap(),
            repeated_oracle.to_bytes().unwrap()
        );
        assert_eq!(text(repeated.file(), "/body/p[1]"), "first final");

        let descending = source
            .edit(|edit| {
                edit.set("/body/p[3]", props! {"text" => "third staged"})?;
                edit.set("/body/p[1]", props! {"text" => "first descending"})?;
                edit.set("/body/p[2]", props! {"text" => "middle staged"})
            })
            .expect("descending paragraph targets");
        let descending_oracle = source
            .set("/body/p[3]", props! {"text" => "third staged"})
            .unwrap()
            .set("/body/p[1]", props! {"text" => "first descending"})
            .unwrap()
            .set("/body/p[2]", props! {"text" => "middle staged"})
            .unwrap();
        assert_eq!(descending.diagnostics().operations(), 3);
        assert_eq!(descending.diagnostics().commits(), 2);
        assert_eq!(
            descending.file().to_bytes().unwrap(),
            descending_oracle.to_bytes().unwrap()
        );
        assert_eq!(text(descending.file(), "/body/p[1]"), "first descending");
        assert_eq!(text(descending.file(), "/body/p[2]"), "middle staged");
        assert_eq!(text(descending.file(), "/body/p[3]"), "third staged");
    });
}

#[test]
fn paragraph_and_run_families_are_incompatible_barriers() {
    each_owner(INPUT, |source| {
        let before = source.to_bytes().unwrap();
        let commit = source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"align" => "right"})?;
                edit.set("/body/p[2]/r[1]", props! {"bold" => true})?;
                edit.set("/body/p[3]", props! {"spaceAfter" => "3pt"})
            })
            .expect("mixed paragraph and run families");
        let oracle = source
            .set("/body/p[1]", props! {"align" => "right"})
            .unwrap()
            .set("/body/p[2]/r[1]", props! {"bold" => true})
            .unwrap()
            .set("/body/p[3]", props! {"spaceAfter" => "3pt"})
            .unwrap();
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 3);
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(property(commit.file(), "/body/p[2]/r[1]", "bold"), "true");
        assert_eq!(source.to_bytes().unwrap(), before);
    });
}

#[test]
fn reads_queries_raw_paths_and_structural_commands_close_direct_runs() {
    let source = eager(INPUT);

    let read = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "before-read"})?;
            assert_eq!(
                edit.get("/body/p[1]")?.view().text()?.as_deref(),
                Some("before-read")
            );
            edit.set("/body/p[2]", props! {"text" => "after-read"})
        })
        .expect("read barrier");
    assert_eq!(read.diagnostics().commits(), 2);
    assert_eq!(text(read.file(), "/body/p[2]"), "after-read");

    let query_result = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "before"})?;
            edit.set(query("p[text=before]"), props! {"text" => "query"})?;
            edit.set("/body/p[2]", props! {"text" => "after"})
        })
        .expect("query barrier");
    assert_eq!(query_result.diagnostics().commits(), 3);
    assert_eq!(text(query_result.file(), "/body/p[1]"), "query");
    assert_eq!(text(query_result.file(), "/body/p[2]"), "after");

    let raw = litchi_ops::path::raw::path("/body/p[2]").expect("raw RTF path");
    let raw_result = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"align" => "right"})?;
            edit.set(raw, props! {"align" => "center"})?;
            edit.set("/body/p[3]", props! {"align" => "justify"})
        })
        .expect("raw path barrier");
    assert_eq!(raw_result.diagnostics().commits(), 3);
    assert_eq!(property(raw_result.file(), "/body/p[2]", "align"), "center");

    let structural = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "before-remove"})?;
            edit.remove("/body/p[4]")?;
            edit.set("/body/p[2]", props! {"text" => "after remove"})
        })
        .expect("structural barrier");
    assert_eq!(structural.diagnostics().commits(), 3);
    assert!(structural.file().get("/body/p[4]").is_err());
    assert_eq!(text(structural.file(), "/body/p[2]"), "after remove");

    let moved = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "before-move"})?;
            edit.move_to("/body/p[4]", Placement::Start)?;
            edit.set("/body/p[2]", props! {"text" => "after move"})
        })
        .expect("move barrier");
    assert_eq!(moved.diagnostics().commits(), 3);
    assert_eq!(text(moved.file(), "/body/p[2]"), "after move");
}

#[test]
fn atomic_batch_coalesces_ordered_sets_while_best_effort_keeps_item_boundaries() {
    let commands = || {
        [
            batch_set("/body/p[1]", Props::new().with("text", "first batch")),
            batch_set("/body/p[2]", Props::new().with("text", "second batch")),
            batch_set("/body/p[3]", Props::new().with("text", "third batch")),
        ]
    };
    each_owner(INPUT, |source| {
        let before = source.to_bytes().unwrap();
        let atomic = source
            .batch(commands(), BatchMode::Atomic)
            .expect("atomic ordered batch");
        assert_eq!(atomic.value().len(), 3);
        assert!(atomic.value().iter().all(|result| {
            matches!(result, litchi_ops::ItemResult::Applied(applied) if applied.changed)
        }));
        assert_eq!(atomic.diagnostics().operations(), 3);
        assert_eq!(atomic.diagnostics().commits(), 1);
        assert_eq!(text(atomic.file(), "/body/p[1]"), "first batch");
        assert_eq!(text(atomic.file(), "/body/p[2]"), "second batch");
        assert_eq!(text(atomic.file(), "/body/p[3]"), "third batch");
        assert_eq!(source.to_bytes().unwrap(), before);

        let best = source
            .batch(commands(), BatchMode::BestEffort)
            .expect("best-effort ordered batch");
        assert_eq!(best.value().len(), 3);
        assert_eq!(best.diagnostics().operations(), 3);
        assert_eq!(best.diagnostics().commits(), 3);
        assert_eq!(text(best.file(), "/body/p[1]"), "first batch");
        assert_eq!(text(best.file(), "/body/p[2]"), "second batch");
        assert_eq!(text(best.file(), "/body/p[3]"), "third batch");
        assert_eq!(source.to_bytes().unwrap(), before);
    });
}

#[test]
fn noops_keep_exact_bytes_and_batch_result_order() {
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let mut same = BatchItem::new("set");
    same.path = Some("/body/p[1]".into());
    same.props = Some(Props::new().with("align", "left"));
    let changed = batch_set("/body/p[2]", Props::new().with("text", "changed"));
    let commit = source
        .batch([same, changed], BatchMode::Atomic)
        .expect("noop plus changed batch");
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
    assert_eq!(text(commit.file(), "/body/p[2]"), "changed");
    assert_eq!(source.to_bytes().unwrap(), before);

    let all_noop = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"align" => "left"})?;
            edit.set("/body/p[2]", props! {"text" => "Second"})
        })
        .expect("all no-op direct sets");
    assert!(!all_noop.diagnostics().changed());
    assert_eq!(all_noop.diagnostics().operations(), 2);
    assert_eq!(all_noop.diagnostics().commits(), 1);
    assert_eq!(all_noop.file().to_bytes().unwrap(), before);
}

#[test]
fn atomic_and_best_effort_batches_preserve_failure_indices_and_rollback() {
    let commands = || {
        [
            batch_set("/body/p[1]", Props::new().with("text", "atomic first")),
            batch_set("/body/p[2]", Props::new().with("unknown-property", true)),
            batch_set("/body/p[3]", Props::new().with("text", "atomic third")),
        ]
    };
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let error = source
        .batch(commands(), BatchMode::Atomic)
        .expect_err("atomic invalid item");
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(source.to_bytes().unwrap(), before);

    let commit = source
        .batch(commands(), BatchMode::BestEffort)
        .expect("best-effort invalid item");
    assert_eq!(commit.value().len(), 3);
    assert!(matches!(
        &commit.value()[0],
        litchi_ops::ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &commit.value()[1],
        litchi_ops::ItemResult::Failed(error)
            if error.item() == Some(1) && error.code() == Code::UnsupportedCapability
    ));
    assert!(matches!(
        &commit.value()[2],
        litchi_ops::ItemResult::Applied(applied) if applied.changed
    ));
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(text(commit.file(), "/body/p[1]"), "atomic first");
    assert_eq!(text(commit.file(), "/body/p[3]"), "atomic third");
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn pending_native_failure_precedes_a_later_invalid_property() {
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let result = source.edit(|edit| {
        // Run replacement cannot contain a line break.  The direct planner
        // admits this item and the native batch reports it when the pending
        // run is flushed for the following command.
        edit.set("/body/p[1]/r[1]", props! {"text" => "first\nsecond"})?;
        edit.set("/body/p[2]", props! {"unknown-property" => true})
    });
    let error = result.expect_err("pending native failure");
    assert_eq!(error.item(), Some(0));
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn pending_native_failure_precedes_operation_limit_excess_in_edit_and_batch() {
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("operation limit")
        .build()
        .expect("limits");
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                // The first item is retained by the direct planner until the
                // second item forces the pending native run to publish.
                edit.set("/body/p[1]/r[1]", props! {"text" => "first\nsecond"})?;
                edit.set("/body/p[2]/r[1]", props! {"text" => "limit item"})
            },
        )
        .expect_err("native failure must precede operation limit");
    assert_eq!(error.item(), Some(0));
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(source.to_bytes().unwrap(), before);

    let error = source
        .batch_with_options(
            [
                batch_set(
                    "/body/p[1]/r[1]",
                    Props::new().with("text", "first\nsecond"),
                ),
                batch_set("/body/p[2]/r[1]", Props::new().with("text", "limit item")),
            ],
            BatchMode::Atomic,
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .expect_err("native failure must precede batch operation limit");
    assert_eq!(error.item(), Some(0));
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn empty_run_text_removal_retargets_the_next_run_in_the_same_paragraph() {
    let input = br"{\rtf1\ansi Plain{\i Italic}\b Bold\par Tail}";
    let source = open_bytes(input).expect("open run-shift fixture");
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            edit.set("/body/p[1]/r[1]", props! {"text" => ""})?;
            // The first run is now gone, so this path refers to the former
            // italic run.  Resolving against the original snapshot would
            // replace Plain's position and lose the native format boundary.
            edit.set("/body/p[1]/r[1]", props! {"text" => "retargeted"})
        })
        .expect("empty run followed by same-paragraph set");
    let oracle = source
        .set("/body/p[1]/r[1]", props! {"text" => ""})
        .unwrap()
        .set("/body/p[1]/r[1]", props! {"text" => "retargeted"})
        .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        oracle.to_bytes().unwrap()
    );
    assert_eq!(text(commit.file(), "/body/p[1]"), "retargetedBold");
    assert_eq!(property(commit.file(), "/body/p[1]/r[1]", "italic"), "true");
    assert_eq!(property(commit.file(), "/body/p[1]/r[2]", "bold"), "true");
    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(source.to_bytes().unwrap(), before);
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
}

#[test]
fn variable_unicode_run_replacements_preserve_ordered_boundaries() {
    let input = br"{\rtf1\ansi\uc1 A\u233?B{\i C\u38634?D}\tab end}";
    let source = open_bytes(input).expect("open unicode run fixture");
    let before = source.to_bytes().unwrap();
    let replacements = [
        ("/body/p[1]/r[1]", "A long 🦀"),
        ("/body/p[1]/r[2]", "é plus 🦀"),
        ("/body/p[1]/r[3]", "B"),
        ("/body/p[1]/r[4]", "C variable"),
        ("/body/p[1]/r[5]", "雪 and 🦀"),
        ("/body/p[1]/r[6]", "D tail"),
        ("/body/p[1]/r[7]", "\tend expanded"),
    ];
    let mut oracle = source.clone();
    for (path, value) in replacements {
        oracle = oracle
            .set(path, props! {"text" => value})
            .expect("sequential unicode replacement");
    }
    let commit = source
        .edit(|edit| {
            for (path, value) in replacements {
                edit.set(path, props! {"text" => value})?;
            }
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("ordered unicode replacement");
    assert_eq!(commit.diagnostics().operations(), replacements.len());
    // Variable-width Unicode replacements can close a native run when the
    // resulting semantic text fragments need new boundaries.  The important
    // contract here is exact sequential equivalence; any resulting native
    // barriers remain observable in the commit count.
    assert!(commit.diagnostics().commits() >= 1);
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        oracle.to_bytes().unwrap()
    );
    assert_eq!(
        text(commit.file(), "/body/p[1]"),
        text(&oracle, "/body/p[1]")
    );
    assert_eq!(source.to_bytes().unwrap(), before);
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
}

#[test]
fn direct_set_limits_roll_over_or_refuse_without_partial_publication() {
    let source = eager(INPUT);
    let before = source.to_bytes().unwrap();
    let operation_limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("operation limit")
        .build()
        .expect("limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(operation_limits),
            |edit| {
                edit.set("/body/p[1]", props! {"text" => "bounded one"})?;
                edit.set("/body/p[2]", props! {"text" => "bounded two"})
            },
        )
        .expect_err("aggregate operation limit");
    assert!(matches!(
        error.leaf(),
        Source::Transaction(TransactionError::Limit {
            kind: LimitKind::Operations,
            ..
        })
    ));
    assert_eq!(source.to_bytes().unwrap(), before);

    let expanded_limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .expect("target limit")
        .build()
        .expect("limits");
    let rolled = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(expanded_limits),
            |edit| {
                edit.set("/body/p[1]", props! {"text" => "expanded one"})?;
                edit.set("/body/p[2]", props! {"text" => "expanded two"})
            },
        )
        .expect("expanded target rollover");
    assert_eq!(rolled.diagnostics().operations(), 2);
    assert_eq!(rolled.diagnostics().commits(), 2);
    assert_eq!(text(rolled.file(), "/body/p[1]"), "expanded one");
    assert_eq!(text(rolled.file(), "/body/p[2]"), "expanded two");
    assert_eq!(source.to_bytes().unwrap(), before);

    for (limits, kind) in [
        (
            SemanticLimits::builder()
                .max_staged_bytes(1)
                .expect("staging limit")
                .build()
                .expect("limits"),
            LimitKind::StagedBytes,
        ),
        (
            SemanticLimits::builder()
                .max_snapshot_bytes(1)
                .expect("snapshot limit")
                .build()
                .expect("limits"),
            LimitKind::SnapshotBytes,
        ),
    ] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.set("/body/p[1]", props! {"text" => "bounded"}),
            )
            .expect_err("individual RTF limit");
        assert!(
            matches!(
                error.leaf(),
                Source::Transaction(TransactionError::Limit { kind: actual, .. })
                    if *actual == kind
            ),
            "{error:?}"
        );
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn filtered_paths_keep_ordinary_resolution_semantics() {
    let source = eager(INPUT);
    for path in [
        "/body/p[1][text=missing]",
        "/body[1][text=missing]/p[1]",
        "/body/p[1]/r[1][text=missing]",
    ] {
        let ordinary = source.set(path, props! {"text" => "unexpected"});
        let transaction = source.edit(|edit| edit.set(path, props! {"text" => "unexpected"}));
        match (ordinary, transaction) {
            (Err(left), Err(right)) => assert_eq!(left.code(), right.code(), "{path}"),
            (Ok(left), Ok(right)) => assert_eq!(
                left.to_bytes().unwrap(),
                right.file().to_bytes().unwrap(),
                "{path}"
            ),
            (left, right) => {
                panic!("different filtered-path admission for {path}: {left:?}, {right:?}")
            },
        }
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn staging_windows_roll_over_and_small_limits_keep_scalar_noop_admission() {
    let source = eager(INPUT);
    let tiny = SemanticLimits::builder()
        .max_staged_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let noop = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(tiny),
            |edit| edit.set("/body/p[1]/r[1]", Props::new()),
        )
        .unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.file().to_bytes().unwrap(), INPUT);

    let bounded = SemanticLimits::builder()
        .max_staged_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    let payload = "x".repeat(700);
    let commit = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(bounded),
            |edit| {
                edit.set("/body/p[1]", props! {"text" => payload.as_str()})?;
                edit.set("/body/p[2]", props! {"text" => payload.as_str()})
            },
        )
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(text(commit.file(), "/body/p[1]"), payload);
    assert_eq!(text(commit.file(), "/body/p[2]"), payload);
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn noop_alignment_before_structural_text_is_a_native_boundary() {
    let source = eager(INPUT);
    let commit = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"align" => "left"})?;
            edit.set("/body/p[2]", props! {"text" => "x\ny"})
        })
        .unwrap();
    let oracle = source
        .set("/body/p[1]", props! {"align" => "left"})
        .unwrap()
        .set("/body/p[2]", props! {"text" => "x\ny"})
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        oracle.to_bytes().unwrap()
    );
    assert_eq!(text(commit.file(), "/body/p[2]"), "x");
    assert_eq!(text(commit.file(), "/body/p[3]"), "y");
}
