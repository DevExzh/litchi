//! PPTX direct shape-set coalescing and transaction barriers.
#![cfg(feature = "pptx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::node::Value;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, NewNode, OpenOptions, Props, SourceMode, props,
};

const LARGE_SHAPE_COUNT: usize = 1_000;

fn fixture(shape_count: usize) -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().expect("create presentation builder");
    let slide = builder
        .presentation()
        .expect("create presentation")
        .add_slide()
        .expect("create slide");
    for index in 0..shape_count {
        slide.add_text_box(
            &format!("original-{index}"),
            i64::try_from(index).expect("shape index") * 100,
            0,
            914_400,
            914_400,
        );
    }
    builder
        .build()
        .expect("build presentation")
        .to_bytes()
        .expect("serialize presentation")
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
    .expect("open presentation")
}

fn shape_path(index: usize) -> String {
    format!("/slide[1]/shape[{index}]")
}

fn shape_item(index: usize, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(shape_path(index).into());
    item.props = Some(properties);
    item
}

fn full_bounds(x: &'static str) -> Props<'static> {
    Props::new()
        .with("x", x)
        .with("y", "1in")
        .with("width", "1in")
        .with("height", "1in")
}

fn meta_item() -> BatchItem<'static> {
    let mut item = BatchItem::new("meta");
    item.dump_version = Some(2);
    item
}

fn text(file: &File, index: usize) -> String {
    let path = shape_path(index);
    file.get(&path)
        .expect("resolve shape")
        .view()
        .text()
        .expect("read shape text")
        .expect("shape has a text body")
        .into_owned()
}

fn length(file: &File, index: usize, key: &str) -> i64 {
    let path = shape_path(index);
    let node = file.get(&path).expect("resolve shape");
    let property = node
        .view()
        .format()
        .expect("read shape properties")
        .find(|property| property.key() == key)
        .expect("shape property");
    let Value::Length(value) = property.value() else {
        panic!("expected a length property");
    };
    value.emu()
}

fn rewrite_final_transform(bytes: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open presentation zip");
    let mut members = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).expect("read zip member");
        let name = member.name().to_owned();
        let mut payload = Vec::new();
        member.read_to_end(&mut payload).expect("read zip payload");
        if name == "ppt/slides/slide1.xml" {
            let mut xml = String::from_utf8(payload).expect("slide XML");
            let start = xml
                .rfind("<a:xfrm>")
                .expect("last shape has a local transform");
            xml.replace_range(
                start..start + "<a:xfrm>".len(),
                "<a:xfrm rot=\"private-invalid\">",
            );
            payload = xml.into_bytes();
        }
        members.push((name, payload));
    }
    drop(archive);

    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, payload) in members {
        writer.start_file(name, options).expect("write zip member");
        writer.write_all(&payload).expect("write zip payload");
    }
    writer
        .finish()
        .expect("finish presentation zip")
        .into_inner()
}

#[test]
fn atomic_batch_coalesces_one_thousand_distinct_shapes_with_per_item_changes() {
    let bytes = fixture(LARGE_SHAPE_COUNT);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize source");
        let items = (1..=LARGE_SHAPE_COUNT)
            .map(|index| {
                let value = if index == 1 {
                    "original-0".to_owned()
                } else {
                    format!("batch-{index}")
                };
                shape_item(index, Props::new().with("text", value))
            })
            .collect::<Vec<_>>();

        let commit = source
            .batch(items, BatchMode::Atomic)
            .expect("coalesced thousand-item batch");
        assert_eq!(commit.value().len(), LARGE_SHAPE_COUNT);
        assert_eq!(commit.diagnostics().operations(), LARGE_SHAPE_COUNT);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert!(commit.diagnostics().changed());
        assert!(matches!(
            &commit.value()[0],
            ItemResult::Applied(applied) if !applied.changed
        ));
        assert!(commit.value()[1..].iter().all(|result| matches!(
            result,
            ItemResult::Applied(applied) if applied.changed
        )));
        assert_eq!(text(commit.file(), 1), "original-0");
        assert_eq!(text(commit.file(), 256), "batch-256");
        assert_eq!(text(commit.file(), LARGE_SHAPE_COUNT), "batch-1000");
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
        assert_eq!(
            commit
                .patch()
                .apply(&source)
                .expect("replay batch")
                .to_bytes()
                .expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse batch")
                .to_bytes()
                .expect("serialize inverse"),
            before
        );
    }
}

#[test]
fn mixed_shape_batch_coalesces_geometry_and_text_and_reports_noops() {
    let bytes = fixture(4);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize source");
        let commit = source
            .batch(
                [
                    shape_item(1, Props::new().with("text", "original-0").with("x", "0in")),
                    shape_item(
                        2,
                        Props::new()
                            .with("text", "changed")
                            .with("x", "1in")
                            .with("width", "2in"),
                    ),
                    shape_item(3, Props::new().with("height", "1in")),
                    shape_item(
                        4,
                        Props::new()
                            .with("text", "last")
                            .with("y", "2in")
                            .with("height", "3in"),
                    ),
                ],
                BatchMode::Atomic,
            )
            .expect("mixed coalesced batch");
        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        let changed = commit
            .value()
            .iter()
            .map(|result| match result {
                ItemResult::Applied(applied) => applied.changed,
                other => panic!("unexpected batch result: {other:?}"),
            })
            .collect::<Vec<_>>();
        assert_eq!(changed, [false, true, false, true]);
        assert_eq!(text(commit.file(), 2), "changed");
        assert_eq!(length(commit.file(), 2, "x"), 914_400);
        assert_eq!(length(commit.file(), 2, "width"), 1_828_800);
        assert_eq!(length(commit.file(), 4, "y"), 1_828_800);
        assert_eq!(length(commit.file(), 4, "height"), 2_743_200);
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
    }
}

#[test]
fn edit_coalesces_mixed_shape_sets_and_replays_one_native_step() {
    let bytes = fixture(3);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize source");
        let commit = source
            .edit(|edit| {
                edit.set(
                    &shape_path(1),
                    props! { "text" => "edited", "x" => "1in", "width" => "2in" },
                )?;
                edit.set(
                    &shape_path(2),
                    props! { "text" => "second", "y" => "2in", "height" => "3in" },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("coalesced closure edit");
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert!(commit.diagnostics().changed());
        assert_eq!(text(commit.file(), 1), "edited");
        assert_eq!(length(commit.file(), 1, "x"), 914_400);
        assert_eq!(length(commit.file(), 2, "height"), 2_743_200);
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
        assert_eq!(
            commit
                .patch()
                .apply(&source)
                .expect("replay closure edit")
                .to_bytes()
                .expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse closure edit")
                .to_bytes()
                .expect("serialize inverse"),
            before
        );
    }
}

#[test]
fn repeated_targets_reads_metadata_and_structure_close_direct_runs() {
    let bytes = fixture(3);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let repeated = source
            .edit(|edit| {
                edit.set(&shape_path(1), props! { "text" => "first" })?;
                edit.set(&shape_path(1), props! { "text" => "second" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeated target edit");
        assert_eq!(repeated.diagnostics().commits(), 2);
        assert_eq!(text(repeated.file(), 1), "second");

        let read_barrier = source
            .edit(|edit| {
                edit.set(&shape_path(1), props! { "text" => "before read" })?;
                let _ = edit.get(&shape_path(2))?.view().text()?;
                edit.set(&shape_path(3), props! { "text" => "after read" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("read barrier edit");
        assert_eq!(read_barrier.diagnostics().commits(), 2);
        assert_eq!(text(read_barrier.file(), 3), "after read");

        let structural_barrier = source
            .edit(|edit| {
                edit.set(&shape_path(1), props! { "text" => "before add" })?;
                let added = edit.add("/", NewNode::new("slide").prop("layout", "Blank"))?;
                assert_eq!(added.to_string(), "/slide[2]");
                edit.set(&shape_path(2), props! { "text" => "after add" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("structural barrier edit");
        assert_eq!(structural_barrier.diagnostics().commits(), 3);
        assert_eq!(text(structural_barrier.file(), 1), "before add");
        assert_eq!(text(structural_barrier.file(), 2), "after add");

        let metadata = source
            .batch(
                [
                    shape_item(1, Props::new().with("text", "before metadata")),
                    meta_item(),
                    shape_item(2, Props::new().with("text", "after metadata")),
                ],
                BatchMode::Atomic,
            )
            .expect("metadata barrier batch");
        assert_eq!(metadata.diagnostics().operations(), 2);
        assert_eq!(metadata.diagnostics().commits(), 2);
        assert!(matches!(
            &metadata.value()[1],
            ItemResult::Applied(applied) if !applied.changed
        ));

        let query = source
            .batch(
                [
                    shape_item(1, Props::new().with("text", "before query")),
                    shape_item(2, Props::new().with("text", "query")),
                    shape_item(3, Props::new().with("text", "after query")),
                ],
                BatchMode::Atomic,
            )
            .expect("direct path batch");
        assert_eq!(query.diagnostics().commits(), 1);
        assert_eq!(text(query.file(), 3), "after query");
    }
}

#[test]
fn selector_and_metadata_are_explicit_batch_barriers() {
    let bytes = fixture(3);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let mut selector = BatchItem::new("set");
        selector.path = Some("/slide[1]/*".into());
        selector.props = Some(Props::new().with("text", "selected"));
        let commit = source
            .batch(
                [
                    shape_item(1, Props::new().with("text", "before selector")),
                    selector,
                    shape_item(3, Props::new().with("text", "after selector")),
                ],
                BatchMode::Atomic,
            )
            .expect("selector barrier batch");
        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(commit.diagnostics().commits(), 3);
        assert_eq!(text(commit.file(), 1), "selected");
        assert_eq!(text(commit.file(), 2), "selected");
        assert_eq!(text(commit.file(), 3), "after selector");
    }
}

#[test]
fn malformed_transform_failure_keeps_the_late_wire_index_and_rolls_back() {
    let bytes = rewrite_final_transform(&fixture(3));
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize malformed source");
        let error = source
            .batch(
                [
                    shape_item(1, Props::new().with("x", "1in")),
                    shape_item(3, Props::new().with("x", "2in")),
                    shape_item(2, Props::new().with("x", "3in")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("malformed native transform must fail atomically");
        assert_eq!(error.item(), Some(1));
        assert!(matches!(
            error.code(),
            Code::CorruptFile | Code::MissingProperty | Code::InvalidValue
        ));
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
    }
}

#[test]
fn an_earlier_deferred_native_failure_precedes_later_set_admission_errors() {
    let bytes = rewrite_final_transform(&fixture(3));
    for eager in [false, true] {
        for later in [
            shape_item(1, Props::new().with("unknownPptxProperty", true)),
            {
                let mut item = BatchItem::new("set");
                item.path = Some("/slide[1]/shape[".into());
                item.props = Some(Props::new().with("text", "bad path"));
                item
            },
            shape_item(1, Props::new()),
        ] {
            let source = open(&bytes, eager);
            let before = source.to_bytes().expect("serialize malformed source");
            let error = source
                .batch(
                    [
                        shape_item(1, full_bounds("1in")),
                        shape_item(3, full_bounds("2in")),
                        later,
                    ],
                    BatchMode::Atomic,
                )
                .expect_err("the pending native failure must win");
            assert_eq!(error.item(), Some(1));
            assert!(matches!(
                error.code(),
                Code::CorruptFile | Code::MissingProperty | Code::InvalidValue
            ));
            assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
        }
    }
}

#[test]
fn best_effort_isolates_a_late_malformed_transform_from_valid_shape_sets() {
    let bytes = rewrite_final_transform(&fixture(3));
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize malformed source");
        let commit = source
            .batch(
                [
                    shape_item(1, full_bounds("1in")),
                    shape_item(3, full_bounds("2in")),
                    shape_item(2, full_bounds("3in")),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batch should isolate malformed item");
        assert!(matches!(
            &commit.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &commit.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1)
                    && matches!(
                        error.code(),
                        Code::CorruptFile | Code::MissingProperty | Code::InvalidValue
                    )
        ));
        assert!(matches!(
            &commit.value()[2],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(length(commit.file(), 1, "x"), 914_400);
        assert_eq!(length(commit.file(), 2, "x"), 2_743_200);
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
        assert_eq!(
            commit
                .patch()
                .apply(&source)
                .expect("replay best-effort result")
                .to_bytes()
                .expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse best-effort result")
                .to_bytes()
                .expect("serialize inverse"),
            before
        );
    }
}

#[test]
fn coalesced_text_batches_split_the_native_text_cap_but_keep_one_commit() {
    let bytes = fixture(2);
    let source = open(&bytes, false);
    let before = source.to_bytes().expect("serialize source");
    let text_bytes = litchi_pptx::opened::Limits::default()
        .max_text_bytes()
        .checked_div(2)
        .expect("nonzero native text limit")
        .saturating_add(1);
    let first = "a".repeat(text_bytes);
    let second = "b".repeat(text_bytes);

    let commit = source
        .batch(
            [
                shape_item(1, Props::new().with("text", first)),
                shape_item(2, Props::new().with("text", second)),
            ],
            BatchMode::Atomic,
        )
        .expect("coalesced text batches split at the native aggregate cap");
    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert!(commit.diagnostics().changed());
    assert_eq!(text(commit.file(), 1).len(), text_bytes);
    assert_eq!(text(commit.file(), 2).len(), text_bytes);
    assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
}

#[test]
fn operation_and_staged_byte_limits_reject_before_publication() {
    let bytes = fixture(3);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().expect("serialize source");
        let operation_limit = SemanticLimits::builder()
            .max_operations(2)
            .expect("valid operation limit")
            .build()
            .expect("build operation limit");
        let operation_error = source
            .batch_with_options(
                [
                    shape_item(1, Props::new().with("text", "one")),
                    shape_item(2, Props::new().with("text", "two")),
                    shape_item(3, Props::new().with("text", "three")),
                ],
                BatchMode::Atomic,
                litchi_ops::EditOptions::new().semantic_limits(operation_limit),
            )
            .expect_err("operation limit");
        assert_eq!(operation_error.code(), Code::LimitExceeded);
        assert_eq!(operation_error.item(), Some(2));
        assert_eq!(source.to_bytes().expect("source unchanged"), before);

        let staged_limit = SemanticLimits::builder()
            .max_staged_bytes(1)
            .expect("valid staged-byte limit")
            .build()
            .expect("build staged-byte limit");
        let staged_error = source
            .batch_with_options(
                [shape_item(1, Props::new().with("text", "too large"))],
                BatchMode::Atomic,
                litchi_ops::EditOptions::new().semantic_limits(staged_limit),
            )
            .expect_err("staged-byte limit");
        assert_eq!(staged_error.code(), Code::LimitExceeded);
        assert_eq!(staged_error.item(), Some(0));
        assert_eq!(source.to_bytes().expect("source unchanged"), before);
    }
}

#[test]
fn alternating_text_and_geometry_commands_share_a_shape_run() {
    let bytes = fixture(LARGE_SHAPE_COUNT);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let commit = source
            .batch(
                (0..LARGE_SHAPE_COUNT).map(|index| {
                    shape_item(
                        index + 1,
                        if index % 2 == 0 {
                            Props::new().with("text", "alternating text")
                        } else {
                            Props::new().with("x", "1in")
                        },
                    )
                }),
                BatchMode::Atomic,
            )
            .expect("alternating same-slide shape commands");
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(commit.value().len(), LARGE_SHAPE_COUNT);
        assert!(commit.value().iter().all(|item| matches!(item,
            ItemResult::Applied(applied) if applied.changed)));
        let output = commit.file().to_bytes().expect("serialize command result");
        let package = litchi_pptx::Package::from_bytes(&output).expect("open result");
        let presentation = package.presentation().expect("presentation");
        let slide = presentation.slide(0).expect("slide access").expect("slide");
        let scene = slide.shapes().expect("shape scene");
        for index in 0..LARGE_SHAPE_COUNT {
            let shape = scene.at(index).expect("shape");
            if index % 2 == 0 {
                assert_eq!(shape.text(), Some("alternating text"));
                assert_eq!(
                    shape.common().bounds().expect("bounds").x(),
                    index as i64 * 100
                );
            } else {
                let original = format!("original-{index}");
                assert_eq!(shape.text(), Some(original.as_str()));
                assert_eq!(shape.common().bounds().expect("bounds").x(), 914_400);
            }
        }
        assert_eq!(source.to_bytes().expect("immutable source"), bytes);
    }
}
