//! Unified immutable writes for legacy PowerPoint presentations.
//!
//! The fixture is authored by the native PPT writer and is then opened through
//! both the byte/eager and filesystem/positional owner routes.  Assertions use
//! only the public [`litchi_ops::File`] surface; the native package is used
//! separately only to verify that a published edit can be reopened.

#![cfg(feature = "ppt")]

use std::io::Cursor;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::node::Value;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Props, SourceMode, props};
use litchi_ppt::writer::Builder;

const FIRST: &str = "/slide[1]/shape[1]";
const SECOND: &str = "/slide[2]/shape[1]";

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new();
    let first = builder.add_slide().expect("first slide");
    builder
        .add_textbox(first, 10, 10, 240, 40, "first")
        .expect("first textbox");
    builder
        .add_textbox(first, 10, 70, 240, 40, "other")
        .expect("second textbox");

    let second = builder.add_slide().expect("second slide");
    builder
        .add_textbox(second, 10, 10, 240, 40, "second")
        .expect("third textbox");
    builder
        .add_textbox(second, 10, 70, 240, 40, "untouched")
        .expect("fourth textbox");

    builder
        .build()
        .expect("build PPT fixture")
        .to_bytes()
        .expect("serialize PPT fixture")
}

fn signed_fixture() -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(fixture())).expect("open PPT fixture");
    let document = ole
        .open_stream(&["PowerPoint Document"])
        .expect("PowerPoint Document stream");
    let current_user = ole
        .open_stream(&["Current User"])
        .expect("Current User stream");

    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .expect("copy document stream");
    writer
        .create_stream(&["Current User"], &current_user)
        .expect("copy Current User stream");
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .expect("add signature stream");
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).expect("write signed PPT");
    bytes.into_inner()
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn source(bytes: &[u8]) -> (tempfile::TempDir, File) {
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("source.ppt");
    std::fs::write(&path, bytes).expect("write PPT source");
    let file = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
        .expect("open positional PPT");
    (directory, file)
}

fn each_owner(bytes: &[u8], mut test: impl FnMut(File)) {
    test(eager(bytes));
    let (_directory, file) = source(bytes);
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
    let property = file
        .get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
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

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn hidden(file: &File, path: &str) -> bool {
    property(file, path, "hidden").as_deref() == Some("true")
}

fn scoped_shapes() -> Query<'static> {
    Query::parse("shape")
        .expect("shape query")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/slide[1]").expect("slide scope"),
        )))
}

fn batch_set(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

#[test]
fn shape_text_and_geometry_round_trip_without_mutating_the_source() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let updated = source
            .set(
                FIRST,
                props! {
                    "text" => "changed & <escaped>",
                    "x" => "1in",
                    "y" => "2in",
                    "width" => "3in",
                    "height" => "4in",
                },
            )
            .expect("shape write");

        assert_eq!(text(&updated, FIRST), "changed & <escaped>");
        assert_eq!(length(&updated, FIRST, "x"), 914_400);
        assert_eq!(length(&updated, FIRST, "y"), 1_828_800);
        assert_eq!(length(&updated, FIRST, "width"), 2_743_200);
        assert_eq!(length(&updated, FIRST, "height"), 3_657_600);
        assert_eq!(text(&updated, SECOND), "second");
        assert_eq!(text(&source, FIRST), "first");
        assert_eq!(source.to_bytes().expect("unchanged source"), before);

        let reopened = eager(&updated.to_bytes().expect("serialize update"));
        assert_eq!(text(&reopened, FIRST), "changed & <escaped>");
        assert_eq!(length(&reopened, FIRST, "height"), 3_657_600);

        let noop = source
            .set(
                FIRST,
                props! {
                    "text" => "first",
                    "x" => "10pt",
                    "y" => "10pt",
                    "width" => "240pt",
                    "height" => "40pt",
                },
            )
            .expect("equal properties are a no-op");
        assert_eq!(noop.to_bytes().expect("no-op bytes"), before);
    });
}

#[test]
fn scoped_queries_and_slide_visibility_write_only_selected_nodes() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let updated = source
            .set(scoped_shapes(), props! { "text" => "scoped" })
            .expect("scoped shape write");
        assert_eq!(text(&updated, "/slide[1]/shape[1]"), "scoped");
        assert_eq!(text(&updated, "/slide[1]/shape[2]"), "scoped");
        assert_eq!(text(&updated, SECOND), "second");

        let hidden_slide = updated
            .set("/slide[1]", props! { "hidden" => true })
            .expect("hide slide");
        assert!(hidden(&hidden_slide, "/slide[1]"));
        assert!(!hidden(&hidden_slide, "/slide[2]"));
        let visible_slide = hidden_slide
            .set("/slide[1]", props! { "hidden" => false })
            .expect("show slide");
        assert!(!hidden(&visible_slide, "/slide[1]"));
    });
}

#[test]
fn edit_history_batch_modes_and_rollback_are_atomic() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "text" => "staged" })?;
                edit.set(SECOND, props! { "text" => "second staged" })
            })
            .expect("transaction");
        assert!(commit.diagnostics.changed());
        assert_eq!(text(&commit.file, FIRST), "staged");
        assert_eq!(
            commit
                .patch
                .apply(&source)
                .expect("replay patch")
                .to_bytes()
                .expect("replayed bytes"),
            commit.file.to_bytes().expect("commit bytes")
        );
        assert_eq!(
            commit
                .patch
                .inverse()
                .apply(&commit.file)
                .expect("inverse patch")
                .to_bytes()
                .expect("inverse bytes"),
            before
        );

        let noop = source
            .edit(|edit| edit.set(FIRST, props! { "text" => "first" }))
            .expect("transactional no-op");
        assert!(noop.patch.is_noop());
        assert!(!noop.diagnostics.changed());

        let poisoned = source.edit(|edit| {
            edit.set(FIRST, props! { "text" => "discarded" })?;
            let _ = edit.set(FIRST, props! { "id" => 1 });
            Ok::<_, litchi_ops::Error>(())
        });
        assert!(poisoned.is_err());
        assert_eq!(source.to_bytes().expect("source after rollback"), before);

        let commands = || {
            [
                batch_set(FIRST, Props::new().with("text", "batch changed")),
                batch_set(FIRST, Props::new().with("id", 1)),
                batch_set(SECOND, Props::new().with("text", "batch later")),
            ]
        };
        let atomic = source
            .batch(commands(), BatchMode::Atomic)
            .expect_err("atomic batch failure");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(source.to_bytes().expect("atomic source"), before);

        let best_effort = source
            .batch(commands(), BatchMode::BestEffort)
            .expect("best effort batch");
        assert_eq!(text(&best_effort.file, FIRST), "batch changed");
        assert_eq!(text(&best_effort.file, SECOND), "batch later");
        assert!(matches!(
            &best_effort.results[1],
            litchi_ops::ItemResult::Failed(error) if error.item() == Some(1)
        ));
        assert_eq!(text(&source, FIRST), "first");
    });
}

#[test]
fn unsupported_properties_invalid_values_and_limits_leave_input_unchanged() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        assert_eq!(
            source
                .set(FIRST, props! { "id" => 1 })
                .expect_err("read-only id")
                .code(),
            Code::UnsupportedProperty
        );
        assert_eq!(
            source
                .set(FIRST, props! { "width" => 0 })
                .expect_err("zero width")
                .code(),
            Code::InvalidValue
        );
        assert_eq!(
            source
                .edit_with_limits(
                    &SemanticLimits::builder()
                        .max_operations(1)
                        .expect("operation limit")
                        .build()
                        .expect("limits"),
                    |edit| {
                        edit.set(FIRST, props! { "text" => "one" })?;
                        edit.set(SECOND, props! { "text" => "two" })
                    },
                )
                .expect_err("operation limit")
                .code(),
            Code::LimitExceeded
        );
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn released_and_signed_sources_refuse_changed_writes_with_typed_errors() {
    let bytes = fixture();
    let released =
        litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().retain_source(false))
            .expect("released PPT");
    assert_eq!(
        released
            .set(FIRST, props! { "text" => "refused" })
            .expect_err("released source")
            .code(),
        Code::UnsupportedCapability
    );

    let signed = litchi_ops::open_bytes(signed_fixture()).expect("signed PPT");
    assert!(matches!(signed, File::Signed(_)));
    assert_eq!(
        signed
            .set(FIRST, props! { "text" => "refused" })
            .expect_err("signed source")
            .code(),
        Code::DocumentProtected
    );
}

#[cfg(any(unix, windows))]
#[test]
fn positional_source_freshness_is_checked_before_a_write() {
    use std::io::Write as _;

    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("stale.ppt");
    std::fs::write(&path, fixture()).expect("write source");
    let source = litchi_ops::open_with(&path, &OpenOptions::new()).expect("open source");
    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(&path)
        .expect("reopen source");
    changed.write_all(b"changed").expect("change source");
    changed.flush().expect("flush source");

    assert_eq!(
        source
            .set(FIRST, props! { "text" => "stale" })
            .expect_err("stale source")
            .code(),
        Code::IoError
    );
    assert_eq!(
        source.to_bytes().expect_err("stale publication").code(),
        Code::IoError
    );
}

#[test]
fn positional_save_reopens_the_native_revision_and_preserves_the_old_handle() {
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("saved.ppt");
    let bytes = fixture();
    std::fs::write(&path, &bytes).expect("write source");
    let source = litchi_ops::open_with(&path, &OpenOptions::new()).expect("open source");
    let updated = source
        .set(FIRST, props! { "text" => "saved revision" })
        .expect("set source");
    updated.save().expect("save source revision");

    let reopened = litchi_ops::open(&path).expect("reopen saved PPT");
    assert_eq!(text(&reopened, FIRST), "saved revision");
    assert_eq!(text(&source, FIRST), "first");
    assert_eq!(source.to_bytes().expect("old source bytes"), bytes);
}
