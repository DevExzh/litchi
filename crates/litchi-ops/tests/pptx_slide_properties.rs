//! Unified PPTX slide name and visibility writes through the public API.
#![cfg(feature = "pptx")]

use litchi_core::FileFormat;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, Props, SourceMode, props,
};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().expect("create PPTX builder");
    {
        let presentation = builder.presentation().expect("presentation writer");
        for text in ["first", "second", "third"] {
            let slide = presentation.add_slide().expect("add slide");
            slide.add_text_box(text, 0, 0, 914_400, 914_400);
        }
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
    .expect("open PPTX")
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .expect("read slide properties")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn hidden(file: &File, path: &str) -> bool {
    property(file, path, "hidden").as_deref() == Some("true")
}

fn scoped_slide_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse slide selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_set(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

#[test]
fn slide_name_and_hidden_round_trip_for_source_and_eager_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let original_bytes = source.to_bytes().expect("serialize source");

        let updated = source
            .set(
                "/slide[1]",
                props! { "name" => "Roadmap", "hidden" => true },
            )
            .expect("set slide name and hidden state");
        assert_eq!(
            property(&updated, "/slide[1]", "name").as_deref(),
            Some("Roadmap")
        );
        assert!(hidden(&updated, "/slide[1]"));
        assert_eq!(
            source.to_bytes().expect("source remains unchanged"),
            original_bytes
        );

        // Element aliases are ASCII case-insensitive, including after a
        // metadata rename; positional paths continue to identify the slide.
        assert_eq!(
            property(&updated, "/SLIDE[1]", "name").as_deref(),
            Some("Roadmap")
        );

        let reopened = open(&updated.to_bytes().expect("serialize update"), eager);
        assert_eq!(
            property(&reopened, "/slide[1]", "name").as_deref(),
            Some("Roadmap")
        );
        assert!(hidden(&reopened, "/slide[1]"));

        let visible = updated
            .set("/slide[1]", props! { "hidden" => false })
            .expect("show slide");
        assert!(!hidden(&visible, "/slide[1]"));
        assert_eq!(
            property(&visible, "/slide[1]", "name").as_deref(),
            Some("Roadmap")
        );

        let noop = updated
            .set(
                "/slide[1]",
                props! { "name" => "Roadmap", "hidden" => true },
            )
            .expect("equal slide properties are a no-op");
        assert_eq!(
            noop.to_bytes().expect("serialize no-op"),
            updated.to_bytes().expect("serialize update")
        );
    }
}

#[test]
fn slide_queries_transactions_replay_inverse_and_reopen_preserve_metadata() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let original_bytes = source.to_bytes().expect("serialize source");
        let commit = source
            .edit(|edit| {
                edit.set("/slide[1]", props! { "name" => "Agenda" })?;
                edit.set("/slide[2]", props! { "name" => "Risks" })?;
                edit.set(
                    scoped_slide_query("slide[name='Agenda'], slide[name='Risks']"),
                    props! { "hidden" => true },
                )
            })
            .expect("transactional slide metadata update");

        assert_eq!(
            property(commit.file(), "/slide[1]", "name").as_deref(),
            Some("Agenda")
        );
        assert_eq!(
            property(commit.file(), "/slide[2]", "name").as_deref(),
            Some("Risks")
        );
        assert!(hidden(commit.file(), "/slide[1]"));
        assert!(hidden(commit.file(), "/slide[2]"));
        assert!(!hidden(commit.file(), "/slide[3]"));
        assert!(commit.diagnostics().changed());

        let replayed = commit.patch().apply(&source).expect("replay slide patch");
        assert_eq!(
            replayed.to_bytes().expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse slide patch");
        assert_eq!(
            restored.to_bytes().expect("serialize inverse"),
            original_bytes
        );

        let reopened = open(&commit.file().to_bytes().expect("serialize commit"), eager);
        assert_eq!(
            property(&reopened, "/slide[1]", "name").as_deref(),
            Some("Agenda")
        );
        assert!(hidden(&reopened, "/slide[2]"));

        let noop = commit
            .file()
            .edit(|edit| {
                edit.set("/slide[1]", props! { "name" => "Agenda" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("equal slide metadata is a transaction no-op");
        assert!(!noop.diagnostics().changed());
        assert!(noop.patch().is_noop());
    }
}

#[test]
fn mixed_slide_properties_validate_atomically_and_batch_modes_report_failures() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let original_bytes = source.to_bytes().expect("serialize source");
        let original_name = property(&source, "/slide[1]", "name");

        let invalid = source
            .set(
                "/slide[1]",
                props! { "name" => "must not publish", "hidden" => "not-a-bool" },
            )
            .expect_err("invalid mixed slide properties");
        assert_eq!(invalid.code(), Code::InvalidValue);
        assert_eq!(
            source.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
        assert_eq!(property(&source, "/slide[1]", "name"), original_name);
        assert!(!hidden(&source, "/slide[1]"));

        let commands = || {
            [
                batch_set(
                    "/slide[1]",
                    Props::new()
                        .with("name", "Batch Agenda")
                        .with("hidden", true),
                ),
                batch_set("/slide[2]", Props::new().with("unsupported", true)),
                batch_set("/slide[3]", Props::new().with("name", "Batch Closing")),
            ]
        };

        let atomic = source
            .batch(commands(), BatchMode::Atomic)
            .expect_err("atomic batch must discard earlier slide updates");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(
            source.to_bytes().expect("source remains unchanged"),
            original_bytes
        );

        let best_effort = source
            .batch(commands(), BatchMode::BestEffort)
            .expect("best-effort batch");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error) if error.item() == Some(1)
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert_eq!(
            property(best_effort.file(), "/slide[1]", "name").as_deref(),
            Some("Batch Agenda")
        );
        assert!(hidden(best_effort.file(), "/slide[1]"));
        assert_eq!(
            property(best_effort.file(), "/slide[3]", "name").as_deref(),
            Some("Batch Closing")
        );
        assert_eq!(
            source.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn slide_metadata_noops_do_not_consume_transaction_changes() {
    let source = open(&fixture(), false);
    let named = source
        .set("/slide[1]", props! { "name" => "Temporary" })
        .expect("seed slide name");
    let before = named.to_bytes().expect("serialize named source");
    let commit = named
        .edit(|edit| {
            edit.set("/slide[1]", props! { "name" => "Temporary" })?;
            edit.set("/slide[1]", props! { "name" => "Temporary" })
        })
        .expect("stage equal metadata twice");
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.file().to_bytes().expect("serialize net no-op"),
        before
    );
}

#[test]
fn pptx_slide_properties_are_exposed_by_the_pptx_file_facade() {
    let file = open(&fixture(), false)
        .set("/slide[1]", props! { "name" => "Named", "hidden" => true })
        .expect("set slide properties");
    assert_eq!(file.format(), Some(FileFormat::Pptx));
    let slide = file.get("/slide[1]").expect("first slide");
    let properties: Vec<_> = slide
        .view()
        .format()
        .expect("slide format")
        .map(|property| property.key().to_owned())
        .collect();
    assert!(properties.iter().any(|key| key == "name"));
    assert!(properties.iter().any(|key| key == "hidden"));
}
