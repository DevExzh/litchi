//! End-to-end coverage for the narrow DOC unified owner.
//!
//! The tests exercise the same immutable owner through creation, reads,
//! paragraph text writes, transactions, publication and reopen.  They also
//! pin the fail-closed source and signature policies advertised by the
//! capability report.
#![cfg(feature = "doc")]

use std::io::Cursor;

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Subject, Verb};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, Conformance, CreateOptions, DocOptions, File, Flavor, OpenOptions,
    Overwrite, SaveOptions, SaveTo, SignatureOrigin, create, create_with, open, open_bytes,
    open_bytes_with, props,
};

fn paragraph_text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .expect("paragraph text")
        .into_owned()
}

fn signed_compound(bytes: Vec<u8>) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("open DOC source");
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..path.len() {
            writer
                .create_storage(&path[..depth])
                .expect("copy DOC storage");
        }
        let data = source.open_stream(&path).expect("copy DOC stream");
        writer
            .create_stream(&path, &data)
            .expect("write DOC stream");
    }
    writer
        .create_stream(&["_signatures"], b"test signature")
        .expect("add DOC signature marker");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write signed DOC");
    output.into_inner()
}

fn doc_fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc")
            .join(name),
    )
    .unwrap_or_else(|error| panic!("read DOC fixture {name}: {error}"))
}

fn multi_paragraph_doc() -> Vec<u8> {
    let mut writer = litchi_doc::Builder::new();
    writer.add_paragraph("first").expect("first paragraph");
    writer.add_paragraph("second").expect("second paragraph");
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build DOC fixture")
        .write_to(&mut output)
        .expect("write DOC fixture");
    output.into_inner()
}

#[test]
fn create_read_set_save_and_reopen_use_one_doc_owner() {
    let file = create(FileFormat::Doc).expect("create DOC");
    assert_eq!(file.format(), Some(FileFormat::Doc));
    assert_eq!(paragraph_text(&file, "/body/p[1]"), "");
    assert!(file.find("/body/p[2]").unwrap().is_none());

    let nodes = file
        .query("p")
        .expect("query DOC paragraphs")
        .collect::<litchi_ops::Result<Vec<_>>>()
        .expect("collect DOC paragraphs");
    assert_eq!(nodes.len(), 1);
    assert_eq!(nodes[0].kind(), "paragraph");
    assert_eq!(nodes[0].positional_path().to_string(), "/body/p[1]");
    assert_eq!(file.view(litchi_ops::view::Mode::Text).unwrap(), "\n");
    let json = serde_json::to_value(
        file.view_json(litchi_ops::view::Mode::Text)
            .expect("DOC JSON view"),
    )
    .expect("serialize DOC JSON view");
    assert!(json.is_object() || json.is_array() || json.is_string());

    let edited = file
        .set(
            "/body/p[1]",
            props! { "text" => "Created through unified DOC" },
        )
        .expect("set DOC paragraph text");
    assert_eq!(paragraph_text(&file, "/body/p[1]"), "");
    assert_eq!(
        paragraph_text(&edited, "/body/p[1]"),
        "Created through unified DOC"
    );

    let bytes = edited.to_bytes().expect("publish DOC bytes");
    let reopened = open_bytes(bytes.clone()).expect("reopen DOC bytes");
    assert_eq!(
        paragraph_text(&reopened, "/body/p[1]"),
        "Created through unified DOC"
    );

    let directory = tempfile::tempdir().expect("DOC save directory");
    let path = directory.path().join("created.doc");
    let saved = edited.save_as(&path).expect("save DOC path");
    assert_eq!(saved.bytes, std::fs::metadata(&path).unwrap().len());
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
    assert_eq!(
        paragraph_text(&open(&path).expect("reopen saved DOC"), "/body/p[1]"),
        "Created through unified DOC"
    );
}

#[test]
fn doc_transaction_coalesces_sets_and_rolls_back_atomically() {
    let file = create(FileFormat::Doc).expect("create DOC");
    let original = file.to_bytes().expect("read original DOC bytes");

    let commit = file
        .edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => "first" })?;
            edit.set("/body/p[1]", props! { "text" => "second" })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("coalesced DOC edit");
    assert_eq!(paragraph_text(commit.file(), "/body/p[1]"), "second");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(commit.diagnostics().leaf_operations(), 2);
    assert!(commit.diagnostics().snapshot_bytes() > 0);

    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse DOC patch");
    assert_eq!(restored.to_bytes().unwrap(), original);
    assert_eq!(
        commit.patch().apply(&file).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );

    let error = file
        .edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => "discard me" })?;
            edit.set("/body/p[2]", props! { "text" => "must fail" })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect_err("a missing paragraph aborts the complete edit");
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(file.to_bytes().unwrap(), original);
    assert_eq!(paragraph_text(&file, "/body/p[1]"), "");
}

#[test]
fn doc_batch_noop_and_limits_keep_exact_source() {
    let file = create(FileFormat::Doc).expect("create DOC");
    let original = file.to_bytes().expect("read original DOC bytes");

    let noop = file
        .set("/body/p[1]", props! { "text" => "" })
        .expect("DOC exact no-op");
    assert_eq!(noop.to_bytes().unwrap(), original);
    let noop_commit = file
        .edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => "" })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("DOC transaction no-op");
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.patch().is_noop());
    assert_eq!(noop_commit.file().to_bytes().unwrap(), original);

    let mut first = BatchItem::new("set");
    first.path = Some("/body/p[1]".into());
    first.props = Some(litchi_ops::Props::new().with("text", "batch first"));
    let mut second = BatchItem::new("set");
    second.path = Some("/body/p[1]".into());
    second.props = Some(litchi_ops::Props::new().with("text", "batch second"));
    let batch = file
        .batch([first, second], BatchMode::Atomic)
        .expect("atomic DOC text batch");
    assert_eq!(batch.value().len(), 2);
    assert_eq!(batch.diagnostics().commits(), 1);
    assert_eq!(paragraph_text(batch.file(), "/body/p[1]"), "batch second");

    let limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("one-operation limit")
        .build()
        .expect("valid limits");
    let error = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.set("/body/p[1]", props! { "text" => "one" })?;
                edit.set("/body/p[1]", props! { "text" => "two" })?;
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect_err("the second DOC mutation exceeds the limit");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(file.to_bytes().unwrap(), original);
}

#[test]
fn doc_query_expansion_uses_the_unified_target_budget() {
    let file = open_bytes(multi_paragraph_doc()).expect("open multi-paragraph DOC");
    let original = file.to_bytes().expect("read source DOC bytes");
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .expect("one-target limit")
        .build()
        .expect("valid limits");
    let query = litchi_ops::query::Query::parse("/body/p").expect("paragraph query");

    let error = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.set(query, props! { "text" => "must not stage" })?;
                Ok::<_, litchi_ops::Error>(())
            },
        )
        .expect_err("the query must stop before the second paragraph");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(file.to_bytes().unwrap(), original);
}

#[test]
fn doc_source_limits_preserve_dimension_and_values_through_facade_ingress() {
    use litchi_doc::{SourceLimitKind, SourceLimits};
    let cases = [
        (
            SourceLimits::builder()
                .max_input_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::InputBytes,
        ),
        (
            SourceLimits::builder()
                .max_entry_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::EntryBytes,
        ),
        (
            SourceLimits::builder()
                .max_aggregate_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::AggregateBytes,
        ),
    ];
    for (limits, expected) in cases {
        let options = OpenOptions::new().doc(DocOptions::new().source_limits(limits));
        let error = open_bytes_with(doc_fixture("Lists.doc"), &options)
            .expect_err("the caller-selected DOC source limit must reject the fixture");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(
            matches!(error.leaf(), litchi_ops::error::Source::Doc(litchi_doc::Error::Limit {
            kind, observed, maximum: 1, ..
        }) if *kind == expected && *observed > 1)
        );
    }
}

#[test]
fn doc_changed_candidate_retains_stream_limits_and_original_bytes() {
    use litchi_doc::{SourceLimitKind, SourceLimits};

    let bytes = multi_paragraph_doc();
    let ole = litchi_cfb::OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let sizes = ole
        .list_streams()
        .iter()
        .map(|path| {
            let path = path.iter().map(String::as_str).collect::<Vec<_>>();
            usize::try_from(ole.stream_len(&path).unwrap()).unwrap()
        })
        .collect::<Vec<_>>();
    let entry = *sizes.iter().max().unwrap();
    let aggregate = sizes.iter().sum::<usize>();
    for (limits, expected, maximum) in [
        (
            SourceLimits::builder()
                .max_entry_bytes(entry)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::EntryBytes,
            entry,
        ),
        (
            SourceLimits::builder()
                .max_aggregate_bytes(aggregate)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::AggregateBytes,
            aggregate,
        ),
    ] {
        let file = open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().doc(DocOptions::new().source_limits(limits)),
        )
        .expect("exact source limits admit original");
        let replacement = "x".repeat(100_000);
        let error = file
            .set("/body/p[1]", props! { "text" => replacement })
            .expect_err("changed candidate exceeds retained source limits");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(error.leaf(),
            litchi_ops::error::Source::DocTransaction(
                litchi_doc::package::edit::Error::Body(litchi_doc::body_text::Error::Invalid(
                    litchi_doc::Error::Limit { kind, observed, maximum: actual }
                ))
            ) if *kind == expected && *actual == maximum as u64 && *observed > *actual
        ));
        assert_eq!(file.to_bytes().unwrap(), bytes);
        assert_eq!(paragraph_text(&file, "/body/p[1]"), "first");
    }
}

#[test]
fn doc_capabilities_and_source_retention_are_narrow_and_explicit() {
    let capabilities = Capabilities::of(FileFormat::Doc);
    for verb in [
        Verb::Get,
        Verb::Find,
        Verb::Query,
        Verb::View,
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
        Verb::Save,
        Verb::Create,
    ] {
        assert!(capabilities.verb(verb).is_supported(), "{verb:?}");
    }
    let paragraph = capabilities
        .element("paragraph")
        .expect("DOC paragraph capability");
    assert!(paragraph.operations().get);
    assert!(paragraph.operations().set);
    assert!(paragraph.operations().add);
    assert!(paragraph.operations().remove);
    let text = paragraph.property("text").expect("DOC text property");
    assert!(text.get());
    assert!(text.set());
    assert!(text.add());
    let run = capabilities.element("run").expect("DOC run capability");
    assert!(run.operations().get && run.operations().set && run.operations().remove);
    assert!(!run.operations().add);
    for key in ["bold", "italic", "underline", "size", "color", "font"] {
        let property = run.property(key).expect("DOC run property");
        assert!(property.get() && property.set(), "{key}");
    }
    for element in [
        "table", "row", "cell", "header", "footer", "footnote", "endnote", "comment",
    ] {
        let element = capabilities.element(element).expect("DOC read element");
        assert!(element.operations().get && !element.operations().set);
    }
    assert!(capabilities.gaps().any(|gap| {
        gap.subject()
            == Subject::Operation {
                verb: "create",
                element: "document",
                context: "template",
            }
    }));

    let bytes = create(FileFormat::Doc)
        .unwrap()
        .to_bytes()
        .expect("create source for retention test");
    let released = open_bytes_with(bytes, &OpenOptions::new().retain_source(false))
        .expect("open released DOC");
    assert_eq!(released.capabilities().retains_source(), Some(false));
    assert_eq!(paragraph_text(&released, "/body/p[1]"), "");
    assert_eq!(
        released
            .set("/body/p[1]", props! { "text" => "refused" })
            .unwrap_err()
            .code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        released.to_bytes().unwrap_err().code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn signed_doc_requires_explicit_unsign_before_text_write() {
    let plain = create(FileFormat::Doc)
        .unwrap()
        .to_bytes()
        .expect("create DOC source");
    let signed_bytes = signed_compound(plain);
    let signed = open_bytes(signed_bytes.clone()).expect("open signed DOC");
    let wrapper = match &signed {
        File::Signed(wrapper) => wrapper,
        other => panic!("expected signed DOC, got {other:?}"),
    };
    assert_eq!(wrapper.origin(), SignatureOrigin::Cfb);
    let error = signed
        .set("/body/p[1]", props! { "text" => "blocked" })
        .expect_err("signed DOC edit");
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(error.suggestion().unwrap().contains("Signed::unsign"));
    assert_eq!(signed.to_bytes().unwrap(), signed_bytes);

    let unsigned = match signed {
        File::Signed(wrapper) => wrapper.unsign(),
        other => panic!("expected signed DOC, got {other:?}"),
    };
    let edited = unsigned
        .set("/body/p[1]", props! { "text" => "authorized" })
        .expect("write after explicit unsign");
    let output = edited.to_bytes().expect("publish unsigned DOC");
    let ole = litchi_cfb::OleFile::open(
        Cursor::new(output.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("open output DOC");
    assert!(!ole.exists(&["_signatures"]));
    assert_eq!(
        paragraph_text(&open_bytes(output).unwrap(), "/body/p[1]"),
        "authorized"
    );
}

#[test]
fn doc_creation_options_and_save_policy_fail_without_side_effects() {
    let template = create_with(
        FileFormat::Doc,
        &CreateOptions::new().flavor(Flavor::Template),
    )
    .expect_err("DOC template creation");
    assert_eq!(template.code(), Code::UnsupportedCapability);
    let strict = create_with(
        FileFormat::Doc,
        &CreateOptions::new().conformance(Conformance::Strict),
    )
    .expect_err("DOC Strict creation");
    assert_eq!(strict.code(), Code::UnsupportedCapability);

    let file = create(FileFormat::Doc).expect("create DOC");
    assert_eq!(
        file.save().unwrap_err().code(),
        Code::FileRequired,
        "a byte-created DOC has no origin path"
    );
    let directory = tempfile::tempdir().expect("DOC save directory");
    let path = directory.path().join("policy.doc");
    file.save_as(&path).expect("first DOC save");
    let existing = std::fs::read(&path).expect("read first DOC save");
    let error = file
        .save_with(
            SaveTo::Path(path.as_path()),
            &SaveOptions::new().overwrite(Overwrite::Refuse),
        )
        .expect_err("refuse existing DOC destination");
    assert_eq!(error.code(), Code::FileExists);
    assert_eq!(std::fs::read(&path).unwrap(), existing);
    let wrong = directory.path().join("policy.docx");
    let error = file
        .save_as(&wrong)
        .expect_err("wrong DOC destination extension");
    assert_eq!(error.code(), Code::InvalidArgument);
    assert!(!wrong.exists());
}

#[test]
fn doc_coalescing_exceeds_leaf_default_and_preserves_path_origin() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("origin.doc");
    create(FileFormat::Doc).unwrap().save_as(&path).unwrap();
    let original = open(&path).unwrap();
    let commit = original
        .edit(|edit| {
            for index in 0..300 {
                edit.set(
                    "/body/p[1]",
                    props! { "text" => format!("paragraph {index}") },
                )?;
            }
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("coalesced DOC run beyond native default 256");
    assert_eq!(commit.diagnostics().commits(), 1);
    commit
        .file()
        .save()
        .expect("edited DOC preserves path origin");
    assert_eq!(
        paragraph_text(&open(&path).unwrap(), "/body/p[1]"),
        "paragraph 299"
    );
}

#[test]
fn doc_owned_ingress_is_detached_from_later_path_changes() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.doc");
    create(FileFormat::Doc).unwrap().save_as(&path).unwrap();
    let file = open(&path).unwrap();
    let node = file.query("p").unwrap().next().unwrap().unwrap();
    let native = match node {
        litchi_ops::node::Node::Doc(node) => node,
        _ => panic!("expected DOC paragraph"),
    };
    // DOC ingress currently detaches the source into an owned CFB buffer.
    file.set("/body/p[1]", props! { "text" => "warm snapshot" })
        .unwrap();
    std::fs::write(&path, b"changed source").unwrap();
    assert!(file.metadata().unwrap().is_some());
    assert_eq!(native.paragraph_text().unwrap().as_deref(), Some(""));
    let edited = file
        .set("/body/p[1]", props! { "text" => "owned" })
        .unwrap();
    assert_eq!(paragraph_text(&edited, "/body/p[1]"), "owned");
}

#[test]
fn doc_outline_reports_paragraphs_without_inventing_headings_or_formatting() {
    let file = open_bytes(multi_paragraph_doc()).unwrap();
    let outline =
        serde_json::to_value(file.view_json(litchi_ops::view::Mode::Outline).unwrap()).unwrap();
    assert_eq!(outline["paragraphs"], 2);
    assert_eq!(outline["elements"][0]["text"], "first");
    assert!(outline.get("headings").is_none());
    assert!(outline.get("tables").is_none());
    let stats =
        serde_json::to_value(file.view_json(litchi_ops::view::Mode::Stats).unwrap()).unwrap();
    assert_eq!(stats["words"], 2);
    assert!(stats.get("styleDistribution").is_none());
}

fn run_prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    let view = node.view();
    view.format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn many_paragraphs(count: usize) -> File {
    let mut builder = litchi_doc::Builder::new();
    for index in 0..count {
        builder
            .add_paragraph(&format!("Paragraph {index:03} body text."))
            .unwrap();
    }
    open_bytes(builder.build().unwrap().to_bytes().unwrap()).unwrap()
}

/// Record 0359a: direct paragraph `bold`/`italic` sets join the body run
/// (one leaf commit), read back as the isolated setter's, and a paragraph
/// whose text the run replaced closes the run before a toggle.
#[test]
fn doc_paragraph_toggles_coalesce_into_one_commit() {
    let file = many_paragraphs(40);
    let commit = file
        .edit(|edit| {
            for index in 1..=40 {
                edit.set(
                    format!("/body/p[{index}]").as_str(),
                    props! { "bold" => true },
                )?;
            }
            edit.set("/body/p[3]", props! { "italic" => true })?;
            // A repeat of a staged value is a no-op inside the run.
            edit.set("/body/p[5]", props! { "bold" => true })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("coalesced toggles");
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(commit.diagnostics().operations(), 42);
    for index in 1..=40 {
        assert_eq!(
            run_prop(commit.file(), &format!("/body/p[{index}]/r[1]"), "bold").as_deref(),
            Some("true"),
            "paragraph {index}"
        );
    }
    assert_eq!(
        run_prop(commit.file(), "/body/p[3]/r[1]", "italic").as_deref(),
        Some("true")
    );
    assert_eq!(run_prop(commit.file(), "/body/p[4]/r[1]", "italic"), None);

    // The isolated setter publishes the same document, one commit per set.
    let mut isolated = file.clone();
    for index in 1..=40 {
        isolated = isolated
            .set(
                format!("/body/p[{index}]").as_str(),
                props! { "bold" => true },
            )
            .unwrap();
    }
    isolated = isolated
        .set("/body/p[3]", props! { "italic" => true })
        .unwrap();
    for index in 1..=40 {
        let path = format!("/body/p[{index}]/r[1]");
        for key in ["bold", "italic"] {
            assert_eq!(
                run_prop(commit.file(), &path, key),
                run_prop(&isolated, &path, key),
                "{path} {key}"
            );
        }
    }

    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());

    // Text then a toggle of the same paragraph: two runs.
    let mixed = file
        .edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => "Replaced text." })?;
            edit.set("/body/p[1]", props! { "bold" => true })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .expect("text then toggle");
    assert_eq!(mixed.diagnostics().commits(), 2);
    assert_eq!(paragraph_text(mixed.file(), "/body/p[1]"), "Replaced text.");
    assert_eq!(
        run_prop(mixed.file(), "/body/p[1]/r[1]", "bold").as_deref(),
        Some("true")
    );

    // A toggle the base already holds changes nothing.
    let bold = commit.file().clone();
    let noop = bold
        .edit(|edit| edit.set("/body/p[2]", props! { "bold" => true }))
        .expect("held toggle");
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.file().to_bytes().unwrap(), bold.to_bytes().unwrap());
}
