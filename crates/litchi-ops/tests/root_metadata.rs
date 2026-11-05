//! OfficeCLI's root-level metadata writes (`set / title=...`) route to the
//! `/metadata` writer for DOCX, XLSX and PPTX (records 0338 and 0342).
#![cfg(all(feature = "docx", feature = "xlsx", feature = "pptx"))]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::batch::{Options, Unsupported};
use litchi_ops::capabilities::Capabilities;
use litchi_ops::{BatchItem, Code, File, PropValue, Props, create, props};

const FORMATS: [FileFormat; 3] = [FileFormat::Docx, FileFormat::Xlsx, FileFormat::Pptx];

fn metadata(file: &File) -> Metadata {
    file.metadata().expect("read metadata").unwrap_or_default()
}

fn root_value(file: &File, key: &str) -> Option<String> {
    file.get("/")
        .expect("root node")
        .view()
        .format()
        .expect("root format")
        .find(|property| property.key() == key)
        .and_then(|property| property.value().as_str().map(str::to_owned))
}

fn core() -> Props<'static> {
    Props::new()
        .with("title", "Q3 Field Report")
        .with("author", "Jane Author")
        .with("subject", "Finance")
        .with("keywords", "report,q3")
        .with("description", "Quarterly field summary.")
        .with("lastModifiedBy", "Editorial")
}

#[test]
fn a_root_set_is_the_metadata_set() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let before = file.to_bytes().expect("bytes");
        let root = file
            .set("/", core())
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        let path = file
            .set("/metadata", core())
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        // The same writer: identical packages, and the input is unchanged.
        assert_eq!(
            root.to_bytes().expect("bytes"),
            path.to_bytes().expect("bytes"),
            "{format:?}"
        );
        assert_eq!(file.to_bytes().expect("bytes"), before, "{format:?}");
        let reopened = litchi_ops::open_bytes(root.to_bytes().expect("bytes")).expect("reopen");
        for file in [&root, &reopened] {
            let value = metadata(file);
            assert_eq!(
                value.title.as_deref(),
                Some("Q3 Field Report"),
                "{format:?}"
            );
            assert_eq!(value.author.as_deref(), Some("Jane Author"), "{format:?}");
            assert_eq!(value.last_modified_by.as_deref(), Some("Editorial"));
            assert!(value.modified.is_none(), "a set never stamps modified");
            assert!(
                root_value(file, "subject").is_some_and(|value| value.contains("Finance")),
                "{format:?}"
            );
        }
        // `null` removes, and an equal value is a no-op.
        let cleared = root
            .set("/", props! { "title" => PropValue::Null })
            .unwrap();
        assert_eq!(metadata(&cleared).title, None, "{format:?}");
        let again = root
            .set("/", props! { "title" => "Q3 Field Report" })
            .unwrap();
        assert_eq!(
            again.to_bytes().unwrap(),
            root.to_bytes().unwrap(),
            "{format:?}"
        );
    }
}

#[test]
fn the_root_refuses_other_keys_by_name() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let error = file
            .set("/", props! { "title" => "T", "calc.mode" => "manual" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty, "{format:?}");
        assert_eq!(error.property(), Some("calc.mode"), "{format:?}");
        let error = file
            .set("/", props! { "title" => "a", "TITLE" => "b" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument, "{format:?}");
    }
    // No OOXML part stores `contentType`.
    let docx = create(FileFormat::Docx).expect("create");
    let error = docx.set("/", props! { "contentType" => "x" }).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(error.property(), Some("content_type"));
    // Every OOXML format writes application properties on the root too
    // (DOCX: record 0342).
    for format in [FileFormat::Docx, FileFormat::Xlsx, FileFormat::Pptx] {
        let file = create(format)
            .unwrap()
            .set(
                "/",
                props! { "extended.company" => "Acme Corp", "extended.manager" => "Dana" },
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        if format == FileFormat::Docx {
            // DOCX `metadata()` is the core properties; the root reads
            // the application ones.
            assert_eq!(
                root_value(&file, "extended.company").as_deref(),
                Some("Acme Corp")
            );
            assert_eq!(
                root_value(&file, "extended.manager").as_deref(),
                Some("Dana")
            );
        } else {
            assert_eq!(metadata(&file).company.as_deref(), Some("Acme Corp"));
            assert_eq!(metadata(&file).manager.as_deref(), Some("Dana"));
        }
    }
}

#[test]
fn warn_skips_the_other_root_keys_and_writes_the_metadata() {
    let warn = Options::new().unsupported(Unsupported::Warn);
    for format in FORMATS {
        let file = create(format).expect("create");
        let mut item = BatchItem::new("set");
        item.path = Some("/".into());
        item.props = Some(
            Props::new()
                .with("author", "Jane")
                .with("theme.color.dk1", "1A1A1A")
                .with("title", "T"),
        );
        let commit = file
            .apply_with(&item, &warn)
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        let skipped: Vec<_> = commit
            .value()
            .warnings
            .iter()
            .map(|warning| warning.property())
            .collect();
        assert_eq!(skipped, [Some("theme.color.dk1")], "{format:?}");
        assert_eq!(metadata(commit.file()).title.as_deref(), Some("T"));
        assert_eq!(metadata(commit.file()).author.as_deref(), Some("Jane"));
        assert!(commit.diagnostics().changed(), "{format:?}");
    }
}

#[test]
fn transactions_replay_and_invert_root_sets() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let commit = file
            .edit(|edit| {
                edit.set("/", props! { "title" => "In a transaction" })?;
                edit.set("/metadata", props! { "author" => "Ada" })?;
                Ok(())
            })
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert!(!commit.patch().is_noop(), "{format:?}");
        let after = commit.file().clone();
        assert_eq!(metadata(&after).title.as_deref(), Some("In a transaction"));
        assert_eq!(metadata(&after).author.as_deref(), Some("Ada"));
        let replayed = commit.patch().apply(&file).expect("replay");
        assert_eq!(metadata(&replayed).author.as_deref(), Some("Ada"));
        let restored = commit.patch().inverse().apply(&after).expect("inverse");
        assert_eq!(metadata(&restored).title, None, "{format:?}");
        // An equal value is a no-op transaction.
        let same = after
            .edit(|edit| {
                edit.set("/", props! { "title" => "In a transaction" })?;
                Ok(())
            })
            .expect("no-op");
        assert!(same.patch().is_noop(), "{format:?}");
    }
}

#[test]
fn the_root_claims_set_for_its_document_property_keys() {
    for format in FORMATS {
        let report = Capabilities::of(format);
        let root = report
            .elements()
            .find(|element| matches!(element.name(), "document" | "workbook" | "presentation"))
            .expect("root element");
        assert!(root.operations().set, "{format:?}");
        for key in ["title", "author", "subject", "lastModifiedBy", "created"] {
            let property = root.property(key).expect(key);
            assert!(property.get() && property.set(), "{format:?} {key}");
        }
        // Derived counts are read only.
        if let Some(property) = root.property("extended.pages") {
            assert!(!property.set(), "{format:?}");
        }
    }
}

/// OfficeCLI's `word/document-formatting.sh:28` (record 0342): DOCX
/// application properties and custom properties on the root, read back as
/// `extended.*`; the core properties and the main document are untouched,
/// and a transaction step replays and inverts them.
#[test]
fn docx_root_writes_application_and_custom_properties() {
    let file = create(FileFormat::Docx).expect("create");
    let file = file
        .set("/", props! { "title" => "Q3 Field Report" })
        .expect("core");
    let before = file.to_bytes().expect("bytes");
    let after = file
        .set(
            "/",
            props! {
                "extended.company" => "Acme Corp",
                "extended.manager" => "Dana Lead",
                "extended.template" => "Normal.dotm",
            },
        )
        .expect("extended");
    assert_eq!(
        root_value(&after, "extended.company").as_deref(),
        Some("Acme Corp")
    );
    assert_eq!(
        root_value(&after, "extended.manager").as_deref(),
        Some("Dana Lead")
    );
    assert_eq!(
        root_value(&after, "extended.template").as_deref(),
        Some("Normal.dotm")
    );
    assert_eq!(metadata(&after).title.as_deref(), Some("Q3 Field Report"));
    let reopened = litchi_ops::open_bytes(after.to_bytes().expect("bytes")).expect("reopen");
    assert_eq!(
        root_value(&reopened, "extended.company").as_deref(),
        Some("Acme Corp")
    );
    // An equal value publishes nothing.
    let same = after
        .set("/", props! { "extended.company" => "Acme Corp" })
        .expect("same");
    assert_eq!(
        same.to_bytes().expect("bytes"),
        after.to_bytes().expect("bytes")
    );
    // Custom properties, in a transaction that replays and inverts.
    let commit = file
        .edit(|edit| {
            edit.set(
                "/",
                props! { "custom.Reviewed" => true, "company" => "Acme" },
            )?;
            Ok(())
        })
        .expect("custom");
    assert!(!commit.patch().is_noop());
    assert_eq!(
        root_value(commit.file(), "extended.company").as_deref(),
        Some("Acme")
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().expect("bytes"), before);
    let replayed = commit.patch().apply(&file).expect("replay");
    assert_eq!(
        root_value(&replayed, "extended.company").as_deref(),
        Some("Acme")
    );
}
