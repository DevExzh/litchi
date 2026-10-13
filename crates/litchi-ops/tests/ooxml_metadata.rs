//! OOXML document metadata through the unified root path (record 0253).
#![cfg(all(feature = "xlsx", feature = "pptx", feature = "xlsb"))]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::capabilities::Capabilities;
use litchi_ops::{Code, CreateOptions, File, PropValue, create, create_with, props};

const FORMATS: [FileFormat; 3] = [FileFormat::Xlsx, FileFormat::Pptx, FileFormat::Xlsb];

fn root_value(file: &File, key: &str) -> Option<String> {
    file.get("/")
        .expect("root node")
        .view()
        .format()
        .expect("root format")
        .find(|property| property.key() == key)
        .and_then(|property| property.value().as_str().map(str::to_owned))
}

fn metadata(file: &File) -> Metadata {
    file.metadata().expect("read metadata").unwrap_or_default()
}

#[test]
fn root_sets_write_core_and_app_properties_and_survive_reopen() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let before = file.to_bytes().expect("original bytes");
        let edited = file
            .set(
                "/",
                props! {
                    "title" => "Quarterly",
                    "creator" => "Ada",
                    "keywords" => "plan, review",
                    "category" => "Finance",
                    "lastModifiedBy" => "Lin",
                    "revisionNumber" => "3",
                    "extended.company" => "Litchi",
                    "extended.manager" => "Grace",
                },
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(file.to_bytes().expect("source bytes"), before);

        let reopened = litchi_ops::open_bytes(edited.to_bytes().expect("bytes")).expect("reopen");
        for file in [&edited, &reopened] {
            let value = metadata(file);
            assert_eq!(value.title.as_deref(), Some("Quarterly"), "{format:?}");
            assert_eq!(value.author.as_deref(), Some("Ada"), "{format:?}");
            assert_eq!(value.keywords.as_deref(), Some("plan, review"));
            assert_eq!(value.category.as_deref(), Some("Finance"));
            assert_eq!(value.last_modified_by.as_deref(), Some("Lin"));
            assert_eq!(value.revision.as_deref(), Some("3"));
            assert_eq!(value.company.as_deref(), Some("Litchi"), "{format:?}");
            assert_eq!(value.manager.as_deref(), Some("Grace"), "{format:?}");
            assert!(value.modified.is_none(), "a set never stamps modified");
            let author = root_value(file, "author");
            assert!(
                author.as_deref().is_some_and(|value| value.contains("Ada")),
                "{format:?} {author:?}"
            );
            assert!(
                root_value(file, "extended.company")
                    .unwrap()
                    .contains("Litchi"),
                "{format:?}"
            );
        }

        let cleared = edited
            .set("/", props! { "title" => PropValue::Null })
            .expect("clear title");
        assert_eq!(metadata(&cleared).title, None);
        assert_eq!(metadata(&cleared).author.as_deref(), Some("Ada"));
    }
}

#[test]
fn unknown_and_duplicate_root_keys_are_typed_refusals() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let error = file.set("/", props! { "titel" => "x" }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty, "{format:?}");
        let error = file
            .set("/", props! { "author" => "a", "creator" => "b" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument, "{format:?}");
    }
}

#[test]
fn equal_values_keep_the_snapshot() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let titled = file.set("/", props! { "title" => "Same" }).expect("set");
        let bytes = titled.to_bytes().expect("bytes");
        let again = titled
            .set("/", props! { "title" => "Same" })
            .expect("no-op");
        assert_eq!(again.to_bytes().expect("bytes"), bytes, "{format:?}");
    }
}

#[test]
fn transactions_replay_and_invert_metadata_sets() {
    for format in FORMATS {
        let file = create(format).expect("create");
        let commit = file
            .edit(|edit| {
                edit.set("/", props! { "title" => "In a transaction" })?;
                Ok(())
            })
            .expect("transaction");
        assert!(!commit.patch.is_noop(), "{format:?}");
        let after = commit.file.clone();
        assert_eq!(metadata(&after).title.as_deref(), Some("In a transaction"));
        let replayed = commit.patch.apply(&file).expect("replay");
        assert_eq!(
            metadata(&replayed).title.as_deref(),
            Some("In a transaction")
        );
        let restored = commit.patch.inverse().apply(&after).expect("inverse");
        assert_eq!(metadata(&restored).title, None);
    }
}

#[test]
fn create_time_metadata_is_authored_and_unstorable_fields_are_refused() {
    for format in FORMATS {
        let file = create_with(
            format,
            &CreateOptions::new().metadata(Metadata {
                title: Some("Created".to_owned()),
                company: Some("Litchi".to_owned()),
                ..Metadata::default()
            }),
        )
        .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(metadata(&file).title.as_deref(), Some("Created"));
        assert_eq!(metadata(&file).company.as_deref(), Some("Litchi"));
        let reopened = litchi_ops::open_bytes(file.to_bytes().expect("bytes")).expect("reopen");
        assert_eq!(metadata(&reopened).company.as_deref(), Some("Litchi"));

        let error = create_with(
            format,
            &CreateOptions::new().metadata(Metadata {
                content_type: Some("x".to_owned()),
                ..Metadata::default()
            }),
        )
        .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty, "{format:?}");
        assert_eq!(error.property(), Some("content_type"));
    }
}

#[test]
fn root_capabilities_advertise_the_settable_keys() {
    for format in FORMATS {
        let report = Capabilities::of(format);
        let root = report
            .elements()
            .find(|element| {
                element.path_segment().is_none()
                    || element.name() == "workbook"
                    || element.name() == "presentation"
            })
            .expect("root element");
        for key in ["title", "author", "extended.company"] {
            let property = root
                .property(key)
                .unwrap_or_else(|| panic!("{format:?} {key}"));
            assert!(property.set(), "{format:?} {key}");
        }
        assert!(!root.property("created").expect("created").set());
    }
}
