//! Legacy OLE metadata through the unified layer (record 0254): initial
//! metadata at creation and the `/metadata` set path for DOC, XLS and PPT.
#![cfg(all(feature = "doc", feature = "xls", feature = "ppt"))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests favor concise panic-driven assertions"
)]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::props::PropValue;
use litchi_ops::{Code, CreateOptions, File, create, create_with, open_bytes, props};

const FORMATS: [FileFormat; 3] = [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt];

fn metadata() -> Metadata {
    Metadata {
        title: Some("Quarterly report".to_owned()),
        author: Some("Ada Lovelace".to_owned()),
        company: Some("Litchi Project".to_owned()),
        created: Some("2026-10-05T09:30:00Z".parse().unwrap()),
        ..Metadata::default()
    }
}

fn properties(file: &File) -> Metadata {
    file.metadata().unwrap().unwrap_or_default()
}

#[test]
fn creation_authors_initial_metadata() {
    for format in FORMATS {
        let file = create_with(format, &CreateOptions::new().metadata(metadata()))
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        let expected = metadata();
        for file in [file.clone(), open_bytes(file.to_bytes().unwrap()).unwrap()] {
            let read = properties(&file);
            assert_eq!(read.title, expected.title, "{format:?}");
            assert_eq!(read.author, expected.author, "{format:?}");
            assert_eq!(read.company, expected.company, "{format:?}");
            assert_eq!(read.created, expected.created, "{format:?}");
        }

        // Empty metadata authors nothing: the output is the plain creator's.
        let plain = create(format).unwrap().to_bytes().unwrap();
        let empty = create_with(format, &CreateOptions::new().metadata(Metadata::default()))
            .unwrap()
            .to_bytes()
            .unwrap();
        assert_eq!(empty, plain, "{format:?}");

        // A field without a property-set slot is refused, never dropped.
        let error = create_with(
            format,
            &CreateOptions::new().metadata(Metadata {
                identifier: Some("urn:x".to_owned()),
                ..metadata()
            }),
        )
        .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{format:?}");
        assert_eq!(error.format(), Some(format));
        assert_eq!(error.property(), Some("identifier"));
    }
}

#[test]
fn metadata_path_sets_removes_and_reverts() {
    for format in FORMATS {
        let file = create(format).unwrap();
        let original = file.to_bytes().unwrap();
        let edited = file
            .set(
                "/metadata",
                props! {
                    "title" => "Retitled",
                    "lastModifiedBy" => "Grace",
                    "revisionNumber" => 3,
                    "modified" => "2026-10-06T08:00:00Z",
                    "custom.Department" => "Research",
                },
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        let read = properties(&edited);
        assert_eq!(read.title.as_deref(), Some("Retitled"), "{format:?}");
        assert_eq!(read.last_modified_by.as_deref(), Some("Grace"));
        assert_eq!(read.revision.as_deref(), Some("3"));
        assert_eq!(read.modified, Some("2026-10-06T08:00:00Z".parse().unwrap()));
        assert_eq!(
            file.to_bytes().unwrap(),
            original,
            "{format:?} source untouched"
        );

        let reopened = open_bytes(edited.to_bytes().unwrap()).unwrap();
        assert_eq!(read_title(&reopened), Some("Retitled".to_owned()));

        let cleared = edited
            .set("/metadata", props! { "title" => PropValue::Null })
            .unwrap();
        assert!(properties(&cleared).title.is_none(), "{format:?}");
        assert_eq!(
            properties(&cleared).last_modified_by.as_deref(),
            Some("Grace")
        );

        // Transactions record the step and invert it exactly.
        let commit = file
            .edit(|edit| {
                edit.set("/metadata", props! { "subject" => "Subject" })?;
                edit.set("/metadata", props! { "category" => "Finance" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .unwrap();
        assert!(commit.diagnostics().changed());
        assert_eq!(
            properties(commit.file()).subject.as_deref(),
            Some("Subject")
        );
        assert_eq!(
            properties(commit.file()).category.as_deref(),
            Some("Finance")
        );
        let restored = commit.patch().inverse().apply(commit.file()).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), original, "{format:?}");

        // Refusals are typed and leave the source untouched.
        for (props, code) in [
            (props! { "identifier" => "x" }, Code::UnsupportedProperty),
            (props! { "created" => "yesterday" }, Code::InvalidValue),
            (props! { "title" => true }, Code::InvalidValue),
        ] {
            let error = file.set("/metadata", props).unwrap_err();
            assert_eq!(error.code(), code, "{format:?}");
            assert_eq!(error.format(), Some(format));
        }
        assert_eq!(file.to_bytes().unwrap(), original);
    }
}

fn read_title(file: &File) -> Option<String> {
    properties(file).title
}

#[test]
fn doc_metadata_and_text_edits_share_one_history() {
    let file = create(FileFormat::Doc).unwrap();
    let original = file.to_bytes().unwrap();
    let commit = file
        .edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => "Body" })?;
            edit.set("/metadata", props! { "title" => "Mixed" })?;
            edit.set("/body/p[1]", props! { "text" => "Body again" })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .unwrap();
    assert_eq!(properties(commit.file()).title.as_deref(), Some("Mixed"));
    assert!(commit.file().text().unwrap().contains("Body again"));
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), original);
}

#[test]
fn xls_sheet_named_metadata_stays_renamable() {
    let file = create(FileFormat::Xls).unwrap();
    let renamed = file
        .set("/Sheet1", props! { "name" => "metadata" })
        .unwrap();
    let renamed = renamed
        .set("/metadata", props! { "name" => "Data" })
        .unwrap();
    assert_eq!(
        renamed.workbook().unwrap().worksheet_names().unwrap(),
        ["Data"]
    );
}
