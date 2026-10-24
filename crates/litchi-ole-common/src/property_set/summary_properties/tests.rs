//! Authoring rules of [`SummaryProperties`] and its stream round trip.

use super::super::binding::Binding;
use super::super::codec::{Edit, PropertySetReader};
use super::super::model::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, SUMMARY_INFORMATION_FMTID, Section, Stream,
    USER_DEFINED_PROPERTIES_FMTID, Value,
};
use super::super::summary_information::{CREATE_DTM, FileTime, LAST_SAVE_DTM, TITLE};
use super::super::{document_summary, summary_information};
use super::{MetadataError, SYSTEM_IDENTIFIER, SummaryProperties};
use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
use std::io::Cursor;

/// Every recognized text name, with the value it reads back through.
fn populated() -> SummaryProperties {
    let mut properties = SummaryProperties::new();
    for (name, value) in [
        ("Title", "Quarterly report"),
        ("subject", "Revenue"),
        ("AUTHOR", "Ada"),
        ("Keywords", "q3; revenue"),
        ("Comments", "Draft"),
        ("Template", "Normal.dot"),
        ("Last Author", "Grace"),
        ("Revision Number", "7"),
        ("Application Name", "Litchi"),
        ("Category", "Finance"),
        ("Format", "On-screen Show (4:3)"),
        ("Manager", "Linus"),
        ("Company", "Litchi Project"),
        ("Content Type", "Report"),
        ("Content Status", "Final"),
        ("Language", "en-US"),
        ("Document Version", "1.2"),
        ("Created", "2026-10-05T09:30:00Z"),
        ("Last Save Time", "2026-10-05T10:45:30.5+02:00"),
        ("Last Printed", "2001-02-03T04:05:06Z"),
        ("Department", "Research"),
    ] {
        properties.set(name, value).unwrap();
    }
    properties
}

fn compound_file(streams: Option<super::Streams>) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], b"host").unwrap();
    if let Some(streams) = streams {
        streams.write_into(&mut writer).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn new_properties_are_empty_and_serialize_nothing() {
    let properties = SummaryProperties::new();
    assert!(properties.is_empty());
    assert_eq!(properties, SummaryProperties::default());
    assert!(properties.to_streams().unwrap().is_none());
    assert_eq!(
        properties.summary_information().codepage(),
        Some(CodePage::WINDOWS_1252)
    );
    assert_eq!(
        properties.document_summary_information().codepage(),
        Some(CodePage::WINDOWS_1252)
    );
    assert!(properties.user_defined_properties().is_none());
}

#[test]
fn recognized_names_set_the_standard_properties() {
    let properties = populated();
    assert!(!properties.is_empty());
    let summary = properties.summary_information();
    assert_eq!(summary.title(), Some("Quarterly report"));
    assert_eq!(summary.subject(), Some("Revenue"));
    assert_eq!(summary.author(), Some("Ada"));
    assert_eq!(summary.keywords(), Some("q3; revenue"));
    assert_eq!(summary.comments(), Some("Draft"));
    assert_eq!(summary.template(), Some("Normal.dot"));
    assert_eq!(summary.last_author(), Some("Grace"));
    assert_eq!(summary.revision_number(), Some("7"));
    assert_eq!(summary.app_name(), Some("Litchi"));
    assert_eq!(
        summary
            .create_time()
            .and_then(FileTime::date_time)
            .map(|time| time.to_rfc3339()),
        Some("2026-10-05T09:30:00+00:00".to_owned())
    );
    assert_eq!(
        summary
            .last_save_time()
            .and_then(FileTime::date_time)
            .map(|time| time.to_rfc3339()),
        Some("2026-10-05T08:45:30.500+00:00".to_owned())
    );
    assert!(summary.last_printed().is_some());

    let document_summary = properties.document_summary_information();
    assert_eq!(document_summary.category(), Some("Finance"));
    assert_eq!(
        document_summary.presentation_format(),
        Some("On-screen Show (4:3)")
    );
    assert_eq!(document_summary.manager(), Some("Linus"));
    assert_eq!(document_summary.company(), Some("Litchi Project"));
    assert_eq!(document_summary.content_type(), Some("Report"));
    assert_eq!(document_summary.content_status(), Some("Final"));
    assert_eq!(document_summary.language(), Some("en-US"));
    assert_eq!(document_summary.document_version(), Some("1.2"));

    let user_defined = properties.user_defined_properties().unwrap();
    assert_eq!(
        user_defined.find_named("department"),
        Some((2, &Value::Lpstr("Research".to_owned())))
    );
}

#[test]
fn unified_metadata_keys_are_aliases() {
    let mut properties = SummaryProperties::new();
    properties.set("description", "Notes").unwrap();
    properties.set("lastModifiedBy", "Grace").unwrap();
    properties.set("revisionNumber", "3").unwrap();
    properties.set("modified", "2026-10-05T00:00:00Z").unwrap();
    let summary = properties.summary_information();
    assert_eq!(summary.comments(), Some("Notes"));
    assert_eq!(summary.last_author(), Some("Grace"));
    assert_eq!(summary.revision_number(), Some("3"));
    assert!(summary.last_save_time().is_some());
    assert!(properties.user_defined_properties().is_none());
}

#[test]
fn refused_values_leave_every_property_unchanged() {
    let mut properties = populated();
    let before = properties.clone();
    for (name, value) in [
        ("Number of Pages", "3"),
        ("Security", "1"),
        ("Hyperlink base", "https://example.com/"),
        ("_MarkAsFinal", "true"),
        ("Revision Number", "seven"),
        ("Created", "yesterday"),
        ("Created", "1600-01-01T00:00:00Z"),
        ("Created", "2026-10-05T09:30:00.123456789Z"),
        ("Title", "nul\0inside"),
        ("Company", "nul\0inside"),
        ("Department", "nul\0inside"),
        ("", "empty name"),
        ("bad\0name", "value"),
    ] {
        assert!(
            properties.set(name, value).is_err(),
            "{name:?} = {value:?} was accepted"
        );
        assert_eq!(properties, before, "{name:?} changed the properties");
    }
    let oversized = "x".repeat((1 << 20) + 1);
    assert!(properties.set("Title", &oversized).is_err());
    assert!(properties.set("Department", &oversized).is_err());
    assert_eq!(properties, before);
}

#[test]
fn text_outside_windows_1252_switches_only_its_section_to_utf16() {
    let mut properties = SummaryProperties::new();
    properties.set("Author", "Ada").unwrap();
    properties.set("Title", "季度报告").unwrap();
    properties.set("Company", "Litchi").unwrap();
    properties.set("部门", "研究").unwrap();

    let summary = properties.summary_information();
    assert_eq!(summary.codepage(), Some(CodePage::Utf16Le));
    assert_eq!(summary.title(), Some("季度报告"));
    assert_eq!(summary.author(), Some("Ada"));
    assert_eq!(
        summary.property(TITLE),
        Some(&Value::Lpwstr("季度报告".to_owned()))
    );
    assert_eq!(
        properties.document_summary_information().codepage(),
        Some(CodePage::WINDOWS_1252)
    );
    let user_defined = properties.user_defined_properties().unwrap();
    assert_eq!(user_defined.page(), Some(CodePage::Utf16Le));
    assert_eq!(
        user_defined.find_named("部门").map(|(_, value)| value),
        Some(&Value::Lpwstr("研究".to_owned()))
    );

    let streams = properties.to_streams().unwrap().unwrap();
    let summary = Stream::parse(streams.summary_information()).unwrap();
    assert_eq!(
        summary_information::Snapshot::from_stream(&summary)
            .unwrap()
            .title(),
        Some("季度报告")
    );
    let document_summary = Stream::parse(streams.document_summary_information()).unwrap();
    assert_eq!(
        document_summary
            .section(USER_DEFINED_PROPERTIES_FMTID)
            .unwrap()
            .find_named("部门")
            .map(|(_, value)| value.clone()),
        Some(Value::Lpwstr("研究".to_owned()))
    );
}

#[test]
fn typed_edits_keep_the_callers_code_page() {
    let mut properties = SummaryProperties::new();
    assert!(
        properties
            .edit_summary_information(|edit| edit.set_title("季度"))
            .is_err()
    );
    assert!(
        properties
            .edit_document_summary_information(|edit| edit.set_company("公司"))
            .is_err()
    );
    assert!(properties.is_empty());
    assert!(
        properties
            .edit_summary_information(|edit| {
                edit.set_codepage(CodePage::Utf16Le)?;
                edit.set_title("季度")
            })
            .unwrap()
    );
    assert!(
        properties
            .edit_document_summary_information(|edit| {
                edit.set_codepage(CodePage::Utf16Le);
                edit.set_company("公司")
            })
            .unwrap()
    );
    assert!(
        properties
            .edit_document_summary_information(|edit| edit.set_slide_count(4))
            .unwrap()
    );
    assert!(
        !properties
            .edit_document_summary_information(|edit| edit.set_slide_count(4))
            .unwrap()
    );
    assert_eq!(
        properties.document_summary_information().company(),
        Some("公司")
    );
    assert_eq!(
        properties.document_summary_information().slide_count(),
        Some(4)
    );
}

#[test]
fn typed_user_defined_values_and_identifiers() {
    let mut properties = SummaryProperties::new();
    properties.set_user_defined("Pages", Value::I4(12)).unwrap();
    properties
        .set_user_defined("Ratio", Value::R8(0.75))
        .unwrap();
    properties
        .set_user_defined("Approved", Value::Bool(true))
        .unwrap();
    properties
        .set_user_defined("Due", Value::Filetime(133_000_000_000_000_000))
        .unwrap();
    properties
        .set_user_defined("Owner", Value::Lpwstr("Ada".to_owned()))
        .unwrap();
    // Replacing keeps the first spelling and identifier.
    properties.set_user_defined("PAGES", Value::I4(13)).unwrap();
    for value in [
        Value::I2(1),
        Value::UI4(1),
        Value::Bstr("text".to_owned()),
        Value::Empty,
    ] {
        assert!(properties.set_user_defined("Other", value).is_err());
    }

    let section = properties.user_defined_properties().unwrap();
    assert_eq!(section.page(), Some(CodePage::WINDOWS_1252));
    assert_eq!(section.find_named("pages"), Some((2, &Value::I4(13))));
    assert_eq!(section.property_name(2), Some("Pages"));
    assert_eq!(section.find_named("Ratio"), Some((3, &Value::R8(0.75))));
    assert_eq!(
        section.find_named("Approved"),
        Some((4, &Value::Bool(true)))
    );
    assert_eq!(
        section.find_named("Due"),
        Some((5, &Value::Filetime(133_000_000_000_000_000)))
    );
    // Text takes the section's code-page form.
    assert_eq!(
        section.find_named("Owner"),
        Some((6, &Value::Lpstr("Ada".to_owned())))
    );
    assert!(section.find_named("Other").is_none());
}

#[test]
fn user_defined_edits_validate_the_whole_section() {
    let mut properties = SummaryProperties::new();
    assert!(!properties.edit_user_defined_properties(|_| Ok(())).unwrap());
    assert!(properties.is_empty());
    let edits: [fn(&mut Section) -> Result<(), litchi_cfb::OleError>; 5] = [
        |section| section.add_named(2, "Small".to_owned(), Value::I2(1)),
        |section| {
            section.clear_page();
            section.add_named(2, "Text".to_owned(), Value::Lpstr("x".to_owned()))
        },
        |section| section.add_named(2, "Text".to_owned(), Value::Lpstr("季".to_owned())),
        |section| section.add_named(2, "名".to_owned(), Value::I4(1)),
        |section| {
            section.format_identifier = DOCUMENT_SUMMARY_INFORMATION_FMTID;
            Ok(())
        },
    ];
    for edit in edits {
        assert!(properties.edit_user_defined_properties(edit).is_err());
        assert!(properties.is_empty());
    }
    assert!(
        properties
            .edit_user_defined_properties(|section| {
                section.add_named(9, "Nine".to_owned(), Value::Bool(false))
            })
            .unwrap()
    );
    assert!(!properties.is_empty());
    // Removing the last property drops the section again.
    assert!(
        properties
            .edit_user_defined_properties(|section| {
                section.remove_named("nine");
                Ok(())
            })
            .unwrap()
    );
    assert!(properties.user_defined_properties().is_none());
    assert!(properties.is_empty());
}

#[test]
fn streams_round_trip_through_the_readers() {
    let properties = populated();
    let streams = properties.to_streams().unwrap().unwrap();

    let summary = Stream::parse(streams.summary_information()).unwrap();
    assert!(summary.tolerance_report().is_clean());
    assert_eq!(summary.version, Stream::VERSION_0);
    assert_eq!(summary.system_identifier, SYSTEM_IDENTIFIER);
    assert_eq!(summary.class_identifier.as_bytes(), &[0; 16]);
    assert_eq!(summary.sections.len(), 1);
    assert_eq!(
        &summary_information::Snapshot::from_stream(&summary).unwrap(),
        properties.summary_information()
    );

    let document_summary = Stream::parse(streams.document_summary_information()).unwrap();
    assert!(document_summary.tolerance_report().is_clean());
    assert_eq!(document_summary.system_identifier, SYSTEM_IDENTIFIER);
    assert_eq!(document_summary.sections.len(), 2);
    assert_eq!(
        document_summary.sections[0].format_identifier,
        DOCUMENT_SUMMARY_INFORMATION_FMTID
    );
    assert_eq!(
        &document_summary::Snapshot::from_stream(&document_summary).unwrap(),
        properties.document_summary_information()
    );
    assert_eq!(
        document_summary.section(USER_DEFINED_PROPERTIES_FMTID),
        properties.user_defined_properties()
    );

    // The projection the unified layer reads.
    let bytes = compound_file(Some(streams));
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    let metadata = ole.get_metadata().unwrap();
    assert_eq!(metadata.codepage, Some(1252));
    assert_eq!(metadata.title.as_deref(), Some("Quarterly report"));
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(metadata.subject.as_deref(), Some("Revenue"));
    assert_eq!(metadata.keywords.as_deref(), Some("q3; revenue"));
    assert_eq!(metadata.comments.as_deref(), Some("Draft"));
    assert_eq!(metadata.last_saved_by.as_deref(), Some("Grace"));
    assert_eq!(metadata.company.as_deref(), Some("Litchi Project"));
    assert_eq!(metadata.category.as_deref(), Some("Finance"));
    assert_eq!(metadata.manager.as_deref(), Some("Linus"));
    assert_eq!(
        metadata.create_time.map(|time| time.to_rfc3339()),
        Some("2026-10-05T09:30:00+00:00".to_owned())
    );
    assert_eq!(
        metadata.custom_properties.get("Department"),
        Some(&Value::Lpstr("Research".to_owned()))
    );
}

#[test]
fn serialization_is_deterministic_and_follows_set_order() {
    let first = populated().to_streams().unwrap().unwrap();
    let second = populated().to_streams().unwrap().unwrap();
    assert_eq!(first, second);

    let mut forward = SummaryProperties::new();
    forward.set("Title", "T").unwrap();
    forward.set("Author", "A").unwrap();
    let mut reverse = SummaryProperties::new();
    reverse.set("Author", "A").unwrap();
    reverse.set("Title", "T").unwrap();
    let forward = forward.to_streams().unwrap().unwrap();
    let reverse = reverse.to_streams().unwrap().unwrap();
    assert_ne!(forward.summary_information(), reverse.summary_information());
    assert_eq!(
        Stream::parse(forward.summary_information())
            .unwrap()
            .sections[0]
            .property_ids()
            .collect::<Vec<_>>(),
        vec![1, TITLE, 4]
    );
}

#[test]
fn a_document_summary_stream_is_written_with_its_code_page() {
    let mut properties = SummaryProperties::new();
    properties
        .edit_summary_information(|edit| {
            edit.set_create_time(FileTime::from_raw(133_000_000_000_000_000))
        })
        .unwrap();
    let streams = properties.to_streams().unwrap().unwrap();
    let (summary, document_summary) = streams.clone().into_parts();
    assert_eq!(summary, streams.summary_information());
    let document_summary = Stream::parse(&document_summary).unwrap();
    assert_eq!(document_summary.sections.len(), 1);
    assert_eq!(
        document_summary.sections[0].page(),
        Some(CodePage::WINDOWS_1252)
    );
    assert!(
        summary_information::Snapshot::from_stream(&Stream::parse(&summary).unwrap())
            .unwrap()
            .property(CREATE_DTM)
            .is_some()
    );
    assert!(
        summary_information::Snapshot::from_stream(&Stream::parse(&summary).unwrap())
            .unwrap()
            .property(LAST_SAVE_DTM)
            .is_none()
    );
}

#[test]
fn written_streams_are_editable_by_the_shared_editor() {
    let bytes = compound_file(populated().to_streams().unwrap());
    let mut editor = Edit::new(bytes, OleFileLimits::default()).unwrap();
    let mut summary = editor
        .property_set(Binding::SummaryInformation)
        .unwrap()
        .unwrap();
    summary.replace(TITLE, Value::Lpstr("Edited".to_owned()));
    editor
        .replace(Binding::SummaryInformation, summary)
        .unwrap();
    let mut ole = OleFile::open(
        Cursor::new(editor.finish().unwrap()),
        OleFileLimits::default(),
    )
    .unwrap();
    let stream = ole.property_set(Binding::SummaryInformation).unwrap();
    assert_eq!(
        summary_information::Snapshot::from_stream(&stream)
            .unwrap()
            .title(),
        Some("Edited")
    );
    let document_summary = ole
        .property_set(Binding::DocumentSummaryInformation)
        .unwrap();
    assert!(
        document_summary
            .section(USER_DEFINED_PROPERTIES_FMTID)
            .unwrap()
            .find_named("Department")
            .is_some()
    );
    assert!(
        document_summary
            .section(SUMMARY_INFORMATION_FMTID)
            .is_none()
    );
}

#[test]
fn section_emptiness_counts_the_code_page() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    assert!(section.is_empty());
    section.set_page(CodePage::WINDOWS_1252);
    assert!(!section.is_empty());
    section.clear_page();
    assert!(section.is_empty());
}

fn full_metadata() -> litchi_core::Metadata {
    use chrono::TimeZone as _;
    litchi_core::Metadata {
        title: Some("Quarterly report".to_owned()),
        subject: Some("Revenue".to_owned()),
        author: Some("Ada".to_owned()),
        keywords: vec!["q3".to_owned(), "revenue".to_owned()],
        description: Some("Draft".to_owned()),
        language: Some("en-US".to_owned()),
        template: Some("Normal.dot".to_owned()),
        last_modified_by: Some("Grace".to_owned()),
        revision: Some("7".to_owned()),
        created: Some(
            chrono::Utc
                .with_ymd_and_hms(2026, 10, 5, 9, 30, 0)
                .unwrap()
                .into(),
        ),
        modified: Some(
            chrono::Utc
                .with_ymd_and_hms(2026, 10, 6, 9, 30, 0)
                .unwrap()
                .into(),
        ),
        last_printed: Some(
            chrono::Utc
                .with_ymd_and_hms(2001, 2, 3, 4, 5, 6)
                .unwrap()
                .into(),
        ),
        page_count: Some(3),
        word_count: Some(300),
        character_count: Some(1800),
        character_count_with_spaces: Some(2100),
        editing_time_minutes: Some(42),
        application: Some("Litchi".to_owned()),
        category: Some("Finance".to_owned()),
        company: Some("Litchi Project".to_owned()),
        manager: Some("Linus".to_owned()),
        content_status: Some("Final".to_owned()),
        content_type: Some("Report".to_owned()),
        version: Some("1.2".to_owned()),
        security: Some(2),
        codepage: Some(1252),
        ..litchi_core::Metadata::default()
    }
}

#[test]
fn unified_metadata_round_trips_through_the_streams() {
    let metadata = full_metadata();
    let properties = SummaryProperties::from_metadata(&metadata).unwrap();
    let bytes = compound_file(properties.to_streams().unwrap());
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    let read: litchi_core::Metadata = ole.get_metadata().unwrap().into();
    assert_eq!(read.keywords, metadata.keywords);
    for (field, left, right) in [
        ("title", &read.title, &metadata.title),
        ("subject", &read.subject, &metadata.subject),
        ("author", &read.author, &metadata.author),
        ("description", &read.description, &metadata.description),
        ("template", &read.template, &metadata.template),
        (
            "last_modified_by",
            &read.last_modified_by,
            &metadata.last_modified_by,
        ),
        ("revision", &read.revision, &metadata.revision),
        ("application", &read.application, &metadata.application),
        ("category", &read.category, &metadata.category),
        ("company", &read.company, &metadata.company),
        ("manager", &read.manager, &metadata.manager),
    ] {
        assert_eq!(left, right, "{field}");
    }
    assert_eq!(read.created, metadata.created);
    assert_eq!(read.modified, metadata.modified);
    assert_eq!(read.last_printed, metadata.last_printed);
    assert_eq!(read.page_count, metadata.page_count);
    assert_eq!(read.word_count, metadata.word_count);
    assert_eq!(read.character_count, metadata.character_count);
    // The unified read projection does not carry these yet; the typed
    // sections do.
    let document_summary = properties.document_summary_information();
    assert_eq!(document_summary.language(), Some("en-US"));
    assert_eq!(document_summary.content_status(), Some("Final"));
    assert_eq!(document_summary.content_type(), Some("Report"));
    assert_eq!(document_summary.document_version(), Some("1.2"));
    assert_eq!(document_summary.character_count_with_spaces(), Some(2100));
    assert_eq!(
        properties
            .summary_information()
            .edit_time()
            .map(FileTime::raw),
        Some(42 * 600_000_000)
    );
    assert_eq!(read.security, metadata.security);
    assert_eq!(read.codepage, metadata.codepage);
    // Empty metadata authors nothing, so a writer keeps its exact output.
    assert!(
        SummaryProperties::from_metadata(&litchi_core::Metadata::default())
            .unwrap()
            .is_empty()
    );
}

#[test]
fn unrepresentable_metadata_fields_are_refused_atomically() {
    let unzoned = chrono::NaiveDate::from_ymd_opt(2026, 1, 2)
        .unwrap()
        .and_hms_opt(3, 4, 5)
        .unwrap();
    let cases: [(&str, litchi_core::Metadata); 6] = [
        (
            "identifier",
            litchi_core::Metadata {
                identifier: Some("urn:x".to_owned()),
                ..litchi_core::Metadata::default()
            },
        ),
        (
            "hyperlink_base",
            litchi_core::Metadata {
                hyperlink_base: Some("https://example.invalid/".to_owned()),
                ..litchi_core::Metadata::default()
            },
        ),
        (
            "created",
            litchi_core::Metadata {
                created: Some(unzoned.into()),
                ..litchi_core::Metadata::default()
            },
        ),
        (
            "modified",
            litchi_core::Metadata {
                modified: Some(unzoned.into()),
                ..litchi_core::Metadata::default()
            },
        ),
        (
            "last_printed",
            litchi_core::Metadata {
                last_printed: Some(unzoned.into()),
                ..litchi_core::Metadata::default()
            },
        ),
        (
            "last_backup",
            litchi_core::Metadata {
                last_backup: Some(unzoned.into()),
                ..litchi_core::Metadata::default()
            },
        ),
    ];
    for (field, mut metadata) in cases {
        metadata.title = Some("kept out".to_owned());
        let mut properties = populated();
        let before = properties.clone();
        let error = properties.merge_metadata(&metadata).unwrap_err();
        assert!(
            matches!(error, MetadataError::Unsupported { field: refused } if refused == field),
            "{field}: {error:?}"
        );
        assert_eq!(properties, before, "{field}");
    }

    let mut properties = populated();
    let before = properties.clone();
    let invalid = litchi_core::Metadata {
        title: Some("set first".to_owned()),
        revision: Some("seven".to_owned()),
        ..litchi_core::Metadata::default()
    };
    let error = properties.merge_metadata(&invalid).unwrap_err();
    assert_eq!(error.field(), "revision");
    assert!(matches!(error, MetadataError::Invalid { .. }));
    assert_eq!(properties, before);
    // The error names the field, never the refused value.
    assert!(!error.to_string().contains("seven"));
}

#[test]
fn metadata_merge_keeps_absent_fields_and_reports_changes() {
    let mut properties = populated();
    let unchanged = litchi_core::Metadata {
        title: Some("Quarterly report".to_owned()),
        ..litchi_core::Metadata::default()
    };
    assert!(!properties.merge_metadata(&unchanged).unwrap());
    let retitle = litchi_core::Metadata {
        title: Some("Annual report".to_owned()),
        ..litchi_core::Metadata::default()
    };
    assert!(properties.merge_metadata(&retitle).unwrap());
    assert_eq!(
        properties.summary_information().title(),
        Some("Annual report")
    );
    assert_eq!(properties.summary_information().author(), Some("Ada"));
}

#[test]
fn metadata_keywords_keep_the_source_separators_until_the_list_changes() {
    let mut properties = populated();
    let same = litchi_core::Metadata {
        keywords: vec!["q3".to_owned(), "revenue".to_owned()],
        ..litchi_core::Metadata::default()
    };
    assert!(!properties.merge_metadata(&same).unwrap());
    assert_eq!(
        properties.summary_information().keywords(),
        Some("q3; revenue")
    );
    let changed = litchi_core::Metadata {
        keywords: vec!["q4".to_owned(), "revenue".to_owned()],
        ..litchi_core::Metadata::default()
    };
    assert!(properties.merge_metadata(&changed).unwrap());
    assert_eq!(
        properties.summary_information().keywords(),
        Some("q4, revenue")
    );
    let before = properties.clone();
    let unstorable = litchi_core::Metadata {
        keywords: vec!["a,b".to_owned()],
        ..litchi_core::Metadata::default()
    };
    assert!(matches!(
        properties.merge_metadata(&unstorable),
        Err(MetadataError::Invalid {
            field: "keywords",
            ..
        })
    ));
    assert_eq!(properties, before);
}

#[test]
fn remove_is_the_inverse_of_set() {
    let mut properties = SummaryProperties::new();
    properties.set("title", "Title").unwrap();
    properties.set("company", "Company").unwrap();
    properties.set("created", "2026-10-05T09:30:00Z").unwrap();
    properties.set("Department", "Research").unwrap();
    for name in ["title", "company", "created", "department"] {
        assert!(properties.remove(name).unwrap(), "{name}");
        assert!(!properties.remove(name).unwrap(), "{name}");
    }
    assert!(properties.is_empty());
    assert_eq!(properties, SummaryProperties::new());
    let before = properties.clone();
    assert!(properties.remove("Number of Pages").is_err());
    assert!(properties.remove("_PID_HLINKS").is_err());
    assert_eq!(properties, before);
}

#[test]
fn editor_bridge_stages_only_changed_sections() {
    let mut writer = OleWriter::new();
    let clsid = [7_u8; 16];
    writer.set_root_clsid(clsid);
    writer.create_stream(&["Workbook"], b"host").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = output.into_inner();
    let source_clsid = OleFile::open(Cursor::new(source.clone()), OleFileLimits::default())
        .unwrap()
        .root_entry()
        .unwrap()
        .clsid
        .clone();
    let blank_clsid = {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], b"host").unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        OleFile::open(Cursor::new(output.into_inner()), OleFileLimits::default())
            .unwrap()
            .root_entry()
            .unwrap()
            .clsid
            .clone()
    };
    assert_ne!(source_clsid, blank_clsid);

    // An exact no-op stages nothing and finishes with the source bytes.
    let editor = Edit::new(source.clone(), OleFileLimits::default()).unwrap();
    let base = SummaryProperties::from_editor(&editor).unwrap();
    assert!(base.is_empty());
    let mut editor = editor;
    assert!(!base.stage_changes(&base, &mut editor).unwrap());
    assert_eq!(editor.finish().unwrap(), source);

    // A missing stream is created with a code page, and the root CLSID of
    // the source survives the rebuild.
    let mut editor = Edit::new(source, OleFileLimits::default()).unwrap();
    let mut edited = base.clone();
    edited.set("title", "Created").unwrap();
    edited.set("Department", "Research").unwrap();
    assert!(edited.stage_changes(&base, &mut editor).unwrap());
    let bytes = editor.finish().unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes.clone()), OleFileLimits::default()).unwrap();
    assert_eq!(ole.root_entry().unwrap().clsid, source_clsid);
    let stream = ole.property_set(Binding::SummaryInformation).unwrap();
    let summary = summary_information::Snapshot::from_stream(&stream).unwrap();
    assert_eq!(summary.title(), Some("Created"));
    assert_eq!(summary.codepage(), Some(CodePage::WINDOWS_1252));
    let document_summary = ole
        .property_set(Binding::DocumentSummaryInformation)
        .unwrap();
    assert_eq!(
        document_summary::Snapshot::from_stream(&document_summary)
            .unwrap()
            .codepage(),
        Some(CodePage::WINDOWS_1252)
    );

    // Removing the last custom property drops the user-defined section.
    let mut editor = Edit::new(bytes, OleFileLimits::default()).unwrap();
    let base = SummaryProperties::from_editor(&editor).unwrap();
    assert_eq!(base.summary_information().title(), Some("Created"));
    let mut edited = base.clone();
    assert!(edited.remove("department").unwrap());
    assert!(edited.stage_changes(&base, &mut editor).unwrap());
    let mut ole = OleFile::open(
        Cursor::new(editor.finish().unwrap()),
        OleFileLimits::default(),
    )
    .unwrap();
    let document_summary = ole
        .property_set(Binding::DocumentSummaryInformation)
        .unwrap();
    assert!(
        document_summary
            .section(USER_DEFINED_PROPERTIES_FMTID)
            .is_none()
    );
}

#[test]
fn blank_sections_read_as_new_properties() {
    let blank_summary = Section::new(SUMMARY_INFORMATION_FMTID);
    let blank_document_summary = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    let properties =
        SummaryProperties::from_sections(Some(&blank_summary), Some(&blank_document_summary), None)
            .unwrap();
    assert_eq!(properties, SummaryProperties::new());
}
