#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    MAX_MAIL_MERGE_STRING_BYTES, MailMerge, MailMergeColumnIndex, MailMergeDataSourceObject,
    MailMergeDataSourceType, MailMergeFieldMapping, RtfDocument, RtfWriter,
};
use std::borrow::Cow;

#[test]
fn parses_and_round_trips_complete_inert_mail_merge_metadata() {
    let source = concat!(
        r#"{\rtf1\ansi{\*\mailmerge\mmlinktoquery"#,
        r#"{\*\mmconnectstr Provider=SQLOLEDB;Server=invalid.example}"#,
        r#"{\*\mmconnectstrdata Integrated Security=SSPI}"#,
        r#"{\*\mmquery DROP TABLE Customers}"#,
        r#"{\*\mmdatasource file:///definitely/not/opened/customers.csv}"#,
        r#"{\*\mmheadersource headers.doc}"#,
        r#"{\*\mmodso\mmodsoactive7\mmodsocoldelim9\mmodsocolumn4"#,
        r#"\mmodsodynaddr1\mmodsofhdr1\mmodsohash123\mmodsolid44\mmodsosrc2"#,
        r#"{\*\mmodsofilter State = 'CA'}{\*\mmodsoname Customers}"#,
        r#"{\*\mmodsosort PostalCode}{\*\mmodsotable Sheet1$}"#,
        r#"{\*\mmodsoudl Provider=Text}{\*\mmodsoudldata Opaque \u20320? data}"#,
        r#"{\*\mmodsouniquetag tag-42}"#,
        r#"{\*\mmodsofldmpdata\mmodsofmcolumn2{\*\mmodsoname PostalCode}"#,
        r#"{\*\mmodsomappedname ZIP}}{\*\mmodsorecipdata 1,1,0}}}Body}"#,
    );

    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "Body");
    let merge = document.mail_merge().unwrap();
    assert!(merge.link_to_query);
    assert_eq!(merge.query.as_deref(), Some("DROP TABLE Customers"));
    assert_eq!(
        merge.data_source.as_deref(),
        Some("file:///definitely/not/opened/customers.csv")
    );
    let object = merge.data_source_object.as_ref().unwrap();
    assert_eq!(object.active_record, Some(7));
    assert_eq!(object.source_type.unwrap().rtf_value(), 2);
    assert_eq!(object.udl_data.as_deref(), Some("Opaque 你 data"));
    assert_eq!(object.field_mappings.len(), 1);
    assert_eq!(object.field_mappings[0].column.get(), 2);
    assert_eq!(object.field_mappings[0].name, "PostalCode");
    assert_eq!(object.field_mappings[0].mapped_name.as_deref(), Some("ZIP"));
    assert_eq!(object.recipient_data, ["1,1,0"]);

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), "Body");
    assert_eq!(reparsed.mail_merge(), document.mail_merge());
}

#[test]
fn ergonomic_api_round_trips_without_interpreting_values() {
    let mut document = RtfDocument::parse(r"{\rtf1 Safe body}").unwrap();
    let mut object = MailMergeDataSourceObject {
        source_type: Some(MailMergeDataSourceType::from_rtf(99)),
        ..MailMergeDataSourceObject::default()
    };
    object.field_mappings.push(
        MailMergeFieldMapping::new(MailMergeColumnIndex::new(0), "Email")
            .with_mapped_name("E-mail Address"),
    );
    let merge = MailMerge {
        connect_string: Some(Cow::Borrowed("shell:never-execute")),
        query: Some(Cow::Borrowed("DELETE FROM Contacts")),
        data_source_object: Some(object),
        ..MailMerge::default()
    };
    document.set_mail_merge(merge).unwrap();

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::from_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), "Safe body");
    assert_eq!(reparsed.mail_merge(), document.mail_merge());
}

#[test]
fn rejects_duplicate_nonstarred_nested_and_oversized_metadata() {
    for malformed in [
        r"{\rtf1{\*\mailmerge{\*\mmquery one}{\*\mmquery two}}}",
        r"{\rtf1{\mailmerge{\*\mmquery one}}}",
        r"{\rtf1{\*\mailmerge{\mmquery one}}}",
        r"{\rtf1{\*\mailmerge{\*\mmquery before{nested}after}}}",
        r"{\rtf1{\*\mailmerge{\*\mmodso{\*\mmodsofldmpdata{\*\mmodsoname MissingColumn}}}}}",
    ] {
        assert!(
            RtfDocument::parse(malformed).is_err(),
            "accepted {malformed}"
        );
    }

    let oversized = "x".repeat(MAX_MAIL_MERGE_STRING_BYTES + 1);
    let malformed = [r"{\rtf1{\*\mailmerge{\*\mmquery ", &oversized, "}}}}"].concat();
    assert!(RtfDocument::parse(&malformed).is_err());
}
