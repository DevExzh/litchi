//! Public API and codec regressions for the mail-merge owner.

use super::model::{MAX_STRING_BYTES, STRICT_W, W};
use super::{
    Conformance, DataType, Destination, FieldMappingType, MainDocumentType, Recipient, Recipients,
    parse_settings_mail_merge,
};
use crate::Error;

const SETTINGS: &str = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:mailMerge><w:mainDocumentType w:val="email"/><w:linkToQuery/><w:dataType w:val="native"/><w:connectString w:val="Provider=Inert&amp;X=1"/><w:query w:val="SELECT * FROM inert"/><w:dataSource r:id="source"/><w:destination w:val="email"/><w:mailSubject w:val="A &amp; B"/><w:mailAsAttachment/><w:viewMergedData/><w:activeRecord w:val="3"/><w:checkErrors w:val="3"/><w:odso><w:table w:val="Sheet1$"/><w:src r:id="source"/><w:colDelim w:val="9"/><w:type w:val="database"/><w:fHdr/><w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="Name"/><w:mappedName w:val="Last Name"/><w:column w:val="0"/><w:lid w:val="en-US"/><w:dynamicAddress/></w:fieldMapData><w:recipientData r:id="recipients"/></w:odso></w:mailMerge><w:trackRevisions/></w:settings>"#;

#[test]
fn parses_and_deterministically_writes_complete_strict_metadata() {
    let value = parse_settings_mail_merge(SETTINGS.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(value.main_document_type(), MainDocumentType::Email);
    assert_eq!(value.data_type(), Some(DataType::Native));
    assert_eq!(value.destination(), Destination::Email);
    assert_eq!(value.query(), Some("SELECT * FROM inert"));
    assert_eq!(
        value.odso().unwrap().field_maps()[0].mapping_type(),
        Some(FieldMappingType::DatabaseColumn)
    );
    let fragment = value.to_xml(Conformance::Strict).unwrap();
    let wrapped = format!(r#"<s:settings xmlns:s="{STRICT_W}">{fragment}</s:settings>"#);
    let reparsed = parse_settings_mail_merge(wrapped.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(reparsed, value);
    assert_eq!(reparsed.to_xml(Conformance::Strict).unwrap(), fragment);
}

#[test]
fn applies_defaults_and_mce_fallback_and_preservation() {
    let xml = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><w:mailMerge mc:Ignorable="x" mc:PreserveAttributes="x:*" x:future="kept"><mc:AlternateContent><mc:Choice Requires="x"><x:dataType/></mc:Choice><mc:Fallback><w:viewMergedData/></mc:Fallback></mc:AlternateContent></w:mailMerge></w:settings>"#
    );
    let value = parse_settings_mail_merge(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(value.main_document_type(), MainDocumentType::FormLetters);
    assert_eq!(value.destination(), Destination::NewDocument);
    assert_eq!(value.active_record(), 1);
    assert_eq!(value.check_errors(), 2);
    assert!(value.view_merged_data());
}

#[test]
fn rejects_malformed_scoped_ordered_bounded_metadata() {
    let invalid = [
        format!(
            r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:dataType w:val="native"/><w:mainDocumentType w:val="email"/></w:mailMerge></w:settings>"#
        ),
        format!(
            r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:dataType w:val="bogus"/></w:mailMerge></w:settings>"#
        ),
        format!(r#"<w:settings xmlns:w="{W}"><w:trackRevisions/><w:mailMerge/></w:settings>"#),
        format!(r#"<w:settings xmlns:w="{W}"><w:zoom><w:mailMerge/></w:zoom></w:settings>"#),
        format!(
            r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:activeRecord w:val="0"/></w:mailMerge></w:settings>"#
        ),
        format!(
            r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:checkErrors w:val="2147483648"/></w:mailMerge></w:settings>"#
        ),
        format!(
            r#"<w:settings xmlns:w="{W}" xmlns:x="urn:fake"><w:mailMerge><w:dataSource x:id="rId1"/></w:mailMerge></w:settings>"#
        ),
        format!(r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:query/></w:mailMerge></w:settings>"#),
    ];
    for xml in invalid {
        assert!(
            parse_settings_mail_merge(xml.as_bytes()).is_err(),
            "accepted {xml}"
        );
    }
    let recipients = format!(
        r#"<w:recipients xmlns:w="{W}"><w:recipientData><w:uniqueTag w:val="AQ="/></w:recipientData></w:recipients>"#
    );
    assert!(Recipients::parse_xml(recipients.as_bytes()).is_err());
}

#[test]
fn rejects_mismatched_end_tags_and_oversized_attributes() {
    let mismatched = format!(r#"<w:settings xmlns:w="{W}"><w:mailMerge></w:settings>"#);
    assert!(parse_settings_mail_merge(mismatched.as_bytes()).is_err());

    let oversized_value = "x".repeat(MAX_STRING_BYTES + 1);
    let oversized = format!(
        r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:query w:val="{oversized_value}"/></w:mailMerge></w:settings>"#
    );
    assert!(matches!(
        parse_settings_mail_merge(oversized.as_bytes()),
        Err(Error::Invalid(message))
            if message == "mail-merge XML attribute is too large"
    ));
}

#[test]
fn round_trips_strict_recipient_data_with_canonical_base64() {
    let mut recipients = Recipients::new();
    recipients
        .add_recipient(Recipient::new(false, Some(7), Some(vec![1, 2, 3])))
        .unwrap();

    let xml = recipients.to_xml(Conformance::Strict).unwrap();
    assert!(xml.contains(STRICT_W));
    assert!(xml.contains("AQID"));
    assert_eq!(Recipients::parse_xml(xml.as_bytes()).unwrap(), recipients);
}
