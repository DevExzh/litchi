use super::{Limits, RevisionType, parse_revisions_with_limits};
use crate::Error;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

fn document(body: &str, namespace: &str) -> String {
    format!(r#"<w:document xmlns:w="{namespace}"><w:body>{body}</w:body></w:document>"#)
}

#[test]
fn property_revisions_project_normative_metadata_in_both_dialects() {
    for namespace in [WORD, STRICT] {
        let xml = document(
            r#"<w:p><w:pPr><w:sectPr><w:sectPrChange w:id="1" w:author="A"><w:sectPr><w:pgSz w:w="11906"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/><w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="200"/></w:tblGrid></w:tblGridChange></w:tblGrid><w:tr><w:tblPrEx><w:tblPrExChange w:id="3" w:author=""><w:tblPrEx/></w:tblPrExChange></w:tblPrEx><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:sectPrChange w:id="4" w:author="B"/></w:sectPr>"#,
            namespace,
        );
        let records = parse_revisions_with_limits(xml.as_bytes(), &[], Limits::default()).unwrap();
        assert_eq!(records.len(), 4);
        assert_eq!(
            records
                .iter()
                .map(|r| r.revision_type())
                .collect::<Vec<_>>(),
            [
                RevisionType::SectionPropertiesChange,
                RevisionType::TableGridChange,
                RevisionType::TablePropertyExceptionsChange,
                RevisionType::SectionPropertiesChange,
            ]
        );
        assert_eq!(records[0].author(), Some("A"));
        assert_eq!(records[1].author(), None);
        assert_eq!(records[1].date(), None);
        assert_eq!(records[2].author(), Some(""));
        assert!(records.iter().all(|r| r.text().is_empty()));
    }
}

#[test]
fn property_history_requires_normative_owner_and_snapshot_cardinality() {
    for body in [
        r#"<w:p><w:sectPrChange w:id="1" w:author="A"/></w:p>"#,
        r#"<w:tbl><w:tblGridChange w:id="1"><w:tblGrid/></w:tblGridChange></w:tbl>"#,
        r#"<w:tblGrid><w:tblGridChange w:id="1"/></w:tblGrid>"#,
        r#"<w:tblGrid><w:tblGridChange w:id="1"></w:tblGridChange></w:tblGrid>"#,
        r#"<w:tblPrEx><w:tblPrExChange w:id="1" w:author="A"/></w:tblPrEx>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><w:sectPr/><w:sectPr/></w:sectPrChange></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><w:tblGrid/></w:sectPrChange></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A">text</w:sectPrChange></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A">&#65;</w:sectPrChange></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><![CDATA[text]]></w:sectPrChange></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><w:sectPr><w:sectPrChange w:id="2" w:author="B"/></w:sectPr></w:sectPrChange></w:sectPr>"#,
        r#"<w:tblGrid><w:tblGridChange><w:tblGrid/></w:tblGridChange></w:tblGrid>"#,
        r#"<w:tblGrid><w:tblGridChange w:id="1" w:author="A"><w:tblGrid/></w:tblGridChange></w:tblGrid>"#,
    ] {
        let xml = document(body, WORD);
        assert!(
            parse_revisions_with_limits(xml.as_bytes(), &[], Limits::default()).is_err(),
            "{body}"
        );
    }
}

#[test]
fn foreign_property_names_do_not_satisfy_revision_snapshot_or_owner() {
    for body in [
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><x:sectPr xmlns:x="urn:foreign"/></w:sectPrChange></w:sectPr>"#,
        r#"<x:sectPr xmlns:x="urn:foreign"><w:sectPrChange w:id="1" w:author="A"/></x:sectPr>"#,
    ] {
        assert!(
            parse_revisions_with_limits(document(body, WORD).as_bytes(), &[], Limits::default())
                .is_err()
        );
    }
    let xml = document(
        r#"<w:sectPr><x:sectPrChange xmlns:x="urn:foreign"/></w:sectPr>"#,
        WORD,
    );
    assert!(
        parse_revisions_with_limits(xml.as_bytes(), &[], Limits::default())
            .unwrap()
            .is_empty()
    );
}

#[test]
fn transitional_numbering_history_retains_original_for_paragraphs_and_fields() {
    let body = r#"<w:p><w:pPr><w:numPr><w:numberingChange w:id="1" w:author="A" w:original="&lt;%1:0:longer than fifteen characters&gt;"/><w:ins w:id="2" w:author="B"/></w:numPr></w:pPr><w:r><w:fldChar w:fldCharType="begin"><w:numberingChange w:id="3" w:author="C" w:original=""/></w:fldChar></w:r></w:p>"#;
    let records =
        parse_revisions_with_limits(document(body, WORD).as_bytes(), &[], Limits::default())
            .unwrap();
    assert_eq!(records.len(), 3);
    assert_eq!(records[0].revision_type(), RevisionType::NumberingChange);
    assert_eq!(
        records[0].original_numbering(),
        Some("<%1:0:longer than fifteen characters>")
    );
    assert_eq!(records[1].original_numbering(), None);
    assert_eq!(records[2].original_numbering(), Some(""));
    assert!(
        parse_revisions_with_limits(document(body, STRICT).as_bytes(), &[], Limits::default())
            .is_err()
    );
    for body in [
        r#"<w:p><w:numberingChange w:id="1" w:author="A"/></w:p>"#,
        r#"<w:numPr><w:numberingChange w:id="1" w:author="A"><w:ins w:id="2" w:author="B"/></w:numberingChange></w:numPr>"#,
        r#"<w:numPr><w:numberingChange w:id="1" w:author="A" w:original="&unknown;"/></w:numPr>"#,
    ] {
        assert!(
            parse_revisions_with_limits(document(body, WORD).as_bytes(), &[], Limits::default())
                .is_err()
        );
    }
}

#[test]
fn numbering_original_and_grid_identifiers_are_charged_to_metadata_limits() {
    let xml = document(
        r#"<w:p><w:pPr><w:numPr><w:numberingChange w:id="1" w:author="A" w:original="&amp;xyz"/></w:numPr></w:pPr></w:p><w:tblGrid><w:tblGridChange w:id="2"><w:tblGrid/></w:tblGridChange></w:tblGrid>"#,
        WORD,
    );
    let limits = Limits {
        max_metadata_bytes: 7,
        max_revisions: 2,
        ..Limits::default()
    };
    assert_eq!(
        parse_revisions_with_limits(xml.as_bytes(), &[], limits)
            .unwrap()
            .len(),
        2
    );
    assert!(matches!(
        parse_revisions_with_limits(
            xml.as_bytes(),
            &[],
            Limits {
                max_metadata_bytes: 6,
                ..limits
            }
        ),
        Err(Error::RevisionLimit {
            resource: "metadata bytes",
            ..
        })
    ));
    assert!(matches!(
        parse_revisions_with_limits(
            xml.as_bytes(),
            &[],
            Limits {
                max_revisions: 1,
                ..limits
            }
        ),
        Err(Error::RevisionLimit {
            resource: "records",
            ..
        })
    ));
    assert!(matches!(
        parse_revisions_with_limits(
            xml.as_bytes(),
            &[],
            Limits {
                max_value_bytes: 7,
                ..limits
            }
        ),
        Err(Error::RevisionLimit {
            resource: "value bytes",
            ..
        })
    ));
}

#[test]
fn original_properties_reject_character_data_and_nested_revision_markers() {
    for content in [
        "text",
        "&#65;",
        "<![CDATA[text]]>",
        r#"<w:pPrChange w:id="2" w:author="B"><w:pPr/></w:pPrChange>"#,
        r#"<w:tblPrChange w:id="2" w:author="B"><w:tblPr/></w:tblPrChange>"#,
        r#"<w:pgSz><w:ins w:id="2" w:author="B"/></w:pgSz>"#,
        r#"<x:wrapper xmlns:x="urn:foreign"><w:ins w:id="2" w:author="B"/></x:wrapper>"#,
    ] {
        let body = format!(
            r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"><w:sectPr>{content}</w:sectPr></w:sectPrChange></w:sectPr>"#
        );
        assert!(
            parse_revisions_with_limits(document(&body, WORD).as_bytes(), &[], Limits::default())
                .is_err(),
            "{content}"
        );
    }
}

#[test]
fn property_and_numbering_owners_reject_duplicate_or_out_of_order_history() {
    for body in [
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"/><w:sectPrChange w:id="2" w:author="B"/></w:sectPr>"#,
        r#"<w:sectPr><w:sectPrChange w:id="1" w:author="A"/><w:pgSz/></w:sectPr>"#,
        r#"<w:tblGrid><w:tblGridChange w:id="1"><w:tblGrid/></w:tblGridChange><w:gridCol/></w:tblGrid>"#,
        r#"<w:tblPrEx><w:tblPrExChange w:id="1" w:author="A"><w:tblPrEx/></w:tblPrExChange><w:tblW/></w:tblPrEx>"#,
        r#"<w:numPr><w:numberingChange w:id="1" w:author="A"/><w:numberingChange w:id="2" w:author="B"/></w:numPr>"#,
        r#"<w:numPr><w:ins w:id="1" w:author="A"/><w:numberingChange w:id="2" w:author="B"/></w:numPr>"#,
        r#"<w:numPr><w:numberingChange w:id="1" w:author="A"/><w:numId w:val="1"/></w:numPr>"#,
        r#"<w:fldChar w:fldCharType="begin"><w:fldData/><w:numberingChange w:id="1" w:author="A"/></w:fldChar>"#,
    ] {
        assert!(
            parse_revisions_with_limits(document(body, WORD).as_bytes(), &[], Limits::default())
                .is_err(),
            "{body}"
        );
    }
}
