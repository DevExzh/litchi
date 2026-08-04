//! Focused tests for the ODF metadata model and XML codec.

use super::*;
use litchi_core::Metadata as CoreMetadata;

#[test]
fn test_odf_metadata_default() {
    let meta = Metadata::default();
    assert!(meta.title.is_none());
    assert!(meta.description.is_none());
    assert!(meta.subject.is_none());
    assert!(meta.keywords.is_empty());
    assert!(meta.creator.is_none());
    assert!(meta.language.is_none());
    assert!(meta.creation_date.is_none());
    assert!(meta.modification_date.is_none());
    assert!(meta.generator.is_none());
    assert!(meta.custom_properties.is_empty());
}

#[test]
fn test_odf_metadata_from_xml_empty() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert!(meta.title.is_none());
    assert!(meta.creator.is_none());
}

#[test]
fn test_odf_metadata_from_xml_title() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
    <dc:title>Test Document</dc:title>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.title, Some("Test Document".to_string()));
}

#[test]
fn test_odf_metadata_from_xml_creator() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
    <dc:creator>John Doe</dc:creator>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.creator, Some("John Doe".to_string()));
}

#[test]
fn test_odf_metadata_from_xml_description() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
    <dc:description>This is a test document</dc:description>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(
        meta.description,
        Some("This is a test document".to_string())
    );
}

#[test]
fn test_odf_metadata_from_xml_subject() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
    <dc:subject>Testing</dc:subject>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.subject, Some("Testing".to_string()));
}

#[test]
fn test_odf_metadata_from_xml_keywords() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <meta:keyword>rust</meta:keyword>
    <meta:keyword>odf</meta:keyword>
    <meta:keyword>testing</meta:keyword>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.keywords, vec!["rust", "odf", "testing"]);
}

#[test]
fn test_odf_metadata_from_xml_language() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/">
<office:meta>
    <dc:language>en-US</dc:language>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.language, Some("en-US".to_string()));
}

#[test]
fn test_odf_metadata_from_xml_dates() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <meta:creation-date>2024-01-15T10:30:00Z</meta:creation-date>
    <dc:date>2024-03-20T14:45:00Z</dc:date>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.creation_date, Some("2024-01-15T10:30:00Z".to_string()));
    assert_eq!(
        meta.modification_date,
        Some("2024-03-20T14:45:00Z".to_string())
    );
}

#[test]
fn test_odf_metadata_from_xml_generator() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <meta:generator>LibreOffice/7.0</meta:generator>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.generator, Some("LibreOffice/7.0".to_string()));
}

#[test]
fn test_odf_metadata_from_xml_statistics() {
    // Note: The parser handles empty document-statistic elements
    // Statistics are parsed from attributes on the Start event
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <meta:document-statistic meta:page-count="5"
                             meta:paragraph-count="42"
                             meta:word-count="350"
                             meta:character-count="2100"
                             meta:table-count="3"
                             meta:image-count="2"
                             meta:object-count="1"></meta:document-statistic>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    // The statistics parsing happens on Start event with attributes
    assert_eq!(meta.statistics.page_count.as_deref(), Some("5"));
    assert_eq!(meta.statistics.paragraph_count.as_deref(), Some("42"));
    assert_eq!(meta.statistics.word_count.as_deref(), Some("350"));
    assert_eq!(meta.statistics.character_count.as_deref(), Some("2100"));
    assert_eq!(meta.statistics.table_count.as_deref(), Some("3"));
    assert_eq!(meta.statistics.image_count.as_deref(), Some("2"));
    assert_eq!(meta.statistics.object_count.as_deref(), Some("1"));
}

#[test]
fn test_odf_metadata_from_xml_user_defined() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <meta:user-defined meta:name="Department">Engineering</meta:user-defined>
    <meta:user-defined meta:name="Project">Alpha</meta:user-defined>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(
        meta.custom_properties.get("Department"),
        Some(&"Engineering".to_string())
    );
    assert_eq!(
        meta.custom_properties.get("Project"),
        Some(&"Alpha".to_string())
    );
}

#[test]
fn test_odf_metadata_from_xml_full() {
    let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:dc="http://purl.org/dc/elements/1.1/"
                  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
<office:meta>
    <dc:title>Full Test Document</dc:title>
    <dc:description>A comprehensive test</dc:description>
    <dc:subject>Testing</dc:subject>
    <dc:creator>Test Author</dc:creator>
    <dc:language>en</dc:language>
    <meta:creation-date>2024-01-01T00:00:00Z</meta:creation-date>
    <dc:date>2024-06-01T00:00:00Z</dc:date>
    <meta:generator>Test Generator</meta:generator>
    <meta:keyword>test</meta:keyword>
    <meta:document-statistic meta:page-count="10"></meta:document-statistic>
</office:meta>
</office:document-meta>"#;

    let meta = Metadata::from_xml(xml).unwrap();
    assert_eq!(meta.title, Some("Full Test Document".to_string()));
    assert_eq!(meta.description, Some("A comprehensive test".to_string()));
    assert_eq!(meta.subject, Some("Testing".to_string()));
    assert_eq!(meta.creator, Some("Test Author".to_string()));
    assert_eq!(meta.language, Some("en".to_string()));
    assert_eq!(meta.creation_date, Some("2024-01-01T00:00:00Z".to_string()));
    assert_eq!(
        meta.modification_date,
        Some("2024-06-01T00:00:00Z".to_string())
    );
    assert_eq!(meta.generator, Some("Test Generator".to_string()));
    assert_eq!(meta.keywords, vec!["test"]);
    assert_eq!(meta.statistics.page_count.as_deref(), Some("10"));
}

#[test]
fn parses_namespaces_entities_and_complete_metadata_without_annotation_leakage() {
    let xml = r#"<?xml version="1.0"?>
<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
        xmlns:d="http://purl.org/dc/elements/1.1/"
        xmlns:x="http://www.w3.org/1999/xlink">
  <o:meta>
<d:title>R&amp;D &#x1F34B;</d:title>
<d:creator>Last Editor</d:creator>
<m:initial-creator>Original Author</m:initial-creator>
<m:printed-by>Print Operator</m:printed-by>
<m:print-date>2025-04-03T02:01:00</m:print-date>
<m:editing-cycles>0000000000000000000000007</m:editing-cycles>
<m:editing-duration>PT1H2M3.0000000001S</m:editing-duration>
<m:template x:href="Templates/A&amp;B.ott" x:title="A &amp; B" m:date="2024-01-02T03:04:05" x:actuate="onRequest"/>
<m:auto-reload x:href="next.fodt" m:delay="PT5M" x:show="replace" x:actuate="onLoad"/>
<m:hyperlink-behaviour o:target-frame-name="_blank" x:show="new"/>
<m:document-statistic m:page-count="184467440737095516160"
  m:table-count="1" m:draw-count="2" m:image-count="3"
  m:ole-object-count="4" m:object-count="5" m:paragraph-count="6"
  m:word-count="7" m:character-count="8" m:frame-count="9"
  m:sentence-count="10" m:syllable-count="11"
  m:non-whitespace-character-count="12" m:row-count="13" m:cell-count="14"/>
<m:user-defined m:name="Flag" m:value-type="boolean">true</m:user-defined>
<m:user-defined m:name="Flag" m:value-type="string"><![CDATA[A&B]]></m:user-defined>
  </o:meta>
  <o:body><o:text><o:annotation><d:creator>Annotation Author</d:creator></o:annotation></o:text></o:body>
</o:document>"#;

    let metadata = Metadata::from_xml(xml).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("R&D 🍋"));
    assert_eq!(metadata.creator.as_deref(), Some("Last Editor"));
    assert_eq!(metadata.initial_creator.as_deref(), Some("Original Author"));
    assert_eq!(metadata.printed_by.as_deref(), Some("Print Operator"));
    assert_eq!(
        metadata.editing_cycles.as_deref(),
        Some("0000000000000000000000007")
    );
    assert_eq!(
        metadata.editing_duration.as_deref(),
        Some("PT1H2M3.0000000001S")
    );

    let template = metadata.template.as_ref().unwrap();
    assert_eq!(template.href.as_deref(), Some("Templates/A&B.ott"));
    assert_eq!(template.title.as_deref(), Some("A & B"));
    assert_eq!(template.actuate.as_deref(), Some("onRequest"));
    let auto_reload = metadata.auto_reload.as_ref().unwrap();
    assert_eq!(auto_reload.href.as_deref(), Some("next.fodt"));
    assert_eq!(auto_reload.delay.as_deref(), Some("PT5M"));
    let hyperlink = metadata.hyperlink_behaviour.as_ref().unwrap();
    assert_eq!(hyperlink.target_frame_name.as_deref(), Some("_blank"));
    assert_eq!(hyperlink.show.as_deref(), Some("new"));

    assert_eq!(
        metadata.statistics.page_count.as_deref(),
        Some("184467440737095516160")
    );
    assert_eq!(metadata.statistics.draw_count.as_deref(), Some("2"));
    assert_eq!(metadata.statistics.ole_object_count.as_deref(), Some("4"));
    assert_eq!(metadata.statistics.frame_count.as_deref(), Some("9"));
    assert_eq!(metadata.statistics.sentence_count.as_deref(), Some("10"));
    assert_eq!(metadata.statistics.syllable_count.as_deref(), Some("11"));
    assert_eq!(
        metadata
            .statistics
            .non_whitespace_character_count
            .as_deref(),
        Some("12")
    );
    assert_eq!(metadata.statistics.row_count.as_deref(), Some("13"));
    assert_eq!(metadata.statistics.cell_count.as_deref(), Some("14"));

    assert_eq!(metadata.user_defined.len(), 2);
    assert_eq!(
        metadata.user_defined[0].value_type,
        UserDefinedValueType::Boolean
    );
    assert_eq!(metadata.user_defined[0].value, "true");
    assert_eq!(
        metadata.user_defined[1].value_type,
        UserDefinedValueType::String
    );
    assert_eq!(metadata.user_defined[1].value, "A&B");
    assert_eq!(
        metadata.custom_properties.get("Flag").map(String::as_str),
        Some("A&B")
    );

    let common: CoreMetadata = metadata.into();
    assert_eq!(common.author.as_deref(), Some("Original Author"));
    assert_eq!(common.last_modified_by.as_deref(), Some("Last Editor"));
    assert_eq!(common.template.as_deref(), Some("Templates/A&B.ott"));
    assert_eq!(
        common.revision.as_deref(),
        Some("0000000000000000000000007")
    );
    assert_eq!(common.page_count, None);
    assert!(common.last_printed_time.is_some());
}

#[test]
fn rejects_invalid_statistic_and_nested_simple_metadata() {
    for xml in [
        r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><o:meta><m:document-statistic m:page-count="-1"/></o:meta></o:document-meta>"#,
        r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:meta><d:title>bad<d:subject>nested</d:subject></d:title></o:meta></o:document-meta>"#,
        r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><o:meta><m:user-defined>missing name</m:user-defined></o:meta></o:document-meta>"#,
    ] {
        assert!(Metadata::from_xml(xml).is_err(), "accepted {xml}");
    }
}

#[test]
fn test_document_statistics_default() {
    let stats = DocumentStatistics::default();
    assert!(stats.page_count.is_none());
    assert!(stats.paragraph_count.is_none());
    assert!(stats.word_count.is_none());
    assert!(stats.character_count.is_none());
    assert!(stats.table_count.is_none());
    assert!(stats.image_count.is_none());
    assert!(stats.object_count.is_none());
}

#[test]
fn test_parse_date_iso8601() {
    let date = Metadata::parse_date(Some("2024-03-15T14:30:00Z".to_string()));
    assert!(date.is_some());
}

#[test]
fn test_parse_date_rfc3339() {
    let date = Metadata::parse_date(Some("2024-03-15T00:00:00+00:00".to_string()));
    assert!(date.is_some());
}

#[test]
fn test_parse_date_none() {
    let date = Metadata::parse_date(None);
    assert!(date.is_none());
}

#[test]
fn test_parse_date_invalid() {
    let date = Metadata::parse_date(Some("not-a-date".to_string()));
    assert!(date.is_none());
}

#[test]
fn test_into_metadata_empty() {
    let odf = Metadata::default();
    let meta: CoreMetadata = odf.into();
    assert!(meta.title.is_none());
    assert!(meta.author.is_none());
    assert!(meta.keywords.is_none());
}

#[test]
fn test_into_metadata_with_data() {
    let odf = Metadata {
        title: Some("Title".to_string()),
        creator: Some("Author".to_string()),
        subject: Some("Subject".to_string()),
        keywords: vec!["a".to_string(), "b".to_string()],
        description: Some("Desc".to_string()),
        creation_date: Some("2024-01-01T00:00:00Z".to_string()),
        modification_date: Some("2024-06-01T00:00:00Z".to_string()),
        generator: Some("App".to_string()),
        statistics: DocumentStatistics {
            page_count: Some("5".to_string()),
            word_count: Some("100".to_string()),
            character_count: Some("500".to_string()),
            ..Default::default()
        },
        ..Default::default()
    };

    let meta: CoreMetadata = odf.into();
    assert_eq!(meta.title, Some("Title".to_string()));
    assert_eq!(meta.author, Some("Author".to_string()));
    assert_eq!(meta.subject, Some("Subject".to_string()));
    assert_eq!(meta.keywords, Some("a, b".to_string()));
    assert_eq!(meta.description, Some("Desc".to_string()));
    assert_eq!(meta.page_count, Some(5));
    assert_eq!(meta.word_count, Some(100));
    assert_eq!(meta.character_count, Some(500));
    assert_eq!(meta.application, Some("App".to_string()));
    assert!(meta.created.is_some());
    assert!(meta.modified.is_some());
}

#[test]
fn test_into_metadata_no_keywords() {
    let odf = Metadata {
        keywords: vec![],
        ..Default::default()
    };

    let meta: CoreMetadata = odf.into();
    assert!(meta.keywords.is_none());
}
