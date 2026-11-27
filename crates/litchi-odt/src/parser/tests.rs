//! Regression tests for the ODT-specific parser owner.

use super::*;

const TEST_TRACK_CHANGES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
    <text:tracked-changes>
        <text:changed-region text:id="change1">
            <text:insertion>
                <office:change-info>
                    <dc:creator>John Doe</dc:creator>
                    <dc:date>2024-03-15T10:30:00</dc:date>
                </office:change-info>
            </text:insertion>
        </text:changed-region>
        <text:changed-region text:id="change2">
            <text:deletion>
                <office:change-info>
                    <dc:creator>Jane Smith</dc:creator>
                    <dc:date>2024-03-15T11:00:00</dc:date>
                </office:change-info>
            </text:deletion>
        </text:changed-region>
        <text:changed-region text:id="change3">
            <text:format-change>
                <office:change-info>
                    <dc:creator>Bob Wilson</dc:creator>
                    <dc:date>2024-03-15T12:00:00</dc:date>
                </office:change-info>
            </text:format-change>
        </text:changed-region>
    </text:tracked-changes>
</office:document-content>"#;

const TEST_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
    <text:p>
        <office:annotation office:name="cmt1">
            <dc:creator>Alice</dc:creator>
            <dc:date>2024-03-15T09:00:00</dc:date>
            <text:p>This is a comment</text:p>
        </office:annotation>
        Some text
    </text:p>
    <text:p>
        <office:annotation office:name="cmt2">
            <dc:creator>Bob</dc:creator>
            <dc:date>2024-03-15T10:00:00</dc:date>
            <text:p>First paragraph</text:p>
            <text:p>Second paragraph</text:p>
        </office:annotation>
        More text
    </text:p>
</office:document-content>"#;

const TEST_SECTIONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <text:section text:name="Introduction" text:style-name="IntroStyle">
        <text:p>Introduction content</text:p>
    </text:section>
    <text:section text:name="ProtectedSection" text:protected="true">
        <text:p>Protected content</text:p>
    </text:section>
    <text:section text:name="Chapter1" text:style-name="ChapterStyle" text:protected="false">
        <text:p>Chapter 1 content</text:p>
    </text:section>
</office:document-content>"#;

const TEST_EMPTY_TRACK_CHANGES: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <text:tracked-changes>
    </text:tracked-changes>
</office:document-content>"#;

const TEST_EMPTY_CONTENT: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
</office:document-content>"#;

#[test]
fn test_parse_track_changes() {
    let changes = Parser::parse_track_changes(TEST_TRACK_CHANGES_XML).unwrap();
    assert_eq!(changes.len(), 3);

    assert_eq!(changes[0].id, "change1");
    assert_eq!(changes[0].change_type, ChangeType::Insertion);
    assert_eq!(changes[0].author.as_deref(), Some("John Doe"));
    assert_eq!(changes[0].date.as_deref(), Some("2024-03-15T10:30:00"));
    assert!(changes[0].content.is_empty());

    assert_eq!(changes[1].id, "change2");
    assert_eq!(changes[1].change_type, ChangeType::Deletion);
    assert_eq!(changes[1].author.as_deref(), Some("Jane Smith"));
    assert_eq!(changes[1].date.as_deref(), Some("2024-03-15T11:00:00"));
    assert!(changes[1].content.is_empty());

    assert_eq!(changes[2].id, "change3");
    assert_eq!(changes[2].change_type, ChangeType::FormatChange);
    assert_eq!(changes[2].author.as_deref(), Some("Bob Wilson"));
    assert_eq!(changes[2].date.as_deref(), Some("2024-03-15T12:00:00"));
    assert!(changes[2].content.is_empty());
}

#[test]
fn test_parse_track_changes_empty() {
    let changes = Parser::parse_track_changes(TEST_EMPTY_TRACK_CHANGES).unwrap();
    assert!(changes.is_empty());
}

#[test]
fn test_parse_track_changes_no_tracked_changes() {
    let changes = Parser::parse_track_changes(TEST_EMPTY_CONTENT).unwrap();
    assert!(changes.is_empty());
}

#[test]
fn parses_tracked_change_metadata_deletions_and_referenced_ranges() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text><t:tracked-changes><t:changed-region t:id="i1"><t:insertion><o:change-info><d:creator>A &amp; B</d:creator><d:date>2026-07-16T10:00:00</d:date><t:p>review note</t:p></o:change-info></t:insertion></t:changed-region><t:changed-region xml:id="d1"><t:deletion><o:change-info><d:creator>Deleter</d:creator><d:date>2026-07-16</d:date><t:p>not deleted text</t:p></o:change-info><t:p>Gone &amp;<t:s t:c="2"/><t:span><![CDATA[X]]></t:span></t:p><t:p>Second<t:tab/></t:p></t:deletion></t:changed-region><t:changed-region t:id="f1"><t:format-change><o:change-info><d:creator>Stylist</d:creator><d:date>2026-07-15</d:date></o:change-info></t:format-change></t:changed-region></t:tracked-changes><t:p>pre<t:change-start t:change-id="i1"/>In&amp;<o:annotation o:name="note"><t:p>hidden comment</t:p></o:annotation><t:span>sert</t:span><t:s t:c="2"/><![CDATA[!]]><t:change-end t:change-id="i1"/>post<t:change t:change-id="d1"/></t:p><t:p><t:change-start t:change-id="i1"/>Again<t:change-end t:change-id="i1"/> and <t:change-start t:change-id="f1"/>Bold<t:change-end t:change-id="f1"/></t:p></o:text></o:body></o:document-content>"#;
    let changes = Parser::parse_track_changes(xml).unwrap();
    assert_eq!(changes.len(), 3);
    assert_eq!(changes[0].id, "i1");
    assert_eq!(changes[0].author.as_deref(), Some("A & B"));
    assert_eq!(changes[0].date.as_deref(), Some("2026-07-16T10:00:00"));
    assert_eq!(changes[0].content, "In&sert  !\nAgain");
    assert_eq!(changes[1].id, "d1");
    assert_eq!(changes[1].change_type, ChangeType::Deletion);
    assert_eq!(changes[1].content, "Gone &  X\nSecond\t");
    assert_eq!(changes[2].id, "f1");
    assert_eq!(changes[2].change_type, ChangeType::FormatChange);
    assert_eq!(changes[2].content, "Bold");
}

#[test]
fn retains_tracked_change_policy_and_schema_attributes() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text><t:tracked-changes t:track-changes="0" t:protection-key="YWJj" t:protection-key-digest-algorithm="urn:example:sha256"><t:changed-region t:id="d1" xml:id="d1"><t:deletion t:merge-last-paragraph="false"><o:change-info><d:creator>A</d:creator><d:date>2026-07-16</d:date></o:change-info><t:p>gone</t:p></t:deletion></t:changed-region><t:changed-region t:id="f1"><t:format-change t:style-name="Emphasis"><o:change-info><d:creator>B</d:creator><d:date>2026-07-16</d:date></o:change-info></t:format-change></t:changed-region></t:tracked-changes></o:text></o:body></o:document-content>"#;
    let tracked = Parser::parse_tracked_changes(xml).unwrap();
    assert_eq!(tracked.track_changes, Some(false));
    assert_eq!(tracked.protection_key.as_deref(), Some("YWJj"));
    assert_eq!(
        tracked.protection_key_digest_algorithm.as_deref(),
        Some("urn:example:sha256")
    );
    assert_eq!(tracked.changes[0].xml_id.as_deref(), Some("d1"));
    assert_eq!(tracked.changes[0].merge_last_paragraph, Some(false));
    assert_eq!(tracked.changes[1].style_name.as_deref(), Some("Emphasis"));
}

#[test]
fn rejects_invalid_tracked_change_policy_and_attributes() {
    let prefix = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text>"#;
    let suffix = "</o:text></o:body></o:document-content>";
    for body in [
        r#"<t:tracked-changes t:track-changes="yes"/>"#,
        r#"<t:tracked-changes t:protection-key="not-base64"/>"#,
        r#"<t:tracked-changes t:protection-key-digest-algorithm="urn:sha256"/>"#,
        r#"<t:tracked-changes><t:changed-region t:id="a" xml:id="b"><t:insertion><o:change-info/></t:insertion></t:changed-region></t:tracked-changes>"#,
        r#"<t:tracked-changes><t:changed-region t:id="a"><t:deletion t:merge-last-paragraph="maybe"><o:change-info/></t:deletion></t:changed-region></t:tracked-changes>"#,
    ] {
        let xml = format!("{prefix}{body}{suffix}");
        assert!(
            Parser::parse_tracked_changes(&xml).is_err(),
            "accepted {body}"
        );
    }
}

#[test]
fn tracked_changes_reject_ambiguous_declarations_and_ranges() {
    let prelude = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:u="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text>"#;
    let info = r"<o:change-info><d:creator>A</d:creator><d:date>D</d:date></o:change-info>";
    let suffix = "</o:text></o:body></o:document-content>";

    let missing_id = format!(
        "{prelude}<t:tracked-changes><t:changed-region><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes>{suffix}"
    );
    assert!(Parser::parse_track_changes(&missing_id).is_err());

    let duplicate_id = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region><t:changed-region t:id=\"x\"><t:deletion>{info}</t:deletion></t:changed-region></t:tracked-changes>{suffix}"
    );
    assert!(Parser::parse_track_changes(&duplicate_id).is_err());

    let multiple_kinds = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion><t:deletion>{info}</t:deletion></t:changed-region></t:tracked-changes>{suffix}"
    );
    assert!(Parser::parse_track_changes(&multiple_kinds).is_err());

    let missing_kind = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"/></t:tracked-changes>{suffix}"
    );
    assert!(Parser::parse_track_changes(&missing_kind).is_err());

    let unknown_marker = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change t:change-id=\"unknown\"/></t:p>{suffix}"
    );
    assert!(Parser::parse_track_changes(&unknown_marker).is_err());

    let unmatched_end = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change-end t:change-id=\"x\"/></t:p>{suffix}"
    );
    assert!(Parser::parse_track_changes(&unmatched_end).is_err());

    let unmatched_start = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change-start t:change-id=\"x\"/>open</t:p>{suffix}"
    );
    assert!(Parser::parse_track_changes(&unmatched_start).is_err());

    let duplicate_attribute = format!(
        "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change t:change-id=\"x\" u:change-id=\"x\"/></t:p>{suffix}"
    );
    assert!(Parser::parse_track_changes(&duplicate_attribute).is_err());
    assert!(Parser::parse_track_changes("<t:tracked-changes>").is_err());
}

#[test]
fn tracked_changes_enforce_nesting_bound() {
    let mut xml = String::from(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text><t:tracked-changes><t:changed-region t:id="d"><t:deletion><o:change-info><d:creator>A</d:creator><d:date>D</d:date></o:change-info><t:p>"#,
    );
    for _ in 0..MAX_SEMANTIC_DEPTH {
        xml.push_str("<t:span>");
    }
    for _ in 0..MAX_SEMANTIC_DEPTH {
        xml.push_str("</t:span>");
    }
    xml.push_str(
        "</t:p></t:deletion></t:changed-region></t:tracked-changes></o:text></o:body></o:document-content>",
    );
    assert!(Parser::parse_track_changes(&xml).is_err());
}

#[test]
fn test_parse_comments() {
    let comments = Parser::parse_comments(TEST_COMMENTS_XML).unwrap();
    assert_eq!(comments.len(), 2);

    // First comment
    assert_eq!(comments[0].id, "cmt1");
    assert_eq!(comments[0].author, Some("Alice".to_string()));
    assert_eq!(comments[0].date, Some("2024-03-15T09:00:00".to_string()));
    assert_eq!(comments[0].content, "This is a comment");

    // Second comment (with multiple paragraphs)
    assert_eq!(comments[1].id, "cmt2");
    assert_eq!(comments[1].author, Some("Bob".to_string()));
    assert_eq!(comments[1].date, Some("2024-03-15T10:00:00".to_string()));
    assert!(comments[1].content.contains("First paragraph"));
    assert!(comments[1].content.contains("Second paragraph"));
}

#[test]
fn test_parse_comments_empty() {
    let comments = Parser::parse_comments(TEST_EMPTY_CONTENT).unwrap();
    assert!(comments.is_empty());
}

#[test]
fn test_parse_sections() {
    let sections = Parser::parse_sections(TEST_SECTIONS_XML).unwrap();
    assert_eq!(sections.len(), 3);

    // First section
    assert_eq!(sections[0].name, "Introduction");
    assert_eq!(sections[0].style, Some("IntroStyle".to_string()));
    assert!(!sections[0].protected);

    // Second section (protected)
    assert_eq!(sections[1].name, "ProtectedSection");
    assert_eq!(sections[1].style, None);
    assert!(sections[1].protected);

    // Third section
    assert_eq!(sections[2].name, "Chapter1");
    assert_eq!(sections[2].style, Some("ChapterStyle".to_string()));
    assert!(!sections[2].protected);
}

#[test]
fn test_parse_sections_empty() {
    let sections = Parser::parse_sections(TEST_EMPTY_CONTENT).unwrap();
    assert!(sections.is_empty());
}

#[test]
fn parses_annotation_metadata_body_and_referenced_range_with_namespace_aliases() {
    let xml = r#"<x:document-content xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><x:body><x:text><t:p>before<x:annotation x:name="c&amp;1"><d:creator>A &amp; B</d:creator><m:date-string>2026-07-16</m:date-string><t:p>First<t:s t:c="2"/>X</t:p><t:list><t:list-item><t:p>Second<![CDATA[!]]></t:p></t:list-item></t:list></x:annotation>R&amp;<t:span>ange</t:span><x:annotation-end x:name="c&amp;1"/>after</t:p></x:text></x:body></x:document-content>"#;
    let comments = Parser::parse_comments(xml).unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].id, "c&1");
    assert_eq!(comments[0].author.as_deref(), Some("A & B"));
    assert_eq!(comments[0].date.as_deref(), Some("2026-07-16"));
    assert_eq!(comments[0].content, "First  X\nSecond!");
    assert_eq!(comments[0].reference.as_deref(), Some("R&ange"));
}

#[test]
fn parses_nested_sections_in_document_order_with_visible_text() {
    let xml = r#"<x:document-content xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:l="http://www.w3.org/1999/xlink"><x:body><x:text><t:section t:name="Outer &amp; Main" t:style-name="S1" t:protected="1" xml:id="outer" t:protection-key="YWJj" t:protection-key-digest-algorithm="urn:sha256" t:display="condition" t:condition="ooow:visible()"><t:section-source l:type="simple" l:href="https://example.invalid/doc.odt" l:show="embed" t:section-name="Remote" t:filter-name="writer8"/><t:p>One &amp;<t:s t:c="2"/></t:p><t:section t:name="Inner"><t:p>Inner <![CDATA[X]]></t:p></t:section><t:p>Last</t:p></t:section><t:section t:name="Empty"><x:dde-source x:name="Feed" x:conversion-mode="keep-text" x:automatic-update="false"/></t:section></x:text></x:body></x:document-content>"#;
    let sections = Parser::parse_sections(xml).unwrap();
    assert_eq!(sections.len(), 3);
    assert_eq!(sections[0].name, "Outer & Main");
    assert_eq!(sections[0].style.as_deref(), Some("S1"));
    assert!(sections[0].protected);
    assert_eq!(sections[0].xml_id.as_deref(), Some("outer"));
    assert_eq!(sections[0].protection_key.as_deref(), Some("YWJj"));
    assert_eq!(
        sections[0].protection_key_digest_algorithm.as_deref(),
        Some("urn:sha256")
    );
    assert_eq!(sections[0].display, SectionDisplay::Condition);
    assert_eq!(sections[0].condition.as_deref(), Some("ooow:visible()"));
    let source = sections[0].source.as_ref().unwrap();
    assert_eq!(
        source.href.as_deref(),
        Some("https://example.invalid/doc.odt")
    );
    assert_eq!(source.section_name.as_deref(), Some("Remote"));
    assert_eq!(source.filter_name.as_deref(), Some("writer8"));
    assert_eq!(sections[0].content, "One &  \nInner X\nLast");
    assert_eq!(sections[1].name, "Inner");
    assert_eq!(sections[1].content, "Inner X");
    assert_eq!(sections[2].name, "Empty");
    assert!(sections[2].content.is_empty());
    let dde = sections[2].dde_source.as_ref().unwrap();
    assert_eq!(dde.name.as_deref(), Some("Feed"));
    assert_eq!(dde.conversion_mode.as_deref(), Some("keep-text"));
    assert_eq!(dde.automatic_update, Some(false));
}

#[test]
fn annotations_and_sections_reject_malformed_or_ambiguous_xml() {
    let namespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    assert!(Parser::parse_comments("<x:annotation>").is_err());
    let nested = r#"<x:annotation xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><x:annotation/></x:annotation>"#;
    assert!(Parser::parse_comments(nested).is_err());
    let missing_name = format!(r#"<t:section xmlns:t="{namespace}"/>"#);
    assert!(Parser::parse_sections(&missing_name).is_err());
    let invalid_boolean =
        format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:protected="yes"/>"#);
    assert!(Parser::parse_sections(&invalid_boolean).is_err());
    let duplicate = format!(
        r#"<t:section xmlns:t="{namespace}" xmlns:u="{namespace}" t:name="A" u:name="B"/>"#
    );
    assert!(Parser::parse_sections(&duplicate).is_err());
    let missing_condition =
        format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:display="condition"/>"#);
    assert!(Parser::parse_sections(&missing_condition).is_err());
    let stray_condition =
        format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:condition="x"/>"#);
    assert!(Parser::parse_sections(&stray_condition).is_err());
    let duplicate_source = format!(
        r#"<t:section xmlns:t="{namespace}" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" t:name="A"><t:section-source/><o:dde-source/></t:section>"#
    );
    assert!(Parser::parse_sections(&duplicate_source).is_err());
    let nonempty_source = format!(
        r#"<t:section xmlns:t="{namespace}" t:name="A"><t:section-source>bad</t:section-source></t:section>"#
    );
    assert!(Parser::parse_sections(&nonempty_source).is_err());
    let incomplete_link = format!(
        r#"<t:section xmlns:t="{namespace}" xmlns:l="http://www.w3.org/1999/xlink" t:name="A"><t:section-source l:href="x"/></t:section>"#
    );
    assert!(Parser::parse_sections(&incomplete_link).is_err());
}

#[test]
fn test_track_change_debug() {
    let change = TrackChange {
        id: "test1".to_string(),
        xml_id: None,
        author: Some("Author".to_string()),
        date: Some("2024-03-15".to_string()),
        comment: None,
        change_type: ChangeType::Insertion,
        style_name: None,
        merge_last_paragraph: None,
        content: "content".to_string(),
    };
    let debug_str = format!("{change:?}");
    assert!(debug_str.contains("TrackChange"));
    assert!(debug_str.contains("test1"));
}

#[test]
fn test_change_type_equality() {
    assert_eq!(ChangeType::Insertion, ChangeType::Insertion);
    assert_eq!(ChangeType::Deletion, ChangeType::Deletion);
    assert_eq!(ChangeType::FormatChange, ChangeType::FormatChange);
    assert_ne!(ChangeType::Insertion, ChangeType::Deletion);
}

#[test]
fn test_change_type_clone() {
    let t1 = ChangeType::Insertion;
    let t2 = t1;
    assert_eq!(t1, t2);
}

#[test]
fn test_change_type_copy() {
    let t1 = ChangeType::Insertion;
    let t2 = t1;
    assert_eq!(t1, t2); // Copy trait allows this
}

#[test]
fn test_comment_debug() {
    let comment = Comment {
        id: "cmt1".to_string(),
        author: Some("Author".to_string()),
        date: Some("2024-03-15".to_string()),
        content: "Comment text".to_string(),
        reference: None,
    };
    let debug_str = format!("{comment:?}");
    assert!(debug_str.contains("Comment"));
    assert!(debug_str.contains("cmt1"));
}

#[test]
fn test_section_debug() {
    let section = Section {
        name: "Sec1".to_string(),
        style: Some("Style1".to_string()),
        protected: true,
        xml_id: None,
        protection_key: None,
        protection_key_digest_algorithm: None,
        display: SectionDisplay::Visible,
        condition: None,
        source: None,
        dde_source: None,
        content: "Content".to_string(),
    };
    let debug_str = format!("{section:?}");
    assert!(debug_str.contains("Section"));
    assert!(debug_str.contains("Sec1"));
}

#[test]
fn test_comment_clone() {
    let comment = Comment {
        id: "cmt1".to_string(),
        author: Some("Author".to_string()),
        date: Some("2024-03-15".to_string()),
        content: "Content".to_string(),
        reference: Some("ref".to_string()),
    };
    let cloned = comment.clone();
    assert_eq!(comment.id, cloned.id);
    assert_eq!(comment.author, cloned.author);
    assert_eq!(comment.content, cloned.content);
}

#[test]
fn test_track_change_clone() {
    let change = TrackChange {
        id: "tc1".to_string(),
        xml_id: None,
        author: Some("Author".to_string()),
        date: Some("2024-03-15".to_string()),
        comment: None,
        change_type: ChangeType::Deletion,
        style_name: None,
        merge_last_paragraph: None,
        content: "Deleted text".to_string(),
    };
    let cloned = change.clone();
    assert_eq!(change.id, cloned.id);
    assert_eq!(change.change_type, cloned.change_type);
}

#[test]
fn test_section_clone() {
    let section = Section {
        name: "Sec1".to_string(),
        style: None,
        protected: false,
        xml_id: None,
        protection_key: None,
        protection_key_digest_algorithm: None,
        display: SectionDisplay::Visible,
        condition: None,
        source: None,
        dde_source: None,
        content: "Text".to_string(),
    };
    let cloned = section.clone();
    assert_eq!(section.name, cloned.name);
    assert_eq!(section.protected, cloned.protected);
}
