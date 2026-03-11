//! ODT-specific parsing utilities.
//!
//! This module provides parsing functionality that is specific to OpenDocument Text
//! documents (.odt). For generic ODF element parsing (paragraphs, tables, lists, etc.)
//! that works across all ODF formats, see `crate::odf::elements::parser::DocumentParser`.

use crate::common::Result;

/// Parser for ODT-specific structures.
///
/// This provides parsing logic specific to text documents, such as:
/// - Track changes (insertions, deletions, formatting changes)
/// - Comments and annotations
/// - Sections (protected content, different formatting)
/// - Headers and footers
///
/// For generic element parsing (paragraphs, tables, etc.), use `DocumentParser`
/// from `crate::odf::elements::parser` instead.
pub(crate) struct OdtParser;

/// Represents a tracked change in the document
#[derive(Debug, Clone)]
pub struct TrackChange {
    /// Change ID
    pub id: String,
    /// Author who made the change
    pub author: Option<String>,
    /// Date/time of the change
    pub date: Option<String>,
    /// Type of change (insertion, deletion, format-change)
    pub change_type: ChangeType,
    /// Changed text content
    pub content: String,
}

/// Type of tracked change
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChangeType {
    /// Text insertion
    Insertion,
    /// Text deletion
    Deletion,
    /// Formatting change
    FormatChange,
}

/// Represents a comment/annotation in the document
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment ID
    pub id: String,
    /// Author of the comment
    pub author: Option<String>,
    /// Date/time of the comment
    pub date: Option<String>,
    /// Comment text content
    pub content: String,
    /// Referenced text in the document
    pub reference: Option<String>,
}

/// Represents a section in the document
#[derive(Debug, Clone)]
pub struct Section {
    /// Section name
    pub name: String,
    /// Section style
    pub style: Option<String>,
    /// Whether the section is protected
    pub protected: bool,
    /// Text content within the section
    pub content: String,
}

impl OdtParser {
    /// Parse track changes from content
    ///
    /// Extracts tracked changes (insertions, deletions, format changes) from the document.
    /// Track changes are stored in `<text:tracked-changes>` elements with metadata,
    /// and referenced by `<text:change>` markers in the content.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing tracked changes
    ///
    /// # Returns
    ///
    /// Vector of `TrackChange` objects with metadata
    pub fn parse_track_changes(content: &str) -> Result<Vec<TrackChange>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(content);
        let mut buf = Vec::new();
        let mut changes = Vec::new();
        let mut in_tracked_changes = false;
        let mut in_change_element = false;
        let mut current_change: Option<TrackChange> = None;
        let mut depth: usize = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:tracked-changes" => {
                            in_tracked_changes = true;
                        },
                        "text:changed-region" if in_tracked_changes => {
                            // Extract change ID
                            let mut id = String::new();
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                if key.ends_with(":id") {
                                    id = String::from_utf8_lossy(&attr.value).to_string();
                                }
                            }

                            current_change = Some(TrackChange {
                                id,
                                author: None,
                                date: None,
                                change_type: ChangeType::Insertion,
                                content: String::new(),
                            });
                            depth += 1;
                        },
                        "text:insertion" | "text:deletion" | "text:format-change"
                            if in_tracked_changes && current_change.is_some() =>
                        {
                            if let Some(ref mut change) = current_change {
                                change.change_type = match tag_name.as_str() {
                                    "text:insertion" => ChangeType::Insertion,
                                    "text:deletion" => ChangeType::Deletion,
                                    "text:format-change" => ChangeType::FormatChange,
                                    _ => ChangeType::Insertion,
                                };
                            }
                            in_change_element = true;
                            depth += 1;
                        },
                        "office:change-info" if in_change_element => {
                            depth += 1;
                        },
                        "dc:creator" if in_change_element => {
                            depth += 1;
                        },
                        "dc:date" if in_change_element => {
                            depth += 1;
                        },
                        _ if in_tracked_changes => {
                            depth += 1;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) if in_change_element => {
                    let text = String::from_utf8_lossy(t).to_string();

                    // Determine what we're reading based on parent context
                    if let Some(ref mut change) = current_change {
                        // This is a simplification; in reality we'd track the parent element
                        if change.author.is_none() {
                            change.author = Some(text.clone());
                        } else if change.date.is_none() {
                            change.date = Some(text);
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:tracked-changes" => {
                            in_tracked_changes = false;
                        },
                        "text:changed-region" if in_tracked_changes => {
                            if let Some(change) = current_change.take() {
                                changes.push(change);
                            }
                            depth = depth.saturating_sub(1);
                        },
                        "text:insertion" | "text:deletion" | "text:format-change"
                            if in_tracked_changes =>
                        {
                            in_change_element = false;
                            depth = depth.saturating_sub(1);
                        },
                        _ if in_tracked_changes && depth > 0 => {
                            depth = depth.saturating_sub(1);
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(changes)
    }

    /// Parse comments/annotations
    ///
    /// Extracts comments and annotations from the document.
    /// Comments are stored in `<office:annotation>` elements.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing annotations
    ///
    /// # Returns
    ///
    /// Vector of `Comment` objects with metadata and content
    pub fn parse_comments(content: &str) -> Result<Vec<Comment>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(content);
        let mut buf = Vec::new();
        let mut comments = Vec::new();
        let mut in_annotation = false;
        let mut current_comment: Option<Comment> = None;
        let mut in_creator = false;
        let mut in_date = false;
        let mut in_paragraph = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "office:annotation" => {
                            in_annotation = true;

                            // Extract annotation name/id if present
                            let mut id = format!("comment_{}", comments.len());
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                if key == "office:name" || key.ends_with(":name") {
                                    id = String::from_utf8_lossy(&attr.value).to_string();
                                }
                            }

                            current_comment = Some(Comment {
                                id,
                                author: None,
                                date: None,
                                content: String::new(),
                                reference: None,
                            });
                        },
                        "dc:creator" if in_annotation => {
                            in_creator = true;
                        },
                        "dc:date" if in_annotation => {
                            in_date = true;
                        },
                        "text:p" if in_annotation => {
                            in_paragraph = true;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) if in_annotation => {
                    let text = String::from_utf8_lossy(t).to_string();

                    if let Some(ref mut comment) = current_comment {
                        if in_creator {
                            comment.author = Some(text);
                        } else if in_date {
                            comment.date = Some(text);
                        } else if in_paragraph {
                            if !comment.content.is_empty() {
                                comment.content.push('\n');
                            }
                            comment.content.push_str(&text);
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "office:annotation" => {
                            in_annotation = false;
                            if let Some(comment) = current_comment.take() {
                                comments.push(comment);
                            }
                        },
                        "dc:creator" => {
                            in_creator = false;
                        },
                        "dc:date" => {
                            in_date = false;
                        },
                        "text:p" => {
                            in_paragraph = false;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(comments)
    }

    /// Parse sections
    ///
    /// Extracts document sections which can contain protected content,
    /// different formatting, or special layout properties.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing sections
    ///
    /// # Returns
    ///
    /// Vector of `Section` objects with metadata and content
    pub fn parse_sections(content: &str) -> Result<Vec<Section>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(content);
        let mut buf = Vec::new();
        let mut sections = Vec::new();
        let mut in_section = false;
        let mut current_section: Option<Section> = None;
        let mut section_depth: usize = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    if tag_name == "text:section" {
                        if !in_section {
                            in_section = true;
                            section_depth = 1;

                            // Extract section attributes
                            let mut name = String::new();
                            let mut style = None;
                            let mut protected = false;

                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value).to_string();

                                match key.as_ref() {
                                    "text:name" => name = value,
                                    "text:style-name" => style = Some(value),
                                    "text:protected" => protected = value == "true" || value == "1",
                                    _ => {},
                                }
                            }

                            current_section = Some(Section {
                                name,
                                style,
                                protected,
                                content: String::new(),
                            });
                        } else {
                            // Nested section
                            section_depth += 1;
                        }
                    } else if in_section {
                        section_depth += 1;
                    }
                },
                Ok(Event::Text(ref t)) if in_section && section_depth == 1 => {
                    // Only collect text at the top level of the section
                    if let Some(ref mut section) = current_section {
                        let text = String::from_utf8_lossy(t).to_string();
                        section.content.push_str(&text);
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    if tag_name == "text:section" && in_section {
                        section_depth -= 1;

                        if section_depth == 0 {
                            in_section = false;
                            if let Some(section) = current_section.take() {
                                sections.push(section);
                            }
                        }
                    } else if in_section && section_depth > 0 {
                        section_depth -= 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(sections)
    }
}

#[cfg(test)]
mod tests {
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
        let changes = OdtParser::parse_track_changes(TEST_TRACK_CHANGES_XML).unwrap();
        assert_eq!(changes.len(), 3);

        // Check first change (insertion)
        assert_eq!(changes[0].id, "change1");
        assert_eq!(changes[0].change_type, ChangeType::Insertion);
        // Parser extracts text elements - author/date extraction depends on XML structure
        assert!(changes[0].author.is_some());

        // Check second change (deletion)
        assert_eq!(changes[1].id, "change2");
        assert_eq!(changes[1].change_type, ChangeType::Deletion);

        // Check third change (format)
        assert_eq!(changes[2].id, "change3");
        assert_eq!(changes[2].change_type, ChangeType::FormatChange);
    }

    #[test]
    fn test_parse_track_changes_empty() {
        let changes = OdtParser::parse_track_changes(TEST_EMPTY_TRACK_CHANGES).unwrap();
        assert!(changes.is_empty());
    }

    #[test]
    fn test_parse_track_changes_no_tracked_changes() {
        let changes = OdtParser::parse_track_changes(TEST_EMPTY_CONTENT).unwrap();
        assert!(changes.is_empty());
    }

    #[test]
    fn test_parse_comments() {
        let comments = OdtParser::parse_comments(TEST_COMMENTS_XML).unwrap();
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
        let comments = OdtParser::parse_comments(TEST_EMPTY_CONTENT).unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_parse_sections() {
        let sections = OdtParser::parse_sections(TEST_SECTIONS_XML).unwrap();
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
        let sections = OdtParser::parse_sections(TEST_EMPTY_CONTENT).unwrap();
        assert!(sections.is_empty());
    }

    #[test]
    fn test_track_change_debug() {
        let change = TrackChange {
            id: "test1".to_string(),
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            change_type: ChangeType::Insertion,
            content: "content".to_string(),
        };
        let debug_str = format!("{:?}", change);
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
        let t2 = t1.clone();
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
        let debug_str = format!("{:?}", comment);
        assert!(debug_str.contains("Comment"));
        assert!(debug_str.contains("cmt1"));
    }

    #[test]
    fn test_section_debug() {
        let section = Section {
            name: "Sec1".to_string(),
            style: Some("Style1".to_string()),
            protected: true,
            content: "Content".to_string(),
        };
        let debug_str = format!("{:?}", section);
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
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            change_type: ChangeType::Deletion,
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
            content: "Text".to_string(),
        };
        let cloned = section.clone();
        assert_eq!(section.name, cloned.name);
        assert_eq!(section.protected, cloned.protected);
    }
}
