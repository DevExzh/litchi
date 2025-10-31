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
