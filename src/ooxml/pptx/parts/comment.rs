/// Comment parts for PowerPoint presentations.
///
/// This module provides types for working with comments in PPTX files.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A comment author.
///
/// Represents information about a comment author from the comment authors part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentAuthor {
    /// Author ID
    pub id: u32,
    /// Author name
    pub name: String,
    /// Author initials
    pub initials: String,
}

/// A comment in a presentation.
///
/// Comments are annotations attached to specific positions on slides.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment author ID
    pub author_id: u32,
    /// Comment text
    pub text: String,
    /// X position in EMUs
    pub x: i64,
    /// Y position in EMUs
    pub y: i64,
    /// Comment date/time as string (ISO 8601 format)
    pub datetime: Option<String>,
    /// Comment index
    pub index: Option<u32>,
}

/// Comments part - contains comments for a slide.
///
/// Corresponds to `/ppt/comments/commentN.xml` in the package.
pub struct CommentsPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> CommentsPart<'a> {
    /// Create a CommentsPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the comments.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return all comments from this part.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let comments_part = CommentsPart::from_part(part)?;
    /// let comments = comments_part.comments()?;
    /// for comment in comments {
    ///     println!("Comment: {}", comment.text);
    /// }
    /// ```
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut comments = Vec::new();
        let mut current_comment: Option<Comment> = None;
        let mut in_text = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    match tag_name.as_ref() {
                        b"cm" => {
                            // Start of a comment element
                            let mut author_id = 0;
                            let x = 0;
                            let y = 0;
                            let mut datetime = None;
                            let mut index = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"authorId" => {
                                        author_id = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    },
                                    b"dt" => {
                                        datetime = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(|s| s.to_string());
                                    },
                                    b"idx" => {
                                        index = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    },
                                    _ => {},
                                }
                            }

                            current_comment = Some(Comment {
                                author_id,
                                text: String::new(),
                                x,
                                y,
                                datetime,
                                index,
                            });
                        },
                        b"pos" => {
                            // Position element
                            if let Some(ref mut comment) = current_comment {
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"x" => {
                                            comment.x = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .and_then(|s| s.parse().ok())
                                                .unwrap_or(0);
                                        },
                                        b"y" => {
                                            comment.y = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .and_then(|s| s.parse().ok())
                                                .unwrap_or(0);
                                        },
                                        _ => {},
                                    }
                                }
                            }
                        },
                        b"text" => {
                            in_text = true;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(e)) if in_text => {
                    if let Some(ref mut comment) = current_comment {
                        let text = std::str::from_utf8(e.as_ref())
                            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        comment.text.push_str(text);
                    }
                },
                Ok(Event::End(e)) => {
                    let tag_name = e.local_name();
                    match tag_name.as_ref() {
                        b"cm" => {
                            if let Some(comment) = current_comment.take() {
                                comments.push(comment);
                            }
                        },
                        b"text" => {
                            in_text = false;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(comments)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

/// Comment authors part - contains author information.
///
/// Corresponds to `/ppt/commentAuthors.xml` in the package.
pub struct CommentAuthorsPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> CommentAuthorsPart<'a> {
    /// Create a CommentAuthorsPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the comment authors.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return all comment authors from this part.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let authors_part = CommentAuthorsPart::from_part(part)?;
    /// let authors = authors_part.authors()?;
    /// for author in authors {
    ///     println!("Author: {}", author.name);
    /// }
    /// ```
    pub fn authors(&self) -> Result<Vec<CommentAuthor>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut authors = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"cmAuthor" {
                        let mut id = 0;
                        let mut name = String::new();
                        let mut initials = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"id" => {
                                    id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                b"name" => {
                                    name = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();
                                },
                                b"initials" => {
                                    initials = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();
                                },
                                _ => {},
                            }
                        }

                        authors.push(CommentAuthor { id, name, initials });
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(authors)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}
