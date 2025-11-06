/// Track changes (revisions) support for DOCX documents.
///
/// This module provides structures and functions for reading tracked changes
/// (revisions) from Word documents. Tracked changes record insertions, deletions,
/// moves, and formatting changes made by document editors.
///
/// # Architecture
///
/// - `Revision`: A single tracked change
/// - `RevisionType`: Type of change (insert, delete, move, format)
/// - `RevisionInfo`: Metadata about who made the change and when
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get all revisions from the document
/// for para in doc.paragraphs()? {
///     for revision in para.revisions()? {
///         println!("Revision by {}: {} - {}",
///             revision.author(),
///             revision.revision_type(),
///             revision.text()
///         );
///         if let Some(date) = revision.date() {
///             println!("  Made on: {}", date);
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::fmt;

/// Type of tracked change.
///
/// Represents the different types of revisions that can be tracked
/// in a Word document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionType {
    /// Text insertion
    Insert,
    /// Text deletion
    Delete,
    /// Move from (cut)
    MoveFrom,
    /// Move to (paste)
    MoveTo,
    /// Formatting change
    FormatChange,
    /// Table insertion
    TableInsert,
    /// Table deletion
    TableDelete,
    /// Custom or unknown revision type
    Unknown,
}

impl fmt::Display for RevisionType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Insert => write!(f, "Insert"),
            Self::Delete => write!(f, "Delete"),
            Self::MoveFrom => write!(f, "Move From"),
            Self::MoveTo => write!(f, "Move To"),
            Self::FormatChange => write!(f, "Format Change"),
            Self::TableInsert => write!(f, "Table Insert"),
            Self::TableDelete => write!(f, "Table Delete"),
            Self::Unknown => write!(f, "Unknown"),
        }
    }
}

/// A tracked change (revision) in a Word document.
///
/// Represents a single change tracked by Word's revision system.
/// Contains information about what changed, who made the change, and when.
///
/// # Field Ordering
///
/// Fields are ordered to maximize CPU cache line utilization:
/// - Strings (24 bytes each on 64-bit systems)
/// - Enums and smaller types
#[derive(Debug, Clone)]
pub struct Revision {
    /// Author who made the change
    author: String,

    /// Date/time of the change (ISO 8601 format)
    date: Option<String>,

    /// Text content affected by this revision
    text: String,

    /// Revision ID
    id: String,

    /// Type of revision
    revision_type: RevisionType,
}

impl Revision {
    /// Create a new Revision.
    ///
    /// # Arguments
    ///
    /// * `revision_type` - Type of revision
    /// * `author` - Author who made the change
    /// * `date` - Date/time of the change
    /// * `id` - Revision ID
    #[inline]
    pub fn new(
        revision_type: RevisionType,
        author: String,
        date: Option<String>,
        id: String,
    ) -> Self {
        Self {
            author,
            date,
            text: String::new(),
            id,
            revision_type,
        }
    }

    /// Get the revision type.
    #[inline]
    pub fn revision_type(&self) -> RevisionType {
        self.revision_type
    }

    /// Get the author who made the change.
    #[inline]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Get the date/time of the change.
    #[inline]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Get the revision ID.
    #[inline]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Get the text content affected by this revision.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content.
    #[inline]
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Append text content.
    #[inline]
    pub fn append_text(&mut self, text: &str) {
        self.text.push_str(text);
    }
}

/// Parse revisions from paragraph XML.
///
/// Extracts all tracked changes (w:ins, w:del, w:moveFrom, w:moveTo) from
/// the paragraph XML.
///
/// # Arguments
///
/// * `xml_bytes` - The raw XML bytes of the paragraph
///
/// # Performance
///
/// Uses streaming XML parsing with pre-allocated SmallVec for efficient
/// storage of typically small revision collections.
///
/// # Example XML Structure
///
/// ```xml
/// <w:p>
///   <w:r>
///     <w:t>Normal text</w:t>
///   </w:r>
///   <w:ins w:id="0" w:author="John Doe" w:date="2024-11-05T10:30:00Z">
///     <w:r>
///       <w:t>inserted text</w:t>
///     </w:r>
///   </w:ins>
///   <w:del w:id="1" w:author="Jane Smith" w:date="2024-11-05T11:00:00Z">
///     <w:r>
///       <w:delText>deleted text</w:delText>
///     </w:r>
///   </w:del>
/// </w:p>
/// ```
pub(crate) fn parse_revisions(xml_bytes: &[u8]) -> Result<SmallVec<[Revision; 4]>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    // Use SmallVec for efficient storage of typically small revision collections
    let mut revisions = SmallVec::new();

    // State tracking for parsing
    let mut in_revision = false;
    let mut in_revision_text = false;
    let mut current_revision: Option<Revision> = None;

    let mut buf = Vec::with_capacity(1024); // Reusable buffer

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                // Check for revision elements
                let revision_type = match local_name {
                    b"ins" => Some(RevisionType::Insert),
                    b"del" => Some(RevisionType::Delete),
                    b"moveFrom" => Some(RevisionType::MoveFrom),
                    b"moveTo" => Some(RevisionType::MoveTo),
                    b"rPrChange" => Some(RevisionType::FormatChange),
                    b"tblIns" => Some(RevisionType::TableInsert),
                    b"tblDel" => Some(RevisionType::TableDelete),
                    _ => None,
                };

                if let Some(rev_type) = revision_type {
                    in_revision = true;

                    // Parse revision attributes
                    let mut author = String::new();
                    let mut date = None;
                    let mut id = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"w:author" | b"author" => {
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    author = s.to_string();
                                }
                            },
                            b"w:date" | b"date" => {
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    date = Some(s.to_string());
                                }
                            },
                            b"w:id" | b"id" => {
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    id = s.to_string();
                                }
                            },
                            _ => {},
                        }
                    }

                    current_revision = Some(Revision::new(rev_type, author, date, id));
                } else if in_revision {
                    // Check for text elements within revision
                    match local_name {
                        b"t" | b"delText" => {
                            in_revision_text = true;
                        },
                        _ => {},
                    }
                }
            },
            Ok(Event::Text(e)) if in_revision_text => {
                // Extract text content from revision
                if let Some(ref mut rev) = current_revision
                    && let Ok(text) = std::str::from_utf8(e.as_ref())
                {
                    rev.append_text(text);
                }
            },
            Ok(Event::End(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"ins" | b"del" | b"moveFrom" | b"moveTo" | b"rPrChange" | b"tblIns"
                    | b"tblDel" => {
                        // Finished parsing a revision
                        in_revision = false;

                        if let Some(revision) = current_revision.take() {
                            revisions.push(revision);
                        }
                    },
                    b"t" | b"delText" => {
                        in_revision_text = false;
                    },
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
    }

    Ok(revisions)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_revision_creation() {
        let rev = Revision::new(
            RevisionType::Insert,
            "John Doe".to_string(),
            Some("2024-11-05T10:30:00Z".to_string()),
            "0".to_string(),
        );

        assert_eq!(rev.revision_type(), RevisionType::Insert);
        assert_eq!(rev.author(), "John Doe");
        assert_eq!(rev.date(), Some("2024-11-05T10:30:00Z"));
        assert_eq!(rev.id(), "0");
        assert_eq!(rev.text(), "");
    }

    #[test]
    fn test_parse_revisions_empty() {
        let xml = b"<w:p><w:r><w:t>Normal text</w:t></w:r></w:p>";
        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 0);
    }

    #[test]
    fn test_parse_insert_revision() {
        let xml = br#"<w:p>
            <w:ins w:id="0" w:author="John Doe" w:date="2024-11-05T10:30:00Z">
                <w:r>
                    <w:t>inserted text</w:t>
                </w:r>
            </w:ins>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);

        let rev = &revisions[0];
        assert_eq!(rev.revision_type(), RevisionType::Insert);
        assert_eq!(rev.author(), "John Doe");
        assert_eq!(rev.date(), Some("2024-11-05T10:30:00Z"));
        assert_eq!(rev.id(), "0");
        assert_eq!(rev.text(), "inserted text");
    }

    #[test]
    fn test_parse_delete_revision() {
        let xml = br#"<w:p>
            <w:del w:id="1" w:author="Jane Smith" w:date="2024-11-05T11:00:00Z">
                <w:r>
                    <w:delText>deleted text</w:delText>
                </w:r>
            </w:del>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);

        let rev = &revisions[0];
        assert_eq!(rev.revision_type(), RevisionType::Delete);
        assert_eq!(rev.author(), "Jane Smith");
        assert_eq!(rev.text(), "deleted text");
    }

    #[test]
    fn test_parse_multiple_revisions() {
        let xml = br#"<w:p>
            <w:ins w:id="0" w:author="Author1">
                <w:r><w:t>inserted</w:t></w:r>
            </w:ins>
            <w:del w:id="1" w:author="Author2">
                <w:r><w:delText>deleted</w:delText></w:r>
            </w:del>
            <w:moveFrom w:id="2" w:author="Author3">
                <w:r><w:t>moved</w:t></w:r>
            </w:moveFrom>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 3);

        assert_eq!(revisions[0].revision_type(), RevisionType::Insert);
        assert_eq!(revisions[1].revision_type(), RevisionType::Delete);
        assert_eq!(revisions[2].revision_type(), RevisionType::MoveFrom);
    }

    #[test]
    fn test_revision_type_display() {
        assert_eq!(format!("{}", RevisionType::Insert), "Insert");
        assert_eq!(format!("{}", RevisionType::Delete), "Delete");
        assert_eq!(format!("{}", RevisionType::MoveFrom), "Move From");
        assert_eq!(format!("{}", RevisionType::MoveTo), "Move To");
        assert_eq!(format!("{}", RevisionType::FormatChange), "Format Change");
    }
}
