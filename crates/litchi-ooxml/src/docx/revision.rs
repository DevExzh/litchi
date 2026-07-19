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
/// use litchi_ooxml::docx::Package;
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
use crate::error::{OoxmlError, Result};
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
    /// Table property change
    TablePropertiesChange,
    /// Table row insertion
    RowInsert,
    /// Table row deletion
    RowDelete,
    /// Table row property change
    RowPropertiesChange,
    /// Table cell insertion
    CellInsert,
    /// Table cell deletion
    CellDelete,
    /// Table cell merge change
    CellMerge,
    /// Table cell property change
    CellPropertiesChange,
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
            Self::TablePropertiesChange => write!(f, "Table Properties Change"),
            Self::RowInsert => write!(f, "Row Insert"),
            Self::RowDelete => write!(f, "Row Delete"),
            Self::RowPropertiesChange => write!(f, "Row Properties Change"),
            Self::CellInsert => write!(f, "Cell Insert"),
            Self::CellDelete => write!(f, "Cell Delete"),
            Self::CellMerge => write!(f, "Cell Merge"),
            Self::CellPropertiesChange => write!(f, "Cell Properties Change"),
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
    let mut in_row_properties = false;
    let mut current_revision: Option<Revision> = None;

    fn revision_type(local_name: &[u8], in_row_properties: bool) -> Option<RevisionType> {
        match local_name {
            b"ins" if in_row_properties => Some(RevisionType::RowInsert),
            b"del" if in_row_properties => Some(RevisionType::RowDelete),
            b"ins" => Some(RevisionType::Insert),
            b"del" => Some(RevisionType::Delete),
            b"moveFrom" => Some(RevisionType::MoveFrom),
            b"moveTo" => Some(RevisionType::MoveTo),
            b"rPrChange" | b"pPrChange" => Some(RevisionType::FormatChange),
            b"tblIns" => Some(RevisionType::TableInsert),
            b"tblDel" => Some(RevisionType::TableDelete),
            b"tblPrChange" => Some(RevisionType::TablePropertiesChange),
            b"trPrChange" => Some(RevisionType::RowPropertiesChange),
            b"cellIns" => Some(RevisionType::CellInsert),
            b"cellDel" => Some(RevisionType::CellDelete),
            b"cellMerge" => Some(RevisionType::CellMerge),
            b"tcPrChange" => Some(RevisionType::CellPropertiesChange),
            _ => None,
        }
    }

    fn revision_from_element(
        element: &quick_xml::events::BytesStart<'_>,
        revision_type: RevisionType,
    ) -> Revision {
        let mut author = String::new();
        let mut date = None;
        let mut id = String::new();

        for attr in element.attributes().flatten() {
            let value = std::str::from_utf8(attr.value.as_ref())
                .ok()
                .and_then(|value| quick_xml::escape::unescape(value).ok())
                .map(|value| value.into_owned())
                .unwrap_or_default();
            match attr.key.as_ref() {
                b"w:author" | b"author" => author = value,
                b"w:date" | b"date" => date = Some(value),
                b"w:id" | b"id" => id = value,
                _ => {},
            }
        }

        Revision::new(revision_type, author, date, id)
    }

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                if local_name == b"trPr" {
                    in_row_properties = true;
                }

                if let Some(rev_type) = revision_type(local_name, in_row_properties) {
                    in_revision = true;
                    current_revision = Some(revision_from_element(&e, rev_type));
                } else if in_revision {
                    // Check for text elements within revision
                    match local_name {
                        b"t" | b"delText" | b"delInstrText" => {
                            in_revision_text = true;
                        },
                        _ => {},
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                let local_name_ref = e.local_name();
                if let Some(rev_type) = revision_type(local_name_ref.as_ref(), in_row_properties) {
                    revisions.push(revision_from_element(&e, rev_type));
                }
            },
            Ok(Event::Text(e)) if in_revision_text => {
                // Extract text content from revision
                if let Some(ref mut rev) = current_revision
                    && let Ok(encoded) = e.decode()
                    && let Ok(text) = quick_xml::escape::unescape(&encoded)
                {
                    rev.append_text(&text);
                }
            },
            Ok(Event::End(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"ins" | b"del" | b"moveFrom" | b"moveTo" | b"rPrChange"
                    | b"pPrChange" | b"tblIns" | b"tblDel" | b"tblPrChange"
                    | b"trPrChange" | b"cellIns" | b"cellDel" | b"cellMerge"
                    | b"tcPrChange" => {
                        // Finished parsing a revision
                        in_revision = false;

                        if let Some(revision) = current_revision.take() {
                            revisions.push(revision);
                        }
                    },
                    b"t" | b"delText" | b"delInstrText" => {
                        in_revision_text = false;
                    },
                    b"trPr" => in_row_properties = false,
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
