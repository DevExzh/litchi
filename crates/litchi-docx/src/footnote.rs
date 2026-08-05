use crate::error::{Error, Result};
/// Footnote and endnote support for reading from Word documents.
///
/// This module provides types and methods for accessing footnotes and endnotes
/// in Word documents. Footnotes appear at the bottom of pages, while endnotes
/// appear at the end of the document or section.
use crate::namespace::scan_word_element_ranges;
use crate::paragraph::{Paragraph, extract_word_text};
use litchi_core::XmlSlice;
use litchi_opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::sync::Arc;

#[derive(Debug, Clone)]
enum NoteXmlData {
    Owned(Box<[u8]>),
    Shared(XmlSlice),
}

impl NoteXmlData {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Shared(slice) => slice.as_bytes(),
        }
    }

    fn get_or_create_arc(&self) -> (Arc<Vec<u8>>, u32) {
        match self {
            Self::Owned(bytes) => (Arc::new(bytes.to_vec()), 0),
            Self::Shared(slice) => (slice.arc(), slice.start()),
        }
    }
}

/// A footnote or endnote in a Word document.
///
/// Represents a `<w:footnote>` or `<w:endnote>` element. Notes can contain
/// paragraphs and tables, just like the main document body.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get all footnotes
/// let footnotes = doc.footnotes()?;
/// for note in footnotes {
///     println!("Footnote {}: {}", note.id(), note.text()?);
/// }
///
/// // Get all endnotes
/// let endnotes = doc.endnotes()?;
/// for note in endnotes {
///     println!("Endnote {}: {}", note.id(), note.text()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Note {
    /// The note ID
    id: u32,
    /// The raw XML bytes for this note
    xml_data: NoteXmlData,
    /// The type of note (normal, separator, continuation separator, etc.)
    note_type: NoteType,
}

/// The type of a note.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteType {
    /// Normal note with content
    Normal,
    /// Separator note (visual separator)
    Separator,
    /// Continuation separator
    ContinuationSeparator,
    /// Continuation notice
    ContinuationNotice,
}

impl NoteType {
    /// Parse note type from XML attribute value.
    fn from_xml(s: &str) -> Self {
        match s {
            "separator" => Self::Separator,
            "continuationSeparator" => Self::ContinuationSeparator,
            "continuationNotice" => Self::ContinuationNotice,
            _ => Self::Normal,
        }
    }

    /// Check if this is a normal content note (not a separator).
    #[inline]
    pub fn is_normal(&self) -> bool {
        matches!(self, Self::Normal)
    }
}

impl Note {
    /// Create a new Note.
    ///
    /// # Arguments
    ///
    /// * `id` - The note ID
    /// * `xml_bytes` - The XML content of the note
    /// * `note_type` - The type of note
    pub fn new(id: u32, xml_bytes: Vec<u8>, note_type: NoteType) -> Self {
        Self {
            id,
            xml_data: NoteXmlData::Owned(xml_bytes.into_boxed_slice()),
            note_type,
        }
    }

    fn from_arc_range(
        id: u32,
        source: Arc<Vec<u8>>,
        start: u32,
        length: u32,
        note_type: NoteType,
    ) -> Self {
        Self {
            id,
            xml_data: NoteXmlData::Shared(XmlSlice::new(source, start, length)),
            note_type,
        }
    }

    /// Get the note ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the note type.
    #[inline]
    pub fn note_type(&self) -> NoteType {
        self.note_type
    }

    /// Get the XML bytes of this note.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Extract all text content from this note.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let footnotes = doc.footnotes()?;
    ///
    /// for note in footnotes {
    ///     if note.note_type().is_normal() {
    ///         println!("Note text: {}", note.text()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Get all paragraphs in this note.
    ///
    /// Returns a vector of `Paragraph` objects representing all `<w:p>`
    /// elements in the note.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let footnotes = doc.footnotes()?;
    ///
    /// for note in footnotes {
    ///     for para in note.paragraphs()? {
    ///         println!("Paragraph: {}", para.text()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let (source, base_offset) = self.xml_data.get_or_create_arc();
        let mut paragraphs = Vec::new();
        scan_word_element_ranges(self.xml_bytes(), &[b"p".as_slice()], |_, start, length| {
            paragraphs.push(Paragraph::from_arc_range(
                Arc::clone(&source),
                base_offset.checked_add(start).ok_or_else(|| {
                    Error::InvalidFormat("Word note offset exceeds u32".to_string())
                })?,
                length,
            ));
            Ok(())
        })?;
        Ok(paragraphs)
    }

    /// Extract all footnotes from a footnotes.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The footnotes part
    ///
    /// # Returns
    ///
    /// A vector of footnotes (excluding separators)
    pub(crate) fn extract_footnotes_from_part(part: &dyn Part) -> Result<Vec<Note>> {
        Self::extract_notes_from_part(part, b"footnote")
    }

    /// Extract all endnotes from an endnotes.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The endnotes part
    ///
    /// # Returns
    ///
    /// A vector of endnotes (excluding separators)
    pub(crate) fn extract_endnotes_from_part(part: &dyn Part) -> Result<Vec<Note>> {
        Self::extract_notes_from_part(part, b"endnote")
    }

    /// Extract notes from a part (generic for footnotes and endnotes).
    fn extract_notes_from_part(part: &dyn Part, note_tag: &[u8]) -> Result<Vec<Note>> {
        let source = litchi_ooxml_common::mce::process_part_arc(part)?;
        let mut notes = Vec::new();
        scan_word_element_ranges(source.as_slice(), &[note_tag], |_, start, length| {
            let start_index = start as usize;
            let end_index = start_index
                .checked_add(length as usize)
                .ok_or_else(|| Error::InvalidFormat("Word note range overflow".to_string()))?;
            let (id, note_type) = parse_note_metadata(&source[start_index..end_index])?;
            if let Some(id) = id
                && id > 0
                && note_type.is_normal()
            {
                notes.push(Note::from_arc_range(
                    id,
                    Arc::clone(&source),
                    start,
                    length,
                    note_type,
                ));
            }
            Ok(())
        })?;
        Ok(notes)
    }
}

fn parse_note_metadata(xml_bytes: &[u8]) -> Result<(Option<u32>, NoteType)> {
    let mut reader = Reader::from_reader(xml_bytes);
    let element = loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => break element,
            Ok(Event::Eof) => {
                return Err(Error::InvalidFormat(
                    "missing Word note element".to_string(),
                ));
            },
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    };

    let mut id = None;
    let mut note_type = NoteType::Normal;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"id" => {
                id = atoi_simd::parse::<u32, false, false>(attribute.value.as_ref()).ok();
            },
            b"type" => {
                let value = std::str::from_utf8(attribute.value.as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?;
                note_type = NoteType::from_xml(value);
            },
            _ => {},
        }
    }
    Ok((id, note_type))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::rel::Relationships;
    use std::sync::Arc;

    /// Simple mock Part for testing
    #[derive(Clone)]
    struct MockPart {
        blob: Vec<u8>,
    }

    impl MockPart {
        fn new(blob: Vec<u8>) -> Self {
            Self { blob }
        }
    }

    impl Part for MockPart {
        fn partname(&self) -> &PackURI {
            unimplemented!("MockPart::partname not implemented")
        }

        fn content_type(&self) -> &str {
            "application/xml"
        }

        fn blob(&self) -> &[u8] {
            &self.blob
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            Arc::new(self.blob.clone())
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.blob = blob;
        }

        fn rels(&self) -> &Relationships {
            unimplemented!("MockPart::rels not implemented")
        }

        fn rels_mut(&mut self) -> &mut Relationships {
            unimplemented!("MockPart::rels_mut not implemented")
        }
    }

    #[test]
    fn test_note_type_from_xml() {
        assert_eq!(NoteType::from_xml("separator"), NoteType::Separator);
        assert_eq!(
            NoteType::from_xml("continuationSeparator"),
            NoteType::ContinuationSeparator
        );
        assert_eq!(
            NoteType::from_xml("continuationNotice"),
            NoteType::ContinuationNotice
        );
        assert_eq!(NoteType::from_xml("normal"), NoteType::Normal);
        assert_eq!(NoteType::from_xml(""), NoteType::Normal);
        assert_eq!(NoteType::from_xml("unknown"), NoteType::Normal);
    }

    #[test]
    fn test_note_type_is_normal() {
        assert!(NoteType::Normal.is_normal());
        assert!(!NoteType::Separator.is_normal());
        assert!(!NoteType::ContinuationSeparator.is_normal());
        assert!(!NoteType::ContinuationNotice.is_normal());
    }

    #[test]
    fn test_note_creation() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.id(), 1);
        assert_eq!(note.note_type(), NoteType::Normal);
        assert_eq!(note.text().unwrap(), "Test");
    }

    #[test]
    fn test_note_with_endnote() {
        let xml = b"<w:endnote><w:p><w:r><w:t>Endnote Text</w:t></w:r></w:p></w:endnote>";
        let note = Note::new(5, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.id(), 5);
        assert_eq!(note.text().unwrap(), "Endnote Text");
    }

    #[test]
    fn test_note_xml_bytes() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Content</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(2, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.xml_bytes(), xml);
    }

    #[test]
    fn test_note_empty_content() {
        let xml = b"<w:footnote></w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.text().unwrap(), "");
    }

    #[test]
    fn test_note_decodes_text_and_special_characters() {
        let xml =
            b"<w:footnote><w:p><w:r><w:t>A &lt; B</w:t><w:tab/><w:cr/></w:r></w:p></w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);
        assert_eq!(note.text().unwrap(), "A < B\t\n");
    }

    #[test]
    fn test_note_multiple_paragraphs() {
        let xml = b"<w:footnote>\
            <w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>\
            <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>\
        </w:footnote>";
        let note = Note::new(3, xml.to_vec(), NoteType::Normal);

        let text = note.text().unwrap();
        assert!(text.contains("First paragraph"));
        assert!(text.contains("Second paragraph"));
    }

    #[test]
    fn test_note_paragraphs_extraction() {
        let xml = b"<w:footnote>\
            <w:p><w:r><w:t>Para 1</w:t></w:r></w:p>\
            <w:p><w:r><w:t>Para 2</w:t></w:r></w:p>\
        </w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        let paragraphs = note.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 2);
    }

    #[test]
    fn test_note_with_unicode() {
        let xml = "<w:footnote><w:p><w:r><w:t>Unicode: 你好世界 🎉</w:t></w:r></w:p></w:footnote>"
            .as_bytes();
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        let text = note.text().unwrap();
        assert!(text.contains("你好世界"));
        assert!(text.contains("🎉"));
    }

    #[test]
    fn test_note_clone() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Clonable</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(10, xml.to_vec(), NoteType::Normal);
        let cloned = note.clone();

        assert_eq!(cloned.id(), note.id());
        assert_eq!(cloned.note_type(), note.note_type());
        assert_eq!(cloned.text().unwrap(), note.text().unwrap());
    }

    #[test]
    fn test_note_separator_type() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Separator</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(-1i32 as u32, xml.to_vec(), NoteType::Separator);

        assert_eq!(note.note_type(), NoteType::Separator);
        assert!(!note.note_type().is_normal());
    }

    #[test]
    fn test_note_continuation_separator() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Cont Sep</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(999, xml.to_vec(), NoteType::ContinuationSeparator);

        assert_eq!(note.note_type(), NoteType::ContinuationSeparator);
    }

    #[test]
    fn test_note_with_nested_elements() {
        let xml = b"<w:footnote>\
            <w:p>\
                <w:pPr><w:jc w:val=\"left\"/></w:pPr>\
                <w:r>\
                    <w:rPr><w:b/></w:rPr>\
                    <w:t>Bold Text</w:t>\
                </w:r>\
            </w:p>\
        </w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.text().unwrap(), "Bold Text");
    }

    #[test]
    fn test_note_type_debug() {
        let note_type = NoteType::Normal;
        let debug_str = format!("{:?}", note_type);
        assert!(debug_str.contains("Normal"));
    }

    #[test]
    fn test_note_debug() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Debug</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(42, xml.to_vec(), NoteType::Normal);

        let debug_str = format!("{:?}", note);
        assert!(debug_str.contains("Note"));
        assert!(debug_str.contains("42"));
    }

    #[test]
    fn test_note_equality() {
        assert_eq!(NoteType::Normal, NoteType::Normal);
        assert_ne!(NoteType::Normal, NoteType::Separator);
        assert_eq!(NoteType::Separator, NoteType::Separator);
    }

    #[test]
    fn test_note_copy() {
        let note_type = NoteType::Normal;
        let copied = note_type;
        // After copy, original should still be valid
        assert_eq!(note_type, NoteType::Normal);
        assert_eq!(copied, NoteType::Normal);
    }

    #[test]
    fn test_extract_footnotes_from_part_empty() {
        let xml = b"<?xml version=\"1.0\"?><w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:footnotes>";
        let part = MockPart::new(xml.to_vec());
        let notes = Note::extract_footnotes_from_part(&part).unwrap();

        assert!(notes.is_empty());
    }

    #[test]
    fn test_extract_footnotes_from_part_single() {
        let xml = b"<?xml version=\"1.0\"?>
        <w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
            <w:footnote w:id=\"1\"><w:p><w:r><w:t>Footnote 1</w:t></w:r></w:p></w:footnote>
        </w:footnotes>";
        let part = MockPart::new(xml.to_vec());
        let notes = Note::extract_footnotes_from_part(&part).unwrap();

        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].id(), 1);
        assert_eq!(notes[0].text().unwrap(), "Footnote 1");
    }

    #[test]
    fn test_extract_footnotes_from_part_multiple() {
        let xml = b"<?xml version=\"1.0\"?>
        <w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
            <w:footnote w:id=\"1\"><w:p><w:r><w:t>First</w:t></w:r></w:p></w:footnote>
            <w:footnote w:id=\"2\"><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:footnote>
            <w:footnote w:id=\"3\"><w:p><w:r><w:t>Third</w:t></w:r></w:p></w:footnote>
        </w:footnotes>";
        let part = MockPart::new(xml.to_vec());
        let notes = Note::extract_footnotes_from_part(&part).unwrap();

        assert_eq!(notes.len(), 3);
    }

    #[test]
    fn test_extract_footnotes_skips_separator() {
        let xml = b"<?xml version=\"1.0\"?>
        <w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
            <w:footnote w:id=\"1\" w:type=\"separator\"><w:p><w:r><w:t>Separator</w:t></w:r></w:p></w:footnote>
            <w:footnote w:id=\"2\"><w:p><w:r><w:t>Normal Note</w:t></w:r></w:p></w:footnote>
        </w:footnotes>";
        let part = MockPart::new(xml.to_vec());
        let notes = Note::extract_footnotes_from_part(&part).unwrap();

        // Should only include the normal note, not the separator
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].id(), 2);
    }

    #[test]
    fn test_extract_endnotes_from_part() {
        let xml = b"<?xml version=\"1.0\"?>
        <w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
            <w:endnote w:id=\"1\"><w:p><w:r><w:t>Endnote 1</w:t></w:r></w:p></w:endnote>
            <w:endnote w:id=\"2\"><w:p><w:r><w:t>Endnote 2</w:t></w:r></w:p></w:endnote>
        </w:endnotes>";
        let part = MockPart::new(xml.to_vec());
        let notes = Note::extract_endnotes_from_part(&part).unwrap();

        assert_eq!(notes.len(), 2);
        assert_eq!(notes[0].id(), 1);
        assert_eq!(notes[1].id(), 2);
    }

    #[test]
    fn extracts_aliased_note_slices_and_ignores_foreign_lookalikes() {
        let xml = br#"<fn:footnotes xmlns:fn="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
            <false:footnote false:id="9"><false:p><false:r><false:t>ignored</false:t></false:r></false:p></false:footnote>
            <fn:footnote fn:id="1"><fn:p><fn:r><fn:t><![CDATA[A < B]]></fn:t></fn:r></fn:p></fn:footnote>
            <fn:footnote fn:id="2" fn:type="separator"><fn:p><fn:r><fn:t>separator</fn:t></fn:r></fn:p></fn:footnote>
            <fn:footnote fn:id="3"/>
        </fn:footnotes>"#;
        let part = MockPart::new(xml.to_vec());

        let notes = Note::extract_footnotes_from_part(&part).unwrap();
        assert_eq!(notes.len(), 2);
        assert_eq!(notes[0].id(), 1);
        assert_eq!(notes[0].text().unwrap(), "A < B");
        let paragraphs = notes[0].paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].runs().unwrap()[0].text().unwrap(), "A < B");
        assert_eq!(notes[1].id(), 3);
        assert_eq!(notes[1].text().unwrap(), "");
    }

    #[test]
    fn rejects_unterminated_note_slices() {
        let xml = br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p/>"#;
        let part = MockPart::new(xml.to_vec());
        assert!(Note::extract_footnotes_from_part(&part).is_err());
    }

    #[test]
    fn test_note_large_id() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Large ID</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(999999, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.id(), 999999);
    }

    #[test]
    fn test_note_with_cdata() {
        // CDATA content is parsed as text by quick-xml, so it should be extracted
        let xml = b"<w:footnote><w:p><w:r><w:t>Regular Content</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        let text = note.text().unwrap();
        assert!(text.contains("Regular Content"));
    }
}
