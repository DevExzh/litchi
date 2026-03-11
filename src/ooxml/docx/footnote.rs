/// Footnote and endnote support for reading from Word documents.
///
/// This module provides types and methods for accessing footnotes and endnotes
/// in Word documents. Footnotes appear at the bottom of pages, while endnotes
/// appear at the end of the document or section.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A footnote or endnote in a Word document.
///
/// Represents a `<w:footnote>` or `<w:endnote>` element. Notes can contain
/// paragraphs and tables, just like the main document body.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
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
    xml_bytes: Vec<u8>,
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
            xml_bytes,
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
        &self.xml_bytes
    }

    /// Extract all text content from this note.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity
        let estimated_capacity = self.xml_bytes.len() / 8;
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    result.push_str(text);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        result.shrink_to_fit();
        Ok(result)
    }

    /// Get all paragraphs in this note.
    ///
    /// Returns a vector of `Paragraph` objects representing all `<w:p>`
    /// elements in the note.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut paragraphs = Vec::new();
        let mut current_para_xml = Vec::with_capacity(4096);
        let mut in_para = false;
        let mut depth = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        current_para_xml.extend_from_slice(b"<w:p");
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b">");
                    } else if in_para {
                        depth += 1;
                        current_para_xml.extend_from_slice(b"<");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.extend_from_slice(b">");

                        if e.local_name().as_ref() == b"p" && depth == 1 {
                            paragraphs.push(Paragraph::new(current_para_xml.clone()));
                            in_para = false;
                        } else {
                            depth -= 1;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"<");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b"/>");
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::CData(e)) if in_para => {
                    current_para_xml.extend_from_slice(b"<![CDATA[");
                    current_para_xml.extend_from_slice(e.as_ref());
                    current_para_xml.extend_from_slice(b"]]>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

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
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut notes = Vec::new();
        let mut current_note_xml = Vec::with_capacity(4096);
        let mut in_note = false;
        let mut depth = 0;
        let mut current_id: Option<u32> = None;
        let mut current_type = NoteType::Normal;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == note_tag && !in_note {
                        in_note = true;
                        depth = 1;
                        current_note_xml.clear();
                        current_id = None;
                        current_type = NoteType::Normal;

                        // Parse attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"id" => {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_id = atoi_simd::parse::<u32>(id_str.as_bytes()).ok();
                                },
                                b"type" => {
                                    let type_str = String::from_utf8_lossy(&attr.value);
                                    current_type = NoteType::from_xml(&type_str);
                                },
                                _ => {},
                            }
                        }

                        // Build opening tag
                        current_note_xml.extend_from_slice(b"<w:");
                        current_note_xml.extend_from_slice(note_tag);
                        current_note_xml.extend_from_slice(b">");
                    } else if in_note {
                        depth += 1;
                        current_note_xml.extend_from_slice(b"<");
                        current_note_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_note_xml.extend_from_slice(b" ");
                            current_note_xml.extend_from_slice(attr.key.as_ref());
                            current_note_xml.extend_from_slice(b"=\"");
                            current_note_xml.extend_from_slice(&attr.value);
                            current_note_xml.extend_from_slice(b"\"");
                        }
                        current_note_xml.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_note {
                        current_note_xml.extend_from_slice(b"</");
                        current_note_xml.extend_from_slice(e.name().as_ref());
                        current_note_xml.extend_from_slice(b">");

                        if e.local_name().as_ref() == note_tag && depth == 1 {
                            // End of note element
                            if let Some(id) = current_id {
                                // Skip separator notes (negative IDs or special types)
                                if id > 0 && current_type.is_normal() {
                                    notes.push(Note::new(
                                        id,
                                        current_note_xml.clone(),
                                        current_type,
                                    ));
                                }
                            }
                            in_note = false;
                        } else {
                            depth -= 1;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    if in_note {
                        current_note_xml.extend_from_slice(b"<");
                        current_note_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_note_xml.extend_from_slice(b" ");
                            current_note_xml.extend_from_slice(attr.key.as_ref());
                            current_note_xml.extend_from_slice(b"=\"");
                            current_note_xml.extend_from_slice(&attr.value);
                            current_note_xml.extend_from_slice(b"\"");
                        }
                        current_note_xml.extend_from_slice(b"/>");
                    }
                },
                Ok(Event::Text(e)) if in_note => {
                    current_note_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::CData(e)) if in_note => {
                    current_note_xml.extend_from_slice(b"<![CDATA[");
                    current_note_xml.extend_from_slice(e.as_ref());
                    current_note_xml.extend_from_slice(b"]]>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(notes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::opc::packuri::PackURI;
    use crate::ooxml::opc::rel::Relationships;
    use std::sync::Arc;

    /// Simple mock Part for testing
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
