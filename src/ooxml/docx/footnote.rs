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
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut buf = Vec::with_capacity(1024);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut buf = Vec::with_capacity(1024);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(notes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_note_type() {
        assert_eq!(NoteType::from_xml("separator"), NoteType::Separator);
        assert_eq!(
            NoteType::from_xml("continuationSeparator"),
            NoteType::ContinuationSeparator
        );
        assert_eq!(NoteType::from_xml("normal"), NoteType::Normal);
        assert!(NoteType::Normal.is_normal());
        assert!(!NoteType::Separator.is_normal());
    }

    #[test]
    fn test_note_creation() {
        let xml = b"<w:footnote><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:footnote>";
        let note = Note::new(1, xml.to_vec(), NoteType::Normal);

        assert_eq!(note.id(), 1);
        assert_eq!(note.note_type(), NoteType::Normal);
        assert_eq!(note.text().unwrap(), "Test");
    }
}
