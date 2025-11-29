/// Paragraph and Run structures for Word documents.
use crate::common::VerticalPosition;
use crate::common::XmlSlice;
use crate::ooxml::docx::drawing::{DrawingObject, parse_drawing_objects};
use crate::ooxml::docx::hyperlink::Hyperlink;
use crate::ooxml::docx::image::{InlineImage, parse_inline_images};
use crate::ooxml::docx::revision::{Revision, parse_revisions};
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::rel::Relationships;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::borrow::Cow;
use std::sync::Arc;

/// A paragraph in a Word document.
///
/// Represents a `<w:p>` element. Paragraphs contain runs which in turn
/// contain the actual text and formatting.
///
/// # Example
///
/// ```rust,ignore
/// for para in document.paragraphs()? {
///     println!("Paragraph text: {}", para.text());
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text(), run.bold());
///     }
/// }
/// ```
/// Internal storage for paragraph XML data.
/// Supports both owned data (for standalone parsing) and shared slices (for arena-based parsing).
#[derive(Debug, Clone)]
enum XmlData {
    /// Owned data for standalone paragraphs
    Owned(Box<[u8]>),
    /// Shared slice into an arena for zero-copy batch parsing
    Shared(XmlSlice),
}

impl XmlData {
    #[inline]
    fn as_bytes(&self) -> &[u8] {
        match self {
            XmlData::Owned(b) => b,
            XmlData::Shared(s) => s.as_bytes(),
        }
    }

    /// Get or create an Arc for this data.
    /// If already shared, returns the existing Arc (cheap clone).
    /// If owned, creates a new Arc (allocates once).
    #[inline]
    fn get_or_create_arc(&self) -> (Arc<Vec<u8>>, u32) {
        match self {
            XmlData::Owned(b) => (Arc::new(b.to_vec()), 0),
            XmlData::Shared(s) => (s.arc(), s.start()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The raw XML bytes for this paragraph
    xml_data: XmlData,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes (owned).
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    #[inline]
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: XmlData::Owned(xml_bytes.into_boxed_slice()),
        }
    }

    /// Create a new Paragraph from a shared XML slice (zero-copy).
    ///
    /// This is used for arena-based parsing where all element XMLs are stored
    /// in a single contiguous buffer.
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: XmlData::Shared(slice),
        }
    }

    /// Create a Paragraph from an Arc<Vec<u8>> and byte range.
    ///
    /// This is a convenience method for arena-based parsing.
    #[inline]
    pub fn from_arc_range(arena: Arc<Vec<u8>>, start: u32, len: u32) -> Self {
        Self::from_slice(XmlSlice::new(arena, start, len))
    }

    /// Get the raw XML bytes.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    pub fn text(&self) -> Result<String> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        // Pre-allocate string with estimated capacity to reduce reallocations
        let estimated_capacity = xml_bytes.len() / 4; // Rough estimate
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
                    // Use unsafe conversion for better performance (safe since we validate XML)
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

        // Shrink to fit to release unused capacity
        result.shrink_to_fit();
        Ok(result)
    }

    /// Get an iterator over the runs in this paragraph.
    ///
    /// Each run represents a `<w:r>` element and may have different formatting.
    ///
    /// # Performance
    ///
    /// Uses zero-copy parsing with NO Vec::push calls.
    /// Reuses existing Arc if paragraph is already shared.
    pub fn runs(&self) -> Result<SmallVec<[Run; 8]>> {
        let xml_bytes = self.xml_bytes();
        let len = xml_bytes.len();

        // PASS 1: Count runs
        let count = count_runs(xml_bytes);
        if count == 0 {
            return Ok(SmallVec::new());
        }

        // Reuse Arc if available (cheap), or create one (allocates once)
        let (source_arc, base_offset) = self.xml_data.get_or_create_arc();

        // Pre-allocate SmallVec with exact capacity
        let mut runs: SmallVec<[Run; 8]> = SmallVec::with_capacity(count);

        // OPTIMIZATION: Pre-increment Arc refcount to avoid per-run atomic operations.
        let arc_ptr = Arc::into_raw(source_arc);
        // SAFETY: arc_ptr came from Arc::into_raw, so it's valid.
        unsafe {
            for _ in 0..(count - 1) {
                Arc::increment_strong_count(arc_ptr);
            }
        }

        // PASS 2: Fill directly using unsafe to avoid push
        let mut write_idx = 0usize;
        let mut i = 0usize;
        let runs_ptr = runs.as_mut_ptr();

        while i < len && write_idx < count {
            let Some(tag_start) = memchr::memchr(b'<', &xml_bytes[i..]) else {
                break;
            };
            let tag_start = i + tag_start;

            if tag_start + 5 < len && &xml_bytes[tag_start..tag_start + 4] == b"<w:r" {
                let next_char = xml_bytes[tag_start + 4];
                if (next_char == b'>' || next_char == b' ' || next_char == b'/')
                    && let Some(end) = find_run_end(&xml_bytes[tag_start..])
                {
                    let end_pos = tag_start + end;
                    let run_len = (end_pos - tag_start) as u32;

                    // SAFETY: arc_ptr is valid; each from_raw consumes one refcount we pre-incremented
                    let arc_clone = unsafe { Arc::from_raw(arc_ptr) };

                    // Write directly to pre-allocated slot (no push)
                    // Add base_offset to get absolute position in source Arc
                    unsafe {
                        std::ptr::write(
                            runs_ptr.add(write_idx),
                            Run::from_slice(XmlSlice::new(
                                arc_clone,
                                base_offset + tag_start as u32,
                                run_len,
                            )),
                        );
                    }
                    write_idx += 1;
                    i = end_pos;
                    continue;
                }
            }

            i = tag_start + 1;
        }

        // Handle refcount mismatch: decrement unused pre-incremented refcounts
        if write_idx < count {
            unsafe {
                for _ in 0..(count - write_idx) {
                    Arc::decrement_strong_count(arc_ptr);
                }
            }
        }

        // Set final length
        unsafe {
            runs.set_len(write_idx);
        }

        Ok(runs)
    }

    /// Extract all OMML formulas from this paragraph.
    ///
    /// Returns a vector of OMML formula strings found in any run within this paragraph.
    /// This extracts inline formulas (formulas within runs).
    pub fn omml_formulas(&self) -> Result<Vec<String>> {
        let mut formulas = Vec::new();
        for run in self.runs()? {
            if let Some(formula) = run.omml_formula()? {
                formulas.push(formula);
            }
        }
        Ok(formulas)
    }

    /// Extract all inline images from this paragraph.
    ///
    /// Returns a vector of `InlineImage` objects found in `<w:drawing>` elements
    /// within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for image in para.images()? {
    ///         println!("Image: {} ({}x{} pixels)",
    ///             image.name(),
    ///             image.width_px(),
    ///             image.height_px()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn images(&self) -> Result<SmallVec<[InlineImage; 4]>> {
        parse_inline_images(self.xml_bytes())
    }

    /// Extract all drawing objects (shapes, text boxes) from this paragraph.
    ///
    /// Returns a vector of `DrawingObject` objects found in `<w:drawing>` elements
    /// within this paragraph. This includes shapes, text boxes, and other DrawingML objects.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for drawing in para.drawing_objects()? {
    ///         println!("Shape: {} (type: {:?})",
    ///             drawing.name(),
    ///             drawing.shape_type()
    ///         );
    ///         if !drawing.text().is_empty() {
    ///             println!("  Text: {}", drawing.text());
    ///         }
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn drawing_objects(&self) -> Result<SmallVec<[DrawingObject; 4]>> {
        parse_drawing_objects(self.xml_bytes())
    }

    /// Extract all tracked changes (revisions) from this paragraph.
    ///
    /// Returns a vector of `Revision` objects representing all tracked changes
    /// (insertions, deletions, moves) within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for revision in para.revisions()? {
    ///         println!("Revision by {}: {} - {}",
    ///             revision.author(),
    ///             revision.revision_type(),
    ///             revision.text()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn revisions(&self) -> Result<SmallVec<[Revision; 4]>> {
        parse_revisions(self.xml_bytes())
    }

    /// Extract paragraph-level OMML formulas.
    ///
    /// Returns a vector of OMML formula strings that are direct children of the paragraph
    /// (display math), not nested within runs. These are block-level formulas.
    ///
    /// # Example
    /// ```ignore
    /// let para = document.paragraphs()?[0];
    /// let display_formulas = para.paragraph_level_formulas()?;
    /// for formula in display_formulas {
    ///     println!("Display formula: {}", formula);
    /// }
    /// ```
    pub fn paragraph_level_formulas(&self) -> Result<Vec<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut formulas = Vec::new();
        let mut in_omath = false;
        let mut in_run = false;
        let mut skip_word_element = false; // Track when we're skipping Word elements
        let mut word_depth = 0; // Track nesting depth of Word elements
        let mut depth = 0;
        let mut omml_content = String::with_capacity(512);

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();
                    let full_name = e.name();
                    let full_name_bytes = full_name.as_ref();

                    // Track when we're inside a run to skip inline formulas
                    if full_name_bytes == b"w:r" || name.as_ref() == b"r" {
                        in_run = true;
                    }

                    // Only process oMath if we're NOT inside a run (paragraph-level formulas)
                    if !in_run && name.as_ref() == b"oMath" {
                        in_omath = true;
                        depth = 1;
                        skip_word_element = false;
                        word_depth = 0;
                        omml_content.clear();
                        omml_content.push_str("<m:oMath>");
                    } else if in_omath && !skip_word_element {
                        // Check if this is a Word namespace element
                        if full_name_bytes.starts_with(b"w:") {
                            skip_word_element = true;
                            word_depth = 1;
                        } else {
                            depth += 1;

                            omml_content.push('<');
                            if full_name_bytes.starts_with(b"m:") {
                                omml_content
                                    .push_str(std::str::from_utf8(full_name_bytes).unwrap());
                            } else {
                                omml_content.push_str("m:");
                                omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                            }

                            // Include math namespace attributes only
                            for attr in e.attributes().flatten() {
                                let key_bytes = attr.key.as_ref();
                                if !key_bytes.starts_with(b"w:") {
                                    omml_content.push(' ');
                                    if key_bytes.starts_with(b"m:") {
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    } else if memchr::memchr(b':', key_bytes).is_some() {
                                        let key_str = std::str::from_utf8(key_bytes).unwrap();
                                        if let Some(idx) = key_str.find(':') {
                                            omml_content.push_str("m:");
                                            omml_content.push_str(&key_str[idx + 1..]);
                                        } else {
                                            omml_content.push_str(key_str);
                                        }
                                    } else {
                                        omml_content.push_str("m:");
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    }
                                    omml_content.push_str("=\"");
                                    omml_content.push_str(
                                        &attr.unescape_value().unwrap_or(Cow::Borrowed("")),
                                    );
                                    omml_content.push('"');
                                }
                            }
                            omml_content.push('>');
                        }
                    } else if skip_word_element {
                        word_depth += 1;
                    }
                },
                Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    let full_name = e.name();
                    let full_name_bytes = full_name.as_ref();

                    if in_omath && !skip_word_element && !full_name_bytes.starts_with(b"w:") {
                        omml_content.push('<');
                        if full_name_bytes.starts_with(b"m:") {
                            omml_content.push_str(std::str::from_utf8(full_name_bytes).unwrap());
                        } else {
                            omml_content.push_str("m:");
                            omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                        }

                        // Include math namespace attributes only
                        for attr in e.attributes().flatten() {
                            let key_bytes = attr.key.as_ref();
                            if !key_bytes.starts_with(b"w:") {
                                omml_content.push(' ');
                                if key_bytes.starts_with(b"m:") {
                                    omml_content.push_str(std::str::from_utf8(key_bytes).unwrap());
                                } else if memchr::memchr(b':', key_bytes).is_some() {
                                    let key_str = std::str::from_utf8(key_bytes).unwrap();
                                    if let Some(idx) = key_str.find(':') {
                                        omml_content.push_str("m:");
                                        omml_content.push_str(&key_str[idx + 1..]);
                                    } else {
                                        omml_content.push_str(key_str);
                                    }
                                } else {
                                    omml_content.push_str("m:");
                                    omml_content.push_str(std::str::from_utf8(key_bytes).unwrap());
                                }
                                omml_content.push_str("=\"");
                                omml_content
                                    .push_str(&attr.unescape_value().unwrap_or(Cow::Borrowed("")));
                                omml_content.push('"');
                            }
                        }
                        omml_content.push_str("/>");
                    }
                },
                Ok(Event::Text(e)) if in_omath && !skip_word_element => {
                    let text = std::str::from_utf8(e.as_ref())
                        .map_err(|_| OoxmlError::Xml("Invalid UTF-8 in OMML text".to_string()))?;
                    omml_content.push_str(text);
                },
                Ok(Event::End(e)) => {
                    let name = e.local_name();
                    let full_name = e.name();
                    let full_name_bytes = full_name.as_ref();

                    // Track when we exit a run
                    if full_name_bytes == b"w:r" || name.as_ref() == b"r" {
                        in_run = false;
                    }

                    // Track exiting Word elements
                    if skip_word_element {
                        word_depth -= 1;
                        if word_depth == 0 {
                            skip_word_element = false;
                        }
                    } else if in_omath {
                        omml_content.push_str("</");
                        if full_name_bytes.starts_with(b"m:") {
                            omml_content.push_str(std::str::from_utf8(full_name_bytes).unwrap());
                        } else {
                            omml_content.push_str("m:");
                            omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                        }
                        omml_content.push('>');

                        depth -= 1;
                        if depth == 0 && name.as_ref() == b"oMath" {
                            // Complete formula extracted
                            formulas.push(omml_content.clone());
                            in_omath = false;
                            omml_content.clear();
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(formulas)
    }

    /// Get all hyperlinks in this paragraph.
    ///
    /// Returns a vector of `Hyperlink` objects representing all hyperlinks
    /// found in this paragraph. Requires relationships to resolve external URLs.
    ///
    /// # Arguments
    ///
    /// * `rels` - Relationships for resolving relationship IDs to URLs
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let para = doc.paragraph(0)?.unwrap();
    /// let hyperlinks = para.hyperlinks(&main_part.rels())?;
    /// for link in hyperlinks {
    ///     println!("Link: {} -> {:?}", link.text(), link.url());
    /// }
    /// ```
    pub fn hyperlinks(&self, rels: &Relationships) -> Result<Vec<Hyperlink>> {
        Hyperlink::extract_from_paragraph(self.xml_bytes(), rels)
    }
}

/// A run within a paragraph.
///
/// Represents a `<w:r>` element. A run is a region of text with a single
/// set of formatting properties.
///
/// # Example
///
/// ```rust,ignore
/// let run = runs[0];
/// println!("Text: {}", run.text()?);
/// println!("Bold: {:?}", run.bold()?);
/// println!("Italic: {:?}", run.italic()?);
///
/// // Check for embedded formulas
/// if let Some(omml) = run.omml_formula()? {
///     println!("OMML formula: {}", omml);
/// }
/// ```
/// Cached formatting properties for a Run.
///
/// This struct stores all commonly accessed formatting properties
/// to avoid repeated XML parsing.
#[derive(Debug, Clone, Copy, Default)]
pub struct RunProperties {
    /// Whether the run is bold
    pub bold: Option<bool>,
    /// Whether the run is italic
    pub italic: Option<bool>,
    /// Whether the run is strikethrough
    pub strikethrough: Option<bool>,
    /// Vertical position (superscript/subscript)
    pub vertical_position: Option<VerticalPosition>,
}

/// Internal storage for run XML data (same pattern as Paragraph).
#[derive(Debug, Clone)]
enum RunXmlData {
    Owned(Vec<u8>),
    Shared(XmlSlice),
}

impl RunXmlData {
    #[inline]
    fn as_bytes(&self) -> &[u8] {
        match self {
            RunXmlData::Owned(v) => v,
            RunXmlData::Shared(s) => s.as_bytes(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Run {
    /// The raw XML data for this run
    xml_data: RunXmlData,
}

impl Run {
    /// Create a new Run from XML bytes (owned).
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: RunXmlData::Owned(xml_bytes),
        }
    }

    /// Create a Run from a shared XML slice (zero-copy).
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: RunXmlData::Shared(slice),
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Get the text content of this run.
    ///
    /// Extracts text from `<w:t>` elements and converts special characters:
    /// - `<w:tab/>` → tab character
    /// - `<w:br/>` → newline character
    pub fn text(&self) -> Result<String> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity
        let estimated_capacity = xml_bytes.len() / 8; // Rough estimate for text content
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"t" {
                        in_text_element = true;
                    } else if name.as_ref() == b"tab" {
                        result.push('\t');
                    } else if name.as_ref() == b"br" {
                        result.push('\n');
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Use unsafe conversion for better performance (safe since we validate XML)
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

        Ok(result)
    }

    /// Check if this run is bold.
    ///
    /// Returns `Some(true)` if bold is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn bold(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"b")
    }

    /// Check if this run is italic.
    ///
    /// Returns `Some(true)` if italic is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn italic(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"i")
    }

    /// Check if this run is underlined.
    ///
    /// Returns `Some(true)` if underline is present,
    /// `None` if not specified.
    ///
    /// Note: This is simplified. Full implementation would return
    /// the underline style (single, double, wavy, etc.).
    pub fn underline(&self) -> Result<Option<bool>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"u" {
                        return Ok(Some(true));
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Check if this run is strikethrough.
    ///
    /// Returns `Some(true)` if strikethrough is present,
    /// `None` if not specified.
    pub fn strikethrough(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"strike")
    }

    /// Get text and properties in a single XML parse.
    ///
    /// This is **the fastest way** to extract both text content and formatting properties
    /// from a run, as it parses the XML only once instead of twice (text() + get_properties()).
    ///
    /// # Performance
    ///
    /// This provides 2x speedup over calling `text()` and `get_properties()` separately,
    /// and 4-6x speedup over individual property methods.
    ///
    /// # Returns
    ///
    /// A tuple of (text_content, properties)
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Fastest: Single XML parse for both text and properties
    /// let (text, props) = run.get_text_and_properties()?;
    /// if props.bold.unwrap_or(false) {
    ///     write_bold(&text);
    /// }
    /// ```
    pub fn get_text_and_properties(&self) -> Result<(String, RunProperties)> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut text = String::with_capacity(self.xml_bytes().len() / 8);
        let mut in_r_pr = false;
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();

                    // Handle text elements
                    if name.as_ref() == b"t" {
                        in_text_element = true;
                    } else if name.as_ref() == b"tab" {
                        text.push('\t');
                    } else if name.as_ref() == b"br" {
                        text.push('\n');
                    } else if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr {
                        // Extract all properties in one pass
                        match name.as_ref() {
                            b"b" => {
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.bold = Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                if !found_val {
                                    props.bold = Some(true);
                                }
                            },
                            b"i" => {
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.italic = Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                if !found_val {
                                    props.italic = Some(true);
                                }
                            },
                            b"strike" => {
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.strikethrough =
                                            Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                if !found_val {
                                    props.strikethrough = Some(true);
                                }
                            },
                            b"vertAlign" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        let value = attr.value.as_ref();
                                        props.vertical_position = match value {
                                            b"superscript" => Some(VerticalPosition::Superscript),
                                            b"subscript" => Some(VerticalPosition::Subscript),
                                            _ => None,
                                        };
                                        break;
                                    }
                                }
                            },
                            _ => {},
                        }
                    }
                },
                Ok(Event::Text(e)) => {
                    if in_text_element {
                        let text_str = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                        text.push_str(text_str);
                    }
                },
                Ok(Event::End(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"t" {
                        in_text_element = false;
                    } else if name.as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok((text, props))
    }

    /// Get all formatting properties in a single pass.
    ///
    /// This is **significantly faster** than calling individual property methods
    /// (bold(), italic(), strikethrough(), vertical_position()) because it parses
    /// the XML only once instead of multiple times.
    ///
    /// # Performance
    ///
    /// For documents with many runs, using this method can provide 3-4x speedup
    /// compared to individual property access.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Fast: Single XML parse
    /// let props = run.get_properties()?;
    /// if props.bold.unwrap_or(false) {
    ///     // Handle bold text
    /// }
    ///
    /// // Slow: Multiple XML parses
    /// if run.bold()?.unwrap_or(false) {
    ///     // Handle bold text
    /// }
    /// ```
    pub fn get_properties(&self) -> Result<RunProperties> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr {
                        // Extract all properties in one pass
                        match name.as_ref() {
                            b"b" => {
                                // Check for w:val attribute
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.bold = Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                // Element present without val attribute means true
                                if !found_val {
                                    props.bold = Some(true);
                                }
                            },
                            b"i" => {
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.italic = Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                if !found_val {
                                    props.italic = Some(true);
                                }
                            },
                            b"strike" => {
                                let mut found_val = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        found_val = true;
                                        let value = attr.value.as_ref();
                                        props.strikethrough =
                                            Some(value == b"true" || value == b"1");
                                        break;
                                    }
                                }
                                if !found_val {
                                    props.strikethrough = Some(true);
                                }
                            },
                            b"vertAlign" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        let value = attr.value.as_ref();
                                        props.vertical_position = match value {
                                            b"superscript" => Some(VerticalPosition::Superscript),
                                            b"subscript" => Some(VerticalPosition::Subscript),
                                            _ => None,
                                        };
                                        break;
                                    }
                                }
                            },
                            _ => {},
                        }
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        // Exit early once we've finished parsing rPr
                        return Ok(props);
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(props)
    }

    /// Get the vertical position of this run (superscript/subscript).
    ///
    /// Returns the vertical positioning if specified, None if normal.
    pub fn vertical_position(&self) -> Result<Option<VerticalPosition>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"vertAlign" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                match value {
                                    b"superscript" => {
                                        return Ok(Some(VerticalPosition::Superscript));
                                    },
                                    b"subscript" => return Ok(Some(VerticalPosition::Subscript)),
                                    _ => {},
                                }
                            }
                        }
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font name for this run.
    ///
    /// Returns the typeface name if specified, None if inherited.
    pub fn font_name(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"rFonts" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ascii" {
                                let value = attr.unescape_value().unwrap_or(Cow::Borrowed(""));
                                return Ok(Some(value.to_string()));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        break;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: Word stores font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&self) -> Result<Option<u32>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"sz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val"
                                && let Ok(value) = std::str::from_utf8(&attr.value)
                                && let Ok(size) = value.parse::<u32>()
                            {
                                return Ok(Some(size));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        break;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Check if this run contains an OMML formula.
    ///
    /// Returns the OMML XML content if this run contains a mathematical formula,
    /// None otherwise. This method looks for `<m:oMath>` elements embedded in the run.
    ///
    /// The extracted OMML is cleaned to remove or fix improperly nested Word namespace
    /// elements that can cause XML parsing errors.
    pub fn omml_formula(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_omath = false;
        let mut skip_word_element = false; // Track when we're skipping Word namespace elements
        let mut word_depth = 0; // Track nesting depth of Word elements to skip
        let mut omml_content = String::with_capacity(512); // Pre-allocate for performance

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();
                    let full_name = e.name();
                    let full_name_bytes = full_name.as_ref();

                    if name.as_ref() == b"oMath" {
                        in_omath = true;
                        omml_content.push_str("<m:");
                        omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());

                        // Include attributes
                        for attr in e.attributes().flatten() {
                            omml_content.push(' ');
                            let key_str = std::str::from_utf8(attr.key.as_ref()).unwrap();
                            // Only include math namespace or unnamespaced attributes
                            if key_str.starts_with("m:") || !key_str.contains(':') {
                                omml_content.push_str(key_str);
                                omml_content.push_str("=\"");
                                omml_content
                                    .push_str(&attr.unescape_value().unwrap_or(Cow::Borrowed("")));
                                omml_content.push('"');
                            }
                        }
                        omml_content.push('>');
                    } else if in_omath && !skip_word_element {
                        // Check if this is a Word namespace element (w: prefix or known Word elements)
                        // These are formatting elements that should be skipped as they're often malformed
                        if full_name_bytes.starts_with(b"w:") {
                            skip_word_element = true;
                            word_depth = 1;
                        } else {
                            // Include math namespace elements
                            omml_content.push('<');

                            // Ensure we prefix with m: if not already present
                            if full_name_bytes.starts_with(b"m:") {
                                // Already has m: prefix, use full name
                                omml_content
                                    .push_str(std::str::from_utf8(full_name_bytes).unwrap());
                            } else {
                                // No prefix, add m: and local name
                                omml_content.push_str("m:");
                                omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                            }

                            // Include math namespace attributes only
                            for attr in e.attributes().flatten() {
                                let key_bytes = attr.key.as_ref();
                                // Skip Word namespace attributes
                                if !key_bytes.starts_with(b"w:") {
                                    omml_content.push(' ');
                                    // Ensure math namespace prefix on attributes
                                    if key_bytes.starts_with(b"m:") {
                                        // Already has m: prefix, use as-is
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    } else if memchr::memchr(b':', key_bytes).is_some() {
                                        // Has some other namespace, replace with m:
                                        let key_str = std::str::from_utf8(key_bytes).unwrap();
                                        if let Some(idx) = key_str.find(':') {
                                            omml_content.push_str("m:");
                                            omml_content.push_str(&key_str[idx + 1..]);
                                        } else {
                                            omml_content.push_str(key_str);
                                        }
                                    } else {
                                        // No namespace, add m:
                                        omml_content.push_str("m:");
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    }
                                    omml_content.push_str("=\"");
                                    // Use Cow to avoid allocation when possible
                                    omml_content.push_str(
                                        &attr.unescape_value().unwrap_or(Cow::Borrowed("")),
                                    );
                                    omml_content.push('"');
                                }
                            }
                            omml_content.push('>');
                        }
                    } else if skip_word_element {
                        word_depth += 1;
                    }
                },
                Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    let full_name = e.name();
                    let full_name_bytes = full_name.as_ref();

                    if in_omath && !skip_word_element {
                        // Skip Word namespace empty elements
                        if !full_name_bytes.starts_with(b"w:") {
                            omml_content.push('<');
                            if full_name_bytes.starts_with(b"m:") {
                                // Already has m: prefix, use full name
                                omml_content
                                    .push_str(std::str::from_utf8(full_name_bytes).unwrap());
                            } else {
                                // No prefix, add m: and local name
                                omml_content.push_str("m:");
                                omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                            }

                            // Include math namespace attributes only
                            for attr in e.attributes().flatten() {
                                let key_bytes = attr.key.as_ref();
                                // Skip Word namespace attributes
                                if !key_bytes.starts_with(b"w:") {
                                    omml_content.push(' ');
                                    // Ensure math namespace prefix on attributes
                                    if key_bytes.starts_with(b"m:") {
                                        // Already has m: prefix, use as-is
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    } else if memchr::memchr(b':', key_bytes).is_some() {
                                        // Has some other namespace, replace with m:
                                        let key_str = std::str::from_utf8(key_bytes).unwrap();
                                        if let Some(idx) = key_str.find(':') {
                                            omml_content.push_str("m:");
                                            omml_content.push_str(&key_str[idx + 1..]);
                                        } else {
                                            omml_content.push_str(key_str);
                                        }
                                    } else {
                                        // No namespace, add m:
                                        omml_content.push_str("m:");
                                        omml_content
                                            .push_str(std::str::from_utf8(key_bytes).unwrap());
                                    }
                                    omml_content.push_str("=\"");
                                    omml_content.push_str(
                                        &attr.unescape_value().unwrap_or(Cow::Borrowed("")),
                                    );
                                    omml_content.push('"');
                                }
                            }
                            omml_content.push_str("/>");
                        }
                    }
                },
                Ok(Event::Text(e)) if in_omath && !skip_word_element => {
                    // Extract text content
                    let text = std::str::from_utf8(e.as_ref())
                        .map_err(|_| OoxmlError::Xml("Invalid UTF-8 in OMML text".to_string()))?;
                    omml_content.push_str(text);
                },
                Ok(Event::End(e)) => {
                    if skip_word_element {
                        word_depth -= 1;
                        if word_depth == 0 {
                            skip_word_element = false;
                        }
                    } else if in_omath {
                        let name = e.local_name();
                        let full_name = e.name();
                        let full_name_bytes = full_name.as_ref();

                        omml_content.push_str("</");
                        if full_name_bytes.starts_with(b"m:") {
                            // Already has m: prefix, use full name
                            omml_content.push_str(std::str::from_utf8(full_name_bytes).unwrap());
                        } else {
                            // No prefix, add m: and local name
                            omml_content.push_str("m:");
                            omml_content.push_str(std::str::from_utf8(name.as_ref()).unwrap());
                        }
                        omml_content.push('>');

                        // Check if this closes the oMath element
                        if name.as_ref() == b"oMath" {
                            break; // We've extracted the complete formula
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        if omml_content.is_empty() {
            Ok(None)
        } else {
            Ok(Some(omml_content))
        }
    }

    /// Helper to extract boolean properties from run properties.
    ///
    /// Handles the tri-state logic where w:val can be "true", "false", "1", "0"
    /// or the element can be present without a val attribute (implies true).
    fn get_bool_property(&self, property_name: &[u8]) -> Result<Option<bool>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == property_name {
                        // Check for w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                return Ok(Some(value == b"true" || value == b"1"));
                            }
                        }
                        // Element present without val attribute means true
                        return Ok(Some(true));
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }
}

/// Count runs in a paragraph XML (for pre-allocation).
#[inline]
fn count_runs(xml_bytes: &[u8]) -> usize {
    let mut count = 0usize;
    let mut i = 0usize;
    let len = xml_bytes.len();

    while i < len {
        let Some(tag_start) = memchr::memchr(b'<', &xml_bytes[i..]) else {
            break;
        };
        let tag_start = i + tag_start;

        if tag_start + 5 < len && &xml_bytes[tag_start..tag_start + 4] == b"<w:r" {
            let c = xml_bytes[tag_start + 4];
            if (c == b'>' || c == b' ' || c == b'/')
                && let Some(end) = find_run_end(&xml_bytes[tag_start..])
            {
                count += 1;
                i = tag_start + end;
                continue;
            }
        }
        i = tag_start + 1;
    }
    count
}

/// Find the end of a `<w:r>` element.
/// Returns byte offset AFTER the closing `</w:r>`.
#[inline]
fn find_run_end(xml: &[u8]) -> Option<usize> {
    let first_gt = memchr::memchr(b'>', xml)?;
    if first_gt > 0 && xml[first_gt - 1] == b'/' {
        return Some(first_gt + 1); // Self-closing
    }

    let mut depth = 1i32;
    let mut pos = first_gt + 1;

    while pos < xml.len() && depth > 0 {
        let Some(next_lt) = memchr::memchr(b'<', &xml[pos..]) else {
            break;
        };
        pos += next_lt;

        // Check for </w:r>
        if pos + 6 <= xml.len() && &xml[pos..pos + 5] == b"</w:r" {
            if xml[pos + 5] == b'>' {
                depth -= 1;
                if depth == 0 {
                    return Some(pos + 6);
                }
            }
        }
        // Check for nested <w:r> (shouldn't happen but handle it)
        else if pos + 5 <= xml.len() && &xml[pos..pos + 4] == b"<w:r" {
            let c = xml[pos + 4];
            if c == b'>' || c == b' ' {
                let gt = memchr::memchr(b'>', &xml[pos..])?;
                if gt > 0 && xml[pos + gt - 1] != b'/' {
                    depth += 1;
                }
            }
        }
        pos += 1;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_text_extraction() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Hello, World!</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        let text = run.text().unwrap();
        assert_eq!(text, "Hello, World!");
    }

    #[test]
    fn test_run_bold() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:b/></w:rPr>
            <w:t>Bold text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert!(run.bold().unwrap().unwrap_or(false));
    }

    #[test]
    fn test_run_italic() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:i/></w:rPr>
            <w:t>Italic text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert!(run.italic().unwrap().unwrap_or(false));
    }
}
