/// Paragraph and Run structures for Word documents.
use crate::common::VerticalPosition;
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::borrow::Cow;

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
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The raw XML bytes for this paragraph
    xml_bytes: Vec<u8>,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes.
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate string with estimated capacity to reduce reallocations
        let estimated_capacity = self.xml_bytes.len() / 4; // Rough estimate
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(1024); // Reusable buffer (increased from 512)

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
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
    pub fn runs(&self) -> Result<SmallVec<[Run; 8]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of typically small run collections
        let mut runs = SmallVec::new();
        let mut current_run_xml = Vec::with_capacity(2048); // Pre-allocate for XML fragments (increased from 1024)
        let mut in_run = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(1024); // Reusable buffer (increased from 512)

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    // Check for w:r (word run) specifically using the full name
                    // This avoids confusion with m:r (math run) which appears in OMML formulas
                    let is_word_run = e.local_name().as_ref() == b"r"
                        && !in_run
                        && (e.name().as_ref() == b"w:r" || e.name().as_ref() == b"r");

                    if is_word_run {
                        in_run = true;
                        depth = 1;
                        current_run_xml.clear();
                        // Pre-allocate with estimated size for run XML
                        current_run_xml.reserve(512);

                        // Build opening tag more efficiently
                        current_run_xml.extend_from_slice(b"<w:r");
                        for attr in e.attributes().flatten() {
                            current_run_xml.push(b' ');
                            current_run_xml.extend_from_slice(attr.key.as_ref());
                            current_run_xml.extend_from_slice(b"=\"");
                            current_run_xml.extend_from_slice(&attr.value);
                            current_run_xml.push(b'"');
                        }
                        current_run_xml.push(b'>');
                    } else if in_run {
                        depth += 1;
                        current_run_xml.push(b'<');
                        current_run_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_run_xml.push(b' ');
                            current_run_xml.extend_from_slice(attr.key.as_ref());
                            current_run_xml.extend_from_slice(b"=\"");
                            current_run_xml.extend_from_slice(&attr.value);
                            current_run_xml.push(b'"');
                        }
                        current_run_xml.push(b'>');
                    }
                },
                Ok(Event::End(ref e)) => {
                    if in_run {
                        current_run_xml.extend_from_slice(b"</");
                        current_run_xml.extend_from_slice(e.name().as_ref());
                        current_run_xml.push(b'>');

                        // Only decrement depth and check for run end when we see w:r or r (without namespace)
                        // This prevents m:r (math run) from being mistaken for the end of w:r (word run)
                        let is_word_run_end = e.local_name().as_ref() == b"r"
                            && (e.name().as_ref() == b"w:r" || e.name().as_ref() == b"r");

                        depth -= 1;
                        if is_word_run_end && depth == 0 {
                            runs.push(Run::new(current_run_xml.clone()));
                            in_run = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_run => {
                    current_run_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_run => {
                    current_run_xml.push(b'<');
                    current_run_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_run_xml.push(b' ');
                        current_run_xml.extend_from_slice(attr.key.as_ref());
                        current_run_xml.extend_from_slice(b"=\"");
                        current_run_xml.extend_from_slice(&attr.value);
                        current_run_xml.push(b'"');
                    }
                    current_run_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut formulas = Vec::new();
        let mut in_omath = false;
        let mut in_run = false;
        let mut skip_word_element = false; // Track when we're skipping Word elements
        let mut word_depth = 0; // Track nesting depth of Word elements
        let mut depth = 0;
        let mut omml_content = String::with_capacity(512);
        let mut buf = Vec::with_capacity(256);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(formulas)
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

#[derive(Debug, Clone)]
pub struct Run {
    /// The raw XML bytes for this run
    xml_bytes: Vec<u8>,
}

impl Run {
    /// Create a new Run from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the text content of this run.
    ///
    /// Extracts text from `<w:t>` elements and converts special characters:
    /// - `<w:tab/>` → tab character
    /// - `<w:br/>` → newline character
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity
        let estimated_capacity = self.xml_bytes.len() / 8; // Rough estimate for text content
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(512); // Reusable buffer (increased from 256)

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut text = String::with_capacity(self.xml_bytes.len() / 8);
        let mut in_r_pr = false;
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut in_r_pr = false;
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(props)
    }

    /// Get the vertical position of this run (superscript/subscript).
    ///
    /// Returns the vertical positioning if specified, None if normal.
    pub fn vertical_position(&self) -> Result<Option<VerticalPosition>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(None)
    }

    /// Get the font name for this run.
    ///
    /// Returns the typeface name if specified, None if inherited.
    pub fn font_name(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(None)
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: Word stores font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&self) -> Result<Option<u32>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_omath = false;
        let mut skip_word_element = false; // Track when we're skipping Word namespace elements
        let mut word_depth = 0; // Track nesting depth of Word elements to skip
        let mut omml_content = String::with_capacity(512); // Pre-allocate for performance
        let mut buf = Vec::with_capacity(256);

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }

        Ok(None)
    }
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
