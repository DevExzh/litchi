#![expect(
    clippy::needless_pass_by_value,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Run text, break, font, and aggregate-property facade.

use crate::UnderlineStyle;
use crate::error::{Error, Result};
use crate::font::{OpenType, Snapshot as OpenTypeSnapshot};
use crate::run_effects::Effects;
use litchi_core::VerticalPosition;
use litchi_ooxml_common::xml::{decode_xml_reference, extract_omml_formulas};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use smallvec::SmallVec;
use std::borrow::Cow;

use super::super::model::{
    Run, RunBreak, RunBreakClear, RunBreakType, RunProperties, RunUnderline,
};
use super::run_properties::{parse_run_underline, update_run_properties};
use super::text::extract_word_text;
use super::xml::is_on;

impl Run {
    /// Get the text content of this run.
    ///
    /// Extracts text from `<w:t>` elements and converts special characters:
    /// - `<w:tab/>` → tab character
    /// - `<w:br/>` → newline character
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Parse all explicit break elements in this run, preserving type and clear behavior.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn breaks(&self) -> Result<SmallVec<[RunBreak; 2]>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);
        let mut breaks = SmallVec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(e) | Event::Empty(e)) if e.local_name().as_ref() == b"br" => {
                    let mut run_break = RunBreak::default();
                    for attribute in e.attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        match attribute.key.local_name().as_ref() {
                            b"type" => {
                                run_break.break_type = match value.as_ref() {
                                    "textWrapping" => RunBreakType::TextWrapping,
                                    "page" => RunBreakType::Page,
                                    "column" => RunBreakType::Column,
                                    _ => {
                                        return Err(Error::InvalidFormat(format!(
                                            "invalid Word run break type '{value}'"
                                        )));
                                    },
                                };
                            },
                            b"clear" => {
                                run_break.clear = match value.as_ref() {
                                    "none" => RunBreakClear::None,
                                    "left" => RunBreakClear::Left,
                                    "right" => RunBreakClear::Right,
                                    "all" => RunBreakClear::All,
                                    _ => {
                                        return Err(Error::InvalidFormat(format!(
                                            "invalid Word run break clear value '{value}'"
                                        )));
                                    },
                                };
                            },
                            _ => {},
                        }
                    }
                    breaks.push(run_break);
                },
                Ok(Event::Eof) => break,
                Err(error) => return Err(Error::Xml(error.to_string())),
                _ => {},
            }
        }
        Ok(breaks)
    }

    /// Count layout-engine page-break hints in this run.
    ///
    /// `<w:lastRenderedPageBreak>` is not an authored break; it records where Word last
    /// paginated content, so it is intentionally exposed separately from [`Self::breaks`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn last_rendered_page_break_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);
        let mut count = 0usize;
        loop {
            match reader.read_event() {
                Ok(Event::Start(e) | Event::Empty(e))
                    if e.local_name().as_ref() == b"lastRenderedPageBreak" =>
                {
                    count = count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(
                            "too many rendered page break markers in one run".to_string(),
                        )
                    })?;
                },
                Ok(Event::Eof) => break,
                Err(error) => return Err(Error::Xml(error.to_string())),
                _ => {},
            }
        }
        Ok(count)
    }

    /// Check if this run is bold.
    ///
    /// Returns `Some(true)` if bold is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn bold(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"b")
    }

    /// Check if this run is italic.
    ///
    /// Returns `Some(true)` if italic is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn italic(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"i")
    }

    /// Check whether this run has a direct underline enabled.
    ///
    /// Returns `Some(false)` for an explicit `w:val="none"` and `None` when the
    /// run has no direct underline property and therefore inherits from styles.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn underline(&self) -> Result<Option<bool>> {
        Ok(self
            .underline_style()?
            .map(|style| style != UnderlineStyle::None))
    }

    /// Return the direct underline pattern, including an explicit
    /// [`UnderlineStyle::None`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn underline_style(&self) -> Result<Option<UnderlineStyle>> {
        Ok(self
            .underline_formatting()?
            .map(|underline| underline.style))
    }

    /// Return the complete direct underline formatting for this run.
    ///
    /// This preserves all `CT_Underline` fields without resolving style or
    /// theme inheritance. A present `<w:u/>` is interpreted as a single
    /// underline for compatibility with documents emitted by Word processors.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn underline_formatting(&self) -> Result<Option<RunUnderline>> {
        parse_run_underline(self.xml_bytes())
    }

    /// Return the typed Word 2010 visual effects attached directly to this run.
    ///
    /// The result is a detached semantic snapshot. Unsupported direct
    /// extension children remain bounded and ordered as
    /// [`crate::run_effects::OpaqueExtension`] values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn effects(&self) -> Result<Effects> {
        Effects::parse(self.xml_bytes())
    }

    /// Read the typed Word 2010 OpenType features attached directly to this run.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn open_type(&self) -> Result<OpenType> {
        OpenType::parse(self.xml_bytes())
    }

    /// Capture a source-preserving OpenType snapshot for an isolated edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn open_type_snapshot(&self) -> Result<OpenTypeSnapshot> {
        OpenTypeSnapshot::from_xml(self.xml_bytes().to_vec())
    }

    /// Replace the modeled OpenType features while preserving every other run
    /// child and unknown extension byte.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_open_type(&mut self, value: OpenType) -> Result<&mut Self> {
        let rewritten = crate::font::open_type::rewrite(self.xml_bytes(), &value)?;
        if rewritten.as_slice() != self.xml_bytes() {
            self.replace_xml(rewritten);
        }
        Ok(self)
    }

    /// Check if this run is strikethrough.
    ///
    /// Returns `Some(true)` if strikethrough is present,
    /// `None` if not specified.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn strikethrough(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"strike")
    }

    /// Get text and properties in a single XML parse.
    ///
    /// This is **the fastest way** to extract both text content and formatting properties
    /// from a run, as it parses the XML only once instead of twice (`text()` + `get_properties()`).
    ///
    /// # Performance
    ///
    /// This provides 2x speedup over calling `text()` and `get_properties()` separately,
    /// and 4-6x speedup over individual property methods.
    ///
    /// # Returns
    ///
    /// A tuple of (`text_content`, properties)
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn get_text_and_properties(&self) -> Result<(String, RunProperties)> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(false);

        let mut props = RunProperties::default();
        let mut text = String::with_capacity(self.xml_bytes().len() / 8);
        let mut in_r_pr = false;
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();

                    if name.as_ref() == b"t" {
                        in_text_element = true;
                    } else if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else {
                        match name.as_ref() {
                            b"tab" => text.push('\t'),
                            b"br" | b"cr" => text.push('\n'),
                            b"noBreakHyphen" => text.push('\u{2011}'),
                            b"softHyphen" => text.push('\u{00ad}'),
                            _ => {},
                        }
                        if in_r_pr {
                            update_run_properties(&mut props, &e)?;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    match name.as_ref() {
                        b"tab" => text.push('\t'),
                        b"br" | b"cr" => text.push('\n'),
                        b"noBreakHyphen" => text.push('\u{2011}'),
                        b"softHyphen" => text.push('\u{00ad}'),
                        _ => {},
                    }
                    if in_r_pr {
                        update_run_properties(&mut props, &e)?;
                    }
                },
                Ok(Event::Text(content)) if in_text_element => {
                    let decoded = content
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    text.push_str(&unescaped);
                },
                Ok(Event::CData(content)) if in_text_element => {
                    let decoded = content
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    text.push_str(&decoded);
                },
                Ok(Event::GeneralRef(reference)) if in_text_element => {
                    text.push_str(&decode_xml_reference(&reference)?);
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
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        props.effects = Effects::parse(self.xml_bytes())?;
        props.open_type = OpenType::parse(self.xml_bytes())?;
        Ok((text, props))
    }

    /// Get all formatting properties in a single pass.
    ///
    /// This is **significantly faster** than calling individual property methods
    /// (`bold()`, `italic()`, `strikethrough()`, `vertical_position()`) because it parses
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn get_properties(&self) -> Result<RunProperties> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr {
                        update_run_properties(&mut props, &e)?;
                    }
                },
                Ok(Event::Empty(e)) if in_r_pr && e.local_name().as_ref() != b"rPr" => {
                    update_run_properties(&mut props, &e)?;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    // The effects snapshot is parsed from the same run fragment below.
                    break;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        props.effects = Effects::parse(self.xml_bytes())?;
        props.open_type = OpenType::parse(self.xml_bytes())?;
        Ok(props)
    }

    /// Get the vertical position of this run (superscript/subscript).
    ///
    /// Returns the vertical positioning if specified, None if normal.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn vertical_position(&self) -> Result<Option<VerticalPosition>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"vertAlign" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
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
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    in_r_pr = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font name for this run.
    ///
    /// Returns the typeface name if specified, None if inherited.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn font_name(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"rFonts" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"ascii" {
                                let value = attr
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .unwrap_or(Cow::Borrowed(""));
                                return Ok(Some(value.to_string()));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    break;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: Word stores font size in half-points (e.g., 24 = 12pt).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn font_size(&self) -> Result<Option<u32>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"sz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = std::str::from_utf8(&attr.value)
                                && let Ok(size) = value.parse::<u32>()
                            {
                                return Ok(Some(size));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    break;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
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
    /// The returned string preserves the exact source XML for the first formula in the run.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn omml_formula(&self) -> Result<Option<String>> {
        Ok(extract_omml_formulas(self.xml_bytes())?.into_iter().next())
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
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == property_name {
                        // Check for w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                return Ok(Some(is_on(value)));
                            }
                        }
                        // Element present without val attribute means true
                        return Ok(Some(true));
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    in_r_pr = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }
}
