//! Run types and implementation for DOCX documents.
use crate::OfficeMath;
use crate::error::{Error, Result};
use crate::run_effects::Effects;
use crate::run_symbols::Symbol;
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::UnderlineStyle;
// Import section types for PageNumberFormat
use super::revision::{RevisionMetadata, RevisionTextMode, RunPropertyChange};
use super::section::PageNumberFormat;

/// Run content type.
#[derive(Debug, Clone)]
pub enum RunContent {
    /// Plain text
    Text(String),
    /// Inline Office Math equation
    OfficeMath(OfficeMath),
    /// Word 2015 extended Unicode symbol
    Symbol(Symbol),
    /// Page number field
    PageNumber(PageNumberFormat),
    /// Page count field (total pages)
    PageCount,
    /// Tab character
    Tab,
    /// Page break
    PageBreak,
    /// Footnote reference
    FootnoteReference(u32),
    /// Endnote reference
    EndnoteReference(u32),
}

/// A mutable run.
///
/// Runs contain text and character formatting.
#[derive(Debug, Clone)]
pub struct MutableRun {
    /// Run content
    pub(crate) content: RunContent,
    /// Run properties
    pub(crate) properties: RunProperties,
    pub(crate) property_change: Option<RunPropertyChange>,
}

#[cfg(feature = "fonts")]
use litchi_fonts::{CollectGlyphs, GlyphMap, Request, Style as FontStyle};

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableRun {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        if let RunContent::Text(text) = &self.content
            && !text.is_empty()
        {
            // Use font name from properties or default to "Calibri" (common Office default)
            let font_name = self
                .properties
                .font_name
                .clone()
                .unwrap_or_else(|| "Calibri".to_string());
            let style = FontStyle::from_flags(
                self.properties.bold == Some(true),
                self.properties.italic == Some(true),
            );
            let bitmap = glyphs.entry(Request::new(font_name, style)).or_default();
            for c in text.chars() {
                bitmap.insert(c);
            }
        }
        glyphs
    }
}

impl MutableRun {
    pub(crate) fn new() -> Self {
        Self {
            content: RunContent::Text(String::new()),
            properties: RunProperties::default(),
            property_change: None,
        }
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: &str) {
        self.content = RunContent::Text(text.to_string());
    }

    /// Replace this run's content with an inline Office Math equation.
    pub fn set_office_math(&mut self, equation: OfficeMath) -> &mut Self {
        self.content = RunContent::OfficeMath(equation);
        self
    }

    /// Parse and replace this run's content with an inline Office Math equation.
    pub fn set_office_math_xml(&mut self, xml: impl Into<String>) -> Result<&mut Self> {
        let equation = OfficeMath::from_xml(xml)?;
        Ok(self.set_office_math(equation))
    }

    /// Get the text content.
    pub fn get_text(&self) -> String {
        match &self.content {
            RunContent::Text(s) => s.clone(),
            _ => String::new(),
        }
    }

    /// Make the text bold.
    pub fn bold(&mut self, bold: bool) -> &mut Self {
        self.properties.bold = Some(bold);
        self
    }

    /// Make the text italic.
    pub fn italic(&mut self, italic: bool) -> &mut Self {
        self.properties.italic = Some(italic);
        self
    }

    /// Set underline style.
    pub fn underline(&mut self, style: UnderlineStyle) -> &mut Self {
        self.properties.underline = Some(style);
        self
    }

    /// Set font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&mut self, size: u32) -> &mut Self {
        self.properties.font_size = Some(size);
        self
    }

    /// Set font name.
    pub fn font_name(&mut self, name: &str) -> &mut Self {
        self.properties.font_name = Some(name.to_string());
        self
    }

    /// Set text color using hex RGB (e.g., "FF0000" for red).
    pub fn color(&mut self, color: &str) -> &mut Self {
        self.properties.color = Some(color.to_string());
        self
    }

    /// Set text highlight color.
    pub fn highlight(&mut self, color: &str) -> &mut Self {
        self.properties.highlight = Some(color.to_string());
        self
    }

    /// Borrow the typed Word 2010 visual effects attached to this new run.
    pub fn effects(&self) -> &Effects {
        &self.properties.effects
    }

    /// Mutably borrow this run's visual effects for validated semantic edits.
    pub fn effects_mut(&mut self) -> &mut Effects {
        &mut self.properties.effects
    }

    /// Replace the visual effects for this run after validating their schema.
    pub fn set_effects(&mut self, effects: Effects) -> Result<&mut Self> {
        effects.validate()?;
        self.properties.effects = effects;
        Ok(self)
    }

    /// Add a line break.
    pub fn add_break(&mut self) -> &mut Self {
        self.properties.has_break = true;
        self
    }

    /// Add a page break.
    pub fn add_page_break(&mut self) -> &mut Self {
        self.content = RunContent::PageBreak;
        self
    }

    /// Add a page number field.
    pub fn add_page_number(&mut self, format: PageNumberFormat) -> &mut Self {
        self.content = RunContent::PageNumber(format);
        self
    }

    /// Add a page count field (total pages).
    pub fn add_page_count(&mut self) -> &mut Self {
        self.content = RunContent::PageCount;
        self
    }

    /// Add a tab character.
    pub fn add_tab(&mut self) -> &mut Self {
        self.content = RunContent::Tab;
        self
    }

    /// Add a footnote reference.
    pub fn add_footnote_reference(&mut self, id: u32) -> &mut Self {
        self.content = RunContent::FootnoteReference(id);
        self
    }

    /// Add an endnote reference.
    pub fn add_endnote_reference(&mut self, id: u32) -> &mut Self {
        self.content = RunContent::EndnoteReference(id);
        self
    }

    /// Whether this run is a reference to the note `id` of the requested
    /// kind (`w:footnoteReference`, ECMA-376 §17.11.14;
    /// `w:endnoteReference`, ECMA-376 §17.11.7).
    pub(crate) fn is_note_reference(&self, footnote: bool, id: u32) -> bool {
        match &self.content {
            RunContent::FootnoteReference(ref_id) => footnote && *ref_id == id,
            RunContent::EndnoteReference(ref_id) => !footnote && *ref_id == id,
            _ => false,
        }
    }

    /// Record the run properties that existed before this formatting revision.
    pub fn set_property_change(
        &mut self,
        metadata: RevisionMetadata,
        previous: &MutableRun,
    ) -> &mut Self {
        self.property_change = Some(RunPropertyChange::snapshot(metadata, previous));
        self
    }

    /// Validate content before this run is serialized inside inert conflict
    /// markup. The public run builder is intentionally broader than the
    /// conflict content model, so conflict wrappers must call this before
    /// emitting any XML.
    pub(crate) fn validate_passive_conflict_content(&self) -> Result<()> {
        match &self.content {
            RunContent::Text(_)
            | RunContent::Symbol(_)
            | RunContent::Tab
            | RunContent::PageBreak => Ok(()),
            RunContent::PageNumber(_) | RunContent::PageCount => Err(Error::InvalidFormat(
                "conflict runs cannot contain field instructions or computed page fields".into(),
            )),
            RunContent::FootnoteReference(_) | RunContent::EndnoteReference(_) => {
                Err(Error::InvalidFormat(
                    "conflict runs cannot contain cross-story note references".into(),
                ))
            },
            RunContent::OfficeMath(_) => Err(Error::InvalidFormat(
                "conflict runs cannot contain raw or foreign Office Math markup".into(),
            )),
        }
    }

    pub(crate) fn to_xml_mode(&self, xml: &mut String, mode: RevisionTextMode) -> Result<()> {
        xml.push_str("<w:r>");

        // Write run properties
        if self.properties.has_properties() || self.property_change.is_some() {
            self.properties.write_open(xml);

            if let Some(bold) = self.properties.bold
                && bold
            {
                xml.push_str("<w:b/>");
            }

            if let Some(italic) = self.properties.italic
                && italic
            {
                xml.push_str("<w:i/>");
            }

            if let Some(underline_style) = self.properties.underline {
                write!(xml, "<w:u w:val=\"{}\"/>", underline_style.as_str())
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(size) = self.properties.font_size {
                write!(xml, "<w:sz w:val=\"{}\"/>", size).map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(ref font_name) = self.properties.font_name {
                write!(
                    xml,
                    "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
                    escape_xml(font_name),
                    escape_xml(font_name)
                )
                .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(ref color) = self.properties.color {
                write!(xml, "<w:color w:val=\"{}\"/>", color)
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(ref highlight) = self.properties.highlight {
                write!(xml, "<w:highlight w:val=\"{}\"/>", highlight)
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if self.properties.no_proof {
                xml.push_str("<w:noProof/>");
            }

            if self.properties.web_hidden {
                xml.push_str("<w:webHidden/>");
            }

            crate::run_effects::codec::write(&self.properties.effects, xml)?;

            if let Some(change) = &self.property_change {
                change.write_xml(xml)?;
            }

            xml.push_str("</w:rPr>");
        }

        // Write content based on type
        match &self.content {
            RunContent::OfficeMath(math) => xml.push_str(math.xml()),
            RunContent::Symbol(symbol) => crate::run_symbols::codec::write_symbol(symbol, xml)?,
            RunContent::Text(text) if !text.is_empty() => {
                let name = if mode == RevisionTextMode::Deleted {
                    "delText"
                } else {
                    "t"
                };
                write!(
                    xml,
                    "<w:{name} xml:space=\"preserve\">{}</w:{name}>",
                    escape_xml(text)
                )
                .map_err(|e| Error::Xml(e.to_string()))?;
            },
            RunContent::PageNumber(format) => {
                // Field begin
                xml.push_str("<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    self.properties.write_open(xml);
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    crate::run_effects::codec::write(&self.properties.effects, xml)?;
                    xml.push_str("</w:rPr>");
                }
                // Field instruction
                let name = if mode == RevisionTextMode::Deleted {
                    "delInstrText"
                } else {
                    "instrText"
                };
                write!(
                    xml,
                    "<w:{name} xml:space=\"preserve\">PAGE \\* {}</w:{name}></w:r><w:r>",
                    format.as_str()
                )
                .map_err(|e| Error::Xml(e.to_string()))?;
                // Field separate
                xml.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    self.properties.write_open(xml);
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    crate::run_effects::codec::write(&self.properties.effects, xml)?;
                    xml.push_str("</w:rPr>");
                }
                // Placeholder text
                if mode == RevisionTextMode::Deleted {
                    xml.push_str("<w:delText>1</w:delText></w:r><w:r>");
                } else {
                    xml.push_str("<w:t>1</w:t></w:r><w:r>");
                }
                // Field end
                xml.push_str("<w:fldChar w:fldCharType=\"end\"/>");
            },
            RunContent::PageCount => {
                xml.push_str("<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    self.properties.write_open(xml);
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    crate::run_effects::codec::write(&self.properties.effects, xml)?;
                    xml.push_str("</w:rPr>");
                }
                if mode == RevisionTextMode::Deleted {
                    xml.push_str("<w:delInstrText xml:space=\"preserve\">NUMPAGES</w:delInstrText></w:r><w:r>");
                } else {
                    xml.push_str(
                        "<w:instrText xml:space=\"preserve\">NUMPAGES</w:instrText></w:r><w:r>",
                    );
                }
                xml.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>");
                if self.properties.has_properties() {
                    self.properties.write_open(xml);
                    if let Some(bold) = self.properties.bold
                        && bold
                    {
                        xml.push_str("<w:b/>");
                    }
                    crate::run_effects::codec::write(&self.properties.effects, xml)?;
                    xml.push_str("</w:rPr>");
                }
                if mode == RevisionTextMode::Deleted {
                    xml.push_str("<w:delText>1</w:delText></w:r><w:r>");
                } else {
                    xml.push_str("<w:t>1</w:t></w:r><w:r>");
                }
                xml.push_str("<w:fldChar w:fldCharType=\"end\"/>");
            },
            RunContent::Tab => {
                xml.push_str("<w:tab/>");
            },
            RunContent::PageBreak => {
                xml.push_str("<w:br w:type=\"page\"/>");
            },
            RunContent::FootnoteReference(id) => {
                write!(xml, "<w:footnoteReference w:id=\"{}\"/>", id)
                    .map_err(|e| Error::Xml(e.to_string()))?;
            },
            RunContent::EndnoteReference(id) => {
                write!(xml, "<w:endnoteReference w:id=\"{}\"/>", id)
                    .map_err(|e| Error::Xml(e.to_string()))?;
            },
            _ => {},
        }

        // Write line break if set
        if self.properties.has_break {
            xml.push_str("<w:br/>");
        }

        xml.push_str("</w:r>");

        Ok(())
    }
}

/// Run properties.
#[derive(Debug, Default, Clone)]
pub(crate) struct RunProperties {
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) underline: Option<UnderlineStyle>,
    pub(crate) font_size: Option<u32>,
    pub(crate) font_name: Option<String>,
    pub(crate) color: Option<String>,
    pub(crate) highlight: Option<String>,
    pub(crate) has_break: bool,
    pub(crate) no_proof: bool,
    pub(crate) web_hidden: bool,
    pub(crate) effects: Effects,
}

impl RunProperties {
    pub(crate) fn has_properties(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.font_size.is_some()
            || self.font_name.is_some()
            || self.color.is_some()
            || self.highlight.is_some()
            || self.no_proof
            || self.web_hidden
            || !self.effects.is_empty()
    }

    pub(crate) fn write_open(&self, xml: &mut String) {
        if self.effects.is_empty() {
            xml.push_str("<w:rPr>");
        } else {
            xml.push_str(
                "<w:rPr xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" \
                 xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" \
                 mc:Ignorable=\"w14\">",
            );
        }
    }

    pub(crate) fn write_values(&self, xml: &mut String) -> Result<()> {
        if self.bold == Some(true) {
            xml.push_str("<w:b/>");
        }
        if self.italic == Some(true) {
            xml.push_str("<w:i/>");
        }
        if let Some(v) = self.underline {
            write!(xml, "<w:u w:val=\"{}\"/>", v.as_str())?;
        }
        if let Some(v) = self.font_size {
            write!(xml, "<w:sz w:val=\"{v}\"/>")?;
        }
        if let Some(v) = &self.font_name {
            write!(
                xml,
                "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
                escape_xml(v),
                escape_xml(v)
            )?;
        }
        if let Some(v) = &self.color {
            write!(xml, "<w:color w:val=\"{}\"/>", escape_xml(v))?;
        }
        if let Some(v) = &self.highlight {
            write!(xml, "<w:highlight w:val=\"{}\"/>", escape_xml(v))?;
        }
        if self.no_proof {
            xml.push_str("<w:noProof/>");
        }
        if self.web_hidden {
            xml.push_str("<w:webHidden/>");
        }
        Ok(())
    }
}
