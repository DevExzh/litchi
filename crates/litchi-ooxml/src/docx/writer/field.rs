//! Field writer support for DOCX documents.

use super::revision::RevisionTextMode;
use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

const MAX_CITATION_SOURCES: usize = 16;
const MAX_CITATION_TEXT_BYTES: usize = 65_536;

/// One source and its source-local options in an inert Word `CITATION` field.
///
/// Word applies volume, prefix, and suffix switches to the source tag that
/// precedes them. This model preserves that order but never looks up the tag,
/// reads bibliography XML, or formats a citation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CitationSource {
    tag: String,
    volume: Option<u32>,
    prefix: Option<String>,
    suffix: Option<String>,
}

impl CitationSource {
    /// Create a source reference from its case-sensitive bibliography tag.
    pub fn new(tag: impl Into<String>) -> Result<Self> {
        let tag = tag.into();
        validate_citation_instruction_text(&tag, "citation source tag", false)?;
        Ok(Self {
            tag,
            volume: None,
            prefix: None,
            suffix: None,
        })
    }

    /// Return the case-sensitive bibliography tag stored by this source.
    pub fn tag(&self) -> &str {
        &self.tag
    }

    /// Replace the bibliography tag.
    pub fn set_tag(&mut self, tag: impl Into<String>) -> Result<()> {
        let tag = tag.into();
        validate_citation_instruction_text(&tag, "citation source tag", false)?;
        self.tag = tag;
        Ok(())
    }

    /// Return the optional source-local volume number.
    pub fn volume(&self) -> Option<u32> {
        self.volume
    }

    /// Set or clear the source-local volume number.
    pub fn set_volume(&mut self, volume: Option<u32>) {
        self.volume = volume;
    }

    /// Return the optional source-local citation prefix.
    pub fn prefix(&self) -> Option<&str> {
        self.prefix.as_deref()
    }

    /// Set or clear the source-local citation prefix.
    pub fn set_prefix(&mut self, prefix: Option<String>) -> Result<()> {
        if let Some(prefix) = &prefix {
            validate_citation_instruction_text(prefix, "citation prefix", false)?;
        }
        self.prefix = prefix;
        Ok(())
    }

    /// Return the optional source-local citation suffix.
    pub fn suffix(&self) -> Option<&str> {
        self.suffix.as_deref()
    }

    /// Set or clear the source-local citation suffix.
    pub fn set_suffix(&mut self, suffix: Option<String>) -> Result<()> {
        if let Some(suffix) = &suffix {
            validate_citation_instruction_text(suffix, "citation suffix", false)?;
        }
        self.suffix = suffix;
        Ok(())
    }

    fn validate(&self) -> Result<()> {
        validate_citation_instruction_text(&self.tag, "citation source tag", false)?;
        if let Some(prefix) = &self.prefix {
            validate_citation_instruction_text(prefix, "citation prefix", false)?;
        }
        if let Some(suffix) = &self.suffix {
            validate_citation_instruction_text(suffix, "citation suffix", false)?;
        }
        Ok(())
    }

    fn append_instruction(&self, instruction: &mut String, additional: bool) {
        if additional {
            instruction.push_str(" \\m ");
        } else {
            instruction.push(' ');
        }
        append_field_argument(instruction, &self.tag);
        if let Some(volume) = self.volume {
            instruction.push_str(" \\v ");
            instruction.push_str(&volume.to_string());
        }
        if let Some(prefix) = &self.prefix {
            instruction.push_str(" \\f ");
            append_quoted_field_argument(instruction, prefix);
        }
        if let Some(suffix) = &self.suffix {
            instruction.push_str(" \\s ");
            append_quoted_field_argument(instruction, suffix);
        }
    }
}

/// Typed, inert authoring data for one Word `CITATION` field.
///
/// The serialized field instruction contains only caller-supplied source tags
/// and documented switches. It never resolves source tags, accesses custom XML
/// bibliography stores, applies a citation style, or refreshes the result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CitationFieldSpec {
    sources: Vec<CitationSource>,
    locale: Option<u32>,
    cached_result: Option<String>,
    dirty: bool,
}

impl CitationFieldSpec {
    /// Create a dirty field with one required primary source.
    pub fn new(primary_source: CitationSource) -> Self {
        Self {
            sources: vec![primary_source],
            locale: None,
            cached_result: None,
            dirty: true,
        }
    }

    /// Return the primary source tag and its source-local switches.
    pub fn primary_source(&self) -> &CitationSource {
        &self.sources[0]
    }

    /// Mutably access the primary source and its source-local switches.
    pub fn primary_source_mut(&mut self) -> &mut CitationSource {
        &mut self.sources[0]
    }

    /// Return every source in field-code order.
    pub fn sources(&self) -> &[CitationSource] {
        &self.sources
    }

    /// Return the optional locale identifier stored with the field.
    pub fn locale(&self) -> Option<u32> {
        self.locale
    }

    /// Set or clear the stored locale identifier.
    pub fn set_locale(&mut self, locale: Option<u32>) {
        self.locale = locale;
    }

    /// Add another source, serialized with a `\\m` switch.
    pub fn add_source(&mut self, source: CitationSource) -> Result<()> {
        source.validate()?;
        if self.sources.len() >= MAX_CITATION_SOURCES {
            return Err(OoxmlError::InvalidFormat(format!(
                "CITATION field supports at most {MAX_CITATION_SOURCES} typed sources"
            )));
        }
        self.sources.push(source);
        Ok(())
    }

    /// Return the caller-supplied cached result, if any.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Set or clear a caller-supplied cached result without generating it.
    pub fn set_cached_result(&mut self, result: Option<String>) -> Result<()> {
        if let Some(result) = &result {
            validate_citation_result_text(result)?;
        }
        self.cached_result = result;
        Ok(())
    }

    /// Return whether the serialized field is marked stale for a word processor.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Set the persisted dirty marker without evaluating the field.
    pub fn set_dirty(&mut self, dirty: bool) {
        self.dirty = dirty;
    }

    /// Build a canonical `CITATION` instruction from the typed metadata.
    pub fn to_instruction(&self) -> Result<String> {
        self.validate()?;
        let mut instruction = String::from("CITATION");
        if let Some(locale) = self.locale {
            instruction.push_str(" \\l ");
            instruction.push_str(&locale.to_string());
        }
        for (index, source) in self.sources.iter().enumerate() {
            source.append_instruction(&mut instruction, index != 0);
        }
        Ok(instruction)
    }

    fn validate(&self) -> Result<()> {
        if self.sources.is_empty() || self.sources.len() > MAX_CITATION_SOURCES {
            return Err(OoxmlError::InvalidFormat(
                "CITATION field requires one through sixteen sources".to_string(),
            ));
        }
        for source in &self.sources {
            source.validate()?;
        }
        if let Some(result) = &self.cached_result {
            validate_citation_result_text(result)?;
        }
        Ok(())
    }
}

fn validate_citation_instruction_text(value: &str, context: &str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.trim().is_empty())
        || value.len() > MAX_CITATION_TEXT_BYTES
        || value.chars().any(|character| character.is_control())
    {
        return Err(OoxmlError::InvalidFormat(format!("invalid {context}")));
    }
    Ok(())
}

fn validate_citation_result_text(value: &str) -> Result<()> {
    if value.len() > MAX_CITATION_TEXT_BYTES
        || value
            .chars()
            .any(|character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}'))
    {
        return Err(OoxmlError::InvalidFormat(
            "invalid cached citation result".to_string(),
        ));
    }
    Ok(())
}

fn append_field_argument(instruction: &mut String, value: &str) {
    if !value.is_empty()
        && value
            .chars()
            .all(|character| !character.is_whitespace() && character != '\\' && character != '"')
    {
        instruction.push_str(value);
    } else {
        append_quoted_field_argument(instruction, value);
    }
}

fn append_quoted_field_argument(instruction: &mut String, value: &str) {
    instruction.push('"');
    for character in value.chars() {
        if matches!(character, '\\' | '"') {
            instruction.push('\\');
        }
        instruction.push(character);
    }
    instruction.push('"');
}

/// A mutable field in a Word document.
///
/// Fields are dynamic content placeholders such as page numbers, dates, cross-references, etc.
///
/// This enum supports both complete fields and individual field characters for complex field structures.
#[derive(Debug, Clone)]
pub enum MutableField {
    /// A complete field with instruction and optional result
    Complete {
        /// Field instruction (e.g., "PAGE", "DATE", "REF MyBookmark")
        instruction: String,
        /// Field result (optional, the displayed value)
        result: Option<String>,
        /// Whether the field is dirty (needs update)
        dirty: bool,
    },
    /// Field begin character
    Begin,
    /// Field instruction text
    Instruction(String),
    /// Field separate character
    Separate {
        /// Whether the field is dirty
        dirty: bool,
    },
    /// Field end character
    End,
}

impl MutableField {
    /// Create a new complete field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction (e.g., "PAGE", "DATE \\@ \"MMMM d, yyyy\"")
    pub fn new(instruction: String) -> Self {
        Self::Complete {
            instruction,
            result: None,
            dirty: true,
        }
    }

    /// Create a field with a result value.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction
    /// * `result` - The current result value
    pub fn with_result(instruction: String, result: String) -> Self {
        Self::Complete {
            instruction,
            result: Some(result),
            dirty: false,
        }
    }

    /// Create a typed, inert `CITATION` field.
    ///
    /// This serializes caller-supplied bibliography tags and switches only. It
    /// does not access source stores, format a citation, or refresh the cached
    /// result.
    pub fn citation(spec: &CitationFieldSpec) -> Result<Self> {
        Ok(Self::Complete {
            instruction: spec.to_instruction()?,
            result: spec.cached_result.clone(),
            dirty: spec.dirty,
        })
    }

    /// Create a field begin character.
    pub fn begin() -> Self {
        Self::Begin
    }

    /// Create a field instruction character.
    pub fn instruction_char(text: String) -> Self {
        Self::Instruction(text)
    }

    /// Create a field separate character.
    pub fn separate() -> Self {
        Self::Separate { dirty: false }
    }

    /// Create a field separate character marked as dirty.
    pub fn separate_dirty() -> Self {
        Self::Separate { dirty: true }
    }

    /// Create a field end character.
    pub fn end() -> Self {
        Self::End
    }

    /// Get the field instruction (for Complete fields only).
    pub fn get_instruction(&self) -> Option<&str> {
        match self {
            Self::Complete { instruction, .. } => Some(instruction),
            Self::Instruction(text) => Some(text),
            _ => None,
        }
    }

    /// Get the field instruction as a string reference.
    ///
    /// # Panics
    /// Panics if this is not a Complete or Instruction field variant.
    pub fn instruction(&self) -> &str {
        self.get_instruction()
            .expect("instruction() called on non-instruction field variant")
    }

    /// Get the field result (alias for get_result).
    pub fn result(&self) -> Option<&str> {
        self.get_result()
    }

    /// Set the field instruction (for Complete fields only).
    pub fn set_instruction(&mut self, new_instruction: String) {
        if let Self::Complete {
            instruction, dirty, ..
        } = self
        {
            *instruction = new_instruction;
            *dirty = true;
        }
    }

    /// Get the field result (for Complete fields only).
    pub fn get_result(&self) -> Option<&str> {
        match self {
            Self::Complete { result, .. } => result.as_deref(),
            _ => None,
        }
    }

    /// Set the field result (for Complete fields only).
    pub fn set_result(&mut self, new_result: Option<String>) {
        if let Self::Complete { result, dirty, .. } = self {
            *result = new_result;
            *dirty = false;
        }
    }

    /// Check if the field is dirty (needs update).
    pub fn is_dirty(&self) -> bool {
        match self {
            Self::Complete { dirty, .. } | Self::Separate { dirty } => *dirty,
            _ => false,
        }
    }

    /// Mark the field as dirty (for Complete fields only).
    pub fn mark_dirty(&mut self) {
        match self {
            Self::Complete { dirty, .. } | Self::Separate { dirty } => *dirty = true,
            _ => {},
        }
    }

    /// Generate XML for this field.
    #[allow(dead_code)]
    pub(crate) fn to_xml(&self) -> Result<String> {
        self.to_xml_mode(RevisionTextMode::Normal)
    }

    pub(crate) fn to_xml_mode(&self, mode: RevisionTextMode) -> Result<String> {
        let mut xml = String::with_capacity(256);

        match self {
            Self::Complete {
                instruction,
                result,
                dirty,
            } => {
                // Field begin
                xml.push_str(r#"<w:fldChar w:fldCharType="begin"/>"#);

                // Field instruction
                let instruction_name = if mode == RevisionTextMode::Deleted { "delInstrText" } else { "instrText" };
                write!(&mut xml, "</w:r><w:r><w:{instruction_name}>{}</w:{instruction_name}>", escape_xml(instruction))?;

                // Separate run
                if *dirty {
                    xml.push_str(
                        r#"</w:r><w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/>"#,
                    );
                } else {
                    xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="separate"/>"#);
                }

                // Field result
                if let Some(res) = result {
                    let text_name = if mode == RevisionTextMode::Deleted { "delText" } else { "t" };
                    write!(&mut xml, "</w:r><w:r><w:{text_name}>{}</w:{text_name}>", escape_xml(res))?;
                }

                // Field end
                xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="end"/>"#);
            },
            Self::Begin => {
                xml.push_str(r#"<w:fldChar w:fldCharType="begin"/>"#);
            },
            Self::Instruction(text) => {
                let name = if mode == RevisionTextMode::Deleted { "delInstrText" } else { "instrText" };
                write!(&mut xml, r#"<w:{name} xml:space="preserve">{}</w:{name}>"#, escape_xml(text))?;
            },
            Self::Separate { dirty } => {
                if *dirty {
                    xml.push_str(r#"<w:fldChar w:fldCharType="separate" w:dirty="true"/>"#);
                } else {
                    xml.push_str(r#"<w:fldChar w:fldCharType="separate"/>"#);
                }
            },
            Self::End => {
                xml.push_str(r#"<w:fldChar w:fldCharType="end"/>"#);
            },
        }

        Ok(xml)
    }

    /// Common field factory methods
    /// Create a PAGE field (page number).
    pub fn page() -> Self {
        Self::new("PAGE".to_string())
    }

    /// Create a NUMPAGES field (total page count).
    pub fn num_pages() -> Self {
        Self::new("NUMPAGES".to_string())
    }

    /// Create a DATE field with optional format.
    ///
    /// # Arguments
    ///
    /// * `format` - Optional date format string (e.g., "MMMM d, yyyy")
    pub fn date(format: Option<&str>) -> Self {
        let instruction = if let Some(fmt) = format {
            format!(r#"DATE \@ "{}""#, fmt)
        } else {
            "DATE".to_string()
        };
        Self::new(instruction)
    }

    /// Create a TIME field with optional format.
    pub fn time(format: Option<&str>) -> Self {
        let instruction = if let Some(fmt) = format {
            format!(r#"TIME \@ "{}""#, fmt)
        } else {
            "TIME".to_string()
        };
        Self::new(instruction)
    }

    /// Create a REF field (cross-reference to a bookmark).
    ///
    /// # Arguments
    ///
    /// * `bookmark_name` - Name of the bookmark to reference
    pub fn reference(bookmark_name: &str) -> Self {
        Self::new(format!("REF {}", bookmark_name))
    }

    /// Create a HYPERLINK field.
    ///
    /// # Arguments
    ///
    /// * `url` - The URL to link to
    pub fn hyperlink(url: &str) -> Self {
        Self::new(format!(r#"HYPERLINK "{}""#, url))
    }

    /// Create a TOC (Table of Contents) field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The complete TOC field instruction (e.g., `TOC \o "1-3" \h \z`)
    /// * `placeholder_text` - Optional placeholder text to display before field update
    pub fn toc(instruction: String, placeholder_text: Option<String>) -> Self {
        Self::Complete {
            instruction,
            result: placeholder_text,
            dirty: true,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_field_creation() {
        let field = MutableField::new("PAGE".to_string());
        assert_eq!(field.instruction(), "PAGE");
        assert!(field.is_dirty());
        assert!(field.result().is_none());
    }

    #[test]
    fn test_field_with_result() {
        let field = MutableField::with_result("PAGE".to_string(), "5".to_string());
        assert_eq!(field.instruction(), "PAGE");
        assert_eq!(field.result(), Some("5"));
        assert!(!field.is_dirty());
    }

    #[test]
    fn test_field_factories() {
        let page = MutableField::page();
        assert_eq!(page.instruction(), "PAGE");

        let date = MutableField::date(Some("MMMM d, yyyy"));
        assert!(date.instruction().contains("DATE"));
        assert!(date.instruction().contains("MMMM d, yyyy"));

        let ref_field = MutableField::reference("MyBookmark");
        assert!(ref_field.instruction().contains("REF MyBookmark"));
    }

    #[test]
    fn typed_citation_serializes_documented_switches_without_evaluation() {
        let mut primary = CitationSource::new("Doe2024").unwrap();
        primary.set_volume(Some(3));
        primary.set_prefix(Some("qtd. in".to_string())).unwrap();
        primary.set_suffix(Some("in press".to_string())).unwrap();
        let mut citation = CitationFieldSpec::new(primary);
        citation.set_locale(Some(1033));
        let mut additional = CitationSource::new("Smith 2025").unwrap();
        additional.set_volume(Some(2));
        citation.add_source(additional).unwrap();
        citation
            .set_cached_result(Some("caller supplied result".to_string()))
            .unwrap();
        citation.set_dirty(false);

        let field = MutableField::citation(&citation).unwrap();
        assert_eq!(
            field.instruction(),
            r#"CITATION \l 1033 Doe2024 \v 3 \f "qtd. in" \s "in press" \m "Smith 2025" \v 2"#
        );
        assert_eq!(field.result(), Some("caller supplied result"));
        assert!(!field.is_dirty());

        let parsed = crate::docx::Field::new(
            field.instruction().to_string(),
            field.result().map(str::to_string),
            field.is_dirty(),
        )
        .citation()
        .unwrap()
        .unwrap();
        assert_eq!(parsed.source_tags(), ["Doe2024", "Smith 2025"]);
        assert_eq!(parsed.switches()[0].name(), 'l');
        assert!(parsed.has_switch('v'));
        assert!(parsed.has_switch('f'));
        assert!(parsed.has_switch('s'));

        assert!(CitationSource::new("source\nname").is_err());
        assert!(
            CitationSource::new("tag")
                .unwrap()
                .set_prefix(Some(String::new()))
                .is_err()
        );
    }

    #[test]
    fn test_field_xml() {
        let mut field = MutableField::with_result("PAGE".to_string(), "1".to_string());
        field.mark_dirty();

        let xml = field.to_xml().unwrap();
        assert!(xml.contains("fldCharType=\"begin\""));
        assert!(xml.contains("instrText"));
        assert!(xml.contains("PAGE"));
        assert!(xml.contains("fldCharType=\"separate\""));
        assert!(xml.contains("dirty=\"true\""));
        assert!(xml.contains("fldCharType=\"end\""));
    }
}
