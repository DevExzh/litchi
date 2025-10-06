/// Paragraph and Run structures for legacy Word documents.
use super::package::Result;
use super::parts::chp::{CharacterProperties, UnderlineStyle, VerticalPosition};

/// A paragraph in a Word document.
///
/// Represents a paragraph in the binary DOC format.
///
/// # Example
///
/// ```rust,ignore
/// for para in document.paragraphs()? {
///     println!("Paragraph text: {}", para.text());
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The text content of this paragraph
    text: String,
    /// Runs within this paragraph
    runs: Vec<Run>,
    // TODO: Add paragraph formatting properties (PAP)
}

impl Paragraph {
    /// Create a new Paragraph from text.
    ///
    /// # Arguments
    ///
    /// * `text` - The text content
    pub(crate) fn new(text: String) -> Self {
        Self {
            text: text.clone(),
            runs: vec![Run::new(text, CharacterProperties::default())],
        }
    }

    /// Create a new Paragraph with runs.
    ///
    /// # Arguments
    ///
    /// * `runs` - The runs within this paragraph
    #[allow(unused)]
    pub(crate) fn with_runs(runs: Vec<Run>) -> Self {
        let text = runs.iter().map(|r| r.text.as_str()).collect::<String>();
        Self { text, runs }
    }

    /// Get the text content of this paragraph.
    ///
    /// # Performance
    ///
    /// Returns a reference to avoid cloning.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the runs in this paragraph.
    ///
    /// Each run represents a region of text with uniform formatting.
    pub fn runs(&self) -> Result<Vec<Run>> {
        Ok(self.runs.clone())
    }

    /// Set the runs for this paragraph (internal use).
    pub(crate) fn set_runs(&mut self, runs: Vec<Run>) {
        self.runs = runs;
    }
}

/// A run within a paragraph.
///
/// Represents a region of text with a single set of formatting properties
/// in the binary DOC format.
///
/// # Example
///
/// ```rust,ignore
/// for run in paragraph.runs()? {
///     println!("Run text: {}", run.text()?);
///     println!("Bold: {:?}", run.bold());
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Run {
    /// The text content of this run
    text: String,
    /// Character formatting properties
    properties: CharacterProperties,
}

impl Run {
    /// Create a new Run from text with character properties.
    pub(crate) fn new(text: String, properties: CharacterProperties) -> Self {
        Self { text, properties }
    }

    /// Get the text content of this run.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Check if this run is bold.
    ///
    /// Returns `Some(true)` if bold is enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn bold(&self) -> Option<bool> {
        self.properties.is_bold
    }

    /// Check if this run is italic.
    ///
    /// Returns `Some(true)` if italic is enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn italic(&self) -> Option<bool> {
        self.properties.is_italic
    }

    /// Check if this run is underlined.
    ///
    /// Returns `Some(true)` if underline is present,
    /// `None` if not specified.
    pub fn underline(&self) -> Option<bool> {
        match self.properties.underline {
            UnderlineStyle::None => None,
            _ => Some(true),
        }
    }

    /// Get the underline style for this run.
    ///
    /// Returns the specific underline style if applied.
    pub fn underline_style(&self) -> UnderlineStyle {
        self.properties.underline
    }

    /// Check if this run is strikethrough.
    pub fn strikethrough(&self) -> Option<bool> {
        self.properties.is_strikethrough
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: DOC format stores font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&self) -> Option<u16> {
        self.properties.font_size
    }

    /// Get the text color as RGB tuple.
    pub fn color(&self) -> Option<(u8, u8, u8)> {
        self.properties.color
    }

    /// Check if text is superscript.
    pub fn is_superscript(&self) -> bool {
        self.properties.vertical_position == VerticalPosition::Superscript
    }

    /// Check if text is subscript.
    pub fn is_subscript(&self) -> bool {
        self.properties.vertical_position == VerticalPosition::Subscript
    }

    /// Check if text is in small caps.
    pub fn small_caps(&self) -> Option<bool> {
        self.properties.is_small_caps
    }

    /// Check if text is in all caps.
    pub fn all_caps(&self) -> Option<bool> {
        self.properties.is_all_caps
    }

    /// Get the character properties for this run.
    ///
    /// Provides access to all formatting properties.
    pub fn properties(&self) -> &CharacterProperties {
        &self.properties
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_paragraph_text() {
        let para = Paragraph::new("Hello, World!".to_string());
        assert_eq!(para.text().unwrap(), "Hello, World!");
    }

    #[test]
    fn test_run_text() {
        let run = Run::new("Test text".to_string(), CharacterProperties::default());
        assert_eq!(run.text().unwrap(), "Test text");
        assert_eq!(run.bold(), None);
        assert_eq!(run.italic(), None);
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_run_with_formatting() {
        let mut props = CharacterProperties::default();
        props.is_bold = Some(true);
        props.is_italic = Some(true);
        props.font_size = Some(24); // 12pt

        let run = Run::new("Formatted text".to_string(), props);
        assert_eq!(run.bold(), Some(true));
        assert_eq!(run.italic(), Some(true));
        assert_eq!(run.font_size(), Some(24));
    }
}

