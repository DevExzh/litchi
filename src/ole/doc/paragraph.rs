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
    /// Paragraph formatting properties (PAP)
    properties: super::parts::pap::ParagraphProperties,
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
            properties: super::parts::pap::ParagraphProperties::default(),
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
        Self { 
            text, 
            runs,
            properties: super::parts::pap::ParagraphProperties::default(),
        }
    }

    /// Create a new Paragraph with text and properties.
    #[allow(dead_code)] // TODO: remove this once we use this function
    pub(crate) fn with_properties(
        text: String,
        properties: super::parts::pap::ParagraphProperties,
    ) -> Self {
        Self {
            text,
            runs: Vec::new(),
            properties,
        }
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

    /// Set the paragraph properties (internal use).
    pub(crate) fn set_properties(&mut self, properties: super::parts::pap::ParagraphProperties) {
        self.properties = properties;
    }

    /// Get the paragraph properties.
    pub fn properties(&self) -> &super::parts::pap::ParagraphProperties {
        &self.properties
    }

    /// Extract all MTEF formulas from this paragraph as LaTeX.
    ///
    /// Returns a vector of LaTeX formula strings found in any run within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     let formulas = para.formulas_as_latex()?;
    ///     for formula in formulas {
    ///         println!("Formula: {}", formula);
    ///     }
    /// }
    /// ```
    pub fn formulas_as_latex(&self) -> Result<Vec<String>> {
        let mut formulas = Vec::new();
        for run in &self.runs {
            if let Some(latex) = run.formula_as_latex()? {
                formulas.push(latex);
            }
        }
        Ok(formulas)
    }

    /// Check if this paragraph contains any formulas.
    pub fn has_formulas(&self) -> bool {
        self.runs.iter().any(|r| r.has_mtef_formula())
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
///
///     // Check for embedded MTEF formulas
///     if let Some(formula_ast) = run.mtef_formula_ast()? {
///         println!("MTEF formula AST with {} nodes", formula_ast.len());
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Run {
    /// The text content of this run
    text: String,
    /// Character formatting properties
    properties: CharacterProperties,
    /// Parsed MTEF formula AST (if this run contains a formula)
    #[cfg(feature = "formula")]
    mtef_formula_ast: Option<Vec<crate::formula::MathNode<'static>>>,
    /// Parsed MTEF formula AST placeholder (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    mtef_formula_ast: Option<Vec<()>>,
}

impl Run {
    /// Create a new Run from text with character properties.
    pub(crate) fn new(text: String, properties: CharacterProperties) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: None,
        }
    }

    /// Create a new Run with MTEF formula AST.
    #[cfg(feature = "formula")]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        mtef_ast: Vec<crate::formula::MathNode<'static>>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: Some(mtef_ast),
        }
    }
    
    /// Create a new Run with MTEF formula AST fallback (when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        _mtef_ast: Vec<()>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: None,
        }
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

    /// Check if this run contains an MTEF formula.
    ///
    /// Returns true if this run contains a parsed MTEF formula AST.
    pub fn has_mtef_formula(&self) -> bool {
        self.mtef_formula_ast.is_some()
    }

    /// Get the MTEF formula AST if this run contains a formula.
    ///
    /// Returns the parsed MTEF formula as AST nodes if this run contains a MathType equation,
    /// None otherwise.
    #[cfg(feature = "formula")]
    pub fn mtef_formula_ast(&self) -> Option<&Vec<crate::formula::MathNode<'static>>> {
        self.mtef_formula_ast.as_ref()
    }
    
    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast(&self) -> Option<&Vec<()>> {
        self.mtef_formula_ast.as_ref()
    }

    /// Get a mutable reference to the MTEF formula AST.
    ///
    /// This allows for modification of the formula AST if needed.
    #[cfg(feature = "formula")]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<crate::formula::MathNode<'static>>> {
        &mut self.mtef_formula_ast
    }
    
    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<()>> {
        &mut self.mtef_formula_ast
    }

    /// Get the MTEF formula as LaTeX string if this run contains a formula.
    ///
    /// Converts the formula AST to LaTeX format for easy display and processing.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for run in paragraph.runs()? {
    ///     if let Some(latex) = run.formula_as_latex()? {
    ///         println!("Formula: {}", latex);
    ///     }
    /// }
    /// ```
    #[cfg(feature = "formula")]
    pub fn formula_as_latex(&self) -> Result<Option<String>> {
        if let Some(ast) = &self.mtef_formula_ast {
            use crate::formula::LatexConverter;
            let mut converter = LatexConverter::new();
            match converter.convert_nodes(ast) {
                Ok(latex) => Ok(Some(latex.to_string())),
                Err(e) => {
                    // Return error message as placeholder
                    Ok(Some(format!("[Formula conversion error: {}]", e)))
                }
            }
        } else {
            Ok(None)
        }
    }
    
    /// Convert formula to LaTeX (fallback when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub fn formula_as_latex(&self) -> Result<Option<String>> {
        if self.mtef_formula_ast.is_some() {
            Ok(Some("[Formula support disabled - enable 'formula' feature]".to_string()))
        } else {
            Ok(None)
        }
    }

    /// Check if this run is an OLE2 embedded object (like an equation or image).
    pub fn is_ole_object(&self) -> bool {
        self.properties.is_ole2
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
        assert!(run.bold().unwrap_or(false));
        assert!(run.italic().unwrap_or(false));
        assert_eq!(run.font_size(), Some(24));
    }
}

