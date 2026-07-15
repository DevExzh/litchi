/// Paragraph and Run structures for legacy Word documents.
use super::package::Result;
use super::parts::chp::{CharacterProperties, UnderlineStyle, VerticalPosition};
use super::parts::revisions::RevisionAuthorTable;
use super::revision::{NumberingRevisionMark, RevisionKind, RevisionMark, decode_dttm};
use std::sync::Arc;

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
    /// Resolved paragraph-formatting revision metadata.
    formatting_revision: Option<RevisionMark>,
    /// Resolved paragraph numbering revision metadata.
    numbering_revision: Option<NumberingRevisionMark>,
}

impl Paragraph {
    /// Create a new Paragraph from text.
    ///
    /// # Arguments
    ///
    /// * `text` - The text content (may be empty if runs will be set separately)
    ///
    /// # Note
    ///
    /// Following Apache POI's design: when creating paragraphs with explicit runs,
    /// pass empty string here and set runs separately to avoid text duplication.
    pub(crate) fn new(text: String) -> Self {
        // Only create default run if text is non-empty
        let runs = if !text.is_empty() {
            vec![Run::new(text.clone(), CharacterProperties::default())]
        } else {
            Vec::new()
        };

        Self {
            text,
            runs,
            properties: super::parts::pap::ParagraphProperties::default(),
            formatting_revision: None,
            numbering_revision: None,
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
            formatting_revision: None,
            numbering_revision: None,
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
            formatting_revision: None,
            numbering_revision: None,
        }
    }

    /// Get the text content of this paragraph.
    ///
    /// # Performance
    ///
    /// Returns a reference to avoid cloning when paragraph has stored text.
    /// If runs are set, returns the stored text which should match run concatenation.
    pub fn text(&self) -> Result<&str> {
        // If we have an explicit text field (for compatibility), use it
        // Otherwise it will be empty and runs contain the actual text
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
        self.text.clear();
        self.text
            .reserve(runs.iter().map(|run| run.text.len()).sum());
        for run in &runs {
            self.text.push_str(&run.text);
        }
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

    /// Tracked paragraph-formatting revision metadata.
    pub fn formatting_revision(&self) -> Option<&RevisionMark> {
        self.formatting_revision.as_ref()
    }

    /// Numbering revision metadata for this paragraph.
    pub fn numbering_revision(&self) -> Option<&NumberingRevisionMark> {
        self.numbering_revision.as_ref()
    }

    /// Whether a numbered list was applied after the previous revision.
    pub fn numbering_revision_list_applied(&self) -> Option<bool> {
        self.properties.numbering_revision_list_applied
    }

    pub(crate) fn resolve_revision(&mut self, authors: &RevisionAuthorTable) -> Result<()> {
        if self.properties.has_formatting_revision == Some(true) {
            let author_index = self
                .properties
                .formatting_revision_author_index
                .unwrap_or(0);
            let author = authors.get(author_index).ok_or_else(|| {
                super::package::DocError::Corrupted(
                    "paragraph revision author index exceeds SttbfRMark".to_string(),
                )
            })?;
            self.formatting_revision = Some(RevisionMark {
                kind: RevisionKind::Formatting,
                author_index,
                author: author.to_string(),
                timestamp: self
                    .properties
                    .formatting_revision_timestamp
                    .map(decode_dttm)
                    .transpose()?
                    .flatten(),
                revision_id: None,
            });
        }
        if let Some(revision) = &self.properties.numbering_revision {
            let author = authors.get(revision.author_index).ok_or_else(|| {
                super::package::DocError::Corrupted(
                    "numbering revision author index exceeds SttbfRMark".to_string(),
                )
            })?;
            self.numbering_revision = Some(NumberingRevisionMark {
                was_numbered: revision.was_numbered,
                author_index: revision.author_index,
                author: author.to_string(),
                timestamp: decode_dttm(revision.timestamp)?,
                placeholder_positions: revision.placeholder_positions,
                number_formats: revision.number_formats,
                numbers: revision.numbers,
                format_string: revision.format_string.clone(),
            });
        }
        Ok(())
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
///
///     // Check for embedded images
///     if let Some(img) = run.image() {
///         println!("Image at offset: {}", img.pic_offset());
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
    /// Using Arc to share AST across multiple runs without cloning (thread-safe)
    #[cfg(feature = "formula")]
    mtef_formula_ast: Option<Arc<Vec<litchi_formula::MathNode<'static>>>>,
    /// Parsed MTEF formula AST placeholder (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    mtef_formula_ast: Option<Arc<Vec<()>>>,
    /// Embedded image (metadata only, data loaded lazily via Document::image_data)
    image: Option<super::image::Image>,
    /// Resolved insertion revision metadata.
    insertion_revision: Option<RevisionMark>,
    /// Resolved deletion revision metadata.
    deletion_revision: Option<RevisionMark>,
    /// Resolved character-formatting revision metadata.
    formatting_revision: Option<RevisionMark>,
}

impl Run {
    /// Create a new Run from text with character properties.
    pub(crate) fn new(text: String, properties: CharacterProperties) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: None,
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
        }
    }

    /// Create a new Run with MTEF formula AST.
    #[cfg(feature = "formula")]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        mtef_ast: Arc<Vec<litchi_formula::MathNode<'static>>>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: Some(mtef_ast),
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
        }
    }

    /// Create a new Run with MTEF formula AST fallback (when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        _mtef_ast: Arc<Vec<()>>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: None,
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
        }
    }

    /// Create a new Run with an embedded image.
    pub(crate) fn with_image(
        text: String,
        properties: CharacterProperties,
        image: super::image::Image,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_ast: None,
            image: Some(image),
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
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

    /// Insertion revision metadata for this run.
    pub fn insertion_revision(&self) -> Option<&RevisionMark> {
        self.insertion_revision.as_ref()
    }

    /// Deletion revision metadata for this run.
    pub fn deletion_revision(&self) -> Option<&RevisionMark> {
        self.deletion_revision.as_ref()
    }

    /// Character-formatting revision metadata for this run.
    pub fn formatting_revision(&self) -> Option<&RevisionMark> {
        self.formatting_revision.as_ref()
    }

    pub(crate) fn resolve_revisions(&mut self, authors: &RevisionAuthorTable) -> Result<()> {
        if self.properties.is_revision_inserted == Some(true) {
            self.insertion_revision = Some(Self::revision_mark(
                RevisionKind::Insertion,
                self.properties.revision_author_index.unwrap_or(0),
                self.properties.revision_timestamp,
                self.properties.revision_id,
                authors,
            )?);
        }
        if self.properties.is_revision_deleted == Some(true) {
            self.deletion_revision = Some(Self::revision_mark(
                RevisionKind::Deletion,
                self.properties.deletion_author_index.unwrap_or(0),
                self.properties.deletion_timestamp,
                self.properties.deletion_revision_id,
                authors,
            )?);
        }
        if self.properties.has_formatting_revision == Some(true) {
            self.formatting_revision = Some(Self::revision_mark(
                RevisionKind::Formatting,
                self.properties
                    .formatting_revision_author_index
                    .unwrap_or(0),
                self.properties.formatting_revision_timestamp,
                None,
                authors,
            )?);
        }
        Ok(())
    }

    fn revision_mark(
        kind: RevisionKind,
        author_index: u16,
        packed_timestamp: Option<u32>,
        revision_id: Option<u16>,
        authors: &RevisionAuthorTable,
    ) -> Result<RevisionMark> {
        let author = authors.get(author_index).ok_or_else(|| {
            super::package::DocError::Corrupted(
                "revision author index exceeds SttbfRMark".to_string(),
            )
        })?;
        Ok(RevisionMark {
            kind,
            author_index,
            author: author.to_string(),
            timestamp: packed_timestamp.map(decode_dttm).transpose()?.flatten(),
            revision_id,
        })
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
    pub fn mtef_formula_ast(&self) -> Option<&Arc<Vec<litchi_formula::MathNode<'static>>>> {
        self.mtef_formula_ast.as_ref()
    }

    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast(&self) -> Option<&Arc<Vec<()>>> {
        self.mtef_formula_ast.as_ref()
    }

    /// Get a mutable reference to the MTEF formula AST.
    ///
    /// This allows for modification of the formula AST if needed.
    #[cfg(feature = "formula")]
    pub fn mtef_formula_ast_mut(
        &mut self,
    ) -> &mut Option<Arc<Vec<litchi_formula::MathNode<'static>>>> {
        &mut self.mtef_formula_ast
    }

    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Arc<Vec<()>>> {
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
            use litchi_formula::LatexConverter;
            let mut converter = LatexConverter::new();
            // Dereference Rc to access the Vec
            match converter.convert_nodes(ast.as_ref()) {
                Ok(latex) => Ok(Some(latex.to_string())),
                Err(e) => {
                    // Return error message as placeholder
                    Ok(Some(format!("[Formula conversion error: {}]", e)))
                },
            }
        } else {
            Ok(None)
        }
    }

    /// Convert formula to LaTeX (fallback when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub fn formula_as_latex(&self) -> Result<Option<String>> {
        if self.mtef_formula_ast.is_some() {
            Ok(Some(
                "[Formula support disabled - enable 'formula' feature]".to_string(),
            ))
        } else {
            Ok(None)
        }
    }

    /// Check if this run is an OLE2 embedded object (like an equation or image).
    pub fn is_ole_object(&self) -> bool {
        self.properties.is_ole2
    }

    /// Check if this run contains an embedded image.
    pub fn has_image(&self) -> bool {
        self.image.is_some()
    }

    /// Get the embedded image if this run contains one.
    ///
    /// Returns the image metadata. Use `Document::image_data()` to get the actual binary data.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(img) = run.image() {
    ///     let data = doc.image_data(img)?;
    ///     // Process image data...
    /// }
    /// ```
    pub fn image(&self) -> Option<&super::image::Image> {
        self.image.as_ref()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_run_revision_authors_and_timestamps() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        let mut run = Run::new(
            "changed".to_string(),
            CharacterProperties {
                is_revision_inserted: Some(true),
                revision_author_index: Some(1),
                revision_timestamp: Some(timestamp),
                revision_id: Some(42),
                ..CharacterProperties::default()
            },
        );
        let authors = RevisionAuthorTable::from_authors(&["Unknown", "Alice"]);
        run.resolve_revisions(&authors).unwrap();
        let revision = run.insertion_revision().unwrap();
        assert_eq!(revision.kind, RevisionKind::Insertion);
        assert_eq!(revision.author, "Alice");
        assert_eq!(revision.revision_id, Some(42));
        assert_eq!(revision.timestamp.unwrap().year, 2026);
        assert!(run.deletion_revision().is_none());
        assert!(run.formatting_revision().is_none());

        let mut formatted = Run::new(
            "formatted".to_string(),
            CharacterProperties {
                has_formatting_revision: Some(true),
                formatting_revision_author_index: Some(1),
                formatting_revision_timestamp: Some(timestamp),
                ..CharacterProperties::default()
            },
        );
        formatted.resolve_revisions(&authors).unwrap();
        let revision = formatted.formatting_revision().unwrap();
        assert_eq!(revision.kind, RevisionKind::Formatting);
        assert_eq!(revision.author, "Alice");
        assert_eq!(revision.timestamp.unwrap().year, 2026);
        assert_eq!(revision.revision_id, None);

        let mut bad_author = Run::new(
            "changed".to_string(),
            CharacterProperties {
                is_revision_inserted: Some(true),
                revision_author_index: Some(2),
                ..CharacterProperties::default()
            },
        );
        assert!(bad_author.resolve_revisions(&authors).is_err());

        let mut bad_time = Run::new(
            "changed".to_string(),
            CharacterProperties {
                is_revision_inserted: Some(true),
                revision_timestamp: Some(63),
                ..CharacterProperties::default()
            },
        );
        assert!(bad_time.resolve_revisions(&authors).is_err());
    }

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
