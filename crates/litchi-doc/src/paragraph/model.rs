/// Paragraph and Run semantic models for legacy Word documents.
use crate::package::Result;
use crate::parts::chp::{CharacterProperties, UnderlineStyle, VerticalPosition};
use crate::parts::revisions::RevisionAuthorTable;
use crate::revision::{
    DisplayFieldRevisionMark, NumberingRevisionMark, RevisionKind, RevisionMark, RevisionReason,
    decode_dttm,
};
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
    properties: crate::parts::pap::ParagraphProperties,
    /// Resolved paragraph-formatting revision metadata.
    formatting_revision: Option<RevisionMark>,
    /// Resolved paragraph numbering revision metadata.
    numbering_revision: Option<NumberingRevisionMark>,
    /// Resolved property revision metadata for a containing table row.
    table_formatting_revision: Option<RevisionMark>,
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
            properties: crate::parts::pap::ParagraphProperties::default(),
            formatting_revision: None,
            numbering_revision: None,
            table_formatting_revision: None,
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
            properties: crate::parts::pap::ParagraphProperties::default(),
            formatting_revision: None,
            numbering_revision: None,
            table_formatting_revision: None,
        }
    }

    /// Create a new Paragraph with text and properties.
    #[allow(dead_code)] // TODO: remove this once we use this function
    pub(crate) fn with_properties(
        text: String,
        properties: crate::parts::pap::ParagraphProperties,
    ) -> Self {
        Self {
            text,
            runs: Vec::new(),
            properties,
            formatting_revision: None,
            numbering_revision: None,
            table_formatting_revision: None,
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
    pub(crate) fn set_properties(&mut self, properties: crate::parts::pap::ParagraphProperties) {
        self.properties = properties;
    }

    /// Get the paragraph properties.
    pub fn properties(&self) -> &crate::parts::pap::ParagraphProperties {
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

    pub(crate) fn table_formatting_revision(&self) -> Option<&RevisionMark> {
        self.table_formatting_revision.as_ref()
    }

    pub(crate) fn resolve_revision(&mut self, authors: &RevisionAuthorTable) -> Result<()> {
        if self.properties.has_formatting_revision == Some(true) {
            let author_index = self
                .properties
                .formatting_revision_author_index
                .unwrap_or(0);
            let author = authors.get(author_index).ok_or_else(|| {
                crate::package::DocError::Corrupted(
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
                reason: None,
                revision_save_id: None,
            });
        }
        if let Some(revision) = &self.properties.numbering_revision {
            let author = authors.get(revision.author_index).ok_or_else(|| {
                crate::package::DocError::Corrupted(
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
        if self.properties.has_table_formatting_revision == Some(true) {
            let author_index = self
                .properties
                .table_formatting_revision_author_index
                .unwrap_or(0);
            let author = authors.get(author_index).ok_or_else(|| {
                crate::package::DocError::Corrupted(
                    "table-row revision author index exceeds SttbfRMark".to_string(),
                )
            })?;
            self.table_formatting_revision = Some(RevisionMark {
                kind: RevisionKind::Formatting,
                author_index,
                author: author.to_string(),
                timestamp: self
                    .properties
                    .table_formatting_revision_timestamp
                    .map(decode_dttm)
                    .transpose()?
                    .flatten(),
                revision_id: None,
                reason: None,
                revision_save_id: None,
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
    /// Owned LaTeX rendered from MTEF while its scoped parser arena is alive.
    mtef_formula_latex: Option<Arc<str>>,
    /// Embedded image (metadata only, data loaded lazily via Document::image_data)
    image: Option<crate::image::Image>,
    /// Resolved insertion revision metadata.
    insertion_revision: Option<RevisionMark>,
    /// Resolved deletion revision metadata.
    deletion_revision: Option<RevisionMark>,
    /// Resolved character-formatting revision metadata.
    formatting_revision: Option<RevisionMark>,
    /// Resolved LISTNUM display-field revision metadata.
    display_field_revision: Option<DisplayFieldRevisionMark>,
}

impl Run {
    /// Create a new Run from text with character properties.
    pub(crate) fn new(text: String, properties: CharacterProperties) -> Self {
        Self {
            text,
            properties,
            mtef_formula_latex: None,
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
            display_field_revision: None,
        }
    }

    /// Create a new Run with MTEF formula AST.
    #[cfg(feature = "formula")]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        mtef_latex: Arc<str>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_latex: Some(mtef_latex),
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
            display_field_revision: None,
        }
    }

    /// Create a new Run with MTEF formula AST fallback (when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub(crate) fn with_mtef_formula(
        text: String,
        properties: CharacterProperties,
        _mtef_latex: Arc<str>,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_latex: None,
            image: None,
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
            display_field_revision: None,
        }
    }

    /// Create a new Run with an embedded image.
    pub(crate) fn with_image(
        text: String,
        properties: CharacterProperties,
        image: crate::image::Image,
    ) -> Self {
        Self {
            text,
            properties,
            mtef_formula_latex: None,
            image: Some(image),
            insertion_revision: None,
            deletion_revision: None,
            formatting_revision: None,
            display_field_revision: None,
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

    /// Revision metadata for a LISTNUM display-field result.
    pub fn display_field_revision(&self) -> Option<&DisplayFieldRevisionMark> {
        self.display_field_revision.as_ref()
    }

    pub(crate) fn resolve_revisions(&mut self, authors: &RevisionAuthorTable) -> Result<()> {
        if self.properties.is_revision_inserted == Some(true) {
            self.insertion_revision = Some(Self::revision_mark(
                RevisionKind::Insertion,
                self.properties.revision_author_index.unwrap_or(0),
                self.properties.revision_timestamp,
                self.properties.revision_id,
                self.properties.insertion_revision_save_id,
                authors,
            )?);
        }
        if self.properties.is_revision_deleted == Some(true) {
            self.deletion_revision = Some(Self::revision_mark(
                RevisionKind::Deletion,
                self.properties.deletion_author_index.unwrap_or(0),
                self.properties.deletion_timestamp,
                self.properties.deletion_revision_id,
                self.properties.deletion_revision_save_id,
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
                self.properties.revision_id,
                self.properties.formatting_revision_save_id,
                authors,
            )?);
        }
        if let Some(revision) = &self.properties.display_field_revision
            && revision.active
        {
            let author = authors.get(revision.author_index).ok_or_else(|| {
                crate::package::DocError::Corrupted(
                    "display-field revision author index exceeds SttbfRMark".to_string(),
                )
            })?;
            self.display_field_revision = Some(DisplayFieldRevisionMark {
                author_index: revision.author_index,
                author: author.to_string(),
                timestamp: decode_dttm(revision.timestamp)?,
                previous_result: revision.previous_result.clone(),
            });
        }
        Ok(())
    }

    fn revision_mark(
        kind: RevisionKind,
        author_index: u16,
        packed_timestamp: Option<u32>,
        revision_id: Option<u16>,
        revision_save_id: Option<u32>,
        authors: &RevisionAuthorTable,
    ) -> Result<RevisionMark> {
        let author = authors.get(author_index).ok_or_else(|| {
            crate::package::DocError::Corrupted(
                "revision author index exceeds SttbfRMark".to_string(),
            )
        })?;
        Ok(RevisionMark {
            kind,
            author_index,
            author: author.to_string(),
            timestamp: packed_timestamp.map(decode_dttm).transpose()?.flatten(),
            reason: revision_id.and_then(RevisionReason::from_raw),
            revision_id,
            revision_save_id,
        })
    }

    /// Check if this run contains an MTEF formula.
    ///
    /// Returns true if this run contains a parsed MTEF formula AST.
    pub fn has_mtef_formula(&self) -> bool {
        self.mtef_formula_latex.is_some()
    }

    /// Get the owned LaTeX rendering for this MTEF formula.
    ///
    /// This replaces the former `'static` AST accessor, whose nodes depended on
    /// allocations owned elsewhere in the document.
    pub fn mtef_formula_latex(&self) -> Option<&str> {
        self.mtef_formula_latex.as_deref()
    }

    /// Replace or remove the owned MTEF rendering.
    pub fn mtef_formula_latex_mut(&mut self) -> &mut Option<Arc<str>> {
        &mut self.mtef_formula_latex
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
        Ok(self.mtef_formula_latex.as_deref().map(str::to_owned))
    }

    /// Convert formula to LaTeX (fallback when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub fn formula_as_latex(&self) -> Result<Option<String>> {
        if self.mtef_formula_latex.is_some() {
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
    pub fn image(&self) -> Option<&crate::image::Image> {
        self.image.as_ref()
    }
}
