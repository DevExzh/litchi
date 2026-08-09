//! Typed `PowerPoint` text-run data models and ergonomic constructors.

/// Text formatting properties for a text run.
///
/// Based on Apache POI's `TextPropCollection` and `CharacterPropertyBags`.
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool mirrors a distinct `CFStyle` bit flag from MS-PPT (bold, italic, \
              underline, shadow, embossed); collapsing them into enums would obscure the \
              one-to-one mapping with the on-disk bit field"
)]
#[derive(Debug, Clone, Default)]
pub struct TextRunFormatting {
    /// Original `CFMasks` value.
    pub property_mask: u32,
    /// Raw `CFStyle` value, when present.
    pub font_style_raw: Option<u16>,
    /// Font size in points
    pub font_size: Option<u16>,
    /// Font color (RGB)
    pub font_color: Option<u32>,
    /// Raw `PowerPoint` `ColorIndexStruct` value
    pub font_color_raw: Option<u32>,
    /// `PowerPoint` color-scheme index when the color is not direct sRGB
    pub font_scheme_color: Option<u8>,
    /// Bold formatting
    pub bold: bool,
    /// Explicit bold value, or `None` when inherited.
    pub bold_explicit: Option<bool>,
    /// Italic formatting
    pub italic: bool,
    /// Explicit italic value, or `None` when inherited.
    pub italic_explicit: Option<bool>,
    /// Underline formatting
    pub underline: bool,
    /// Explicit underline value, or `None` when inherited.
    pub underline_explicit: Option<bool>,
    /// Shadow formatting
    pub shadow: bool,
    /// Explicit shadow value, or `None` when inherited.
    pub shadow_explicit: Option<bool>,
    /// Whether the run originated from double-byte input, when specified.
    pub fe_hint: Option<bool>,
    /// Whether Kumimoji formatting is active, when specified.
    pub kumi: Option<bool>,
    /// De facto legacy strikethrough value from the MS-PPT `unused3` bit.
    pub legacy_strikethrough: Option<bool>,
    /// Embossed/relief formatting
    pub embossed: bool,
    /// Explicit emboss value, or `None` when inherited.
    pub embossed_explicit: Option<bool>,
    /// `PowerPoint` 9 additional-property run grouping identifier.
    pub pp9_run_id: Option<u8>,
    /// Baseline position as a percentage of line height
    pub baseline_position: Option<i16>,
    /// Font name
    pub font_name: Option<String>,
    /// Zero-based font reference in the `PowerPoint` font collection
    pub font_index: Option<u16>,
    /// East Asian font reference
    pub asian_font_index: Option<u16>,
    /// ANSI font reference
    pub ansi_font_index: Option<u16>,
    /// Symbol font reference
    pub symbol_font_index: Option<u16>,
}

/// Paragraph alignment stored by `TextPFException`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphAlignment {
    /// Left for horizontal text, top for vertical text.
    Left,
    /// Center for horizontal text, middle for vertical text.
    Center,
    /// Right for horizontal text, bottom for vertical text.
    Right,
    /// Flush both horizontal or vertical edges.
    Justify,
    /// Distribute space between characters.
    Distributed,
    /// Thai distributed justification.
    ThaiDistributed,
    /// Low Kashida justification.
    JustifyLow,
}

/// Vertical placement of characters within the line height.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphFontAlignment {
    /// Place characters on the font baseline.
    Roman,
    /// Hang characters from the top of the line.
    Hanging,
    /// Center characters within the line height.
    Center,
    /// Anchor characters to the bottom of the line.
    UpholdFixed,
}

/// Paragraph text direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTextDirection {
    /// Left-to-right text flow.
    LeftToRight,
    /// Right-to-left text flow.
    RightToLeft,
}

/// Alignment at a paragraph tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTabAlignment {
    /// Left-aligned tab stop.
    Left,
    /// Center-aligned tab stop.
    Center,
    /// Right-aligned tab stop.
    Right,
    /// Decimal-point-aligned tab stop.
    Decimal,
}

/// A paragraph tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphTabStop {
    /// Signed offset in `PowerPoint` master units.
    pub position: i16,
    /// How text aligns at the stop.
    pub alignment: ParagraphTabAlignment,
}

/// Formatting explicitly carried by one `PowerPoint` paragraph run.
#[derive(Debug, Clone, Default)]
pub struct ParagraphRunFormatting {
    /// Original `PFMasks` value.
    pub property_mask: u32,
    /// Paragraph indentation level.
    pub indent_level: u16,
    /// Raw `BulletFlags` value, when present.
    pub bullet_flags_raw: Option<u16>,
    /// Whether the paragraph has a bullet, when explicitly specified.
    pub bullet_enabled: Option<bool>,
    /// Whether the bullet font override is active, when explicitly specified.
    pub bullet_font_enabled: Option<bool>,
    /// Whether the bullet color override is active, when explicitly specified.
    pub bullet_color_enabled: Option<bool>,
    /// Whether the bullet size override is active, when explicitly specified.
    pub bullet_size_enabled: Option<bool>,
    /// Raw UTF-16 code unit used as the bullet character.
    pub bullet_character: Option<u16>,
    /// Zero-based bullet font reference.
    pub bullet_font_index: Option<u16>,
    /// Raw `BulletSize` value.
    pub bullet_size: Option<i16>,
    /// Normalized direct bullet color in `0xRRGGBB` form.
    pub bullet_color: Option<u32>,
    /// Raw bullet `ColorIndexStruct` value.
    pub bullet_color_raw: Option<u32>,
    /// Bullet color-scheme index when the color is not direct sRGB.
    pub bullet_scheme_color: Option<u8>,
    /// Paragraph alignment.
    pub alignment: Option<ParagraphAlignment>,
    /// Raw line-spacing value.
    pub line_spacing: Option<i16>,
    /// Raw space-before value.
    pub space_before: Option<i16>,
    /// Raw space-after value.
    pub space_after: Option<i16>,
    /// Left margin in master units.
    pub left_margin: Option<i16>,
    /// First-line indent in master units.
    pub indent: Option<i16>,
    /// Default tab size in master units.
    pub default_tab_size: Option<i16>,
    /// Explicit paragraph tab stops.
    pub tab_stops: Option<Vec<ParagraphTabStop>>,
    /// Character alignment within the line height.
    pub font_alignment: Option<ParagraphFontAlignment>,
    /// Whether East Asian character wrapping is active, when explicitly specified.
    pub character_wrap: Option<bool>,
    /// Whether wrapping occurs at word boundaries, when explicitly specified.
    pub word_wrap: Option<bool>,
    /// Whether hanging punctuation is allowed, when explicitly specified.
    pub overflow: Option<bool>,
    /// Raw `PFWrapFlags` value, when present.
    pub wrap_flags_raw: Option<u16>,
    /// Paragraph text direction.
    pub text_direction: Option<ParagraphTextDirection>,
}

impl ParagraphRunFormatting {
    /// Decode the bullet UTF-16 code unit as a Unicode scalar value.
    ///
    /// Returns `None` when no bullet character is present or the stored unit is
    /// an unpaired surrogate.
    #[must_use]
    pub fn bullet_char(&self) -> Option<char> {
        char::from_u32(u32::from(self.bullet_character?))
    }
}

/// A text range carrying paragraph-level formatting.
#[derive(Debug, Clone)]
pub struct ParagraphRun {
    /// Text covered by this paragraph style, including any stored paragraph marker.
    pub text: String,
    /// Paragraph formatting properties.
    pub formatting: ParagraphRunFormatting,
    /// Start index in Unicode scalar values in the full text.
    pub start_index: usize,
    /// Length in Unicode scalar values.
    pub length: usize,
}

impl ParagraphRun {
    /// Create a paragraph run with explicit formatting.
    #[must_use]
    pub fn with_formatting(
        text: String,
        start_index: usize,
        formatting: ParagraphRunFormatting,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
        }
    }
}

/// A text run with formatting.
///
/// Based on Apache POI's `RichTextRun`.
///
/// # Example
///
/// ```rust,ignore
/// let run = TextRun::new("x^2 + y^2 = z^2".to_string(), 0);
/// println!("Text: {}", run.text);
///
/// // Check for embedded MTEF formulas
/// if let Some(formula_ast) = run.mtef_formula_ast() {
///     println!("MTEF formula AST with {} nodes", formula_ast.len());
/// }
/// ```
#[derive(Debug, Clone)]
pub struct TextRun {
    /// Text content
    pub text: String,
    /// Formatting properties
    pub formatting: TextRunFormatting,
    /// Start index in the full text
    pub start_index: usize,
    /// Length in characters
    pub length: usize,
    /// Parsed MTEF formula AST (if this run contains a formula)
    #[cfg(feature = "formula")]
    mtef_formula_ast: Option<Vec<litchi_formula::MathNode<'static>>>,
    /// Parsed MTEF formula AST placeholder (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    mtef_formula_ast: Option<Vec<()>>,
}

impl TextRun {
    /// Create a new text run.
    #[must_use]
    pub fn new(text: String, start_index: usize) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting: TextRunFormatting::default(),
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Create a text run with formatting.
    #[must_use]
    pub fn with_formatting(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Create a text run with MTEF formula AST.
    #[cfg(feature = "formula")]
    #[must_use]
    pub fn with_mtef_formula(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
        mtef_ast: Vec<litchi_formula::MathNode<'static>>,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: Some(mtef_ast),
        }
    }

    /// Create a text run with MTEF formula AST fallback (when formula feature is disabled).
    #[cfg(not(feature = "formula"))]
    pub fn with_mtef_formula(
        text: String,
        start_index: usize,
        formatting: TextRunFormatting,
        _mtef_ast: Vec<()>,
    ) -> Self {
        let length = text.chars().count();
        Self {
            text,
            formatting,
            start_index,
            length,
            mtef_formula_ast: None,
        }
    }

    /// Check if this text run contains an MTEF formula.
    ///
    /// Returns true if this run contains a parsed MTEF formula AST.
    #[must_use]
    pub fn has_mtef_formula(&self) -> bool {
        self.mtef_formula_ast.is_some()
    }

    /// Get the MTEF formula AST if this run contains a formula.
    ///
    /// Returns the parsed MTEF formula as AST nodes if this run contains a `MathType` equation,
    /// None otherwise.
    #[cfg(feature = "formula")]
    #[must_use]
    pub fn mtef_formula_ast(&self) -> Option<&Vec<litchi_formula::MathNode<'static>>> {
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
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<litchi_formula::MathNode<'static>>> {
        &mut self.mtef_formula_ast
    }

    #[cfg(not(feature = "formula"))]
    pub fn mtef_formula_ast_mut(&mut self) -> &mut Option<Vec<()>> {
        &mut self.mtef_formula_ast
    }
}
