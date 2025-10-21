use crate::formula::ast::{AccentType, Fence, LargeOperator, MathNode, MatrixFence};

/// Element types in OMML
///
/// This enum represents all possible OMML elements that can appear in
/// Office Math Markup Language documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)] // Some variants constructed indirectly or reserved for future use
pub enum ElementType {
    Math,
    Run,
    Text,
    Fraction,
    Numerator,
    Denominator,
    Radical,
    Degree,
    Base,
    Superscript,
    Subscript,
    SubSup,
    SuperscriptElement,
    SubscriptElement,
    Delimiter,
    Nary,
    LowerLimit,
    UpperLimit,
    LimLow,
    LimUpp,
    Integrand,
    Function,
    FunctionName,
    Matrix,
    MatrixRow,
    MatrixCell,
    Accent,
    AccentProperties,
    Bar,
    Box,
    Properties,
    Phantom,
    GroupChar,
    BorderBox,
    EqArr,
    EqArrPr,
    Limit,
    Spacing,
    PreScript,
    PostScript,
    Character,
    Position,
    VerticalAlignment,
    Lit,
    Scr,
    Sty,
    Nor,
    Unknown,
}

/// Properties for OMML elements
///
/// This struct contains all possible properties that can be specified in OMML elements.
/// Not all properties are used by all elements; each element type uses only the relevant
/// properties for its specific purpose.
#[derive(Debug, Clone, Default)]
#[allow(dead_code)] // Many fields are used conditionally based on element type
pub struct ElementProperties {
    // Style and formatting
    pub style: Option<String>,
    pub size: Option<String>,
    pub color: Option<String>,
    pub font: Option<String>,

    // Layout and positioning
    pub spacing: Option<String>,
    pub alignment: Option<String>,
    pub vertical_alignment: Option<String>,

    // Visibility and rendering
    pub hide: Option<bool>,
    pub strike_through: Option<bool>,
    pub double_strike_through: Option<bool>,
    pub underline: Option<String>,
    pub overline: Option<String>,

    // Characters and symbols
    pub chr: Option<String>, // Character for accents, operators, etc.

    // Math-specific properties
    pub math_variant: Option<String>,
    pub script_level: Option<i32>,
    pub display_style: Option<bool>,

    // Size and scaling
    pub min_size: Option<String>,
    pub max_size: Option<String>,

    // Spacing and margins
    pub left_margin: Option<String>,
    pub right_margin: Option<String>,
    pub top_margin: Option<String>,
    pub bottom_margin: Option<String>,

    // Operator properties
    pub operator_form: Option<String>, // prefix, infix, postfix
    pub operator_spacing: Option<String>,
    pub operator_stretch: Option<String>,

    // Fraction properties
    pub fraction_line_thickness: Option<String>,
    pub fraction_type: Option<String>, // bar, noBar, skewed

    // Matrix properties
    pub matrix_alignment: Option<String>,
    pub matrix_row_spacing: Option<String>,
    pub matrix_column_spacing: Option<String>,

    // Accent properties
    pub accent_position: Option<String>, // top, bottom

    // Box properties
    pub box_alignment: Option<String>,
    pub box_differential: Option<bool>,
    pub box_operator_emulation: Option<bool>,
    pub box_break: Option<bool>,
    pub box_no_break: Option<bool>,

    // Phantom properties
    pub phantom_show: Option<bool>,
    pub phantom_zero_width: Option<bool>,
    pub phantom_zero_ascent: Option<bool>,
    pub phantom_zero_descent: Option<bool>,
    pub phantom_transparent: Option<bool>,

    // Border box properties
    pub border_hide_top: Option<bool>,
    pub border_hide_bottom: Option<bool>,
    pub border_hide_left: Option<bool>,
    pub border_hide_right: Option<bool>,
    pub border_strike_horizontal: Option<bool>,
    pub border_strike_vertical: Option<bool>,
    pub border_strike_bltr: Option<bool>, // bottom-left to top-right
    pub border_strike_tlbr: Option<bool>, // top-left to bottom-right

    // Equation array properties
    pub eq_arr_base_alignment: Option<String>,
    pub eq_arr_max_distance: Option<String>,
    pub eq_arr_object_distance: Option<String>,
    pub eq_arr_row_spacing: Option<String>,
    pub eq_arr_row_spacing_rule: Option<String>,

    // N-ary operator properties
    pub nary_hide_sub: Option<bool>,
    pub nary_hide_sup: Option<bool>,
    pub nary_operator_grow: Option<bool>,

    // Delimiter properties
    pub delimiter_grow: Option<bool>,
    pub delimiter_shape: Option<String>, // centered, match
    pub delimiter_separator_char: Option<String>,
    pub delimiter_open_char: Option<String>,
    pub delimiter_close_char: Option<String>,

    // Radical properties
    pub radical_hide_degree: Option<bool>,

    // Run properties
    pub run_literal: Option<bool>,
    pub run_normal_text: Option<String>,
    pub run_math_style: Option<String>,
}

/// High-performance text buffer for collecting text content
#[derive(Debug)]
pub struct TextBuffer {
    buffer: String,
    has_content: bool,
}

impl TextBuffer {
    pub fn new() -> Self {
        Self {
            buffer: String::new(),
            has_content: false,
        }
    }

    pub fn push_str(&mut self, s: &str) {
        if !s.is_empty() {
            self.buffer.push_str(s);
            self.has_content = true;
        }
    }

    pub fn clear(&mut self) {
        self.buffer.clear();
        self.has_content = false;
    }

    pub fn is_empty(&self) -> bool {
        !self.has_content
    }

    pub fn as_str(&self) -> &str {
        &self.buffer
    }

    /// Convert into owned String (used for final text extraction)
    #[inline]
    #[allow(dead_code)]
    pub fn into_string(self) -> String {
        self.buffer
    }
}

impl Default for TextBuffer {
    fn default() -> Self {
        Self::new()
    }
}

/// Context for tracking element state during parsing
#[derive(Debug)]
pub struct ElementContext<'arena> {
    pub element_type: ElementType,
    pub children: Vec<MathNode<'arena>>,
    pub text: TextBuffer,
    pub properties: ElementProperties,
    pub attributes: Vec<quick_xml::events::attributes::Attribute<'static>>,

    // Core mathematical components
    pub base: Option<Vec<MathNode<'arena>>>,
    pub numerator: Option<Vec<MathNode<'arena>>>,
    pub denominator: Option<Vec<MathNode<'arena>>>,
    pub subscript: Option<Vec<MathNode<'arena>>>,
    pub superscript: Option<Vec<MathNode<'arena>>>,
    pub degree: Option<Vec<MathNode<'arena>>>,
    pub lower_limit: Option<Vec<MathNode<'arena>>>,
    pub upper_limit: Option<Vec<MathNode<'arena>>>,
    pub integrand: Option<Vec<MathNode<'arena>>>,

    // Matrix and array structures
    pub matrix_rows: Vec<Vec<Vec<MathNode<'arena>>>>,
    pub matrix_cells: Vec<MathNode<'arena>>,
    pub eq_array_rows: Vec<Vec<MathNode<'arena>>>,

    // Function and operator data
    pub function_name: Option<String>,
    pub operator: Option<LargeOperator>,
    pub accent_type: Option<AccentType>,

    // Fencing and delimiters
    pub fence_open: Option<Fence>,
    pub fence_close: Option<Fence>,
    pub matrix_fence: Option<MatrixFence>,

    // Scripts and limits
    pub pre_scripts: Vec<MathNode<'arena>>,
    pub post_scripts: Vec<MathNode<'arena>>,

    // Spacing and layout
    pub spacing_nodes: Vec<MathNode<'arena>>,

    // Group and character data
    pub character_data: Option<String>,
}

impl<'arena> ElementContext<'arena> {
    pub fn new(element_type: ElementType) -> Self {
        Self {
            element_type,
            children: Vec::new(),
            text: TextBuffer::new(),
            properties: ElementProperties::default(),
            attributes: Vec::new(),
            base: None,
            numerator: None,
            denominator: None,
            subscript: None,
            superscript: None,
            degree: None,
            lower_limit: None,
            upper_limit: None,
            integrand: None,
            matrix_rows: Vec::new(),
            matrix_cells: Vec::new(),
            eq_array_rows: Vec::new(),
            function_name: None,
            operator: None,
            accent_type: None,
            fence_open: None,
            fence_close: None,
            matrix_fence: None,
            pre_scripts: Vec::new(),
            post_scripts: Vec::new(),
            spacing_nodes: Vec::new(),
            character_data: None,
        }
    }

    /// Clear the context for reuse
    pub fn clear(&mut self) {
        self.children.clear();
        self.text.clear();
        self.properties = ElementProperties::default();
        self.attributes.clear();
        self.base = None;
        self.numerator = None;
        self.denominator = None;
        self.subscript = None;
        self.superscript = None;
        self.degree = None;
        self.lower_limit = None;
        self.upper_limit = None;
        self.integrand = None;
        self.matrix_rows.clear();
        self.matrix_cells.clear();
        self.eq_array_rows.clear();
        self.function_name = None;
        self.fence_open = None;
        self.fence_close = None;
        self.operator = None;
        self.accent_type = None;
        self.matrix_fence = None;
        self.pre_scripts.clear();
        self.post_scripts.clear();
        self.spacing_nodes.clear();
        self.character_data = None;
    }

    /// Check if the context has any content
    #[inline]
    #[allow(dead_code)]
    pub fn has_content(&self) -> bool {
        !self.children.is_empty()
            || !self.text.is_empty()
            || self.base.is_some()
            || self.numerator.is_some()
            || self.denominator.is_some()
            || self.subscript.is_some()
            || self.superscript.is_some()
            || self.degree.is_some()
            || self.lower_limit.is_some()
            || self.upper_limit.is_some()
            || self.integrand.is_some()
            || !self.matrix_rows.is_empty()
            || !self.matrix_cells.is_empty()
            || !self.eq_array_rows.is_empty()
            || self.function_name.is_some()
            || self.character_data.is_some()
    }

    /// Get the total number of child nodes across all collections
    #[inline]
    #[allow(dead_code)]
    pub fn total_child_count(&self) -> usize {
        self.children.len()
            + self.pre_scripts.len()
            + self.post_scripts.len()
            + self.spacing_nodes.len()
            + self.matrix_cells.len()
            + self
                .eq_array_rows
                .iter()
                .map(|row| row.len())
                .sum::<usize>()
            + self
                .matrix_rows
                .iter()
                .map(|row| row.iter().map(|cell| cell.len()).sum::<usize>())
                .sum::<usize>()
    }

    /// Reserve capacity for expected number of children
    /// Used for performance optimization when the number of children is known in advance
    #[inline]
    #[allow(dead_code)]
    pub fn reserve_children(&mut self, capacity: usize) {
        self.children.reserve(capacity);
    }

    /// Reserve capacity for matrix rows
    /// Used when parsing matrix structures to avoid reallocations
    #[inline]
    #[allow(dead_code)]
    pub fn reserve_matrix_rows(&mut self, capacity: usize) {
        self.matrix_rows.reserve(capacity);
    }

    /// Reserve capacity for equation array rows
    /// Used when parsing equation arrays to avoid reallocations
    #[inline]
    #[allow(dead_code)]
    pub fn reserve_eq_array_rows(&mut self, capacity: usize) {
        self.eq_array_rows.reserve(capacity);
    }

    /// Check if this is a structural element (contains other elements)
    #[inline]
    #[allow(dead_code)]
    pub fn is_structural(&self) -> bool {
        matches!(
            self.element_type,
            ElementType::Math
                | ElementType::Fraction
                | ElementType::Radical
                | ElementType::Superscript
                | ElementType::Subscript
                | ElementType::SubSup
                | ElementType::Delimiter
                | ElementType::Nary
                | ElementType::Function
                | ElementType::Matrix
                | ElementType::Accent
                | ElementType::Bar
                | ElementType::Box
                | ElementType::Phantom
                | ElementType::GroupChar
                | ElementType::BorderBox
                | ElementType::EqArr
        )
    }

    /// Check if this is a leaf element (contains only text/symbols)
    #[inline]
    #[allow(dead_code)]
    pub fn is_leaf(&self) -> bool {
        matches!(
            self.element_type,
            ElementType::Text | ElementType::Character | ElementType::Run
        )
    }
}
