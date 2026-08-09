//! Legacy BIFF8 conditional formatting (`CONDFMT` and `CF`).

pub(crate) const CONDFMT_RECORD_TYPE: u16 = 0x01b0;
pub(crate) const CF_RECORD_TYPE: u16 = 0x01b1;
pub(crate) const CONDFMT12_RECORD_TYPE: u16 = 0x0879;
pub(crate) const CF12_RECORD_TYPE: u16 = 0x087a;
pub(crate) const CFEX_RECORD_TYPE: u16 = 0x087b;
pub(super) fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

/// Inclusive worksheet range affected by conditional formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u8,
    pub(super) last_column: u8,
}

impl Range {
    #[must_use]
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    #[must_use]
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    #[must_use]
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    #[must_use]
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

/// Comparison performed by a cell-value conditional formatting rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Comparison {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

/// Type of condition used by a legacy rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleKind {
    CellValue(Comparison),
    Formula,
}

/// Conditional number-format override.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NumberFormat {
    Identifier(u8),
    Custom(String),
}

/// Raw BIFF font differential block with common typed properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    pub(super) raw: Vec<u8>,
    pub(super) name: Option<String>,
}

impl Font {
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    #[must_use]
    pub fn height_twips(&self) -> Option<u32> {
        let value = read_u32(&self.raw, 64);
        (value != u32::MAX).then_some(value)
    }
    #[must_use]
    pub fn is_italic(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0002 != 0
    }
    #[must_use]
    pub fn is_outline(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0008 != 0
    }
    #[must_use]
    pub fn has_shadow(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0010 != 0
    }
    #[must_use]
    pub fn is_struck_out(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0080 != 0
    }
    #[must_use]
    pub fn weight(&self) -> u16 {
        read_u16(&self.raw, 72)
    }
    #[must_use]
    pub fn escapement(&self) -> u16 {
        read_u16(&self.raw, 74)
    }
    #[must_use]
    pub fn underline(&self) -> u8 {
        self.raw[76]
    }
    #[must_use]
    pub fn color_index(&self) -> i32 {
        crate::utils::wrap_u32_to_i32(read_u32(&self.raw, 80))
    }
    #[must_use]
    pub fn raw_data(&self) -> &[u8] {
        &self.raw
    }
}

/// Text alignment differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Alignment {
    pub(super) horizontal: u8,
    pub(super) vertical: u8,
    pub(super) wrap_text: bool,
    pub(super) rotation: u8,
    pub(super) absolute_indent: u8,
    pub(super) relative_indent: i32,
    pub(super) shrink_to_fit: bool,
    pub(super) merge_cell: bool,
    pub(super) reading_order: u8,
}

impl Alignment {
    #[must_use]
    pub fn horizontal(&self) -> u8 {
        self.horizontal
    }
    #[must_use]
    pub fn vertical(&self) -> u8 {
        self.vertical
    }
    #[must_use]
    pub fn wraps_text(&self) -> bool {
        self.wrap_text
    }
    #[must_use]
    pub fn rotation(&self) -> u8 {
        self.rotation
    }
    #[must_use]
    pub fn absolute_indent(&self) -> u8 {
        self.absolute_indent
    }
    #[must_use]
    pub fn relative_indent(&self) -> i32 {
        self.relative_indent
    }
    #[must_use]
    pub fn shrinks_to_fit(&self) -> bool {
        self.shrink_to_fit
    }
    #[must_use]
    pub fn merges_cell(&self) -> bool {
        self.merge_cell
    }
    #[must_use]
    pub fn reading_order(&self) -> u8 {
        self.reading_order
    }
}

/// Cell border differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Border {
    pub(super) styles: [u8; 5],
    pub(super) colors: [u8; 5],
    pub(super) diagonal_down: bool,
    pub(super) diagonal_up: bool,
}

impl Border {
    /// Left, right, top, bottom, and diagonal styles.
    #[must_use]
    pub fn styles(&self) -> &[u8; 5] {
        &self.styles
    }
    /// Left, right, top, bottom, and diagonal color indexes.
    #[must_use]
    pub fn color_indexes(&self) -> &[u8; 5] {
        &self.colors
    }
    #[must_use]
    pub fn has_diagonal_down(&self) -> bool {
        self.diagonal_down
    }
    #[must_use]
    pub fn has_diagonal_up(&self) -> bool {
        self.diagonal_up
    }
}

/// Fill pattern differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Pattern {
    pub(super) fill_pattern: u8,
    pub(super) foreground_color_index: u8,
    pub(super) background_color_index: u8,
}

impl Pattern {
    #[must_use]
    pub fn fill_pattern(&self) -> u8 {
        self.fill_pattern
    }
    #[must_use]
    pub fn foreground_color_index(&self) -> u8 {
        self.foreground_color_index
    }
    #[must_use]
    pub fn background_color_index(&self) -> u8 {
        self.background_color_index
    }
}

/// Cell protection differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Protection {
    pub(super) locked: bool,
    pub(super) hidden: bool,
}

impl Protection {
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.locked
    }
    #[must_use]
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }
}

/// Differential formatting applied when a rule evaluates to true.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub(super) options: u32,
    pub(super) new_border: bool,
    pub(super) number_format: Option<NumberFormat>,
    pub(super) font: Option<Font>,
    pub(super) alignment: Option<Alignment>,
    pub(super) border: Option<Border>,
    pub(super) pattern: Option<Pattern>,
    pub(super) protection: Option<Protection>,
}

impl Style {
    #[must_use]
    pub fn number_format(&self) -> Option<&NumberFormat> {
        self.number_format.as_ref()
    }
    #[must_use]
    pub fn font(&self) -> Option<&Font> {
        self.font.as_ref()
    }
    #[must_use]
    pub fn alignment(&self) -> Option<&Alignment> {
        self.alignment.as_ref()
    }
    #[must_use]
    pub fn border(&self) -> Option<&Border> {
        self.border.as_ref()
    }
    #[must_use]
    pub fn pattern(&self) -> Option<&Pattern> {
        self.pattern.as_ref()
    }
    #[must_use]
    pub fn protection(&self) -> Option<&Protection> {
        self.protection.as_ref()
    }
    #[must_use]
    pub fn applies_border_to_range_outline(&self) -> bool {
        self.new_border
    }
    #[must_use]
    pub fn is_pattern_style_modified(&self) -> bool {
        self.options & 0x0001_0000 == 0
    }
    #[must_use]
    pub fn is_pattern_foreground_modified(&self) -> bool {
        self.options & 0x0002_0000 == 0
    }
    #[must_use]
    pub fn is_pattern_background_modified(&self) -> bool {
        self.options & 0x0004_0000 == 0
    }
}

/// One legacy conditional formatting rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rule {
    pub(super) kind: RuleKind,
    pub(super) style: Style,
    pub(super) formula1_tokens: Vec<u8>,
    pub(super) formula2_tokens: Vec<u8>,
    pub(super) formula1_rendered: Option<String>,
    pub(super) formula2_rendered: Option<String>,
}

impl Rule {
    #[must_use]
    pub fn kind(&self) -> RuleKind {
        self.kind
    }
    #[must_use]
    pub fn style(&self) -> &Style {
        &self.style
    }
    #[must_use]
    pub fn formula1_tokens(&self) -> &[u8] {
        &self.formula1_tokens
    }
    #[must_use]
    pub fn formula2_tokens(&self) -> &[u8] {
        &self.formula2_tokens
    }
    #[must_use]
    pub fn formula1_rendered(&self) -> Option<&str> {
        self.formula1_rendered.as_deref()
    }
    #[must_use]
    pub fn formula2_rendered(&self) -> Option<&str> {
        self.formula2_rendered.as_deref()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Rule12Kind {
    CellValue(Comparison),
    Formula,
    ColorScale,
    DataBar,
    Filter,
    IconSet,
}

/// Office 2007 future conditional-formatting rule. Visual payloads remain inert bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rule12 {
    pub(super) kind: Rule12Kind,
    pub(super) priority: u16,
    pub(super) stop_if_true: bool,
    pub(super) template: u16,
    pub(super) differential_format: Vec<u8>,
    pub(super) formula1_tokens: Vec<u8>,
    pub(super) formula2_tokens: Vec<u8>,
    pub(super) active_formula_tokens: Vec<u8>,
    pub(super) formula1_rendered: Option<String>,
    pub(super) formula2_rendered: Option<String>,
    pub(super) active_formula_rendered: Option<String>,
    pub(super) template_parameters: [u8; 16],
    pub(super) rule_payload: Vec<u8>,
}
impl Rule12 {
    #[must_use]
    pub fn kind(&self) -> Rule12Kind {
        self.kind
    }
    #[must_use]
    pub fn priority(&self) -> u16 {
        self.priority
    }
    #[must_use]
    pub fn stop_if_true(&self) -> bool {
        self.stop_if_true
    }
    #[must_use]
    pub fn template(&self) -> u16 {
        self.template
    }
    #[must_use]
    pub fn differential_format(&self) -> &[u8] {
        &self.differential_format
    }
    #[must_use]
    pub fn formula1_tokens(&self) -> &[u8] {
        &self.formula1_tokens
    }
    #[must_use]
    pub fn formula2_tokens(&self) -> &[u8] {
        &self.formula2_tokens
    }
    #[must_use]
    pub fn active_formula_tokens(&self) -> &[u8] {
        &self.active_formula_tokens
    }
    #[must_use]
    pub fn formula1_rendered(&self) -> Option<&str> {
        self.formula1_rendered.as_deref()
    }
    #[must_use]
    pub fn formula2_rendered(&self) -> Option<&str> {
        self.formula2_rendered.as_deref()
    }
    #[must_use]
    pub fn active_formula_rendered(&self) -> Option<&str> {
        self.active_formula_rendered.as_deref()
    }
    #[must_use]
    pub fn template_parameters(&self) -> &[u8; 16] {
        &self.template_parameters
    }
    #[must_use]
    pub fn rule_payload(&self) -> &[u8] {
        &self.rule_payload
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formatting12 {
    pub(super) identifier: u16,
    pub(super) tough_recalculation: bool,
    pub(super) enclosing_range: Range,
    pub(super) ranges: Vec<Range>,
    pub(super) rules: Vec<Rule12>,
}
impl Formatting12 {
    #[must_use]
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    #[must_use]
    pub fn requires_tough_recalculation(&self) -> bool {
        self.tough_recalculation
    }
    #[must_use]
    pub fn enclosing_range(&self) -> Range {
        self.enclosing_range
    }
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
    #[must_use]
    pub fn rules(&self) -> &[Rule12] {
        &self.rules
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub(super) identifier: u16,
    pub(super) legacy_rule_index: Option<u16>,
    pub(super) priority: u16,
    pub(super) active: bool,
    pub(super) stop_if_true: bool,
    pub(super) template: u8,
    pub(super) differential_format: Vec<u8>,
    pub(super) template_parameters: [u8; 16],
    pub(super) future_rule: Option<Rule12>,
}
impl Extension {
    #[must_use]
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    #[must_use]
    pub fn legacy_rule_index(&self) -> Option<u16> {
        self.legacy_rule_index
    }
    #[must_use]
    pub fn priority(&self) -> u16 {
        self.priority
    }
    #[must_use]
    pub fn active(&self) -> bool {
        self.active
    }
    #[must_use]
    pub fn stop_if_true(&self) -> bool {
        self.stop_if_true
    }
    #[must_use]
    pub fn template(&self) -> u8 {
        self.template
    }
    #[must_use]
    pub fn differential_format(&self) -> &[u8] {
        &self.differential_format
    }
    #[must_use]
    pub fn template_parameters(&self) -> &[u8; 16] {
        &self.template_parameters
    }
    #[must_use]
    pub fn future_rule(&self) -> Option<&Rule12> {
        self.future_rule.as_ref()
    }
}

/// A range set and its one-to-three legacy conditional formatting rules.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formatting {
    pub(super) identifier: u16,
    pub(super) tough_recalculation: bool,
    pub(super) enclosing_range: Range,
    pub(super) ranges: Vec<Range>,
    pub(super) rules: Vec<Rule>,
}

impl Formatting {
    #[must_use]
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    #[must_use]
    pub fn requires_tough_recalculation(&self) -> bool {
        self.tough_recalculation
    }
    #[must_use]
    pub fn enclosing_range(&self) -> Range {
        self.enclosing_range
    }
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
    #[must_use]
    pub fn rules(&self) -> &[Rule] {
        &self.rules
    }
}
