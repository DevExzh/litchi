/// Passive Asian grid and line-breaking compatibility requests.
///
/// These flags are retained for round trips only. This crate does not apply
/// Thai or Asian line breaking, character-grid snapping, or punctuation layout.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentAsianGridCompatibility {
    /// `\ApplyBrkRules`: use line-breaking rules compatible with Thai text.
    pub apply_thai_line_breaking_rules: bool,
    /// `\snaptogridincell`: snap text to the grid inside table cells.
    pub snap_text_to_grid_inside_table: bool,
    /// `\wrppunct`: allow hanging punctuation in the character grid.
    pub allow_hanging_punctuation: bool,
    /// `\asianbrkrule`: use Asian line-breaking rules with the character grid.
    pub use_asian_line_breaking_rules: bool,
    /// `\toplinepunct`: compress punctuation at the start of a line.
    pub compress_punctuation_at_line_start: bool,
}

impl DocumentAsianGridCompatibility {
    /// Return whether every Asian grid compatibility request was omitted.
    pub fn is_empty(&self) -> bool {
        !self.apply_thai_line_breaking_rules
            && !self.snap_text_to_grid_inside_table
            && !self.allow_hanging_punctuation
            && !self.use_asian_line_breaking_rules
            && !self.compress_punctuation_at_line_start
    }
}
