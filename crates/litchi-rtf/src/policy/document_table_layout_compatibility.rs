#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive legacy table-layout compatibility requests.
///
/// These flags are retained for round trips only. This crate does not alter
/// table borders, widths, row placement, line heights, autofit, or styles in
/// response to them.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentTableLayoutCompatibility {
    /// `\otblrul`: combine table borders using Word 5.x for Macintosh rules.
    pub combine_borders_like_word_5: bool,
    /// `\alntblind`: do not align table rows independently.
    pub do_not_align_rows_independently: bool,
    /// `\lytcalctblwd`: do not lay out tables using raw width.
    pub do_not_use_raw_table_width: bool,
    /// `\lyttblrtgr`: do not allow table rows to lay out apart.
    pub keep_rows_together: bool,
    /// `\nolnhtadjtbl`: do not adjust line height in tables.
    pub do_not_adjust_line_height: bool,
    /// `\nobrkwrptbl`: do not break wrapped tables across pages.
    pub do_not_break_wrapped_tables_across_pages: bool,
    /// `\nogrowautofit`: do not let autofit tables grow into page margins.
    pub prevent_autofit_growth_into_margins: bool,
    /// `\newtblstyruls`: use the table-style rules introduced by Word 2003.
    pub use_word_2003_table_style_rules: bool,
}

impl DocumentTableLayoutCompatibility {
    /// Return whether every table-layout compatibility request was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.combine_borders_like_word_5
            && !self.do_not_align_rows_independently
            && !self.do_not_use_raw_table_width
            && !self.keep_rows_together
            && !self.do_not_adjust_line_height
            && !self.do_not_break_wrapped_tables_across_pages
            && !self.prevent_autofit_growth_into_margins
            && !self.use_word_2003_table_style_rules
    }
}
