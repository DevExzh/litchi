#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive compatibility requests for Word 2003-era layout behavior.
///
/// These flags are retained for round trips only. This crate does not alter
/// tables, floating objects, numbering, line breaking, typography, or pagination.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentWord2003Compatibility {
    /// `\noafcnsttbl`: do not resize `AutoFit` tables around floating shapes.
    pub preserve_autofit_table_width_around_shapes: bool,
    /// `\noindnmbrts`: use hanging indents as numbering tab stops.
    pub use_hanging_indent_as_numbering_tab: bool,
    /// `\felnbrelev`: use alternate East Asian kinsoku characters.
    pub use_legacy_kinsoku_characters: bool,
    /// `\indrlsweleven`: use legacy paragraph indentation around floating objects.
    pub use_legacy_floating_object_indentation: bool,
    /// `\nocxsptable`: allow contextual paragraph spacing inside tables.
    pub allow_contextual_spacing_in_tables: bool,
    /// `\notcvasp`: ignore cell vertical alignment when a floating object is present.
    pub ignore_cell_vertical_alignment_with_floating_objects: bool,
    /// `\notvatxbx`: ignore vertical alignment in text boxes.
    pub ignore_text_box_vertical_alignment: bool,
    /// `\spltpgpar`: move a paragraph mark after a terminal page break.
    pub split_page_break_paragraph: bool,
    /// `\hwelev`: use fixed-width Hangul syllables.
    pub use_fixed_width_hangul: bool,
    /// `\afelev`: use legacy `AutoFit` width expansion.
    pub use_legacy_autofit_width_expansion: bool,
    /// `\cachedcolbal`: use cached paragraph data for column balancing.
    pub use_cached_column_balancing: bool,
    /// `\utinl`: underline the generated numbering suffix when applicable.
    pub underline_numbering_suffix: bool,
    /// `\notbrkcnstfrctbl`: do not split tall rows around floating tables.
    pub do_not_split_rows_around_floating_tables: bool,
    /// `\krnprsnet`: use ANSI rather than Unicode font kerning pairs.
    pub use_ansi_kerning_pairs: bool,
}

impl DocumentWord2003Compatibility {
    /// Return whether every Word 2003 compatibility request was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.preserve_autofit_table_width_around_shapes
            && !self.use_hanging_indent_as_numbering_tab
            && !self.use_legacy_kinsoku_characters
            && !self.use_legacy_floating_object_indentation
            && !self.allow_contextual_spacing_in_tables
            && !self.ignore_cell_vertical_alignment_with_floating_objects
            && !self.ignore_text_box_vertical_alignment
            && !self.split_page_break_paragraph
            && !self.use_fixed_width_hangul
            && !self.use_legacy_autofit_width_expansion
            && !self.use_cached_column_balancing
            && !self.underline_numbering_suffix
            && !self.do_not_split_rows_around_floating_tables
            && !self.use_ansi_kerning_pairs
    }
}
