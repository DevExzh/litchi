#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive legacy document-level line-spacing compatibility requests.
///
/// These flags are retained for round trips only. This crate does not change
/// line layout, paragraph spacing, page breaking, or raised/lowered text.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentLineSpacingCompatibility {
    /// `\noextrasprl`: do not add extra line height for raised/lowered text.
    pub suppress_extra_spacing_for_raised_lowered_text: bool,
    /// `\sprstsp`: suppress extra line spacing at the top of a page.
    pub suppress_extra_spacing_at_top_of_page: bool,
    /// `\sprsspbf`: suppress paragraph space-before after a hard break.
    pub suppress_space_before_after_hard_break: bool,
    /// `\sprslnsp`: suppress extra line spacing using `WordPerfect` 5.x rules.
    pub suppress_wordperfect_extra_line_spacing: bool,
    /// `\sprsbsp`: suppress extra line spacing at the bottom of a page.
    pub suppress_extra_spacing_at_bottom_of_page: bool,
}

impl DocumentLineSpacingCompatibility {
    /// Return whether no legacy line-spacing compatibility request is present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.suppress_extra_spacing_for_raised_lowered_text
            && !self.suppress_extra_spacing_at_top_of_page
            && !self.suppress_space_before_after_hard_break
            && !self.suppress_wordperfect_extra_line_spacing
            && !self.suppress_extra_spacing_at_bottom_of_page
    }
}
