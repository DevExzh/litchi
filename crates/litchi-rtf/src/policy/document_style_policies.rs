#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive document style and theme-editing policies.
///
/// These values are retained for round-tripping only. This crate does not
/// lock, replace, remove, or automatically apply any theme or style.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentStylePolicies {
    /// `\linkstyles`: styles are requested to update from a linked template.
    /// This crate retains the request but never resolves or applies a template.
    pub update_styles_from_template: bool,
    /// Whether modification of theme information is locked (`stylelocktheme`).
    pub lock_theme: bool,
    /// Whether replacement of the complete quick-format style set is locked (`stylelockqfset`).
    pub lock_quick_format_set: bool,
    /// Whether numbered Normal paragraphs retain Normal rather than an alternate list style.
    pub use_normal_style_for_lists: bool,
}

impl DocumentStylePolicies {
    /// Return whether all three policy flags were omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.update_styles_from_template
            && !self.lock_theme
            && !self.lock_quick_format_set
            && !self.use_normal_style_for_lists
    }
}
