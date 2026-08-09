#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive legacy automatic-layout compatibility requests.
///
/// These flags are retained for round trips only. This crate does not change
/// shape, footnote, paragraph-spacing, or tab layout in response to them.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentLegacyLayoutCompatibility {
    /// `\splytwnine`: do not lay out `AutoShapes` using Word 97 behavior.
    pub do_not_use_word_97_shape_layout: bool,
    /// `\ftnlytwnine`: use pre-Word 2000 footnote layout behavior.
    pub use_legacy_footnote_layout: bool,
    /// `\htmautsp`: use HTML paragraph automatic spacing.
    pub use_html_paragraph_auto_spacing: bool,
    /// `\useltbaln`: preserve the last tab alignment.
    pub preserve_last_tab_alignment: bool,
    /// `\oldas`: use Word 95 automatic spacing.
    pub use_word_95_auto_spacing: bool,
}

impl DocumentLegacyLayoutCompatibility {
    /// Return whether every legacy layout request was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.do_not_use_word_97_shape_layout
            && !self.use_legacy_footnote_layout
            && !self.use_html_paragraph_auto_spacing
            && !self.preserve_last_tab_alignment
            && !self.use_word_95_auto_spacing
    }
}
