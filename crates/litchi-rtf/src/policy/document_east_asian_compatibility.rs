#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive Word 6-era East Asian typography compatibility requests.
///
/// These flags are retained for round trips only. This crate does not apply
/// legacy character balancing, spacing, underlining, character translation,
/// or line-breaking behavior.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentEastAsianCompatibility {
    /// `\dntblnsbdb`: do not balance SBCS and DBCS characters.
    pub do_not_balance_sbcs_dbcs: bool,
    /// `\expshrtn`: expand character spacing at SHIFT+RETURN line endings.
    pub expand_spacing_at_shift_return: bool,
    /// `\nospaceforul`: do not add space for underlining.
    pub do_not_add_space_for_underline: bool,
    /// `\noultrlspc`: do not underline trailing spaces.
    pub do_not_underline_trailing_spaces: bool,
    /// `\noxlattoyen`: do not translate backslash to a Yen sign.
    pub do_not_translate_backslash_to_yen: bool,
    /// `\lnbrkrule`: use pre-Word 97 Asian line-breaking rules.
    pub use_legacy_line_breaking_rules: bool,
}

impl DocumentEastAsianCompatibility {
    /// Return whether every compatibility request was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.do_not_balance_sbcs_dbcs
            && !self.expand_spacing_at_shift_return
            && !self.do_not_add_space_for_underline
            && !self.do_not_underline_trailing_spaces
            && !self.do_not_translate_backslash_to_yen
            && !self.use_legacy_line_breaking_rules
    }
}
