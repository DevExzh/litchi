/// Passive document-level style and formatting restriction declarations.
///
/// These values preserve RTF metadata only. They do not restrict editing or
/// cause this crate to enforce, apply, or synthesize any protection behavior.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentStyleRestrictions {
    /// `\stylelock`: the document declares style and formatting restrictions.
    pub restrictions_present: bool,
    /// `\stylelockenforced`: the declared restrictions are marked as enforced.
    pub enforced: bool,
    /// `\stylelockbackcomp`: legacy protection keywords were emitted for
    /// compatibility with older readers.
    pub backward_compatibility: bool,
    /// `\autofmtoverride`: AutoFormat is permitted to override the declared
    /// style restrictions. This is retained as metadata only.
    pub allow_auto_format_override: bool,
}

impl DocumentStyleRestrictions {
    /// Return whether no style-restriction declaration was present.
    pub fn is_empty(&self) -> bool {
        !self.restrictions_present
            && !self.enforced
            && !self.backward_compatibility
            && !self.allow_auto_format_override
    }
}
