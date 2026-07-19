/// Passive document-output and compatibility flags from the RTF header.
///
/// These values are retained for round-tripping only. In particular, `muser`
/// does not change this crate's compatibility behavior and `psover` does not
/// cause this crate to perform any printing.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentOutputSettings {
    /// Whether the `muser` Word 97 compatibility marker was present.
    pub word97_compatibility_marker: bool,
    /// Whether PostScript-over-text output was requested with `psover`.
    pub postscript_over_text: bool,
}

impl DocumentOutputSettings {
    /// Return whether neither passive output flag is present.
    pub fn is_empty(&self) -> bool {
        !self.word97_compatibility_marker && !self.postscript_over_text
    }
}
