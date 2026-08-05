//! IWA-native adapters for archive-free paragraph flow values.

pub use litchi_iwa_text::paragraph::flow::{
    Flow as ParagraphFlow, Hyphenation as ParagraphHyphenation,
};

/// Decode native iWork's boolean automatic-hyphenation field.
pub(crate) const fn hyphenation_from_native(value: bool) -> ParagraphHyphenation {
    if value {
        ParagraphHyphenation::Automatic
    } else {
        ParagraphHyphenation::Prevented
    }
}

/// Encode the semantic hyphenation policy into native iWork's boolean field.
pub(crate) const fn hyphenation_to_native(value: ParagraphHyphenation) -> bool {
    matches!(value, ParagraphHyphenation::Automatic)
}
