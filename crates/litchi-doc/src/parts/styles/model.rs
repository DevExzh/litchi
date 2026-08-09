//! Typed stylesheet records and the parsed stylesheet container.

use crate::leniency::ToleranceReport;

/// General stylesheet information stored in `Stshif`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleSheetHeader {
    /// Number of entries in [`StyleSheet::styles`].
    pub style_count: u16,
    /// Size of each fixed `Stdf` prefix (10 or 18 bytes).
    pub stdf_size: u16,
    /// Largest built-in style identifier known when the file was saved, plus one.
    pub max_builtin_style: u16,
    /// Count of fixed-index style slots. This is always 15 for Word 97+.
    pub fixed_style_count: u16,
    /// Built-in style-name version.
    pub builtin_name_version: u16,
    /// Default ASCII font index.
    pub ascii_font: i16,
    /// Default East Asian font index.
    pub east_asian_font: i16,
    /// Default non-ASCII font index.
    pub other_font: i16,
}

/// The four style kinds encoded by `StdfBase.stk`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleKind {
    /// Paragraph style.
    Paragraph,
    /// Character style.
    Character,
    /// Table style.
    Table,
    /// Numbering style.
    Numbering,
}

/// Miscellaneous flags stored in `StdfBase` and `GRFSTD`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct StyleFlags {
    /// Paragraph heights for this style need to be recalculated.
    pub invalidate_height: bool,
    /// User formatting is automatically merged into the style.
    pub auto_redefine: bool,
    /// The style is hidden from the application UI.
    pub hidden: bool,
    /// Legacy language compatibility properties have been applied.
    pub legacy_languages_set: bool,
    /// The legacy compatibility language represents no-proofing.
    pub copy_language: bool,
    /// Character style used for new e-mail messages.
    pub personal_compose: bool,
    /// Character style used for e-mail replies.
    pub personal_reply: bool,
    /// Character style used for e-mail senders.
    pub personal: bool,
    /// The style is hidden from the simplified styles UI.
    pub semi_hidden: bool,
    /// The style cannot be applied through the application UI.
    pub locked: bool,
    /// The style becomes visible after it is used.
    pub unhide_when_used: bool,
    /// The style is shown in the quick-style gallery.
    pub quick_format: bool,
}

/// Word 2000-and-later metadata appended to an `StdfBase`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StylePost2000 {
    /// Linked style index, or `None` when the encoded index is zero.
    pub linked_style: Option<u16>,
    /// Whether the style stores its pre-revision formatting.
    pub has_original_style: bool,
    /// Revision-save identifier of the last style modification.
    pub revision_id: u32,
    /// Legacy HTML font category.
    pub html_font_category: u8,
    /// UI ordering priority (0 through 99).
    pub priority: u16,
}

/// Previous formatting and attribution stored by a revision-marked style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleRevisionMark {
    /// Date and time at which the style was revision-marked.
    pub timestamp: Option<crate::DateTime>,
    /// Signed index into the document's `SttbfRMark` author table.
    pub author_index: i16,
    /// Resolved revision author when the stylesheet belongs to a complete document.
    pub author: Option<String>,
    /// Previous paragraph formatting for a paragraph style.
    pub paragraph_properties: Option<Vec<u8>>,
    /// Previous character formatting.
    pub character_properties: Vec<u8>,
}

/// One non-empty style definition from the stylesheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleDefinition {
    /// Zero-based index used by `sprmPIstd` and `sprmTIstd`.
    pub index: u16,
    /// Invariant built-in style identifier, or `0x0FFE` for a user style.
    pub invariant_id: u16,
    /// Style kind.
    pub kind: StyleKind,
    /// Parent style index, if this style inherits from another style.
    pub base_style: Option<u16>,
    /// Style automatically applied to the following paragraph.
    pub next_style: u16,
    /// Primary style name.
    pub name: String,
    /// Alternate names following the primary name in the `Xstz`.
    pub aliases: Vec<String>,
    /// Raw UPX payloads, in the kind-specific order prescribed by MS-DOC.
    pub property_sets: Vec<Vec<u8>>,
    /// Optional Word 2000-and-later metadata.
    pub post_2000: Option<StylePost2000>,
    /// Parsed revision attribution and previous formatting, when present.
    pub revision: Option<StyleRevisionMark>,
    /// Style behavior flags.
    pub flags: StyleFlags,
    /// Exact `STD` bytes, excluding the `LPStd` length and outer alignment byte.
    pub raw_std: Vec<u8>,
    /// Alignment byte following an odd-sized `STD`, when present.
    pub outer_padding: Option<u8>,
}

impl StyleDefinition {
    /// The table-property UPX for a table style.
    #[must_use]
    pub fn table_properties(&self) -> Option<&[u8]> {
        (self.kind == StyleKind::Table)
            .then(|| self.property_sets.first().map(Vec::as_slice))
            .flatten()
    }

    /// The current paragraph-property UPX for a paragraph or table style.
    pub fn paragraph_properties(&self) -> Option<&[u8]> {
        match self.kind {
            StyleKind::Paragraph => self.property_sets.first().map(Vec::as_slice),
            StyleKind::Table => self.property_sets.get(1).map(Vec::as_slice),
            StyleKind::Character | StyleKind::Numbering => None,
        }
    }

    /// The current character-property UPX for a paragraph, character, or table style.
    pub fn character_properties(&self) -> Option<&[u8]> {
        match self.kind {
            StyleKind::Paragraph => self.property_sets.get(1).map(Vec::as_slice),
            StyleKind::Character => self.property_sets.first().map(Vec::as_slice),
            StyleKind::Table => self.property_sets.get(2).map(Vec::as_slice),
            StyleKind::Numbering => None,
        }
    }
}

/// Parsed Word stylesheet with null style slots retained.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleSheet {
    /// Non-structural defects repaired during a lenient parse.
    pub(super) tolerance: ToleranceReport,
    pub(super) header: StyleSheetHeader,
    pub(super) styles: Vec<Option<StyleDefinition>>,
    pub(super) stshi_tail: Vec<u8>,
}

impl StyleSheet {
    /// Non-structural defects a lenient parse repaired.
    ///
    /// Always empty after a strict parse.
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance
    }

    /// General stylesheet information.
    #[must_use]
    pub fn header(&self) -> &StyleSheetHeader {
        &self.header
    }

    /// All style slots, including required null fixed-index slots.
    #[must_use]
    pub fn styles(&self) -> &[Option<StyleDefinition>] {
        &self.styles
    }

    /// Resolve one style index to a non-empty definition.
    #[must_use]
    pub fn get(&self, index: u16) -> Option<&StyleDefinition> {
        self.styles.get(usize::from(index))?.as_ref()
    }

    /// Uninterpreted STSHI extension bytes following the 18-byte Stshif.
    #[must_use]
    pub fn stshi_tail(&self) -> &[u8] {
        &self.stshi_tail
    }
}
