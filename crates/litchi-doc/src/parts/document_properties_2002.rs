//! Typed, lossless Word 2002 Document Properties extension.

use super::document_properties_97::DopExtensionError;
use super::document_properties_2000::Dop2000;

const DOP2000_SIZE: usize = 544;
const DOP2002_SIZE: usize = 594;
const EXTENSION_SIZE: usize = DOP2002_SIZE - DOP2000_SIZE;
const FLAGS: usize = 4;
const DEFAULT_TABLE_STYLE: usize = 6;
const FEATURE_SET: usize = 8;
const STYLE_FILTER: usize = 10;
const BOOKLET_PAGES: usize = 12;
const TEXT_CODE_PAGE: usize = 14;
const MAIN_REVISION_CP: usize = 18;
const FOOTNOTE_REVISION_CP: usize = 22;
const HEADER_REVISION_CP: usize = 26;
const COMMENT_REVISION_CP: usize = 30;
const ENDNOTE_REVISION_CP: usize = 34;
const TEXTBOX_REVISION_CP: usize = 38;
const HEADER_TEXTBOX_REVISION_CP: usize = 42;
const ROOT_RSID: usize = 46;
const MAX_CP: u32 = i32::MAX as u32;

/// Line terminator selected for automation-driven text export.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextLineEnding {
    CarriageReturnLineFeed,
    CarriageReturn,
    LineFeed,
    LineFeedCarriageReturn,
    UnicodeOrCodePageDefault,
}

impl TextLineEnding {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::CarriageReturnLineFeed),
            1 => Ok(Self::CarriageReturn),
            2 => Ok(Self::LineFeed),
            3 => Ok(Self::LineFeedCarriageReturn),
            4 => Ok(Self::UnicodeOrCodePageDefault),
            _ => Err(DopExtensionError::new(format!(
                "invalid Dop2002 text line-ending value {raw}"
            ))),
        }
    }

    const fn raw(self) -> u16 {
        match self {
            Self::CarriageReturnLineFeed => 0,
            Self::CarriageReturn => 1,
            Self::LineFeed => 2,
            Self::LineFeedCarriageReturn => 3,
            Self::UnicodeOrCodePageDefault => 4,
        }
    }
}

/// Word 2002 document feature restriction mask.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentFeatureSet(u16);

impl DocumentFeatureSet {
    pub const INTERNET_EXPLORER_4: u16 = 0x0001;
    pub const INTERNET_EXPLORER_5: u16 = 0x0002;
    pub const WORD_95: u16 = 0x0004;
    pub const WORD_97: u16 = 0x0008;
    pub const WORD_HTML: u16 = 0x0010;
    pub const WORD_RTF: u16 = 0x0020;
    pub const EAST_ASIAN_WORD_95: u16 = 0x0040;
    pub const PLAIN_TEXT_EMAIL: u16 = 0x0080;
    pub const INTERNET_EXPLORER_6: u16 = 0x0100;
    pub const WORD_XML: u16 = 0x0200;
    pub const RTF_EMAIL: u16 = 0x0400;
    pub const PRE_WORD_2007: u16 = 0x0800;
    pub const PLAIN_TEXT: u16 = 0x1000;

    #[must_use]
    pub const fn from_raw(raw: u16) -> Self {
        Self(raw)
    }

    #[must_use]
    pub const fn raw(self) -> u16 {
        self.0
    }

    #[must_use]
    pub const fn contains(self, feature: u16) -> bool {
        self.0 & feature != 0
    }
}

/// Suggested style-pane filtering flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StylePaneFormatFilter(u16);

impl StylePaneFormatFilter {
    pub const DEFAULT: Self = Self(0x5024);

    #[must_use]
    pub const fn from_raw(raw: u16) -> Self {
        Self(raw)
    }

    #[must_use]
    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// Code page used by encoded-text export. This codec never performs the export.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TextCodePage(u32);

impl TextCodePage {
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        Self(raw)
    }

    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }
}

/// Earliest character position at which revisions can occur in each story.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct RevisionBoundaries {
    pub main_document: u32,
    pub footnotes: u32,
    pub headers: u32,
    pub comments: u32,
    pub endnotes: u32,
    pub textboxes: u32,
    pub header_textboxes: u32,
}

impl RevisionBoundaries {
    fn validate_cp_domain(self) -> Result<(), DopExtensionError> {
        for (story, cp) in self.entries() {
            if cp > MAX_CP {
                return Err(DopExtensionError::new(format!(
                    "Dop2002 {story} revision CP exceeds signed CP range"
                )));
            }
        }
        Ok(())
    }

    fn entries(self) -> [(&'static str, u32); 7] {
        [
            ("main document", self.main_document),
            ("footnote", self.footnotes),
            ("header", self.headers),
            ("comment", self.comments),
            ("endnote", self.endnotes),
            ("textbox", self.textboxes),
            ("header textbox", self.header_textboxes),
        ]
    }
}

/// Story lengths needed for contextual CP validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct StoryCharacterCounts {
    pub main_document: u32,
    pub footnotes: u32,
    pub headers: u32,
    pub comments: u32,
    pub endnotes: u32,
    pub textboxes: u32,
    pub header_textboxes: u32,
}

impl StoryCharacterCounts {
    fn entries(self) -> [(&'static str, u32); 7] {
        [
            ("main document", self.main_document),
            ("footnote", self.footnotes),
            ("header", self.headers),
            ("comment", self.comments),
            ("endnote", self.endnotes),
            ("textbox", self.textboxes),
            ("header textbox", self.header_textboxes),
        ]
    }
}

/// Typed, lossless Word 2002 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dop2002 {
    raw: [u8; EXTENSION_SIZE],
    pub do_not_embed_system_fonts: bool,
    pub enforce_feature_compatibility: bool,
    pub live_recovered: bool,
    pub embed_smart_tags: bool,
    pub smart_tag_xml_on_html_export: bool,
    pub smart_tag_scan_complete: bool,
    pub book_fold_printing: bool,
    pub reverse_book_fold_printing: bool,
    pub text_line_ending: TextLineEnding,
    pub hide_format_consistency_cues: bool,
    pub show_review_markup_area: bool,
    pub show_comments: bool,
    pub show_insertions_deletions: bool,
    pub show_property_revisions: bool,
    pub default_table_style: u16,
    pub feature_set: DocumentFeatureSet,
    pub style_filter: StylePaneFormatFilter,
    pub booklet_pages: u16,
    pub text_code_page: TextCodePage,
    pub revision_boundaries: RevisionBoundaries,
    pub root_revision_save_id: u32,
}

impl Dop2002 {
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2002_SIZE {
            return Err(DopExtensionError::new("Dop2002 is shorter than 594 bytes"));
        }
        Dop2000::parse(dop)?;
        let extension = &dop[DOP2000_SIZE..DOP2002_SIZE];
        let mut raw = [0u8; EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        let flags = le_u16(extension, FLAGS);
        let reverse_book_fold_printing = flags & (1 << 7) != 0;
        let book_fold_printing = flags & (1 << 6) != 0;
        if reverse_book_fold_printing && !book_fold_printing {
            return Err(DopExtensionError::new(
                "reverse book-fold printing requires book-fold printing",
            ));
        }
        let revision_boundaries = RevisionBoundaries {
            main_document: le_u32(extension, MAIN_REVISION_CP),
            footnotes: le_u32(extension, FOOTNOTE_REVISION_CP),
            headers: le_u32(extension, HEADER_REVISION_CP),
            comments: le_u32(extension, COMMENT_REVISION_CP),
            endnotes: le_u32(extension, ENDNOTE_REVISION_CP),
            textboxes: le_u32(extension, TEXTBOX_REVISION_CP),
            header_textboxes: le_u32(extension, HEADER_TEXTBOX_REVISION_CP),
        };
        revision_boundaries.validate_cp_domain()?;
        Ok(Self {
            raw,
            do_not_embed_system_fonts: flags & 1 != 0,
            enforce_feature_compatibility: flags & 2 != 0,
            live_recovered: flags & 4 != 0,
            embed_smart_tags: flags & 8 != 0,
            smart_tag_xml_on_html_export: flags & 0x10 != 0,
            smart_tag_scan_complete: flags & 0x20 != 0,
            book_fold_printing,
            reverse_book_fold_printing,
            text_line_ending: TextLineEnding::parse((flags >> 8) & 7)?,
            hide_format_consistency_cues: flags & (1 << 11) != 0,
            show_review_markup_area: flags & (1 << 12) != 0,
            show_comments: flags & (1 << 13) != 0,
            show_insertions_deletions: flags & (1 << 14) != 0,
            show_property_revisions: flags & (1 << 15) != 0,
            default_table_style: le_u16(extension, DEFAULT_TABLE_STYLE),
            feature_set: DocumentFeatureSet::from_raw(le_u16(extension, FEATURE_SET)),
            style_filter: StylePaneFormatFilter::from_raw(le_u16(extension, STYLE_FILTER)),
            booklet_pages: le_u16(extension, BOOKLET_PAGES),
            text_code_page: TextCodePage::from_raw(le_u32(extension, TEXT_CODE_PAGE)),
            revision_boundaries,
            root_revision_save_id: le_u32(extension, ROOT_RSID),
        })
    }

    pub fn validate_default_table_style(
        &self,
        style_count: usize,
    ) -> Result<(), DopExtensionError> {
        if usize::from(self.default_table_style) >= style_count {
            Err(DopExtensionError::new(format!(
                "Dop2002 default table style {} exceeds stylesheet",
                self.default_table_style
            )))
        } else {
            Ok(())
        }
    }

    pub fn validate_revision_boundaries(
        &self,
        story_counts: StoryCharacterCounts,
    ) -> Result<(), DopExtensionError> {
        for ((story, cp), (_, count)) in self
            .revision_boundaries
            .entries()
            .into_iter()
            .zip(story_counts.entries())
        {
            if cp > count {
                return Err(DopExtensionError::new(format!(
                    "Dop2002 {story} revision CP {cp} exceeds story length {count}"
                )));
            }
        }
        Ok(())
    }

    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2002_SIZE {
            return Err(DopExtensionError::new(
                "Dop2002 target is shorter than 594 bytes",
            ));
        }
        if self.reverse_book_fold_printing && !self.book_fold_printing {
            return Err(DopExtensionError::new(
                "reverse book-fold printing requires book-fold printing",
            ));
        }
        self.revision_boundaries.validate_cp_domain()?;
        let mut flags = u16::from(self.do_not_embed_system_fonts);
        flags |= u16::from(self.enforce_feature_compatibility) << 1;
        flags |= u16::from(self.live_recovered) << 2;
        flags |= u16::from(self.embed_smart_tags) << 3;
        flags |= u16::from(self.smart_tag_xml_on_html_export) << 4;
        flags |= u16::from(self.smart_tag_scan_complete) << 5;
        flags |= u16::from(self.book_fold_printing) << 6;
        flags |= u16::from(self.reverse_book_fold_printing) << 7;
        flags |= self.text_line_ending.raw() << 8;
        flags |= u16::from(self.hide_format_consistency_cues) << 11;
        flags |= u16::from(self.show_review_markup_area) << 12;
        flags |= u16::from(self.show_comments) << 13;
        flags |= u16::from(self.show_insertions_deletions) << 14;
        flags |= u16::from(self.show_property_revisions) << 15;
        put_u16(&mut self.raw, FLAGS, flags);
        put_u16(&mut self.raw, DEFAULT_TABLE_STYLE, self.default_table_style);
        put_u16(&mut self.raw, FEATURE_SET, self.feature_set.raw());
        put_u16(&mut self.raw, STYLE_FILTER, self.style_filter.raw());
        put_u16(&mut self.raw, BOOKLET_PAGES, self.booklet_pages);
        put_u32(&mut self.raw, TEXT_CODE_PAGE, self.text_code_page.raw());
        put_u32(
            &mut self.raw,
            MAIN_REVISION_CP,
            self.revision_boundaries.main_document,
        );
        put_u32(
            &mut self.raw,
            FOOTNOTE_REVISION_CP,
            self.revision_boundaries.footnotes,
        );
        put_u32(
            &mut self.raw,
            HEADER_REVISION_CP,
            self.revision_boundaries.headers,
        );
        put_u32(
            &mut self.raw,
            COMMENT_REVISION_CP,
            self.revision_boundaries.comments,
        );
        put_u32(
            &mut self.raw,
            ENDNOTE_REVISION_CP,
            self.revision_boundaries.endnotes,
        );
        put_u32(
            &mut self.raw,
            TEXTBOX_REVISION_CP,
            self.revision_boundaries.textboxes,
        );
        put_u32(
            &mut self.raw,
            HEADER_TEXTBOX_REVISION_CP,
            self.revision_boundaries.header_textboxes,
        );
        put_u32(&mut self.raw, ROOT_RSID, self.root_revision_save_id);
        dop[DOP2000_SIZE..DOP2002_SIZE].copy_from_slice(&self.raw);
        Ok(())
    }
}

fn le_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        data[offset..offset + 4]
            .try_into()
            .expect("fixed-width slice"),
    )
}

fn put_u16(data: &mut [u8], offset: usize, value: u16) {
    data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parts::document_properties_97::Dop95;

    fn valid_dop2002() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2002_SIZE];
        dop[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        dop
    }

    #[test]
    fn copts60_mirror_uses_base_offset_eight() {
        let mut dop = valid_dop2002();
        put_u16(&mut dop, 8, 0x1234);
        put_u16(&mut dop, 10, 0x5678);
        put_u32(&mut dop, 84, 0x9abc_1234);
        assert_eq!(
            Dop95::parse(&dop).unwrap().compatibility().copts60(),
            0x1234
        );
    }

    #[test]
    fn parses_complete_typed_extension() {
        let mut dop = valid_dop2002();
        let flags = 1 | 2 | 4 | 8 | (1 << 6) | (1 << 7) | (4 << 8) | (1 << 13);
        put_u16(&mut dop, DOP2000_SIZE + FLAGS, flags);
        put_u16(&mut dop, DOP2000_SIZE + DEFAULT_TABLE_STYLE, 7);
        put_u16(
            &mut dop,
            DOP2000_SIZE + FEATURE_SET,
            DocumentFeatureSet::WORD_XML,
        );
        put_u16(&mut dop, DOP2000_SIZE + STYLE_FILTER, 0x5024);
        put_u16(&mut dop, DOP2000_SIZE + BOOKLET_PAGES, 8);
        put_u32(&mut dop, DOP2000_SIZE + TEXT_CODE_PAGE, 1252);
        put_u32(&mut dop, DOP2000_SIZE + MAIN_REVISION_CP, 42);
        put_u32(&mut dop, DOP2000_SIZE + ROOT_RSID, 0x1234_5678);
        let value = Dop2002::parse(&dop).unwrap();
        assert!(value.reverse_book_fold_printing);
        assert_eq!(
            value.text_line_ending,
            TextLineEnding::UnicodeOrCodePageDefault
        );
        assert!(value.show_comments);
        assert!(value.feature_set.contains(DocumentFeatureSet::WORD_XML));
        assert_eq!(value.text_code_page.raw(), 1252);
        assert_eq!(value.revision_boundaries.main_document, 42);
    }

    #[test]
    fn round_trip_preserves_undefined_prefix() {
        let mut dop = valid_dop2002();
        dop[DOP2000_SIZE..DOP2000_SIZE + 4].copy_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        let parsed = Dop2002::parse(&dop).unwrap();
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(output, dop);
    }

    #[test]
    fn rejects_invalid_line_ending_reverse_fold_and_cp() {
        let mut line = valid_dop2002();
        put_u16(&mut line, DOP2000_SIZE + FLAGS, 5 << 8);
        assert!(Dop2002::parse(&line).is_err());

        let mut reverse = valid_dop2002();
        put_u16(&mut reverse, DOP2000_SIZE + FLAGS, 1 << 7);
        assert!(Dop2002::parse(&reverse).is_err());

        let mut cp = valid_dop2002();
        put_u32(&mut cp, DOP2000_SIZE + MAIN_REVISION_CP, 0x8000_0000);
        assert!(Dop2002::parse(&cp).is_err());
    }

    #[test]
    fn validates_contextual_style_and_story_bounds() {
        let mut dop = valid_dop2002();
        put_u16(&mut dop, DOP2000_SIZE + DEFAULT_TABLE_STYLE, 3);
        put_u32(&mut dop, DOP2000_SIZE + MAIN_REVISION_CP, 5);
        let value = Dop2002::parse(&dop).unwrap();
        assert!(value.validate_default_table_style(3).is_err());
        assert!(value.validate_default_table_style(4).is_ok());
        assert!(
            value
                .validate_revision_boundaries(StoryCharacterCounts {
                    main_document: 4,
                    ..StoryCharacterCounts::default()
                })
                .is_err()
        );
    }
}
