use super::{
    BookmarkEntry, CommentEntry, DocumentAssociatedStrings, FootnoteEntry, FormattingRevision,
    GlossaryMetadata, HeaderAnchor, HeaderFooterParagraph, NumberingWriter, ProofingTables,
    SavedByTable, SmartTagEntry, SmartTagRecognizerRange, WritableParagraph, WritableTable,
    WriterPicture, WriterShape,
};
use crate::encryption::EncryptionProfile;
use std::collections::HashMap;
use zeroize::Zeroizing;

/// DOC file writer
///
/// Provides methods to create and modify DOC files.
pub struct Writer {
    /// Paragraphs in the document
    pub(in crate::writer::core) paragraphs: Vec<WritableParagraph>,
    /// Tables in the document
    pub(in crate::writer::core) tables: Vec<WritableTable>,
    /// Document properties
    pub(in crate::writer::core) properties: HashMap<String, String>,
    /// Header/footer paragraphs (`None` means the story is not set).
    /// Indices map to plcfHdd entries (following Apache POI HeaderStories indexing):
    /// 0..5: footnote/endnote separators (unused here)
    /// 6: even header, 7: odd header, 10: first header
    /// 8: even footer, 9: odd footer, 11: first footer
    pub(in crate::writer::core) header_even: Option<Vec<HeaderFooterParagraph>>,
    pub(in crate::writer::core) header_odd: Option<Vec<HeaderFooterParagraph>>,
    pub(in crate::writer::core) header_first: Option<Vec<HeaderFooterParagraph>>,
    pub(in crate::writer::core) footer_even: Option<Vec<HeaderFooterParagraph>>,
    pub(in crate::writer::core) footer_odd: Option<Vec<HeaderFooterParagraph>>,
    pub(in crate::writer::core) footer_first: Option<Vec<HeaderFooterParagraph>>,
    /// Footnote entries
    pub(in crate::writer::core) footnotes: Vec<FootnoteEntry>,
    /// Endnote entries
    pub(in crate::writer::core) endnotes: Vec<FootnoteEntry>,
    /// Comments
    pub(in crate::writer::core) comments: Vec<CommentEntry>,
    /// Standard bookmarks
    pub(in crate::writer::core) bookmarks: Vec<BookmarkEntry>,
    /// Embedded smart-tag bookmarks and property bags.
    pub(in crate::writer::core) smart_tags: Vec<SmartTagEntry>,
    /// Smart-tag recognizer processing-state ranges.
    pub(in crate::writer::core) smart_tag_recognizer_ranges: Vec<SmartTagRecognizerRange>,
    /// Optional spelling and grammar proofing-state PLCFs.
    pub(in crate::writer::core) proofing_tables: ProofingTables,
    /// Mandatory fixed associated-document string table.
    pub(in crate::writer::core) associated_strings: DocumentAssociatedStrings,
    /// Optional Word 97/2000 save-history table.
    pub(in crate::writer::core) saved_by_table: Option<SavedByTable>,
    /// Optional glossary-only AutoText metadata over the main story.
    pub(in crate::writer::core) glossary_metadata: Option<GlossaryMetadata>,
    /// Optional distinct AutoText-only document attached to this template.
    pub(in crate::writer::core) attached_glossary: Option<Box<Writer>>,
    /// Property revision metadata for the writer's single document section
    pub(in crate::writer::core) section_formatting_revision: Option<FormattingRevision>,
    /// Explicit column geometry for the writer's single document section.
    pub(in crate::writer::core) section_columns: Option<crate::section::columns::Layout>,
    /// Whether section columns are populated from right to left.
    pub(in crate::writer::core) section_right_to_left: bool,
    /// Section-wide glyph and line flow.
    pub(in crate::writer::core) section_text_flow: crate::TextFlow,
    /// Explicit page-border edges and placement for the single section.
    pub(in crate::writer::core) section_page_borders: Option<crate::section::borders::Borders>,
    /// Numbering writer for list tables
    pub(in crate::writer::core) numbering: NumberingWriter,
    /// User-defined styles appended after the fifteen fixed style slots
    pub(in crate::writer::core) styles: Vec<crate::writer::stylesheet::StyleDefinition>,
    /// Inline pictures embedded via [`Writer::insert_picture`]
    pub(in crate::writer::core) pictures: Vec<WriterPicture>,
    /// Primitive drawing shapes embedded via [`Writer::insert_floating_shape`]
    pub(in crate::writer::core) shapes: Vec<WriterShape>,
    /// Text boxes anchored in the header story, in insertion order.
    pub(in crate::writer::core) header_shapes: Vec<WriterShape>,
    /// Pictures anchored in the header story, in insertion order.
    pub(in crate::writer::core) header_pictures: Vec<WriterPicture>,
    /// Anchor paragraphs appended to header paragraph lists, in insertion
    /// order (one per header floating item).
    pub(in crate::writer::core) header_anchors: Vec<HeaderAnchor>,
    /// Next shape id to allocate (shared by pictures and drawing shapes).
    pub(in crate::writer::core) next_shape_id: u32,
    /// Password-to-open settings. The password is wiped when replaced, cleared, or dropped.
    pub(in crate::writer::core) encryption: Option<WriterEncryption>,
    /// Complete inert MS-OVBA project written under the MS-DOC `Macros` storage.
    pub(in crate::writer::core) vba_project: Option<litchi_vba::Payload>,
}

pub(in crate::writer::core) struct WriterEncryption {
    pub(in crate::writer::core) profile: EncryptionProfile,
    pub(in crate::writer::core) password: Zeroizing<String>,
}
