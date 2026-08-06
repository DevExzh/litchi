//! Writer-facing document state, input models, and authoring methods.

use crate::CommentDateTime;
use crate::encryption::{EncryptionProfile, validate_writer_password};
use crate::parts::pap::{
    Borders as ParagraphBorders, DropCap, FontAlignment, FrameAnchor, FrameHeight,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalPosition,
    LegacyAutoNumbering, LegacyBorderPosition, LegacyBorderStyle, PhysicalJustification,
    Shading as ParagraphShading, TabStop, TextBoxTightWrap,
};
use crate::parts::{list_names::ListNamesTable, list_templates::ListTemplateTable};
use crate::writer::bookmarks::BookmarkEntry;
use crate::writer::comments::CommentEntry;
use crate::writer::footnotes::FootnoteEntry;
use crate::writer::numbering::{ListFormatOverride, ListStructure, NumberingWriter};
use crate::writer::revisions::{
    DisplayFieldRevision, FormattingRevision, NumberingRevision, TextRevision,
};
use crate::writer::smart_tags::SmartTagEntry;
use crate::{
    AssociatedStringSlot, DocumentAssociatedStrings, GlossaryMetadata, ProofingFeature,
    ProofingStateTable, ProofingTables, SavedByTable, SmartTagRecognizerRange,
};
use std::collections::HashMap;
use zeroize::Zeroizing;

pub(super) const WORD_DOCUMENT_CLSID: [u8; 16] = [
    0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
pub(super) const VBA_PROJECT_STORAGE_NAME: &str = "Macros";

mod codec;
mod error;
mod formatting;
mod semantic;
mod state;
mod story;
mod validation;

pub use error::WriteError;
pub use formatting::{CharacterFormatting, LineSpacing, ParagraphFormatting};
pub use state::Writer;
pub use story::{HeaderFooterParagraph, HeaderKind};

pub(crate) use codec::pack_dttm;
pub(super) use codec::utf16_code_unit_len;
pub(super) use state::WriterEncryption;
pub(super) use story::{
    BookmarkTableData, CommentStoryData, FloatingAnchorKind, HeaderAnchor, HeaderStoryData,
    MainReferenceKind, NoteStoryData, RevisionWriterData, TableCell, TableRow, TextRun,
    WritableParagraph, WritableTable, WriterPicture, WriterShape, writable_paragraph_from_runs,
};
#[cfg(test)]
pub(super) use story::{HEADER_SLOT_EVEN, HEADER_SLOT_FIRST, HEADER_SLOT_ODD};
pub(super) use validation::{HeaderFieldState, checked_text_fc, validate_header_footer_paragraphs};
