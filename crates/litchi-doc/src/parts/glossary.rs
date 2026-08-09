//! Passive `AutoText` and formatted `AutoCorrect` metadata for legacy DOC files.

use super::super::package::{Error as PackageError, Result};
use super::super::paragraph::{Paragraph, Run};
use super::chp_bin_table::ChpBinTable;
use super::fib::FileInformationBlock;
use super::fields::{Field, FieldStory, FieldText, FieldsTable, HyperlinkField, NonPlcfFields};
use super::pap_bin_table::PapBinTable;
use super::paragraph_extractor::ParagraphExtractor;
use super::piece_table::PieceTable;
use super::revisions::RevisionAuthorTable;
use super::styles::StyleSheet;
use super::text::TextExtractor;
use std::collections::HashSet;
use std::sync::Arc;

const STTBF_GLSY_FIB_INDEX: usize = 9;
const PLCF_GLSY_FIB_INDEX: usize = 10;
const STTB_GLSY_STYLE_FIB_INDEX: usize = 83;
const MAX_ITEM_NAME_UNITS: usize = 32;
const MAX_STYLE_USE_COUNT: u8 = 0x32;
const MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;
const FIB_PAGE_BYTES: usize = 512;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], context: &str) -> Result<String> {
    char::decode_utf16(
        data.chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
    )
    .collect::<std::result::Result<String, _>>()
    .map_err(|_| corrupted(format!("{context} contains invalid UTF-16")))
}

/// The inert classification recorded by `LEGOXTR_V11.flego`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GlossaryItemKind {
    NamedAutoText,
    FormattedAutoCorrect,
}

impl GlossaryItemKind {
    fn from_byte(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(Self::NamedAutoText),
            0x0A => Ok(Self::FormattedAutoCorrect),
            _ => Err(corrupted(format!(
                "SttbfGlsy contains invalid flego value 0x{value:02X}"
            ))),
        }
    }

    fn as_byte(self) -> u8 {
        match self {
            Self::NamedAutoText => 0x00,
            Self::FormattedAutoCorrect => 0x0A,
        }
    }
}

/// One `AutoText` or formatted `AutoCorrect` item and its main-story CP range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryItem {
    name: String,
    kind: GlossaryItemKind,
    style_index: Option<u16>,
    start_cp: u32,
    end_cp: u32,
}

impl GlossaryItem {
    pub fn try_new(
        name: impl Into<String>,
        kind: GlossaryItemKind,
        style_index: Option<u16>,
        start_cp: u32,
        end_cp: u32,
    ) -> Result<Self> {
        let item = Self {
            name: name.into(),
            kind,
            style_index,
            start_cp,
            end_cp,
        };
        validate_item_shape(&item, 0)?;
        Ok(item)
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub fn kind(&self) -> GlossaryItemKind {
        self.kind
    }
    #[must_use]
    pub fn style_index(&self) -> Option<u16> {
        self.style_index
    }
    #[must_use]
    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    #[must_use]
    pub fn end_cp(&self) -> u32 {
        self.end_cp
    }
}

/// One style-name slot parallel to the style indices in `SttbfGlsy`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryStyle {
    name: String,
    use_count: u8,
}

impl GlossaryStyle {
    pub fn try_new(name: impl Into<String>, use_count: u8) -> Result<Self> {
        let style = Self {
            name: name.into(),
            use_count,
        };
        validate_style(&style, 0)?;
        Ok(style)
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub fn use_count(&self) -> u8 {
        self.use_count
    }
}

/// Serialized forms of the three glossary tables, ready for FIB slots 9, 10, and 83.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryTables {
    sttbf_glsy: Vec<u8>,
    plcf_glsy: Vec<u8>,
    sttb_glsy_style: Vec<u8>,
}

impl GlossaryTables {
    #[must_use]
    pub fn item_table(&self) -> &[u8] {
        &self.sttbf_glsy
    }
    #[must_use]
    pub fn position_table(&self) -> &[u8] {
        &self.plcf_glsy
    }
    #[must_use]
    pub fn style_table(&self) -> &[u8] {
        &self.sttb_glsy_style
    }
    #[must_use]
    pub fn into_parts(self) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        (self.sttbf_glsy, self.plcf_glsy, self.sttb_glsy_style)
    }
}

/// Cross-validated metadata from `SttbfGlsy`, `PlcfGlsy`, and `SttbGlsyStyle`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryMetadata {
    items: Vec<GlossaryItem>,
    styles: Vec<GlossaryStyle>,
    terminal_cp: u32,
    ignored_cp: u32,
    main_text_length: u32,
}

/// An AutoText-only document attached to a template through `FibBase.pnNext`.
///
/// Fields, links, embedded objects, and macros remain passive. Stored text,
/// formatting, pictures, and drawing metadata can be inspected without
/// resolving or activating content.
pub struct AttachedGlossary {
    fib: FileInformationBlock,
    metadata: GlossaryMetadata,
    text_extractor: TextExtractor,
    fields_table: FieldsTable,
    chp_bin_table: Option<ChpBinTable>,
    pap_bin_table: Option<PapBinTable>,
    stylesheet: Option<StyleSheet>,
    revision_authors: RevisionAuthorTable,
    images: Vec<super::super::image::Image>,
    shapes: Vec<crate::shape::Shape>,
    shape_anchors: Vec<super::spa::ShapeAnchor>,
    header_shape_anchors: Vec<super::spa::ShapeAnchor>,
    textbox_entries: Vec<super::textbox::TextBoxEntry>,
    header_textbox_entries: Vec<super::textbox::TextBoxEntry>,
}

impl AttachedGlossary {
    pub(crate) fn parse(
        main_fib: &FileInformationBlock,
        word_document: &[u8],
        table_stream: &[u8],
        data_stream: Option<&[u8]>,
    ) -> Result<Option<Self>> {
        let page = main_fib.next_fib_page();
        if page == 0 {
            return Ok(None);
        }
        if main_fib.is_glossary_document() {
            return Err(corrupted(
                "glossary-only FIB has a forbidden nonzero pnNext",
            ));
        }
        if !main_fib.is_template() {
            return Err(corrupted(
                "non-template document has a forbidden nonzero pnNext",
            ));
        }
        let offset = usize::from(page)
            .checked_mul(FIB_PAGE_BYTES)
            .ok_or_else(|| corrupted("attached glossary FIB offset overflows"))?;
        let main_fib_size = main_fib
            .minimum_serialized_size()
            .ok_or_else(|| corrupted("template FIB pointer array is truncated"))?;
        if offset < main_fib_size {
            return Err(corrupted("attached glossary FIB overlaps the template FIB"));
        }
        let main_cb_mac_u32 = main_fib
            .word_document_size()
            .ok_or_else(|| corrupted("template FIB does not contain cbMac"))?;
        let main_cb_mac = usize::try_from(main_cb_mac_u32)
            .map_err(|_| corrupted("template cbMac is too large"))?;
        if main_cb_mac > word_document.len() {
            return Err(corrupted(
                "template cbMac extends beyond the WordDocument stream",
            ));
        }
        if offset >= main_cb_mac {
            return Err(corrupted(
                "attached glossary FIB starts at or beyond template cbMac",
            ));
        }
        let fib = FileInformationBlock::parse_at(word_document, offset)?;
        if !fib.is_glossary_document() {
            return Err(corrupted("pnNext does not address a glossary-only FIB"));
        }
        if fib.next_fib_page() != 0 {
            return Err(corrupted(
                "attached glossary FIB has a forbidden nonzero pnNext",
            ));
        }
        if fib.which_table_stream() != main_fib.which_table_stream() {
            return Err(corrupted(
                "template and attached glossary select different table streams",
            ));
        }
        for (index, name) in [(12, "PlcfBteChpx"), (13, "PlcfBtePapx")] {
            let main = main_fib
                .get_table_pointer(index)
                .ok_or_else(|| corrupted(format!("template FIB does not contain {name}")))?;
            let attached = fib.get_table_pointer(index).ok_or_else(|| {
                corrupted(format!("attached glossary FIB does not contain {name}"))
            })?;
            if main != attached {
                return Err(corrupted(format!(
                    "template and attached glossary do not share {name}"
                )));
            }
        }
        let attached_cb_mac = fib
            .word_document_size()
            .ok_or_else(|| corrupted("attached glossary FIB does not contain cbMac"))?;
        if attached_cb_mac != main_cb_mac_u32 {
            return Err(corrupted(
                "template and attached glossary do not share cbMac",
            ));
        }
        let metadata = GlossaryMetadata::parse(&fib, table_stream)?
            .ok_or_else(|| corrupted("attached glossary metadata is absent"))?;
        let text_extractor = TextExtractor::new(&fib, word_document, table_stream)?;
        metadata.validate_text_boundaries(&text_extractor)?;
        let fields_table = FieldsTable::parse(&fib, table_stream)?;
        let revision_authors = RevisionAuthorTable::parse(&fib, table_stream)?;
        let mut stylesheet = (fib.version() >= 0x00C1)
            .then(|| StyleSheet::parse(&fib, table_stream))
            .transpose()?;
        if let Some(stylesheet) = &mut stylesheet {
            stylesheet.resolve_revision_authors(&revision_authors)?;
        }
        let piece_table = table_slice(&fib, table_stream, 33)?.and_then(PieceTable::parse);
        let chpx_data = table_slice(&fib, table_stream, 12)?;
        let chp_bin_table = piece_table.as_ref().and_then(|piece_table| {
            chpx_data.and_then(|data| ChpBinTable::parse(data, word_document, piece_table))
        });
        let papx_data = table_slice(&fib, table_stream, 13)?;
        let pap_bin_table =
            if let (Some(piece_table), Some(data)) = (piece_table.as_ref(), papx_data) {
                PapBinTable::parse(
                    data,
                    word_document,
                    data_stream,
                    piece_table,
                    stylesheet.as_ref(),
                )?
            } else {
                None
            };
        let images = collect_images(&text_extractor, chp_bin_table.as_ref(), data_stream);
        let _ = table_slice(&fib, table_stream, crate::shape::FIB_INDEX_DGG_INFO)?;
        let shapes = crate::shape::extract_dgg_shapes(&fib, table_stream)
            .map_err(|error| corrupted(format!("invalid attached glossary drawing: {error}")))?;
        let shape_anchors =
            parse_shape_anchors(&fib, table_stream, super::spa::FIB_INDEX_PLC_SPA_MOM)?;
        let header_shape_anchors =
            parse_shape_anchors(&fib, table_stream, super::spa::FIB_INDEX_PLC_SPA_HDR)?;
        let textbox_entries = parse_textbox_entries(
            &fib,
            table_stream,
            super::textbox::FIB_INDEX_PLCF_TXBX_TXT,
            fib.get_textbox_range(),
        )?;
        let header_textbox_entries = parse_textbox_entries(
            &fib,
            table_stream,
            super::textbox::FIB_INDEX_PLCF_HDR_TXBX_TXT,
            fib.get_header_textbox_range(),
        )?;
        Ok(Some(Self {
            fib,
            metadata,
            text_extractor,
            fields_table,
            chp_bin_table,
            pap_bin_table,
            stylesheet,
            revision_authors,
            images,
            shapes,
            shape_anchors,
            header_shape_anchors,
            textbox_entries,
            header_textbox_entries,
        }))
    }

    /// The secondary AutoText-only FIB.
    #[must_use]
    pub fn fib(&self) -> &FileInformationBlock {
        &self.fib
    }

    /// Cross-validated item, style, and character-position metadata.
    #[must_use]
    pub fn metadata(&self) -> &GlossaryMetadata {
        &self.metadata
    }

    /// The complete stored AutoText-only main story.
    #[must_use]
    pub fn text(&self) -> &str {
        self.text_extractor.text()
    }

    /// Get one entry's stored content without its structural final character.
    #[must_use]
    pub fn item_text(&self, index: usize) -> Option<&str> {
        let item = self.metadata.items().get(index)?;
        Some(
            self.text_extractor
                .text_at_range(item.start_cp(), item.end_cp().saturating_sub(1)),
        )
    }

    /// Return the strictly parsed field-character tables for all attached stories.
    ///
    /// Field instructions, cached results, external targets, controls, and
    /// macro names remain inert and are never resolved, activated, refreshed,
    /// or executed.
    #[must_use]
    pub fn fields_table(&self) -> &FieldsTable {
        &self.fields_table
    }

    /// Return stored instruction and cached-result text for every attached story.
    ///
    /// Ranges are resolved against the secondary FIB's story character counts.
    /// This reads existing text only and performs no field evaluation or
    /// external action.
    pub fn fields(&self) -> Result<Vec<FieldText>> {
        self.fields_table
            .field_texts(|story, start, end| self.field_story_text(story, start, end))
    }

    /// Return stored instruction and cached-result text for one attached field.
    pub fn field_text(&self, field: &Field) -> Result<FieldText> {
        FieldText::from_field(field, |start, end| {
            self.field_story_text(field.story, start, end)
        })
    }

    /// Return typed, inert `HYPERLINK` fields from all attached stories.
    ///
    /// Targets and bookmarks are returned exactly as stored. They are never
    /// opened, followed, resolved, or refreshed.
    pub fn hyperlink_fields(&self) -> Result<Vec<HyperlinkField>> {
        Ok(self
            .fields()?
            .iter()
            .filter_map(FieldText::hyperlink_field)
            .collect())
    }

    /// Reconstruct the five field kinds excluded from `Plcfld` by MS-DOC.
    ///
    /// This scans each attached story once and returns typed, inert `TC`, `TA`,
    /// `XE`, `RD`, and `PRIVATE` metadata. Unbalanced or unrecognized field
    /// characters are ignored. Referenced documents are never opened and
    /// conversion payloads are never interpreted.
    #[must_use]
    pub fn non_plcf_fields(&self) -> NonPlcfFields {
        NonPlcfFields::from_story_texts(FieldStory::ALL.into_iter().filter_map(|story| {
            let (start, end) = story.range(&self.fib)?;
            Some((story, self.text_extractor.text_at_range(start, end)))
        }))
    }

    /// Extract formatted paragraphs from every stored glossary subdocument.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let text = Arc::new(self.text_extractor.text().to_string());
        let mut output = Vec::new();
        for (_, start_cp, end_cp) in self.fib.get_all_subdoc_ranges() {
            if start_cp >= end_cp {
                continue;
            }
            let extractor = ParagraphExtractor::new_with_range_and_stylesheet(
                Arc::clone(&text),
                self.pap_bin_table.as_ref(),
                self.chp_bin_table.as_ref(),
                (start_cp, end_cp),
                self.stylesheet.as_ref(),
            )?;
            for (_, properties, runs) in extractor.extract_paragraphs()? {
                let mut run_objects = Vec::with_capacity(runs.len());
                for (text, properties) in runs {
                    let image_offset = properties.pic_offset.filter(|offset| {
                        self.images
                            .binary_search_by_key(offset, super::super::image::Image::pic_offset)
                            .is_ok()
                    });
                    let mut run = if let Some(offset) = image_offset {
                        Run::with_image(text, properties, super::super::image::Image::new(offset))
                    } else {
                        Run::new(text, properties)
                    };
                    run.resolve_revisions(&self.revision_authors)?;
                    run_objects.push(run);
                }
                let mut paragraph = Paragraph::new(String::new());
                paragraph.set_runs(run_objects);
                paragraph.set_properties(properties);
                paragraph.resolve_revision(&self.revision_authors)?;
                output.push(paragraph);
            }
        }
        Ok(output)
    }

    /// Pictures referenced by formatted runs in the attached story.
    ///
    /// Pass an entry to [`super::super::document::Document::image_data`] to
    /// retrieve its payload from the template's shared Data stream.
    #[must_use]
    pub fn images(&self) -> &[super::super::image::Image] {
        &self.images
    }

    /// Floating `OfficeArt` shapes stored by the secondary FIB.
    #[must_use]
    pub fn shapes(&self) -> &[crate::shape::Shape] {
        &self.shapes
    }

    /// Floating-shape anchors in the attached main story.
    #[must_use]
    pub fn shape_positions(&self) -> &[super::spa::ShapeAnchor] {
        &self.shape_anchors
    }

    /// Floating-shape anchors in attached header/footer stories.
    #[must_use]
    pub fn header_shape_positions(&self) -> &[super::spa::ShapeAnchor] {
        &self.header_shape_anchors
    }

    /// Text boxes stored in the attached main story.
    #[must_use]
    pub fn text_boxes(&self) -> Vec<super::textbox::TextBox> {
        resolve_text_boxes(
            &self.text_extractor,
            &self.textbox_entries,
            self.fib.get_textbox_range(),
        )
    }

    /// Text boxes stored in attached header/footer stories.
    #[must_use]
    pub fn header_text_boxes(&self) -> Vec<super::textbox::TextBox> {
        resolve_text_boxes(
            &self.text_extractor,
            &self.header_textbox_entries,
            self.fib.get_header_textbox_range(),
        )
    }

    fn field_story_text(&self, story: FieldStory, start: u32, end: u32) -> Result<String> {
        if start > end {
            return Err(corrupted("field text range has its start after its end"));
        }
        let (story_start, story_end) = story
            .range(&self.fib)
            .ok_or_else(|| corrupted("field table refers to an absent attached story"))?;
        let start = story_start
            .checked_add(start)
            .ok_or_else(|| corrupted("field text range start overflows"))?;
        let end = story_start
            .checked_add(end)
            .ok_or_else(|| corrupted("field text range end overflows"))?;
        if end > story_end {
            return Err(corrupted(
                "field text range exceeds its attached document story",
            ));
        }
        Ok(self.text_extractor.text_at_range(start, end).to_string())
    }
}

fn table_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start = usize::try_from(offset).map_err(|_| {
        corrupted(format!(
            "attached glossary table {index} offset is too large"
        ))
    })?;
    let length = usize::try_from(length).map_err(|_| {
        corrupted(format!(
            "attached glossary table {index} length is too large"
        ))
    })?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("attached glossary table {index} range overflows")))?;
    table_stream.get(start..end).map(Some).ok_or_else(|| {
        corrupted(format!(
            "attached glossary table {index} extends beyond the table stream"
        ))
    })
}

fn collect_images(
    text: &TextExtractor,
    chp: Option<&ChpBinTable>,
    data_stream: Option<&[u8]>,
) -> Vec<super::super::image::Image> {
    let (Some(chp), Some(data_stream)) = (chp, data_stream) else {
        return Vec::new();
    };
    let mut offsets = HashSet::new();
    let mut images = Vec::new();
    for run in chp.runs() {
        let Some(offset) = run.properties.pic_offset else {
            continue;
        };
        let run_text = text.text_at_range(run.start_cp, run.end_cp);
        let run_text = run_text
            .strip_suffix('\r')
            .or_else(|| run_text.strip_suffix('\u{7}'))
            .unwrap_or(run_text);
        if offsets.insert(offset)
            && let Ok(Some(image)) =
                super::super::image::extract_image(data_stream, run_text, &run.properties)
        {
            images.push(image);
        }
    }
    images.sort_unstable_by_key(super::super::image::Image::pic_offset);
    images
}

fn parse_shape_anchors(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    index: usize,
) -> Result<Vec<super::spa::ShapeAnchor>> {
    let Some(data) = table_slice(fib, table_stream, index)? else {
        return Ok(Vec::new());
    };
    super::spa::parse_plcf_spa(data)
}

fn parse_textbox_entries(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    index: usize,
    story_range: Option<(u32, u32)>,
) -> Result<Vec<super::textbox::TextBoxEntry>> {
    let Some(data) = table_slice(fib, table_stream, index)? else {
        return Ok(Vec::new());
    };
    let entries = super::textbox::parse_plcf_txbx_txt(data)?;
    let story_length = story_range
        .and_then(|(start, end)| end.checked_sub(start))
        .ok_or_else(|| corrupted(format!("attached glossary table {index} has no text story")))?;
    if entries
        .iter()
        .any(|entry| entry.start_cp > entry.end_cp || entry.end_cp > story_length)
    {
        return Err(corrupted(format!(
            "attached glossary table {index} has an invalid text range"
        )));
    }
    Ok(entries)
}

fn resolve_text_boxes(
    text: &TextExtractor,
    entries: &[super::textbox::TextBoxEntry],
    range: Option<(u32, u32)>,
) -> Vec<super::textbox::TextBox> {
    let Some((story_start, story_end)) = range else {
        return Vec::new();
    };
    entries
        .iter()
        .filter_map(|entry| {
            let start = story_start.checked_add(entry.start_cp)?;
            let end = story_start.checked_add(entry.end_cp)?;
            if start > end || end > story_end {
                return None;
            }
            let raw = text.text_at_range(start, end);
            Some(super::textbox::TextBox {
                shape_id: entry.shape_id,
                text: raw.strip_suffix('\r').unwrap_or(raw).to_string(),
                header_kind: None,
            })
        })
        .collect()
}

impl GlossaryMetadata {
    pub fn try_new(
        items: Vec<GlossaryItem>,
        styles: Vec<GlossaryStyle>,
        terminal_cp: u32,
        ignored_cp: u32,
        main_text_length: u32,
    ) -> Result<Self> {
        let value = Self {
            items,
            styles,
            terminal_cp,
            ignored_cp,
            main_text_length,
        };
        validate_metadata(&value)?;
        Ok(value)
    }

    /// Parse the three FIB-addressed tables only for an AutoText-only document.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let pointers = [
            (STTBF_GLSY_FIB_INDEX, "SttbfGlsy"),
            (PLCF_GLSY_FIB_INDEX, "PlcfGlsy"),
            (STTB_GLSY_STYLE_FIB_INDEX, "SttbGlsyStyle"),
        ];
        if !fib.is_glossary_document() {
            if pointers.iter().any(|(index, _)| {
                fib.get_table_pointer(*index)
                    .is_some_and(|(_, length)| length != 0)
            }) {
                return Err(corrupted(
                    "glossary table data is present while FibBase.fGlsy is clear",
                ));
            }
            return Ok(None);
        }

        let item = table_range(fib, table_stream, pointers[0])?;
        let positions = table_range(fib, table_stream, pointers[1])?;
        let styles = table_range(fib, table_stream, pointers[2])?;
        Self::parse_table_bytes(item, positions, styles, fib.get_main_doc_range().1).map(Some)
    }

    /// Parse complete raw payloads for the three parallel glossary tables.
    pub fn parse_table_bytes(
        sttbf_glsy: &[u8],
        plcf_glsy: &[u8],
        sttb_glsy_style: &[u8],
        main_text_length: u32,
    ) -> Result<Self> {
        let styles = parse_styles(sttb_glsy_style)?;
        let raw_items = parse_items(sttbf_glsy)?;
        let expected_cp_count = raw_items
            .len()
            .checked_add(2)
            .ok_or_else(|| corrupted("PlcfGlsy CP count overflows"))?;
        let expected_size = expected_cp_count
            .checked_mul(4)
            .ok_or_else(|| corrupted("PlcfGlsy byte count overflows"))?;
        if plcf_glsy.len() != expected_size {
            return Err(corrupted(format!(
                "PlcfGlsy has {} bytes; expected {expected_size}",
                plcf_glsy.len()
            )));
        }
        if plcf_glsy.len() > MAX_TABLE_BYTES {
            return Err(corrupted("PlcfGlsy exceeds the table size cap"));
        }
        let mut cps = Vec::with_capacity(expected_cp_count);
        for index in 0..expected_cp_count {
            cps.push(read_u32(plcf_glsy, index * 4, "PlcfGlsy CP")?);
        }
        let items = raw_items
            .into_iter()
            .enumerate()
            .map(|(index, raw)| GlossaryItem {
                name: raw.name,
                kind: raw.kind,
                style_index: raw.style_index,
                start_cp: cps[index],
                end_cp: cps[index + 1],
            })
            .collect();
        Self::try_new(
            items,
            styles,
            cps[expected_cp_count - 2],
            cps[expected_cp_count - 1],
            main_text_length,
        )
    }

    #[must_use]
    pub fn items(&self) -> &[GlossaryItem] {
        &self.items
    }
    #[must_use]
    pub fn styles(&self) -> &[GlossaryStyle] {
        &self.styles
    }
    #[must_use]
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    #[must_use]
    pub fn ignored_cp(&self) -> u32 {
        self.ignored_cp
    }
    #[must_use]
    pub fn main_text_length(&self) -> u32 {
        self.main_text_length
    }

    #[must_use]
    pub fn style_for_item(&self, index: usize) -> Option<&GlossaryStyle> {
        self.items
            .get(index)?
            .style_index
            .and_then(|style| self.styles.get(usize::from(style)))
    }

    pub(crate) fn validate_text_boundaries(&self, text: &TextExtractor) -> Result<()> {
        for (index, item) in self.items.iter().enumerate() {
            if !text.is_cp_boundary(item.start_cp) || !text.is_cp_boundary(item.end_cp) {
                return Err(corrupted(format!(
                    "glossary item {index} range splits a UTF-16 surrogate pair"
                )));
            }
        }
        for (name, cp) in [
            ("terminal", self.terminal_cp),
            ("ignored", self.ignored_cp),
            ("ccpText", self.main_text_length),
        ] {
            if !text.is_cp_boundary(cp) {
                return Err(corrupted(format!(
                    "glossary {name} CP splits a UTF-16 surrogate pair"
                )));
            }
        }
        Ok(())
    }

    pub(crate) fn validate_utf16_bytes(&self, bytes: &[u8]) -> Result<()> {
        let required_bytes = usize::try_from(self.main_text_length)
            .ok()
            .and_then(|units| units.checked_mul(2))
            .ok_or_else(|| corrupted("glossary ccpText byte count overflows"))?;
        let bytes = bytes
            .get(..required_bytes)
            .ok_or_else(|| corrupted("generated glossary text is truncated"))?;
        let is_boundary = |cp: u32| {
            let Ok(index) = usize::try_from(cp) else {
                return false;
            };
            if index > bytes.len() / 2 {
                return false;
            }
            if index == 0 || index == bytes.len() / 2 {
                return true;
            }
            let previous_offset = (index - 1) * 2;
            let current_offset = index * 2;
            let previous = u16::from_le_bytes([bytes[previous_offset], bytes[previous_offset + 1]]);
            let current = u16::from_le_bytes([bytes[current_offset], bytes[current_offset + 1]]);
            !(matches!(previous, 0xD800..=0xDBFF) && matches!(current, 0xDC00..=0xDFFF))
        };
        for (index, item) in self.items.iter().enumerate() {
            if !is_boundary(item.start_cp) || !is_boundary(item.end_cp) {
                return Err(corrupted(format!(
                    "glossary item {index} range splits a UTF-16 surrogate pair"
                )));
            }
        }
        for (name, cp) in [
            ("terminal", self.terminal_cp),
            ("ignored", self.ignored_cp),
            ("ccpText", self.main_text_length),
        ] {
            if !is_boundary(cp) {
                return Err(corrupted(format!(
                    "glossary {name} CP splits a UTF-16 surrogate pair"
                )));
            }
        }
        Ok(())
    }

    /// Serialize the three complete table payloads deterministically.
    pub fn to_table_bytes(&self) -> Result<GlossaryTables> {
        validate_metadata(self)?;
        let item_size = sttb_size(
            self.items.iter().map(|item| item.name.as_str()),
            4,
            "SttbfGlsy",
        )?;
        let mut sttbf_glsy = Vec::with_capacity(item_size);
        write_sttb_header(&mut sttbf_glsy, self.items.len(), 4, "SttbfGlsy")?;
        for item in &self.items {
            write_string(&mut sttbf_glsy, &item.name, "SttbfGlsy")?;
            sttbf_glsy.push(item.kind.as_byte());
            sttbf_glsy.push(0);
            sttbf_glsy.extend_from_slice(&item.style_index.unwrap_or(u16::MAX).to_le_bytes());
        }

        let mut plcf_glsy = Vec::with_capacity((self.items.len() + 2) * 4);
        for item in &self.items {
            plcf_glsy.extend_from_slice(&item.start_cp.to_le_bytes());
        }
        plcf_glsy.extend_from_slice(&self.terminal_cp.to_le_bytes());
        plcf_glsy.extend_from_slice(&self.ignored_cp.to_le_bytes());

        let style_size = sttb_size(
            self.styles.iter().map(|style| style.name.as_str()),
            1,
            "SttbGlsyStyle",
        )?;
        let mut sttb_glsy_style = Vec::with_capacity(style_size);
        write_sttb_header(&mut sttb_glsy_style, self.styles.len(), 1, "SttbGlsyStyle")?;
        for style in &self.styles {
            write_string(&mut sttb_glsy_style, &style.name, "SttbGlsyStyle")?;
            sttb_glsy_style.push(style.use_count);
        }
        Ok(GlossaryTables {
            sttbf_glsy,
            plcf_glsy,
            sttb_glsy_style,
        })
    }
}

struct RawItem {
    name: String,
    kind: GlossaryItemKind,
    style_index: Option<u16>,
}

fn table_range<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    (index, name): (usize, &str),
) -> Result<&'a [u8]> {
    let (offset, length) = fib
        .get_table_pointer(index)
        .ok_or_else(|| corrupted(format!("FIB does not contain the {name} pointer")))?;
    if length == 0 {
        return Err(corrupted(format!("glossary document has no {name}")));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    if length > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn parse_items(data: &[u8]) -> Result<Vec<RawItem>> {
    let count = parse_sttb_header(data, 4, "SttbfGlsy")?;
    let mut items = Vec::with_capacity(count);
    let mut offset = 6usize;
    for index in 0..count {
        let (name, next) = parse_string(data, offset, MAX_ITEM_NAME_UNITS, "SttbfGlsy", index)?;
        offset = next;
        let extra = data
            .get(offset..offset + 4)
            .ok_or_else(|| corrupted(format!("SttbfGlsy item {index} extra data is truncated")))?;
        let kind = GlossaryItemKind::from_byte(extra[0])?;
        let raw_style = u16::from_le_bytes([extra[2], extra[3]]);
        let style_index = if raw_style == u16::MAX {
            None
        } else if i16::try_from(raw_style).is_ok() {
            Some(raw_style)
        } else {
            return Err(corrupted(format!(
                "SttbfGlsy item {index} has a negative style index"
            )));
        };
        if kind == GlossaryItemKind::FormattedAutoCorrect && style_index.is_some() {
            return Err(corrupted(format!(
                "SttbfGlsy AutoCorrect item {index} uses a style"
            )));
        }
        items.push(RawItem {
            name,
            kind,
            style_index,
        });
        offset += 4;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfGlsy has trailing bytes"));
    }
    Ok(items)
}

fn parse_styles(data: &[u8]) -> Result<Vec<GlossaryStyle>> {
    let count = parse_sttb_header(data, 1, "SttbGlsyStyle")?;
    let mut styles = Vec::with_capacity(count);
    let mut offset = 6usize;
    for index in 0..count {
        let (name, next) = parse_string(data, offset, u16::MAX as usize, "SttbGlsyStyle", index)?;
        offset = next;
        let use_count = *data
            .get(offset)
            .ok_or_else(|| corrupted(format!("SttbGlsyStyle entry {index} is truncated")))?;
        styles.push(GlossaryStyle { name, use_count });
        offset += 1;
    }
    if offset != data.len() {
        return Err(corrupted("SttbGlsyStyle has trailing bytes"));
    }
    Ok(styles)
}

fn parse_sttb_header(data: &[u8], extra: u16, name: &str) -> Result<usize> {
    if data.len() > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    if data.len() < 6
        || read_u16(data, 0, &format!("{name} fExtend"))? != u16::MAX
        || read_u16(data, 4, &format!("{name} cbExtra"))? != extra
    {
        return Err(corrupted(format!("{name} has an invalid header")));
    }
    Ok(usize::from(read_u16(data, 2, &format!("{name} cData"))?))
}

fn parse_string(
    data: &[u8],
    offset: usize,
    max_units: usize,
    table: &str,
    index: usize,
) -> Result<(String, usize)> {
    let units = usize::from(read_u16(
        data,
        offset,
        &format!("{table} string {index} length"),
    )?);
    if units > max_units {
        return Err(corrupted(format!(
            "{table} string {index} exceeds {max_units} UTF-16 code units"
        )));
    }
    let start = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{table} string offset overflows")))?;
    let end = start
        .checked_add(
            units
                .checked_mul(2)
                .ok_or_else(|| corrupted(format!("{table} string size overflows")))?,
        )
        .ok_or_else(|| corrupted(format!("{table} string range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{table} string {index} is truncated")))?;
    Ok((
        decode_utf16(bytes, &format!("{table} string {index}"))?,
        end,
    ))
}

fn validate_item_shape(item: &GlossaryItem, index: usize) -> Result<()> {
    if item.name.encode_utf16().count() > MAX_ITEM_NAME_UNITS {
        return Err(corrupted(format!(
            "glossary item {index} name exceeds 32 UTF-16 code units"
        )));
    }
    if item.start_cp >= item.end_cp {
        return Err(corrupted(format!(
            "glossary item {index} has an empty or reversed CP range"
        )));
    }
    if item.kind == GlossaryItemKind::FormattedAutoCorrect && item.style_index.is_some() {
        return Err(corrupted(format!(
            "formatted AutoCorrect item {index} cannot use a style"
        )));
    }
    if item
        .style_index
        .is_some_and(|value| value > i16::MAX as u16)
    {
        return Err(corrupted(format!(
            "glossary item {index} style index exceeds the signed 16-bit range"
        )));
    }
    Ok(())
}

fn validate_style(style: &GlossaryStyle, index: usize) -> Result<()> {
    if style.name.encode_utf16().count() > u16::MAX as usize {
        return Err(corrupted(format!(
            "glossary style {index} name exceeds 65535 UTF-16 code units"
        )));
    }
    if style.use_count > MAX_STYLE_USE_COUNT {
        return Err(corrupted(format!(
            "glossary style {index} use count exceeds 0x32"
        )));
    }
    Ok(())
}

fn validate_metadata(value: &GlossaryMetadata) -> Result<()> {
    if value.items.len() > u16::MAX as usize || value.styles.len() > u16::MAX as usize {
        return Err(corrupted("glossary STTB count exceeds 65535 entries"));
    }
    if value.terminal_cp >= value.ignored_cp || value.ignored_cp >= value.main_text_length {
        return Err(corrupted(
            "PlcfGlsy terminal CPs are not strictly increasing within ccpText",
        ));
    }
    let mut actual_uses = vec![0u8; value.styles.len()];
    for (index, item) in value.items.iter().enumerate() {
        validate_item_shape(item, index)?;
        let expected_end = value
            .items
            .get(index + 1)
            .map_or(value.terminal_cp, |next| next.start_cp);
        if item.end_cp != expected_end {
            return Err(corrupted(format!(
                "glossary item {index} range is not contiguous with PlcfGlsy"
            )));
        }
        if item.end_cp >= value.main_text_length {
            return Err(corrupted(format!(
                "glossary item {index} range is outside ccpText"
            )));
        }
        if let Some(style_index) = item.style_index {
            let count = actual_uses
                .get_mut(usize::from(style_index))
                .ok_or_else(|| {
                    corrupted(format!("glossary item {index} style index is out of range"))
                })?;
            *count = count
                .checked_add(1)
                .ok_or_else(|| corrupted("glossary style use count overflows"))?;
        }
    }
    for (index, (style, actual)) in value.styles.iter().zip(actual_uses).enumerate() {
        validate_style(style, index)?;
        if style.use_count != actual {
            return Err(corrupted(format!(
                "glossary style {index} records {} uses but {actual} items refer to it",
                style.use_count
            )));
        }
    }
    Ok(())
}

fn sttb_size<'a>(
    strings: impl Iterator<Item = &'a str>,
    extra: usize,
    name: &str,
) -> Result<usize> {
    let mut size = 6usize;
    for string in strings {
        size = size
            .checked_add(2)
            .and_then(|value| value.checked_add(string.encode_utf16().count().checked_mul(2)?))
            .and_then(|value| value.checked_add(extra))
            .ok_or_else(|| corrupted(format!("{name} serialized size overflows")))?;
    }
    if size > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    Ok(size)
}

fn write_sttb_header(data: &mut Vec<u8>, count: usize, extra: u16, name: &str) -> Result<()> {
    let count = u16::try_from(count)
        .map_err(|_| corrupted(format!("{name} contains more than 65535 strings")))?;
    data.extend_from_slice(&u16::MAX.to_le_bytes());
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&extra.to_le_bytes());
    Ok(())
}

fn write_string(data: &mut Vec<u8>, value: &str, table: &str) -> Result<()> {
    let count = u16::try_from(value.encode_utf16().count())
        .map_err(|_| corrupted(format!("{table} string length exceeds u16")))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> GlossaryMetadata {
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new("Greeting", GlossaryItemKind::NamedAutoText, Some(0), 1, 5)
                    .unwrap(),
                GlossaryItem::try_new("teh", GlossaryItemKind::FormattedAutoCorrect, None, 5, 9)
                    .unwrap(),
            ],
            vec![GlossaryStyle::try_new("Normal", 1).unwrap()],
            9,
            10,
            11,
        )
        .unwrap()
    }

    #[test]
    fn round_trips_all_three_parallel_tables() {
        let metadata = sample();
        let tables = metadata.to_table_bytes().unwrap();
        let parsed = GlossaryMetadata::parse_table_bytes(
            tables.item_table(),
            tables.position_table(),
            tables.style_table(),
            11,
        )
        .unwrap();
        assert_eq!(parsed, metadata);
        assert_eq!(parsed.style_for_item(0).unwrap().name(), "Normal");
        assert!(parsed.style_for_item(1).is_none());
    }

    #[test]
    fn rejects_cross_table_and_lexical_inconsistencies() {
        let metadata = sample();
        let tables = metadata.to_table_bytes().unwrap();

        let mut items = tables.item_table().to_vec();
        let first_extra = 6 + 2 + "Greeting".encode_utf16().count() * 2;
        items[first_extra] = 0x09;
        assert!(
            GlossaryMetadata::parse_table_bytes(
                &items,
                tables.position_table(),
                tables.style_table(),
                11,
            )
            .is_err()
        );

        let mut positions = tables.position_table().to_vec();
        positions[4..8].copy_from_slice(&1u32.to_le_bytes());
        assert!(
            GlossaryMetadata::parse_table_bytes(
                tables.item_table(),
                &positions,
                tables.style_table(),
                11,
            )
            .is_err()
        );

        let wrong_count = vec![GlossaryStyle::try_new("Normal", 0).unwrap()];
        assert!(
            GlossaryMetadata::try_new(metadata.items.clone(), wrong_count, 9, 10, 11,).is_err()
        );
    }

    #[test]
    fn rejects_autocorrect_style_and_resource_limit_violations() {
        assert!(
            GlossaryItem::try_new("bad", GlossaryItemKind::FormattedAutoCorrect, Some(0), 0, 1,)
                .is_err()
        );
        assert!(
            GlossaryItem::try_new("x".repeat(33), GlossaryItemKind::NamedAutoText, None, 0, 1,)
                .is_err()
        );
        assert!(GlossaryStyle::try_new("Normal", 0x33).is_err());
    }
}
