//! Construction of independent Pages packages without bundled templates.

mod table_bootstrap;

pub(super) use table_bootstrap::bootstrap_first_table_graph;

use plist::Value;
use prost::Message;

use litchi_pages::section::{PageNumbering, Start};

use super::editor::PagesEditor;
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::identity::IWorkDocumentIdentity;
use crate::protobuf::{tp, tsa, tsd, tsk, tsp, tss, tst, tswp};
use crate::text::{ParagraphList, preset_style_object};
use crate::{IWorkPackage, IWorkThemeArchive, IWorkThemeExtensions, Result};

const DOCUMENT_ARCHIVE_ENTRY: &str = "Index/Document.iwa";
const STYLESHEET_ARCHIVE_ENTRY: &str = "Index/DocumentStylesheet.iwa";
const ANNOTATION_ARCHIVE_ENTRY: &str = "Index/AnnotationAuthorStorage.iwa";
const METADATA_ARCHIVE_ENTRY: &str = "Index/Metadata.iwa";
const PROPERTIES_ENTRY: &str = "Metadata/Properties.plist";
const DOCUMENT_IDENTIFIER_ENTRY: &str = "Metadata/DocumentIdentifier";
const BUILD_HISTORY_ENTRY: &str = "Metadata/BuildVersionHistory.plist";

const DEFAULT_LANGUAGE: &str = "en";
const DEFAULT_LOCALE: &str = "en_US";
const THEME_NAME: &str = "Litchi Blank";
const SECTION_NAME: &str = "Blank";
const LIST_STYLE_IDENTIFIER: &str = "text-0-liststyle-None";
const PARAGRAPH_STYLE_IDENTIFIER: &str = "text-20-paragraphstyle-Body 1";
const CHARACTER_STYLE_IDENTIFIER: &str = "character-style-null";
const LINE_STYLE_IDENTIFIER: &str = "line-0-shapestyle";
const SHAPE_STYLE_IDENTIFIER: &str = "shape-0-shapestyle";
const TEXT_BOX_STYLE_IDENTIFIER: &str = "textbox-0-shapestyle";
const IMAGE_STYLE_IDENTIFIER: &str = "image-0-imageStyle";
const MOVIE_STYLE_IDENTIFIER: &str = "movie-0-movieStyle";
const DRAWING_LINE_STYLE_IDENTIFIER: &str = "drawingline-0-drawinglineStyle";
const TOC_ENTRY_STYLE_IDENTIFIER: &str = "toc-entry-style-default";
const DROP_CAP_STYLE_IDENTIFIER: &str = "dropcap-style-0";
const COLUMN_STYLE_IDENTIFIER: &str = "column-style-default";
const CAPTION_STYLE_IDENTIFIER: &str = "captions-0-paragraphstyle-Caption Title";
const SVG_IMPORT_STYLE_IDENTIFIER: &str = "svgimport-0-shapestyle";
const BUILD_HISTORY_VALUE: &str = "Created by litchi-iwa";
const FILE_FORMAT_VERSION_STRING: &str = "14.4.1";
const FILE_FORMAT_VERSION: [u32; 3] = [14, 4, 1];
const PACKAGE_VERSION: [u32; 3] = [2, 0, 0];
const MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

const LETTER_WIDTH_POINTS: f32 = 612.0;
const LETTER_HEIGHT_POINTS: f32 = 792.0;
const BODY_MARGIN_POINTS: f32 = 72.0;
const HEADER_FOOTER_MARGIN_POINTS: f32 = 36.0;
const DEFAULT_PAGE_SCALE: f32 = 1.0;
const TEXT_BOX_PADDING_POINTS: f32 = 4.0;
const TEXT_BOX_COLUMN_COUNT: u32 = 1;
const DEFAULT_TEXT_PRESET_INDEX: u32 = 0;
const DEFAULT_STYLE_OVERRIDE_COUNT: u32 = 0;
const INITIAL_SAVE_TOKEN: u64 = 1;
const INITIAL_REVISION_SEQUENCE: i32 = 0;
const INITIAL_PAGE_NUMBER: u32 = 1;
#[cfg(test)]
const COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID: u64 = 3;
const TABLE_INFO_OBJECT_ID: u64 = 9;
const SOURCE_TABLE_SHEET_OBJECT_ID: u64 = 8;
const TABLE_MODEL_OBJECT_ID: u64 = 10;
const TABLE_LIST_STYLE_OBJECT_ID: u64 = 11;
const TABLE_PARAGRAPH_STYLE_OBJECT_ID: u64 = 12;
const TABLE_CHARACTER_STYLE_OBJECT_ID: u64 = 13;
const TABLE_SHAPE_STYLE_OBJECT_ID: u64 = 14;
const TABLE_MEDIA_STYLE_OBJECT_ID: u64 = 15;
const TABLE_DROP_CAP_STYLE_OBJECT_ID: u64 = 16;
const TABLE_SHEET_STYLE_OBJECT_ID: u64 = 17;
const TABLE_STYLE_OBJECT_ID: u64 = 18;
const TABLE_CELL_STYLE_OBJECT_ID: u64 = 19;
const TABLE_PRESET_OBJECT_ID: u64 = 20;
const TABLE_TILE_OBJECT_ID: u64 = 22;
const TABLE_ROW_HEADERS_OBJECT_ID: u64 = 23;
const TABLE_COLUMN_HEADERS_OBJECT_ID: u64 = 24;
const TABLE_STRING_LIST_OBJECT_ID: u64 = 25;
const TABLE_STYLE_LIST_OBJECT_ID: u64 = 26;
const TABLE_FORMULA_LIST_OBJECT_ID: u64 = 27;
const TABLE_FORMAT_LIST_OBJECT_ID: u64 = 28;
const TABLE_UID_MAP_OBJECT_ID: u64 = 29;
const TABLE_STROKE_SIDECAR_OBJECT_ID: u64 = 30;
const TABLE_CALCULATION_ENGINE_OBJECT_ID: u64 = 31;
const TABLE_STYLE_NETWORK_OBJECT_ID: u64 = 32;
const TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID: u64 = 33;
const TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID: u64 = 34;
const TABLE_FORMULA_OWNER_OBJECT_ID: u64 = 39;
const TABLE_STYLESHEET_OBJECT_ID: u64 = 40;
const TABLE_ATTACHMENT_OBJECT_ID: u64 = 41;
const TABLE_STYLE_OBJECT_IDS: &[u64] = &[
    TABLE_LIST_STYLE_OBJECT_ID,
    TABLE_PARAGRAPH_STYLE_OBJECT_ID,
    TABLE_CHARACTER_STYLE_OBJECT_ID,
    TABLE_SHAPE_STYLE_OBJECT_ID,
    TABLE_MEDIA_STYLE_OBJECT_ID,
    TABLE_DROP_CAP_STYLE_OBJECT_ID,
    TABLE_SHEET_STYLE_OBJECT_ID,
    TABLE_STYLE_OBJECT_ID,
    TABLE_CELL_STYLE_OBJECT_ID,
    TABLE_STYLESHEET_OBJECT_ID,
];
const TABLE_DOCUMENT_OBJECT_IDS: &[u64] = &[
    TABLE_INFO_OBJECT_ID,
    TABLE_MODEL_OBJECT_ID,
    TABLE_PRESET_OBJECT_ID,
    TABLE_TILE_OBJECT_ID,
    TABLE_ROW_HEADERS_OBJECT_ID,
    TABLE_COLUMN_HEADERS_OBJECT_ID,
    TABLE_STRING_LIST_OBJECT_ID,
    TABLE_STYLE_LIST_OBJECT_ID,
    TABLE_FORMULA_LIST_OBJECT_ID,
    TABLE_FORMAT_LIST_OBJECT_ID,
    TABLE_UID_MAP_OBJECT_ID,
    TABLE_STROKE_SIDECAR_OBJECT_ID,
    TABLE_STYLE_NETWORK_OBJECT_ID,
    TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID,
    TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID,
];
const TABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_DRAWABLE_FLAGS: u32 = 3;
const TABLE_ATTACHMENT_OFFSET_TYPE: u32 = 0;
const TABLE_ATTACHMENT_OFFSET_POINTS: f32 = 0.0;
const TABLE_ROTATION_DEGREES: f32 = 0.0;
const TABLE_CALCULATION_COMPONENT_VERSION: [u32; 3] = [3, 2, 10];
const DEFAULT_TABLE_WIDTH_POINTS: f32 = 468.0;
const DEFAULT_TABLE_ROW_HEIGHT_POINTS: f32 = 22.73;
const DEFAULT_TABLE_X_POINTS: f32 = 0.0;
const DEFAULT_TABLE_Y_POINTS: f32 = 0.0;

const COLOR_PRESET_COUNT: usize = 30;
const GRADIENT_FILL_PRESET_COUNT: usize = 6;
const IMAGE_FILL_PRESET_COUNT: usize = 6;
const SHADOW_PRESET_COUNT: usize = 8;
const LINE_STYLE_PRESET_COUNT: usize = 1;
const SHAPE_STYLE_PRESET_COUNT: usize = 1;
const TEXT_BOX_STYLE_PRESET_COUNT: usize = 1;
const IMAGE_STYLE_PRESET_COUNT: usize = 1;
const MOVIE_STYLE_PRESET_COUNT: usize = 1;
const DRAWING_LINE_STYLE_PRESET_COUNT: usize = 1;
const TOC_ENTRY_STYLE_PRESET_COUNT: usize = 1;
const TOC_SETTINGS_PRESET_COUNT: usize = 1;
const CHARACTER_STYLE_PRESET_COUNT: usize = 1;
const DROP_CAP_STYLE_PRESET_COUNT: usize = 1;
const CAPTION_STYLE_PRESET_COUNT: usize = 1;
const SVG_IMPORT_STYLE_PRESET_COUNT: usize = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u64)]
enum PagesObjectId {
    Document = 1,
    PackageMetadata = 2,
    // `COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID` is reserved for TSCKDocumentSupport.
    // Keep the low identifier range available to the shared source-built table graph.
    Stylesheet = 101,
    Theme = 102,
    Body = 103,
    Settings = 104,
    Section = 105,
    SectionTemplate = 106,
    ListStyle = 107,
    ParagraphStyle = 108,
    CharacterStyle = 109,
    LineStyle = 110,
    ShapeStyle = 111,
    TextBoxStyle = 112,
    ImageStyle = 113,
    MovieStyle = 114,
    DrawingLineStyle = 115,
    TocEntryStyle = 116,
    DropCapStyle = 117,
    BaseColumnStyle = 118,
    ColumnStyle = 119,
    TocSettings = 120,
    CaptionStyle = 121,
    SvgImportStyle = 122,
    HeaderPrimary = 123,
    HeaderEven = 124,
    HeaderFirst = 125,
    FooterPrimary = 126,
    FooterEven = 127,
    FooterFirst = 128,
    AnnotationAuthorStorage = 129,
    BulletListStyle = 130,
    NumberedListStyle = 131,
    ParagraphStyleTitle = 132,
    ParagraphStyleSubtitle = 133,
    ParagraphStyleHeading = 134,
    ParagraphStyleHeading2 = 135,
    ParagraphStyleHeading3 = 136,
    ParagraphStyleHeadingRed = 137,
    ParagraphStyleCaption = 138,
    ParagraphStyleHeaderFooter = 139,
    ParagraphStyleFootnote = 140,
    ParagraphStyleLabel = 141,
    ParagraphStyleLabelDark = 142,
}

impl PagesObjectId {
    const fn value(self) -> u64 {
        self as u64
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PagesParagraphStylePreset {
    Title,
    Subtitle,
    Heading,
    Heading2,
    Heading3,
    HeadingRed,
    Body,
    Caption,
    HeaderFooter,
    Footnote,
    Label,
    LabelDark,
}

impl PagesParagraphStylePreset {
    const ALL: [Self; 12] = [
        Self::Title,
        Self::Subtitle,
        Self::Heading,
        Self::Heading2,
        Self::Heading3,
        Self::HeadingRed,
        Self::Body,
        Self::Caption,
        Self::HeaderFooter,
        Self::Footnote,
        Self::Label,
        Self::LabelDark,
    ];

    const fn object_id(self) -> PagesObjectId {
        match self {
            Self::Title => PagesObjectId::ParagraphStyleTitle,
            Self::Subtitle => PagesObjectId::ParagraphStyleSubtitle,
            Self::Heading => PagesObjectId::ParagraphStyleHeading,
            Self::Heading2 => PagesObjectId::ParagraphStyleHeading2,
            Self::Heading3 => PagesObjectId::ParagraphStyleHeading3,
            Self::HeadingRed => PagesObjectId::ParagraphStyleHeadingRed,
            Self::Body => PagesObjectId::ParagraphStyle,
            Self::Caption => PagesObjectId::ParagraphStyleCaption,
            Self::HeaderFooter => PagesObjectId::ParagraphStyleHeaderFooter,
            Self::Footnote => PagesObjectId::ParagraphStyleFootnote,
            Self::Label => PagesObjectId::ParagraphStyleLabel,
            Self::LabelDark => PagesObjectId::ParagraphStyleLabelDark,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Title => "Title",
            Self::Subtitle => "Subtitle",
            Self::Heading => "Heading",
            Self::Heading2 => "Heading 2",
            Self::Heading3 => "Heading 3",
            Self::HeadingRed => "Heading Red",
            Self::Body => "Body",
            Self::Caption => "Caption",
            Self::HeaderFooter => "Header & Footer",
            Self::Footnote => "Footnote",
            Self::Label => "Label",
            Self::LabelDark => "Label Dark",
        }
    }

    const fn style_identifier(self) -> &'static str {
        match self {
            Self::Title => "text-0-paragraphstyle-Title",
            Self::Subtitle => "text-6-paragraphstyle-Subtitle",
            Self::Heading => "text-11-paragraphstyle-Heading 1",
            Self::Heading2 => "text-12-paragraphstyle-Heading 2",
            Self::Heading3 => "text-13-paragraphstyle-Heading 3",
            Self::HeadingRed => "text-14-paragraphstyle-Heading 4",
            Self::Body => PARAGRAPH_STYLE_IDENTIFIER,
            Self::Caption => "text-26-paragraphstyle-Caption 3",
            Self::HeaderFooter => "text-29-paragraphstyle-Header & Footer",
            Self::Footnote => "text-31-paragraphstyle-Footnote Text",
            Self::Label => "text-32-paragraphstyle-Label",
            Self::LabelDark => "text-33-paragraphstyle-Label Dark",
        }
    }
}

const LAST_PARAGRAPH_STYLE_OBJECT_ID: u64 =
    PagesParagraphStylePreset::LabelDark.object_id().value();

const STYLESHEET_OBJECTS: [PagesObjectId; 29] = [
    PagesObjectId::Stylesheet,
    PagesObjectId::ListStyle,
    PagesObjectId::BulletListStyle,
    PagesObjectId::NumberedListStyle,
    PagesObjectId::ParagraphStyleTitle,
    PagesObjectId::ParagraphStyleSubtitle,
    PagesObjectId::ParagraphStyleHeading,
    PagesObjectId::ParagraphStyleHeading2,
    PagesObjectId::ParagraphStyleHeading3,
    PagesObjectId::ParagraphStyleHeadingRed,
    PagesObjectId::ParagraphStyle,
    PagesObjectId::ParagraphStyleCaption,
    PagesObjectId::ParagraphStyleHeaderFooter,
    PagesObjectId::ParagraphStyleFootnote,
    PagesObjectId::ParagraphStyleLabel,
    PagesObjectId::ParagraphStyleLabelDark,
    PagesObjectId::CharacterStyle,
    PagesObjectId::LineStyle,
    PagesObjectId::ShapeStyle,
    PagesObjectId::TextBoxStyle,
    PagesObjectId::ImageStyle,
    PagesObjectId::MovieStyle,
    PagesObjectId::DrawingLineStyle,
    PagesObjectId::TocEntryStyle,
    PagesObjectId::DropCapStyle,
    PagesObjectId::BaseColumnStyle,
    PagesObjectId::ColumnStyle,
    PagesObjectId::CaptionStyle,
    PagesObjectId::SvgImportStyle,
];

const DOCUMENT_OBJECTS: [PagesObjectId; 13] = [
    PagesObjectId::Document,
    PagesObjectId::Theme,
    PagesObjectId::Body,
    PagesObjectId::Settings,
    PagesObjectId::Section,
    PagesObjectId::SectionTemplate,
    PagesObjectId::TocSettings,
    PagesObjectId::HeaderPrimary,
    PagesObjectId::HeaderEven,
    PagesObjectId::HeaderFirst,
    PagesObjectId::FooterPrimary,
    PagesObjectId::FooterEven,
    PagesObjectId::FooterFirst,
];

const IDENTIFIED_STYLES: [(PagesObjectId, &str); 25] = [
    (PagesObjectId::ListStyle, LIST_STYLE_IDENTIFIER),
    (
        PagesObjectId::ParagraphStyleTitle,
        PagesParagraphStylePreset::Title.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleSubtitle,
        PagesParagraphStylePreset::Subtitle.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleHeading,
        PagesParagraphStylePreset::Heading.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleHeading2,
        PagesParagraphStylePreset::Heading2.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleHeading3,
        PagesParagraphStylePreset::Heading3.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleHeadingRed,
        PagesParagraphStylePreset::HeadingRed.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyle,
        PagesParagraphStylePreset::Body.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleCaption,
        PagesParagraphStylePreset::Caption.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleHeaderFooter,
        PagesParagraphStylePreset::HeaderFooter.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleFootnote,
        PagesParagraphStylePreset::Footnote.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleLabel,
        PagesParagraphStylePreset::Label.style_identifier(),
    ),
    (
        PagesObjectId::ParagraphStyleLabelDark,
        PagesParagraphStylePreset::LabelDark.style_identifier(),
    ),
    (PagesObjectId::CharacterStyle, CHARACTER_STYLE_IDENTIFIER),
    (PagesObjectId::LineStyle, LINE_STYLE_IDENTIFIER),
    (PagesObjectId::ShapeStyle, SHAPE_STYLE_IDENTIFIER),
    (PagesObjectId::TextBoxStyle, TEXT_BOX_STYLE_IDENTIFIER),
    (PagesObjectId::ImageStyle, IMAGE_STYLE_IDENTIFIER),
    (PagesObjectId::MovieStyle, MOVIE_STYLE_IDENTIFIER),
    (
        PagesObjectId::DrawingLineStyle,
        DRAWING_LINE_STYLE_IDENTIFIER,
    ),
    (PagesObjectId::TocEntryStyle, TOC_ENTRY_STYLE_IDENTIFIER),
    (PagesObjectId::DropCapStyle, DROP_CAP_STYLE_IDENTIFIER),
    (PagesObjectId::BaseColumnStyle, COLUMN_STYLE_IDENTIFIER),
    (PagesObjectId::CaptionStyle, CAPTION_STYLE_IDENTIFIER),
    (PagesObjectId::SvgImportStyle, SVG_IMPORT_STYLE_IDENTIFIER),
];

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum PagesMessageType {
    AnnotationAuthorStorage = 213,
    Stylesheet = 401,
    Storage = 2_001,
    CharacterStyle = 2_021,
    ParagraphStyle = 2_022,
    ListStyle = 2_023,
    ColumnStyle = 2_024,
    ShapeStyle = 2_025,
    TocEntryStyle = 2_026,
    TocSettings = 2_051,
    MediaStyle = 3_016,
    Document = 10_000,
    Theme = 10_001,
    Section = 10_011,
    Settings = 10_012,
    DropCapStyle = 10_024,
    SectionTemplate = 10_143,
    PackageMetadata = 11_006,
}

impl PagesMessageType {
    const fn value(self) -> u32 {
        self as u32
    }
}

/// Builder for a new, independent Pages document.
///
/// The resulting package is encoded from typed protobuf values and newly
/// generated identities. It does not copy or embed an Apple template.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesDocumentBuilder {
    body_text: String,
    language: String,
    locale: String,
    initial_tables: Vec<InitialPagesTable>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct InitialPagesTable {
    name: String,
    rows: usize,
    columns: usize,
}

impl Default for PagesDocumentBuilder {
    fn default() -> Self {
        Self {
            body_text: String::new(),
            language: DEFAULT_LANGUAGE.to_owned(),
            locale: DEFAULT_LOCALE.to_owned(),
            initial_tables: Vec::new(),
        }
    }
}

impl PagesDocumentBuilder {
    /// Start a blank Pages document with US Letter page geometry.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the initial body text.
    pub fn body_text(mut self, body_text: impl Into<String>) -> Self {
        self.body_text = body_text.into();
        self
    }

    /// Set the BCP 47 document language, such as `en` or `fr`.
    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = language.into();
        self
    }

    /// Set the document creation locale, such as `en_US` or `fr_FR`.
    pub fn locale(mut self, locale: impl Into<String>) -> Self {
        self.locale = locale.into();
        self
    }

    /// Append one empty native table to the initial document body.
    ///
    /// The table owns independent cell storage and is immediately writable
    /// through [`PagesEditor::set_table_cell`](super::editor::PagesEditor::set_table_cell).
    pub fn body_table(mut self, name: impl Into<String>, rows: usize, columns: usize) -> Self {
        self.initial_tables.push(InitialPagesTable {
            name: name.into(),
            rows,
            columns,
        });
        self
    }

    /// Build a mutable editor for the new document.
    pub fn build(self) -> Result<PagesEditor> {
        PagesEditor::from_package(self.build_package()?)
    }

    /// Build the underlying package for lower-level IWA manipulation.
    pub(crate) fn build_package(self) -> Result<IWorkPackage> {
        crate::text::TextLanguageTag::new(self.language.as_str())?;
        if self.locale.trim().is_empty() {
            return Err(crate::Error::InvalidFormat(
                "Pages document locale cannot be empty".to_owned(),
            ));
        }
        for table in &self.initial_tables {
            validate_initial_table(table)?;
        }

        let first_table = self.initial_tables.first();

        let identity = IWorkDocumentIdentity::generate();
        let mut package = IWorkPackage::new();
        package.replace_archive(
            DOCUMENT_ARCHIVE_ENTRY,
            &document_archive(&self.body_text, &self.language, &self.locale, first_table)?,
        )?;
        package.replace_archive(STYLESHEET_ARCHIVE_ENTRY, &stylesheet_archive()?)?;
        package.replace_archive(
            ANNOTATION_ARCHIVE_ENTRY,
            &annotation_author_storage_archive()?,
        )?;
        package.replace_archive(
            METADATA_ARCHIVE_ENTRY,
            &metadata_archive(&identity, first_table.is_some())?,
        )?;
        if let Some(table) = first_table {
            install_initial_table_graph(&mut package, table, &self.language, &self.locale)?;
        }
        insert_property_lists(&mut package, &identity)?;
        if self.initial_tables.len() > 1 {
            let mut editor = PagesEditor::from_package(package)?;
            for table in &self.initial_tables[1..] {
                let anchor = editor.body_text()?.encode_utf16().count();
                editor.add_table(anchor, &table.name, table.rows, table.columns)?;
            }
            package = editor.package().clone();
        }
        Ok(package)
    }
}

fn document_archive(
    body_text: &str,
    language: &str,
    locale: &str,
    initial_table: Option<&InitialPagesTable>,
) -> Result<Archive> {
    let document = tp::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive {
                locale_identifier: Some(locale.to_owned()),
                annotation_author_storage: Some(reference(PagesObjectId::AnnotationAuthorStorage)),
                creation_locale_identifier: Some(locale.to_owned()),
                prevent_image_conversion_on_open: Some(true),
                has_user_defined_locale: Some(false),
                ..Default::default()
            },
            document_language: Some(language.to_owned()),
            calculation_engine: initial_table
                .map(|_| raw_reference(TABLE_CALCULATION_ENGINE_OBJECT_ID)),
            function_browser_state: initial_table
                .map(|_| raw_reference(TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID)),
            custom_format_list: initial_table
                .map(|_| raw_reference(TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID)),
            ..Default::default()
        },
        stylesheet: Some(reference(PagesObjectId::Stylesheet)),
        body_storage: Some(reference(PagesObjectId::Body)),
        theme: Some(reference(PagesObjectId::Theme)),
        settings: Some(reference(PagesObjectId::Settings)),
        page_width: Some(LETTER_WIDTH_POINTS),
        page_height: Some(LETTER_HEIGHT_POINTS),
        left_margin: Some(BODY_MARGIN_POINTS),
        right_margin: Some(BODY_MARGIN_POINTS),
        top_margin: Some(BODY_MARGIN_POINTS),
        bottom_margin: Some(BODY_MARGIN_POINTS),
        header_margin: Some(HEADER_FOOTER_MARGIN_POINTS),
        footer_margin: Some(HEADER_FOOTER_MARGIN_POINTS),
        page_scale: Some(DEFAULT_PAGE_SCALE),
        ..Default::default()
    };
    let table_anchor = initial_table
        .map(|_| {
            u32::try_from(body_text.encode_utf16().count()).map_err(|_| {
                crate::Error::InvalidFormat(
                    "Pages table anchor exceeds the UTF-16 index limit".to_owned(),
                )
            })
        })
        .transpose()?;
    let mut body_contents = body_text.to_owned();
    if initial_table.is_some() {
        body_contents.push('\u{fffc}');
    }
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        style_sheet: Some(reference(PagesObjectId::Stylesheet)),
        text: vec![body_contents],
        in_document: Some(true),
        table_para_style: Some(object_table(Some(PagesObjectId::ParagraphStyle))),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_table(Some(PagesObjectId::ListStyle))),
        table_layout_style: Some(object_table(Some(PagesObjectId::ColumnStyle))),
        table_para_starts: Some(zero_para_data()),
        table_section: Some(object_table(Some(PagesObjectId::Section))),
        table_attachment: table_anchor.map(|character_index| tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index,
                object: Some(raw_reference(TABLE_ATTACHMENT_OBJECT_ID)),
            }],
        }),
        table_language: Some(tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language.to_owned()),
            }],
        }),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_table(None)),
        ..Default::default()
    };

    let theme = IWorkThemeArchive::new(
        tss::ThemeArchive {
            theme_identifier: Some(THEME_NAME.to_owned()),
            document_stylesheet: Some(reference(PagesObjectId::Stylesheet)),
            color_presets: repeated(COLOR_PRESET_COUNT, black),
            ..Default::default()
        },
        IWorkThemeExtensions {
            drawing: Some(tsd::ThemePresetsArchive {
                gradient_fill_presets: repeated(GRADIENT_FILL_PRESET_COUNT, solid_fill),
                image_fill_presets: repeated(IMAGE_FILL_PRESET_COUNT, solid_fill),
                shadow_presets: repeated(SHADOW_PRESET_COUNT, tsd::ShadowArchive::default),
                line_style_presets: repeated_reference(
                    LINE_STYLE_PRESET_COUNT,
                    PagesObjectId::LineStyle,
                ),
                shape_style_presets: repeated_reference(
                    SHAPE_STYLE_PRESET_COUNT,
                    PagesObjectId::ShapeStyle,
                ),
                textbox_style_presets: repeated_reference(
                    TEXT_BOX_STYLE_PRESET_COUNT,
                    PagesObjectId::TextBoxStyle,
                ),
                image_style_presets: repeated_reference(
                    IMAGE_STYLE_PRESET_COUNT,
                    PagesObjectId::ImageStyle,
                ),
                movie_style_presets: repeated_reference(
                    MOVIE_STYLE_PRESET_COUNT,
                    PagesObjectId::MovieStyle,
                ),
                drawing_line_style_presets: repeated_reference(
                    DRAWING_LINE_STYLE_PRESET_COUNT,
                    PagesObjectId::DrawingLineStyle,
                ),
            }),
            text: Some(tswp::ThemePresetsArchive {
                list_style_presets: [
                    PagesObjectId::ListStyle,
                    PagesObjectId::BulletListStyle,
                    PagesObjectId::NumberedListStyle,
                ]
                .into_iter()
                .map(reference)
                .collect(),
                toc_entry_style_presets: repeated_reference(
                    TOC_ENTRY_STYLE_PRESET_COUNT,
                    PagesObjectId::TocEntryStyle,
                ),
                toc_settings_presets: repeated_reference(
                    TOC_SETTINGS_PRESET_COUNT,
                    PagesObjectId::TocSettings,
                ),
                character_style_presets: repeated_reference(
                    CHARACTER_STYLE_PRESET_COUNT,
                    PagesObjectId::CharacterStyle,
                ),
                paragraph_style_presets: PagesParagraphStylePreset::ALL
                    .into_iter()
                    .map(PagesParagraphStylePreset::object_id)
                    .map(reference)
                    .collect(),
                dropcap_style_presets: repeated_reference(
                    DROP_CAP_STYLE_PRESET_COUNT,
                    PagesObjectId::DropCapStyle,
                ),
                ..Default::default()
            }),
            chart: None,
            table: initial_table.map(|_| tst::ThemePresetsArchive {
                table_style_presets: vec![raw_reference(TABLE_PRESET_OBJECT_ID)],
                ..Default::default()
            }),
            application: Some(tsa::ThemePresetsArchive {
                caption_style_presets: repeated_reference(
                    CAPTION_STYLE_PRESET_COUNT,
                    PagesObjectId::CaptionStyle,
                ),
                svg_import_style_presets: repeated_reference(
                    SVG_IMPORT_STYLE_PRESET_COUNT,
                    PagesObjectId::SvgImportStyle,
                ),
            }),
        },
    );
    let theme_references = [
        PagesObjectId::Stylesheet,
        PagesObjectId::ListStyle,
        PagesObjectId::BulletListStyle,
        PagesObjectId::NumberedListStyle,
    ]
    .into_iter()
    .chain(
        PagesParagraphStylePreset::ALL
            .into_iter()
            .map(PagesParagraphStylePreset::object_id),
    )
    .chain([
        PagesObjectId::CharacterStyle,
        PagesObjectId::LineStyle,
        PagesObjectId::ShapeStyle,
        PagesObjectId::TextBoxStyle,
        PagesObjectId::ImageStyle,
        PagesObjectId::MovieStyle,
        PagesObjectId::DrawingLineStyle,
        PagesObjectId::TocEntryStyle,
        PagesObjectId::TocSettings,
        PagesObjectId::DropCapStyle,
        PagesObjectId::CaptionStyle,
        PagesObjectId::SvgImportStyle,
    ])
    .collect::<Vec<_>>();

    let mut objects = vec![
        object(
            PagesObjectId::Document,
            PagesMessageType::Document,
            document,
            &[
                PagesObjectId::Stylesheet,
                PagesObjectId::Body,
                PagesObjectId::Theme,
                PagesObjectId::Settings,
                PagesObjectId::AnnotationAuthorStorage,
            ],
        )?,
        raw_object(
            PagesObjectId::Theme,
            PagesMessageType::Theme,
            theme.encode()?,
            &theme_references,
        )?,
        object(
            PagesObjectId::Body,
            PagesMessageType::Storage,
            body,
            &[
                PagesObjectId::Stylesheet,
                PagesObjectId::ParagraphStyle,
                PagesObjectId::ListStyle,
                PagesObjectId::ColumnStyle,
                PagesObjectId::Section,
            ],
        )?,
        object(
            PagesObjectId::Settings,
            PagesMessageType::Settings,
            tp::SettingsArchive {
                body: Some(true),
                headers: Some(false),
                footers: Some(false),
                language: Some(language.to_owned()),
                creation_locale: Some(locale.to_owned()),
                last_locale: Some(locale.to_owned()),
                ..Default::default()
            },
            &[],
        )?,
        object(
            PagesObjectId::Section,
            PagesMessageType::Section,
            tp::SectionArchive {
                inherit_previous_header_footer: Some(true),
                section_template_first_page_different: Some(false),
                section_template_even_odd_pages_different: Some(false),
                section_start_kind: Some(Start::NextPage.as_raw()),
                section_page_number_kind: Some(PageNumbering::ContinueFromPrevious.as_raw()),
                section_page_number_start: Some(INITIAL_PAGE_NUMBER),
                first_section_template_page: Some(reference(PagesObjectId::SectionTemplate)),
                even_section_template_page: Some(reference(PagesObjectId::SectionTemplate)),
                odd_section_template_page: Some(reference(PagesObjectId::SectionTemplate)),
                name: Some(SECTION_NAME.to_owned()),
                section_template_first_page_hides_header_footer: Some(false),
                ..Default::default()
            },
            &[PagesObjectId::SectionTemplate],
        )?,
        object(
            PagesObjectId::SectionTemplate,
            PagesMessageType::SectionTemplate,
            tp::SectionTemplateArchive {
                headers: [
                    PagesObjectId::HeaderPrimary,
                    PagesObjectId::HeaderEven,
                    PagesObjectId::HeaderFirst,
                ]
                .into_iter()
                .map(reference)
                .collect(),
                footers: [
                    PagesObjectId::FooterPrimary,
                    PagesObjectId::FooterEven,
                    PagesObjectId::FooterFirst,
                ]
                .into_iter()
                .map(reference)
                .collect(),
                ..Default::default()
            },
            &[
                PagesObjectId::HeaderPrimary,
                PagesObjectId::HeaderEven,
                PagesObjectId::HeaderFirst,
                PagesObjectId::FooterPrimary,
                PagesObjectId::FooterEven,
                PagesObjectId::FooterFirst,
            ],
        )?,
        object(
            PagesObjectId::HeaderPrimary,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::HeaderEven,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::HeaderFirst,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::FooterPrimary,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::FooterEven,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::FooterFirst,
            PagesMessageType::Storage,
            header_footer_storage(),
            &[PagesObjectId::ParagraphStyle, PagesObjectId::ListStyle],
        )?,
        object(
            PagesObjectId::TocSettings,
            PagesMessageType::TocSettings,
            tswp::TocSettingsArchive::default(),
            &[],
        )?,
    ];
    if initial_table.is_some() {
        for identifier in [
            TABLE_CALCULATION_ENGINE_OBJECT_ID,
            TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID,
            TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID,
        ] {
            append_object_reference(&mut objects[0], identifier);
        }
        append_object_reference(&mut objects[1], TABLE_PRESET_OBJECT_ID);
        append_object_reference(&mut objects[2], TABLE_ATTACHMENT_OBJECT_ID);
        objects.push(raw_object_with_id(
            TABLE_ATTACHMENT_OBJECT_ID,
            TABLE_ATTACHMENT_MESSAGE_TYPE,
            tswp::DrawableAttachmentArchive {
                drawable: Some(raw_reference(TABLE_INFO_OBJECT_ID)),
                h_offset_type: Some(TABLE_ATTACHMENT_OFFSET_TYPE),
                h_offset: Some(TABLE_ATTACHMENT_OFFSET_POINTS),
                v_offset_type: Some(TABLE_ATTACHMENT_OFFSET_TYPE),
                v_offset: Some(TABLE_ATTACHMENT_OFFSET_POINTS),
            }
            .encode_to_vec(),
            &[TABLE_INFO_OBJECT_ID],
        )?);
    }
    Ok(Archive { objects })
}

fn stylesheet_archive() -> Result<Archive> {
    let mut objects = vec![object(
        PagesObjectId::Stylesheet,
        PagesMessageType::Stylesheet,
        tss::StylesheetArchive {
            styles: STYLESHEET_OBJECTS
                .iter()
                .copied()
                .skip(1)
                .map(reference)
                .collect(),
            identifier_to_style_map: IDENTIFIED_STYLES
                .iter()
                .map(|(identifier, package_identifier)| {
                    tss::stylesheet_archive::IdentifiedStyleEntry {
                        identifier: (*package_identifier).to_owned(),
                        style: reference(*identifier),
                    }
                })
                .collect(),
            is_locked: Some(false),
            can_cull_styles: Some(true),
            ..Default::default()
        },
        &STYLESHEET_OBJECTS[1..],
    )?];
    objects.extend([
        object(
            PagesObjectId::ListStyle,
            PagesMessageType::ListStyle,
            tswp::ListStyleArchive {
                super_: style("None", LIST_STYLE_IDENTIFIER),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::CharacterStyle,
            PagesMessageType::CharacterStyle,
            tswp::CharacterStyleArchive {
                super_: style("None", CHARACTER_STYLE_IDENTIFIER),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::LineStyle,
            PagesMessageType::ShapeStyle,
            shape_style_archive(
                "Line",
                LINE_STYLE_IDENTIFIER,
                tsd::FillArchive::default(),
                Some(PagesObjectId::ParagraphStyle),
            ),
            &[PagesObjectId::Stylesheet, PagesObjectId::ParagraphStyle],
        )?,
        object(
            PagesObjectId::ShapeStyle,
            PagesMessageType::ShapeStyle,
            shape_style_archive(
                "Shape",
                SHAPE_STYLE_IDENTIFIER,
                solid_fill(),
                Some(PagesObjectId::ParagraphStyle),
            ),
            &[PagesObjectId::Stylesheet, PagesObjectId::ParagraphStyle],
        )?,
        object(
            PagesObjectId::TextBoxStyle,
            PagesMessageType::ShapeStyle,
            shape_style_archive(
                "Text Box",
                TEXT_BOX_STYLE_IDENTIFIER,
                tsd::FillArchive::default(),
                Some(PagesObjectId::ParagraphStyle),
            ),
            &[PagesObjectId::Stylesheet, PagesObjectId::ParagraphStyle],
        )?,
        object(
            PagesObjectId::ImageStyle,
            PagesMessageType::MediaStyle,
            media_style_archive("Image", IMAGE_STYLE_IDENTIFIER),
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::MovieStyle,
            PagesMessageType::MediaStyle,
            media_style_archive("Movie", MOVIE_STYLE_IDENTIFIER),
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::DrawingLineStyle,
            PagesMessageType::ShapeStyle,
            shape_style_archive(
                "Drawing Line",
                DRAWING_LINE_STYLE_IDENTIFIER,
                tsd::FillArchive::default(),
                None,
            ),
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::TocEntryStyle,
            PagesMessageType::TocEntryStyle,
            tswp::TocEntryStyleArchive {
                super_: tswp::ParagraphStyleArchive {
                    super_: style("TOC", TOC_ENTRY_STYLE_IDENTIFIER),
                    override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                    ..Default::default()
                },
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::DropCapStyle,
            PagesMessageType::DropCapStyle,
            tswp::DropCapStyleArchive {
                super_: style("Drop Cap", DROP_CAP_STYLE_IDENTIFIER),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::BaseColumnStyle,
            PagesMessageType::ColumnStyle,
            tswp::ColumnStyleArchive {
                super_: style("None", COLUMN_STYLE_IDENTIFIER),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::ColumnStyle,
            PagesMessageType::ColumnStyle,
            tswp::ColumnStyleArchive {
                super_: tss::StyleArchive {
                    parent: Some(reference(PagesObjectId::BaseColumnStyle)),
                    is_variation: Some(true),
                    stylesheet: Some(reference(PagesObjectId::Stylesheet)),
                    ..Default::default()
                },
                override_count: Some(1),
                column_properties: Some(tswp::ColumnStylePropertiesArchive {
                    writing_direction: Some(
                        tswp::WritingDirectionType::KWritingDirectionLeftToRight as i32,
                    ),
                    ..Default::default()
                }),
            },
            &[PagesObjectId::Stylesheet, PagesObjectId::BaseColumnStyle],
        )?,
        object(
            PagesObjectId::CaptionStyle,
            PagesMessageType::ParagraphStyle,
            tswp::ParagraphStyleArchive {
                super_: style("Object Title", CAPTION_STYLE_IDENTIFIER),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?,
        object(
            PagesObjectId::SvgImportStyle,
            PagesMessageType::ShapeStyle,
            shape_style_archive(
                "SVG Import",
                SVG_IMPORT_STYLE_IDENTIFIER,
                solid_fill(),
                None,
            ),
            &[PagesObjectId::Stylesheet],
        )?,
    ]);
    for preset in PagesParagraphStylePreset::ALL {
        objects.push(object(
            preset.object_id(),
            PagesMessageType::ParagraphStyle,
            tswp::ParagraphStyleArchive {
                super_: style(preset.name(), preset.style_identifier()),
                override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
                ..Default::default()
            },
            &[PagesObjectId::Stylesheet],
        )?);
    }
    objects.push(preset_style_object(
        PagesObjectId::BulletListStyle.value(),
        PagesObjectId::Stylesheet.value(),
        ParagraphList::Bullet,
    )?);
    objects.push(preset_style_object(
        PagesObjectId::NumberedListStyle.value(),
        PagesObjectId::Stylesheet.value(),
        ParagraphList::Numbered,
    )?);
    Ok(Archive { objects })
}

fn metadata_archive(identity: &IWorkDocumentIdentity, has_initial_table: bool) -> Result<Archive> {
    let mut stylesheet = component(PagesObjectId::Stylesheet, "DocumentStylesheet");
    stylesheet.object_uuid_map_entries = STYLESHEET_OBJECTS
        .iter()
        .copied()
        .map(object_uuid)
        .collect();
    if has_initial_table {
        stylesheet
            .object_uuid_map_entries
            .extend(TABLE_STYLE_OBJECT_IDS.iter().copied().map(object_uuid_raw));
    }

    let mut document = component(PagesObjectId::Document, "Document");
    document.object_uuid_map_entries = DOCUMENT_OBJECTS.iter().copied().map(object_uuid).collect();
    if has_initial_table {
        document.object_uuid_map_entries.extend(
            TABLE_DOCUMENT_OBJECT_IDS
                .iter()
                .copied()
                .chain(std::iter::once(TABLE_ATTACHMENT_OBJECT_ID))
                .map(object_uuid_raw),
        );
    }
    document.external_references = std::iter::once(None)
        .chain(STYLESHEET_OBJECTS.iter().copied().map(Some))
        .map(|object_identifier| tsp::ComponentExternalReference {
            component_identifier: PagesObjectId::Stylesheet.value(),
            object_identifier: object_identifier.map(PagesObjectId::value),
            is_weak: None,
        })
        .collect();
    document
        .external_references
        .push(tsp::ComponentExternalReference {
            component_identifier: PagesObjectId::AnnotationAuthorStorage.value(),
            object_identifier: None,
            is_weak: None,
        });
    if has_initial_table {
        document
            .external_references
            .extend(
                TABLE_STYLE_OBJECT_IDS
                    .iter()
                    .copied()
                    .map(|object_identifier| tsp::ComponentExternalReference {
                        component_identifier: PagesObjectId::Stylesheet.value(),
                        object_identifier: Some(object_identifier),
                        is_weak: None,
                    }),
            );
        document
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: TABLE_CALCULATION_ENGINE_OBJECT_ID,
                object_identifier: None,
                is_weak: None,
            });
    }

    let annotation = component(
        PagesObjectId::AnnotationAuthorStorage,
        "AnnotationAuthorStorage",
    );

    let mut components = vec![stylesheet, annotation];
    if has_initial_table {
        let mut calculation = component_raw(
            TABLE_CALCULATION_ENGINE_OBJECT_ID,
            "CalculationEngine",
            &TABLE_CALCULATION_COMPONENT_VERSION,
        );
        calculation.object_uuid_map_entries = [
            TABLE_CALCULATION_ENGINE_OBJECT_ID,
            TABLE_FORMULA_OWNER_OBJECT_ID,
        ]
        .into_iter()
        .map(object_uuid_raw)
        .collect();
        calculation.external_references = vec![tsp::ComponentExternalReference {
            component_identifier: PagesObjectId::Document.value(),
            object_identifier: Some(TABLE_INFO_OBJECT_ID),
            is_weak: None,
        }];
        components.push(calculation);
    }
    components.push(document);

    let metadata = tsp::PackageMetadata {
        last_object_identifier: LAST_PARAGRAPH_STYLE_OBJECT_ID,
        revision: Some(tsp::DocumentRevision {
            sequence_32: Some(INITIAL_REVISION_SEQUENCE),
            identifier: Some(identity.version_uuid().to_owned()),
            sequence_64: None,
        }),
        components,
        read_version: PACKAGE_VERSION.to_vec(),
        write_version: PACKAGE_VERSION.to_vec(),
        file_format_version: FILE_FORMAT_VERSION.to_vec(),
        save_token: Some(INITIAL_SAVE_TOKEN),
        ..Default::default()
    };
    Ok(Archive {
        objects: vec![object(
            PagesObjectId::PackageMetadata,
            PagesMessageType::PackageMetadata,
            metadata,
            &[],
        )?],
    })
}

fn annotation_author_storage_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            PagesObjectId::AnnotationAuthorStorage,
            PagesMessageType::AnnotationAuthorStorage,
            tsk::AnnotationAuthorStorageArchive::default(),
            &[],
        )?],
    })
}

fn component(identifier: PagesObjectId, locator: &str) -> tsp::ComponentInfo {
    component_raw(identifier.value(), locator, &PACKAGE_VERSION)
}

fn component_raw(identifier: u64, locator: &str, version: &[u32]) -> tsp::ComponentInfo {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        document_read_version: version.to_vec(),
        document_write_version: version.to_vec(),
        component_read_version: version.to_vec(),
        save_token: Some(INITIAL_SAVE_TOKEN),
        ..Default::default()
    }
}

fn object_uuid(identifier: PagesObjectId) -> tsp::ObjectUuidMapEntry {
    object_uuid_raw(identifier.value())
}

fn object_uuid_raw(identifier: u64) -> tsp::ObjectUuidMapEntry {
    let bytes = litchi_core::id::generate_guid_bytes();
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: u64::from_be_bytes([
                bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
                bytes[15],
            ]),
            upper: u64::from_be_bytes([
                bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
            ]),
        },
    }
}

fn insert_property_lists(
    package: &mut IWorkPackage,
    identity: &IWorkDocumentIdentity,
) -> Result<()> {
    let mut properties = plist::Dictionary::new();
    for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
        properties.insert(
            key.to_owned(),
            Value::String(identity.document_uuid().to_owned()),
        );
    }
    properties.insert(
        "fileFormatVersion".to_owned(),
        Value::String(FILE_FORMAT_VERSION_STRING.to_owned()),
    );
    properties.insert(
        "hasExternalReferenceOrMissingOrUnmaterializedRemoteData".to_owned(),
        Value::Boolean(false),
    );
    properties.insert("isMultiPage".to_owned(), Value::Boolean(false));
    properties.insert(
        "privateUUID".to_owned(),
        Value::String(identity.private_uuid().to_owned()),
    );
    properties.insert(
        "versionUUID".to_owned(),
        Value::String(identity.version_uuid().to_owned()),
    );
    properties.insert(
        "revision".to_owned(),
        Value::String(format!(
            "{INITIAL_REVISION_SEQUENCE}::{}",
            identity.version_uuid()
        )),
    );

    let mut encoded_properties = Vec::new();
    Value::Dictionary(properties)
        .to_writer_binary(&mut encoded_properties)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Pages properties: {error}"
            ))
        })?;
    package.insert_entry(PROPERTIES_ENTRY, encoded_properties)?;
    package.insert_entry(
        DOCUMENT_IDENTIFIER_ENTRY,
        identity.document_uuid().as_bytes().to_vec(),
    )?;

    let mut build_history = Vec::new();
    Value::Array(vec![Value::String(BUILD_HISTORY_VALUE.to_owned())])
        .to_writer_binary(&mut build_history)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Pages build history: {error}"
            ))
        })?;
    package.insert_entry(BUILD_HISTORY_ENTRY, build_history)?;
    Ok(())
}

fn object(
    identifier: PagesObjectId,
    message_type: PagesMessageType,
    message: impl Message,
    references: &[PagesObjectId],
) -> Result<ArchiveObject> {
    raw_object(
        identifier,
        message_type,
        message.encode_to_vec(),
        references,
    )
}

fn raw_object(
    identifier: PagesObjectId,
    message_type: PagesMessageType,
    data: Vec<u8>,
    references: &[PagesObjectId],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier.value(),
        vec![RawMessage {
            type_: message_type.value(),
            data,
        }],
    )?;
    let message_info = &mut object.archive_info.message_infos[0];
    message_info.versions = MESSAGE_VERSION.to_vec();
    message_info.object_references = references
        .iter()
        .copied()
        .map(PagesObjectId::value)
        .collect();
    Ok(object)
}

fn raw_object_with_id(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?;
    let message_info = &mut object.archive_info.message_infos[0];
    message_info.versions = MESSAGE_VERSION.to_vec();
    message_info.object_references = references.to_vec();
    Ok(object)
}

fn append_object_reference(object: &mut ArchiveObject, identifier: u64) {
    let references = &mut object.archive_info.message_infos[0].object_references;
    if !references.contains(&identifier) {
        references.push(identifier);
    }
}

fn validate_initial_table(table: &InitialPagesTable) -> Result<()> {
    if table.name.is_empty() || table.name.contains('\0') {
        return Err(crate::Error::InvalidFormat(
            "Pages table names must be non-empty and contain no NUL".to_owned(),
        ));
    }
    if table.rows == 0 || table.columns == 0 {
        return Err(crate::Error::InvalidFormat(
            "Pages tables must contain at least one row and one column".to_owned(),
        ));
    }
    u32::try_from(table.rows)
        .and_then(|_| u32::try_from(table.columns))
        .map_err(|_| crate::Error::InvalidFormat("Pages table dimensions exceed u32".to_owned()))?;
    Ok(())
}

fn install_initial_table_graph(
    package: &mut IWorkPackage,
    table: &InitialPagesTable,
    language: &str,
    locale: &str,
) -> Result<()> {
    let source = crate::numbers::NumbersDocumentBuilder::new()
        .table_name(&table.name)
        .table_dimensions(table.rows, table.columns)
        .language(language)
        .locale(locale)
        .build_package()?;

    let mut document_objects = source.archive(DOCUMENT_ARCHIVE_ENTRY)?.objects;
    document_objects.retain(|object| {
        object
            .archive_info
            .identifier
            .is_some_and(|identifier| TABLE_DOCUMENT_OBJECT_IDS.contains(&identifier))
    });
    let table_info = document_objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(TABLE_INFO_OBJECT_ID))
        .ok_or_else(|| {
            crate::Error::InvalidFormat("source-built table info is missing".to_owned())
        })?;
    let message = table_info
        .messages
        .iter_mut()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or_else(|| {
            crate::Error::InvalidFormat("source-built table info payload is missing".to_owned())
        })?;
    let mut info = tst::TableInfoArchive::decode(message.data.as_slice())?;
    info.super_.parent = Some(reference(PagesObjectId::Body));
    info.super_.geometry = Some(tsd::GeometryArchive {
        position: Some(tsp::Point {
            x: DEFAULT_TABLE_X_POINTS,
            y: DEFAULT_TABLE_Y_POINTS,
        }),
        size: Some(tsp::Size {
            width: DEFAULT_TABLE_WIDTH_POINTS,
            height: DEFAULT_TABLE_ROW_HEIGHT_POINTS * table.rows as f32,
        }),
        flags: Some(TABLE_DRAWABLE_FLAGS),
        angle: Some(TABLE_ROTATION_DEGREES),
    });
    message.data = info.encode_to_vec();
    let references = &mut table_info.archive_info.message_infos[0].object_references;
    references.retain(|identifier| *identifier != SOURCE_TABLE_SHEET_OBJECT_ID);
    if !references.contains(&PagesObjectId::Body.value()) {
        references.push(PagesObjectId::Body.value());
    }

    package.update_archive(DOCUMENT_ARCHIVE_ENTRY, |archive| {
        for object in document_objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;

    let mut style_objects = source.archive(STYLESHEET_ARCHIVE_ENTRY)?.objects;
    style_objects.retain(|object| {
        object
            .archive_info
            .identifier
            .is_some_and(|identifier| TABLE_STYLE_OBJECT_IDS.contains(&identifier))
    });
    package.update_archive(STYLESHEET_ARCHIVE_ENTRY, |archive| {
        for object in style_objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;

    let mut calculation = source.archive("Index/CalculationEngine.iwa")?;
    calculation.objects.retain(|object| {
        matches!(
            object.archive_info.identifier,
            Some(TABLE_CALCULATION_ENGINE_OBJECT_ID | TABLE_FORMULA_OWNER_OBJECT_ID)
        )
    });
    package.replace_archive("Index/CalculationEngine.iwa", &calculation)?;
    Ok(())
}

fn object_table(identifier: Option<PagesObjectId>) -> tswp::ObjectAttributeTable {
    tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: identifier.map(reference),
        }],
    }
}

fn header_footer_storage() -> tswp::StorageArchive {
    tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Header as i32),
        style_sheet: Some(reference(PagesObjectId::Stylesheet)),
        in_document: Some(true),
        table_para_style: Some(object_table(Some(PagesObjectId::ParagraphStyle))),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_table(Some(PagesObjectId::ListStyle))),
        table_para_starts: Some(zero_para_data()),
        table_para_bidi: Some(zero_para_data()),
        ..Default::default()
    }
}

fn zero_para_data() -> tswp::ParaDataAttributeTable {
    tswp::ParaDataAttributeTable {
        entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        }],
    }
}

fn style(name: &str, identifier: &str) -> tss::StyleArchive {
    tss::StyleArchive {
        name: Some(name.to_owned()),
        style_identifier: Some(identifier.to_owned()),
        stylesheet: Some(reference(PagesObjectId::Stylesheet)),
        ..Default::default()
    }
}

fn shape_style_archive(
    name: &str,
    identifier: &str,
    fill: tsd::FillArchive,
    paragraph_style: Option<PagesObjectId>,
) -> tswp::ShapeStyleArchive {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: style(name, identifier),
            override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
            shape_properties: Some(tsd::ShapeStylePropertiesArchive {
                fill: Some(fill),
                opacity: Some(1.0),
                ..Default::default()
            }),
        },
        override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
        shape_properties: Some(tswp::ShapeStylePropertiesArchive {
            shrink_to_fit: Some(false),
            vertical_alignment: Some(
                tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop as i32,
            ),
            columns: Some(tswp::ColumnsArchive {
                equal_columns: Some(tswp::columns_archive::EqualColumnsArchive {
                    count: Some(TEXT_BOX_COLUMN_COUNT),
                    gap: None,
                }),
                non_equal_columns: None,
            }),
            padding: Some(tswp::PaddingArchive {
                left: Some(TEXT_BOX_PADDING_POINTS),
                top: Some(TEXT_BOX_PADDING_POINTS),
                right: Some(TEXT_BOX_PADDING_POINTS),
                bottom: Some(TEXT_BOX_PADDING_POINTS),
            }),
            default_text_preset_index: Some(DEFAULT_TEXT_PRESET_INDEX),
            paragraph_style: paragraph_style.map(reference),
            vertical_text_40: Some(false),
            ..Default::default()
        }),
    }
}

fn media_style_archive(name: &str, identifier: &str) -> tsd::MediaStyleArchive {
    tsd::MediaStyleArchive {
        super_: style(name, identifier),
        override_count: Some(DEFAULT_STYLE_OVERRIDE_COUNT),
        ..Default::default()
    }
}

fn repeated<T>(count: usize, make: impl Fn() -> T) -> Vec<T> {
    std::iter::repeat_with(make).take(count).collect()
}

fn repeated_reference(count: usize, identifier: PagesObjectId) -> Vec<tsp::Reference> {
    repeated(count, || reference(identifier))
}

fn reference(identifier: PagesObjectId) -> tsp::Reference {
    raw_reference(identifier.value())
}

fn raw_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    }
}

fn black() -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.0),
        g: Some(0.0),
        b: Some(0.0),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(1.0),
        ..Default::default()
    }
}

fn solid_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        color: Some(black()),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    #[test]
    fn generated_theme_exposes_distinct_canonical_list_presets() {
        let package = PagesDocumentBuilder::new().build_package().unwrap();
        let document = package.archive(DOCUMENT_ARCHIVE_ENTRY).unwrap();
        let theme = document.object(PagesObjectId::Theme.value()).unwrap();
        let theme = IWorkThemeArchive::decode(&theme.messages[0].data).unwrap();
        let preset_ids = theme
            .extensions
            .text
            .unwrap()
            .list_style_presets
            .into_iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        assert_eq!(
            preset_ids,
            [
                PagesObjectId::ListStyle.value(),
                PagesObjectId::BulletListStyle.value(),
                PagesObjectId::NumberedListStyle.value(),
            ]
        );
        let stylesheet = package.archive(STYLESHEET_ARCHIVE_ENTRY).unwrap();
        assert!(
            preset_ids
                .iter()
                .all(|identifier| stylesheet.object(*identifier).is_some())
        );
        assert_eq!(
            crate::package_metadata::package_last_object_identifier(&package).unwrap(),
            Some(LAST_PARAGRAPH_STYLE_OBJECT_ID)
        );
    }

    #[test]
    fn generated_theme_exposes_unique_native_paragraph_style_catalog() {
        let package = PagesDocumentBuilder::new().build_package().unwrap();
        let styles = crate::text::paragraph_alignment::native::named_paragraph_styles(
            &package,
            PagesObjectId::ParagraphStyle.value(),
        )
        .unwrap();
        assert_eq!(
            styles
                .iter()
                .map(crate::text::NamedParagraphStyle::name)
                .collect::<Vec<_>>(),
            [
                "Title",
                "Subtitle",
                "Heading",
                "Heading 2",
                "Heading 3",
                "Heading Red",
                "Body",
                "Caption",
                "Header & Footer",
                "Footnote",
                "Label",
                "Label Dark",
            ]
        );
        assert_eq!(
            styles
                .iter()
                .map(|style| litchi_iwa_text::paragraph::style::raw::native_id(style.id()))
                .collect::<Vec<_>>(),
            PagesParagraphStylePreset::ALL
                .into_iter()
                .map(PagesParagraphStylePreset::object_id)
                .map(PagesObjectId::value)
                .collect::<Vec<_>>()
        );

        let stylesheet = package.archive(STYLESHEET_ARCHIVE_ENTRY).unwrap();
        let stylesheet = tss::StylesheetArchive::decode(
            stylesheet
                .object(PagesObjectId::Stylesheet.value())
                .unwrap()
                .messages
                .first()
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert!(PagesParagraphStylePreset::ALL.into_iter().all(|preset| {
            stylesheet.identifier_to_style_map.iter().any(|entry| {
                entry.identifier == preset.style_identifier()
                    && entry.style.identifier == preset.object_id().value()
            })
        }));

        let document = package.archive(DOCUMENT_ARCHIVE_ENTRY).unwrap();
        let theme_references = &document
            .object(PagesObjectId::Theme.value())
            .unwrap()
            .archive_info
            .message_infos[0]
            .object_references;
        assert!(
            PagesParagraphStylePreset::ALL
                .into_iter()
                .map(PagesParagraphStylePreset::object_id)
                .map(PagesObjectId::value)
                .all(|identifier| theme_references.contains(&identifier))
        );
    }

    const EXPECTED_ENTRIES: [&str; 7] = [
        DOCUMENT_ARCHIVE_ENTRY,
        STYLESHEET_ARCHIVE_ENTRY,
        ANNOTATION_ARCHIVE_ENTRY,
        METADATA_ARCHIVE_ENTRY,
        PROPERTIES_ENTRY,
        DOCUMENT_IDENTIFIER_ENTRY,
        BUILD_HISTORY_ENTRY,
    ];

    #[test]
    fn creates_and_edits_pages_entirely_from_typed_objects() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Created from scratch")
            .build()
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), "Created from scratch");

        editor.set_body_text("Updated through CRUD").unwrap();
        let encoded = editor.to_bytes().unwrap();
        let reopened = PagesEditor::from_bytes(&encoded).unwrap();
        assert_eq!(reopened.body_text().unwrap(), "Updated through CRUD");
        assert_eq!(reopened.sections().len(), 1);
    }

    #[test]
    fn repeated_body_table_calls_append_independent_tables() {
        let body = "Tables 🙂\n";
        let editor = PagesDocumentBuilder::new()
            .body_text(body)
            .body_table("First", 2, 3)
            .body_table("Second", 4, 1)
            .build()
            .unwrap();
        let tables = editor.tables().unwrap();
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0].name, "First");
        assert_eq!((tables[0].rows, tables[0].columns), (2, 3));
        assert_eq!(
            tables[0].anchor_character_index,
            body.encode_utf16().count()
        );
        assert_eq!(tables[1].name, "Second");
        assert_eq!((tables[1].rows, tables[1].columns), (4, 1));
        assert_eq!(
            tables[1].anchor_character_index,
            body.encode_utf16().count() + 1
        );
        assert_eq!(
            editor
                .table(tables[0].model_object_id)
                .unwrap()
                .cell_count(),
            0
        );
        assert_eq!(
            editor
                .table(tables[1].model_object_id)
                .unwrap()
                .cell_count(),
            0
        );
    }

    #[test]
    fn blank_package_contains_only_generated_required_entries() {
        let package = PagesDocumentBuilder::new().build_package().unwrap();
        assert_eq!(package.entry_names().collect::<Vec<_>>(), EXPECTED_ENTRIES);
        assert_eq!(package.len(), EXPECTED_ENTRIES.len());
        assert!(package.entry_names().all(|name| !name.starts_with("Data/")));
        assert!(
            package
                .entry_names()
                .all(|name| !name.starts_with("preview"))
        );
        assert_eq!(
            PagesEditor::from_package(package)
                .unwrap()
                .body_text()
                .unwrap(),
            ""
        );
    }

    #[test]
    fn scratch_package_has_native_save_scaffolding() {
        let package = PagesDocumentBuilder::new().build_package().unwrap();
        assert!(package.iwa_entry_names().all(|member| {
            package
                .archive(member)
                .unwrap()
                .object(COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID)
                .is_none()
        }));

        let document = package.archive(DOCUMENT_ARCHIVE_ENTRY).unwrap();
        let document_message = tp::DocumentArchive::decode(
            document
                .object(PagesObjectId::Document.value())
                .unwrap()
                .messages
                .first()
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert_eq!(
            document_message
                .super_
                .super_
                .annotation_author_storage
                .unwrap()
                .identifier,
            PagesObjectId::AnnotationAuthorStorage.value()
        );
        let annotation = package.archive(ANNOTATION_ARCHIVE_ENTRY).unwrap();
        let annotation = annotation
            .object(PagesObjectId::AnnotationAuthorStorage.value())
            .unwrap();
        assert_eq!(
            annotation.messages[0].type_,
            PagesMessageType::AnnotationAuthorStorage.value()
        );
        assert!(
            tsk::AnnotationAuthorStorageArchive::decode(annotation.messages[0].data.as_slice())
                .unwrap()
                .annotation_author
                .is_empty()
        );
        let template = document
            .object(PagesObjectId::SectionTemplate.value())
            .unwrap();
        let template =
            tp::SectionTemplateArchive::decode(template.messages.first().unwrap().data.as_slice())
                .unwrap();
        assert_eq!(template.headers.len(), 3);
        assert_eq!(template.footers.len(), 3);
        assert!(
            template
                .headers
                .iter()
                .chain(&template.footers)
                .all(|reference| document.object(reference.identifier).is_some())
        );

        let stylesheet = package.archive(STYLESHEET_ARCHIVE_ENTRY).unwrap();
        let stylesheet = tss::StylesheetArchive::decode(
            stylesheet
                .object(PagesObjectId::Stylesheet.value())
                .unwrap()
                .messages
                .first()
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert!(stylesheet.identifier_to_style_map.iter().any(|entry| {
            entry.identifier == CHARACTER_STYLE_IDENTIFIER
                && entry.style.identifier == PagesObjectId::CharacterStyle.value()
        }));
        assert!(stylesheet.identifier_to_style_map.iter().any(|entry| {
            entry.identifier == TEXT_BOX_STYLE_IDENTIFIER
                && entry.style.identifier == PagesObjectId::TextBoxStyle.value()
        }));
    }

    #[test]
    fn new_documents_receive_independent_package_and_object_identities() {
        let first = PagesDocumentBuilder::new().build_package().unwrap();
        let second = PagesDocumentBuilder::new().build_package().unwrap();
        assert_ne!(
            first.entry(DOCUMENT_IDENTIFIER_ENTRY),
            second.entry(DOCUMENT_IDENTIFIER_ENTRY)
        );

        let first_properties =
            Value::from_reader(Cursor::new(first.entry(PROPERTIES_ENTRY).unwrap())).unwrap();
        let properties = first_properties.as_dictionary().unwrap();
        let document_uuid = properties["documentUUID"].as_string().unwrap();
        assert_ne!(
            document_uuid,
            properties["versionUUID"].as_string().unwrap()
        );
        assert_ne!(
            document_uuid,
            properties["privateUUID"].as_string().unwrap()
        );

        let first_metadata = first.archive(METADATA_ARCHIVE_ENTRY).unwrap();
        let second_metadata = second.archive(METADATA_ARCHIVE_ENTRY).unwrap();
        let first_message = &first_metadata.objects[0].messages[0].data;
        let second_message = &second_metadata.objects[0].messages[0].data;
        let first = tsp::PackageMetadata::decode(first_message.as_slice()).unwrap();
        let second = tsp::PackageMetadata::decode(second_message.as_slice()).unwrap();
        let first_uuids = first
            .components
            .iter()
            .flat_map(|component| &component.object_uuid_map_entries)
            .map(|entry| (entry.uuid.upper, entry.uuid.lower))
            .collect::<Vec<_>>();
        let second_uuids = second
            .components
            .iter()
            .flat_map(|component| &component.object_uuid_map_entries)
            .map(|entry| (entry.uuid.upper, entry.uuid.lower))
            .collect::<Vec<_>>();
        assert_eq!(
            first_uuids.len(),
            STYLESHEET_OBJECTS.len() + DOCUMENT_OBJECTS.len()
        );
        assert_eq!(
            first_uuids
                .iter()
                .copied()
                .collect::<std::collections::HashSet<_>>()
                .len(),
            first_uuids.len()
        );
        assert_ne!(first_uuids, second_uuids);
    }

    #[test]
    fn rejects_empty_language_and_locale() {
        assert!(PagesDocumentBuilder::new().language(" ").build().is_err());
        assert!(PagesDocumentBuilder::new().locale("").build().is_err());
    }
}
