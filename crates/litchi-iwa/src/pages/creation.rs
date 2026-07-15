//! Construction of independent Pages packages without bundled templates.

use plist::Value;
use prost::Message;

use super::editor::{PagesEditor, PagesSectionPageNumbering, PagesSectionStart};
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::identity::IWorkDocumentIdentity;
use crate::protobuf::{tp, tsa, tsd, tsk, tsp, tss, tswp};
use crate::{IWorkPackage, IWorkThemeArchive, IWorkThemeExtensions, Result};

const DOCUMENT_ARCHIVE_ENTRY: &str = "Index/Document.iwa";
const STYLESHEET_ARCHIVE_ENTRY: &str = "Index/DocumentStylesheet.iwa";
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
const COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID: u64 = 3;

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
const LIST_STYLE_PRESET_COUNT: usize = 1;
const TOC_ENTRY_STYLE_PRESET_COUNT: usize = 1;
const TOC_SETTINGS_PRESET_COUNT: usize = 1;
const CHARACTER_STYLE_PRESET_COUNT: usize = 1;
const PARAGRAPH_STYLE_PRESET_COUNT: usize = 1;
const DROP_CAP_STYLE_PRESET_COUNT: usize = 1;
const CAPTION_STYLE_PRESET_COUNT: usize = 1;
const SVG_IMPORT_STYLE_PRESET_COUNT: usize = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u64)]
enum PagesObjectId {
    Document = 1,
    PackageMetadata = 2,
    // `COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID` is reserved for TSCKDocumentSupport.
    Stylesheet = COLLABORATION_DOCUMENT_SUPPORT_OBJECT_ID + 1,
    Theme = 5,
    Body = 6,
    Settings = 7,
    Section = 8,
    SectionTemplate = 9,
    ListStyle = 10,
    ParagraphStyle = 11,
    CharacterStyle = 12,
    LineStyle = 13,
    ShapeStyle = 14,
    TextBoxStyle = 15,
    ImageStyle = 16,
    MovieStyle = 17,
    DrawingLineStyle = 18,
    TocEntryStyle = 19,
    DropCapStyle = 20,
    BaseColumnStyle = 21,
    ColumnStyle = 22,
    TocSettings = 23,
    CaptionStyle = 24,
    SvgImportStyle = 25,
    HeaderPrimary = 26,
    HeaderEven = 27,
    HeaderFirst = 28,
    FooterPrimary = 29,
    FooterEven = 30,
    FooterFirst = 31,
}

impl PagesObjectId {
    const fn value(self) -> u64 {
        self as u64
    }
}

const STYLESHEET_OBJECTS: [PagesObjectId; 16] = [
    PagesObjectId::Stylesheet,
    PagesObjectId::ListStyle,
    PagesObjectId::ParagraphStyle,
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

const IDENTIFIED_STYLES: [(PagesObjectId, &str); 14] = [
    (PagesObjectId::ListStyle, LIST_STYLE_IDENTIFIER),
    (PagesObjectId::ParagraphStyle, PARAGRAPH_STYLE_IDENTIFIER),
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
}

impl Default for PagesDocumentBuilder {
    fn default() -> Self {
        Self {
            body_text: String::new(),
            language: DEFAULT_LANGUAGE.to_owned(),
            locale: DEFAULT_LOCALE.to_owned(),
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

    /// Build a mutable editor for the new document.
    pub fn build(self) -> Result<PagesEditor> {
        PagesEditor::from_package(self.build_package()?)
    }

    /// Build the underlying package for lower-level IWA manipulation.
    pub fn build_package(self) -> Result<IWorkPackage> {
        if self.language.trim().is_empty() {
            return Err(crate::Error::InvalidFormat(
                "Pages document language cannot be empty".to_owned(),
            ));
        }
        if self.locale.trim().is_empty() {
            return Err(crate::Error::InvalidFormat(
                "Pages document locale cannot be empty".to_owned(),
            ));
        }

        let identity = IWorkDocumentIdentity::generate();
        let mut package = IWorkPackage::new();
        package.replace_archive(
            DOCUMENT_ARCHIVE_ENTRY,
            &document_archive(&self.body_text, &self.language, &self.locale)?,
        )?;
        package.replace_archive(STYLESHEET_ARCHIVE_ENTRY, &stylesheet_archive()?)?;
        package.replace_archive(METADATA_ARCHIVE_ENTRY, &metadata_archive(&identity)?)?;
        insert_property_lists(&mut package, &identity)?;
        Ok(package)
    }
}

fn document_archive(body_text: &str, language: &str, locale: &str) -> Result<Archive> {
    let document = tp::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive {
                locale_identifier: Some(locale.to_owned()),
                creation_locale_identifier: Some(locale.to_owned()),
                prevent_image_conversion_on_open: Some(true),
                has_user_defined_locale: Some(false),
                ..Default::default()
            },
            document_language: Some(language.to_owned()),
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
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        style_sheet: Some(reference(PagesObjectId::Stylesheet)),
        text: vec![body_text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_table(Some(PagesObjectId::ParagraphStyle))),
        table_para_data: Some(zero_para_data()),
        table_list_style: Some(object_table(Some(PagesObjectId::ListStyle))),
        table_layout_style: Some(object_table(Some(PagesObjectId::ColumnStyle))),
        table_para_starts: Some(zero_para_data()),
        table_section: Some(object_table(Some(PagesObjectId::Section))),
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
                list_style_presets: repeated_reference(
                    LIST_STYLE_PRESET_COUNT,
                    PagesObjectId::ListStyle,
                ),
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
                paragraph_style_presets: repeated_reference(
                    PARAGRAPH_STYLE_PRESET_COUNT,
                    PagesObjectId::ParagraphStyle,
                ),
                dropcap_style_presets: repeated_reference(
                    DROP_CAP_STYLE_PRESET_COUNT,
                    PagesObjectId::DropCapStyle,
                ),
                ..Default::default()
            }),
            chart: None,
            table: None,
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

    Ok(Archive {
        objects: vec![
            object(
                PagesObjectId::Document,
                PagesMessageType::Document,
                document,
                &[
                    PagesObjectId::Stylesheet,
                    PagesObjectId::Body,
                    PagesObjectId::Theme,
                    PagesObjectId::Settings,
                ],
            )?,
            raw_object(
                PagesObjectId::Theme,
                PagesMessageType::Theme,
                theme.encode()?,
                &[
                    PagesObjectId::Stylesheet,
                    PagesObjectId::ListStyle,
                    PagesObjectId::ParagraphStyle,
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
                ],
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
                    section_start_kind: Some(PagesSectionStart::NextPage.as_raw()),
                    section_page_number_kind: Some(
                        PagesSectionPageNumbering::ContinueFromPrevious.as_raw(),
                    ),
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
        ],
    })
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
            PagesObjectId::ParagraphStyle,
            PagesMessageType::ParagraphStyle,
            tswp::ParagraphStyleArchive {
                super_: style("Body", PARAGRAPH_STYLE_IDENTIFIER),
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
    Ok(Archive { objects })
}

fn metadata_archive(identity: &IWorkDocumentIdentity) -> Result<Archive> {
    let mut stylesheet = component(PagesObjectId::Stylesheet, "DocumentStylesheet");
    stylesheet.object_uuid_map_entries = STYLESHEET_OBJECTS
        .iter()
        .copied()
        .map(object_uuid)
        .collect();

    let mut document = component(PagesObjectId::Document, "Document");
    document.object_uuid_map_entries = DOCUMENT_OBJECTS.iter().copied().map(object_uuid).collect();
    document.external_references = std::iter::once(None)
        .chain(STYLESHEET_OBJECTS.iter().copied().map(Some))
        .map(|object_identifier| tsp::ComponentExternalReference {
            component_identifier: PagesObjectId::Stylesheet.value(),
            object_identifier: object_identifier.map(PagesObjectId::value),
            is_weak: None,
        })
        .collect();

    let metadata = tsp::PackageMetadata {
        last_object_identifier: PagesObjectId::FooterFirst.value(),
        revision: Some(tsp::DocumentRevision {
            sequence_32: Some(INITIAL_REVISION_SEQUENCE),
            identifier: Some(identity.version_uuid().to_owned()),
            sequence_64: None,
        }),
        components: vec![stylesheet, document],
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

fn component(identifier: PagesObjectId, locator: &str) -> tsp::ComponentInfo {
    tsp::ComponentInfo {
        identifier: identifier.value(),
        preferred_locator: locator.to_owned(),
        document_read_version: PACKAGE_VERSION.to_vec(),
        document_write_version: PACKAGE_VERSION.to_vec(),
        component_read_version: PACKAGE_VERSION.to_vec(),
        save_token: Some(INITIAL_SAVE_TOKEN),
        ..Default::default()
    }
}

fn object_uuid(identifier: PagesObjectId) -> tsp::ObjectUuidMapEntry {
    let bytes = litchi_core::id::generate_guid_bytes();
    tsp::ObjectUuidMapEntry {
        identifier: identifier.value(),
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
    tsp::Reference {
        identifier: identifier.value(),
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

    const EXPECTED_ENTRIES: [&str; 6] = [
        DOCUMENT_ARCHIVE_ENTRY,
        STYLESHEET_ARCHIVE_ENTRY,
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
