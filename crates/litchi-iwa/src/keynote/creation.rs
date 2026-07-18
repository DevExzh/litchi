//! Construction of independent Keynote packages without bundled templates.

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::identity::IWorkDocumentIdentity;
use crate::package_metadata::{add_component_external_reference, add_component_object_uuids};
use crate::protobuf::{kn, tsa, tsce, tsd, tsk, tsp, tss, tst, tswp};
use crate::wire::{parse_wire_fields, patch_length_delimited_field};
use crate::{IWorkPackage, IWorkThemeArchive, IWorkThemeExtensions, Result};
use plist::Value;
use prost::Message;

use super::editor::KeynoteEditor;

mod slide_number;

const DOCUMENT_ARCHIVE_ENTRY: &str = "Index/Document.iwa";
const TEMPLATE_ARCHIVE_ENTRY: &str = "Index/TemplateSlide-8.iwa";
const SLIDE_ARCHIVE_ENTRY: &str = "Index/Slide-14.iwa";
const CALCULATION_ARCHIVE_ENTRY: &str = "Index/CalculationEngine.iwa";
const STYLESHEET_ARCHIVE_ENTRY: &str = "Index/DocumentStylesheet.iwa";
const VIEW_STATE_ARCHIVE_ENTRY: &str = "Index/ViewState.iwa";
const ANNOTATION_ARCHIVE_ENTRY: &str = "Index/AnnotationAuthorStorage.iwa";
const DOCUMENT_METADATA_ARCHIVE_ENTRY: &str = "Index/DocumentMetadata.iwa";
const PACKAGE_METADATA_ARCHIVE_ENTRY: &str = "Index/Metadata.iwa";

const DEFAULT_LANGUAGE: &str = "en";
const DEFAULT_LOCALE: &str = "en_US";
const DEFAULT_TITLE: &str = "Presentation Title";
const DEFAULT_SUBTITLE: &str = "Presentation Subtitle";
const DEFAULT_LAYOUT_NAME: &str = "Title & Subtitle";
const DEFAULT_WIDTH: f32 = 1_920.0;
const DEFAULT_HEIGHT: f32 = 1_080.0;
const TEXT_WRAP_TYPE: u32 = 4;
const TEXT_WRAP_DIRECTION: u32 = 2;
const TEXT_WRAP_FIT_TYPE: u32 = 1;
const TEXT_WRAP_MARGIN: f32 = 12.0;
const TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;

const DOCUMENT: u64 = 1;
const METADATA: u64 = 2;
// Keynote reserves identifier 3 for its lazily-created TSCKDocumentSupport root.
// Keeping generated objects above that slot allows an opened document to save.
// Keep the native Numbers table-template range (9..=40) free. A hidden,
// source-built table scaffold occupies those identifiers so a blank Keynote
// presentation can create its first table without copying an Apple template.
const THEME: u64 = 100;
const SHOW: u64 = 101;
const LIVE_NODE: u64 = 102;
const TEMPLATE_NODE: u64 = 103;
const TEMPLATE_SLIDE: u64 = 104;
const TEMPLATE_TITLE: u64 = 105;
const TEMPLATE_BODY: u64 = 106;
const TEMPLATE_TITLE_STORAGE: u64 = 107;
const TEMPLATE_BODY_STORAGE: u64 = 108;
const TEMPLATE_GUIDES: u64 = 109;
const LIVE_SLIDE: u64 = 110;
const LIVE_TITLE: u64 = 111;
const LIVE_BODY: u64 = 112;
const LIVE_TITLE_STORAGE: u64 = 113;
const LIVE_BODY_STORAGE: u64 = 114;
const LIVE_NOTE: u64 = 115;
const LIVE_NOTE_STORAGE: u64 = 116;
const LIVE_GUIDES: u64 = 117;
const SLIDE_STYLE: u64 = 118;
const LIST_STYLE: u64 = 119;
const PARAGRAPH_STYLE: u64 = 120;
const CHARACTER_STYLE: u64 = 121;
const SHAPE_STYLE: u64 = 122;
const MEDIA_STYLE: u64 = 123;
const DROP_CAP_STYLE: u64 = 124;
const TEMPLATE_SLIDE_NUMBER: u64 = 125;
const LIVE_SLIDE_NUMBER: u64 = 126;
const CALCULATION_ENGINE: u64 = 127;
const FUNCTION_BROWSER_STATE: u64 = 128;
const CUSTOM_FORMAT_LIST: u64 = 129;
const VIEW_STATE: u64 = 130;
const UI_STATE: u64 = 131;
const ANNOTATION_AUTHOR_STORAGE: u64 = 132;
const DOCUMENT_METADATA: u64 = 133;
const SOUNDTRACK: u64 = 134;
const LIVE_VIDEO_COLLECTION: u64 = 135;
const DEFAULT_LIVE_VIDEO_SOURCE: u64 = 136;
const STYLESHEET: u64 = 137;

const TABLE_INFO_TEMPLATE: u64 = 9;
const TABLE_MODEL_TEMPLATE: u64 = 10;
const TABLE_PRESET_TEMPLATE: u64 = 20;
const TABLE_TEMPLATE_DOCUMENT_OBJECTS: &[u64] = &[
    TABLE_INFO_TEMPLATE,
    TABLE_MODEL_TEMPLATE,
    TABLE_PRESET_TEMPLATE,
    22,
    23,
    24,
    25,
    26,
    27,
    28,
    29,
    30,
    32,
];
const TABLE_TEMPLATE_STYLE_OBJECTS: &[u64] = &[11, 12, 13, 14, 15, 16, 17, 18, 19, 40];
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;

const STYLESHEET_OBJECTS: &[u64] = &[
    STYLESHEET,
    SLIDE_STYLE,
    LIST_STYLE,
    PARAGRAPH_STYLE,
    CHARACTER_STYLE,
    SHAPE_STYLE,
    MEDIA_STYLE,
    DROP_CAP_STYLE,
];
const DOCUMENT_OBJECTS: &[u64] = &[
    DOCUMENT,
    THEME,
    SHOW,
    LIVE_NODE,
    TEMPLATE_NODE,
    SOUNDTRACK,
    LIVE_VIDEO_COLLECTION,
    DEFAULT_LIVE_VIDEO_SOURCE,
    FUNCTION_BROWSER_STATE,
    CUSTOM_FORMAT_LIST,
];
const TEMPLATE_OBJECTS: &[u64] = &[
    TEMPLATE_SLIDE,
    TEMPLATE_TITLE,
    TEMPLATE_BODY,
    TEMPLATE_TITLE_STORAGE,
    TEMPLATE_BODY_STORAGE,
    TEMPLATE_GUIDES,
    TEMPLATE_SLIDE_NUMBER,
];
const SLIDE_OBJECTS: &[u64] = &[
    LIVE_SLIDE,
    LIVE_TITLE,
    LIVE_BODY,
    LIVE_TITLE_STORAGE,
    LIVE_BODY_STORAGE,
    LIVE_NOTE,
    LIVE_NOTE_STORAGE,
    LIVE_GUIDES,
    LIVE_SLIDE_NUMBER,
];

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum KeynoteMessageType {
    Document = 1,
    Show = 2,
    UiState = 3,
    SlideNode = 4,
    Slide = 5,
    Placeholder = 7,
    SlideStyle = 9,
    Theme = 10,
    Note = 15,
    Soundtrack = 21,
    LiveVideoSource = 184,
    LiveVideoCollection = 185,
    ViewState = 210,
    AnnotationAuthorStorage = 213,
    CustomFormatList = 222,
    Stylesheet = 401,
    FunctionBrowserState = 601,
    Storage = 2_001,
    CharacterStyle = 2_021,
    ParagraphStyle = 2_022,
    ListStyle = 2_023,
    ShapeStyle = 2_025,
    MediaStyle = 3_016,
    GuideStorage = 3_047,
    CalculationEngine = 4_000,
    DropCapStyle = 10_024,
    PackageMetadata = 11_006,
    DocumentMetadata = 11_011,
}

impl KeynoteMessageType {
    const fn value(self) -> u32 {
        self as u32
    }
}

/// Builder for a new, independent Keynote presentation.
///
/// Every archive, relationship, and identity is encoded from typed values. No
/// Apple presentation, blank package, or other prebuilt template is copied.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteDocumentBuilder {
    title: String,
    subtitle: String,
    presenter_notes: String,
    language: String,
    locale: String,
    width: f32,
    height: f32,
    slide_number_visible: bool,
}

impl Default for KeynoteDocumentBuilder {
    fn default() -> Self {
        Self {
            title: DEFAULT_TITLE.to_owned(),
            subtitle: DEFAULT_SUBTITLE.to_owned(),
            presenter_notes: String::new(),
            language: DEFAULT_LANGUAGE.to_owned(),
            locale: DEFAULT_LOCALE.to_owned(),
            width: DEFAULT_WIDTH,
            height: DEFAULT_HEIGHT,
            slide_number_visible: false,
        }
    }
}

impl KeynoteDocumentBuilder {
    /// Start a presentation containing one title-and-subtitle slide.
    pub fn new() -> Self {
        Self::default()
    }

    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = title.into();
        self
    }

    pub fn subtitle(mut self, subtitle: impl Into<String>) -> Self {
        self.subtitle = subtitle.into();
        self
    }

    pub fn presenter_notes(mut self, notes: impl Into<String>) -> Self {
        self.presenter_notes = notes.into();
        self
    }

    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = language.into();
        self
    }

    pub fn locale(mut self, locale: impl Into<String>) -> Self {
        self.locale = locale.into();
        self
    }

    /// Set the slide dimensions in points.
    pub fn slide_size(mut self, width: f32, height: f32) -> Self {
        self.width = width;
        self.height = height;
        self
    }

    /// Show the native slide-number placeholder on the initial slide.
    ///
    /// The placeholder graph is always created so visibility can be toggled
    /// later without consulting an Apple template.
    pub const fn slide_number_visible(mut self, visible: bool) -> Self {
        self.slide_number_visible = visible;
        self
    }

    /// Build a mutable editor for the generated presentation.
    pub fn build(self) -> Result<KeynoteEditor> {
        KeynoteEditor::from_package(self.build_package()?)
    }

    /// Build the underlying package for lower-level IWA manipulation.
    pub fn build_package(self) -> Result<IWorkPackage> {
        self.validate()?;
        let identity = IWorkDocumentIdentity::generate();
        let template_identifier = fresh_tsp_uuid();
        let mut package = IWorkPackage::new();
        package.replace_archive(
            DOCUMENT_ARCHIVE_ENTRY,
            &document_archive(&self, template_identifier)?,
        )?;
        package.replace_archive(TEMPLATE_ARCHIVE_ENTRY, &template_archive(&self)?)?;
        package.replace_archive(SLIDE_ARCHIVE_ENTRY, &slide_archive(&self)?)?;
        package.replace_archive(STYLESHEET_ARCHIVE_ENTRY, &stylesheet_archive()?)?;
        package.replace_archive(
            CALCULATION_ARCHIVE_ENTRY,
            &calculation_archive(&self.locale)?,
        )?;
        package.replace_archive(VIEW_STATE_ARCHIVE_ENTRY, &view_state_archive()?)?;
        package.replace_archive(
            ANNOTATION_ARCHIVE_ENTRY,
            &annotation_author_storage_archive()?,
        )?;
        package.replace_archive(
            DOCUMENT_METADATA_ARCHIVE_ENTRY,
            &document_metadata_archive()?,
        )?;
        package.replace_archive(
            PACKAGE_METADATA_ARCHIVE_ENTRY,
            &metadata_archive(&identity)?,
        )?;
        install_table_template(&mut package, &self.language, &self.locale)?;
        add_plists(&mut package, &identity)?;
        Ok(package)
    }

    fn validate(&self) -> Result<()> {
        crate::text::TextLanguageTag::new(self.language.as_str())?;
        if self.locale.trim().is_empty() {
            return Err(crate::Error::InvalidFormat(
                "Keynote document locale cannot be empty".to_owned(),
            ));
        }
        if !self.width.is_finite()
            || !self.height.is_finite()
            || self.width <= 0.0
            || self.height <= 0.0
        {
            return Err(crate::Error::InvalidFormat(
                "Keynote slide dimensions must be finite and positive".to_owned(),
            ));
        }
        Ok(())
    }
}

/// Install the native table style/storage scaffold used to create the first
/// slide table. The scaffold is generated by the crate's typed Numbers builder,
/// has no slide parent, and is therefore never rendered as presentation
/// content. Keynote and Numbers share these TST/TSCE archive types.
fn install_table_template(package: &mut IWorkPackage, language: &str, locale: &str) -> Result<()> {
    let source = crate::numbers::NumbersDocumentBuilder::new()
        .table_name("Table Template")
        .table_dimensions(5, 4)
        .language(language)
        .locale(locale)
        .build_package()?;

    let mut table_objects = source.archive(DOCUMENT_ARCHIVE_ENTRY)?.objects;
    table_objects.retain(|object| {
        object
            .archive_info
            .identifier
            .is_some_and(|identifier| TABLE_TEMPLATE_DOCUMENT_OBJECTS.contains(&identifier))
    });
    let table_info = table_objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(TABLE_INFO_TEMPLATE))
        .ok_or_else(|| {
            crate::Error::InvalidFormat("source-built Keynote table template is missing".to_owned())
        })?;
    let message = table_info
        .messages
        .iter_mut()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or_else(|| {
            crate::Error::InvalidFormat(
                "source-built Keynote table-info payload is missing".to_owned(),
            )
        })?;
    let mut decoded = tst::TableInfoArchive::decode(message.data.as_slice())?;
    decoded.super_.parent = None;
    message.data = decoded.encode_to_vec();
    for info in &mut table_info.archive_info.message_infos {
        info.object_references.retain(|identifier| *identifier != 8);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| *identifier != 8);
        }
    }
    package.update_archive(DOCUMENT_ARCHIVE_ENTRY, |archive| {
        for object in table_objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;

    let mut style_objects = source.archive(STYLESHEET_ARCHIVE_ENTRY)?.objects;
    style_objects.retain(|object| {
        object
            .archive_info
            .identifier
            .is_some_and(|identifier| TABLE_TEMPLATE_STYLE_OBJECTS.contains(&identifier))
    });
    package.update_archive(STYLESHEET_ARCHIVE_ENTRY, |archive| {
        for object in style_objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;

    package.update_archive(DOCUMENT_ARCHIVE_ENTRY, |archive| {
        let theme = archive.object_mut(THEME).ok_or_else(|| {
            crate::Error::InvalidFormat("source-built Keynote theme is missing".to_owned())
        })?;
        let message_index = theme
            .messages
            .iter()
            .position(|message| message.type_ == KeynoteMessageType::Theme.value())
            .ok_or_else(|| {
                crate::Error::InvalidFormat(
                    "source-built Keynote theme payload is missing".to_owned(),
                )
            })?;
        let message_type = theme.messages[message_index].type_;
        let mut decoded = IWorkThemeArchive::decode(&theme.messages[message_index].data)?;
        decoded.extensions.table = Some(tst::ThemePresetsArchive {
            table_style_presets: vec![reference(TABLE_PRESET_TEMPLATE)],
            ..Default::default()
        });
        theme.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data: decoded.encode()?,
            },
        )?;
        let references = &mut theme.archive_info.message_infos[message_index].object_references;
        if !references.contains(&TABLE_PRESET_TEMPLATE) {
            references.push(TABLE_PRESET_TEMPLATE);
        }
        Ok(())
    })?;

    add_component_object_uuids(package, DOCUMENT, TABLE_TEMPLATE_DOCUMENT_OBJECTS)?;
    add_component_object_uuids(package, STYLESHEET, TABLE_TEMPLATE_STYLE_OBJECTS)?;
    for &identifier in TABLE_TEMPLATE_STYLE_OBJECTS {
        add_component_external_reference(package, DOCUMENT, STYLESHEET, identifier)?;
    }
    Ok(())
}

impl KeynoteEditor {
    pub fn builder() -> KeynoteDocumentBuilder {
        KeynoteDocumentBuilder::new()
    }

    pub fn create() -> Result<Self> {
        KeynoteDocumentBuilder::new().build()
    }
}

fn document_archive(builder: &KeynoteDocumentBuilder, template_id: tsp::Uuid) -> Result<Archive> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive {
                locale_identifier: Some(builder.locale.clone()),
                annotation_author_storage: Some(reference(ANNOTATION_AUTHOR_STORAGE)),
                creation_locale_identifier: Some(builder.locale.clone()),
                prevent_image_conversion_on_open: Some(true),
                has_user_defined_locale: Some(false),
                should_measure_negatively_tracked_text_correctly: Some(true),
                use_optimized_text_vertical_alignment: Some(true),
                should_allow_ligatures_in_minimally_tracked_text: Some(true),
                formatting_symbols: Some(formatting_symbols()),
                ..Default::default()
            },
            document_language: Some(builder.language.clone()),
            calculation_engine: Some(reference(CALCULATION_ENGINE)),
            view_state: Some(reference(VIEW_STATE)),
            function_browser_state: Some(reference(FUNCTION_BROWSER_STATE)),
            needs_media_compatibility_upgrade: Some(false),
            template_identifier: Some("Application/Litchi/Blank/Wide".to_owned()),
            custom_format_list: Some(reference(CUSTOM_FORMAT_LIST)),
            collaborative_media_compatibility_upgrade_did_fail: Some(false),
            can_use_hevc: Some(false),
            is_content_source: Some(false),
            ..Default::default()
        },
        show: reference(SHOW),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(THEME),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(LIVE_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: builder.width,
            height: builder.height,
        },
        stylesheet: reference(STYLESHEET),
        loop_presentation: Some(false),
        mode: Some(kn::show_archive::KnShowMode::KKnShowModeNormal as i32),
        autoplay_transition_delay: Some(5.0),
        autoplay_build_delay: Some(2.0),
        idle_timer_active: Some(false),
        idle_timer_delay: Some(900.0),
        soundtrack: Some(reference(SOUNDTRACK)),
        automatically_plays_upon_open: Some(false),
        ..Default::default()
    };
    let live_node = slide_node(LIVE_SLIDE, template_id, true, builder.slide_number_visible);
    let template_node = slide_node(TEMPLATE_SLIDE, template_id, false, false);

    let common_theme = IWorkThemeArchive::new(
        tss::ThemeArchive {
            theme_identifier: Some("Litchi Blank".to_owned()),
            document_stylesheet: Some(reference(STYLESHEET)),
            color_presets: repeated(30, black),
            ..Default::default()
        },
        IWorkThemeExtensions {
            drawing: Some(tsd::ThemePresetsArchive {
                gradient_fill_presets: repeated(6, gradient_fill),
                image_fill_presets: repeated(6, image_fill),
                shadow_presets: shadow_presets(),
                line_style_presets: repeated_reference(6, SHAPE_STYLE),
                shape_style_presets: repeated_reference(6, SHAPE_STYLE),
                textbox_style_presets: repeated_reference(1, SHAPE_STYLE),
                image_style_presets: repeated_reference(6, MEDIA_STYLE),
                movie_style_presets: repeated_reference(6, MEDIA_STYLE),
                drawing_line_style_presets: repeated_reference(1, SHAPE_STYLE),
            }),
            text: Some(tswp::ThemePresetsArchive {
                list_style_presets: repeated_reference(5, LIST_STYLE),
                character_style_presets: repeated_reference(6, CHARACTER_STYLE),
                paragraph_style_presets: repeated_reference(13, PARAGRAPH_STYLE),
                dropcap_style_presets: repeated_reference(6, DROP_CAP_STYLE),
                ..Default::default()
            }),
            chart: None,
            table: None,
            application: Some(tsa::ThemePresetsArchive {
                caption_style_presets: repeated_reference(2, PARAGRAPH_STYLE),
                svg_import_style_presets: repeated_reference(1, SHAPE_STYLE),
            }),
        },
    )
    .encode()?;
    let common_fields = parse_wire_fields(&common_theme)?;
    let common_field = common_fields.first().ok_or_else(|| {
        crate::Error::InvalidFormat("generated Keynote theme wrapper is empty".to_owned())
    })?;
    let common_payload = &common_theme[common_field.payload_start..common_field.end];
    let theme = kn::ThemeArchive {
        super_: tss::ThemeArchive::default(),
        templates: vec![reference(TEMPLATE_NODE)],
        uuid: Some(fresh_uuid()),
        default_template_slide_node: Some(reference(TEMPLATE_NODE)),
        default_template_slide_node_reference: Some(reference(TEMPLATE_NODE)),
        default_template_slide_node_is_our_best_guess: Some(false),
        live_video_source_collection: Some(reference(LIVE_VIDEO_COLLECTION)),
        ..Default::default()
    };
    let theme =
        patch_length_delimited_field(&theme.encode_to_vec(), 1, true, Some(common_payload))?;

    Ok(Archive {
        objects: vec![
            object(
                DOCUMENT,
                KeynoteMessageType::Document,
                document,
                &[
                    SHOW,
                    CALCULATION_ENGINE,
                    VIEW_STATE,
                    FUNCTION_BROWSER_STATE,
                    CUSTOM_FORMAT_LIST,
                    ANNOTATION_AUTHOR_STORAGE,
                ],
            )?,
            raw_object(
                THEME,
                KeynoteMessageType::Theme,
                theme,
                &[
                    STYLESHEET,
                    TEMPLATE_NODE,
                    LIST_STYLE,
                    PARAGRAPH_STYLE,
                    CHARACTER_STYLE,
                    SHAPE_STYLE,
                    MEDIA_STYLE,
                    DROP_CAP_STYLE,
                    LIVE_VIDEO_COLLECTION,
                ],
            )?,
            object(
                SHOW,
                KeynoteMessageType::Show,
                show,
                &[THEME, LIVE_NODE, STYLESHEET, SOUNDTRACK],
            )?,
            object(
                LIVE_NODE,
                KeynoteMessageType::SlideNode,
                live_node,
                &[LIVE_SLIDE],
            )?,
            object(
                TEMPLATE_NODE,
                KeynoteMessageType::SlideNode,
                template_node,
                &[TEMPLATE_SLIDE],
            )?,
            object(
                SOUNDTRACK,
                KeynoteMessageType::Soundtrack,
                kn::Soundtrack {
                    volume: Some(1.0),
                    mode: Some(kn::soundtrack::SoundtrackMode::KKnSoundtrackModePlayOnce as i32),
                    movie_media: Vec::new(),
                },
                &[],
            )?,
            object_with_versions(
                LIVE_VIDEO_COLLECTION,
                KeynoteMessageType::LiveVideoCollection,
                kn::LiveVideoSourceCollection {
                    sources: Vec::new(),
                    default_source: Some(reference(DEFAULT_LIVE_VIDEO_SOURCE)),
                },
                &[DEFAULT_LIVE_VIDEO_SOURCE],
                &[11, 2, 4],
            )?,
            object_with_versions(
                DEFAULT_LIVE_VIDEO_SOURCE,
                KeynoteMessageType::LiveVideoSource,
                kn::LiveVideoSource {
                    name: Some("Default Camera".to_owned()),
                    collaboration_command_usage_state: Some(
                        kn::LiveVideoSourceCollaborationCommandUsageState {
                            has_multiple_collaboration_command_usage_tokens: Some(false),
                            ..Default::default()
                        },
                    ),
                    symbol_image_identifier: Some(0),
                    symbol_tint_color_identifier: Some(0),
                    is_default_source: Some(true),
                    ..Default::default()
                },
                &[],
                &[11, 2, 4],
            )?,
            object(
                FUNCTION_BROWSER_STATE,
                KeynoteMessageType::FunctionBrowserState,
                tsa::FunctionBrowserStateArchive {
                    current_function: Some(0),
                    ..Default::default()
                },
                &[],
            )?,
            object(
                CUSTOM_FORMAT_LIST,
                KeynoteMessageType::CustomFormatList,
                tsk::CustomFormatListArchive::default(),
                &[],
            )?,
        ],
    })
}

fn template_archive(builder: &KeynoteDocumentBuilder) -> Result<Archive> {
    let slide = slide(
        TEMPLATE_TITLE,
        TEMPLATE_BODY,
        TEMPLATE_GUIDES,
        TEMPLATE_SLIDE_NUMBER,
        false,
        None,
        None,
        Some(DEFAULT_LAYOUT_NAME.to_owned()),
    );
    Ok(Archive {
        objects: vec![
            object(
                TEMPLATE_SLIDE,
                KeynoteMessageType::Slide,
                slide,
                &[
                    SLIDE_STYLE,
                    TEMPLATE_TITLE,
                    TEMPLATE_BODY,
                    TEMPLATE_GUIDES,
                    TEMPLATE_SLIDE_NUMBER,
                ],
            )?,
            object(
                TEMPLATE_TITLE,
                KeynoteMessageType::Placeholder,
                placeholder(
                    TEMPLATE_SLIDE,
                    TEMPLATE_TITLE_STORAGE,
                    title_geometry(builder),
                    kn::placeholder_archive::Kind::KKindTitlePlaceholder,
                ),
                &[TEMPLATE_SLIDE, TEMPLATE_TITLE_STORAGE, SHAPE_STYLE],
            )?,
            object(
                TEMPLATE_BODY,
                KeynoteMessageType::Placeholder,
                placeholder(
                    TEMPLATE_SLIDE,
                    TEMPLATE_BODY_STORAGE,
                    body_geometry(builder),
                    kn::placeholder_archive::Kind::KKindBodyPlaceholder,
                ),
                &[TEMPLATE_SLIDE, TEMPLATE_BODY_STORAGE, SHAPE_STYLE],
            )?,
            object(
                TEMPLATE_TITLE_STORAGE,
                KeynoteMessageType::Storage,
                text_storage(String::new(), &builder.language),
                &[STYLESHEET, PARAGRAPH_STYLE, LIST_STYLE],
            )?,
            object(
                TEMPLATE_BODY_STORAGE,
                KeynoteMessageType::Storage,
                text_storage(String::new(), &builder.language),
                &[STYLESHEET, PARAGRAPH_STYLE, LIST_STYLE],
            )?,
            object(
                TEMPLATE_GUIDES,
                KeynoteMessageType::GuideStorage,
                tsd::GuideStorageArchive::default(),
                &[],
            )?,
            object(
                TEMPLATE_SLIDE_NUMBER,
                KeynoteMessageType::Placeholder,
                slide_number::placeholder(
                    TEMPLATE_SLIDE,
                    slide_number::PlaceholderContext::Template,
                ),
                &[TEMPLATE_SLIDE, SHAPE_STYLE],
            )?,
        ],
    })
}

fn slide_archive(builder: &KeynoteDocumentBuilder) -> Result<Archive> {
    let slide = slide(
        LIVE_TITLE,
        LIVE_BODY,
        LIVE_GUIDES,
        LIVE_SLIDE_NUMBER,
        builder.slide_number_visible,
        Some(TEMPLATE_SLIDE),
        Some(LIVE_NOTE),
        None,
    );
    Ok(Archive {
        objects: vec![
            object(
                LIVE_SLIDE,
                KeynoteMessageType::Slide,
                slide,
                &[
                    SLIDE_STYLE,
                    LIVE_TITLE,
                    LIVE_BODY,
                    LIVE_GUIDES,
                    TEMPLATE_SLIDE,
                    LIVE_NOTE,
                    LIVE_SLIDE_NUMBER,
                ],
            )?,
            object(
                LIVE_TITLE,
                KeynoteMessageType::Placeholder,
                placeholder(
                    LIVE_SLIDE,
                    LIVE_TITLE_STORAGE,
                    title_geometry(builder),
                    kn::placeholder_archive::Kind::KKindTitlePlaceholder,
                ),
                &[LIVE_SLIDE, LIVE_TITLE_STORAGE, SHAPE_STYLE],
            )?,
            object(
                LIVE_BODY,
                KeynoteMessageType::Placeholder,
                placeholder(
                    LIVE_SLIDE,
                    LIVE_BODY_STORAGE,
                    body_geometry(builder),
                    kn::placeholder_archive::Kind::KKindBodyPlaceholder,
                ),
                &[LIVE_SLIDE, LIVE_BODY_STORAGE, SHAPE_STYLE],
            )?,
            object(
                LIVE_TITLE_STORAGE,
                KeynoteMessageType::Storage,
                text_storage(builder.title.clone(), &builder.language),
                &[STYLESHEET, PARAGRAPH_STYLE, LIST_STYLE],
            )?,
            object(
                LIVE_BODY_STORAGE,
                KeynoteMessageType::Storage,
                text_storage(builder.subtitle.clone(), &builder.language),
                &[STYLESHEET, PARAGRAPH_STYLE, LIST_STYLE],
            )?,
            object(
                LIVE_NOTE,
                KeynoteMessageType::Note,
                kn::NoteArchive {
                    contained_storage: reference(LIVE_NOTE_STORAGE),
                },
                &[LIVE_NOTE_STORAGE],
            )?,
            object(
                LIVE_NOTE_STORAGE,
                KeynoteMessageType::Storage,
                note_storage(builder.presenter_notes.clone(), &builder.language),
                &[STYLESHEET, PARAGRAPH_STYLE, LIST_STYLE],
            )?,
            object(
                LIVE_GUIDES,
                KeynoteMessageType::GuideStorage,
                tsd::GuideStorageArchive::default(),
                &[],
            )?,
            object(
                LIVE_SLIDE_NUMBER,
                KeynoteMessageType::Placeholder,
                slide_number::placeholder(LIVE_SLIDE, slide_number::PlaceholderContext::Live),
                &[LIVE_SLIDE, SHAPE_STYLE],
            )?,
        ],
    })
}

#[allow(deprecated)]
fn slide_node(
    slide_id: u64,
    template_id: tsp::Uuid,
    live: bool,
    slide_number_visible: bool,
) -> kn::SlideNodeArchive {
    kn::SlideNodeArchive {
        slide: Some(reference(slide_id)),
        depth: Some(1),
        thumbnails_are_dirty: Some(true),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        has_note: Some(live),
        is_slide_number_visible: Some(slide_number_visible),
        build_event_count: Some(0),
        build_event_count_cache_version: Some(1),
        has_explicit_builds: Some(false),
        has_explicit_builds_cache_version: Some(1),
        background_is_no_fill_or_color_fill_with_alpha: Some(false),
        template_slide_id: Some(template_id),
        ..Default::default()
    }
}

fn slide(
    title_id: u64,
    body_id: u64,
    guides_id: u64,
    slide_number_id: u64,
    slide_number_visible: bool,
    template_slide_id: Option<u64>,
    note_id: Option<u64>,
    name: Option<String>,
) -> kn::SlideArchive {
    let mut owned_drawables = vec![reference(title_id), reference(body_id)];
    let mut drawables_z_order = owned_drawables.clone();
    if slide_number_visible {
        owned_drawables.push(reference(slide_number_id));
        drawables_z_order.push(reference(slide_number_id));
    }
    kn::SlideArchive {
        style: reference(SLIDE_STYLE),
        transition: transition(),
        title_placeholder: Some(reference(title_id)),
        body_placeholder: Some(reference(body_id)),
        slide_number_placeholder: Some(reference(slide_number_id)),
        owned_drawables,
        drawables_z_order,
        instructional_text_map: Some(kn::slide_archive::InstructionalTextMap::default()),
        name,
        template_slide: template_slide_id.map(reference),
        user_defined_guide_storage: Some(reference(guides_id)),
        in_document: true,
        note: note_id.map(reference),
        ..Default::default()
    }
}

#[allow(deprecated)]
fn placeholder(
    parent_id: u64,
    storage_id: u64,
    geometry: tsd::GeometryArchive,
    kind: kn::placeholder_archive::Kind,
) -> kn::PlaceholderArchive {
    let pathsource = geometry.size.as_ref().map(rectangle_path_source);
    kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(geometry),
                    parent: Some(reference(parent_id)),
                    exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                        r#type: Some(TEXT_WRAP_TYPE),
                        direction: Some(TEXT_WRAP_DIRECTION),
                        fit_type: Some(TEXT_WRAP_FIT_TYPE),
                        margin: Some(TEXT_WRAP_MARGIN),
                        alpha_threshold: Some(TEXT_WRAP_ALPHA_THRESHOLD),
                        is_html_wrap: Some(false),
                    }),
                    locked: Some(false),
                    aspect_ratio_locked: Some(false),
                    title_hidden: Some(false),
                    caption_hidden: Some(false),
                    ..Default::default()
                },
                style: Some(reference(SHAPE_STYLE)),
                pathsource,
                stroke_pattern_offset_distance: Some(0.0),
                ..Default::default()
            },
            deprecated_storage: Some(reference(storage_id)),
            owned_storage: Some(reference(storage_id)),
            is_text_box: Some(true),
            ..Default::default()
        },
        kind: Some(kind as i32),
    }
}

fn rectangle_path_source(size: &tsp::Size) -> tsd::PathSourceArchive {
    use tsp::path::{Element, ElementType};

    let point = |x, y| tsp::Point { x, y };
    let element = |r#type: ElementType, points| Element {
        r#type: r#type as i32,
        points,
    };
    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size {
                width: size.width,
                height: size.height,
            }),
            path: Some(tsp::Path {
                elements: vec![
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                    element(ElementType::LineTo, vec![point(size.width, 0.0)]),
                    element(ElementType::LineTo, vec![point(size.width, size.height)]),
                    element(ElementType::LineTo, vec![point(0.0, size.height)]),
                    element(ElementType::CloseSubpath, Vec::new()),
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn title_geometry(builder: &KeynoteDocumentBuilder) -> tsd::GeometryArchive {
    geometry(
        builder.width * 0.1,
        builder.height * 0.24,
        builder.width * 0.8,
        builder.height * 0.2,
    )
}

fn body_geometry(builder: &KeynoteDocumentBuilder) -> tsd::GeometryArchive {
    geometry(
        builder.width * 0.15,
        builder.height * 0.52,
        builder.width * 0.7,
        builder.height * 0.14,
    )
}

fn geometry(x: f32, y: f32, width: f32, height: f32) -> tsd::GeometryArchive {
    tsd::GeometryArchive {
        position: Some(tsp::Point { x, y }),
        size: Some(tsp::Size { width, height }),
        flags: Some(3),
        angle: Some(0.0),
    }
}

fn transition() -> kn::TransitionArchive {
    kn::TransitionArchive {
        attributes: kn::TransitionAttributesArchive {
            animation_attributes: Some(kn::AnimationAttributesArchive {
                animation_type: Some("Transition".to_owned()),
                effect: Some("none".to_owned()),
                duration: Some(1.0),
                delay: Some(0.5),
                is_automatic: Some(false),
                writing_direction_is_rtl: Some(false),
                ..Default::default()
            }),
            ..Default::default()
        },
    }
}

fn text_storage(text: String, language: &str) -> tswp::StorageArchive {
    tswp::StorageArchive {
        style_sheet: Some(reference(STYLESHEET)),
        text: vec![text],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(Some(PARAGRAPH_STYLE))),
        table_para_data: Some(para_data_table()),
        table_list_style: Some(object_attribute_table(Some(LIST_STYLE))),
        table_para_starts: Some(para_data_table()),
        table_language: Some(tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language.to_owned()),
            }],
        }),
        table_para_bidi: Some(para_data_table()),
        ..Default::default()
    }
}

fn note_storage(text: String, language: &str) -> tswp::StorageArchive {
    tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Note as i32),
        ..text_storage(text, language)
    }
}

fn object_attribute_table(identifier: Option<u64>) -> tswp::ObjectAttributeTable {
    tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: identifier.map(reference),
        }],
    }
}

fn para_data_table() -> tswp::ParaDataAttributeTable {
    tswp::ParaDataAttributeTable {
        entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        }],
    }
}

fn stylesheet_archive() -> Result<Archive> {
    let styles = &STYLESHEET_OBJECTS[1..];
    Ok(Archive {
        objects: vec![
            object(
                STYLESHEET,
                KeynoteMessageType::Stylesheet,
                tss::StylesheetArchive {
                    styles: styles.iter().copied().map(reference).collect(),
                    is_locked: Some(false),
                    can_cull_styles: Some(true),
                    ..Default::default()
                },
                styles,
            )?,
            object(
                SLIDE_STYLE,
                KeynoteMessageType::SlideStyle,
                kn::SlideStyleArchive {
                    super_: style("Slide", "litchi-slide-default"),
                    override_count: Some(3),
                    slide_properties: Some(kn::SlideStylePropertiesArchive {
                        fill: Some(tsd::FillArchive {
                            color: Some(white()),
                            ..Default::default()
                        }),
                        title_placeholder_visibility: Some(true),
                        body_placeholder_visibility: Some(true),
                        ..Default::default()
                    }),
                },
                &[STYLESHEET],
            )?,
            object(
                LIST_STYLE,
                KeynoteMessageType::ListStyle,
                tswp::ListStyleArchive {
                    super_: style("None", "litchi-list-none"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                PARAGRAPH_STYLE,
                KeynoteMessageType::ParagraphStyle,
                tswp::ParagraphStyleArchive {
                    super_: style("Body", "litchi-body"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                CHARACTER_STYLE,
                KeynoteMessageType::CharacterStyle,
                tswp::CharacterStyleArchive {
                    super_: style("Default", "litchi-character-default"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                SHAPE_STYLE,
                KeynoteMessageType::ShapeStyle,
                tswp::ShapeStyleArchive {
                    super_: tsd::ShapeStyleArchive {
                        super_: style("Text", "litchi-text-default"),
                        override_count: Some(0),
                        ..Default::default()
                    },
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                MEDIA_STYLE,
                KeynoteMessageType::MediaStyle,
                tsd::MediaStyleArchive {
                    super_: style("Media", "litchi-media-default"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                DROP_CAP_STYLE,
                KeynoteMessageType::DropCapStyle,
                tswp::DropCapStyleArchive {
                    super_: style("Drop Cap", "litchi-dropcap"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
        ],
    })
}

fn calculation_archive(locale: &str) -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            CALCULATION_ENGINE,
            KeynoteMessageType::CalculationEngine,
            tsce::CalculationEngineArchive {
                dependency_tracker: tsce::DependencyTrackerArchive {
                    number_of_formulas: Some(0),
                    ..Default::default()
                },
                saved_locale_identifier: Some(locale.to_owned()),
                ..Default::default()
            },
            &[],
        )?],
    })
}

#[allow(deprecated)]
fn view_state_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![
            object(
                VIEW_STATE,
                KeynoteMessageType::ViewState,
                tsk::ViewStateArchive {
                    view_state_root: reference(UI_STATE),
                    ..Default::default()
                },
                &[UI_STATE],
            )?,
            object(
                UI_STATE,
                KeynoteMessageType::UiState,
                kn::UiStateArchive {
                    desktop_slide_view_content_fits_window: Some(true),
                    desktop_canvas_view_scale: Some(0.5),
                    show_slide_guides: Some(true),
                    show_template_guides: Some(true),
                    shows_comments: Some(true),
                    shows_ruler: Some(false),
                    desktop_navigator_view_width: Some(128.0),
                    editing_disabled: Some(false),
                    ..Default::default()
                },
                &[],
            )?,
        ],
    })
}

fn annotation_author_storage_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            ANNOTATION_AUTHOR_STORAGE,
            KeynoteMessageType::AnnotationAuthorStorage,
            tsk::AnnotationAuthorStorageArchive::default(),
            &[],
        )?],
    })
}

fn document_metadata_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            DOCUMENT_METADATA,
            KeynoteMessageType::DocumentMetadata,
            tsp::DocumentMetadata {
                is_in_collaboration_mode: Some(false),
                ..Default::default()
            },
            &[],
        )?],
    })
}

fn metadata_archive(identity: &IWorkDocumentIdentity) -> Result<Archive> {
    let mut document = component(DOCUMENT, "Document", &[2, 4, 0], DOCUMENT_OBJECTS);
    document.external_references = std::iter::once(external(STYLESHEET, None))
        .chain(
            STYLESHEET_OBJECTS
                .iter()
                .copied()
                .map(|identifier| external(STYLESHEET, Some(identifier))),
        )
        .chain([
            external(TEMPLATE_SLIDE, None),
            external(LIVE_SLIDE, None),
            external(CALCULATION_ENGINE, None),
            external(VIEW_STATE, None),
            external(ANNOTATION_AUTHOR_STORAGE, None),
        ])
        .collect();
    let mut template = component(
        TEMPLATE_SLIDE,
        "TemplateSlide-8",
        &[2, 0, 0],
        TEMPLATE_OBJECTS,
    );
    template.external_references = STYLESHEET_OBJECTS
        .iter()
        .copied()
        .map(|identifier| external(STYLESHEET, Some(identifier)))
        .collect();
    let mut slide = component(LIVE_SLIDE, "Slide-14", &[2, 0, 0], SLIDE_OBJECTS);
    slide.external_references = std::iter::once(external(TEMPLATE_SLIDE, None))
        .chain(
            STYLESHEET_OBJECTS
                .iter()
                .copied()
                .map(|identifier| external(STYLESHEET, Some(identifier))),
        )
        .collect();
    let metadata = tsp::PackageMetadata {
        last_object_identifier: STYLESHEET,
        revision: Some(tsp::DocumentRevision {
            sequence_32: Some(0),
            identifier: Some(identity.version_uuid().to_owned()),
            sequence_64: None,
        }),
        components: vec![
            component(DOCUMENT_METADATA, "DocumentMetadata", &[2, 0, 0], &[]),
            template,
            slide,
            component(
                STYLESHEET,
                "DocumentStylesheet",
                &[2, 0, 0],
                STYLESHEET_OBJECTS,
            ),
            component(
                CALCULATION_ENGINE,
                "CalculationEngine",
                &[2, 0, 0],
                &[CALCULATION_ENGINE],
            ),
            component(VIEW_STATE, "ViewState", &[2, 0, 0], &[VIEW_STATE, UI_STATE]),
            component(
                ANNOTATION_AUTHOR_STORAGE,
                "AnnotationAuthorStorage",
                &[2, 0, 0],
                &[ANNOTATION_AUTHOR_STORAGE],
            ),
            document,
        ],
        read_version: vec![2, 4, 0],
        write_version: vec![11, 2, 4],
        file_format_version: vec![14, 4, 1],
        save_token: Some(1),
        ..Default::default()
    };
    Ok(Archive {
        objects: vec![object(
            METADATA,
            KeynoteMessageType::PackageMetadata,
            metadata,
            &[],
        )?],
    })
}

fn component(
    identifier: u64,
    locator: &str,
    version: &[u32],
    objects: &[u64],
) -> tsp::ComponentInfo {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        document_read_version: version.to_vec(),
        document_write_version: version.to_vec(),
        save_token: Some(1),
        object_uuid_map_entries: objects.iter().copied().map(object_uuid).collect(),
        ..Default::default()
    }
}

fn external(
    component_identifier: u64,
    object_identifier: Option<u64>,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier,
        is_weak: None,
    }
}

fn add_plists(package: &mut IWorkPackage, identity: &IWorkDocumentIdentity) -> Result<()> {
    let mut properties = plist::Dictionary::new();
    for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
        properties.insert(
            key.to_owned(),
            Value::String(identity.document_uuid().to_owned()),
        );
    }
    properties.insert(
        "fileFormatVersion".to_owned(),
        Value::String("14.4.1".to_owned()),
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
        Value::String(format!("0::{}", identity.version_uuid())),
    );
    let mut encoded = Vec::new();
    Value::Dictionary(properties)
        .to_writer_binary(&mut encoded)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Keynote properties: {error}"
            ))
        })?;
    package.insert_entry("Metadata/Properties.plist", encoded)?;
    package.insert_entry(
        "Metadata/DocumentIdentifier",
        identity.document_uuid().as_bytes().to_vec(),
    )?;
    let mut history = Vec::new();
    Value::Array(vec![Value::String("Created by litchi-iwa".to_owned())])
        .to_writer_binary(&mut history)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Keynote build history: {error}"
            ))
        })?;
    package.insert_entry("Metadata/BuildVersionHistory.plist", history)?;
    Ok(())
}

fn formatting_symbols() -> tsk::FormattingSymbolsArchive {
    tsk::FormattingSymbolsArchive {
        version: Some("4302.00*".to_owned()),
        calendar: Some("gregorian".to_owned()),
        numbering_system: Some("latn".to_owned()),
        am_symbol: Some("AM".to_owned()),
        pm_symbol: Some("PM".to_owned()),
        short_date_pattern: Some("M/d/yy".to_owned()),
        medium_date_pattern: Some("MMM d, y".to_owned()),
        long_date_pattern: Some("MMMM d, y".to_owned()),
        full_date_pattern: Some("EEEE, MMMM d, y".to_owned()),
        short_time_pattern: Some("HH:mm".to_owned()),
        medium_time_pattern: Some("HH:mm:ss".to_owned()),
        long_time_pattern: Some("HH:mm:ss z".to_owned()),
        full_time_pattern: Some("HH:mm:ss zzzz".to_owned()),
        decimal_separator: Some(".".to_owned()),
        grouping_separator: Some(",".to_owned()),
        currency_decimal_separator: Some(".".to_owned()),
        currency_grouping_separator: Some(",".to_owned()),
        plus_sign: Some("+".to_owned()),
        minus_sign: Some("-".to_owned()),
        exponential_symbol: Some("E".to_owned()),
        percent_symbol: Some("%".to_owned()),
        per_mille_symbol: Some("‰".to_owned()),
        infinity_symbol: Some("+∞".to_owned()),
        nan_symbol: Some("NaN".to_owned()),
        decimal_pattern: Some("#,##0.###".to_owned()),
        scientific_pattern: Some("#E0".to_owned()),
        percent_pattern: Some("#,##0%".to_owned()),
        currency_pattern: Some("¤#,##0.00".to_owned()),
        currency_code: Some("USD".to_owned()),
        ..Default::default()
    }
}

fn object_uuid(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: fresh_tsp_uuid(),
    }
}

fn fresh_tsp_uuid() -> tsp::Uuid {
    let bytes = litchi_core::id::generate_guid_bytes();
    tsp::Uuid {
        upper: u64::from_be_bytes(bytes[..8].try_into().expect("eight-byte UUID half")),
        lower: u64::from_be_bytes(bytes[8..].try_into().expect("eight-byte UUID half")),
    }
}

fn fresh_uuid() -> String {
    litchi_core::id::generate_guid_braced()
        .trim_matches(['{', '}'])
        .to_owned()
}

fn object(
    identifier: u64,
    message_type: KeynoteMessageType,
    message: impl Message,
    references: &[u64],
) -> Result<ArchiveObject> {
    raw_object(
        identifier,
        message_type,
        message.encode_to_vec(),
        references,
    )
}

fn object_with_versions(
    identifier: u64,
    message_type: KeynoteMessageType,
    message: impl Message,
    references: &[u64],
    versions: &[u32],
) -> Result<ArchiveObject> {
    let mut object = raw_object(
        identifier,
        message_type,
        message.encode_to_vec(),
        references,
    )?;
    object.archive_info.message_infos[0].versions = versions.to_vec();
    Ok(object)
}

fn raw_object(
    identifier: u64,
    message_type: KeynoteMessageType,
    data: Vec<u8>,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type.value(),
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = vec![1, 0, 5];
    object.archive_info.message_infos[0].object_references = references.to_vec();
    Ok(object)
}

fn style(name: &str, identifier: &str) -> tss::StyleArchive {
    tss::StyleArchive {
        name: Some(name.to_owned()),
        style_identifier: Some(identifier.to_owned()),
        stylesheet: Some(reference(STYLESHEET)),
        ..Default::default()
    }
}

fn repeated<T>(count: usize, make: impl Fn() -> T) -> Vec<T> {
    std::iter::repeat_with(make).take(count).collect()
}

fn repeated_reference(count: usize, identifier: u64) -> Vec<tsp::Reference> {
    repeated(count, || reference(identifier))
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn black() -> tsp::Color {
    color(0.0)
}

fn white() -> tsp::Color {
    color(1.0)
}

fn color(component: f32) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(component),
        g: Some(component),
        b: Some(component),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(1.0),
        ..Default::default()
    }
}

fn gradient_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        gradient: Some(tsd::GradientArchive {
            r#type: Some(tsd::gradient_archive::GradientType::Linear as i32),
            stops: vec![
                tsd::gradient_archive::GradientStop {
                    color: Some(white()),
                    fraction: Some(0.0),
                    inflection: Some(0.5),
                },
                tsd::gradient_archive::GradientStop {
                    color: Some(black()),
                    fraction: Some(1.0),
                    inflection: Some(0.5),
                },
            ],
            opacity: Some(1.0),
            advanced_gradient: Some(false),
            anglegradient: Some(tsd::AngleGradientArchive {
                gradientangle: Some(3.0 * std::f32::consts::FRAC_PI_2),
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn image_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        image: Some(tsd::ImageFillArchive {
            technique: Some(tsd::image_fill_archive::ImageFillTechnique::Tile as i32),
            tint: Some(white()),
            fillsize: Some(tsp::Size {
                width: 300.0,
                height: 300.0,
            }),
            interprets_untagged_image_data_as_generic: Some(true),
            referencecolor: Some(white()),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn shadow_presets() -> Vec<tsd::ShadowArchive> {
    vec![
        drop_shadow(90.0, 2.0, 5, 0.5),
        drop_shadow(45.0, 5.0, 4, 0.5),
        drop_shadow(90.0, 8.0, 10, 0.5),
        drop_shadow(45.0, 0.0, 10, 0.75),
        curved_shadow(0.665_777, 15, 1.0, -0.123_095_46),
        curved_shadow(1.0, 15, 0.75, 0.164_878_73),
        contact_shadow(0.0, 28, 0.75, 0.173_648_18),
        contact_shadow(9.418_75, 30, 0.75, 0.25),
    ]
}

fn drop_shadow(angle: f32, offset: f32, radius: i32, opacity: f32) -> tsd::ShadowArchive {
    shadow(
        angle,
        offset,
        radius,
        opacity,
        tsd::shadow_archive::ShadowType::TsdDropShadow,
    )
}

fn curved_shadow(offset: f32, radius: i32, opacity: f32, curve: f32) -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        curved_shadow: Some(tsd::CurvedShadowArchive { curve: Some(curve) }),
        ..shadow(
            90.0,
            offset,
            radius,
            opacity,
            tsd::shadow_archive::ShadowType::TsdCurvedShadow,
        )
    }
}

fn contact_shadow(offset: f32, radius: i32, opacity: f32, height: f32) -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        contact_shadow: Some(tsd::ContactShadowArchive {
            height: Some(height),
            ..Default::default()
        }),
        ..shadow(
            0.0,
            offset,
            radius,
            opacity,
            tsd::shadow_archive::ShadowType::TsdContactShadow,
        )
    }
}

fn shadow(
    angle: f32,
    offset: f32,
    radius: i32,
    opacity: f32,
    kind: tsd::shadow_archive::ShadowType,
) -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        color: Some(black()),
        angle: Some(angle),
        offset: Some(offset),
        radius: Some(radius),
        opacity: Some(opacity),
        is_enabled: Some(true),
        r#type: Some(kind as i32),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn generated_package_contains_only_synthetic_required_entries() {
        const EXPECTED_ENTRIES: [&str; 12] = [
            DOCUMENT_ARCHIVE_ENTRY,
            TEMPLATE_ARCHIVE_ENTRY,
            SLIDE_ARCHIVE_ENTRY,
            STYLESHEET_ARCHIVE_ENTRY,
            CALCULATION_ARCHIVE_ENTRY,
            VIEW_STATE_ARCHIVE_ENTRY,
            ANNOTATION_ARCHIVE_ENTRY,
            DOCUMENT_METADATA_ARCHIVE_ENTRY,
            PACKAGE_METADATA_ARCHIVE_ENTRY,
            "Metadata/Properties.plist",
            "Metadata/DocumentIdentifier",
            "Metadata/BuildVersionHistory.plist",
        ];

        let first = KeynoteDocumentBuilder::new().build_package().unwrap();
        let second = KeynoteDocumentBuilder::new().build_package().unwrap();
        assert_eq!(first.entry_names().collect::<Vec<_>>(), EXPECTED_ENTRIES);
        assert_eq!(first.len(), EXPECTED_ENTRIES.len());
        assert!(first.entry_names().all(|name| !name.starts_with("Data/")));
        assert!(first.entry_names().all(|name| !name.starts_with("preview")));
        assert_ne!(
            first.entry("Metadata/DocumentIdentifier"),
            second.entry("Metadata/DocumentIdentifier")
        );
    }

    #[test]
    fn generated_presentation_round_trips_and_is_editable() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Quarterly Review")
            .subtitle("Built without a template")
            .presenter_notes("Start with the result.")
            .build()
            .unwrap();

        let slides = editor.slides().unwrap();
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].title.as_deref(), Some("Quarterly Review"));
        assert_eq!(slides[0].body.as_deref(), Some("Built without a template"));
        assert_eq!(slides[0].notes.as_deref(), Some("Start with the result."));

        editor.set_slide_title(0, "Updated").unwrap();
        editor.set_slide_body(0, "Body").unwrap();
        editor.set_slide_notes(0, "Notes").unwrap();
        let bytes = editor.to_bytes().unwrap();
        let reopened = KeynoteEditor::from_bytes(&bytes).unwrap();
        let slide = &reopened.slides().unwrap()[0];
        assert_eq!(slide.title.as_deref(), Some("Updated"));
        assert_eq!(slide.body.as_deref(), Some("Body"));
        assert_eq!(slide.notes.as_deref(), Some("Notes"));
    }

    #[test]
    fn generated_presentation_can_add_and_remove_slides() {
        let mut editor = KeynoteEditor::create().unwrap();
        let layout = editor.default_slide_layout().unwrap();
        let created = editor.add_slide(layout).unwrap();
        assert_eq!(created.index, 1);
        editor.set_slide_title(1, "Second").unwrap();
        assert_eq!(editor.slides().unwrap()[1].title.as_deref(), Some("Second"));
        let removed = editor.remove_slide(0).unwrap();
        assert_eq!(removed.title.as_deref(), Some(DEFAULT_TITLE));
        assert_eq!(editor.slides().unwrap().len(), 1);
    }

    #[test]
    fn generated_presentation_materializes_native_slide_number_placeholders() {
        let mut hidden = KeynoteEditor::create().unwrap();
        assert_eq!(
            hidden.slides().unwrap()[0].is_slide_number_visible,
            Some(false)
        );
        hidden.set_slide_number_visible(0, true).unwrap();
        assert_eq!(
            hidden.slides().unwrap()[0].is_slide_number_visible,
            Some(true)
        );
        hidden.set_slide_number_visible(0, false).unwrap();
        assert_eq!(
            hidden.slides().unwrap()[0].is_slide_number_visible,
            Some(false)
        );

        let mut visible = KeynoteDocumentBuilder::new()
            .slide_number_visible(true)
            .build()
            .unwrap();
        assert_eq!(
            visible.slides().unwrap()[0].is_slide_number_visible,
            Some(true)
        );

        let layout = visible.default_slide_layout().unwrap();
        let created = visible.add_slide(layout).unwrap();
        assert_eq!(created.is_slide_number_visible, Some(false));
        visible.set_slide_number_visible(1, true).unwrap();
        assert_eq!(
            visible.slides().unwrap()[1].is_slide_number_visible,
            Some(true)
        );
        assert_eq!(visible.slide_text_storages(1).unwrap().len(), 2);

        let reopened = KeynoteEditor::from_bytes(&visible.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slides()
                .unwrap()
                .iter()
                .map(|slide| slide.is_slide_number_visible)
                .collect::<Vec<_>>(),
            [Some(true), Some(true)]
        );
    }

    #[test]
    fn invalid_slide_dimensions_are_rejected() {
        for (width, height) in [(0.0, 1.0), (1.0, f32::NAN), (f32::INFINITY, 1.0)] {
            assert!(
                KeynoteDocumentBuilder::new()
                    .slide_size(width, height)
                    .build_package()
                    .is_err()
            );
        }
    }
}
