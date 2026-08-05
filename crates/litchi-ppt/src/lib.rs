//! PowerPoint (`.ppt`) presentation support.
//
// This module provides parsing of Microsoft PowerPoint presentations
// in the legacy binary format (.ppt files), which uses OLE2 structured storage.
//
// # Architecture
//
// The module is organized around these key types:
// - `Package`: The overall .ppt file package (OLE container)
// - `Presentation`: The main presentation content and API
// - `Slide`: Individual slide content and API
// - `Shape`, `TextBox`, `Placeholder`: Shape and placeholder support
//
// # PPT File Structure
//
// A .ppt file is an OLE2 structured storage containing several streams:
// - **PowerPoint Document**: Main presentation stream containing document properties
// - **Pictures**: Embedded pictures and images
// - **\x05SummaryInformation**: Document metadata
//
// # Example
//
//! ```rust,no_run
//! use litchi_ppt::{Package, shapes::ShapeEnum};
//
// // Open a presentation
// let mut package = Package::open("presentation.ppt")?;
// let pres = package.presentation()?;
//
// // Extract all text
// let text = pres.text()?;
// println!("Presentation text: {}", text);
//
// // Access slides and shapes
// for slide in pres.slides()? {
//     println!("Slide: {}", slide.text()?);
//
//     // Access individual shapes
//     for shape in slide.shapes()? {
//         match shape {
//             ShapeEnum::TextBox(textbox) => {
//                 println!("Text box: {}", textbox.text());
//             }
//             ShapeEnum::Placeholder(placeholder) => {
//                 println!("Placeholder type: {:?}", placeholder.placeholder_type());
//             }
//             _ => {}
//         }
//     }
// }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

#![forbid(unsafe_code)]

mod consts;
mod officeart_wire;

pub use consts::PptRecordType;
// Core modules
pub mod package;
pub mod presentation;
pub mod presentation_advisor;
pub mod sound_collection;
pub mod view_info;

// Non-zoom (outline and slide-sorter) view metadata
pub mod non_zoom_view;

/// PPT file writing
pub mod writer;

/// Slide module with factory and enhanced implementation
pub mod slide;

// Submodules (organized by functionality)
pub mod parsers;
pub mod persist;
pub mod records;
pub mod shapes;
pub mod text;

/// Semantic owners for content embedded in a legacy presentation.
pub mod embedded;

// PowerPoint projections over the format-neutral OfficeArt substrate.
pub mod odraw;

// PowerPoint host metadata over format-neutral Office Graph views.
pub mod chart;

// Legacy compatibility modules
pub mod bookmark_summary;
pub mod broadcast;
pub mod client_anchor;
pub mod client_data;
pub mod color_scheme;
pub mod comments;
pub mod current_user;
pub mod document_atom;
pub mod document_comparison;
pub mod document_properties;
pub mod document_structure;
mod encryption;
pub mod envelope;
pub mod envelope_data;
pub mod escher_textbox;
pub mod external_media;
pub mod font;
pub mod header_footer;
pub mod html_publish;
pub mod hyperlink;
pub mod kinsoku;
pub mod main_master;
pub mod master_style;
pub mod modify_password;
pub mod named_shows;
pub mod picture_bullets;
pub mod placeholder_atom;
pub mod print_options;
pub mod privacy;
pub mod prog_tag_extensions;
pub mod prog_tags;
pub mod recolor;
pub mod routing_slip;
pub mod shape_flags;
pub mod shape_programmable_tags;
pub mod slide_extension;
pub mod slide_round_trip;
pub mod slide_show_settings;
pub mod slide_sync;
pub mod smart_tags;
pub mod style_text_prop;
pub mod text_bookmark;
pub mod text_extensions;
pub mod text_format_exception;
pub mod text_interaction;
pub mod text_metachar;
pub mod text_prop;
pub mod text_ruler;
pub mod text_run;
pub mod text_si_exception;
pub mod text_special_info;
pub mod vba_info;
pub mod view_set_info;

// Re-export main types for convenience
pub use document_comparison::{
    PowerPointDiffFlags, PowerPointDiffNode, PowerPointDiffRecordHeaders, PowerPointDiffTree10,
    PowerPointDiffType, PowerPointDocDiffFlags, PowerPointElementType,
    PowerPointMainMasterDiffFlags, PowerPointReviewingToolbarStates, PowerPointShapeDiffFlags,
    PowerPointSlideCreationEntry, PowerPointSlideDiffFlags, PowerPointSlideListTable10,
    PowerPointTableDiffFlags, PowerPointTextDiffFlags,
};
pub use encryption::PptEncryptionProfile;
pub use non_zoom_view::{
    PowerPointNoZoomViewInfo, PowerPointNonZoomViewKind, PowerPointOutlineSorterViewInfo,
    PowerPointOutlineSorterViewInformation,
};
pub use package::{Package, PptEncryptionKind, PptError, PptOpenOptions};
pub use presentation::{ParsedCustomShow, ParsedSlideComments, Presentation};
pub use presentation_advisor::{PowerPointAdvisorRule, PowerPointPresentationAdvisorSettings};
pub use slide::{
    ParsedComment, ParsedSlideTiming, Slide, SlideData, SlideDirectory, SlideDirectoryEntry,
    SlideFactory, SpeakerNotes,
};
pub use sound_collection::{
    EmbeddedPowerPointSound, PowerPointBuiltinSoundId, PowerPointSoundCollection,
};
pub use view_info::{
    PowerPointGuide, PowerPointGuideOrientation, PowerPointRatio, PowerPointSlideViewInfo,
    PowerPointSlideViewInformation, PowerPointSlideViewPreferences, PowerPointViewKind,
    PowerPointViewOrigin, PowerPointZoomViewInfo,
};

// Re-export record types
pub use parsers::PptRecordParser;
pub use records::{DocumentInfo, PptRecord, SlideAtomsSet, SlideInfo};

// Re-export persist types
pub use persist::{PersistMapping, PersistPtrHolder};

// Re-export shape types
pub use shapes::{AutoShape, Placeholder, PlaceholderSize, PlaceholderType, Shape, TextBox};

// Re-export legacy types
pub use bookmark_summary::{PowerPointBookmark, PowerPointBookmarkSummary};
pub use broadcast::{PowerPointBroadcast, PowerPointBroadcastProperties, PowerPointBroadcasts};
pub use client_anchor::{
    OFFICE_ART_CLIENT_ANCHOR_RECORD_TYPE, PowerPointClientAnchor, PowerPointClientAnchorData,
    PowerPointClientAnchorEncoding, PowerPointClientAnchorLimits, PowerPointRect,
    PowerPointSmallRect,
};
pub use client_data::{
    OFFICE_ART_CLIENT_DATA_RECORD_TYPE, PowerPointClientData, PowerPointClientDataChild,
    PowerPointClientDataChildKind, PowerPointClientDataLimits,
};
pub use color_scheme::{
    PowerPointColorScheme, PowerPointColorSchemeAtom, PowerPointColorSchemeAtomKind,
    PowerPointSchemeColor,
};
pub use comments::{Author, Authors};
pub use current_user::CurrentUser;
pub use document_atom::{
    PowerPointDocumentAtom, PowerPointDocumentDimensions, PowerPointSlideSizeType,
};
pub use document_properties::{
    PowerPoint10DocumentProperties, PowerPoint12DocumentProperties, PowerPointCustomTableStyles,
    PowerPointGridSpacing, PowerPointPhotoAlbumFrameShape, PowerPointPhotoAlbumLayout,
    PowerPointPhotoAlbumSettings,
};
pub use document_structure::{PowerPointCustomTableStylesPlacement, PowerPointDocumentStructure};
pub use embedded::object::Editor;
pub use embedded::object::{
    Collection, ColorFollow, ContainerKind, Control, Definition, DimensionPolicy, DrawAspect,
    EmbedPreferences, ExternalObject, LinkInfo, Metadata, ObjectSubtype, ObjectType, UnknownRecord,
    UpdateMode,
};
pub use embedded::reference::{Reference, Target};
pub use envelope::PowerPointEnvelopeSettings;
pub use envelope_data::{
    MSO_ENVELOPE_CLSID, MsoAttachment, MsoEnvelope, MsoEnvelopeText, MsoEnvelopeVersion,
    MsoFollowUpStatus, MsoImportance, MsoPropertyValue, MsoRecipientCollection,
    MsoRecipientProperties, MsoRecipientProperty, MsoSecurityFlags, MsoSensitivity,
    PowerPointEnvelopeData, PowerPointEnvelopePayload,
};
pub use escher_textbox::EscherTextboxWrapper;
pub use external_media::{
    CdAudio, CdTime, EmbeddedWav, LinkedAudio, LinkedAudioKind, Media, Movie, MovieKind, Object,
    Video,
};
pub use font::{
    EmbeddedPowerPointFont, PowerPointFont, PowerPointFontCollection, PowerPointFontCollections,
    PowerPointFontEmbeddingFlags,
};
pub use header_footer::{
    PowerPointDateTimeFormatId, PowerPointHeaderFooter, PowerPointHeaderFooterDisplayText,
    PowerPointHeaderFooterOptions, PowerPointHeaderFooterParent,
    PowerPointHeaderFooterParentOrdinal, PowerPointHeaderFooterScope, PowerPointHeaderFooters,
    PowerPointScopedHeaderFooterDisplayText,
};
pub use html_publish::{
    PowerPointCodePage, PowerPointHtmlDocumentSettings, PowerPointHtmlPublishSettings,
    PowerPointWebFrameColors, PowerPointWebOutput, PowerPointWebScreenSize,
};
pub use hyperlink::{Hyperlink, HyperlinkExtension, Hyperlinks};
pub use hyperlink::{
    Interaction, InteractionAction, InteractionJump, InteractionLimits, InteractionLinkTarget,
    InteractionTrigger, InteractiveInfoAtom, MacroNameAtom, ShapeInteractionEntry,
};
pub use kinsoku::{
    BaseKinsokuSettings, KinsokuLanguage, KinsokuLevel, PowerPoint9KinsokuSettings,
    PowerPointKinsoku,
};
pub use main_master::{
    PowerPoint12MainMasterMetadata, PowerPointContentMasterInfo, PowerPointMainMasterTextStyles,
    PowerPointMainMasterTextStylesSource,
};
pub use master_style::{TextMasterStyle, TextMasterStyleLevel};
pub use modify_password::PowerPointModifyPassword;
pub use named_shows::{PowerPointNamedShow, PowerPointNamedShows};
pub use picture_bullets::{PictureBullet, PictureBulletCollection, PictureBulletType};
pub use placeholder_atom::{
    PowerPointPlaceholderAtom, PowerPointPlaceholderContext, PowerPointPlaceholderEntry,
    PowerPointPlaceholderKind, PowerPointPlaceholderLimits, PowerPointPlaceholderProjection,
    PowerPointPlaceholderSize, PowerPointPresentationPlaceholderEntry,
};
pub use print_options::{PowerPointPrintColorMode, PowerPointPrintOptions, PowerPointPrintTarget};
pub use privacy::PowerPointPrivacySettings;
pub use prog_tag_extensions::{
    PowerPoint9DocBinaryTagExtension, PowerPoint9SlideBinaryTagExtension,
    PowerPoint10DocBinaryTagExtension, PowerPoint10SlideBinaryTagExtension,
    PowerPoint11DocBinaryTagExtension, PowerPoint12DocBinaryTagExtension,
    PowerPoint12SlideBinaryTagExtension, PowerPointDocBinaryTagExtension,
    PowerPointDocumentTagExtensions, PowerPointSlideBinaryTagExtension,
    PowerPointSlideTagExtensions,
};
pub use prog_tags::{
    PowerPointProgBinaryTag, PowerPointProgBinaryTagVersion, PowerPointProgStringTag,
    PowerPointProgTag, PowerPointProgTagLimits, PowerPointProgTagScope, PowerPointProgTags,
};
pub use recolor::{
    PowerPointRecolorBitmapType, PowerPointRecolorBrush, PowerPointRecolorEntry,
    PowerPointRecolorHatch, PowerPointRecolorInfo, PowerPointRecolorLimits,
    PowerPointRecolorPattern, PowerPointRecolorSource, PowerPointWideColor,
    PowerPointWmfBrushStyle, PowerPointWmfHatchStyle,
};
pub use routing_slip::{Address, CurrentRecipient, Slip, Text};
pub use shape_flags::{
    PowerPointPresentationShapeFlagEntry, PowerPointShapeFlagEntry, PowerPointShapeFlagLimits,
    PowerPointShapeFlagProjection, PowerPointShapeFlags, PowerPointShapeFlags10,
};
pub use shape_programmable_tags::{
    PowerPointPresentationShapeProgrammableTagsEntry, PowerPointShapeBinaryTag,
    PowerPointShapeBinaryTagPayload, PowerPointShapeBinaryTagVersion,
    PowerPointShapeProgrammableTag, PowerPointShapeProgrammableTagLimits,
    PowerPointShapeProgrammableTags, PowerPointShapeProgrammableTagsEntry,
    PowerPointShapeStringTag, PowerPointShapeStyleAtom,
};
pub use slide_extension::{
    HeaderFooterDefaults, HeaderFooterPlaceholder, NewPlaceholder, ShapeChecksums, ShapeMetadata,
    SlideExtension,
};
pub use slide_round_trip::{
    PowerPoint12SlideRoundTripMetadata, PowerPointAnimationPackage, PowerPointColorMapping,
    PowerPointColorMappingKind, PowerPointColorMappingValues, PowerPointColorSchemeIndex,
    PowerPointContentMasterReference, PowerPointEmbeddedXmlPackage, PowerPointThemeKind,
    PowerPointThemePackage,
};
pub use slide_show_settings::{
    PowerPointColorIndex, PowerPointColorIndexKind, PowerPointSlideShowFlags,
    PowerPointSlideShowSettings,
};
pub use slide_sync::{PowerPointSlideSyncInfo, PowerPointSystemTime};
pub use smart_tags::{
    PowerPointSmartTag, PowerPointSmartTagProperty, PowerPointSmartTagStore, PowerPointSmartTagType,
};
pub use style_text_prop::{PowerPointStyleTextPropAtom, PowerPointTextCFRun, PowerPointTextPFRun};
pub use text_bookmark::PowerPointTextBookmark;
pub use text_extensions::{
    TextCharacterExtension9, TextCharacterExtension10, TextDefaultsExtension9,
    TextDefaultsExtension10, TextMasterStyleExtension9, TextMasterStyleExtension9Level,
    TextMasterStyleExtension10, TextParagraphExtension9, TextSpecialInfoExtension9,
    TextSpecialInfoExtension11, TextStyleExtension9, TextStyleExtension9Run, TextStyleExtension10,
    TextStyleExtension11, VersionedTextDefaults, VersionedTextMasterStyles,
};
pub use text_format_exception::{
    PowerPointBulletFlags, PowerPointCFStyle, PowerPointTextCFException, PowerPointTextPFException,
    PowerPointWrapFlags,
};
pub use text_interaction::{
    PowerPointShapeTextInteractionEntry, PowerPointTextBodyInteractions, PowerPointTextInteraction,
    PowerPointTextInteractionLimits, PowerPointTextRange, PowerPointTextType,
};
pub use text_metachar::{PowerPointMetacharKind, PowerPointTextMetachar};
pub use text_prop::{TextProp, TextPropCollection, TextPropType, TextTabStop};
pub use text_ruler::{TextRuler, TextRulerLevel, parse_default_text_ruler};
pub use text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphRun, ParagraphRunFormatting,
    ParagraphTabAlignment, ParagraphTabStop, ParagraphTextDirection, TextRun, TextRunExtractor,
    TextRunFormatting,
};
pub use text_si_exception::{
    PowerPointOutlineTextRef, PowerPointSpellingFlags, PowerPointTextSpecialInfoDefaults,
};
pub use text_special_info::{
    PowerPointMasterTextPropLevels, PowerPointMasterTextPropRun, PowerPointTextSIException,
    PowerPointTextSIRun, PowerPointTextSpecialInfoRuns,
};
pub use vba_info::{
    PowerPointVbaInfo, PowerPointVbaProjectCompression, PowerPointVbaProjectError,
    PowerPointVbaProjectLimits, PowerPointVbaProjectStorage,
};
pub use view_set_info::{
    PowerPointNormalViewSet, PowerPointNormalViewSetInfo, PowerPointNormalViewSetPayload,
    PowerPointNotesTextViewInfo, PowerPointViewBarState,
};

// Re-export writer types
pub use writer::{
    FreeformGeometry, GeometryRect, PowerPointSmartTagDefinition, PowerPointSmartTagIndex,
    PptWriteError, PptWriter, ShapePathType, ShapeProperties, ShapeType, TabAlign, TabStop,
    TextAlignment, TextDirection, TextFontAlign,
};

// Animation and transition support
pub mod animation;
pub mod transition;

// Re-export transition types for ergonomic read access
pub use animation::{EditorLimits, LegacyShapeAnimation, Scope, Timeline};
pub use transition::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
