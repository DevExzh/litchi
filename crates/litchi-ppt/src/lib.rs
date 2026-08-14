//! `PowerPoint` (`.ppt`) presentation support.
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
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "modules and re-exports are grouped by functional area rather than by item kind"
)]

mod consts;
mod officeart_wire;

pub use consts::RecordType;
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
/// Source-checked reordering of slides in an opened legacy presentation.
pub mod slide_order;

// Submodules (organized by functionality)
pub mod parsers;
pub mod persist;
pub mod records;
pub mod shapes;
pub mod text;
/// Source-checked editing of text in an existing parsed shape.
pub mod text_edit;

/// Semantic owners for content embedded in a legacy presentation.
pub mod embedded;

// PowerPoint projections over the format-neutral OfficeArt substrate.
pub mod odraw;

// PowerPoint host metadata over format-neutral Office Graph views.
pub mod chart;

/// Bounded native diagram inventory over MS-PPT build metadata and OfficeArt shapes.
pub mod diagram;

// Format-specific semantic owners
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
#[cfg(feature = "encryption")]
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
/// Contextual inventory of the legacy presentation's main, title, notes, and
/// handout masters.
pub mod master;
/// Lossless, snapshot-isolated authoring of one contextual master layout.
pub mod master_layout;
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
#[cfg(feature = "vba-inspection")]
pub mod vba_info;
pub mod view_set_info;

/// Bounded, non-mutating semantic validation of legacy PowerPoint sources.
pub mod validation;

// Re-export main types for convenience
pub use document_comparison::{
    DiffFlags, DiffNode, DiffRecordHeaders, DiffTree10, DiffType, DocDiffFlags, ElementType,
    MainMasterDiffFlags, ReviewingToolbarStates, ShapeDiffFlags, SlideCreationEntry,
    SlideDiffFlags, SlideListTable10, TableDiffFlags, TextDiffFlags,
};
#[cfg(feature = "encryption")]
pub use encryption::EncryptionProfile;
pub use non_zoom_view::{
    NoZoomViewInfo, NonZoomViewKind, OutlineSorterViewInfo, OutlineSorterViewInformation,
};
pub use package::{EncryptionKind, Error, OpenOptions, Package, RecordLimits, SourceBackedPackage};
pub use presentation::{ParsedCustomShow, ParsedSlideComments, Presentation};
pub use presentation_advisor::{AdvisorRule, PresentationAdvisorSettings};
pub use slide::{
    ParsedComment, ParsedSlideTiming, Slide, SlideData, SlideDirectory, SlideDirectoryEntry,
    SlideFactory, SpeakerNotes,
};
pub use sound_collection::{BuiltinId, Sound};
pub use validation::{
    PptValidationError, PptValidationLimits, validate_source, validate_source_with_limits,
};
pub use view_info::{
    Guide, GuideOrientation, Ratio, SlideViewInfo, SlideViewInformation, SlideViewPreferences,
    ViewKind, ViewOrigin, ZoomViewInfo,
};

// Re-export record types
pub use parsers::RecordParser;
pub use records::{DocumentInfo, Record, SlideAtomsSet, SlideInfo};

// Re-export persist types
pub use persist::{PersistMapping, PersistPtrHolder};

// Re-export shape types
pub use shapes::{AutoShape, Placeholder, PlaceholderSize, PlaceholderType, Shape, TextBox};

// Re-export legacy types
pub use bookmark_summary::{Bookmark, Summary};
pub use broadcast::{Broadcast, BroadcastProperties, Broadcasts};
pub use client_anchor::Anchor;
pub use client_data::{
    ClientData, ClientDataChild, ClientDataChildKind, ClientDataLimits,
    OFFICE_ART_CLIENT_DATA_RECORD_TYPE,
};
pub use color_scheme::{ColorScheme, ColorSchemeAtom, ColorSchemeAtomKind, SchemeColor};
pub use comments::{Author, Authors, Catalog};
pub use current_user::CurrentUser;
pub use document_atom::{DocumentAtom, DocumentDimensions, SlideSizeType};
pub use document_properties::{
    CustomTableStyles, DocumentProperties10, DocumentProperties12, GridSpacing,
    PhotoAlbumFrameShape, PhotoAlbumLayout, PhotoAlbumSettings,
};
pub use document_structure::{CustomTableStylesPlacement, DocumentStructure};
pub use embedded::object::Editor;
pub use embedded::object::{
    Collection, ColorFollow, ContainerKind, Control, Definition, DimensionPolicy, DrawAspect,
    EmbedPreferences, ExternalObject, LinkInfo, Metadata, ObjectSubtype, ObjectType, UnknownRecord,
    UpdateMode,
};
pub use embedded::reference::{Reference, Target};
pub use envelope::EnvelopeSettings;
pub use envelope_data::{
    EnvelopeData, EnvelopePayload, MSO_ENVELOPE_CLSID, MsoAttachment, MsoEnvelope, MsoEnvelopeText,
    MsoEnvelopeVersion, MsoFollowUpStatus, MsoImportance, MsoPropertyValue, MsoRecipientCollection,
    MsoRecipientProperties, MsoRecipientProperty, MsoSecurityFlags, MsoSensitivity,
};
pub use escher_textbox::EscherTextboxWrapper;
pub use external_media::{
    CdAudio, CdTime, EmbeddedWav, LinkedAudio, LinkedAudioKind, Media, Movie, MovieKind, Object,
    Video,
};
pub use font::{
    Commit as FontCommit, EmbeddedFont, EotMetadata, Facet as FontFacet, Font, FontCollection,
    FontCollections, FontEmbeddingFlags, Limits as FontLimits, PackageLimits as FontPackageLimits,
    PackageOptions as FontPackageOptions, Patch as FontPatch, Revision as FontRevision,
    Scope as FontScope, Snapshot as FontSnapshot, Transaction as FontTransaction,
};
pub use header_footer::{
    DateTimeFormatId, HeaderFooter, HeaderFooterDisplayText, HeaderFooterOptions,
    HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope, HeaderFooters,
    ScopedHeaderFooterDisplayText,
};
pub use html_publish::{
    CodePage, HtmlDocumentSettings, HtmlPublishSettings, WebFrameColors, WebOutput, WebScreenSize,
};
pub use hyperlink::{Hyperlink, HyperlinkExtension, Hyperlinks};
pub use hyperlink::{
    Interaction, InteractionAction, InteractionJump, InteractionLimits, InteractionLinkTarget,
    InteractionTrigger, InteractiveInfoAtom, MacroNameAtom, ShapeInteractionEntry,
};
pub use kinsoku::{BaseKinsokuSettings, Kinsoku, KinsokuLanguage, KinsokuLevel, KinsokuSettings9};
pub use main_master::{
    ContentMasterInfo, MainMasterMetadata12, MainMasterTextStyles, MainMasterTextStylesSource,
};
pub use master_style::{TextMasterStyle, TextMasterStyleLevel};
pub use modify_password::ModifyPassword;
pub use named_shows::{NamedShow, NamedShows};
pub use picture_bullets::{PictureBullet, PictureBulletCollection, PictureBulletType};
pub use placeholder_atom::{
    AtomPlaceholderSize, PlaceholderAtom, PlaceholderContext, PlaceholderEntry, PlaceholderKind,
    PlaceholderLimits, PlaceholderProjection, PresentationPlaceholderEntry,
};
pub use print_options::{PrintColorMode, PrintOptions, PrintTarget};
pub use privacy::PrivacySettings;
pub use prog_tag_extensions::{
    DocBinaryTagExtension, DocBinaryTagExtension9, DocBinaryTagExtension10,
    DocBinaryTagExtension11, DocBinaryTagExtension12, DocumentTagExtensions,
    SlideBinaryTagExtension, SlideBinaryTagExtension9, SlideBinaryTagExtension10,
    SlideBinaryTagExtension12, SlideTagExtensions,
};
pub use prog_tags::{
    ProgBinaryTag, ProgBinaryTagVersion, ProgStringTag, ProgTag, ProgTagLimits, ProgTagScope,
    ProgTags,
};
pub use recolor::{
    RecolorBitmapType, RecolorBrush, RecolorEntry, RecolorHatch, RecolorInfo, RecolorLimits,
    RecolorPattern, RecolorSource, WideColor, WmfBrushStyle, WmfHatchStyle,
};
pub use routing_slip::{Address, CurrentRecipient, Slip, Text};
pub use shape_flags::{
    PresentationShapeFlagEntry, ShapeFlagEntry, ShapeFlagLimits, ShapeFlagProjection, ShapeFlags,
    ShapeFlags10,
};
pub use shape_programmable_tags::{
    PresentationShapeProgrammableTagsEntry, ShapeBinaryTag, ShapeBinaryTagPayload,
    ShapeBinaryTagVersion, ShapeProgrammableTag, ShapeProgrammableTagLimits, ShapeProgrammableTags,
    ShapeProgrammableTagsEntry, ShapeStringTag, ShapeStyleAtom,
};
pub use slide_extension::{
    HeaderFooterDefaults, HeaderFooterPlaceholder, NewPlaceholder, ShapeChecksums, ShapeMetadata,
    SlideExtension,
};
pub use slide_round_trip::{
    AnimationPackage, ColorMapping, ColorMappingKind, ColorMappingValues, ColorSchemeIndex,
    ContentMasterReference, EmbeddedXmlPackage, SlideRoundTripMetadata12, ThemeKind, ThemePackage,
};
pub use slide_show_settings::{ColorIndex, ColorIndexKind, SlideShowFlags, SlideShowSettings};
pub use slide_sync::{
    Change as SlideSyncChange, ChangeSet as SlideSyncChangeSet, Commit as SlideSyncCommit,
    Editor as SlideSyncEditor, LibraryUrl, Limits as SlideSyncLimits, MAX_TEXT_BYTES,
    Revision as SlideSyncRevision, ServerId, Snapshot as SlideSyncSnapshot, Synchronization,
    SystemTime,
};
pub use smart_tags::{SmartTag, SmartTagProperty, SmartTagStore, SmartTagType};
pub use style_text_prop::{StyleTextPropAtom, TextCFRun, TextPFRun};
pub use text_bookmark::TextBookmark;
pub use text_extensions::{
    TextCharacterExtension9, TextCharacterExtension10, TextDefaultsExtension9,
    TextDefaultsExtension10, TextMasterStyleExtension9, TextMasterStyleExtension9Level,
    TextMasterStyleExtension10, TextParagraphExtension9, TextSpecialInfoExtension9,
    TextSpecialInfoExtension11, TextStyleExtension9, TextStyleExtension9Run, TextStyleExtension10,
    TextStyleExtension11, VersionedTextDefaults, VersionedTextMasterStyles,
};
pub use text_format_exception::{
    BulletFlags, CFStyle, TextCFException, TextPFException, WrapFlags,
};
pub use text_interaction::{
    ShapeTextInteractionEntry, TextBodyInteractions, TextInteraction, TextInteractionLimits,
    TextRange, TextType,
};
pub use text_metachar::{MetacharKind, TextMetachar};
pub use text_prop::{TextProp, TextPropCollection, TextPropType, TextTabStop};
pub use text_ruler::{TextRuler, TextRulerLevel, parse_default_text_ruler};
pub use text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphRun, ParagraphRunFormatting,
    ParagraphTabAlignment, ParagraphTabStop, ParagraphTextDirection, TextRun, TextRunExtractor,
    TextRunFormatting,
};
pub use text_si_exception::{OutlineTextRef, SpellingFlags, TextSpecialInfoDefaults};
pub use text_special_info::{
    MasterTextPropLevels, MasterTextPropRun, TextSIException, TextSIRun, TextSpecialInfoRuns,
};
#[cfg(feature = "vba-inspection")]
pub use vba_info::{
    VbaInfo, VbaProjectCompression, VbaProjectError, VbaProjectLimits, VbaProjectStorage,
};
pub use view_set_info::{
    NormalViewSet, NormalViewSetInfo, NormalViewSetPayload, NotesTextViewInfo, ViewBarState,
};

// Re-export writer types
pub use writer::{
    FreeformGeometry, GeometryRect, ShapePathType, ShapeProperties, ShapeType, SmartTagDefinition,
    SmartTagIndex, TabAlign, TabStop, TextAlignment, TextDirection, TextFontAlign, WriteError,
    Writer,
};

// Animation and transition support
pub mod animation;
pub mod transition;

// Re-export transition types for ergonomic read access
pub use animation::{EditorLimits, LegacyShapeAnimation, Scope, Timeline};
pub use transition::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
