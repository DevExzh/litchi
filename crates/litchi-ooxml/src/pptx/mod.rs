//! PowerPoint (.pptx) presentation support.
//!
//! This module provides parsing and manipulation of Microsoft PowerPoint presentations
//! in the Office Open XML (OOXML) format (.pptx files).
//!
//! The concise facade delegates canonical semantic values to focused crates and
//! keeps package identities below safe, name-first APIs. Shapes are exposed as
//! borrowed data-bearing enums rather than copied XML plus a type discriminator.
//!
//! # Architecture
//!
//! The module is organized around these key types:
//! - `Package`: The overall .pptx file package (entry point)
//! - `Presentation`: The main presentation content and API
//! - `Slide`: Individual slide content and API
//! - `SlideMaster`: Slide master for themes and default formatting
//! - `SlideLayout`: Layout templates for slides
//! - `PresentationPart`, `SlidePart`, etc.: Lower-level part wrappers
//!
//! # Example: Reading a Presentation
//!
//! ```rust,no_run
//! use litchi_ooxml::pptx::Package;
//!
//! // Open a presentation
//! let pkg = Package::open("presentation.pptx")?;
//! let mut pres = pkg.presentation()?;
//!
//! // Get presentation info
//! println!("Slides: {}", pres.slide_count()?);
//! if let (Some(w), Some(h)) = (pres.slide_width()?, pres.slide_height()?) {
//!     println!("Slide size: {}x{} EMUs", w, h);
//! }
//!
//! // Access slides and extract text
//! for (idx, slide) in pres.slides()?.iter_mut().enumerate() {
//!     println!("\nSlide {}: {}", idx + 1, slide.name()?);
//!     println!("Content:\n{}", slide.text()?);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Example: Accessing Slide Masters
//!
//! ```rust,no_run
//! use litchi_ooxml::pptx::Package;
//!
//! let pkg = Package::open("presentation.pptx")?;
//! let mut pres = pkg.presentation()?;
//!
//! // Get slide masters
//! for master in pres.slide_masters()?.iter_mut() {
//!     println!("Master: {}", master.name()?);
//!     let layout_rids = master.slide_layout_rids()?;
//!     println!("  Has {} layouts", layout_rids.len());
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod actions;
mod animation_relationships;
pub mod animations;
pub mod backgrounds;
pub mod changes_information;
pub mod color_map;
pub mod comments;
pub mod controls;
pub mod customshow;
pub mod extended_guides;
pub mod format;
pub mod handout;
pub mod hyperlinks;
pub mod ink;
pub mod laser;
pub mod master_layout;
pub mod media;
pub mod media_parts;
mod namespace;
pub mod ole;
pub mod ole_object;
pub mod package;
pub mod parts;
pub mod presentation;
pub mod presentation_properties;
pub mod presentation_structure;
pub mod protection;
pub mod revision_information;
pub mod sections;
pub mod shape;
mod shapes;
pub mod show_events;
pub mod slide;
pub mod slide_sync;
pub mod smartart;
pub mod table;
pub mod template;
pub mod theme;
pub mod tracks;
pub mod vba_project;
pub mod view_properties;
pub mod writer;

pub use animations::{Direction, Effect, EffectInstance, Sequence};
pub use backgrounds::{GradientStop, GradientType, PatternType, PictureStyle, SlideBackground};
pub use changes_information::{
    CHANGES_INFORMATION_CONTENT_TYPE, CHANGES_INFORMATION_RELATIONSHIP_TYPE, ChangesData,
    ChangesInformation, ChangesInformationPart, ChangesNamespaceDeclaration,
    DocumentChangeDescriptor, DocumentChangeKind, DocumentChangesList, load_changes_information,
    store_changes_information,
};
pub use color_map::{ColorMap, ColorMapOverride, ColorMapSlot, ThemeColorRole};
pub use comments::{
    Author, Comment, Comments, Conformance, List, add_presentation_comment,
    add_presentation_comment_author, find_presentation_comment, find_presentation_comment_author,
    load_presentation_comments, parse_comment_authors, parse_slide_comments,
    remove_presentation_comment, remove_presentation_comment_author,
    reorder_presentation_comment_authors, reorder_presentation_comments,
    replace_presentation_comment, replace_presentation_comment_author, store_presentation_comments,
    update_presentation_comment, update_presentation_comment_author, write_comment_authors,
    write_slide_comments,
};
pub use controls::{
    ControlBinary, ControlDescriptor, Persistence as ControlPersistence, SlideControl,
};
pub use customshow::{CustomShow, CustomShowList};
pub use extended_guides::{
    ExtendedGuide, ExtendedGuideColor, ExtendedGuideColorKind, ExtendedGuideList,
    ExtendedGuideOrientation, PresentationExtendedGuides,
};
pub use format::{ImageFormat, TextFormat};
pub use handout::{HandoutHeaderFooter, HandoutLayout, HandoutMaster};
pub use hyperlinks::Hyperlink;
pub use ink::{Annotation, INK_CONTENT_TYPE, StoredInkAnnotation, store_slide_ink_annotation};
pub use litchi_drawingml::coord;
pub use litchi_pptx::actions::{Jump, Kind, Setting, Target, Trigger};
pub(crate) mod slide_patch;
pub use laser::store_slide_laser_trace;
pub use litchi_opc::FontEmbedding;
pub use litchi_pptx::font;
pub use litchi_pptx::laser::{LASER_TRACE_EXTENSION_URI, Trace, TracePoint};
pub use litchi_pptx::tag;
pub use litchi_pptx::time;
pub use master_layout::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderKind,
    PlaceholderSpec, SlideLayoutKind, add_slide_layout, add_slide_master, remove_slide_layout,
    store_placeholder_shape, validate_master_layout_graph,
};
pub use media::{Media, MediaFormat, MediaType};
pub use media_parts::{
    Bookmark, Data, ExtensionList, Fade, Picture, Poster, Resource, Transform, Trim,
    load_slide_media, parse_slide_media, store_slide_media, write_slide_media_pictures,
};
pub use ole::{Mode, Object, PayloadKind};
pub use ole_object::{AuthoredOleObject, OleObjectFrame, add_ole_object};
pub use package::Package;
pub use parts::{
    CHART_COLOR_STYLE_CONTENT_TYPE, CHART_COLOR_STYLE_RELATIONSHIP_TYPE, CHART_STYLE_CONTENT_TYPE,
    CHART_STYLE_RELATIONSHIP_TYPE, ChartColorStyleDocument, ChartColorStyleInfo,
    ChartColorStyleMethod, ChartColorStylePart, ChartData, ChartExAxis, ChartExAxisScaling,
    ChartExAxisTitle, ChartExAxisUnit, ChartExAxisUnits, ChartExAxisUnitsLabel, ChartExBinning,
    ChartExBinningChoice, ChartExChart, ChartExChartSpaceFormatting, ChartExChartTitle,
    ChartExClosedSide, ChartExColorKind, ChartExColorPosition, ChartExDataLabel,
    ChartExDataLabelPosition, ChartExDataLabelVisibility, ChartExDataLabels, ChartExDataPoint,
    ChartExDataSet, ChartExDimension, ChartExDocument, ChartExDoubleOrAutomatic,
    ChartExDrawingPayload, ChartExElementVisibility, ChartExExternalData,
    ChartExExternalDataTarget, ChartExFormatOverride, ChartExFormula, ChartExFormulaDirection,
    ChartExGeoAddress, ChartExGeoCache, ChartExGeoCacheEntry, ChartExGeoChildEntitiesQuery,
    ChartExGeoChildEntitiesQueryResult, ChartExGeoClear, ChartExGeoData, ChartExGeoDataEntityQuery,
    ChartExGeoDataEntityQueryResult, ChartExGeoDataPointQuery, ChartExGeoDataPointToEntityQuery,
    ChartExGeoDataPointToEntityQueryResult, ChartExGeoEntity, ChartExGeoEntityType,
    ChartExGeoHierarchyEntity, ChartExGeoLocation, ChartExGeoLocationQuery,
    ChartExGeoLocationQueryResult, ChartExGeoMappingLevel, ChartExGeoParentEntitiesQueryResult,
    ChartExGeoPolygon, ChartExGeoProjection, ChartExGeography, ChartExGridlines,
    ChartExHeaderFooter, ChartExInfo, ChartExLayoutProperties, ChartExLegend, ChartExNumberFormat,
    ChartExNumericDimensionType, ChartExNumericLevel, ChartExNumericPoint, ChartExOffset,
    ChartExPageMargins, ChartExPageOrientation, ChartExPageSetup, ChartExParentLabelLayout,
    ChartExPart, ChartExPlotArea, ChartExPlotAreaRegion, ChartExPlotSurface,
    ChartExPositionAlignment, ChartExPrintSettings, ChartExQuartileMethod,
    ChartExRegionLabelLayout, ChartExSeriesDataReference, ChartExSeriesLayout, ChartExSidePosition,
    ChartExSolidColor, ChartExStringDimensionType, ChartExStringLevel, ChartExStringPoint,
    ChartExText, ChartExTickLabels, ChartExTickMarkType, ChartExTickMarks,
    ChartExValueColorPositions, ChartExValueColors, ChartSeries, ChartStyleColor,
    ChartStyleColorKind, ChartStyleColorTransform, ChartStyleColorTransformKind,
    ChartStyleColorValue, ChartStyleDocument, ChartStyleEntry, ChartStyleEntryKind,
    ChartStyleFontIndex, ChartStyleFontReference, ChartStyleInfo, ChartStyleMarkerLayout,
    ChartStyleMarkerSymbol, ChartStylePart, ChartStylePayload, ChartStyleReference,
    ChartStyleVariation, ChartType,
};
pub use parts::{
    MasterVisibility, NotesSize, PhotoAlbumFrame, PhotoAlbumLayout, PresentationConformance,
    PresentationCustomerDataList, PresentationDefaultTextStyle, PresentationKinsokuSettings,
    PresentationMetadata, PresentationModificationVerifier, PresentationPhotoAlbum,
    SlideHeaderFooterVisibility, SlideLayoutMetadata, SlideLayoutReference, SlideMasterTextStyle,
    SlideMasterTextStyles, SlideSize,
};
pub use presentation::{Chart, Presentation};
pub use presentation_properties::{
    BrowserSupport, Color, ColorKind, Extension, HtmlPublish, HtmlTarget, OpaqueExtension, Print,
    PrintColorMode, PrintOutput, Properties, Show, ShowExtension, ShowMode, SlideSelection, Web,
    WebColor, WebScreenSize, load_from_package as load_presentation_properties,
};
pub use presentation_structure::{
    PresentationSlideReference, PresentationStructure, add_custom_show, add_custom_show_slide,
    add_section, add_section_slide, find_custom_show, find_section, load_presentation_structure,
    remove_custom_show, remove_custom_show_slide, remove_section, remove_section_slide,
    reorder_custom_show_slides, reorder_custom_shows, reorder_section_slides, reorder_sections,
    replace_custom_show, replace_section, store_presentation_structure,
    synchronize_presentation_structure_after_slide_mutation, update_custom_show, update_section,
};
pub use protection::{
    CryptoAlgorithm, ModifyVerifier, PresentationProtection, ProtectionType, SlideProtection,
};
pub use revision_information::{
    ClientRevision, REVISION_INFORMATION_CONTENT_TYPE, REVISION_INFORMATION_RELATIONSHIP_TYPE,
    RevisionInformation, RevisionInformationPart, RevisionNamespaceDeclaration,
    load_revision_information, store_revision_information,
};
pub use sections::{Section, SectionList};
pub use show_events::{
    Event, EventDraft, EventKind, SHOW_EVENT_EXTENSION_URI, store_slide_show_events,
};
pub use slide::{Slide, SlideLayout, SlideMaster};
pub use slide_sync::{
    SLIDE_SYNC_CONTENT_TYPE, SLIDE_SYNC_RELATIONSHIP_TYPE, SlideSyncDateTime, SlideSyncOffset,
    SlideSyncProperties, SlideSyncPropertiesPart, load_slide_sync_properties,
    store_slide_sync_properties,
};
pub use smartart::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};
pub use theme::{
    AuthoredTheme, SystemColorKind, ThemeColorScheme, ThemeColorSlot, ThemeColorValue,
    ThemeFontFace, ThemeFontScheme, ThemeOverride, ThemeScriptFont, add_theme,
    attach_theme_to_master, remove_theme_override, store_theme_color_scheme,
    store_theme_font_scheme, store_theme_override, theme_override, validate_theme_graph,
};
pub use vba_project::VbaProject;
pub use view_properties::{
    CommonSlideView, CommonView, GridSpacing, Guide, GuideOrientation, NormalView, OutlineSlide,
    OutlineView, Point, Ratio, RestoredPane, SimpleView, SlideLikeView, SorterView, SplitterState,
    ViewKind, ViewProperties, load_from_package as load_view_properties,
};
pub use writer::{MutablePresentation, MutableShape, MutableSlide};
