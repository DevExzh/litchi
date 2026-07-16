/// PowerPoint (.ppt) presentation support.
///
/// This module provides parsing of Microsoft PowerPoint presentations
/// in the legacy binary format (.ppt files), which uses OLE2 structured storage.
///
/// # Architecture
///
/// The module is organized around these key types:
/// - `Package`: The overall .ppt file package (OLE container)
/// - `Presentation`: The main presentation content and API
/// - `Slide`: Individual slide content and API
/// - `Shape`, `TextBox`, `Placeholder`: Shape and placeholder support
///
/// # PPT File Structure
///
/// A .ppt file is an OLE2 structured storage containing several streams:
/// - **PowerPoint Document**: Main presentation stream containing document properties
/// - **Pictures**: Embedded pictures and images
/// - **\x05SummaryInformation**: Document metadata
///
/// # Example
///
/// ```rust,no_run
/// use litchi_ole::ppt::{Package, shapes::ShapeEnum};
///
/// // Open a presentation
/// let mut package = Package::open("presentation.ppt")?;
/// let pres = package.presentation()?;
///
/// // Extract all text
/// let text = pres.text()?;
/// println!("Presentation text: {}", text);
///
/// // Access slides and shapes
/// for slide in pres.slides()? {
///     println!("Slide: {}", slide.text()?);
///
///     // Access individual shapes
///     for shape in slide.shapes()? {
///         match shape {
///             ShapeEnum::TextBox(textbox) => {
///                 println!("Text box: {}", textbox.text());
///             }
///             ShapeEnum::Placeholder(placeholder) => {
///                 println!("Placeholder type: {:?}", placeholder.placeholder_type());
///             }
///             _ => {}
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
// Core modules
pub mod package;
pub mod presentation;

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

// Drawing layer (Escher) support
pub mod escher;

// Legacy compatibility modules
pub mod comments;
pub mod current_user;
pub mod document_properties;
pub mod escher_textbox;
pub mod font;
pub mod hyperlink;
pub mod kinsoku;
pub mod master_style;
pub mod picture_bullets;
pub mod slide_extension;
pub mod slide_round_trip;
pub mod slide_sync;
pub mod smart_tags;
pub mod text_extensions;
pub mod text_prop;
pub mod text_ruler;
pub mod text_run;

// Re-export main types for convenience
pub use package::Package;
pub use presentation::Presentation;
pub use slide::{Slide, SlideData, SlideFactory};

// Re-export record types
pub use parsers::PptRecordParser;
pub use records::{DocumentInfo, PptRecord, SlideAtomsSet, SlideInfo};

// Re-export persist types
pub use persist::{PersistMapping, PersistPtrHolder};

// Re-export shape types
pub use shapes::{AutoShape, Placeholder, PlaceholderSize, PlaceholderType, Shape, TextBox};

// Re-export legacy types
pub use comments::{PowerPointCommentAuthor, PowerPointCommentAuthors};
pub use current_user::CurrentUser;
pub use document_properties::{
    PowerPoint10DocumentProperties, PowerPoint12DocumentProperties, PowerPointGridSpacing,
    PowerPointPhotoAlbumFrameShape, PowerPointPhotoAlbumLayout, PowerPointPhotoAlbumSettings,
};
pub use escher_textbox::EscherTextboxWrapper;
pub use font::{
    EmbeddedPowerPointFont, PowerPointFont, PowerPointFontCollection, PowerPointFontCollections,
    PowerPointFontEmbeddingFlags,
};
pub use hyperlink::{
    InteractionAction, InteractionJump, InteractionLinkTarget, InteractionTrigger,
    PowerPointInteraction,
};
pub use hyperlink::{PowerPointHyperlink, PowerPointHyperlinkExtension, PowerPointHyperlinks};
pub use kinsoku::{
    BaseKinsokuSettings, KinsokuLanguage, KinsokuLevel, PowerPoint9KinsokuSettings,
    PowerPointKinsoku,
};
pub use master_style::{TextMasterStyle, TextMasterStyleLevel};
pub use picture_bullets::{PictureBullet, PictureBulletCollection, PictureBulletType};
pub use slide_extension::{
    PowerPoint12PlaceholderMetadata, PowerPoint12ShapeMetadata, PowerPoint12SlideExtension,
    PowerPointHeaderFooterDefaults, PowerPointHeaderFooterPlaceholder, PowerPointNewPlaceholder,
    PowerPointShapeChecksums,
};
pub use slide_round_trip::{
    PowerPoint12SlideRoundTripMetadata, PowerPointAnimationPackage,
    PowerPointContentMasterReference,
};
pub use slide_sync::{PowerPointSlideSyncInfo, PowerPointSystemTime};
pub use smart_tags::{
    PowerPointSmartTag, PowerPointSmartTagProperty, PowerPointSmartTagStore, PowerPointSmartTagType,
};
pub use text_extensions::{
    TextCharacterExtension9, TextCharacterExtension10, TextDefaultsExtension9,
    TextDefaultsExtension10, TextMasterStyleExtension9, TextMasterStyleExtension9Level,
    TextMasterStyleExtension10, TextParagraphExtension9, TextSpecialInfoExtension9,
    TextSpecialInfoExtension11, TextStyleExtension9, TextStyleExtension9Run, TextStyleExtension10,
    TextStyleExtension11, VersionedTextDefaults, VersionedTextMasterStyles,
};
pub use text_prop::{TextProp, TextPropCollection, TextPropType, TextTabStop};
pub use text_ruler::{TextRuler, TextRulerLevel, parse_default_text_ruler};
pub use text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphRun, ParagraphRunFormatting,
    ParagraphTabAlignment, ParagraphTabStop, ParagraphTextDirection, TextRun, TextRunExtractor,
    TextRunFormatting,
};

// Re-export writer types
pub use writer::{
    FreeformGeometry, GeometryRect, PptWriteError, PptWriter, ShapePathType, ShapeProperties,
    ShapeType, TabAlign, TabStop, TextAlignment, TextDirection, TextFontAlign,
};

// Animation and transition support
pub mod animation;
pub mod transition;
