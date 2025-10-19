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
/// use litchi::ppt::Package;
///
/// // Open a presentation
/// let package = Package::open("presentation.ppt")?;
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
///             litchi::ppt::shapes::Shape::TextBox(textbox) => {
///                 println!("Text box: {}", textbox.text()?);
///             }
///             litchi::ppt::shapes::Shape::Placeholder(placeholder) => {
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

/// Slide module with factory and enhanced implementation
pub mod slide;

// Submodules (organized by functionality)
pub mod records;
pub mod parsers;
pub mod persist;
pub mod text;
pub mod shapes;

// Drawing layer (Escher) support
pub mod escher;

// Legacy compatibility modules
pub mod current_user;
pub mod text_run;
pub mod text_prop;
pub mod escher_textbox;

// Re-export main types for convenience
pub use package::Package;
pub use presentation::Presentation;
pub use slide::{Slide, SlideFactory, SlideData};

// Re-export record types
pub use records::{PptRecord, DocumentInfo, SlideInfo, SlideAtomsSet};
pub use parsers::PptRecordParser;

// Re-export persist types
pub use persist::{PersistPtrHolder, PersistMapping};

// Re-export shape types
pub use shapes::{Shape, TextBox, Placeholder, PlaceholderType, PlaceholderSize, AutoShape};

// Re-export legacy types
pub use current_user::CurrentUser;
pub use text_run::{TextRun, TextRunExtractor, TextRunFormatting};
pub use text_prop::{TextProp, TextPropCollection, TextPropType};
pub use escher_textbox::EscherTextboxWrapper;
