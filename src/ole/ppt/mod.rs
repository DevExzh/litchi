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
pub mod package;
pub mod presentation;
pub mod shapes;
pub mod slide;
pub mod record_parser;

pub use package::Package;
pub use presentation::Presentation;
pub use shapes::{Shape, TextBox, Placeholder, PlaceholderType, PlaceholderSize, AutoShape};
pub use slide::Slide;
pub use record_parser::{PptRecordParser, DocumentInfo, SlideInfo};
