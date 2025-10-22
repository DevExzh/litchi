pub mod autoshape;
/// Shape and placeholder parsing for PowerPoint presentations.
///
/// This module provides functionality to parse shapes and placeholders
/// from PPT binary format, following the Apache POI HSLF structure.
///
/// # Architecture
///
/// The module is organized around these key types:
/// - `Shape`: Base trait for all shape types
/// - `TextBox`: Text box shapes
/// - `Placeholder`: Placeholder shapes for titles, content, etc.
/// - `AutoShape`: Auto shapes (rectangles, ovals, etc.)
/// - `PictureShape`: Picture shapes with embedded images
/// - `EscherRecord`: Parser for Escher binary records
///
/// # PPT Shape Structure
///
/// Shapes in PPT are stored in Escher format within the slide data.
/// Each shape has properties like position, size, text content, and formatting.
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let pres = pkg.presentation()?;
///
/// for slide in pres.slides()? {
///     for shape in slide.shapes()? {
///         match shape {
///             Shape::TextBox(textbox) => {
///                 println!("Text box: {}", textbox.text()?);
///             }
///             Shape::Placeholder(placeholder) => {
///                 println!("Placeholder type: {:?}", placeholder.placeholder_type());
///             }
///             _ => {}
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod escher;
pub mod geometry;
pub mod picture;
pub mod placeholder;
pub mod shape;
pub mod shape_enum;
pub mod textbox;

// Re-export the trait and type
pub use shape::{Shape, ShapeType};

// Re-export the high-performance enum
pub use shape_enum::ShapeEnum;

// Re-export concrete shape types
pub use autoshape::AutoShape;
pub use picture::PictureShape;
#[cfg(feature = "imgconv")]
pub use picture::extract_blip_id_from_escher;
pub use placeholder::{Placeholder, PlaceholderSize, PlaceholderType};
pub use textbox::TextBox;
