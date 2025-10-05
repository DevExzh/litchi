/// Parts for PowerPoint presentation documents.
///
/// This module contains wrapper types for different XML parts in a .pptx package,
/// following the structure of the python-pptx library.
pub mod presentation;
pub mod slide;

pub use presentation::PresentationPart;
pub use slide::{SlidePart, SlideLayoutPart, SlideMasterPart};

