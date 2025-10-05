//! PowerPoint (.pptx) presentation support.
//!
//! This module provides parsing and manipulation of Microsoft PowerPoint presentations
//! in the Office Open XML (OOXML) format (.pptx files).
//!
//! The implementation follows the structure and API design of the python-pptx library,
//! adapted for Rust with performance optimizations and zero-copy parsing where possible.
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
//! use litchi::ooxml::pptx::Package;
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
//! use litchi::ooxml::pptx::Package;
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

pub mod package;
pub mod parts;
pub mod presentation;
pub mod shapes;
pub mod slide;

pub use package::Package;
pub use presentation::Presentation;
pub use slide::{Slide, SlideLayout, SlideMaster};
