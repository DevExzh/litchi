//! Unified PowerPoint presentation module.
//!
//! This module provides a unified API for working with PowerPoint presentations in both
//! legacy (.ppt) and modern (.pptx) formats. The format is automatically detected
//! and handled transparently.
//!
//! # Architecture
//!
//! The module provides a format-agnostic API following the python-pptx design:
//! - `Presentation`: The main presentation API (auto-detects format)
//! - `Slide`: Individual slide with shapes and content
//! - `Shape`: Shape elements on slides
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::Presentation;
//!
//! // Open any PowerPoint presentation (.ppt or .pptx) - format auto-detected
//! let pres = Presentation::open("presentation.ppt")?;
//!
//! // Extract all text
//! let text = pres.text()?;
//! println!("Presentation text: {}", text);
//!
//! // Get slide count
//! let count = pres.slide_count()?;
//! println!("Slides: {}", count);
//!
//! // Access slides
//! for (i, slide) in pres.slides()?.iter().enumerate() {
//!     println!("Slide {}: {}", i + 1, slide.text()?);
//! }
//! # Ok::<(), litchi::common::Error>(())
//! ```

// Submodule declarations
mod types;
mod prs;
mod slide;

// Re-exports
pub use prs::Presentation;
pub use slide::Slide;
