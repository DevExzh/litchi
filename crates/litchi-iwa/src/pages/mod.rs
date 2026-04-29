//! Pages Document Support
//!
//! This module provides comprehensive support for parsing Apple Pages documents,
//! including text extraction, section management, and document structure analysis.
//!
//! ## Features
//!
//! - Document metadata extraction
//! - Section and paragraph parsing
//! - Text style information
//! - Floating drawables (images, shapes)
//! - Header and footer extraction
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi::iwa::pages::PagesDocument;
//!
//! let doc = PagesDocument::open("document.pages")?;
//! let text = doc.text()?;
//! let sections = doc.sections()?;
//!
//! for section in sections {
//!     println!("Section: {:?}", section.heading);
//!     for para in &section.paragraphs {
//!         println!("  {}", para);
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod document;
pub mod section;

pub use document::PagesDocument;
pub use section::{PagesSection, PagesSectionType};
