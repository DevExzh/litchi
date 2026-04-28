//! OpenDocument Presentation (.odp) implementation.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating
//! OpenDocument Presentation documents (.odp files), which are the open standard
//! equivalent of Microsoft PowerPoint presentations.
//!
//! # Implementation Progress
//!
//! ## ✅ Reading (`presentation.rs`, `parser.rs`, `slide.rs`) - COMPLETE
//! - ✅ `Presentation::open()` - Load from file path
//! - ✅ `Presentation::from_bytes()` - Load from memory
//! - ✅ `slides()` - Get all slides
//! - ✅ `slide_count()` - Count slides
//! - ✅ `Slide::shapes()` - Get shapes on a slide
//! - ✅ `Slide::text()` - Extract text from slide
//! - ✅ `Slide::layout()` - Get slide layout name
//! - ✅ `Shape` parsing (text boxes, rectangles, ellipses, images)
//! - ✅ Master page parsing
//! - ✅ Metadata extraction
//! - ✅ Style parsing
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `PresentationBuilder::new()` - Create new presentations
//! - ✅ `add_slide()` - Add slides
//! - ✅ `add_shape()` - Add shapes (text boxes, rectangles, etc.)
//! - ✅ `set_slide_layout()` - Set slide layout
//! - ✅ `set_title()` / `set_author()` - Set metadata
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutablePresentation` - Modify existing presentations
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Slide transitions (fade, wipe, push, etc.)
//! - ⚠️ Animations (entrance, emphasis, exit, motion paths)
//! - ⚠️ Speaker notes (notes pages)
//! - ⚠️ Multimedia embedding (audio, video)
//! - ⚠️ Custom slide layouts
//! - ⚠️ Advanced shape properties (gradients, shadows, 3D effects)
//! - ⚠️ Connector lines (arrows between shapes)
//! - ⚠️ Slide master editing
//! - ⚠️ SmartArt/diagrams
//! - ⚠️ Slide timings
//! - ⚠️ Action buttons and hyperlinks
//! - ⚠️ Embedded charts
//! - ⚠️ Presentation protection
//!
//! # References
//! - ODF Specification: §10 (Presentation Content)
//! - odfpy: `odf/draw.py`, `odf/presentation.py`
//! - ODF Toolkit: Simple API - Presentation class

mod builder;
mod mutable;
mod parser;
mod presentation;
mod slide;

pub use builder::PresentationBuilder;
pub use mutable::MutablePresentation;
pub use presentation::Presentation;
pub use slide::{Shape, Slide};
