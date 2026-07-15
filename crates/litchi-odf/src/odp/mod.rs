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
//! - ✅ Speaker notes
//! - ✅ Slide transition and automatic-timing style resolution
//! - ✅ Inert ODF/SMIL timing trees and legacy presentation effects
//! - ✅ Inert audio/video plugin references and parameters
//! - ✅ Inert shape hyperlinks, presentation actions, and script bindings
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `PresentationBuilder::new()` - Create new presentations
//! - ✅ `add_slide()` - Add slides
//! - ✅ `add_shape()` - Add shapes (text boxes, rectangles, etc.)
//! - ✅ `set_slide_layout()` - Set slide layout
//! - ✅ `set_title()` / `set_author()` - Set metadata
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutablePresentation` - Modify existing presentations
//! - ✅ Slide transitions, timings, and transition sounds
//! - ✅ Modern and legacy animation-tree creation and namespace-preserving round trips
//! - ✅ Package-contained audio/video embedding and mutable preservation
//! - ✅ Shape hyperlink/action creation and inert round trips
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Custom slide layouts
//! - ⚠️ Advanced shape properties (gradients, shadows, 3D effects)
//! - ⚠️ Slide master editing
//! - ⚠️ SmartArt/diagrams
//! - ⚠️ Embedded charts
//! - ⚠️ Presentation protection
//!
//! # References
//! - ODF Specification: §10 (Presentation Content)
//! - odfpy: `odf/draw.py`, `odf/presentation.py`
//! - ODF Toolkit: Simple API - Presentation class

mod action;
mod animation;
mod builder;
mod legacy_animation;
mod media;
mod mutable;
mod parser;
mod presentation;
mod slide;
mod transition;

pub use action::{
    DrawingHyperlink, HyperlinkShow, PresentationAction, PresentationEffect,
    PresentationEffectDirection, PresentationEventListener, ScriptEventListener,
    ShapeEventListener,
};
pub use animation::{
    AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
};
pub use builder::PresentationBuilder;
pub use legacy_animation::{LegacyAnimationKind, LegacyAnimationNode};
pub use media::{MediaActuate, MediaParameter, MediaReference, MediaShow};
pub use mutable::MutablePresentation;
pub use presentation::Presentation;
pub use slide::{DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, Shape, Slide};
pub use transition::{
    SlideTransition, TransitionDirection, TransitionSound, TransitionSoundShow, TransitionSpeed,
    TransitionStyle, TransitionType,
};
