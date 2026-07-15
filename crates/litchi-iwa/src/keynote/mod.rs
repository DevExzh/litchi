//! Keynote Presentation Support
//!
//! This module provides comprehensive support for parsing Apple Keynote presentations,
//! including slide extraction, build animations, and multimedia content.
//!
//! ## Features
//!
//! - Slide extraction with content
//! - Master slide identification
//! - Build animations and transitions
//! - Speaker notes
//! - Multimedia references
//!
//! ## Example
//!
//! ```rust,no_run
//! use litchi_iwa::keynote::KeynoteDocument;
//!
//! let doc = KeynoteDocument::open("presentation.key")?;
//! let slides = doc.slides()?;
//!
//! for slide in slides {
//!     if let Some(title) = &slide.title {
//!         println!("Slide {}: {}", slide.index + 1, title);
//!     }
//!     for text in &slide.text_content {
//!         println!("  - {}", text);
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod document;
pub mod editor;
pub mod show;
pub mod slide;

pub use document::KeynoteDocument;
pub use editor::{
    KeynoteBuildAcceleration, KeynoteBuildChunkInfo, KeynoteBuildCustomParameters,
    KeynoteBuildInfo, KeynoteBuildSettings, KeynoteBuildStart, KeynoteEditor,
    KeynoteEmphasisAction, KeynoteFlipDirection, KeynoteGradient, KeynoteGradientAngle,
    KeynoteGradientKind, KeynoteGradientStop, KeynoteHorizontalBuildDirection,
    KeynoteJiggleIntensity, KeynoteKeyboardBuild, KeynoteKeyboardDirection, KeynoteMotionPath,
    KeynoteMotionPathNode, KeynoteMotionPathNodeType, KeynoteMotionPathPoint, KeynoteMotionSubpath,
    KeynoteMoveAction, KeynoteObjectBuildEffect, KeynoteOpacityAction, KeynoteRgbColorSpace,
    KeynoteRgbaColor, KeynoteRotationAction, KeynoteRotationDirection, KeynoteScaleAction,
    KeynoteShowMode, KeynoteShowSettings, KeynoteSlideBackground, KeynoteSlideInfo,
    KeynoteSlideLayoutId, KeynoteSlideLayoutInfo, KeynoteSlideMovieInfo, KeynoteSlideMovieKind,
    KeynoteSlideTextInfo, KeynoteSlideTextPlaceholder, KeynoteSlideTextRole,
    KeynoteSoundtrackItemInfo, KeynoteSoundtrackMode, KeynoteSoundtrackSettings,
    KeynoteSwooshDirection, KeynoteTransitionAcceleration, KeynoteTransitionAnimationParameters,
    KeynoteTransitionCustomParameters, KeynoteTransitionDirection, KeynoteTransitionEffect,
    KeynoteTransitionMosaicType, KeynoteTransitionSettings, KeynoteTransitionTextDelivery,
    RemovedKeynoteSlideMovie, RemovedKeynoteTextBox,
};
pub use show::KeynoteShow;
pub use slide::{BuildAnimation, KeynoteSlide, SlideTransition};
