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
//! use litchi_keynote::Package;
//!
//! let doc = Package::open("presentation.key")?;
//! let slides = doc.slides()?;
//!
//! for slide in slides {
//!     if let Some(title) = slide.title() {
//!         println!("Slide {}: {}", slide.index() + 1, title);
//!     }
//!     for text in slide.text_content() {
//!         println!("  - {}", text);
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod creation;
pub mod editor;

pub use creation::KeynoteDocumentBuilder;
pub use editor::Effect;
pub use editor::{
    Angle, Background, BuildAcceleration, BuildStart, Gradient, KeynoteBuildChunkInfo,
    KeynoteBuildCustomParameters, KeynoteBuildInfo, KeynoteBuildSettings, KeynoteBuildTimingCurve,
    KeynoteEditor, KeynoteEmphasisAction, KeynoteFlipDirection, KeynoteHorizontalBuildDirection,
    KeynoteJiggleIntensity, KeynoteKeyboardBuild, KeynoteKeyboardDirection, KeynoteMotionPath,
    KeynoteMotionPathNode, KeynoteMotionPathNodeType, KeynoteMotionPathPoint, KeynoteMotionSubpath,
    KeynoteMoveAction, KeynoteObjectBuildEffect, KeynoteOpacityAction, KeynoteRotationAction,
    KeynoteRotationDirection, KeynoteScaleAction, KeynoteSlideAudioInfo, KeynoteSlideChartInfo,
    KeynoteSlideImageInfo, KeynoteSlideImageKind, KeynoteSlideInfo, KeynoteSlideLayoutId,
    KeynoteSlideLayoutInfo, KeynoteSlideMovieInfo, KeynoteSlideShapeInfo, KeynoteSlideTable,
    KeynoteSlideTableInfo, KeynoteSlideTextInfo, KeynoteSlideTextRole, KeynoteSoundtrackItemInfo,
    KeynoteSwooshDirection, KeynoteTableCellParagraphList, KeynoteTableCellParagraphListBullet,
    KeynoteTableCellParagraphListBulletGeometry, KeynoteTableCellParagraphListIndentation,
    KeynoteTableCellParagraphListLabelColor, KeynoteTableCellParagraphListLevel,
    KeynoteTableCellParagraphListLevelPlacement, KeynoteTableCellParagraphListNumberFormat,
    KeynoteTableCellParagraphListNumberScale, KeynoteTableCellParagraphListNumberTiering,
    KeynoteTableCellParagraphListNumbering, KeynoteTableCellParagraphListPlacement,
    KeynoteTableCellParagraphTabStops, KeynoteTableCellTextBackground,
    KeynoteTableCellTextBaselineShift, KeynoteTableCellTextCapitalization,
    KeynoteTableCellTextCharacterSpacing, KeynoteTableCellTextColor,
    KeynoteTableCellTextDecorations, KeynoteTableCellTextFont, KeynoteTableCellTextLigatures,
    KeynoteTableCellTextOutline, KeynoteTableCellTextScript, KeynoteTableCellTextShadow,
    KeynoteTableCellTextStyle, KeynoteTableCellUpdate, KeynoteTableCellValue,
    KeynoteTableDimension, KeynoteTableDimensionSize, KeynoteTablePoints,
    KeynoteTableTitleSettings, Kind, MovieKind, Opaque, RemovedKeynoteSlideAudio,
    RemovedKeynoteSlideChart, RemovedKeynoteSlideImage, RemovedKeynoteSlideMovie,
    RemovedKeynoteSlideShape, RemovedKeynoteSlideTable, RemovedKeynoteTextBox, RgbColorSpace, Rgba,
    Stop,
};
pub use litchi_keynote::Seconds;
pub use litchi_keynote::build::{AnimationType, Build};
pub use litchi_keynote::document::Document;
pub use litchi_keynote::show::{Mode, Settings, Show, Size};
pub use litchi_keynote::slide::{Slide, Transition};
pub use litchi_keynote::transition::{
    Acceleration, AccelerationKind, AnimationParameters, CustomParameters, Direction, MosaicType,
    TextDelivery, TextDeliveryKind,
};
