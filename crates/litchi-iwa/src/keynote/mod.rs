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

pub use creation::Builder;
pub use editor::Effect;
pub use editor::{
    BuildAcceleration, BuildChunkInfo, BuildCustomParameters, BuildInfo, BuildSettings, BuildStart,
    BuildTimingCurve, Editor, EmphasisAction, FlipDirection, HorizontalBuildDirection,
    JiggleIntensity, KeyboardBuild, KeyboardDirection, MotionPath, MotionPathNode,
    MotionPathNodeType, MotionPathPoint, MotionSubpath, MoveAction, MovieKind, ObjectBuildEffect,
    OpacityAction, RemovedKeynoteSlideChart, RemovedKeynoteSlideImage, RemovedKeynoteSlideShape,
    RemovedKeynoteSlideTable, RemovedKeynoteTextBox, RotationAction, RotationDirection,
    ScaleAction, SlideChartInfo, SlideImageInfo, SlideImageKind, SlideInfo, SlideLayoutId,
    SlideLayoutInfo, SlideShapeInfo, SlideTable, SlideTableInfo, SlideTextInfo, SlideTextRole,
    SwooshDirection, TableCellParagraphList, TableCellParagraphListBullet,
    TableCellParagraphListBulletGeometry, TableCellParagraphListIndentation,
    TableCellParagraphListLabelColor, TableCellParagraphListLevel,
    TableCellParagraphListLevelPlacement, TableCellParagraphListNumberFormat,
    TableCellParagraphListNumberScale, TableCellParagraphListNumberTiering,
    TableCellParagraphListNumbering, TableCellParagraphListPlacement, TableCellParagraphTabStops,
    TableCellTextBackground, TableCellTextBaselineShift, TableCellTextCapitalization,
    TableCellTextCharacterSpacing, TableCellTextColor, TableCellTextDecorations, TableCellTextFont,
    TableCellTextLigatures, TableCellTextOutline, TableCellTextScript, TableCellTextShadow,
    TableCellTextStyle, TableCellUpdate, TableCellValue, TableDimension, TableDimensionSize,
    TablePoints,
};
pub use litchi_keynote::build::{AnimationType, Build};
pub use litchi_keynote::show::{Mode, Settings, Show, Size};
pub use litchi_keynote::slide::{Slide, Transition};
pub use litchi_keynote::transition::{
    Acceleration, AccelerationKind, AnimationParameters, CustomParameters, Direction, MosaicType,
    TextDelivery, TextDeliveryKind,
};
pub use litchi_keynote::{Seconds, SlideSelector, TableSelector};
