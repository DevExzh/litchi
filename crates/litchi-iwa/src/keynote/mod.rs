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

pub mod creation;
pub mod document;
pub mod editor;
pub mod show;
pub mod slide;

pub use creation::KeynoteDocumentBuilder;
pub use document::KeynoteDocument;
pub use editor::{
    KeynoteBuildAcceleration, KeynoteBuildChunkInfo, KeynoteBuildCustomParameters,
    KeynoteBuildInfo, KeynoteBuildSettings, KeynoteBuildStart, KeynoteBuildTimingCurve,
    KeynoteEditor, KeynoteEmphasisAction, KeynoteFlipDirection, KeynoteGradient,
    KeynoteGradientAngle, KeynoteGradientKind, KeynoteGradientStop,
    KeynoteHorizontalBuildDirection, KeynoteJiggleIntensity, KeynoteKeyboardBuild,
    KeynoteKeyboardDirection, KeynoteMotionPath, KeynoteMotionPathNode, KeynoteMotionPathNodeType,
    KeynoteMotionPathPoint, KeynoteMotionSubpath, KeynoteMoveAction, KeynoteObjectBuildEffect,
    KeynoteOpacityAction, KeynoteRgbColorSpace, KeynoteRgbaColor, KeynoteRotationAction,
    KeynoteRotationDirection, KeynoteScaleAction, KeynoteShowMode, KeynoteShowSettings,
    KeynoteSlideAudioInfo, KeynoteSlideAudioOptions, KeynoteSlideBackground, KeynoteSlideChartInfo,
    KeynoteSlideImageInfo, KeynoteSlideImageKind, KeynoteSlideImageOptions, KeynoteSlideInfo,
    KeynoteSlideLayoutId, KeynoteSlideLayoutInfo, KeynoteSlideMovieInfo, KeynoteSlideMovieKind,
    KeynoteSlideMovieOptions, KeynoteSlideShapeInfo, KeynoteSlideShapeKind, KeynoteSlideTable,
    KeynoteSlideTableInfo, KeynoteSlideTextInfo, KeynoteSlideTextPlaceholder, KeynoteSlideTextRole,
    KeynoteSoundtrackItemInfo, KeynoteSoundtrackMode, KeynoteSoundtrackSettings,
    KeynoteSwooshDirection, KeynoteTableAxisIndex, KeynoteTableCellCheckboxFormat,
    KeynoteTableCellCurrencyFormat, KeynoteTableCellDataFormat, KeynoteTableCellDateTimeFormat,
    KeynoteTableCellDecimalPlaces, KeynoteTableCellDurationFormat, KeynoteTableCellDurationStyle,
    KeynoteTableCellDurationUnit, KeynoteTableCellDurationUnitRange, KeynoteTableCellDurationUnits,
    KeynoteTableCellFixedDecimalPlaces, KeynoteTableCellFractionFormat, KeynoteTableCellInset,
    KeynoteTableCellInsets, KeynoteTableCellLayout, KeynoteTableCellNegativeNumberStyle,
    KeynoteTableCellNumberFormat, KeynoteTableCellNumeralSystemFormat,
    KeynoteTableCellParagraphIndents, KeynoteTableCellParagraphLineSpacing,
    KeynoteTableCellParagraphList, KeynoteTableCellParagraphListBullet,
    KeynoteTableCellParagraphListBulletGeometry, KeynoteTableCellParagraphListIndentation,
    KeynoteTableCellParagraphListLabelColor, KeynoteTableCellParagraphListLevel,
    KeynoteTableCellParagraphListLevelPlacement, KeynoteTableCellParagraphListNumberFormat,
    KeynoteTableCellParagraphListNumberTiering, KeynoteTableCellParagraphListNumbering,
    KeynoteTableCellParagraphListPlacement, KeynoteTableCellParagraphSpacing,
    KeynoteTableCellParagraphTabStops, KeynoteTableCellPercentageFormat,
    KeynoteTableCellPopUpMenuFormat, KeynoteTableCellPopUpMenuInitialSelection,
    KeynoteTableCellPopUpMenuItem, KeynoteTableCellRegion, KeynoteTableCellScientificFormat,
    KeynoteTableCellSliderDisplayFormat, KeynoteTableCellSliderFormat, KeynoteTableCellSliderRange,
    KeynoteTableCellStarRatingFormat, KeynoteTableCellStepperDisplayFormat,
    KeynoteTableCellStepperFormat, KeynoteTableCellStepperRange, KeynoteTableCellTextAlignment,
    KeynoteTableCellTextBackground, KeynoteTableCellTextBaselineShift,
    KeynoteTableCellTextCapitalization, KeynoteTableCellTextCharacterSpacing,
    KeynoteTableCellTextColor, KeynoteTableCellTextDecorations, KeynoteTableCellTextFont,
    KeynoteTableCellTextFormat, KeynoteTableCellTextLigatures, KeynoteTableCellTextOutline,
    KeynoteTableCellTextScript, KeynoteTableCellTextShadow, KeynoteTableCellTextStyle,
    KeynoteTableCellTextWrap, KeynoteTableCellThousandsSeparator, KeynoteTableCellUpdate,
    KeynoteTableCellValue, KeynoteTableCellVerticalAlignment, KeynoteTableColumnDeletion,
    KeynoteTableColumnInsertion, KeynoteTableDimension, KeynoteTableDimensionSize,
    KeynoteTableFormulaAxisReference, KeynoteTableFormulaBinaryOperator,
    KeynoteTableFormulaCachedValue, KeynoteTableFormulaCellReference,
    KeynoteTableFormulaExpression, KeynoteTableHeaderCount, KeynoteTableHeaderSettings,
    KeynoteTableHiddenAxes, KeynoteTablePoints, KeynoteTableRowDeletion, KeynoteTableRowInsertion,
    KeynoteTableSortColumnIndex, KeynoteTableSortDirection, KeynoteTableSortOrder,
    KeynoteTableSortRowRange, KeynoteTableSortRule, KeynoteTableSortScope,
    KeynoteTableTitleSettings, KeynoteTransitionAcceleration, KeynoteTransitionAnimationParameters,
    KeynoteTransitionCustomParameters, KeynoteTransitionDirection, KeynoteTransitionEffect,
    KeynoteTransitionMosaicType, KeynoteTransitionSettings, KeynoteTransitionTextDelivery,
    RemovedKeynoteSlideAudio, RemovedKeynoteSlideChart, RemovedKeynoteSlideImage,
    RemovedKeynoteSlideMovie, RemovedKeynoteSlideShape, RemovedKeynoteSlideTable,
    RemovedKeynoteTextBox,
};
pub use show::KeynoteShow;
pub use slide::{BuildAnimation, KeynoteSlide, SlideTransition};
