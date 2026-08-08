//! Typed Keynote semantic models and bounded native package APIs.
//!
//! This crate owns the presentation vocabulary used by Keynote readers and
//! editors plus the concrete `.key` package boundary. Archive objects,
//! protobuf messages, native identifiers, and component names remain private.
//!
//! # Edit slide playback state
//!
//! Select slides by an exact navigator name or a checked semantic position;
//! native IWA identifiers never enter the public transaction.
//!
//! ```no_run
//! use std::fs::OpenOptions;
//! use std::io::Write as _;
//!
//! use litchi_keynote::{Package, SlideSelector};
//!
//! let package = Package::open("input.key")?;
//! let mut edit = package.edit();
//! edit.skip_slide(SlideSelector::name("Appendix"))?;
//! let commit = edit.commit()?;
//! let mut output = OpenOptions::new()
//!     .write(true)
//!     .create_new(true)
//!     .open("output.key")?;
//! output.write_all(commit.package().source_bytes())?;
//! output.sync_all()?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Reorder slides
//!
//! The destination is a checked final position; the immutable source remains
//! available and the commit carries an exact-source-checked inverse patch.
//!
//! ```no_run
//! use litchi_keynote::{Package, Position, SlideSelector};
//!
//! let package = Package::open("input.key")?;
//! let mut edit = package.edit_slide_order();
//! edit.move_slide(SlideSelector::name("Appendix"), Position::new(0))?;
//! let commit = edit.commit()?;
//! let restored = commit.package().apply_slide_order(&commit.patch().inverse())?;
//! assert_eq!(restored.package().source_bytes(), package.source_bytes());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Edit presentation settings
//!
//! Read or stage dimensions and playback behavior without native identifiers
//! or generated protobuf values. Commits retain an exact reversible patch.
//!
//! ```no_run
//! use litchi_keynote::{Mode, Package, Size};
//!
//! let package = Package::open("input.key")?;
//! let before = package.show_settings()?;
//! let mut edit = package.edit_show_settings()?;
//! edit.settings_mut().set_size(Size::new(1920.0, 1080.0)?);
//! edit.settings_mut().set_mode(Some(Mode::SelfPlaying))?;
//! let commit = edit.commit()?;
//! assert_eq!(commit.package().show_settings()?.mode(), Some(Mode::SelfPlaying));
//! let restored = commit
//!     .package()
//!     .apply_show_settings(&commit.patch().inverse())?;
//! assert_eq!(restored.package().show_settings()?, before);
//! assert_eq!(restored.package().source_bytes(), package.source_bytes());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

#![forbid(unsafe_code)]

mod error;
mod package;
mod selector;
mod time;

pub mod background;
pub mod build;
pub mod chart;
pub mod document;
pub mod show;
pub mod slide;
pub mod soundtrack;
pub mod transition;

pub use background::{Angle, Background, Gradient, Kind, Opaque, Stop};
pub use build::{AnimationType, Build};
pub use chart::ChartSelector;
pub use document::Document;
pub use error::{Error, Result};
pub use litchi_core::Position;
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__semantic_document_from_prepared_source;
pub use package::{
    Commit, Diagnostics, Edit, EditError, Limits, MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES,
    MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package, Patch, PayloadLimitKind,
    ReadError, ReadOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsError, SemanticPath,
    ShowSettingsCommit, ShowSettingsDiagnostics, ShowSettingsEdit, ShowSettingsError,
    ShowSettingsLimitKind, ShowSettingsPatch, SlideOrderCommit, SlideOrderDiagnostics,
    SlideOrderEdit, SlideOrderError, SlideOrderLimitKind, SlideOrderPatch, Stats,
    TextStorageFailure,
};
pub use selector::{SlideSelector, SlideSelectorError, SlideSelectorResult};
pub use show::{Mode, Settings, Show, Size};
pub use slide::media::MovieKind;
pub use slide::{Slide, Transition};
pub use time::Seconds;
pub use transition::Effect;
