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
//! commit.package().write_to(&mut output)?;
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
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Edit presentation settings
//!
//! Read or stage dimensions and playback behavior without native identifiers
//! or generated protobuf values. Commits retain an exact reversible patch.
//!
//! ```no_run
//! use litchi_keynote::{Package, show::{Mode, Size}};
//!
//! let package = Package::open("input.key")?;
//! let before = package.show_settings()?;
//! let edit = package.edit_show_settings()?;
//! let mut settings = edit.settings();
//! settings.set_size(Size::new(1920.0, 1080.0)?);
//! settings.set_mode(Some(Mode::SelfPlaying))?;
//! let commit = edit.set(settings).commit()?;
//! assert_eq!(commit.package().show_settings()?.mode(), Some(Mode::SelfPlaying));
//! let restored = commit
//!     .package()
//!     .apply_show_settings(&commit.patch().inverse())?;
//! assert_eq!(restored.package().show_settings()?, before);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Edit one slide transition
//!
//! Read, replace, or clear a modern transition with the same selector-first,
//! exact-source transaction model. Clearing writes Keynote's native no-effect
//! transition; it does not synthesize a transition envelope for a legacy-only
//! slide. Patches are reversible against the exact committed package bytes.
//!
//! ```no_run
//! use std::io;
//!
//! use litchi_keynote::{Effect, Package};
//!
//! let package = Package::open("input.key")?;
//! let before = package
//!     .slide_transition("Appendix")?
//!     .ok_or_else(|| io::Error::other("slide has no modern transition"))?;
//! let mut replacement = before.clone();
//! replacement.set_effect(Some(Effect::Dissolve))?;
//!
//! let mut edit = package.edit_slide_transition("Appendix")?;
//! edit.set_transition(replacement)?;
//! let commit = edit.commit()?;
//! assert_eq!(
//!     commit.package().slide_transition("Appendix")?,
//!     commit.patch().after().cloned(),
//! );
//!
//! let restored = commit
//!     .package()
//!     .apply_slide_transition(&commit.patch().inverse())?;
//! assert_eq!(restored.package().slide_transition("Appendix")?, Some(before));
//!
//! let mut clear = restored.package().edit_slide_transition("Appendix")?;
//! clear.clear()?;
//! let cleared = clear.commit()?;
//! assert_eq!(cleared.package().slide_transition("Appendix")?.unwrap().effect(), Some(&Effect::None));
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Edit speaker notes
//!
//! Existing notes storage can be edited by an exact slide name or checked
//! position with notes-relative UTF-16 spans. A missing notes graph is kept
//! distinct from an existing empty storage and is not synthesized implicitly.
//!
//! ```no_run
//! use std::io;
//!
//! use litchi_keynote::{Package, TextSpan};
//!
//! let package = Package::open("input.key")?;
//! let before = package
//!     .slide_notes("Appendix")?
//!     .ok_or_else(|| io::Error::other("slide has no existing notes storage"))?;
//! let mut edit = package.edit_slide_notes("Appendix")?;
//! edit.replace(TextSpan::from_utf16_indexes(0, 0)?, "Draft: ")?;
//! let commit = edit.commit()?;
//! assert!(commit.package().slide_notes("Appendix")?.unwrap().starts_with("Draft: "));
//!
//! let restored = commit
//!     .package()
//!     .apply_slide_notes(&commit.patch().inverse())?;
//! assert_eq!(restored.package().slide_notes("Appendix")?, Some(before));
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Edit slide title or body text
//!
//! Title and body placeholders are distinct semantic roles. Only an existing,
//! exclusively owned placeholder storage is edited; this API never exposes or
//! synthesizes native object identifiers. An absent role placeholder is `None`,
//! while `Some("")` denotes an existing placeholder with empty storage.
//!
//! ```no_run
//! use std::io;
//!
//! use litchi_keynote::{Package, TextSpan};
//!
//! let package = Package::open("input.key")?;
//! let title_before = package
//!     .slide_title("Appendix")?
//!     .ok_or_else(|| io::Error::other("slide has no existing title storage"))?;
//! let body_before = package
//!     .slide_body("Appendix")?
//!     .ok_or_else(|| io::Error::other("slide has no existing body storage"))?;
//!
//! let mut title = package.edit_slide_title("Appendix")?;
//! title.replace(TextSpan::from_utf16_indexes(0, 0)?, "Draft: ")?;
//! let title_commit = title.commit()?;
//! let mut body = title_commit.package().edit_slide_body("Appendix")?;
//! body.set("Draft body")?;
//! let body_commit = body.commit()?;
//! assert!(body_commit.package().slide_title("Appendix")?.unwrap().starts_with("Draft: "));
//! assert_eq!(body_commit.package().slide_body("Appendix")?, Some("Draft body".to_owned()));
//!
//! let title_restored = body_commit
//!     .package()
//!     .apply_slide_text(&body_commit.patch().inverse())?;
//! let restored = title_restored
//!     .package()
//!     .apply_slide_text(&title_commit.patch().inverse())?;
//! assert_eq!(restored.package().slide_title("Appendix")?, Some(title_before));
//! assert_eq!(restored.package().slide_body("Appendix")?, Some(body_before));
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
pub use litchi_iwa_text::{TextPosition, TextSpan};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__semantic_document_from_prepared_source;
pub use package::{
    Commit, Diagnostics, Edit, EditError, Limits, MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES,
    MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package, Patch, PayloadLimitKind,
    ReadError, ReadOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsError, SemanticPath,
    SlideNotesCommit, SlideNotesDiagnostics, SlideNotesEdit, SlideNotesError, SlideNotesLimitKind,
    SlideNotesPatch, SlideOrderCommit, SlideOrderDiagnostics, SlideOrderEdit, SlideOrderError,
    SlideOrderLimitKind, SlideOrderPatch, SlideTextCommit, SlideTextDiagnostics, SlideTextEdit,
    SlideTextError, SlideTextLimitKind, SlideTextPatch, SlideTextRole, SlideTransitionCommit,
    SlideTransitionDiagnostics, SlideTransitionEdit, SlideTransitionError,
    SlideTransitionLimitKind, SlideTransitionPatch, Stats, TextStorageFailure, WriteError,
};
pub use selector::{SlideSelector, SlideSelectorError, SlideSelectorResult};
pub use slide::media::MovieKind;
pub use slide::{Slide, Transition};
pub use time::Seconds;
pub use transition::Effect;
