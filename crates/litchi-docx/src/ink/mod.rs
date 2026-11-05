//! Bounded, inert InkML annotation metadata in relationship-owned Word stories.
//!
//! [`crate::Package::ink`] returns an immutable semantic inventory. Shared
//! payloads are parsed once and retained without copying their OPC allocation.
//! Generic non-Ink content parts are not annotations. Recognition, rendering,
//! trace-coordinate interpretation are not implemented.
//! [`crate::Package::edit_ink`] supports source-checked insertion, replacement and
//! removal. Detached [`Draft`] values prepare bounded X/Y integer InkML; package
//! commits allocate resource names and relationships. Drawing hosts require a
//! caller-supplied image fallback. Existing hosts with unmodeled dependencies
//! refuse edits.
//! This is not a generic content-part or Word product-compatibility validator.
//! Orphan Ink parts and parts referenced only by inactive or foreign hosts are
//! outside this active-annotation inventory.
//!
//! ```
//! use litchi_docx::ink::{
//!     BaseProfile, BrushDraft, ContextDraft, ContextKind, Destination, Draft,
//!     Style, TraceDraft,
//! };
//! use litchi_core::Position;
//!
//! # fn insert(package: &litchi_docx::Package) -> Result<(), Box<dyn std::error::Error>> {
//! let payload = Draft::default()
//!     .context(ContextDraft::new(ContextKind::InkDrawing).with_xml_id("ctx")?)?
//!     .brush(BrushDraft::new("brush")?)?
//!     .trace(TraceDraft::new("0 0, 10 20")?
//!         .with_context_ref("#ctx")?.with_brush_ref("#brush")?)?
//!     .finish()?;
//! let mut edit = package.edit_ink()?;
//! edit.insert(Destination::main(Position::new(0)), payload,
//!     Style::Base(BaseProfile::InkContent))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! # Ok(())
//! # }
//! ```

mod authoring;
mod codec;
mod graph;
mod host;
mod limit;
mod model;
pub(crate) mod package;
mod placement;
mod trace;
mod transaction;
mod xml;

pub use authoring::{
    AnchorGeometry, BaseProfile, FallbackImage, Geometry, HorizontalAlignment, HorizontalPosition,
    HorizontalRelativeFrom, ImageDimensions, ImageType, Placement, Point, Style, VerticalAlignment,
    VerticalPosition, VerticalRelativeFrom, WrapMode,
};
pub use litchi_drawingml::ink::builder::{
    AuthoringLimits, BrushDraft, BrushPropertyDraft, ContextDraft, Draft, TraceDraft,
};
pub use litchi_drawingml::ink::{
    BrushPropertyName, ContextKind, EffectiveBrushProperty, Guid, InkEffect, Prepared, SemanticType,
};
pub use model::{Annotation, BrushProperty, Context, Limits, Location, Snapshot};
pub use transaction::{Commit, Destination, Diagnostics, Edit, EditOptions, Patch};

/// Error of the ink family: the crate error. Ink budget refusals are
/// [`crate::Error::InkLimit`], which carries the exceeded [`LimitKind`].
pub type Error = crate::Error;

pub use limit::LimitKind;

pub(crate) const CONTENT_TYPE: &str = "application/inkml+xml";
