//! Bounded native diagram inventory for legacy PowerPoint.
//!
//! A native PPT diagram is represented here as the relationship between a
//! `DiagramBuild` record from [MS-PPT] and an OfficeArt shape tree from
//! [MS-ODRAW]. The facade exposes identifiers, build metadata, associated
//! shapes, and borrowed inert record payloads. Its contextual slide owner can
//! publish only validated fixed-width build metadata; it does not calculate
//! layout, render a diagram, play its build, or author a complete SmartArt
//! object.
//!
//! The wire owners remain [`crate::animation::diagram_build`] and
//! [`crate::odraw`].  This module only validates their context and composes
//! their typed views.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes, parse_bytes_with_limits, parse_with_limits};
pub use model::{
    Build, Diagram, EditLimits, Id, Inventory, Limits, Payload, PayloadKind, ShapeRef,
};
pub use package::{
    SlideCommit, SlideEditor, SlideLimits, SlidePatch, SlideRevision, SlideSnapshot,
};
pub use transaction::{Change, Commit, Patch, Revision, Snapshot, Transaction};
