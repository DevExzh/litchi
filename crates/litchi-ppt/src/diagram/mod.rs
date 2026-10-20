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
mod edit;
mod model;
mod package;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes};
pub use edit::Edit;
pub use model::{
    Build, Diagram, EditLimits, Id, Inventory, Limits, Payload, PayloadKind, ShapeRef,
};
pub use package::{SlideLimits, SlideRevision, SlideSnapshot};
pub use transaction::{Change, Revision, Snapshot};

crate::scoped::family!("diagram");
