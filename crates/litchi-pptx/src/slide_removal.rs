//! Source-bound whole-slide removal plans and their durable patches.
//!
//! [`crate::opened::Snapshot::plan_slide_removal`] captures a complete-package
//! fingerprint and an exact durable [`Patch`] over only the presentation
//! owner and the selected slide, and returns a [`Plan`];
//! [`crate::opened::Edit::stage_slide_removal`] stages it in a primary edit.
//! [`crate::Package::apply_slide_removal`] replays a durable patch (or its
//! inverse) onto a package with the recorded revision.

pub use crate::opened::remove_plan::{Patch, Plan};
