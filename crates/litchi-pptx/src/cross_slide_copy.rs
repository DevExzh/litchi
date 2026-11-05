//! Source-bound cross-presentation slide-copy plans and their durable
//! patches.
//!
//! [`crate::opened::Snapshot::plan_cross_slide_copy`] proves a slide's
//! dependency closure from a source presentation into a destination and
//! returns a [`Plan`]; [`crate::opened::Edit::stage_cross_slide_copy`]
//! stages it in a primary edit of the destination.
//! [`crate::Package::apply_cross_slide_copy`] replays a durable [`Patch`]
//! (or its inverse) onto a destination with the recorded revisions.

pub use crate::opened::cross_copy_plan::{Patch, Plan};
