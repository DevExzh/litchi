//! Source-bound same-package whole-slide copy plans.
//!
//! [`crate::opened::Snapshot::plan_slide_copy`] validates and materializes
//! the complete candidate of a whole-slide copy without mutation and returns
//! a [`Plan`]; [`crate::opened::Edit::stage_slide_copy`] stages it in a
//! primary edit, whose commit publishes the new package. The plan's durable
//! patch replays through [`crate::Package::apply`].

pub use crate::opened::copy_plan::{Part, Plan};
