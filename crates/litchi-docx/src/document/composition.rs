//! Bounded document edit planning and history.
//!
//! Composition values retain planning state only. Publication belongs to an
//! eager or source-backed package, so a plan cannot publish a standalone
//! document snapshot by itself.

pub use super::transaction::{
    Composition, History, JoinError, Prepared, ThreeWayError, ThreeWayPlan,
};
pub use litchi_core::patch::{
    CompositionLimits, HistoryLimits, MergeChoice, SubEditConflict, SubEditJoinFailure,
    ThreeWayMergeFailure,
};
