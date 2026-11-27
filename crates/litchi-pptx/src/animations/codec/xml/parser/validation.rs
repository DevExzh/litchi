//! Parser-local validation seam.
//!
//! The parser keeps its bounds and namespace checks colocated while delegating
//! their implementation to the animation codec's shared validation owner.

pub(super) use super::super::super::validation::*;
