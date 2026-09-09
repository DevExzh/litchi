//! Compatibility re-exports for the neutral table coordinate vocabulary.
//!
//! The coordinate types are owned by `litchi-iwa-common` so all concrete
//! iWork format owners can share the same compact, archive-free selectors.
//! This module preserves the existing Numbers paths while the migration
//! completes.

pub use litchi_iwa_common::table::coordinate::{
    AddressError, CellPosition, CellRange, Error, Result,
};
