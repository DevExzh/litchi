//! Typed gradient shade stops from `[MS-ODRAW]` `fillShadeColors`.
//!
//! The facade keeps this property family contextual: [`Stops`] borrows the
//! original `IMsoArray`, `codec` owns its fixed-width wire conversion, and
//! `validation` owns the cross-element ordering and range rules.  No color
//! resolution or rendering is performed.

mod codec;
mod model;
mod validation;

pub use codec::{encode, parse, parse_payload};
pub use model::{MAX_STOPS, Position, Stop, Stops};

#[cfg(test)]
mod tests;
