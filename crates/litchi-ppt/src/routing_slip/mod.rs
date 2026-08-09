//! Inert document routing-slip metadata from legacy `PowerPoint` files.
//!
//! The public model is kept separate from the strict binary codec. The model
//! uses names scoped by this module (`Slip`, `Address`, and `Text`), while the
//! codec owns the `[MS-PPT]` record layout and bounded validation. Undefined
//! bytes remain part of the model so an untouched routing slip can round-trip
//! without discarding source data or changing recipient order.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Address, CurrentRecipient, Slip, Text};
