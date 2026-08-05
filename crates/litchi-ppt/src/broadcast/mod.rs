//! Inert PowerPoint 9 presentation-broadcast metadata from MS-PPT 2.4.17.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Broadcast, BroadcastProperties, Broadcasts};
