//! PowerPoint projections over format-neutral OfficeArt shapes.
//!
//! [MS-ODRAW] deliberately leaves ClientData and ClientTextbox payloads
//! to the host application. Keeping their interpretation here prevents the
//! shared drawing crate from acquiring a PowerPoint dependency while giving
//! PPT callers a concise, typed shape API.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{anchor, parse, text_from_drawing, text_from_textbox};
pub use model::{Anchor, FrameKind, Placeholder, ShapeExt};

pub(crate) use codec::textbox;
