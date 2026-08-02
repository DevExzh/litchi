//! Pictures, shapes, embedded objects, and legacy drawing records.

pub(crate) mod legacy_drawing;
pub(crate) mod legacy_text_box;
pub(crate) mod object;
pub(crate) mod page_border;
pub mod picture;
pub(crate) mod picture_compatibility;
pub mod shape;

pub(crate) use crate::{border, types};
