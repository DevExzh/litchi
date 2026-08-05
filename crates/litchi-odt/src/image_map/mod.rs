//! ODF `draw:image-map`: client-side image maps with clickable areas.
//!
//! An image map attaches link targets to geometric areas of an image frame.
//! Everything here is inert: link targets are stored verbatim and are never
//! resolved, followed, fetched, or rendered, and event-listener content is
//! preserved without interpretation.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_image_maps;
pub use model::{ImageMap, ImageMapArea, ImageMapAreaShape};
