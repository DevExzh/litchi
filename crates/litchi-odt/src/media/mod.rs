//! Image media discovery for packaged and flat OpenDocument families.

mod codec;
mod model;

pub use model::{Image, ImageFrame, ImagePart, ImageSource};

pub use codec::{scan_content_images, scan_flat_images, scan_packaged_images};
