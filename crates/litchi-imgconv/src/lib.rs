//! Bounded decoding and rendering for image payloads used by Litchi.
//!
//! OfficeArt BLIP and BStore grammar lives in `litchi-odraw`. This crate is
//! the optional codec layer: it consumes borrowed OfficeArt image views and
//! performs decompression, file adaptation, rasterization, and encoding.

#![allow(missing_docs)]

mod codec;
pub mod emf;
mod file;
pub mod pict;
pub mod svg;
pub mod svg_utils;
pub mod wmf;

pub use codec::{Limits, Options, convert, decode_data, to_jpeg, to_png, to_svg, to_webp};
pub use file::Convert;
