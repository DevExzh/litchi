//! Bounded decoding and rendering for image payloads used by Litchi.
//!
//! OfficeArt BLIP and BStore grammar lives in `litchi-odraw`. This crate is
//! the optional codec layer: it consumes borrowed OfficeArt image views and
//! performs decompression, file adaptation, rasterization, and encoding.

#![allow(missing_docs)]
// `zerocopy` derives generate non-ASCII helper identifiers for the packed EMF
// record definitions. Rust only permits this lint allowance at crate scope.
#![allow(non_ascii_idents)]

mod codec;
pub mod dib;
pub mod emf;
pub mod emfplus;
mod file;
pub mod metafile_bitmap;
pub mod pict;
mod raster;
pub mod svg;
pub mod svg_utils;
pub mod wmf;

pub use codec::{
    ConversionDiagnostic, ConversionReport, ConvertedFormat, ConvertedImage, InputFormat, Limits,
    Options, OutputFormat, convert, convert_metafile, decode_data, to_jpeg, to_png, to_svg,
    to_webp,
};
pub use file::Convert;
