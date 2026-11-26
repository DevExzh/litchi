#![no_main]

#[path = "common/mod.rs"]
mod common;

use image::ImageFormat;
use libfuzzer_sys::fuzz_target;
use litchi_imgconv::pict::convert_pict_with_options;
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fuzz_target!(|data: &[u8]| {
    let limits = common::small_limits();
    let options = Options {
        limits,
        ..Options::default()
    };
    let png = convert_pict_with_options(data, ImageFormat::Png, options);
    if let Ok(png) = png {
        assert!(png.len() <= limits.max_output_bytes);
    }
    // A requested width exercises the scaled-down raster path of the same
    // strict helper.
    let _ = convert_pict_with_options(
        data,
        ImageFormat::Png,
        Options {
            width: Some(64),
            ..options
        },
    );
    let _ = convert_metafile(data, InputFormat::Pict, OutputFormat::Auto, options);
});
