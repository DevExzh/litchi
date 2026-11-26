#![no_main]

#[path = "common/mod.rs"]
mod common;

use image::ImageFormat;
use libfuzzer_sys::fuzz_target;
use litchi_imgconv::pict::convert_pict_with_options;
use litchi_imgconv::Options;

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
    // A requested width plays the picture at a scaled-down size instead.
    let _ = convert_pict_with_options(
        data,
        ImageFormat::Png,
        Options {
            width: Some(64),
            ..options
        },
    );
});
