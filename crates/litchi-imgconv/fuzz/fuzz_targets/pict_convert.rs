#![no_main]

#[path = "common/mod.rs"]
mod common;

use image::ImageFormat;
use libfuzzer_sys::fuzz_target;
use litchi_imgconv::Options;
use litchi_imgconv::pict::convert_pict_with_options;

fuzz_target!(|data: &[u8]| {
    let limits = common::raster_limits();
    // A requested width plays the picture at a scaled-down size instead of
    // its own; the input length decides which, so both stay reachable with
    // one playback per execution.
    let width = (data.len() % 2 == 1).then_some(64);
    let options = Options {
        width,
        limits,
        ..Options::default()
    };
    if let Ok(png) = convert_pict_with_options(data, ImageFormat::Png, options) {
        assert!(png.len() <= limits.max_output_bytes);
    }
});
