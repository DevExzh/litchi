#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::pict::{PictConverter, PictParser, PictToRasterOptions};

fuzz_target!(|data: &[u8]| {
    let Ok(parser) = PictParser::new(data, &common::small_limits()) else {
        return;
    };
    let limits = litchi_imgconv::Limits {
        max_output_bytes: 1 << 20,
        ..common::raster_limits()
    };
    let options = PictToRasterOptions {
        max_width: limits.max_width,
        max_height: limits.max_height,
        max_pixels: limits.max_pixels,
        max_output_bytes: limits.max_output_bytes,
        ..PictToRasterOptions::default()
    };
    let converter = PictConverter::new(parser, options);
    if let Ok(png) = converter.convert_to_png() {
        common::assert_decodable_png(&png, &limits);
    }
});
