#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::pict::{PictConverter, PictParser, PictToRasterOptions};

fuzz_target!(|data: &[u8]| {
    let Ok(parser) = PictParser::new(data, &common::small_limits()) else {
        return;
    };
    let max_output_bytes = 1 << 20;
    let options = PictToRasterOptions {
        max_width: 256,
        max_height: 256,
        max_pixels: 256 * 256,
        max_output_bytes,
        ..PictToRasterOptions::default()
    };
    let converter = PictConverter::new(parser, options);
    if let Ok(png) = converter.convert_to_png() {
        assert!(png.len() <= max_output_bytes);
    }
});
