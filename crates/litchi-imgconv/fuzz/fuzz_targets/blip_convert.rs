#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::{Options, decode_data, to_png, to_svg};
use litchi_odraw::image::Blip;

fuzz_target!(|data: &[u8]| {
    let Ok(blip) = Blip::parse(data, common::art_limits()) else {
        return;
    };
    let limits = common::small_limits();
    let _ = decode_data(&blip, &limits);
    let options = Options {
        limits,
        ..Options::default()
    };
    if let Ok(png) = to_png(&blip, options) {
        assert!(png.len() <= limits.max_output_bytes);
    }
    if let Ok(svg) = to_svg(&blip, options) {
        assert!(svg.len() <= limits.max_output_bytes);
    }
});
