#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fuzz_target!(|data: &[u8]| {
    let limits = common::raster_limits();
    let options = Options {
        limits,
        ..Options::default()
    };
    let png = convert_metafile(data, InputFormat::Emf, OutputFormat::Png, options);
    if let Ok(image) = png {
        common::assert_decodable_png(&image.bytes, &limits);
    }
});
