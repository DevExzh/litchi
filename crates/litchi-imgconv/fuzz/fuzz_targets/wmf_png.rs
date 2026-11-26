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
    let png = convert_metafile(data, InputFormat::Wmf, OutputFormat::Png, options);
    if let Ok(image) = png {
        assert!(image.bytes.len() <= limits.max_output_bytes);
    }
});
