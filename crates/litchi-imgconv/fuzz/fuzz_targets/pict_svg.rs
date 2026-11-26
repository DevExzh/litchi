#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fuzz_target!(|data: &[u8]| {
    let limits = common::small_limits();
    let options = Options {
        limits,
        ..Options::default()
    };
    if let Ok(svg) = convert_metafile(data, InputFormat::Pict, OutputFormat::Svg, options) {
        common::assert_well_formed_svg(&svg.bytes, limits.max_output_bytes);
    }
});
