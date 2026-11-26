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
        assert!(svg.bytes.len() <= limits.max_output_bytes);
    }
    // The strict SVG helper rejects any picture that playback only approximates.
    let _ = litchi_imgconv::pict::convert_pict_to_svg_with_options(data, options);
    let _ = convert_metafile(data, InputFormat::Pict, OutputFormat::Auto, options);
});
