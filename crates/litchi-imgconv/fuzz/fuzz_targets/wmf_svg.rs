#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fuzz_target!(|data: &[u8]| {
    let _ = convert_metafile(data, InputFormat::Wmf, OutputFormat::Svg, Options::default());
});
