#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::{fuzz_mutator, fuzz_target};
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fuzz_target!(|data: &[u8]| {
    let limits = common::small_limits();
    let options = Options {
        limits,
        ..Options::default()
    };
    let first = convert_metafile(data, InputFormat::Emf, OutputFormat::Svg, options);
    if let Ok(svg) = &first {
        common::assert_well_formed_svg(&svg.bytes, limits.max_output_bytes);
    }
    // Conversion is a pure function of its input. Repeating it costs a second
    // playback, so only a quarter of the inputs check it.
    if data.len() % 4 == 0 {
        let second = convert_metafile(data, InputFormat::Emf, OutputFormat::Svg, options);
        match (&first, &second) {
            (Ok(first), Ok(second)) => {
                assert!(first.bytes == second.bytes && first.report == second.report)
            },
            (Err(first), Err(second)) => assert_eq!(first.to_string(), second.to_string()),
            _ => panic!("repeating a conversion changed its outcome"),
        }
    }
});

fuzz_mutator!(|data: &mut [u8], size: usize, max_size: usize, seed: u32| {
    common::mutate::mutate::<common::mutate::Emf>(data, size, max_size, seed)
});
