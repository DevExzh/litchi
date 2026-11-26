#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::{Convert, Options, decode_data, to_png, to_svg};
use litchi_odraw::image::{Blip, File, Kind};

fuzz_target!(|data: &[u8]| {
    let Ok(blip) = Blip::parse(data, common::art_limits()) else {
        return;
    };
    let limits = common::raster_limits();
    let _ = decode_data(&blip, &limits);
    let options = Options {
        limits,
        ..Options::default()
    };
    if let Ok(png) = to_png(&blip, options) {
        assert!(png.len() <= limits.max_output_bytes);
    }
    if let Ok(svg) = to_svg(&blip, options) {
        common::assert_well_formed_svg(svg.as_bytes(), limits.max_output_bytes);
    }

    // The file wrapper names the recommended representation. Extraction of
    // metafile and bitmap kinds repeats the conversions above, so only the
    // native pass-through kinds run it here.
    let file = File::new(blip, None, 0, common::art_limits());
    let name = file.out_name();
    assert!(!name.is_empty() && !name.contains(['/', '\\']));
    if matches!(file.kind(), Kind::Jpeg | Kind::CmykJpeg | Kind::Png)
        && let Ok(bytes) = file.extract(options)
    {
        assert!(bytes.len() <= limits.max_output_bytes);
    }
});
