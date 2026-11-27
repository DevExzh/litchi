#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::dib::{AlphaInterpretation, ColorUsage, Dib, DibLimits};

fn limits() -> DibLimits {
    DibLimits {
        max_input_bytes: 1 << 20,
        max_width: 256,
        max_height: 256,
        max_pixels: 256 * 256,
        max_palette_entries: 4096,
        max_decoded_bytes: 1 << 20,
        max_output_bytes: 1 << 20,
    }
}

fuzz_target!(|data: &[u8]| {
    let Some((&selector, rest)) = data.split_first() else {
        return;
    };
    let limits = limits();
    let usage = ColorUsage::from_raw(u32::from(selector & 3)).unwrap_or_default();
    let alpha = match (selector >> 2) & 3 {
        0 => AlphaInterpretation::Ignore,
        1 => AlphaInterpretation::Straight,
        _ => AlphaInterpretation::Premultiplied,
    };

    if let Ok(dib) = Dib::parse_with_usage(rest, usage, limits) {
        let _ = dib.to_bmp();
        let _ = dib.to_dynamic_image();
        let _ = dib.to_dynamic_image_with_alpha(alpha);
        if let Ok(png) = dib.to_png() {
            assert!(png.len() <= limits.max_output_bytes);
        }
    }

    let parts = rest.len() / 2;
    let _ = Dib::parse_parts(&rest[..parts], &rest[parts..], usage, limits);
});
