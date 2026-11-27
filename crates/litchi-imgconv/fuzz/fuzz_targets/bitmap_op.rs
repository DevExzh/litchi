#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::dib::{ColorUsage, DibLimits};
use litchi_imgconv::metafile_bitmap::{
    BitmapOp, BlitGrid, BrushPattern, RasterComposite, Rop2, StretchPolicy, parse_emf_object_dib,
    xor_bracket_paint,
};

/// EMF bitmap record types: `BITBLT`, `STRETCHBLT`, `MASKBLT`, `PLGBLT`,
/// `SETDIBITSTODEVICE`, `STRETCHDIBITS`, `ALPHABLEND` and `TRANSPARENTBLT`.
const EMF_BITMAP_RECORDS: [u32; 8] = [76, 77, 78, 79, 80, 81, 114, 116];

/// WMF bitmap functions: `BITBLT` and `STRETCHBLT` with and without a bitmap,
/// `SETDIBTODEV` and `STRETCHDIB`.
const WMF_BITMAP_FUNCTIONS: [u16; 6] = [0x0922, 0x0940, 0x0b23, 0x0b41, 0x0d33, 0x0f43];

fn dib_limits() -> DibLimits {
    DibLimits {
        max_input_bytes: 1 << 20,
        max_width: 256,
        max_height: 256,
        max_pixels: 256 * 256,
        max_palette_entries: 256,
        max_decoded_bytes: 4 << 20,
        max_output_bytes: 1 << 20,
    }
}

fn stretch_policy(selector: u8) -> StretchPolicy {
    match selector & 3 {
        0 => StretchPolicy::BlackOnWhite,
        1 => StretchPolicy::WhiteOnBlack,
        2 => StretchPolicy::ColorOnColor,
        _ => StretchPolicy::Halftone,
    }
}

fn exercise_operation(operation: &BitmapOp<'_>, pattern_color: u32) {
    let _ = operation.paint();
    let _ = operation.is_plain_blit();
    let _ = operation.destination_rect();
    let _ = operation.to_svg_image();
    let Ok(Some(pixels)) = operation.blit_pixels() else {
        return;
    };
    // Compose the decoded blit the way playback does for a destination-
    // dependent raster operation.
    let Ok(destination) = operation.destination_rect() else {
        return;
    };
    let destination = [
        f64::from(destination.x),
        f64::from(destination.y),
        f64::from(destination.width),
        f64::from(destination.height),
    ];
    let painted = pixels.window.painted(destination);
    let (columns, rows) = pixels.window.pixels();
    let Some(grid) = BlitGrid::for_pixels(painted, columns, rows) else {
        return;
    };
    if let Ok(mut composite) = RasterComposite::new(grid) {
        let _ = composite.apply(operation.rop3, painted, Some(&pixels.image), pattern_color);
        let _ = composite.to_svg_image(1 << 20);
    }
}

fn exercise_pattern(payload: &[u8], usage: ColorUsage, colors: [u32; 3]) {
    // The first sixteen bytes pick the bitmap-info and bitmap-bits ranges the
    // way an EMF brush record's offsets and sizes do.
    let Some((header, rest)) = payload.split_first_chunk::<16>() else {
        return;
    };
    let word = |index: usize| {
        u32::from_le_bytes([
            header[index * 4],
            header[index * 4 + 1],
            header[index * 4 + 2],
            header[index * 4 + 3],
        ])
    };
    let Ok(dib) = parse_emf_object_dib(
        rest,
        [word(0), word(1), word(2), word(3)],
        usage,
        dib_limits(),
    ) else {
        return;
    };
    let [text, background, xor] = colors;
    for pattern in [
        BrushPattern::from_dib(&dib),
        BrushPattern::from_monochrome_dib(&dib),
    ]
    .into_iter()
    .flatten()
    {
        let _ = pattern.is_monochrome();
        let _ = pattern.retained_bytes();
        if let Ok(png) = pattern.tile_png(text, background, 1 << 20) {
            let _ = BrushPattern::svg_definition("p", &png, (8.0, 8.0), (0.0, 0.0));
        }
        for mode in 1..=16 {
            if let Some(rop2) = Rop2::from_mode(mode) {
                let _ = pattern.under_rop2(rop2, text, background);
                let _ = xor_bracket_paint(rop2, text, xor);
            }
        }
    }
}

fuzz_target!(|data: &[u8]| {
    let Some((header, payload)) = data.split_first_chunk::<8>() else {
        return;
    };
    let selector = header[0];
    let stretch = stretch_policy(header[1]);
    let pattern_color = u32::from_le_bytes([header[2], header[3], header[4], 0]);
    let background = u32::from_le_bytes([header[5], header[6], header[7], 0]);
    let limits = dib_limits();

    match selector % 4 {
        0 => {
            let record = EMF_BITMAP_RECORDS[usize::from(selector / 4) % EMF_BITMAP_RECORDS.len()];
            if let Ok(operation) = BitmapOp::parse_emf(record, payload, stretch, limits) {
                exercise_operation(&operation, pattern_color);
            }
        },
        1 => {
            let function =
                WMF_BITMAP_FUNCTIONS[usize::from(selector / 4) % WMF_BITMAP_FUNCTIONS.len()];
            // A WMF record is whole 16-bit words; its size field counts the
            // six header bytes too.
            let params = &payload[..payload.len() & !1];
            let size_words = u32::try_from((params.len() + 6) / 2).unwrap_or(u32::MAX);
            if let Ok(operation) =
                BitmapOp::parse_wmf(size_words, function, params, stretch, limits)
            {
                exercise_operation(&operation, pattern_color);
            }
        },
        _ => {
            let usage = match selector / 4 % 3 {
                0 => ColorUsage::Rgb,
                1 => ColorUsage::PaletteEntries,
                _ => ColorUsage::PaletteIndices,
            };
            exercise_pattern(
                payload,
                usage,
                [pattern_color, background, pattern_color ^ background],
            );
        },
    }
});
