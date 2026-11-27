#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_movie_caption_codec::{
    DecodeOptions, MovieCaptionWrite, decode_movie_caption, rewrite_movie_caption,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }
    let options = DecodeOptions::new(data.len().max(1), MAX_FIELDS, MAX_WORK_BYTES, 64)
        .with_max_output_bytes(MAX_OUTPUT_BYTES.max(data.len()));
    let Ok(snapshot) = decode_movie_caption(data, options) else {
        return;
    };
    black_box((snapshot.title_identifier(), snapshot.caption_identifier()));

    for identifier in [
        1_u64,
        0x1_0000_0000,
        u64::from(data.first().copied().unwrap_or(1)),
    ] {
        let before = data.to_vec();
        if let Ok(output) =
            rewrite_movie_caption(data, MovieCaptionWrite::title(identifier), options)
        {
            debug_assert_eq!(before, data);
            let readback =
                decode_movie_caption(&output, options.with_max_output_bytes(output.len()));
            let _ = black_box(readback);
        }
        if let Ok(output) =
            rewrite_movie_caption(data, MovieCaptionWrite::caption(identifier), options)
        {
            debug_assert_eq!(before, data);
            let readback =
                decode_movie_caption(&output, options.with_max_output_bytes(output.len()));
            let _ = black_box(readback);
        }
    }
});
