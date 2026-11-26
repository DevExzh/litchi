#![no_main]

use image::ImageFormat;
use libfuzzer_sys::fuzz_target;
use litchi_imgconv::pict::convert_pict;

fuzz_target!(|data: &[u8]| {
    let _ = convert_pict(data, ImageFormat::Png, None, None);
    let _ = convert_pict(data, ImageFormat::Png, Some(64), None);
});
