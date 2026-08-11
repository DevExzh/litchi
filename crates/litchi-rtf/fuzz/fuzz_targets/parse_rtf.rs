#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = litchi_rtf::decompress(data);
    if let Ok(doc) = litchi_rtf::Document::from_bytes(data) {
        let _ = doc.text();
        let mut paragraphs = doc.body().paragraphs();
        let selected = data
            .first()
            .copied()
            .map_or(0, |value| usize::from(value) % 32);
        if paragraphs.nth(selected).is_some() {
            let _ = paragraphs.next();
        }
    }
});
