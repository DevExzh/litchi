#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Drives raw bytes through litchi-doc's .doc reader.
    if let Ok(mut pkg) = litchi_doc::Package::from_reader(std::io::Cursor::new(data)) {
        if let Ok(doc) = pkg.document() {
            let _ = doc.text();
            let _ = doc.paragraphs();
        }
    }
});
