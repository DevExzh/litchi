#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(pkg) = litchi_docx::Package::from_reader(std::io::Cursor::new(data)) {
        if let Ok(doc) = pkg.document() {
            let _ = doc.text();
        }
    }
});
