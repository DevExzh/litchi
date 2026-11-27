#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(package) = litchi_odt::Package::from_bytes(data) {
        if let Ok(doc) = package.document() {
            let _ = doc.text();
        }
    }
});
