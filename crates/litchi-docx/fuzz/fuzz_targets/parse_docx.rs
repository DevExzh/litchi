#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(pkg) = litchi_docx::Package::from_reader(std::io::Cursor::new(data)) {
        if let Ok(doc) = pkg.document() {
            let _ = doc.text();
            if let Ok(blocks) = doc.body_blocks() {
                for block in blocks.iter().take(64) {
                    let _ = block.transaction_position();
                    let _ = block.self_contained_xml();
                    if let Some(control) = block.content_control() {
                        let _ = control.tag();
                        let _ = control.alias();
                        let _ = control.text();
                        let _ = control.blocks();
                    }
                }
            }
        }
    }
});
