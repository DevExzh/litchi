#![no_main]

//! Apple iWork packages through the unified `File` (record 0306): open under
//! bounded options, then read the text, the outline and the stats, query
//! every element, and check that a no-op publication is the input itself.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi::view::Mode;
use litchi::{File, FileFormat, OpenOptions};

const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024;

fuzz_target!(|data: &[u8]| {
    let options = OpenOptions::new().max_input_bytes(MAX_INPUT_BYTES);
    let Ok(file) = litchi::open_bytes_with(data.to_vec(), &options) else {
        return;
    };
    let kinds: &[&str] = match file.format() {
        Some(FileFormat::Pages) => &["p", "section", "tbl", "tc", "header", "footer", "footnote"],
        Some(FileFormat::Keynote) => &["slide", "placeholder", "notes", "shape", "table"],
        Some(FileFormat::Numbers) => &["sheet", "table", "cell"],
        _ => return,
    };
    assert_reads(&file, kinds);
    if let Ok(bytes) = file.to_bytes() {
        assert_eq!(bytes, data, "a no-op publication is the input");
    }
});

fn assert_reads(file: &File, kinds: &[&str]) {
    black_box(file.text().ok());
    black_box(file.view(Mode::Outline).ok());
    black_box(file.view(Mode::Stats).ok());
    for kind in kinds {
        let Ok(nodes) = file.query(*kind) else {
            continue;
        };
        for node in nodes.take(256) {
            let Ok(node) = node else {
                break;
            };
            let path = node.path().to_string();
            black_box(node.view().text().ok());
            assert!(file.get(path.as_str()).is_ok(), "{path} resolves again");
        }
    }
}
