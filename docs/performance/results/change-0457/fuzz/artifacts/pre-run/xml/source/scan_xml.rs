#![no_main]

use std::io::{BufReader, Cursor};

use libfuzzer_sys::fuzz_target;
use litchi_odf_common::core::private::{XmlStreamLimits, scan_xml};

const MAX_INPUT: usize = 64 * 1024;

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT {
        return;
    }
    let limits = XmlStreamLimits::new(MAX_INPUT as u64, 64, 32_768, 16_384, 16_384)
        .expect("finite XML fuzz profile");
    let contiguous = scan_xml(Cursor::new(data), limits, |_, _| Ok(()));
    // One-byte reads exercise split UTF-8, references, names, and markup
    // delimiters. Provider chunking cannot change the logical XML contract.
    let fragmented = scan_xml(
        BufReader::with_capacity(1, Cursor::new(data)),
        limits,
        |_, _| Ok(()),
    );
    assert_eq!(contiguous.is_ok(), fragmented.is_ok());
    if let (Ok(contiguous), Ok(fragmented)) = (contiguous, fragmented) {
        assert_eq!(contiguous, fragmented);
        assert_eq!(contiguous.bytes(), data.len() as u64);
    }
});
