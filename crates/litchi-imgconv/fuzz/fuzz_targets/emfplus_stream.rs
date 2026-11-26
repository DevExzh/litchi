#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::emfplus::{
    EmfPlusRecordIter, ParserLimits, try_extract_emfplus_comment_body, validate_complete_stream,
};

fuzz_target!(|data: &[u8]| {
    let limits = ParserLimits {
        max_bytes: 1 << 20,
        max_records: 20_000,
        max_object_slots: 64,
    };
    let _ = validate_complete_stream(data, limits);
    let _ = try_extract_emfplus_comment_body(data, limits);

    let Ok(iter) = EmfPlusRecordIter::new(data, limits) else {
        return;
    };
    for parsed in iter {
        let Ok(record) = parsed else {
            break;
        };
        let _ = record.bytes();
        let _ = record.object_fragment(limits);
    }
});
