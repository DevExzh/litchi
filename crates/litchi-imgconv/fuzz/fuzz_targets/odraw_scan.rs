#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_odraw::Parser;
use litchi_odraw::image::{Id, all, delay, get, record, scan, store};

fuzz_target!(|data: &[u8]| {
    let limits = common::art_limits();
    let _ = scan(data, limits);
    let _ = delay(data, limits);

    if let Ok(Some(store)) = store(data, limits) {
        let _ = all(&store, None);
        let _ = all(&store, Some(data));
        if let Ok(id) = Id::new(1) {
            let _ = get(&store, id, None);
        }
        for file in all(&store, Some(data)).unwrap_or_default() {
            let _ = file.blip();
            let _ = file.data();
            let _ = file.into_owned();
        }
    }

    let parser = Parser::new(data, limits.record_limits());
    for parsed in parser.records().take(20_000) {
        let Ok(raw) = parsed else {
            break;
        };
        let _ = record(&raw, limits);
        let _ = litchi_odraw::image::record_with_delay(&raw, Some(data), limits);
    }
});
