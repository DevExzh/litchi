#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_odraw::Parser;
use litchi_odraw::image::write::copy;
use litchi_odraw::image::{Blip, Id, all, delay, get, record, scan, store};

fuzz_target!(|data: &[u8]| {
    let limits = common::art_limits();
    let _ = scan(data, limits);

    // A parsed BLIP survives being copied record-for-record: the copy parses
    // again to the same kind, instance, UIDs and file data.
    if let Ok(blip) = Blip::parse(data, limits) {
        let mut copied = Vec::new();
        if copy(&mut copied, &blip, limits).is_ok() {
            let again = Blip::parse(&copied, limits).expect("a copied BLIP parses again");
            assert_eq!(again.raw_kind(), blip.raw_kind());
            assert_eq!(again.record().instance(), blip.record().instance());
            assert_eq!(again.uids(), blip.uids());
            assert_eq!(again.data(), blip.data());
        }
    }
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
