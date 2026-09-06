#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_opc::OpcPackage;

fuzz_target!(|data: &[u8]| {
    // Primary entrypoint: parse a ZIP-backed OPC package from a byte slice.
    if let Ok(pkg) = OpcPackage::from_bytes(data) {
        // Exercise package-level relationships.
        let rels = pkg.rels();
        for rel in rels.iter() {
            let _ = rel.r_id();
            let _ = rel.reltype();
            let _ = rel.target_ref();
            let _ = rel.is_external();
            // Resolving target partname can fail on malformed inputs; ignore.
            let _ = rel.target_partname();
        }

        // Iterate all parts and inspect surface API.
        for part in pkg.iter_parts() {
            let _ = part.partname();
            let _ = part.content_type();
            let _ = part.blob().len();

            // Walk part-level relationships too.
            let prels = part.rels();
            for rel in prels.iter() {
                let _ = rel.r_id();
                let _ = rel.reltype();
                let _ = rel.target_ref();
                let _ = rel.is_external();
                let _ = rel.target_partname();
            }
        }

        // Try to follow the main document relationship if present.
        let _ = pkg.main_document_part();
        let _ = pkg.part_count();
    }
});
